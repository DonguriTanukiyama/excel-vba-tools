# -*- coding: utf-8 -*-
u"""Excel に取り込む前に、.bas を機械で確かめる。

  python tools/check_bas.py           # リポジトリの .bas を全部（フォルダごとに見る）
  python tools/check_bas.py --demo    # この検査自体の自己確認

見るもの
  1. Shift-JIS（cp932）で保存されていて、改行が CRLF か。VBA エディタは UTF-8 の日本語を読めない
  2. 変数・引数・プロシージャの名前が、VBA の予約語と同じでないか。**大文字と小文字は区別されない。**
     2026-09-13、`Dim eNum As Long` が Excel で「構文エラー」になった。VBA には eNum と Enum の区別が無い
  3. 1つのプロシージャの中で、同じ名前を2回宣言していないか（引数と変数の重なりも）
  4. 同じフォルダのモジュールに、同じ名前の Public プロシージャが無いか
     （同じフォルダの .bas は同じブックに取り込む前提。呼ぶ側で名前があいまいになる）

**コンパイルの代わりにはならない。**型の食い違い・宣言していない変数・実行時の誤りは、Excel で動かすまで分からない。
"""
import glob
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

# VBA の言語仕様（MS-VBAL の Reserved Identifiers）から
RESERVED = frozenset(w.lower() for w in u"""
Abs AddressOf And Any Array As Attribute Boolean ByRef Byte ByVal Call Case CBool CByte CCur CDate
CDbl CDec CDecl CInt Circle CLng CLngLng CLngPtr Close Const CSng CStr Currency CVar CVErr Date Debug
Decimal Declare DefBool DefByte DefCur DefDate DefDbl DefDec DefInt DefLng DefLngLng DefLngPtr DefObj
DefSng DefStr DefVar Dim Do DoEvents Double Each Else ElseIf Empty End EndIf Enum Eqv Erase Event Exit
False Fix For Friend Function Get Global GoSub GoTo If Imp Implements In Input InputB Int Integer Is
LBound Len LenB Let Like LineInput Lock Long LongLong LongPtr Loop LSet Me Mod New Next Not Nothing
Null On Open Option Optional Or ParamArray Preserve Print Private PSet Public Put RaiseEvent ReDim Rem
Resume Return RSet Scale Seek Select Set Sgn Shared Single Spc Static Stop String Sub Tab Then To True
Type TypeOf UBound Unlock Until Variant Wend While With WithEvents Write Xor
""".split())

PROC_HEAD = re.compile(r"^(?:(Public|Private|Friend)\s+)?(?:Static\s+)?"
                       r"(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)\s*\(", re.I)
PROC_END = re.compile(r"^End\s+(?:Sub|Function|Property)\b", re.I)
# ReDim は宣言済みの配列の作り直しなので数えない
DECL = re.compile(r"^(?:Dim|Static|(?:(?:Private|Public|Global)\s+)?Const|"
                  r"(?:Private|Public|Global)(?!\s+(?:Sub|Function|Property|Declare|Enum|Type|Event|Static|Const)\b))"
                  r"\s+(.+)$", re.I)
FOR = re.compile(r"^For\s+(?:Each\s+)?(\w+)", re.I)
NAME = re.compile(r"^(?:(?:Optional|ByVal|ByRef|ParamArray|WithEvents)\s+)*(\w+)", re.I)


def strip_line(line):
    u"""文字列の中身と、' から後ろの注釈を消す（"" は残す）。"""
    out, in_str, i = [], False, 0
    while i < len(line):
        ch = line[i]
        if in_str:
            if ch == u'"':
                if line[i + 1:i + 2] == u'"':
                    i += 2
                    continue
                in_str = False
                out.append(ch)
        elif ch == u'"':
            in_str = True
            out.append(ch)
        elif ch == u"'":
            break
        else:
            out.append(ch)
        i += 1
    return u"".join(out)


def code_lines(text):
    u"""注釈と文字列を消し、行の続き（ _）をつないだ行。[(最初の行番号, 行)]"""
    out, buf, first = [], u"", 0
    for no, raw in enumerate(text.split(u"\n"), 1):
        s = strip_line(raw.rstrip(u"\r")).rstrip()
        if not buf:
            first = no
        if s == u"_" or s.endswith(u" _"):
            buf += s[:-1] + u" "
            continue
        out.append((first, (buf + s).strip()))
        buf = u""
    return out


def statements(line):
    u"""「:」で文に分ける。名前つき引数の「:=」では分けない。"""
    return [s.strip() for s in re.split(u":(?!=)", line) if s.strip()]


def split_top(s):
    u"""かっこの外のカンマで分ける。"""
    parts, depth, cur = [], 0, u""
    for ch in s:
        if ch == u"(":
            depth += 1
        elif ch == u")":
            depth -= 1
        if ch == u"," and depth == 0:
            parts.append(cur)
            cur = u""
        else:
            cur += ch
    parts.append(cur)
    return [p.strip() for p in parts if p.strip()]


def params_of(st, start):
    u"""プロシージャの ( の直後から、対応する ) の手前まで。"""
    depth, i = 1, start
    while i < len(st) and depth:
        if st[i] == u"(":
            depth += 1
        elif st[i] == u")":
            depth -= 1
        i += 1
    return st[start:i - 1]


def module_problems(name, data):
    u"""1つのモジュールの問題の一覧と、Public プロシージャ [(名前, 場所)] を返す。"""
    bad, public = [], []
    if data.count(b"\n") != data.count(b"\r\n"):
        bad.append(u"%s: 改行が CRLF でない（.gitattributes の取り出し設定を確かめる）" % name)
    try:
        data.decode("utf-8")
        if any(b > 0x7F for b in bytearray(data)):
            bad.append(u"%s: UTF-8 のまま（VBA エディタで日本語が化ける。Shift-JIS で保存する）" % name)
            return bad, public
    except UnicodeDecodeError:
        pass
    try:
        text = data.decode("cp932")
    except UnicodeDecodeError as e:
        bad.append(u"%s: Shift-JIS として読めない（%s）" % (name, e))
        return bad, public

    def declare(item, no, scope):
        m = NAME.match(item)
        if not m:
            return
        nm, where = m.group(1), u"%s:%d" % (name, no)
        if nm.lower() in RESERVED:
            bad.append(u"%s: 名前が予約語と同じ: %s（VBA は大文字と小文字を区別しない）" % (where, nm))
        if scope is None:
            return
        if nm.lower() in scope:
            bad.append(u"%s: 同じ名前を2回宣言している: %s（先の宣言は %d 行目）" % (where, nm, scope[nm.lower()]))
        else:
            scope[nm.lower()] = no

    scope = None
    for no, line in code_lines(text):
        for st in statements(line):
            m = PROC_HEAD.match(st)
            if m:
                proc, where = m.group(2), u"%s:%d" % (name, no)
                if proc.lower() in RESERVED:
                    bad.append(u"%s: プロシージャの名前が予約語と同じ: %s" % (where, proc))
                if (m.group(1) or u"Public").lower() == u"public":
                    public.append((proc, where))
                scope = {}
                for p in split_top(params_of(st, m.end())):
                    declare(p, no, scope)
            elif PROC_END.match(st):
                scope = None
            elif DECL.match(st):
                for p in split_top(DECL.match(st).group(1)):
                    declare(p, no, scope)
            elif FOR.match(st):
                declare(FOR.match(st).group(1), no, None)   # For の変数は宣言ではない。予約語だけ見る
    return bad, public


def problems(files):
    u"""files = {ファイル名: バイト列}。問題の一覧を返す。空なら問題なし。"""
    out, public = [], {}
    for name in sorted(files):
        bad, pub = module_problems(name, files[name])
        out += bad
        for proc, where in pub:
            public.setdefault(proc.lower(), []).append((proc, where))
    for k in sorted(public):
        if len(public[k]) > 1:
            out.append(u"同じ名前の Public プロシージャが複数のモジュールにある: %s（%s）"
                       % (public[k][0][0], u" / ".join(w for _, w in public[k])))
    return out


def demo():
    def mod(*lines):
        return u"\r\n".join(lines + (u"",)).encode("cp932")

    ok = mod(u'Attribute VB_Name = "A"', u"Option Explicit",
             u"Public Sub Run(ByVal path As String, ByRef rec() As Variant)",
             u'    Dim errNo As Long, errText As String: Const K = "a:b"',
             u'    MsgBox "eNum は文字の中なら問題ない", Title:="t"   \' Dim Enum も注釈の中なら問題ない',
             u"    ReDim rec(1 To 2)",
             u"End Sub")
    assert problems({u"A.bas": ok}) == [], problems({u"A.bas": ok})

    # 実際に Excel で「構文エラー」になった行
    got = problems({u"B.bas": mod(u'Attribute VB_Name = "B"',
                                  u"Public Function ReadTable(ByVal path As String) As Variant",
                                  u"    Dim eNum As Long, eDesc As String",
                                  u"End Function")})
    assert len(got) == 1 and u"B.bas:3" in got[0] and u": eNum（" in got[0], got

    # 行の続き（ _）をまたいだ引数と、変数の重なり
    got = problems({u"C.bas": mod(u'Attribute VB_Name = "C"',
                                  u"Private Sub Fill(ByVal r As Long, _",
                                  u"                 ByRef rec() As Variant)",
                                  u"    Dim i As Long, r As Long",
                                  u"End Sub")})
    assert len(got) == 1 and u": r（" in got[0], got

    # 2つのモジュールに同じ Public プロシージャ（Public を書かない Sub も Public）
    got = problems({u"D.bas": mod(u'Attribute VB_Name = "D"', u"Public Sub SelfTest()", u"End Sub"),
                    u"E.bas": mod(u'Attribute VB_Name = "E"', u"Sub SelfTest()", u"End Sub",
                                  u"Private Sub Check()", u"End Sub"),
                    u"F.bas": mod(u'Attribute VB_Name = "F"', u"Private Sub Check()", u"End Sub")})
    assert len(got) == 1 and u"SelfTest" in got[0], got

    # UTF-8 のまま / 改行が LF
    got = problems({u"G.bas": u"Attribute VB_Name = \"G\"\r\n' 日本語\r\n".encode("utf-8")})
    assert len(got) == 1 and u"UTF-8" in got[0], got
    got = problems({u"H.bas": b'Attribute VB_Name = "H"\nOption Explicit\n'})
    assert len(got) == 1 and u"CRLF" in got[0], got
    print(u"demo OK（予約語 eNum / 行の続きをまたぐ重複宣言 / Public の重なり / UTF-8 / LF / 文字列と注釈は見ない）")


def main():
    if "--demo" in sys.argv:
        return demo()
    paths = [a for a in sys.argv[1:] if not a.startswith("--")] or sorted(
        glob.glob(os.path.join(ROOT, u"**", u"*.bas"), recursive=True))
    # 同じフォルダの .bas は、同じブックに取り込む前提で一緒に見る
    groups = {}
    for p in paths:
        groups.setdefault(os.path.dirname(os.path.abspath(p)), []).append(p)
    bad, seen = [], []
    for d in sorted(groups):
        files = dict((os.path.relpath(os.path.abspath(p), ROOT).replace(os.sep, u"/"), io.open(p, "rb").read())
                     for p in groups[d])
        bad += problems(files)
        seen += [u"%s %d行" % (n, files[n].count(b"\n")) for n in sorted(files)]
    if not seen:
        print(u"見つからない: .bas が1つも無い（%s）" % ROOT)
        sys.exit(2)
    if bad:
        print(u"**問題 %d 件**" % len(bad))
        for b in bad:
            print(u"  " + b)
        sys.exit(1)
    print(u"OK: %s（予約語・同じ名前の宣言・文字コード・改行）" % u" / ".join(seen))


if __name__ == "__main__":
    main()
