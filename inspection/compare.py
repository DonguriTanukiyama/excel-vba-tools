# -*- coding: utf-8 -*-
"""VBA の出力と、正解（oracle.py）を1セルずつ突き合わせる。

  python inspection/oracle.py                     # 正解を作る  → <元の名前>_正解.csv / _報告書_正解.csv
  （Excel で GenbaNorm.Normalize を実行）  # VBA の出力 → <元の名前>_VBA.csv
  python inspection/compare.py                    # 整えた表を突き合わせる
  （Excel で GenbaReport.MakeReports）     # VBA の出力 → <元の名前>_報告書_VBA.csv と PDF
  python inspection/compare.py --report           # 報告書を突き合わせ、PDF が日数ぶんあるか数える
  python inspection/compare.py --demo             # この突き合わせ自体の自己確認

一致しないセルは、行・列・正解・VBA の値を全部出す。**件数だけで済ませない。**
"""
import csv
import io
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT = os.path.join(HERE, u"点検記録_2026-09.xlsx")


def load(path):
    with io.open(path, encoding="utf-8-sig", newline="") as f:
        return [row for row in csv.reader(f) if row]


def trim(rows):
    u"""報告書の CSV は行ごとに列の数が違う。後ろの空欄は書き出す側で付いたり付かなかったりするので、落として比べる。"""
    out = []
    for r in rows:
        r = list(r)
        while r and r[-1] == u"":
            r.pop()
        out.append(r)
    return out


def diff(want, got):
    u"""返すのは不一致の一覧 [(行, 列, 正解, VBA), ...]。空なら全部一致。"""
    bad = []
    if want[0] != got[0]:
        bad.append((u"見出し", u"", u" | ".join(want[0]), u" | ".join(got[0])))
    w = dict((r[0], r) for r in want[1:])
    g = dict((r[0], r) for r in got[1:])
    for k in sorted(set(w) | set(g), key=lambda x: (int(x) if x.isdigit() else 10 ** 9, x)):
        if k not in g:
            bad.append((k, u"（行）", u"ある", u"VBA に無い"))
            continue
        if k not in w:
            bad.append((k, u"（行）", u"無い", u"VBA だけにある"))
            continue
        for i in range(max(len(w[k]), len(g[k]))):
            a = w[k][i] if i < len(w[k]) else u"（列なし）"
            b = g[k][i] if i < len(g[k]) else u"（列なし）"
            if a != b:
                col = want[0][i] if i < len(want[0]) else u"%d列目" % (i + 1)
                bad.append((k, col, a, b))
    return bad


def pdfs(stem, want):
    u"""日ごとの PDF があるか。返すのは (見つからない・壊れている一覧, 各 PDF のページ数)。"""
    days = []
    for r in want[1:]:
        d = r[0].split(u"/")[0]
        if d not in days:
            days.append(d)
    try:
        from pypdf import PdfReader
    except ImportError:
        PdfReader = None
    bad, pages = [], {}
    for d in days:
        p = os.path.join(stem + u"_点検報告書", (u"日付不明" if d == u"読めない" else d) + u".pdf")
        if not os.path.exists(p):
            bad.append(u"PDF が無い: %s" % p)
            continue
        with open(p, "rb") as f:
            if f.read(5) != b"%PDF-":
                bad.append(u"PDF ではない: %s" % p)
                continue
        if PdfReader:
            pages[d] = len(PdfReader(p).pages)
    return bad, pages


def demo():
    head = [u"行", u"日付", u"担当"]
    want = [head, [u"3", u"2026-09-01", u"山田太郎"], [u"4", u"2026-09-02", u"佐藤花子"]]
    assert diff(want, [list(r) for r in want]) == [], u"同じ表で不一致が出た"
    got = [head, [u"3", u"2026/09/01", u"山田太郎"], [u"4", u"2026-09-02", u"佐藤花子"]]
    d = diff(want, got)
    assert d == [(u"3", u"日付", u"2026-09-01", u"2026/09/01")], d
    got = [head, [u"3", u"2026-09-01", u"山田太郎"]]
    assert diff(want, got) == [(u"4", u"（行）", u"ある", u"VBA に無い")], diff(want, got)
    got = [head, [u"3", u"2026-09-01"], [u"4", u"2026-09-02", u"佐藤花子"]]
    assert diff(want, got) == [(u"3", u"担当", u"山田太郎", u"（列なし）")], diff(want, got)
    assert trim([[u"a", u"", u"b", u"", u""], [u"c"]]) == [[u"a", u"", u"b"], [u"c"]]
    print(u"demo OK（一致 / 値の違い / 行の欠け / 列の欠け / 後ろの空欄）")


def main():
    if "--demo" in sys.argv:
        return demo()
    rep = "--report" in sys.argv
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    stem = os.path.splitext(args[0] if args else DEFAULT)[0]
    tag = u"_報告書" if rep else u""
    want_p, got_p = stem + tag + u"_正解.csv", stem + tag + u"_VBA.csv"
    for p in (want_p, got_p):
        if not os.path.exists(p):
            print(u"見つからない: %s" % p)
            sys.exit(2)
    want, got = load(want_p), load(got_p)
    if rep:
        want, got = trim(want), trim(got)
    bad = diff(want, got)
    cells = sum(len(r) - 1 if rep else len(r) for r in want[1:])
    if bad:
        print(u"**不一致 %d 件**（正解 %d行）" % (len(bad), len(want) - 1))
        for k, col, a, b in bad:
            print(u"  %s  %s  正解=[%s]  VBA=[%s]" % (k, col, a, b))
        sys.exit(1)
    if not rep:
        print(u"VBA == Python（%d行 × %d列・全%dセル一致）" % (len(want) - 1, len(want[0]), cells))
        return

    missing, pages = pdfs(stem, want)
    days = len(set(r[0].split(u"/")[0] for r in want[1:]))
    print(u"報告書 VBA == Python（%d日分・%d欄・値 %dセル一致）" % (days, len(want) - 1, cells))
    if missing:
        print(u"**PDF の問題 %d 件**" % len(missing))
        for m in missing:
            print(u"  " + m)
        sys.exit(1)
    if pages:
        multi = dict((d, n) for d, n in pages.items() if n != 1)
        print(u"PDF %d 件。%s" % (len(pages), u"すべて1ページ" if not multi else u"2ページ以上: %s" % multi))
    else:
        print(u"PDF %d 件（pypdf が無いのでページ数は数えていない）" % days)


if __name__ == "__main__":
    main()
