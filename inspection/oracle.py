# -*- coding: utf-8 -*-
"""VBA 版の「正解」を作る。

同じフォルダの genba_norm.py（rag-ja からの写し）の関数（parse_date / canon / parse_num / classify）をそのまま使い、
VBA 版と同じ内容を CSV に出す。compare.py がこの CSV と VBA の出力を1セルずつ突き合わせる。

  python inspection/oracle.py                   # 同じフォルダの 点検記録_2026-09.xlsx
  python inspection/oracle.py <点検表の xlsx>

出すもの
  <元の名前>_正解.csv         整えた表（GenbaNorm.Normalize の正解）
  <元の名前>_報告書_正解.csv  日ごとの報告書に書くはずの内容（GenbaReport.MakeReports の正解）

ponytail: 行のループは genba_norm.norm_inspection() と同じ規則を写してある（十数行）。
          norm_inspection() を行単位の関数に割ったら、ここはそれを呼ぶだけにする。
"""
import csv
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import genba_norm as G  # noqa: E402

CHECK_KEYS = (u"掻き取り", u"異音", u"発泡", u"色")
FLAGGED = (G.ABNORMAL, G.CONDITIONAL, G.BLANK, G.AMBIGUOUS)   # 「確認が必要な項目」に載せる順


def _head(rows):
    u"""2段の見出しをつないだ列名と、点検欄の列・備考の列。"""
    head1, head2 = rows[0], rows[1]
    cols = [(G.z2h(a).strip() + u" " + G.z2h(b).strip()).strip()
            for a, b in zip(head1, list(head2) + [u""] * len(head1))]
    check = [i for i, c in enumerate(cols) if any(k in c for k in CHECK_KEYS)]
    note = next((i for i, c in enumerate(cols) if u"備考" in c and i not in check), None)
    return cols, check, note


def table(path):
    rows, _, reader = G.read_xlsx(path)
    base = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"(\d{4})-(\d{2})", base)
    dy, dm = (int(m.group(1)), int(m.group(2))) if m else (None, None)
    cols, check, _ = _head(rows)

    out = [[u"行", u"日付", u"担当", u"部署", u"汚泥厚"] + [cols[i] for i in check]]
    for r, row in enumerate(rows[2:], start=3):
        if not any((u"%s" % x).strip() for x in row):
            continue
        d = G.parse_date(row[0], dy, dm)
        num = G.parse_num(row[2])
        out.append([u"%d" % r,
                    d.isoformat() if d else u"読めない",
                    G.canon(u"担当", row[1]),
                    G.canon(u"部署", row[7]) if len(row) > 7 else u"",
                    u"%.1f" % num if num is not None else u"読めない"]
                   + [G.classify(row[i] if i < len(row) else u"")[0] for i in check])
    return out, reader


def _join(names):
    u"""出てきた順に、重ねずに「、」でつなぐ（GenbaReport の AddUnique と同じ）。"""
    seen = []
    for x in names:
        if x and x not in seen:
            seen.append(x)
    return u"、".join(seen)


def report(path):
    u"""GenbaReport が報告書のセルに書くはずの内容。[欄の名前, 値, ...] の行を返す。"""
    out, _ = table(path)
    rows = G.read_xlsx(path)[0]
    _, check, note = _head(rows)
    head, body = out[0], out[1:]
    nck = len(check)

    def raw(b):
        u"""点検欄の原文と備考。VBA の ReadTable が右側の列に持たせているもの。"""
        row = rows[int(b[0]) - 1]
        texts = [G.classify(row[i] if i < len(row) else u"")[1] for i in check]
        memo = (u"%s" % row[note]).strip() if note is not None and note < len(row) else u""
        return texts, memo

    days = []
    for b in body:
        if b[1] not in days:
            days.append(b[1])

    rec = [[u"欄", u"値"]]
    for day in days:
        today = [b for b in body if b[1] == day]
        name = (day + u"/%s").__mod__
        rec.append([name(u"日付"), day])
        rec.append([name(u"部署"), _join(b[3] for b in today)])
        rec.append([name(u"担当"), _join(b[2] for b in today)])
        rec.append([name(u"点検回数"), u"%d" % len(today)])
        for kind in G.KINDS:
            rec.append([name(kind), u"%d" % sum(b[5:5 + nck].count(kind) for b in today)])

        items = []
        for kind in FLAGGED:
            for b in today:
                texts = raw(b)[0]
                items += [[kind, head[5 + k], texts[k], b[0], b[2]]
                          for k in range(nck) if b[5 + k] == kind]
        if not items:
            rec.append([name(u"確認が必要な項目/1"), u"なし"])
        for i, it in enumerate(items, 1):
            rec.append([name(u"確認が必要な項目/%d" % i)] + it)

        rec.append([name(u"点検結果/0"), u"原票の行", u"担当", head[4]] + head[5:5 + nck] + [u"備考"])
        for i, b in enumerate(today, 1):
            rec.append([name(u"点検結果/%d" % i), b[0], b[2], b[4]] + b[5:5 + nck] + [raw(b)[1]])
    return rec


def write(dst, rows):
    with io.open(dst, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(rows)


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, u"点検記録_2026-09.xlsx")
    stem = os.path.splitext(src)[0]

    out, reader = table(src)
    write(stem + u"_正解.csv", out)
    dates = sum(1 for r in out[1:] if r[1] != u"読めない")
    print(u"%s  %d行（日付確定 %d）  読み取り: %s" % (stem + u"_正解.csv", len(out) - 1, dates, reader))

    rec = report(src)
    write(stem + u"_報告書_正解.csv", rec)
    days = len(set(r[0].split(u"/")[0] for r in rec[1:]))
    print(u"%s  %d日分・%d欄" % (stem + u"_報告書_正解.csv", days, len(rec) - 1))


if __name__ == "__main__":
    main()
