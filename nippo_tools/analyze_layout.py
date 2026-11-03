"""日報(.xls)の日別シートで、分類見出しの行位置と記入欄の容量・満杯率を集計する（再現用）。

使い方:  python3 analyze_layout.py <日報xlsが入ったフォルダ>
依存:    pip install xlrd   （日報の中身は出力しない。集計値のみ）
前提:    分類見出しは D 列、内容は Q 列(index16)、末尾は A 列の「発注品」行。
"""
import collections
import glob
import hashlib
import re
import sys

import xlrd

LABELS = ("不動産関係", "テナント関係", "工事関係", "特記事項")


def main(folder):
    caps = collections.defaultdict(collections.Counter)
    full, count, patterns = collections.Counter(), collections.Counter(), collections.Counter()
    seen, sheets, missing = set(), 0, 0
    for f in sorted(glob.glob(folder.rstrip("/") + "/*.xls")):
        h = hashlib.md5(open(f, "rb").read()).hexdigest()
        if h in seen:
            continue
        seen.add(h)
        for sh in xlrd.open_workbook(f).sheets():
            if not re.fullmatch(r"\(\d+\)", sh.name):
                continue
            sheets += 1

            def cv(r, c):
                try:
                    return sh.cell_value(r, c)
                except IndexError:
                    return ""

            pos = {}
            for r in range(15, 70):
                lab = str(cv(r, 3)).replace("　", "").strip()
                if lab in LABELS and lab not in pos:
                    pos[lab] = r
            if len(pos) < 4:
                missing += 1
                continue
            patterns[tuple(sorted(pos.items(), key=lambda x: x[1]))] += 1
            end = next((r for r in range(40, 80)
                        if str(cv(r, 0)).replace("　", "").strip().startswith("発注品")), None)
            order = sorted(pos.items(), key=lambda x: x[1])
            for i, (lab, s) in enumerate(order):
                e = order[i + 1][1] if i + 1 < len(order) else (end or s + 10)
                cap = e - s
                filled = sum(1 for r in range(s, e) if str(cv(r, 16)).strip())
                caps[lab][cap] += 1
                count[lab] += 1
                full[lab] += filled >= cap
    print(f"日別シート={sheets} 見出し欠落={missing} 見出し位置のパターン数={len(patterns)}")
    for k, v in patterns.most_common(3):
        print("  ", v, "日:", [(l, r + 1) for l, r in k])
    for lab in LABELS:
        print(f"{lab}: 容量(行)別の日数={dict(caps[lab])}  満杯の日={full[lab]}/{count[lab]} ({full[lab] / count[lab]:.1%})")


if __name__ == "__main__":
    main(sys.argv[1])
