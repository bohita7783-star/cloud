"""勤務日報(.xls)の書き方傾向を再集計するスクリプト（再現用）。

使い方:  python3 analyze_nippo.py <日報xlsが入ったフォルダ>
依存:    pip install xlrd
注意:    日報データそのものは出力・保存しない（集計値のみ表示）。同一内容のファイルは1回だけ数える。
前提:    日別シート名が (1)〜(31)、内容欄=Q列(index16)、時刻=G/M列、作業者=BS列、
         ブロック名=D列(「日常業務」「不動産関係」「テナント関係」「工事関係」「特記事項」)
"""
import collections
import glob
import hashlib
import re
import sys
import unicodedata

import xlrd


def read_rows(folder):
    seen, rows = set(), []
    for f in sorted(glob.glob(folder.rstrip("/") + "/*.xls")):
        h = hashlib.md5(open(f, "rb").read()).hexdigest()
        if h in seen:
            continue
        seen.add(h)
        wb = xlrd.open_workbook(f)
        for sh in wb.sheets():
            if not re.fullmatch(r"\(\d+\)", sh.name):
                continue

            def cell(r, c):
                try:
                    return sh.cell_value(r, c)
                except IndexError:
                    return ""

            block = None
            for r in range(15, 56):
                lab = str(cell(r, 3)).replace("　", "").strip()
                if lab:
                    block = lab
                t = str(cell(r, 16)).strip()
                if t:
                    rows.append((block, t))
    return len(seen), rows


def main(folder):
    nfiles, rows = read_rows(folder)
    body = [r for r in rows if r[0] != "日常業務"]
    print(f"ファイル数(重複除外)={nfiles}  記入行={len(rows)}  日常業務除く={len(body)}")
    print("ブロック別:", dict(collections.Counter(b for b, _ in rows)))
    lead = collections.Counter(t[0] for _, t in body)
    print("先頭「・」:", lead["・"], " 先頭「【」:", lead["【"], " 先頭「➡」:", lead["➡"])
    al = " ".join(t for _, t in body)
    print("異常なし/異常無し/異常無:", al.count("異常なし"), al.count("異常無し"),
          len(re.findall("異常無(?!し)", al)))

    def w(s):
        return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)

    ws = sorted(w(t) for _, t in body)
    print("表示幅 p50/p90/p99/max:", ws[len(ws) // 2], ws[int(len(ws) * .9)], ws[int(len(ws) * .99)], ws[-1])


if __name__ == "__main__":
    main(sys.argv[1])
