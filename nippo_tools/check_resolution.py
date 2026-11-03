"""実際の日報(.xls)の全日別シートに対し、追記先の決定(nippo_target.resolve)を読み取り専用で試す。

使い方:  python3 check_resolution.py <日報xlsが入ったフォルダ>
出力:    件数の集計のみ（日報の内容は出さない）。ファイルは一切書き換えない。
"""
import collections
import glob
import hashlib
import re
import sys

import xlrd

from nippo_target import SECTIONS, resolve


def main(folder):
    seen, total = set(), 0
    status = collections.defaultdict(collections.Counter)
    reasons = collections.Counter()
    for f in sorted(glob.glob(folder.rstrip("/") + "/*.xls")):
        h = hashlib.md5(open(f, "rb").read()).hexdigest()
        if h in seen:
            continue
        seen.add(h)
        month = int(re.search(r"2026_(\d+)__", f).group(1))
        for sh in xlrd.open_workbook(f).sheets():
            if not re.fullmatch(r"\(\d+\)", sh.name):
                continue
            day = int(sh.name.strip("()"))

            def cell(r, c, sh=sh):
                try:
                    return sh.cell_value(r, c)
                except IndexError:
                    return ""

            total += 1
            for sec in SECTIONS:
                for mode, gaps in (("最後の次", False), ("空き行も使う", True)):
                    res = resolve(cell, sec, 1, sheet_name=sh.name, year=2026, month=month, day=day, fill_gaps=gaps)
                    status[(sec, mode)][res.status] += 1
                    if res.status != "OK":
                        reasons[(mode, re.sub(r"[0-9]+", "N", res.reason))] += 1
    print(f"日別シート={total}（1行追記を各分類に試行）")
    for sec in SECTIONS:
        for mode in ("最後の次", "空き行も使う"):
            d = status[(sec, mode)]
            print(f"  {sec} [{mode}]: OK={d['OK']} NeedsReview={d['NeedsReview']} ({d['NeedsReview'] / total:.1%})")
    print("NeedsReview の理由:", dict(reasons))


if __name__ == "__main__":
    main(sys.argv[1])
