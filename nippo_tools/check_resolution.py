"""実際の日報(.xls)の全日別シートに対し、追記先の決定(nippo_target.resolve)を読み取り専用で試す。

使い方:  python3 check_resolution.py <日報xlsが入ったフォルダ>
出力:    件数の集計のみ（日報の内容は出さない）。ファイルは一切書き換えない。
ファイル名: 年月は名前から読む（例: 2026_11__x.xls／【名店街】日報（2026年11月）.xls／2026-11.xls）。
判定:    OK=自動転記してよい / Candidate=設備員の承認が必要 / NeedsReview=自動処理を止める
"""
import collections
import glob
import hashlib
import re
import sys

import xlrd

from nippo_filename import year_month_from_filename
from nippo_target import SECTIONS, resolve

MODES = (("末尾のみ", False), ("末尾→候補", True))


def main(folder):
    seen, total = set(), 0
    status = collections.defaultdict(collections.Counter)
    reasons = collections.Counter()
    for f in sorted(glob.glob(folder.rstrip("/") + "/*.xls")):
        h = hashlib.md5(open(f, "rb").read()).hexdigest()
        if h in seen:
            continue
        seen.add(h)
        year, month = year_month_from_filename(f)
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
                for mode, cand in MODES:
                    res = resolve(cell, sec, 1, sheet_name=sh.name, year=year, month=month,
                                  day=day, offer_candidates=cand)
                    status[(sec, mode)][res.status] += 1
                    if res.status == "NeedsReview":
                        reasons[(mode, re.sub(r"[0-9]+", "N", res.reason))] += 1
    print(f"日別シート={total}（1行追記を各分類に試行）")
    for sec in SECTIONS:
        for mode, _ in MODES:
            d = status[(sec, mode)]
            print(f"  {sec} [{mode}]: OK={d['OK']} Candidate(要承認)={d['Candidate']} "
                  f"NeedsReview={d['NeedsReview']} ({d['NeedsReview'] / total:.1%})")
    print("NeedsReview の理由:")
    for (mode, why), n in sorted(reasons.items()):
        print(f"  [{mode}] {why}: {n}")


if __name__ == "__main__":
    main(sys.argv[1])
