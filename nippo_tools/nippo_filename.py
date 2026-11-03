"""日報ファイル名から対象年月を安全に抽出する。"""
import os
import re


def year_month_from_filename(path):
    """例: 2026_11__xxx.xls / 【名店街】日報（2026年11月）.xls / 2026-11.xls"""
    name = os.path.basename(path)
    m = re.search(r"(?<!\d)(20\d{2})[_-](\d{1,2})(?!\d)|(?<!\d)(20\d{2})年(\d{1,2})月", name)
    if m is None:
        raise ValueError(f"ファイル名から年・月を識別できません: {name}")
    year = int(m.group(1) or m.group(3))
    month = int(m.group(2) or m.group(4))
    if not 1 <= month <= 12:
        raise ValueError(f"月が範囲外です: {name}")
    return year, month
