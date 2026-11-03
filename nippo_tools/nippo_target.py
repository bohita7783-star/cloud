"""日報の「追記先の決定」基準実装（Excel非依存。読み取り関数 cell(r, c) を渡して使う）。

セルの座標は 0 始まり（r=行, c=列）。実機(Excel COM等)では同じ形の読み取り関数を用意する。
ルール（仕様書v0.1＋突き合わせ結果）:
  - 日付は「対象年・月・日・シート名」をセットで照合（年 A4 / 月 K4 / 日 S4 / シート名 "(日)"）。
  - 分類見出しは D 列を毎回探索（固定の行番号は使わない）。見出しが欠落・重複なら NeedsReview。
  - 記入欄 = 見出し行〜次の見出し(最後は A 列「発注品」)の直前。
  - 追記位置 = 「最後に記入のある行」の次。対象行の G/M/Q/BS が空であること。
  - 複数行を書く場合は連続した空き行が必要。足りなければ NeedsReview（行挿入はしない）。
"""
from dataclasses import dataclass
from typing import Callable, Optional

SECTIONS = ("不動産関係", "テナント関係", "工事関係", "特記事項")
C_LABEL, C_G, C_M, C_Q, C_BS, C_A = 3, 6, 12, 16, 70, 0
R_DATE = 3                      # 4行目
C_YEAR, C_MONTH, C_DAY = 0, 10, 18
SCAN_FROM, SCAN_TO = 15, 70


@dataclass
class Resolution:
    status: str                 # "OK" / "NeedsReview"
    rows: tuple = ()            # 書き込む行（0始まり）
    reason: str = ""


def _s(v) -> str:
    return "" if v is None else str(v).replace("　", " ").strip()


def _num(v) -> Optional[int]:
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def check_date(cell: Callable, sheet_name: str, year: int, month: int, day: int) -> Optional[str]:
    """一致すれば None、不一致なら理由を返す。"""
    if sheet_name.strip() != f"({day})":
        return f"シート名 {sheet_name!r} が日 {day} と不一致"
    got = (_num(cell(R_DATE, C_YEAR)), _num(cell(R_DATE, C_MONTH)), _num(cell(R_DATE, C_DAY)))
    if got != (year, month, day):
        return f"表内の日付 {got} が対象 {(year, month, day)} と不一致"
    return None


def find_sections(cell: Callable) -> tuple:
    """(見出し行の辞書, エラー理由)。各見出しがちょうど1回出ることを要求。"""
    pos, errs = {}, []
    for r in range(SCAN_FROM, SCAN_TO):
        lab = _s(cell(r, C_LABEL))
        if lab in SECTIONS:
            if lab in pos:
                errs.append(f"見出し {lab} が重複")
            pos[lab] = pos.get(lab, r)
    for lab in SECTIONS:
        if lab not in pos:
            errs.append(f"見出し {lab} が見つからない")
    return pos, "; ".join(errs)


def _end_row(cell: Callable, start: int) -> Optional[int]:
    for r in range(start, SCAN_TO + 10):
        if _s(cell(r, C_A)).startswith("発注品"):
            return r
    return None


def _free(cell: Callable, r: int) -> bool:
    return not any(_s(cell(r, c)) for c in (C_G, C_M, C_Q, C_BS))


def resolve(cell: Callable, section: str, n_lines: int, *, sheet_name: str,
            year: int, month: int, day: int, fill_gaps: bool = False) -> Resolution:
    """fill_gaps=False（既定）: 最後の記入の次の行に追記（時系列を崩さない・保守的）。
    fill_gaps=True: 記入欄の途中の空き行（連続 n_lines 行）も使う。"""
    bad = check_date(cell, sheet_name, year, month, day)
    if bad:
        return Resolution("NeedsReview", reason=bad)
    pos, err = find_sections(cell)
    if err:
        return Resolution("NeedsReview", reason=err)
    if section not in pos:
        return Resolution("NeedsReview", reason=f"未知の分類 {section!r}")
    order = sorted(pos.values())
    start = pos[section]
    nxt = [p for p in order if p > start]
    end = nxt[0] if nxt else _end_row(cell, start)
    if end is None:
        return Resolution("NeedsReview", reason="記入欄の終端(発注品)が見つからない")
    filled = [r for r in range(start, end) if _s(cell(r, C_Q))]
    first = (filled[-1] + 1) if filled else start
    if fill_gaps:
        run = next((a for a in range(start, end - n_lines + 1)
                    if all(_free(cell, r) for r in range(a, a + n_lines))), None)
        if run is None:
            return Resolution("NeedsReview", reason=f"追記欄不足（連続{n_lines}行の空きなし）")
        return Resolution("OK", rows=tuple(range(run, run + n_lines)))
    rows = tuple(range(first, first + n_lines))
    if rows[-1] >= end:
        return Resolution("NeedsReview", reason=f"追記欄不足（空き{max(end - first, 0)}行／必要{n_lines}行）")
    for r in rows:
        if any(_s(cell(r, c)) for c in (C_G, C_M, C_Q, C_BS)):
            return Resolution("NeedsReview", reason=f"{r + 1}行目が空きでない（時刻・内容・作業者のいずれかが記入済み）")
    return Resolution("OK", rows=rows)
