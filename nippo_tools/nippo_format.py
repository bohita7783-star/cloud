"""勤務日報の本文整形（基準実装）。

スマホ入力アプリ／転記処理に移植するときの「正解」を定義する。
ルール：
  1. 行頭は必ず「・」。ただし【…】の見出し行（【作業届】等すべて）と「➡」始まりの続き行は付けない。
  2. 半角カナ → 全角、「･」→「・」、全角英字 → 半角英字。
  3. 時刻は「22：00」（全角コロン）、範囲は「～」。
  4. 「異常無し」「異常無」→「異常なし」に統一。
  5. 表示幅 WRAP_WIDTH（半角換算・既定80）で折り返し、2行目以降は全角スペース1つで字下げ（「・」は付けない）。
"""
import re
import unicodedata

WRAP_WIDTH = 80  # 推定値：内容欄は9pt・縮小表示ON。縮まず読める目安。実機で要確認
INDENT = "　"
NO_DOT_STARTS = ("【", "➡")


def _width(s: str) -> int:
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in s)


def _normalize(s: str) -> str:
    # 半角カナは濁点・半濁点が別文字なので、連続した塊ごとにまとめて変換する
    s = re.sub("[\uff66-\uff9f]+", lambda m: unicodedata.normalize("NFKC", m.group()), s)
    s = re.sub("[Ａ-Ｚａ-ｚ]+", lambda m: unicodedata.normalize("NFKC", m.group()), s)
    s = s.replace("･", "・").replace("〜", "～").replace("~", "～")
    # 時刻：22:00 / ２２：００ → 22：00
    s = re.sub(r"(?<!\d)(\d{1,2})[:：](\d{2})(?!\d)", r"\1：\2", s)
    s = re.sub(r"異常無し|異常無(?!し)", "異常なし", s)
    return s


def _wrap(first_prefix: str, text: str) -> list:
    lines, cur = [], first_prefix
    for ch in text:
        if _width(cur + ch) > WRAP_WIDTH:
            lines.append(cur)
            cur = INDENT
        cur += ch
    lines.append(cur)
    return lines


def format_entry(text: str) -> list:
    """利用者が入力した1項目（改行可）を、日報の行リストへ整形する。"""
    result = []
    for raw in text.splitlines():
        raw = _normalize(raw).strip()
        if not raw:
            continue
        if raw.startswith(NO_DOT_STARTS):
            result.extend(_wrap("", raw))
        else:
            body = raw.lstrip("・ 　")
            result.extend(_wrap("・", body))
    return result
