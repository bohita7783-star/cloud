import unittest
from nippo_target import resolve, check_date

def grid(sections=None, texts=None, year=2026, month=11, day=3, extra=None, end=52):
    """合成シート: sections={見出し: 行}, texts={行: Q列の文}"""
    g = {(3, 0): year, (3, 10): month, (3, 18): day, (end, 0): "発注品"}
    for lab, r in (sections or {"不動産関係": 24, "テナント関係": 30, "工事関係": 36, "特記事項": 42}).items():
        g[(r, 3)] = lab
    for r, t in (texts or {}).items():
        g[(r, 16)] = t
    g.update(extra or {})
    return lambda r, c: g.get((r, c), "")

KW = dict(sheet_name="(3)", year=2026, month=11, day=3)

class T(unittest.TestCase):
    def test_empty_section_uses_first_row(self):
        r = resolve(grid(), "テナント関係", 1, **KW)
        self.assertEqual((r.status, r.rows), ("OK", (30,)))

    def test_appends_after_last_filled(self):
        r = resolve(grid(texts={30: "a", 31: "b"}), "テナント関係", 1, **KW)
        self.assertEqual(r.rows, (32,))

    def test_gap_is_not_reused(self):
        # 31行目が空でも、最後の記入(33行目)の次に書く（時系列を崩さない）
        r = resolve(grid(texts={30: "a", 32: "c"}), "テナント関係", 1, **KW)
        self.assertEqual(r.rows, (33,))

    def test_shifted_layout_is_found(self):
        secs = {"不動産関係": 25, "テナント関係": 31, "工事関係": 37, "特記事項": 43}
        r = resolve(grid(sections=secs), "テナント関係", 1, **KW)
        self.assertEqual(r.rows, (31,))

    def test_full_section_needs_review(self):
        t = {r: "x" for r in range(30, 36)}
        r = resolve(grid(texts=t), "テナント関係", 1, **KW)
        self.assertEqual(r.status, "NeedsReview"); self.assertIn("不足", r.reason)

    def test_multiline_needs_consecutive_rows(self):
        t = {r: "x" for r in range(30, 34)}      # 空きは34,35の2行
        self.assertEqual(resolve(grid(texts=t), "テナント関係", 2, **KW).rows, (34, 35))
        self.assertEqual(resolve(grid(texts=t), "テナント関係", 3, **KW).status, "NeedsReview")

    def test_last_section_ends_at_hatchuhin(self):
        r = resolve(grid(texts={r: "x" for r in range(42, 52)}), "特記事項", 1, **KW)
        self.assertEqual(r.status, "NeedsReview")

    def test_row_with_time_or_worker_is_not_free(self):
        r = resolve(grid(extra={(30, 6): 0.5}), "テナント関係", 1, **KW)   # G列に時刻
        self.assertEqual(r.status, "NeedsReview")
        r = resolve(grid(extra={(30, 70): "全員"}), "テナント関係", 1, **KW)  # BS列に作業者
        self.assertEqual(r.status, "NeedsReview")

    def test_date_mismatch(self):
        for kw, word in ((dict(sheet_name="(4)"), "シート名"), (dict(day=4, sheet_name="(4)"), "日付"),
                         (dict(year=2027), "日付"), (dict(month=10), "日付")):
            with self.subTest(kw=kw):
                a = {**KW, **kw}
                r = resolve(grid(), "テナント関係", 1, **a)
                self.assertEqual(r.status, "NeedsReview"); self.assertIn(word, r.reason)

    def test_year_is_checked(self):
        self.assertIsNone(check_date(grid(), "(3)", 2026, 11, 3))
        self.assertIsNotNone(check_date(grid(year=2025), "(3)", 2026, 11, 3))

    def test_missing_or_duplicate_section(self):
        s = {"不動産関係": 24, "テナント関係": 30, "工事関係": 36}
        self.assertEqual(resolve(grid(sections=s), "テナント関係", 1, **KW).status, "NeedsReview")
        g = grid(extra={(45, 3): "テナント関係"})
        self.assertIn("重複", resolve(g, "テナント関係", 1, **KW).reason)

    def test_fill_gaps_uses_middle_row(self):
        g = grid(texts={30: "a", 32: "c", 35: "z"})        # 末尾35行目まで使用済みだが途中に空き
        self.assertEqual(resolve(g, "テナント関係", 1, **KW).status, "NeedsReview")
        r = resolve(g, "テナント関係", 1, fill_gaps=True, **KW)
        self.assertEqual((r.status, r.rows), ("OK", (31,)))

    def test_fill_gaps_needs_consecutive(self):
        g = grid(texts={30: "a", 32: "c", 34: "e", 35: "z"})  # 空きは31,33の単独のみ
        self.assertEqual(resolve(g, "テナント関係", 2, fill_gaps=True, **KW).status, "NeedsReview")


if __name__ == "__main__":
    unittest.main()
