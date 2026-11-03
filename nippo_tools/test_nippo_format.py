import unittest
from nippo_format import format_entry, _width, WRAP_WIDTH


class T(unittest.TestCase):
    def test_dot_added(self):
        self.assertEqual(format_entry("「すし昌」天井漏水確認　排水済"), ["・「すし昌」天井漏水確認　排水済"])

    def test_dot_not_duplicated(self):
        self.assertEqual(format_entry("・漏水確認"), ["・漏水確認"])

    def test_halfwidth_kana_dot(self):
        self.assertEqual(format_entry("･ﾏｰﾙﾌﾞﾗﾝｼｭ　ｴｱｺﾝ更新"), ["・マールブランシュ　エアコン更新"])

    def test_sakugyotodoke_no_dot(self):
        self.assertEqual(format_entry("【作業届】"), ["【作業届】"])

    def test_any_heading_no_dot(self):
        self.assertEqual(format_entry("【ホテル】"), ["【ホテル】"])

    def test_arrow_no_dot(self):
        self.assertEqual(format_entry("➡現場異常報告書を不動産へ送信済"), ["➡現場異常報告書を不動産へ送信済"])

    def test_multiline(self):
        self.assertEqual(format_entry("漏水確認\n➡排水済"), ["・漏水確認", "➡排水済"])

    def test_abnormal_unify(self):
        for s in ("異常無し", "異常無", "異常なし"):
            self.assertEqual(format_entry("調査実施するも" + s), ["・調査実施するも異常なし"])

    def test_alpha_halfwidth(self):
        self.assertEqual(format_entry("ＡＬＳＯＫ来館"), ["・ALSOK来館"])

    def test_time_format(self):
        self.assertEqual(format_entry("22:00～24:00　害虫駆除"), ["・22：00～24：00　害虫駆除"])
        self.assertEqual(format_entry("翌4:30"), ["・翌4：30"])

    def test_wrap(self):
        out = format_entry("あ" * 80)
        self.assertGreater(len(out), 1)
        self.assertTrue(out[0].startswith("・"))
        for line in out[1:]:
            self.assertTrue(line.startswith("　"))
            self.assertFalse(line.startswith("・"))
        for line in out:
            self.assertLessEqual(_width(line), WRAP_WIDTH)
        self.assertEqual("".join(l.lstrip("・　") for l in out), "あ" * 80)

    def test_wrap_arrow_continuation_no_dot(self):
        out = format_entry("➡" + "い" * 80)
        self.assertTrue(out[0].startswith("➡"))
        self.assertTrue(all(not l.startswith("・") for l in out))

    def test_empty_lines_dropped(self):
        self.assertEqual(format_entry("\n\n漏水\n\n"), ["・漏水"])


if __name__ == "__main__":
    unittest.main()
