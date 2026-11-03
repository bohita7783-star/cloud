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


class SpecExamples(unittest.TestCase):
    """別AI作成「入力項目仕様書 v0.1」§2 の文章例7件と、追加の誤変換防止ケース。"""

    def test_spec_examples(self):
        cases = [
            ("区画25の天井漏水を確認", "・区画25の天井漏水を確認"),
            ("・異常無", "・異常なし"),
            ("【ホテル】", "【ホテル】"),
            ("【作業届】", "【作業届】"),
            ("➡翌朝再確認する", "➡翌朝再確認する"),
            ("･警報なし", "・警報なし"),
            ("ﾌﾞﾚｰｶｰを確認", "・ブレーカーを確認"),
        ]
        for src, exp in cases:
            with self.subTest(src=src):
                self.assertEqual(format_entry(src), [exp])

    def test_abnormal_not_overreplaced(self):
        # 「異常無く」「異常無かった」「異常無線」は別の語。「異常なしく」等にしてはいけない
        for s in ("異常無く復旧", "異常無かった", "異常無線"):
            with self.subTest(s=s):
                self.assertEqual(format_entry(s), ["・" + s])

    def test_abnormal_at_punctuation(self):
        self.assertEqual(format_entry("調査実施するも異常無。"), ["・調査実施するも異常なし。"])
        self.assertEqual(format_entry("異常無（確認済）"), ["・異常なし（確認済）"])

    def test_no_wrap_mode(self):
        out = format_entry("あ" * 80, wrap_width=None)
        self.assertEqual(out, ["・" + "あ" * 80])


    def test_fullwidth_time_digits(self):
        self.assertEqual(format_entry("２２：００～２３：３０"), ["・22：00～23：30"])


if __name__ == "__main__":
    unittest.main()
