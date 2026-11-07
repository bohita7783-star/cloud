"""2.5.0 report layout: written like completed company reports. All data is fictional."""
import datetime
import sys
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import converter_mobile as CM
import report_layout as RL

TODAY = datetime.date(2026, 11, 8)


def case(**kw):
    c = {'ID': 'CASE-1', '件名': '架空区画 天井漏水あり', '発生日時': '2026-11-07T10:30:00+09:00',
         '発生場所': '架空ホール', '状況概要': '架空店の店長より、客席の天井から漏水していると連絡有り。',
         '発生区分': '故障', '職種': '設備', '案件類型': '漏水', '現場確認を実施': True, '一次対応を実施': True,
         '現場確認開始': '10:35', '現場確認終了': '10:40'}
    c.update(kw)
    return c


def lay(c, timeline=(), contacts=()):
    c = dict(c, _contacts=list(contacts))
    return CM.build_updates_with_notes(c, [], list(timeline), today=TODAY)[0]


class TestHeader(unittest.TestCase):
    def test_dropdown_strings_and_dates(self):
        r = lay(case())
        self.assertEqual(r['G3'], '故　障　')            # exact list string of the form
        self.assertEqual(r['G4'], '設　備')
        self.assertEqual(r['W3'], '漏水')
        self.assertEqual(r['AB2'], '（第１報）')
        self.assertEqual(r['AU1'], RL.excel_serial_day(TODAY))           # 発行日 = report day
        self.assertEqual(r['G5'], RL.excel_serial_day(datetime.date(2026, 11, 7)))
        self.assertEqual((r['X5'], r['AC5']), (10, 30))
        self.assertEqual((r['G8'], r['G9']), ('有', '有'))
        self.assertAlmostEqual(r['L8'], (10 * 60 + 35) / 1440)

    def test_no_test_markers_and_no_page2_overwrite(self):
        r = lay(case())
        self.assertEqual(r['G6'], '架空区画 天井漏水あり')
        for v in r.values():
            self.assertNotIn('検証', str(v))
            self.assertNotIn('Field Survey', str(v))
        for cell in ('G53', 'X53', 'AC53', 'G54', 'AP54'):    # page 2 = formulas of page 1
            self.assertNotIn(cell, r)

    def test_profile_and_people(self):
        r = lay(case(**{'報告の段階': '（第２報）', '部署': '架空支店', 'ビル名': '架空ビル', '報告者': '架空 太郎',
                        '報告者の役職': '班長', '当事者': '架空、仮名', '当事者の区分': '【当社の社員】',
                        '現場確認の実施者': '架空', '現場確認の内容': '現場確認', '一次対応の実施者': '架空、仮名',
                        '一次対応の内容': 'バケツで受ける', '処置の実施者': '架空', '処置の確認者': '架空課長'}))
        self.assertEqual((r['AB2'], r['AP2'], r['AP3'], r['AP4'], r['BC4']), ('（第２報）', '架空支店', '架空ビル', '架空 太郎', '班長'))
        self.assertEqual((r['G7'], r['S7'], r['AC9'], r['AP9'], r['AR32'], r['AR34']),
                         ('架空、仮名', '【当社の社員】', '架空、仮名', 'バケツで受ける', '架空', '架空課長'))

    def test_unknown_stage_falls_back_to_first(self):
        self.assertEqual(lay(case(**{'報告の段階': '第9報'}))['AB2'], '（第１報）')


class TestChronology(unittest.TestCase):
    def test_one_line_per_row_with_time_on_first_line_only(self):
        long = 'あ' * 60 + '。'
        r = lay(case(状況概要=long), [{'対応日時': '2026-11-07T10:50:00+09:00', '実施内容・確認結果': '天井内に受けを設置する。'}])
        self.assertEqual(r['C15'], '10：30')
        self.assertEqual(r['G15'], 'あ' * 51)
        self.assertNotIn('C16', r)                             # continuation line: no time
        self.assertEqual(r['G16'], 'あ' * 9 + '。')
        self.assertEqual((r['C17'], r['G17']), ('10：50', '天井内に受けを設置する。'))

    def test_line_never_starts_with_punctuation(self):
        lines = RL.wrap('あ' * 51 + '。次', 51)
        self.assertEqual(lines, ['あ' * 51 + '。', '次'])

    def test_other_day_written_as_date_with_time_in_text(self):
        r = lay(case(), [{'対応日時': '2026-11-08T09:20:00+09:00', '実施内容・確認結果': '再調査する。'}])
        self.assertEqual((r['C16'], r['G16']), ('11/8', '9:20　再調査する。'))

    def test_measurements_person_and_contact(self):
        r = lay(case(状況概要='発見。'), [{'対応日時': '2026-11-07T11:00:00+09:00', '実施内容・確認結果': '排水する。',
                                       '測定数値など': '約1L', '担当': '架空', '連絡先・連絡結果': '店長へ説明'}])
        self.assertEqual(r['G16'], '排水する。（約1L）（担当：架空）')
        self.assertEqual(r['G17'], '連絡：店長へ説明')

    def test_plan_goes_to_the_bottom_rows(self):
        r = lay(case(**{'恒久対策・今後の予定': '修理を提案予定。'}))
        self.assertEqual(r['G26'], '修理を提案予定。')
        self.assertNotIn('G27', r)
        r = lay(case(**{'恒久対策・今後の予定': 'い' * 60}))
        self.assertEqual((r['G26'], r['G27']), ('い' * 51, 'い' * 9))

    def test_overflow_stops_instead_of_dropping(self):
        many = [{'対応日時': f'2026-11-07T{11 + i}:00:00+09:00', '実施内容・確認結果': 'う' * 40} for i in range(8)]
        r = lay(case(), many[:8])  # 1 + 8 = 9 rows: fits
        self.assertIn('G23', r)
        with self.assertRaisesRegex(ValueError, '行になり、様式の.*入りません'):
            lay(case(状況概要='え' * 400), many)  # 8 + 8 = 16 rows > 13
        with self.assertRaisesRegex(ValueError, '恒久対策・今後の予定'):
            lay(case(**{'恒久対策・今後の予定': 'お' * 120}))

    def test_cause_and_action_rows(self):
        r = lay(case(**{'原因・考察': '経年使用によるものと思われます。', '応急処置': 'か' * 30, '処置報告の区分': '恒久処置：'}))
        self.assertEqual(r['C28'], '経年使用によるものと思われます。')
        self.assertEqual((r['G32'], r['N32'], r['G33']), ('恒久処置：', 'か' * 24, 'か' * 6))
        with self.assertRaisesRegex(ValueError, '原因・考察'):
            lay(case(**{'原因・考察': 'き' * 300}))
        with self.assertRaisesRegex(ValueError, '処置報告の内容'):
            lay(case(**{'応急処置': 'く' * 200}))
        self.assertEqual(lay(case(**{'応急処置': '受けを設置'}))['G32'], '応急処置：')


class TestContacts(unittest.TestCase):
    def test_slots_like_completed_reports(self):
        r = lay(case(), contacts=[
            {'区分': 'オーナー・テナント', '手段': 'メール', '相手': '架空', '所属': '架空不動産株式会社'},
            {'区分': 'オーナー・テナント', '手段': 'メール', '相手': '仮名 様', '所属': '架空不動産株式会社'},
            {'区分': '社内', '手段': 'メール', '相手': '架空課長'},
            {'区分': '対応業者', '手段': '電話', '相手': '架空設備', '所属': '架空工業'}])
        self.assertEqual((r['G11'], r['K11'], r['T11']), ('【ﾒｰﾙ】', '架空　様', '架空不動産株式会社'))
        self.assertEqual((r['G12'], r['K12']), ('【ﾒｰﾙ】', '仮名 様'))   # left column first
        self.assertEqual((r['AH11'], r['AH12'], r['G10']), ('：', '：', '：'))  # unused: blank-form placeholder
        self.assertEqual((r['G14'], r['K14'], r['Q14']), ('【ﾒｰﾙ】', '架空課長', 'ﾒｰﾙ'))
        self.assertEqual((r['G13'], r['K13']), ('【電話】', '架空設備（架空工業）'))

    def test_too_many_contacts_stop(self):
        five = [{'区分': 'オーナー・テナント', '手段': 'メール', '相手': f'架空{i}'} for i in range(5)]
        with self.assertRaisesRegex(ValueError, '4件まで'):
            lay(case(), contacts=five)
        with self.assertRaisesRegex(ValueError, '区分が違います'):
            lay(case(), contacts=[{'区分': '知らない区分', '相手': 'x'}])


class TestPhotosAndWarnings(unittest.TestCase):
    def test_photo_captions_three_lines(self):
        photos = [{'撮影場所': '架空区画 店内', '写真右側の説明文': 'ホール天井\n天井内にバケツ設置'},
                  {'撮影場所': '架空区画', '写真右側の説明文': '配線', '注釈・マーキングの指示': '赤枠が焼損箇所'}]
        _, p, _ = CM.build_updates_with_notes(case(_contacts=[]), photos, [], today=TODAY)
        self.assertEqual((p['B1'], p['P5'], p['N7'], p['N10'], p['N13']), ('架空区画 天井漏水あり', 1, '架空区画 店内', 'ホール天井', '天井内にバケツ設置'))
        self.assertEqual((p['N29'], p['N32']), ('配線', '赤枠が焼損箇所'))
        # 2.6.0: 5 caption lines in the 2026-10 form (場所 + 説明 + 注釈); the 3rd slot uses the
        # template's own lines (N45/N48/... in the 2026-10-31 form)
        five = {'撮影場所': '場所', '写真右側の説明文': 'a\nb\nc', '注釈・マーキングの指示': '赤丸'}
        _, p, _ = CM.build_updates_with_notes(case(_contacts=[]), [{}, {}, five], [], today=TODAY)
        self.assertEqual([p.get(c) for c in ('N45', 'N48', 'N51', 'N54', 'N57')], ['場所', 'a', 'b', 'c', '赤丸'])
        with self.assertRaisesRegex(ValueError, '5行まで'):
            CM.build_updates_with_notes(case(_contacts=[]), [dict(five, 写真右側の説明文='a\nb\nc\nd')], [], today=TODAY)
        with self.assertRaisesRegex(ValueError, '3枚まで'):
            CM.build_updates_with_notes(case(_contacts=[]), [{}, {}, {}, {}], [], today=TODAY)

    def test_long_title_is_a_warning_not_a_loss(self):
        _, _, warn = CM.build_updates_with_notes(case(件名='け' * 40, _contacts=[]), [], [], today=TODAY)
        self.assertTrue(any('件名' in w for w in warn))

    def test_every_written_cell_is_cleared_or_permitted(self):
        r = lay(case(**{'部署': 'x'}), contacts=[{'区分': '関係機関', '手段': '電話', '相手': '消防署'}])
        self.assertTrue(set(r) <= RL.REPORT_CELLS)


if __name__ == '__main__':
    unittest.main()
