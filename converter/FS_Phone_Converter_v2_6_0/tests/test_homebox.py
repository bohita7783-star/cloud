"""HomeBox package input (2.4.0). Synthetic data only."""
import hashlib
import io
import json
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE.parent), str(HERE)]

from PIL import Image
import converter_mobile as CM
import homebox_adapter as HB
import homebox_convert as HC
from test_approval import ApprovalBase


def jpeg(color='navy', size=(400, 300)):
    b = io.BytesIO()
    Image.new('RGB', size, color).save(b, 'JPEG')
    return b.getvalue()


def make_package(path, n_photos=2, mutate=None):
    photos = []
    files = {}
    for i in range(1, n_photos + 1):
        name = f'CASE-20261107-0001_EQ-000006_{i:03d}_銘板.jpg'
        data = jpeg(('navy', 'teal', 'maroon', 'olive')[i % 4])
        files['photos/' + name] = data
        photos.append({'ID': f'PH-{i:04d}', '写真番号': i, '撮影日時': f'2026-11-07T10:3{i}:00+09:00',
                       '撮影場所': 'B1 架空機械室', '元写真': 'photos/' + name, '写真右側の説明文': f'架空の写真説明{i}',
                       '注釈・マーキングの指示': None, '種別': '銘板', 'sha256': hashlib.sha256(data).hexdigest()})
    pkg = {
        'schema': 'hb-report-package/1', 'created_at': '2026-11-07T11:00:00+09:00',
        'app': {'name': 'HomeBox', 'version': '0.1.0'},
        'case': {'ID': 'CASE-20261107-0001', '案件識別番号': 'R-0001', '設備ID': 'EQ-000006',
                 '件名': '架空 熱源水ポンプ 異音', '発生日時': '2026-11-07T10:30:00+09:00', '発生区分': '故障',
                 '職種': '設備', '発生場所': 'B1 架空機械室', '状況概要': '架空の状況概要。軸受付近から異音。',
                 '現場確認を実施': True, '現場確認開始': '10:35', '現場確認終了': '10:50',
                 '一次対応を実施': False, '復旧・完了状況': '調査中'},
        'equipment': {'id': 'EQ-000006', 'name': '架空 熱源水ポンプ', 'location': 'B1 架空機械室'},
        'photos': photos,
        'timeline': [{'ID': 'TL-0001', '対応日時': '2026-11-07T10:40:00+09:00', '担当': '架空 担当',
                      '実施内容・確認結果': '架空：運転電流を測定', '測定数値など': '15.2 A'}],
    }
    if mutate:
        mutate(pkg, files)
    with zipfile.ZipFile(path, 'w') as z:
        z.writestr('package.json', json.dumps(pkg, ensure_ascii=False))
        for n, d in files.items():
            z.writestr(n, d)
    return path


class TestHomeBoxPackage(ApprovalBase):
    def convert(self, package, name='hb'):
        out = self.dir / f'{name}.xlsx'
        return CM.make_report_from_homebox(package, self.template, out, self.dir / f'{name}.json',
                                           approval_manifest=self.manifest_path, trusted_keys=self.trusted), out

    def test_package_converts_with_approved_template(self):
        pkg = make_package(self.dir / 'HB-REPORT-CASE-20261107-0001.zip')
        review, out = self.convert(pkg)
        self.assertEqual(review['input']['kind'], 'homebox-package')
        self.assertEqual(review['input']['equipment_id'], 'EQ-000006')
        self.assertEqual(review['photo_count'], 2)
        self.assertEqual(review['timeline_count'], 1)
        self.assertIn('案件識別番号', review['unmapped_nonempty_fields'])
        self.assertTrue(review['approval']['approval_id'].startswith('APV-'))
        with zipfile.ZipFile(out) as z:
            sheet = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            self.assertIn('架空 熱源水ポンプ 異音', sheet)
            self.assertIn('xl/media/image2.jpeg', z.namelist())
            self.assertNotIn('xl/media/image3.jpeg', z.namelist())

    def test_cli_converter_writes_report_and_log(self):
        pkg = make_package(self.dir / 'HB-REPORT-CASE-20261107-0001.zip')
        out_dir = self.dir / 'FS_変換結果'
        output, audit, review = HC.convert(pkg, self.template, out_dir,
                                           approval_manifest=self.manifest_path, trusted_keys=self.trusted)
        self.assertTrue(output.is_file() and audit.is_file())
        self.assertIn('架空 熱源水ポンプ 異音', output.name)

    def test_package_and_template_untouched(self):
        pkg = make_package(self.dir / 'HB-REPORT-x.zip')
        before = (pkg.read_bytes(), self.template.read_bytes())
        self.convert(pkg)
        self.assertEqual(before, (pkg.read_bytes(), self.template.read_bytes()))

    def test_rejections(self):
        cases = {
            'photo changed after recording': (lambda p, f: f.__setitem__(next(iter(f)), jpeg('white')), '記録した時と違います'),
            'missing photo file': (lambda p, f: f.pop(next(iter(f))), 'ありません'),
            'path traversal': (lambda p, f: p['photos'][0].__setitem__('元写真', 'photos/../x.jpg'), 'ファイル名が不正'),
            'unknown schema': (lambda p, f: p.__setitem__('schema', 'hb-report-package/9'), '対応していない形式'),
            'unknown case key': (lambda p, f: p['case'].__setitem__('秘密の項目', 'x'), '知らない項目'),
            'equipment mismatch': (lambda p, f: p['equipment'].__setitem__('id', 'EQ-999999'), '設備ID'),
            'extra file': (lambda p, f: f.__setitem__('notes.txt', b'x'), '想定外のファイル'),
            'unused photo': (lambda p, f: f.__setitem__('photos/extra.jpg', jpeg()), '使われていない写真'),
            'nested object': (lambda p, f: p['case'].__setitem__('件名', {'a': 1}), '型が違います'),
        }
        for label, (mutate, msg) in cases.items():
            with self.subTest(label=label):
                pkg = make_package(self.dir / f'bad_{abs(hash(label))}.zip', mutate=mutate)
                with self.assertRaisesRegex(ValueError, msg):
                    HB.read_package(pkg)

    def test_audit_negative_cases(self):
        # The four cases from the v2.4.0 audit, plus neighbours.
        def wrong_eq(p, f):
            before = p['photos'][0]['元写真']
            after = before.replace('EQ-000006', 'EQ-999999')
            p['photos'][0]['元写真'] = after
            f[after] = f.pop(before)
        def bad_name(p, f):
            before = p['photos'][0]['元写真']
            after = 'photos/random_name.jpg'
            p['photos'][0]['元写真'] = after
            f[after] = f.pop(before)
        def no_case_eq(p, f):
            p['case']['設備ID'] = None
            p['equipment'] = {}
        cases = {
            'photo without sha256': (lambda p, f: p['photos'][0].pop('sha256'), 'SHA-256'),
            'short sha256': (lambda p, f: p['photos'][0].__setitem__('sha256', 'abc'), 'SHA-256'),
            'photo of another equipment': (wrong_eq, '別の設備'),
            'photo name breaks rule': (bad_name, '命名規則'),
            'unknown top-level key': (lambda p, f: p.__setitem__('unexpected_field', 'x'), '知らない項目'),
            'unknown equipment key': (lambda p, f: p['equipment'].__setitem__('secret', 'x'), '知らない項目'),
            'unknown app key': (lambda p, f: p['app'].__setitem__('build', 'x'), '知らない項目'),
            'equipment only on one side': (lambda p, f: p['case'].__setitem__('設備ID', None), '一致しません'),
            'photos on a case without equipment': (no_case_eq, '設備が決まっていない'),
            'photos is not a list': (lambda p, f: p.__setitem__('photos', {'a': 1}), '形式が違います'),
        }
        for label, (mutate, msg) in cases.items():
            with self.subTest(label=label):
                pkg = make_package(self.dir / f'neg_{abs(hash(label))}.zip', mutate=mutate)
                with self.assertRaisesRegex(ValueError, msg):
                    HB.read_package(pkg)

    def test_capture_time_evidence_is_accepted(self):
        def add(p, f):
            p['photos'][0]['撮影日時の根拠'] = 'EXIF'
            p['photos'][0]['登録日時'] = '2026-11-07T10:45:00+09:00'
        pkg = make_package(self.dir / 'HB-REPORT-evidence.zip', mutate=add)
        case, photos, timeline, info = HB.read_package(pkg)
        self.assertNotIn('撮影日時の根拠', photos[0])

    def test_more_than_three_photos_stops_without_dropping(self):
        pkg = make_package(self.dir / 'HB-REPORT-4.zip', n_photos=4)
        with self.assertRaisesRegex(ValueError, '3枚まで'):
            self.convert(pkg)

    def test_package_needs_approval_like_everything_else(self):
        pkg = make_package(self.dir / 'HB-REPORT-y.zip')
        self.manifest_path.unlink()
        with self.assertRaisesRegex(ValueError, '承認済み一覧がありません'):
            CM.make_report_from_homebox(pkg, self.template, self.dir / 'z.xlsx', None,
                                        trusted_keys=self.trusted, approvals_dir=self.approvals)


if __name__ == '__main__':
    import unittest
    unittest.main()
