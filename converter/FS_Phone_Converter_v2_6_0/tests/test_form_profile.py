"""2.6.0: form detection and photos on a BLANK form. Everything is synthetic."""
import io
import json
import sys
import unittest
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

from lxml import etree as ET
from PIL import Image

import approval as AP
import converter_mobile as CM
import engine_policy as EP
import form_profile as FP
import template_inspector as TI
from synthetic_template import build_blank_template_parts, build_template_parts, write_zip
from test_approval import ApprovalBase, rezip
from test_homebox import make_package

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'


def plan_of(parts):
    return FP.analyze(parts.keys(), parts.__getitem__)


class TestDetection(unittest.TestCase):
    def test_baseline_report_is_recognised(self):
        plan = plan_of(build_template_parts())
        self.assertEqual(plan['mode'], 'baseline', plan['problems'])
        self.assertEqual([s['number'] for s in plan['slots']], ['P5', 'P24', 'P43'])
        self.assertEqual(plan['slots'][2]['captions'], ['N45', 'N47', 'N50', 'N53', 'N56'])  # 2026-10-17

    def test_blank_forms(self):
        plan = plan_of(build_blank_template_parts('none'))
        self.assertEqual(plan['mode'], 'blank', plan['problems'])
        self.assertIsNone(plan['photo_drawing'])
        self.assertEqual(plan['slots'][2]['captions'], ['N45', 'N48', 'N51', 'N54', 'N57'])  # 2026-10-31
        plan = plan_of(build_blank_template_parts('logo'))
        self.assertEqual((plan['mode'], plan['photo_drawing']), ('blank', 'xl/drawings/drawing2.xml'))

    def test_leftovers_and_other_versions_are_refused(self):
        plan = plan_of(build_blank_template_parts('mark'))
        self.assertEqual(plan['mode'], 'refused')
        self.assertIn('図形', ' '.join(plan['problems']))
        plan = plan_of(build_blank_template_parts('incell'))
        self.assertEqual(plan['mode'], 'refused')
        self.assertIn('セル内画像', ' '.join(plan['problems']))
        parts = build_blank_template_parts('none')
        parts['xl/worksheets/sheet1.xml'] = parts['xl/worksheets/sheet1.xml'].replace(b'G15:BI15', b'G15:BH15')
        plan = plan_of(parts)
        self.assertEqual(plan['mode'], 'refused')
        self.assertIn('枠の位置', plan['problems'][0])
        parts = build_blank_template_parts('none')
        parts['xl/worksheets/sheet2.xml'] = parts['xl/worksheets/sheet2.xml'].replace(b'P3:S3', b'P3:R3')
        self.assertIn('日付の枠', ' '.join(plan_of(parts)['problems']))

    def test_inspector_blocks_and_explains(self):
        import tempfile
        with tempfile.TemporaryDirectory() as td:
            bad = write_zip(Path(td) / 'x.xlsx', build_blank_template_parts('incell'))
            pending = TI.inspect_template(bad)
            self.assertTrue(any('セル内画像' in b for b in pending['inspection']['blockers']))
            good = write_zip(Path(td) / 'y.xlsx', build_blank_template_parts('none'))
            pending = TI.inspect_template(good)
            self.assertEqual(pending['inspection']['blockers'], [])
            self.assertEqual(pending['form']['photo_mode'], 'blank')
            report = TI.render_report(pending)
            self.assertIn('空様式の写真枠に写真を貼る', report)
            self.assertIn('N48', report)
            cats = {p['name']: p['category'] for p in pending['parts']}
            self.assertEqual(cats['xl/drawings/drawing1.xml'], 'immutable')
            self.assertEqual(cats['xl/worksheets/sheet2.xml'], 'modifiable')

    def test_placement_fits_inside_the_slot(self):
        sheet = ET.fromstring(build_blank_template_parts('none')['xl/worksheets/sheet2.xml'])
        col_px, row_emu = FP.sheet_geometry(sheet, 8)
        slot = FP.DEFAULT_SLOTS[0]
        bw, bh = FP.slot_box(slot, col_px, row_emu)
        self.assertEqual((bw // FP.EMU_PER_PX, bh // FP.EMU_PER_PX), (11 * 37, 17 * 18))  # same as Excel
        for w, h in ((1600, 1200), (900, 1600), (1600, 400)):
            col, coff, row, roff, cx, cy = FP.place(slot, col_px, row_emu, w, h)
            self.assertLessEqual(cx, bw)
            self.assertLessEqual(cy, bh)
            self.assertAlmostEqual(cx / cy, w / h, places=2)
            self.assertTrue(1 <= col <= 11 and 4 <= row <= 20)


class TestBlankFormConversion(ApprovalBase):
    def use_blank(self, kind='none'):
        self.template = write_zip(self.dir / f'blank_{kind}.xlsx', build_blank_template_parts(kind))
        self.manifest_path = self.approve(self.template)

    def convert(self, n_photos, name='hb', mutate=None):
        pkg = make_package(self.dir / f'HB-REPORT-{name}.zip', n_photos=n_photos, mutate=mutate)
        out = self.dir / f'{name}.xlsx'
        review = CM.make_report_from_homebox(pkg, self.template, out, self.dir / f'{name}.json',
                                             approval_manifest=self.manifest_path, trusted_keys=self.trusted)
        return review, out.read_bytes()

    def test_photos_go_into_a_blank_form_without_drawing(self):
        self.use_blank('none')
        for n in (0, 1, 3):
            review, out = self.convert(n, name=f'n{n}')
            self.assertEqual(review['photo_count'], n)
            self.assertEqual(review['part_verification']['photo_mode'], 'blank')
            with zipfile.ZipFile(io.BytesIO(out)) as z:
                names = set(z.namelist())
                self.assertEqual({x for x in names if x.startswith('xl/media/hb_photo')},
                                 {f'xl/media/hb_photo{i}.jpeg' for i in range(1, n + 1)})
                if not n:
                    self.assertNotIn(EP.GEN_DRAWING, names)
                    continue
                draw = ET.fromstring(z.read(EP.GEN_DRAWING))
                self.assertEqual(len(draw.findall('.//{%s}pic' % XDR)), n)
                sheet = z.read('xl/worksheets/sheet2.xml').decode('utf-8')
                self.assertIn('<drawing', sheet)
                ct = z.read('[Content_Types].xml').decode('utf-8')
                self.assertIn('Extension="jpeg"', ct)
                self.assertIn('/xl/drawings/hb_photos.xml', ct)
                for i in range(1, n + 1):
                    with Image.open(io.BytesIO(z.read(f'xl/media/hb_photo{i}.jpeg'))) as im:
                        self.assertEqual(im.size, (400, 300))  # own shape kept, not padded
                # caption lines of the 2026-10-31 form: 場所 then 説明
                self.assertIn('<c r="N7" t="inlineStr"><is><t>B1 架空機械室</t>', sheet)
                self.assertIn('架空の写真説明1', sheet)

    def test_existing_drawing_is_kept_and_extended(self):
        self.use_blank('logo')
        review, out = self.convert(2, name='logo')
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            draw = ET.fromstring(z.read('xl/drawings/drawing2.xml'))
            self.assertEqual(len(draw), 3)  # the logo shape + 2 photos
            self.assertEqual(draw[0].find('.//{%s}cNvPr' % XDR).get('name'), '架空ロゴ枠')
            self.assertIn('xl/drawings/_rels/drawing2.xml.rels', z.namelist())
            self.assertNotIn(EP.GEN_DRAWING, z.namelist())

    def test_more_photos_than_slots_stop(self):
        self.use_blank('none')
        with self.assertRaisesRegex(ValueError, '3枚まで|写真枠は3つ'):
            self.convert(4, name='four')

    def test_refused_template_cannot_be_approved(self):
        import approval_sign as SG
        bad = write_zip(self.dir / 'incell.xlsx', build_blank_template_parts('incell'))
        with self.assertRaisesRegex(ValueError, 'セル内画像'):
            SG.sign_pending(bad, TI.inspect_template(bad), self.pem, 'synthetic-passphrase-123', '架空 承認者', True)

    def test_old_manifest_without_form_is_refused(self):
        self.use_blank('none')
        m = json.loads(self.manifest_path.read_text(encoding='utf-8'))
        m['form']['photo_slots'][2]['captions'][1] = 'N47'
        self.manifest_path.write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '署名|配置'):
            AP.verify_manifest(self.template, self.manifest_path, self.trusted)

    def _tampered_output(self, edit):
        """Run the engine, then edit its output and verify again: every edit must be caught."""
        self.use_blank('none')
        approved = AP.verify_manifest(self.template, self.manifest_path, self.trusted)
        plan = approved['_plan']
        photo = io.BytesIO()
        Image.new('RGB', (400, 300), 'teal').save(photo, 'JPEG')
        img, _, _ = CM.photo_to_jpeg(photo.getvalue())
        case = {'件名': '架空 件名', '発生日時': '2026-11-07T10:30:00+09:00'}
        r, p, _ = CM.build_updates_with_notes(dict(case, _contacts=[]), [{'写真右側の説明文': '架空'}], [], slots=plan['slots'])
        out, _ = CM.transform_package(self.template, r, p, {'hb_photo1.jpeg': img}, approval=approved, plan=plan)
        changed = edit(out)
        self._last_changed = changed != out
        self._last_out = out
        return lambda: CM.verify_output_against_approval(
            self.template, changed, approved, set(CM.CLEAR_REPORT) | set(r), CM.clear_photo_cells(plan['slots']) | set(p),
            [], plan=plan, photo_count=1)

    def test_tampering_with_the_placed_pictures_is_caught(self):
        def read(out, name):
            with zipfile.ZipFile(io.BytesIO(out)) as z:
                return z.read(name)
        cases = {
            'link on picture': lambda o: rezip(o, {EP.GEN_DRAWING: read(o, EP.GEN_DRAWING).replace(
                b'<xdr:cNvPr id="2" name="&#20889;&#30495; 1"/>', b'<xdr:cNvPr id="2" name="&#20889;&#30495; 1" descr="x"/>').replace(
                b'name="\xe5\x86\x99\xe7\x9c\x9f 1"/>', b'name="\xe5\x86\x99\xe7\x9c\x9f 1" descr="x"/>')}),
            'external image': lambda o: rezip(o, {EP.GEN_DRAWING_RELS: read(o, EP.GEN_DRAWING_RELS).replace(
                b'Target="../media/hb_photo1.jpeg"', b'Target="https://example.invalid/x.jpeg" TargetMode="External"')}),
            'extra part': lambda o: rezip(o, {'xl/media/hb_photo9.jpeg': read(o, 'xl/media/hb_photo1.jpeg')}),
            'picture moved': lambda o: rezip(o, {EP.GEN_DRAWING: read(o, EP.GEN_DRAWING).replace(
                b'<xdr:row>5</xdr:row>', b'<xdr:row>30</xdr:row>', 1)}),
        }
        for label, edit in cases.items():
            with self.subTest(label):
                check = self._tampered_output(edit)
                self.assertTrue(self._last_changed, 'the edit must really change the output')
                with self.assertRaises(ValueError):
                    check()

    def test_untampered_output_passes(self):
        check = self._tampered_output(lambda o: o)
        self.assertEqual(check()['photo_mode'], 'blank')


if __name__ == '__main__':
    unittest.main()
