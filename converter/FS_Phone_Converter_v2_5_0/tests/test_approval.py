"""v2.3.3 approval tests. Every workbook, photo, key and name is synthetic."""
import base64
import datetime
import io
import json
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

from lxml import etree as ET
from PIL import Image

import approval as AP
import approval_sign as SG
import converter_mobile as CM
import engine_policy as EP
import template_inspector as TI
from synthetic_template import (build_template_parts, write_zip, LEGACY, LEGACY_PHOTO, LEGACY_CREATOR, EXT_PATH)
from test_converter import make_workbook

PASS = 'synthetic-passphrase-123'
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'


def rezip(data, changes=None, remove=()):
    changes = changes or {}
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(data)) as zin, zipfile.ZipFile(out, 'w') as zout:
        for info in zin.infolist():
            if info.filename in remove:
                continue
            zout.writestr(info, changes.pop(info.filename, zin.read(info.filename)))
        for name, value in changes.items():
            zout.writestr(name, value)
    return out.getvalue()


class ApprovalBase(unittest.TestCase):
    def setUp(self):
        self.td = tempfile.TemporaryDirectory()
        self.addCleanup(self.td.cleanup)
        self.dir = Path(self.td.name)
        self.template = write_zip(self.dir / 'synthetic_blank_template.xlsx', build_template_parts())
        self.pem, self.entry = SG.generate_keypair(PASS, '架空 承認者')
        self.trusted = self.write_trusted([self.entry])
        self.approvals = self.dir / 'approvals'
        self.approvals.mkdir()
        self.manifest_path = self.approve(self.template)

    def write_trusted(self, entries, name='trusted.json'):
        p = self.dir / name
        p.write_text(json.dumps({'schema': AP.TRUSTED_SCHEMA, 'keys': entries}, ensure_ascii=False), encoding='utf-8')
        return p

    def approve(self, template, pem=None, valid_days=None, now=None, folder=None):
        pending = TI.inspect_template(template, now=now)
        manifest = SG.sign_pending(template, pending, pem or self.pem, PASS, '架空 承認者', True, valid_days, now=now)
        folder = folder or self.approvals
        path = folder / (manifest['template']['sha256'] + '.approved.json')
        path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
        return path

    def field_inputs(self):
        field = make_workbook(self.dir / 'synthetic_field_survey.xlsx')
        photos = self.dir / 'synthetic_photos.zip'
        b = io.BytesIO()
        Image.new('RGB', (400, 300), 'navy').save(b, 'JPEG')
        with zipfile.ZipFile(photos, 'w') as z:
            z.writestr('photoB.jpg', b.getvalue())
        return field, photos

    def run_report(self, template=None, manifest=None, name='out'):
        field, photos = self.field_inputs()
        out = self.dir / f'{name}.xlsx'
        audit = self.dir / f'{name}.json'
        review = CM.make_report(str(field), str(photos), str(template or self.template), str(out), str(audit),
                                case_id='ID_B', approval_manifest=manifest, trusted_keys=self.trusted,
                                approvals_dir=self.approvals)
        return review, out.read_bytes()

    def good_output(self):
        approved = AP.verify_manifest(self.template, self.manifest_path, self.trusted)
        ru = {'G6': '【検証用】架空の新しい件名', 'AP6': '架空室'}
        pu = {'N10': '架空の写真説明'}
        img = io.BytesIO()
        Image.new('RGB', (120, 90), 'red').save(img, 'JPEG')
        out, log = CM.transform_package(self.template, ru, pu, {'image1.jpeg': img.getvalue()}, approval=approved)
        return approved, out, ru, pu

    def verify(self, approved, out, ru, pu):
        return CM.verify_output_against_approval(
            self.template, out, approved, set(CM.CLEAR_REPORT) | set(ru), set(CM.CLEAR_PHOTO) | set(pu),
            [str(v) for v in list(ru.values()) + list(pu.values())])


class TestInspectorAndSigning(ApprovalBase):
    def test_inspector_classifies_parts_and_hides_values(self):
        pending = TI.inspect_template(self.template)
        self.assertEqual(pending['status'], 'pending')
        cats = {p['name']: p['category'] for p in pending['parts']}
        self.assertEqual(cats['xl/worksheets/sheet1.xml'], 'modifiable')
        self.assertEqual(cats['xl/media/image4.emf'], 'immutable')
        self.assertEqual(cats['xl/ctrlProps/ctrlProp1.xml'], 'immutable')
        self.assertEqual(cats['xl/calcChain.xml'], 'delete')
        report = TI.render_report(pending)
        for secret in list(LEGACY.values()) + list(LEGACY_PHOTO.values()) + [LEGACY_CREATOR, EXT_PATH, '架空設備マスタ']:
            self.assertNotIn(secret, report)
            self.assertNotIn(secret, json.dumps(pending, ensure_ascii=False))
        self.assertIn('外部への参照', report)
        self.assertIn('補助リスト', report)

    def test_pending_manifest_is_not_an_approval(self):
        pending = TI.inspect_template(self.template)
        p = self.dir / 'pending.json'
        p.write_text(json.dumps(pending, ensure_ascii=False), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '未承認'):
            AP.verify_manifest(self.template, p, self.trusted)

    def test_signing_requires_review_confirmation(self):
        pending = TI.inspect_template(self.template)
        with self.assertRaisesRegex(ValueError, 'Excelでの確認'):
            SG.sign_pending(self.template, pending, self.pem, PASS, '架空 承認者', False)

    def test_signing_refuses_if_template_changed_after_inspection(self):
        pending = TI.inspect_template(self.template)
        parts = build_template_parts()
        parts['xl/worksheets/sheet3.xml'] += b' '
        write_zip(self.template, parts)
        with self.assertRaisesRegex(ValueError, '検査した後'):
            SG.sign_pending(self.template, pending, self.pem, PASS, '架空 承認者', True)


class TestApprovedConversion(ApprovalBase):
    def test_approved_template_passes_end_to_end(self):
        review, out = self.run_report()
        self.assertEqual(review['version'], CM.VERSION)
        self.assertTrue(review['approval']['approval_id'].startswith('APV-'))
        self.assertEqual(review['approval_manifest']['template_sha256'], AP.sha256_file(self.template))
        pv = review['part_verification']
        self.assertGreater(pv['immutable_parts_byte_identical'], 5)
        self.assertIn('xl/calcChain.xml', pv['deleted_parts'])
        self.assertEqual(pv['added_parts'], [])
        self.assertIn('converter_mobile.py', review['program']['file_sha256'])
        audit = json.loads((self.dir / 'out.json').read_text(encoding='utf-8'))
        self.assertEqual(audit['approval']['approval_id'], review['approval']['approval_id'])

    def test_approved_special_parts_are_preserved_byte_for_byte(self):
        _, out = self.run_report()
        with zipfile.ZipFile(self.template) as a, zipfile.ZipFile(io.BytesIO(out)) as b:
            for part in ('xl/media/image4.emf', 'xl/drawings/vmlDrawing1.vml', 'xl/ctrlProps/ctrlProp1.xml',
                         'xl/worksheets/sheet3.xml', 'xl/externalLinks/externalLink1.xml',
                         'xl/externalLinks/_rels/externalLink1.xml.rels', 'xl/drawings/drawing1.xml'):
                self.assertEqual(a.read(part), b.read(part), part)
            self.assertIn(b'externalLinks/externalLink1.xml', b.read('xl/_rels/workbook.xml.rels'))
            self.assertNotIn(b'calcChain', b.read('xl/_rels/workbook.xml.rels'))
            out_text = b''.join(b.read(n) for n in ('xl/sharedStrings.xml', 'docProps/core.xml', 'xl/worksheets/sheet1.xml'))
            for legacy in list(LEGACY.values()) + [LEGACY_CREATOR]:
                self.assertNotIn(legacy.encode('utf-8'), out_text)

    def test_source_files_never_modified(self):
        before = self.template.read_bytes()
        mbefore = self.manifest_path.read_bytes()
        self.run_report()
        self.assertEqual(before, self.template.read_bytes())
        self.assertEqual(mbefore, self.manifest_path.read_bytes())

    def test_engine_changes_stay_inside_policy(self):
        _, out, _, _ = self.good_output()
        approved = AP.verify_manifest(self.template, self.manifest_path, self.trusted)
        _, log = CM.transform_package(self.template, {}, {}, {}, approval=approved)
        allowed = set(EP.MODIFIABLE_PARTS) | set(EP.DELETE_PARTS)
        self.assertTrue(set(log['changed_parts']) <= allowed, set(log['changed_parts']) - allowed)
        self.assertTrue(set(log['deleted_parts']) <= allowed, set(log['deleted_parts']) - allowed)


class TestStops(ApprovalBase):
    def test_one_byte_change_in_template_stops(self):
        data = bytearray(self.template.read_bytes())
        i = data.index(b'SYNTHETIC_CHECKBOX_SHAPE')
        data[i] = ord('Z')
        self.template.write_bytes(bytes(data))
        with self.assertRaisesRegex(ValueError, 'APPROVAL STOP: 承認された原本と違う'):
            AP.verify_manifest(self.template, self.manifest_path, self.trusted)
        with self.assertRaisesRegex(ValueError, 'APPROVAL STOP: このひな形の承認済み一覧がありません'):
            self.run_report()

    def test_unapproved_additions_to_template_stop(self):
        additions = {
            'comment': {'xl/comments1.xml': b'<comments xmlns="%s"/>' % S.encode()},
            'image': {'xl/media/image9.png': b'\x89PNG synthetic'},
            'external link': {'xl/externalLinks/externalLink2.xml': b'<externalLink/>'},
            'sheet': {'xl/worksheets/sheet4.xml': build_template_parts()['xl/worksheets/sheet3.xml']},
        }
        for label, add in additions.items():
            with self.subTest(label=label):
                tampered = self.dir / f'tampered_{label.replace(" ", "_")}.xlsx'
                tampered.write_bytes(rezip(self.template.read_bytes(), dict(add)))
                with self.assertRaisesRegex(ValueError, 'APPROVAL STOP'):
                    AP.verify_manifest(tampered, self.manifest_path, self.trusted)

    def test_parts_added_to_output_stop(self):
        approved, out, ru, pu = self.good_output()
        for name, data in (('xl/comments1.xml', b'<comments/>'), ('xl/media/image9.png', b'x'),
                           ('xl/externalLinks/externalLink2.xml', b'<externalLink/>'), ('xl/worksheets/sheet9.xml', b'<worksheet/>')):
            with self.subTest(part=name):
                with self.assertRaisesRegex(ValueError, '承認されていない部品'):
                    self.verify(approved, rezip(out, {name: data}), ru, pu)

    def test_immutable_part_changed_during_conversion_stops(self):
        approved, out, ru, pu = self.good_output()
        for part in ('xl/media/image4.emf', 'xl/ctrlProps/ctrlProp1.xml', 'xl/worksheets/sheet3.xml',
                     'xl/externalLinks/_rels/externalLink1.xml.rels'):
            with self.subTest(part=part):
                with zipfile.ZipFile(io.BytesIO(out)) as z:
                    changed = z.read(part) + b' '
                with self.assertRaisesRegex(ValueError, '変更禁止の部品が変わりました'):
                    self.verify(approved, rezip(out, {part: changed}), ru, pu)
        with self.assertRaisesRegex(ValueError, '削除が許可されていない'):
            self.verify(approved, rezip(out, remove={'xl/media/image4.emf'}), ru, pu)

    def test_external_link_reference_removed_from_workbook_rels_stops(self):
        approved, out, ru, pu = self.good_output()
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            rels = z.read('xl/_rels/workbook.xml.rels').replace(b'externalLinks/externalLink1.xml', b'externalLinks/gone.xml')
        with self.assertRaisesRegex(ValueError, 'workbook.xml.rels'):
            self.verify(approved, rezip(out, {'xl/_rels/workbook.xml.rels': rels}), ru, pu)

    def test_unexpected_cell_change_in_modifiable_sheet_stops(self):
        approved, out, ru, pu = self.good_output()
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        cell = root.find('.//{%s}c[@r="A2"]' % S)
        cell.find('{%s}v' % S).text = '0'
        with self.assertRaisesRegex(ValueError, '許可されていないセル.*A2'):
            self.verify(approved, rezip(out, {'xl/worksheets/sheet1.xml': ET.tostring(root)}), ru, pu)
        root.find('.//{%s}c[@r="A2"]' % S).find('{%s}v' % S).text = cell.find('{%s}v' % S).text
        merge = root.find('{%s}mergeCells' % S)
        merge[0].set('ref', 'A1:G1')
        with self.assertRaisesRegex(ValueError, 'セル以外の設定'):
            self.verify(approved, rezip(out, {'xl/worksheets/sheet1.xml': ET.tostring(root)}), ru, pu)

    def test_personal_info_residue_stops(self):
        # document properties left as they were in the template
        with mock.patch.object(CM, 'scrub_doc_properties', lambda xml, core: xml):
            with self.assertRaisesRegex(ValueError, '作成者'):
                self.run_report(name='residue_props')
        # old incident strings left in unreferenced shared strings
        with mock.patch.object(CM, 'scrub_orphan_shared_strings', lambda shared, sheets: (shared, 0)):
            with self.assertRaisesRegex(ValueError, '使われていない共有文字列に内容が残っています'):
                self.run_report(name='residue_shared')

    def test_photo_with_exif_stops(self):
        approved, out, ru, pu = self.good_output()
        exif = Image.Exif()
        exif[0x010F] = 'SYNTHETIC_CAMERA'
        b = io.BytesIO()
        Image.new('RGB', (120, 90), 'red').save(b, 'JPEG', exif=exif)
        with self.assertRaisesRegex(ValueError, 'EXIF'):
            self.verify(approved, rezip(out, {'xl/media/image1.jpeg': b.getvalue()}), ru, pu)

    def _mutated_sheet_output(self, original, edit):
        """Change only sheet1 XML of a valid generated report (all synthetic)."""
        with zipfile.ZipFile(io.BytesIO(original)) as z:
            root=ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        edit(root)
        return rezip(original, {'xl/worksheets/sheet1.xml': ET.tostring(root, encoding='UTF-8', xml_declaration=True)})

    def test_unapproved_empty_row_rejected(self):
        approved, out, ru, pu = self.good_output()
        def mutate(root):
            sd=root.find('{%s}sheetData' % S)
            ET.SubElement(sd, '{%s}row' % S, r='999', customHeight='1', ht='99')
        malicious=self._mutated_sheet_output(out,mutate)
        with self.assertRaisesRegex(ValueError, '未承認の行'):
            self.verify(approved,malicious,ru,pu)

    def test_duplicated_cell_rejected(self):
        approved, out, ru, pu = self.good_output()
        def mutate(root):
            sd=root.find('{%s}sheetData' % S)
            row=next(r for r in sd if r.get('r')=='6')
            c=next(c for c in row if c.get('r')=='G6')
            row.append(ET.fromstring(ET.tostring(c)))
        malicious=self._mutated_sheet_output(out,mutate)
        with self.assertRaisesRegex(ValueError, '重複セル'):
            self.verify(approved,malicious,ru,pu)

    def test_extra_sheetdata_attributes_rejected(self):
        approved, out, ru, pu = self.good_output()
        def mutate(root):
            sd=root.find('{%s}sheetData' % S)
            sd.set('unexpected', 'yes')
        malicious=self._mutated_sheet_output(out,mutate)
        with self.assertRaisesRegex(ValueError, 'sheetDataの設定'):
            self.verify(approved,malicious,ru,pu)

    def test_immutable_part_changed_during_conversion_stops(self):
        approved, out, ru, pu = self.good_output()
        for part in ('xl/media/image4.emf', 'xl/ctrlProps/ctrlProp1.xml', 'xl/worksheets/sheet3.xml',
                     'xl/externalLinks/_rels/externalLink1.xml.rels'):
            with self.subTest(part=part):
                with zipfile.ZipFile(io.BytesIO(out)) as z:
                    changed = z.read(part) + b' '
                with self.assertRaisesRegex(ValueError, '変更禁止の部品が変わりました'):
                    self.verify(approved, rezip(out, {part: changed}), ru, pu)
        with self.assertRaisesRegex(ValueError, '削除が許可されていない'):
            self.verify(approved, rezip(out, remove={'xl/media/image4.emf'}), ru, pu)

    def test_external_link_reference_removed_from_workbook_rels_stops(self):
        approved, out, ru, pu = self.good_output()
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            rels = z.read('xl/_rels/workbook.xml.rels').replace(b'externalLinks/externalLink1.xml', b'externalLinks/gone.xml')
        with self.assertRaisesRegex(ValueError, 'workbook.xml.rels'):
            self.verify(approved, rezip(out, {'xl/_rels/workbook.xml.rels': rels}), ru, pu)

    def test_unexpected_cell_change_in_modifiable_sheet_stops(self):
        approved, out, ru, pu = self.good_output()
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        cell = root.find('.//{%s}c[@r="A2"]' % S)
        cell.find('{%s}v' % S).text = '0'
        with self.assertRaisesRegex(ValueError, '許可されていないセル.*A2'):
            self.verify(approved, rezip(out, {'xl/worksheets/sheet1.xml': ET.tostring(root)}), ru, pu)
        root.find('.//{%s}c[@r="A2"]' % S).find('{%s}v' % S).text = cell.find('{%s}v' % S).text
        merge = root.find('{%s}mergeCells' % S)
        merge[0].set('ref', 'A1:G1')
        with self.assertRaisesRegex(ValueError, 'セル以外の設定'):
            self.verify(approved, rezip(out, {'xl/worksheets/sheet1.xml': ET.tostring(root)}), ru, pu)

    def test_personal_info_residue_stops(self):
        # document properties left as they were in the template
        with mock.patch.object(CM, 'scrub_doc_properties', lambda xml, core: xml):
            with self.assertRaisesRegex(ValueError, '作成者'):
                self.run_report(name='residue_props')
        # old incident strings left in unreferenced shared strings
        with mock.patch.object(CM, 'scrub_orphan_shared_strings', lambda shared, sheets: (shared, 0)):
            with self.assertRaisesRegex(ValueError, '使われていない共有文字列に内容が残っています'):
                self.run_report(name='residue_shared')

    def test_photo_with_exif_stops(self):
        approved, out, ru, pu = self.good_output()
        exif = Image.Exif()
        exif[0x010F] = 'SYNTHETIC_CAMERA'
        b = io.BytesIO()
        Image.new('RGB', (120, 90), 'red').save(b, 'JPEG', exif=exif)
        with self.assertRaisesRegex(ValueError, 'EXIF'):
            self.verify(approved, rezip(out, {'xl/media/image1.jpeg': b.getvalue()}), ru, pu)


    def test_reordered_rows_rejected(self):
        approved, out, ru, pu = self.good_output()
        def mutate(root):
            sd=root.find('{%s}sheetData' % S)
            rows=list(sd)
            sd.remove(rows[0]); sd.append(rows[0])
        malicious=self._mutated_sheet_output(out,mutate)
        with self.assertRaisesRegex(ValueError, '並び順'):
            self.verify(approved,malicious,ru,pu)


class TestManifestTrust(ApprovalBase):
    def test_missing_manifest_stops(self):
        self.manifest_path.unlink()
        with self.assertRaisesRegex(ValueError, '承認済み一覧がありません'):
            self.run_report()

    def test_tampered_manifest_stops(self):
        m = json.loads(self.manifest_path.read_text(encoding='utf-8'))
        for p in m['parts']:
            if p['name'] == 'xl/media/image4.emf':
                p['category'] = 'modifiable'
        tampered = self.dir / 'tampered.approved.json'
        tampered.write_text(json.dumps(m, ensure_ascii=False), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '署名が一致しません'):
            AP.verify_manifest(self.template, tampered, self.trusted)

    def test_flag_flip_on_pending_cannot_forge_approval(self):
        pending = TI.inspect_template(self.template)
        pending['status'] = 'approved'
        pending['approval'] = {'approval_id': 'APV-FAKE', 'approver': '偽', 'approver_key_id': self.entry['key_id']}
        pending['signature'] = {'alg': 'Ed25519', 'key_id': self.entry['key_id'], 'value': base64.b64encode(b'\0' * 64).decode()}
        forged = self.dir / 'forged.approved.json'
        forged.write_text(json.dumps(pending, ensure_ascii=False), encoding='utf-8')
        with self.assertRaisesRegex(ValueError, '署名が一致しません'):
            AP.verify_manifest(self.template, forged, self.trusted)

    def test_manifest_of_another_template_stops(self):
        other = write_zip(self.dir / 'other_template.xlsx', build_template_parts(variant='（別様式）'))
        other_manifest = self.approve(other, folder=self.dir)
        with self.assertRaisesRegex(ValueError, '承認された原本と違うファイル'):
            AP.verify_manifest(self.template, other_manifest, self.trusted)

    def test_untrusted_or_revoked_signer_stops(self):
        pem2, _ = SG.generate_keypair(PASS, '未登録の人')
        stranger = self.approve(self.template, pem=pem2, folder=self.dir)
        with self.assertRaisesRegex(ValueError, '登録されていない承認者'):
            AP.verify_manifest(self.template, stranger, self.trusted)
        revoked = self.write_trusted([dict(self.entry, revoked=True)], 'revoked.json')
        with self.assertRaisesRegex(ValueError, '失効'):
            AP.verify_manifest(self.template, self.manifest_path, revoked)

    def test_expired_approval_stops(self):
        past = AP.now_utc() - datetime.timedelta(days=30)
        path = self.approve(self.template, valid_days=1, now=past, folder=self.dir)
        with self.assertRaisesRegex(ValueError, '有効期限'):
            AP.verify_manifest(self.template, path, self.trusted)

    def test_engine_policy_change_requires_reapproval(self):
        with mock.patch.object(EP, 'ENGINE_POLICY_VERSION', 'report-2026-10-17/engine-9.9.9'):
            with self.assertRaisesRegex(ValueError, '再承認'):
                AP.verify_manifest(self.template, self.manifest_path, self.trusted)


if __name__ == '__main__':
    unittest.main()
