"""Regenerate the files in samples/ from SYNTHETIC data only.

  python tests/make_samples.py

A throwaway key pair is created for the sample signature and its private key
is never written to disk. The sample public key is NOT registered in
trusted_approvers.json, so the sample approval cannot be used for real runs.
"""
import io
import json
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path[:0] = [str(ROOT), str(HERE)]

from PIL import Image
import approval as AP
import approval_sign as SG
import converter_mobile as CM
import template_inspector as TI
from synthetic_template import build_template_parts, write_zip
from test_converter import make_workbook


def main():
    out = ROOT / 'samples'
    out.mkdir(exist_ok=True)
    template = write_zip(out / 'synthetic_blank_template.xlsx', build_template_parts())
    pending = TI.inspect_template(template)
    (out / 'inspection_report_sample.md').write_text(TI.render_report(pending), encoding='utf-8')
    (out / 'approval_manifest.pending_sample.json').write_text(json.dumps(pending, ensure_ascii=False, indent=2), encoding='utf-8')
    passphrase = 'throwaway-sample-passphrase'
    pem, entry = SG.generate_keypair(passphrase, 'サンプル承認者（架空）')
    manifest = SG.sign_pending(template, pending, pem, passphrase, 'サンプル承認者（架空）', True, valid_days=365)
    del pem
    (out / 'approval_manifest.approved_sample.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')
    trusted = {'schema': AP.TRUSTED_SCHEMA, 'note': 'サンプル専用。本番の trusted_approvers.json にコピーしないこと。', 'keys': [entry]}
    (out / 'sample_trusted_approvers.json').write_text(json.dumps(trusted, ensure_ascii=False, indent=2), encoding='utf-8')
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        field = make_workbook(td / 'synthetic_field_survey.xlsx')
        photos = td / 'synthetic_photos.zip'
        b = io.BytesIO()
        Image.new('RGB', (400, 300), 'navy').save(b, 'JPEG')
        with zipfile.ZipFile(photos, 'w') as z:
            z.writestr('photoB.jpg', b.getvalue())
        review = CM.make_report(str(field), str(photos), str(template), str(td / 'report.xlsx'), str(td / 'audit.json'),
                                case_id='ID_B', approval_manifest=out / 'approval_manifest.approved_sample.json',
                                trusted_keys=out / 'sample_trusted_approvers.json')
        (out / 'conversion_audit_sample.json').write_text((td / 'audit.json').read_text(encoding='utf-8'), encoding='utf-8')
    print('samples written:', sorted(p.name for p in out.iterdir()))
    print('approval_id:', review['approval']['approval_id'])


if __name__ == '__main__':
    main()
