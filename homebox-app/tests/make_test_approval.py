"""Create SYNTHETIC test inputs for tests/e2e_excel.js and a served test copy.

  python homebox-app/tests/make_test_approval.py <out_dir>

Writes into <out_dir>:
  template.xlsx              synthetic company-style template (EMF, VML, checkbox, extra sheet, external link)
  approved.json              approval manifest signed with a THROWAWAY key (private key is never written)
  tampered.json              same manifest with the approver name changed (signature must fail)
  trusted_approvers.json     the throwaway public key (put this ONLY into a served test copy of homebox-app)
  photo_exif.jpg, photo2.jpg two different synthetic photos (the first carries EXIF DateTimeOriginal)
"""
import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
CONV = ROOT / 'converter' / 'FS_Phone_Converter_v2_5_0'
sys.path[:0] = [str(CONV), str(CONV / 'tests')]

from PIL import Image
import approval as AP
import approval_sign as SG
import template_inspector as TI
from synthetic_template import build_template_parts, write_zip

out = Path(sys.argv[1])
out.mkdir(parents=True, exist_ok=True)
tpl = write_zip(out / 'template.xlsx', build_template_parts())
passphrase = 'throwaway-test-passphrase'
pem, entry = SG.generate_keypair(passphrase, '架空 承認者')
(out / 'trusted_approvers.json').write_text(json.dumps({'schema': AP.TRUSTED_SCHEMA, 'keys': [entry]}, ensure_ascii=False), encoding='utf-8')
manifest = SG.sign_pending(tpl, TI.inspect_template(tpl), pem, passphrase, '架空 承認者', True)
del pem
(out / 'approved.json').write_text(json.dumps(manifest, ensure_ascii=False), encoding='utf-8')
manifest['approval']['approver'] = 'なりすまし（架空）'
(out / 'tampered.json').write_text(json.dumps(manifest, ensure_ascii=False), encoding='utf-8')
im = Image.new('RGB', (1200, 900), (40, 70, 120))
exif = Image.Exif()
ifd = exif.get_ifd(0x8769)
ifd[0x9003] = '2026:11:01 09:15:00'
ifd[0x9011] = '+09:00'
im.save(out / 'photo_exif.jpg', exif=exif, quality=85)
Image.new('RGB', (1200, 900), (150, 60, 40)).save(out / 'photo2.jpg', quality=85)
print('written to', out)
