#!/usr/bin/env python3
# Repro checks for HomeBox converter v2.4.0, synthetic files only.
# Usage: python homebox_v240_regression_checks.py /path/to/FS_Phone_Converter_v2_4_0
from pathlib import Path
import sys, tempfile, zipfile

if len(sys.argv) != 2:
    raise SystemExit('Usage: python homebox_v240_regression_checks.py /path/to/FS_Phone_Converter_v2_4_0')
root = Path(sys.argv[1]).resolve()
sys.path[:0] = [str(root), str(root / 'tests')]
import converter_mobile as CM
import homebox_adapter as HB
from test_homebox import make_package
from lxml import etree as ET

results = []
with tempfile.TemporaryDirectory() as scratch:
    d = Path(scratch)

    # 1. A new row not present in the approved template should be rejected.
    with zipfile.ZipFile(root / 'samples' / 'synthetic_blank_template.xlsx') as z:
        original = z.read('xl/worksheets/sheet1.xml')
    ws = CM.xml_root(original)
    sd = ws.find('m:sheetData', CM.ns)
    ET.SubElement(sd, '{' + CM.SS + '}row', r='999')
    edited = ET.tostring(ws)
    try:
        CM._check_sheet(original, edited, set(), 'synthetic report')
        results.append(('未承認の空行', '未検出（要修正）'))
    except ValueError:
        results.append(('未承認の空行', '検出済み'))

    # 2. Every photo is required to carry a nonempty SHA-256, per DATA_CONTRACT.
    def remove_hash(pkg, files):
        del pkg['photos'][0]['sha256']
    p = make_package(d / 'photo_without_hash.zip', mutate=remove_hash)
    try:
        HB.read_package(p)
        results.append(('写真ハッシュの欠落', '未検出（要修正）'))
    except ValueError:
        results.append(('写真ハッシュの欠落', '検出済み'))

    # 3. The equipment ID embedded in the photo's filename must match case equipment.
    def cross_equipment(pkg, files):
        before = pkg['photos'][0]['元写真']
        after = before.replace('EQ-000006', 'EQ-999999')
        pkg['photos'][0]['元写真'] = after
        files[after] = files.pop(before)
    p = make_package(d / 'photo_wrong_equipment.zip', mutate=cross_equipment)
    try:
        HB.read_package(p)
        results.append(('写真名と案件の設備ID不一致', '未検出（要修正）'))
    except ValueError:
        results.append(('写真名と案件の設備ID不一致', '検出済み'))

    # 4. The package contract says unknown items must cause a stop.
    def unknown_root(pkg, files):
        pkg['unexpected_field'] = 'synthetic'
    p = make_package(d / 'unknown_root_key.zip', mutate=unknown_root)
    try:
        HB.read_package(p)
        results.append(('package.jsonの不明なトップレベル項目', '未検出（要修正）'))
    except ValueError:
        results.append(('package.jsonの不明なトップレベル項目', '検出済み'))

print('HomeBox converter negative checks (synthetic data only):')
for name, state in results:
    print(f' - {name}: {state}')
print('未検出:', sum('未検出' in status for _, status in results), '/', len(results))
