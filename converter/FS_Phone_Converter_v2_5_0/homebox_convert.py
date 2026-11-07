#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""HomeBox report package -> company Excel (approved template only).

  bash ~/FS_Phone_Converter_v2_4_1/run_homebox.sh            newest HB-REPORT-*.zip in Download
  python homebox_convert.py --package path/to/HB-REPORT-....zip

Output: Download/FS_変換結果/<件名>_会社書式_要確認_<日時>_<id>.xlsx (+ .json audit log)
The package, the template and the approval files are never modified.
"""
import argparse
import datetime
import json
import re
import sys
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from converter_mobile import make_report_from_homebox, VERSION
from case_exports import output_report_name

HERE = Path(__file__).resolve().parent
TEMPLATE = HERE / 'company_template.xlsx'
DOWNLOADS = Path.home() / 'storage' / 'downloads'
REPORT_DIR = DOWNLOADS / 'FS_変換結果'
PACKAGE_RE = re.compile(r'^HB-REPORT-.+\.zip$', re.I)


def newest_package(folder):
    found = [p for p in folder.glob('**/HB-REPORT-*.zip') if PACKAGE_RE.match(p.name) and REPORT_DIR not in p.parents]
    return max(found, key=lambda p: p.stat().st_mtime) if found else None


def convert(package, template=TEMPLATE, out_dir=REPORT_DIR, **approval_args):
    package = Path(package)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime('%Y%m%d%H%M')
    unique = uuid.uuid4().hex[:8]
    import homebox_adapter as HB
    case, _, _, _ = HB.read_package(package)
    output = out_dir / output_report_name(case.get('件名'), stamp, unique)
    audit = output.with_suffix('.json')
    review = make_report_from_homebox(package, template, output, audit, **approval_args)
    return output, audit, review


def main(argv=None):
    ap = argparse.ArgumentParser(description='HomeBox -> company Excel ' + VERSION)
    ap.add_argument('--package')
    ap.add_argument('--template', default=str(TEMPLATE))
    ap.add_argument('--out-dir', default=str(REPORT_DIR))
    ap.add_argument('--approval')
    ap.add_argument('--trusted')
    a = ap.parse_args(argv)
    package = Path(a.package) if a.package else newest_package(DOWNLOADS)
    if not package or not package.is_file():
        print('HomeBoxの報告書パッケージ（HB-REPORT-*.zip）が見つかりません。HomeBoxで「報告書パッケージを書き出す」を押してください。')
        return 1
    print('=== HomeBox → 会社書式', VERSION, '===')
    print('パッケージ:', package.name)
    try:
        output, audit, review = convert(package, a.template, a.out_dir,
                                        approval_manifest=a.approval, trusted_keys=a.trusted)
    except Exception as ex:
        print('停止しました:', ex)
        print('元のファイルは変更していません。必要なら手作業で報告書を作成してください。')
        return 2
    print('完成（要確認）:', output)
    print('変換ログ:', audit)
    print('承認ID:', review['approval']['approval_id'], '／ 写真', review['photo_count'], '枚 ／ 対応経過', review['timeline_count'], '件')
    if review.get('unmapped_nonempty_fields'):
        print('報告書に自動転記されない入力があります（提出前に確認）:', '、'.join(review['unmapped_nonempty_fields']))
    return 0


if __name__ == '__main__':
    sys.exit(main())
