#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Inspect a candidate company template and write an approval CANDIDATE.

Usage (admin PC):
  python template_inspector.py <blank_template.xlsx> --out <folder>

Writes:
  inspection_report.md          human-readable report for the reviewer
  approval_manifest.pending.json  approval candidate (status "pending", unsigned)

The report never prints cell values, comment text, document-property values
or external-link paths: only part names, sizes, hashes, counts and flags.
Inspection alone never approves anything; see approval_sign.py.
"""
import argparse
import datetime
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # find sibling modules in any launch mode

from lxml import etree as ET

import approval as AP
import engine_policy as EP

INSPECTOR_VERSION = '2.4.2-homebox-preview'
SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
NS = {'m': SS}


def _xml(data):
    return ET.fromstring(data, parser=ET.XMLParser(resolve_entities=False, no_network=True, huge_tree=False))


def _sheets(z):
    try:
        wb = _xml(z.read('xl/workbook.xml'))
        rels = _xml(z.read('xl/_rels/workbook.xml.rels'))
    except Exception:
        return []
    targets = {r.get('Id'): r.get('Target', '') for r in rels}
    result = []
    for sh in wb.findall('.//m:sheet', NS):
        t = targets.get(sh.get('{%s}id' % REL), '')
        part = t.lstrip('/') if t.startswith('/') else posixpath.normpath(posixpath.join('xl', t))
        result.append({'name': sh.get('name', ''), 'state': sh.get('state', 'visible'), 'part': part,
                       'example_like': bool(re.search(r'(記入例|サンプル|sample|example)', sh.get('name', ''), re.I))})
    return result


def _external_relationships(z):
    """Count external relationships per .rels part; never return the targets."""
    found = []
    for name in sorted(n for n in z.namelist() if n.endswith('.rels')):
        try:
            root = _xml(z.read(name))
        except Exception:
            found.append({'rels': name, 'count': None, 'kinds': ['読めません']})
            continue
        kinds = []
        for r in root:
            if r.get('TargetMode', '').lower() == 'external':
                target = r.get('Target', '')
                kind = ('Web' if target.lower().startswith(('http:', 'https:')) else
                        'ネットワーク/ファイル' if target.lower().startswith(('file:', '\\\\')) or re.match(r'^[a-z]:', target, re.I) else '相対パス')
                kinds.append(kind + '（' + r.get('Type', '').rsplit('/', 1)[-1] + '）')
        if kinds:
            found.append({'rels': name, 'count': len(kinds), 'kinds': kinds})
    return found


def _doc_props_presence(z):
    out = {}
    for part, fields in (('docProps/core.xml', ('creator', 'lastModifiedBy', 'title', 'subject', 'description', 'keywords')),
                         ('docProps/app.xml', ('Company', 'Manager'))):
        if part not in z.namelist():
            continue
        root = _xml(z.read(part))
        for el in root.iter():
            tag = el.tag.split('}')[-1]
            if tag in fields:
                out[part + ':' + tag] = '値あり' if (el.text or '').strip() else '空'
    return out


def _counts(z):
    names = z.namelist()
    comments = 0
    for n in names:
        if re.match(r'xl/comments\d*\.xml$', n):
            try:
                comments += len(_xml(z.read(n)).findall('.//m:comment', NS))
            except Exception:
                pass
    defined = 0
    try:
        defined = len(_xml(z.read('xl/workbook.xml')).findall('.//m:definedName', NS))
    except Exception:
        pass
    return {'comments': comments, 'defined_names': defined,
            'ctrl_props': sum(1 for n in names if n.startswith('xl/ctrlProps/')),
            'vml': sum(1 for n in names if n.lower().endswith('.vml')),
            'emf': sum(1 for n in names if n.lower().endswith('.emf')),
            'media_other_than_photo_slots': sum(1 for n in names if n.startswith('xl/media/') and n not in EP.ALLOWED_PHOTO_PARTS)}


def inspect_template(path, now=None):
    path = Path(path)
    now = now or AP.now_utc()
    inventory = AP.part_inventory(path)
    with zipfile.ZipFile(path) as z:
        sheets = _sheets(z)
        external = _external_relationships(z)
        props = _doc_props_presence(z)
        counts = _counts(z)
    parts = []
    for p in inventory:
        cat = EP.category_for(p['name'])
        parts.append(dict(p, category=cat, category_label=EP.CATEGORY_LABELS[cat],
                          purpose=EP.purpose_for(p['name']), risk_flags=EP.risk_flags(p['name'])))
    names = {p['name'] for p in parts}
    missing_required = [n for n in EP.REQUIRED_PARTS if n not in names]
    blockers = []
    if missing_required:
        blockers.append('変換に必要な部品がありません: ' + ', '.join(missing_required))
    sheet_parts = [s['part'] for s in sheets]
    if sheet_parts[:2] != [EP.SHEET_REPORT, EP.SHEET_PHOTO]:
        blockers.append('1枚目が報告書、2枚目が写真帳という前提と違います')
    pending = {
        'schema': AP.SCHEMA,
        'status': 'pending',
        'layout_id': EP.LAYOUT_ID,
        'engine_policy_version': EP.ENGINE_POLICY_VERSION,
        'template': {'file_name': path.name, 'sha256': AP.sha256_file(path), 'size': path.stat().st_size},
        'parts': parts,
        'inspection': {
            'inspector_version': INSPECTOR_VERSION,
            'inspected_at': AP.iso(now),
            'sheets': sheets,
            'external_relationships': external,
            'document_properties': props,
            'counts': counts,
            'blockers': blockers,
        },
    }
    return pending


def render_report(pending):
    ins = pending['inspection']
    parts = pending['parts']
    lines = []
    w = lines.append
    w('# ひな形検査報告（承認候補）')
    w('')
    w('> この報告は機械的な検査結果です。**これだけでは承認になりません。** 管理者がExcelで中身を確認し、approval_sign.py で署名したものだけが承認済みになります。')
    w('> セルの値・コメント本文・文書情報の値・外部リンクの参照先は、この報告には表示しません。Excelで直接確認してください。')
    w('')
    w('| 項目 | 内容 |')
    w('|---|---|')
    w(f"| ファイル名 | {pending['template']['file_name']} |")
    w(f"| ファイル全体のSHA-256 | `{pending['template']['sha256']}` |")
    w(f"| サイズ | {pending['template']['size']:,} バイト |")
    w(f"| 様式の種類 | {pending['layout_id']} |")
    w(f"| 変換ツールの変更範囲の版 | {pending['engine_policy_version']} |")
    w(f"| 検査日時（UTC） | {ins['inspected_at']} |")
    w(f"| 検査ツール | {ins['inspector_version']} |")
    w('')
    if ins['blockers']:
        w('## 承認できない理由（先に解消が必要）')
        for b in ins['blockers']:
            w('- ' + b)
        w('')
    w('## シート構成')
    w('')
    w('| 順番 | シート名 | 表示状態 | 部品 | 変換ツールの扱い | 注意 |')
    w('|---|---|---|---|---|---|')
    for i, s in enumerate(ins['sheets'], 1):
        role = {EP.SHEET_REPORT: '報告書として書き込む', EP.SHEET_PHOTO: '写真帳として書き込む'}.get(s['part'], '変更しない（承認後はバイト一致を検査）')
        note = []
        if s['state'] != 'visible':
            note.append('**非表示**')
        if s['example_like']:
            note.append('**記入例・サンプルらしい名前**')
        w(f"| {i} | {s['name']} | {s['state']} | {s['part']} | {role} | {'、'.join(note)} |")
    w('')
    c = ins['counts']
    w('## 人が確認すべきもの')
    w('')
    w('| 確認項目 | 件数 | Excelでの確認場所 |')
    w('|---|---|---|')
    w(f"| コメント | {c['comments']} | 校閲 → コメントの表示 |")
    w(f"| 外部への参照 | {sum(e['count'] or 0 for e in ins['external_relationships'])} | データ → リンクの編集（参照先パスを確認） |")
    w(f"| 名前の定義 | {c['defined_names']} | 数式 → 名前の管理 |")
    w(f"| チェックボックス等 | {c['ctrl_props']} | 開発 → デザインモードで各コントロール |")
    w(f"| VML図形 | {c['vml']} | チェックボックス・コメント枠の正体 |")
    w(f"| EMF画像 | {c['emf']} | 各シートの図・ロゴ |")
    w(f"| 写真枠以外の画像 | {c['media_other_than_photo_slots']} | 各シートの図・画像に過去の写真が無いか |")
    w('')
    if ins['external_relationships']:
        w('外部参照のある部品（参照先は表示しません）:')
        for e in ins['external_relationships']:
            w(f"- {e['rels']}: {e['count']}件（{'、'.join(e['kinds'])}）")
        w('')
    if ins['document_properties']:
        w('文書情報（値は表示しません。変換時に空にする項目は出力側だけで空にします）:')
        for k, v in ins['document_properties'].items():
            w(f'- {k}: {v}')
        w('')
    w('## 部品一覧')
    w('')
    w('| 部品 | 用途 | 変更区分 | 要確認 | サイズ | SHA-256（先頭16桁） |')
    w('|---|---|---|---|---|---|')
    for p in parts:
        w(f"| {p['name']} | {p['purpose']} | {p['category_label']} | {'、'.join(p['risk_flags'])} | {p['size']:,} | `{p['sha256'][:16]}` |")
    w('')
    w('## 変更区分の意味')
    w('')
    w('- **変更禁止**: 変換後もバイト単位で同一であることを検査します。')
    w('- **指定範囲のみ変更可**: 変換ツールのコード（engine_policy.py）に書かれた範囲だけ変わることを検査します。')
    for name, (purpose, what) in EP.MODIFIABLE_PARTS.items():
        if name in {p['name'] for p in parts}:
            w(f'  - {name}（{purpose}）: {what}')
    w('- **削除予定**: 変換時に削除します。' + '、'.join(f'{n}（{v[1]}）' for n, v in EP.DELETE_PARTS.items()))
    w('- **生成予定**: この版では新しく作る部品はありません。それ以外の部品が増えたら停止します。')
    return '\n'.join(lines) + '\n'


def main(argv=None):
    ap = argparse.ArgumentParser(description='Company template inspector ' + INSPECTOR_VERSION)
    ap.add_argument('template')
    ap.add_argument('--out', required=True, help='報告書と承認候補を書き出すフォルダー')
    a = ap.parse_args(argv)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    pending = inspect_template(a.template)
    report_path = out / 'inspection_report.md'
    pending_path = out / 'approval_manifest.pending.json'
    for p in (report_path, pending_path):
        if p.exists():
            sys.exit(f'既にあります（上書きしません）: {p}')
    report_path.write_text(render_report(pending), encoding='utf-8')
    pending_path.write_text(json.dumps(pending, ensure_ascii=False, indent=2), encoding='utf-8')
    print('検査報告:', report_path)
    print('承認候補:', pending_path)
    if pending['inspection']['blockers']:
        print('注意: 承認できない理由があります。報告書の先頭を確認してください。')
        return 2
    return 0


if __name__ == '__main__':
    sys.exit(main())
