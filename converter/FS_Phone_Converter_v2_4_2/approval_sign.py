#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Approver tools. Run ONLY on the approver's own PC, never on the field phone.

  python approval_sign.py keygen --name "承認者名" --out <folder>
      -> approver_private_key.pem (passphrase-encrypted, keep secret)
         approver_public_entry.json (send to whoever maintains trusted_approvers.json)

  python approval_sign.py sign --template blank.xlsx --pending approval_manifest.pending.json
         --key approver_private_key.pem --approver "承認者名" --i-reviewed-in-excel
         [--valid-days 365] --out approvals/
      -> approvals/<template sha256>.approved.json (read-only)

Signing re-inspects the template and refuses if it differs from the pending
candidate, so the file that was reviewed is exactly the file that is approved.
"""
import argparse
import base64
import datetime
import getpass
import json
import os
import stat
import sys
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))  # find sibling modules in any launch mode

import approval as AP
import engine_policy as EP
import template_inspector as TI

REVIEW_STATEMENT = ('承認者は、このひな形をExcelで開き、全シート（非表示を含む）、コメント、外部リンクの参照先、'
                    '名前の定義、図形・画像・EMF、チェックボックス、文書情報を確認し、過去の案件の情報や'
                    '不要な個人情報が含まれていないことを確認した。')


def _crypto():
    try:
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PrivateKey
        from cryptography.hazmat.primitives import serialization
    except ImportError:
        sys.exit('cryptography が必要です: pip install cryptography')
    return Ed25519PrivateKey, serialization


def generate_keypair(passphrase, name):
    if not passphrase or len(passphrase) < 12:
        raise ValueError('秘密鍵のパスフレーズは12文字以上にしてください')
    Ed25519PrivateKey, serialization = _crypto()
    private = Ed25519PrivateKey.generate()
    pem = private.private_bytes(serialization.Encoding.PEM, serialization.PrivateFormat.PKCS8,
                                serialization.BestAvailableEncryption(passphrase.encode('utf-8')))
    raw = private.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw)
    entry = {'key_id': AP.key_id_for(raw), 'name': name, 'public_key': base64.b64encode(raw).decode('ascii'),
             'registered_at': AP.iso(AP.now_utc()), 'revoked': False}
    return pem, entry


def sign_pending(template_path, pending, private_pem, passphrase, approver, reviewed, valid_days=None, now=None):
    if not reviewed:
        raise ValueError('Excelでの確認が済んでいない承認候補には署名しません（--i-reviewed-in-excel が必要）')
    if not approver or not approver.strip():
        raise ValueError('承認者名が必要です')
    if pending.get('schema') != AP.SCHEMA or pending.get('status') != 'pending':
        raise ValueError('承認候補（pending）の一覧ではありません')
    if pending['inspection'].get('blockers'):
        raise ValueError('承認できない理由が残っています: ' + '; '.join(pending['inspection']['blockers']))
    now = now or AP.now_utc()
    fresh = TI.inspect_template(template_path, now=now)
    if fresh['template']['sha256'] != pending['template']['sha256']:
        raise ValueError('検査した後にひな形が変わっています。もう一度検査してください')
    if [(p['name'], p['sha256'], p['category']) for p in fresh['parts']] != \
       [(p['name'], p['sha256'], p['category']) for p in pending['parts']]:
        raise ValueError('承認候補の部品一覧が、現在のひな形・変換ツールと一致しません')
    if pending.get('engine_policy_version') != EP.ENGINE_POLICY_VERSION:
        raise ValueError('承認候補を作った変換ツールの版が、署名ツールの版と違います')
    _, serialization = _crypto()
    private = serialization.load_pem_private_key(private_pem, password=passphrase.encode('utf-8'))
    raw = private.public_key().public_bytes(serialization.Encoding.Raw, serialization.PublicFormat.Raw)
    key_id = AP.key_id_for(raw)
    manifest = dict(pending)
    manifest['status'] = 'approved'
    manifest['approval'] = {
        'approval_id': 'APV-' + now.strftime('%Y%m%d') + '-' + uuid.uuid4().hex[:8],
        'approver': approver.strip(),
        'approver_key_id': key_id,
        'approved_at': AP.iso(now),
        'valid_until': AP.iso(now + datetime.timedelta(days=valid_days)) if valid_days else None,
        'review_statement': REVIEW_STATEMENT,
    }
    signature = private.sign(AP.canonical(AP.signed_body(manifest)))
    manifest['signature'] = {'alg': 'Ed25519', 'key_id': key_id, 'value': base64.b64encode(signature).decode('ascii')}
    return manifest


def write_read_only(path, text):
    path = Path(path)
    if path.exists():
        raise FileExistsError(f'既にあります（上書きしません）: {path}')
    path.write_text(text, encoding='utf-8')
    os.chmod(path, stat.S_IRUSR | stat.S_IRGRP | stat.S_IROTH)


def main(argv=None):
    ap = argparse.ArgumentParser(description='Approver tools (admin PC only)')
    sub = ap.add_subparsers(dest='cmd', required=True)
    k = sub.add_parser('keygen')
    k.add_argument('--name', required=True)
    k.add_argument('--out', required=True)
    s = sub.add_parser('sign')
    s.add_argument('--template', required=True)
    s.add_argument('--pending', required=True)
    s.add_argument('--key', required=True)
    s.add_argument('--approver', required=True)
    s.add_argument('--i-reviewed-in-excel', action='store_true')
    s.add_argument('--valid-days', type=int)
    s.add_argument('--out', required=True)
    a = ap.parse_args(argv)
    if a.cmd == 'keygen':
        out = Path(a.out)
        out.mkdir(parents=True, exist_ok=True)
        p1 = getpass.getpass('秘密鍵のパスフレーズ（12文字以上）: ')
        if p1 != getpass.getpass('もう一度: '):
            sys.exit('パスフレーズが一致しません')
        pem, entry = generate_keypair(p1, a.name)
        priv = out / 'approver_private_key.pem'
        if priv.exists():
            sys.exit('秘密鍵が既にあります（上書きしません）')
        priv.write_bytes(pem)
        os.chmod(priv, stat.S_IRUSR | stat.S_IWUSR)
        write_read_only(out / 'approver_public_entry.json', json.dumps(entry, ensure_ascii=False, indent=2))
        print('秘密鍵（他人に渡さない・スマホに置かない）:', priv)
        print('公開鍵の登録用:', out / 'approver_public_entry.json', ' key_id =', entry['key_id'])
        return 0
    pending = json.loads(Path(a.pending).read_text(encoding='utf-8'))
    passphrase = getpass.getpass('秘密鍵のパスフレーズ: ')
    manifest = sign_pending(a.template, pending, Path(a.key).read_bytes(), passphrase,
                            a.approver, a.i_reviewed_in_excel, a.valid_days)
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    target = out / (manifest['template']['sha256'] + '.approved.json')
    write_read_only(target, json.dumps(manifest, ensure_ascii=False, indent=2))
    print('承認済み一覧:', target, ' 承認ID =', manifest['approval']['approval_id'])
    return 0


if __name__ == '__main__':
    sys.exit(main())
