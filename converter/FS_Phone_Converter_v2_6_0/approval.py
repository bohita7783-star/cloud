#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Approved-template manifests: inventory, canonical form, Ed25519 verification.

Why a signature: an approval is valid only if it was signed with an approver's
PRIVATE key, which stays on the approver's PC. Editing the JSON (e.g. changing
"status" or a part's category) breaks the signature, so a flag flip cannot
forge an approval. The converter only holds PUBLIC keys (trusted_approvers.json).

Limit (documented in the admin procedure): someone who can edit this program
or trusted_approvers.json on the phone can still bypass checks. That is why the
audit log records the SHA-256 of the program files and the trusted-key file.
"""
import base64
import datetime
import hashlib
import json
import posixpath
import zipfile
from pathlib import Path

import engine_policy as EP

SCHEMA = 'fs-converter-approval/1'
TRUSTED_SCHEMA = 'fs-converter-trusted-approvers/1'
MAX_PARTS = 2000
MAX_TOTAL_BYTES = 200 * 1024 * 1024
HERE = Path(__file__).resolve().parent
DEFAULT_TRUSTED = HERE / 'trusted_approvers.json'
DEFAULT_APPROVALS_DIR = HERE / 'approvals'


class ApprovalError(ValueError):
    def __init__(self, message):
        super().__init__('APPROVAL STOP: ' + message)


def sha256_bytes(data):
    return hashlib.sha256(data).hexdigest()


def sha256_file(path):
    h = hashlib.sha256()
    with open(path, 'rb') as f:
        for block in iter(lambda: f.read(1 << 20), b''):
            h.update(block)
    return h.hexdigest()


def canonical(obj):
    return json.dumps(obj, ensure_ascii=False, sort_keys=True, separators=(',', ':')).encode('utf-8')


def signed_body(manifest):
    return {k: v for k, v in manifest.items() if k != 'signature'}


def key_id_for(public_raw):
    return 'ed25519:' + sha256_bytes(public_raw)[:16]


def now_utc():
    return datetime.datetime.now(datetime.timezone.utc).replace(microsecond=0)


def iso(dt):
    return dt.astimezone(datetime.timezone.utc).isoformat().replace('+00:00', 'Z')


def parse_iso(text):
    return datetime.datetime.fromisoformat(str(text).replace('Z', '+00:00'))


def part_inventory(template_path):
    """Name, SHA-256 and size of every ZIP member. Refuses unsafe archives."""
    out = []
    total = 0
    with zipfile.ZipFile(template_path) as z:
        infos = [i for i in z.infolist() if not i.is_dir()]
        names = [i.filename for i in infos]
        if len(names) > MAX_PARTS:
            raise ApprovalError('部品数が多すぎます')
        if len(names) != len(set(names)):
            raise ApprovalError('同じ名前の部品が重複しています')
        for info in infos:
            n = info.filename
            if n.startswith('/') or '\\' in n or '..' in n.split('/'):
                raise ApprovalError('不正な部品名があります')
            total += info.file_size
            if total > MAX_TOTAL_BYTES:
                raise ApprovalError('展開後のサイズが大きすぎます')
            data = z.read(info)
            out.append({'name': n, 'sha256': sha256_bytes(data), 'size': len(data)})
    return sorted(out, key=lambda p: p['name'])


def load_trusted(path=None):
    path = Path(path or DEFAULT_TRUSTED)
    if not path.is_file():
        raise ApprovalError('承認者の公開鍵一覧（trusted_approvers.json）がありません')
    try:
        data = json.loads(path.read_text(encoding='utf-8'))
    except Exception as ex:
        raise ApprovalError('承認者の公開鍵一覧を読めません') from ex
    if data.get('schema') != TRUSTED_SCHEMA:
        raise ApprovalError('承認者の公開鍵一覧の形式が違います')
    keys = {}
    for entry in data.get('keys', []):
        raw = base64.b64decode(entry['public_key'])
        if len(raw) != 32 or key_id_for(raw) != entry.get('key_id'):
            raise ApprovalError('公開鍵一覧に壊れた鍵があります')
        keys[entry['key_id']] = dict(entry, raw=raw)
    return keys, sha256_file(path)


def _verify_signature(public_raw, signature_b64, body):
    try:
        from cryptography.hazmat.primitives.asymmetric.ed25519 import Ed25519PublicKey
        from cryptography.exceptions import InvalidSignature
    except ImportError as ex:
        raise ApprovalError('署名の検証に必要な cryptography がありません（Termux: pkg install python-cryptography）') from ex
    try:
        Ed25519PublicKey.from_public_bytes(public_raw).verify(base64.b64decode(signature_b64), canonical(body))
    except (InvalidSignature, ValueError) as ex:
        raise ApprovalError('承認一覧の署名が一致しません（改ざん、または別の鍵で署名されています）') from ex


def resolve_manifest_path(template_path, approvals_dir=None):
    folder = Path(approvals_dir or DEFAULT_APPROVALS_DIR)
    return folder / (sha256_file(template_path) + '.approved.json')


def verify_manifest(template_path, manifest_path=None, trusted_path=None, approvals_dir=None, now=None):
    """Return the verified manifest, or raise ApprovalError. Never modifies files."""
    template_path = Path(template_path)
    actual_sha = sha256_file(template_path)
    manifest_path = Path(manifest_path) if manifest_path else resolve_manifest_path(template_path, approvals_dir)
    if not manifest_path.is_file():
        raise ApprovalError('このひな形の承認済み一覧がありません（' + manifest_path.name + '）。管理者の承認が必要です')
    raw_bytes = manifest_path.read_bytes()
    try:
        manifest = json.loads(raw_bytes.decode('utf-8'))
    except Exception as ex:
        raise ApprovalError('承認済み一覧を読めません') from ex
    if manifest.get('schema') != SCHEMA:
        raise ApprovalError('承認済み一覧の形式が違います')
    if manifest.get('status') != 'approved' or 'signature' not in manifest or 'approval' not in manifest:
        raise ApprovalError('承認候補（未承認）の一覧です。管理者の署名がありません')
    sig = manifest['signature']
    approval = manifest['approval']
    if sig.get('alg') != 'Ed25519':
        raise ApprovalError('署名方式が違います')
    keys, trusted_sha = load_trusted(trusted_path)
    key = keys.get(sig.get('key_id'))
    if key is None:
        raise ApprovalError('登録されていない承認者の鍵で署名されています')
    if key.get('revoked'):
        raise ApprovalError('失効した承認者の鍵で署名されています')
    if approval.get('approver_key_id') != sig.get('key_id'):
        raise ApprovalError('承認者と署名の鍵が一致しません')
    _verify_signature(key['raw'], sig.get('value', ''), signed_body(manifest))
    # From here on the content is authentic; check it applies to THIS file and THIS engine.
    if manifest.get('template', {}).get('sha256') != actual_sha:
        raise ApprovalError('承認された原本と違うファイルです（原本が変更された、または別のひな形の一覧です）')
    if manifest.get('layout_id') != EP.LAYOUT_ID:
        raise ApprovalError('承認対象の様式の種類が、この変換ツールと違います')
    if manifest.get('engine_policy_version') != EP.ENGINE_POLICY_VERSION:
        raise ApprovalError('変換ツールの変更範囲が承認時から変わっています。再承認が必要です')
    now = now or now_utc()
    if approval.get('valid_until') and now > parse_iso(approval['valid_until']):
        raise ApprovalError('承認の有効期限が切れています')
    actual_parts = part_inventory(template_path)
    listed = {p['name']: p for p in manifest.get('parts', [])}
    if {p['name'] for p in actual_parts} != set(listed):
        raise ApprovalError('原本の部品構成が承認済み一覧と違います')
    missing = [n for n in EP.REQUIRED_PARTS if n not in listed]
    if missing:
        raise ApprovalError('変換に必要な部品がひな形にありません: ' + ', '.join(missing))
    # 2.6.0: which form this is and how photos go in, derived again from the file itself
    import form_profile as FP
    with zipfile.ZipFile(template_path) as z:
        try:
            plan = FP.analyze_zip(z)
        except FP.FormError as ex:
            raise ApprovalError(str(ex)) from ex
    if plan['mode'] == 'refused':
        raise ApprovalError('このひな形は使えません：' + ' / '.join(plan['problems'][:3]))
    if manifest.get('form') != FP.summary(plan):
        raise ApprovalError('承認された様式の配置（写真枠など）と、変換ツールが読み取った配置が違います。再承認が必要です')
    if plan['mode'] == 'baseline':
        missing = [n for n in EP.BASELINE_REQUIRED if n not in listed]
        if missing:
            raise ApprovalError('変換に必要な部品がひな形にありません: ' + ', '.join(missing))
    for p in actual_parts:
        entry = listed[p['name']]
        if entry.get('sha256') != p['sha256']:
            raise ApprovalError('原本の部品が承認時と違います: ' + p['name'])
        if entry.get('category') != EP.category_for(p['name'], plan):
            raise ApprovalError('承認一覧の変更区分が変換ツールのコードと一致しません: ' + p['name'])
    return dict(manifest, _manifest_sha256=sha256_bytes(raw_bytes), _manifest_file=manifest_path.name,
                _trusted_sha256=trusted_sha, _plan=plan)
