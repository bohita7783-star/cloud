#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Read a HomeBox report package (hb-report-package/1) for the approved engine.

The package is a ZIP written by the HomeBox app:
  package.json        case / equipment / photos / timeline (Japanese keys = Field Survey names)
  photos/<file>.jpg   original photos, unmodified

Output has exactly the shape read_field_tables() produces for Field Survey,
so the report engine and its approval checks are reused unchanged.
Nothing here writes files.
"""
import hashlib
import json
import posixpath
import re
import zipfile

from converter_mobile import zip_guard, MAX_TEXT_LENGTH

SCHEMA = 'hb-report-package/2'
ACCEPTED_SCHEMAS = ('hb-report-package/1', 'hb-report-package/2')
MAX_PACKAGE = 160 * 1024 * 1024

CASE_KEYS_V1 = {
    'ID', '案件識別番号', '設備ID', '発生日時', '件名', '発生区分', '職種', '対象区分', '発生場所', '案件類型',
    '発見者・報告者', '現場確認を実施', '現場確認開始', '現場確認終了', '一次対応を実施', '一次対応開始',
    '一次対応終了', '状況概要', '影響範囲', '原因・考察', '応急処置', '恒久対策・今後の予定',
    'オーナー・テナント報告', '対応業者への連絡', '社内報告', '報告先・提出状況', '復旧・完了状況',
}
CASE_KEYS = {
    'ID', '案件識別番号', '設備ID', '発生日時', '件名', '発生区分', '職種', '対象区分', '発生場所', '案件類型',
    '発見者・報告者', '現場確認を実施', '現場確認開始', '現場確認終了', '一次対応を実施', '一次対応開始',
    '一次対応終了', '状況概要', '影響範囲', '原因・考察', '応急処置', '恒久対策・今後の予定',
    'オーナー・テナント報告', '対応業者への連絡', '社内報告', '報告先・提出状況', '復旧・完了状況',
    # 2.5.0 (hb-report-package/2): fields of the company form
    '報告の段階', '部署', 'ビル名', '報告者', '報告者の役職', '当事者', '当事者の区分', '連絡先',
    '現場確認の実施者', '現場確認の内容', '一次対応の実施者', '一次対応の内容',
    '処置報告の区分', '処置の実施者', '処置の確認者',
}
CONTACT_KEYS = {'区分', '手段', '相手', '所属'}
MAX_CONTACTS = 20
PHOTO_KEYS = {'ID', '写真番号', '撮影日時', '撮影日時の根拠', '登録日時', '撮影場所', '元写真', '写真右側の説明文',
              '注釈・マーキングの指示', '種別', 'sha256'}
TOP_KEYS = {'schema', 'created_at', 'app', 'case', 'equipment', 'photos', 'timeline', 'contacts'}
EQUIPMENT_KEYS = {'id', 'name', 'category', 'location'}
APP_KEYS = {'name', 'version'}
MAX_PHOTOS_IN_PACKAGE = 20
MAX_TIMELINE_IN_PACKAGE = 50
SHA_RE = re.compile(r'^[0-9a-f]{64}$')
EQ_RE = re.compile(r'^EQ-\d{6}$')
# 案件ID_設備ID_連番_種別.拡張子  or  PHYYYYMMDD_設備ID_連番_種別.拡張子 (photo taken outside a case)
PHOTO_NAME_RE = re.compile(r'^(?:CASE-\d{8}-\d{4}|PH\d{8})_(EQ-\d{6})_\d{3}_[^_/\\]{1,20}\.(?:jpe?g|png|webp)$', re.I)
TIMELINE_KEYS = {'ID', '対応日時', '担当', '実施内容・確認結果', '測定数値など', '連絡先・連絡結果'}
ID_RE = re.compile(r'^[A-Za-z0-9][A-Za-z0-9_.-]{0,63}$')
PHOTO_MEMBER_RE = re.compile(r'^photos/[^/\\]{1,120}\.(?:jpe?g|png|webp)$', re.I)


def _stop(message):
    raise ValueError('HOMEBOX STOP: ' + message)


def _scalar_dict(obj, allowed, label):
    if not isinstance(obj, dict):
        _stop(label + ' の形式が違います')
    unknown = sorted(set(obj) - allowed)
    if unknown:
        _stop(label + ' に知らない項目があります: ' + ', '.join(unknown[:5]))
    for k, v in obj.items():
        if v is not None and not isinstance(v, (str, int, float, bool)):
            _stop(f'{label} の {k} の値の型が違います')
        if isinstance(v, str) and len(v) > MAX_TEXT_LENGTH:
            _stop(f'{label} の {k} が長すぎます')
    return dict(obj)


def read_package(path):
    """Return (case, photo_rows, timeline, info). Raises ValueError('HOMEBOX STOP: ...')."""
    with open(path, 'rb') as fh:
        raw = fh.read()
    if len(raw) > MAX_PACKAGE:
        _stop('パッケージが大きすぎます')
    with zipfile.ZipFile(path) as z:
        zip_guard(z)
        names = set(z.namelist())
        if 'package.json' not in names:
            _stop('package.json がありません（HomeBoxの報告書パッケージではありません）')
        try:
            data = json.loads(z.read('package.json').decode('utf-8'))
        except Exception:
            _stop('package.json を読めません')
        if not isinstance(data, dict):
            _stop('package.json の形式が違います')
        if data.get('schema') not in ACCEPTED_SCHEMAS:
            _stop('対応していない形式です: ' + str(data.get('schema')))
        unknown = sorted(set(data) - TOP_KEYS)
        if unknown:
            _stop('package.json に知らない項目があります: ' + ', '.join(unknown[:5]))
        _scalar_dict(data.get('app') or {}, APP_KEYS, 'アプリ情報')
        case = _scalar_dict(data.get('case'), CASE_KEYS, '案件')
        if not ID_RE.match(str(case.get('ID') or '')):
            _stop('案件IDが不正です')
        eq = _scalar_dict(data.get('equipment') or {}, EQUIPMENT_KEYS, '設備情報')
        case_eq = case.get('設備ID')
        if case_eq is not None and not EQ_RE.match(str(case_eq)):
            _stop('案件の設備IDの形式が違います')
        if (case_eq or eq.get('id')) and eq.get('id') != case_eq:
            _stop('案件の設備IDと設備情報が一致しません')
        raw_photos = data.get('photos') or []
        raw_timeline = data.get('timeline') or []
        if not isinstance(raw_photos, list) or not isinstance(raw_timeline, list):
            _stop('写真・対応経過の形式が違います')
        raw_contacts = data.get('contacts') or []
        if data.get('schema') == 'hb-report-package/1' and ('contacts' in data or set(case) - CASE_KEYS_V1):
            _stop('hb-report-package/1 に /2 の項目が入っています')
        if not isinstance(raw_contacts, list) or len(raw_contacts) > MAX_CONTACTS:
            _stop('通報・連絡の形式または件数が違います')
        case['_contacts'] = [_scalar_dict(c, CONTACT_KEYS, f'連絡{i}') for i, c in enumerate(raw_contacts, 1)]
        if len(raw_photos) > MAX_PHOTOS_IN_PACKAGE or len(raw_timeline) > MAX_TIMELINE_IN_PACKAGE:
            _stop('写真または対応経過の件数が多すぎます')
        if raw_photos and not case_eq:
            _stop('設備が決まっていない案件に写真が入っています（写真は必ず設備に結び付けます）')
        photos = []
        seen_ids, seen_names = set(), set()
        members_used = set()
        for i, p in enumerate(raw_photos, 1):
            p = _scalar_dict(p, PHOTO_KEYS, f'写真{i}')
            pid = str(p.get('ID') or '')
            if not ID_RE.match(pid) or pid in seen_ids:
                _stop(f'写真{i}のIDが不正または重複しています')
            seen_ids.add(pid)
            member = str(p.get('元写真') or '')
            if not PHOTO_MEMBER_RE.match(member) or '..' in member:
                _stop(f'写真{i}のファイル名が不正です')
            if member not in names:
                _stop(f'写真{i}の画像がパッケージにありません: {posixpath.basename(member)}')
            base = posixpath.basename(member)
            m = PHOTO_NAME_RE.match(base)
            if not m:
                _stop(f'写真{i}のファイル名が命名規則に合いません: {base}')
            if m.group(1) != case_eq:
                _stop(f'写真{i}は別の設備（{m.group(1)}）の写真です。案件の設備は {case_eq}')
            if not SHA_RE.match(str(p.get('sha256') or '')):
                _stop(f'写真{i}にSHA-256（64桁）がありません')
            if base in seen_names:
                _stop('同じ名前の写真が重複しています: ' + base)
            seen_names.add(base)
            data_bytes = z.read(member)
            digest = hashlib.sha256(data_bytes).hexdigest()
            if p['sha256'] != digest:
                _stop(f'写真{i}の内容がHomeBoxで記録した時と違います（{base}）')
            members_used.add(member)
            row = dict(p)
            row['元写真'] = base
            row['異常案件_ID'] = case['ID']
            row['createAt'] = p.get('撮影日時') or ''
            row.pop('撮影日時の根拠', None)
            row.pop('登録日時', None)
            photos.append(row)
        extra = sorted(n for n in names if n.startswith('photos/') and n not in members_used)
        if extra:
            _stop('案件に使われていない写真がパッケージに入っています: ' + ', '.join(posixpath.basename(n) for n in extra[:3]))
        unexpected = sorted(n for n in names if n != 'package.json' and not n.startswith('photos/'))
        if unexpected:
            _stop('想定外のファイルがパッケージに入っています: ' + ', '.join(unexpected[:3]))
        timeline = []
        for i, t in enumerate(raw_timeline, 1):
            t = _scalar_dict(t, TIMELINE_KEYS, f'対応経過{i}')
            t['異常案件_ID'] = case['ID']
            timeline.append(t)
    info = {
        'kind': 'homebox-package',
        'schema': data.get('schema'),
        'package_sha256': hashlib.sha256(raw).hexdigest(),
        'app': {k: str(v) for k, v in (data.get('app') or {}).items() if k in ('name', 'version')},
        'created_at': str(data.get('created_at') or ''),
        'case_id': case['ID'],
        'equipment_id': case.get('設備ID') or None,
        'equipment_name': (eq.get('name') if isinstance(eq, dict) else None),
    }
    return case, photos, timeline, info
