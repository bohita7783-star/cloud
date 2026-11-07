#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Which parts of the company workbook the conversion engine may touch.

This file is the single source of truth shared by:
  * converter_mobile.transform_package (what the engine actually changes),
  * template_inspector (how parts are classified for the approval candidate),
  * approval.verify_manifest (an approval must agree with this policy).

If the engine starts touching another part, this policy and its version must
change, which invalidates every existing approval (re-approval required).
Standard library only, so the inspector can run on an admin PC without Pillow.
"""
import re

ENGINE_POLICY_VERSION = 'report-2026-10-17/engine-2.5.0'
LAYOUT_ID = 'genba-ijou-sokuho-2026-10-17'

SHEET_REPORT = 'xl/worksheets/sheet1.xml'
SHEET_PHOTO = 'xl/worksheets/sheet2.xml'
PHOTO_PARTS = tuple(f'xl/media/image{i}.jpeg' for i in (1, 2, 3))

CATEGORY_LABELS = {
    'immutable': '変更禁止',
    'modifiable': '指定範囲のみ変更可',
    'delete': '削除予定',
    'generate': '生成予定',
}

# Parts the engine rewrites. Value: (purpose, what may change).
MODIFIABLE_PARTS = {
    SHEET_REPORT: ('報告書シート', '入力欄として指定されたセルの値のみ。数式の計算済みの値は消去'),
    SHEET_PHOTO: ('写真帳シート', '写真番号・場所・説明・加工指示のセルの値のみ。数式の計算済みの値は消去'),
    'xl/sharedStrings.xml': ('共有文字列', '参照されなくなった文字列の中身を空にする（番号は変えない）'),
    'xl/drawings/drawing2.xml': ('写真帳の図形配置', '旧加工指示の図形と未使用の写真枠を外す'),
    'xl/drawings/_rels/drawing2.xml.rels': ('写真帳の図形の参照', '未使用の写真枠への参照を外す'),
    'xl/workbook.xml': ('ブック設定', '再計算の指定（calcPr）のみ'),
    'xl/_rels/workbook.xml.rels': ('ブックの参照', '計算順序（calcChain）への参照を外す。外部リンク等は保持'),
    '[Content_Types].xml': ('部品の種類一覧', '計算順序（calcChain）の登録を外す'),
    'docProps/core.xml': ('文書情報', '作成者・更新者・題名などを空にする'),
    'docProps/app.xml': ('アプリ情報', '会社名・管理者を空にする'),
}
for _p in PHOTO_PARTS:
    MODIFIABLE_PARTS[_p] = ('写真枠の画像', '新しい写真に差し替え。使わない枠は削除')

DELETE_PARTS = {
    'xl/calcChain.xml': ('計算順序', '数式を再計算させるため削除'),
}
# The engine never adds parts in this version.
GENERATE_PARTS = {}

# Parts that must exist for the engine to work at all.
REQUIRED_PARTS = (SHEET_REPORT, SHEET_PHOTO, 'xl/sharedStrings.xml', 'xl/workbook.xml',
                  'xl/drawings/drawing2.xml', 'xl/drawings/_rels/drawing2.xml.rels') + PHOTO_PARTS

# Risk categories (same rules as the v2.3.2 stop gate). Used for reporting;
# for an UNAPPROVED workbook the converter still stops on them.
RISKY_PART_PREFIXES = (
    'customXml/', 'xl/externalLinks/', 'xl/embeddings/', 'xl/activeX/',
    'xl/persons/', 'xl/threadedComments/', 'xl/queryTables/',
    'xl/pivotCache/', 'xl/pivotTables/', 'xl/ctrlProps/',
    'xl/connections', 'xl/comments', 'xl/metadata.xml',
    'xl/worksheets/comments', 'xl/ink/', 'xl/model/',
)
RISKY_EXACT_PARTS = {'docProps/custom.xml', 'docProps/thumbnail.jpeg',
                     'docProps/thumbnail.png', 'docProps/thumbnail.wmf'}
ALLOWED_PHOTO_PARTS = set(PHOTO_PARTS)


def category_for(name):
    if name in MODIFIABLE_PARTS:
        return 'modifiable'
    if name in DELETE_PARTS:
        return 'delete'
    if name in GENERATE_PARTS:
        return 'generate'
    return 'immutable'


def purpose_for(name):
    for table in (MODIFIABLE_PARTS, DELETE_PARTS, GENERATE_PARTS):
        if name in table:
            return table[name][0]
    lower = name.lower()
    rules = (
        (r'^xl/worksheets/sheet\d+\.xml$', 'ワークシート（変換ツールは触らない）'),
        (r'^xl/worksheets/_rels/', 'ワークシートの参照'),
        (r'^xl/drawings/vmldrawing', 'VML図形（チェックボックス・コメント枠など）'),
        (r'^xl/drawings/drawing\d+\.xml$', '図形配置'),
        (r'^xl/drawings/_rels/', '図形の参照'),
        (r'^xl/media/.*\.emf$', 'EMF画像（ロゴ・図など）'),
        (r'^xl/media/', '画像'),
        (r'^xl/ctrlprops/', 'フォームコントロール（チェックボックス等）'),
        (r'^xl/externallinks/', '外部ブックへのリンク'),
        (r'^xl/comments', 'セルのコメント'),
        (r'^xl/threadedcomments/', 'スレッド形式のコメント'),
        (r'^xl/persons/', 'コメント投稿者の情報'),
        (r'^customxml/', '管理用XML（SharePoint等）'),
        (r'^docprops/custom\.xml$', '独自の文書情報'),
        (r'^docprops/thumbnail', '縮小表示画像'),
        (r'^xl/printersettings/', '印刷設定'),
        (r'^xl/styles\.xml$', '書式'),
        (r'^xl/theme/', 'テーマ'),
        (r'^xl/richdata/', 'セル内画像の情報'),
        (r'^xl/metadata\.xml$', 'セルのメタデータ'),
        (r'^xl/embeddings/', '埋め込みファイル'),
        (r'\.rels$', '参照関係'),
    )
    for pattern, label in rules:
        if re.search(pattern, lower):
            return label
    return 'その他'


def risk_flags(name):
    """Reasons a part needs a human look before approval (never values)."""
    flags = []
    lower = name.lower()
    if name in RISKY_EXACT_PARTS or any(name.startswith(x) for x in RISKY_PART_PREFIXES):
        flags.append('要確認の内部部品')
    if name.startswith('xl/media/') and name not in ALLOWED_PHOTO_PARTS:
        flags.append('写真枠以外の画像・EMF')
    if lower.endswith('.vml') or '/vml' in lower:
        flags.append('VML図形')
    if re.fullmatch(r'xl/worksheets/sheet\d+\.xml', name) and name not in (SHEET_REPORT, SHEET_PHOTO):
        flags.append('変換ツールが触らない追加シート')
    return flags

# Document-property fields blanked in the OUTPUT copy (never in the source).
DOCPROP_CORE_BLANK = {'creator', 'lastModifiedBy', 'title', 'subject', 'description',
                      'keywords', 'category', 'contentStatus', 'identifier'}
DOCPROP_APP_BLANK = {'Company', 'Manager'}
