#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""How one incident is written into the company form "現場異常報告" (layout 2026-10-17 / 2026-10-31).

Derived from the company's own sheets in the form (記入例・作成方法) and from how
completed reports are actually filled in:

* 対応・処置内容 (rows 15-27): chronological, ONE LINE PER ROW. Column C holds the
  time ("11：40", full-width colon) on the first line of each entry only; a
  different day is written "5/6" in C and the time goes at the start of the text.
  Long sentences continue on the next row. The last rows hold the plan
  ("修理を提案予定。"). Rules from 作成方法: 時系列で / 知らない人が読んでも分かるように / 5W2H.
* 原因 (rows 28-31) and 処置報告 (rows 32-35) are also one line per row.
* Dropdown cells must hold the form's exact list strings (e.g. "故　障　", "【ﾒｰﾙ】").
* Page 2 (事故報告書) is formulas that copy page 1; it is not written.

Width units: a full-width character counts 1, a half-width character 0.5. The
per-line capacities come from the merged cell widths of the form.

Standard library only (also runs inside Pyodide in the HomeBox app).
"""
import datetime
import re
import unicodedata

LAYOUT_VERSION = 'genba-ijou-sokuho-2026-10-17/layout-2.5.0'


class LayoutError(ValueError):
    """The case does not fit the form; nothing may be dropped silently."""


# ---- per-line capacities (full-width characters), from the form's column widths ----
BODY_WIDTH = 51            # G15:BI15 … G27:BI27
CAUSE_WIDTH = 55           # C28:BI28 … C31:BI31
ACTION_FIRST_WIDTH = 24    # N32:AM32 (after "応急処置：")
ACTION_NEXT_WIDTH = 31     # G33:AM33 … G35:AM35
BODY_ROWS = range(15, 28)  # 13 rows: chronology, then the plan at the bottom
CAUSE_ROWS = range(28, 32)
ACTION_ROWS = ('N32', 'G33', 'G34', 'G35')

# Single-line fields: (cell, label, capacity). Too long -> a warning (the text
# is kept in the cell; Excel may show it clipped).
SHORT_FIELDS = {
    'G6': ('件名', 27), 'AP6': ('発生場所', 18), 'AP2': ('部署', 18), 'AP3': ('ビル名', 18),
    'AP4': ('報告者', 11), 'AN5': ('発見・通報者', 20), 'G7': ('当事者', 11), 'AG7': ('連絡先', 27),
    'AC8': ('現場確認の実施者', 9), 'AP8': ('現場確認の内容', 18),
    'AC9': ('一次対応の実施者', 9), 'AP9': ('一次対応の内容', 18),
    'AR32': ('処置の実施者', 16), 'AR34': ('処置の確認者', 16),
}
PHOTO_TITLE_WIDTH = 37      # 写真 sheet B1:O2
PHOTO_CAPTION_WIDTH = 16    # N7:S9 etc.

# ---- the form's dropdown strings (hidden columns BN-BR of the form) ----
STAGES = ('（第１報）', '（第２報）', '（第３報）', '（最終）', '(事故報告)')
CATEGORY = {'事件': '事　件', '事故': '事　故', '故障': '故　障　', 'クレーム': 'クレーム', '苦情': 'クレーム'}
TRADE = {'設備': '設　備', '清掃': '清　掃', '警備': '警　備', '駐車場': '駐車場', '客室': '客　室',
         '管理人': '管理人', '受付': '受　付', '事務': '事　務'}
CONTENT = ('火災', '地震', '風水害', '停電', '漏水', '排水詰り', '機器の故障', '施設の損傷')
PARTY_KIND = ('【当社の社員】', '【当社のパート】', '【委託先従業員】', '【オーナー】', '【テナント従業員】',
              '【一般客】', '【周辺住民】', '【業者】')
ACTION_KIND = ('応急処置：', '恒久処置：', '観察中：')
MEANS_LABEL = {'メール': '【ﾒｰﾙ】', 'ﾒｰﾙ': '【ﾒｰﾙ】', '電話': '【電話】', '口頭': '【口頭】', 'FAX': '【FAX】', '訪問': '【訪問】'}
MEANS_DROPDOWN = {'メール': 'ﾒｰﾙ', 'ﾒｰﾙ': 'ﾒｰﾙ', '電話': '電話'}

# 通報・現状報告 (rows 10-14): (means cell, name cell, organisation cell, dropdown cell)
CONTACT_SLOTS = {
    '関係機関': [('G10', 'K10', None, None), ('U10', 'Y10', None, None), ('AH10', 'AL10', 'AU10', None)],
    # completed reports fill the left column top to bottom first (row 11, row 12), then the right
    'オーナー・テナント': [('G11', 'K11', 'T11', None), ('G12', 'K12', 'T12', None),
                      ('AH11', 'AL11', 'AU11', None), ('AH12', 'AL12', 'AU12', None)],
    '対応業者': [('G13', 'K13', None, None), ('AH13', 'AL13', None, None)],
    '社内': [('G14', 'K14', 'U14', 'Q14'), ('AH14', 'AL14', 'AV14', 'AR14')],
}
CONTACT_KINDS = tuple(CONTACT_SLOTS)

# Every cell this layout may write on the report sheet (the engine clears/permits these).
REPORT_CELLS = (
    {'AU1', 'AB2', 'AP2', 'AP3', 'G3', 'W3', 'G4', 'AP4', 'BC4', 'G5', 'X5', 'AC5', 'AN5', 'G6', 'AP6',
     'G7', 'S7', 'AG7', 'G8', 'L8', 'S8', 'AC8', 'AP8', 'G9', 'L9', 'S9', 'AC9', 'AP9',
     'G32', 'AR32', 'AR34'}
    | {f'{c}{r}' for r in BODY_ROWS for c in ('C', 'G')}
    | {f'C{r}' for r in CAUSE_ROWS} | set(ACTION_ROWS)
    | {c for slots in CONTACT_SLOTS.values() for slot in slots for c in slot if c}
)


def text_width(s):
    return sum(1 if unicodedata.east_asian_width(ch) in ('F', 'W', 'A') else 0.5 for ch in str(s))


_NO_LINE_START = set('、。，．）」』】〕〉》？！ー・：')


def wrap(text, width, first_width=None):
    """Greedy wrap by display width; keeps the user's own line breaks and indents.
    A line never starts with 、。）」 etc. (it stays on the previous line instead)."""
    lines = []
    for para in str(text or '').replace('\r\n', '\n').split('\n'):
        para = para.rstrip()
        if not para.strip():
            continue
        cur, w = '', 0.0
        for ch in para:
            limit = first_width if (first_width and not lines) else width
            cw = text_width(ch)
            if cur and w + cw > limit and ch not in _NO_LINE_START:
                lines.append(cur)
                cur, w = '', 0.0
            cur += ch
            w += cw
        if cur:
            lines.append(cur)
    return lines


def form_choice(table, value):
    v = str(value or '').strip()
    if not v:
        return None
    key = v.replace('　', '').replace(' ', '')
    if isinstance(table, dict):
        return table.get(key, v)
    return v


def time_label(dt):
    return f'{dt.hour}：{dt.minute:02d}'


def excel_serial_day(d):
    return (d - datetime.date(1899, 12, 30)).days


def _clean(v):
    return str(v).strip() if v not in (None, '') else ''


def layout_report(case, timeline, contacts, parse_dt, parse_bool, parse_hhmm, today=None):
    """Return (cells, warnings). cells = {address: value} for the report sheet.

    parse_dt/parse_bool/parse_hhmm are the engine's strict parsers (passed in so
    this module stays independent and the same rules apply everywhere)."""
    warnings = []
    incident = parse_dt(case['発生日時'], '発生日時')
    today = today or datetime.date.today()
    cells = {
        'AU1': excel_serial_day(today),                       # 発行日 = the day the report is made
        'AB2': case.get('報告の段階') if case.get('報告の段階') in STAGES else '（第１報）',
        'G3': form_choice(CATEGORY, case.get('発生区分')),
        'W3': form_choice(None, case.get('案件類型')),
        'G4': form_choice(TRADE, case.get('職種')),
        'G5': excel_serial_day(incident.date()), 'X5': incident.hour, 'AC5': incident.minute,
        'AN5': _clean(case.get('発見者・報告者')) or None,
        'G6': _clean(case.get('件名')), 'AP6': _clean(case.get('発生場所')),
    }
    for key, cell in (('部署', 'AP2'), ('ビル名', 'AP3'), ('報告者', 'AP4'), ('報告者の役職', 'BC4'),
                      ('当事者', 'G7'), ('当事者の区分', 'S7'), ('連絡先', 'AG7'),
                      ('現場確認の実施者', 'AC8'), ('現場確認の内容', 'AP8'),
                      ('一次対応の実施者', 'AC9'), ('一次対応の内容', 'AP9'),
                      ('処置の実施者', 'AR32'), ('処置の確認者', 'AR34')):
        if _clean(case.get(key)):
            cells[cell] = _clean(case.get(key))
    for flag, cell, times in (('現場確認を実施', 'G8', (('現場確認開始', 'L8'), ('現場確認終了', 'S8'))),
                              ('一次対応を実施', 'G9', (('一次対応開始', 'L9'), ('一次対応終了', 'S9')))):
        done = parse_bool(case.get(flag), flag)
        cells[cell] = '有' if done else '無'
        for label, tcell in times:
            if case.get(label) not in (None, ''):
                t = parse_hhmm(case[label])
                if not t or not re.fullmatch(r'\d{2}:\d{2}', t):
                    raise LayoutError('時刻の形式が違います: ' + label)
                h, m = map(int, t.split(':'))
                cells[tcell] = (h * 60 + m) / 1440

    # ---- 通報・現状報告 ----
    used = {k: 0 for k in CONTACT_SLOTS}
    for i, c in enumerate(contacts or [], 1):
        kind = c.get('区分')
        if kind not in CONTACT_SLOTS:
            raise LayoutError(f'連絡{i}の区分が違います: {kind}')
        n = used[kind]
        if n >= len(CONTACT_SLOTS[kind]):
            raise LayoutError(f'「{kind}」への連絡は様式に{len(CONTACT_SLOTS[kind])}件までです（{n + 1}件目が入りません）')
        used[kind] += 1
        means_cell, name_cell, org_cell, dd_cell = CONTACT_SLOTS[kind][n]
        means = _clean(c.get('手段'))
        name = _clean(c.get('相手'))
        org = _clean(c.get('所属'))
        if means:
            cells[means_cell] = MEANS_LABEL.get(means, f'【{means}】')
        if dd_cell and means in MEANS_DROPDOWN:
            cells[dd_cell] = MEANS_DROPDOWN[means]
        if kind == 'オーナー・テナント' and name and not name.endswith('様'):
            name += '　様'
        if org and not org_cell:
            name = f'{name}（{org}）' if name else org
        if name:
            cells[name_cell] = name
        if org and org_cell:
            cells[org_cell] = org

    # unused slots show the blank form's '：' placeholder (never a previous report's 【ﾒｰﾙ】)
    for kind, slots in CONTACT_SLOTS.items():
        for means_cell, *_ in slots[used[kind]:]:
            cells[means_cell] = '：'

    # ---- 対応・処置内容: chronology (rows 15..) + plan in rows 26-27 ----
    plan = wrap(case.get('恒久対策・今後の予定'), BODY_WIDTH)
    entries = []
    opening = _clean(case.get('状況概要'))
    if case.get('影響範囲'):
        opening += '\n影響範囲：' + _clean(case['影響範囲'])
    entries.append((incident, opening))

    def sort_key(rec):
        dt = rec.get('対応日時')
        return (1, datetime.datetime.max) if dt in (None, '') else (0, parse_dt(dt, '対応日時'))
    for rec in sorted(timeline or [], key=sort_key):
        text = _clean(rec.get('実施内容・確認結果'))
        if not text:
            raise LayoutError('対応経過の「実施内容・確認結果」が空です')
        if _clean(rec.get('測定数値など')):
            text += f"（{_clean(rec['測定数値など'])}）"
        who = _clean(rec.get('担当'))
        if who and who not in text:
            text += f'（担当：{who}）'
        if _clean(rec.get('連絡先・連絡結果')):
            text += '\n連絡：' + _clean(rec['連絡先・連絡結果'])
        dt = parse_dt(rec['対応日時'], '対応日時') if rec.get('対応日時') not in (None, '') else None
        entries.append((dt, text))

    body = []  # (C value or None, G text)
    for dt, text in entries:
        prefix = ''
        label = None
        if dt is not None:
            if dt.date() == incident.date():
                label = time_label(dt)
            else:
                label = f'{dt.month}/{dt.day}'
                prefix = f'{dt.hour}:{dt.minute:02d}　'
        lines = wrap(prefix + text, BODY_WIDTH)
        for j, line in enumerate(lines):
            body.append((label if j == 0 else None, line))
    capacity = len(BODY_ROWS) - (2 if plan else 0)
    if len(plan) > 2:
        raise LayoutError(f'「恒久対策・今後の予定」が{len(plan)}行になり、様式の2行に入りません（1行約{BODY_WIDTH}文字）')
    if len(body) > capacity:
        raise LayoutError(f'対応・処置内容が{len(body)}行になり、様式の{capacity}行に入りません'
                          f'（1行約{BODY_WIDTH}文字）。文章を短くするか、続報（第２報）に分けてください')
    rows = list(BODY_ROWS)
    for (label, line), r in zip(body, rows):
        if label:
            cells[f'C{r}'] = label
        cells[f'G{r}'] = line
    # completed reports start the plan at row 26 (row 27 only for a second line)
    for line, r in zip(plan, rows[len(rows) - 2:]):
        cells[f'G{r}'] = line

    # ---- 原因 ----
    cause = wrap(case.get('原因・考察'), CAUSE_WIDTH)
    if len(cause) > len(CAUSE_ROWS):
        raise LayoutError(f'「原因・考察」が{len(cause)}行になり、様式の{len(CAUSE_ROWS)}行に入りません（1行約{CAUSE_WIDTH}文字）')
    for line, r in zip(cause, CAUSE_ROWS):
        cells[f'C{r}'] = line

    # ---- 処置報告 ----
    action = wrap(case.get('応急処置'), ACTION_NEXT_WIDTH, first_width=ACTION_FIRST_WIDTH)
    if len(action) > len(ACTION_ROWS):
        raise LayoutError(f'「処置報告の内容」が{len(action)}行になり、様式の{len(ACTION_ROWS)}行に入りません'
                          f'（1行目約{ACTION_FIRST_WIDTH}文字、2行目以降約{ACTION_NEXT_WIDTH}文字）')
    if action:
        kind = case.get('処置報告の区分') if case.get('処置報告の区分') in ACTION_KIND else '応急処置：'
        cells['G32'] = kind
        for line, cell in zip(action, ACTION_ROWS):
            cells[cell] = line

    for cell, (label, cap) in SHORT_FIELDS.items():
        v = cells.get(cell)
        if isinstance(v, str) and text_width(v) > cap:
            warnings.append(f'「{label}」が長めです（約{cap}文字まで。今は約{int(text_width(v))}文字）。印刷で切れていないか確認してください')
    cells = {k: v for k, v in cells.items() if v not in (None, '')}
    assert set(cells) <= REPORT_CELLS, set(cells) - REPORT_CELLS
    return cells, warnings


def layout_photos(case, photos, issue_day, slots):
    """Photo sheet: title, date, and per slot № / 場所 / 対象 / 詳細 (説明の2行目 or 注釈)."""
    warnings = []
    title = _clean(case.get('件名'))
    cells = {'B1': title, 'P3': issue_day}
    if text_width(title) > PHOTO_TITLE_WIDTH:
        warnings.append(f'写真帳の題名（件名）が長めです（約{PHOTO_TITLE_WIDTH}文字まで）')
    for i, photo in enumerate(photos):
        _, num, loc, desc, mark, _ = slots[i]
        lines = [x.strip() for x in str(photo.get('写真右側の説明文') or '').replace('\r\n', '\n').split('\n') if x.strip()]
        note = _clean(photo.get('注釈・マーキングの指示'))
        if len(lines) > 2 or (len(lines) == 2 and note):
            raise LayoutError(f'写真{i + 1}の説明は「対象」「詳細」の2行まで（詳細の行は注釈と共用）です')
        cells[num] = i + 1
        cells[loc] = _clean(photo.get('撮影場所'))
        cells[desc] = lines[0] if lines else ''
        cells[mark] = lines[1] if len(lines) > 1 else note
        for cell in (loc, desc, mark):
            if isinstance(cells.get(cell), str) and text_width(cells[cell]) > PHOTO_CAPTION_WIDTH:
                warnings.append(f'写真{i + 1}の説明「{cells[cell][:8]}…」が長めです（1行約{PHOTO_CAPTION_WIDTH}文字まで）')
    return {k: v for k, v in cells.items() if v not in (None, '')}, warnings
