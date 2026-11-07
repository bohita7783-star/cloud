#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Which company form a template is, where its boxes are, and how photos go in.

2.6.0. Completed reports show that the company form exists in several versions
(the photo sheet's 3rd caption column moved by one row between 2026-10-17 and
2026-10-31; older copies have more rows, other columns, or photos placed IN the
cells). Writing fixed cell addresses into the wrong version would put text into
hidden cells without any error. So before anything is written:

  * the report sheet must have exactly the boxes (merged ranges) the layout
    writes into (REPORT_MERGES, taken from the 2026-10-17 / 2026-10-31 form);
  * the photo sheet's slots are DERIVED from its merged ranges (picture area,
    No. box, caption lines), not assumed;
  * how photos can be placed is decided from what the photo sheet contains:
      baseline  the vetted 2026-10-17 report used as template (3 pictures and
                3 marking shapes; photos are swapped, markings removed)
      blank     a blank form whose photo slots are empty: new pictures are
                placed over the slots (floating pictures, as in 2026-10-17)
      refused   anything else, e.g. pictures placed in cells (Excel "place in
                cell", richData) or pictures/shapes left in the slots.

The result (plan) is deterministic for a given file, so the inspector, the
approval check and the converter all derive the same plan from the same bytes.
Never reports cell values.
"""
import posixpath
import re

from lxml import etree as ET

SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR = 'http://schemas.openxmlformats.org/package/2006/relationships'
XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
NS = {'m': SS, 'xdr': XDR, 'a': A}
DRAWING_TYPE = REL + '/drawing'

PROFILE_ID = 'genba-ijou-sokuho-2026-10'
PROFILE_LABEL = '現場異常速報（2026年10月の配置）'
SHEET_REPORT = 'xl/worksheets/sheet1.xml'
SHEET_PHOTO = 'xl/worksheets/sheet2.xml'
SHEET_PHOTO_RELS = 'xl/worksheets/_rels/sheet2.xml.rels'

# Every cell the report layout writes, and the box (merged range) it heads in the
# 2026-10-17 and 2026-10-31 forms (identical: 359 merged ranges). A template whose
# boxes differ is a different version of the form and is refused.
REPORT_MERGES = {
    'AU1': 'AU1:BI1', 'AB2': 'AB2:AJ2', 'AP2': 'AP2:BI2', 'AP3': 'AP3:BI3', 'G3': 'G3:R3', 'W3': 'W3:AH3',
    'G4': 'G4:V4', 'AP4': 'AP4:BA4', 'BC4': 'BC4:BH4', 'G5': 'G5:S5', 'X5': 'X5:Z5', 'AC5': 'AC5:AD5',
    'AN5': 'AN5:BI5', 'G6': 'G6:AI6', 'AP6': 'AP6:BI6', 'G7': 'G7:R7', 'S7': 'S7:AB7', 'AG7': 'AG7:BI7',
    'G8': 'G8:K8', 'L8': 'L8:P8', 'S8': 'S8:W8', 'AC8': 'AC8:AL8', 'AP8': 'AP8:BI8',
    'G9': 'G9:K9', 'L9': 'L9:P9', 'S9': 'S9:W9', 'AC9': 'AC9:AL9', 'AP9': 'AP9:BI9',
    'G10': 'G10:J10', 'K10': 'K10:T10', 'U10': 'U10:X10', 'Y10': 'Y10:AG10', 'AH10': 'AH10:AK10',
    'AL10': 'AL10:AS10', 'AU10': 'AU10:BH10',
    'G11': 'G11:J11', 'K11': 'K11:R11', 'T11': 'T11:AF11', 'AH11': 'AH11:AK11', 'AL11': 'AL11:AS11', 'AU11': 'AU11:BH11',
    'G12': 'G12:J12', 'K12': 'K12:R12', 'T12': 'T12:AF12', 'AH12': 'AH12:AK12', 'AL12': 'AL12:AS12', 'AU12': 'AU12:BH12',
    'G13': 'G13:J13', 'K13': 'K13:AG13', 'AH13': 'AH13:AK13', 'AL13': 'AL13:BI13',
    'G14': 'G14:J14', 'K14': 'K14:P14', 'Q14': 'Q14:S14', 'U14': 'U14:AE14', 'AH14': 'AH14:AK14',
    'AL14': 'AL14:AQ14', 'AR14': 'AR14:AT14', 'AV14': 'AV14:BH14',
    'C28': 'C28:BI28', 'C29': 'C29:BI29', 'C30': 'C30:BI30', 'C31': 'C31:BI31',
    'G32': 'G32:M32', 'N32': 'N32:AM32', 'AR32': 'AR32:BI32', 'G33': 'G33:AM33', 'G34': 'G34:AM34',
    'AR34': 'AR34:BI34', 'G35': 'G35:AM35',
}
REPORT_MERGES.update({f'C{r}': f'C{r}:F{r}' for r in range(15, 28)})
REPORT_MERGES.update({f'G{r}': f'G{r}:BI{r}' for r in range(15, 28)})

# Photo sheet of this form family: title B1:O2, date P3:S3; each slot is a picture
# area B..L (at least 12 rows), a "No." box N:O and the number box P:S on its first
# two rows, and caption lines N:S below them.
PHOTO_TITLE = ('B1', 'B1:O2')
PHOTO_DATE = ('P3', 'P3:S3')
MIN_SLOT_ROWS = 12
MIN_CAPTION_LINES = 3
MAX_SLOTS = 9


class FormError(ValueError):
    pass


def _xml(data):
    return ET.fromstring(data, parser=ET.XMLParser(resolve_entities=False, no_network=True, huge_tree=False))


def col_index(letters):
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref):
    m = re.fullmatch(r'([A-Z]+)([0-9]+)', ref)
    if not m:
        raise FormError('不正なセル番地: ' + ref)
    return col_index(m.group(1)), int(m.group(2))


def parse_range(ref):
    a, _, b = ref.partition(':')
    c1, r1 = split_ref(a)
    c2, r2 = split_ref(b or a)
    return c1, r1, c2, r2


def merges_of(sheet_root):
    return [m.get('ref') for m in sheet_root.findall('m:mergeCells/m:mergeCell', NS)]


def resolve(base_part, target):
    if target.startswith('/'):
        return target.lstrip('/')
    return posixpath.normpath(posixpath.join(posixpath.dirname(base_part), target))


def rels_name(part):
    return posixpath.join(posixpath.dirname(part), '_rels', posixpath.basename(part) + '.rels')


def check_report(sheet_root):
    """Problems (empty list = this is the expected report form)."""
    by_head = {}
    for ref in merges_of(sheet_root):
        by_head.setdefault(ref.split(':')[0], ref)
    wrong = sorted((c for c, ref in REPORT_MERGES.items() if by_head.get(c) != ref),
                   key=lambda a: (split_ref(a)[1], split_ref(a)[0]))
    if wrong:
        return [f'報告書シートの枠の位置が対応している様式と違います（{len(wrong)}か所、例：{"、".join(wrong[:4])}）。'
                f'{PROFILE_LABEL}の空様式を使ってください']
    return []


def derive_photo_slots(sheet_root):
    """Photo slots from the merged ranges of the photo sheet (no cell values used)."""
    merges = [parse_range(r) + (r,) for r in merges_of(sheet_root)]
    have = {r for *_, r in merges}
    problems = []
    for cell, ref in (PHOTO_TITLE, PHOTO_DATE):
        if ref not in have:
            problems.append(f'写真シートの{"題名" if cell == "B1" else "日付"}の枠（{ref}）がありません')
    B, L, N, O, P, S = (col_index(x) for x in 'B L N O P S'.split())
    slots = []
    areas = 0
    for c1, r1, c2, r2, ref in sorted(merges, key=lambda m: m[1]):
        if (c1, c2) != (B, L) or r2 - r1 + 1 < MIN_SLOT_ROWS:
            continue
        areas += 1
        label = f'N{r1}:O{r1 + 1}'
        number = f'P{r1}:S{r1 + 1}'
        captions = sorted((m for m in merges if m[0] == N and m[2] == S and r1 + 2 <= m[1] and m[3] <= r2),
                          key=lambda m: m[1])
        if label not in have or number not in have or len(captions) < MIN_CAPTION_LINES:
            problems.append(f'写真枠 {ref} の番号欄・説明欄の形が対応している様式と違います')
            continue
        slots.append({'picture': ref, 'number': f'P{r1}', 'captions': [f'N{m[1]}' for m in captions],
                      'rows': [r1, r2], 'cols': [c1, c2]})
    if not areas:
        problems.append('写真シートに写真枠（B〜L列の結合範囲）が見つかりません')
    if len(slots) > MAX_SLOTS:
        problems.append('写真枠が多すぎます')
    return slots, problems


def _anchor_rows(anchor):
    frm = anchor.find('xdr:from', NS)
    if frm is None:  # absoluteAnchor: position not tied to rows; treat as inside to be safe
        return (0, 10 ** 6)
    row = int(frm.findtext('xdr:row', '0', NS))
    to = anchor.find('xdr:to', NS)
    return row + 1, (int(to.findtext('xdr:row', '0', NS)) + 1 if to is not None else row + 1)


def photo_plan(names, read):
    """Decide how photos are placed. `names` is the set of part names, `read(name)` returns bytes."""
    problems = []
    if SHEET_PHOTO not in names:
        raise FormError('写真シート（2枚目）がありません')
    sheet = _xml(read(SHEET_PHOTO))
    slots, slot_problems = derive_photo_slots(sheet)
    problems += slot_problems
    in_cell = [c.get('r') for c in sheet.findall('m:sheetData/m:row/m:c', NS) if c.get('vm')]
    if in_cell:
        problems.append(f'写真シートのセルの中に画像（セル内画像）が{len(in_cell)}個あります（{"、".join(in_cell[:4])}）。'
                        '前の案件の写真です。写真を削除した空様式を承認してください')
    drawing = None
    if SHEET_PHOTO_RELS in names:
        for rel in _xml(read(SHEET_PHOTO_RELS)):
            if rel.get('Type') == DRAWING_TYPE:
                drawing = resolve(SHEET_PHOTO, rel.get('Target', ''))
    plan = {'profile_id': PROFILE_ID, 'label': PROFILE_LABEL, 'slots': slots,
            'title': PHOTO_TITLE[0], 'date': PHOTO_DATE[0],
            'photo_drawing': drawing, 'photo_drawing_rels': None, 'photo_sheet_rels': SHEET_PHOTO_RELS in names,
            'mode': None, 'problems': problems}
    if drawing is not None:
        if drawing not in names:
            raise FormError('写真シートの図形の部品がありません')
        root = _xml(read(drawing))
        drels = rels_name(drawing)
        plan['photo_drawing_rels'] = drels if drels in names else None
        anchors = list(root)
        pics = [a for a in anchors if a.find('xdr:pic', NS) is not None]
        shapes = [a for a in anchors if a.find('xdr:pic', NS) is None]
        if _is_baseline(names, read, drawing, drels, pics, shapes):
            plan['mode'] = 'baseline'
            return plan
        if pics:
            problems.append(f'写真シートに写真（図）が{len(pics)}枚残っています。前の案件の写真です。削除した空様式を承認してください')
        inside = []
        for a in shapes:
            rows = _anchor_rows(a)
            if any(rows[0] <= s['rows'][1] and rows[1] >= s['rows'][0] for s in slots):
                inside.append(a)
        if inside:
            problems.append(f'写真枠の上に図形（赤丸・矢印など）が{len(inside)}個残っています。削除した空様式を承認してください')
    plan['mode'] = 'refused' if problems else 'blank'
    return plan


def _is_baseline(names, read, drawing, drels, pics, shapes):
    """The vetted 2026-10-17 report: 3 photos (image1..3.jpeg via rId1..3) and 3 marking shapes."""
    if drawing != 'xl/drawings/drawing2.xml' or drels not in names or len(pics) != 3 or len(shapes) != 3:
        return False
    targets = {r.get('Id'): posixpath.basename(r.get('Target', '')) for r in _xml(read(drels))}
    if {targets.get(f'rId{i}') for i in (1, 2, 3)} != {f'image{i}.jpeg' for i in (1, 2, 3)}:
        return False
    return all(f'xl/media/image{i}.jpeg' in names for i in (1, 2, 3))


def analyze(names, read):
    """Full form plan: report check + photo plan. Raises nothing for a mismatch; lists problems."""
    names = set(names)
    problems = []
    if SHEET_REPORT not in names:
        problems.append('報告書シート（1枚目）がありません')
    else:
        problems += check_report(_xml(read(SHEET_REPORT)))
    plan = photo_plan(names, read)
    plan['problems'] = problems + plan['problems']
    if plan['problems']:
        plan['mode'] = 'refused'
    return plan


def analyze_zip(z):
    return analyze(z.namelist(), z.read)


def summary(plan):
    """What the approval signs and what the app shows (no values)."""
    return {'profile_id': plan['profile_id'], 'photo_mode': plan['mode'],
            'photo_slots': [{'picture': s['picture'], 'number': s['number'], 'captions': s['captions']} for s in plan['slots']],
            'photo_drawing': plan['photo_drawing'], 'photo_drawing_rels': plan['photo_drawing_rels'],
            'photo_sheet_rels': plan['photo_sheet_rels']}


MODE_LABELS = {'baseline': '旧方式（検証済みの2026-10-17記入済み報告書を原本にする）',
               'blank': '空様式の写真枠に写真を貼る',
               'refused': '使えない'}


# --- sizes on the photo sheet, for placing a picture over a slot ---

EMU_PER_PX = 9525
EMU_PER_PT = 12700


def max_digit_width(styles_xml):
    """Approximate pixel width of '0' in the workbook's default font (Excel's column unit)."""
    try:
        font = _xml(styles_xml).find('m:fonts/m:font', NS)
        size = float(font.find('m:sz', NS).get('val'))
        name = font.find('m:name', NS).get('val')
    except Exception:
        return 7
    base = 8 if re.search(r'[　-鿿＀-￯]', name or '') else 7
    return max(5, round(base * size / 11))


def sheet_geometry(sheet_root, mdw):
    fmt = sheet_root.find('m:sheetFormatPr', NS)
    default_h = float(fmt.get('defaultRowHeight', '15')) if fmt is not None else 15.0
    default_w = None
    if fmt is not None and fmt.get('defaultColWidth'):
        default_w = float(fmt.get('defaultColWidth'))
    widths = {}
    for col in sheet_root.findall('m:cols/m:col', NS):
        for i in range(int(col.get('min')), int(col.get('max')) + 1):
            widths[i] = float(col.get('width', '8.43'))
    heights = {}
    for row in sheet_root.findall('m:sheetData/m:row', NS):
        if row.get('ht'):
            heights[int(row.get('r'))] = float(row.get('ht'))

    def col_px(i):
        w = widths.get(i, default_w if default_w is not None else 8.43)
        return int(((256 * w + int(128 / mdw)) / 256) * mdw)

    def row_emu(r):
        return round(heights.get(r, default_h) * EMU_PER_PT)
    return col_px, row_emu


def slot_box(slot, col_px, row_emu):
    """(width EMU, height EMU) of a slot's picture area."""
    c1, c2 = slot['cols']
    r1, r2 = slot['rows']
    w = sum(col_px(i) for i in range(c1, c2 + 1)) * EMU_PER_PX
    h = sum(row_emu(r) for r in range(r1, r2 + 1))
    return w, h


def place(slot, col_px, row_emu, img_w, img_h, margin=0.04):
    """Anchor (col0, colOff, row0, rowOff, cx, cy) that fits the picture into the slot, centred."""
    bw, bh = slot_box(slot, col_px, row_emu)
    scale = min(bw * (1 - margin) / img_w, bh * (1 - margin) / img_h)
    cx, cy = int(img_w * scale), int(img_h * scale)
    x, y = (bw - cx) // 2, (bh - cy) // 2
    col = slot['cols'][0]
    while x >= col_px(col) * EMU_PER_PX and col < slot['cols'][1]:
        x -= col_px(col) * EMU_PER_PX
        col += 1
    row = slot['rows'][0]
    while y >= row_emu(row) and row < slot['rows'][1]:
        y -= row_emu(row)
        row += 1
    return col - 1, int(x), row - 1, int(y), cx, cy


# Slots of the 2026-10-31 form, used only where no template is at hand (the app's
# "様式に入るか確認" preview). Conversion always uses the slots derived from the template.
DEFAULT_SLOTS = [
    {'picture': 'B5:L21', 'number': 'P5', 'captions': ['N7', 'N10', 'N13', 'N16', 'N19'], 'rows': [5, 21], 'cols': [2, 12]},
    {'picture': 'B24:L40', 'number': 'P24', 'captions': ['N26', 'N29', 'N32', 'N35', 'N38'], 'rows': [24, 40], 'cols': [2, 12]},
    {'picture': 'B43:L58', 'number': 'P43', 'captions': ['N45', 'N48', 'N51', 'N54', 'N57'], 'rows': [43, 58], 'cols': [2, 12]},
]
