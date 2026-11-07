"""Build a SYNTHETIC company-style template for tests and samples.

Everything here is invented. It imitates only the structure the converter
expects (report sheet, photo sheet with 3 photo slots) plus the "awkward"
parts real templates have: an EMF picture, VML, a checkbox (ctrlProps), an
auxiliary sheet, an external workbook link, calcChain and document properties.
The cleared cells contain clearly fake "旧案件_" text so residue checks have
something to look for.
"""
import io
import zipfile
from lxml import etree as ET
from PIL import Image

S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT = 'http://schemas.openxmlformats.org/package/2006/content-types'
X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
DC = 'http://purl.org/dc/elements/1.1/'
CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
EXT_PATH = 'file:///C:/架空フォルダ/架空設備マスタ.xlsx'
LEGACY = {
    'G6': '旧案件_架空の漏水件名', 'AP6': '旧案件_架空の地下二階', 'G15': '旧案件_状況の文章ABCDEF',
    'AP9': '旧案件_架空の担当者', 'C28': '旧案件_原因の文章GHIJKL',
}
LEGACY_PHOTO = {'B1': '旧案件_写真帳の題名', 'N10': '旧案件_写真の説明MNOPQR'}
STATIC = {'A1': '現場異常速報（架空様式）', 'A2': '記入欄は太枠内'}
AUX_TEXT = ['機械室', '電気室', '屋上', '受水槽室']
LEGACY_CREATOR = '旧作成者_架空'


def _x(root):
    return ET.tostring(root, encoding='UTF-8', xml_declaration=True, standalone=True)


def _jpeg(color, size=(120, 90)):
    b = io.BytesIO()
    Image.new('RGB', size, color).save(b, 'JPEG', quality=85)
    return b.getvalue()


def _sheet(cells, shared, extra=()):
    ws = ET.Element('{%s}worksheet' % S, nsmap={None: S, 'r': R})
    sd = ET.SubElement(ws, '{%s}sheetData' % S)
    rows = {}
    import re
    for addr in sorted(cells, key=lambda a: (int(re.sub(r'\D', '', a)), len(re.sub(r'\d', '', a)), re.sub(r'\d', '', a))):
        r = int(re.sub(r'\D', '', addr))
        row = rows.get(r)
        if row is None:
            row = rows[r] = ET.SubElement(sd, '{%s}row' % S, r=str(r))
        val = cells[addr]
        c = ET.SubElement(row, '{%s}c' % S, r=addr)
        if isinstance(val, tuple):  # formula with cached value
            ET.SubElement(c, '{%s}f' % S).text = val[0]
            ET.SubElement(c, '{%s}v' % S).text = str(val[1])
        elif isinstance(val, (int, float)):
            ET.SubElement(c, '{%s}v' % S).text = str(val)
        else:
            if val not in shared:
                shared.append(val)
            c.set('t', 's')
            ET.SubElement(c, '{%s}v' % S).text = str(shared.index(val))
    for tag, attrs in extra:
        ET.SubElement(ws, '{%s}%s' % (S, tag), attrs)
    return ws


def build_template_parts(variant=''):
    """Return {part name: bytes}. `variant` changes static text to make a different template."""
    shared = []
    report_cells = dict(STATIC)
    report_cells['A1'] = STATIC['A1'] + variant
    report_cells.update(LEGACY)
    report_cells['AU1'] = 46300
    report_cells['B60'] = ('SUM(1,2)', 3)
    sheet1 = _sheet(report_cells, shared)
    merge = ET.SubElement(sheet1, '{%s}mergeCells' % S, count='1')
    ET.SubElement(merge, '{%s}mergeCell' % S, ref='A1:F1')
    ET.SubElement(sheet1, '{%s}pageMargins' % S, left='0.7', right='0.7', top='0.75', bottom='0.75', header='0.3', footer='0.3')
    ET.SubElement(sheet1, '{%s}drawing' % S, {'{%s}id' % R: 'rId1'})
    ET.SubElement(sheet1, '{%s}legacyDrawing' % S, {'{%s}id' % R: 'rId2'})
    photo_cells = {'A1': '写真帳'}
    photo_cells.update(LEGACY_PHOTO)
    sheet2 = _sheet(photo_cells, shared)
    ET.SubElement(sheet2, '{%s}drawing' % S, {'{%s}id' % R: 'rId1'})
    sheet3 = _sheet({f'A{i}': t for i, t in enumerate(AUX_TEXT, 1)}, shared)

    sst = ET.Element('{%s}sst' % S, nsmap={None: S}, count=str(len(shared)), uniqueCount=str(len(shared)))
    for t in shared:
        ET.SubElement(ET.SubElement(sst, '{%s}si' % S), '{%s}t' % S).text = t

    wb = ET.Element('{%s}workbook' % S, nsmap={None: S, 'r': R})
    sheets = ET.SubElement(wb, '{%s}sheets' % S)
    for i, name in enumerate(('報告書', '写真', '補助リスト'), 1):
        ET.SubElement(sheets, '{%s}sheet' % S, {'name': name, 'sheetId': str(i), '{%s}id' % R: f'rId{i}'})
    ET.SubElement(wb, '{%s}calcPr' % S, calcId='191029')

    wbrels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    for i in (1, 2, 3):
        ET.SubElement(wbrels, '{%s}Relationship' % P, Id=f'rId{i}', Type=R + '/worksheet', Target=f'worksheets/sheet{i}.xml')
    ET.SubElement(wbrels, '{%s}Relationship' % P, Id='rId4', Type=R + '/sharedStrings', Target='sharedStrings.xml')
    ET.SubElement(wbrels, '{%s}Relationship' % P, Id='rId5', Type=R + '/calcChain', Target='calcChain.xml')
    ET.SubElement(wbrels, '{%s}Relationship' % P, Id='rId6', Type=R + '/externalLink', Target='externalLinks/externalLink1.xml')

    ext = ET.Element('{%s}externalLink' % S, nsmap={None: S, 'r': R})
    ET.SubElement(ext, '{%s}externalBook' % S, {'{%s}id' % R: 'rId1'})
    extrels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    ET.SubElement(extrels, '{%s}Relationship' % P, Id='rId1', Type=R + '/externalLinkPath', Target=EXT_PATH, TargetMode='External')

    s1rels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    ET.SubElement(s1rels, '{%s}Relationship' % P, Id='rId1', Type=R + '/drawing', Target='../drawings/drawing1.xml')
    ET.SubElement(s1rels, '{%s}Relationship' % P, Id='rId2', Type=R + '/vmlDrawing', Target='../drawings/vmlDrawing1.vml')
    ET.SubElement(s1rels, '{%s}Relationship' % P, Id='rId3', Type=R + '/ctrlProp', Target='../ctrlProps/ctrlProp1.xml')
    s2rels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    ET.SubElement(s2rels, '{%s}Relationship' % P, Id='rId1', Type=R + '/drawing', Target='../drawings/drawing2.xml')

    def drawing(pics, shapes):
        dr = ET.Element('{%s}wsDr' % X, nsmap={'xdr': X, 'a': A, 'r': R})
        for rid in pics:
            anchor = ET.SubElement(dr, '{%s}twoCellAnchor' % X)
            pic = ET.SubElement(anchor, '{%s}pic' % X)
            ET.SubElement(ET.SubElement(pic, '{%s}blipFill' % X), '{%s}blip' % A, {'{%s}embed' % R: rid})
            if shapes:
                mark = ET.SubElement(dr, '{%s}twoCellAnchor' % X)
                ET.SubElement(mark, '{%s}sp' % X)
        return dr
    d1 = drawing(['rId1'], False)
    d1rels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    ET.SubElement(d1rels, '{%s}Relationship' % P, Id='rId1', Type=R + '/image', Target='../media/image4.emf')
    d2 = drawing(['rId1', 'rId2', 'rId3'], True)
    d2rels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    for i in (1, 2, 3):
        ET.SubElement(d2rels, '{%s}Relationship' % P, Id=f'rId{i}', Type=R + '/image', Target=f'../media/image{i}.jpeg')

    calc = ET.Element('{%s}calcChain' % S, nsmap={None: S})
    ET.SubElement(calc, '{%s}c' % S, r='B60', i='1')

    core = ET.Element('{%s}coreProperties' % CP, nsmap={'cp': CP, 'dc': DC})
    ET.SubElement(core, '{%s}creator' % DC).text = LEGACY_CREATOR
    ET.SubElement(core, '{%s}lastModifiedBy' % CP).text = '旧更新者_架空'
    app = ET.Element('{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Properties')
    ET.SubElement(app, '{http://schemas.openxmlformats.org/officeDocument/2006/extended-properties}Company').text = '架空ビル管理株式会社'

    types = ET.Element('{%s}Types' % CT, nsmap={None: CT})
    for ext_, ct in (('rels', 'application/vnd.openxmlformats-package.relationships+xml'), ('xml', 'application/xml'),
                     ('jpeg', 'image/jpeg'), ('emf', 'image/x-emf'), ('vml', 'application/vnd.openxmlformats-officedocument.vmlDrawing')):
        ET.SubElement(types, '{%s}Default' % CT, Extension=ext_, ContentType=ct)
    for part, ct in (('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
                     ('/xl/calcChain.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'),
                     ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml')):
        ET.SubElement(types, '{%s}Override' % CT, PartName=part, ContentType=ct)

    rootrels = ET.Element('{%s}Relationships' % P, nsmap={None: P})
    ET.SubElement(rootrels, '{%s}Relationship' % P, Id='rId1', Type=R + '/officeDocument', Target='xl/workbook.xml')

    return {
        '[Content_Types].xml': _x(types),
        '_rels/.rels': _x(rootrels),
        'xl/workbook.xml': _x(wb),
        'xl/_rels/workbook.xml.rels': _x(wbrels),
        'xl/worksheets/sheet1.xml': _x(sheet1),
        'xl/worksheets/sheet2.xml': _x(sheet2),
        'xl/worksheets/sheet3.xml': _x(sheet3),
        'xl/worksheets/_rels/sheet1.xml.rels': _x(s1rels),
        'xl/worksheets/_rels/sheet2.xml.rels': _x(s2rels),
        'xl/sharedStrings.xml': _x(sst),
        'xl/calcChain.xml': _x(calc),
        'xl/drawings/drawing1.xml': _x(d1),
        'xl/drawings/_rels/drawing1.xml.rels': _x(d1rels),
        'xl/drawings/drawing2.xml': _x(d2),
        'xl/drawings/_rels/drawing2.xml.rels': _x(d2rels),
        'xl/drawings/vmlDrawing1.vml': b'<xml xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="SYNTHETIC_CHECKBOX_SHAPE"/></xml>',
        'xl/ctrlProps/ctrlProp1.xml': b'<?xml version="1.0" encoding="UTF-8"?><formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" objectType="CheckBox" checked="Checked"/>',
        'xl/externalLinks/externalLink1.xml': _x(ext),
        'xl/externalLinks/_rels/externalLink1.xml.rels': _x(extrels),
        'xl/media/image1.jpeg': _jpeg('#cccccc'),
        'xl/media/image2.jpeg': _jpeg('#bbbbbb'),
        'xl/media/image3.jpeg': _jpeg('#aaaaaa'),
        'xl/media/image4.emf': b'\x01\x00\x00\x00SYNTHETIC_EMF_LOGO_' + bytes(range(64)),
        'docProps/core.xml': _x(core),
        'docProps/app.xml': _x(app),
    }


def write_zip(path, parts, stored=True):
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_STORED if stored else zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    return path
