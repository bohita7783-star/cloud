"""Synthetic-only test suite for v2.3.2-security-preview.
Run: python -m unittest discover -s tests -v
No company workbooks, photos or user records are included.
"""
import io
import os
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from converter_mobile import (
    VERSION, rows_from_sheet, get_incidents, strict_datetime, hhmm, parse_bool,
    build_updates, read_field_tables, transform_package, image_to_jpeg,
)
from case_exports import build_case_excel_bytes, scrub_doc_properties

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'
DC='http://purl.org/dc/elements/1.1/'
CP='http://schemas.openxmlformats.org/package/2006/metadata/core-properties'

def col(i):
    s=''
    while i:
        i,k=divmod(i-1,26)
        s=chr(65+k)+s
    return s


def mk_sheet(headers, data, shared_map):
    root=ET.Element('{%s}worksheet'%S)
    ET.SubElement(root,'{%s}dimension'%S,ref='A1:%s%d'%(col(len(headers)),len(data)+1))
    sd=ET.SubElement(root,'{%s}sheetData'%S)
    for row_i,row in enumerate([headers]+data,1):
        row_el=ET.SubElement(sd,'{%s}row'%S,r=str(row_i))
        for column_i, val in enumerate(row,1):
            if val is None: continue
            c=ET.SubElement(row_el,'{%s}c'%S,r=col(column_i)+str(row_i))
            if isinstance(val,bool):
                c.set('t','b'); ET.SubElement(c,'{%s}v'%S).text='1' if val else '0'
            elif isinstance(val,(int,float)):
                ET.SubElement(c,'{%s}v'%S).text=str(val)
            else:
                c.set('t','s')
                text=str(val)
                if text not in shared_map:shared_map[text]=len(shared_map)
                ET.SubElement(c,'{%s}v'%S).text=str(shared_map[text])
    return ET.tostring(root,encoding='utf-8',xml_declaration=True)


def make_workbook(path,extra_col=False,missing_header=False):
    base=['ID','件名','発生日時','発生場所','状況概要']
    if extra_col: base+=['備考'+str(i) for i in range(6,34)]+['設備ID']
    if missing_header:base[1]='改名された件名'
    vals_a=['ID_A','CASE_A_PRIVATE_20261107','2026-11-07T08:00:00Z','架空室A','PRIVATE_A_NOTE_ABCDEF']
    vals_b=['ID_B','CASE_B_VISIBLE_20261107','2026-11-07T10:00:00Z','架空室B','PUBLIC_B_NOTE_ABCDEF']
    if extra_col:
        vals_a+=['a']*28+['EQ-A'];vals_b+=['b']*28+['EQ-B']
    sh1=(base,[vals_a,vals_b])
    sh2=(['ID','異常案件_ID','対応日時','実施内容・確認結果'],[
        ['T_A','ID_A','2026-11-07T08:10:00Z','PRIVATE_A_TIMELINE_ABCDEF'],
        ['T_B','ID_B',None,'PUBLIC_B_TIMELINE_ABCDEF']])
    sh3=(['ID','異常案件_ID','写真番号','元写真','写真右側の説明文'],[
        ['P_A','ID_A',1,'photoA.jpg','PRIVATE_A_PHOTO_ABCDEF'],
        ['P_B','ID_B',1,'photoB.jpg','PUBLIC_B_PHOTO_ABCDEF']])
    shared={}
    sheetxml={f'xl/worksheets/sheet{i}.xml':mk_sheet(*data,shared) for i,data in enumerate((sh1,sh2,sh3),1)}
    ss=ET.Element('{%s}sst'%S, count=str(len(shared)),uniqueCount=str(len(shared)))
    for t,i in sorted(shared.items(), key=lambda kv:kv[1]):
        ET.SubElement(ET.SubElement(ss,'{%s}si'%S),'{%s}t'%S).text=t
    wb=ET.Element('{%s}workbook'%S)
    sheets=ET.SubElement(wb,'{%s}sheets'%S)
    for i,name in enumerate(('異常案件','対応経過','写真帳'),1):
        ET.SubElement(sheets,'{%s}sheet'%S, {'name':name,'sheetId':str(i),'{%s}id'%R:'rId'+str(i)})
    rels=ET.Element('{%s}Relationships'%P)
    for i in range(1,4):ET.SubElement(rels,'{%s}Relationship'%P,Id='rId'+str(i),Target='worksheets/sheet%d.xml'%i)
    core=ET.Element('{%s}coreProperties'%CP)
    ET.SubElement(core,'{%s}creator'%DC).text='PRIVATE_Author'
    ET.SubElement(core,'{%s}lastModifiedBy'%CP).text='PRIVATE_EDITOR'
    def xml(root):return ET.tostring(root,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(path,'w') as z:
        z.writestr('xl/workbook.xml',xml(wb))
        z.writestr('xl/_rels/workbook.xml.rels',xml(rels))
        z.writestr('xl/sharedStrings.xml',xml(ss))
        for name, data in sheetxml.items():z.writestr(name,data)
        z.writestr('docProps/core.xml',xml(core))
    return path


class TestConverter(unittest.TestCase):
    def setUp(self):self.td=tempfile.TemporaryDirectory();self.addCleanup(self.td.cleanup)
    def workbook(self,**kwargs):return make_workbook(Path(self.td.name)/'test.xlsx',**kwargs)

    def test_version_unique(self):
        self.assertEqual(VERSION,'2.4.2-homebox-preview')

    def test_34th_column_preserved(self):
        b=self.workbook(extra_col=True)
        rows=rows_from_sheet(str(b),'異常案件')
        self.assertEqual(len(rows[0]),34)
        self.assertEqual(rows[0][33],'設備ID')
        self.assertEqual(rows[1][33],'EQ-A')
        self.assertEqual(rows[2][33],'EQ-B')

    def test_missing_required_header_rejected(self):
        b=self.workbook(missing_header=True)
        with self.assertRaisesRegex(ValueError,'必須列'):
            get_incidents(b)

    def test_time_zone_utc(self):
        self.assertEqual(strict_datetime('2026-11-06T10:00:00Z','test').strftime('%H:%M'),'19:00')

    def test_time_zone_plus0900(self):
        self.assertEqual(strict_datetime('2026-11-06T10:00:00+09:00','test').strftime('%H:%M'),'10:00')

    def test_time_formats(self):
        self.assertEqual(hhmm('9:05'),'09:05')
        self.assertEqual(hhmm('10:30:00'),'10:30')

    def test_invalid_time_rejected(self):
        with self.assertRaises(ValueError):hhmm('25:00')

    def test_false_string_not_truthy(self):
        self.assertFalse(parse_bool('FALSE','flag'))
        case={'件名':'架空故障','発生日時':'2026-11-07T01:00:00Z','発生場所':'架空室',
          '現場確認を実施':'FALSE','一次対応を実施':'TRUE'}
        rep,_=build_updates(case,[],[])
        self.assertIsNone(rep['G8'])
        self.assertEqual(rep['G9'],'有')

    def test_mixed_and_missing_timeline_dates(self):
        case={'件名':'架空故障','発生日時':'2026-11-07T01:00:00Z','発生場所':'架空室'}
        times=[{'対応日時':None,'実施内容・確認結果':'日時未設定'},
            {'対応日時':'2026-11-07T10:30:00+09:00','実施内容・確認結果':'日時あり'}]
        rep,_=build_updates(case,[],times)
        self.assertEqual(rep['G17'],'日時あり')
        self.assertEqual(rep['G18'],'日時未設定')

    def test_case_specific_excel_no_hidden_other_case(self):
        b=self.workbook(extra_col=True)
        output=build_case_excel_bytes(b,'ID_B')
        with zipfile.ZipFile(io.BytesIO(output)) as z:
            rows=rows_from_sheet_extracted(output, '異常案件')
            self.assertEqual(len(rows),2)
            self.assertEqual(rows[1][33],'EQ-B')
            for s in ('xl/sharedStrings.xml','docProps/core.xml','xl/worksheets/sheet1.xml'):
                data=z.read(s)
                self.assertNotIn(b'PRIVATE_A',data)
            self.assertNotIn('PRIVATE_EDITOR', z.read('docProps/core.xml').decode('utf-8'))
            self.assertEqual(len(rows_from_sheet_extracted(output,'写真帳')),2)

    def test_transform_removes_calc_chain_and_requests_recalculation(self):
        xdr='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
        a='http://schemas.openxmlformats.org/drawingml/2006/main'
        template=Path(self.td.name)/'dummy_company_template.xlsx'
        empty=ET.Element('{%s}worksheet'%S)
        ET.SubElement(empty,'{%s}sheetData'%S)
        draw=ET.Element('{%s}wsDr'%xdr)
        relationships=ET.Element('{%s}Relationships'%P)
        for idx in range(1,4):
            anchor=ET.SubElement(draw,'{%s}twoCellAnchor'%xdr)
            pic=ET.SubElement(anchor,'{%s}pic'%xdr)
            ET.SubElement(ET.SubElement(pic,'{%s}blipFill'%xdr),'{%s}blip'%a,{'{%s}embed'%R:f'rId{idx}'})
            mark=ET.SubElement(draw,'{%s}twoCellAnchor'%xdr)
            ET.SubElement(mark,'{%s}sp'%xdr)
            ET.SubElement(relationships,'{%s}Relationship'%P,Id=f'rId{idx}',Target=f'../media/image{idx}.jpeg')
        workbook=ET.Element('{%s}workbook'%S)
        ET.SubElement(workbook,'{%s}calcPr'%S, calcId='191029')
        workbookrels=ET.Element('{%s}Relationships'%P)
        ET.SubElement(workbookrels,'{%s}Relationship'%P,Id='rId22',Target='calcChain.xml')
        CT='http://schemas.openxmlformats.org/package/2006/content-types'
        contenttypes=ET.Element('{%s}Types'%CT)
        ET.SubElement(contenttypes,'{%s}Override'%CT,PartName='/xl/calcChain.xml',ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml')
        def xml(root):return ET.tostring(root,encoding='utf-8')
        with zipfile.ZipFile(template,'w') as z:
            z.writestr('xl/worksheets/sheet1.xml',xml(empty))
            z.writestr('xl/worksheets/sheet2.xml',xml(empty))
            z.writestr('xl/drawings/drawing2.xml',xml(draw))
            z.writestr('xl/drawings/_rels/drawing2.xml.rels',xml(relationships))
            z.writestr('xl/sharedStrings.xml',xml(ET.Element('{%s}sst'%S)))
            z.writestr('xl/workbook.xml',xml(workbook))
            z.writestr('xl/_rels/workbook.xml.rels',xml(workbookrels))
            z.writestr('[Content_Types].xml',xml(contenttypes))
            z.writestr('xl/calcChain.xml',b'<calcChain />')
            for idx in range(1,4):z.writestr(f'xl/media/image{idx}.jpeg',b'fake')
        output,audit=transform_package(template,{}, {}, {})
        with zipfile.ZipFile(io.BytesIO(output)) as z:
            self.assertNotIn('xl/calcChain.xml',z.namelist())
            wb=ET.fromstring(z.read('xl/workbook.xml'))
            calc=wb.find('{%s}calcPr'%S)
            self.assertEqual(calc.get('fullCalcOnLoad'),'1')
            self.assertEqual(calc.get('forceFullCalc'),'1')
            self.assertNotIn(b'calcChain.xml',z.read('xl/_rels/workbook.xml.rels'))
            self.assertNotIn(b'calcChain.xml',z.read('[Content_Types].xml'))
        self.assertEqual(len(audit['deleted_parts']),4)

    def test_email_jpeg_keeps_rotation_and_strips_exif(self):
        from PIL import Image
        image=Image.new('RGB',(30,10),'navy')
        exif=Image.Exif()
        exif[274]=6
        buff=io.BytesIO();image.save(buff,'JPEG',exif=exif)
        result,size=image_to_jpeg(buff.getvalue(),(80,80))
        self.assertEqual(size,[30,10])
        with Image.open(io.BytesIO(result)) as output:
            self.assertEqual(output.size,(80,80))
            self.assertFalse(output.getexif().get(274))
            self.assertEqual(output.format,'JPEG')

    def test_original_untouched(self):
        b=self.workbook()
        before=b.read_bytes()
        _=build_case_excel_bytes(b,'ID_B')
        self.assertEqual(before,b.read_bytes())


def rows_from_sheet_extracted(data,name):
    with tempfile.TemporaryDirectory() as tmp:
        f=Path(tmp)/'extracted.xlsx';f.write_bytes(data)
        return rows_from_sheet(str(f),name)


if __name__=='__main__':unittest.main()
