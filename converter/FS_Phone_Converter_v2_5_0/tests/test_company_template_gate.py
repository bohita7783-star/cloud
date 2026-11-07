"""Security gate tests: every workbook and image is synthetic. No company data."""
import io, zipfile, tempfile, unittest, sys
from pathlib import Path
from lxml import etree as ET
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
from converter_mobile import (check_company_report_package, transform_package, xml_root, ns, Q)
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
X='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
P='http://schemas.openxmlformats.org/package/2006/relationships'

def xml(obj):return ET.tostring(obj,encoding='utf-8',xml_declaration=True)

def clean_package(calc=False,after_calc=False):
    sheet=ET.Element('{%s}worksheet'%S); ET.SubElement(sheet,'{%s}sheetData'%S)
    dr=ET.Element('{%s}wsDr'%X)
    rel=ET.Element('{%s}Relationships'%P)
    for i in range(1,4):
        anchor=ET.SubElement(dr,'{%s}twoCellAnchor'%X)
        pic=ET.SubElement(anchor,'{%s}pic'%X)
        ET.SubElement(ET.SubElement(pic,'{%s}blipFill'%X),'{%s}blip'%A,{'{%s}embed'%R:'rId'+str(i)})
        marker=ET.SubElement(dr,'{%s}twoCellAnchor'%X)
        ET.SubElement(marker,'{%s}sp'%X)
        ET.SubElement(rel,'{%s}Relationship'%P,Id='rId'+str(i),Target=f'../media/image{i}.jpeg')
    wb=ET.Element('{%s}workbook'%S)
    if calc:ET.SubElement(wb,'{%s}calcPr'%S)
    if after_calc: ET.SubElement(wb,'{%s}extLst'%S)
    p={
        'xl/worksheets/sheet1.xml':xml(sheet),
        'xl/worksheets/sheet2.xml':xml(sheet),
        'xl/workbook.xml':xml(wb),
        'xl/drawings/drawing2.xml':xml(dr),
        'xl/drawings/_rels/drawing2.xml.rels':xml(rel),
        'xl/sharedStrings.xml':xml(ET.Element('{%s}sst'%S)),
    }
    p.update({f'xl/media/image{i}.jpeg':b'fake'+bytes([i]) for i in range(1,4)})
    return p

def aszip(parts):
    buffer=io.BytesIO()
    with zipfile.ZipFile(buffer,'w') as z:
        for p,v in parts.items():z.writestr(p,v)
    return buffer.getvalue()

class TestCompanyTemplateGate(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.addCleanup(self.tmp.cleanup)
    def source(self,parts):
        path=Path(self.tmp.name)/'synthetic_empty_company_template.xlsx'
        path.write_bytes(aszip(parts))
        return path
    def test_clean_synthetic_passes(self):
        p=clean_package()
        with zipfile.ZipFile(io.BytesIO(aszip(p))) as z:
            self.assertTrue(check_company_report_package(z)['risky_ooxml_parts_checked'])
        x,_=transform_package(self.source(p),{}, {}, {})
        self.assertGreater(len(x),50)
    def test_rejects_comments(self):
        p=clean_package();p['xl/comments1.xml']=b'<comments />'
        with self.assertRaisesRegex(ValueError,'SAFETY STOP'):
            transform_package(self.source(p),{}, {}, {})
    def test_rejects_custom_xml_and_properties(self):
        for name in ('customXml/item1.xml','docProps/custom.xml'):
            with self.subTest(name=name):
                p=clean_package();p[name]=b'<x>synthetic</x>'
                with self.assertRaisesRegex(ValueError,'未承認'):
                    transform_package(self.source(p),{}, {}, {})
    def test_rejects_external_link_without_modifying_source(self):
        p=clean_package()
        rel=ET.Element('{%s}Relationships'%P)
        ET.SubElement(rel,'{%s}Relationship'%P,Id='rId9',Target='file:///example.xlsx',TargetMode='External')
        p['xl/_rels/workbook.xml.rels']=xml(rel)
        src=self.source(p);before=src.read_bytes()
        with self.assertRaisesRegex(ValueError,'外部参照'):
            transform_package(src,{}, {}, {})
        self.assertEqual(src.read_bytes(),before)
    def test_rejects_worksheet_not_handled_by_converter(self):
        p=clean_package();p['xl/worksheets/sheet3.xml']=p['xl/worksheets/sheet1.xml']
        with self.assertRaisesRegex(ValueError,'追加ワークシート'):
            transform_package(self.source(p),{}, {}, {})
    def test_rejects_unknown_images_including_emf(self):
        for part in ('xl/media/image4.jpeg','xl/media/image5.emf'):
            with self.subTest(part=part):
                p=clean_package();p[part]=b'SYNTHETIC'
                with self.assertRaisesRegex(ValueError,'追加の写真'):
                    transform_package(self.source(p),{}, {}, {})
    def test_rejects_hidden_example_sheet(self):
        p=clean_package()
        wb=ET.Element('{%s}workbook'%S);ss=ET.SubElement(wb,'{%s}sheets'%S)
        ET.SubElement(ss,'{%s}sheet'%S,name='記入例',sheetId='1',state='hidden')
        p['xl/workbook.xml']=xml(wb)
        with self.assertRaisesRegex(ValueError,'非表示シート'):
            transform_package(self.source(p),{}, {}, {})
    def test_calcpr_inserted_before_extLst(self):
        p=clean_package(after_calc=True)
        out,_=transform_package(self.source(p),{}, {}, {})
        with zipfile.ZipFile(io.BytesIO(out)) as z:
            wb=xml_root(z.read('xl/workbook.xml'))
        self.assertEqual([c.tag.split('}')[-1] for c in wb],['calcPr','extLst'])
        self.assertEqual(wb.find('m:calcPr',ns).get('fullCalcOnLoad'),'1')

if __name__=='__main__':unittest.main()
