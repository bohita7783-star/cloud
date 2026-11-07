#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Field Survey XLSX + photos ZIP -> company format (safety-first TEST DRAFT).
Termux Python/lxml/Pillow, vetted 2026-10-17 template, one incident and up to 3 photos.
No network calls; refuses overwrite; NOT a company-approved/print-verified report.
"""
import argparse, os, sys, re, zipfile, tempfile, posixpath, io, hashlib, json, datetime, copy
from pathlib import Path
from PIL import Image, ImageOps
from lxml import etree as ET

SS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
DML='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
DRAW='http://schemas.openxmlformats.org/drawingml/2006/main'
REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
PR='http://schemas.openxmlformats.org/package/2006/relationships'
ns={'m':SS,'xdr':DML,'a':DRAW,'r':REL}
Q=lambda s:'{%s}'%SS+s




def serial_to_datetime(v):
    return datetime.datetime(1899,12,30)+datetime.timedelta(days=float(v))

REQUIRED_FIELD_HEADERS = {
    '異常案件': ('ID', '件名', '発生日時', '発生場所', '状況概要'),
    '写真帳': ('ID', '異常案件_ID', '写真番号', '元写真', '写真右側の説明文'),
    '対応経過': ('ID', '異常案件_ID', '対応日時', '実施内容・確認結果'),
}
MAX_FIELD_COLUMNS = 256
MAX_FIELD_ROWS = 50000


def rows_from_sheet(filename, name):
    """Read all populated Field Survey columns, refusing ambiguous schema.

    This is a narrow XLSX table parser, not a general Excel interpreter.
    Field Survey exports may have more than 33 columns after YAML changes.
    """
    with zipfile.ZipFile(filename) as z:
        zip_guard(z, 45 * 1024 * 1024)
        wb=xml_root(z.read('xl/workbook.xml'))
        rels=xml_root(z.read('xl/_rels/workbook.xml.rels'))
        relmap={r.get('Id'):r.get('Target') for r in rels}
        sheet=next((s for s in wb.findall('.//m:sheet',ns) if s.get('name')==name), None)
        if sheet is None:raise ValueError(f'Missing Field Survey sheet: {name}')
        target=relmap.get(sheet.get('{%s}id'%REL))
        if not target:raise ValueError('Missing relationship for '+name)
        xmlpath=(target.lstrip('/') if target.startswith('/') else posixpath.normpath(posixpath.join('xl',target)))
        sh=xml_root(z.read(xmlpath))
        strings=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            sr=xml_root(z.read('xl/sharedStrings.xml'))
            strings=[''.join(si.itertext()) for si in sr.findall('m:si',ns)]
        source_rows=sh.findall('.//m:sheetData/m:row',ns)
        maxcol=1
        maxrow=1
        for row in source_rows:
            ridx=int(row.get('r','0'))
            if not 1<=ridx<=MAX_FIELD_ROWS:raise ValueError('Field Survey row number outside safety limit')
            maxrow=max(maxrow,ridx)
            for cell in row.findall('m:c',ns):
                col=cell_column_index(cell.get('r',''))
                if col>MAX_FIELD_COLUMNS:raise ValueError('Field Survey columns exceed parser safety limit')
                maxcol=max(maxcol,col)
        output=[[None]*maxcol for _ in range(maxrow)]
        for row in source_rows:
            values=output[int(row.get('r'))-1]
            for c in row.findall('m:c',ns):
                col=cell_column_index(c.get('r'))-1
                typ=c.get('t')
                v=c.find('m:v',ns)
                inline=c.find('m:is',ns)
                if typ=='inlineStr':val=''.join(inline.itertext()) if inline is not None else ''
                elif typ=='s':
                    idx=int(v.text) if v is not None and v.text else -1
                    if not 0<=idx<len(strings):raise ValueError('Invalid shared-string index')
                    val=strings[idx]
                elif typ=='b':val=bool(int(v.text)) if v is not None else False
                elif typ=='str':val=v.text if v is not None else ''
                elif v is not None and v.text is not None:
                    val=float(v.text)
                    if val.is_integer():val=int(val)
                else:val=None
                values[col]=val
        headers=[str(v).strip() if v is not None else '' for v in output[0]]
        meaningful=[v for v in headers if v]
        if len(meaningful)!=len(set(meaningful)):
            raise ValueError(f'{name}: duplicate Field Survey header')
        required=REQUIRED_FIELD_HEADERS.get(name,())
        missing=[v for v in required if v not in meaningful]
        if missing:raise ValueError(f'{name}: 必須列がありません: '+', '.join(missing))
        for i,h in enumerate(headers):
            if not h and any(row[i] not in (None,'') for row in output[1:]):
                raise ValueError(f'{name}: blank heading for column {i+1}')
        return output


def header_row_dict(rows):
    names=[str(n).strip() if n is not None else '' for n in rows[0]]
    return [dict(zip(names, row)) for row in rows[1:] if row[0] not in (None,'')]

def cell_column_index(a):
    s=re.match(r'^[A-Z]+',a).group();v=0
    for ch in s:v=v*26+ord(ch)-64
    return v

def get_or_insert_cell(root,addr):
    sheet_data=root.find('m:sheetData',ns)
    row_idx=int(re.search(r'[0-9]+$',addr).group())
    row=sheet_data.find(f"m:row[@r='{row_idx}']",ns)
    if row is None:
        row=ET.Element(Q('row'),r=str(row_idx))
        idx=0
        for existing in sheet_data:
            if existing.tag!=Q('row'):continue
            if int(existing.attrib['r'])>row_idx:break
            idx+=1
        sheet_data.insert(idx,row)
    cell=row.find(f"m:c[@r='{addr}']",ns)
    if cell is None:
        cell=ET.Element(Q('c'),r=addr)
        insert_idx=0
        for n in row:
            if n.tag==Q('c') and cell_column_index(n.attrib['r'])<cell_column_index(addr):insert_idx+=1
        row.insert(insert_idx,cell)
    return cell

# --- v2 safety-focused replacement. 2026-11-04. ---
# This is an auditable TEST DRAFT converter, NOT a company-approved final report.
# It preserves original OOXML parts instead of round-tripping Excel through a
# generic spreadsheet writer (the template contains EMF and print resources).
import shutil
import uuid
import math
import stat

import engine_policy as EP

VERSION = '2.6.0-form-profile'
# Kept for reference only. Since 2.3.3 the template must match a SIGNED approval manifest
# (approval.verify_manifest) instead of this single hard-coded hash.
LEGACY_COMPANY_TEMPLATE_SHA256 = '569966343fa14d98a84bb45ccc27f9a54cfacb30b161f3bd9cc2cd536cdd53cc'
UNMAPPED_CASE_FIELDS = ('案件識別番号','対象区分','オーナー・テナント報告','対応業者への連絡',
                        '社内報告','報告先・提出状況','復旧・完了状況')
MAX_PHOTOS = 3  # photo slots of the 2026-10 form (the template decides; see form_profile)
MAX_ZIP_INPUT = 160 * 1024 * 1024
MAX_ZIP_MEMBERS = 800
MAX_IMAGE_SIZE = 30 * 1024 * 1024
MAX_IMAGE_PIXELS = 32_000_000
MAX_TEXT_LENGTH = 2500
# 2.6.0: the photo slots are derived from the template (form_profile.photo_plan).
# The baseline (vetted 2026-10-17 report) keeps its 3 picture parts image1..3.jpeg.
BASELINE_IMAGES = ('image1.jpeg', 'image2.jpeg', 'image3.jpeg')
# Incident-specific fields of the original completed report; must be blanked on
# every conversion. Header/static dropdown tables elsewhere are preserved.
CLEAR_REPORT = {
    'AB2','AU1','G3','W3','G4','X4','AP4','BC4','G5','X5','AC5','AN5',
    'G6','AP6','G7','S7','AG7','G8','L8','S8','AC8','AP8',
    'G9','L9','S9','AC9','AP9',
    'K10','Y10','K11','T11','AL11','AU11','K12','T12','AL12','AU12',
    'G13','K13','AH13','AL13','K14','Q14','U14','AL14','AR14','AV14',
    'AR32','AR34','G32','N32','BE43',
    'G36','K36','T36','AH36','AL36','AU36','G37','K37','T37','AH37','AL37','AU37',
    'C28','C29','C30','N33','N34','N35','G26',
}
CLEAR_REPORT.update(f'{col}{row}' for row in range(15,28) for col in ('C','G'))
# 2.5.0: every cell the report layout writes is cleared first, except cells whose
# template value must survive when the case leaves them empty (site name, and the
# '：' placeholders of the notification means cells).
import report_layout as _RL
import form_profile as FP
KEEP_TEMPLATE_IF_EMPTY = {'AP2','AP3'}  # means cells are always rewritten ('：' when unused)
CLEAR_REPORT.update(_RL.REPORT_CELLS - KEEP_TEMPLATE_IF_EMPTY)
CLEAR_REPORT.update({'AR33','AR35'})


def clear_photo_cells(slots):
    """Photo-sheet cells blanked before writing: title, date, every slot's number and caption lines."""
    cells = {FP.PHOTO_TITLE[0], FP.PHOTO_DATE[0]}
    for slot in slots:
        cells.add(slot['number'])
        cells.update(slot['captions'])
    return cells


CLEAR_PHOTO = clear_photo_cells(FP.DEFAULT_SLOTS)  # preview / reference only


def xml_root(data):
    return ET.fromstring(data, parser=ET.XMLParser(resolve_entities=False, no_network=True, huge_tree=False))


def encode_xml(root):
    return ET.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def zip_guard(z, max_bytes=MAX_ZIP_INPUT):
    members=z.infolist()
    if len(members)>MAX_ZIP_MEMBERS:raise ValueError('ZIP has too many members')
    names=[x.filename for x in members]
    if len(names)!=len(set(names)):raise ValueError('ZIP has duplicated member names')
    if sum(x.file_size for x in members)>max_bytes:raise ValueError('ZIP uncompressed content too large')
    for x in members:
        p=x.filename
        if p.startswith('/') or '\\' in p or any(seg in ('..','') for seg in p.split('/')[:-1]):
            raise ValueError('Unsafe ZIP member: '+p)
        if x.file_size>max_bytes:raise ValueError('Individual ZIP member is too large: '+p)
    failed=z.testzip()
    if failed:raise ValueError('Corrupt ZIP entry: '+failed)


def validate_path_separation(field, photos, template, out, audit=None):
    # photos=None: a HomeBox package carries its photos inside the same file.
    src=[Path(x).expanduser().resolve(strict=True) for x in (field,photos,template) if x is not None]
    if len(set(src))!=len(src):raise ValueError('The source paths must be distinct')
    out_path=Path(out).expanduser().resolve(strict=False)
    audit_path=Path(audit).expanduser().resolve(strict=False) if audit else None
    if out_path in src or (audit_path and audit_path in src):
        raise ValueError('SAFETY STOP: output or audit path points to a source file')
    if audit_path and audit_path==out_path:
        raise ValueError('Output XLSX and audit JSON may not use the same path')
    if out_path.exists() or (audit_path and audit_path.exists()):
        raise FileExistsError('SAFETY STOP: output file already exists; choose another name')
    if not out_path.parent.is_dir():raise ValueError('Output directory does not exist')
    if audit_path and not audit_path.parent.is_dir():raise ValueError('Audit directory does not exist')
    if out_path.suffix.lower()!='.xlsx':raise ValueError('Output must be an .xlsx file')
    return src,out_path,audit_path


def get_incidents(field):
    """Read all incidents for explicit selection. No data is written."""
    with zipfile.ZipFile(field) as z: zip_guard(z,45*1024*1024)
    cases=header_row_dict(rows_from_sheet(str(field),'異常案件'))
    keys=[str(c.get('ID') or '').strip() for c in cases]
    if not all(keys) or len(keys)!=len(set(keys)):
        raise ValueError('案件IDの欠落または重複があります')
    return cases


def get_related_rows(field, case_id):
    case=next((c for c in get_incidents(field) if str(c['ID'])==str(case_id)),None)
    if case is None: raise ValueError('選択した案件IDがExcelにありません')
    photos=[p for p in header_row_dict(rows_from_sheet(str(field),'写真帳'))
            if str(p.get('異常案件_ID'))==str(case_id)]
    timelines=[t for t in header_row_dict(rows_from_sheet(str(field),'対応経過'))
               if str(t.get('異常案件_ID'))==str(case_id)]
    return case,photos,timelines


def read_field_tables(field, case_id=None, photo_actions=None):
    cases=get_incidents(field)
    if case_id is None:
        if len(cases)!=1:
            raise ValueError(f'案件が{len(cases)}件あります。件名で対象案件を指定してください')
        case=cases[0]
    else:
        case=next((c for c in cases if str(c['ID'])==str(case_id)),None)
        if case is None:raise ValueError('案件IDが見つかりません')
    original_photos=[p for p in header_row_dict(rows_from_sheet(str(field),'写真帳'))
                     if str(p.get('異常案件_ID'))==str(case['ID'])]
    timeline=[t for t in header_row_dict(rows_from_sheet(str(field),'対応経過'))
              if str(t.get('異常案件_ID'))==str(case['ID'])]
    return select_case_rows(case,original_photos,timeline,photo_actions)


def select_case_rows(case, original_photos, timeline, photo_actions=None):
    """Shared by Field Survey XLSX input and HomeBox package input (2.4.0)."""
    for key in ('件名','発生日時','発生場所','状況概要'):
        if case.get(key) in (None,''):raise ValueError('必須の案件情報がありません: '+key)
    if any(len(str(v))>MAX_TEXT_LENGTH for v in case.values() if v is not None):
        raise ValueError('案件文章の長さが上限を超えました')
    photo_actions=photo_actions or {}
    exported_ids={str(p.get('ID')) for p in original_photos}
    if set(photo_actions)-exported_ids:raise ValueError('選択案件に存在しない写真IDが指定されました')
    rows=[];seen_ids=set()
    for item in original_photos:
        row=dict(item)
        rid=str(row.get('ID') or '')
        if not rid or rid in seen_ids:raise ValueError('写真行IDの欠落・重複')
        seen_ids.add(rid)
        edit=photo_actions.get(rid,{})
        if edit.get('remove'):continue
        if 'description' in edit:row['写真右側の説明文']=str(edit['description'])
        if 'location' in edit:row['撮影場所']=str(edit['location'])
        if not str(row.get('写真右側の説明文') or '').strip():
            raise ValueError('写真の説明文が空です: '+str(row.get('写真番号')))
        if any(len(str(v))>MAX_TEXT_LENGTH for v in row.values() if v is not None):
            raise ValueError('写真の説明が長すぎます')
        # We accept exported duplicates in numbers, then renumber in selected order.
        # Source record ID is used as stable identity, NOT its possibly stale number.
        rows.append(row)
    # If the user intentionally removes rows, missing photo positions are
    # compacted, to avoid an empty photo1 slot containing a prior incident's photo.
    def ordered(row):
        number=row.get('写真番号')
        try: n=float(number)
        except (ValueError,TypeError):n=float('inf')
        return (n,str(row.get('createAt') or ''),str(row['ID']))
    rows.sort(key=ordered)
    if len(rows)>MAX_PHOTOS:
        raise ValueError(f'写真が{len(rows)}件あります。今回の会社原本の枠は{MAX_PHOTOS}枚までなので使用する写真を選択してください')
    for i,row in enumerate(rows,1):row['写真番号']=i
    if len(timeline)>8:raise ValueError('対応経過は8件まで。省略せずに扱う追加機能が必要です')
    return case,rows,timeline

def strict_datetime(v, name):
    if isinstance(v,(int,float)) and math.isfinite(v) and 1 < v < 120000:
        val=serial_to_datetime(v)
    elif isinstance(v,str):
        try:
            val=datetime.datetime.fromisoformat(v.replace('Z','+00:00'))
            if val.tzinfo is not None:
                val=val.astimezone(datetime.timezone(datetime.timedelta(hours=9))).replace(tzinfo=None)
        except Exception as ex:raise ValueError(f'Invalid date in {name}: {v}') from ex
    else:raise ValueError(f'Missing/invalid date in {name}')
    if not 2000<=val.year<=2100:raise ValueError('Out-of-range date in '+name)
    return val


def safe_text(v):
    return str(v).strip() if v not in (None,'') else ''


def parse_bool(value, name):
    """Strict conversion: the string 'FALSE' must never count as truthy."""
    if isinstance(value, bool):return value
    if value is None or value=='':return False
    if isinstance(value, (int,float)) and value in (0,1):return bool(value)
    if isinstance(value,str):
        v=value.strip().casefold()
        if v in ('true','1','yes','はい','有'):return True
        if v in ('false','0','no','いいえ','無',''):return False
    raise ValueError(f'Invalid Boolean field {name}: {value!r}')


def hhmm(v):
    if v is None or v=='':return None
    if isinstance(v,(int,float)) and math.isfinite(v):
        total=round((float(v)%1)*1440)
        return f'{(total//60)%24:02d}:{total%60:02d}'
    if isinstance(v,str):
        s=v.strip()
        match=re.fullmatch(r'(\d{1,2}):(\d{2})(?::(\d{2}))?',s)
        if match:
            h, m, sec = int(match[1]),int(match[2]),int(match[3] or 0)
            if h>23 or m>59 or sec>59:raise ValueError('Invalid time: '+s)
            return f'{h:02d}:{m:02d}'
        return strict_datetime(s,'time').strftime('%H:%M')
    raise ValueError('Unsupported time format')


def image_to_jpeg(data, size):
    if len(data)>MAX_IMAGE_SIZE:raise ValueError('Photo larger than 30MB')
    with Image.open(io.BytesIO(data)) as src:
        if src.format not in ('JPEG','PNG','WEBP'):raise ValueError('Unrecognized photo type')
        if src.width*src.height>MAX_IMAGE_PIXELS:raise ValueError('Photo pixel count too large')
        source_size=[src.width,src.height]
        src=ImageOps.exif_transpose(src).convert('RGB')
        src.thumbnail((1600,1600),Image.Resampling.LANCZOS)
        fitted=ImageOps.contain(src,size,Image.Resampling.LANCZOS)
        dst=Image.new('RGB',size,'white')
        dst.paste(fitted,((size[0]-fitted.width)//2,(size[1]-fitted.height)//2))
        buf=io.BytesIO();dst.save(buf,'JPEG',quality=80,optimize=True,subsampling=0)
        return buf.getvalue(),source_size


def photo_to_jpeg(data):
    """2.6.0 blank form: the photo keeps its own shape (long edge <= 1600 px), no EXIF."""
    if len(data)>MAX_IMAGE_SIZE:raise ValueError('Photo larger than 30MB')
    with Image.open(io.BytesIO(data)) as src:
        if src.format not in ('JPEG','PNG','WEBP'):raise ValueError('Unrecognized photo type')
        if src.width*src.height>MAX_IMAGE_PIXELS:raise ValueError('Photo pixel count too large')
        source_size=[src.width,src.height]
        src=ImageOps.exif_transpose(src).convert('RGB')
        src.thumbnail((1600,1600),Image.Resampling.LANCZOS)
        buf=io.BytesIO();src.save(buf,'JPEG',quality=80,optimize=True,subsampling=0)
        return buf.getvalue(),source_size,src.size


def photo_part_name(plan, i):
    """Media file name (under xl/media/) of photo i (0-based) for this template's plan."""
    if plan['mode']=='baseline':return BASELINE_IMAGES[i]
    return posixpath.basename(EP.GEN_MEDIA.format(i+1))


def read_photos(zip_path, rows, template, photo_actions=None, plan=None):
    photo_actions=photo_actions or {}
    with zipfile.ZipFile(zip_path) as z,zipfile.ZipFile(template) as original:
        zip_guard(z)
        if plan is None:plan=FP.analyze_zip(original)
        if len(rows)>len(plan['slots']):
            raise ValueError(f'写真が{len(rows)}枚あります。この様式の写真枠は{len(plan["slots"])}つです')
        names={}
        for item in z.infolist():
            if item.is_dir():continue
            bn=posixpath.basename(item.filename)
            if bn in names:raise ValueError('写真ZIP内に同じファイル名があります: '+bn)
            names[bn]=item.filename
        images={}; info=[];hashes={}
        for i,row in enumerate(rows):
            rid=str(row['ID']);edit=photo_actions.get(rid,{})
            basename=row.get('元写真')
            if 'image_bytes' in edit:
                data=edit['image_bytes'];origin='差し替え画像'
            else:
                if not isinstance(basename,str) or basename not in names:
                    raise ValueError('写真がZIPにありません。編集画面で差し替え可能です: '+str(basename))
                data=z.read(names[basename]);origin='Field Survey写真ZIP'
            checksum=hashlib.sha256(data).hexdigest()
            if checksum in hashes:
                raise ValueError(f'写真{i+1}と写真{hashes[checksum]}が完全に同一です。編集画面で差し替えるか、片方を外してください')
            hashes[checksum]=i+1
            part=photo_part_name(plan,i)
            if plan['mode']=='baseline':
                with Image.open(io.BytesIO(original.read('xl/media/'+part))) as im:
                    target_size=im.size
                if max(target_size)>1600:
                    ratio=1600/max(target_size)
                    target_size=tuple(max(1,round(v*ratio)) for v in target_size)
                images[part],source_size=image_to_jpeg(data,target_size)
            else:
                images[part],source_size,target_size=photo_to_jpeg(data)
            info.append({'number':i+1,'file':basename or '(元写真なし)',
                         'origin':origin,'source_size':source_size,
                         'embedded_size':list(target_size),'embedded_bytes':len(images[part]),'jpeg_quality':80,'max_long_edge':1600,'sha256':checksum,
                         'caption':row['写真右側の説明文']})
        return images,info

def scrub_sheet(xml_bytes, clear_cells, updates):
    root=xml_root(xml_bytes)
    for addr in sorted(clear_cells):
        c=root.find(f".//m:sheetData/m:row/m:c[@r='{addr}']",ns)
        if c is not None:
            for child in list(c):
                if child.tag in (Q('v'),Q('is'),Q('f')):c.remove(child)
            c.attrib.pop('t',None)
    # Remove stale cached formula values from both report and accident summary.
    old_formula_caches=0
    for cell in root.findall('.//m:sheetData/m:row/m:c',ns):
        formula=cell.find('m:f',ns)
        if formula is None:continue
        val=cell.find('m:v',ns)
        if val is not None:
            cell.remove(val);old_formula_caches+=1
        cell.attrib.pop('t',None)
    for addr,val in updates.items():
        cell=get_or_insert_cell(root,addr)
        for child in list(cell):
            if child.tag in (Q('v'),Q('is'),Q('f')):cell.remove(child)
        cell.attrib.pop('t',None)
        if val is None:continue
        if isinstance(val,(int,float,bool)):
            ET.SubElement(cell,Q('v')).text=str(int(val) if isinstance(val,(bool,int)) else val)
        else:
            cell.set('t','inlineStr')
            isel=ET.SubElement(cell,Q('is'))
            el=ET.SubElement(isel,Q('t'))
            if '\n' in str(val):el.set('{http://www.w3.org/XML/1998/namespace}space','preserve')
            el.text=str(val)
    return encode_xml(root),old_formula_caches



def scrub_doc_properties(xml:bytes, core:bool) -> bytes:
    """Remove identifying metadata from an exported copy, not the source."""
    root=xml_root(xml)
    restricted=EP.DOCPROP_CORE_BLANK if core else EP.DOCPROP_APP_BLANK
    for child in root.iter():
        if child.tag.split('}')[-1] in restricted:
            child.text=''
            for nested in list(child):child.remove(nested)
    return encode_xml(root)

def scrub_orphan_shared_strings(shared_xml, sheets):
    """Empty no-longer-referenced sharedStrings entries without reindexing any cell."""
    used=set()
    for sheet in sheets:
        root=xml_root(sheet)
        for c in root.findall('.//m:sheetData/m:row/m:c',ns):
            if c.get('t')=='s':
                v=c.find('m:v',ns)
                if v is not None and v.text is not None:used.add(int(v.text))
    root=xml_root(shared_xml)
    cleared=0
    for idx,si in enumerate(root.findall('m:si',ns)):
        if idx not in used:
            for child in list(si):si.remove(child)
            ET.SubElement(si,Q('t')).text=''
            cleared+=1
    return encode_xml(root),cleared


def build_updates_with_notes(case, photos, timeline, today=None, slots=None):
    """2.5.0: lay the case out like completed company reports (see report_layout.py).
    2.6.0: photo slots come from the template (None = the 2026-10-31 form, for previews)."""
    import report_layout as RL
    incident_date=strict_datetime(case['発生日時'],'発生日時')
    issue_day=(incident_date.date()-datetime.date(1899,12,30)).days
    r,warn=RL.layout_report(case,timeline,case.get('_contacts') or [],strict_datetime,parse_bool,hhmm,today=today)
    p,pwarn=RL.layout_photos(case,photos,issue_day,slots or FP.DEFAULT_SLOTS)
    return r,p,warn+pwarn


def build_updates(case, photos, timeline, slots=None):
    r,p,_=build_updates_with_notes(case,photos,timeline,slots=slots)
    return r,p


# Strict, review-first blocklist for a report produced from an Excel package.
# Unlike case_exports, this protects the COMPANY REPORT branch.  It does NOT
# remove external links, formulas, EMF, or arbitrary original parts: on risky
# packages it refuses to generate any result.  A separate approval workflow
# for a CLEAN blank template is needed; no user-level bypass is supplied.
RISKY_PART_PREFIXES = EP.RISKY_PART_PREFIXES
RISKY_EXACT_PARTS = EP.RISKY_EXACT_PARTS
ALLOWED_PHOTO_PARTS = EP.ALLOWED_PHOTO_PARTS


def check_company_report_package(z, phase='template'):
    """Fail closed for high-risk OOXML artifacts (not a full DLP scan).

    Generic XML and untouched drawing content can still leak information:
    therefore a separately reviewed clean master is still indispensable.
    All file names are listed as *categories*, not logged with their values.
    """
    names=set(z.namelist())
    problems=[]
    for name in sorted(names):
        lower=name.lower()
        if name in RISKY_EXACT_PARTS or any(name.startswith(x) for x in RISKY_PART_PREFIXES):
            problems.append(f'未承認の内部部品: {name}')
        elif name.startswith('xl/media/') and name not in ALLOWED_PHOTO_PARTS:
            problems.append(f'追加の写真・EMF等: {name}（承認済み空白様式の検査が必要）')
        elif lower.endswith('.vml') or '/vml' in lower:
            problems.append(f'VML図形・コメント等: {name}')
        elif re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name) and name not in (
                'xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml'):
            # This converter modifies only two sheets. A worksheet left behind
            # could contain complete text from earlier incidents.
            problems.append(f'未処理の追加ワークシート: {name}')
        elif name.endswith('.rels'):
            try:
                relroot=xml_root(z.read(name))
                for rel in relroot:
                    if rel.get('TargetMode','').lower()=='external':
                        problems.append(f'外部参照の関係: {name}（自動切断しない）')
                        break
            except ET.XMLSyntaxError:
                problems.append(f'関係XMLを検証できません: {name}')
    if 'xl/workbook.xml' in names:
        try:
            wb=xml_root(z.read('xl/workbook.xml'))
            for sh in wb.findall('.//m:sheet',ns):
                nm=sh.get('name','')
                if sh.get('state','visible')!='visible':
                    problems.append('非表示シートが存在（内容は検査していません）')
                if re.search(r'(記入例|サンプル|sample|example)',nm,re.I):
                    problems.append('記入例・サンプルのシートが存在')
            for item in wb.findall('.//m:definedName',ns):
                v=item.text or ''
                if re.search(r'\[[^\]]+\.(?:xlsx?|xlsm|xlsb)\]',v,re.I):
                    problems.append('定義名に外部ブック参照が存在')
        except ET.XMLSyntaxError:
            problems.append('workbook.xmlを検証できません')
    # Detect hyperlinks and external references embedded directly in formulas;
    # the approved workflow must determine whether these are essential.
    for name in sorted(n for n in names if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n)):
        try:
            root=xml_root(z.read(name))
            for f in root.findall('.//m:f',ns):
                value=f.text or ''
                if ('HYPERLINK(' in value.upper() or
                    re.search(r'\[[^\]]+\.(?:xlsx?|xlsm|xlsb)\]',value,re.I)):
                    problems.append(f'シート数式に外部参照候補: {name}')
                    break
        except ET.XMLSyntaxError:
            problems.append(f'ワークシートXMLを検証できません: {name}')
    if problems:
        limit=10
        msg='; '.join(dict.fromkeys(problems[:limit]))
        if len(problems)>limit:msg+=f'; 他{len(problems)-limit}件'
        raise ValueError('SAFETY STOP: 会社様式の'+phase+'に未承認の情報残留要因があります。'+msg+
                         '。原本は変更していません。承認済み空白ひな形を準備してください。')
    return {'risky_ooxml_parts_checked': True,
            'company_report_extra_sheets':0,
            'unapproved_media_parts':0,
            'external_relationships_detected':0}


CT_NS='http://schemas.openxmlformats.org/package/2006/content-types'
DRAWING_CT='application/vnd.openxmlformats-officedocument.drawing+xml'
SHEET_TAIL=('legacyDrawing','legacyDrawingHF','drawingHF','picture','oleObjects','controls',
            'webPublishItems','tableParts','extLst')


def _new_rid(existing, base):
    i=1
    while f'{base}{i}' in existing:i+=1
    return f'{base}{i}'


def _picture_anchor(cnv_id, number, rid, placement):
    col,coff,row,roff,cx,cy=placement
    X=lambda t:'{%s}%s'%(DML,t)
    D=lambda t:'{%s}%s'%(DRAW,t)
    anchor=ET.Element(X('oneCellAnchor'))
    frm=ET.SubElement(anchor,X('from'))
    for tag,val in (('col',col),('colOff',coff),('row',row),('rowOff',roff)):
        ET.SubElement(frm,X(tag)).text=str(val)
    ET.SubElement(anchor,X('ext'),cx=str(cx),cy=str(cy))
    pic=ET.SubElement(anchor,X('pic'))
    nv=ET.SubElement(pic,X('nvPicPr'))
    ET.SubElement(nv,X('cNvPr'),id=str(cnv_id),name=f'写真 {number}')
    ET.SubElement(ET.SubElement(nv,X('cNvPicPr')),D('picLocks'),noChangeAspect='1')
    fill=ET.SubElement(pic,X('blipFill'))
    blip=ET.SubElement(fill,D('blip'),nsmap={'r':REL})
    blip.set('{%s}embed'%REL,rid)
    ET.SubElement(ET.SubElement(fill,D('stretch')),D('fillRect'))
    sp=ET.SubElement(pic,X('spPr'))
    xfrm=ET.SubElement(sp,D('xfrm'))
    ET.SubElement(xfrm,D('off'),x='0',y='0')
    ET.SubElement(xfrm,D('ext'),cx=str(cx),cy=str(cy))
    ET.SubElement(ET.SubElement(sp,D('prstGeom'),prst='rect'),D('avLst'))
    ET.SubElement(anchor,X('clientData'))
    return anchor


def place_blank_photos(z, names, plan, sheet2_root, embedded):
    """2.6.0 blank form: put each photo over its slot as a floating picture.

    Existing drawing content and relationships are kept as they are; new parts use
    the fixed names of engine_policy (GEN_*). sheet2_root gets a <drawing> element
    only when the photo sheet had no drawing. Returns (changed parts, added parts).
    """
    tsheet=xml_root(z.read(EP.SHEET_PHOTO))
    mdw=FP.max_digit_width(z.read('xl/styles.xml') if 'xl/styles.xml' in names else b'')
    col_px,row_emu=FP.sheet_geometry(tsheet,mdw)
    changes={};added={}
    drawing=plan['photo_drawing']
    if drawing:
        droot=xml_root(z.read(drawing))
        drels_name=plan['photo_drawing_rels'] or FP.rels_name(drawing)
        drels=xml_root(z.read(drels_name)) if plan['photo_drawing_rels'] else ET.Element('{%s}Relationships'%PR,nsmap={None:PR})
    else:
        drawing=EP.GEN_DRAWING;drels_name=EP.GEN_DRAWING_RELS
        droot=ET.Element('{%s}wsDr'%DML,nsmap={'xdr':DML,'a':DRAW})
        drels=ET.Element('{%s}Relationships'%PR,nsmap={None:PR})
        srels=xml_root(z.read(EP.SHEET_PHOTO_RELS)) if plan['photo_sheet_rels'] else ET.Element('{%s}Relationships'%PR,nsmap={None:PR})
        rid=_new_rid({r.get('Id') for r in srels},'rIdHbDr')
        ET.SubElement(srels,'{%s}Relationship'%PR,Id=rid,Type=FP.DRAWING_TYPE,Target='../drawings/'+posixpath.basename(EP.GEN_DRAWING))
        el=ET.Element(Q('drawing'),nsmap={'r':REL})
        el.set('{%s}id'%REL,rid)
        tail={Q(t) for t in SHEET_TAIL}
        sheet2_root.insert(next((i for i,c in enumerate(sheet2_root) if c.tag in tail),len(sheet2_root)),el)
        (changes if plan['photo_sheet_rels'] else added)[EP.SHEET_PHOTO_RELS]=encode_xml(srels)
    ids=[int(x.get('id')) for x in droot.iter('{%s}cNvPr'%DML) if (x.get('id') or '').isdigit()]
    next_id=max(ids,default=1)+1
    existing={r.get('Id') for r in drels}
    for i in range(len(embedded)):
        part=photo_part_name(plan,i);data=embedded[part]
        with Image.open(io.BytesIO(data)) as im:w,h=im.size
        rid=_new_rid(existing,'rIdHb');existing.add(rid)
        ET.SubElement(drels,'{%s}Relationship'%PR,Id=rid,Type=REL+'/image',Target='../media/'+part)
        droot.append(_picture_anchor(next_id+i,i+1,rid,FP.place(plan['slots'][i],col_px,row_emu,w,h)))
        added['xl/media/'+part]=data
    (changes if plan['photo_drawing'] else added)[drawing]=encode_xml(droot)
    (changes if plan['photo_drawing_rels'] else added)[drels_name]=encode_xml(drels)
    return changes,added


def transform_package(template, report_updates, photo_updates, embedded, approval=None, plan=None):
    """Build the report package.

    approval=None   : v2.3.2 behaviour, stop on any risky part (unapproved template).
    approval=<dict> : a manifest already verified by approval.verify_manifest for THIS
                      file; approved parts (EMF, VML, checkboxes, extra sheets, links)
                      pass through and the output is verified part by part.
    plan            : form_profile plan of this template (derived again when omitted).
    """
    if approval is not None:
        with open(template,'rb') as fh:
            if hashlib.sha256(fh.read()).hexdigest()!=approval.get('template',{}).get('sha256'):
                raise ValueError('APPROVAL STOP: 承認済み一覧と原本のハッシュ値が一致しません')
    with zipfile.ZipFile(template) as z:
        zip_guard(z)
        names=set(z.namelist())
        if approval is None:
            check_company_report_package(z,'原本')
        if plan is None:plan=(approval or {}).get('_plan') or FP.analyze_zip(z)
        if plan['mode']=='refused':
            raise ValueError('SAFETY STOP: このひな形は使えません：'+' / '.join(plan['problems'][:3]))
        baseline=plan['mode']=='baseline'
        if approval is None and not baseline:
            raise ValueError('SAFETY STOP: 空様式から作るには承認済み一覧が必要です')
        required=['xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml','xl/sharedStrings.xml']
        if baseline:
            required+=['xl/drawings/drawing2.xml','xl/drawings/_rels/drawing2.xml.rels']+['xl/media/'+n for n in BASELINE_IMAGES]
        if not set(required)<=names:raise ValueError('Unsupported original workbook parts')
        if len(embedded)>len(plan['slots']):raise ValueError('写真が写真枠の数より多いです')
        sheet1,fc1=scrub_sheet(z.read(required[0]),CLEAR_REPORT,report_updates)
        sheet2,fc2=scrub_sheet(z.read(required[1]),clear_photo_cells(plan['slots']),photo_updates)
        shared,orphans=scrub_orphan_shared_strings(z.read('xl/sharedStrings.xml'),[sheet1,sheet2]+[
            z.read(n) for n in sorted(names) if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',n) and n not in ('xl/worksheets/sheet1.xml','xl/worksheets/sheet2.xml')])
        changes={};added={};remove_parts=set();removed_drawings=0;removed_markings=0
        if baseline:
            drawing=xml_root(z.read('xl/drawings/drawing2.xml'))
            relationships=xml_root(z.read('xl/drawings/_rels/drawing2.xml.rels'))
            rid_targets={r.get('Id'):posixpath.basename(r.get('Target','')) for r in relationships}
            if {rid_targets.get('rId1'),rid_targets.get('rId2'),rid_targets.get('rId3')}!=set(BASELINE_IMAGES):
                raise ValueError('Unexpected photo relationship mapping')
            for anchor in list(drawing):
                sp=anchor.find('xdr:sp',ns)
                if sp is not None:
                    drawing.remove(anchor);removed_markings+=1;continue
                pic=anchor.find('xdr:pic',ns)
                if pic is None:continue
                blip=pic.find('.//a:blip',ns)
                if blip is None:raise ValueError('Unexpected photo shape')
                rid=blip.get('{%s}embed'%REL)
                filename=rid_targets.get(rid)
                if filename not in BASELINE_IMAGES:
                    raise ValueError('Unexpected image anchor')
                if filename not in embedded:
                    drawing.remove(anchor);removed_drawings+=1
            if removed_markings!=3 or removed_drawings!=3-len(embedded):
                raise ValueError('Original drawing structure does not match the vetted baseline')
            remove_parts=set('xl/media/'+filename for filename in BASELINE_IMAGES if filename not in embedded)
            for r in list(relationships):
                if rid_targets.get(r.get('Id')) in {posixpath.basename(n) for n in remove_parts}:
                    relationships.remove(r)
            changes['xl/drawings/drawing2.xml']=encode_xml(drawing)
            changes['xl/drawings/_rels/drawing2.xml.rels']=encode_xml(relationships)
            for name, data in embedded.items():changes['xl/media/'+name]=data
        elif embedded:
            s2=xml_root(sheet2)
            c,added=place_blank_photos(z,names,plan,s2,embedded)
            changes.update(c);sheet2=encode_xml(s2)
        remove_calc={'xl/calcChain.xml'} & names
        remove_parts|=remove_calc
        # Rebuild calculation metadata: removed formula caches must be recalculated.
        workbook=xml_root(z.read('xl/workbook.xml'))
        calc=workbook.find('m:calcPr',ns)
        if calc is None:
            calc=ET.Element(Q('calcPr'))
            following={'oleSize','customWorkbookViews','pivotCaches','smartTagPr',
                       'smartTagTypes','webPublishing','fileRecoveryPr',
                       'webPublishObjects','extLst'}
            insert_at=next((i for i,node in enumerate(workbook)
                           if node.tag.split('}')[-1] in following),len(workbook))
            workbook.insert(insert_at,calc)
        calc.set('fullCalcOnLoad','1')
        calc.set('forceFullCalc','1')
        calc.set('calcMode','auto')
        if 'xl/_rels/workbook.xml.rels' in names:
            wb_rels=xml_root(z.read('xl/_rels/workbook.xml.rels'))
            for rel in list(wb_rels):
                if 'calcChain' in (rel.get('Target') or ''):
                    wb_rels.remove(rel)
        else:wb_rels=None
        if '[Content_Types].xml' in names:
            content_types=xml_root(z.read('[Content_Types].xml'))
            for override in list(content_types):
                if override.get('PartName')=='/xl/calcChain.xml':content_types.remove(override)
            if not baseline and embedded:
                if not any(d.get('Extension','').lower()=='jpeg' for d in content_types.findall('{%s}Default'%CT_NS)):
                    content_types.insert(0,ET.Element('{%s}Default'%CT_NS,Extension='jpeg',ContentType='image/jpeg'))
                if EP.GEN_DRAWING in added:
                    ET.SubElement(content_types,'{%s}Override'%CT_NS,PartName='/'+EP.GEN_DRAWING,ContentType=DRAWING_CT)
        else:content_types=None
        changes.update({
            'xl/workbook.xml':encode_xml(workbook),
            'xl/worksheets/sheet1.xml':sheet1,'xl/worksheets/sheet2.xml':sheet2,
            'xl/sharedStrings.xml':shared,
        })
        if wb_rels is not None:changes['xl/_rels/workbook.xml.rels']=encode_xml(wb_rels)
        if content_types is not None:changes['[Content_Types].xml']=encode_xml(content_types)
        if 'docProps/core.xml' in names:
            changes['docProps/core.xml']=scrub_doc_properties(z.read('docProps/core.xml'),True)
        if 'docProps/app.xml' in names:
            changes['docProps/app.xml']=scrub_doc_properties(z.read('docProps/app.xml'),False)
        if set(added)&names:raise ValueError('SAFETY STOP: 作る部品と同じ名前の部品がひな形にあります')
        buffer=io.BytesIO()
        with zipfile.ZipFile(buffer,'w') as outzip:
            for entry in z.infolist():
                if entry.filename in remove_parts:continue
                outzip.writestr(entry,changes.get(entry.filename,z.read(entry.filename)))
            for name in sorted(added):
                outzip.writestr(name,added[name],compress_type=zipfile.ZIP_DEFLATED)
    result=buffer.getvalue()
    log={
        'photo_mode':plan['mode'],'photo_slots':len(plan['slots']),
        'changed_parts':sorted(changes),'added_parts':sorted(added),'deleted_parts':sorted(remove_parts),
        'cleared_formula_caches':fc1+fc2,'cleared_orphan_shared_strings':orphans,
        'removed_old_markings':removed_markings,'removed_unused_photo_anchors':removed_drawings}
    if approval is None:
        with zipfile.ZipFile(io.BytesIO(result)) as produced:
            check_company_report_package(produced,'出力')
    else:
        allowed_report=set(CLEAR_REPORT)|set(report_updates)
        allowed_photo=clear_photo_cells(plan['slots'])|set(photo_updates)
        inputs=[str(v) for v in list(report_updates.values())+list(photo_updates.values()) if v not in (None,'')]
        log['part_verification']=verify_output_against_approval(
            template,result,approval,allowed_report,allowed_photo,inputs,plan=plan,photo_count=len(embedded))
    return result,log


# --- 2.3.3: verify an output package against a SIGNED approval manifest. ---
# Every rule below is independent of the code that produced the output: it
# recomputes what the engine is allowed to have done and compares.

def _astop(message):
    raise ValueError('APPROVAL STOP: '+message)


def _c14n(el):
    # exclusive C14N: namespace declarations inherited from the document do not
    # count as a difference, only the element's own content does.
    return ET.tostring(el, method='c14n', exclusive=True)


def _text_of_cell(cell, shared):
    inline=cell.find('m:is',ns)
    if inline is not None:return ''.join(inline.itertext())
    v=cell.find('m:v',ns)
    if cell.get('t')=='s' and v is not None and v.text is not None:
        idx=int(v.text)
        if not 0<=idx<len(shared):_astop('壊れた共有文字列の参照があります')
        return ''.join(shared[idx].itertext())
    if cell.get('t') in ('str','inlineStr') and v is not None:return v.text or ''
    return ''


def _check_sheet(tpl_bytes, out_bytes, allowed, label, added_drawing_rid=None):
    t=xml_root(tpl_bytes);o=xml_root(out_bytes)
    if dict(t.attrib)!=dict(o.attrib):_astop(label+': シートの設定が変わっています')
    def is_added(x):
        return (added_drawing_rid is not None and x.tag==Q('drawing') and
                x.get('{%s}id'%REL)==added_drawing_rid and not len(x) and len(x.attrib)==1)
    if added_drawing_rid is not None:
        if t.find('m:drawing',ns) is not None or sum(1 for x in o if is_added(x))!=1:
            _astop(label+': 写真の図形配置への参照が想定と違います')
    if [_c14n(x) for x in t if x.tag!=Q('sheetData')]!=[_c14n(x) for x in o if x.tag!=Q('sheetData') and not is_added(x)]:
        _astop(label+': セル以外の設定（結合・印刷範囲・入力規則など）が変わっています')
    # Check the entire sheetData structure, not just previously existing rows/cells.
    # An extra blank row or a duplicate cell could otherwise escape the dict-based
    # comparison below and alter layout or hide data in the XLSX package.
    td=t.find('m:sheetData',ns);od=o.find('m:sheetData',ns)
    if td is None or od is None:_astop(label+': sheetDataがありません')
    if dict(td.attrib)!=dict(od.attrib):_astop(label+': sheetDataの設定が変わっています')
    if [_c14n(x) for x in td if x.tag!=Q('row')] != [_c14n(x) for x in od if x.tag!=Q('row')]:
        _astop(label+': sheetData内の想定外の要素が変わっています')
    tlist=td.findall('m:row',ns);olist=od.findall('m:row',ns)
    def row_map(rows,where):
        keys=[r.get('r') for r in rows]
        if any(k is None or not k.isdigit() or int(k)<1 for k in keys) or len(keys)!=len(set(keys)):
            _astop(label+f': {where}の行番号に欠落・重複があります')
        return {r.get('r'):r for r in rows}
    trows=row_map(tlist,'原本');orows=row_map(olist,'出力')
    # 2.4.1: rows must stay in ascending order (Excel repairs/misreads reordered rows).
    okeys=[int(r.get('r')) for r in olist]
    if okeys!=sorted(okeys):_astop(label+': 行の並び順が変わっています')
    for key,trow in trows.items():
        orow=orows.get(key)
        if orow is None or dict(orow.attrib)!=dict(trow.attrib):
            _astop(label+f': 行の設定が変わっています（{key}行）')
        if [_c14n(x) for x in trow if x.tag!=Q('c')] != [_c14n(x) for x in orow if x.tag!=Q('c')]:
            _astop(label+f': 行内のセル以外の要素が変わっています（{key}行）')
    for key,orow in orows.items():
        if key in trows:continue
        children=list(orow)
        if dict(orow.attrib)!={'r':key} or not children or any(x.tag!=Q('c') for x in children):
            _astop(label+f': 未承認の行が追加されました（{key}行）')
        if any(c.get('r') not in allowed for c in children):
            _astop(label+f': 新しい行に許可されていないセルがあります（{key}行）')
    def cell_map(rows,where):
        result={}
        for row in rows:
            rowid=row.get('r')
            for cell in row.findall('m:c',ns):
                addr=cell.get('r')
                if not addr or not re.fullmatch(r'[A-Z]+[1-9][0-9]*',addr):
                    _astop(label+f': {where}に不正なセル番地があります')
                if int(re.search(r'[0-9]+$',addr).group())!=int(rowid) or addr in result:
                    _astop(label+f': {where}に行違い・重複セルがあります（{addr}）')
                result[addr]=cell
        return result
    tc=cell_map(tlist,'原本');oc=cell_map(olist,'出力')
    changed=[]
    for addr in sorted(set(tc)|set(oc)):
        a=tc.get(addr);b=oc.get(addr)
        if a is not None and b is not None and _c14n(a)==_c14n(b):continue
        if addr in allowed:
            changed.append(addr);continue
        if a is not None and b is not None and a.find('m:f',ns) is not None:
            stripped=copy.deepcopy(a)
            for v in stripped.findall('m:v',ns):stripped.remove(v)
            stripped.attrib.pop('t',None)
            if _c14n(stripped)==_c14n(b):continue
        _astop(label+f': 変更が許可されていないセルが変わりました（{addr}）')
    return changed


def _without(root, predicate):
    root=copy.deepcopy(root)
    for child in list(root):
        if predicate(child):root.remove(child)
    return root


def _blank_props(xml_bytes, fields):
    root=xml_root(xml_bytes)
    for el in root.iter():
        if el.tag.split('}')[-1] in fields:
            el.text=''
            for nested in list(el):el.remove(nested)
    return root


BLANK_PICTURE_TAGS={'{%s}%s'%(DML,t) for t in ('oneCellAnchor','from','col','colOff','row','rowOff','ext','pic','nvPicPr',
                     'cNvPr','cNvPicPr','blipFill','spPr','clientData')}|{'{%s}%s'%(DRAW,t) for t in (
                     'picLocks','blip','stretch','fillRect','xfrm','off','ext','prstGeom','avLst')}


def expected_blank_additions(plan, photo_count):
    """Parts the engine may create for a blank form with this many photos (and nothing else)."""
    if not photo_count:return set()
    added={EP.GEN_MEDIA.format(i) for i in range(1,photo_count+1)}
    if plan['photo_drawing']:
        if not plan['photo_drawing_rels']:added.add(FP.rels_name(plan['photo_drawing']))
    else:
        added|={EP.GEN_DRAWING,EP.GEN_DRAWING_RELS}
        if not plan['photo_sheet_rels']:added.add(EP.SHEET_PHOTO_RELS)
    return added


def _verify_blank_photos(a, b, plan, photo_count):
    """Independent check of the pictures placed on a blank form. Returns the added drawing r:id (or None)."""
    if not photo_count:return None
    drawing=plan['photo_drawing'] or EP.GEN_DRAWING
    drels=plan['photo_drawing_rels'] or FP.rels_name(drawing)
    od=xml_root(b.read(drawing));orels=list(xml_root(b.read(drels)))
    tchildren=list(xml_root(a.read(drawing))) if plan['photo_drawing'] else []
    trels=list(xml_root(a.read(plan['photo_drawing_rels']))) if plan['photo_drawing_rels'] else []
    if plan['photo_drawing'] and dict(xml_root(a.read(drawing)).attrib)!=dict(od.attrib):
        _astop('写真帳の図形配置の設定が変わっています')
    ochildren=list(od)
    if [_c14n(x) for x in tchildren]!=[_c14n(x) for x in ochildren[:len(tchildren)]]:
        _astop('写真帳の元の図形が変わりました')
    if [_c14n(x) for x in trels]!=[_c14n(x) for x in orels[:len(trels)]]:
        _astop('写真帳の元の図形の参照が変わりました')
    new=ochildren[len(tchildren):];newrels=orels[len(trels):]
    if len(new)!=photo_count or len(newrels)!=photo_count:
        _astop('写真の図の数が写真の枚数と違います')
    rid_number={}
    for i,r in enumerate(newrels,1):
        if (r.get('Type')!=REL+'/image' or r.get('TargetMode') is not None or
                r.get('Target')!='../media/'+posixpath.basename(EP.GEN_MEDIA.format(i))):
            _astop('写真の図の参照先が想定と違います')
        rid_number[r.get('Id')]=i
    X=lambda t:'{%s}%s'%(DML,t)
    for i,anchor in enumerate(new,1):
        if anchor.tag!=X('oneCellAnchor') or [c.tag for c in anchor]!=[X('from'),X('ext'),X('pic'),X('clientData')]:
            _astop('写真の図の形が想定と違います')
        if any(el.tag not in BLANK_PICTURE_TAGS for el in anchor.iter()):
            _astop('写真の図に想定外の要素（リンク・説明文など）があります')
        cnv=anchor.find('.//xdr:cNvPr',ns)
        if set(cnv.attrib)!={'id','name'} or cnv.get('name')!=f'写真 {i}':
            _astop('写真の図の名前が想定と違います')
        blip=anchor.find('.//a:blip',ns)
        if set(blip.attrib)!={'{%s}embed'%REL} or rid_number.get(blip.get('{%s}embed'%REL))!=i:
            _astop('写真の図が別の画像を指しています')
        slot=plan['slots'][i-1]
        row=int(anchor.find('xdr:from/xdr:row',ns).text)+1;col=int(anchor.find('xdr:from/xdr:col',ns).text)+1
        if not (slot['rows'][0]<=row<=slot['rows'][1] and slot['cols'][0]<=col<=slot['cols'][1]):
            _astop(f'写真{i}が写真枠{i}の外に置かれています')
    if plan['photo_drawing']:return None
    # a new drawing: the photo sheet got exactly one new relationship to it
    srels_t=list(xml_root(a.read(EP.SHEET_PHOTO_RELS))) if plan['photo_sheet_rels'] else []
    srels_o=list(xml_root(b.read(EP.SHEET_PHOTO_RELS)))
    if [_c14n(x) for x in srels_t]!=[_c14n(x) for x in srels_o[:len(srels_t)]] or len(srels_o)!=len(srels_t)+1:
        _astop('写真シートの参照が想定外に変わりました')
    r=srels_o[-1]
    if r.get('Type')!=FP.DRAWING_TYPE or r.get('TargetMode') is not None or r.get('Target')!='../drawings/'+posixpath.basename(EP.GEN_DRAWING):
        _astop('写真シートの参照先が想定と違います')
    return r.get('Id')


def verify_output_against_approval(template, output_bytes, approval, allowed_report, allowed_photo, input_texts,
                                   plan=None, photo_count=None):
    listed={p['name']:p for p in approval['parts']}
    with zipfile.ZipFile(template) as a, zipfile.ZipFile(io.BytesIO(output_bytes)) as b:
        zip_guard(b)
        if plan is None:plan=approval.get('_plan') or FP.analyze_zip(a)
        baseline=plan['mode']=='baseline'
        tnames=set(a.namelist());onames=set(b.namelist())
        if tnames!=set(listed):_astop('原本の部品構成が承認済み一覧と違います')
        added=onames-tnames
        if baseline:
            if added-set(EP.GENERATE_PARTS):
                _astop('承認されていない部品が出力に追加されました: '+', '.join(sorted(added)[:5]))
        else:
            media=sorted(n for n in added if re.fullmatch(r'xl/media/hb_photo\d+\.jpeg',n))
            count=len(media) if photo_count is None else photo_count
            if added!=expected_blank_additions(plan,count):
                _astop('承認されていない部品が出力に追加された、または写真の部品が足りません: '+', '.join(sorted(added)[:5]))
            photo_count=count
        deleted=tnames-onames
        bad_delete=deleted-set(EP.DELETE_PARTS)-(set(EP.PHOTO_PARTS) if baseline else set())
        if bad_delete:_astop('削除が許可されていない部品が消えました: '+', '.join(sorted(bad_delete)[:5]))
        immutable_ok=0
        for name in sorted(tnames&onames):
            cat=EP.category_for(name,plan)
            if cat=='modifiable':continue
            if hashlib.sha256(b.read(name)).hexdigest()!=listed[name]['sha256']:
                _astop('変更禁止の部品が変わりました: '+name)
            immutable_ok+=1
        added_rid=None if baseline else _verify_blank_photos(a,b,plan,photo_count)
        # --- modifiable parts, each against its own rule ---
        cells_r=_check_sheet(a.read(EP.SHEET_REPORT),b.read(EP.SHEET_REPORT),allowed_report,'報告書シート')
        cells_p=_check_sheet(a.read(EP.SHEET_PHOTO),b.read(EP.SHEET_PHOTO),allowed_photo,'写真帳シート',added_rid)
        # shared strings: same count, no reindexing; unreferenced entries must be empty
        ts_root=xml_root(a.read('xl/sharedStrings.xml'));os_root=xml_root(b.read('xl/sharedStrings.xml'))
        if dict(ts_root.attrib)!=dict(os_root.attrib):_astop('共有文字列の設定が変わっています')
        ts=ts_root.findall('m:si',ns);os_=os_root.findall('m:si',ns)
        if len(ts)!=len(os_):_astop('共有文字列の数が変わっています')
        used=set()
        sheet_roots={}
        for n in sorted(x for x in onames if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',x)):
            r=xml_root(b.read(n));sheet_roots[n]=r
            for c in r.findall('.//m:sheetData/m:row/m:c',ns):
                v=c.find('m:v',ns)
                if c.get('t')=='s' and v is not None and v.text:used.add(int(v.text))
        blanked=0
        for i,(x,y) in enumerate(zip(ts,os_)):
            ytext=''.join(y.itertext())
            if i not in used and ytext.strip():
                _astop('使われていない共有文字列に内容が残っています（過去の案件の文字の可能性）')
            if _c14n(x)!=_c14n(y):
                if i in used:_astop('使われている共有文字列が書き換わりました')
                blanked+=1
        # drawing of the photo sheet (baseline): only marking shapes and unused photo anchors removed
        kept={posixpath.basename(n) for n in EP.PHOTO_PARTS if n in onames} if baseline else set()
        if baseline:
            trels=xml_root(a.read('xl/drawings/_rels/drawing2.xml.rels'))
            rid_to_file={r.get('Id'):posixpath.basename(r.get('Target','')) for r in trels}
            def drop_anchor(anchor):
                if anchor.find('xdr:sp',ns) is not None:return True
                blip=anchor.find('.//a:blip',ns)
                if blip is None:return False
                return rid_to_file.get(blip.get('{%s}embed'%REL)) not in kept
            if _c14n(_without(xml_root(a.read('xl/drawings/drawing2.xml')),drop_anchor))!=_c14n(xml_root(b.read('xl/drawings/drawing2.xml'))):
                _astop('写真帳の図形が想定外に変わりました')
            removed_files={posixpath.basename(n) for n in EP.PHOTO_PARTS}-kept
            if _c14n(_without(trels,lambda r:posixpath.basename(r.get('Target','')) in removed_files))!=_c14n(xml_root(b.read('xl/drawings/_rels/drawing2.xml.rels'))):
                _astop('写真帳の図形の参照が想定外に変わりました')
        # workbook.xml: only calcPr may differ, with exactly these attributes added
        tw=xml_root(a.read('xl/workbook.xml'));ow=xml_root(b.read('xl/workbook.xml'))
        tcalc=tw.find('m:calcPr',ns);ocalc=ow.find('m:calcPr',ns)
        if ocalc is None:_astop('再計算の指定がありません')
        expect=dict(tcalc.attrib) if tcalc is not None else {}
        expect.update({'fullCalcOnLoad':'1','forceFullCalc':'1','calcMode':'auto'})
        if dict(ocalc.attrib)!=expect:_astop('再計算の指定以外の計算設定が変わっています')
        is_calc=lambda e:e.tag==Q('calcPr')
        if _c14n(_without(tw,is_calc))!=_c14n(_without(ow,is_calc)):_astop('ブック設定が想定外に変わりました')
        # workbook rels and content types: only the calcChain entry removed
        part='xl/_rels/workbook.xml.rels'
        if part in tnames:
            if _c14n(_without(xml_root(a.read(part)),lambda r:'calcChain' in (r.get('Target') or '')))!=_c14n(xml_root(b.read(part))):
                _astop(part+' が想定外に変わりました（外部リンク等の参照は保持が必要）')
        part='[Content_Types].xml'
        if part in tnames:
            tct=_without(xml_root(a.read(part)),lambda o:o.get('PartName')=='/xl/calcChain.xml')
            oct_=xml_root(b.read(part))
            def addition(el):
                if baseline or not photo_count:return False
                if el.tag=='{%s}Default'%CT_NS and dict(el.attrib)=={'Extension':'jpeg','ContentType':'image/jpeg'}:
                    return not any(d.get('Extension','').lower()=='jpeg' for d in tct.findall('{%s}Default'%CT_NS))
                return (EP.GEN_DRAWING in added and el.tag=='{%s}Override'%CT_NS and
                        dict(el.attrib)=={'PartName':'/'+EP.GEN_DRAWING,'ContentType':DRAWING_CT})
            if _c14n(tct)!=_c14n(_without(oct_,addition)):
                _astop(part+' が想定外に変わりました（外部リンク等の参照は保持が必要）')
        # document properties: only the listed fields emptied
        legacy_props=[]
        for part,fields in (('docProps/core.xml',EP.DOCPROP_CORE_BLANK),('docProps/app.xml',EP.DOCPROP_APP_BLANK)):
            if part in tnames:
                for el in xml_root(a.read(part)).iter():
                    if el.tag.split('}')[-1] in fields and (el.text or '').strip():legacy_props.append(el.text.strip())
                if _c14n(_blank_props(a.read(part),fields))!=_c14n(xml_root(b.read(part))):
                    _astop(part+' の作成者などが消えていない、または他の文書情報が変わりました')
        # photo slots: new JPEG, same frame size, no EXIF
        photos_checked=0
        if not baseline:
            for n in sorted(x for x in added if x.startswith('xl/media/')):
                with Image.open(io.BytesIO(b.read(n))) as im:
                    if im.format!='JPEG' or max(im.size)>1600:_astop('写真の画像の形式・大きさが違います: '+n)
                    if len(im.getexif()) or 'exif' in im.info:_astop('写真に撮影情報（EXIF）が残っています: '+n)
                photos_checked+=1
        for n in (EP.PHOTO_PARTS if baseline else ()):
            if n not in onames:continue
            with Image.open(io.BytesIO(a.read(n))) as im:tsize=im.size
            if max(tsize)>1600:
                ratio=1600/max(tsize);tsize=tuple(max(1,round(v*ratio)) for v in tsize)
            with Image.open(io.BytesIO(b.read(n))) as im:
                if im.format!='JPEG' or im.size!=tuple(tsize):_astop('写真枠の画像の形式・大きさが違います: '+n)
                if len(im.getexif()) or 'exif' in im.info:_astop('写真に撮影情報（EXIF）が残っています: '+n)
            photos_checked+=1
        # residue: text that sat in the cleared cells of the template must not survive anywhere active
        legacy=set()
        tshared=ts
        for part,allowed in ((EP.SHEET_REPORT,allowed_report),(EP.SHEET_PHOTO,allowed_photo)):
            for c in xml_root(a.read(part)).findall('m:sheetData/m:row/m:c',ns):
                if c.get('r') in allowed:
                    t=_text_of_cell(c,tshared).strip()
                    if len(t)>=4:legacy.add(t)
        legacy.update(x for x in legacy_props if len(x)>=2)
        # Text that also sits in the template's KEPT, non-formula cells (labels, the
        # hidden dropdown lists such as （第１報）/故　障　/【当社の社員】) is part of the
        # form itself, not incident residue. Formula caches do not count: they are removed.
        kept_static=[]
        for part,allowed in ((EP.SHEET_REPORT,allowed_report),(EP.SHEET_PHOTO,allowed_photo)):
            for c in xml_root(a.read(part)).findall('m:sheetData/m:row/m:c',ns):
                if c.get('r') not in allowed and c.find('m:f',ns) is None:
                    t=_text_of_cell(c,tshared).strip()
                    if t:kept_static.append(t)
        legacy={m for m in legacy if not any(m in t for t in kept_static)}
        joined_inputs='\n'.join(input_texts)
        legacy={m for m in legacy if m not in joined_inputs}
        oshared=os_
        active=[]
        for part in (EP.SHEET_REPORT,EP.SHEET_PHOTO):
            for c in sheet_roots[part].findall('m:sheetData/m:row/m:c',ns):active.append(_text_of_cell(c,oshared))
        for part in ('docProps/core.xml','docProps/app.xml'):
            if part in onames:active.append(' '.join(xml_root(b.read(part)).itertext()))
        for m in legacy:
            if any(m in t for t in active):
                _astop('過去の案件の文字列が出力に残っています（消去対象外のセル、または文書情報）')
    return {'approval_id':approval['approval']['approval_id'],
            'immutable_parts_byte_identical':immutable_ok,
            'photo_mode':plan['mode'],
            'modifiable_parts_checked':sorted(n for n in tnames&onames if EP.category_for(n,plan)=='modifiable'),
            'deleted_parts':sorted(deleted),'added_parts':sorted(added),
            'changed_cells_report':len(cells_r),'changed_cells_photo':len(cells_p),
            'blanked_shared_strings':blanked,'photos_checked':photos_checked,
            'legacy_markers_checked':len(legacy)}


def verify_package(source, output_bytes, photo_count, expected,old_markers,plan=None):
    with zipfile.ZipFile(source) as a,zipfile.ZipFile(io.BytesIO(output_bytes)) as b:
        zip_guard(b)
        if plan is None:plan=FP.analyze_zip(a)
        for part in a.namelist():
            if part.lower().endswith('.emf') or '/printerSettings/' in part:
                if b.read(part)!=a.read(part):raise ValueError('Preservation failure: '+part)
        if plan['mode']=='baseline':
            rel=b.read('xl/drawings/_rels/drawing2.xml.rels')
            draw=b.read('xl/drawings/drawing2.xml')
            if len(xml_root(draw).findall('.//xdr:pic',ns))!=photo_count:raise ValueError('Photo anchor count incorrect')
            if len(xml_root(rel))!=photo_count:raise ValueError('Photo reference count incorrect')
        elif photo_count:
            draw=xml_root(b.read(plan['photo_drawing'] or EP.GEN_DRAWING))
            before=len(xml_root(a.read(plan['photo_drawing'])).findall('.//xdr:pic',ns)) if plan['photo_drawing'] else 0
            if len(draw.findall('.//xdr:pic',ns))-before!=photo_count:raise ValueError('Photo anchor count incorrect')
        sheet=xml_root(b.read('xl/worksheets/sheet1.xml'))
        pic_sheet=xml_root(b.read('xl/worksheets/sheet2.xml'))
        def text_at(root,addr):
            c=root.find(f".//m:sheetData/m:row/m:c[@r='{addr}']",ns)
            if c is None:return None
            inline=c.find('m:is',ns)
            if inline is not None:return ''.join(inline.itertext())
            v=c.find('m:v',ns)
            return v.text if v is not None else None
        checks=[(sheet,'G6',expected['title']),(sheet,'AP6',expected['location']),(pic_sheet,FP.PHOTO_TITLE[0],expected['title'])]
        checks+=[(pic_sheet,slot['number'],str(i+1)) for i,slot in enumerate(plan['slots'][:photo_count])]
        for root,addr,val in checks:
            if text_at(root,addr)!=val:raise ValueError('Output verification failed: '+addr)
        for xml in (sheet,pic_sheet):
            for c in xml.findall('.//m:sheetData/m:row/m:c',ns):
                f=c.find('m:f',ns)
                if f is not None and c.find('m:v',ns) is not None:
                    raise ValueError('Stale formula cache remains')
        # Check both inline strings and actively referenced shared strings.
        # Unreferenced shared-string entries have already been blanked.
        shared=xml_root(b.read('xl/sharedStrings.xml')).findall('m:si',ns)
        active=[]
        for root in (sheet,pic_sheet):
            for c in root.findall('.//m:sheetData/m:row/m:c',ns):
                inline=c.find('m:is',ns)
                if inline is not None:active.append(''.join(inline.itertext()))
                elif c.get('t')=='s':
                    v=c.find('m:v',ns)
                    if v is not None and v.text is not None:
                        idx=int(v.text)
                        if not 0<=idx<len(shared):raise ValueError('Broken shared-string reference')
                        active.append(''.join(shared[idx].itertext()))
        source_and_input=' '.join(map(str,expected.values()))
        for legacy in old_markers:
            if legacy and len(legacy)>=7 and legacy not in source_and_input:
                if any(legacy in item for item in active):
                    raise ValueError('Legacy incident text is still active')
        return {'zip_crc':'ok','photo_anchors':photo_count,'emf_preserved':True,'printer_settings_preserved':True,
                'old_AP9_cleared':True,'formula_caches_cleared':True}


def legacy_markers_from_template(template):
    marked=('G6','G15','G16','G17','G18','G20','G21','G26','C28','N32','AP9')
    with zipfile.ZipFile(template) as z:
        root=xml_root(z.read('xl/worksheets/sheet1.xml'))
        shared=xml_root(z.read('xl/sharedStrings.xml')).findall('m:si',ns)
        result=[]
        for a in marked:
            c=root.find(f".//m:sheetData/m:row/m:c[@r='{a}']",ns)
            if c is None:continue
            v=c.find('m:v',ns)
            if c.get('t')=='s' and v is not None:
                result.append(''.join(shared[int(v.text)].itertext()).strip())
            elif c.get('t')=='inlineStr':
                inline=c.find('m:is',ns)
                if inline is not None:result.append(''.join(inline.itertext()).strip())
        return result


def write_new_file(path, data):
    """Create-new semantics; refuse overwrite even if path appeared after validation.
    The work is completely prepared and validated before any output is published.
    """
    # Publishing to Android shared storage may not support hardlinks. Write an
    # O_EXCL file only after all validations, flush, then check the size; on
    # ordinary exceptions delete it. Power loss may still leave a partial file.
    fd=os.open(str(path), os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
    try:
        with os.fdopen(fd,'wb') as handle:
            handle.write(data)
            handle.flush()
            os.fsync(handle.fileno())
        if path.stat().st_size!=len(data):raise IOError('Output size mismatch')
    except BaseException:
        try:path.unlink()
        except Exception:pass
        raise


def _program_hashes():
    here=Path(__file__).resolve().parent
    out={}
    for n in ('converter_mobile.py','approval.py','engine_policy.py','case_exports.py'):
        f=here/n
        if f.is_file():out[n]=hashlib.sha256(f.read_bytes()).hexdigest()
    return out


def make_report(field, photos, template, out, audit_path=None, case_id=None, photo_actions=None,
                approval_manifest=None, trusted_keys=None, approvals_dir=None):
    import approval as AP
    sources,output,audit=validate_path_separation(field,photos,template,out,audit_path)
    fs,zp,corp=sources
    hashes_before={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sources}
    # 2.3.3: the template must match a manifest signed by a registered approver.
    approved=AP.verify_manifest(corp,approval_manifest,trusted_keys,approvals_dir)
    case,photo_rows,timeline=read_field_tables(fs, case_id, photo_actions)
    return _finish_report(sources,hashes_before,output,audit,corp,zp,case,photo_rows,timeline,approved,
                          photo_actions,len(get_incidents(fs)),{'kind':'field-survey-xlsx'})


def make_report_from_homebox(package, template, out, audit_path=None,
                             approval_manifest=None, trusted_keys=None, approvals_dir=None):
    """2.4.0: same approved engine, input from a HomeBox report package (hb-report-package/1)."""
    import approval as AP
    import homebox_adapter as HB
    sources,output,audit=validate_path_separation(package,None,template,out,audit_path)
    pkg,corp=sources
    hashes_before={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in sources}
    approved=AP.verify_manifest(corp,approval_manifest,trusted_keys,approvals_dir)
    case,original_photos,timeline,info=HB.read_package(pkg)
    case,photo_rows,timeline=select_case_rows(case,original_photos,timeline,None)
    return _finish_report(sources,hashes_before,output,audit,corp,pkg,case,photo_rows,timeline,approved,
                          None,1,info)


def _finish_report(sources,hashes_before,output,audit,corp,zp,case,photo_rows,timeline,approved,
                   photo_actions,source_total,input_info):
    plan=approved['_plan']
    embedded,details=read_photos(zp,photo_rows,corp,photo_actions,plan)
    report_update,photo_update,layout_warnings=build_updates_with_notes(case,photo_rows,timeline,slots=plan['slots'])
    content,change_log=transform_package(corp,report_update,photo_update,embedded,approval=approved,plan=plan)
    expected={'title':safe_text(case['件名']),
              'location':safe_text(case['発生場所'])}
    verified=verify_package(corp,content,len(photo_rows),expected,legacy_markers_from_template(corp),plan)
    review={
        'status':'TEST_DRAFT_NOT_APPROVED','version':VERSION,
        'project_case_id':case['ID'],'project_title':case.get('件名'),
        'photo_count':len(photo_rows),'timeline_count':len(timeline),
        'embedded_jpeg_quality':80,'embedded_max_long_edge':1600,'time_zone_for_aware_timestamps':'Asia/Tokyo',
        'photos':details,'source_hashes_sha256':hashes_before,
        'selected_photo_row_ids':[str(r['ID']) for r in photo_rows],
        'source_incident_total':source_total,
        'input':input_info,
        'unmapped_nonempty_fields':[key for key in UNMAPPED_CASE_FIELDS if case.get(key) not in (None,'')],
        'manual_checks':['Photo content vs incident, location and description (no automatic visual validation)',
            'Incident fields, reporter identity, and previously filled company template fields',
            'Unmapped Field Survey inputs are listed in unmapped_nonempty_fields and need manual review',
            'All formula-derived second-page fields must be recalculated in desktop Excel',
            'Printed layout and existing pictures, shapes, EMF; do not submit without approval'],
        'form':FP.summary(plan),
        'known_limitations':['Only templates with a signed approval manifest of the 2026-10 form (report boxes checked, photo slots derived)','One chosen incident, photos up to the number of photo slots, report rows as laid out',
            'No direct Field Survey annotations drawn on images','Unmapped questionnaire columns require manual review',
            'Photo source is retained in the original photos ZIP; embedded images are downscaled'],
        'checks':verified,'ooxml':change_log,
        'layout_version':__import__('report_layout').LAYOUT_VERSION,'layout_warnings':layout_warnings,
        'approval':{k:approved['approval'].get(k) for k in ('approval_id','approver','approver_key_id','approved_at','valid_until')},
        'approval_manifest':{'file':approved['_manifest_file'],'sha256':approved['_manifest_sha256'],
                             'template_sha256':approved['template']['sha256'],
                             'trusted_keys_sha256':approved['_trusted_sha256']},
        'program':{'version':VERSION,'engine_policy_version':EP.ENGINE_POLICY_VERSION,'file_sha256':_program_hashes()},
        'part_verification':change_log.get('part_verification'),
    }
    if [hashlib.sha256(p.read_bytes()).hexdigest() for p in sources]!=list(hashes_before.values()):
        raise RuntimeError('Source file changed while processing')
    if audit:
        write_new_file(audit,json.dumps(review,ensure_ascii=False,indent=2).encode('utf-8'))
    try:write_new_file(output,content)
    except BaseException:
        if audit:
            try:audit.unlink()
            except Exception:pass
        raise
    return review


if __name__=='__main__':
    p=argparse.ArgumentParser(description='Field Survey verified TEST DRAFT converter '+VERSION)
    p.add_argument('--field',required=True)
    p.add_argument('--photos',required=True)
    p.add_argument('--template',required=True)
    p.add_argument('--out',required=True)
    p.add_argument('--audit')
    p.add_argument('--case-id')
    p.add_argument('--approval',help='承認済み一覧（省略時は approvals/<原本SHA-256>.approved.json）')
    p.add_argument('--trusted',help='承認者の公開鍵一覧（省略時は trusted_approvers.json）')
    args=p.parse_args()
    try:
        result=make_report(args.field,args.photos,args.template,args.out,args.audit,case_id=args.case_id,
                           approval_manifest=args.approval,trusted_keys=args.trusted)
        print(json.dumps(result,ensure_ascii=False,indent=2))
    except Exception as exc:
        print('SAFETY STOP:',type(exc).__name__,str(exc),file=sys.stderr)
        sys.exit(1)
