#!/usr/bin/env python3
# coding: utf-8
"""Produce incident-titled, case-filtered Field Survey XLSX + photo ZIP copies.

Field Survey's own Android export name is unchanged: these copies are created
only after export. No input is renamed, deleted, or overwritten.
"""
from __future__ import annotations
import datetime
import hashlib
import io
import os
import re
import unicodedata
import uuid
import zipfile
from pathlib import Path
from lxml import etree as ET
from converter_mobile import (
    get_incidents, get_related_rows, rows_from_sheet, zip_guard,
    xml_root, encode_xml, scrub_doc_properties, scrub_orphan_shared_strings, ns, SS, REL,
)

SHEET_NAMES = ('異常案件', '対応経過', '写真帳')
RESERVED = {'CON','PRN','AUX','NUL', *(f'COM{i}' for i in range(1,10)), *(f'LPT{i}' for i in range(1,10))}
M='{'+SS+'}'

def filename_subject(raw: object, limit: int = 48) -> str:
    text = unicodedata.normalize('NFC', str(raw or '')).strip()
    # Control, invisible bidi, Windows-forbidden and path separators are removed.
    text = ''.join(' ' if unicodedata.category(ch).startswith('C') or ch in '<>:"/\\|?*' else ch for ch in text)
    text = re.sub(r'\s+', ' ', text).strip(' .')
    if not text: text = '件名未設定'
    if text.split('.')[0].upper() in RESERVED: text='案件_'+text
    text=text[:limit].rstrip(' .')
    return text or '件名未設定'

def case_source_names(title: str, stamp: str, unique: str):
    if not re.fullmatch(r'\d{12}|手動選択', stamp):
        raise ValueError('書き出し日時の形式が不正です')
    if stamp=='手動選択':stamp=datetime.datetime.now().strftime('%Y%m%d%H%M')
    if not re.fullmatch(r'[0-9a-f]{8}',unique):raise ValueError('識別子が不正です')
    base=f'{filename_subject(title)}_FieldSurvey_{unique}'
    return f'{base}_{stamp}.xlsx', f'{base}_photos_{stamp}.zip'

def output_report_name(title: str, stamp: str, unique: str):
    if stamp=='手動選択':stamp=datetime.datetime.now().strftime('%Y%m%d%H%M')
    if not re.fullmatch(r'\d{12}',stamp):raise ValueError('書き出し日時の形式が不正です')
    if not re.fullmatch(r'[0-9a-f]{8}',unique):raise ValueError('識別子が不正です')
    return f'{filename_subject(title)}_会社書式_要確認_{stamp}_{unique}.xlsx'

def _sheet_path_map(z: zipfile.ZipFile):
    wb=xml_root(z.read('xl/workbook.xml'))
    rels=xml_root(z.read('xl/_rels/workbook.xml.rels'))
    by_id={x.get('Id'): x.get('Target') for x in rels}
    import posixpath
    mapping={}
    for sh in wb.findall('.//m:sheet',ns):
        target=by_id[sh.get('{'+REL+'}id')]
        if target.startswith('/'): path=target.lstrip('/')
        else: path=posixpath.normpath(posixpath.join('xl',target))
        if not path.startswith('xl/worksheets/') or path not in z.namelist():
            raise ValueError('未知のシート構成を検出')
        mapping[sh.get('name')]=path
    if set(mapping)!=set(SHEET_NAMES):
        raise ValueError('Field Surveyの3シート構造が想定と異なります')
    return mapping

def _filtered_sheet(xml:bytes, keep_record_rows:set[int]):
    root=xml_root(xml)
    data=root.find('m:sheetData',ns)
    if data is None:raise ValueError('sheetDataがありません')
    source_rows=list(data.findall('m:row',ns))
    # Safe for the simple data tables exported by Field Survey: no workbook formulas.
    if any(x.find('m:f',ns) is not None for r in source_rows for x in r.findall('m:c',ns)):
        raise ValueError('数式入りのField Surveyは案件別のコピーを自動作成できません')
    chosen=[r for r in source_rows if int(r.get('r','0'))==1 or int(r.get('r','0')) in keep_record_rows]
    if not chosen or int(chosen[0].get('r','0'))!=1:
        raise ValueError('見出し行がありません')
    for r in source_rows:data.remove(r)
    for index,r in enumerate(chosen,1):
        r.set('r',str(index))
        for c in r.findall('m:c',ns):
            address=c.get('r','')
            col=re.match(r'^[A-Z]+',address)
            if col is None: raise ValueError('セル番地が不正です')
            c.set('r',col.group()+str(index))
        data.append(r)
    dim=root.find('m:dimension',ns)
    if dim is not None:
        rng=dim.get('ref','A1')
        end=re.search(r'([A-Z]+)\d+$',rng)
        if end:dim.set('ref',f'A1:{end.group(1)}{len(chosen)}')
    # These derived ranges may point to old row indexes and should not be stale.
    for tag in ('autoFilter','mergeCells','conditionalFormatting','dataValidations'):
        if root.find('m:'+tag,ns) is not None:
            raise ValueError(f'案件別コピーに未対応のシート機能があります: {tag}')
    return encode_xml(root)


def audit_no_case_data_leaks(zout_bytes:bytes, excluded_strings:set[str], allowed_strings:set[str]):
    """Best-effort active XML string detection; not an exhaustive privacy audit."""
    with zipfile.ZipFile(io.BytesIO(zout_bytes)) as z:
        for item in z.infolist():
            if not item.filename.endswith('.xml'):continue
            try:root=xml_root(z.read(item))
            except Exception:continue
            content=' '.join(root.itertext())
            for marker in excluded_strings:
                if len(marker)>=10 and marker not in allowed_strings and marker in content:
                    raise ValueError('案件外の文章が出力に残っています: '+item.filename)

def build_case_excel_bytes(source:Path,case_id:str) -> bytes:
    case,photo_rows,timeline_rows=get_related_rows(source,case_id)
    valid={
        '異常案件':{str(case_id)},
        '対応経過':{str(r['ID']) for r in timeline_rows},
        '写真帳':{str(r['ID']) for r in photo_rows},
    }
    with zipfile.ZipFile(source) as zin:
        zip_guard(zin,45*1024*1024)
        mapping=_sheet_path_map(zin)
        changes={}
        for name,path in mapping.items():
            rows=rows_from_sheet(str(source),name)
            headers=[str(c or '') for c in rows[0]]
            key_col=headers.index('ID')
            indices={i for i,r in enumerate(rows[1:],2) if str(r[key_col]) in valid[name]}
            expected=len(valid[name])
            if len(indices)!=expected:raise ValueError('対象案件の記録IDが重複・欠落しています: '+name)
            changes[path]=_filtered_sheet(zin.read(path),indices)
        # Old cases can remain as unused strings after deleting worksheet rows.
        # Scrub entries without reindexing references (safer for simple exports).
        if 'xl/sharedStrings.xml' in zin.namelist():
            changes['xl/sharedStrings.xml'],_=scrub_orphan_shared_strings(
                zin.read('xl/sharedStrings.xml'),[changes[p] for p in mapping.values()])
        # A stale thumbnail may disclose the original case: fail closed rather than guess.
        unsafe=[n for n in zin.namelist() if n.startswith(('docProps/thumbnail',
                'xl/comments','xl/embeddings','customXml/'))]
        if unsafe:raise ValueError('案件別出力に未対応の埋込情報があります: '+', '.join(unsafe[:3]))
        for file,core in (('docProps/core.xml',True),('docProps/app.xml',False)):
            if file in zin.namelist():changes[file]=scrub_doc_properties(zin.read(file),core)
        # Verify that all excluded rows were deleted. This does not rule out
        # case data embedded in other XLSX components; external review remains mandatory.
        excluded=set();allowed=set()
        for tab,path in mapping.items():
            rows=rows_from_sheet(str(source),tab)
            headers=[str(c or '') for c in rows[0]]
            id_col=headers.index('ID')
            for row in rows[1:]:
                if str(row[id_col]) in valid[tab]:allowed.update(str(x) for x in row if isinstance(x,str))
                else:excluded.update(str(x) for x in row if isinstance(x,str))
        b=io.BytesIO()
        with zipfile.ZipFile(b,'w') as zout:
            for info in zin.infolist():zout.writestr(info,changes.get(info.filename,zin.read(info.filename)))
    out=b.getvalue()
    with zipfile.ZipFile(io.BytesIO(out)) as z:zip_guard(z,45*1024*1024)
    audit_no_case_data_leaks(out, excluded, allowed)
    return out

def build_case_photo_zip_bytes(source_zip:Path,field:Path,case_id:str)->bytes:
    _,photos,_=get_related_rows(field,case_id)
    names=[str(r.get('元写真') or '') for r in photos]
    if any(not n for n in names) or len(names)!=len(set(names)):
        raise ValueError('写真のファイル名が欠落または重複しています')
    with zipfile.ZipFile(source_zip) as zin:
        zip_guard(zin)
        by_name={}
        for member in zin.infolist():
            if member.is_dir():continue
            key=Path(member.filename).name
            if key in by_name:raise ValueError('ZIP内に同じ写真名があります')
            by_name[key]=member
        missing=[n for n in names if n not in by_name]
        if missing:raise ValueError('写真ZIPに案件の写真がありません: '+', '.join(missing[:3]))
        b=io.BytesIO()
        with zipfile.ZipFile(b,'w',compression=zipfile.ZIP_STORED) as zout:
            for name in names:
                info=by_name[name]
                data=zin.read(info)
                zout.writestr(name,data)
        out=b.getvalue()
    with zipfile.ZipFile(io.BytesIO(out)) as z:zip_guard(z)
    return out

def write_exclusive(path: Path,data:bytes):
    # No replacement of existing files, even if they are not source files.
    with path.open('xb') as f:
        f.write(data)
        f.flush()
        os.fsync(f.fileno())

def save_case_sources(field, photos, case_id, stamp, output_dir:Path,unique=None):
    field=Path(field);photos=Path(photos);output_dir=Path(output_dir)
    if not output_dir.is_dir():raise ValueError('出力フォルダーがありません')
    case,_,_=get_related_rows(field,case_id)
    uid=unique or uuid.uuid4().hex[:8]
    excel_name,zip_name=case_source_names(str(case.get('件名')),stamp,uid)
    out_excel=output_dir/excel_name;out_zip=output_dir/zip_name
    # All expensive validation occurs before any output is written.
    data_excel=build_case_excel_bytes(field,case_id)
    data_zip=build_case_photo_zip_bytes(photos,field,case_id)
    created=[]
    try:
        write_exclusive(out_excel,data_excel);created.append(out_excel)
        write_exclusive(out_zip,data_zip);created.append(out_zip)
    except Exception:
        for path in created:path.unlink(missing_ok=True)
        raise
    # Confirm the new pair is parseable and contains only the requested case.
    try:
        remaining=get_incidents(out_excel)
        _,rphotos,rtimes=get_related_rows(out_excel,case_id)
        if len(remaining)!=1 or str(remaining[0]['ID'])!=str(case_id):raise ValueError('案件分離に失敗')
        with zipfile.ZipFile(out_zip) as z:
            if len(z.namelist())!=len(rphotos):raise ValueError('写真の分離に失敗')
    except Exception:
        out_excel.unlink(missing_ok=True);out_zip.unlink(missing_ok=True)
        raise
    return out_excel,out_zip
