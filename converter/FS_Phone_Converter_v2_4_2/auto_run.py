#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Field Survey v2.1 input discovery. No Android sandbox bypass.

Search accessible shared-storage folders read-only, including nested folders.
When not found, optionally use Android document picker via Termux:API.
The workbook transformer is the *unchanged* audited v2.0-test engine.
"""
from __future__ import annotations
import os, re, sys, uuid, shutil, tempfile, subprocess
from pathlib import Path
from converter_mobile import make_report, VERSION, get_incidents, get_related_rows, MAX_PHOTOS
from case_exports import output_report_name, save_case_sources

HERE=Path(__file__).resolve().parent
TEMPLATE=HERE/'company_template.xlsx'
HOME=Path.home()
DOWNLOADS=HOME/'storage'/'downloads'
DOCUMENTS=HOME/'storage'/'documents'
SHARED=HOME/'storage'/'shared'
REPORT_DIR=DOWNLOADS/'FS_変換結果'
# Scope the scan to places where a user expects exports, avoiding indiscriminate
# searching of the phone's private areas and large photo/music collections.
ROOTS=[(DOWNLOADS,4),(DOCUMENTS,4),
       (SHARED/'Android'/'media'/'net.pufstudio.fieldsurvey',5)]
XLSX=re.compile(r'^(?P<base>.+)_(?P<stamp>\d{12})(?P<copy>(?:\s*\(\d+\))*)\.xlsx$',re.I)
PHOTO=re.compile(r'^(?P<base>.+)_photos_(?P<stamp>\d{12})(?P<copy>(?:\s*\(\d+\))*)\.zip$',re.I)
SKIP_DIRS={'FS_変換結果','$RECYCLE.BIN','.thumbnails','.git'}

def scan_roots():
    custom=HERE/'scan_folders.txt'
    roots=list(ROOTS)
    if custom.is_file():
        for raw in custom.read_text(encoding='utf-8-sig').splitlines():
            path=raw.strip()
            if path and not path.startswith('#'):
                roots.append((Path(os.path.expandvars(os.path.expanduser(path))),4))
    unique=set()
    for loc,depth in roots:
        try:p=loc.resolve(strict=True)
        except (OSError,RuntimeError):continue
        if not p.is_dir() or str(p) in unique:continue
        unique.add(str(p))
        yield p,depth

def collect_files():
    excel=[];photo=[];seen=set()
    for root,depth_limit in scan_roots():
        for here,dirs,files in os.walk(root,followlinks=False):
            current=Path(here)
            relative=len(current.relative_to(root).parts)
            dirs[:]=[d for d in dirs if relative<depth_limit and d not in SKIP_DIRS and not d.startswith('.')]
            for name in files:
                m=XLSX.fullmatch(name)
                p=PHOTO.fullmatch(name)
                if not m and not p:continue
                file=current/name
                try:
                    if not file.is_file() or file.stat().st_size==0:continue
                    real=file.resolve(strict=True)
                except (OSError,RuntimeError):continue
                if str(real) in seen:continue
                seen.add(str(real))
                match=m or p
                # Preserve exact copy suffix, do not guess which duplicates belong together.
                key=(match['base'].casefold(),match['stamp'],re.sub(r'\s+','',match['copy']))
                row=(real,key,match['stamp'])
                (excel if m else photo).append(row)
    return excel,photo

def list_pairs():
    exc,zipfiles=collect_files()
    photos_by_key={}; exc_by_key={}
    for zp,key,stamp in zipfiles:photos_by_key.setdefault(key,[]).append(zp)
    for field,key,stamp in exc:exc_by_key.setdefault(key,[]).append(field)
    choices=[]
    for key,fields in exc_by_key.items():
        zs=photos_by_key.get(key,[])
        matched=set()
        # Primary rule: exported spreadsheet and photo ZIP are in one directory.
        for field in fields:
            for zp in zs:
                if field.parent==zp.parent:
                    try:mod=max(field.stat().st_mtime,zp.stat().st_mtime)
                    except OSError:continue
                    choices.append((field,zp,key[1],mod))
                    matched.add((field,zp))
        # Secondary rule: allow different directories ONLY if exactly one XLSX
        # and one ZIP share the same exact base/timestamp/copy suffix.
        if not matched and len(fields)==1 and len(zs)==1:
            field,zp=fields[0],zs[0]
            try:mod=max(field.stat().st_mtime,zp.stat().st_mtime)
            except OSError:continue
            choices.append((field,zp,key[1],mod))
    choices.sort(key=lambda x:x[3],reverse=True)
    return choices

def choose_from_pairs(choices):
    print('\n見つかった組み合わせ（新しい順）:',flush=True)
    for i,(field,photos,stamp,_) in enumerate(choices,1):
        print(f'{i}. {field.name}\n   {field.parent}\n   写真: {photos.name}',flush=True)
    if len(choices)==1:
        print('1組のみ見つかりました。選択: 1')
        return choices[0][:3]
    try:
        entry=input(f'番号 [1〜{len(choices)}]（Enter=選択せず中止 / 0=中止）: ').strip()
        if not entry:return None
        number=int(entry)
    except (EOFError,ValueError):return None
    if number==0:return None
    if not 1<=number<=len(choices):return None
    return choices[number-1][:3]

def pick_via_android():
    command=shutil.which('termux-storage-get')
    if command is None:
        print('Androidのファイル選択機能にはTermux:APIアプリと「pkg install termux-api」が必要。')
        return None
    # Temp files are inside Termux app sandbox; always removed on exit.
    temp=tempfile.TemporaryDirectory(prefix='fs_pick_')
    folder=Path(temp.name)
    fields=folder/'chosen.xlsx';photos=folder/'chosen_photos.zip'
    for text,target in [('① Field Surveyから出力したExcel (.xlsx)',fields),
                        ('② 同じ出力時刻の写真ZIP (_photos_...zip)',photos)]:
        print(text+'をAndroidのファイル選択画面で選択してください。',flush=True)
        try:
            result=subprocess.run([command,str(target)],check=False)
        except (OSError,KeyboardInterrupt):
            temp.cleanup();return None
        if result.returncode or not target.is_file() or target.stat().st_size==0:
            print('選択に失敗したかキャンセルされました。')
            temp.cleanup();return None
    # Original filenames lost through the document picker. Converter validates
    # workbook incident ID vs ZIP photo entries, not only matching names.
    if not zipfile_ok(fields) or not zipfile_ok(photos):
        print('選択ファイルがExcel/ZIP形式として開けません。')
        temp.cleanup();return None
    return fields,photos,'手動選択',temp

def zipfile_ok(file):
    import zipfile
    return zipfile.is_zipfile(file)

def choose_case(field):
    cases=get_incidents(field)
    print(f'\nExcel内の案件：{len(cases)}件（名前で選択）')
    for n,c in enumerate(cases,1):
        print(f'{n}. {c.get("発生日時","")} / {c.get("件名", "(件名なし)")} / {c.get("発生場所", "")}',flush=True)
    if len(cases)==1:
        print('案件が1件なので自動選択します。')
        return str(cases[0]['ID'])
    try:answer=input(f'案件番号 [1〜{len(cases)}]（0=中止）: ').strip()
    except EOFError:return None
    if not answer.isdigit() or not 1<=int(answer)<=len(cases):return None
    return str(cases[int(answer)-1]['ID'])


def choose_photos(field,caseid):
    _,photos,_=get_related_rows(field,caseid)
    print(f'\nこの案件の写真：{len(photos)}枚')
    for i,p in enumerate(photos,1):
        caption=str(p.get('写真右側の説明文') or '')[:65]
        print(f'{i}. 写真番号{p.get("写真番号")} / {caption}')
    if len(photos)>MAX_PHOTOS:
        print(f'会社Excelへ入れる写真は最大{MAX_PHOTOS}枚。番号を選んでください。')
        try:choice=input('採用する写真番号（例: 1,3,4 / 0=中止）: ').strip()
        except EOFError:return None
        if choice=='0':return None
        try:ids={int(n.strip()) for n in choice.split(',')}
        except ValueError:return None
        if not ids or max(ids)>len(photos) or min(ids)<1 or len(ids)>MAX_PHOTOS:return None
        return {str(p['ID']):{'remove':i not in ids} for i,p in enumerate(photos,1)}
    return {}

def run_conversion(field,photos,stamp):
    REPORT_DIR.mkdir(parents=True,exist_ok=True)
    print('\n選択Excel:',field,'\n写真ZIP:',photos,'\n会社原本:',TEMPLATE,flush=True)
    try:
        caseid=choose_case(field)
        if not caseid:return 2
        case,_,_=get_related_rows(field,caseid)
        actions=choose_photos(field,caseid)
        if actions is None:return 2
        unique=uuid.uuid4().hex[:8]
        output=REPORT_DIR/output_report_name(case.get('件名'),stamp,unique)
        audit=output.with_suffix('.json')
        review=make_report(str(field),str(photos),str(TEMPLATE),str(output),str(audit),case_id=caseid,photo_actions=actions)
        print('\n✅件名付き会社書式Excelを生成:',output)
        print('監査ログ:',audit)
        try:
            raw_excel,raw_zip=save_case_sources(field,photos,caseid,stamp,REPORT_DIR,unique=unique)
            print('✅案件別に分けた元Excel:',raw_excel)
            print('✅案件別に分けた写真ZIP:',raw_zip)
            print('※Field Survey本体の最初の出力ファイルは名前を変更していません。')
            print('※写真の差し替え・説明文の手直しは会社Excel側のみ適用。')
        except Exception as exc:
            print('⚠案件別の元Excel・写真ZIPコピーは作成できませんでした:',type(exc).__name__,exc)
            print('会社書式Excelは生成済みです。元ファイルにも変更はありません。')
        print(f'写真: {review["photo_count"]}枚、対応経過: {review["timeline_count"]}件')
        print('会社提出前の目視・印刷確認は必須です。')
        return 0
    except Exception as exc:
        print('\n❌変換できません:',type(exc).__name__,exc)
        print('入力データ・会社原本は変更していません。')
        return 1

def main():
    print('=== Field Survey → 会社書式 v2.4.2-homebox-preview（件名ファイル名・複数案件・写真選択）/ 変換エンジン '+VERSION+' ===')
    print('※試作・正式提出不可。社内資料は会社の許可がある場合のみ利用。')
    if not TEMPLATE.is_file():
        print('会社原本がありません:',TEMPLATE)
        return 2
    choices=list_pairs()
    temp=None
    try:
        if choices:
            selected=choose_from_pairs(choices)
            if selected is None:print('中止しました。');return 2
            field,photos,stamp=selected
        else:
            print('共有フォルダーの中にExcelと写真ZIPの組が見つかりませんでした。')
            try:option=input('Androidファイル選択画面から2ファイルを直接選びますか？ [y/N]: ').strip().lower()
            except EOFError:option=''
            if option!='y':return 2
            picked=pick_via_android()
            if picked is None:return 2
            field,photos,stamp,temp=picked
        return run_conversion(field,photos,stamp)
    finally:
        if temp:temp.cleanup()

if __name__=='__main__':sys.exit(main())
