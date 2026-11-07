#!/usr/bin/env python3
# coding: utf-8
"""Local-only browser UI for Field Survey -> company Excel test converter.
No external network; session token required; no changes to source files.
"""
from __future__ import annotations
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse,parse_qs,quote
from email.parser import BytesParser
from email.policy import default
from pathlib import Path
from html import escape as esc
import io,os,sys,random,secrets,subprocess,shutil,hashlib,zipfile,uuid,json,datetime,traceback
from PIL import Image,ImageOps
from auto_run import list_pairs,TEMPLATE,REPORT_DIR
from converter_mobile import get_incidents,get_related_rows,make_report,MAX_PHOTOS,MAX_IMAGE_SIZE,zip_guard
from case_exports import output_report_name,save_case_sources

TOKEN=secrets.token_urlsafe(22)
MAX_POST=100*1024*1024

CSS='''
:root{font-family:system-ui,'Noto Sans JP',sans-serif;color:#242c39;background:#f5f6fa}
*{box-sizing:border-box}body{margin:0} main{max-width:890px;margin:0 auto;padding:17px 13px 80px}
h1{font-size:1.35rem}h2{font-size:1.12rem;margin-top:22px}p{line-height:1.55}
.card{background:#fff;border:1px solid #d5dce8;border-radius:13px;padding:15px;margin:12px 0}
.warning{border-left:4px solid #cc8119;padding:12px;background:#fff8e6}
small,.muted{color:#5c697c}.pill{display:inline-block;background:#e8eef9;border-radius:12px;padding:4px 9px;margin:4px 4px 4px 0}
label{display:block;font-weight:600;margin:12px 0 5px}input[type=text],textarea,select{width:100%;padding:12px;border:1px solid #a9b7c9;border-radius:8px;font-size:1rem;background:white}
textarea{min-height:98px}button,.btn{background:#335c95;color:white;border:0;padding:13px 16px;border-radius:9px;font-size:1rem;cursor:pointer;display:inline-block;text-decoration:none}
a{color:#234d89}button:disabled{opacity:.6}.photo{display:grid;grid-template-columns:140px 1fr;gap:15px;border-top:1px solid #e0e4eb;padding:18px 0}
.photo img{max-width:100%;max-height:160px;object-fit:contain;border-radius:7px;background:#eef}
.row{display:flex;align-items:center;gap:10px;flex-wrap:wrap}.row label{margin:0} .mono{font-family:monospace;overflow-wrap:anywhere}
@media(max-width:520px){.photo{grid-template-columns:1fr}.photo img{width:180px;max-height:180px}}
'''

def format_page(body,title='写真帳変換'):
    return ('<!doctype html><html lang="ja"><meta name="viewport" content="width=device-width,initial-scale=1">'
        '<meta charset="UTF-8"><title>'+esc(title)+'</title><style>'+CSS+'</style><main>'
        '<h1>現場異常報告書 変換・編集（v2.4.2-homebox-preview 検証版）</h1>'
        '<div class="warning">これはField Survey本体の記録を変更しません。会社提出前に内容・写真・印刷結果を必ず確認してください。</div>'
        +body+'</main></html>')

def find_photo_bytes(zp,name):
    if not isinstance(name,str) or not name:return None
    with zipfile.ZipFile(zp) as z:
        zip_guard(z)
        hits=[n for n in z.namelist() if Path(n).name==name]
        if len(hits)!=1:return None
        b=z.read(hits[0])
        if len(b)>MAX_IMAGE_SIZE:return None
        return b

def duplicates(zp,rows):
    seen={}; captions={}; flags={}
    for r in rows:
        rid=str(r['ID'])
        caption=str(r.get('写真右側の説明文') or '').strip()
        if len(caption)>=5:
            if caption in captions:
                flags[rid]='説明文が別の写真と同一です。写真ごとに内容を確認してください'
                flags[captions[caption]]='説明文が別の写真と同一です。写真ごとに内容を確認してください'
            else: captions[caption]=rid
        try: data=find_photo_bytes(zp,r.get('元写真'))
        except Exception:data=None
        if data is not None:
            sha=hashlib.sha256(data).hexdigest()
            if sha in seen:
                flags[rid]='画像が別の行と完全一致しています'
                flags[seen[sha]]='画像が別の行と完全一致しています'
            else:seen[sha]=rid
    return flags

class Handler(BaseHTTPRequestHandler):
    def log_message(self,*args):pass
    def send_html(self,body,status=200):
        data=format_page(body).encode('utf-8')
        self.send_response(status);self.send_header('Content-Type','text/html; charset=utf-8')
        self.send_header('Content-Length',str(len(data)))
        self.send_header('Cache-Control','no-store');self.send_header('X-Content-Type-Options','nosniff')
        self.end_headers();self.wfile.write(data)
    def query(self):return parse_qs(urlparse(self.path).query)
    def token_valid(self,q):return q.get('token',[None])[0]==TOKEN
    def qpath(self,path,**kw):
        return path+'?'+ '&'.join(f'{quote(k)}={quote(str(v))}' for k,v in [('token',TOKEN),*kw.items()])
    def do_GET(self):
        q=self.query();parsed=urlparse(self.path)
        if not self.token_valid(q):return self.send_html('<p>アクセスキーがありません。</p>',403)
        try:
            if parsed.path=='/image':return self.photo_get(q)
            if parsed.path=='/download':return self.download_get(q)
            return self.main_get(q)
        except Exception as ex:
            self.send_html('<div class="card"><b>読込みエラー</b><p>'+esc(str(ex))+'</p></div><a href="'+self.qpath('/')+'">戻る</a>',400)
    def photo_get(self,q):
        pairs=list_pairs();idx=int(q['pair'][0]);rid=q['rid'][0];caseid=q['case'][0]
        field,zp,_,_=pairs[idx];_,rows,_=get_related_rows(field,caseid)
        row=next((r for r in rows if str(r['ID'])==rid),None)
        if row is None:raise ValueError('写真行がありません')
        raw=find_photo_bytes(zp,row.get('元写真'))
        if raw is None:raise ValueError('写真ZIP内に元画像がありません')
        with Image.open(io.BytesIO(raw)) as img:
            if img.width*img.height>32_000_000:raise ValueError('画像が大きすぎます')
            img=ImageOps.exif_transpose(img).convert('RGB');img.thumbnail((540,400))
            buffer=io.BytesIO();img.save(buffer,'JPEG',quality=77);data=buffer.getvalue()
        self.send_response(200);self.send_header('Content-Type','image/jpeg')
        self.send_header('Content-Length',str(len(data)));self.send_header('Cache-Control','no-store')
        self.end_headers();self.wfile.write(data)
    def download_get(self,q):
        name=q.get('name',[''])[0]
        if not name or Path(name).name!=name or not (name.endswith('.xlsx') or name.endswith('.json') or name.endswith('.zip')):raise ValueError('不正なファイル名')
        file=REPORT_DIR/name
        if not file.is_file():raise FileNotFoundError(name)
        data=file.read_bytes()
        self.send_response(200)
        self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' if name.endswith('.xlsx') else ('application/zip' if name.endswith('.zip') else 'application/json'))
        self.send_header('Content-Disposition',"attachment; filename*=UTF-8''"+quote(name))
        self.send_header('Content-Length',str(len(data)))
        self.end_headers();self.wfile.write(data)
    def main_get(self,q):
        pairs=list_pairs()
        if not pairs:
            return self.send_html('<div class="card">Excel＋写真ZIPの組が見つかりません。Field Surveyでエクスポートしてください。</div>')
        try:idx=int(q.get('pair',[0])[0])
        except ValueError:raise ValueError('不正なファイル選択')
        if idx<0 or idx>=len(pairs):raise ValueError('存在しないファイル組')
        field,zp,stamp,_=pairs[idx]
        cases=get_incidents(field)
        chooser='<form method="get"><input type="hidden" name="token" value="'+esc(TOKEN)+'"><label>① 書き出したExcel＋写真ZIPを選ぶ</label><select name="pair" onchange="this.form.submit()">'
        for i,(f,z,s,_) in enumerate(pairs):chooser+=f'<option value="{i}" '+('selected' if idx==i else '')+'>'+esc(s+' / '+f.name[:55])+'</option>'
        chooser+='</select></form>'
        caseid=q.get('case',[str(cases[0]['ID']) if cases else ''])[0]
        if caseid not in [str(x['ID']) for x in cases]:caseid=str(cases[0]['ID']) if cases else ''
        if not caseid:raise ValueError('このExcelには案件がありません')
        chooser+='<form method="get"><input type="hidden" name="token" value="'+esc(TOKEN)+'"><input type="hidden" name="pair" value="'+str(idx)+'"><label>② 件名で対象の案件を選ぶ（'+str(len(cases))+'件）</label><select name="case" onchange="this.form.submit()">'
        for c in cases:
            id=str(c['ID']); label=esc(str(c.get('発生日時') or '')+' / '+str(c.get('件名') or '(件名なし)')+' / '+str(c.get('発生場所') or ''))
            chooser+='<option value="'+esc(id)+'" '+('selected' if id==caseid else '')+'>'+label+'</option>'
        chooser+='</select></form>'
        case,rows,timeline=get_related_rows(field,caseid)
        flags=duplicates(zp,rows)
        info='<p class="muted">対応経過：'+str(len(timeline))+'件 / 登録写真：'+str(len(rows))+'件 / 会社原本の写真枠：最大3枚</p>'
        if len(rows)>3:info+='<div class="warning">登録写真が3枚を超えています。使用する3枚以内を選んでください。</div>'
        if not rows:info+='<div class="warning">この案件には写真がありません。写真なしの検証用Excelを作成できます。</div>'
        items=''
        for i,r in enumerate(rows):
            rid=str(r['ID']); img=self.qpath('/image',pair=idx,case=caseid,rid=rid)
            flag=('<div class="warning">'+esc(flags[rid])+' — 不要な行を外すか画像を差し替えてください。</div>') if rid in flags else ''
            items+='<div class="photo"><div><img loading="lazy" src="'+img+'" onerror="this.alt=\'画像未取得（差し替え可能）\';this.removeAttribute(\'src\')"><small>写真行 '+str(i+1)+'<br>'+esc(str(r.get('元写真') or '(元写真なし)')[:42])+'</small></div><div>'
            items+='<div class="row"><input type="checkbox" checked name="keep_'+str(i)+'" id="k'+str(i)+'"><label for="k'+str(i)+'">この写真を使う</label></div>'+flag
            items+='<label>右側の説明文（写真ごとに編集可能）</label><textarea name="caption_'+str(i)+'">'+esc(str(r.get('写真右側の説明文') or ''))+'</textarea>'
            items+='<label>撮影場所</label><input type="text" name="location_'+str(i)+'" value="'+esc(str(r.get('撮影場所') or ''),quote=True)+'">'
            items+='<label>画像を差し替える（必要なときだけ）</label><input type="file" accept="image/jpeg,image/png,image/webp" name="replace_'+str(i)+'">'
            items+='<small>元のField Surveyの写真は削除・変更しません。今回の変換結果だけ変更します。</small></div></div>'
        body='<div class="card">'+chooser+'</div><div class="card"><h2>③ 写真ごとに確認・削除・修正</h2>'+info
        body+='<form method="post" action="'+self.qpath('/convert')+'" enctype="multipart/form-data">'
        body+='<input type="hidden" name="pair" value="'+str(idx)+'"><input type="hidden" name="case" value="'+esc(caseid,quote=True)+'">'
        body+=items+'<button type="submit">選んだ案件だけ会社Excelへ変換</button></form></div>'
        self.send_html(body)
    def do_POST(self):
        q=self.query()
        if not self.token_valid(q):return self.send_html('アクセスできません',403)
        if urlparse(self.path).path!='/convert':return self.send_html('不正な処理',404)
        try:
            total=int(self.headers.get('Content-Length','0'))
            if not 0<total<=MAX_POST:raise ValueError('送信データが大きすぎるか空です')
            cont=self.headers.get('Content-Type','')
            if 'multipart/form-data' not in cont:raise ValueError('不正な送信形式')
            mime=b'Content-Type: '+cont.encode('ascii')+b'\r\nMIME-Version: 1.0\r\n\r\n'+self.rfile.read(total)
            msg=BytesParser(policy=default).parsebytes(mime)
            if not msg.is_multipart():raise ValueError('フォームデータを解釈できません')
            fields={};uploads={}
            for part in msg.iter_parts():
                name=part.get_param('name',header='content-disposition')
                if not name:continue
                data=part.get_payload(decode=True) or b''
                filename=part.get_filename()
                if filename and data:uploads[name]=data
                else:
                    if len(data)>100_000:raise ValueError('文章が長すぎます')
                    fields[name]=data.decode('utf-8')
            pairs=list_pairs();idx=int(fields['pair']);f,z,stamp,_=pairs[idx]
            caseid=fields['case'];case,rows,_=get_related_rows(f,caseid)
            edits={}
            for i,row in enumerate(rows):
                rid=str(row['ID']);edit={}
                if 'keep_'+str(i) not in fields:edit['remove']=True
                else:
                    edit['description']=fields.get('caption_'+str(i),'')
                    edit['location']=fields.get('location_'+str(i),'')
                    b=uploads.get('replace_'+str(i))
                    if b:
                        if len(b)>MAX_IMAGE_SIZE:raise ValueError('差し替え写真が30MBを超えています')
                        edit['image_bytes']=b
                edits[rid]=edit
            REPORT_DIR.mkdir(exist_ok=True,parents=True)
            unique=uuid.uuid4().hex[:8]
            output=REPORT_DIR/output_report_name(case.get('件名'),stamp,unique)
            audit=output.with_suffix('.json')
            result=make_report(str(f),str(z),str(TEMPLATE),str(output),str(audit),case_id=caseid,photo_actions=edits)
            filename=output.name
            body='<div class="card"><h2>変換完了（検証用）</h2><p>'+esc(str(case.get('件名') or ''))+'</p>'
            body+='<p>写真：'+str(result['photo_count'])+'枚、対応経過：'+str(result['timeline_count'])+'件</p>'
            body+='<p><a class="btn" href="'+self.qpath('/download',name=filename)+'">① 件名付きの会社書式Excelを開く・保存</a></p>'
            body+='<p><a href="'+self.qpath('/download',name=audit.name)+'">変換検証ログ（JSON）</a></p>'
            try:
                raw_excel,raw_zip=save_case_sources(f,z,caseid,stamp,REPORT_DIR,unique=unique)
                body+='<p><a class="btn" href="'+self.qpath('/download',name=raw_excel.name)+'">② 件名付きField Survey元Excel（案件別コピー）</a></p>'
                body+='<p><a href="'+self.qpath('/download',name=raw_zip.name)+'">③ 件名付き元写真ZIP（案件別コピー）</a></p>'
            except Exception as copy_ex:
                body+='<div class="warning">案件別コピーを作れませんでした：'+esc(str(copy_ex))+'</div>'
            if result.get('unmapped_nonempty_fields'):
                body+='<div class="warning">報告書へ自動転記されない入力欄があります。提出前に確認：'+esc('、'.join(result['unmapped_nonempty_fields']))+'</div>'
            body+='<p class="muted">Field Surveyアプリが最初に付ける名前は変わりません。案件選択後に件名付きのコピーを生成します。</p>'
            body+='<p class="muted">元のField Survey XLSX・元写真ZIP・会社原本は変更せず、修正は会社書式Excelのみに適用されます。</p>'
            body+='<p class="muted">保存先：端末内 Download/FS_変換結果。印刷・目視確認を省略できません。</p></div>'
            body+='<p><a href="'+self.qpath('/',pair=idx,case=caseid)+'">編集画面に戻る</a></p>'
            self.send_html(body)
        except Exception as ex:
            self.send_html('<div class="card"><h2>変換できなかった項目</h2><p>'+esc(str(ex))+'</p><p>入力Excel、元写真ZIP、会社原本はいずれも変更していません。</p></div><a href="'+self.qpath('/')+'">編集画面に戻る</a>',400)

def main():
    if not TEMPLATE.is_file():sys.exit('会社原本 company_template.xlsx がありません。管理者が承認したひな形と approvals/ の承認済み一覧を置いてください（承認手順書_管理者向け.md）。')
    server=ThreadingHTTPServer(('127.0.0.1',0),Handler)
    url=f'http://127.0.0.1:{server.server_address[1]}/?token={TOKEN}'
    print('=== ブラウザで写真帳編集・件名保存 v2.4.2-homebox-preview ===',flush=True)
    print('アドレス:',url,'\n終了するときはCTRL+C。',flush=True)
    opener=shutil.which('termux-open-url')
    if opener:
        try:subprocess.Popen([opener,url],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
        except OSError:pass
    try:server.serve_forever(poll_interval=.4)
    except KeyboardInterrupt:print('終了しました。')
    finally:server.server_close()
if __name__=='__main__':main()
