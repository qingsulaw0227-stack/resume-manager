from pathlib import Path
import json, zipfile, copy, uuid, datetime, threading, os, re, base64, io
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from urllib.parse import unquote
from lxml import etree as E
from document import generate_docx, GROUPS
from parser import parse_resume_files
ROOT=Path(__file__).resolve().parent
DATA=Path(os.environ.get('RESUME_DATA_DIR',ROOT/'data')).resolve()
DATA.mkdir(parents=True,exist_ok=True);(DATA/'generated').mkdir(exist_ok=True)
PORT=int(os.environ.get('PORT','8765'))
NS={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
LOCK=threading.Lock();STATE=DATA/'state.json'
def txt(e):return ''.join(e.xpath('.//w:t/text()',namespaces=NS))
def state():
 return json.loads(STATE.read_text()) if STATE.exists() else {'person':{},'entries':[],'profiles':{'综合':{'ids':[],'language':'zh'}},'history':[]}
def save(s):
 p=STATE.with_suffix('.tmp');p.write_text(json.dumps(s,ensure_ascii=False,indent=2));p.replace(STATE)
class Handler(BaseHTTPRequestHandler):
 def response(self,obj,status=200):
  data=json.dumps(obj,ensure_ascii=False).encode();self.send_response(status);self.send_header('Content-Type','application/json; charset=utf-8');self.send_header('Content-Length',str(len(data)));self.end_headers();self.wfile.write(data)
 def do_GET(self):
  if self.path=='/health':return self.response({'ok':True})
  if self.path=='/api/state':return self.response(state())
  if self.path=='/':p=ROOT/'index.html';ctype='text/html; charset=utf-8'
  elif self.path.startswith('/files/'):
   name=unquote(self.path[7:]);p=DATA/'generated'/name
   if '/' in name or '\\' in name or not name.endswith('.docx'):return self.response({'error':'无效文件'},400)
   ctype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
  else:return self.response({'error':'未找到'},404)
  if not p.is_file():return self.response({'error':'文件不存在'},404)
  data=p.read_bytes();self.send_response(200);self.send_header('Content-Type',ctype);self.send_header('Content-Length',str(len(data)));self.end_headers();self.wfile.write(data)
 def do_POST(self):
  if self.headers.get('Origin') not in [None,f'http://127.0.0.1:{PORT}',f'http://localhost:{PORT}']:return self.response({'error':'来源不允许'},403)
  try:
   size=int(self.headers.get('Content-Length','0'))
   limit=50_000_000 if self.path=='/api/bulk-import' else 8_000_000
   if size>limit:raise ValueError('上传内容过大，请减少单次文件数量')
   data=json.loads(self.rfile.read(size))
   with LOCK:
    s=state()
    if self.path=='/api/bulk-import':
     incoming=data.get('files',[])
     if not isinstance(incoming,list) or not 1<=len(incoming)<=30:raise ValueError('每次请选择 1–30 份文件')
     files=[]
     for item in incoming:
      name=str(item.get('name',''));blob=base64.b64decode(item.get('content',''),validate=True)
      if len(blob)>10_000_000:raise ValueError(name+' 超过 10 MB')
      files.append({'name':name,'blob':blob})
     candidates,warnings=parse_resume_files(files)
     def key(value):return re.sub(r'[^\w\u4e00-\u9fff]','',value).lower()
     existing={key(e['title']):e for e in s['entries']};added=0;updated=0
     for candidate in candidates:
      old=existing.get(key(candidate['title']))
      if old:
       old.setdefault('sources',[]);old['sources']=sorted(set(old['sources']+candidate['sources']))
       old.setdefault('conflicts',[]);old['conflicts']=sorted(set(old['conflicts']+candidate['conflicts']))
       if candidate['raw'] and candidate['raw'] not in old.get('raw',''):old['raw']=(old.get('raw','')+'\n\n'+candidate['raw']).strip()
       updated+=1
      else:
       candidate['id']=uuid.uuid4().hex;candidate['revisions']=[];s['entries'].append(candidate);existing[key(candidate['title'])]=candidate;added+=1
     save(s);return self.response({'added':added,'merged':updated,'warnings':warnings,'entries':s['entries']})
    if self.path=='/api/person':
     p=data.get('person',{})
     if not isinstance(p,dict) or not all(isinstance(p.get(k,''),str) and len(p.get(k,''))<1000 for k in ['name','english_name','contact']):raise ValueError('个人信息格式无效')
     if not p.get('name','').strip():raise ValueError('请填写姓名')
     photo=p.get('photo','')
     if not isinstance(photo,str) or len(photo)>4000000:raise ValueError('照片过大')
     if photo:
      from PIL import Image
      image=Image.open(io.BytesIO(base64.b64decode(photo,validate=True)))
      if image.format not in ['PNG','JPEG'] or image.width*image.height>12000000:raise ValueError('请选择 1200 万像素以内的 PNG/JPEG 照片')
      image.verify()
     s['person']={k:p.get(k,'').strip() for k in ['name','english_name','contact','photo']};save(s);return self.response({'ok':True})
    if self.path=='/api/import':
     name=str(data.get('name',''));blob=base64.b64decode(data['content'],validate=True)
     if len(blob)>5000000:raise ValueError('文件超过 5 MB')
     if name.lower().endswith('.docx'):
      with zipfile.ZipFile(io.BytesIO(blob)) as z:
       info=z.getinfo('word/document.xml')
       if info.file_size>10000000:raise ValueError('文档解压后过大')
       doc=E.fromstring(z.read(info),E.XMLParser(resolve_entities=False,no_network=True))
       text='\n'.join(txt(p) for p in doc.findall('.//w:p',NS))
     elif name.lower().endswith(('.txt','.md')):text=blob.decode('utf-8-sig')
     else:raise ValueError('支持 DOCX、UTF-8 TXT 和 Markdown 文件')
     if len(text)>100000:raise ValueError('原文超过 10 万字，请拆分导入，未保存截断内容')
     return self.response({'text':text})
    if self.path=='/api/entry':
     e=data['entry'];keys=['title','date','description','english_title','english_date','english','raw']
     if e.get('group') not in GROUPS:raise ValueError('请选择经历类别')
     if not all(isinstance(e.get(k,''),str) and len(e.get(k,''))<=100000 for k in keys):raise ValueError('文本无效或过长')
     if not e.get('title','').strip():raise ValueError('请填写经历名称和职位')
     old=next((x for x in s['entries'] if x['id']==e.get('id')),None)
     if e.get('id') and old is None:raise ValueError('经历不存在')
     clean={k:e.get(k,'').strip() for k in keys};clean.update(id=old['id'] if old else uuid.uuid4().hex,group=e['group'])
     if old:
      for k in ['sources','conflicts']:
       if k in old:clean[k]=copy.deepcopy(old[k])
     changed=old and any(clean[k]!=old.get(k,'') for k in ['title','date','description'])
     clean['english_pending']=bool((changed or (old and old.get('english_pending'))) and not e.get('confirm_english'))
     clean['revisions']=copy.deepcopy(old.get('revisions',[])) if old else []
     if old:clean['revisions'].append({k:v for k,v in old.items() if k!='revisions'})
     if old:s['entries'][s['entries'].index(old)]=clean
     else:s['entries'].append(clean)
     save(s);return self.response(clean)
    if self.path=='/api/profile':
     name=data.get('name','').strip()
     if not name or len(name)>40 or re.search(r'[\\/\x00-\x1f]',name):raise ValueError('方向名称请使用 1–40 个普通字符')
     if name in s['profiles']:raise ValueError('该方向已经存在')
     s['profiles'][name]={'ids':[],'language':'zh'};save(s);return self.response({'name':name})
    profile=data.get('profile');ids=data.get('ids');lang=data.get('language','zh')
    if profile not in s['profiles']:raise ValueError('无效分类')
    if not isinstance(ids,list) or not all(isinstance(x,str) and x in [c['id'] for c in s['entries']] for x in ids):raise ValueError('无效经历')
    if lang not in ['zh','en','both']:raise ValueError('无效语言')
    if self.path=='/api/generate':
     stamp=datetime.datetime.now().strftime('%Y%m%d-%H%M%S');name=f'resume-{stamp}-{uuid.uuid4().hex[:8]}.docx'
     generate_docx(s['person'],s['entries'],ids,lang,DATA/'generated'/name)
     layout_report=None
     if data.get('auto_layout',True):
      from layout import auto_fit
      try:layout_report=auto_fit(DATA/'generated'/name,DATA/'layout-checks'/name[:-5])
      except Exception as exc:layout_report={'balanced':False,'message':'自动排版检查未完成，已保留基础版 Word。请在 WPS 检查。'}
     item={'name':name,'profile':profile,'ids':ids,'language':lang,'time':datetime.datetime.now().isoformat(timespec='seconds'),'layout':layout_report,'entries_snapshot':[e for e in s['entries'] if e['id'] in ids],'person_snapshot':s['person']}
     s['profiles'][profile]={'ids':ids,'language':lang};s['history'].insert(0,item);save(s);return self.response({k:v for k,v in item.items() if k!='entries_snapshot'})
    if self.path=='/api/save':s['profiles'][profile]={'ids':ids,'language':lang};save(s);return self.response({'ok':True})
    return self.response({'error':'未找到'},404)
  except (ValueError,TypeError,KeyError,zipfile.BadZipFile,E.XMLSyntaxError,OSError) as e:return self.response({'error':str(e)},400)
if __name__=='__main__':
 print(f'简历入口 http://127.0.0.1:{PORT}',flush=True)
 ThreadingHTTPServer((os.environ.get('BIND_HOST','127.0.0.1'),PORT),Handler).serve_forever()
