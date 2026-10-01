"""Bounded, render-measured fitting. Never changes text or hides overflow."""
import copy, os, sys, tempfile, zipfile
from pathlib import Path
from lxml import etree as E
import fitz
import shutil,subprocess

def convert_to_pdf(doc_path,profile,out,stem,verbose=False):
 executable=os.environ.get('SOFFICE_PATH') or shutil.which('soffice') or shutil.which('libreoffice')
 if not executable:raise RuntimeError('未安装 LibreOffice，无法验证实际分页')
 cmd=[executable,'-env:UserInstallation='+Path(profile).resolve().as_uri(),'--headless','--convert-to','pdf','--outdir',str(out),str(Path(doc_path).resolve())]
 result=subprocess.run(cmd,capture_output=True,text=True,timeout=60)
 pdf=Path(out)/(stem+'.pdf')
 return (str(pdf) if pdf.exists() else '',result.stderr)
N={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'};W='{'+N['w']+'}'
def write_doc(source,root,path):
 with zipfile.ZipFile(source) as zi,zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as zo:
  for info in zi.infolist():zo.writestr(info,E.tostring(root,xml_declaration=True,encoding='UTF-8',standalone=True) if info.filename=='word/document.xml' else zi.read(info.filename))
def tune(base,gain):
 r=copy.deepcopy(base);body=r.find('w:body',N)
 for sect in r.findall('.//w:sectPr',N):
  m=sect.find('w:pgMar',N)
  for a,b in [('left','right'),('top','bottom')]:
   v=str(round((int(m.get(W+a))+int(m.get(W+b)))/2));m.set(W+a,v);m.set(W+b,v)
 for child in list(body)[1:]:
  if child.tag==W+'sectPr':continue
  paragraphs=[child] if child.tag==W+'p' else child.findall('.//w:p',N)
  for p in paragraphs:
   if not p.findall('.//w:t',N):continue
   pp=p.find('w:pPr',N)
   if pp is None:pp=E.Element(W+'pPr');p.insert(0,pp)
   heading=pp.find('w:pBdr',N) is not None
   sp=pp.find('w:spacing',N)
   if sp is None:sp=E.SubElement(pp,W+'spacing')
   # Headings retain their font. Section spacing changes uniformly.
   for key in ['before','after']:
    old=int(sp.get(W+key,'0'))
    if old:sp.set(W+key,str(max(20,round(old*(1+gain*.7)))))
   if not heading:
    baseline=int(sp.get(W+'line','242'))
    if sp.get(W+'lineRule')!='exact':baseline=242
    line=max(216,round(baseline+gain*120));sp.set(W+'line',str(line));sp.set(W+'lineRule','exact')
    for run in p.findall('w:r',N):
     rp=run.find('w:rPr',N)
     if rp is None:rp=E.Element(W+'rPr');run.insert(0,rp)
     for tag in ['sz','szCs']:
      sz=rp.find('w:'+tag,N)
      if sz is None:sz=E.SubElement(rp,W+tag);sz.set(W+'val','21')
      old=int(sz.get(W+'val','21'));size=min(24,max(19,old+round(min(0,gain)*2)))
      sz.set(W+'val',str(size))
 return r

def fit_section(source,base,folder):
 folder.mkdir();best=None;attempt=0
 def measure(gain):
  nonlocal attempt,best
  attempt+=1;r=tune(base,gain);path=folder/f'pass-{attempt}.docx';write_doc(source,r,path)
  profile=folder/f'profile-{attempt}';out=folder/f'pdf-{attempt}';profile.mkdir();out.mkdir()
  pdf,log=convert_to_pdf(str(path),str(profile),str(out),path.stem,False)
  if not pdf:raise RuntimeError('排版检查无法完成，未能渲染 Word')
  doc=fitz.open(pdf);pages=len(doc);page=doc[0];blocks=page.get_text('blocks');boxes=[b[:4] for b in blocks if b[3]>b[1]]
  top=min(x[1] for x in boxes);bottom=max(x[3] for x in boxes);gap=page.rect.height-bottom-top
  result={'root':r,'pages':pages,'remaining_pt':round(gap,1),'gain':round(gain,3),'pdf':pdf}
  if pages==1 and gap>=-4 and (best is None or abs(gap)<abs(best['remaining_pt'])):best=result
  doc.close();return result
 initial=measure(0)
 if initial['pages']==1 and abs(initial['remaining_pt'])<=10:return initial
 low,high=(-1,0) if initial['pages']>1 or initial['remaining_pt']<0 else (0,1)
 edge=measure(low if high==0 else high)
 if high==0 and edge['pages']>1:return edge
 if high==1 and edge['pages']==1 and edge['remaining_pt']>10:return edge
 for _ in range(7):
  mid=(low+high)/2;result=measure(mid)
  if result['pages']>1 or result['remaining_pt']<0:high=mid
  else:low=mid
  if result['pages']==1 and abs(result['remaining_pt'])<=10:break
 return best or result

def auto_fit(path,qa_dir):
 with zipfile.ZipFile(path) as z:root=E.fromstring(z.read('word/document.xml'))
 body=root.find('w:body',N);nodes=list(body);breaks=[i for i,e in enumerate(nodes) if e.find('w:pPr/w:sectPr',N) is not None]
 roots=[]
 if breaks:
  ix=breaks[0];cn=copy.deepcopy(root);cb=cn.find('w:body',N)
  for e in list(cb)[ix:]:cb.remove(e)
  cb.append(copy.deepcopy(nodes[ix].find('w:pPr/w:sectPr',N)));roots.append(cn)
  en=copy.deepcopy(root);eb=en.find('w:body',N)
  for e in list(eb)[:ix+1]:eb.remove(e)
  roots.append(en)
 else:roots=[root]
 qa_dir=Path(qa_dir);qa_dir.mkdir(parents=True,exist_ok=True)
 with tempfile.TemporaryDirectory(prefix='resume-fit-') as tmp:
  results=[fit_section(path,r,Path(tmp)/str(i)) for i,r in enumerate(roots)]
  merged=copy.deepcopy(results[0]['root']);mb=merged.find('w:body',N)
  if len(results)==2:
   sect=mb[-1];mb.remove(sect);p=mb[-1];pp=p.find('w:pPr',N)
   if pp is None:pp=E.Element(W+'pPr');p.insert(0,pp)
   pp.append(sect)
   for e in results[1]['root'].find('w:body',N):mb.append(copy.deepcopy(e))
  for i,res in enumerate(results):
   doc=fitz.open(res['pdf'])
   for j,page in enumerate(doc):page.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(str(qa_dir/f'section-{i+1}-page-{j+1}.png'))
   doc.close()
  write_doc(path,merged,qa_dir/'fitted.docx');(qa_dir/'fitted.docx').replace(path)
  reports=[{k:v for k,v in r.items() if k not in ['root','pdf']} for r in results]
  ok=all(r['pages']==1 and abs(r['remaining_pt'])<=10 for r in reports)
  return {'balanced':ok,'sections':reports,'message':'已按实际渲染调整：每种语言一页，上下留白差不超过约 3.5 毫米。仍需 WPS 核对。' if ok else '已在字号 9.5–12 磅和有限行距内调整；在当前调整范围内未找到兼顾一页与均匀留白的版式，请增减经历或允许改写。'}
