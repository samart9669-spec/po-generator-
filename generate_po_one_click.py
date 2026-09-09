#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""One-click Seamless PO generator.

Uses only Python standard library. Put the monthly plan and PO template in INPUT/;
run the launcher to create OUTPUT/PO_Generated.xlsx.
"""
from __future__ import annotations
import datetime, re, sys, zipfile, html
import xml.etree.ElementTree as ET
from pathlib import Path

NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main'; REL='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
R='{%s}'%REL; ET.register_namespace('',NS); ET.register_namespace('r',REL)
Q=lambda t:'{%s}%s'%(NS,t)
MODEL={
 'AiR':('09_📦 AiR แผนผลิต',31,167,{'M9':'Air seamless M9','M10':'Air seamless M10','M11':'Air seamless M11'}),
 'Freedom':('12_📦 Freedom แผนผลิต',31,124,{'M9':'FreeDom M9','M10':'FreeDom M10','M11':'FreeDom M11'}),
 'Kids':('15_📦 Kids แผนผลิต',24,44,{'M11':'Kid seamless M9'})}
MONTH={'M9':'M','M10':'N','M11':'O'}; PACK={'Pack1':1,'Pack3':3,'Pack5':5,'Doz':12}
COLOR={'BL':'Black','NB':'Navy Blue','GR':'Grey','WH':'White','BI':'Blue indigo','DG':'Dk.Grey','MX':'Mix'}

def norm(s): return re.sub(r'\s+',' ',(s or '').strip()).lower()
def colnum(c):
 n=0
 for x in c:n=n*26+ord(x)-64
 return n
def ref(s):
 m=re.match(r'([A-Z]+)(\d+)',s); return m.group(1),int(m.group(2))
def shared(z):
 if 'xl/sharedStrings.xml' not in z.namelist():return []
 root=ET.fromstring(z.read('xl/sharedStrings.xml')); out=[]
 for si in root.findall(Q('si')):out.append(''.join(t.text or '' for t in si.iter(Q('t'))))
 return out
def workbook(z):
 wb=ET.fromstring(z.read('xl/workbook.xml')); rr=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
 mp={x.attrib['Id']:x.attrib['Target'] for x in rr}; out=[]
 for s in wb.find(Q('sheets')):
  p=mp[s.attrib[R+'id']]; p=('xl/'+p if not p.startswith('/') else p.lstrip('/'))
  out.append({'name':s.attrib['name'],'path':p})
 return wb,rr,out
def values(z,path,ss):
 root=ET.fromstring(z.read(path)); d={}
 for row in root.findall('.//'+Q('row')):
  rn=int(row.attrib['r'])
  for c in row.findall(Q('c')):
   typ=c.attrib.get('t'); v=c.find(Q('v')); val=None
   if typ=='s' and v is not None:
    try:val=ss[int(v.text)]
    except:pass
   elif typ=='inlineStr':val=''.join(t.text or '' for t in c.iter(Q('t')))
   elif v is not None:
    try: val=float(v.text) if any(x in v.text for x in '.Ee') else int(v.text)
    except: val=v.text
   d[(rn,ref(c.attrib['r'])[0])]=val
 return d
def pack(sku):
 if sku.startswith('M5R-'):return 'Pack5'
 if sku.startswith('B3R-'):return 'Pack3'
 if sku.startswith('FTR-'):return 'Doz'
 return 'Pack1'
def base(sku):
 p=sku.split('-'); return '-'.join(p[:4]) if len(p)>=4 else sku
def sku_detail(sku):
 p=sku.split('-')
 return (p[-2],COLOR.get(p[-1],p[-1])) if len(p)>=2 else ('','')
def choose(inp):
 xs=[p for p in inp.glob('*.xlsx') if not p.name.startswith('~$')]
 plans=sorted([p for p in xs if ('แผน' in p.name or 'plan' in p.name.lower()) and 'PO' not in p.name],key=lambda p:p.stat().st_mtime,reverse=True)
 tmpls=sorted([p for p in xs if 'PO' in p.name.lower()],key=lambda p:p.stat().st_mtime,reverse=True)
 if not plans:raise RuntimeError('ไม่พบไฟล์แผนสั่งผลิตใน INPUT/')
 if not tmpls:raise RuntimeError('ไม่พบไฟล์ PO Template ใน INPUT/')
 return plans[0],tmpls[0]
def template_master(z,shs,ss):
 m={}; order=[]
 for sh in shs:
  if not any(sh['name'].startswith(x) for x in ('Air seamless','FreeDom','Kid seamless')):continue
  v=values(z,sh['path'],ss)
  for r in range(16,121):
   sty=v.get((r,'C'))
   if not isinstance(sty,str) or sty.count('-')<3:continue
   bk=base(sty); order.append(bk) if bk not in order else None
   m.setdefault(bk,{'description':v.get((r,'D')) or '','pack':v.get((r,'E')) or pack(sty),'price':v.get((r,'J')),'currency':'USD' if sh['name'].strip().startswith('Kid') else 'CNY'})
 return m,{x:i for i,x in enumerate(order)}
def make_lines(z,ss,cfg,month,m):
 name,start,end,_=cfg; _,_,shs=workbook(z); sh=next((x for x in shs if norm(x['name'])==norm(name)),None)
 if not sh:raise RuntimeError('ไม่พบชีต '+name)
 v=values(z,sh['path'],ss); groups={}; mc=MONTH[month]
 for r in range(start+1,end+1):
  sku=v.get((r,'E')); qty=v.get((r,mc))
  if not isinstance(sku,str) or '-' not in sku or not isinstance(qty,(int,float)) or qty<=0:continue
  p=pack(sku); key=sku.rsplit('-',1)[0] if p!='Pack1' else sku
  g=groups.setdefault((key,p),0); groups[(key,p)]=g+float(qty)
 out=[]
 for (key,p),qty in groups.items():
  bk=base(key)
  if bk not in m:raise RuntimeError(f'Missing PO mapping: {key} ({month})')
  size,color=sku_detail(key) if p=='Pack1' else (key.split('-')[-1],'Mix')
  t=m[bk]; q=int(round(qty)); ps=PACK[p]
  out.append({'sku':key,'qty':q,'size':size,'color':color,'pack':p,'description':t['description'],'price':t['price'],'currency':t['currency']})
 return out

def cellfrag(row,refx):
 m=re.search(r'<c\b[^>]*r="'+re.escape(refx)+r'"[^>]*(?:/>|>.*?</c>)',row,re.S); return m.group(0) if m else None
def attrs(f):
 if not f:return ''
 m=re.match(r'<c\b([^>]*)>',f,re.S); a=m.group(1) if m else ''
 a=re.sub(r'\s+t="[^"]*"','',a); a=re.sub(r'\s+r="[^"]*"','',a); return a.rstrip('/').rstrip()
def makecell(r,a,val=None,string=False):
 a=' r="'+r+'"'+a
 if val is None:return '<c'+a+'/>'
 if string:return '<c'+a+' t="inlineStr"><is><t>'+html.escape(str(val),quote=False)+'</t></is></c>'
 return '<c'+a+'><v>'+str(val)+'</v></c>'
def repl(row,r,val=None,string=False):
 f=cellfrag(row,r); n=makecell(r,attrs(f),val,string)
 if f:return row.replace(f,n,1)
 return row[:-2]+n+'</row>' if row.endswith('/>') else row.replace('</row>',n+'</row>',1)
def rows(sheet):
 sm=re.search(r'<sheetData\b[^>]*>(.*?)</sheetData>',sheet,re.S); body=sm.group(1); basepos=sm.start(1)
 out=[]
 for m in re.finditer(r'<row\b[^>]*\br="(\d+)"[^>]*(?:/>|>.*?</row>)',body,re.S):out.append((int(m.group(1)),m.group(0),basepos+m.start(),basepos+m.end()))
 return out,sm
def shift(rx,new):
 old=int(re.search(r'<row\b[^>]*\br="(\d+)"',rx).group(1)); rx=re.sub(r'(\br=")'+str(old)+r'"',r'\g<1>'+str(new)+'"',rx,1); rx=re.sub(r'(\br=")([A-Z]+)'+str(old)+r'"',lambda m:m.group(1)+m.group(2)+str(new)+'"',rx); return rx
def clear(rx,r):
 for c in 'BCDEFGHIJK':rx=repl(rx,f'{c}{r}')
 return rx
def locate(sheet):
 rs,_=rows(sheet); its=[]
 for r,rx,_,_ in rs:
  b=cellfrag(rx,f'B{r}'); c=cellfrag(rx,f'C{r}')
  if 16<=r<=120 and b and c and re.search(r'<v>\d+(?:\.0+)?</v>',b) and ' t="s"' in c:its.append(r)
 if not its:raise RuntimeError('ไม่พบพื้นที่รายการสินค้าใน PO Template')
 total=next((r for r,_,_,_ in rs if r>max(its) and r<=120 and (cellfrag(_ if False else next(rx for rr,rx,_,_ in rs if rr==r),f'K{r}') or cellfrag(next(rx for rr,rx,_,_ in rs if rr==r),f'H{r}'))),max(its)+1)
 return min(its),max(its),total

def patch(sheet,lines,model,month):
 rs,sm=rows(sheet); first,last,total=locate(sheet); mp={r:rx for r,rx,_,_ in rs}; style=mp[first]; totalrx=mp[total]; desired=first+len(lines); delta=desired-total
 body=sm.group(1); firstpos=next(st for r,rx,st,en in rs if r==first)-sm.start(1); prefix=body[:firstpos]
 gen=[]
 for i,x in enumerate(lines):
  r=first+i; rx=shift(style,r); rx=clear(rx,r)
  vals={'B':i+1,'C':x['sku'],'D':x['description'],'E':x['pack'],'F':x['size'],'G':x['color'],'H':x['qty'],'I':(x['qty']/PACK[x['pack']] if PACK[x['pack']]>1 else None),'J':x['price'],'K':x['qty']/PACK[x['pack']]*x['price'] if x['price'] not in (None,'') else None}
  for c,v in vals.items():rx=repl(rx,f'{c}{r}',v,not isinstance(v,(int,float)) and v is not None)
  gen.append(rx)
 tr=shift(totalrx,desired); tr=clear(tr,desired); tq=sum(x['qty'] for x in lines); ta=sum(x['qty']/PACK[x['pack']]*x['price'] for x in lines if isinstance(x['price'],(int,float)))
 for c,v in {'H':'Total','I':tq,'J':ta,'K':ta}.items():tr=repl(tr,f'{c}{desired}',v,not isinstance(v,(int,float)))
 post=''.join(shift(rx,r+delta) for r,rx,_,_ in rs if r>total)
 newbody=prefix+''.join(gen)+tr+post
 out=sheet[:sm.start(1)]+newbody+sheet[sm.end(1):]
 out=re.sub(r'(<mergeCell\b[^>]*\bref=")([A-Z]+)(\d+)(:[A-Z]+)(\d+)(")',lambda m:m.group(1)+m.group(2)+str(int(m.group(3))+(delta if int(m.group(3))>=total else 0))+m.group(4)+str(int(m.group(5))+(delta if int(m.group(5))>=total else 0))+m.group(6),out)
 return out.encode(),tq,ta

def add_sheet(parts,tmpl_shs,source,newname,wtx,rtx,ctx):
 src=next(s for s in tmpl_shs if norm(s['name'])==norm(source)); nums=[int(m.group(1)) for k in parts if (m:=re.search(r'sheet(\d+)\.xml$',k))]; n=max(nums)+1; path=f'xl/worksheets/sheet{n}.xml'; parts[path]=parts[src['path']]
 ridnums=[int(x) for x in re.findall(r'Id="rId(\d+)"',rtx.decode())]; rid='rId'+str(max(ridnums+[0])+1)
 sidnums=[int(x) for x in re.findall(r'<sheet\b[^>]*sheetId="(\d+)"',wtx.decode())]; sid=str(max(sidnums+[0])+1)
 w=wtx.decode().replace('</sheets>',f'<sheet name="{newname}" sheetId="{sid}" r:id="{rid}" /></sheets>',1)
 rr=rtx.decode().replace('</Relationships>',f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml" /></Relationships>',1)
 cc=ctx.decode().replace('</Types>',f'<Override PartName="/{path}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" /></Types>',1)
 return path,w.encode(),rr.encode(),cc.encode()

def main():
 basep=Path(__file__).resolve().parent; inp=basep/'INPUT'; out=basep/'OUTPUT'; out.mkdir(exist_ok=True)
 plan,tmpl=choose(inp)
 with zipfile.ZipFile(plan) as zp,zipfile.ZipFile(tmpl) as zt:
  ssp=shared(zp); sst=shared(zt); wbr,rr,shs=workbook(zt); master,order=template_master(zt,shs,sst)
  parts={n:zt.read(n) for n in zt.namelist()}; wtx=parts['xl/workbook.xml']; rtx=parts['xl/_rels/workbook.xml.rels']; ctx=parts['[Content_Types].xml']; generated=[]
  for model,(sn,st,en,targets) in MODEL.items():
   for month,source in targets.items():
    lines=make_lines(zp,ssp,(sn,st,en,targets),month,master)
    if not lines:continue
    lines.sort(key=lambda x:(order.get(base(x['sku']),10**9),x['sku']))
    path,wtx,rtx,ctx=add_sheet(parts,shs,source,f'GEN_{model}_{month}',wtx,rtx,ctx); parts[path],qty,amt=patch(parts[path],lines,model,month); generated.append((model,month,len(lines),qty,amt))
  parts['xl/workbook.xml']=wtx; parts['xl/_rels/workbook.xml.rels']=rtx; parts['[Content_Types].xml']=ctx
  wb=ET.fromstring(wtx); calc=wb.find(Q('calcPr')) or ET.SubElement(wb,Q('calcPr')); calc.set('fullCalcOnLoad','1'); calc.set('forceFullCalc','1'); calc.set('calcMode','auto'); parts['xl/workbook.xml']=ET.tostring(wb,encoding='utf-8',xml_declaration=True)
  target=out/'PO_Generated.xlsx'
  with zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as zo:
   for n,b in parts.items():zo.writestr(n,b)
 with (out/'PO_Generated_Summary.txt').open('w',encoding='utf8') as f:
  f.write(f'Generated {datetime.datetime.now():%Y-%m-%d %H:%M}\nPlan: {plan.name}\nTemplate: {tmpl.name}\n\n'); [f.write(f'{a} {b}: {c} lines | {d:,.0f} pcs | {e:,.2f}\n') for a,b,c,d,e in generated]
 print('SUCCESS'); print('Output:',out/'PO_Generated.xlsx'); [print(x) for x in generated]
if __name__=='__main__':
 try:main()
 except Exception as e: print('ERROR:',e);sys.exit(1)
