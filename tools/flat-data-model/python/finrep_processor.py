"""Browser-compatible Excel-to-flat-CSV processing engine."""
import csv, io, json, openpyxl
FIELDS=['Id_Tab','Row','Column','Cod_Conto','Cod_Dest2','Cod_Dest3','Cod_Dest4','Cod_Dest5','Cod_Categoria','Formula','Calculation_Logic','DB_Storage_Sign','EBA_Sign','Coordinate']
def rgb(c): return getattr(c.fill.fgColor,'rgb',None)
def between(t,a,b):
 s=t.find(a);e=t.find(b,s+len(a));return None if s<0 or e<0 else t[s+len(a):e].strip().strip("'").strip()
def last(t,a):
 s=t.find(a);return None if s<0 else t[s+len(a):].strip().strip("'").strip()
def parse(v):
 if v is None:return None
 t=str(v).replace('\n','');A='Account= ';D2='Dest 2 = ';D3='Dest 3 = ';D4='Dest 4 = ';D5='Dest 5 = ';C='Category= ';F='Formula= ';S='DB Storage Sign= ';E='EBA Sign= ';L='Calculation logic= '
 r={k:None for k in FIELDS[3:13]}
 if A in t and F not in t and D2 in t:r.update(Cod_Conto=between(t,A,D2),Cod_Dest2=between(t,D2,D3),Cod_Dest3=between(t,D3,D4),Cod_Dest4=between(t,D4,D5),Cod_Dest5=between(t,D5,C),Cod_Categoria=between(t,C,S),DB_Storage_Sign=between(t,S,E),EBA_Sign=between(t,E,L) if L in t else last(t,E),Calculation_Logic=last(t,L) if L in t else None)
 elif A not in t and F in t and L not in t:r.update(Formula=between(t,F,S),DB_Storage_Sign=between(t,S,E),EBA_Sign=last(t,E))
 else:return None
 return r if r['EBA_Sign'] is not None else None
def process_workbook(path):
 w=openpyxl.load_workbook(path,data_only=False);names=w.sheetnames
 if not names:raise ValueError('The Excel file does not contain valid worksheets.')
 ref=w[names[0]];color=rgb(ref['A1']);rs=ref['A2']._style;cs=ref['B1']._style;cols={};rows={};maxc={};maxr={}
 for n in names:
  sh=w[n];cc=cr=1
  for row in sh.iter_rows():
   for c in row:
    if c.column==1:continue
    if c.value is not None and (c._style==cs or rgb(c)==color):cc+=1;cols[(n,cc)]='c'+str(c.value).zfill(4)
    elif c.value is None:break
   break
  for col in sh.iter_cols():
   for c in col:
    if c.row==1:continue
    if c.value is not None and (c._style==rs or rgb(c)==color):cr+=1;rows[(n,cr)]='r'+str(c.value).zfill(4)
    elif c.value is None:break
   break
  maxc[n]=cc;maxr[n]=cr
 rec=[]
 for n in names:
  if n in {'Test_Formula','Macro'}:continue
  sh=w[n]
  for rr in sh.iter_rows(min_row=2,max_row=maxr[n],min_col=2,max_col=maxc[n]):
   for c in rr:
    if rgb(c)==color:continue
    p=parse(c.value)
    if p:rec.append({'Id_Tab':n,'Row':rows.get((n,c.row),c.row),'Column':cols.get((n,c.column),c.column),**p,'Coordinate':c.coordinate})
 out=io.StringIO(newline='');wr=csv.DictWriter(out,fieldnames=FIELDS,lineterminator='\n');wr.writeheader();wr.writerows(rec)
 return json.dumps({'records':rec,'csv':out.getvalue(),'count':len(rec)},ensure_ascii=False)
