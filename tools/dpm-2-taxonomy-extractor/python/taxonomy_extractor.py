import json, os, io, pandas as pd
FIXED_PATTERN = r"(?i)tab[\\/].*\.json$"
SHEETS=['documentInfo_All','factValue_All','datapoint_dpm1_df_All','datapoint_dpm1_agg_df_All','datapoint_dpm2_df_All','parameter_df_All','parameter_allowedValue_df_All','cell_allowedValue_df_All','Mapping_dpm2_df_All']
def extract(path):
 with open(path,encoding='utf-8') as f:data=json.load(f)
 doc=pd.DataFrame(data.get('documentInfo',{}).get('namespaces',{}).items(),columns=['namespace','url']);doc['documentType']=data.get('documentInfo',{}).get('documentType')
 facts=[];d1=[];d2=[];pars=[];pvals=[];cvals=[]
 for table,cols in data.get('tableTemplates',{}).items():
  for col,items in cols.items():
   fv=items.get('factValue',{}); dims=fv.get('dimensions',{})
   for k,v in dims.items():facts.append({'dimensions_ext':k,'dimensions':v,'dimensions_code':str(v).replace('$',''),'propertiesFrom':(fv.get('propertiesFrom') or [None])[0],'table':table,'file_path':path})
   codes={str(v).replace('$','') for v in dims.values()}
   dp=items.get('datapoint',{}).get('propertyGroups',{})
   for name,obj in dp.items():
    dec=obj.get('decimals'); documentation=obj.get('eba:documentation',{}).copy(); allowed=documentation.pop('AllowedValue',{}) or {}
    for k,v in obj.get('dimensions',{}).items():d1.append({'type_xbrl':dec,'metric_dimension':k,'value':v,'datapoint':name,'table':table,'file_path':path})
    d2.append({'type_xbrl':'$textClosedList' if allowed else dec,**documentation,'datapoint':name,'table':table,'file_path':path})
    for v,c in allowed.items():cvals.append({'value':v,'code':c,'datapoint':name,'table':table,'file_path':path})
   for element,obj in items.items():
    if element in codes and isinstance(obj,dict):
     documentation=obj.get('eba:documentation',{}).copy(); allowed=documentation.pop('AllowedValue',{}) or {}; constraints=obj.get('tc:constraints',{})
     pars.append({**documentation,**{('constraints_'+k if k=='type' else k):v for k,v in constraints.items()},'parameter':element,'table':table,'file_path':path})
     for v,c in allowed.items():pvals.append({'value':v,'code':c,'parameter':element,'table':table,'file_path':path})
  doc['table']=table;doc['file_path']=path
 return doc,pd.DataFrame(facts),pd.DataFrame(d1),pd.DataFrame(d2),pd.DataFrame(pars),pd.DataFrame(pvals),pd.DataFrame(cvals)
def process_taxonomy(paths_json,exclude_json,specific_json,output_path):
 paths=json.loads(paths_json);exclude=[x.lower() for x in json.loads(exclude_json) if x.strip()];specific=[x.lower() for x in json.loads(specific_json) if x.strip()]
 selected=[p for p in paths if '/tab/' in p.lower().replace('\\','/') and p.lower().endswith('.json') and not any(x in p.lower() for x in exclude) and (not specific or any(x in p.lower() for x in specific))]
 if not selected:raise ValueError('No JSON file matched the fixed tab/*.json pattern and filters.')
 allframes=[[] for _ in range(7)];failed=[]
 for p in selected:
  try:
   for i,df in enumerate(extract(p)):allframes[i].append(df)
  except Exception as e:failed.append(f'{os.path.basename(p)}: {e}')
 frames=[pd.concat(x,ignore_index=True) if x else pd.DataFrame() for x in allframes]
 doc,fact,d1,d2,par,pval,cval=frames
 if 'CellCode' in d2.columns:d2['cellcode']=d2.get('cellcode').fillna(d2['CellCode']) if 'cellcode' in d2.columns else d2['CellCode']
 d2=d2.drop(columns=['CellCode'],errors='ignore')
 agg=d1.groupby(['datapoint','table','file_path'],as_index=False,dropna=False).agg({'metric_dimension':'/'.join,'value':'/'.join}) if not d1.empty else pd.DataFrame()
 for c in ['SheetVID','cellcode','type','datapoint','table']:
  if c not in d2.columns:d2[c]=None
 mapping=d2[['SheetVID','cellcode','type','datapoint','table']].copy();parts=mapping['cellcode'].astype('string').str.strip('{}').str.split(',',expand=True)
 for i in range(3):
  if i not in parts.columns:parts[i]=None
 parts=parts.iloc[:,:3];parts.columns=['FR_TABLE_U','FR_ROW','FR_COLUMN'];parts=parts.apply(lambda x:x.str.strip());mapping[['FR_TABLE_U','FR_ROW','FR_COLUMN']]=parts;mapping['FR_TABLE_L']=mapping['FR_TABLE_U'].str[0].str.lower()+mapping['FR_TABLE_U'].str[1:]
 if not par.empty and {'table','headerCode','parameter'}.issubset(par.columns):
  pt=par[['table','headerCode','parameter']].sort_values(['table','headerCode']).groupby('table',as_index=False).agg({'headerCode':lambda x:','.join(x.astype(str)),'parameter':lambda x:','.join(x.astype(str))});mapping=mapping.merge(pt,on='table',how='left')
 else:mapping['headerCode']=None;mapping['parameter']=None
 mapping=mapping.drop(columns=['cellcode','table'],errors='ignore')[['SheetVID','FR_TABLE_U','FR_TABLE_L','FR_ROW','FR_COLUMN','datapoint','parameter','type','headerCode']]
 dfs=dict(zip(SHEETS,[doc,fact,d1,agg,d2,par,pval,cval,mapping]))
 with pd.ExcelWriter(output_path,engine='openpyxl') as writer:
  for name,df in dfs.items():df.to_excel(writer,sheet_name=name,index=False);ws=writer.sheets[name];ws.freeze_panes='A2';ws.auto_filter.ref=ws.dimensions
 return json.dumps({'processed_files':len(selected)-len(failed),'matched_files':len(selected),'mapping_rows':len(mapping),'sheets':[{'sheet':k,'rows':len(v),'columns':v.shape[1]} for k,v in dfs.items()],'preview':mapping.head(200).where(pd.notna(mapping),None).to_dict('records'),'columns':list(mapping.columns),'failed':failed},default=str)
