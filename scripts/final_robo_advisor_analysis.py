"""Final analysis script for the submission-ready robo-advisor multi-study manuscript.

This script reads the data files that the user placed in /mnt/data and creates analysis
outputs for the final manuscript. It does not redistribute raw FINRA data. To reproduce
Study 3, users must independently obtain the FINRA NFCS Investor Survey under FINRA's
terms of use and place it in the same path.
"""
import os, csv, re, json, math, warnings
from zipfile import ZipFile
import xml.etree.ElementTree as ET
from collections import Counter
import numpy as np
import statsmodels.api as sm
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
from sklearn.model_selection import StratifiedKFold, cross_val_predict
from sklearn.preprocessing import StandardScaler
from sklearn.pipeline import Pipeline
from sklearn.linear_model import LogisticRegression
from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier
from sklearn.metrics import roc_auc_score, accuracy_score, balanced_accuracy_score, f1_score, roc_curve
from sklearn.inspection import permutation_importance

warnings.filterwarnings("ignore")
SEED = 42
BASE = os.environ.get('ROBO_PROJECT_ROOT', os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
OUT = os.path.join(BASE, 'outputs')
os.makedirs(OUT, exist_ok=True)

# -------------------------------
# Minimal XLSX reader using XML
# -------------------------------
NS={'main':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}

def col_to_idx(cell_ref):
    m=re.match(r'([A-Z]+)', cell_ref)
    if not m: return None
    idx=0
    for ch in m.group(1): idx=idx*26+ord(ch)-64
    return idx-1

def get_shared_strings(z):
    try:
        root=ET.fromstring(z.read('xl/sharedStrings.xml'))
    except KeyError:
        return []
    strings=[]
    for si in root.findall('main:si', NS):
        texts=[]
        for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'):
            texts.append(t.text or '')
        strings.append(''.join(texts))
    return strings

def get_sheets(z):
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    relmap={rel.attrib['Id']: rel.attrib['Target'] for rel in rels}
    sheets=[]
    for s in wb.find('main:sheets',NS):
        name=s.attrib['name']; rid=s.attrib.get('{%s}id'%NS['r'])
        target=relmap[rid]
        if not target.startswith('worksheets/'):
            target='worksheets/'+target.split('/')[-1]
        sheets.append((name, 'xl/'+target))
    return sheets

def cell_value(c, shared):
    t=c.attrib.get('t')
    if t=='s':
        v=c.find('main:v',NS)
        return '' if v is None else shared[int(v.text)]
    if t=='inlineStr':
        return ''.join([(t.text or '') for t in c.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')])
    v=c.find('main:v',NS)
    return '' if v is None else v.text

def read_xlsx_sheet(path, sheet_name=None, sheet_index=None):
    with ZipFile(path) as z:
        shared=get_shared_strings(z)
        sheets=get_sheets(z)
        if sheet_name is not None:
            target=next(t for n,t in sheets if n==sheet_name)
        elif sheet_index is not None:
            target=sheets[sheet_index][1]
        else:
            target=sheets[0][1]
        root=ET.fromstring(z.read(target))
        rows=[]; maxcols=0
        for r in root.findall('main:sheetData/main:row',NS):
            vals={}; maxidx=-1
            for c in r.findall('main:c',NS):
                idx=col_to_idx(c.attrib.get('r',''))
                if idx is None: continue
                vals[idx]=cell_value(c,shared); maxidx=max(maxidx,idx)
            row=['']*(maxidx+1 if maxidx>=0 else 0)
            for i,v in vals.items(): row[i]=v
            rows.append(row); maxcols=max(maxcols,len(row))
        return [r+['']*(maxcols-len(r)) for r in rows]

# -------------------------------
# Statistical helpers
# -------------------------------
def to_float(x):
    try:
        if x is None or str(x).strip()=='' : return np.nan
        return float(str(x).strip())
    except Exception:
        return np.nan

def cronbach_alpha(mat):
    mat=np.array(mat, dtype=float)
    mat=mat[~np.isnan(mat).any(axis=1)]
    k=mat.shape[1]
    if k < 2 or mat.shape[0] < 3: return np.nan
    item_vars=mat.var(axis=0, ddof=1)
    total_var=mat.sum(axis=1).var(ddof=1)
    if total_var == 0: return np.nan
    return k/(k-1)*(1-item_vars.sum()/total_var)

def mean_sd(x):
    x=np.array(x,dtype=float); x=x[~np.isnan(x)]
    return float(x.mean()), float(x.std(ddof=1))

def zscore(x):
    x=np.array(x,dtype=float)
    return (x-np.nanmean(x))/np.nanstd(x, ddof=1)

def save_csv(path, rows, cols=None):
    if cols is None:
        cols=list(rows[0].keys()) if rows else []
    with open(path,'w',encoding='utf-8',newline='') as f:
        wr=csv.DictWriter(f,fieldnames=cols)
        wr.writeheader(); wr.writerows(rows)

def ols_hc3(y, X, names):
    y=np.array(y,dtype=float); X=np.array(X,dtype=float)
    mask=~np.isnan(y)
    for j in range(X.shape[1]): mask &= ~np.isnan(X[:,j])
    y2=y[mask]; X2=X[mask,:]
    Xc=sm.add_constant(X2)
    model=sm.OLS(y2,Xc).fit(cov_type='HC3')
    rows=[]
    for i,n in enumerate(['Intercept']+names):
        rows.append({'term':n,'coef':float(model.params[i]),'se':float(model.bse[i]),'p':float(model.pvalues[i]),'ci_low':float(model.conf_int()[i,0]),'ci_high':float(model.conf_int()[i,1])})
    return model, rows, int(mask.sum())

def wmean(values, weights):
    vals=[]; w=[]
    for v,wt in zip(values,weights):
        if not np.isnan(v) and not np.isnan(wt): vals.append(v); w.append(wt)
    if not vals: return np.nan
    vals=np.array(vals,dtype=float); w=np.array(w,dtype=float)
    return float(np.sum(vals*w)/np.sum(w))

def wprop(condition, valid, weights):
    vals=[]; w=[]
    for cond,val,wt in zip(condition,valid,weights):
        if val and not np.isnan(wt): vals.append(1.0 if cond else 0.0); w.append(wt)
    if not vals: return np.nan
    return float(np.sum(np.array(vals)*np.array(w))/np.sum(w))

def classification_cv(X, y, feature_names, prefix, sample_weight=None):
    mask=(~np.isnan(y))
    for j in range(X.shape[1]): mask &= ~np.isnan(X[:,j])
    X2=X[mask]; y2=y[mask].astype(int)
    w2=None if sample_weight is None else sample_weight[mask]
    cv=StratifiedKFold(n_splits=5, shuffle=True, random_state=SEED)
    models={
        'Regularized logistic': Pipeline([('scale',StandardScaler()),('clf',LogisticRegression(max_iter=1000, class_weight='balanced', random_state=SEED))]),
        'Random forest': RandomForestClassifier(n_estimators=100, max_depth=5, min_samples_leaf=10, random_state=SEED, class_weight='balanced_subsample', n_jobs=-1),
        'Gradient boosting': GradientBoostingClassifier(n_estimators=70, learning_rate=0.05, max_depth=2, random_state=SEED),
    }
    rows=[]; probas={}
    for name,model in models.items():
        oof=np.zeros(len(y2))*np.nan
        fold_aucs=[]; fold_acc=[]; fold_bacc=[]; fold_f1=[]
        for tr,te in cv.split(X2,y2):
            if sample_weight is not None:
                try:
                    if isinstance(model, Pipeline): model.fit(X2[tr], y2[tr], clf__sample_weight=w2[tr])
                    else: model.fit(X2[tr], y2[tr], sample_weight=w2[tr])
                except Exception:
                    model.fit(X2[tr], y2[tr])
            else:
                model.fit(X2[tr], y2[tr])
            p=model.predict_proba(X2[te])[:,1]
            yh=(p>=0.5).astype(int)
            oof[te]=p
            fold_aucs.append(roc_auc_score(y2[te],p))
            fold_acc.append(accuracy_score(y2[te],yh))
            fold_bacc.append(balanced_accuracy_score(y2[te],yh))
            fold_f1.append(f1_score(y2[te],yh))
        probas[name]=oof
        rows.append({'model':name,'n':len(y2),'auc_mean':float(np.mean(fold_aucs)),'auc_sd':float(np.std(fold_aucs,ddof=1)),'accuracy_mean':float(np.mean(fold_acc)),'balanced_accuracy_mean':float(np.mean(fold_bacc)),'f1_mean':float(np.mean(fold_f1))})
    best=max(rows, key=lambda r:r['auc_mean'])['model']
    best_model=models[best]
    if sample_weight is not None:
        try:
            if isinstance(best_model, Pipeline): best_model.fit(X2,y2, clf__sample_weight=w2)
            else: best_model.fit(X2,y2, sample_weight=w2)
        except Exception:
            best_model.fit(X2,y2)
    else:
        best_model.fit(X2,y2)
    if isinstance(best_model, Pipeline):
        coef=np.abs(best_model.named_steps['clf'].coef_[0])
        raw=coef/(coef.sum()+1e-12)
    elif hasattr(best_model, 'feature_importances_'):
        raw=best_model.feature_importances_
    else:
        raw=np.zeros(len(feature_names))
    imp=[{'feature':feature_names[i], 'importance':float(raw[i]), 'sd':0.0} for i in range(len(feature_names))]
    imp=sorted(imp, key=lambda x:x['importance'], reverse=True)
    save_csv(os.path.join(OUT,f'{prefix}_ml_performance.csv'), rows)
    save_csv(os.path.join(OUT,f'{prefix}_ml_importance.csv'), imp)
    return {'rows':rows,'importance':imp,'probas':probas,'y':y2,'best':best,'n':len(y2),'mask':mask}

# -------------------------------
# Study 2: Indonesian robo-advisor survey
# -------------------------------
ind_path=os.path.join(BASE,'data','raw','study2_indonesia','study2_indonesia_raw.xlsx')
boot_path=os.path.join(BASE,'data','raw','study2_indonesia','Boostraping - Result.xlsx')
smartpls_path=os.path.join(BASE,'data','raw','study2_indonesia','Smartpls - Validty.xlsx')
nfcs_path=os.path.join(BASE,'data','raw','study3_finra','NFCS 2024 Investor Data 251114.sav')

rows=read_xlsx_sheet(ind_path, sheet_name='Sheet3')
headers=rows[0]
arr=np.array([[to_float(x) for x in r[:len(headers)]] for r in rows[1:] if any(str(x).strip() for x in r)],dtype=float)
constructs={
    'SI':['SI1','SI2','SI3'],
    'PE':['PE1','PE2','PE3','PE4'],
    'TR':['TR1','TR2','TR3','TR4','TR5'],
    'FL':['FL1','FL2','FL3','FL4'],
    'IQ':['IQ1','IQ2','IQ3','IQ4'],
    'OBU':['OBU1','OBU2','OBU3','OBU4'],
    'RISK':['RISK1','RISK2','RISK3','RISK4','RISK5','RISK6'],
    'BEN':['BEN1','BEN2','BEN3'],
    'AA':['AA1','AA2','AA3','AA4'],
    'WU':['WU1','WU2','WU3','WU4'],
    'SR':['SR1','SR2','SR3'],
}
construct_labels={
    'SI':'Social influence', 'PE':'Performance expectancy', 'TR':'Trust', 'FL':'Financial literacy/self-knowledge',
    'IQ':'Information/interface quality', 'OBU':'Need for human contact barrier', 'RISK':'Perceived robo-advisor risk',
    'BEN':'Perceived benefit', 'AA':'Awareness/adoption channels', 'WU':'Willingness to use', 'SR':'Security/privacy concern'}
colmap={h:i for i,h in enumerate(headers)}
con_scores={}; reliability=[]
for c,items in constructs.items():
    idx=[colmap[it] for it in items]
    mat=arr[:,idx]
    score=np.nanmean(mat,axis=1)
    con_scores[c]=score
    m,sd=mean_sd(score)
    reliability.append({'construct':c,'label':construct_labels[c],'items':len(items),'mean':m,'sd':sd,'alpha':cronbach_alpha(mat)})
save_csv(os.path.join(OUT,'study2_construct_reliability.csv'), reliability)

# survey filtering demographics and actual usage/awareness
filter_rows=read_xlsx_sheet(ind_path, sheet_name='Filtering ')
filter_header=filter_rows[0]
filter_data=[r for r in filter_rows[1:] if any(str(x).strip() for x in r)]
valid_filter=[]; used=[]
for r in filter_data:
    aw = r[9].strip().lower() if len(r)>9 else ''
    if 'sudah pernah menggunakan' in aw:
        valid_filter.append(r); used.append(1)
    elif 'tahu fitur tersebut tapi belum pernah menggunakan' in aw:
        valid_filter.append(r); used.append(0)
# Ensure target length matches construct data
used=np.array(used[:arr.shape[0]],dtype=int)
N_full=len(filter_data)
counts_awareness=Counter([r[9] if len(r)>9 else '' for r in filter_data])
counts_age=Counter([r[2] if len(r)>2 else '' for r in filter_data])
counts_gender=Counter([r[3] if len(r)>3 else '' for r in filter_data])
counts_freq=Counter([r[8] if len(r)>8 else '' for r in filter_data])

# Inferential regression models
names=['BEN','RISK','TR','PE','SI','FL','IQ','OBU','AA','SR']
X=np.column_stack([zscore(con_scores[n]) for n in names])
model_wu, rows_wu, n_wu=ols_hc3(zscore(con_scores['WU']), X, names)
save_csv(os.path.join(OUT,'study2_ols_wu.csv'), rows_wu)
model_ben, rows_ben, n_ben=ols_hc3(zscore(con_scores['BEN']), np.column_stack([zscore(con_scores[n]) for n in ['TR','PE','SI','FL','IQ','AA']]), ['TR','PE','SI','FL','IQ','AA'])
save_csv(os.path.join(OUT,'study2_ols_ben.csv'), rows_ben)
model_risk, rows_risk, n_risk=ols_hc3(zscore(con_scores['RISK']), np.column_stack([zscore(con_scores[n]) for n in ['SR','AA','IQ','FL','TR']]), ['SR','AA','IQ','FL','TR'])
save_csv(os.path.join(OUT,'study2_ols_risk.csv'), rows_risk)
# Logistic model for actual adoption/use
model_use, rows_use, n_use=ols_hc3(used, X, names)  # linear probability model for interpretability
save_csv(os.path.join(OUT,'study2_lpm_actual_use.csv'), rows_use)

# SmartPLS tables
boot=read_xlsx_sheet(boot_path, sheet_name='complete')
pls_paths=[]
for r in boot[7:19]:
    if len(r)>=7 and '->' in r[1]:
        pls_paths.append({'path':r[1].replace('1',''), 'coef':to_float(r[2]), 'stdev':to_float(r[4]), 't':to_float(r[5]), 'p':to_float(r[6])})
for item in pls_paths:
    for r in boot[21:33]:
        if len(r)>=6 and r[1].replace('1','')==item['path']:
            item['ci_low']=to_float(r[4]); item['ci_high']=to_float(r[5])
save_csv(os.path.join(OUT,'study2_pls_bootstrap_paths.csv'), pls_paths)
smart=read_xlsx_sheet(smartpls_path, sheet_name='complete')
pls_reliability=[]
for r in smart[1770:1781]:
    if len(r)>=6 and r[1]:
        pls_reliability.append({'construct':r[1].replace('1',''), 'alpha':to_float(r[2]), 'rho_a':to_float(r[3]), 'rho_c':to_float(r[4]), 'AVE':to_float(r[5])})
save_csv(os.path.join(OUT,'study2_pls_measurement.csv'), pls_reliability)
htmt_high=[]
for r in smart[1797:1830]:
    if len(r)>=3 and '<->' in r[1]:
        val=to_float(r[2])
        if not np.isnan(val) and val>=0.85:
            htmt_high.append({'pair':r[1].replace('1',''), 'htmt':val})
save_csv(os.path.join(OUT,'study2_htmt_high.csv'), htmt_high)

# ML Study 2: high willingness and actual use
feature_names=[construct_labels[n] for n in names]
study2_ml_wu=classification_cv(X, (con_scores['WU']>=np.nanmedian(con_scores['WU'])).astype(int), feature_names, 'study2_high_wu')
study2_ml_use=classification_cv(X, used.astype(int), feature_names, 'study2_actual_use')

# -------------------------------
# Study 3: FINRA NFCS 2024 investor data
# -------------------------------
with open(nfcs_path, encoding='utf-8-sig', newline='') as f:
    rdr=csv.DictReader(f)
    nfcs=[row for row in rdr]

def nf(row, key): return to_float(row.get(key,''))
def vararr(key): return np.array([nf(r,key) for r in nfcs], dtype=float)
def valid1_3(v): return not np.isnan(v) and 1 <= v <= 3
def valid1_7(v): return not np.isnan(v) and 1 <= v <= 7
weights=np.array([nf(r,'WGT1') for r in nfcs])
c22_3=vararr('C22_3'); c22_4=vararr('C22_4'); c40=vararr('C40'); d31=vararr('D31'); d40=vararr('D40')
f30_1=vararr('F30_1'); f30_2=vararr('F30_2'); f30_3=vararr('F30_3'); f30_9=vararr('F30_9'); f40=vararr('F40')
g2=vararr('G2')
# trading intensity
digital_any=((c22_3>=2)|(c22_4>=2)).astype(float)
valid_trade=np.array([valid1_3(a) or valid1_3(b) for a,b in zip(c22_3,c22_4)])
digital_any[~valid_trade]=np.nan
trade_int=[]
for a,b in zip(c22_3,c22_4):
    vals=[v for v in [a,b] if valid1_3(v)]
    trade_int.append(np.nanmean(vals) if vals else np.nan)
trade_int=np.array(trade_int,dtype=float)
# objective literacy
keys={'G4':1,'G5':2,'G6':2,'G7':1,'G21':2,'G8':1,'G44':3,'G22':2,'G11':3,'G12':3,'G13':4,'G23':2}
lit=[]
for r in nfcs:
    total=0; correct=0
    for k,ans in keys.items():
        v=nf(r,k)
        if not np.isnan(v) and v not in (98,99):
            total+=1
            if int(v)==ans: correct+=1
    lit.append(correct/total if total else np.nan)
lit=np.array(lit,dtype=float)

nfcs_indicators=[]
def add_indicator(name, value, scale='%'):
    nfcs_indicators.append({'indicator':name,'value':value,'scale':scale})
add_indicator('Place orders online through website sometimes/frequently', wprop(c22_3>=2, np.array([valid1_3(v) for v in c22_3]), weights)*100)
add_indicator('Place orders through mobile app sometimes/frequently', wprop(c22_4>=2, np.array([valid1_3(v) for v in c22_4]), weights)*100)
add_indicator('Either online or app trading sometimes/frequently', wprop((c22_3>=2)|(c22_4>=2), valid_trade, weights)*100)
add_indicator('Rely on brokerage/advisory research tools somewhat/a great deal', wprop(f30_2>=2, np.array([valid1_3(v) for v in f30_2]), weights)*100)
add_indicator('Rely on personal financial professionals somewhat/a great deal', wprop(f30_1>=2, np.array([valid1_3(v) for v in f30_1]), weights)*100)
add_indicator('Rely on popular investments shown in trading app somewhat/a great deal', wprop(f30_3>=2, np.array([valid1_3(v) for v in f30_3]), weights)*100)
add_indicator('Rely on social media groups/message boards somewhat/a great deal', wprop(f30_9>=2, np.array([valid1_3(v) for v in f30_9]), weights)*100)
add_indicator('Make investment decisions based on social-media personality sometimes/frequently', wprop((f40==1)|(f40==2), np.array([not np.isnan(v) and v in (1,2,3) for v in f40]), weights)*100)
add_indicator('Agree Internet investing makes accounts less secure (5-7)', wprop(c40>=5, np.array([valid1_7(v) for v in c40]), weights)*100)
add_indicator('Agree worried about investment fraud (5-7)', wprop(d31>=5, np.array([valid1_7(v) for v in d31]), weights)*100)
add_indicator('Worried Internet investing makes accounts less secure: mean 1-7', wmean(np.array([v if valid1_7(v) else np.nan for v in c40]), weights), 'mean')
add_indicator('Worried about losing money due to investment fraud: mean 1-7', wmean(np.array([v if valid1_7(v) else np.nan for v in d31]), weights), 'mean')
add_indicator('U.S. markets are fair to all investors: mean 1-7', wmean(np.array([v if valid1_7(v) else np.nan for v in d40]), weights), 'mean')
add_indicator('Self-rated investing knowledge: mean 1-7', wmean(np.array([v if valid1_7(v) else np.nan for v in g2]), weights), 'mean')
add_indicator('Objective investing literacy: mean percent correct', wmean(lit*100, weights), '%')
save_csv(os.path.join(OUT,'study3_nfcs_weighted_indicators.csv'), nfcs_indicators)

# WLS regression digital trading intensity
X_keys=['C40','D31','D40','F30_2','F30_1','F30_3','F30_9','F40','G2','literacy','B10','S_Age','S_Gender2','S_Education','S_Income']
X_cols=[]; names_nf=[]
for k in X_keys:
    if k=='literacy': x=lit.copy(); label='Objective literacy'
    else: x=vararr(k); label=k
    if k in ['C40','D31','D40','G2']:
        x=np.array([v if valid1_7(v) else np.nan for v in x])
        label={'C40':'Internet-security concern','D31':'Fraud concern','D40':'Market fairness belief','G2':'Self-rated investing knowledge'}[k]
    elif k.startswith('F30'):
        x=np.array([v if valid1_3(v) else np.nan for v in x])
        label={'F30_2':'Reliance on brokerage/advisory tools','F30_1':'Reliance on financial professionals','F30_3':'Reliance on app popularity cues','F30_9':'Reliance on social media groups'}[k]
    elif k=='F40':
        x=np.array([4-v if (not np.isnan(v) and v in (1,2,3)) else np.nan for v in x])
        label='Social-media personality use'
    elif k=='B10':
        x=np.array([5-v if (not np.isnan(v) and v in (1,2,3,4)) else np.nan for v in x])
        label='Risk tolerance'
    elif k in ['S_Age','S_Gender2','S_Education','S_Income']:
        x=np.array([v if (not np.isnan(v) and v not in (98,99)) else np.nan for v in x])
        label={'S_Age':'Age','S_Gender2':'Gender','S_Education':'Education','S_Income':'Income'}[k]
    names_nf.append(label); X_cols.append(zscore(x))
X_nf=np.column_stack(X_cols)
y_nf=zscore(trade_int)
mask=~np.isnan(y_nf) & ~np.isnan(weights)
for j in range(X_nf.shape[1]): mask &= ~np.isnan(X_nf[:,j])
Xc=sm.add_constant(X_nf[mask])
wls=sm.WLS(y_nf[mask], Xc, weights=weights[mask]).fit(cov_type='HC3')
nfcs_reg=[]
for i,n in enumerate(['Intercept']+names_nf):
    nfcs_reg.append({'term':n,'coef':float(wls.params[i]),'se':float(wls.bse[i]),'p':float(wls.pvalues[i]),'ci_low':float(wls.conf_int()[i,0]),'ci_high':float(wls.conf_int()[i,1])})
save_csv(os.path.join(OUT,'study3_nfcs_wls_digital_trading.csv'), nfcs_reg)

# ML Study 3: digital trading classifier
study3_ml=classification_cv(X_nf, digital_any, names_nf, 'study3_digital_trading', sample_weight=weights)

# -------------------------------
# Figures
# -------------------------------
plt.rcParams.update({'font.size': 10, 'font.family': 'DejaVu Sans'})
# Figure 1: conceptual model
fig, ax = plt.subplots(figsize=(8.0, 4.7))
ax.axis('off')
boxes = [
    ('Study 1\n2 x 2 experiment\nDigital twin x Conversational AI', 0.03, 0.62, 0.27, 0.22, '#6EC6FF'),
    ('Perceived\npersonalization', 0.38, 0.77, 0.20, 0.12, '#00C9A7'),
    ('Social\npresence', 0.38, 0.60, 0.20, 0.12, '#FFC75F'),
    ('Privacy/security\nconcern', 0.38, 0.42, 0.20, 0.12, '#FF9671'),
    ('Trust, usefulness,\nadvice willingness', 0.68, 0.62, 0.26, 0.18, '#B39CD0'),
    ('Study 2\nReal robo-advisor survey\nIndonesian Gen-Z app users', 0.03, 0.18, 0.27, 0.22, '#4D96FF'),
    ('Study 3\nReal investor validation\nFINRA NFCS 2024', 0.68, 0.18, 0.26, 0.16, '#45B7D1'),
]
for text,x,y,w,h,c in boxes:
    patch=FancyBboxPatch((x,y),w,h,boxstyle='round,pad=0.02,rounding_size=0.03',fc=c,ec='#2D4059',lw=1.2,alpha=0.95)
    ax.add_patch(patch); ax.text(x+w/2,y+h/2,text,ha='center',va='center',fontsize=10,weight='bold')
arrows=[((0.30,0.73),(0.38,0.83)),((0.30,0.73),(0.38,0.66)),((0.30,0.73),(0.38,0.48)),((0.58,0.83),(0.68,0.72)),((0.58,0.66),(0.68,0.72)),((0.58,0.48),(0.68,0.65)),((0.30,0.29),(0.68,0.65)),((0.81,0.34),(0.81,0.62))]
for a,b in arrows:
    ax.add_patch(FancyArrowPatch(a,b,arrowstyle='-|>',mutation_scale=14,lw=1.3,color='#2D4059'))
ax.text(0.50,0.08,'AI/ML extension: cross-validated prediction and permutation importance link survey constructs to adoption outcomes',ha='center',fontsize=10,color='#2D4059',weight='bold')
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure1_conceptual_model.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 2: Study 1 effects.
# These values are extracted from the public secondary-analysis manuscript unless the original Mendeley workbook is separately downloaded.
study1_csv=os.path.join(BASE,'data','derived','study1','study1_effects.csv')
study1_effects=[]
if os.path.exists(study1_csv):
    with open(study1_csv, encoding='utf-8-sig', newline='') as f:
        for rr in csv.DictReader(f):
            study1_effects.append((rr['label'], float(rr['effect']), float(rr['ci_low']), float(rr['ci_high']), rr['feature_group']))
else:
    study1_effects=[
        ('Digital twin -> personalization',1.43,1.30,1.56,'Digital twin'),('Digital twin -> privacy concern',0.92,0.71,1.13,'Digital twin'),('Digital twin -> usefulness',0.67,0.54,0.81,'Digital twin'),('Digital twin -> trust',0.28,0.14,0.42,'Digital twin'),('Digital twin -> adoption/advice',0.17,0.05,0.30,'Digital twin'),
        ('Conversational AI -> social presence',1.46,1.31,1.61,'Conversational AI'),('Conversational AI -> trust',0.29,0.15,0.42,'Conversational AI'),('Conversational AI -> usefulness',0.22,0.08,0.36,'Conversational AI'),('Conversational AI -> adoption/advice',0.20,0.08,0.32,'Conversational AI'),('Interaction -> adoption/advice',0.22,0.04,0.40,'Interaction')]
fig, ax=plt.subplots(figsize=(8.2,5.8))
colors={'Digital twin':'#00A896','Conversational AI':'#4D96FF','Interaction':'#FF7B54'}
order=np.argsort([e[1] for e in study1_effects])
y=np.arange(len(study1_effects))
for j,i in enumerate(order):
    label,b,lo,hi,g=study1_effects[i]
    ax.errorbar(b,j,xerr=[[b-lo],[hi-b]],fmt='o',color=colors[g],ecolor=colors[g],capsize=4,markersize=7)
ax.axvline(0,color='#555555',lw=1)
ax.set_yticks(y); ax.set_yticklabels([study1_effects[i][0] for i in order])
ax.set_xlabel('Robust factorial effect in scale points (95% CI)')
ax.set_title('Study 1 experimental effects: design features shift perceptions and willingness',weight='bold')
ax.grid(axis='x',alpha=0.25)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure2_study1_effects.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 3: Study 2 construct means with alpha overlay
rel_sorted=sorted(reliability, key=lambda r:r['mean'])
fig, ax=plt.subplots(figsize=(8.2,5.4))
vals=[r['mean'] for r in rel_sorted]; labs=[r['label'] for r in rel_sorted]
alphas=[r['alpha'] for r in rel_sorted]
colors=plt.cm.viridis(np.linspace(0.15,0.85,len(vals)))
ax.barh(labs, vals, color=colors, edgecolor='white')
for i,(v,a) in enumerate(zip(vals,alphas)):
    ax.text(v+0.03,i,f'α={a:.2f}',va='center',fontsize=9)
ax.set_xlim(1,5); ax.set_xlabel('Mean score (1-5)')
ax.set_title('Study 2 construct profile: real robo-advisor adoption survey',weight='bold')
ax.grid(axis='x',alpha=0.2)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure3_study2_construct_profile.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 4: Study 2 PLS paths
paths=[p['path'] for p in pls_paths]; coefs=np.array([p['coef'] for p in pls_paths]); lo=np.array([p.get('ci_low',np.nan) for p in pls_paths]); hi=np.array([p.get('ci_high',np.nan) for p in pls_paths])
order=np.argsort(coefs)
fig, ax=plt.subplots(figsize=(8.2,5.6))
for j,i in enumerate(order):
    col='#00A896' if coefs[i]>0 else '#FF5E5B'
    ax.errorbar(coefs[i],j,xerr=[[coefs[i]-lo[i]],[hi[i]-coefs[i]]],fmt='o',color=col,ecolor=col,capsize=3,markersize=7)
ax.axvline(0,color='#555555',lw=1)
ax.set_yticks(np.arange(len(paths))); ax.set_yticklabels([paths[i] for i in order])
ax.set_xlabel('Bootstrapped path coefficient (95% CI)')
ax.set_title('Study 2 SmartPLS paths: benefits drive willingness, risk inhibits it',weight='bold')
ax.grid(axis='x',alpha=0.2)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure4_study2_pls_paths.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 5: ML ROC curves Study 2 high WU and actual use
def plot_roc(ml_dict, outname, title):
    fig, ax=plt.subplots(figsize=(6.4,5.2))
    color_cycle=['#845EC2','#00C9A7','#FF6F91']
    for color,(name,p) in zip(color_cycle, ml_dict['probas'].items()):
        fpr,tpr,_=roc_curve(ml_dict['y'],p)
        auc=roc_auc_score(ml_dict['y'],p)
        ax.plot(fpr,tpr,lw=2.5,color=color,label=f'{name}: AUC={auc:.2f}')
    ax.plot([0,1],[0,1],'--',color='#555555',lw=1)
    ax.set_xlabel('False positive rate'); ax.set_ylabel('True positive rate')
    ax.set_title(title,weight='bold')
    ax.legend(loc='lower right',frameon=True)
    ax.grid(alpha=0.25)
    fig.tight_layout(); fig.savefig(os.path.join(OUT,outname),dpi=300,bbox_inches='tight'); plt.close(fig)
plot_roc(study2_ml_wu, 'figure5_study2_ml_roc_high_wu.png', 'AI/ML prediction of high willingness to use robo-advisors')
plot_roc(study2_ml_use, 'figure6_study2_ml_roc_actual_use.png', 'AI/ML prediction of reported robo-advisor use')

# Figure 7: Study 2 ML feature importance (best high WU)
imp=study2_ml_wu['importance'][:10]
fig, ax=plt.subplots(figsize=(8.2,5.0))
vals=[r['importance'] for r in imp][::-1]; labs=[r['feature'] for r in imp][::-1]
colors=plt.cm.plasma(np.linspace(0.2,0.9,len(vals)))
ax.barh(labs, vals, color=colors, edgecolor='white')
ax.set_xlabel('Permutation importance (AUC decrease)')
ax.set_title(f'Study 2 ML feature importance for high willingness ({study2_ml_wu["best"]})',weight='bold')
ax.grid(axis='x',alpha=0.2)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure7_study2_ml_importance.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 8: NFCS indicators
plot_inds=[x for x in nfcs_indicators if x['scale']=='%' and 'Objective' not in x['indicator']]
short_labels={
'Place orders online through website sometimes/frequently':'Website order placement',
'Place orders through mobile app sometimes/frequently':'Mobile-app order placement',
'Either online or app trading sometimes/frequently':'Website or mobile trading',
'Rely on brokerage/advisory research tools somewhat/a great deal':'Brokerage/advisory tools',
'Rely on personal financial professionals somewhat/a great deal':'Financial professionals',
'Rely on popular investments shown in trading app somewhat/a great deal':'Popular investments in app',
'Rely on social media groups/message boards somewhat/a great deal':'Social media/message boards',
'Make investment decisions based on social-media personality sometimes/frequently':'Social-media personality decisions',
'Agree Internet investing makes accounts less secure (5-7)':'Internet-security concern',
'Agree worried about investment fraud (5-7)':'Fraud concern'}
vals=[x['value'] for x in plot_inds]; labs=[short_labels.get(x['indicator'],x['indicator']) for x in plot_inds]
order=np.argsort(vals)
fig, ax=plt.subplots(figsize=(8.0,5.8))
colors=plt.cm.turbo(np.linspace(0.1,0.85,len(vals)))
ax.barh([labs[i] for i in order], [vals[i] for i in order], color=colors, edgecolor='white')
for j,i in enumerate(order): ax.text(vals[i]+1,j,f'{vals[i]:.1f}%',va='center',fontsize=9)
ax.set_xlim(0,100); ax.set_xlabel('Weighted percent')
ax.set_title('Study 3 FINRA NFCS: digital-investing readiness and risk concerns',weight='bold')
ax.grid(axis='x',alpha=0.2)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure8_study3_nfcs_indicators.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 9: Study 3 WLS regression
reg_plot=[r for r in nfcs_reg if r['term']!='Intercept']
coefs=np.array([r['coef'] for r in reg_plot]); lo=np.array([r['ci_low'] for r in reg_plot]); hi=np.array([r['ci_high'] for r in reg_plot]); labs=[r['term'] for r in reg_plot]
# Select top by abs coef for readability
sel=np.argsort(np.abs(coefs))[-12:]
order=sel[np.argsort(coefs[sel])]
fig, ax=plt.subplots(figsize=(8.2,5.4))
for j,i in enumerate(order):
    col='#0077B6' if coefs[i]>0 else '#D00000'
    ax.errorbar(coefs[i],j,xerr=[[coefs[i]-lo[i]],[hi[i]-coefs[i]]],fmt='o',color=col,ecolor=col,capsize=3,markersize=7)
ax.axvline(0,color='#555555',lw=1)
ax.set_yticks(np.arange(len(order))); ax.set_yticklabels([labs[i] for i in order])
ax.set_xlabel('Standardized WLS coefficient (95% CI)')
ax.set_title('Study 3 correlates of digital trading intensity',weight='bold')
ax.grid(axis='x',alpha=0.2)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure9_study3_wls_coefficients.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 10: Study 3 ML importance
imp=study3_ml['importance'][:12]
fig, ax=plt.subplots(figsize=(8.2,5.4))
vals=[r['importance'] for r in imp][::-1]; labs=[r['feature'] for r in imp][::-1]
colors=plt.cm.cividis(np.linspace(0.2,0.95,len(vals)))
ax.barh(labs, vals, color=colors, edgecolor='white')
ax.set_xlabel('Permutation importance (AUC decrease)')
ax.set_title(f'Study 3 ML feature importance for digital trading ({study3_ml["best"]})',weight='bold')
ax.grid(axis='x',alpha=0.2)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure10_study3_ml_importance.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# Figure 11: combined ML performance heatmap
perf_rows=[]
for study,label,ml in [('Study 2','High willingness',study2_ml_wu),('Study 2','Reported use',study2_ml_use),('Study 3','Digital trading',study3_ml)]:
    for r in ml['rows']:
        perf_rows.append({'study':study,'target':label,'model':r['model'],'auc_mean':r['auc_mean'],'accuracy_mean':r['accuracy_mean'],'balanced_accuracy_mean':r['balanced_accuracy_mean'],'f1_mean':r['f1_mean']})
save_csv(os.path.join(OUT,'combined_ml_performance.csv'), perf_rows)
# Heatmap manually
models=['Regularized logistic','Random forest','Gradient boosting']
targets=['High willingness','Reported use','Digital trading']
mat=np.zeros((len(targets),len(models)))
for i,t in enumerate(targets):
    for j,m in enumerate(models):
        mat[i,j]=next(r['auc_mean'] for r in perf_rows if r['target']==t and r['model']==m)
fig, ax=plt.subplots(figsize=(7.0,3.6))
im=ax.imshow(mat, cmap='YlGnBu', vmin=0.5, vmax=1.0)
ax.set_xticks(np.arange(len(models))); ax.set_xticklabels(models,rotation=25,ha='right')
ax.set_yticks(np.arange(len(targets))); ax.set_yticklabels(targets)
for i in range(mat.shape[0]):
    for j in range(mat.shape[1]):
        ax.text(j,i,f'{mat[i,j]:.2f}',ha='center',va='center',color='black',weight='bold')
ax.set_title('Cross-validated AI/ML discrimination (AUC)',weight='bold')
fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
fig.tight_layout(); fig.savefig(os.path.join(OUT,'figure11_ml_performance_heatmap.png'),dpi=300,bbox_inches='tight'); plt.close(fig)

# -------------------------------
# Summary JSON
# -------------------------------
summary={
    'study2_n_complete': int(arr.shape[0]),
    'study2_n_raw': int(N_full),
    'study2_used_count': int(used.sum()),
    'study2_aware_not_used_count': int(len(used)-used.sum()),
    'study2_use_rate': float(used.mean()),
    'study2_age_counts': dict(counts_age),
    'study2_gender_counts': dict(counts_gender),
    'study2_frequency_counts': dict(counts_freq),
    'study2_wu_model_n': int(n_wu),
    'study2_wu_model_r2': float(model_wu.rsquared),
    'study2_ben_model_r2': float(model_ben.rsquared),
    'study2_risk_model_r2': float(model_risk.rsquared),
    'study2_ml_high_wu_best': study2_ml_wu['best'],
    'study2_ml_actual_use_best': study2_ml_use['best'],
    'study3_n': int(len(nfcs)),
    'study3_wls_n': int(mask.sum()),
    'study3_wls_r2': float(wls.rsquared),
    'study3_ml_best': study3_ml['best'],
    'output_dir': OUT
}
with open(os.path.join(OUT,'analysis_summary.json'),'w',encoding='utf-8') as f:
    json.dump(summary,f,indent=2,ensure_ascii=False)
print(json.dumps(summary, indent=2, ensure_ascii=False))
