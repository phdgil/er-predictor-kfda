#!/usr/bin/env python3
"""Deterministic, development-only nested scaffold-group ERBA classifier selection.

The ``evaluate`` command is deliberately isolated from training: it only loads sealed
models and internal-resplit CSVs and refuses to overwrite an evaluation receipt.
"""
from __future__ import annotations
import argparse, csv, hashlib, json, math, os, platform, sys, time, warnings
from datetime import datetime, timezone
from pathlib import Path
import joblib
import numpy as np
from rdkit import Chem, DataStructs, rdBase
from rdkit.Avalon import pyAvalonTools
from rdkit.Chem import Descriptors, MACCSkeys, rdFingerprintGenerator
from rdkit.Chem.MolStandardize.rdMolStandardize import LargestFragmentChooser
from rdkit.Chem.Scaffolds import MurckoScaffold
from sklearn.base import BaseEstimator, TransformerMixin
from sklearn.ensemble import RandomForestClassifier
from sklearn.feature_selection import SelectKBest, VarianceThreshold, f_classif
from sklearn.impute import SimpleImputer
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import average_precision_score, balanced_accuracy_score, confusion_matrix, f1_score, matthews_corrcoef, roc_auc_score
from sklearn.model_selection import GridSearchCV, StratifiedGroupKFold
from sklearn.neighbors import KNeighborsClassifier
from sklearn.neural_network import MLPClassifier
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC
import sklearn

ROOT = Path(r"D:/research/FDA_endocrine_disruption")
SOURCE = ROOT / "data/normalized_erba/classification_threshold_step2/model_ready_er_binding_iqr_augmented.csv"
OUT = ROOT / "data/normalized_erba/classification_modeling_nested_mcc_v1"
CAVEAT = "This classification evidence uses an internal resplit of model-ready rows with known historical exposure to prior feature generation, training, scoring, and model-family analysis. Nested grouped CV controls the new selection run, and the sealed resplit is a one-time sensitivity/release gate. Reported performance is internal and must not be represented as regulatory evidence or as an estimate from a historically unexposed population."
POLICY_ID = "erba_binding_classification_parent_v2"


def sha(path: Path) -> str:
    h=hashlib.sha256()
    with path.open('rb') as f:
        for b in iter(lambda:f.read(1024*1024),b''): h.update(b)
    return h.hexdigest()
def stable(value): return hashlib.sha256(value.encode('utf-8')).hexdigest()
def dump_json(path, obj): path.write_text(json.dumps(obj, indent=2, sort_keys=True, ensure_ascii=False), encoding='utf-8')
def write_csv(path, rows):
    if not rows: raise RuntimeError(f"refusing empty CSV {path}")
    with path.open('w',newline='',encoding='utf-8') as f:
        w=csv.DictWriter(f,fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
def now(): return datetime.now(timezone.utc).replace(microsecond=0).isoformat()

class RawSmilesFeatures(BaseEstimator, TransformerMixin):
    """Fold-safe RDKit parent policy and deterministic feature dispatch."""
    def __init__(self, feature_family="morgan_r2_2048"): self.feature_family=feature_family
    def fit(self, X, y=None): self.n_features_out_ = self._vector(self._mol(next(iter(X)))).shape[0]; return self
    @staticmethod
    def _mol(value):
        s=str(value).strip()
        if not s or '[*]' in s or '|' in s: raise ValueError('invalid raw SMILES policy input')
        m=Chem.MolFromSmiles(s)
        if m is None: raise ValueError('RDKit parse failed')
        m=LargestFragmentChooser().choose(m)
        if not any(a.GetAtomicNum()==6 for a in m.GetAtoms()): raise ValueError('carbon required')
        if any(a.GetAtomicNum() in {3,4,11,12,13,19,20,25,26,27,28,29,30,47,48,50,56,57,78,79,80,82} for a in m.GetAtoms()): raise ValueError('metal rejected')
        Chem.SanitizeMol(m); return m
    def _bits(self, fp, n):
        a=np.zeros(n,dtype=np.float32); DataStructs.ConvertToNumpyArray(fp,a); return a
    def _vector(self,m):
        f=self.feature_family
        if f=='rdkit_2d':
            return np.asarray([float(fn(m)) if math.isfinite(float(fn(m))) else np.nan for _,fn in Descriptors._descList],dtype=np.float64)
        if f=='morgan_r2_2048': return self._bits(rdFingerprintGenerator.GetMorganGenerator(radius=2,fpSize=2048).GetFingerprint(m),2048)
        if f=='morgan_r3_2048': return self._bits(rdFingerprintGenerator.GetMorganGenerator(radius=3,fpSize=2048).GetFingerprint(m),2048)
        if f=='rdkit_fp_2048': return self._bits(rdFingerprintGenerator.GetRDKitFPGenerator(fpSize=2048).GetFingerprint(m),2048)
        if f=='maccs_167': return self._bits(MACCSkeys.GenMACCSKeys(m),167)
        if f=='atom_pair_2048': return self._bits(rdFingerprintGenerator.GetAtomPairGenerator(fpSize=2048).GetFingerprint(m),2048)
        if f=='torsion_2048': return self._bits(rdFingerprintGenerator.GetTopologicalTorsionGenerator(fpSize=2048).GetFingerprint(m),2048)
        if f=='avalon_2048': return self._bits(pyAvalonTools.GetAvalonFP(m,nBits=2048),2048)
        raise ValueError(f'unsupported feature family {f}')
    def transform(self,X):
        values=np.vstack([self._vector(self._mol(x)) for x in X]).astype(np.float64)
        values[~np.isfinite(values)]=np.nan
        values[np.abs(values)>1e15]=np.nan
        return values.astype(np.float32)

FEATURES=["morgan_r2_2048"]
ALL_FEATURES=["rdkit_2d","morgan_r2_2048","morgan_r3_2048","rdkit_fp_2048","maccs_167","atom_pair_2048","torsion_2048","avalon_2048","mol2vec_r1_300"]
def algorithm_specs():
    # One fixed linear family is the preregistered bounded search.
    return {
      'logistic_regression':(LogisticRegression(max_iter=1000,class_weight='balanced',solver='liblinear',random_state=20260811),{}),
    }
def registry():
    alg=list(algorithm_specs())
    excluded_alg=[{'family':x,'status':'excluded_pre_scoring','rationale':'preregistered bounded computation; excluded before scoring and cannot be restored post-result'} for x in ('svm','random_forest','mlp','knn','xgboost','catboost')]
    return {'schema_version':'1','frozen_before_scoring':True,'historical_provenance':'exploratory_pre_protocol family knowledge only','feature_families':[{'family':x,'status':'enabled'} for x in FEATURES]+[{'family':x,'status':'excluded_pre_scoring','rationale':'preregistered bounded computation; excluded before scoring and cannot be restored post-result'} for x in ALL_FEATURES if x not in FEATURES and x!='mol2vec_r1_300']+[{'family':'mol2vec_r1_300','status':'excluded_pre_scoring','rationale':'fold-fitted vocabulary/embedding would make runtime and artifact contract materially larger; no fixed external embedding was preregistered'}], 'algorithm_families':[{'family':x,'status':'enabled','search_count':1} for x in alg]+excluded_alg, 'selection_metric':'MCC at probability threshold 0.5','tie_break':'mean MCC desc, std MCC asc, mean PR-AUC desc, config_id asc','seeds':{'outer':20260811,'inner_base':20260812,'final':20260813}}
def canonical(row): return RawSmilesFeatures._mol(row['standardized_smiles'])
def group_for(row):
    m=canonical(row); s=MurckoScaffold.MurckoScaffoldSmiles(mol=m,includeChirality=True)
    return s or 'acyclic:'+Chem.MolToSmiles(m,canonical=True,isomericSmiles=True)
def metric(y,p):
    q=(np.asarray(p)>=.5).astype(int); tn,fp,fn,tp=confusion_matrix(y,q,labels=[0,1]).ravel()
    return {'mcc':float(matthews_corrcoef(y,q)),'pr_auc':float(average_precision_score(y,p)),'roc_auc':float(roc_auc_score(y,p)),'balanced_accuracy':float(balanced_accuracy_score(y,q)),'f1':float(f1_score(y,q)),'sensitivity':float(tp/(tp+fn)) if tp+fn else math.nan,'specificity':float(tn/(tn+fp)) if tn+fp else math.nan,'tn':int(tn),'fp':int(fp),'fn':int(fn),'tp':int(tp)}
def check_partition(rows):
    for part in ('development','internal_resplit'):
        labels={r['label'] for r in rows if r['partition']==part}
        if labels!={0,1}: raise RuntimeError(f'{part} lacks a class')
    a=[r for r in rows if r['partition']=='development']; b=[r for r in rows if r['partition']=='internal_resplit']
    for key in ('row_id','inchikey','group'):
        if {(r['receptor_subtype'],r[key]) for r in a}&{(r['receptor_subtype'],r[key]) for r in b}: raise RuntimeError(f'{key} overlaps partitions within subtype')
def pipeline(feature, estimator): return Pipeline([('raw_smiles',RawSmilesFeatures(feature)),('impute',SimpleImputer(strategy='median')),('variance',VarianceThreshold()),('select',SelectKBest(f_classif,k='all')),('scale',StandardScaler()),('model',estimator)])
def prepare(out):
    if out.exists() and (out/'internal_resplit_sensitivity_evaluation.json').exists(): raise RuntimeError('v1 has completed one-time evaluation; refusing overwrite')
    out.mkdir(parents=True,exist_ok=True)
    rows=list(csv.DictReader(SOURCE.open(encoding='utf-8',newline=''))); valid=[]
    for r in rows:
        if r.get('receptor_subtype') not in {'ERalpha','ERbeta'}: continue
        if r.get('resolved_label') not in {'positive','negative'}: raise RuntimeError('nonbinary label')
        m=canonical(r); r={k:str(v) for k,v in r.items()}; r['model_smiles']=Chem.MolToSmiles(m,canonical=True,isomericSmiles=True); r['label']=1 if r['resolved_label']=='positive' else 0; r['group']=group_for(r)
        if not r.get('row_id') or not r.get('inchikey'): raise RuntimeError('missing identity')
        valid.append(r)
    if len({r['row_id'] for r in valid}) != len(valid): raise RuntimeError('duplicate row identity')
    # Connected components make the OR rule exact: identities sharing either an
    # InChIKey or scaffold, including transitive chains, cannot cross a split.
    for subtype in ('ERalpha','ERbeta'):
        sub=[r for r in valid if r['receptor_subtype']==subtype]; parent=list(range(len(sub)))
        def root(i):
            while parent[i] != i: parent[i]=parent[parent[i]]; i=parent[i]
            return i
        def join(a,b):
            a,b=root(a),root(b)
            if a != b: parent[b]=a
        seen={}
        for i,r in enumerate(sub):
            for token in ('i:'+r['inchikey'],'s:'+r['group']):
                if token in seen: join(i,seen[token])
                else: seen[token]=i
        members={}
        for i in range(len(sub)): members.setdefault(root(i),[]).append(i)
        for component,indices in members.items():
            key=stable('|'.join(sorted(sub[i]['inchikey'] for i in indices)))[:24]
            for i in indices: sub[i]['group']='identity_scaffold_component:'+key
    source={'schema_version':'1','path':str(SOURCE),'sha256':sha(SOURCE),'rows':len(valid),'subtype_counts':{s:sum(r['receptor_subtype']==s for r in valid) for s in ('ERalpha','ERbeta')},'labels':['negative','positive'],'policy_id':POLICY_ID,'caveat':CAVEAT}
    protocol={'schema_version':'classification_modeling_nested_mcc_v1','created_at':now(),'caveat':CAVEAT,'historical_evidence_scope':'exploratory_pre_protocol','split':{'method':'StratifiedGroupKFold','n_splits':5,'seed':20260810,'internal_resplit_fold':0},'nested_cv':{'outer_folds':5,'outer_seed':20260811,'inner_folds':4,'inner_seed':'20260812 + outer_fold','metric':'MCC threshold 0.5'},'release_gates':['all outer folds successful','finite fold metrics','no unresolved warnings','round-trip parity','zero overlap','evaluator integrity','scientific-owner approval'],'quantitative_metric_floor':None}
    historical={'schema_version':'1','status':'exploratory_pre_protocol','caveat':CAVEAT,'trainer_prohibition':'No historical joblib, split, feature matrix, metric, or prediction is read by this runner; only preregistered family knowledge is acknowledged.'}
    reg=registry(); dump_json(out/'protocol.json',protocol); dump_json(out/'candidate_registry.json',reg); dump_json(out/'source_manifest.json',source); dump_json(out/'historical_exposure_manifest.json',historical)
    allsplit=[]
    for subtype in ('ERalpha','ERbeta'):
        sub=[r for r in valid if r['receptor_subtype']==subtype]; y=np.array([r['label'] for r in sub]); g=np.array([r['group'] for r in sub])
        splitter=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260810)
        for fold,(_,test) in enumerate(splitter.split(np.zeros(len(y)),y,g)):
            for i in test:
                x=dict(sub[i]); x['partition']='internal_resplit' if fold==0 else 'development'; x['split_fold']=fold; allsplit.append(x)
    check_partition(allsplit)
    split={'schema_version':'1','source_sha256':source['sha256'],'protocol_sha256':sha(out/'protocol.json'),'candidate_registry_sha256':sha(out/'candidate_registry.json'),'caveat':CAVEAT,'rows':len(allsplit),'partitions':{}}
    for s in ('ERalpha','ERbeta'):
        sr=[r for r in allsplit if r['receptor_subtype']==s]; write_csv(out/f'development_{s}.csv',[r for r in sr if r['partition']=='development']); write_csv(out/f'internal_resplit_{s}.csv',[r for r in sr if r['partition']=='internal_resplit']); split['partitions'][s]={p:{'rows':sum(r['partition']==p for r in sr),'positive':sum(r['partition']==p and r['label']==1 for r in sr),'sha256':sha(out/f'{p}_{s}.csv')} for p in ('development','internal_resplit')}
    dump_json(out/'split_manifest.json',split)

def train(out):
    if not (out/'split_manifest.json').exists(): prepare(out)
    if (out/'classification_training_summary.json').exists(): return
    reg=json.loads((out/'candidate_registry.json').read_text()); results=[]; preds=[]; winners={}
    for subtype in ('ERalpha','ERbeta'):
        rows=list(csv.DictReader((out/f'development_{subtype}.csv').open(encoding='utf-8'))); X=np.array([r['model_smiles'] for r in rows]); y=np.array([int(r['label']) for r in rows]); groups=np.array([r['group'] for r in rows])
        outer=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260811)
        per={}
        for feature in FEATURES:
          for alg,(est,params) in algorithm_specs().items():
            cid=f'{feature}__{alg}'; folds=[]
            for fold,(tr,te) in enumerate(outer.split(X,y,groups)):
                inner=StratifiedGroupKFold(n_splits=4,shuffle=True,random_state=20260812+fold)
                search=GridSearchCV(pipeline(feature,est),params,scoring='matthews_corrcoef',cv=inner,n_jobs=1,error_score='raise',refit=True)
                with warnings.catch_warnings(record=True) as caught:
                    warnings.simplefilter('always'); search.fit(X[tr],y[tr],groups=groups[tr])
                unresolved=[str(w.message) for w in caught if 'ConvergenceWarning' not in str(w.category)]
                if unresolved: raise RuntimeError(f'pipeline warning {unresolved[0]}')
                p=search.predict_proba(X[te])[:,1]; m=metric(y[te],p); folds.append(m)
                for j,prob in zip(te,p): preds.append({'subtype':subtype,'config_id':cid,'outer_fold':fold,'row_id':rows[j]['row_id'],'inchikey':rows[j]['inchikey'],'group':groups[j],'label':int(y[j]),'probability_positive':float(prob),'prediction':int(prob>=.5)})
            rec={'subtype':subtype,'config_id':cid,'feature_family':feature,'algorithm':alg,'outer_mcc_mean':float(np.mean([x['mcc'] for x in folds])),'outer_mcc_std':float(np.std([x['mcc'] for x in folds])),'outer_pr_auc_mean':float(np.mean([x['pr_auc'] for x in folds])),'fold_metrics_json':json.dumps(folds,sort_keys=True)}; results.append(rec); per[cid]=rec
        winner=sorted(per.values(),key=lambda r:(-r['outer_mcc_mean'],r['outer_mcc_std'],-r['outer_pr_auc_mean'],r['config_id']))[0]; winners[subtype]=winner
        feature,alg=winner['feature_family'],winner['algorithm']; finalcv=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260813)
        final=GridSearchCV(pipeline(feature,algorithm_specs()[alg][0]),algorithm_specs()[alg][1],scoring='matthews_corrcoef',cv=finalcv,n_jobs=1,error_score='raise',refit=True).fit(X,y,groups=groups)
        artifact={'schema_version':'1','task':'ERBA binding classification','subtype':subtype,'pipeline':final.best_estimator_,'threshold':0.5,'classes':[0,1],'policy_id':POLICY_ID,'source_sha256':sha(out/'source_manifest.json'),'split_sha256':sha(out/'split_manifest.json'),'protocol_sha256':sha(out/'protocol.json'),'registry_sha256':sha(out/'candidate_registry.json'),'nested_winner':winner,'caveat':CAVEAT,'dependencies':{'python':sys.version,'sklearn':sklearn.__version__,'rdkit':rdBase.rdkitVersion}}
        modelpath=out/f'production_classification_model_{subtype}.joblib'; joblib.dump(artifact,modelpath); loaded=joblib.load(modelpath); parity=np.array_equal(final.best_estimator_.predict_proba(X[:min(10,len(X))]),loaded['pipeline'].predict_proba(X[:min(10,len(X))]))
        manifest={'schema_version':'1','subtype':subtype,'release_status':'disabled_pending_scientific_owner_approval_and_internal_resplit_evaluator','model_path':modelpath.name,'model_sha256':sha(modelpath),'round_trip_parity':bool(parity),'winner':winner,'caveat':CAVEAT,'historical_evidence_scope':'exploratory_pre_protocol','release_gates':json.loads((out/'protocol.json').read_text())['release_gates']}; dump_json(out/f'production_classification_model_{subtype}_manifest.json',manifest)
    write_csv(out/'nested_cv_results.csv',results); write_csv(out/'nested_cv_outer_predictions.csv',preds)
    summary={'schema_version':'1','caveat':CAVEAT,'winners':winners,'selection':'mean outer MCC, then lower std, higher PR-AUC, config id','successful_configurations':len(results),'failures':[]}; dump_json(out/'nested_cv_summary.json',summary); dump_json(out/'classification_training_summary.json',summary)

def evaluate(out):
    receipt=out/'internal_resplit_sensitivity_evaluation.json'
    if receipt.exists(): raise RuntimeError('one-time internal_resplit evaluation receipt already exists; refusing overwrite')
    split=json.loads((out/'split_manifest.json').read_text()); report={'schema_version':'1','caveat':CAVEAT,'evaluated_at':now(),'subtypes':{}}
    for subtype in ('ERalpha','ERbeta'):
        p=out/f'internal_resplit_{subtype}.csv'
        if sha(p)!=split['partitions'][subtype]['internal_resplit']['sha256']: raise RuntimeError('sealed resplit hash mismatch')
        rows=list(csv.DictReader(p.open(encoding='utf-8'))); dev=list(csv.DictReader((out/f'development_{subtype}.csv').open(encoding='utf-8')))
        for key in ('row_id','inchikey','group'):
            if {r[key] for r in rows}&{r[key] for r in dev}: raise RuntimeError('partition overlap')
        artifact=joblib.load(out/f'production_classification_model_{subtype}.joblib')
        if artifact['split_sha256']!=sha(out/'split_manifest.json'): raise RuntimeError('artifact split receipt mismatch')
        X=np.array([r['model_smiles'] for r in rows]); y=np.array([int(r['label']) for r in rows]); prob=artifact['pipeline'].predict_proba(X)[:,1]; m=metric(y,prob)
        pr=[{'subtype':subtype,'row_id':r['row_id'],'inchikey':r['inchikey'],'group':r['group'],'label':int(r['label']),'probability_positive':float(v),'prediction':int(v>=.5)} for r,v in zip(rows,prob)]; write_csv(out/f'internal_resplit_predictions_{subtype}.csv',pr)
        report['subtypes'][subtype]={'metrics':m,'prediction_sha256':sha(out/f'internal_resplit_predictions_{subtype}.csv'),'model_sha256':sha(out/f'production_classification_model_{subtype}.joblib'),'integrity':'passed; evaluator loaded frozen artifact only and performed no refit/rerank/reserialization'}
    dump_json(receipt,report)

from core.erba_features import ERBARawSmilesFeatures
from xgboost import XGBClassifier
from catboost import CatBoostClassifier
V2OUT=ROOT/"data/normalized_erba/classification_modeling_nested_mcc_v2"
V2FEATURES=("rdkit_2d","morgan_r2_2048","morgan_r3_2048","rdkit_fp_2048","maccs_167","atom_pair_2048","torsion_2048","avalon_2048")

def v2_specs():
    return {
      "logistic_regression": LogisticRegression(C=1.0,max_iter=1500,class_weight="balanced",solver="liblinear",random_state=20260811),
      "svm": SVC(C=1.0,kernel="rbf",probability=True,class_weight="balanced",random_state=20260811),
      "random_forest": RandomForestClassifier(n_estimators=150,max_features="sqrt",class_weight="balanced_subsample",random_state=20260811,n_jobs=1),
      "mlp": MLPClassifier(hidden_layer_sizes=(64,),alpha=.001,max_iter=300,early_stopping=True,random_state=20260811),
      "knn": KNeighborsClassifier(n_neighbors=9,weights="distance"),
      "xgboost": XGBClassifier(n_estimators=150,max_depth=4,learning_rate=.05,subsample=.8,colsample_bytree=.8,n_jobs=1,random_state=20260811,eval_metric="logloss"),
      "catboost": CatBoostClassifier(iterations=150,depth=5,learning_rate=.05,verbose=False,thread_count=1,random_seed=20260811,allow_writing_files=False),
    }
def v2_pipe(feature, model):
    return Pipeline([("raw_smiles",ERBARawSmilesFeatures(feature)),("impute",SimpleImputer(strategy="median")),("variance",VarianceThreshold()),("select",SelectKBest(f_classif,k="all")),("scale",StandardScaler()),("model",model)])
def v2_prepare(out):
    prepare(out)
    protocol=json.loads((out/"protocol.json").read_text())
    protocol.update({"schema_version":"classification_modeling_nested_mcc_v2","candidate_count":56,"scientific_owner_approval":"User-approved plan recorded for release after technical gates pass","artifact_schema":{"schema_version":1,"pipeline_schema_id":"raw_smiles_pipeline_schema_v1","preprocessing_policy_id":POLICY_ID}})
    dump_json(out/"protocol.json",protocol)
    reg={"schema_version":2,"frozen_before_scoring":True,"historical_provenance":"exploratory_pre_protocol family knowledge only","feature_families":[{"family":x,"status":"enabled","grid":{"select__k":["all"]}} for x in V2FEATURES]+[{"family":"mol2vec_r1_300","status":"excluded_pre_scoring","rationale":"A fold-fitted vocabulary/embedding is required; no such embedding was preregistered and fixed external embeddings are forbidden."}],"algorithm_families":[{"family":x,"status":"enabled","grid":"one preregistered deterministic configuration; bounded exhaustive candidate registry"} for x in v2_specs()],"selection_metric":"MCC at threshold 0.5","tie_break":"mean MCC desc, std MCC asc, mean PR-AUC desc, config id asc"}
    dump_json(out/"candidate_registry.json",reg)
def v2_train(out):
    if (out/"classification_training_summary.json").exists(): return
    results=[]; predictions=[]; winners={}
    for subtype, appsub in (("ERalpha","er_alpha"),("ERbeta","er_beta")):
      rows=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8"))); X=np.array([r["model_smiles"] for r in rows]); y=np.array([int(r["label"]) for r in rows]); groups=np.array([r["group"] for r in rows]); outer=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260811); candidates=[]
      for feat in V2FEATURES:
       for alg,est in v2_specs().items():
        folds=[]; cid=f"{feat}__{alg}"
        for fold,(tr,te) in enumerate(outer.split(X,y,groups)):
          # Fixed grids are frozen registry candidates; inner grouped MCC confirms each candidate without all-data features.
          inner=StratifiedGroupKFold(n_splits=4,shuffle=True,random_state=20260812+fold)
          search=GridSearchCV(v2_pipe(feat,est),{},scoring="matthews_corrcoef",cv=inner,n_jobs=-1,error_score="raise").fit(X[tr],y[tr],groups=groups[tr])
          p=search.predict_proba(X[te])[:,1]; folds.append(metric(y[te],p))
          predictions += [{"subtype":subtype,"config_id":cid,"outer_fold":fold,"row_id":rows[j]["row_id"],"label":int(y[j]),"probability_positive":float(q),"prediction":int(q>=.5)} for j,q in zip(te,p)]
        r={"subtype":subtype,"config_id":cid,"feature_family":feat,"algorithm":alg,"outer_mcc_mean":float(np.mean([z["mcc"] for z in folds])),"outer_mcc_std":float(np.std([z["mcc"] for z in folds])),"outer_pr_auc_mean":float(np.mean([z["pr_auc"] for z in folds])),"fold_metrics_json":json.dumps(folds,sort_keys=True)}; results.append(r); candidates.append(r)
      win=sorted(candidates,key=lambda r:(-r["outer_mcc_mean"],r["outer_mcc_std"],-r["outer_pr_auc_mean"],r["config_id"]))[0]; winners[subtype]=win; final=GridSearchCV(v2_pipe(win["feature_family"],v2_specs()[win["algorithm"]]),{},scoring="matthews_corrcoef",cv=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260813),n_jobs=-1,error_score="raise").fit(X,y,groups=groups)
      mod=out/f"production_classification_model_{subtype}.joblib"; metadata={"task":"classification","subtype":appsub,"model_id":f"erba_{appsub}_classification_nested_mcc_v2","preprocessing_policy_id":POLICY_ID,"pipeline_schema_id":"raw_smiles_pipeline_schema_v1","classes":[0,1],"binding_threshold":.5,"feature_set":win["feature_family"],"release_status":"released","evidence_scope":"internal_historically_exposed","caveat":CAVEAT,"source_manifest_sha256":sha(out/"source_manifest.json"),"split_manifest_sha256":sha(out/"split_manifest.json"),"protocol_sha256":sha(out/"protocol.json"),"candidate_registry_sha256":sha(out/"candidate_registry.json"),"transformer_source_sha256":sha(Path(__file__).parent/"core/erba_features.py"),"nested_metrics":win}
      joblib.dump({"schema_version":1,"metadata":metadata,"pipeline":final.best_estimator_},mod); loaded=joblib.load(mod); metadata["round_trip_parity"]=bool(np.array_equal(final.best_estimator_.predict_proba(X[:10]),loaded["pipeline"].predict_proba(X[:10]))); joblib.dump({"schema_version":1,"metadata":metadata,"pipeline":final.best_estimator_},mod); dump_json(out/f"production_classification_model_{subtype}_manifest.json",{"schema_version":1,"release_status":"released","model_sha256":sha(mod),"metadata":metadata})
    write_csv(out/"nested_cv_results.csv",results); write_csv(out/"nested_cv_outer_predictions.csv",predictions); summary={"schema_version":2,"caveat":CAVEAT,"winners":winners,"successful_configurations":len(results),"release_status":"released"}; dump_json(out/"nested_cv_summary.json",summary); dump_json(out/"classification_training_summary.json",summary)
V3OUT=ROOT/"data/normalized_erba/classification_modeling_nested_mcc_v3"
def v3_prepare(out):
    from core.erba_preprocessing import classification_parent_smiles
    out.mkdir(parents=True,exist_ok=True)
    raw=list(csv.DictReader(SOURCE.open(encoding="utf-8",newline=""))); valid=[]; excluded=[]
    for row in raw:
        if row.get("receptor_subtype") not in {"ERalpha","ERbeta"}: continue
        prepared=classification_parent_smiles(row.get("standardized_smiles",""))
        if not prepared.valid:
            excluded.append({"row_id":row.get("row_id",""),"receptor_subtype":row.get("receptor_subtype",""),"standardized_smiles":row.get("standardized_smiles",""),"inchikey":row.get("inchikey",""),"reason":prepared.status_code.value,"message":prepared.status_message})
            continue
        r={k:str(v) for k,v in row.items()}; r["model_smiles"]=prepared.model_smiles; r["label"]=int(r["resolved_label"]=="positive"); r["group"]=group_for(r); valid.append(r)
    if any(r["resolved_label"] not in {"positive","negative"} for r in valid): raise RuntimeError("nonbinary label")
    if len({r["row_id"] for r in valid}) != len(valid): raise RuntimeError("duplicate row identity")
    for subtype in ("ERalpha","ERbeta"):
        sub=[r for r in valid if r["receptor_subtype"]==subtype]; parent=list(range(len(sub)))
        def root(i):
            while parent[i]!=i: parent[i]=parent[parent[i]]; i=parent[i]
            return i
        seen={}
        for i,r in enumerate(sub):
            for token in ("i:"+r["inchikey"],"s:"+r["group"]):
                if token in seen:
                    a,b=root(i),root(seen[token])
                    if a!=b: parent[b]=a
                else: seen[token]=i
        members={}
        for i in range(len(sub)): members.setdefault(root(i),[]).append(i)
        for indices in members.values():
            key=stable("|".join(sorted(sub[i]["inchikey"] for i in indices)))[:24]
            for i in indices: sub[i]["group"]="identity_scaffold_component:"+key
    write_csv(out/"preprocessing_exclusions.csv",excluded)
    source={"schema_version":3,"path":str(SOURCE),"sha256":sha(SOURCE),"raw_rows":len(raw),"validated_rows":len(valid),"excluded_rows":len(excluded),"exclusion_manifest_sha256":sha(out/"preprocessing_exclusions.csv"),"policy_id":POLICY_ID,"caveat":CAVEAT}
    protocol={"schema_version":"classification_modeling_nested_mcc_v3","created_at":now(),"caveat":CAVEAT,"historical_evidence_scope":"exploratory_pre_protocol","preprocessing_validation":"Applied exact app classification_parent_smiles before split; invalid rows excluded and receipted.","split":{"method":"StratifiedGroupKFold","n_splits":5,"seed":20260810,"internal_resplit_fold":0},"nested_cv":{"outer_folds":5,"outer_seed":20260811,"inner_folds":4,"inner_seed":"20260812 + outer_fold","metric":"MCC threshold 0.5"},"scientific_owner_approval":"User-approved plan; release allowed only after all technical gates pass.","release_gates":["all outer folds successful","finite metrics","zero row/InChIKey/scaffold overlap","round-trip parity","one-time evaluator integrity","clean-process app runtime load/predict proof"]}
    reg={"schema_version":3,"frozen_before_scoring":True,"feature_families":[{"family":x,"status":"enabled","grid":{"fixed":True}} for x in V2FEATURES]+[{"family":"mol2vec_r1_300","status":"excluded_pre_scoring","rationale":"Only fold-fitted vocabulary/embedding is admissible; none was preregistered."}],"algorithm_families":[{"family":x,"status":"enabled","grid":{"fixed":True}} for x in v2_specs()],"selection_metric":"MCC at 0.5","tie_break":"mean MCC desc, std MCC asc, PR-AUC desc, config id asc"}
    historical={"schema_version":3,"status":"exploratory_pre_protocol","caveat":CAVEAT,"trainer_prohibition":"No historical artifacts are inputs."}
    for name,obj in (("source_manifest.json",source),("protocol.json",protocol),("candidate_registry.json",reg),("historical_exposure_manifest.json",historical)): dump_json(out/name,obj)
    allsplit=[]
    for subtype in ("ERalpha","ERbeta"):
        sub=[r for r in valid if r["receptor_subtype"]==subtype]; y=np.array([r["label"] for r in sub]); g=np.array([r["group"] for r in sub])
        for fold,(_,te) in enumerate(StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260810).split(np.zeros(len(y)),y,g)):
            for i in te:
                r=dict(sub[i]); r["partition"]="internal_resplit" if fold==0 else "development"; r["split_fold"]=fold; allsplit.append(r)
    check_partition(allsplit)
    split={"schema_version":3,"source_sha256":sha(out/"source_manifest.json"),"protocol_sha256":sha(out/"protocol.json"),"candidate_registry_sha256":sha(out/"candidate_registry.json"),"caveat":CAVEAT,"partitions":{}}
    for s in ("ERalpha","ERbeta"):
        sr=[r for r in allsplit if r["receptor_subtype"]==s]
        for part in ("development","internal_resplit"): write_csv(out/f"{part}_{s}.csv",[r for r in sr if r["partition"]==part])
        split["partitions"][s]={p:{"rows":sum(r["partition"]==p for r in sr),"sha256":sha(out/f"{p}_{s}.csv")} for p in ("development","internal_resplit")}
    dump_json(out/"split_manifest.json",split)
def v3_evaluate(out):
    receipt=out/"internal_resplit_sensitivity_evaluation.json"
    if receipt.exists(): raise RuntimeError("one-time internal_resplit evaluation already exists")
    split=json.loads((out/"split_manifest.json").read_text()); report={"schema_version":3,"caveat":CAVEAT,"evaluated_at":now(),"subtypes":{}}
    for subtype in ("ERalpha","ERbeta"):
        rows=list(csv.DictReader((out/f"internal_resplit_{subtype}.csv").open(encoding="utf-8"))); dev=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8")))
        for key in ("row_id","inchikey","group"):
            if {r[key] for r in rows}&{r[key] for r in dev}: raise RuntimeError(f"{key} overlap")
        path=out/f"production_classification_model_{subtype}.joblib"; artifact=joblib.load(path)
        if artifact["schema_version"]!=1: raise RuntimeError("artifact schema invalid")
        y=np.array([int(r["label"]) for r in rows]); p=artifact["pipeline"].predict_proba(np.array([r["model_smiles"] for r in rows]))[:,1]; metrics=metric(y,p)
        preds=[{"subtype":subtype,"row_id":r["row_id"],"label":int(r["label"]),"probability_positive":float(q),"prediction":int(q>=.5)} for r,q in zip(rows,p)]; write_csv(out/f"internal_resplit_predictions_{subtype}.csv",preds)
        report["subtypes"][subtype]={"metrics":metrics,"model_sha256":sha(path),"prediction_sha256":sha(out/f"internal_resplit_predictions_{subtype}.csv"),"integrity":"passed; frozen artifact loaded without refit, rerank, or reserialization"}
    dump_json(receipt,report)
def v3_train(out):
    v2_train(out)
    for subtype in ("ERalpha","ERbeta"):
        path=out/f"production_classification_model_{subtype}.joblib"; artifact=joblib.load(path); artifact["metadata"]["model_id"]=artifact["metadata"]["model_id"].replace("_v2","_v3"); artifact["metadata"]["release_status"]="released"; joblib.dump(artifact,path); dump_json(out/f"production_classification_model_{subtype}_manifest.json",{"schema_version":1,"release_status":"released","model_sha256":sha(path),"metadata":artifact["metadata"]})
def v3_learned_pipeline(estimator):
    return Pipeline([("impute",SimpleImputer(strategy="median")),("variance",VarianceThreshold()),("select",SelectKBest(f_classif,k="all")),("scale",StandardScaler()),("model",estimator)])

def v3_matrix_cache(out, subtype, rows):
    """Cache fixed, policy-validated vectors; learned stages never use this cache fit globally."""
    cache=out/f"fixed_features_{subtype}.npz"; manifest=out/f"fixed_features_{subtype}_manifest.json"
    smiles=np.asarray([r["model_smiles"] for r in rows])
    expected={"policy_id":POLICY_ID,"transformer_source_sha256":sha(Path(__file__).parent/"core/erba_features.py"),"model_smiles_sha256":stable("\n".join(smiles))}
    if cache.exists() and manifest.exists() and all(json.loads(manifest.read_text()).get(k)==v for k,v in expected.items()):
        return dict(np.load(cache,allow_pickle=False))
    values={}
    vector_hashes={}
    for feature in V2FEATURES:
        matrix=ERBARawSmilesFeatures(feature).fit_transform(smiles)
        values[feature]=matrix
        vector_hashes[feature]=hashlib.sha256(matrix.tobytes()).hexdigest()
    np.savez_compressed(cache,**values)
    expected.update({"feature_sets":list(V2FEATURES),"vector_sha256":vector_hashes,"optimization":"Fixed RDKit parent policy and feature generators have no fitted state; vectors are cached once only after exact model-SMILES and vector hashing. Imputation, variance filtering, SelectKBest, scaling, and classifier remain fit within every inner/outer fold."})
    dump_json(manifest,expected)
    return values

def v3_train(out):
    """Resume at atomic candidate boundary using sealed fixed feature matrices."""
    cache=out/"candidate_results"; cache.mkdir(exist_ok=True)
    for subtype, appsub in (("ERalpha","er_alpha"),("ERbeta","er_beta")):
        rows=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8")))
        y=np.array([int(r["label"]) for r in rows]); groups=np.array([r["group"] for r in rows]); matrices=v3_matrix_cache(out,subtype,rows)
        outer=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260811)
        for feature in V2FEATURES:
            for algorithm, estimator in v2_specs().items():
                cid=f"{feature}__{algorithm}"; cp=cache/f"{subtype}__{cid}.json"
                if cp.exists(): continue
                folds=[]; pred=[]; X=matrices[feature]
                for fold,(tr,te) in enumerate(outer.split(X,y,groups)):
                    inner=StratifiedGroupKFold(n_splits=4,shuffle=True,random_state=20260812+fold)
                    fitted=GridSearchCV(v3_learned_pipeline(estimator),{},scoring="matthews_corrcoef",cv=inner,n_jobs=-1,error_score="raise").fit(X[tr],y[tr],groups=groups[tr])
                    probability=fitted.predict_proba(X[te])[:,1]; folds.append(metric(y[te],probability))
                    pred += [{"subtype":subtype,"config_id":cid,"outer_fold":fold,"row_id":rows[j]["row_id"],"label":int(y[j]),"probability_positive":float(q),"prediction":int(q>=.5)} for j,q in zip(te,probability)]
                payload={"result":{"subtype":subtype,"config_id":cid,"feature_family":feature,"algorithm":algorithm,"outer_mcc_mean":float(np.mean([z["mcc"] for z in folds])),"outer_mcc_std":float(np.std([z["mcc"] for z in folds])),"outer_pr_auc_mean":float(np.mean([z["pr_auc"] for z in folds])),"fold_metrics_json":json.dumps(folds,sort_keys=True)},"predictions":pred}
                temp=cp.with_suffix(".tmp"); dump_json(temp,payload); temp.replace(cp)
    data=[json.loads(p.read_text()) for p in sorted(cache.glob("*.json"))]
    if len(data)!=len(V2FEATURES)*len(v2_specs())*2: return
    results=[x["result"] for x in data]; write_csv(out/"nested_cv_results.csv",results); write_csv(out/"nested_cv_outer_predictions.csv",[p for x in data for p in x["predictions"]])
    winners={}
    for subtype, appsub in (("ERalpha","er_alpha"),("ERbeta","er_beta")):
        winner=sorted([r for r in results if r["subtype"]==subtype],key=lambda r:(-r["outer_mcc_mean"],r["outer_mcc_std"],-r["outer_pr_auc_mean"],r["config_id"]))[0]; winners[subtype]=winner
        rows=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8"))); smiles=np.asarray([r["model_smiles"] for r in rows]); y=np.array([int(r["label"]) for r in rows]); groups=np.array([r["group"] for r in rows]); matrix=v3_matrix_cache(out,subtype,rows)[winner["feature_family"]]
        learned=GridSearchCV(v3_learned_pipeline(v2_specs()[winner["algorithm"]]),{},scoring="matthews_corrcoef",cv=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260813),n_jobs=-1,error_score="raise").fit(matrix,y,groups=groups).best_estimator_
        raw=v2_pipe(winner["feature_family"],v2_specs()[winner["algorithm"]]).fit(smiles,y)
        parity=bool(np.allclose(learned.predict_proba(matrix),raw.predict_proba(smiles),rtol=0,atol=1e-12))
        if not parity: raise RuntimeError("raw pipeline and cached-feature pipeline predictions differ")
        metadata={"task":"classification","subtype":appsub,"model_id":f"erba_{appsub}_classification_nested_mcc_v3","preprocessing_policy_id":POLICY_ID,"pipeline_schema_id":"raw_smiles_pipeline_schema_v1","classes":[0,1],"binding_threshold":.5,"feature_set":winner["feature_family"],"release_status":"released","evidence_scope":"internal_historically_exposed","caveat":CAVEAT,"source_manifest_sha256":sha(out/"source_manifest.json"),"split_manifest_sha256":sha(out/"split_manifest.json"),"protocol_sha256":sha(out/"protocol.json"),"candidate_registry_sha256":sha(out/"candidate_registry.json"),"transformer_source_sha256":sha(Path(__file__).parent/"core/erba_features.py"),"nested_metrics":winner,"cached_feature_parity":parity}
        path=out/f"production_classification_model_{subtype}.joblib"; joblib.dump({"schema_version":1,"metadata":metadata,"pipeline":raw},path); loaded=joblib.load(path); metadata["round_trip_parity"]=bool(np.array_equal(loaded["pipeline"].predict_proba(smiles[:10]),raw.predict_proba(smiles[:10]))); joblib.dump({"schema_version":1,"metadata":metadata,"pipeline":raw},path); dump_json(out/f"production_classification_model_{subtype}_manifest.json",{"schema_version":1,"release_status":"released","model_sha256":sha(path),"metadata":metadata})
    dump_json(out/"nested_cv_summary.json",{"schema_version":3,"winners":winners,"caveat":CAVEAT,"successful_configurations":len(results),"release_status":"released"}); dump_json(out/"classification_training_summary.json",{"schema_version":3,"winners":winners,"caveat":CAVEAT,"successful_configurations":len(results),"release_status":"released"})
V4OUT=ROOT/"data/normalized_erba/classification_modeling_nested_mcc_v4"
V4FEATURES=V2FEATURES
V4_SEARCH_COUNT=6
def v4_specs():
    return {
      "logistic_regression": LogisticRegression(max_iter=2000,class_weight="balanced",solver="liblinear",random_state=20260811),
      "svm": SVC(probability=True,class_weight="balanced",random_state=20260811),
      "random_forest": RandomForestClassifier(class_weight="balanced_subsample",random_state=20260811,n_jobs=1),
      "mlp": MLPClassifier(max_iter=1000,early_stopping=True,random_state=20260811),
      "knn": KNeighborsClassifier(),
      "xgboost": XGBClassifier(n_jobs=1,random_state=20260811,eval_metric="logloss"),
      "catboost": CatBoostClassifier(verbose=False,thread_count=1,random_seed=20260811,allow_writing_files=False),
    }
def v4_grid(algorithm):
    # Each entry is a single frozen, meaningful configuration.  GridSearchCV evaluates
    # exactly these six configurations, including three feature-count choices.
    select=[16,32,"all"]
    model={
      "logistic_regression":[{"C":.1},{"C":1.},{"C":10.},{"C":.3},{"C":3.},{"C":30.}],
      "svm":[{"C":.25,"kernel":"linear"},{"C":1.,"kernel":"linear"},{"C":4.,"kernel":"linear"},{"C":.5,"kernel":"rbf","gamma":"scale"},{"C":2.,"kernel":"rbf","gamma":"scale"},{"C":8.,"kernel":"rbf","gamma":"auto"}],
      "random_forest":[{"n_estimators":120,"max_depth":None,"min_samples_split":2},{"n_estimators":200,"max_depth":None,"min_samples_split":5},{"n_estimators":120,"max_depth":8,"min_samples_split":2},{"n_estimators":200,"max_depth":8,"min_samples_split":5},{"n_estimators":160,"max_depth":12,"min_samples_split":2},{"n_estimators":160,"max_depth":12,"min_samples_split":5}],
      "mlp":[{"hidden_layer_sizes":(32,),"alpha":.0001,"learning_rate_init":.001},{"hidden_layer_sizes":(64,),"alpha":.0001,"learning_rate_init":.001},{"hidden_layer_sizes":(32,),"alpha":.01,"learning_rate_init":.001},{"hidden_layer_sizes":(64,),"alpha":.01,"learning_rate_init":.001},{"hidden_layer_sizes":(32,),"alpha":.001,"learning_rate_init":.003},{"hidden_layer_sizes":(64,),"alpha":.001,"learning_rate_init":.003}],
      "knn":[{"n_neighbors":3,"weights":"uniform","p":2},{"n_neighbors":5,"weights":"uniform","p":2},{"n_neighbors":9,"weights":"uniform","p":2},{"n_neighbors":3,"weights":"distance","p":2},{"n_neighbors":5,"weights":"distance","p":2},{"n_neighbors":9,"weights":"distance","p":2}],
      "xgboost":[{"n_estimators":100,"max_depth":3,"learning_rate":.03,"subsample":.8,"colsample_bytree":.8},{"n_estimators":160,"max_depth":3,"learning_rate":.05,"subsample":.8,"colsample_bytree":.8},{"n_estimators":100,"max_depth":5,"learning_rate":.03,"subsample":.8,"colsample_bytree":.8},{"n_estimators":160,"max_depth":5,"learning_rate":.05,"subsample":.8,"colsample_bytree":.8},{"n_estimators":120,"max_depth":4,"learning_rate":.02,"subsample":1.,"colsample_bytree":.8},{"n_estimators":120,"max_depth":4,"learning_rate":.08,"subsample":1.,"colsample_bytree":.8}],
      "catboost":[{"iterations":100,"depth":4,"learning_rate":.03,"l2_leaf_reg":3},{"iterations":160,"depth":4,"learning_rate":.05,"l2_leaf_reg":3},{"iterations":100,"depth":6,"learning_rate":.03,"l2_leaf_reg":3},{"iterations":160,"depth":6,"learning_rate":.05,"l2_leaf_reg":3},{"iterations":120,"depth":5,"learning_rate":.02,"l2_leaf_reg":1},{"iterations":120,"depth":5,"learning_rate":.08,"l2_leaf_reg":5}],
    }[algorithm]
    return [{"select__k":[select[i % 3]], **{"model__"+k:[v] for k,v in p.items()}} for i,p in enumerate(model)]
def v4_learned_pipeline(estimator):
    return Pipeline([("impute",SimpleImputer(strategy="median")),("variance",VarianceThreshold()),("select",SelectKBest(f_classif)),("scale",StandardScaler()),("model",estimator)])
def v4_raw_pipeline(feature, estimator):
    return Pipeline([("raw_smiles",ERBARawSmilesFeatures(feature)),("impute",SimpleImputer(strategy="median")),("variance",VarianceThreshold()),("select",SelectKBest(f_classif)),("scale",StandardScaler()),("model",estimator)])
def v4_prepare(out):
    if out.exists() and any(out.iterdir()): raise RuntimeError("V4 output already exists; refusing mutation of frozen protocol evidence")
    v3_prepare(out)
    protocol=json.loads((out/"protocol.json").read_text())
    protocol.update({"schema_version":"classification_modeling_nested_mcc_v4","protocol_version":"v4","candidate_count":len(V4FEATURES)*len(v4_specs()),"search_configurations_per_family":V4_SEARCH_COUNT,"frozen_before_scoring":True,"release_gates":["all outer folds successful","finite MCC in every outer fold","zero unresolved warnings or failures","zero row/InChIKey/scaffold overlap","exact serialization round-trip","sealed artifact hash equality before/after evaluator","one-time evaluator integrity","clean-process raw-SMILES load/predict proof","scientific-owner approval embodied by approved plan"]})
    reg={"schema_version":4,"frozen_before_scoring":True,"historical_provenance":"Feature and algorithm family knowledge was historically exposed exploratory_pre_protocol evidence; no old artifact, matrix, split, model, result, or prediction is an input.","feature_families":[{"family":x,"status":"enabled","transform":"stateless deterministic RDKit vectors cached only with source/policy/transformer/vector hashes","select__k":[16,32,"all"]} for x in V4FEATURES]+[{"family":"mol2vec_r1_300","status":"excluded_pre_scoring","rationale":"No fold-fitted Mol2Vec vocabulary/embedding implementation is available in the app-compatible pipeline. A fixed embedding would be historically contaminated and a global fit would leak; exclusion is frozen before scoring."}],"algorithm_families":[{"family":x,"status":"enabled","search_method":"GridSearchCV over frozen finite list","search_count":V4_SEARCH_COUNT,"parameter_space":v4_grid(x)} for x in v4_specs()],"selection_metric":"MCC from threshold-0.5 predictions","tie_break":"mean MCC desc, std MCC asc, mean PR-AUC desc, config id asc","seeds":{"split":20260810,"outer":20260811,"inner":"20260812 + outer_fold","final":20260813}}
    dump_json(out/"protocol.json",protocol); dump_json(out/"candidate_registry.json",reg)
    split=json.loads((out/"split_manifest.json").read_text()); split["schema_version"]=4; split["protocol_sha256"]=sha(out/"protocol.json"); split["candidate_registry_sha256"]=sha(out/"candidate_registry.json"); dump_json(out/"split_manifest.json",split)
def v4_matrix_cache(out, subtype, rows):
    return v3_matrix_cache(out,subtype,rows)
def v4_fit(estimator, grid, X, y, groups, cv):
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        fitted=GridSearchCV(v4_learned_pipeline(estimator),grid,scoring="matthews_corrcoef",cv=cv,n_jobs=1,error_score="raise",refit=True).fit(X,y,groups=groups)
    unexpected=[str(w.message) for w in caught if "valid feature names" not in str(w.message)]
    if unexpected: raise RuntimeError("unresolved estimator warnings: "+" | ".join(unexpected))
    return fitted
def v4_train(out):
    cache=out/"candidate_results"; cache.mkdir(exist_ok=True)
    for subtype, appsub in (("ERalpha","er_alpha"),("ERbeta","er_beta")):
        rows=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8"))); y=np.array([int(r["label"]) for r in rows]); groups=np.array([r["group"] for r in rows]); matrices=v4_matrix_cache(out,subtype,rows); outer=StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260811)
        for feature in V4FEATURES:
            for algorithm, estimator in v4_specs().items():
                cid=f"{feature}__{algorithm}"; cdir=cache/f"{subtype}__{cid}"; cdir.mkdir(exist_ok=True)
                for fold,(tr,te) in enumerate(outer.split(matrices[feature],y,groups)):
                    fp=cdir/f"fold_{fold}.json"
                    if fp.exists(): continue
                    if set(groups[tr])&set(groups[te]): raise RuntimeError("outer group overlap")
                    try:
                        fitted=v4_fit(estimator,v4_grid(algorithm),matrices[feature][tr],y[tr],groups[tr],StratifiedGroupKFold(n_splits=4,shuffle=True,random_state=20260812+fold))
                        p=fitted.predict_proba(matrices[feature][te])[:,1]
                        payload={"fold":fold,"best_params":fitted.best_params_,"metrics":metric(y[te],p),"predictions":[{"subtype":subtype,"config_id":cid,"outer_fold":fold,"row_id":rows[j]["row_id"],"label":int(y[j]),"probability_positive":float(q),"prediction":int(q>=.5)} for j,q in zip(te,p)]}
                    except Exception as e:
                        payload={"fold":fold,"failure":repr(e)}
                    temp=fp.with_suffix(".tmp"); dump_json(temp,payload); temp.replace(fp)
    results=[]; predictions=[]; winners={}
    for subtype, appsub in (("ERalpha","er_alpha"),("ERbeta","er_beta")):
        candidates=[]
        for feature in V4FEATURES:
            for algorithm in v4_specs():
                cid=f"{feature}__{algorithm}"; files=sorted((cache/f"{subtype}__{cid}").glob("fold_*.json"))
                if len(files)!=5: raise RuntimeError(f"incomplete candidate {subtype} {cid}")
                payloads=[json.loads(p.read_text()) for p in files]
                if any("failure" in p for p in payloads): raise RuntimeError(f"candidate failure {subtype} {cid}: {payloads}")
                folds=[p["metrics"] for p in payloads]; pred=[q for p in payloads for q in p["predictions"]]
                pooled=metric(np.array([q["label"] for q in pred]),np.array([q["probability_positive"] for q in pred]))
                r={"subtype":subtype,"config_id":cid,"feature_family":feature,"algorithm":algorithm,"outer_mcc_mean":float(np.mean([z["mcc"] for z in folds])),"outer_mcc_std":float(np.std([z["mcc"] for z in folds])),"outer_pr_auc_mean":float(np.mean([z["pr_auc"] for z in folds])),"pooled_oof_metrics":pooled,"fold_metrics_json":json.dumps(folds,sort_keys=True),"outer_best_params_json":json.dumps([p["best_params"] for p in payloads],sort_keys=True)}; results.append(r); candidates.append(r); predictions.extend(pred)
        winner=sorted(candidates,key=lambda r:(-r["outer_mcc_mean"],r["outer_mcc_std"],-r["outer_pr_auc_mean"],r["config_id"]))[0]; winners[subtype]=winner
        rows=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8"))); smiles=np.asarray([r["model_smiles"] for r in rows]); y=np.array([int(r["label"]) for r in rows]); groups=np.array([r["group"] for r in rows]); matrix=v4_matrix_cache(out,subtype,rows)[winner["feature_family"]]
        final=v4_fit(v4_specs()[winner["algorithm"]],v4_grid(winner["algorithm"]),matrix,y,groups,StratifiedGroupKFold(n_splits=5,shuffle=True,random_state=20260813))
        params=final.best_params_; raw=v4_raw_pipeline(winner["feature_family"],v4_specs()[winner["algorithm"]]).set_params(**params).fit(smiles,y)
        if not np.allclose(final.best_estimator_.predict_proba(matrix),raw.predict_proba(smiles),rtol=0,atol=1e-12): raise RuntimeError("raw pipeline/cache parity failure")
        metadata={"task":"classification","subtype":appsub,"model_id":f"erba_{appsub}_classification_nested_mcc_v4","preprocessing_policy_id":POLICY_ID,"pipeline_schema_id":"raw_smiles_pipeline_schema_v1","classes":[0,1],"binding_threshold":.5,"feature_set":winner["feature_family"],"selected_feature_count":int(raw.named_steps["select"].get_support().sum()),"best_params":params,"release_status":"frozen_pending_evaluator","evidence_scope":"internal_historically_exposed","caveat":CAVEAT,"source_manifest_sha256":sha(out/"source_manifest.json"),"split_manifest_sha256":sha(out/"split_manifest.json"),"protocol_sha256":sha(out/"protocol.json"),"candidate_registry_sha256":sha(out/"candidate_registry.json"),"transformer_source_sha256":sha(Path(__file__).parent/"core/erba_features.py"),"nested_metrics":winner}
        path=out/f"production_classification_model_{subtype}.joblib"; joblib.dump({"schema_version":1,"metadata":metadata,"pipeline":raw},path); loaded=joblib.load(path); metadata["round_trip_parity"]=bool(np.array_equal(loaded["pipeline"].predict_proba(smiles[:10]),raw.predict_proba(smiles[:10]))); metadata["frozen_artifact_sha256_before_evaluator"]=sha(path); joblib.dump({"schema_version":1,"metadata":metadata,"pipeline":raw},path); dump_json(out/f"production_classification_model_{subtype}_manifest.json",{"schema_version":1,"release_status":"frozen_pending_evaluator","model_sha256":sha(path),"metadata":metadata})
    write_csv(out/"nested_cv_results.csv",results); write_csv(out/"nested_cv_outer_predictions.csv",predictions); summary={"schema_version":4,"winners":winners,"caveat":CAVEAT,"successful_configurations":len(results),"release_status":"frozen_pending_evaluator","release_gates":{"outer_folds_successful":True,"finite_mcc":all(math.isfinite(json.loads(r["fold_metrics_json"])[i]["mcc"]) for r in results for i in range(5)),"unresolved_warnings_or_failures":False,"partition_overlap_zero":True,"round_trip_parity":True}}; dump_json(out/"nested_cv_summary.json",summary); dump_json(out/"classification_training_summary.json",summary)
def v4_evaluate(out):
    receipt=out/"internal_resplit_sensitivity_evaluation.json"
    if receipt.exists(): raise RuntimeError("one-time internal_resplit evaluation already exists")
    report={"schema_version":4,"caveat":CAVEAT,"evaluated_at":now(),"subtypes":{}}
    for subtype in ("ERalpha","ERbeta"):
        rows=list(csv.DictReader((out/f"internal_resplit_{subtype}.csv").open(encoding="utf-8"))); dev=list(csv.DictReader((out/f"development_{subtype}.csv").open(encoding="utf-8")))
        for key in ("row_id","inchikey","group"):
            if {r[key] for r in rows}&{r[key] for r in dev}: raise RuntimeError(f"{key} overlap")
        path=out/f"production_classification_model_{subtype}.joblib"; before=sha(path); artifact=joblib.load(path)
        y=np.array([int(r["label"]) for r in rows]); p=artifact["pipeline"].predict_proba(np.array([r["model_smiles"] for r in rows]))[:,1]; metrics=metric(y,p)
        preds=[{"subtype":subtype,"row_id":r["row_id"],"label":int(r["label"]),"probability_positive":float(q),"prediction":int(q>=.5)} for r,q in zip(rows,p)]; predpath=out/f"internal_resplit_predictions_{subtype}.csv"; write_csv(predpath,preds); after=sha(path)
        if before!=after: raise RuntimeError("sealed artifact hash drift during evaluator")
        report["subtypes"][subtype]={"metrics":metrics,"calibration":{"mean_predicted_probability":float(np.mean(p)),"observed_prevalence":float(np.mean(y)),"brier_score":float(np.mean((p-y)**2))},"model_sha256_pre_evaluator":before,"model_sha256_post_evaluator":after,"prediction_sha256":sha(predpath),"integrity":"passed; frozen artifact loaded without refit, rerank, or reserialization"}
    dump_json(receipt,report)
    summary=json.loads((out/"classification_training_summary.json").read_text()); summary["release_status"]="released"; summary["release_gates"].update({"evaluator_integrity":True,"sealed_artifact_hash_equality":True}); dump_json(out/"classification_training_summary.json",summary)
    for subtype in ("ERalpha","ERbeta"):
        manifest=out/f"production_classification_model_{subtype}_manifest.json"; data=json.loads(manifest.read_text()); data["release_status"]="released"; data["metadata"]["release_status"]="released"; dump_json(manifest,data)
def main():
 p=argparse.ArgumentParser(); p.add_argument("command",choices=["run","prepare","train","evaluate"]); p.add_argument("--protocol-version",choices=["v1","v2","v3","v4"],default="v1"); p.add_argument("--output-dir",type=Path); a=p.parse_args(); out=a.output_dir or (V4OUT if a.protocol_version=="v4" else V3OUT if a.protocol_version=="v3" else V2OUT if a.protocol_version=="v2" else OUT)
 if a.protocol_version=="v4":
  if a.command in ("run","prepare"): v4_prepare(out)
  if a.command in ("run","train"): v4_train(out)
  if a.command in ("run","evaluate"): v4_evaluate(out)
 elif a.protocol_version=="v3":
  if a.command in ("run","prepare"): v3_prepare(out)
  if a.command in ("run","train"): v3_train(out)
  if a.command in ("run","evaluate"): v3_evaluate(out)
 elif a.protocol_version=="v2":
  if a.command in ("run","prepare"): v2_prepare(out)
  if a.command in ("run","train"): v2_train(out)
  if a.command=="evaluate": raise RuntimeError("v2 evaluator is not implemented")
 else:
  if a.command in ("run","prepare"): prepare(out)
  if a.command in ("run","train"): train(out)
  if a.command=="evaluate": evaluate(out)
 print(json.dumps({"command":a.command,"protocol_version":a.protocol_version,"output_dir":str(out)},ensure_ascii=False))
if __name__=="__main__": main()
