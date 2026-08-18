"""V7 raw-SMILES classification primitives using the app transformer."""
from __future__ import annotations
import csv, hashlib, json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import numpy as np
from rdkit import Chem
from rdkit.Chem.Scaffolds import MurckoScaffold
from sklearn.feature_selection import SelectKBest, VarianceThreshold, f_classif
from sklearn.impute import SimpleImputer
from sklearn.metrics import average_precision_score, balanced_accuracy_score, brier_score_loss, confusion_matrix, f1_score, matthews_corrcoef, roc_auc_score
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from core.erba_features import ERBARawSmilesFeatures
from core.erba_preprocessing import classification_parent_smiles
from core.contracts import HISTORICAL_EXPOSURE_CAVEAT_FULL

ROOT=Path(r"D:/research/FDA_endocrine_disruption")
SOURCE=ROOT/"data/normalized_erba/classification_threshold_step2/model_ready_er_binding_iqr_augmented.csv"
OUT=ROOT/"data/normalized_erba/classification_modeling_nested_mcc_v7"
POLICY_ID="erba_binding_classification_parent_v2"
PIPELINE_SCHEMA_ID="raw_smiles_pipeline_schema_v1"
CAVEAT=HISTORICAL_EXPOSURE_CAVEAT_FULL
FEATURES=("rdkit_2d","morgan_r2_2048","morgan_r3_2048","rdkit_fp_2048","maccs_167","atom_pair_2048","torsion_2048","avalon_2048")

def sha(path):
 h=hashlib.sha256()
 with Path(path).open("rb") as f:
  for b in iter(lambda:f.read(1048576),b""): h.update(b)
 return h.hexdigest()
def text_sha(value): return hashlib.sha256(value.encode("utf-8")).hexdigest()
def dump(path,value): Path(path).write_text(json.dumps(value,sort_keys=True,indent=2,ensure_ascii=False,allow_nan=False),encoding="utf-8")
def load(path): return json.loads(Path(path).read_text(encoding="utf-8"))
def write_rows(path, rows):
 if not rows: raise RuntimeError("refusing to write empty rows")
 with Path(path).open("w",newline="",encoding="utf-8") as f:
  w=csv.DictWriter(f,fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
def rows(path):
 with Path(path).open(encoding="utf-8",newline="") as f: return list(csv.DictReader(f))
def model_smiles(raw_smiles):
 prepared=classification_parent_smiles(raw_smiles)
 if not prepared.valid or not prepared.model_smiles: raise ValueError(f"invalid raw-SMILES policy input: {prepared.status_code.value}")
 return prepared.model_smiles
def group_for(raw_smiles):
 m=Chem.MolFromSmiles(model_smiles(raw_smiles),sanitize=True)
 if m is None: raise ValueError("RDKit parse failed")
 scaffold=MurckoScaffold.MurckoScaffoldSmiles(mol=m,includeChirality=True)
 return scaffold or "acyclic:"+Chem.MolToSmiles(m,canonical=True,isomericSmiles=True)
def pipeline(feature,estimator,memory=None): return Pipeline([("raw_smiles",ERBARawSmilesFeatures(feature_set=feature)),("impute",SimpleImputer(strategy="median")),("variance",VarianceThreshold()),("select",SelectKBest(f_classif)),("scale",StandardScaler()),("model",estimator)],memory=memory)
def metric(y,p):
 y=np.asarray(y); p=np.asarray(p); q=(p>=.5).astype(int); tn,fp,fn,tp=confusion_matrix(y,q,labels=[0,1]).ravel()
 return {"mcc":float(matthews_corrcoef(y,q)),"pr_auc":float(average_precision_score(y,p)),"roc_auc":float(roc_auc_score(y,p)),"balanced_accuracy":float(balanced_accuracy_score(y,q)),"f1":float(f1_score(y,q)) if (tp+fp) else 0.0,"sensitivity":float(tp/(tp+fn)),"specificity":float(tn/(tn+fp)),"tn":int(tn),"fp":int(fp),"fn":int(fn),"tp":int(tp),"brier_score":float(brier_score_loss(y,p))}
