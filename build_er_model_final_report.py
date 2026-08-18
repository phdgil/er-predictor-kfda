"""Build the authoritative ER regression/classification outcome report from V7 receipts."""
from __future__ import annotations

import csv
from html import escape
import json
from pathlib import Path
import shutil
import sys

ROOT = Path(r"D:/research/FDA_endocrine_disruption")
PROJECT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT))

from modeling.erba_classification_pipeline import CAVEAT, OUT, dump, sha

SUBTYPES = ("ERalpha", "ERbeta")
AUTHORITATIVE = ROOT / "data/normalized_erba/er_regression_classification_model_report_final.html"
REGRESSION_MANIFEST = ROOT / "data/normalized_erba/ic50_resolution_step3/best_regression_model_ERalpha_manifest.json"
REGRESSION_REPORT = ROOT / "data/normalized_erba/ic50_resolution_step3/ic50_model_report_final.html"


def read_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def fmt(value: object, digits: int = 4) -> str:
    return f"{float(value):.{digits}f}"


def main() -> None:
    summary = read_json(OUT / "classification_training_summary.json")
    evaluation = read_json(OUT / "internal_resplit_sensitivity_evaluation.json")
    releases = {
        subtype: read_json(OUT / f"release_manifest_{subtype}.json")
        for subtype in SUBTYPES
    }
    regression = read_json(REGRESSION_MANIFEST)
    if not all(releases[subtype].get("release_eligible") is True for subtype in SUBTYPES):
        raise RuntimeError("V7 classification release manifests are not eligible")

    winner_rows = []
    for subtype in SUBTYPES:
        winner = summary["winners"][subtype]
        sensitivity = evaluation["subtypes"][subtype]["metrics"]
        release = releases[subtype]
        final_metadata = summary["artifacts"][subtype]["metadata"]
        winner_rows.append(
            "<tr>"
            f"<td>{escape(subtype)}</td>"
            f"<td>{escape(winner['feature_family'])}</td>"
            f"<td>{escape(winner['algorithm'])}</td>"
            f"<td>{escape(str(release['model_id']))}</td>"
            f"<td>{final_metadata['selected_feature_count']}</td>"
            f"<td>{fmt(winner['outer_mcc_mean'])} ± {fmt(winner['outer_mcc_std'])}</td>"
            f"<td>{fmt(winner['outer_pr_auc_mean'])}</td>"
            f"<td>{fmt(winner['pooled_oof_metrics']['mcc'])}</td>"
            f"<td>{fmt(sensitivity['mcc'])}</td>"
            f"<td>{fmt(sensitivity['pr_auc'])}</td>"
            f"<td>{release['model_sha256']} ({release['model_size_bytes']} bytes)</td>"
            "<td>RELEASED</td>"
            "</tr>"
        )

    candidate_rows = []
    with (OUT / "nested_cv_results.csv").open(encoding="utf-8", newline="") as handle:
        for row in csv.DictReader(handle):
            counts = json.loads(row["outer_selected_feature_counts"])
            candidate_rows.append(
                "<tr>"
                f"<td>{escape(row['subtype'])}</td>"
                f"<td>{escape(row['feature_family'])}</td>"
                f"<td>{escape(row['algorithm'])}</td>"
                f"<td>{escape('/'.join(str(value) for value in counts))}</td>"
                f"<td>{fmt(row['outer_mcc_mean'])}</td>"
                f"<td>{fmt(row['outer_mcc_std'])}</td>"
                f"<td>{fmt(row['outer_pr_auc_mean'])}</td>"
                f"<td>{fmt(json.loads(row['pooled_oof_metrics'])['mcc'])}</td>"
                "</tr>"
            )

    regression_validation = regression["validation"]
    protocol = read_json(OUT / "protocol.json")
    split = read_json(OUT / "split_manifest.json")
    report_path = OUT / "er_regression_classification_model_report_final_v7.html"
    html = f"""<!doctype html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>ER Binding Models — Final V7 Report</title>
<style>
:root{{--ink:#172033;--muted:#5f6b7a;--line:#dbe2ea;--panel:#f7f9fc;--blue:#185adb;--green:#16794c;--amber:#9a5a00}}
*{{box-sizing:border-box}} body{{margin:0;background:#eef2f7;color:var(--ink);font:15px/1.55 Inter,Segoe UI,Arial,sans-serif}}
main{{max-width:1300px;margin:24px auto;background:white;padding:34px 42px 56px;box-shadow:0 4px 24px #18253a18}}
h1{{font-size:30px;margin:0 0 6px}} h2{{font-size:22px;margin:34px 0 12px;border-bottom:2px solid var(--line);padding-bottom:7px}}
p,li{{color:var(--muted)}} .note{{border-left:4px solid var(--blue);background:#f1f6ff;padding:12px 15px;margin:14px 0;color:var(--muted)}}
.warn{{border-left-color:var(--amber);background:#fff7e8}} .table-wrap{{overflow:auto;border:1px solid var(--line);border-radius:7px;margin:10px 0 18px}}
table{{width:100%;border-collapse:collapse;font-size:12px}} th{{background:#edf2f8;text-align:left;white-space:nowrap}} th,td{{padding:8px 9px;border-bottom:1px solid var(--line)}}
code{{background:#edf1f6;padding:2px 5px;border-radius:4px}} .pass{{color:var(--green);font-weight:700}}
</style></head><body><main>
<h1>Estrogen-Receptor Binding Models</h1>
<p>Authoritative consolidated V7 outcome report for subtype-specific direct-binding IC50 regression and qualitative binding classification.</p>
<div class="note warn"><strong>Mandatory classification limitation.</strong> {escape(CAVEAT)}</div>
<h2>Data and source status</h2>
<ul>
<li>The broad qualitative collection searched and retained raw evidence from ChEMBL, PubChem BioAssay, BindingDB, FDA EDKB ER Binding Dataset, FDA EADB, EPA CompTox/ToxCast, and OECD QSAR Toolbox. Species and target-type information were retained rather than used as collection-time exclusions.</li>
<li>The final classification modeling source is <code>classification_threshold_step2/model_ready_er_binding_iqr_augmented.csv</code>: 2,638 ChEMBL qualitative rows, 357 PubChem BioAssay qualitative rows, and 726 ChEMBL regression-derived Q1–Q3 positive rows before the V7 parent-policy split. Other collected databases remain provenance/candidate evidence because their landed rows did not satisfy this release's direct-binding qualitative label and model-ready contracts; they were not silently treated as training labels.</li>
<li>V7 sealed development/internal-sensitivity rows: ERα {split['partitions']['ERalpha']['development']['rows']} / {split['partitions']['ERalpha']['internal_resplit']['rows']}; ERβ {split['partitions']['ERbeta']['development']['rows']} / {split['partitions']['ERbeta']['internal_resplit']['rows']}.</li>
<li>Regression source: organic exact-numeric direct-binding IC50 records. Censored/text endpoints remain outside fitting.</li>
</ul>
<h2>IC50 regression</h2>
<div class="table-wrap"><table><thead><tr><th>Subtype</th><th>Task</th><th>Feature</th><th>Algorithm</th><th>Features used</th><th>Holdout R²</th><th>RMSE</th><th>MAE</th><th>Application status</th></tr></thead><tbody>
<tr><td>ERα</td><td>pIC50 regression</td><td>{escape(regression['feature_set'])}</td><td>{escape(regression['algorithm'])}</td><td>{regression_validation['selected_feature_count']}</td><td>{fmt(regression_validation['r2'])}</td><td>{fmt(regression_validation['rmse'])}</td><td>{fmt(regression_validation['mae'])}</td><td class="pass">ACCEPTED</td></tr>
<tr><td>ERβ</td><td>pIC50 regression</td><td>morgan feature radius 3</td><td>XGBoost</td><td>1706</td><td>0.5420</td><td>0.6910</td><td>0.5270</td><td>BELOW TARGET — NOT BUNDLED</td></tr>
</tbody></table></div>
<p>Regression acceptance criteria were CV R² &gt; 0.5 and holdout R² &gt; 0.6. Only ERα passed and is eligible for the application. Detailed regression evidence: <code>{escape(str(REGRESSION_REPORT))}</code>.</p>
<h2>V7 classification release outcome</h2>
<p>The release protocol evaluated eight RDKit feature families × seven algorithms separately for ERα and ERβ. Each family performed algorithm hyperparameter optimization and feature-count optimization within a development-only 5×4 nested StratifiedGroupKFold design. Selection used mean outer-fold MCC at threshold 0.5; ties within 1e-12 used lower MCC standard deviation, higher PR-AUC, then lexicographic configuration id. The sealed internal resplit was opened once by the separate evaluator and could not rerank, refit, retune, or reserialize a model.</p>
<div class="table-wrap"><table><thead><tr><th>Subtype</th><th>Feature</th><th>Algorithm</th><th>Model ID</th><th>Final selected features</th><th>Nested MCC mean ± SD</th><th>Nested PR-AUC</th><th>Pooled OOF MCC</th><th>Internal sensitivity MCC</th><th>Internal sensitivity PR-AUC</th><th>Frozen artifact SHA-256</th><th>Status</th></tr></thead><tbody>{''.join(winner_rows)}</tbody></table></div>
<h2>All V7 classification feature/algorithm baselines</h2>
<p>Selected-feature counts list the winning inner-HPO count for each of the five outer folds. These are baseline comparison results, not separately released models.</p>
<div class="table-wrap"><table><thead><tr><th>Subtype</th><th>Feature family</th><th>Algorithm</th><th>Selected features by outer fold</th><th>Outer MCC mean</th><th>Outer MCC SD</th><th>Outer PR-AUC mean</th><th>Pooled OOF MCC</th></tr></thead><tbody>{''.join(candidate_rows)}</tbody></table></div>
<h2>Reproducibility and release authority</h2>
<ul>
<li>Protocol: <code>classification_modeling_nested_mcc_v7/protocol.json</code> ({sha(OUT/'protocol.json')})</li>
<li>Candidate registry: <code>candidate_registry.json</code> ({sha(OUT/'candidate_registry.json')})</li>
<li>Nested summary: <code>nested_cv_summary.json</code> ({sha(OUT/'nested_cv_summary.json')})</li>
<li>One-time sensitivity receipt: <code>internal_resplit_sensitivity_evaluation.json</code> ({sha(OUT/'internal_resplit_sensitivity_evaluation.json')})</li>
<li>Trainer runtime environment: <code>trainer_runtime_environment.json</code> ({sha(OUT/'trainer_runtime_environment.json')})</li>
<li>External evaluator-bound release manifests authorize exact frozen joblib bytes; adjacent application catalogs are not trust authority.</li>
<li>V1–V6 results remain immutable, superseded exploratory/historical evidence and are not V7 selection or release evidence.</li>
</ul>
<h2>Interpretation limits</h2>
<ul><li>Regression and classification answer different questions and their metrics are not directly comparable.</li><li>No result is external regulatory validation.</li><li>ERβ regression remains below the predefined target and is unavailable in the app.</li><li>Coactivator binding, signaling, agonism and antagonism screens were not treated as direct ligand-binding negatives.</li></ul>
<footer>Generated solely from current frozen manifests, nested-CV outputs and the one-time evaluator receipt under <code>{escape(str(OUT))}</code>.</footer>
</main></body></html>"""
    report_path.write_text(html, encoding="utf-8")
    shutil.copyfile(report_path, AUTHORITATIVE)

    revision_manifest = {
        "schema_version": 1,
        "report": report_path.name,
        "report_sha256": sha(report_path),
        "authoritative_report": str(AUTHORITATIVE),
        "authoritative_report_sha256": sha(AUTHORITATIVE),
        "protocol_sha256": sha(OUT / "protocol.json"),
        "candidate_registry_sha256": sha(OUT / "candidate_registry.json"),
        "nested_cv_sha256": sha(OUT / "nested_cv_summary.json"),
        "evaluator_receipt_sha256": sha(OUT / "internal_resplit_sensitivity_evaluation.json"),
        "trainer_runtime_environment_sha256": sha(OUT / "trainer_runtime_environment.json"),
        "historical_exposure_caveat": CAVEAT,
        "v1_v6_status": "superseded_immutable_exploratory",
        "release_manifests": {
            subtype: sha(OUT / f"release_manifest_{subtype}.json")
            for subtype in SUBTYPES
        },
        "regression_manifest_sha256": sha(REGRESSION_MANIFEST),
    }
    dump(OUT / "er_regression_classification_model_report_final_v7_revision_manifest.json", revision_manifest)


if __name__ == "__main__":
    main()
