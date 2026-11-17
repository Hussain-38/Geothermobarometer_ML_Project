"""
STEP 08.5.1–08.8 — Integrated validation, method comparison, XGBoost diagnostics and SHAP sensitivity.

Place the two authorized input workbooks in ./data/ beside this repository:
  Step08_5_XGBoost_Experimental_Analysis.xlsx
  Step08_2_ML_Modelling_Dataset_and_Validation_Readiness.xlsx

Run in Spyder:
  %runfile '.../Step08_5_to_08_8_Integrated_Analysis.py'

Important:
- ML targets are experimental errors, not direct method accuracy.
- S1 has only seven independent publications and is exploratory.
- Method/output mappings remain provisional and require formula-level verification.
- SHAP models are refitted fold-wise using stored best parameters and grouped splits.
  This is a sensitivity refit, not exact reproduction of the original Step08.5 pipeline.
"""
from pathlib import Path
import json
import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

try:
    from xgboost import XGBRegressor
    import shap
except ImportError as exc:
    raise SystemExit(
        "Missing dependency. In a system terminal run: "
        "python -m pip install xgboost shap pandas numpy scikit-learn matplotlib openpyxl"
    ) from exc

BASE = Path(__file__).resolve().parent
DATA = BASE / "data"
OUT = BASE / "results" / "Step08_5_to_08_8"
OUT.mkdir(parents=True, exist_ok=True)

XGB_FILE = DATA / "Step08_5_XGBoost_Experimental_Analysis.xlsx"
DATA_FILE = DATA / "Step08_2_ML_Modelling_Dataset_and_Validation_Readiness.xlsx"

for path in (XGB_FILE, DATA_FILE):
    if not path.exists():
        raise FileNotFoundError(f"Required input not found: {path}")

TASKS = {
    "S1_Temperature": ("S1", "Temperature"),
    "S2_Temperature": ("S2", "Temperature"),
    "S2_Pressure": ("S2", "Pressure"),
}

def load_sheet(path, sheet):
    return pd.read_excel(path, sheet_name=sheet)

def main():
    oof = load_sheet(XGB_FILE, "OOF_Predictions")
    fold_metrics = load_sheet(XGB_FILE, "Outer_Fold_Metrics")
    model_comparison = load_sheet(XGB_FILE, "Model_Comparison")
    gain_importance = load_sheet(XGB_FILE, "XGB_Importance_Summary")

    # A. Publication-grouped OOF audit
    validation_rows = []
    task_data = {}
    for sheet, (dataset, prop) in TASKS.items():
        d = load_sheet(DATA_FILE, sheet)
        task_data[sheet] = d
        q = oof[(oof.Dataset == dataset) & (oof.Property == prop)]
        for (target, feature_set, model), g in q.groupby(
            ["Target", "Feature_Set", "Model"]
        ):
            unique_ids = g.Analysis_Record_ID.nunique()
            max_pub_folds = g.groupby("Publication_Group").Outer_Fold.nunique().max()
            validation_rows.append({
                "Task": sheet, "Target": target, "Feature_Set": feature_set,
                "Model": model, "Expected_records": len(d), "OOF_rows": len(g),
                "Unique_analysis_IDs": unique_ids,
                "Duplicate_rows": len(g) - unique_ids,
                "Max_folds_per_publication": max_pub_folds,
                "Coverage_fraction": unique_ids / len(d),
                "Status": "PASS" if len(g) == len(d) and unique_ids == len(d)
                          and max_pub_folds == 1 else "REVIEW"
            })
    validation = pd.DataFrame(validation_rows)

    # B. Pooled out-of-fold performance of combined XGBoost error models
    oof_rows = []
    q = oof[(oof.Feature_Set == "Combined") & (oof.Model == "XGBoost")]
    for (dataset, prop, target), g in q.groupby(["Dataset", "Property", "Target"]):
        observed = g.Observed.to_numpy()
        predicted = g.Predicted.to_numpy()
        residual = observed - predicted
        total_ss = np.sum((observed - observed.mean()) ** 2)
        oof_rows.append({
            "Dataset": dataset, "Property": prop, "Target": target, "N": len(g),
            "OOF_MAE": np.mean(np.abs(residual)),
            "OOF_RMSE": np.sqrt(np.mean(residual ** 2)),
            "OOF_R2": 1 - np.sum(residual ** 2) / total_ss if total_ss else np.nan,
            "Mean_residual": residual.mean(),
            "Meaning": "Predictive accuracy for error target; not direct method accuracy."
        })
    oof_summary = pd.DataFrame(oof_rows)

    # C. Direct method comparison on matched observations
    method_rows, matched_rows = [], []
    for sheet, (dataset, prop) in TASKS.items():
        d = task_data[sheet].copy()
        slot_sets = {
            slot: set(g.Observation_ID.dropna())
            for slot, g in d.groupby("Output_Slot_ID")
        }
        common = set.intersection(*slot_sets.values())
        matched = d[d.Observation_ID.isin(common)].copy()
        matched["Abs_Error"] = matched.Experimental_Error_Target.abs()
        matched["Squared_Error"] = matched.Experimental_Error_Target ** 2
        matched_rows.append({
            "Task": sheet, "Matched_observations": len(common),
            "Publications": matched.CV_Group_ID.nunique(),
            "Matched_records": len(matched)
        })
        for slot, g in matched.groupby("Output_Slot_ID"):
            by_pub = g.groupby("CV_Group_ID").agg(
                pub_mae=("Abs_Error", "mean"),
                pub_bias=("Experimental_Error_Target", "mean")
            )
            method_rows.append({
                "Task": sheet, "Property": prop, "Output_Slot_ID": slot,
                "Method_Label": g.Method_Label.iloc[0],
                "Method_Candidate_ID": g.Method_Candidate_ID.iloc[0],
                "Mapping_Status": g.Mapping_Status.iloc[0],
                "N_matched": len(g), "N_publications": g.CV_Group_ID.nunique(),
                "MAE": g.Abs_Error.mean(),
                "RMSE": np.sqrt(g.Squared_Error.mean()),
                "Signed_Bias": g.Experimental_Error_Target.mean(),
                "Median_Absolute_Error": g.Abs_Error.median(),
                "Publication_Balanced_MAE": by_pub.pub_mae.mean(),
                "Publication_Balanced_Bias": by_pub.pub_bias.mean()
            })
    method_performance = pd.DataFrame(method_rows)
    matched_audit = pd.DataFrame(matched_rows)

    # D. Fold-wise SHAP sensitivity refit.
    # Use Output_Slot_ID only; Method_Candidate_ID is nested within output slot.
    shap_rows, shap_fold_rows = [], []
    for sheet, (dataset, prop) in TASKS.items():
        d = task_data[sheet]
        numeric_cols = [c for c in d.columns if c.startswith(f"X_{dataset}_")]
        X = pd.concat([
            d[numeric_cols].apply(pd.to_numeric, errors="coerce"),
            pd.get_dummies(d["Output_Slot_ID"].fillna("MISSING").astype(str),
                           prefix="Output_Slot", dtype=float)
        ], axis=1)
        y = d["Experimental_Error_Target"].astype(float).to_numpy()
        saved = oof[
            (oof.Dataset == dataset) & (oof.Property == prop) &
            (oof.Target == "Signed_Error") & (oof.Feature_Set == "Combined") &
            (oof.Model == "XGBoost")
        ]
        params_df = fold_metrics[
            (fold_metrics.Dataset == dataset) & (fold_metrics.Property == prop) &
            (fold_metrics.Target == "Signed_Error") &
            (fold_metrics.Feature_Set == "Combined") &
            (fold_metrics.Model == "XGBoost")
        ]
        for fold in sorted(saved.Outer_Fold.unique()):
            held_publications = saved.loc[
                saved.Outer_Fold == fold, "Publication_Group"
            ].unique()
            test = d.CV_Group_ID.isin(held_publications).to_numpy()
            params = json.loads(
                params_df.loc[params_df.Outer_Fold == fold, "Best_Parameters"].iloc[0]
            )
            model = XGBRegressor(
                **params, objective="reg:squarederror",
                random_state=42, n_jobs=1
            )
            model.fit(X.loc[~test], y[~test])
            prediction = model.predict(X.loc[test])
            shap_values = np.asarray(
                shap.TreeExplainer(model).shap_values(X.loc[test])
            )
            if shap_values.ndim == 3:
                shap_values = shap_values[:, :, 0]
            for j, feature in enumerate(X.columns):
                shap_rows.append({
                    "Task": sheet, "Fold": fold, "Feature": feature,
                    "Mean_Absolute_SHAP": np.mean(np.abs(shap_values[:, j])),
                    "Mean_SHAP": np.mean(shap_values[:, j]),
                    "N_test": int(test.sum())
                })
            error = y[test] - prediction
            ss = np.sum((y[test] - y[test].mean()) ** 2)
            shap_fold_rows.append({
                "Task": sheet, "Fold": fold, "N_test": int(test.sum()),
                "MAE": np.mean(np.abs(error)),
                "RMSE": np.sqrt(np.mean(error ** 2)),
                "R2": 1 - np.sum(error ** 2) / ss if ss else np.nan
            })

    shap_detail = pd.DataFrame(shap_rows)
    shap_summary_rows = []
    for task, g in shap_detail.groupby("Task"):
        vals = g.groupby("Feature").apply(
            lambda z: np.average(z.Mean_Absolute_SHAP, weights=z.N_test),
            include_groups=False
        ).sort_values(ascending=False)
        for rank, (feature, value) in enumerate(vals.items(), start=1):
            shap_summary_rows.append({
                "Task": task, "Rank": rank, "Feature": feature,
                "Weighted_Mean_Absolute_SHAP": value
            })
    shap_summary = pd.DataFrame(shap_summary_rows)
    shap_fold_metrics = pd.DataFrame(shap_fold_rows)

    # E. Figures
    fig, axes = plt.subplots(1, 3, figsize=(18, 7))
    for ax, task in zip(axes, TASKS):
        sub = method_performance[
            method_performance.Task == task
        ].sort_values("Publication_Balanced_MAE")
        labels = sub.Output_Slot_ID.str.replace(r"^S[12]:", "", regex=True)
        ax.barh(labels, sub.Publication_Balanced_MAE)
        ax.invert_yaxis()
        ax.set_title(task.replace("_", " "))
        ax.set_xlabel("Publication-balanced MAE")
        ax.grid(axis="x", alpha=.25)
    fig.suptitle("Matched experimental performance by output variant")
    fig.tight_layout()
    fig.savefig(OUT / "Figure_1_Matched_Method_Performance.png", dpi=300, bbox_inches="tight")
    plt.close(fig)

    fig, axes = plt.subplots(1, 3, figsize=(18, 7))
    for ax, task in zip(axes, TASKS):
        sub = shap_summary[shap_summary.Task == task].head(12).sort_values(
            "Weighted_Mean_Absolute_SHAP"
        )
        ax.barh(sub.Feature, sub.Weighted_Mean_Absolute_SHAP)
        ax.set_title(task.replace("_", " "))
        ax.set_xlabel("Weighted mean |SHAP|")
        ax.grid(axis="x", alpha=.25)
    fig.suptitle("Fold-wise SHAP sensitivity analysis")
    fig.tight_layout()
    fig.savefig(OUT / "Figure_2_SHAP_Sensitivity.png", dpi=300, bbox_inches="tight")
    plt.close(fig)

    # F. Workbook
    output = OUT / "Step08_5_to_08_8_Integrated_Analysis.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        validation.to_excel(writer, sheet_name="OOF_Validation", index=False)
        oof_summary.to_excel(writer, sheet_name="XGB_OOF_Summary", index=False)
        model_comparison.to_excel(writer, sheet_name="Model_Comparison", index=False)
        method_performance.to_excel(writer, sheet_name="Matched_Method_Performance", index=False)
        matched_audit.to_excel(writer, sheet_name="Matched_Sample_Audit", index=False)
        shap_summary.to_excel(writer, sheet_name="SHAP_Sensitivity_Summary", index=False)
        shap_fold_metrics.to_excel(writer, sheet_name="SHAP_Fold_Metrics", index=False)
        gain_importance.to_excel(writer, sheet_name="Original_XGB_Gain_Importance", index=False)

    print(f"Completed. Results saved to: {OUT}")
    print("\nMatched method performance:")
    print(method_performance.sort_values(
        ["Task", "Publication_Balanced_MAE"]
    )[["Task", "Output_Slot_ID", "N_matched", "N_publications",
       "Publication_Balanced_MAE", "Publication_Balanced_Bias"]].to_string(index=False))

if __name__ == "__main__":
    main()
