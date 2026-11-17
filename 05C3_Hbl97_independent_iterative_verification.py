# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 15:47:43 2026

@author: hussa
"""


# ============================================================================
# STEP 05C.3 — INDEPENDENT Hbl97 ITERATIVE AND STOICHIOMETRIC VERIFICATION
#
# Publication: Gerya et al. (1997)
# Input: Step05C1_Hbl97_Independent_Chemical_Reconstruction.xlsx
#
# Objectives:
# 1. Independently solve coupled Hbl97 temperature-pressure equations.
# 2. Test convergence from multiple initial P-T estimates.
# 3. Compare converged solutions with original workbook results.
# 4. Audit chemical normalization and stoichiometric consistency.
# 5. Quantify sensitivity to initial conditions.
# 6. Generate a detailed scientific verification report.
#
# IMPORTANT:
# This script does not modify original calibration workbooks.
# ============================================================================

from pathlib import Path
from datetime import datetime
import logging
import math
import hashlib
import numpy as np
import pandas as pd

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================================
# 1. PROJECT SETTINGS
# ============================================================================

PROJECT = Path(
    r"D:\Data Science Assignment\My Projects\Geothermobarometer_ML_Project"
)

INPUT_FILE = (
    PROJECT
    / "05C1_Hbl97_independent_chemical_reconstruction"
    / "Step05C1_Hbl97_Independent_Chemical_Reconstruction.xlsx"
)

OUTPUT_DIR = (
    PROJECT
    / "05C3_Hbl97_independent_iterative_verification"
)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

REPORT = (
    OUTPUT_DIR
    / "Step05C3_Hbl97_Independent_Iterative_Verification.xlsx"
)

LOG_FILE = LOG_DIR / "step05C3_execution_log.txt"

MAX_ITERATIONS = 10000
CONVERGENCE_TOLERANCE = 1e-10

# Additional numerical stability tests.
INITIAL_GUESSES = [
    ("Original workbook input", 0.0, 0.0),
    ("Temperature +100 C", 100.0, 0.0),
    ("Temperature -100 C", -100.0, 0.0),
    ("Pressure +2 kbar", 0.0, 2.0),
    ("Pressure -2 kbar", 0.0, -2.0),
    ("Temperature +50 C and pressure +1 kbar", 50.0, 1.0),
    ("Temperature -50 C and pressure -1 kbar", -50.0, -1.0),
]


# ============================================================================
# 2. LOGGING
# ============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(
            LOG_FILE,
            mode="w",
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("Hbl97_Step05C3")


# ============================================================================
# 3. UTILITIES
# ============================================================================

def safe_float(value):

    try:
        result = float(value)

        if not math.isfinite(result):
            return None

        return result

    except (TypeError, ValueError):
        return None


def calculate_rmse(values):

    values = np.asarray(values, dtype=float)

    if len(values) == 0:
        return np.nan

    return float(np.sqrt(np.mean(values ** 2)))


def calculate_mae(values):

    values = np.asarray(values, dtype=float)

    if len(values) == 0:
        return np.nan

    return float(np.mean(np.abs(values)))


# ============================================================================
# 4. ORIGINAL Hbl97 EQUATIONS
# ============================================================================

def hbl97_temperature(P_kbar, Si, Mg, Fe2):

    """
    Original Hbl97 workbook equation from cell BO.

    T = (6119 - 28.4*P + 114*Mg/(Mg+Fe2))
        / (8.181 - 1.987*ln(8.489-Si)) - 273.15
    """

    denominator = 8.181 - 1.987 * math.log(8.489 - Si)

    if denominator == 0:
        raise ValueError("Temperature denominator is zero.")

    ratio = Mg / (Mg + Fe2)

    return (
        (6119 - 28.4 * P_kbar + 114 * ratio)
        / denominator
        - 273.15
    )


def hbl97_pressure(T_C, Al, Mg, Fe2):

    """
    Original Hbl97 workbook equation from cell BP.

    P = [2543 - 4.744*(T+273.15)
         +175*Mg/(Mg+Fe2)
         +1.987*(T+273.15)*ln(Al+1.433)] / 148.1
    """

    T_K = T_C + 273.15

    ratio = Mg / (Mg + Fe2)

    return (
        2543
        - 4.744 * T_K
        + 175 * ratio
        + 1.987 * T_K * math.log(Al + 1.433)
    ) / 148.1


# ============================================================================
# 5. COUPLED ITERATIVE SOLVER
# ============================================================================

def solve_hbl97(
    initial_T,
    initial_P,
    Si,
    Al,
    Mg,
    Fe2
):

    T = float(initial_T)
    P = float(initial_P)

    history = []

    for iteration in range(1, MAX_ITERATIONS + 1):

        T_new = hbl97_temperature(
            P, Si, Mg, Fe2
        )

        P_new = hbl97_pressure(
            T, Al, Mg, Fe2
        )

        delta_T = T_new - T
        delta_P = P_new - P

        history.append({
            "Iteration": iteration,
            "Temperature_C": T_new,
            "Pressure_kbar": P_new,
            "Delta_T_C": delta_T,
            "Delta_P_kbar": delta_P
        })

        if (
            abs(delta_T) <= CONVERGENCE_TOLERANCE
            and
            abs(delta_P) <= CONVERGENCE_TOLERANCE
        ):

            return {
                "Converged": True,
                "Iterations": iteration,
                "Temperature_C": T_new,
                "Pressure_kbar": P_new,
                "Final_Delta_T": delta_T,
                "Final_Delta_P": delta_P,
                "History": history
            }

        if not (
            math.isfinite(T_new)
            and math.isfinite(P_new)
        ):
            break

        T = T_new
        P = P_new

    return {
        "Converged": False,
        "Iterations": len(history),
        "Temperature_C": T,
        "Pressure_kbar": P,
        "Final_Delta_T": np.nan,
        "Final_Delta_P": np.nan,
        "History": history
    }


# ============================================================================
# 6. CHEMICAL STOICHIOMETRIC AUDIT
# ============================================================================

def audit_chemistry(record):

    Si = safe_float(record["Si_13"])
    Al = safe_float(record["Al_13"])
    Fe2 = safe_float(record["Fe2_13"])
    Fe3 = safe_float(record["Fe3_13"])
    Mg = safe_float(record["Mg_13"])

    values = [Si, Al, Fe2, Fe3, Mg]

    if any(v is None for v in values):
        return {
            "Status": "NOT EVALUATED",
            "Reason": "Missing chemical input"
        }

    cation_sum = sum(values)

    # These are partial checks, not a complete amphibole
    # charge-balance calculation.
    xmg = Mg / (Mg + Fe2)

    return {
        "Status": "PASS" if cation_sum > 0 else "FAIL",
        "Si_Al_Fe2_Fe3_Mg_Sum": cation_sum,
        "XMg": xmg,
        "Mg_Fe2_Ratio": Mg / Fe2 if Fe2 > 0 else np.nan,
        "Fe3_FeTotal_Ratio": (
            Fe3 / (Fe2 + Fe3)
            if (Fe2 + Fe3) > 0
            else np.nan
        ),
        "Si_Positive": Si > 0,
        "Al_Positive": Al > 0,
        "Fe2_Positive": Fe2 > 0,
        "Mg_Positive": Mg > 0
    }


# ============================================================================
# 7. LOAD STEP 05C1 REPORT
# ============================================================================

def load_input():

    if not INPUT_FILE.exists():

        raise FileNotFoundError(
            "Step 05C1 input report not found:\n"
            + str(INPUT_FILE)
        )

    logger.info("Opening Step 05C1 report.")

    wb = load_workbook(
        INPUT_FILE,
        read_only=True,
        data_only=True
    )

    if "Sample_Chemistry" not in wb.sheetnames:

        raise ValueError(
            "Sample_Chemistry worksheet is missing."
        )

    ws = wb["Sample_Chemistry"]

    rows = list(
        ws.iter_rows(values_only=True)
    )

    wb.close()

    if len(rows) < 3:

        raise ValueError(
            "No sample chemistry records found."
        )

    headers = list(rows[1])

    records = []

    for row in rows[2:]:

        if not row or not row[1]:
            continue

        records.append(
            dict(zip(headers, row))
        )

    logger.info(
        "Chemical records loaded: %s",
        len(records)
    )

    return records


# ============================================================================
# 8. VERIFY FORMULA INPUTS
# ============================================================================

def verify_required_inputs(record):

    required = [
        "Input_T_C",
        "Input_P_kbar",
        "Si_13",
        "Al_13",
        "Fe2_13",
        "Mg_13",
        "Python_T_C",
        "Python_P_kbar",
        "Excel_T_C",
        "Excel_P_kbar"
    ]

    missing = [
        name
        for name in required
        if safe_float(record.get(name)) is None
    ]

    if missing:

        raise ValueError(
            "Missing required fields: "
            + ", ".join(missing)
        )


# ============================================================================
# 9. INDEPENDENT ITERATIVE VALIDATION
# ============================================================================

def perform_validation(records):

    results = []
    sensitivity_results = []
    convergence_history = []
    chemistry_results = []

    logger.info(
        "Starting independent coupled P-T solution."
    )

    for index, record in enumerate(records, start=1):

        verify_required_inputs(record)

        sample = str(record["Sample_ID"])

        T0 = float(record["Input_T_C"])
        P0 = float(record["Input_P_kbar"])

        Si = float(record["Si_13"])
        Al = float(record["Al_13"])
        Fe2 = float(record["Fe2_13"])
        Mg = float(record["Mg_13"])

        logger.info(
            "Solving sample %s (%s/%s)",
            sample,
            index,
            len(records)
        )

        # ----------------------------------------------------
        # Main solution
        # ----------------------------------------------------

        solution = solve_hbl97(
            T0, P0,
            Si, Al, Mg, Fe2
        )

        if not solution["Converged"]:

            logger.warning(
                "Convergence failure: %s",
                sample
            )

        T_final = solution["Temperature_C"]
        P_final = solution["Pressure_kbar"]

        original_T = float(record["Excel_T_C"])
        original_P = float(record["Excel_P_kbar"])

        delta_original_T = T_final - original_T
        delta_original_P = P_final - original_P

        # ----------------------------------------------------
        # Independent equation residuals
        # ----------------------------------------------------

        T_check = hbl97_temperature(
            P_final, Si, Mg, Fe2
        )

        P_check = hbl97_pressure(
            T_final, Al, Mg, Fe2
        )

        residual_T = T_check - T_final
        residual_P = P_check - P_final

        # ----------------------------------------------------
        # Multiple initial-condition tests
        # ----------------------------------------------------

        for label, dT, dP in INITIAL_GUESSES:

            test = solve_hbl97(
                T0 + dT,
                P0 + dP,
                Si, Al, Mg, Fe2
            )

            sensitivity_results.append({

                "Sample_ID": sample,

                "Initial_Guess": label,

                "Initial_T_C": T0 + dT,

                "Initial_P_kbar": P0 + dP,

                "Converged": test["Converged"],

                "Iterations": test["Iterations"],

                "Final_T_C": test["Temperature_C"],

                "Final_P_kbar": test["Pressure_kbar"],

                "Difference_From_Main_T": (
                    test["Temperature_C"] - T_final
                    if test["Converged"] and solution["Converged"]
                    else np.nan
                ),

                "Difference_From_Main_P": (
                    test["Pressure_kbar"] - P_final
                    if test["Converged"] and solution["Converged"]
                    else np.nan
                )
            })

        # ----------------------------------------------------
        # Store main convergence history
        # ----------------------------------------------------

        for item in solution["History"]:

            convergence_history.append({

                "Sample_ID": sample,

                **item

            })

        # ----------------------------------------------------
        # Chemical checks
        # ----------------------------------------------------

        chemistry = audit_chemistry(record)

        chemistry_results.append({

            "Sample_ID": sample,

            **chemistry

        })

        # ----------------------------------------------------
        # Final status
        # ----------------------------------------------------

        status = (
            "PASS"
            if (
                solution["Converged"]
                and abs(residual_T) <= 1e-7
                and abs(residual_P) <= 1e-7
                and chemistry["Status"] == "PASS"
            )
            else "REVIEW"
        )

        results.append({

            "Sample_ID": sample,

            "Excel_Row": record["Excel_Row"],

            "Reference": record["Reference"],

            "Original_Input_T_C": T0,

            "Original_Input_P_kbar": P0,

            "Original_Excel_T_C": original_T,

            "Original_Excel_P_kbar": original_P,

            "Independent_T_C": T_final,

            "Independent_P_kbar": P_final,

            "Difference_T_vs_Excel_C": delta_original_T,

            "Difference_P_vs_Excel_kbar": delta_original_P,

            "Equation_Residual_T_C": residual_T,

            "Equation_Residual_P_kbar": residual_P,

            "Iterations": solution["Iterations"],

            "Converged": solution["Converged"],

            "Stoichiometric_Check": chemistry["Status"],

            "Overall_Status": status

        })

    return (
        results,
        sensitivity_results,
        convergence_history,
        chemistry_results
    )


# ============================================================================
# 10. CREATE SUMMARY
# ============================================================================

def create_summary(
    results,
    sensitivity,
    chemistry
):

    df = pd.DataFrame(results)
    ds = pd.DataFrame(sensitivity)

    converged = df["Converged"].sum()

    passed = (
        df["Overall_Status"] == "PASS"
    ).sum()

    T_errors = df.loc[
        df["Converged"],
        "Difference_T_vs_Excel_C"
    ].to_numpy()

    P_errors = df.loc[
        df["Converged"],
        "Difference_P_vs_Excel_kbar"
    ].to_numpy()

    sensitivity_pass = 0

    if len(ds):

        grouped = ds.groupby("Sample_ID")

        for _, group in grouped:

            valid = group[
                group["Converged"]
            ]

            if len(valid) == len(INITIAL_GUESSES):

                t_range = (
                    valid["Final_T_C"].max()
                    - valid["Final_T_C"].min()
                )

                p_range = (
                    valid["Final_P_kbar"].max()
                    - valid["Final_P_kbar"].min()
                )

                if t_range <= 1e-6 and p_range <= 1e-6:
                    sensitivity_pass += 1

    summary = [

        ("Samples evaluated", len(df)),

        ("Converged samples", int(converged)),

        ("Samples passing all numerical checks", int(passed)),

        ("Samples requiring review", int(len(df) - passed)),

        ("Initial-condition sensitivity PASS", sensitivity_pass),

        ("Total initial-condition tests",
         len(df) * len(INITIAL_GUESSES)),

        ("Temperature MAE vs Excel (C)",
         calculate_mae(T_errors)),

        ("Temperature RMSE vs Excel (C)",
         calculate_rmse(T_errors)),

        ("Temperature maximum absolute difference (C)",
         float(np.max(np.abs(T_errors))) if len(T_errors) else np.nan),

        ("Pressure MAE vs Excel (kbar)",
         calculate_mae(P_errors)),

        ("Pressure RMSE vs Excel (kbar)",
         calculate_rmse(P_errors)),

        ("Pressure maximum absolute difference (kbar)",
         float(np.max(np.abs(P_errors))) if len(P_errors) else np.nan),

        ("Maximum temperature equation residual (C)",
         float(df["Equation_Residual_T_C"].abs().max())),

        ("Maximum pressure equation residual (kbar)",
         float(df["Equation_Residual_P_kbar"].abs().max())),

        ("Convergence tolerance",
         CONVERGENCE_TOLERANCE),

        ("Maximum iterations",
         MAX_ITERATIONS),

        ("Number of initial guesses per sample",
         len(INITIAL_GUESSES))

    ]

    return pd.DataFrame(
        summary,
        columns=["Metric", "Value"]
    )


# ============================================================================
# 11. REPORT FORMATTING
# ============================================================================

def format_report(path):

    wb = load_workbook(path)

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="244062"
            )

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=True
            )

        ws.row_dimensions[1].height = 32

        for column_cells in ws.iter_cols():

            column_index = column_cells[0].column

            maximum = 0

            for cell in column_cells:

                if cell.value is not None:

                    maximum = max(
                        maximum,
                        len(str(cell.value))
                    )

            ws.column_dimensions[
                get_column_letter(column_index)
            ].width = min(
                max(maximum + 2, 12),
                42
            )

    wb.save(path)


# ============================================================================
# 12. MAIN
# ============================================================================

def main():

    logger.info("=" * 75)

    logger.info(
        "STEP 05C.3 — INDEPENDENT Hbl97 ITERATIVE VERIFICATION"
    )

    logger.info("=" * 75)

    logger.info("Input: %s", INPUT_FILE)

    logger.info("Output: %s", OUTPUT_DIR)

    records = load_input()

    (
        results,
        sensitivity,
        history,
        chemistry
    ) = perform_validation(records)

    df_results = pd.DataFrame(results)

    df_sensitivity = pd.DataFrame(sensitivity)

    df_history = pd.DataFrame(history)

    df_chemistry = pd.DataFrame(chemistry)

    df_summary = create_summary(
        results,
        sensitivity,
        chemistry
    )

    # --------------------------------------------------------
    # Formula documentation
    # --------------------------------------------------------

    formula_records = [

        {
            "Component": "Temperature equation",
            "Workbook cell": "BO",
            "Expression": (
                "(6119 - 28.4*P + 114*Mg/(Mg+Fe2)) "
                "/ (8.181 - 1.987*LN(8.489-Si)) - 273.15"
            ),
            "Source": "Original Hbl97 workbook"
        },

        {
            "Component": "Pressure equation",
            "Workbook cell": "BP",
            "Expression": (
                "(2543 - 4.744*(T+273.15) "
                "+ 175*Mg/(Mg+Fe2) "
                "+ 1.987*(T+273.15)*LN(Al+1.433)) / 148.1"
            ),
            "Source": "Original Hbl97 workbook"
        },

        {
            "Component": "Iteration method",
            "Workbook cell": "Independent Python solver",
            "Expression": (
                "Repeated substitution of T=f(P) and P=g(T) "
                "until both changes satisfy tolerance."
            ),
            "Source": "Independent numerical implementation"
        },

        {
            "Component": "Convergence tolerance",
            "Workbook cell": "Python setting",
            "Expression": str(CONVERGENCE_TOLERANCE),
            "Source": "Numerical acceptance setting"
        },

        {
            "Component": "Initial conditions",
            "Workbook cell": "Python setting",
            "Expression": (
                "Original input plus six perturbed starting conditions."
            ),
            "Source": "Numerical sensitivity experiment"
        }

    ]

    df_formulas = pd.DataFrame(formula_records)

    # --------------------------------------------------------
    # Limitations
    # --------------------------------------------------------

    limitations = [

        (
            "Equation provenance",
            "Equations were transcribed from original workbook cells BO and BP."
        ),

        (
            "Iteration",
            "Independent fixed-point iteration is used. Convergence does not "
            "prove global uniqueness of the mathematical solution."
        ),

        (
            "Chemical validation",
            "The chemical checks are partial consistency checks. They do not "
            "replace a complete amphibole charge-balance and site-allocation audit."
        ),

        (
            "Fe3+",
            "Fe3+ affects the reconstructed composition through the existing "
            "Step 05C1 results. This step does not independently determine "
            "the geologically correct oxidation state."
        ),

        (
            "Calibration validity",
            "Agreement with the original workbook does not independently "
            "validate the empirical calibration against experimental data."
        ),

        (
            "Initial guesses",
            "Seven starting conditions are tested. Stability over these "
            "conditions is not proof of stability for every possible initial value."
        ),

        (
            "Scientific interpretation",
            "The equations are evaluated using the chemical variables "
            "reconstructed in Step 05C1. Their scientific applicability "
            "must still be assessed against the original publication."
        )

    ]

    df_limitations = pd.DataFrame(
        limitations,
        columns=["Item", "Description"]
    )

    # --------------------------------------------------------
    # Save report
    # --------------------------------------------------------

    logger.info("Generating Excel report.")

    with pd.ExcelWriter(
        REPORT,
        engine="openpyxl"
    ) as writer:

        df_summary.to_excel(
            writer,
            sheet_name="Validation_Summary",
            index=False
        )

        df_results.to_excel(
            writer,
            sheet_name="Iterative_Results",
            index=False
        )

        df_sensitivity.to_excel(
            writer,
            sheet_name="Initial_Guess_Sensitivity",
            index=False
        )

        df_history.to_excel(
            writer,
            sheet_name="Convergence_History",
            index=False
        )

        df_chemistry.to_excel(
            writer,
            sheet_name="Chemical_Consistency",
            index=False
        )

        df_formulas.to_excel(
            writer,
            sheet_name="Formula_Documentation",
            index=False
        )

        df_limitations.to_excel(
            writer,
            sheet_name="Method_and_Limitations",
            index=False
        )

    format_report(REPORT)

    # --------------------------------------------------------
    # Final statistics
    # --------------------------------------------------------

    logger.info("=" * 75)

    logger.info("STEP 05C.3 COMPLETED")

    logger.info("Samples evaluated: %s", len(df_results))

    logger.info(
        "Converged samples: %s",
        int(df_results["Converged"].sum())
    )

    logger.info(
        "PASS: %s",
        int((df_results["Overall_Status"] == "PASS").sum())
    )

    logger.info(
        "REVIEW: %s",
        int((df_results["Overall_Status"] == "REVIEW").sum())
    )

    logger.info(
        "Initial condition tests: %s",
        len(df_sensitivity)
    )

    logger.info(
        "Report: %s",
        REPORT
    )

    logger.info(
        "Log: %s",
        LOG_FILE
    )

    logger.info("=" * 75)

    print("\nSTEP 05C.3 REPORT GENERATED")

    print("\nExcel report:")
    print(REPORT)

    print("\nExecution log:")
    print(LOG_FILE)

    print(
        "\nIMPORTANT: Numerical convergence is not equivalent "
        "to independent geological calibration validation."
    )


# ============================================================================
# 13. EXECUTION
# ============================================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as exc:

        logger.exception(
            "STEP 05C.3 FAILED: %s",
            exc
        )

        raise
