# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 16:17:22 2026

@author: hussa
"""


# ============================================================================
# STEP 05C.4.2
# Hbl97 Fe3+, OPTIMIZATION AND NORMALIZATION AUDIT
#
# Independent reconstruction of:
# 1. f1-f10 optimization factors
# 2. Minimum and maximum factor selection
# 3. fav
# 4. Initial-to-revised cation transformation
# 5. Fe2+/Fe3+ partitioning
# 6. Sigma13(Mn) normalization
#
# IMPORTANT:
# Arithmetic reconstruction is not equivalent to geological validation.
# The origin of Fe3+ and the scientific justification of the procedure
# remain subject to publication-based verification.
# ============================================================================

from pathlib import Path
import logging
import hashlib

import xlrd
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

SOURCE = (
    PROJECT
    / "Geothermobarometers Calibrations"
    / "Hbl P-T (1997).xls"
)

PREVIOUS_REPORT = (
    PROJECT
    / "05C4_1_Hbl97_original_chemical_reconstruction"
    / "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

OUTPUT_DIR = (
    PROJECT
    / "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit"
)

LOG_DIR = OUTPUT_DIR / "logs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

REPORT = (
    OUTPUT_DIR
    / "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

LOG_FILE = (
    LOG_DIR
    / "step05C4_2_execution_log.txt"
)


# ============================================================================
# 2. CONFIGURATION
# ============================================================================

FIRST_SAMPLE_ROW = 10
LAST_SAMPLE_ROW = 22

# Original workbook columns.

FACTOR_COLUMNS = {
    "f1": "AE",
    "f2": "AF",
    "f3": "AG",
    "f4": "AH",
    "f5": "AI",
    "Minimum_f1_f5": "AJ",
    "f6": "AK",
    "f7": "AL",
    "f8": "AM",
    "f9": "AN",
    "f10": "AO",
    "Maximum_f6_f10": "AP",
    "fav": "AQ"
}

INITIAL_COLUMNS = {
    "Si": "U",
    "Ti": "V",
    "Al": "W",
    "Fe_total": "X",
    "Mn": "Y",
    "Mg": "Z",
    "Ca": "AA",
    "Na": "AB",
    "K": "AC"
}

REVISED_COLUMNS = {
    "Si": "AT",
    "Ti": "AU",
    "Al": "AV",
    "Fe2": "AW",
    "Fe3": "AX",
    "Mn": "AY",
    "Mg": "AZ",
    "Ca": "BA",
    "Na": "BB",
    "K": "BC"
}

MN13_COLUMNS = {
    "Si": "BH",
    "Ti": "BI",
    "Al": "BJ",
    "Fe2": "BK",
    "Fe3": "BL",
    "Mn": "BM",
    "Mg": "BN"
}

NON_IRON_ELEMENTS = [
    "Si", "Ti", "Al", "Mn", "Mg", "Ca", "Na", "K"
]

MN13_ELEMENTS = [
    "Si", "Ti", "Al", "Fe2", "Fe3", "Mn", "Mg"
]

TOL = 1e-7


# ============================================================================
# 3. LOGGING
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
    ],
    force=True
)

logger = logging.getLogger("Hbl97_Step05C42")


# ============================================================================
# 4. UTILITIES
# ============================================================================

def col_index(letter):

    result = 0

    for char in letter.upper():

        result = (
            result * 26
            + ord(char) - ord("A") + 1
        )

    return result - 1


def get_cell(sheet, row, column):

    return sheet.cell_value(
        row - 1,
        col_index(column)
    )


def number(value):

    if value in ("", None):

        return np.nan

    try:

        value = float(value)

        if not np.isfinite(value):

            return np.nan

        return value

    except (ValueError, TypeError):

        return np.nan


def close_enough(a, b, tolerance=TOL):

    if not (
        np.isfinite(a)
        and np.isfinite(b)
    ):

        return False

    return abs(a - b) <= tolerance


def status(a, b):

    if not (
        np.isfinite(a)
        and np.isfinite(b)
    ):

        return "NOT EVALUATED"

    if close_enough(a, b):

        return "PASS"

    return "FAIL"


def sha256_file(path):

    digest = hashlib.sha256()

    with open(path, "rb") as f:

        for chunk in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):

            digest.update(chunk)

    return digest.hexdigest()


# ============================================================================
# 5. READ ORIGINAL SAMPLE RECORDS
# ============================================================================

def read_samples(sheet):

    samples = []

    for row in range(
        FIRST_SAMPLE_ROW,
        LAST_SAMPLE_ROW + 1
    ):

        sample = get_cell(
            sheet,
            row,
            "B"
        )

        if not str(sample).strip():

            continue

        record = {

            "Sample": str(sample).strip(),

            "Excel_Row": row,

            "Fe3_Hbl": number(
                get_cell(sheet, row, "H")
            )

        }

        for element, column in FACTOR_COLUMNS.items():

            record[element] = number(
                get_cell(sheet, row, column)
            )

        for element, column in INITIAL_COLUMNS.items():

            record["Initial_" + element] = number(
                get_cell(sheet, row, column)
            )

        for element, column in REVISED_COLUMNS.items():

            record["Revised_" + element] = number(
                get_cell(sheet, row, column)
            )

        for element, column in MN13_COLUMNS.items():

            record["MN13_" + element] = number(
                get_cell(sheet, row, column)
            )

        samples.append(record)

    return samples


# ============================================================================
# 6. RECONSTRUCT OPTIMIZATION FACTORS
# ============================================================================

def reconstruct_optimization(record):

    f_first = [
        record["f1"],
        record["f2"],
        record["f3"],
        record["f4"],
        record["f5"]
    ]

    f_second = [
        record["f6"],
        record["f7"],
        record["f8"],
        record["f9"],
        record["f10"]
    ]

    valid_first = [
        x for x in f_first
        if np.isfinite(x)
    ]

    valid_second = [
        x for x in f_second
        if np.isfinite(x)
    ]

    if (
        len(valid_first) != 5
        or len(valid_second) != 5
    ):

        minimum = np.nan
        maximum = np.nan
        fav = np.nan

    else:

        minimum = min(valid_first)

        maximum = max(valid_second)

        fav = (
            minimum + maximum
        ) / 2.0

    return {

        "Sample": record["Sample"],

        "Minimum_Reconstructed": minimum,

        "Minimum_Original": record["Minimum_f1_f5"],

        "Minimum_Check": status(
            minimum,
            record["Minimum_f1_f5"]
        ),

        "Maximum_Reconstructed": maximum,

        "Maximum_Original": record["Maximum_f6_f10"],

        "Maximum_Check": status(
            maximum,
            record["Maximum_f6_f10"]
        ),

        "fav_Reconstructed": fav,

        "fav_Original": record["fav"],

        "fav_Check": status(
            fav,
            record["fav"]
        )

    }


# ============================================================================
# 7. INITIAL TO REVISED CATION AUDIT
# ============================================================================

def reconstruct_revised(record):

    fav = record["fav"]

    results = []

    for element in NON_IRON_ELEMENTS:

        initial = record[
            "Initial_" + element
        ]

        original = record[
            "Revised_" + element
        ]

        predicted = (

            initial * fav

            if (
                np.isfinite(initial)
                and np.isfinite(fav)
            )

            else np.nan

        )

        results.append({

            "Sample": record["Sample"],

            "Element": element,

            "Initial_Cation": initial,

            "fav": fav,

            "Predicted_Revised": predicted,

            "Original_Revised": original,

            "Difference": (
                predicted - original
                if (
                    np.isfinite(predicted)
                    and np.isfinite(original)
                )
                else np.nan
            ),

            "Check": status(
                predicted,
                original
            )

        })

    return results


# ============================================================================
# 8. IRON PARTITIONING AUDIT
# ============================================================================

def reconstruct_iron(record):

    total_initial = record["Initial_Fe_total"]

    fav = record["fav"]

    total_predicted = (

        total_initial * fav

        if (
            np.isfinite(total_initial)
            and np.isfinite(fav)
        )

        else np.nan

    )

    fe3_original = record["Fe3_Hbl"]

    fe2_predicted = (

        total_predicted - fe3_original

        if (
            np.isfinite(total_predicted)
            and np.isfinite(fe3_original)
        )

        else np.nan

    )

    fe2_original = record["Revised_Fe2"]

    fe3_revised = record["Revised_Fe3"]

    iron_sum = (

        fe2_original + fe3_revised

        if (
            np.isfinite(fe2_original)
            and np.isfinite(fe3_revised)
        )

        else np.nan

    )

    fe3_difference = (

        fe3_revised - fe3_original

        if (
            np.isfinite(fe3_revised)
            and np.isfinite(fe3_original)
        )

        else np.nan

    )

    ratio = (

        fe3_original / total_predicted

        if (
            np.isfinite(fe3_original)
            and np.isfinite(total_predicted)
            and total_predicted > 0
        )

        else np.nan

    )

    return {

        "Sample": record["Sample"],

        "Initial_Fe_Total": total_initial,

        "fav": fav,

        "Predicted_Revised_Fe_Total": total_predicted,

        "Original_H_Fe3": fe3_original,

        "Predicted_Fe2": fe2_predicted,

        "Original_Revised_Fe2": fe2_original,

        "Original_Revised_Fe3": fe3_revised,

        "Fe2_Prediction_Check": status(
            fe2_predicted,
            fe2_original
        ),

        "Fe3_H_vs_AX_Check": status(
            fe3_original,
            fe3_revised
        ),

        "Fe2_plus_Fe3": iron_sum,

        "Fe_Total_Check": status(
            iron_sum,
            total_predicted
        ),

        "Implied_Fe3_FeTotal_Ratio": ratio,

        "Fe3_Origin": (
            "UNRESOLVED: workbook value preserved; "
            "independent analytical origin not established"
        )

    }


# ============================================================================
# 9. SIGMA13(MN) NORMALIZATION AUDIT
# ============================================================================

def reconstruct_mn13(record):

    revised = [
        record["Revised_" + element]
        for element in MN13_ELEMENTS
    ]

    valid = all(
        np.isfinite(x)
        for x in revised
    )

    if valid:

        revised_sum = sum(revised)

        factor = (
            13.0 / revised_sum
            if revised_sum > 0
            else np.nan
        )

    else:

        revised_sum = np.nan

        factor = np.nan

    output = {

        "Sample": record["Sample"],

        "Revised_Sigma13_Cation_Sum": revised_sum,

        "Predicted_MN13_Factor": factor

    }

    for element in MN13_ELEMENTS:

        original = record[
            "MN13_" + element
        ]

        revised_value = record[
            "Revised_" + element
        ]

        predicted = (

            revised_value * factor

            if (
                np.isfinite(revised_value)
                and np.isfinite(factor)
            )

            else np.nan

        )

        output[element + "_Predicted"] = predicted

        output[element + "_Original"] = original

        output[element + "_Check"] = status(
            predicted,
            original
        )

    predicted_sum = sum(

        output[element + "_Predicted"]

        for element in MN13_ELEMENTS

        if np.isfinite(
            output[element + "_Predicted"]
        )

    )

    output["Predicted_Sigma13"] = predicted_sum

    output["Sigma13_Check"] = status(
        predicted_sum,
        13.0
    )

    return output


# ============================================================================
# 10. INTEGRATED SAMPLE AUDIT
# ============================================================================

def integrated_audit(
    record,
    optimization,
    iron,
    mn13
):

    checks = [

        optimization["fav_Check"],

        iron["Fe2_Prediction_Check"],

        iron["Fe3_H_vs_AX_Check"],

        iron["Fe_Total_Check"],

        mn13["Sigma13_Check"]

    ]

    revised_checks = [

        x["Check"]

        for x in reconstruct_revised(record)

    ]

    checks.extend(revised_checks)

    passed = checks.count("PASS")

    failed = checks.count("FAIL")

    not_evaluated = checks.count("NOT EVALUATED")

    return {

        "Sample": record["Sample"],

        "Total_Checks": len(checks),

        "PASS": passed,

        "FAIL": failed,

        "NOT_EVALUATED": not_evaluated,

        "Arithmetic_Audit": (

            "FAIL"
            if failed > 0
            else (
                "INCOMPLETE"
                if not_evaluated > 0
                else "PASS"
            )
        )

    }


# ============================================================================
# 11. SCIENTIFIC LIMITATIONS
# ============================================================================

def limitations():

    return [

        {

            "Issue": "Fe3+ analytical origin",

            "Finding": (
                "H column and revised AX values are compared."
            ),

            "Limitation": (
                "Their numerical agreement does not establish "
                "whether Fe3+ was measured or calculated."
            ),

            "Required_Evidence": (
                "Original publication and calculation instructions."
            )

        },

        {

            "Issue": "Optimization factors",

            "Finding": (
                "Minimum, maximum and fav relationships "
                "are reconstructed from numerical records."
            ),

            "Limitation": (
                "The physical or geological meaning of f1-f10 "
                "is not established by arithmetic alone."
            ),

            "Required_Evidence": (
                "Original equations and methodological description."
            )

        },

        {

            "Issue": "Revised cations",

            "Finding": (
                "Initial non-iron cations are tested against "
                "multiplication by fav."
            ),

            "Limitation": (
                "Numerical reproduction does not establish "
                "chemical validity."
            ),

            "Required_Evidence": (
                "Published normalization methodology."
            )

        },

        {

            "Issue": "Sigma13(Mn)",

            "Finding": (
                "A common factor is tested against a sum of 13."
            ),

            "Limitation": (
                "The normalization basis and selected cation set "
                "require scientific verification."
            ),

            "Required_Evidence": (
                "Original workbook instructions and publication."
            )

        },

        {

            "Issue": "Structural formula",

            "Finding": (
                "This script does not allocate crystallographic sites."
            ),

            "Limitation": (
                "T, C, B and A site occupancies are not certified."
            ),

            "Required_Evidence": (
                "Appropriate amphibole recalculation procedure."
            )

        }

    ]


# ============================================================================
# 12. REPORT GENERATION
# ============================================================================

def save_report(data):

    with pd.ExcelWriter(
        REPORT,
        engine="openpyxl"
    ) as writer:

        for name, records in data.items():

            pd.DataFrame(records).to_excel(

                writer,

                sheet_name=name[:31],

                index=False

            )

    wb = load_workbook(REPORT)

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        if ws.max_row > 1:

            ws.auto_filter.ref = ws.dimensions

        ws.row_dimensions[1].height = 35

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
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        for column_cells in ws.iter_cols():

            index = column_cells[0].column

            width = max(

                (
                    len(str(cell.value))
                    for cell in column_cells
                    if cell.value is not None
                ),

                default=12

            )

            ws.column_dimensions[
                get_column_letter(index)
            ].width = min(
                max(width + 2, 12),
                48
            )

    wb.save(REPORT)


# ============================================================================
# 13. MAIN
# ============================================================================

def main():

    logger.info("=" * 75)

    logger.info(
        "STEP 05C.4.2 — Hbl97 Fe3+ OPTIMIZATION AUDIT"
    )

    logger.info("=" * 75)

    if not SOURCE.exists():

        raise FileNotFoundError(
            "Original XLS workbook not found: "
            + str(SOURCE)
        )

    logger.info(
        "Original workbook SHA256: %s",
        sha256_file(SOURCE)
    )

    book = xlrd.open_workbook(
        str(SOURCE),
        on_demand=True
    )

    sheet = book.sheet_by_index(0)

    samples = read_samples(sheet)

    logger.info(
        "Samples recovered: %s",
        len(samples)
    )

    if len(samples) != 13:

        logger.warning(
            "Expected 13 samples; recovered %s",
            len(samples)
        )

    optimization_records = []

    revised_records = []

    iron_records = []

    mn13_records = []

    summary_records = []

    for index, record in enumerate(samples, start=1):

        opt = reconstruct_optimization(record)

        revised = reconstruct_revised(record)

        iron = reconstruct_iron(record)

        mn13 = reconstruct_mn13(record)

        summary = integrated_audit(
            record,
            opt,
            iron,
            mn13
        )

        optimization_records.append(opt)

        revised_records.extend(revised)

        iron_records.append(iron)

        mn13_records.append(mn13)

        summary_records.append(summary)

        logger.info(
            "Processed %s/%s: %s",
            index,
            len(samples),
            record["Sample"]
        )

    book.release_resources()

    # --------------------------------------------------------
    # Report sheets
    # --------------------------------------------------------

    report_data = {

        "Optimization_Audit": optimization_records,

        "Revised_Cation_Audit": revised_records,

        "Iron_Partition_Audit": iron_records,

        "Sigma13_Normalization": mn13_records,

        "Integrated_Sample_Audit": summary_records,

        "Scientific_Limitations": limitations()

    }

    save_report(report_data)

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    summary_df = pd.DataFrame(summary_records)

    total_pass = int(
        summary_df["PASS"].sum()
    )

    total_fail = int(
        summary_df["FAIL"].sum()
    )

    total_pending = int(
        summary_df["NOT_EVALUATED"].sum()
    )

    logger.info("=" * 75)

    logger.info(
        "STEP 05C.4.2 COMPLETED"
    )

    logger.info(
        "Samples: %s",
        len(samples)
    )

    logger.info(
        "Arithmetic checks passed: %s",
        total_pass
    )

    logger.info(
        "Arithmetic checks failed: %s",
        total_fail
    )

    logger.info(
        "Checks not evaluated: %s",
        total_pending
    )

    logger.info(
        "Excel report: %s",
        REPORT
    )

    logger.info(
        "Execution log: %s",
        LOG_FILE
    )

    logger.info("=" * 75)

    print("\nSTEP 05C.4.2 COMPLETED")

    print("\nSamples:", len(samples))

    print("PASS:", total_pass)

    print("FAIL:", total_fail)

    print("NOT EVALUATED:", total_pending)

    print("\nExcel report:")
    print(REPORT)

    print("\nExecution log:")
    print(LOG_FILE)

    print(
        "\nIMPORTANT: Arithmetic agreement does not "
        "certify Fe3+ derivation, mineral formula validity "
        "or geological calibration validity."
    )


# ============================================================================
# 14. EXECUTION
# ============================================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logger.exception(
            "STEP 05C.4.2 FAILED: %s",
            error
        )

        raise
