# -*- coding: utf-8 -*-
"""
Created on Tue Nov 17 11:41:16 2026

@author: hussa
"""


# -*- coding: utf-8 -*-
"""
===============================================================================
STEP 06.3.2
Hbl97 — EXACT ORIGINAL OPTIMIZATION RECONSTRUCTION
===============================================================================

Purpose:
1. Independently reconstruct original Hbl97 factors f1-f10.
2. Recalculate minimum and maximum constraints.
3. Recalculate fav.
4. Independently reconstruct Fe3+ and Fe2+.
5. Reconstruct all ten revised cation values.
6. Compare calculations against the original XLS workbook.
7. Cross-check initial cations against Step 05C4.1.
8. Generate sample-level and overall verification reports.

Scientific basis:
Numerical equations recovered from the original Hbl P-T (1997).xls
workbook using its cached numerical outputs.

IMPORTANT:
This is an exact numerical reconstruction audit.
It does not independently certify Fe3+ analytical provenance,
amphibole site occupancy, or geological calibration validity.

The original workbook is never modified.
===============================================================================
"""

from pathlib import Path
from datetime import datetime
import hashlib
import logging
import math

import xlrd
import openpyxl

from openpyxl.styles import (
    Font,
    PatternFill,
    Border,
    Side,
    Alignment
)

from openpyxl.utils import get_column_letter


# =============================================================================
# A. PROJECT PATHS
# =============================================================================

ROOT = Path(
    r"D:\Data Science Assignment\My Projects\Geothermobarometer_ML_Project"
)

ORIGINAL_XLS = (
    ROOT /
    "Geothermobarometers Calibrations" /
    "Hbl P-T (1997).xls"
)

CHEM_FILE = (
    ROOT /
    "05C4_1_Hbl97_original_chemical_reconstruction" /
    "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

PREVIOUS_OPT_FILE = (
    ROOT /
    "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit" /
    "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

OUTPUT_DIR = (
    ROOT /
    "06_3_2_Hbl97_Exact_Original_Optimization_Reconstruction"
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OUTPUT_FILE = (
    OUTPUT_DIR /
    "Step06_3_2_Hbl97_Exact_Original_Optimization_Reconstruction.xlsx"
)

LOG_FILE = (
    LOG_DIR /
    "step06_3_2_execution_log.txt"
)


# =============================================================================
# B. SETTINGS
# =============================================================================

EXPECTED_SAMPLES = 13

# Strict numerical comparison tolerance.
# Excel cached values may contain minor floating-point differences.

ABS_TOL = 1e-8
REL_TOL = 1e-8

ELEMENTS = [
    "Si",
    "Ti",
    "Al",
    "Fe_total",
    "Mn",
    "Mg",
    "Ca",
    "Na",
    "K"
]

REVISED_ELEMENTS = [
    "Si",
    "Ti",
    "Al",
    "Fe2",
    "Fe3",
    "Mn",
    "Mg",
    "Ca",
    "Na",
    "K"
]

# Original workbook column positions, zero-based.

ORIGINAL_COLUMNS = {

    "Sample": 1,

    "Fe3_H": 7,

    "Si": 20,
    "Ti": 21,
    "Al": 22,
    "Fe_total": 23,
    "Mn": 24,
    "Mg": 25,
    "Ca": 26,
    "Na": 27,
    "K": 28,

    "Initial_Sum": 29,

    "f1": 30,
    "f2": 31,
    "f3": 32,
    "f4": 33,
    "f5": 34,

    "Minimum": 35,

    "f6": 36,
    "f7": 37,
    "f8": 38,
    "f9": 39,
    "f10": 40,

    "Maximum": 41,
    "fav": 42,

    "Fe3_AR": 43,

    "Revised_Si": 45,
    "Revised_Ti": 46,
    "Revised_Al": 47,
    "Revised_Fe2": 48,
    "Revised_Fe3": 49,
    "Revised_Mn": 50,
    "Revised_Mg": 51,
    "Revised_Ca": 52,
    "Revised_Na": 53,
    "Revised_K": 54,

    "Revised_Sum": 55

}


# =============================================================================
# C. LOGGING
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(
            LOG_FILE,
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ],
    force=True
)

logger = logging.getLogger("STEP06_3_2")


# =============================================================================
# D. GENERAL FUNCTIONS
# =============================================================================

def sha256_file(path):

    digest = hashlib.sha256()

    with open(path, "rb") as file:

        for block in iter(
            lambda: file.read(1024 * 1024),
            b""
        ):

            digest.update(block)

    return digest.hexdigest()


def number(value):

    if value is None:
        return None

    try:

        result = float(value)

        if math.isfinite(result):
            return result

    except (TypeError, ValueError):

        pass

    return None


def numerical_check(
    calculated,
    original,
    abs_tol=ABS_TOL,
    rel_tol=REL_TOL
):

    if calculated is None or original is None:
        return False

    return math.isclose(
        calculated,
        original,
        rel_tol=rel_tol,
        abs_tol=abs_tol
    )


def difference(calculated, original):

    if calculated is None or original is None:
        return None

    return calculated - original


def status(calculated, original):

    return (
        "PASS"
        if numerical_check(calculated, original)
        else "REVIEW"
    )


def read_xlsx_sheet(path, sheet_name):

    if not path.exists():

        raise FileNotFoundError(
            f"Required file missing:\n{path}"
        )

    wb = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=True
    )

    if sheet_name not in wb.sheetnames:

        wb.close()

        raise ValueError(
            f"Worksheet not found: {sheet_name}"
        )

    ws = wb[sheet_name]

    iterator = ws.iter_rows(
        values_only=True
    )

    headers = next(iterator)

    records = []

    for row in iterator:

        records.append(
            dict(zip(headers, row))
        )

    wb.close()

    return records


# =============================================================================
# E. READ ORIGINAL HBL97 WORKBOOK
# =============================================================================

def read_original_workbook():

    logger.info(
        "Opening original Hbl97 XLS workbook."
    )

    book = xlrd.open_workbook(
        str(ORIGINAL_XLS),
        on_demand=True
    )

    ws = book.sheet_by_name(
        "Hbl P-T (1997)"
    )

    records = {}

    # Original numerical sample rows: Excel 10–22.

    for row_index in range(9, 22):

        sample_value = ws.cell_value(
            row_index,
            ORIGINAL_COLUMNS["Sample"]
        )

        if not sample_value:
            continue

        sample = str(sample_value).strip()

        record = {

            "Sample": sample,

            "Excel_Row": row_index + 1

        }

        for name, column in ORIGINAL_COLUMNS.items():

            if name in [
                "Sample",
                "Excel_Row"
            ]:
                continue

            record[name] = number(
                ws.cell_value(
                    row_index,
                    column
                )
            )

        records[sample] = record

    book.release_resources()

    return records


# =============================================================================
# F. READ INDEPENDENT CHEMICAL RECONSTRUCTION
# =============================================================================

def read_initial_cations():

    logger.info(
        "Reading Step 05C4.1 initial cation reconstruction."
    )

    records = read_xlsx_sheet(
        CHEM_FILE,
        "Original_Calculation_Blocks"
    )

    output = {}

    for record in records:

        if (
            record.get("Calculation_Block")
            != "Initial_Cations"
        ):
            continue

        sample = str(
            record["Sample"]
        ).strip()

        output[sample] = {

            element: number(
                record.get(element)
            )

            for element in ELEMENTS

        }

    return output


# =============================================================================
# G. EXACT ORIGINAL HBL97 OPTIMIZATION EQUATIONS
# =============================================================================

def calculate_hbl97_factors(c):

    """
    Reconstruct original Hbl97 optimization equations.

    Equations recovered from the original workbook.

    Z = sum of initial cations.

    f1 = 16/Z
    f2 = 8/Si
    f3 = 15/(Z-Na-K)
    f4 = 2/Ca
    f5 = 1

    f6 = 8/(Si+Al)
    f7 = 15/(Z-K)
    f8 = 13/(Z-Na-K-Ca-Mn)
    f9 = 36/(46-Al-Si-Ti)
    f10 = 46/(46+Fe_total)

    fA = minimum of capped f1-f5.
    fB = maximum of f6-f10.
    fav = (fA+fB)/2.
    """

    missing = [
        element
        for element in ELEMENTS
        if c.get(element) is None
    ]

    if missing:

        raise ValueError(
            f"Missing initial cations: {missing}"
        )

    Si = c["Si"]
    Ti = c["Ti"]
    Al = c["Al"]
    Fe = c["Fe_total"]
    Mn = c["Mn"]
    Mg = c["Mg"]
    Ca = c["Ca"]
    Na = c["Na"]
    K = c["K"]

    if any(
        c[element] < 0
        for element in ELEMENTS
    ):

        raise ValueError(
            "Negative initial cation abundance."
        )

    Z = sum(
        c[element]
        for element in ELEMENTS
    )

    denominators = {

        "f1": Z,

        "f2": Si,

        "f3": Z - Na - K,

        "f4": Ca,

        "f6": Si + Al,

        "f7": Z - K,

        "f8": Z - Na - K - Ca - Mn,

        "f9": 46.0 - Al - Si - Ti,

        "f10": 46.0 + Fe

    }

    for name, denominator in denominators.items():

        if denominator <= 0:

            raise ValueError(
                f"Invalid denominator for {name}: "
                f"{denominator}"
            )

    factors = {

        "f1": 16.0 / Z,

        "f2": 8.0 / Si,

        "f3": 15.0 / (
            Z - Na - K
        ),

        "f4": 2.0 / Ca,

        "f5": 1.0,

        "f6": 8.0 / (
            Si + Al
        ),

        "f7": 15.0 / (
            Z - K
        ),

        "f8": 13.0 / (
            Z - Na - K - Ca - Mn
        ),

        "f9": 36.0 / (
            46.0 - Al - Si - Ti
        ),

        "f10": 46.0 / (
            46.0 + Fe
        )

    }

    # Original upper constraint.

    group_A = [
        min(factors[f"f{i}"], 1.0)
        for i in range(1, 6)
    ]

    fA = min(group_A)

    # Original lower constraint.

    group_B = [
        factors[f"f{i}"]
        for i in range(6, 11)
    ]

    fB = max(group_B)

    # Original averaging procedure.

    fav = (
        fA + fB
    ) / 2.0

    return {

        **factors,

        "Z": Z,

        "fA": fA,

        "fB": fB,

        "fav": fav,

        "Constraint_Interval_Valid": (
            fB <= fA + ABS_TOL
        )

    }


# =============================================================================
# H. RECONSTRUCT IRON PARTITION
# =============================================================================

def calculate_iron(c, factors):

    fav = factors["fav"]

    Fe_initial = c["Fe_total"]

    # Total revised iron.

    Fe_total_revised = (
        Fe_initial * fav
    )

    # Original workbook Fe3 calculation.

    Fe3_raw = (
        46.0 * (1.0 - fav)
    )

    # H-column negative-value instruction:
    # Negative Fe3 values are replaced with zero.

    Fe3_H_final = max(
        0.0,
        Fe3_raw
    )

    # Revised ferrous iron.

    Fe2_revised = (
        Fe_total_revised - Fe3_H_final
    )

    return {

        "Fe_Total_Initial": Fe_initial,

        "Fe_Total_Revised": Fe_total_revised,

        "Fe3_Raw": Fe3_raw,

        "Fe3_H_Final": Fe3_H_final,

        "Fe2_Revised": Fe2_revised,

        "Fe2_Negative": (
            Fe2_revised < -ABS_TOL
        ),

        "Fe3_Exceeds_Revised_Fe": (
            Fe3_H_final
            > Fe_total_revised + ABS_TOL
        )

    }


# =============================================================================
# I. RECONSTRUCT ALL REVISED CATIONS
# =============================================================================

def calculate_revised_cations(c, iron, factors):

    fav = factors["fav"]

    revised = {

        "Si": c["Si"] * fav,

        "Ti": c["Ti"] * fav,

        "Al": c["Al"] * fav,

        "Fe2": iron["Fe2_Revised"],

        "Fe3": iron["Fe3_H_Final"],

        "Mn": c["Mn"] * fav,

        "Mg": c["Mg"] * fav,

        "Ca": c["Ca"] * fav,

        "Na": c["Na"] * fav,

        "K": c["K"] * fav

    }

    revised_sum = sum(
        revised.values()
    )

    return revised, revised_sum


# =============================================================================
# J. INITIAL CATION COMPARISON
# =============================================================================

def build_input_comparison(
    sample,
    original,
    independent
):

    output = []

    for element in ELEMENTS:

        original_value = original[element]

        independent_value = independent[element]

        output.append({

            "Sample": sample,

            "Element": element,

            "Original_XLS_Value": original_value,

            "Independent_Reconstruction": independent_value,

            "Difference": difference(
                independent_value,
                original_value
            ),

            "Status": status(
                independent_value,
                original_value
            )

        })

    return output


# =============================================================================
# K. FACTOR COMPARISON
# =============================================================================

def build_factor_comparison(
    sample,
    factors,
    original
):

    output = []

    for i in range(1, 11):

        name = f"f{i}"

        calculated = factors[name]

        original_value = original[name]

        output.append({

            "Sample": sample,

            "Factor": name,

            "Independent_Value": calculated,

            "Original_XLS_Value": original_value,

            "Difference": difference(
                calculated,
                original_value
            ),

            "Absolute_Difference": (
                abs(calculated - original_value)
                if original_value is not None
                else None
            ),

            "Status": status(
                calculated,
                original_value
            )

        })

    return output


# =============================================================================
# L. REVISED CATION COMPARISON
# =============================================================================

def build_revised_comparison(
    sample,
    revised,
    original
):

    original_keys = {

        "Si": "Revised_Si",

        "Ti": "Revised_Ti",

        "Al": "Revised_Al",

        "Fe2": "Revised_Fe2",

        "Fe3": "Revised_Fe3",

        "Mn": "Revised_Mn",

        "Mg": "Revised_Mg",

        "Ca": "Revised_Ca",

        "Na": "Revised_Na",

        "K": "Revised_K"

    }

    output = []

    for element, original_key in original_keys.items():

        calculated = revised[element]

        original_value = original[original_key]

        output.append({

            "Sample": sample,

            "Element": element,

            "Independent_Value": calculated,

            "Original_XLS_Value": original_value,

            "Difference": difference(
                calculated,
                original_value
            ),

            "Absolute_Difference": (
                abs(calculated - original_value)
                if original_value is not None
                else None
            ),

            "Status": status(
                calculated,
                original_value
            )

        })

    return output


# =============================================================================
# M. SAMPLE-LEVEL VERIFICATION
# =============================================================================

def build_sample_summary(
    sample,
    c,
    original,
    factors,
    iron,
    revised,
    revised_sum,
    input_results,
    factor_results,
    revised_results
):

    input_pass = sum(
        r["Status"] == "PASS"
        for r in input_results
    )

    factor_pass = sum(
        r["Status"] == "PASS"
        for r in factor_results
    )

    revised_pass = sum(
        r["Status"] == "PASS"
        for r in revised_results
    )

    # Constraint summaries.

    minimum_check = numerical_check(
        factors["fA"],
        original["Minimum"]
    )

    maximum_check = numerical_check(
        factors["fB"],
        original["Maximum"]
    )

    fav_check = numerical_check(
        factors["fav"],
        original["fav"]
    )

    # Iron comparisons.

    fe3_ar_check = numerical_check(
        iron["Fe3_Raw"],
        original["Fe3_AR"]
    )

    fe3_h_check = numerical_check(
        iron["Fe3_H_Final"],
        original["Fe3_H"]
    )

    fe2_check = numerical_check(
        iron["Fe2_Revised"],
        original["Revised_Fe2"]
    )

    # Initial cation sum.

    initial_sum = sum(c.values())

    initial_sum_check = numerical_check(
        initial_sum,
        original["Initial_Sum"]
    )

    # Revised sum.

    revised_sum_check = numerical_check(
        revised_sum,
        original["Revised_Sum"]
    )

    # All numerical checks.

    all_checks = (

        [r["Status"] == "PASS" for r in input_results]

        + [r["Status"] == "PASS" for r in factor_results]

        + [r["Status"] == "PASS" for r in revised_results]

        + [
            minimum_check,
            maximum_check,
            fav_check,
            fe3_ar_check,
            fe3_h_check,
            fe2_check,
            initial_sum_check,
            revised_sum_check
        ]

    )

    total_pass = sum(all_checks)

    total_fail = len(all_checks) - total_pass

    return {

        "Sample": sample,

        "Initial_Cation_Sum_Independent": initial_sum,

        "Initial_Cation_Sum_Original": original["Initial_Sum"],

        "Initial_Sum_Check": (
            "PASS" if initial_sum_check else "REVIEW"
        ),

        "fA_Independent": factors["fA"],

        "fA_Original": original["Minimum"],

        "fA_Check": (
            "PASS" if minimum_check else "REVIEW"
        ),

        "fB_Independent": factors["fB"],

        "fB_Original": original["Maximum"],

        "fB_Check": (
            "PASS" if maximum_check else "REVIEW"
        ),

        "fav_Independent": factors["fav"],

        "fav_Original": original["fav"],

        "fav_Check": (
            "PASS" if fav_check else "REVIEW"
        ),

        "Constraint_Interval_Valid": (
            factors["Constraint_Interval_Valid"]
        ),

        "Fe3_Raw_Independent": iron["Fe3_Raw"],

        "Fe3_AR_Original": original["Fe3_AR"],

        "Fe3_AR_Check": (
            "PASS" if fe3_ar_check else "REVIEW"
        ),

        "Fe3_H_Final_Independent": iron["Fe3_H_Final"],

        "Fe3_H_Original": original["Fe3_H"],

        "Fe3_H_Check": (
            "PASS" if fe3_h_check else "REVIEW"
        ),

        "Fe2_Independent": iron["Fe2_Revised"],

        "Fe2_Original": original["Revised_Fe2"],

        "Fe2_Check": (
            "PASS" if fe2_check else "REVIEW"
        ),

        "Revised_Cation_Sum_Independent": revised_sum,

        "Revised_Cation_Sum_Original": original["Revised_Sum"],

        "Revised_Sum_Check": (
            "PASS" if revised_sum_check else "REVIEW"
        ),

        "Input_Cation_Passes": input_pass,

        "Factor_Passes": factor_pass,

        "Revised_Cation_Passes": revised_pass,

        "Total_Checks": len(all_checks),

        "Total_Pass": total_pass,

        "Total_Review": total_fail,

        "Arithmetic_Status": (
            "PASS" if total_fail == 0 else "REVIEW"
        ),

        "Scientific_Certification": "NOT ESTABLISHED"

    }


# =============================================================================
# N. EXCEL REPORT WRITING
# =============================================================================

def write_records(wb, sheet_name, records):

    ws = wb.create_sheet(
        sheet_name
    )

    if not records:
        return ws

    headers = list(
        records[0].keys()
    )

    ws.append(headers)

    for record in records:

        ws.append([
            record.get(header)
            for header in headers
        ])

    return ws


def format_workbook(wb):

    header_fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    header_font = Font(
        bold=True,
        color="FFFFFF"
    )

    thin = Side(
        style="thin",
        color="D9E2F3"
    )

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        if ws.max_row > 1 and ws.max_column > 0:

            ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:

            cell.fill = header_fill
            cell.font = header_font

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        for row in ws.iter_rows():

            for cell in row:

                cell.border = Border(
                    bottom=thin
                )

                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

        for column in range(
            1,
            ws.max_column + 1
        ):

            maximum = 0

            for row in range(
                1,
                min(ws.max_row, 300) + 1
            ):

                value = ws.cell(
                    row,
                    column
                ).value

                if value is not None:

                    maximum = max(
                        maximum,
                        len(str(value))
                    )

            ws.column_dimensions[
                get_column_letter(column)
            ].width = min(
                max(14, maximum + 2),
                45
            )

        ws.row_dimensions[1].height = 35


# =============================================================================
# O. MAIN EXECUTION
# =============================================================================

def main():

    logger.info("=" * 85)

    logger.info(
        "STEP 06.3.2 — EXACT ORIGINAL HBL97 OPTIMIZATION RECONSTRUCTION"
    )

    logger.info("=" * 85)

    # ---------------------------------------------------------
    # Verify inputs
    # ---------------------------------------------------------

    for path in [
        ORIGINAL_XLS,
        CHEM_FILE
    ]:

        if not path.exists():

            raise FileNotFoundError(
                f"Required input missing: {path}"
            )

        logger.info(
            "Verified input: %s",
            path
        )

    source_hash = sha256_file(
        ORIGINAL_XLS
    )

    logger.info(
        "Original XLS SHA256: %s",
        source_hash
    )

    # ---------------------------------------------------------
    # Load original workbook and independent chemistry
    # ---------------------------------------------------------

    original_records = read_original_workbook()

    independent_records = read_initial_cations()

    samples = sorted(
        set(original_records)
        & set(independent_records)
    )

    if len(samples) != EXPECTED_SAMPLES:

        raise ValueError(
            f"Expected {EXPECTED_SAMPLES} matched samples; "
            f"recovered {len(samples)}."
        )

    logger.info(
        "Matched samples: %d",
        len(samples)
    )

    # ---------------------------------------------------------
    # Output containers
    # ---------------------------------------------------------

    input_output = []

    factor_output = []

    revised_output = []

    iron_output = []

    summary_output = []

    formula_output = []

    # ---------------------------------------------------------
    # Sample processing
    # ---------------------------------------------------------

    for sample in samples:

        logger.info(
            "Processing %s",
            sample
        )

        original = original_records[sample]

        independent = independent_records[sample]

        # Use the original workbook's initial cations
        # for the exact workbook reconstruction.

        c_original = {

            element: original[element]

            for element in ELEMENTS

        }

        # Compare initial cations against Step 05C4.1.

        input_results = build_input_comparison(
            sample,
            c_original,
            independent
        )

        input_output.extend(
            input_results
        )

        # Reconstruct original factors.

        factors = calculate_hbl97_factors(
            c_original
        )

        factor_results = build_factor_comparison(
            sample,
            factors,
            original
        )

        factor_output.extend(
            factor_results
        )

        # Reconstruct iron.

        iron = calculate_iron(
            c_original,
            factors
        )

        # Reconstruct all revised cations.

        revised, revised_sum = calculate_revised_cations(
            c_original,
            iron,
            factors
        )

        revised_results = build_revised_comparison(
            sample,
            revised,
            original
        )

        revised_output.extend(
            revised_results
        )

        # Iron audit.

        iron_output.append({

            "Sample": sample,

            "Initial_Fe_Total": iron["Fe_Total_Initial"],

            "fav": factors["fav"],

            "Revised_Fe_Total": iron["Fe_Total_Revised"],

            "Independent_Fe3_Raw": iron["Fe3_Raw"],

            "Original_AR_Fe3": original["Fe3_AR"],

            "Independent_Fe3_H_Final": iron["Fe3_H_Final"],

            "Original_H_Fe3": original["Fe3_H"],

            "Independent_Fe2": iron["Fe2_Revised"],

            "Original_AW_Fe2": original["Revised_Fe2"],

            "Fe3_AR_Check": status(
                iron["Fe3_Raw"],
                original["Fe3_AR"]
            ),

            "Fe3_H_Check": status(
                iron["Fe3_H_Final"],
                original["Fe3_H"]
            ),

            "Fe2_Check": status(
                iron["Fe2_Revised"],
                original["Revised_Fe2"]
            ),

            "Negative_Fe2": iron["Fe2_Negative"],

            "Fe3_Exceeds_Revised_Fe": (
                iron["Fe3_Exceeds_Revised_Fe"]
            )

        })

        # Formula summary.

        formula_output.append({

            "Sample": sample,

            "Z": factors["Z"],

            **{
                f"f{i}": factors[f"f{i}"]
                for i in range(1, 11)
            },

            "fA": factors["fA"],

            "fB": factors["fB"],

            "fav": factors["fav"],

            "Fe3_Raw": iron["Fe3_Raw"],

            "Fe3_Final": iron["Fe3_H_Final"],

            "Fe2_Final": iron["Fe2_Revised"],

            "Revised_Cation_Sum": revised_sum

        })

        # Sample-level summary.

        summary = build_sample_summary(
            sample,
            c_original,
            original,
            factors,
            iron,
            revised,
            revised_sum,
            input_results,
            factor_results,
            revised_results
        )

        summary_output.append(
            summary
        )

    # =============================================================================
    # P. GLOBAL STATISTICS
    # =============================================================================

    total_input = len(input_output)

    input_pass = sum(
        r["Status"] == "PASS"
        for r in input_output
    )

    total_factors = len(factor_output)

    factor_pass = sum(
        r["Status"] == "PASS"
        for r in factor_output
    )

    total_revised = len(revised_output)

    revised_pass = sum(
        r["Status"] == "PASS"
        for r in revised_output
    )

    total_checks = sum(
        r["Total_Checks"]
        for r in summary_output
    )

    total_pass = sum(
        r["Total_Pass"]
        for r in summary_output
    )

    total_review = total_checks - total_pass

    fully_reproduced = sum(
        r["Arithmetic_Status"] == "PASS"
        for r in summary_output
    )

    # =============================================================================
    # Q. CREATE EXCEL REPORT
    # =============================================================================

    wb = openpyxl.Workbook()

    ws = wb.active
    ws.title = "Read_Me"

    readme = [

        ["Item", "Description"],

        ["Step",
         "06.3.2 — Exact original Hbl97 optimization reconstruction"],

        ["Source workbook",
         ORIGINAL_XLS.name],

        ["Source SHA256",
         source_hash],

        ["Matched samples",
         len(samples)],

        ["Initial cation basis",
         "Original workbook U:AC numerical values"],

        ["Independent chemistry comparison",
         "Step 05C4.1 Original_Calculation_Blocks"],

        ["f8 equation",
         "13/(Z-Na-K-Ca-Mn)"],

        ["f9 equation",
         "36/(46-Al-Si-Ti)"],

        ["f10 equation",
         "46/(46+Fe_total)"],

        ["fA",
         "Minimum of capped f1-f5"],

        ["fB",
         "Maximum of f6-f10"],

        ["fav",
         "(fA+fB)/2"],

        ["Fe3+",
         "46*(1-fav), with negative H-column values set to zero"],

        ["Fe2+",
         "Revised total Fe minus final Fe3+"],

        ["Numerical tolerance",
         f"Absolute={ABS_TOL}; Relative={REL_TOL}"],

        ["Scientific status",
         "Numerical reproduction does not certify geological validity."],

        ["Execution time",
         datetime.now().isoformat()]

    ]

    for row in readme:
        ws.append(row)

    # Overall summary.

    ws = wb.create_sheet(
        "Overall_Summary",
        1
    )

    overall = [

        ["Metric", "Value"],

        ["Samples", len(samples)],

        ["Initial cation comparisons", total_input],

        ["Initial cation PASS", input_pass],

        ["Factor comparisons", total_factors],

        ["Factor PASS", factor_pass],

        ["Factor REVIEW", total_factors - factor_pass],

        ["Revised cation comparisons", total_revised],

        ["Revised cation PASS", revised_pass],

        ["Revised cation REVIEW",
         total_revised - revised_pass],

        ["Integrated comparisons", total_checks],

        ["Integrated PASS", total_pass],

        ["Integrated REVIEW", total_review],

        ["Samples fully reproduced", fully_reproduced],

        ["Samples requiring numerical review",
         len(samples) - fully_reproduced],

        ["Scientific certification",
         "NOT ESTABLISHED"],

        ["Independent geological validation",
         "PENDING"]

    ]

    for row in overall:
        ws.append(row)

    # Detail sheets.

    write_records(
        wb,
        "Initial_Cation_Comparison",
        input_output
    )

    write_records(
        wb,
        "Factor_Reconstruction",
        factor_output
    )

    write_records(
        wb,
        "Iron_Reconstruction",
        iron_output
    )

    write_records(
        wb,
        "Revised_Cation_Comparison",
        revised_output
    )

    write_records(
        wb,
        "Sample_Summary",
        summary_output
    )

    write_records(
        wb,
        "Independent_Equations",
        formula_output
    )

    # Scientific limitations.

    limitations = [

        {
            "Issue": "Fe3+ provenance",

            "Finding": "Fe3+ is reconstructed using the original Hbl97 optimization procedure.",

            "Limitation": "This is not an independent spectroscopic Fe3+ measurement.",

            "Required_Evidence": "Independent oxidation-state evidence where available."
        },

        {
            "Issue": "Amphibole site allocation",

            "Finding": "The original workbook's numerical optimization is reproduced.",

            "Limitation": "Exact numerical reproduction does not independently validate crystallographic occupancy.",

            "Required_Evidence": "Independent structural formula and site-allocation validation."
        },

        {
            "Issue": "Calibration accuracy",

            "Finding": "Original workbook calculations can be independently reproduced.",

            "Limitation": "Reproduction does not establish geological accuracy or precision.",

            "Required_Evidence": "Independent geological and experimental validation."
        },

        {
            "Issue": "Publication provenance",

            "Finding": "The original workbook contains the 13 sample calculations.",

            "Limitation": "Sample-level publication provenance requires separate documentary verification.",

            "Required_Evidence": "Original publication and analytical data crosswalk."
        }

    ]

    write_records(
        wb,
        "Scientific_Limitations",
        limitations
    )

    format_workbook(wb)

    wb.save(
        OUTPUT_FILE
    )

    # =============================================================================
    # R. FINAL CONSOLE SUMMARY
    # =============================================================================

    logger.info("=" * 85)

    logger.info(
        "STEP 06.3.2 COMPLETED"
    )

    logger.info(
        "Samples: %d",
        len(samples)
    )

    logger.info(
        "Initial cation comparisons: %d | PASS: %d",
        total_input,
        input_pass
    )

    logger.info(
        "Factor comparisons: %d | PASS: %d",
        total_factors,
        factor_pass
    )

    logger.info(
        "Revised cation comparisons: %d | PASS: %d",
        total_revised,
        revised_pass
    )

    logger.info(
        "Integrated comparisons: %d | PASS: %d | REVIEW: %d",
        total_checks,
        total_pass,
        total_review
    )

    logger.info(
        "Fully reproduced samples: %d/%d",
        fully_reproduced,
        len(samples)
    )

    logger.info(
        "Scientific certification: NOT ESTABLISHED"
    )

    logger.info(
        "Excel report: %s",
        OUTPUT_FILE
    )

    logger.info(
        "Execution log: %s",
        LOG_FILE
    )

    logger.info("=" * 85)


# =============================================================================
# S. EXECUTION
# =============================================================================

if __name__ == "__main__":

    try:

        main()

    except Exception:

        logger.exception(
            "STEP 06.3.2 FAILED"
        )

        raise
