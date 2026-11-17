# -*- coding: utf-8 -*-
"""
Created on Tue Nov 17 11:05:07 2026

@author: hussa
"""


# =============================================================================
# STEP 06.3
# INDEPENDENT VERIFICATION OF THE PUBLISHED Fe3+ OPTIMIZATION METHOD
#
# Scientific basis:
# Holland & Blundy (1994), Appendix B, p. 445.
#
# Purpose:
# 1. Independently reconstruct f1-f10.
# 2. Recalculate upper and lower constraints.
# 3. Independently calculate fav.
# 4. Reconstruct Fe3+ and Fe2+.
# 5. Compare with original Hbl97 workbook.
# 6. Investigate negative intermediate Fe3+ values.
#
# IMPORTANT:
# Arithmetic agreement is not equivalent to independent analytical
# determination of Fe3+ or crystallographic site validation.
# =============================================================================

from pathlib import Path
import math
import hashlib
import logging
from datetime import datetime

import xlrd
import openpyxl

from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment
)
from openpyxl.utils import get_column_letter


# =============================================================================
# A. PROJECT PATHS
# =============================================================================

ROOT = Path(
    r"D:\Data Science Assignment\My Projects\Geothermobarometer_ML_Project"
)

SOURCE_DIR = ROOT / "Geothermobarometers Calibrations"

ORIGINAL_XLS = SOURCE_DIR / "Hbl P-T (1997).xls"

CHEM_FILE = (
    ROOT /
    "05C4_1_Hbl97_original_chemical_reconstruction" /
    "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

OPT_FILE = (
    ROOT /
    "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit" /
    "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

STEP061_FILE = (
    ROOT /
    "06_1_Hbl97_Independent_Iron_Normalization_Verification" /
    "Step06_1_Hbl97_Independent_Verification.xlsx"
)

OUTPUT_DIR = (
    ROOT /
    "06_3_Hbl97_Published_Fe3_Method_Verification"
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OUTPUT_FILE = (
    OUTPUT_DIR /
    "Step06_3_Hbl97_Published_Fe3_Method_Verification.xlsx"
)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / "step06_3_execution_log.txt"


# =============================================================================
# B. LOGGING
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
    ]
)

logger = logging.getLogger("STEP06_3")


# =============================================================================
# C. CONSTANTS
# =============================================================================

TOL = 1e-8

OXYGEN_BASIS = 23

FE_FACTOR = 46.0

EXPECTED_SAMPLES = 13

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


def clean_number(value):

    if value is None:
        return None

    try:

        value = float(value)

        if math.isfinite(value):
            return value

    except (TypeError, ValueError):
        pass

    return None


def close_enough(a, b):

    if a is None or b is None:
        return False

    return abs(a - b) <= TOL


def read_xlsx_records(path, sheet_name):

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

    rows = ws.iter_rows(values_only=True)

    headers = next(rows)

    records = []

    for row in rows:

        record = dict(
            zip(headers, row)
        )

        records.append(record)

    wb.close()

    return records


# =============================================================================
# E. INDEPENDENT RECONSTRUCTION OF HOLLAND-BLUNDY 1994 FACTORS
# =============================================================================

def calculate_published_factors(c):

    Si = c["Si"]
    Ti = c["Ti"]
    Al = c["Al"]
    Fe = c["Fe_total"]
    Mn = c["Mn"]
    Mg = c["Mg"]
    Ca = c["Ca"]
    Na = c["Na"]
    K = c["K"]

    Z = sum(
        c[element]
        for element in ELEMENTS
    )

    if any(
        value is None
        for value in c.values()
    ):

        raise ValueError(
            "Incomplete initial cation composition."
        )

    denominators = {
        "f1": Z,
        "f2": Si,
        "f3": Z - Na - K,
        "f4": Ca,
        "f6": Si + Al,
        "f7": Z - K,
        "f8": Z - Na - K - Ca,
        "f9": 46.0 + Al + Si + Ti
    }

    for name, value in denominators.items():

        if value <= 0:

            raise ValueError(
                f"Invalid denominator for {name}: {value}"
            )

    # Published Appendix B equations, page 445.

    factors = {

        "f1": 16.0 / Z,

        "f2": 8.0 / Si,

        "f3": 15.0 / (Z - Na - K),

        "f4": 2.0 / Ca,

        "f5": 1.0,

        "f6": 8.0 / (Si + Al),

        "f7": 15.0 / (Z - K),

        "f8": 12.9 / (Z - Na - K - Ca),

        "f9": 36.0 / (46.0 + Al + Si + Ti),

        "f10": 1.0 - (Fe / 46.0)

    }

    # Group A: upper constraints.
    #
    # Any factor > 1 is set to 1.

    group_A = [
        min(factors[f"f{i}"], 1.0)
        for i in range(1, 6)
    ]

    fA = min(group_A)

    # Group B: lower constraints.

    group_B = [
        factors[f"f{i}"]
        for i in range(6, 11)
    ]

    fB = max(group_B)

    # A group B factor > 1 means the analysis cannot
    # be made stoichiometric under this procedure.

    group_B_invalid = any(
        value > 1.0 + TOL
        for value in group_B
    )

    # Published average factor.

    fav = (fA + fB) / 2.0

    # Fe3+ from the 23-oxygen calculation.

    Fe3_raw = FE_FACTOR * (1.0 - fav)

    # Workbook instruction: negative H-column Fe3+ is set to zero.

    Fe3_final = max(0.0, Fe3_raw)

    # Revised total Fe.

    Fe_total_revised = fav * Fe

    # Ferrous iron after ferric iron allocation.

    Fe2_final = Fe_total_revised - Fe3_final

    return {

        **factors,

        "Z": Z,

        "fA": fA,

        "fB": fB,

        "fav": fav,

        "Group_B_Invalid": group_B_invalid,

        "Fe3_Raw": Fe3_raw,

        "Fe3_Final": Fe3_final,

        "Fe_Total_Revised": Fe_total_revised,

        "Fe2_Final": Fe2_final

    }


# =============================================================================
# F. LOAD SOURCE DATA
# =============================================================================

def main():

    logger.info("=" * 79)

    logger.info(
        "STEP 06.3 — PUBLISHED Fe3+ METHOD VERIFICATION"
    )

    logger.info("=" * 79)

    required_files = [
        ORIGINAL_XLS,
        CHEM_FILE,
        OPT_FILE,
        STEP061_FILE
    ]

    for path in required_files:

        if not path.exists():

            raise FileNotFoundError(
                f"Missing required input:\n{path}"
            )

        logger.info(
            "Verified: %s",
            path.name
        )

    source_hash = sha256_file(
        ORIGINAL_XLS
    )

    logger.info(
        "Original workbook SHA256: %s",
        source_hash
    )

    # Load independently reconstructed initial cations.

    calculation_blocks = read_xlsx_records(
        CHEM_FILE,
        "Original_Calculation_Blocks"
    )

    initial_records = {

        record["Sample"]: record

        for record in calculation_blocks

        if record["Calculation_Block"] == "Initial_Cations"

    }

    # Load original optimization factors.

    optimization_records = read_xlsx_records(
        CHEM_FILE,
        "Optimization_Factors"
    )

    optimization_map = {

        record["Sample"]: record

        for record in optimization_records

    }

    # Load previous arithmetic audit.

    previous_audit = read_xlsx_records(
        OPT_FILE,
        "Optimization_Audit"
    )

    previous_map = {

        record["Sample"]: record

        for record in previous_audit

    }

    # Load previous independent verification.

    step061_records = read_xlsx_records(
        STEP061_FILE,
        "Fe3_Reconciliation"
    )

    step061_map = {

        record["Sample"]: record

        for record in step061_records

    }

    # Open original XLS workbook.

    xls = xlrd.open_workbook(
        str(ORIGINAL_XLS),
        on_demand=True
    )

    ws = xls.sheet_by_name(
        "Hbl P-T (1997)"
    )

    original_map = {}

    for row_index in range(9, 22):

        sample = ws.cell_value(
            row_index,
            1
        )

        if not sample:
            continue

        original_map[str(sample)] = {

            "Excel_Row": row_index + 1,

            "Fe3_H": clean_number(
                ws.cell_value(row_index, 7)
            ),

            "Fe3_AR": clean_number(
                ws.cell_value(row_index, 43)
            ),

            "Fe2_AW": clean_number(
                ws.cell_value(row_index, 48)
            ),

            "Fe3_AX": clean_number(
                ws.cell_value(row_index, 49)
            ),

            "f1": clean_number(
                ws.cell_value(row_index, 30)
            ),

            "f2": clean_number(
                ws.cell_value(row_index, 31)
            ),

            "f3": clean_number(
                ws.cell_value(row_index, 32)
            ),

            "f4": clean_number(
                ws.cell_value(row_index, 33)
            ),

            "f5": clean_number(
                ws.cell_value(row_index, 34)
            ),

            "f6": clean_number(
                ws.cell_value(row_index, 36)
            ),

            "f7": clean_number(
                ws.cell_value(row_index, 37)
            ),

            "f8": clean_number(
                ws.cell_value(row_index, 38)
            ),

            "f9": clean_number(
                ws.cell_value(row_index, 39)
            ),

            "f10": clean_number(
                ws.cell_value(row_index, 40)
            ),

            "Original_Minimum": clean_number(
                ws.cell_value(row_index, 35)
            ),

            "Original_Maximum": clean_number(
                ws.cell_value(row_index, 41)
            ),

            "Original_fav": clean_number(
                ws.cell_value(row_index, 42)
            )

        }

    xls.release_resources()

    samples = sorted(
        set(initial_records)
        & set(optimization_map)
        & set(original_map)
    )

    if len(samples) != EXPECTED_SAMPLES:

        raise ValueError(
            f"Expected {EXPECTED_SAMPLES} samples, "
            f"recovered {len(samples)}."
        )

    logger.info(
        "Samples available for independent verification: %s",
        len(samples)
    )


    # =============================================================================
    # G. SAMPLE-LEVEL VERIFICATION
    # =============================================================================

    factor_results = []

    iron_results = []

    summary_results = []

    for sample in samples:

        initial = initial_records[sample]

        original = original_map[sample]

        reconstructed = optimization_map[sample]

        cations = {

            element: clean_number(
                initial[element]
            )

            for element in ELEMENTS

        }

        result = calculate_published_factors(
            cations
        )

        # ---------------------------------------------------------
        # Factor comparison
        # ---------------------------------------------------------

        factor_checks = []

        for i in range(1, 11):

            name = f"f{i}"

            independent_value = result[name]

            original_value = original[name]

            extracted_value = clean_number(
                reconstructed.get(name)
            )

            original_check = close_enough(
                independent_value,
                original_value
            )

            extracted_check = close_enough(
                independent_value,
                extracted_value
            )

            factor_checks.extend([
                original_check,
                extracted_check
            ])

            factor_results.append({

                "Sample": sample,

                "Factor": name,

                "Independent_Published_Value": independent_value,

                "Original_XLS_Value": original_value,

                "Extracted_Report_Value": extracted_value,

                "Original_Check": (
                    "PASS" if original_check else "FAIL"
                ),

                "Report_Check": (
                    "PASS" if extracted_check else "FAIL"
                )

            })

        # ---------------------------------------------------------
        # fA, fB and fav comparison
        # ---------------------------------------------------------

        checks = {

            "fA": close_enough(
                result["fA"],
                original["Original_Minimum"]
            ),

            "fB": close_enough(
                result["fB"],
                original["Original_Maximum"]
            ),

            "fav": close_enough(
                result["fav"],
                original["Original_fav"]
            )

        }

        # ---------------------------------------------------------
        # Fe3 and Fe2 comparison
        # ---------------------------------------------------------

        raw_check = close_enough(
            result["Fe3_Raw"],
            original["Fe3_AR"]
        )

        final_check = close_enough(
            result["Fe3_Final"],
            original["Fe3_H"]
        )

        fe3_revised_check = close_enough(
            result["Fe3_Final"],
            original["Fe3_AX"]
        )

        fe2_check = close_enough(
            result["Fe2_Final"],
            original["Fe2_AW"]
        )

        iron_results.append({

            "Sample": sample,

            "Initial_Fe_Total": cations["Fe_total"],

            "Z_Initial_Cations": result["Z"],

            "fA_Independent": result["fA"],

            "fB_Independent": result["fB"],

            "fav_Independent": result["fav"],

            "Original_fav": original["Original_fav"],

            "Fe3_Raw_Independent": result["Fe3_Raw"],

            "Fe3_AR_Original": original["Fe3_AR"],

            "Fe3_Final_Independent": result["Fe3_Final"],

            "Fe3_H_Original": original["Fe3_H"],

            "Fe3_AX_Original": original["Fe3_AX"],

            "Fe2_Independent": result["Fe2_Final"],

            "Fe2_AW_Original": original["Fe2_AW"],

            "Group_B_Invalid": result["Group_B_Invalid"],

            "Raw_Fe3_Check": "PASS" if raw_check else "FAIL",

            "Final_H_Fe3_Check": "PASS" if final_check else "FAIL",

            "Revised_AX_Fe3_Check": (
                "PASS" if fe3_revised_check else "FAIL"
            ),

            "Fe2_Check": "PASS" if fe2_check else "FAIL"

        })

        # ---------------------------------------------------------
        # Sample-level integrated status
        # ---------------------------------------------------------

        all_checks = (
            factor_checks
            + list(checks.values())
            + [
                raw_check,
                final_check,
                fe3_revised_check,
                fe2_check
            ]
        )

        failures = sum(
            not value
            for value in all_checks
        )

        summary_results.append({

            "Sample": sample,

            "Factor_Checks": len(factor_checks),

            "Factor_Passes": sum(factor_checks),

            "Factor_Failures": len(factor_checks) - sum(factor_checks),

            "fA_Check": "PASS" if checks["fA"] else "FAIL",

            "fB_Check": "PASS" if checks["fB"] else "FAIL",

            "fav_Check": "PASS" if checks["fav"] else "FAIL",

            "Fe3_Raw_Check": "PASS" if raw_check else "FAIL",

            "Fe3_Final_Check": "PASS" if final_check else "FAIL",

            "Fe2_Check": "PASS" if fe2_check else "FAIL",

            "Group_B_Validity": (
                "FAIL" if result["Group_B_Invalid"] else "PASS"
            ),

            "Total_Failures": failures,

            "Independent_Method_Status": (
                "PASS" if failures == 0 else "REVIEW"
            )

        })

        logger.info(
            "Processed %s",
            sample
        )


    # =============================================================================
    # H. CREATE REPORT
    # =============================================================================

    wb = openpyxl.Workbook()

    ws_readme = wb.active

    ws_readme.title = "Read_Me"

    readme = [

        ["Item", "Description"],

        ["Step", "06.3 — Independent published Fe3+ method verification"],

        ["Scientific source",
         "Holland and Blundy (1994), Appendix B, page 445"],

        ["Original workbook", ORIGINAL_XLS.name],

        ["Original workbook SHA256", source_hash],

        ["Samples", len(samples)],

        ["Oxygen basis", "23 oxygens"],

        ["Fe conversion factor", 46.0],

        ["Tolerance", TOL],

        ["Purpose",
         "Independent arithmetic reconstruction of published Fe3+ method"],

        ["Limitation",
         "Does not establish independent analytical Fe3+ measurement"],

        ["Site allocation",
         "Requires a separate mineralogical validation"],

        ["Execution time", datetime.now().isoformat()]

    ]

    for row in readme:
        ws_readme.append(row)

    # Write data sheets.

    def add_sheet(name, records):

        ws = wb.create_sheet(name)

        if not records:
            return ws

        headers = list(records[0].keys())

        ws.append(headers)

        for record in records:

            ws.append([
                record.get(header)
                for header in headers
            ])

        return ws

    add_sheet(
        "Factor_Recalculation",
        factor_results
    )

    add_sheet(
        "Iron_Recalculation",
        iron_results
    )

    add_sheet(
        "Sample_Summary",
        summary_results
    )

    # ---------------------------------------------------------
    # Scientific limitations
    # ---------------------------------------------------------

    limitations = [

        {
            "Issue": "Fe3+ analytical provenance",

            "Finding": "The published procedure reconstructs Fe3+ from stoichiometric constraints.",

            "Limitation": "This is a calculated Fe3+ estimate, not an independent spectroscopic measurement.",

            "Required_Evidence": "Independent Fe oxidation-state measurements, if available."
        },

        {
            "Issue": "Negative raw Fe3+",

            "Finding": "A negative value may arise when fav exceeds 1.",

            "Limitation": "The original workbook instructs that negative H-column values be set to zero.",

            "Required_Evidence": "Confirm the original workbook implementation and distinguish raw AR from final H."
        },

        {
            "Issue": "Stoichiometric constraints",

            "Finding": "Ten constraints are reconstructed from Appendix B.",

            "Limitation": "A mathematical PASS does not prove all mineralogical assumptions.",

            "Required_Evidence": "Independent mineralogical and analytical validation."
        },

        {
            "Issue": "Crystallographic sites",

            "Finding": "Cation totals and normalization can be reproduced.",

            "Limitation": "Numerical agreement does not certify site occupancy.",

            "Required_Evidence": "Published allocation scheme, site capacities and charge-balance checks."
        },

        {
            "Issue": "Natural sample provenance",

            "Finding": "The workbook identifies the samples as Xiao et al. (2011).",

            "Limitation": "Publication attribution alone does not independently verify sample-level source records.",

            "Required_Evidence": "Original Xiao et al. analytical tables and sample crosswalk."
        }

    ]

    add_sheet(
        "Scientific_Limitations",
        limitations
    )

    # ---------------------------------------------------------
    # Overall summary
    # ---------------------------------------------------------

    total_factor_checks = sum(
        x["Factor_Checks"]
        for x in summary_results
    )

    total_factor_passes = sum(
        x["Factor_Passes"]
        for x in summary_results
    )

    total_failures = sum(
        x["Total_Failures"]
        for x in summary_results
    )

    overall = [

        ["Metric", "Value"],

        ["Samples", len(samples)],

        ["Factor comparisons", total_factor_checks],

        ["Factor PASS", total_factor_passes],

        ["Factor FAIL", total_factor_checks - total_factor_passes],

        ["Integrated failures", total_failures],

        ["Samples fully passing",
         sum(
             x["Independent_Method_Status"] == "PASS"
             for x in summary_results
         )],

        ["Samples requiring review",
         sum(
             x["Independent_Method_Status"] == "REVIEW"
             for x in summary_results
         )],

        ["Overall arithmetic status",
         "PASS" if total_failures == 0 else "REVIEW"],

        ["Scientific certification",
         "NOT ESTABLISHED"],

        ["ML eligibility",
         "Not automatically approved by this audit"]

    ]

    ws_overall = wb.create_sheet(
        "Overall_Summary",
        1
    )

    for row in overall:
        ws_overall.append(row)

    # ---------------------------------------------------------
    # Formatting
    # ---------------------------------------------------------

    header_fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    header_font = Font(
        color="FFFFFF",
        bold=True
    )

    thin = Side(
        style="thin",
        color="D9E2F3"
    )

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

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

        for col_idx in range(
            1,
            ws.max_column + 1
        ):

            maximum = 0

            for row_idx in range(
                1,
                min(ws.max_row, 200) + 1
            ):

                value = ws.cell(
                    row_idx,
                    col_idx
                ).value

                if value is not None:

                    maximum = max(
                        maximum,
                        len(str(value))
                    )

            ws.column_dimensions[
                get_column_letter(col_idx)
            ].width = min(
                max(maximum + 2, 13),
                48
            )

        ws.row_dimensions[1].height = 32

    wb.save(
        OUTPUT_FILE
    )

    # =============================================================================
    # I. FINAL CONSOLE SUMMARY
    # =============================================================================

    logger.info("=" * 79)

    logger.info(
        "STEP 06.3 COMPLETED"
    )

    logger.info(
        "Samples: %s",
        len(samples)
    )

    logger.info(
        "Factor comparisons: %s",
        total_factor_checks
    )

    logger.info(
        "Factor PASS: %s",
        total_factor_passes
    )

    logger.info(
        "Integrated failures: %s",
        total_failures
    )

    logger.info(
        "Overall arithmetic status: %s",
        "PASS" if total_failures == 0 else "REVIEW"
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

    logger.info("=" * 79)


# =============================================================================
# J. EXECUTION
# =============================================================================

if __name__ == "__main__":

    try:

        main()

    except Exception:

        logger.exception(
            "STEP 06.3 FAILED"
        )

        raise
