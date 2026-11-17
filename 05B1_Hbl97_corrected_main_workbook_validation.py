# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 15:22:27 2026

@author: hussa
"""


# -*- coding: utf-8 -*-
"""
STEP 05B.1 — CORRECTED Hbl97 MAIN WORKBOOK VALIDATION

Gerya et al. (1997)

Purpose:
1. Inspect all 1121 worksheet rows.
2. Identify actual samples using sample IDs and oxide inputs.
3. Exclude blank and invalid composition rows from numerical metrics.
4. Independently reproduce published T-P output equations.
5. Compare Python results with cached Excel outputs.
6. Calculate MAE, RMSE and maximum absolute error.
7. Generate an auditable Excel report.

Original XLS file is never modified.

IMPORTANT:
This validates the final T-P equations using the workbook's
stored intermediate chemical variables. It does not yet
independently validate the entire chemical normalization.
"""

from pathlib import Path
from datetime import datetime
import logging
import math
import hashlib

import numpy as np
import pandas as pd
import xlrd
import openpyxl

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 1. PROJECT SETTINGS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent

INPUT_DIR = PROJECT_DIR / "Geothermobarometers Calibrations"

SOURCE_FILE = INPUT_DIR / "Hbl P-T (1997).xls"

OUTPUT_DIR = (
    PROJECT_DIR /
    "05B1_Hbl97_corrected_main_validation"
)

LOG_DIR = OUTPUT_DIR / "logs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

REPORT_PATH = (
    OUTPUT_DIR /
    "Step05B1_Corrected_Main_Workbook_Validation.xlsx"
)

LOG_PATH = (
    LOG_DIR /
    "step05B1_execution_log.txt"
)

FONT_NAME = "Times New Roman"

ABS_TOL = 1e-7
REL_TOL = 1e-9


# ============================================================
# 2. LOGGING
# ============================================================

logging.basicConfig(
    filename=str(LOG_PATH),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    force=True
)


def log(message):

    print(message)
    logging.info(message)


# ============================================================
# 3. GENERAL FUNCTIONS
# ============================================================

def is_number(value):

    return (
        isinstance(value, (int, float, np.integer, np.floating))
        and np.isfinite(value)
    )


def safe_log(value):

    if not is_number(value) or value <= 0:

        raise ValueError(
            "Invalid logarithm argument: {}".format(value)
        )

    return math.log(value)


def compare_values(python_value, excel_value):

    difference = abs(
        python_value - excel_value
    )

    passed = np.isclose(
        python_value,
        excel_value,
        atol=ABS_TOL,
        rtol=REL_TOL
    )

    return difference, passed


def get_hash(path):

    digest = hashlib.sha256()

    with open(path, "rb") as file:

        for block in iter(
            lambda: file.read(1024 * 1024), b""
        ):

            digest.update(block)

    return digest.hexdigest()


def calculate_metrics(errors):

    if len(errors) == 0:

        return {
            "N": 0,
            "MAE": np.nan,
            "RMSE": np.nan,
            "Maximum_Absolute_Error": np.nan
        }

    errors = np.asarray(errors, dtype=float)

    return {

        "N": len(errors),

        "MAE": float(
            np.mean(np.abs(errors))
        ),

        "RMSE": float(
            np.sqrt(np.mean(errors ** 2))
        ),

        "Maximum_Absolute_Error": float(
            np.max(np.abs(errors))
        )
    }


# ============================================================
# 4. READ SOURCE WORKBOOK
# ============================================================

def load_source():

    if not SOURCE_FILE.exists():

        raise FileNotFoundError(
            "Source workbook not found: {}".format(
                SOURCE_FILE
            )
        )

    log("Opening original XLS workbook read-only.")

    workbook = xlrd.open_workbook(
        str(SOURCE_FILE),
        on_demand=False
    )

    sheet = workbook.sheet_by_name(
        "Hbl P-T (1997)"
    )

    log("Worksheet: {}".format(sheet.name))
    log("Total worksheet rows: {}".format(sheet.nrows))
    log("Total worksheet columns: {}".format(sheet.ncols))

    return workbook, sheet


# ============================================================
# 5. VALIDATE SAMPLE COMPOSITION
# ============================================================

def inspect_row(sheet, row_index):

    excel_row = row_index + 1

    # Excel columns:
    #
    # A = reference
    # B = sample ID
    # D = input temperature
    # E = input pressure
    # F = calculated temperature
    # G = calculated pressure
    # J:R = nine oxide inputs

    reference = sheet.cell_value(
        row_index, 0
    )

    sample_id = sheet.cell_value(
        row_index, 1
    )

    input_T = sheet.cell_value(
        row_index, 3
    )

    input_P = sheet.cell_value(
        row_index, 4
    )

    excel_T = sheet.cell_value(
        row_index, 5
    )

    excel_P = sheet.cell_value(
        row_index, 6
    )

    oxide_values = [
        sheet.cell_value(row_index, col)
        for col in range(9, 18)
    ]

    # Intermediate variables used in equations.

    Si = sheet.cell_value(row_index, 59)  # BH
    Al = sheet.cell_value(row_index, 61)  # BJ
    Fe = sheet.cell_value(row_index, 62)  # BK
    Mg = sheet.cell_value(row_index, 65)  # BN

    # Check sample identification.

    if not str(sample_id).strip():

        return {

            "Excel_Row": excel_row,

            "Sample_ID": "",

            "Reference": reference,

            "Validation_Status": "NOT EVALUATED",

            "Reason": "No sample identifier",

        }

    # Check nine original oxide inputs.

    if not all(
        is_number(value)
        for value in oxide_values
    ):

        return {

            "Excel_Row": excel_row,

            "Sample_ID": sample_id,

            "Reference": reference,

            "Validation_Status": "NOT EVALUATED",

            "Reason": "Incomplete oxide composition",

        }

    # Check intermediate chemical variables.

    if not all(
        is_number(value)
        for value in [Si, Al, Fe, Mg]
    ):

        return {

            "Excel_Row": excel_row,

            "Sample_ID": sample_id,

            "Reference": reference,

            "Validation_Status": "NOT EVALUATED",

            "Reason": "Missing intermediate chemical variables",

        }

    # Check required output values.

    if not all(
        is_number(value)
        for value in [input_T, input_P, excel_T, excel_P]
    ):

        return {

            "Excel_Row": excel_row,

            "Sample_ID": sample_id,

            "Reference": reference,

            "Validation_Status": "NOT EVALUATED",

            "Reason": "Invalid input or cached Excel output",

        }

    return {

        "Excel_Row": excel_row,

        "Sample_ID": sample_id,

        "Reference": reference,

        "Input_T_C": input_T,

        "Input_P_kbar": input_P,

        "Excel_T_C": excel_T,

        "Excel_P_kbar": excel_P,

        "Oxide_Inputs": oxide_values,

        "Si_Hbl": Si,

        "Al_Hbl": Al,

        "Fe_Hbl": Fe,

        "Mg_Hbl": Mg,

        "Validation_Status": "ELIGIBLE",

        "Reason": "Valid sample composition"

    }


# ============================================================
# 6. INDEPENDENT TEMPERATURE AND PRESSURE EQUATIONS
# ============================================================

def calculate_published_equations(record):

    Si = record["Si_Hbl"]

    Al = record["Al_Hbl"]

    Fe = record["Fe_Hbl"]

    Mg = record["Mg_Hbl"]

    input_T = record["Input_T_C"]

    input_P = record["Input_P_kbar"]

    # Published equation 16.
    #
    # T(K) = (6119 - 28.4P + 114 XMg)
    #        / (8.181 - 1.987 ln(8.489-Si))
    #
    # XMg = Mg / (Fe + Mg)

    if Fe + Mg == 0:

        raise ValueError("Fe + Mg equals zero.")

    XMg = Mg / (Fe + Mg)

    temperature_denominator = (
        8.181
        - 1.987 * safe_log(8.489 - Si)
    )

    if temperature_denominator == 0:

        raise ValueError(
            "Temperature denominator equals zero."
        )

    T_kelvin = (

        6119
        - 28.4 * input_P
        + 114 * XMg

    ) / temperature_denominator

    T_celsius = T_kelvin - 273.15

    # Published equation 17.
    #
    # P(kbar) = [2543 - 4.744T + 175XMg
    #          + 1.987T ln(Al + 1.433)] / 148.1
    #
    # Here T is Kelvin.

    input_T_kelvin = input_T + 273.15

    P_kbar = (

        2543
        - 4.744 * input_T_kelvin
        + 175 * XMg
        + 1.987 * input_T_kelvin
        * safe_log(Al + 1.433)

    ) / 148.1

    return {

        "XMg": XMg,

        "Python_T_C": T_celsius,

        "Python_P_kbar": P_kbar

    }


# ============================================================
# 7. PROCESS ALL WORKSHEET ROWS
# ============================================================

def run_validation(sheet):

    log("=" * 65)
    log("INSPECTING ALL WORKSHEET ROWS")
    log("=" * 65)

    audit_records = []

    eligible_records = []

    for row_index in range(9, sheet.nrows):

        record = inspect_row(
            sheet,
            row_index
        )

        if record["Validation_Status"] != "ELIGIBLE":

            audit_records.append(record)

            continue

        try:

            calculated = calculate_published_equations(
                record
            )

            t_difference, t_pass = compare_values(
                calculated["Python_T_C"],
                record["Excel_T_C"]
            )

            p_difference, p_pass = compare_values(
                calculated["Python_P_kbar"],
                record["Excel_P_kbar"]
            )

            overall = (
                "PASS"
                if t_pass and p_pass
                else "FAIL"
            )

            output = {

                "Excel_Row": record["Excel_Row"],

                "Sample_ID": record["Sample_ID"],

                "Reference": record["Reference"],

                "Input_T_C": record["Input_T_C"],

                "Input_P_kbar": record["Input_P_kbar"],

                "Si_Hbl": record["Si_Hbl"],

                "Al_Hbl": record["Al_Hbl"],

                "Fe_Hbl": record["Fe_Hbl"],

                "Mg_Hbl": record["Mg_Hbl"],

                "XMg": calculated["XMg"],

                "Python_T_C": calculated["Python_T_C"],

                "Excel_T_C": record["Excel_T_C"],

                "Temperature_Error": t_difference,

                "Temperature_Status": (
                    "PASS" if t_pass else "FAIL"
                ),

                "Python_P_kbar": calculated["Python_P_kbar"],

                "Excel_P_kbar": record["Excel_P_kbar"],

                "Pressure_Error": p_difference,

                "Pressure_Status": (
                    "PASS" if p_pass else "FAIL"
                ),

                "Overall_Status": overall

            }

            eligible_records.append(output)

            audit_records.append({

                "Excel_Row": record["Excel_Row"],

                "Sample_ID": record["Sample_ID"],

                "Reference": record["Reference"],

                "Validation_Status": overall,

                "Reason": "Numerical comparison completed"

            })

        except Exception as error:

            audit_records.append({

                "Excel_Row": record["Excel_Row"],

                "Sample_ID": record["Sample_ID"],

                "Reference": record["Reference"],

                "Validation_Status": "NOT EVALUATED",

                "Reason": str(error)

            })

    log("Rows scanned: {}".format(sheet.nrows - 9))

    log("Eligible samples: {}".format(len(eligible_records)))

    log(
        "PASS: {}".format(
            sum(
                x["Overall_Status"] == "PASS"
                for x in eligible_records
            )
        )
    )

    log(
        "FAIL: {}".format(
            sum(
                x["Overall_Status"] == "FAIL"
                for x in eligible_records
            )
        )
    )

    log(
        "NOT EVALUATED: {}".format(
            sum(
                x["Validation_Status"] == "NOT EVALUATED"
                for x in audit_records
            )
        )
    )

    return audit_records, eligible_records


# ============================================================
# 8. METRICS
# ============================================================

def calculate_summary(records):

    t_errors = [
        x["Python_T_C"] - x["Excel_T_C"]
        for x in records
    ]

    p_errors = [
        x["Python_P_kbar"] - x["Excel_P_kbar"]
        for x in records
    ]

    t_metrics = calculate_metrics(t_errors)

    p_metrics = calculate_metrics(p_errors)

    passed = sum(
        x["Overall_Status"] == "PASS"
        for x in records
    )

    failed = sum(
        x["Overall_Status"] == "FAIL"
        for x in records
    )

    return [

        {
            "Metric": "Eligible samples",
            "Temperature": len(records),
            "Pressure": len(records)
        },

        {
            "Metric": "PASS",
            "Temperature": sum(
                x["Temperature_Status"] == "PASS"
                for x in records
            ),
            "Pressure": sum(
                x["Pressure_Status"] == "PASS"
                for x in records
            )
        },

        {
            "Metric": "FAIL",
            "Temperature": sum(
                x["Temperature_Status"] == "FAIL"
                for x in records
            ),
            "Pressure": sum(
                x["Pressure_Status"] == "FAIL"
                for x in records
            )
        },

        {
            "Metric": "MAE",
            "Temperature": t_metrics["MAE"],
            "Pressure": p_metrics["MAE"]
        },

        {
            "Metric": "RMSE",
            "Temperature": t_metrics["RMSE"],
            "Pressure": p_metrics["RMSE"]
        },

        {
            "Metric": "Maximum absolute error",
            "Temperature": t_metrics["Maximum_Absolute_Error"],
            "Pressure": p_metrics["Maximum_Absolute_Error"]
        },

        {
            "Metric": "Overall sample status",
            "Temperature": "",
            "Pressure": (
                "PASS" if passed == len(records)
                else "REQUIRES INVESTIGATION"
            )
        }

    ]


# ============================================================
# 9. CREATE EXCEL REPORT
# ============================================================

def write_report(
    audit_records,
    eligible_records,
    summary_records,
    source_hash
):

    wb = openpyxl.Workbook()

    wb.remove(wb.active)

    def add_sheet(name, title, records):

        ws = wb.create_sheet(name)

        ws.sheet_view.showGridLines = False

        ws.append([title])

        ws["A1"].font = Font(
            name=FONT_NAME,
            size=16,
            bold=True,
            color="FFFFFF"
        )

        ws["A1"].fill = PatternFill(
            "solid",
            fgColor="17365D"
        )

        ws.row_dimensions[1].height = 30

        if not records:

            records = [
                {"Information": "No records"}
            ]

        headers = list(records[0].keys())

        ws.append(headers)

        for cell in ws[2]:

            cell.font = Font(
                name=FONT_NAME,
                size=10,
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="305496"
            )

            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        for record in records:

            ws.append([
                record.get(header, "")
                for header in headers
            ])

        for row in ws.iter_rows(min_row=3):

            for cell in row:

                cell.font = Font(
                    name=FONT_NAME,
                    size=10
                )

                cell.alignment = Alignment(
                    vertical="top",
                    wrap_text=True
                )

        for column in range(1, ws.max_column + 1):

            letter = get_column_letter(column)

            max_length = 0

            for row in range(2, ws.max_row + 1):

                value = ws.cell(row, column).value

                if value is not None:

                    max_length = max(
                        max_length,
                        min(len(str(value)), 50)
                    )

            ws.column_dimensions[letter].width = min(
                max(max_length + 3, 14),
                45
            )

        ws.freeze_panes = "A3"

        if ws.max_row >= 3:

            ws.auto_filter.ref = ws.dimensions

        return ws

    # Summary sheet

    summary = add_sheet(
        "Validation_Summary",
        "STEP 05B.1 — VALIDATION SUMMARY",
        summary_records
    )

    summary.append([
        "Source workbook SHA256",
        source_hash,
        ""
    ])

    summary.append([
        "Validation timestamp",
        datetime.now().isoformat(timespec="seconds"),
        ""
    ])

    summary.append([
        "Numerical tolerance",
        ABS_TOL,
        ""
    ])

    # Eligible samples

    add_sheet(
        "Eligible_Samples",
        "ELIGIBLE SAMPLE NUMERICAL VALIDATION",
        eligible_records
    )

    # Full row audit

    add_sheet(
        "Complete_Row_Audit",
        "COMPLETE WORKSHEET ROW AUDIT",
        audit_records
    )

    # Method documentation

    notes = [

        {
            "Item": "Published temperature equation",
            "Description": (
                "T(K) = (6119 - 28.4P + 114XMg) / "
                "(8.181 - 1.987 ln(8.489-Si))"
            )
        },

        {
            "Item": "Published pressure equation",
            "Description": (
                "P(kbar) = [2543 - 4.744T + 175XMg "
                "+ 1.987T ln(Al+1.433)] / 148.1"
            )
        },

        {
            "Item": "Mg number",
            "Description": "XMg = Mg/(Fe+Mg)"
        },

        {
            "Item": "Temperature comparison",
            "Description": (
                "Python temperature is compared with cached Excel F."
            )
        },

        {
            "Item": "Pressure comparison",
            "Description": (
                "Python pressure is compared with cached Excel G."
            )
        },

        {
            "Item": "Important limitation",
            "Description": (
                "Intermediate variables are taken from the original "
                "workbook. Full oxide-to-cation normalization and "
                "iterative convergence are not independently validated."
            )
        },

        {
            "Item": "Sample selection",
            "Description": (
                "Requires a sample identifier, all nine oxide inputs, "
                "valid intermediate variables and numerical outputs."
            )
        }

    ]

    add_sheet(
        "Method_and_Limitations",
        "METHOD AND SCIENTIFIC LIMITATIONS",
        notes
    )

    wb.save(REPORT_PATH)


# ============================================================
# 10. MAIN
# ============================================================

def main():

    log("=" * 70)

    log("STEP 05B.1 — CORRECTED Hbl97 VALIDATION")

    log("=" * 70)

    workbook, sheet = load_source()

    audit_records, eligible_records = run_validation(
        sheet
    )

    summary_records = calculate_summary(
        eligible_records
    )

    source_hash = get_hash(SOURCE_FILE)

    write_report(
        audit_records,
        eligible_records,
        summary_records,
        source_hash
    )

    log("=" * 70)

    log("STEP 05B.1 COMPLETED")

    log("Worksheet rows: {}".format(sheet.nrows))

    log("Eligible samples: {}".format(len(eligible_records)))

    log("Report: {}".format(REPORT_PATH))

    log("Log: {}".format(LOG_PATH))

    log("=" * 70)

    workbook.release_resources()


if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logging.exception(
            "STEP 05B.1 FAILED: %s",
            error
        )

        print("\nSTEP 05B.1 FAILED")
        print(error)

        raise
