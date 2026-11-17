# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 15:35:20 2026

@author: hussa
"""


# -*- coding: utf-8 -*-
"""
STEP 05C.1–05C.2
Hbl97 INDEPENDENT AMPHIBOLE CHEMICAL RECONSTRUCTION

Source:
Gerya et al. (1997), printed pp. 527–529.
Original Hbl P-T (1997).xls workbook.

Tasks:
1. Extract original oxide compositions.
2. Reconstruct oxide-to-cation calculations.
3. Reconstruct initial oxygen normalization.
4. Reconstruct stoichiometric factors f1-f10.
5. Reconstruct Fe3+ allocation and revised cations.
6. Reconstruct normalization to 13 cations.
7. Compare every available intermediate with Excel.
8. Produce a detailed audit report.

IMPORTANT:
This is an independent Python implementation of the
observed workbook calculation sequence. Agreement with
Excel establishes computational reproduction, not
independent experimental validation of the calibration.

Original XLS workbook remains unchanged.
"""

from pathlib import Path
from datetime import datetime
import logging
import hashlib
import math
import sys

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
    "05C1_Hbl97_independent_chemical_reconstruction"
)

LOG_DIR = OUTPUT_DIR / "logs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

REPORT_PATH = (
    OUTPUT_DIR /
    "Step05C1_Hbl97_Independent_Chemical_Reconstruction.xlsx"
)

LOG_PATH = (
    LOG_DIR /
    "step05C1_execution_log.txt"
)

FONT_NAME = "Times New Roman"

ABS_TOL = 1e-8
REL_TOL = 1e-9

# Original workbook worksheet
SHEET_NAME = "Hbl P-T (1997)"


# ============================================================
# 2. OXIDE MOLECULAR WEIGHTS
# ============================================================

MW = {

    "SiO2": 60.0843,

    "TiO2": 79.8788,

    "Al2O3": 101.9612,

    "FeO": 71.8464,

    "MnO": 70.9375,

    "MgO": 40.3044,

    "CaO": 56.0774,

    "Na2O": 61.979,

    "K2O": 94.196

}

# Oxygen contributions per oxide molecule.
OXYGENS = {

    "SiO2": 2,

    "TiO2": 2,

    "Al2O3": 3,

    "FeO": 1,

    "MnO": 1,

    "MgO": 1,

    "CaO": 1,

    "Na2O": 1,

    "K2O": 1

}

# Cation multiplicities.
CATIONS = {

    "SiO2": 1,

    "TiO2": 1,

    "Al2O3": 2,

    "FeO": 1,

    "MnO": 1,

    "MgO": 1,

    "CaO": 1,

    "Na2O": 2,

    "K2O": 2

}


# ============================================================
# 3. LOGGING
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
# 4. GENERAL FUNCTIONS
# ============================================================

def sha256_file(path):

    digest = hashlib.sha256()

    with open(path, "rb") as file:

        for block in iter(
            lambda: file.read(1024 * 1024),
            b""
        ):

            digest.update(block)

    return digest.hexdigest()


def is_number(value):

    return (

        isinstance(
            value,
            (int, float, np.integer, np.floating)
        )

        and np.isfinite(value)

    )


def safe_divide(a, b):

    if not is_number(a) or not is_number(b):

        raise ValueError("Non-numerical division.")

    if abs(b) < 1e-14:

        raise ValueError("Division by zero.")

    return a / b


def safe_log(value):

    if not is_number(value) or value <= 0:

        raise ValueError(
            "Invalid logarithm argument: {}".format(value)
        )

    return math.log(value)


def compare_value(calculated, original):

    if not is_number(calculated) or not is_number(original):

        return np.nan, "NOT EVALUATED"

    difference = abs(calculated - original)

    passed = np.isclose(

        calculated,

        original,

        atol=ABS_TOL,

        rtol=REL_TOL

    )

    return (

        difference,

        "PASS" if passed else "FAIL"

    )


# ============================================================
# 5. SOURCE WORKBOOK
# ============================================================

def load_source():

    if not SOURCE_FILE.exists():

        raise FileNotFoundError(
            "Original calibration workbook not found:\n{}".format(
                SOURCE_FILE
            )
        )

    log("Opening original XLS workbook.")

    workbook = xlrd.open_workbook(
        str(SOURCE_FILE),
        on_demand=False
    )

    if SHEET_NAME not in workbook.sheet_names():

        raise ValueError(
            "Expected worksheet not found: {}".format(
                SHEET_NAME
            )
        )

    sheet = workbook.sheet_by_name(SHEET_NAME)

    log("Worksheet: {}".format(sheet.name))

    log("Rows: {}".format(sheet.nrows))

    log("Columns: {}".format(sheet.ncols))

    return workbook, sheet


# ============================================================
# 6. EXTRACT OXIDE INPUTS
# ============================================================

def extract_sample(sheet, row):

    # xlrd uses zero-based indexing.
    #
    # Excel:
    # B = sample identifier
    # J:R = nine oxide weight percentages
    # D:E = input T and P
    # H = Fe3+ value used in revised formula
    #
    # Row 10 in Excel = index 9 in Python.

    sample_id = sheet.cell_value(row, 1)

    reference = sheet.cell_value(row, 0)

    oxide_values = [

        sheet.cell_value(row, col)

        for col in range(9, 18)

    ]

    input_T = sheet.cell_value(row, 3)

    input_P = sheet.cell_value(row, 4)

    Fe3_input = sheet.cell_value(row, 7)

    if not str(sample_id).strip():

        return None, "Missing sample identifier"

    if not all(is_number(x) for x in oxide_values):

        return None, "Incomplete oxide composition"

    if any(x < 0 for x in oxide_values):

        return None, "Negative oxide concentration"

    if sum(oxide_values) <= 0:

        return None, "Zero oxide total"

    if not is_number(Fe3_input):

        return None, "Missing Fe3+ value"

    # Workbook instruction:
    # Negative Fe3+ values in column H must be changed to zero.

    Fe3_used = max(0.0, Fe3_input)

    oxide_names = list(MW.keys())

    oxide = dict(
        zip(oxide_names, oxide_values)
    )

    return {

        "Excel_Row": row + 1,

        "Sample_ID": str(sample_id).strip(),

        "Reference": str(reference),

        "Oxides": oxide,

        "Input_T_C": input_T,

        "Input_P_kbar": input_P,

        "Fe3_Input": Fe3_input,

        "Fe3_Used": Fe3_used

    }, None


# ============================================================
# 7. INITIAL OXYGEN NORMALIZATION
# ============================================================

def calculate_initial_cations(sample):

    oxide = sample["Oxides"]

    # Workbook equation:
    #
    # T = 23 / sum(oxide_moles * oxygen_number)

    oxygen_sum = 0.0

    for name in MW:

        oxygen_sum += (

            oxide[name] / MW[name]

        ) * OXYGENS[name]

    ratio = safe_divide(23.0, oxygen_sum)

    result = {

        "T_Ratio": ratio

    }

    # Initial cations U:AC.

    initial = {}

    for name in MW:

        initial[name] = (

            oxide[name]

            / MW[name]

            * CATIONS[name]

            * ratio

        )

    result["Si"] = initial["SiO2"]

    result["Ti"] = initial["TiO2"]

    result["Al"] = initial["Al2O3"]

    result["Fe_total"] = initial["FeO"]

    result["Mn"] = initial["MnO"]

    result["Mg"] = initial["MgO"]

    result["Ca"] = initial["CaO"]

    result["Na"] = initial["Na2O"]

    result["K"] = initial["K2O"]

    result["Initial_Cation_Sum"] = sum(
        initial.values()
    )

    return result


# ============================================================
# 8. STOICHIOMETRIC CONSTRAINTS
# ============================================================

def calculate_stoichiometry(initial):

    Si = initial["Si"]

    Ti = initial["Ti"]

    Al = initial["Al"]

    Fe = initial["Fe_total"]

    Mn = initial["Mn"]

    Mg = initial["Mg"]

    Ca = initial["Ca"]

    Na = initial["Na"]

    K = initial["K"]

    total = initial["Initial_Cation_Sum"]

    f = {}

    # f1-f5

    f["f1"] = safe_divide(16.0, total)

    f["f2"] = safe_divide(8.0, Si)

    f["f3"] = safe_divide(
        15.0,
        total - Na - K
    )

    f["f4"] = safe_divide(2.0, Ca)

    f["f5"] = 1.0

    minimum = min(f.values())

    # f6-f10

    f["f6"] = safe_divide(
        8.0,
        Si + Al
    )

    f["f7"] = safe_divide(
        15.0,
        total - K
    )

    f["f8"] = safe_divide(
        13.0,
        total - Ca - Mn - Na - K
    )

    f["f9"] = safe_divide(
        36.0,
        46.0 - Si - Ti - Al
    )

    f["f10"] = safe_divide(
        46.0,
        Fe + 46.0
    )

    maximum = max(
        f[name]
        for name in [
            "f6", "f7", "f8", "f9", "f10"
        ]
    )

    # Workbook formula:
    # fav = 0.5 * min(f1:f5) + 0.5 * max(f6:f10)

    fav = 0.5 * minimum + 0.5 * maximum

    # Workbook formula:
    # Fe3+ = 46 * (1 - fav)

    Fe3_calculated = 46.0 * (1.0 - fav)

    return {

        "Factors": f,

        "Minimum_f1_f5": minimum,

        "Maximum_f6_f10": maximum,

        "fav": fav,

        "Fe3_Calculated": Fe3_calculated

    }


# ============================================================
# 9. REVISED CATION FORMULA
# ============================================================

def calculate_revised_cations(initial, sample):

    Si = initial["Si"]

    Ti = initial["Ti"]

    Al = initial["Al"]

    Fe = initial["Fe_total"]

    Mn = initial["Mn"]

    Mg = initial["Mg"]

    Ca = initial["Ca"]

    Na = initial["Na"]

    K = initial["K"]

    fav = initial["fav"]

    Fe3 = sample["Fe3_Used"]

    revised = {

        "Si": Si * fav,

        "Ti": Ti * fav,

        "Al": Al * fav,

        "Fe2": Fe * fav - Fe3,

        "Fe3": Fe3,

        "Mn": Mn * fav,

        "Mg": Mg * fav,

        "Ca": Ca * fav,

        "Na": Na * fav,

        "K": K * fav

    }

    revised["New_Sum"] = sum(
        revised.values()
    )

    return revised


# ============================================================
# 10. NORMALIZATION TO 13 CATIONS
# ============================================================

def normalize_to_13(revised):

    # Workbook BF:
    # Sum(Si, Ti, Al, Fe2+, Fe3+, Mn, Mg)

    sum_mn = sum(

        revised[name]

        for name in [
            "Si", "Ti", "Al", "Fe2", "Fe3", "Mn", "Mg"
        ]

    )

    ratio_13 = safe_divide(
        13.0,
        sum_mn
    )

    normalized = {

        name: value * ratio_13

        for name, value in revised.items()

        if name != "New_Sum"

    }

    return {

        "Sum_Mn": sum_mn,

        "Ratio_13": ratio_13,

        "Normalized": normalized

    }


# ============================================================
# 11. FINAL TEMPERATURE AND PRESSURE
# ============================================================

def calculate_Hbl97(normalized, sample):

    Si = normalized["Si"]

    Al = normalized["Al"]

    Fe2 = normalized["Fe2"]

    Mg = normalized["Mg"]

    input_T = sample["Input_T_C"]

    input_P = sample["Input_P_kbar"]

    XMg = safe_divide(
        Mg,
        Mg + Fe2
    )

    # Workbook BO:
    #
    # T = (6119 - 28.4P + 114XMg)
    #     / (8.181 - 1.987 ln(8.489-Si))
    #     - 273.15

    T_C = (

        safe_divide(

            6119 - 28.4 * input_P + 114 * XMg,

            8.181 - 1.987 * safe_log(8.489 - Si)

        )

        - 273.15

    )

    # Workbook BP:
    #
    # P = [2543 - 4.744T + 175XMg
    #      + 1.987T ln(Al + 1.433)] / 148.1
    #
    # T is Kelvin.

    T_K = input_T + 273.15

    P_kbar = (

        2543

        - 4.744 * T_K

        + 175 * XMg

        + 1.987 * T_K * safe_log(Al + 1.433)

    ) / 148.1

    return {

        "XMg": XMg,

        "T_C": T_C,

        "P_kbar": P_kbar

    }


# ============================================================
# 12. EXCEL REFERENCE VALUES
# ============================================================

def get_excel_reference(sheet, row):

    # Exact zero-based column mapping.
    #
    # T:20
    # U:21 ... AC:29
    # AD:30
    # AE:31 ... AI:35
    # AJ:36
    # AK:37 ... AO:41
    # AP:42
    # AQ:43
    # AR:44
    # AT:46 ... BD:55
    # BF:57
    # BG:58
    # BH:59 ... BN:65
    # BO:66
    # BP:67

    return {

        "T_Ratio": sheet.cell_value(row, 19),

        "Si": sheet.cell_value(row, 20),

        "Ti": sheet.cell_value(row, 21),

        "Al": sheet.cell_value(row, 22),

        "Fe_total": sheet.cell_value(row, 23),

        "Mn": sheet.cell_value(row, 24),

        "Mg": sheet.cell_value(row, 25),

        "Ca": sheet.cell_value(row, 26),

        "Na": sheet.cell_value(row, 27),

        "K": sheet.cell_value(row, 28),

        "Initial_Cation_Sum": sheet.cell_value(row, 29),

        "f1": sheet.cell_value(row, 30),

        "f2": sheet.cell_value(row, 31),

        "f3": sheet.cell_value(row, 32),

        "f4": sheet.cell_value(row, 33),

        "f5": sheet.cell_value(row, 34),

        "f_min": sheet.cell_value(row, 35),

        "f6": sheet.cell_value(row, 36),

        "f7": sheet.cell_value(row, 37),

        "f8": sheet.cell_value(row, 38),

        "f9": sheet.cell_value(row, 39),

        "f10": sheet.cell_value(row, 40),

        "f_max": sheet.cell_value(row, 41),

        "fav": sheet.cell_value(row, 42),

        "Fe3_Calculated": sheet.cell_value(row, 43),

        "Revised_Si": sheet.cell_value(row, 45),

        "Revised_Ti": sheet.cell_value(row, 46),

        "Revised_Al": sheet.cell_value(row, 47),

        "Revised_Fe2": sheet.cell_value(row, 48),

        "Revised_Fe3": sheet.cell_value(row, 49),

        "Revised_Mn": sheet.cell_value(row, 50),

        "Revised_Mg": sheet.cell_value(row, 51),

        "Revised_Ca": sheet.cell_value(row, 52),

        "Revised_Na": sheet.cell_value(row, 53),

        "Revised_K": sheet.cell_value(row, 54),

        "Revised_Sum": sheet.cell_value(row, 55),

        "Sum_Mn": sheet.cell_value(row, 57),

        "Ratio_13": sheet.cell_value(row, 58),

        "13C_Si": sheet.cell_value(row, 59),

        "13C_Ti": sheet.cell_value(row, 60),

        "13C_Al": sheet.cell_value(row, 61),

        "13C_Fe2": sheet.cell_value(row, 62),

        "13C_Fe3": sheet.cell_value(row, 63),

        "13C_Mn": sheet.cell_value(row, 64),

        "13C_Mg": sheet.cell_value(row, 65),

        "T_Excel": sheet.cell_value(row, 66),

        "P_Excel": sheet.cell_value(row, 67)

    }


# ============================================================
# 13. BUILD DETAILED COMPARISON RECORDS
# ============================================================

def build_comparisons(sample, initial, stoich, revised, norm, final, excel):

    comparisons = []

    def add(name, calculated, reference):

        difference, status = compare_value(
            calculated,
            reference
        )

        comparisons.append({

            "Excel_Row": sample["Excel_Row"],

            "Sample_ID": sample["Sample_ID"],

            "Variable": name,

            "Python_Value": calculated,

            "Excel_Value": reference,

            "Absolute_Difference": difference,

            "Status": status

        })

    add(
        "Oxygen_Ratio",
        initial["T_Ratio"],
        excel["T_Ratio"]
    )

    for name in [
        "Si", "Ti", "Al", "Fe_total",
        "Mn", "Mg", "Ca", "Na", "K"
    ]:

        add(
            "Initial_" + name,
            initial[name],
            excel[name]
        )

    add(
        "Initial_Cation_Sum",
        initial["Initial_Cation_Sum"],
        excel["Initial_Cation_Sum"]
    )

    for name, value in stoich["Factors"].items():

        add(
            name,
            value,
            excel[name]
        )

    add(
        "Minimum_f1_f5",
        stoich["Minimum_f1_f5"],
        excel["f_min"]
    )

    add(
        "Maximum_f6_f10",
        stoich["Maximum_f6_f10"],
        excel["f_max"]
    )

    add(
        "fav",
        stoich["fav"],
        excel["fav"]
    )

    add(
        "Calculated_Fe3",
        stoich["Fe3_Calculated"],
        excel["Fe3_Calculated"]
    )

    revised_map = {

        "Revised_Si": "Si",

        "Revised_Ti": "Ti",

        "Revised_Al": "Al",

        "Revised_Fe2": "Fe2",

        "Revised_Fe3": "Fe3",

        "Revised_Mn": "Mn",

        "Revised_Mg": "Mg",

        "Revised_Ca": "Ca",

        "Revised_Na": "Na",

        "Revised_K": "K"

    }

    for excel_name, python_name in revised_map.items():

        add(
            excel_name,
            revised[python_name],
            excel[excel_name]
        )

    add(
        "Revised_Sum",
        revised["New_Sum"],
        excel["Revised_Sum"]
    )

    add(
        "Sum_Mn",
        norm["Sum_Mn"],
        excel["Sum_Mn"]
    )

    add(
        "Ratio_13",
        norm["Ratio_13"],
        excel["Ratio_13"]
    )

    normalized_map = {

        "13C_Si": "Si",

        "13C_Ti": "Ti",

        "13C_Al": "Al",

        "13C_Fe2": "Fe2",

        "13C_Fe3": "Fe3",

        "13C_Mn": "Mn",

        "13C_Mg": "Mg"

    }

    for excel_name, python_name in normalized_map.items():

        add(
            excel_name,
            norm["Normalized"][python_name],
            excel[excel_name]
        )

    add(
        "Final_Temperature",
        final["T_C"],
        excel["T_Excel"]
    )

    add(
        "Final_Pressure",
        final["P_kbar"],
        excel["P_Excel"]
    )

    return comparisons


# ============================================================
# 14. MAIN SAMPLE PROCESSING
# ============================================================

def process_all(sheet):

    sample_records = []

    comparison_records = []

    rejected_records = []

    log("=" * 70)

    log("INDEPENDENT CHEMICAL RECONSTRUCTION")

    log("=" * 70)

    for row in range(9, sheet.nrows):

        sample, reason = extract_sample(
            sheet,
            row
        )

        if sample is None:

            rejected_records.append({

                "Excel_Row": row + 1,

                "Reason": reason

            })

            continue

        try:

            initial = calculate_initial_cations(
                sample
            )

            stoich = calculate_stoichiometry(
                initial
            )

            revised = calculate_revised_cations(
                {

                    **initial,

                    "fav": stoich["fav"]

                },

                sample

            )

            norm = normalize_to_13(
                revised
            )

            final = calculate_Hbl97(
                norm["Normalized"],
                sample
            )

            excel = get_excel_reference(
                sheet,
                row
            )

            comparisons = build_comparisons(

                sample,

                initial,

                stoich,

                revised,

                norm,

                final,

                excel

            )

            comparison_records.extend(comparisons)

            sample_records.append({

                "Excel_Row": sample["Excel_Row"],

                "Sample_ID": sample["Sample_ID"],

                "Reference": sample["Reference"],

                "Input_T_C": sample["Input_T_C"],

                "Input_P_kbar": sample["Input_P_kbar"],

                "Fe3_Input": sample["Fe3_Input"],

                "Fe3_Used": sample["Fe3_Used"],

                "T_Ratio_Python": initial["T_Ratio"],

                "Initial_Si": initial["Si"],

                "Initial_Al": initial["Al"],

                "Initial_Fe_Total": initial["Fe_total"],

                "Initial_Mg": initial["Mg"],

                "Initial_Cation_Sum": initial["Initial_Cation_Sum"],

                "Minimum_f1_f5": stoich["Minimum_f1_f5"],

                "Maximum_f6_f10": stoich["Maximum_f6_f10"],

                "fav": stoich["fav"],

                "Calculated_Fe3": stoich["Fe3_Calculated"],

                "Revised_Si": revised["Si"],

                "Revised_Al": revised["Al"],

                "Revised_Fe2": revised["Fe2"],

                "Revised_Fe3": revised["Fe3"],

                "Revised_Mg": revised["Mg"],

                "Revised_Cation_Sum": revised["New_Sum"],

                "Sum_Mn": norm["Sum_Mn"],

                "Ratio_13": norm["Ratio_13"],

                "Si_13": norm["Normalized"]["Si"],

                "Al_13": norm["Normalized"]["Al"],

                "Fe2_13": norm["Normalized"]["Fe2"],

                "Fe3_13": norm["Normalized"]["Fe3"],

                "Mg_13": norm["Normalized"]["Mg"],

                "XMg": final["XMg"],

                "Python_T_C": final["T_C"],

                "Excel_T_C": excel["T_Excel"],

                "Python_P_kbar": final["P_kbar"],

                "Excel_P_kbar": excel["P_Excel"]

            })

        except Exception as error:

            rejected_records.append({

                "Excel_Row": row + 1,

                "Sample_ID": sample["Sample_ID"],

                "Reason": str(error)

            })

    log("Eligible samples: {}".format(len(sample_records)))

    log("Rejected or incomplete rows: {}".format(len(rejected_records)))

    return (
        sample_records,
        comparison_records,
        rejected_records
    )


# ============================================================
# 15. SUMMARY STATISTICS
# ============================================================

def build_summary(samples, comparisons, rejected):

    df = pd.DataFrame(comparisons)

    summary = []

    if not df.empty:

        for variable, group in df.groupby("Variable"):

            evaluated = group[
                group["Status"].isin(["PASS", "FAIL"])
            ]

            errors = evaluated[
                "Python_Value"
            ].astype(float) - evaluated[
                "Excel_Value"
            ].astype(float)

            summary.append({

                "Variable": variable,

                "Records": len(group),

                "Evaluated": len(evaluated),

                "Passed": int(
                    (group["Status"] == "PASS").sum()
                ),

                "Failed": int(
                    (group["Status"] == "FAIL").sum()
                ),

                "MAE": (
                    float(np.mean(np.abs(errors)))
                    if len(errors) else np.nan
                ),

                "RMSE": (
                    float(np.sqrt(np.mean(errors ** 2)))
                    if len(errors) else np.nan
                ),

                "Maximum_Absolute_Error": (
                    float(np.max(np.abs(errors)))
                    if len(errors) else np.nan
                )

            })

    metadata = [

        {
            "Metric": "Eligible samples",
            "Value": len(samples)
        },

        {
            "Metric": "Rejected or incomplete rows",
            "Value": len(rejected)
        },

        {
            "Metric": "Total intermediate comparisons",
            "Value": len(comparisons)
        },

        {
            "Metric": "Passing comparisons",
            "Value": sum(
                x["Status"] == "PASS"
                for x in comparisons
            )
        },

        {
            "Metric": "Failed comparisons",
            "Value": sum(
                x["Status"] == "FAIL"
                for x in comparisons
            )
        },

        {
            "Metric": "Not evaluated comparisons",
            "Value": sum(
                x["Status"] == "NOT EVALUATED"
                for x in comparisons
            )
        }

    ]

    return summary, metadata


# ============================================================
# 16. REPORT FORMATTING
# ============================================================

def add_dataframe_sheet(wb, name, title, dataframe):

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

    if dataframe.empty:

        dataframe = pd.DataFrame(
            [{"Information": "No records"}]
        )

    headers = list(dataframe.columns)

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

    for record in dataframe.itertuples(
        index=False,
        name=None
    ):

        ws.append(list(record))

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

    for col in range(1, ws.max_column + 1):

        letter = get_column_letter(col)

        maximum = 0

        for row in range(2, ws.max_row + 1):

            value = ws.cell(row, col).value

            if value is not None:

                maximum = max(

                    maximum,

                    min(len(str(value)), 50)

                )

        ws.column_dimensions[letter].width = min(

            max(maximum + 3, 14),

            45

        )

    ws.freeze_panes = "A3"

    if ws.max_row >= 3:

        ws.auto_filter.ref = ws.dimensions

    return ws


# ============================================================
# 17. SAVE REPORT
# ============================================================

def save_report(
    samples,
    comparisons,
    rejected,
    summary,
    metadata,
    source_hash
):

    wb = openpyxl.Workbook()

    wb.remove(wb.active)

    add_dataframe_sheet(

        wb,

        "Validation_Summary",

        "STEP 05C.1–05C.2 — SUMMARY",

        pd.DataFrame(metadata)

    )

    add_dataframe_sheet(

        wb,

        "Sample_Chemistry",

        "INDEPENDENT AMPHIBOLE CHEMICAL RECONSTRUCTION",

        pd.DataFrame(samples)

    )

    add_dataframe_sheet(

        wb,

        "Intermediate_Audit",

        "PYTHON VERSUS ORIGINAL EXCEL INTERMEDIATE VALUES",

        pd.DataFrame(comparisons)

    )

    add_dataframe_sheet(

        wb,

        "Error_Statistics",

        "INTERMEDIATE NUMERICAL ERROR STATISTICS",

        pd.DataFrame(summary)

    )

    add_dataframe_sheet(

        wb,

        "Rejected_Rows",

        "INCOMPLETE OR INVALID ROWS",

        pd.DataFrame(rejected)

    )

    documentation = pd.DataFrame([

        {

            "Item": "Publication",

            "Description": (
                "Gerya et al. (1997), printed pages 527–529."
            )

        },

        {

            "Item": "Original workbook",

            "Description": SOURCE_FILE.name

        },

        {

            "Item": "Source SHA256",

            "Description": source_hash

        },

        {

            "Item": "Oxide inputs",

            "Description": "Excel columns J:R"

        },

        {

            "Item": "Initial normalization",

            "Description": (
                "23 oxygen basis, as implemented in the original workbook."
            )

        },

        {

            "Item": "Stoichiometric adjustment",

            "Description": (
                "Original workbook factors f1–f10, minimum, maximum "
                "and fav."
            )

        },

        {

            "Item": "Fe3+ handling",

            "Description": (
                "Original workbook instruction: negative column H "
                "values are changed to zero."
            )

        },

        {

            "Item": "13-cation normalization",

            "Description": (
                "Original workbook normalization using Si, Ti, Al, "
                "Fe2+, Fe3+, Mn and Mg."
            )

        },

        {

            "Item": "Final T-P calculation",

            "Description": (
                "Original workbook equations in columns BO and BP."
            )

        },

        {

            "Item": "Scientific limitation",

            "Description": (
                "Agreement with workbook arithmetic does not independently "
                "establish the validity of the empirical calibration, "
                "experimental uncertainty or iterative convergence."
            )

        },

        {

            "Item": "Fe3+ distinction",

            "Description": (
                "The calculated Fe3+ estimate and the Fe3+ value actually "
                "used in column H are reported separately."
            )

        }

    ])

    add_dataframe_sheet(

        wb,

        "Method_and_Limitations",

        "METHOD DOCUMENTATION AND LIMITATIONS",

        documentation

    )

    wb.save(REPORT_PATH)


# ============================================================
# 18. MAIN
# ============================================================

def main():

    log("=" * 75)

    log("STEP 05C.1–05C.2 — Hbl97 CHEMICAL RECONSTRUCTION")

    log("=" * 75)

    log("Python executable: {}".format(sys.executable))

    log("Source: {}".format(SOURCE_FILE))

    workbook, sheet = load_source()

    samples, comparisons, rejected = process_all(
        sheet
    )

    summary, metadata = build_summary(

        samples,

        comparisons,

        rejected

    )

    source_hash = sha256_file(
        SOURCE_FILE
    )

    save_report(

        samples,

        comparisons,

        rejected,

        summary,

        metadata,

        source_hash

    )

    log("=" * 75)

    log("STEP 05C.1–05C.2 COMPLETED")

    log("Eligible samples: {}".format(len(samples)))

    log("Intermediate comparisons: {}".format(len(comparisons)))

    log(
        "PASS comparisons: {}".format(
            sum(x["Status"] == "PASS" for x in comparisons)
        )
    )

    log(
        "FAIL comparisons: {}".format(
            sum(x["Status"] == "FAIL" for x in comparisons)
        )
    )

    log("Rejected/incomplete rows: {}".format(len(rejected)))

    log("Excel report: {}".format(REPORT_PATH))

    log("Execution log: {}".format(LOG_PATH))

    log("=" * 75)

    workbook.release_resources()


if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logging.exception(
            "STEP 05C.1–05C.2 FAILED: %s",
            error
        )

        print("\nSTEP 05C.1–05C.2 FAILED")

        print(str(error))

        raise
