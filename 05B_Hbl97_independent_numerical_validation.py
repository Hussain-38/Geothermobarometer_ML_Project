# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 15:07:18 2026

@author: hussa
"""

# -*- coding: utf-8 -*-
"""
STEP 05B — INDEPENDENT Hbl97 NUMERICAL VALIDATION

Purpose:
1. Independently reproduce authors' V2, V3 and V4 calculations.
2. Compare Python outputs with original Excel cached outputs.
3. Independently check final T-P equations in the main workbook.
4. Generate a traceable numerical validation report.

Original XLS workbooks are read-only.
No source workbook is modified.

IMPORTANT:
This is a numerical reproduction audit, not a complete
scientific validation of the published calibration.
"""

from pathlib import Path
from datetime import datetime
import logging
import math
import numpy as np
import xlrd
import openpyxl

from openpyxl.styles import (
    Font, PatternFill, Alignment,
    Border, Side
)
from openpyxl.utils import get_column_letter


# ============================================================
# 1. PROJECT SETTINGS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parent

INPUT_DIR = PROJECT_DIR / "Geothermobarometers Calibrations"

MAIN_FILE = INPUT_DIR / "Hbl P-T (1997).xls"

AUTHORS_FILE = (
    INPUT_DIR / "Hbl P-T (1997) authors' spreadsheet.xls"
)

OUTPUT_DIR = (
    PROJECT_DIR /
    "05B_Hbl97_independent_validation"
)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

REPORT_PATH = (
    OUTPUT_DIR /
    "Step05B_Hbl97_Independent_Validation.xlsx"
)

LOG_PATH = (
    LOG_DIR /
    "step05B_execution_log.txt"
)

FONT_NAME = "Times New Roman"

TOLERANCE = 1e-7

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

    if value <= 0:
        raise ValueError(
            "Logarithm argument must be positive."
        )

    return math.log(value)


def compare_values(python_value, excel_value):

    difference = abs(
        python_value - excel_value
    )

    passed = np.isclose(
        python_value,
        excel_value,
        rtol=1e-9,
        atol=TOLERANCE
    )

    return difference, passed


def read_cell(sheet, coordinate):

    cell = sheet[coordinate]

    return cell.value


# ============================================================
# 4. INDEPENDENT IMPLEMENTATION OF AUTHORS' EQUATIONS
# ============================================================

def calculate_hbl97_version(
    Si,
    Al_total,
    Fe3,
    total_cations,
    version
):

    # Common temperature calculation

    Xalt = (
        8 / 15.5
        - Si / total_cations
        + 0.07531
    )

    Tk = (
        4701 /
        (
            1.825
            - 1.987 * safe_log(Xalt)
        )
    )

    if version == "V4":

        Xalm = (
            (Si + Al_total - 8) / 2
            + 1.397
        )

        Xalf = (
            Xalm /
            (
                Xalm
                + Fe3 / 2
                + 1.397
            )
        )

        Pb = (
            -400.6
            - 1551 * Xalf
            + 1.463 * Tk
            + 1.987 * Tk * safe_log(Xalm)
        ) / 0.2034

    elif version == "V3":

        Xalm = (
            (Si + Al_total - 8) / 2
            + 0.763
        )

        Xalf = (
            Xalm /
            (
                Xalm
                + Fe3 / 2
                + 0.763
            )
        )

        Pb = (
            -425
            - 1719 * Xalf
            + 2.75 * Tk
            + 1.987 * Tk * safe_log(Xalm)
        ) / 0.2

    elif version == "V2":

        Xalm = (
            (Si + Al_total - 8) / 2
            + 2.158
        )

        Xalf = np.nan

        Pb = (
            -1341
            + 0.9522 * Tk
            + 1.987 * Tk * safe_log(Xalm)
        ) / 0.2152

    else:

        raise ValueError(
            "Unknown calibration version."
        )

    Temperature_C = Tk - 273.15

    Pressure_kbar = (Pb + 1) / 1000

    return {
        "Xalt": Xalt,
        "Tk": Tk,
        "Xalm": Xalm,
        "Xalf": Xalf,
        "Pb": Pb,
        "Temperature_C": Temperature_C,
        "Pressure_kbar": Pressure_kbar
    }


# ============================================================
# 5. VALIDATE AUTHORS' SPREADSHEET
# ============================================================

def validate_authors_workbook():

    log("=" * 65)
    log("VALIDATING AUTHORS' Hbl97 SPREADSHEET")
    log("=" * 65)

    if not AUTHORS_FILE.exists():

        raise FileNotFoundError(
            "Authors' workbook not found: "
            + str(AUTHORS_FILE)
        )

    workbook = xlrd.open_workbook(
        str(AUTHORS_FILE),
        on_demand=False
    )

    values_sheet = workbook.sheet_by_name(
        "values"
    )

    calculation_sheet = workbook.sheet_by_name(
        "calculation"
    )

    # Read cached Excel values.
    # xlrd uses zero-based row and column indices.

    Si = values_sheet.cell_value(4, 1)

    Al_total = values_sheet.cell_value(25, 1)

    Fe3 = values_sheet.cell_value(9, 1)

    total_cations = values_sheet.cell_value(26, 1)

    reference_T = values_sheet.cell_value(2, 1)

    reference_P = values_sheet.cell_value(3, 1)

    log(f"Si: {Si}")
    log(f"Al total: {Al_total}")
    log(f"Fe3+: {Fe3}")
    log(f"Total cations: {total_cations}")

    input_data = {
        "Si": Si,
        "Al_total": Al_total,
        "Fe3": Fe3,
        "Total_cations": total_cations,
        "Reference_T_C": reference_T,
        "Reference_P_kbar": reference_P
    }

    # Excel output cell locations:
    #
    # V4: B12 = T, B13 = P
    # V3: B20 = T, B21 = P
    # V2: B27 = T, B28 = P

    output_cells = {
        "V4": (11, 1, 12, 1),
        "V3": (19, 1, 20, 1),
        "V2": (26, 1, 27, 1)
    }

    results = []

    for version in ["V2", "V3", "V4"]:

        python_result = calculate_hbl97_version(
            Si,
            Al_total,
            Fe3,
            total_cations,
            version
        )

        t_row, t_col, p_row, p_col = (
            output_cells[version]
        )

        excel_T = calculation_sheet.cell_value(
            t_row, t_col
        )

        excel_P = calculation_sheet.cell_value(
            p_row, p_col
        )

        t_difference, t_pass = compare_values(
            python_result["Temperature_C"],
            excel_T
        )

        p_difference, p_pass = compare_values(
            python_result["Pressure_kbar"],
            excel_P
        )

        status = (
            "PASS"
            if t_pass and p_pass
            else "FAIL"
        )

        results.append({

            "Version": version,

            "Python_T_C":
                python_result["Temperature_C"],

            "Excel_T_C": excel_T,

            "Temperature_Difference":
                t_difference,

            "Temperature_Status":
                "PASS" if t_pass else "FAIL",

            "Python_P_kbar":
                python_result["Pressure_kbar"],

            "Excel_P_kbar": excel_P,

            "Pressure_Difference":
                p_difference,

            "Pressure_Status":
                "PASS" if p_pass else "FAIL",

            "Xalt": python_result["Xalt"],

            "Xalm": python_result["Xalm"],

            "Xalf": python_result["Xalf"],

            "Overall_Status": status
        })

        log(
            f"{version}: "
            f"T = {python_result['Temperature_C']:.8f} °C | "
            f"P = {python_result['Pressure_kbar']:.8f} kbar | "
            f"{status}"
        )

    return input_data, results


# ============================================================
# 6. VALIDATE MAIN WORKBOOK FINAL EQUATIONS
# ============================================================

def validate_main_workbook():

    log("=" * 65)
    log("VALIDATING MAIN Hbl97 WORKBOOK")
    log("=" * 65)

    if not MAIN_FILE.exists():

        raise FileNotFoundError(
            "Main workbook not found: "
            + str(MAIN_FILE)
        )

    workbook = xlrd.open_workbook(
        str(MAIN_FILE),
        on_demand=False
    )

    sheet = workbook.sheet_by_name(
        "Hbl P-T (1997)"
    )

    results = []

    # Excel columns, converted to zero-based indices:
    #
    # D = 3     Input temperature
    # E = 4     Input pressure
    # F = 5     Calculated temperature
    # G = 6     Calculated pressure
    # BH = 59
    # BJ = 61
    # BK = 62
    # BN = 65

    for row in range(9, sheet.nrows):

        try:

            D = sheet.cell_value(row, 3)

            E = sheet.cell_value(row, 4)

            F = sheet.cell_value(row, 5)

            G = sheet.cell_value(row, 6)

            BH = sheet.cell_value(row, 59)

            BJ = sheet.cell_value(row, 61)

            BK = sheet.cell_value(row, 62)

            BN = sheet.cell_value(row, 65)

            values = [D, E, F, G, BH, BJ, BK, BN]

            if not all(is_number(v) for v in values):
                continue

            if BN + BK == 0:
                continue

            if 8.489 - BH <= 0:
                continue

            if BJ + 1.433 <= 0:
                continue

            # Independently reproduce Excel BO equation.

            python_T = (

                (
                    6119
                    - 28.4 * E
                    + 114 * BN / (BN + BK)
                )
                /
                (
                    8.181
                    - 1.987 * safe_log(8.489 - BH)
                )

                - 273.15
            )

            # Independently reproduce Excel BP equation.

            python_P = (

                (
                    2543
                    - 4.744 * (D + 273.15)
                    + 175 * BN / (BN + BK)
                    + 1.987 * (D + 273.15)
                    * safe_log(BJ + 1.433)
                )
                / 148.1
            )

            t_difference, t_pass = compare_values(
                python_T, F
            )

            p_difference, p_pass = compare_values(
                python_P, G
            )

            results.append({

                "Excel_Row": row + 1,

                "Sample": sheet.cell_value(row, 1),

                "Input_T_C": D,

                "Input_P_kbar": E,

                "Python_T_C": python_T,

                "Excel_T_C": F,

                "Temperature_Difference": t_difference,

                "Temperature_Status":
                    "PASS" if t_pass else "FAIL",

                "Python_P_kbar": python_P,

                "Excel_P_kbar": G,

                "Pressure_Difference": p_difference,

                "Pressure_Status":
                    "PASS" if p_pass else "FAIL",

                "Overall_Status":
                    "PASS" if t_pass and p_pass else "FAIL"
            })

        except (ValueError, ZeroDivisionError):

            continue

    log(
        "Main workbook valid numerical rows: "
        + str(len(results))
    )

    return results


# ============================================================
# 7. CREATE PROFESSIONAL EXCEL REPORT
# ============================================================

def create_report(input_data, author_results, main_results):

    wb = openpyxl.Workbook()

    ws = wb.active

    ws.title = "Validation_Summary"

    navy = "17365D"

    blue = "D9EAF7"

    green = "E2F0D9"

    red = "FCE4D6"

    grey = "E7E6E6"

    thin = Side(
        style="thin",
        color="D9E1F2"
    )

    def format_sheet(sheet, title, subtitle):

        sheet.sheet_view.showGridLines = False

        sheet.merge_cells("A1:J1")

        sheet["A1"] = title

        sheet["A1"].font = Font(
            name=FONT_NAME,
            size=16,
            bold=True,
            color="FFFFFF"
        )

        sheet["A1"].fill = PatternFill(
            "solid",
            fgColor=navy
        )

        sheet["A1"].alignment = Alignment(
            horizontal="left",
            vertical="center"
        )

        sheet.row_dimensions[1].height = 30

        sheet.merge_cells("A2:J2")

        sheet["A2"] = subtitle

        sheet["A2"].font = Font(
            name=FONT_NAME,
            size=10,
            italic=True,
            color="666666"
        )

        sheet.freeze_panes = "A5"

    format_sheet(
        ws,
        "STEP 05B | Hbl97 INDEPENDENT VALIDATION",
        "Independent Python reproduction against original Excel workbooks."
    )

    summary = [

        ("Validation date", datetime.now().strftime("%Y-%m-%d %H:%M")),

        ("Authors workbook", AUTHORS_FILE.name),

        ("Main workbook", MAIN_FILE.name),

        ("Authors' versions checked", len(author_results)),

        ("Main workbook numerical rows", len(main_results)),

        ("Numerical tolerance", TOLERANCE),

        ("Authors' version tests passed",
         sum(r["Overall_Status"] == "PASS" for r in author_results)),

        ("Main workbook rows passed",
         sum(r["Overall_Status"] == "PASS" for r in main_results)),

        ("Scientific validation status",
         "PARTIAL NUMERICAL REPRODUCTION ONLY")
    ]

    for row in summary:
        ws.append(list(row))

    for row in range(4, 4 + len(summary)):

        ws.cell(row, 1).font = Font(
            name=FONT_NAME,
            size=10,
            bold=True
        )

        ws.cell(row, 2).font = Font(
            name=FONT_NAME,
            size=10
        )

    # --------------------------------------------------------
    # INPUT DATA
    # --------------------------------------------------------

    input_ws = wb.create_sheet("Author_Input_Data")

    format_sheet(
        input_ws,
        "AUTHOR WORKBOOK | INPUT DATA",
        "Values read from the original authors' spreadsheet."
    )

    input_ws.append(["Parameter", "Value"])

    for key, value in input_data.items():

        input_ws.append([key, value])

    # --------------------------------------------------------
    # AUTHOR VERSION VALIDATION
    # --------------------------------------------------------

    author_ws = wb.create_sheet("Version_Validation")

    format_sheet(
        author_ws,
        "V2–V4 | INDEPENDENT VALIDATION",
        "Python calculations compared with cached Excel outputs."
    )

    headers = list(author_results[0].keys())

    author_ws.append(headers)

    for record in author_results:

        author_ws.append([
            record[h] for h in headers
        ])

    # --------------------------------------------------------
    # MAIN WORKBOOK VALIDATION
    # --------------------------------------------------------

    main_ws = wb.create_sheet("Main_Workbook_Validation")

    format_sheet(
        main_ws,
        "MAIN WORKBOOK | EQUATION VALIDATION",
        "Independent evaluation of final T and P equations using stored intermediate cells."
    )

    if main_results:

        headers = list(main_results[0].keys())

        main_ws.append(headers)

        for record in main_results:

            main_ws.append([
                record[h] for h in headers
            ])

    else:

        main_ws.append([
            "No rows passed the numerical input checks."
        ])

    # --------------------------------------------------------
    # EQUATION DOCUMENTATION
    # --------------------------------------------------------

    eq_ws = wb.create_sheet("Equation_Documentation")

    format_sheet(
        eq_ws,
        "EQUATION DOCUMENTATION",
        "Transcription of equations from the authors' spreadsheet and main workbook."
    )

    equations = [

        ["Version", "Equation component", "Source cell", "Python implementation"],

        ["V2", "Xalt", "B23", "8/15.5 - Si/Kat_sum + 0.07531"],

        ["V2", "Temperature K", "B24", "4701/(1.825 - 1.987 ln(Xalt))"],

        ["V2", "Xalm", "B25", "(Si + Al - 8)/2 + 2.158"],

        ["V2", "Pressure", "B26", "(-1341 + 0.9522 Tk + 1.987 Tk ln(Xalm))/0.2152"],

        ["V3", "Xalm", "B17", "(Si + Al - 8)/2 + 0.763"],

        ["V3", "Xalf", "B18", "Xalm/(Xalm + Fe3+/2 + 0.763)"],

        ["V3", "Pressure", "B19", "(-425 - 1719 Xalf + 2.75 Tk + 1.987 Tk ln(Xalm))/0.2"],

        ["V4", "Xalm", "B9", "(Si + Al - 8)/2 + 1.397"],

        ["V4", "Xalf", "B10", "Xalm/(Xalm + Fe3+/2 + 1.397)"],

        ["V4", "Pressure", "B11", "(-400.6 - 1551 Xalf + 1.463 Tk + 1.987 Tk ln(Xalm))/0.2034"],

        ["Main", "Temperature", "BO", "Formula transcribed from main workbook"],

        ["Main", "Pressure", "BP", "Formula transcribed from main workbook"]

    ]

    for row in equations:
        eq_ws.append(row)

    # --------------------------------------------------------
    # NOTES AND LIMITATIONS
    # --------------------------------------------------------

    notes_ws = wb.create_sheet("Limitations_and_Notes")

    format_sheet(
        notes_ws,
        "SCIENTIFIC LIMITATIONS",
        "Important interpretation requirements."
    )

    notes = [

        "1. Successful numerical reproduction does not establish scientific validity.",

        "2. V2, V3 and V4 are the versions identified in the authors' spreadsheet.",

        "3. The authors' spreadsheet identifies V3 as calibrated on natural data.",

        "4. The authors' spreadsheet recommends calcic amphiboles.",

        "5. The authors' spreadsheet warns against plagioclase with An < 10.",

        "6. The stated uncertainties in the authors' instructions are approximately ±1.2 kbar and ±37 °C.",

        "7. V5 is described as an experimental calibration requiring the Ptform program.",

        "8. The main workbook uses an iterative T-P calculation structure.",

        "9. Main workbook validation here checks final equations using cached intermediate values.",

        "10. Full chemical normalization, Fe3+ allocation and published equation verification remain separate tasks.",

        "11. One authors' test composition cannot establish calibration performance across the compositional domain.",

        "12. Do not use these results as evidence of ML predictive accuracy."
    ]

    for note in notes:
        notes_ws.append([note])

    # --------------------------------------------------------
    # PROFESSIONAL FORMATTING
    # --------------------------------------------------------

    for sheet in wb.worksheets:

        sheet.sheet_view.showGridLines = False

        for row in sheet.iter_rows():

            for cell in row:

                cell.font = Font(
                    name=FONT_NAME,
                    size=10,
                    bold=cell.row in (1, 4),
                    color=(
                        "FFFFFF"
                        if cell.row == 1
                        else "000000"
                    )
                )

                cell.alignment = Alignment(
                    vertical="center",
                    wrap_text=True
                )

                cell.border = Border(
                    bottom=thin
                )

        # Column widths

        for col in range(1, sheet.max_column + 1):

            letter = get_column_letter(col)

            max_length = 0

            for row in range(1, min(sheet.max_row, 1000) + 1):

                value = sheet.cell(row, col).value

                if value is not None:

                    max_length = max(
                        max_length,
                        len(str(value))
                    )

            sheet.column_dimensions[letter].width = min(
                max(max_length + 3, 14),
                48
            )

        # Header styling

        for cell in sheet[4]:

            if cell.value is not None:

                cell.fill = PatternFill(
                    "solid",
                    fgColor=blue
                )

                cell.font = Font(
                    name=FONT_NAME,
                    size=10,
                    bold=True,
                    color=navy
                )

    # Apply PASS / FAIL colors

    for sheet in [author_ws, main_ws]:

        for row in sheet.iter_rows():

            for cell in row:

                if cell.value == "PASS":

                    cell.fill = PatternFill(
                        "solid",
                        fgColor=green
                    )

                elif cell.value == "FAIL":

                    cell.fill = PatternFill(
                        "solid",
                        fgColor=red
                    )

    for sheet in [author_ws, main_ws]:

        if sheet.max_row >= 5:

            sheet.auto_filter.ref = sheet.dimensions

    wb.save(REPORT_PATH)

    log("Excel report saved: " + str(REPORT_PATH))


# ============================================================
# 8. MAIN EXECUTION
# ============================================================

def main():

    log("=" * 65)

    log("STEP 05B — INDEPENDENT Hbl97 NUMERICAL VALIDATION")

    log("=" * 65)

    if not MAIN_FILE.exists():
        raise FileNotFoundError(str(MAIN_FILE))

    if not AUTHORS_FILE.exists():
        raise FileNotFoundError(str(AUTHORS_FILE))

    input_data, author_results = (
        validate_authors_workbook()
    )

    main_results = validate_main_workbook()

    create_report(
        input_data,
        author_results,
        main_results
    )

    log("=" * 65)

    log("STEP 05B COMPLETED")

    log(
        "Authors' version checks: "
        + str(len(author_results))
    )

    log(
        "Main workbook numerical rows: "
        + str(len(main_results))
    )

    log("Report: " + str(REPORT_PATH))

    log("Log: " + str(LOG_PATH))

    log("=" * 65)


if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logging.exception(
            "STEP 05B FAILED: %s",
            error
        )

        print("\nSTEP 05B FAILED")

        print(str(error))

        raise

