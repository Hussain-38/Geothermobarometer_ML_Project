
# -*- coding: utf-8 -*-
"""
STEP 02: GEOTHERMOBAROMETER CALIBRATION WORKBOOK AUDIT

Compatible with Windows and Spyder.

FUNCTIONS
---------
1. Inventory XLS, XLSX and XLSM workbooks.
2. Record workbook metadata and SHA-256 checksums.
3. Inspect worksheet names and dimensions.
4. Extract formulas from XLSX/XLSM.
5. Extract XLS formulas using Microsoft Excel COM, if available.
6. Record stored formula values where accessible.
7. Generate a formatted Excel audit report.
8. Generate an execution log.

IMPORTANT
---------
Original workbooks are never intentionally modified.
This is a structural formula audit, NOT independent validation
of the published thermodynamic calibration equations.
"""

# ============================================================
# 1. IMPORT LIBRARIES
# ============================================================

from pathlib import Path
from datetime import datetime
import hashlib
import logging
import sys
import traceback

import pandas as pd
import openpyxl

from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# 2. PROJECT SETTINGS
# ============================================================

PROJECT_DIR = Path(
    r"D:\Data Science Assignment\My Projects\Geothermobarometer_ML_Project"
)

OUTPUT_DIR = PROJECT_DIR / "02_calibration_formula_audit"

LOG_DIR = OUTPUT_DIR / "logs"

REPORT_PATH = OUTPUT_DIR / "Step02_Calibration_Formula_Audit.xlsx"

LOG_PATH = LOG_DIR / "step02_audit_log.txt"

MAX_FORMULAS_PER_WORKBOOK = 100000

FONT_NAME = "Times New Roman"

HEADER_COLOR = "244062"


# ============================================================
# 3. CREATE OUTPUT DIRECTORIES
# ============================================================

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR.mkdir(parents=True, exist_ok=True)


# ============================================================
# 4. LOGGING CONFIGURATION
# ============================================================

logger = logging.getLogger("STEP02")

logger.setLevel(logging.INFO)

logger.handlers.clear()

file_handler = logging.FileHandler(
    LOG_PATH,
    mode="w",
    encoding="utf-8"
)

file_handler.setFormatter(
    logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )
)

logger.addHandler(file_handler)

console_handler = logging.StreamHandler(sys.stdout)

console_handler.setFormatter(
    logging.Formatter("%(levelname)s: %(message)s")
)

logger.addHandler(console_handler)


def log(message):
    """Write messages to console and log file."""
    logger.info(str(message))


# ============================================================
# 5. SHA-256 CHECKSUM
# ============================================================

def calculate_sha256(file_path):

    digest = hashlib.sha256()

    with open(file_path, "rb") as file:

        while True:

            chunk = file.read(1024 * 1024)

            if not chunk:
                break

            digest.update(chunk)

    return digest.hexdigest()


# ============================================================
# 6. SELECT INPUT DIRECTORY
# ============================================================

def select_input_directory():

    print("\n" + "=" * 60)

    print("STEP 02: CALIBRATION WORKBOOK AUDIT")

    print("=" * 60)

    print(
        "\nEnter the FULL folder path containing your "
        "geothermobarometer calibration Excel files."
    )

    print(
        '\nExample: D:\\Geothermobarometer\\Calibration_Files'
    )

    entered_path = input(
        "\nCalibration workbook folder: "
    ).strip().strip('"')

    if not entered_path:

        raise ValueError(
            "No calibration workbook folder was provided."
        )

    folder = Path(entered_path).expanduser()

    if not folder.exists():

        raise FileNotFoundError(
            "Folder does not exist: {}".format(folder)
        )

    if not folder.is_dir():

        raise NotADirectoryError(
            "Selected path is not a directory."
        )

    return folder.resolve()


# ============================================================
# 7. FIND EXCEL WORKBOOKS
# ============================================================

def find_workbooks(input_folder):

    supported_extensions = {
        ".xls",
        ".xlsx",
        ".xlsm"
    }

    files = []

    for file_path in input_folder.rglob("*"):

        if not file_path.is_file():
            continue

        if file_path.name.startswith("~$"):
            continue

        if file_path.suffix.lower() not in supported_extensions:
            continue

        try:

            if file_path.resolve() == REPORT_PATH.resolve():
                continue

        except Exception:
            pass

        files.append(file_path)

    return sorted(
        files,
        key=lambda x: str(x).lower()
    )


# ============================================================
# 8. ERROR RECORDING
# ============================================================

def record_error(error_data, workbook, stage, error):

    error_data.append({

        "Workbook": workbook,

        "Stage": stage,

        "Error": str(error)

    })

    log(
        "WARNING | {} | {} | {}".format(
            workbook,
            stage,
            error
        )
    )


# ============================================================
# 9. INSPECT XLSX AND XLSM WORKBOOKS
# ============================================================

def inspect_modern_workbook(
    file_path,
    formula_data,
    worksheet_data,
    error_data
):

    formula_count = 0

    try:

        workbook_formulas = openpyxl.load_workbook(

            file_path,

            read_only=True,

            data_only=False,

            keep_vba=(
                file_path.suffix.lower() == ".xlsm"
            )

        )

        workbook_values = openpyxl.load_workbook(

            file_path,

            read_only=True,

            data_only=True,

            keep_vba=(
                file_path.suffix.lower() == ".xlsm"
            )

        )

        try:

            for worksheet in workbook_formulas.worksheets:

                value_sheet = workbook_values[
                    worksheet.title
                ]

                worksheet_data.append({

                    "Workbook": file_path.name,

                    "Worksheet": worksheet.title,

                    "Rows": worksheet.max_row,

                    "Columns": worksheet.max_column,

                    "Visibility": worksheet.sheet_state,

                    "Inspection_Method": "openpyxl"

                })

                for row in worksheet.iter_rows():

                    for cell in row:

                        formula = cell.value

                        if (
                            isinstance(formula, str)
                            and formula.startswith("=")
                        ):

                            cached_value = value_sheet[
                                cell.coordinate
                            ].value

                            formula_data.append({

                                "Workbook": file_path.name,

                                "Worksheet": worksheet.title,

                                "Cell": cell.coordinate,

                                "Formula": formula,

                                "Stored_Cached_Value": cached_value,

                                "Formula_Length": len(formula)

                            })

                            formula_count += 1

                            if (
                                formula_count
                                >= MAX_FORMULAS_PER_WORKBOOK
                            ):

                                record_error(

                                    error_data,

                                    file_path.name,

                                    "Formula extraction limit",

                                    "Configured formula limit reached."

                                )

                                return (
                                    formula_count,
                                    "Formula extraction capped"
                                )

        finally:

            workbook_formulas.close()

            workbook_values.close()

        return formula_count, "Completed"

    except Exception as exc:

        record_error(

            error_data,

            file_path.name,

            "XLSX/XLSM inspection",

            exc

        )

        return formula_count, "Failed"



# ============================================================
# 9. OPTIMIZED XLSX AND XLSM WORKBOOK INSPECTION
# ============================================================

def inspect_modern_workbook(
    file_path,
    formula_data,
    worksheet_data,
    error_data
):

    formula_count = 0

    workbook_formulas = None
    workbook_values = None

    try:

        log("Opening workbook: {}".format(file_path.name))

        # Open formula version
        workbook_formulas = openpyxl.load_workbook(
            file_path,
            read_only=True,
            data_only=False,
            keep_vba=(file_path.suffix.lower() == ".xlsm")
        )

        # Open cached-value version
        workbook_values = openpyxl.load_workbook(
            file_path,
            read_only=True,
            data_only=True,
            keep_vba=(file_path.suffix.lower() == ".xlsm")
        )

        for worksheet in workbook_formulas.worksheets:

            sheet_name = worksheet.title

            log(
                "Inspecting worksheet: {} | {}".format(
                    file_path.name,
                    sheet_name
                )
            )

            value_sheet = workbook_values[sheet_name]

            # Record original worksheet dimensions
            original_dimension = worksheet.calculate_dimension()

            original_rows = worksheet.max_row
            original_columns = worksheet.max_column

            # Detect potentially inflated worksheet dimensions.
            # Resetting dimensions allows the XML stream to determine
            # the actual populated cell extent.
            if (
                original_rows is not None
                and original_rows > 100000
            ) or (
                original_columns is not None
                and original_columns > 500
            ):

                log(
                    "WARNING: Unusually large worksheet dimensions: "
                    "{} | {}".format(
                        sheet_name,
                        original_dimension
                    )
                )

                worksheet.reset_dimensions()
                value_sheet.reset_dimensions()

            worksheet_data.append({

                "Workbook": file_path.name,

                "Worksheet": sheet_name,

                "Rows": original_rows,

                "Columns": original_columns,

                "Reported_Dimension": original_dimension,

                "Visibility": worksheet.sheet_state,

                "Inspection_Method": "Optimized openpyxl streaming"

            })

            rows_scanned = 0

            # IMPORTANT:
            # Read formula and cached-value worksheets together.
            # This avoids repeated random-access XML parsing.

            formula_iterator = worksheet.iter_rows()

            value_iterator = value_sheet.iter_rows()

            for formula_row, value_row in zip(
                formula_iterator,
                value_iterator
            ):

                rows_scanned += 1

                for cell, value_cell in zip(
                    formula_row,
                    value_row
                ):

                    formula = cell.value

                    if (
                        isinstance(formula, str)
                        and formula.startswith("=")
                    ):

                        cached_value = value_cell.value

                        formula_data.append({

                            "Workbook": file_path.name,

                            "Worksheet": sheet_name,

                            "Cell": cell.coordinate,

                            "Formula": formula,

                            "Stored_Cached_Value": cached_value,

                            "Formula_Length": len(formula)

                        })

                        formula_count += 1

                        if (
                            formula_count
                            >= MAX_FORMULAS_PER_WORKBOOK
                        ):

                            record_error(

                                error_data,

                                file_path.name,

                                "Formula extraction limit",

                                "Configured formula extraction "
                                "limit reached."

                            )

                            log(
                                "Formula extraction limit reached: "
                                "{}".format(file_path.name)
                            )

                            return (
                                formula_count,
                                "Formula extraction capped"
                            )

                # Periodic progress reporting
                if rows_scanned % 10000 == 0:

                    log(
                        "Progress: {} | {} | Rows scanned: {} | "
                        "Formulas found: {}".format(
                            file_path.name,
                            sheet_name,
                            rows_scanned,
                            formula_count
                        )
                    )

            log(
                "Worksheet completed: {} | {} | "
                "Rows scanned: {} | Formulas: {}".format(
                    file_path.name,
                    sheet_name,
                    rows_scanned,
                    formula_count
                )
            )

        log(
            "Workbook completed: {} | Total formulas: {}".format(
                file_path.name,
                formula_count
            )
        )

        return formula_count, "Completed"

    except Exception as exc:

        record_error(

            error_data,

            file_path.name,

            "XLSX/XLSM inspection",

            exc

        )

        return formula_count, "Failed"

    finally:

        if workbook_formulas is not None:

            workbook_formulas.close()

        if workbook_values is not None:

            workbook_values.close()



# ============================================================
# 11. EXTRACT FORMULAS FROM LEGACY XLS USING EXCEL
# ============================================================

def extract_legacy_formulas(
    files,
    formula_data,
    error_data
):

    xls_files = [

        file_path

        for file_path in files

        if file_path.suffix.lower() == ".xls"

    ]

    if not xls_files:

        return "No legacy XLS files require extraction."

    try:

        import win32com.client

    except ImportError:

        message = (
            "pywin32 is unavailable. "
            "Legacy XLS formulas were not extracted. "
            "Install pywin32 and Microsoft Excel."
        )

        log(message)

        return message

    excel = None

    total_formulas = 0

    try:

        excel = win32com.client.DispatchEx(
            "Excel.Application"
        )

        excel.Visible = False

        excel.DisplayAlerts = False

        excel.AskToUpdateLinks = False

        for file_path in xls_files:

            workbook = None

            workbook_formula_count = 0

            try:

                log(
                    "Extracting XLS formulas: {}".format(
                        file_path.name
                    )
                )

                workbook = excel.Workbooks.Open(

                    str(file_path.resolve()),

                    UpdateLinks=0,

                    ReadOnly=True,

                    IgnoreReadOnlyRecommended=True,

                    AddToMru=False

                )

                for worksheet in workbook.Worksheets:

                    try:

                        # -4123 = Excel xlCellTypeFormulas

                        formula_cells = (
                            worksheet.UsedRange.SpecialCells(-4123)
                        )

                        for cell in formula_cells.Cells:

                            formula = cell.Formula

                            formula_data.append({

                                "Workbook": file_path.name,

                                "Worksheet": str(worksheet.Name),

                                "Cell": str(cell.Address),

                                "Formula": str(formula),

                                "Stored_Cached_Value": cell.Value,

                                "Formula_Length": len(str(formula))

                            })

                            workbook_formula_count += 1

                            total_formulas += 1

                            if (
                                workbook_formula_count
                                >= MAX_FORMULAS_PER_WORKBOOK
                            ):

                                record_error(

                                    error_data,

                                    file_path.name,

                                    "Formula extraction limit",

                                    "Configured formula limit reached."

                                )

                                break

                    except Exception as exc:

                        message = str(exc).lower()

                        if "no cells were found" not in message:

                            record_error(

                                error_data,

                                file_path.name,

                                "XLS worksheet formula extraction",

                                exc

                            )

                    if (
                        workbook_formula_count
                        >= MAX_FORMULAS_PER_WORKBOOK
                    ):

                        break

                workbook.Close(SaveChanges=False)

                workbook = None

            except Exception as exc:

                record_error(

                    error_data,

                    file_path.name,

                    "XLS Excel COM extraction",

                    exc

                )

            finally:

                if workbook is not None:

                    try:

                        workbook.Close(SaveChanges=False)

                    except Exception:

                        pass

        return (
            "Excel COM extraction completed. "
            "Formulas extracted: {}".format(total_formulas)
        )

    except Exception as exc:

        message = (
            "Excel COM could not be initialized: {}".format(exc)
        )

        log(message)

        return message

    finally:

        if excel is not None:

            try:

                excel.Quit()

            except Exception:

                pass


# ============================================================
# 12. FORMAT EXCEL REPORT
# ============================================================

def format_report(report_path):

    workbook = openpyxl.load_workbook(report_path)

    try:

        for worksheet in workbook.worksheets:

            worksheet.freeze_panes = "A2"

            if worksheet.max_row >= 1:

                worksheet.auto_filter.ref = worksheet.dimensions

            for cell in worksheet[1]:

                cell.font = Font(

                    name=FONT_NAME,

                    size=10,

                    bold=True,

                    color="FFFFFF"

                )

                cell.fill = PatternFill(

                    fill_type="solid",

                    fgColor=HEADER_COLOR

                )

                cell.alignment = Alignment(

                    horizontal="center",

                    vertical="center",

                    wrap_text=True

                )

            for row in worksheet.iter_rows(min_row=2):

                for cell in row:

                    cell.font = Font(

                        name=FONT_NAME,

                        size=9

                    )

                    cell.alignment = Alignment(

                        vertical="top",

                        wrap_text=True

                    )

            for column in worksheet.columns:

                column_number = column[0].column

                letter = get_column_letter(column_number)

                maximum_length = 12

                for cell in column:

                    if cell.value is not None:

                        maximum_length = max(

                            maximum_length,

                            min(len(str(cell.value)), 55)

                        )

                worksheet.column_dimensions[
                    letter
                ].width = min(maximum_length + 2, 57)

            worksheet.row_dimensions[1].height = 30

        workbook.save(report_path)

    finally:

        workbook.close()


# ============================================================
# 13. MAIN ANALYSIS
# ============================================================

def main():

    log("=" * 60)

    log("STEP 02 STARTED")

    log("Project: {}".format(PROJECT_DIR))

    log("Output: {}".format(OUTPUT_DIR))

    log("=" * 60)

    if not PROJECT_DIR.exists():

        raise FileNotFoundError(
            "Project directory not found: {}".format(PROJECT_DIR)
        )

    input_folder = select_input_directory()

    files = find_workbooks(input_folder)

    if not files:

        raise FileNotFoundError(
            "No Excel calibration workbooks found in: {}".format(
                input_folder
            )
        )

    log("Input folder: {}".format(input_folder))

    log("Workbooks found: {}".format(len(files)))

    workbook_data = []

    worksheet_data = []

    formula_data = []

    error_data = []

    # --------------------------------------------------------
    # WORKBOOK INVENTORY
    # --------------------------------------------------------

    for file_path in files:

        log("Inspecting: {}".format(file_path.name))

        try:

            workbook_data.append({

                "Workbook": file_path.name,

                "Full_Path": str(file_path.resolve()),

                "Extension": file_path.suffix.lower(),

                "Size_KB": round(
                    file_path.stat().st_size / 1024,
                    2
                ),

                "Modified": datetime.fromtimestamp(
                    file_path.stat().st_mtime
                ).isoformat(timespec="seconds"),

                "SHA256": calculate_sha256(file_path)

            })

        except Exception as exc:

            record_error(

                error_data,

                file_path.name,

                "File metadata",

                exc

            )

        if file_path.suffix.lower() in {".xlsx", ".xlsm"}:

            inspect_modern_workbook(

                file_path,

                formula_data,

                worksheet_data,

                error_data

            )

        elif file_path.suffix.lower() == ".xls":

            inspect_legacy_xls(

                file_path,

                worksheet_data,

                error_data

            )

    # --------------------------------------------------------
    # LEGACY XLS FORMULA EXTRACTION
    # --------------------------------------------------------

    com_status = extract_legacy_formulas(

        files,

        formula_data,

        error_data

    )

    # --------------------------------------------------------
    # WORKBOOK SUMMARY
    # --------------------------------------------------------

    summary_rows = []

    for item in workbook_data:

        name = item["Workbook"]

        sheet_count = sum(

            1

            for row in worksheet_data

            if row["Workbook"] == name

        )

        formula_count = sum(

            1

            for row in formula_data

            if row["Workbook"] == name

        )

        has_errors = any(

            row["Workbook"] == name

            for row in error_data

        )

        summary_rows.append({

            "Workbook": name,

            "Extension": item["Extension"],

            "Size_KB": item["Size_KB"],

            "Worksheet_Count": sheet_count,

            "Formula_Count_Extracted": formula_count,

            "Inventory_Status": (

                "Review errors"

                if has_errors

                else "Inventory completed"

            ),

            "SHA256": item["SHA256"],

            "Full_Path": item["Full_Path"]

        })

    # --------------------------------------------------------
    # AUDIT NOTES
    # --------------------------------------------------------

    notes = [

        (
            "Run_Date_Time",
            datetime.now().isoformat(timespec="seconds")
        ),

        ("Input_Folder", str(input_folder)),

        ("Output_Folder", str(OUTPUT_DIR)),

        ("Workbook_Count", len(workbook_data)),

        ("Worksheet_Count", len(worksheet_data)),

        ("Formula_Count", len(formula_data)),

        ("Error_Warning_Count", len(error_data)),

        ("Legacy_XLS_Status", com_status),

        (
            "Scientific_Limitation",
            "Formula extraction does not independently validate "
            "published thermodynamic calibration equations."
        ),

        (
            "Cached_Value_Limitation",
            "Stored cached values can be absent or outdated if "
            "Excel has not recalculated and saved the workbook."
        ),

        (
            "Source_Protection",
            "Source workbooks are opened read-only where applicable. "
            "No source workbook is saved or modified by this script."
        )

    ]

    # --------------------------------------------------------
    # CREATE REPORT TABLES
    # --------------------------------------------------------

    summary_df = pd.DataFrame(summary_rows)

    worksheet_df = pd.DataFrame(

        worksheet_data,

        columns=[

            "Workbook",

            "Worksheet",

            "Rows",

            "Columns",

            "Visibility",

            "Inspection_Method"

        ]

    )

    formula_df = pd.DataFrame(

        formula_data,

        columns=[

            "Workbook",

            "Worksheet",

            "Cell",

            "Formula",

            "Stored_Cached_Value",

            "Formula_Length"

        ]

    )

    error_df = pd.DataFrame(

        error_data,

        columns=[

            "Workbook",

            "Stage",

            "Error"

        ]

    )

    notes_df = pd.DataFrame(

        notes,

        columns=[

            "Parameter",

            "Value"

        ]

    )

    # --------------------------------------------------------
    # EXPORT EXCEL REPORT
    # --------------------------------------------------------

    with pd.ExcelWriter(

        REPORT_PATH,

        engine="openpyxl"

    ) as writer:

        summary_df.to_excel(

            writer,

            sheet_name="Workbook_Summary",

            index=False

        )

        worksheet_df.to_excel(

            writer,

            sheet_name="Worksheet_Inventory",

            index=False

        )

        formula_df.to_excel(

            writer,

            sheet_name="Formula_Cells",

            index=False

        )

        error_df.to_excel(

            writer,

            sheet_name="Errors_and_Warnings",

            index=False

        )

        notes_df.to_excel(

            writer,

            sheet_name="Audit_Notes",

            index=False

        )

    # Apply formatting

    format_report(REPORT_PATH)

    # --------------------------------------------------------
    # FINAL SUMMARY
    # --------------------------------------------------------

    log("=" * 60)

    log("STEP 02 FINISHED")

    log("Workbooks inventoried: {}".format(len(workbook_data)))

    log("Worksheets inventoried: {}".format(len(worksheet_data)))

    log("Formulas extracted: {}".format(len(formula_data)))

    log("Errors and warnings: {}".format(len(error_data)))

    log("Excel report: {}".format(REPORT_PATH))

    log("Execution log: {}".format(LOG_PATH))

    log("=" * 60)

    print("\nSTEP 02 REPORT GENERATED")

    print("\nExcel report:")

    print(REPORT_PATH)

    print("\nExecution log:")

    print(LOG_PATH)

    print("\nInspect the Errors_and_Warnings worksheet.")


# ============================================================
# 14. EXECUTE
# ============================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as exc:

        logger.exception(
            "STEP 02 FAILED: {}".format(exc)
        )

        print("\nSTEP 02 FAILED")

        print(exc)

        print("\nFULL ERROR TRACEBACK:")

        traceback.print_exc()

        print("\nCheck the log file:")

        print(LOG_PATH)
