# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 14:36:48 2026

@author: hussa
"""


# -*- coding: utf-8 -*-
"""
======================================================================
STEP 04 — MASTER GEOTHERMOBAROMETER CALIBRATION DATABASE
======================================================================

PROJECT:
    Geothermobarometer Machine Learning Research Project

INPUT:
    03C1_corrected_dependency_parser/
    Step03C1_Corrected_Dependency_Analysis.xlsx

OPTIONAL ORIGINAL SOURCE DIRECTORY:
    Geothermobarometers Calibrations/

OUTPUT:
    04_master_calibration_database/
        Step04_Master_Geothermobarometer_Database.xlsx
        logs/
            step04_execution_log.txt

OBJECTIVES:
    1. Inventory calibration workbooks.
    2. Inventory individual worksheets.
    3. Create a master calibration register.
    4. Create a publication and DOI register.
    5. Create an equation verification register.
    6. Record mineral assemblages and input variables.
    7. Record calibration applicability ranges.
    8. Map verified Excel output cells.
    9. Assess possible duplicate workbook files.
   10. Establish scientific verification status.

SCIENTIFIC SAFEGUARDS:
    - A workbook is not automatically a unique calibration.
    - A worksheet is not automatically a unique calibration.
    - Filename similarity is not proof of scientific duplication.
    - No equation, DOI, mineral assemblage or P-T range is invented.
    - Unverified fields are explicitly marked Pending verification.
    - Original calibration files are never modified.

COMPATIBLE WITH:
    Spyder / Anaconda / Python
======================================================================
"""

# ==============================================================
# 1. IMPORT LIBRARIES
# ==============================================================

from pathlib import Path
from datetime import datetime

import hashlib
import logging
import re
import sys
import unicodedata

import openpyxl

from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


# ==============================================================
# 2. PROJECT DIRECTORIES
# ==============================================================

PROJECT_DIR = Path(__file__).resolve().parent

INPUT_FILE = (
    PROJECT_DIR
    / "03C1_corrected_dependency_parser"
    / "Step03C1_Corrected_Dependency_Analysis.xlsx"
)

# Original calibration folder.
# This is optional for Step 04.

SOURCE_CALIBRATION_DIR = (
    PROJECT_DIR
    / "Geothermobarometers Calibrations"
)

OUTPUT_DIR = (
    PROJECT_DIR
    / "04_master_calibration_database"
)

LOG_DIR = OUTPUT_DIR / "logs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = (
    OUTPUT_DIR
    / "Step04_Master_Geothermobarometer_Database.xlsx"
)

LOG_FILE = (
    LOG_DIR
    / "step04_execution_log.txt"
)


# ==============================================================
# 3. DESIGN SETTINGS
# ==============================================================

FONT_NAME = "Times New Roman"

TITLE_SIZE = 16
SUBTITLE_SIZE = 11
HEADER_SIZE = 10
BODY_SIZE = 10

DARK_BLUE = "17365D"
MEDIUM_BLUE = "4472C4"
LIGHT_BLUE = "D9EAF7"
LIGHT_GREY = "F2F2F2"
LIGHT_GREEN = "E2F0D9"
LIGHT_YELLOW = "FFF2CC"
WHITE = "FFFFFF"
DARK_GREY = "595959"

# Supported calibration workbook formats.

SUPPORTED_EXTENSIONS = {

    ".xls",
    ".xlsx",
    ".xlsm"

}


# ==============================================================
# 4. LOGGING
# ==============================================================

logger = logging.getLogger("STEP04")

logger.setLevel(logging.INFO)

logger.handlers.clear()

formatter = logging.Formatter(
    "%(asctime)s | %(levelname)s | %(message)s"
)

file_handler = logging.FileHandler(
    LOG_FILE,
    mode="w",
    encoding="utf-8"
)

file_handler.setFormatter(formatter)

logger.addHandler(file_handler)

console_handler = logging.StreamHandler(sys.stdout)

console_handler.setFormatter(formatter)

logger.addHandler(console_handler)


def log(message):

    logger.info(message)


# ==============================================================
# 5. GENERAL HELPERS
# ==============================================================

def sha256_file(file_path, chunk_size=1024 * 1024):

    digest = hashlib.sha256()

    with open(file_path, "rb") as file:

        for chunk in iter(
            lambda: file.read(chunk_size),
            b""
        ):

            digest.update(chunk)

    return digest.hexdigest()


def clean_text(value):

    if value is None:

        return ""

    return str(value).strip()


def normalized_stem(filename):

    """
    Generate a filename comparison key.

    This is only for finding possible duplicate names.
    It does not establish calibration identity.
    """

    stem = Path(filename).stem

    stem = unicodedata.normalize(
        "NFKC",
        stem
    )

    stem = stem.casefold()

    stem = re.sub(
        r"\s+",
        " ",
        stem
    )

    return stem.strip()


def safe_value(value):

    """
    Prevent formula-looking text from becoming active Excel formulas.
    """

    if isinstance(value, str):

        if value.startswith("="):

            return "'" + value

    return value


# ==============================================================
# 6. LOAD STEP 03C.1 REPORT
# ==============================================================

def load_source_inventory():

    log("Reading Step 03C.1 report...")

    if not INPUT_FILE.exists():

        raise FileNotFoundError(
            "Step 03C.1 report not found:\n{}".format(
                INPUT_FILE
            )
        )

    source_wb = load_workbook(

        INPUT_FILE,

        read_only=True,

        data_only=True

    )

    required_sheet = "Worksheet_Summary"

    if required_sheet not in source_wb.sheetnames:

        raise ValueError(
            "Worksheet_Summary is missing from Step 03C.1 report."
        )

    ws = source_wb[required_sheet]

    rows = ws.iter_rows(
        values_only=True
    )

    headers = next(rows)

    headers = [
        clean_text(value)
        for value in headers
    ]

    required_columns = {

        "Workbook",

        "Worksheet",

        "Formula_Count",

        "Unique_Patterns"

    }

    missing = required_columns - set(headers)

    if missing:

        raise ValueError(
            "Missing required inventory columns: {}".format(
                missing
            )
        )

    records = []

    for row in rows:

        record = dict(
            zip(headers, row)
        )

        if not record.get("Workbook"):

            continue

        records.append(record)

    source_wb.close()

    log(
        "Worksheet records recovered: {}".format(
            len(records)
        )
    )

    return records


# ==============================================================
# 7. BUILD WORKBOOK INVENTORY
# ==============================================================

def build_workbook_inventory(worksheet_records):

    workbook_dict = {}

    for record in worksheet_records:

        filename = clean_text(
            record["Workbook"]
        )

        if filename not in workbook_dict:

            workbook_dict[filename] = {

                "Workbook": filename,

                "Worksheet_Count": 0,

                "Total_Formula_Count": 0,

                "Unique_Patterns_Summed": 0,

                "File_Type": Path(filename).suffix.lower(),

                "Filename_Stem": normalized_stem(filename),

                "Method_Identity": "Pending verification",

                "Publication_Association": "Pending verification",

                "Scientific_Status": "Not yet verified"

            }

        entry = workbook_dict[filename]

        entry["Worksheet_Count"] += 1

        entry["Total_Formula_Count"] += int(
            record["Formula_Count"] or 0
        )

        entry["Unique_Patterns_Summed"] += int(
            record["Unique_Patterns"] or 0
        )

    result = []

    for number, entry in enumerate(

        workbook_dict.values(),

        start=1

    ):

        entry["Workbook_ID"] = "WB-{:03d}".format(
            number
        )

        result.append(entry)

    return result


# ==============================================================
# 8. BUILD WORKSHEET INVENTORY
# ==============================================================

def build_worksheet_inventory(worksheet_records):

    result = []

    for number, record in enumerate(

        worksheet_records,

        start=1

    ):

        result.append({

            "Worksheet_ID": "WS-{:04d}".format(number),

            "Workbook": clean_text(
                record["Workbook"]
            ),

            "Worksheet": clean_text(
                record["Worksheet"]
            ),

            "Formula_Count": record["Formula_Count"],

            "Unique_Patterns": record["Unique_Patterns"],

            "Worksheet_Role": "Pending verification",

            "Input_or_Output": "Pending verification",

            "Scientific_Interpretation": "Pending verification"

        })

    return result


# ==============================================================
# 9. BUILD CALIBRATION REGISTER
# ==============================================================

def build_calibration_register(workbook_inventory):

    """
    One provisional candidate per source workbook.

    These records are not assumed to represent unique
    scientific calibration methods.

    Multiple workbooks can later be mapped to the same
    Calibration_ID after publication-level verification.
    """

    result = []

    for number, workbook in enumerate(

        workbook_inventory,

        start=1

    ):

        result.append({

            "Calibration_ID": "CAL-{:03d}".format(number),

            "Source_Workbook_ID": workbook["Workbook_ID"],

            "Source_Workbook": workbook["Workbook"],

            "Method_Name": "",

            "Method_Abbreviation": "",

            "Method_Type": "Pending verification",

            "Publication_Year": "",

            "Calibration_Version": "Pending verification",

            "Version_Relationship": "Pending verification",

            "Mineral_Assemblage": "",

            "Equation_Number": "",

            "Temperature_Equation": "",

            "Pressure_Equation": "",

            "Temperature_Units": "",

            "Pressure_Units": "",

            "Calibration_Identity_Status": "Unresolved",

            "Duplicate_or_Revision_Status": "Unresolved",

            "Scientific_Verification_Status": "Pending"

        })

    return result


# ==============================================================
# 10. BUILD PUBLICATION REGISTER
# ==============================================================

def build_publication_register(calibration_register):

    result = []

    for calibration in calibration_register:

        result.append({

            "Publication_ID": (
                "PUB-" + calibration["Calibration_ID"].split("-")[1]
            ),

            "Calibration_ID": calibration["Calibration_ID"],

            "Authors": "",

            "Publication_Year": "",

            "Publication_Title": "",

            "Journal": "",

            "Volume": "",

            "Issue": "",

            "Pages": "",

            "DOI": "",

            "DOI_Verification": "Pending",

            "Primary_or_Secondary_Source": "Pending verification",

            "Equation_Location": "",

            "Publication_Notes": ""

        })

    return result


# ==============================================================
# 11. BUILD EQUATION REGISTER
# ==============================================================

def build_equation_register(calibration_register):

    result = []

    for calibration in calibration_register:

        result.append({

            "Equation_ID": (
                "EQ-" + calibration["Calibration_ID"].split("-")[1]
            ),

            "Calibration_ID": calibration["Calibration_ID"],

            "Publication_ID": (
                "PUB-" + calibration["Calibration_ID"].split("-")[1]
            ),

            "Equation_Number": "",

            "Equation_Description": "",

            "Published_Equation": "",

            "Excel_Formula": "",

            "Python_Implementation": "",

            "Independent_Reproduction": "Not started",

            "Numerical_Comparison": "Not started",

            "Equation_Verification_Status": "Pending"

        })

    return result


# ==============================================================
# 12. BUILD MINERAL ASSEMBLAGE REGISTER
# ==============================================================

def build_mineral_register(calibration_register):

    result = []

    for calibration in calibration_register:

        result.append({

            "Calibration_ID": calibration["Calibration_ID"],

            "Method_Name": "",

            "Mineral_1": "",

            "Mineral_2": "",

            "Mineral_3": "",

            "Mineral_4": "",

            "Additional_Minerals": "",

            "Required_Assemblage": "",

            "Assemblage_Constraints": "",

            "Reference_Location": "",

            "Verification_Status": "Pending"

        })

    return result


# ==============================================================
# 13. BUILD INPUT VARIABLE REGISTER
# ==============================================================

def build_input_register(calibration_register):

    result = []

    for calibration in calibration_register:

        result.append({

            "Calibration_ID": calibration["Calibration_ID"],

            "Variable_Name": "",

            "Chemical_Oxide_or_Component": "",

            "Mineral": "",

            "Required_Unit": "",

            "Normalization_Method": "",

            "Cation_Basis": "",

            "Formula_Cell_or_Source": "",

            "Publication_Reference": "",

            "Verification_Status": "Pending"

        })

    return result


# ==============================================================
# 14. BUILD CALIBRATION RANGE REGISTER
# ==============================================================

def build_range_register(calibration_register):

    result = []

    for calibration in calibration_register:

        result.append({

            "Calibration_ID": calibration["Calibration_ID"],

            "Temperature_Min": "",

            "Temperature_Max": "",

            "Temperature_Unit": "",

            "Pressure_Min": "",

            "Pressure_Max": "",

            "Pressure_Unit": "",

            "Calibration_Domain": "",

            "Applicability_Conditions": "",

            "Extrapolation_Allowed": "Pending verification",

            "Publication_Reference": "",

            "Verification_Status": "Pending"

        })

    return result


# ==============================================================
# 15. BUILD EXCEL OUTPUT MAPPING
# ==============================================================

def build_output_mapping(calibration_register):

    result = []

    for calibration in calibration_register:

        result.append({

            "Calibration_ID": calibration["Calibration_ID"],

            "Source_Workbook": calibration["Source_Workbook"],

            "Source_Worksheet": "",

            "Temperature_Output_Cell": "",

            "Pressure_Output_Cell": "",

            "Cached_Temperature": "",

            "Cached_Pressure": "",

            "Output_Units": "",

            "Output_Cell_Verification": "Pending",

            "Python_Comparison_Status": "Pending"

        })

    return result


# ==============================================================
# 16. BUILD VERIFICATION STATUS REGISTER
# ==============================================================

def build_verification_register(calibration_register):

    result = []

    checks = [

        "Unique method identified",

        "Primary publication located",

        "DOI verified",

        "Equation located in publication",

        "Mineral assemblage verified",

        "Chemical input variables verified",

        "Normalization procedure verified",

        "Calibration ranges verified",

        "Excel output cells verified",

        "Python equation implemented",

        "Python reproduces Excel output",

        "Independent scientific review completed"

    ]

    for calibration in calibration_register:

        for check in checks:

            result.append({

                "Calibration_ID": calibration["Calibration_ID"],

                "Source_Workbook": calibration["Source_Workbook"],

                "Verification_Item": check,

                "Status": "Pending",

                "Evidence_or_Source": "",

                "Comments": ""

            })

    return result


# ==============================================================
# 17. DUPLICATE WORKBOOK ASSESSMENT
# ==============================================================

def build_duplicate_assessment(workbook_inventory):

    log("Assessing possible duplicate workbook files...")

    result = []

    if not SOURCE_CALIBRATION_DIR.exists():

        log(
            "Original calibration folder not found. "
            "File hashing will be skipped."
        )

        for workbook in workbook_inventory:

            result.append({

                "Workbook": workbook["Workbook"],

                "Filename_Stem": workbook["Filename_Stem"],

                "Source_File_Found": "Not checked",

                "SHA256": "",

                "Exact_File_Duplicate": "Not assessed",

                "Filename_Similarity": "Not assessed",

                "Scientific_Duplication": "Not assessed",

                "Comments": (
                    "Original source directory unavailable."
                )

            })

        return result

    # ----------------------------------------------------------
    # Search for source files
    # ----------------------------------------------------------

    source_files = [

        path

        for path in SOURCE_CALIBRATION_DIR.rglob("*")

        if (

            path.is_file()

            and

            path.suffix.lower() in SUPPORTED_EXTENSIONS

        )

    ]

    source_by_name = defaultdict(list)

    for path in source_files:

        source_by_name[path.name.casefold()].append(path)

    hash_map = {}

    stem_map = defaultdict(list)

    for path in source_files:

        try:

            digest = sha256_file(path)

        except Exception as error:

            log(
                "Could not hash {}: {}".format(
                    path.name,
                    error
                )
            )

            digest = ""

        stem = normalized_stem(path.name)

        hash_map[str(path)] = digest

        stem_map[stem].append(path)

    # ----------------------------------------------------------
    # Match inventory names to source files
    # ----------------------------------------------------------

    for workbook in workbook_inventory:

        filename = workbook["Workbook"]

        matches = source_by_name.get(
            filename.casefold(),
            []
        )

        if not matches:

            result.append({

                "Workbook": filename,

                "Filename_Stem": normalized_stem(filename),

                "Source_File_Found": "No exact filename match",

                "SHA256": "",

                "Exact_File_Duplicate": "Unknown",

                "Filename_Similarity": "Review required",

                "Scientific_Duplication": "Not established",

                "Comments": (
                    "Check original source folder manually."
                )

            })

            continue

        for path in matches:

            digest = hash_map.get(
                str(path),
                ""
            )

            same_hash = (

                bool(digest)

                and

                sum(

                    1

                    for other_hash in hash_map.values()

                    if other_hash == digest

                ) > 1

            )

            same_stem_files = stem_map[
                normalized_stem(path.name)
            ]

            if same_hash:

                duplicate_status = (
                    "Exact byte-level match exists"
                )

            else:

                duplicate_status = (
                    "No exact byte-level match detected"
                )

            if len(same_stem_files) > 1:

                similarity = (
                    "Same filename stem; review versions"
                )

            else:

                similarity = (
                    "No same-stem file detected"
                )

            result.append({

                "Workbook": filename,

                "Filename_Stem": normalized_stem(path.name),

                "Source_File_Found": "Yes",

                "SHA256": digest,

                "Exact_File_Duplicate": duplicate_status,

                "Filename_Similarity": similarity,

                "Scientific_Duplication": (
                    "Not established by file hash or filename"
                ),

                "Comments": str(path)

            })

    return result


# ==============================================================
# 18. EXCEL WORKSHEET FORMATTING
# ==============================================================


def format_worksheet(
    worksheet,
    title,
    subtitle,
    header_row=4
):

    worksheet.sheet_view.showGridLines = False

    worksheet.freeze_panes = "A{}".format(
        header_row + 1
    )

    # ----------------------------------------------------------
    # TITLE ROW
    # ----------------------------------------------------------

    last_column = max(
        worksheet.max_column,
        1
    )

    worksheet.merge_cells(
        start_row=1,
        start_column=1,
        end_row=1,
        end_column=last_column
    )

    title_cell = worksheet.cell(1, 1)

    title_cell.value = title

    title_cell.font = Font(
        name=FONT_NAME,
        size=TITLE_SIZE,
        bold=True,
        color=WHITE
    )

    title_cell.fill = PatternFill(
        fill_type="solid",
        fgColor=DARK_BLUE
    )

    title_cell.alignment = Alignment(
        vertical="center"
    )

    worksheet.row_dimensions[1].height = 32

    # ----------------------------------------------------------
    # SUBTITLE ROW
    # ----------------------------------------------------------

    worksheet.merge_cells(
        start_row=2,
        start_column=1,
        end_row=2,
        end_column=last_column
    )

    subtitle_cell = worksheet.cell(2, 1)

    subtitle_cell.value = subtitle

    subtitle_cell.font = Font(
        name=FONT_NAME,
        size=SUBTITLE_SIZE,
        italic=True,
        color=DARK_GREY
    )

    subtitle_cell.alignment = Alignment(
        vertical="center",
        wrap_text=True
    )

    worksheet.row_dimensions[2].height = 30

    # ----------------------------------------------------------
    # HEADER FORMATTING
    # ----------------------------------------------------------

    for cell in worksheet[header_row]:

        cell.font = Font(
            name=FONT_NAME,
            size=HEADER_SIZE,
            bold=True,
            color=WHITE
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor=MEDIUM_BLUE
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        cell.border = Border(
            bottom=Side(
                style="thin",
                color=WHITE
            )
        )

    worksheet.row_dimensions[
        header_row
    ].height = 38

    # ----------------------------------------------------------
    # BODY FORMATTING
    # ----------------------------------------------------------

    for row in worksheet.iter_rows(
        min_row=header_row + 1
    ):

        for cell in row:

            cell.font = Font(
                name=FONT_NAME,
                size=BODY_SIZE,
                color="000000"
            )

            cell.alignment = Alignment(
                vertical="top",
                wrap_text=True
            )

            if cell.row % 2 == 0:

                cell.fill = PatternFill(
                    fill_type="solid",
                    fgColor="F5F8FC"
                )

    # ----------------------------------------------------------
    # FILTERS
    # ----------------------------------------------------------

    if worksheet.max_row >= header_row:

        worksheet.auto_filter.ref = (
            "A{}:{}{}".format(
                header_row,
                get_column_letter(last_column),
                worksheet.max_row
            )
        )

    # ----------------------------------------------------------
    # SAFE COLUMN WIDTH FORMATTING
    # ----------------------------------------------------------

    for column_number in range(
        1,
        last_column + 1
    ):

        column_letter = get_column_letter(
            column_number
        )

        maximum = 0

        for row_number in range(
            header_row,
            worksheet.max_row + 1
        ):

            cell = worksheet.cell(
                row=row_number,
                column=column_number
            )

            # Ignore merged cells.
            if cell.__class__.__name__ == "MergedCell":
                continue

            if cell.value is not None:

                maximum = max(
                    maximum,
                    min(
                        len(str(cell.value)),
                        45
                    )
                )

        worksheet.column_dimensions[
            column_letter
        ].width = min(
            max(maximum + 3, 15),
            42
        )


# ==============================================================
# 19. WRITE DATAFRAME-LIKE RECORDS
# ==============================================================

def write_records(

    workbook,

    sheet_name,

    records,

    title,

    subtitle

):

    if not records:

        records = [

            {

                "Information": "No records available"

            }

        ]

    headers = list(
        records[0].keys()
    )

    worksheet = workbook.create_sheet(
        sheet_name
    )

    # Title row.

    worksheet.append([
        title
    ])

    # Subtitle row.

    worksheet.append([
        subtitle
    ])

    # Blank row.

    worksheet.append([])

    # Header row.

    worksheet.append(headers)

    # Data rows.

    for record in records:

        values = [

            safe_value(record.get(header, ""))

            for header in headers

        ]

        worksheet.append(values)

    format_worksheet(

        worksheet,

        title,

        subtitle,

        header_row=4

    )

    return worksheet


# ==============================================================
# 20. ADD STATUS DROPDOWNS
# ==============================================================

def add_status_validation(

    worksheet,

    column_name,

    header_row=4

):

    headers = [

        cell.value

        for cell in worksheet[header_row]

    ]

    if column_name not in headers:

        return

    column_number = headers.index(
        column_name
    ) + 1

    column_letter = get_column_letter(
        column_number
    )

    validation = DataValidation(

        type="list",

        formula1='"Pending,In progress,Verified,Not applicable,Unresolved"',

        allow_blank=True

    )

    worksheet.add_data_validation(
        validation
    )

    validation.add(

        "{}5:{}{}".format(

            column_letter,

            column_letter,

            max(worksheet.max_row, 5)

        )

    )


# ==============================================================
# 21. READ ME WORKSHEET
# ==============================================================

def build_readme(

    workbook_count,

    worksheet_count,

    total_formula_count

):

    return [

        {

            "Parameter": "Project",

            "Value": "Geothermobarometer Machine Learning Research"

        },

        {

            "Parameter": "Step",

            "Value": "04 — Master Calibration Database"

        },

        {

            "Parameter": "Source report",

            "Value": INPUT_FILE.name

        },

        {

            "Parameter": "Source workbooks",

            "Value": workbook_count

        },

        {

            "Parameter": "Formula-bearing worksheets",

            "Value": worksheet_count

        },

        {

            "Parameter": "Formula records represented",

            "Value": total_formula_count

        },

        {

            "Parameter": "Database generation date",

            "Value": datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            )

        },

        {

            "Parameter": "Important limitation",

            "Value": (
                "Workbook-level records are provisional candidates, "
                "not verified unique scientific calibrations."
            )

        },

        {

            "Parameter": "DOI policy",

            "Value": (
                "DOIs must be linked to the correct calibration "
                "using the primary publication."
            )

        },

        {

            "Parameter": "Equation policy",

            "Value": (
                "Only equations checked against the original "
                "publication may be marked verified."
            )

        },

        {

            "Parameter": "ML readiness",

            "Value": (
                "This database is a calibration metadata register. "
                "It is not yet the final ML training dataset."
            )

        },

        {

            "Parameter": "Original workbook preservation",

            "Value": (
                "No original calibration workbook was modified."
            )

        }

    ]


# ==============================================================
# 22. MAIN FUNCTION
# ==============================================================

def main():

    log("=" * 75)

    log(
        "STEP 04 — MASTER GEOTHERMOBAROMETER CALIBRATION DATABASE"
    )

    log("=" * 75)

    # ----------------------------------------------------------
    # A. Load Step 03C.1 inventory
    # ----------------------------------------------------------

    worksheet_records = load_source_inventory()

    # ----------------------------------------------------------
    # B. Build workbook inventory
    # ----------------------------------------------------------

    workbook_inventory = build_workbook_inventory(
        worksheet_records
    )

    # ----------------------------------------------------------
    # C. Build worksheet inventory
    # ----------------------------------------------------------

    worksheet_inventory = build_worksheet_inventory(
        worksheet_records
    )

    # ----------------------------------------------------------
    # D. Build calibration register
    # ----------------------------------------------------------

    calibration_register = build_calibration_register(
        workbook_inventory
    )

    # ----------------------------------------------------------
    # E. Build publication register
    # ----------------------------------------------------------

    publication_register = build_publication_register(
        calibration_register
    )

    # ----------------------------------------------------------
    # F. Build equation register
    # ----------------------------------------------------------

    equation_register = build_equation_register(
        calibration_register
    )

    # ----------------------------------------------------------
    # G. Build mineral assemblage register
    # ----------------------------------------------------------

    mineral_register = build_mineral_register(
        calibration_register
    )

    # ----------------------------------------------------------
    # H. Build input variable register
    # ----------------------------------------------------------

    input_register = build_input_register(
        calibration_register
    )

    # ----------------------------------------------------------
    # I. Build calibration range register
    # ----------------------------------------------------------

    range_register = build_range_register(
        calibration_register
    )

    # ----------------------------------------------------------
    # J. Build Excel output mapping
    # ----------------------------------------------------------

    output_mapping = build_output_mapping(
        calibration_register
    )

    # ----------------------------------------------------------
    # K. Build verification register
    # ----------------------------------------------------------

    verification_register = build_verification_register(
        calibration_register
    )

    # ----------------------------------------------------------
    # L. Duplicate assessment
    # ----------------------------------------------------------

    duplicate_assessment = build_duplicate_assessment(
        workbook_inventory
    )

    # ----------------------------------------------------------
    # M. Create output workbook
    # ----------------------------------------------------------

    output_wb = Workbook()

    # Remove default sheet.

    default_sheet = output_wb.active

    output_wb.remove(default_sheet)

    # ----------------------------------------------------------
    # N. Create all worksheets
    # ----------------------------------------------------------

    write_records(

        output_wb,

        "Read_Me",

        build_readme(

            len(workbook_inventory),

            len(worksheet_inventory),

            sum(

                int(
                    record["Total_Formula_Count"]
                )

                for record in workbook_inventory

            )

        ),

        "STEP 04 — MASTER CALIBRATION DATABASE",

        "Project documentation and database overview."

    )

    write_records(

        output_wb,

        "Workbook_Inventory",

        workbook_inventory,

        "WORKBOOK INVENTORY",

        "Original source workbook inventory. Workbook count is not calibration count."

    )

    write_records(

        output_wb,

        "Worksheet_Inventory",

        worksheet_inventory,

        "WORKSHEET INVENTORY",

        "Formula-bearing worksheets identified in Step 03C.1."

    )

    write_records(

        output_wb,

        "Calibration_Register",

        calibration_register,

        "MASTER CALIBRATION REGISTER",

        "One provisional record per source workbook. Scientific identity requires publication verification."

    )

    write_records(

        output_wb,

        "Publication_Register",

        publication_register,

        "PUBLICATION AND DOI REGISTER",

        "Enter primary publication information after verifying method-to-paper correspondence."

    )

    write_records(

        output_wb,

        "Equation_Register",

        equation_register,

        "PUBLISHED EQUATION REGISTER",

        "Record published equations and independently reproduced Python formulations."

    )

    write_records(

        output_wb,

        "Mineral_Assemblage",

        mineral_register,

        "MINERAL ASSEMBLAGE REGISTER",

        "Record the exact mineral requirements from each primary calibration publication."

    )

    write_records(

        output_wb,

        "Input_Variables",

        input_register,

        "CHEMICAL INPUT VARIABLE REGISTER",

        "Record mineral-specific chemical inputs, units and normalization procedures."

    )

    write_records(

        output_wb,

        "Calibration_Ranges",

        range_register,

        "CALIBRATION APPLICABILITY RANGES",

        "Record published temperature, pressure and compositional applicability conditions."

    )

    write_records(

        output_wb,

        "Excel_Output_Mapping",

        output_mapping,

        "EXCEL OUTPUT CELL MAPPING",

        "Final temperature and pressure cells must be independently verified."

    )

    write_records(

        output_wb,

        "Verification_Status",

        verification_register,

        "SCIENTIFIC VERIFICATION TRACKER",

        "Track the evidence supporting each calibration and its implementation."

    )

    write_records(

        output_wb,

        "Duplicate_Assessment",

        duplicate_assessment,

        "POSSIBLE DUPLICATE WORKBOOK ASSESSMENT",

        "Byte-level duplication and filename similarity do not establish scientific calibration identity."

    )

    # ----------------------------------------------------------
    # O. Add status dropdowns
    # ----------------------------------------------------------

    add_status_validation(

        output_wb["Calibration_Register"],

        "Scientific_Verification_Status"

    )

    add_status_validation(

        output_wb["Equation_Register"],

        "Equation_Verification_Status"

    )

    add_status_validation(

        output_wb["Verification_Status"],

        "Status"

    )

    # ----------------------------------------------------------
    # P. Workbook properties
    # ----------------------------------------------------------

    output_wb.properties.title = (
        "Master Geothermobarometer Calibration Database"
    )

    output_wb.properties.subject = (
        "Calibration inventory and scientific verification"
    )

    output_wb.properties.creator = (
        "Geothermobarometer ML Research Project"
    )

    output_wb.properties.description = (
        "Provisional calibration metadata register generated "
        "from the corrected Step 03C.1 workbook inventory."
    )

    # ----------------------------------------------------------
    # Q. Save workbook
    # ----------------------------------------------------------

    log("Saving master calibration database...")

    output_wb.save(
        OUTPUT_FILE
    )

    output_wb.close()

    # ----------------------------------------------------------
    # R. Final execution report
    # ----------------------------------------------------------

    total_formulas = sum(

        int(
            record["Total_Formula_Count"]
        )

        for record in workbook_inventory

    )

    log("=" * 75)

    log(
        "STEP 04 COMPLETED SUCCESSFULLY"
    )

    log(
        "Source workbooks: {}".format(
            len(workbook_inventory)
        )
    )

    log(
        "Formula-bearing worksheets: {}".format(
            len(worksheet_inventory)
        )
    )

    log(
        "Formula records represented: {}".format(
            total_formulas
        )
    )

    log(
        "Provisional calibration candidates: {}".format(
            len(calibration_register)
        )
    )

    log(
        "Publication records: {}".format(
            len(publication_register)
        )
    )

    log(
        "Equation records: {}".format(
            len(equation_register)
        )
    )

    log(
        "Verification tasks: {}".format(
            len(verification_register)
        )
    )

    log(
        "Excel output: {}".format(
            OUTPUT_FILE
        )
    )

    log(
        "Execution log: {}".format(
            LOG_FILE
        )
    )

    log("=" * 75)

    print("\nSTEP 04 COMPLETED SUCCESSFULLY")

    print("\nSource workbooks:", len(workbook_inventory))

    print(
        "Formula-bearing worksheets:",
        len(worksheet_inventory)
    )

    print(
        "Formula records:",
        total_formulas
    )

    print(
        "Provisional calibration candidates:",
        len(calibration_register)
    )

    print("\nGenerated Excel database:")

    print(OUTPUT_FILE)

    print("\nExecution log:")

    print(LOG_FILE)

    print(
        "\nIMPORTANT: Calibration identities and equations "
        "remain pending scientific verification."
    )


# ==============================================================
# 23. EXECUTION
# ==============================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logger.exception(
            "STEP 04 FAILED: {}".format(error)
        )

        print("\nSTEP 04 FAILED.")

        print(
            "Please provide the complete Spyder console error."
        )

        raise
