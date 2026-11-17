# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 15:56:06 2026

@author: hussa
"""


# ============================================================================
# STEP 05C.4 — Hbl97 AMPHIBOLE FORMULA INPUT AUDIT
#
# Phase 1: Inspect original chemical inputs and reconstructed chemistry.
#
# This is an input audit, NOT yet a complete IMA amphibole recalculation.
#
# Original files remain unchanged.
# ============================================================================

from pathlib import Path
from datetime import datetime
import logging
import hashlib
import warnings

import xlrd
import pandas as pd
import numpy as np

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# ============================================================================
# 1. PROJECT SETTINGS
# ============================================================================

PROJECT = Path(
    r"D:\Data Science Assignment\My Projects\Geothermobarometer_ML_Project"
)

SOURCE_XLS = (
    PROJECT
    / "Geothermobarometers Calibrations"
    / "Hbl P-T (1997).xls"
)

STEP05C1 = (
    PROJECT
    / "05C1_Hbl97_independent_chemical_reconstruction"
    / "Step05C1_Hbl97_Independent_Chemical_Reconstruction.xlsx"
)

STEP05C3 = (
    PROJECT
    / "05C3_Hbl97_independent_iterative_verification"
    / "Step05C3_Hbl97_Independent_Iterative_Verification.xlsx"
)

OUTPUT_DIR = (
    PROJECT
    / "05C4_Hbl97_amphibole_formula_input_audit"
)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

REPORT = (
    OUTPUT_DIR
    / "Step05C4_Hbl97_Amphibole_Input_Audit.xlsx"
)

LOG_FILE = LOG_DIR / "step05C4_execution_log.txt"


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
    ],
    force=True
)

logger = logging.getLogger("Hbl97_Chemical_Audit")


# ============================================================================
# 3. UTILITIES
# ============================================================================

def sha256_file(path):

    h = hashlib.sha256()

    with open(path, "rb") as f:

        for chunk in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):

            h.update(chunk)

    return h.hexdigest()


def safe_value(value):

    if value is None:
        return ""

    if isinstance(value, float):

        if not np.isfinite(value):
            return ""

    return value


def clean_text(value):

    if value is None:
        return ""

    return str(value).strip()


# ============================================================================
# 4. CHECK INPUT FILES
# ============================================================================

def check_inputs():

    files = [
        SOURCE_XLS,
        STEP05C1,
        STEP05C3
    ]

    records = []

    for path in files:

        exists = path.exists()

        records.append({

            "File": path.name,

            "Path": str(path),

            "Exists": exists,

            "Size_MB": (
                round(path.stat().st_size / 1024**2, 3)
                if exists else None
            ),

            "SHA256": (
                sha256_file(path)
                if exists else ""
            )

        })

        if not exists:

            logger.warning(
                "Input file unavailable: %s",
                path
            )

    return records


# ============================================================================
# 5. INSPECT ORIGINAL XLS WORKBOOK
# ============================================================================

def inspect_original_workbook():

    logger.info("Opening original Hbl97 XLS workbook.")

    book = xlrd.open_workbook(
        str(SOURCE_XLS),
        on_demand=True
    )

    workbook_inventory = []

    cell_inventory = []

    candidate_headers = []

    for sheet_index in range(book.nsheets):

        sheet = book.sheet_by_index(sheet_index)

        logger.info(
            "Inspecting worksheet: %s",
            sheet.name
        )

        workbook_inventory.append({

            "Sheet_Index": sheet_index,

            "Worksheet": sheet.name,

            "Rows": sheet.nrows,

            "Columns": sheet.ncols

        })

        for row_index in range(sheet.nrows):

            nonempty = []

            for col_index in range(sheet.ncols):

                value = sheet.cell_value(
                    row_index,
                    col_index
                )

                if value not in ("", None):

                    text = clean_text(value)

                    nonempty.append(
                        f"{xlrd.formula.colname(col_index)}"
                        f"{row_index + 1}={text}"
                    )

                    cell_inventory.append({

                        "Worksheet": sheet.name,

                        "Excel_Row": row_index + 1,

                        "Excel_Column": (
                            xlrd.formula.colname(col_index)
                        ),

                        "Cell": (
                            f"{xlrd.formula.colname(col_index)}"
                            f"{row_index + 1}"
                        ),

                        "Value": safe_value(value),

                        "Cell_Type": sheet.cell_type(
                            row_index,
                            col_index
                        )

                    })

            if nonempty:

                # Preserve rows that may contain chemical labels.
                candidate_headers.append({

                    "Worksheet": sheet.name,

                    "Excel_Row": row_index + 1,

                    "Nonempty_Cell_Count": len(nonempty),

                    "Row_Content": " | ".join(nonempty[:30])

                })

    book.release_resources()

    return (
        workbook_inventory,
        cell_inventory,
        candidate_headers
    )


# ============================================================================
# 6. INSPECT STEP 05C1 RECONSTRUCTION
# ============================================================================

def inspect_reconstruction():

    logger.info("Inspecting Step 05C1 report.")

    wb = load_workbook(
        STEP05C1,
        read_only=True,
        data_only=True
    )

    sheet_records = []

    header_records = []

    sample_records = []

    for ws in wb.worksheets:

        logger.info(
            "Inspecting Step05C1 worksheet: %s",
            ws.title
        )

        sheet_records.append({

            "Worksheet": ws.title,

            "Rows": ws.max_row,

            "Columns": ws.max_column

        })

        # Record all nonempty rows near the beginning.
        for row_index, row in enumerate(
            ws.iter_rows(
                min_row=1,
                max_row=min(ws.max_row, 15),
                values_only=True
            ),
            start=1
        ):

            content = []

            for col_index, value in enumerate(row):

                if value not in ("", None):

                    content.append({

                        "Worksheet": ws.title,

                        "Excel_Row": row_index,

                        "Excel_Column": (
                            xlrd.formula.colname(col_index)
                        ),

                        "Value": safe_value(value)

                    })

            header_records.extend(content)

        # Extract Sample_Chemistry records without assuming
        # any particular chemical column names.
        if ws.title == "Sample_Chemistry":

            rows = list(
                ws.iter_rows(values_only=True)
            )

            for row_index, row in enumerate(
                rows,
                start=1
            ):

                sample_records.append({

                    "Excel_Row": row_index,

                    "Values": " | ".join(
                        clean_text(v)
                        for v in row
                        if v not in ("", None)
                    )

                })

    wb.close()

    return (
        sheet_records,
        header_records,
        sample_records
    )


# ============================================================================
# 7. IDENTIFY POTENTIAL CHEMICAL VARIABLES
# ============================================================================

CHEMICAL_TERMS = [

    "SIO2",
    "TIO2",
    "AL2O3",
    "FEO",
    "FEOT",
    "FE2O3",
    "MNO",
    "MGO",
    "CAO",
    "NA2O",
    "K2O",
    "CR2O3",
    "NIO",
    "H2O",
    "F",
    "CL",

    "SI",
    "TI",
    "AL",
    "FE2",
    "FE3",
    "FEOT",
    "MN",
    "MG",
    "CA",
    "NA",
    "K",

    "OXIDE",
    "CATION",
    "FORMULA",
    "NORMALIZATION",
    "ANALYSIS",
    "SAMPLE",
    "CHARGE",
    "SITE",
    "APFU"

]


def identify_chemical_fields(records):

    output = []

    for record in records:

        text = str(
            record.get("Value", "")
        ).upper().replace(" ", "")

        matched = []

        for term in CHEMICAL_TERMS:

            if term in text:

                matched.append(term)

        if matched:

            output.append({

                **record,

                "Potential_Chemical_Terms": ", ".join(
                    matched
                )

            })

    return output


# ============================================================================
# 8. INSPECT STEP 05C3 FORMULA DOCUMENTATION
# ============================================================================

def inspect_step05c3():

    if not STEP05C3.exists():

        return []

    wb = load_workbook(
        STEP05C3,
        read_only=True,
        data_only=True
    )

    output = []

    for sheet_name in wb.sheetnames:

        ws = wb[sheet_name]

        if sheet_name not in [
            "Formula_Documentation",
            "Chemical_Consistency",
            "Method_and_Limitations"
        ]:

            continue

        for row_index, row in enumerate(
            ws.iter_rows(values_only=True),
            start=1
        ):

            values = [
                safe_value(v)
                for v in row
            ]

            output.append({

                "Worksheet": sheet_name,

                "Excel_Row": row_index,

                "Content": " | ".join(
                    clean_text(v)
                    for v in values
                    if v not in ("", None)
                )

            })

    wb.close()

    return output


# ============================================================================
# 9. GENERATE AUDIT REPORT
# ============================================================================

def generate_report():

    logger.info("=" * 75)

    logger.info(
        "STEP 05C.4 — Hbl97 AMPHIBOLE INPUT AUDIT"
    )

    logger.info("=" * 75)

    file_records = check_inputs()

    if not SOURCE_XLS.exists():

        raise FileNotFoundError(SOURCE_XLS)

    if not STEP05C1.exists():

        raise FileNotFoundError(STEP05C1)

    (
        original_sheets,
        original_cells,
        original_rows
    ) = inspect_original_workbook()

    (
        reconstruction_sheets,
        reconstruction_headers,
        reconstruction_rows
    ) = inspect_reconstruction()

    chemical_fields = identify_chemical_fields(
        reconstruction_headers
    )

    formula_documentation = inspect_step05c3()

    original_df = pd.DataFrame(original_cells)

    if len(original_df):

        original_chemical_fields = (
            identify_chemical_fields(original_cells)
        )

    else:

        original_chemical_fields = []

    # --------------------------------------------------------
    # Required input availability assessment
    # --------------------------------------------------------

    required_inputs = [

        "SiO2",

        "TiO2",

        "Al2O3",

        "FeO / total Fe",

        "MnO",

        "MgO",

        "CaO",

        "Na2O",

        "K2O",

        "Cr2O3 (if measured)",

        "NiO (if measured)",

        "H2O (if measured)",

        "F (if measured)",

        "Cl (if measured)",

        "Independent Fe3+/FeTotal measurement (if available)"

    ]

    availability_records = []

    original_text = " ".join(
        clean_text(r["Value"]).upper()
        for r in original_cells
    )

    for variable in required_inputs:

        normalized = variable.upper()

        if variable == "FeO / total Fe":

            found = (
                "FEO" in original_text
                or "FEOT" in original_text
            )

        elif variable.startswith("Independent Fe3+"):

            found = (
                "MOSSBAUER" in original_text
                or "FE3+/FETOTAL" in original_text
            )

        elif "if measured" in variable:

            found = (
                normalized.split(" ")[0].replace(
                    "(IF", ""
                ) in original_text
            )

        else:

            found = normalized.replace(
                " ", ""
            ) in original_text

        availability_records.append({

            "Chemical_Input": variable,

            "Potentially_Present": found,

            "Assessment": (
                "Potentially present; verify cell meaning"
                if found
                else "Not identified by automatic text search"
            )

        })

    # --------------------------------------------------------
    # Save Excel report
    # --------------------------------------------------------

    logger.info("Writing audit workbook.")

    with pd.ExcelWriter(
        REPORT,
        engine="openpyxl"
    ) as writer:

        pd.DataFrame(file_records).to_excel(
            writer,
            sheet_name="Input_File_Inventory",
            index=False
        )

        pd.DataFrame(original_sheets).to_excel(
            writer,
            sheet_name="Original_Workbook_Sheets",
            index=False
        )

        pd.DataFrame(original_rows).to_excel(
            writer,
            sheet_name="Original_Row_Inventory",
            index=False
        )

        pd.DataFrame(original_cells).to_excel(
            writer,
            sheet_name="Original_Cell_Inventory",
            index=False
        )

        pd.DataFrame(reconstruction_sheets).to_excel(
            writer,
            sheet_name="Step05C1_Sheets",
            index=False
        )

        pd.DataFrame(reconstruction_headers).to_excel(
            writer,
            sheet_name="Step05C1_Headers",
            index=False
        )

        pd.DataFrame(reconstruction_rows).to_excel(
            writer,
            sheet_name="Step05C1_Rows",
            index=False
        )

        pd.DataFrame(original_chemical_fields).to_excel(
            writer,
            sheet_name="Original_Chemical_Fields",
            index=False
        )

        pd.DataFrame(chemical_fields).to_excel(
            writer,
            sheet_name="Reconstructed_Chemical_Fields",
            index=False
        )

        pd.DataFrame(availability_records).to_excel(
            writer,
            sheet_name="Chemical_Input_Availability",
            index=False
        )

        pd.DataFrame(formula_documentation).to_excel(
            writer,
            sheet_name="Step05C3_Method_Review",
            index=False
        )

    # --------------------------------------------------------
    # Formatting
    # --------------------------------------------------------

    wb = load_workbook(REPORT)

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        if ws.max_row > 1:

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
                wrap_text=True,
                vertical="center"
            )

        ws.row_dimensions[1].height = 30

        for column_cells in ws.iter_cols():

            col_idx = column_cells[0].column

            maximum = max(
                (
                    len(str(c.value))
                    for c in column_cells
                    if c.value is not None
                ),
                default=12
            )

            ws.column_dimensions[
                get_column_letter(col_idx)
            ].width = min(
                max(maximum + 2, 12),
                45
            )

    wb.save(REPORT)

    # --------------------------------------------------------
    # Completion
    # --------------------------------------------------------

    logger.info("=" * 75)

    logger.info("STEP 05C.4 INPUT AUDIT COMPLETED")

    logger.info(
        "Original worksheets: %s",
        len(original_sheets)
    )

    logger.info(
        "Original cells inventoried: %s",
        len(original_cells)
    )

    logger.info(
        "Original chemical candidates: %s",
        len(original_chemical_fields)
    )

    logger.info(
        "Reconstructed chemical fields: %s",
        len(chemical_fields)
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

    print("\nSTEP 05C.4 INPUT AUDIT COMPLETED")

    print("\nExcel report:")
    print(REPORT)

    print("\nExecution log:")
    print(LOG_FILE)

    print(
        "\nIMPORTANT: This is an input inventory only. "
        "No IMA site allocation or Fe3+ determination "
        "has been certified."
    )


# ============================================================================
# 10. EXECUTION
# ============================================================================

if __name__ == "__main__":

    try:

        generate_report()

    except Exception as exc:

        logger.exception(
            "STEP 05C.4 FAILED: %s",
            exc
        )

        raise
