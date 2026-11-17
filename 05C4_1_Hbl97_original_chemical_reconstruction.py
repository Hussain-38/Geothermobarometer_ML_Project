# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 16:00:31 2026

@author: hussa
"""


# ============================================================================
# STEP 05C.4.1 — Hbl97 ORIGINAL CHEMICAL RECONSTRUCTION
#
# Objectives:
# 1. Extract original oxide analyses from the Hbl97 XLS workbook.
# 2. Preserve the original Fe3+ (Hbl) field separately.
# 3. Independently calculate oxide mole and cation proportions.
# 4. Calculate oxygen proportions and alternative normalization factors.
# 5. Compare reconstructed cations with original workbook calculations.
# 6. Audit the f1-f10, fav and revised cation calculation blocks.
# 7. Preserve original workbook values and identify discrepancies.
#
# IMPORTANT:
# This script does not certify an amphibole structural formula.
# It does not assume that the workbook Fe3+ field is independently measured.
# It does not silently assign crystallographic sites.
# ============================================================================

from pathlib import Path
from datetime import datetime
import logging
import hashlib
import math

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

OUTPUT_DIR = (
    PROJECT
    / "05C4_1_Hbl97_original_chemical_reconstruction"
)

LOG_DIR = OUTPUT_DIR / "logs"

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

REPORT = (
    OUTPUT_DIR
    / "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

LOG_FILE = (
    LOG_DIR
    / "step05C4_1_execution_log.txt"
)


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

logger = logging.getLogger("Hbl97_Reconstruction")


# ============================================================================
# 3. ORIGINAL WORKBOOK STRUCTURE
# ============================================================================

SHEET_INDEX = 0

FIRST_SAMPLE_ROW = 9       # Excel row 10, zero-based index
LAST_SAMPLE_ROW = 21       # Excel row 22, zero-based index

# Original analytical oxide columns.
OXIDE_COLUMNS = {

    "SiO2": "J",
    "TiO2": "K",
    "Al2O3": "L",
    "FeO_total": "M",
    "MnO": "N",
    "MgO": "O",
    "CaO": "P",
    "Na2O": "Q",
    "K2O": "R"

}

# Original workbook Fe3+ field.
FE3_FIELD_COLUMN = "H"

# Original workbook calculation blocks.
ORIGINAL_CATION_COLUMNS = {

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

ORIGINAL_REVISED_COLUMNS = {

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

ORIGINAL_MN13_COLUMNS = {

    "Si": "BH",
    "Ti": "BI",
    "Al": "BJ",
    "Fe2": "BK",
    "Fe3": "BL",
    "Mn": "BM",
    "Mg": "BN"

}

# Oxide molecular weights (g/mol).
# Standard atomic-weight-derived values.
OXIDE_DATA = {

    "SiO2": {
        "MW": 60.0843,
        "Cations": 1,
        "Oxygen": 2
    },

    "TiO2": {
        "MW": 79.866,
        "Cations": 1,
        "Oxygen": 2
    },

    "Al2O3": {
        "MW": 101.9613,
        "Cations": 2,
        "Oxygen": 3
    },

    "FeO_total": {
        "MW": 71.844,
        "Cations": 1,
        "Oxygen": 1
    },

    "MnO": {
        "MW": 70.9374,
        "Cations": 1,
        "Oxygen": 1
    },

    "MgO": {
        "MW": 40.3044,
        "Cations": 1,
        "Oxygen": 1
    },

    "CaO": {
        "MW": 56.0774,
        "Cations": 1,
        "Oxygen": 1
    },

    "Na2O": {
        "MW": 61.9789,
        "Cations": 2,
        "Oxygen": 1
    },

    "K2O": {
        "MW": 94.196,
        "Cations": 2,
        "Oxygen": 1
    }

}


# ============================================================================
# 4. UTILITIES
# ============================================================================

def excel_col_number(letter):

    return xlrd.formula.colname(
        0
    ) if False else xlrd.formula.colname(
        0
    )


def column_index(letter):

    result = 0

    for char in letter.upper():

        result = (
            result * 26
            + ord(char) - ord("A") + 1
        )

    return result - 1


def get_value(sheet, row, column):

    return sheet.cell_value(
        row,
        column_index(column)
    )


def numeric(value):

    if value in ("", None):

        return np.nan

    try:

        value = float(value)

        if not np.isfinite(value):

            return np.nan

        return value

    except (ValueError, TypeError):

        return np.nan


def sha256_file(path):

    digest = hashlib.sha256()

    with open(path, "rb") as file:

        for chunk in iter(
            lambda: file.read(1024 * 1024),
            b""
        ):

            digest.update(chunk)

    return digest.hexdigest()


def make_safe_dataframe(records):

    df = pd.DataFrame(records)

    if len(df.columns) == 0:

        return pd.DataFrame({
            "Information": ["No records available"]
        })

    return df


# ============================================================================
# 5. EXTRACT ORIGINAL SAMPLE RECORDS
# ============================================================================

def extract_samples(sheet):

    logger.info(
        "Extracting original chemical analyses..."
    )

    records = []

    for row in range(
        FIRST_SAMPLE_ROW,
        LAST_SAMPLE_ROW + 1
    ):

        sample = get_value(
            sheet,
            row,
            "B"
        )

        reference = get_value(
            sheet,
            row,
            "A"
        )

        if not str(sample).strip():

            continue

        record = {

            "Excel_Row": row + 1,

            "Sample": str(sample).strip(),

            "Reference": reference,

            "Fe3_Hbl_Original": numeric(
                get_value(
                    sheet,
                    row,
                    FE3_FIELD_COLUMN
                )
            )

        }

        for oxide, column in OXIDE_COLUMNS.items():

            record[oxide] = numeric(
                get_value(
                    sheet,
                    row,
                    column
                )
            )

        records.append(record)

    logger.info(
        "Chemical samples extracted: %s",
        len(records)
    )

    return records


# ============================================================================
# 6. INDEPENDENT OXIDE-TO-CATION RECONSTRUCTION
# ============================================================================

def reconstruct_sample(record):

    result = {

        "Sample": record["Sample"],

        "Excel_Row": record["Excel_Row"],

        "Reference": record["Reference"],

        "Fe3_Hbl_Original": record[
            "Fe3_Hbl_Original"
        ]

    }

    total_oxide = 0.0

    total_oxygen = 0.0

    raw_cations = {}

    oxygen_proportions = {}

    # --------------------------------------------------------
    # Calculate mole and cation proportions.
    # --------------------------------------------------------

    for oxide, specification in OXIDE_DATA.items():

        wt = record.get(oxide, np.nan)

        if not np.isfinite(wt):

            result[oxide + "_wt_percent"] = np.nan

            raw_cations[oxide] = np.nan

            oxygen_proportions[oxide] = np.nan

            continue

        if wt < 0:

            raise ValueError(
                "Negative oxide concentration: "
                + record["Sample"]
                + " | "
                + oxide
            )

        mole_proportion = (
            wt / specification["MW"]
        )

        cation_proportion = (
            mole_proportion
            * specification["Cations"]
        )

        oxygen_proportion = (
            mole_proportion
            * specification["Oxygen"]
        )

        result[oxide + "_wt_percent"] = wt

        result[oxide + "_mole_proportion"] = (
            mole_proportion
        )

        result[oxide + "_cation_proportion"] = (
            cation_proportion
        )

        result[oxide + "_oxygen_proportion"] = (
            oxygen_proportion
        )

        raw_cations[oxide] = cation_proportion

        oxygen_proportions[oxide] = oxygen_proportion

        total_oxide += wt

        total_oxygen += oxygen_proportion

    result["Total_Reported_Oxides_wt_percent"] = (
        total_oxide
    )

    result["Unnormalized_Oxygen_Sum"] = (
        total_oxygen
    )

    # --------------------------------------------------------
    # Alternative oxygen normalization factors.
    #
    # These are computational comparisons only.
    # They are NOT automatically accepted as the
    # correct Hbl97 normalization method.
    # --------------------------------------------------------

    if total_oxygen > 0:

        factor_23 = 23.0 / total_oxygen

        factor_24 = 24.0 / total_oxygen

    else:

        factor_23 = np.nan

        factor_24 = np.nan

    result["Factor_23_Oxygen"] = factor_23

    result["Factor_24_Oxygen"] = factor_24

    # --------------------------------------------------------
    # Calculate normalized cation proportions.
    # --------------------------------------------------------

    for oxide, specification in OXIDE_DATA.items():

        raw = raw_cations[oxide]

        if np.isfinite(raw):

            result[
                oxide + "_cation_per_23O"
            ] = raw * factor_23

            result[
                oxide + "_cation_per_24O"
            ] = raw * factor_24

        else:

            result[
                oxide + "_cation_per_23O"
            ] = np.nan

            result[
                oxide + "_cation_per_24O"
            ] = np.nan

    # --------------------------------------------------------
    # Separate total iron from workbook Fe3+.
    #
    # No assumption is made that H is independently measured.
    # --------------------------------------------------------

    fe_total_23 = result.get(
        "FeO_total_cation_per_23O",
        np.nan
    )

    fe3_original = result[
        "Fe3_Hbl_Original"
    ]

    result["Fe_Total_Cation_23O"] = fe_total_23

    result["Fe3_Hbl_Original_Separate"] = (
        fe3_original
    )

    if (
        np.isfinite(fe_total_23)
        and np.isfinite(fe3_original)
    ):

        result["Fe2_Implied_23O"] = (
            fe_total_23 - fe3_original
        )

        result["Fe3_FeTotal_Ratio_Implied"] = (

            fe3_original / fe_total_23

            if fe_total_23 > 0

            else np.nan

        )

    else:

        result["Fe2_Implied_23O"] = np.nan

        result["Fe3_FeTotal_Ratio_Implied"] = np.nan

    # --------------------------------------------------------
    # Total cations for the 23-oxygen scenario.
    # --------------------------------------------------------

    cation_columns = [

        oxide + "_cation_per_23O"

        for oxide in OXIDE_DATA

    ]

    result["Total_Cations_23O"] = sum(

        result.get(column, 0.0)

        for column in cation_columns

        if np.isfinite(
            result.get(column, np.nan)
        )

    )

    result["Total_Cations_24O"] = sum(

        result.get(
            oxide + "_cation_per_24O",
            0.0
        )

        for oxide in OXIDE_DATA

        if np.isfinite(
            result.get(
                oxide + "_cation_per_24O",
                np.nan
            )
        )

    )

    return result


# ============================================================================
# 7. COMPARE WITH ORIGINAL WORKBOOK CATIONS
# ============================================================================

def compare_original_cations(sheet, reconstructed):

    logger.info(
        "Comparing independently reconstructed cations..."
    )

    records = []

    for item in reconstructed:

        row = item["Excel_Row"] - 1

        sample = item["Sample"]

        for element, column in ORIGINAL_CATION_COLUMNS.items():

            original = numeric(
                get_value(
                    sheet,
                    row,
                    column
                )
            )

            calculated = np.nan

            oxide_mapping = {

                "Si": "SiO2",

                "Ti": "TiO2",

                "Al": "Al2O3",

                "Fe_total": "FeO_total",

                "Mn": "MnO",

                "Mg": "MgO",

                "Ca": "CaO",

                "Na": "Na2O",

                "K": "K2O"

            }

            oxide = oxide_mapping[element]

            calculated = item.get(
                oxide + "_cation_per_23O",
                np.nan
            )

            difference = (

                calculated - original

                if (
                    np.isfinite(calculated)
                    and np.isfinite(original)
                )

                else np.nan

            )

            records.append({

                "Sample": sample,

                "Element": element,

                "Original_Workbook_Value": original,

                "Independent_23O_Value": calculated,

                "Difference": difference,

                "Absolute_Difference": (
                    abs(difference)
                    if np.isfinite(difference)
                    else np.nan
                )

            })

    return records


# ============================================================================
# 8. EXTRACT ORIGINAL CALCULATION BLOCKS
# ============================================================================

def extract_original_calculation_blocks(sheet, samples):

    logger.info(
        "Extracting original intermediate calculations..."
    )

    blocks = {

        "Initial_Cations": ORIGINAL_CATION_COLUMNS,

        "Revised_Cations": ORIGINAL_REVISED_COLUMNS,

        "Mn13_Cations": ORIGINAL_MN13_COLUMNS

    }

    output = []

    for sample_record in samples:

        row = sample_record["Excel_Row"] - 1

        for block_name, mapping in blocks.items():

            record = {

                "Sample": sample_record["Sample"],

                "Excel_Row": row + 1,

                "Calculation_Block": block_name

            }

            for element, column in mapping.items():

                record[element] = numeric(
                    get_value(
                        sheet,
                        row,
                        column
                    )
                )

            output.append(record)

    return output


# ============================================================================
# 9. AUDIT f1-f10 AND fav
# ============================================================================

def audit_optimization_factors(sheet, samples):

    logger.info(
        "Extracting f1-f10 optimization factors..."
    )

    factor_columns = {

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

    output = []

    for sample_record in samples:

        row = sample_record["Excel_Row"] - 1

        record = {

            "Sample": sample_record["Sample"],

            "Excel_Row": row + 1

        }

        for name, column in factor_columns.items():

            record[name] = numeric(
                get_value(
                    sheet,
                    row,
                    column
                )
            )

        output.append(record)

    return output


# ============================================================================
# 10. EXTRACT ORIGINAL TEMPERATURE AND PRESSURE
# ============================================================================

def extract_original_PT(sheet, samples):

    output = []

    for sample_record in samples:

        row = sample_record["Excel_Row"] - 1

        output.append({

            "Sample": sample_record["Sample"],

            "Input_T_C": numeric(
                get_value(sheet, row, "D")
            ),

            "Input_P_kbar": numeric(
                get_value(sheet, row, "E")
            ),

            "Calculated_T_C": numeric(
                get_value(sheet, row, "F")
            ),

            "Calculated_P_kbar": numeric(
                get_value(sheet, row, "G")
            ),

            "T97_C": numeric(
                get_value(sheet, row, "BO")
            ),

            "P97_kbar": numeric(
                get_value(sheet, row, "BP")
            )

        })

    return output


# ============================================================================
# 11. CHEMICAL INPUT QUALITY CONTROL
# ============================================================================

def create_quality_control(records):

    output = []

    for record in records:

        missing = []

        negative = []

        for oxide in OXIDE_COLUMNS:

            value = record.get(oxide, np.nan)

            if not np.isfinite(value):

                missing.append(oxide)

            elif value < 0:

                negative.append(oxide)

        total = sum(

            record.get(oxide, 0.0)

            for oxide in OXIDE_COLUMNS

            if np.isfinite(
                record.get(oxide, np.nan)
            )

        )

        output.append({

            "Sample": record["Sample"],

            "Missing_Oxides": ", ".join(missing),

            "Negative_Oxides": ", ".join(negative),

            "Total_Oxides_wt_percent": total,

            "Fe3_Hbl_Original": record[
                "Fe3_Hbl_Original"
            ],

            "Fe3_Field_Status": (

                "Available"
                if np.isfinite(
                    record["Fe3_Hbl_Original"]
                )
                else "Missing"

            ),

            "Input_QC": (

                "PASS"
                if (
                    not missing
                    and not negative
                    and total > 0
                )
                else "REVIEW"

            )

        })

    return output


# ============================================================================
# 12. WORKBOOK INTERPRETATION RECORD
# ============================================================================

def create_method_inventory():

    return [

        {

            "Workbook_Feature": "H column",

            "Original_Label": "Fe3+ (Hbl)",

            "Location": "H9",

            "Interpretation": (
                "Original workbook Fe3+ field. "
                "Its derivation must be established."
            ),

            "Scientific_Status": "Unverified origin"

        },

        {

            "Workbook_Feature": "J:R",

            "Original_Label": "Nine major oxide inputs",

            "Location": "J9:R9",

            "Interpretation": (
                "Original amphibole oxide analytical inputs."
            ),

            "Scientific_Status": "Extracted"

        },

        {

            "Workbook_Feature": "T:AD",

            "Original_Label": "Initial cation calculation",

            "Location": "T9:AD9",

            "Interpretation": (
                "Original intermediate cation calculations."
            ),

            "Scientific_Status": "Extracted"

        },

        {

            "Workbook_Feature": "AE:AQ",

            "Original_Label": "f1-f10 and fav",

            "Location": "AE9:AQ9",

            "Interpretation": (
                "Original optimization-related calculation block."
            ),

            "Scientific_Status": "Extracted; equations require review"

        },

        {

            "Workbook_Feature": "AR",

            "Original_Label": "Fe3+ (Hbl)",

            "Location": "AR9",

            "Interpretation": (
                "Intermediate Fe3+ field."
            ),

            "Scientific_Status": "Extracted"

        },

        {

            "Workbook_Feature": "AT:BD",

            "Original_Label": "Revised cations",

            "Location": "AT9:BD9",

            "Interpretation": (
                "Original revised cation calculation block."
            ),

            "Scientific_Status": "Extracted"

        },

        {

            "Workbook_Feature": "BF:BN",

            "Original_Label": "to Sigma13(Mn)",

            "Location": "BF9:BN9",

            "Interpretation": (
                "Original normalization block labelled Sigma13(Mn)."
            ),

            "Scientific_Status": "Extracted; normalization assumptions require review"

        },

        {

            "Workbook_Feature": "BO:BP",

            "Original_Label": "T (97), P (97)",

            "Location": "BO9:BP9",

            "Interpretation": (
                "Original Hbl97 temperature and pressure outputs."
            ),

            "Scientific_Status": "Extracted"

        }

    ]


# ============================================================================
# 13. SAVE EXCEL REPORT
# ============================================================================

def save_report(data):

    logger.info(
        "Generating independent reconstruction report..."
    )

    with pd.ExcelWriter(
        REPORT,
        engine="openpyxl"
    ) as writer:

        for sheet_name, records in data.items():

            df = make_safe_dataframe(records)

            df.to_excel(

                writer,

                sheet_name=sheet_name[:31],

                index=False

            )

    # --------------------------------------------------------
    # Professional formatting
    # --------------------------------------------------------

    wb = load_workbook(REPORT)

    for ws in wb.worksheets:

        ws.freeze_panes = "A2"

        if ws.max_row > 1:

            ws.auto_filter.ref = ws.dimensions

        ws.row_dimensions[1].height = 32

        for cell in ws[1]:

            cell.font = Font(

                bold=True,

                color="FFFFFF"

            )

            cell.fill = PatternFill(

                fill_type="solid",

                fgColor="244062"

            )

            cell.alignment = Alignment(

                horizontal="center",

                vertical="center",

                wrap_text=True

            )

        for column_cells in ws.iter_cols():

            column_number = column_cells[0].column

            maximum = max(

                (

                    len(str(cell.value))

                    for cell in column_cells

                    if cell.value is not None

                ),

                default=12

            )

            ws.column_dimensions[
                get_column_letter(column_number)
            ].width = min(

                max(maximum + 2, 12),

                45

            )

    wb.save(REPORT)


# ============================================================================
# 14. MAIN EXECUTION
# ============================================================================

def main():

    logger.info("=" * 75)

    logger.info(
        "STEP 05C.4.1 — Hbl97 ORIGINAL CHEMICAL RECONSTRUCTION"
    )

    logger.info("=" * 75)

    if not SOURCE.exists():

        raise FileNotFoundError(
            "Original Hbl97 workbook not found: "
            + str(SOURCE)
        )

    logger.info(
        "Source workbook: %s",
        SOURCE
    )

    logger.info(
        "SHA256: %s",
        sha256_file(SOURCE)
    )

    # --------------------------------------------------------
    # Open original XLS workbook.
    # --------------------------------------------------------

    book = xlrd.open_workbook(
        str(SOURCE),
        on_demand=True
    )

    sheet = book.sheet_by_index(
        SHEET_INDEX
    )

    logger.info(
        "Worksheet: %s",
        sheet.name
    )

    logger.info(
        "Rows: %s | Columns: %s",
        sheet.nrows,
        sheet.ncols
    )

    # --------------------------------------------------------
    # Extract original analytical data.
    # --------------------------------------------------------

    samples = extract_samples(sheet)

    if not samples:

        raise ValueError(
            "No sample records were identified."
        )

    # --------------------------------------------------------
    # Independent oxide reconstruction.
    # --------------------------------------------------------

    logger.info(
        "Calculating independent oxide proportions..."
    )

    reconstructed = [

        reconstruct_sample(record)

        for record in samples

    ]

    # --------------------------------------------------------
    # Compare with original calculations.
    # --------------------------------------------------------

    comparison = compare_original_cations(

        sheet,

        reconstructed

    )

    # --------------------------------------------------------
    # Extract original calculation blocks.
    # --------------------------------------------------------

    original_blocks = extract_original_calculation_blocks(

        sheet,

        samples

    )

    # --------------------------------------------------------
    # Extract optimization factors.
    # --------------------------------------------------------

    optimization = audit_optimization_factors(

        sheet,

        samples

    )

    # --------------------------------------------------------
    # Original P-T results.
    # --------------------------------------------------------

    original_PT = extract_original_PT(

        sheet,

        samples

    )

    # --------------------------------------------------------
    # Quality control.
    # --------------------------------------------------------

    quality = create_quality_control(
        samples
    )

    # --------------------------------------------------------
    # Close source workbook.
    # --------------------------------------------------------

    book.release_resources()

    # --------------------------------------------------------
    # Generate report.
    # --------------------------------------------------------

    report_data = {

        "Sample_Inputs": samples,

        "Independent_Reconstruction": reconstructed,

        "Original_Cation_Comparison": comparison,

        "Original_Calculation_Blocks": original_blocks,

        "Optimization_Factors": optimization,

        "Original_PT_Results": original_PT,

        "Chemical_Input_QC": quality,

        "Method_Inventory": create_method_inventory()

    }

    save_report(report_data)

    # --------------------------------------------------------
    # Summary statistics.
    # --------------------------------------------------------

    comparison_df = pd.DataFrame(comparison)

    qc_df = pd.DataFrame(quality)

    if len(comparison_df):

        valid = comparison_df[
            "Absolute_Difference"
        ].dropna()

        mae = (
            float(valid.mean())
            if len(valid)
            else np.nan
        )

        maximum_error = (
            float(valid.max())
            if len(valid)
            else np.nan
        )

    else:

        mae = np.nan

        maximum_error = np.nan

    logger.info("=" * 75)

    logger.info(
        "STEP 05C.4.1 COMPLETED"
    )

    logger.info(
        "Samples extracted: %s",
        len(samples)
    )

    logger.info(
        "Independent reconstructions: %s",
        len(reconstructed)
    )

    logger.info(
        "Cation comparisons: %s",
        len(comparison)
    )

    logger.info(
        "Mean absolute cation difference (23O comparison): %s",
        mae
    )

    logger.info(
        "Maximum absolute cation difference: %s",
        maximum_error
    )

    logger.info(
        "Chemical QC PASS: %s",
        int(
            (qc_df["Input_QC"] == "PASS").sum()
        )
    )

    logger.info(
        "Chemical QC REVIEW: %s",
        int(
            (qc_df["Input_QC"] == "REVIEW").sum()
        )
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

    print("\nSTEP 05C.4.1 COMPLETED")

    print("\nSamples reconstructed:", len(samples))

    print("\nExcel report:")
    print(REPORT)

    print("\nExecution log:")
    print(LOG_FILE)

    print(
        "\nIMPORTANT: Alternative oxygen normalizations "
        "are diagnostic comparisons, not certified "
        "Hbl97 structural formulae."
    )


# ============================================================================
# 15. EXECUTION
# ============================================================================

if __name__ == "__main__":

    try:

        main()

    except Exception as error:

        logger.exception(
            "STEP 05C.4.1 FAILED: %s",
            error
        )

        raise
