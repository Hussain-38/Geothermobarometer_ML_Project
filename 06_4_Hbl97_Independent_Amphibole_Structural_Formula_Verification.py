
# -*- coding: utf-8 -*-
"""
===============================================================================
STEP 06.4 — INDEPENDENT Hbl97 AMPHIBOLE STRUCTURAL FORMULA VERIFICATION
===============================================================================

Project:
Explainable Machine Learning-Based Evaluation of Geothermobarometers

Purpose:
1. Independently inspect reconstructed amphibole cations.
2. Reconstruct Fe2+ from total Fe and original Hbl97 Fe3+.
3. Investigate T, C, B and A site occupancies.
4. Calculate formal cation charge and charge residual.
5. Cross-check sample identities against Step 06.3.2.
6. Identify site-capacity excesses and unallocated cations.
7. Produce a comprehensive Excel audit report.

IMPORTANT:
This is a diagnostic structural-formula investigation.
It does not certify IMA site allocation, Fe3+ oxidation state,
or the geological validity of the Hbl97 calibration.

No site occupancy is artificially forced to its ideal value.
===============================================================================
"""

from pathlib import Path
from datetime import datetime
import logging
import hashlib
import math

import openpyxl
from openpyxl.styles import (
    Font, PatternFill, Border,
    Side, Alignment
)
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. PROJECT PATHS
# =============================================================================

ROOT = Path(
    r"D:\Data Science Assignment\My Projects\Geothermobarometer_ML_Project"
)

CHEM_FILE = (
    ROOT /
    "05C4_1_Hbl97_original_chemical_reconstruction" /
    "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

OPT_FILE = (
    ROOT /
    "06_3_2_Hbl97_Exact_Original_Optimization_Reconstruction" /
    "Step06_3_2_Hbl97_Exact_Original_Optimization_Reconstruction.xlsx"
)

OUTPUT_DIR = (
    ROOT /
    "06_4_Hbl97_Independent_Amphibole_Structural_Formula_Verification"
)

LOG_DIR = OUTPUT_DIR / "logs"

OUTPUT_FILE = (
    OUTPUT_DIR /
    "Step06_4_Hbl97_Independent_Amphibole_Structural_Formula_Verification.xlsx"
)

LOG_FILE = LOG_DIR / "step06_4_execution_log.txt"


# =============================================================================
# 2. SCIENTIFIC SETTINGS
# =============================================================================

EXPECTED_SAMPLES = 13

ABS_TOL = 1e-8

SITE_TOL = 0.02

OXYGEN_NORMALIZATION = 23.0

OXYGEN_CHARGE = 46.0

SITE_CAPACITY = {
    "T": 8.0,
    "C": 5.0,
    "B": 2.0,
    "A": 1.0
}


# =============================================================================
# 3. LOGGING
# =============================================================================

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
LOG_DIR.mkdir(parents=True, exist_ok=True)

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

logger = logging.getLogger("STEP06_4")


# =============================================================================
# 4. GENERAL FUNCTIONS
# =============================================================================

def read_sheet(path, sheet_name):

    if not path.exists():
        raise FileNotFoundError(
            f"Required workbook not found:\n{path}"
        )

    wb = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=True
    )

    try:

        if sheet_name not in wb.sheetnames:

            raise ValueError(
                f"Worksheet '{sheet_name}' not found in {path.name}.\n"
                f"Available sheets: {wb.sheetnames}"
            )

        ws = wb[sheet_name]

        iterator = ws.iter_rows(values_only=True)

        headers = next(iterator)

        headers = [
            str(h).strip() if h is not None else None
            for h in headers
        ]

        records = []

        for row in iterator:

            record = {}

            for header, value in zip(headers, row):

                if header is not None:
                    record[header] = value

            records.append(record)

        return records

    finally:

        wb.close()


def to_float(value):

    if value is None:
        return None

    try:

        value = float(value)

        if math.isfinite(value):
            return value

    except (ValueError, TypeError, OverflowError):

        pass

    return None


def sha256_file(path):

    digest = hashlib.sha256()

    with open(path, "rb") as f:

        for block in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):

            digest.update(block)

    return digest.hexdigest()


def close_enough(a, b, tolerance=ABS_TOL):

    if a is None or b is None:
        return False

    return abs(a - b) <= tolerance


def occupancy_status(value, capacity):

    difference = value - capacity

    if abs(difference) <= SITE_TOL:
        return "WITHIN_TOLERANCE"

    if difference > 0:
        return "EXCEEDS_CAPACITY"

    return "BELOW_CAPACITY"


def check_nonnegative(name, values):

    for key, value in values.items():

        if value is None:

            raise ValueError(
                f"{name}: missing value for {key}"
            )

        if value < -ABS_TOL:

            raise ValueError(
                f"{name}: negative value for {key}: {value}"
            )


# =============================================================================
# 5. READ INDEPENDENT CHEMICAL RECONSTRUCTION
# =============================================================================

def load_chemistry():

    logger.info(
        "Reading Step 05C4.1 independent reconstruction."
    )

    records = read_sheet(
        CHEM_FILE,
        "Independent_Reconstruction"
    )

    output = {}

    required_columns = [
        "Sample",
        "SiO2_cation_per_23O",
        "TiO2_cation_per_23O",
        "Al2O3_cation_per_23O",
        "FeO_total_cation_per_23O",
        "MnO_cation_per_23O",
        "MgO_cation_per_23O",
        "CaO_cation_per_23O",
        "Na2O_cation_per_23O",
        "K2O_cation_per_23O",
        "Fe3_Hbl_Original"
    ]

    if not records:
        raise ValueError(
            "Independent_Reconstruction worksheet is empty."
        )

    missing_columns = [
        col for col in required_columns
        if col not in records[0]
    ]

    if missing_columns:

        raise ValueError(
            f"Missing required chemical columns: {missing_columns}"
        )

    for r in records:

        sample = r.get("Sample")

        if sample is None:
            continue

        sample = str(sample).strip()

        if not sample:
            continue

        if sample in output:

            raise ValueError(
                f"Duplicate sample detected: {sample}"
            )

        c = {

            "Si": to_float(
                r["SiO2_cation_per_23O"]
            ),

            "Ti": to_float(
                r["TiO2_cation_per_23O"]
            ),

            "Al": to_float(
                r["Al2O3_cation_per_23O"]
            ),

            "Fe_total": to_float(
                r["FeO_total_cation_per_23O"]
            ),

            "Mn": to_float(
                r["MnO_cation_per_23O"]
            ),

            "Mg": to_float(
                r["MgO_cation_per_23O"]
            ),

            "Ca": to_float(
                r["CaO_cation_per_23O"]
            ),

            "Na": to_float(
                r["Na2O_cation_per_23O"]
            ),

            "K": to_float(
                r["K2O_cation_per_23O"]
            )

        }

        fe3 = to_float(
            r["Fe3_Hbl_Original"]
        )

        output[sample] = {

            "Sample": sample,

            "Excel_Row": r.get("Excel_Row"),

            "Reference": r.get("Reference"),

            "Cations": c,

            "Fe3_Original": fe3

        }

    logger.info(
        "Independent chemical records: %d",
        len(output)
    )

    return output


# =============================================================================
# 6. READ STEP 06.3.2 SAMPLE IDENTITIES
# =============================================================================

def load_optimization_samples():

    logger.info(
        "Reading Step 06.3.2 sample identities."
    )

    records = read_sheet(
        OPT_FILE,
        "Sample_Summary"
    )

    samples = set()

    for r in records:

        sample = r.get("Sample")

        if sample is not None:

            sample = str(sample).strip()

            if sample:
                samples.add(sample)

    if not samples:

        raise ValueError(
            "No sample identities found in Step 06.3.2."
        )

    logger.info(
        "Step 06.3.2 samples recovered: %d",
        len(samples)
    )

    return samples


# =============================================================================
# 7. VALIDATE INPUTS
# =============================================================================

def validate_inputs(chemistry, optimization_samples):

    chemistry_samples = set(chemistry)

    if len(chemistry_samples) != EXPECTED_SAMPLES:

        raise ValueError(
            f"Expected {EXPECTED_SAMPLES} chemistry samples; "
            f"found {len(chemistry_samples)}."
        )

    missing_from_optimization = (
        chemistry_samples - optimization_samples
    )

    missing_from_chemistry = (
        optimization_samples - chemistry_samples
    )

    if missing_from_optimization:

        raise ValueError(
            "Samples missing from Step 06.3.2: "
            f"{sorted(missing_from_optimization)}"
        )

    if missing_from_chemistry:

        logger.warning(
            "Additional samples in Step 06.3.2: %s",
            sorted(missing_from_chemistry)
        )

    for sample, record in chemistry.items():

        c = record["Cations"]

        check_nonnegative(
            sample,
            c
        )

        fe3 = record["Fe3_Original"]

        if fe3 is None or fe3 < 0:

            raise ValueError(
                f"{sample}: invalid original Fe3+ value."
            )

        if fe3 > c["Fe_total"] + ABS_TOL:

            raise ValueError(
                f"{sample}: Fe3+ exceeds total Fe."
            )

    return sorted(chemistry_samples)


# =============================================================================
# 8. IRON PARTITIONING
# =============================================================================

def partition_iron(c, fe3_original):

    fe_total = c["Fe_total"]

    fe3 = fe3_original

    fe2 = fe_total - fe3

    difference = fe_total - (fe2 + fe3)

    return {

        "Fe_total": fe_total,

        "Fe3": fe3,

        "Fe2": fe2,

        "Fe_partition_sum": fe2 + fe3,

        "Fe_partition_difference": difference,

        "Fe_partition_status": (
            "PASS"
            if abs(difference) <= ABS_TOL
            else "REVIEW"
        ),

        "Fe3_FeTotal_Ratio": (
            fe3 / fe_total
            if fe_total > 0
            else None
        )

    }


# =============================================================================
# 9. T-SITE ALLOCATION
# =============================================================================

def allocate_T_site(c):

    Si = c["Si"]

    Al = c["Al"]

    Si_T = min(Si, 8.0)

    Al_IV = min(
        Al,
        max(0.0, 8.0 - Si_T)
    )

    Al_VI = Al - Al_IV

    T_total = Si_T + Al_IV

    return {

        "Si_T": Si_T,

        "Al_IV": Al_IV,

        "Al_VI": Al_VI,

        "T_total": T_total,

        "T_capacity": 8.0,

        "T_difference": T_total - 8.0,

        "T_status": occupancy_status(
            T_total,
            8.0
        )

    }


# =============================================================================
# 10. C-SITE ALLOCATION
# =============================================================================

def allocate_C_site(c, iron, T):

    Al_VI = T["Al_VI"]

    Ti_C = c["Ti"]

    Fe3_C = iron["Fe3"]

    Fe2_C = iron["Fe2"]

    Mn_C = c["Mn"]

    Mg_C = c["Mg"]

    C_total = (
        Al_VI
        + Ti_C
        + Fe3_C
        + Fe2_C
        + Mn_C
        + Mg_C
    )

    return {

        "Al_VI_C": Al_VI,

        "Ti_C": Ti_C,

        "Fe3_C": Fe3_C,

        "Fe2_C": Fe2_C,

        "Mn_C": Mn_C,

        "Mg_C": Mg_C,

        "C_total": C_total,

        "C_capacity": 5.0,

        "C_difference": C_total - 5.0,

        "C_status": occupancy_status(
            C_total,
            5.0
        )

    }


# =============================================================================
# 11. B-SITE ALLOCATION
# =============================================================================

def allocate_B_site(c):

    Ca = c["Ca"]

    Na = c["Na"]

    Ca_B = min(Ca, 2.0)

    Na_B = min(
        Na,
        max(0.0, 2.0 - Ca_B)
    )

    Ca_unallocated = Ca - Ca_B

    Na_unallocated = Na - Na_B

    B_total = Ca_B + Na_B

    return {

        "Ca_B": Ca_B,

        "Na_B": Na_B,

        "Ca_unallocated": Ca_unallocated,

        "Na_unallocated": Na_unallocated,

        "B_total": B_total,

        "B_capacity": 2.0,

        "B_difference": B_total - 2.0,

        "B_status": occupancy_status(
            B_total,
            2.0
        )

    }


# =============================================================================
# 12. A-SITE ALLOCATION
# =============================================================================

def allocate_A_site(c, B):

    Na_A = B["Na_unallocated"]

    K_A = c["K"]

    A_total = Na_A + K_A

    return {

        "Na_A": Na_A,

        "K_A": K_A,

        "A_total": A_total,

        "A_capacity": 1.0,

        "A_difference": A_total - 1.0,

        "A_status": occupancy_status(
            A_total,
            1.0
        )

    }


# =============================================================================
# 13. FORMAL CHARGE BALANCE
# =============================================================================

def calculate_charge(c, iron):

    # Assumed oxidation states:
    #
    # Si  = +4
    # Ti  = +4
    # Al  = +3
    # Fe2 = +2
    # Fe3 = +3
    # Mn  = +2
    # Mg  = +2
    # Ca  = +2
    # Na  = +1
    # K   = +1
    #
    # The 23-oxygen normalization is compared against
    # 46 negative charge units, corresponding to the
    # conventional 22 O2- + 2 OH- anion basis.

    positive_charge = (

        4 * c["Si"]

        + 4 * c["Ti"]

        + 3 * c["Al"]

        + 2 * iron["Fe2"]

        + 3 * iron["Fe3"]

        + 2 * c["Mn"]

        + 2 * c["Mg"]

        + 2 * c["Ca"]

        + c["Na"]

        + c["K"]

    )

    residual = positive_charge - OXYGEN_CHARGE

    return {

        "Positive_Cation_Charge": positive_charge,

        "Assumed_Oxygen_Charge": OXYGEN_CHARGE,

        "Charge_Residual": residual,

        "Absolute_Charge_Residual": abs(residual),

        "Charge_Status": (
            "PASS"
            if abs(residual) <= SITE_TOL
            else "REVIEW"
        )

    }


# =============================================================================
# 14. COMPLETE SAMPLE FORMULA
# =============================================================================

def calculate_formula(sample, record):

    c = record["Cations"]

    iron = partition_iron(
        c,
        record["Fe3_Original"]
    )

    T = allocate_T_site(c)

    C = allocate_C_site(
        c,
        iron,
        T
    )

    B = allocate_B_site(c)

    A = allocate_A_site(
        c,
        B
    )

    charge = calculate_charge(
        c,
        iron
    )

    total_cations = sum(c.values())

    allocated_site_sum = (
        T["T_total"]
        + C["C_total"]
        + B["B_total"]
        + A["A_total"]
    )

    # Explicit status assessment.
    #
    # This is not a certification of site allocation.

    site_excess = any([

        T["T_status"] == "EXCEEDS_CAPACITY",

        C["C_status"] == "EXCEEDS_CAPACITY",

        B["B_status"] == "EXCEEDS_CAPACITY",

        A["A_status"] == "EXCEEDS_CAPACITY"

    ])

    unallocated_ca_flag = (
        B["Ca_unallocated"] > SITE_TOL
    )

    if site_excess or unallocated_ca_flag:

        overall_site_status = "REVIEW"

    elif any([

        T["T_status"] == "BELOW_CAPACITY",

        C["C_status"] == "BELOW_CAPACITY",

        B["B_status"] == "BELOW_CAPACITY",

        A["A_status"] == "BELOW_CAPACITY"

    ]):

        overall_site_status = "BELOW_CAPACITY"

    else:

        overall_site_status = "WITHIN_TOLERANCE"

    return {

        "Sample": sample,

        "Excel_Row": record["Excel_Row"],

        "Reference": record["Reference"],

        "Normalization_Oxygen": OXYGEN_NORMALIZATION,

        # Initial cations

        "Si": c["Si"],

        "Ti": c["Ti"],

        "Al_total": c["Al"],

        "Fe_total": c["Fe_total"],

        "Fe3": iron["Fe3"],

        "Fe2": iron["Fe2"],

        "Mn": c["Mn"],

        "Mg": c["Mg"],

        "Ca": c["Ca"],

        "Na": c["Na"],

        "K": c["K"],

        "Total_Cations": total_cations,

        # IMPORTANT: all iron fields are retained.

        "Fe_partition_sum": iron["Fe_partition_sum"],

        "Fe_partition_difference": iron[
            "Fe_partition_difference"
        ],

        "Fe_partition_status": iron[
            "Fe_partition_status"
        ],

        "Fe3_FeTotal_Ratio": iron[
            "Fe3_FeTotal_Ratio"
        ],

        # T site

        **T,

        # C site

        **C,

        # B site

        **B,

        # A site

        **A,

        # Charge

        **charge,

        # Totals

        "Allocated_Site_Cation_Sum": allocated_site_sum,

        "Unallocated_Ca": B["Ca_unallocated"],

        "Unallocated_Na": B["Na_unallocated"],

        "Site_Allocation_Status": overall_site_status,

        "Scientific_Certification": "NOT ESTABLISHED"

    }


# =============================================================================
# 15. IRON AND SAMPLE CROSS-CHECK
# =============================================================================

def build_crosscheck(sample, formula):

    # Fe2+ and Fe3+ are checked independently against
    # the original total-iron input.

    fe_total = formula["Fe_total"]

    fe2 = formula["Fe2"]

    fe3 = formula["Fe3"]

    return {

        "Sample": sample,

        "Original_Fe3": fe3,

        "Reconstructed_Fe2": fe2,

        "Initial_Total_Fe": fe_total,

        "Fe2_plus_Fe3": fe2 + fe3,

        "Fe_Partition_Difference": (
            fe_total - fe2 - fe3
        ),

        "Fe_Partition_Check": (
            "PASS"
            if close_enough(
                fe_total,
                fe2 + fe3
            )
            else "REVIEW"
        ),

        "Fe3_FeTotal_Ratio": formula[
            "Fe3_FeTotal_Ratio"
        ],

        "Arithmetic_Status": formula[
            "Fe_partition_status"
        ],

        "Oxidation_State_Certification": (
            "NOT ESTABLISHED"
        )

    }


# =============================================================================
# 16. REPORT RECORD WRITER
# =============================================================================

def write_records(wb, sheet_name, records):

    ws = wb.create_sheet(sheet_name)

    if not records:
        ws.append(["No records"])
        return ws

    headers = list(records[0].keys())

    ws.append(headers)

    for record in records:

        ws.append([
            record.get(header)
            for header in headers
        ])

    return ws


# =============================================================================
# 17. EXCEL FORMATTING
# =============================================================================

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

        if ws.max_row > 1:

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

        for col in range(1, ws.max_column + 1):

            maximum = 0

            for row in range(
                1,
                min(ws.max_row, 300) + 1
            ):

                value = ws.cell(
                    row,
                    col
                ).value

                if value is not None:

                    maximum = max(
                        maximum,
                        len(str(value))
                    )

            ws.column_dimensions[
                get_column_letter(col)
            ].width = min(
                max(14, maximum + 2),
                42
            )

        ws.row_dimensions[1].height = 36


# =============================================================================
# 18. MAIN EXECUTION
# =============================================================================

def main():

    logger.info("=" * 85)

    logger.info(
        "STEP 06.4 — INDEPENDENT AMPHIBOLE STRUCTURAL FORMULA VERIFICATION"
    )

    logger.info("=" * 85)

    # -------------------------------------------------------------------------
    # Verify files
    # -------------------------------------------------------------------------

    for path in [CHEM_FILE, OPT_FILE]:

        if not path.exists():

            raise FileNotFoundError(
                f"Required file not found:\n{path}"
            )

        logger.info(
            "Verified input: %s",
            path
        )

    chemistry_hash = sha256_file(
        CHEM_FILE
    )

    optimization_hash = sha256_file(
        OPT_FILE
    )

    logger.info(
        "Chemical workbook SHA256: %s",
        chemistry_hash
    )

    logger.info(
        "Optimization workbook SHA256: %s",
        optimization_hash
    )

    # -------------------------------------------------------------------------
    # Load inputs
    # -------------------------------------------------------------------------

    chemistry = load_chemistry()

    optimization_samples = load_optimization_samples()

    samples = validate_inputs(
        chemistry,
        optimization_samples
    )

    logger.info(
        "Matched samples: %d",
        len(samples)
    )

    # -------------------------------------------------------------------------
    # Process samples
    # -------------------------------------------------------------------------

    formula_records = []

    site_records = []

    charge_records = []

    iron_records = []

    sample_summary = []

    for index, sample in enumerate(samples, start=1):

        logger.info(
            "Processing %d/%d: %s",
            index,
            len(samples),
            sample
        )

        formula = calculate_formula(
            sample,
            chemistry[sample]
        )

        formula_records.append(formula)

        # -----------------------------------------------------
        # Site occupancy report
        # -----------------------------------------------------

        site_records.append({

            "Sample": sample,

            "T_Occupancy": formula["T_total"],

            "T_Capacity": 8.0,

            "T_Difference": formula["T_difference"],

            "T_Status": formula["T_status"],

            "C_Occupancy": formula["C_total"],

            "C_Capacity": 5.0,

            "C_Difference": formula["C_difference"],

            "C_Status": formula["C_status"],

            "B_Occupancy": formula["B_total"],

            "B_Capacity": 2.0,

            "B_Difference": formula["B_difference"],

            "B_Status": formula["B_status"],

            "A_Occupancy": formula["A_total"],

            "A_Capacity": 1.0,

            "A_Difference": formula["A_difference"],

            "A_Status": formula["A_status"],

            "Unallocated_Ca": formula["Unallocated_Ca"],

            "Unallocated_Na": formula["Unallocated_Na"],

            "Overall_Site_Status": formula[
                "Site_Allocation_Status"
            ]

        })

        # -----------------------------------------------------
        # Charge report
        # -----------------------------------------------------

        charge_records.append({

            "Sample": sample,

            "Positive_Cation_Charge": formula[
                "Positive_Cation_Charge"
            ],

            "Assumed_Oxygen_Charge": formula[
                "Assumed_Oxygen_Charge"
            ],

            "Charge_Residual": formula[
                "Charge_Residual"
            ],

            "Absolute_Charge_Residual": formula[
                "Absolute_Charge_Residual"
            ],

            "Charge_Status": formula[
                "Charge_Status"
            ]

        })

        # -----------------------------------------------------
        # Iron report
        # -----------------------------------------------------

        iron_record = build_crosscheck(
            sample,
            formula
        )

        iron_records.append(
            iron_record
        )

        # -----------------------------------------------------
        # Sample summary
        # -----------------------------------------------------

        sample_summary.append({

            "Sample": sample,

            "Total_Cations": formula[
                "Total_Cations"
            ],

            "Fe_total": formula[
                "Fe_total"
            ],

            "Fe2": formula[
                "Fe2"
            ],

            "Fe3": formula[
                "Fe3"
            ],

            "T_Status": formula[
                "T_status"
            ],

            "C_Status": formula[
                "C_status"
            ],

            "B_Status": formula[
                "B_status"
            ],

            "A_Status": formula[
                "A_status"
            ],

            "Charge_Status": formula[
                "Charge_Status"
            ],

            "Site_Allocation_Status": formula[
                "Site_Allocation_Status"
            ],

            "Fe_Partition_Status": formula[
                "Fe_partition_status"
            ],

            "Scientific_Certification": (
                "NOT ESTABLISHED"
            )

        })

    # =============================================================================
    # 19. SUMMARY STATISTICS
    # =============================================================================

    site_review = sum(

        r["Site_Allocation_Status"] == "REVIEW"

        for r in formula_records

    )

    site_below = sum(

        r["Site_Allocation_Status"] == "BELOW_CAPACITY"

        for r in formula_records

    )

    charge_pass = sum(

        r["Charge_Status"] == "PASS"

        for r in formula_records

    )

    iron_pass = sum(

        r["Fe_partition_status"] == "PASS"

        for r in formula_records

    )

    iron_review = sum(

        r["Fe_partition_status"] != "PASS"

        for r in formula_records

    )

    crosscheck_pass = sum(

        r["Fe_Partition_Check"] == "PASS"

        for r in iron_records

    )

    # =============================================================================
    # 20. CREATE EXCEL REPORT
    # =============================================================================

    logger.info(
        "Generating Excel report."
    )

    wb = openpyxl.Workbook()

    # -------------------------------------------------------------------------
    # Read Me
    # -------------------------------------------------------------------------

    ws = wb.active

    ws.title = "Read_Me"

    readme = [

        ["Item", "Description"],

        [
            "Step",
            "06.4 Independent Hbl97 amphibole structural formula verification"
        ],

        [
            "Chemical source",
            CHEM_FILE.name
        ],

        [
            "Optimization source",
            OPT_FILE.name
        ],

        [
            "Normalization",
            "23 oxygen basis"
        ],

        [
            "Iron source",
            "Original Hbl97 Fe3+ field"
        ],

        [
            "Fe2+ calculation",
            "Total Fe minus original Hbl97 Fe3+"
        ],

        [
            "T-site convention",
            "Si first, followed by tetrahedral Al"
        ],

        [
            "C-site convention",
            "AlVI, Ti, Fe3+, Fe2+, Mn and Mg"
        ],

        [
            "B-site convention",
            "Ca first, followed by Na"
        ],

        [
            "A-site convention",
            "Remaining Na and K"
        ],

        [
            "Charge calculation",
            "Formal cation charge compared with 46 negative charge units"
        ],

        [
            "Site tolerance",
            SITE_TOL
        ],

        [
            "Scientific limitation",
            "Simplified diagnostic allocation; not IMA certification"
        ],

        [
            "Execution time",
            datetime.now().isoformat()
        ],

        [
            "Chemical workbook SHA256",
            chemistry_hash
        ],

        [
            "Optimization workbook SHA256",
            optimization_hash
        ]

    ]

    for row in readme:
        ws.append(row)

    # -------------------------------------------------------------------------
    # Overall summary
    # -------------------------------------------------------------------------

    overall = [

        ["Metric", "Result"],

        ["Total samples", len(samples)],

        ["Iron partition arithmetic PASS", iron_pass],

        ["Iron partition arithmetic REVIEW", iron_review],

        ["Iron cross-check PASS", crosscheck_pass],

        ["Charge-balance diagnostic PASS", charge_pass],

        ["Site-allocation REVIEW", site_review],

        ["Site-allocation BELOW CAPACITY", site_below],

        [
            "Site allocation certification",
            "NOT ESTABLISHED"
        ],

        [
            "Fe3+ determination certification",
            "NOT ESTABLISHED"
        ],

        [
            "Geological calibration certification",
            "NOT ESTABLISHED"
        ],

        [
            "Overall scientific interpretation",
            "DIAGNOSTIC RESULTS REQUIRE MINERALOGICAL REVIEW"
        ]

    ]

    ws = wb.create_sheet(
        "Overall_Summary"
    )

    for row in overall:
        ws.append(row)

    # -------------------------------------------------------------------------
    # Detailed worksheets
    # -------------------------------------------------------------------------

    write_records(
        wb,
        "Structural_Formula",
        formula_records
    )

    write_records(
        wb,
        "Site_Occupancy",
        site_records
    )

    write_records(
        wb,
        "Charge_Balance",
        charge_records
    )

    write_records(
        wb,
        "Iron_Crosscheck",
        iron_records
    )

    write_records(
        wb,
        "Sample_Summary",
        sample_summary
    )

    # -------------------------------------------------------------------------
    # Scientific limitations
    # -------------------------------------------------------------------------

    limitations = [

        {

            "Issue": "Fe3+ provenance",

            "Finding": (
                "Original Hbl97 Fe3+ values were used."
            ),

            "Limitation": (
                "The present calculations do not independently "
                "measure iron oxidation state."
            ),

            "Required_Evidence": (
                "Independent analytical Fe3+ determination, "
                "where available."
            )

        },

        {

            "Issue": "Site allocation",

            "Finding": (
                "An explicit sequential T-C-B-A allocation "
                "convention was applied."
            ),

            "Limitation": (
                "This convention does not certify crystallographic "
                "site occupancy or establish the full published "
                "amphibole formula procedure."
            ),

            "Required_Evidence": (
                "Validation against the applicable published "
                "amphibole structural formula methodology."
            )

        },

        {

            "Issue": "Charge balance",

            "Finding": (
                "Formal cation charge was compared with 46 "
                "negative charge units."
            ),

            "Limitation": (
                "Residuals may reflect analytical uncertainty, "
                "normalization, oxidation-state assumptions "
                "or anion assumptions."
            ),

            "Required_Evidence": (
                "Independent analytical and formula validation."
            )

        },

        {

            "Issue": "Site capacity",

            "Finding": (
                "Nominal T=8, C=5, B=2 and A=1 capacities "
                "were investigated."
            ),

            "Limitation": (
                "The simplified allocation is diagnostic and "
                "may not represent the full mineralogical "
                "distribution of cations."
            ),

            "Required_Evidence": (
                "Comparison with the relevant amphibole "
                "classification and site-allocation procedure."
            )

        },

        {

            "Issue": "Geothermobarometer accuracy",

            "Finding": (
                "This step investigates structural formula "
                "consistency only."
            ),

            "Limitation": (
                "It does not assess prediction accuracy or "
                "precision of the Hbl97 calibration."
            ),

            "Required_Evidence": (
                "Independent pressure-temperature reference "
                "conditions and comparative calibration evaluation."
            )

        }

    ]

    write_records(
        wb,
        "Scientific_Limitations",
        limitations
    )

    # -------------------------------------------------------------------------
    # Format and save
    # -------------------------------------------------------------------------

    format_workbook(wb)

    wb.save(OUTPUT_FILE)

    # Reopen to verify that the report was written successfully.

    check_wb = openpyxl.load_workbook(
        OUTPUT_FILE,
        read_only=True,
        data_only=True
    )

    try:

        expected_sheets = [

            "Read_Me",
            "Overall_Summary",
            "Structural_Formula",
            "Site_Occupancy",
            "Charge_Balance",
            "Iron_Crosscheck",
            "Sample_Summary",
            "Scientific_Limitations"

        ]

        missing = [
            s for s in expected_sheets
            if s not in check_wb.sheetnames
        ]

        if missing:

            raise RuntimeError(
                f"Output report missing worksheets: {missing}"
            )

        if (
            check_wb["Sample_Summary"].max_row
            != EXPECTED_SAMPLES + 1
        ):

            raise RuntimeError(
                "Unexpected sample count in final Excel report."
            )

    finally:

        check_wb.close()

    # =============================================================================
    # 21. FINAL CONSOLE REPORT
    # =============================================================================

    logger.info("=" * 85)

    logger.info(
        "STEP 06.4 COMPLETED"
    )

    logger.info(
        "Samples: %d",
        len(samples)
    )

    logger.info(
        "Iron partition PASS: %d/%d",
        iron_pass,
        len(samples)
    )

    logger.info(
        "Iron partition REVIEW: %d",
        iron_review
    )

    logger.info(
        "Iron cross-check PASS: %d/%d",
        crosscheck_pass,
        len(samples)
    )

    logger.info(
        "Charge-balance diagnostic PASS: %d/%d",
        charge_pass,
        len(samples)
    )

    logger.info(
        "Site-allocation REVIEW: %d",
        site_review
    )

    logger.info(
        "Site-allocation BELOW CAPACITY: %d",
        site_below
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
# 22. EXECUTION
# =============================================================================

if __name__ == "__main__":

    try:

        main()

    except Exception:

        logger.exception(
            "STEP 06.4 FAILED"
        )

        raise
