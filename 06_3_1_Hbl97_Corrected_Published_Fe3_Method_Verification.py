# -*- coding: utf-8 -*-
"""
Created on Tue Nov 17 11:33:10 2026

@author: hussa
"""


# -*- coding: utf-8 -*-
"""
STEP 06.3.1
Hbl97 — CORRECTED INDEPENDENT Fe3+ METHOD VERIFICATION

Project:
Geothermobarometer ML Project

Scientific basis:
Holland & Blundy (1994), Appendix B.
Independent implementation cross-check:
Thermobar documentation, amphibole ferric/ferrous recalculation.
Mutch et al. spreadsheet implementation.

IMPORTANT:
1. Original workbook is read-only.
2. Both f8 conventions (12.9 and 13.0) are reported.
3. Both averaging conventions are reported.
4. No scientific certification is automatically assigned.
5. Site allocation is an explicit computational reconstruction.
"""

from pathlib import Path
from datetime import datetime
import hashlib
import logging
import math

import xlrd
import openpyxl

from openpyxl.styles import (
    Font, PatternFill, Border, Side, Alignment
)
from openpyxl.utils import get_column_letter


# =============================================================================
# A. PATHS
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

OPT_FILE = (
    ROOT /
    "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit" /
    "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

OUTPUT_DIR = (
    ROOT /
    "06_3_1_Hbl97_Corrected_Published_Fe3_Method_Verification"
)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = (
    OUTPUT_DIR /
    "Step06_3_1_Hbl97_Corrected_Fe3_Verification.xlsx"
)

LOG_FILE = LOG_DIR / "step06_3_1_execution_log.txt"


# =============================================================================
# B. SETTINGS
# =============================================================================

TOL = 1e-6
EXPECTED_SAMPLES = 13

ELEMENTS = [
    "Si", "Ti", "Al", "Fe_total",
    "Mn", "Mg", "Ca", "Na", "K"
]

FE_OXYGEN_FACTOR = 46.0


# =============================================================================
# C. LOGGING
# =============================================================================

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler()
    ],
    force=True
)

logger = logging.getLogger("STEP06_3_1")


# =============================================================================
# D. GENERAL FUNCTIONS
# =============================================================================

def sha256_file(path):

    digest = hashlib.sha256()

    with open(path, "rb") as f:

        for block in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):
            digest.update(block)

    return digest.hexdigest()


def number(value):

    if value is None:
        return None

    try:
        value = float(value)

        if math.isfinite(value):
            return value

    except (ValueError, TypeError):
        pass

    return None


def compare(a, b, tolerance=TOL):

    if a is None or b is None:
        return False

    return abs(a - b) <= tolerance


def read_sheet(path, sheet_name):

    if not path.exists():
        raise FileNotFoundError(path)

    wb = openpyxl.load_workbook(
        path,
        read_only=True,
        data_only=True
    )

    if sheet_name not in wb.sheetnames:

        wb.close()

        raise ValueError(
            f"Missing worksheet: {sheet_name}"
        )

    ws = wb[sheet_name]

    rows = ws.iter_rows(values_only=True)

    headers = next(rows)

    records = []

    for row in rows:

        records.append(
            dict(zip(headers, row))
        )

    wb.close()

    return records


# =============================================================================
# E. INITIAL AMPHIBOLE SITE ALLOCATION
# =============================================================================

def allocate_initial_sites(c):

    """
    Allocate the initial 23O cations to T, C, B and A sites.

    This is a computational site-allocation reconstruction.
    It is not a certification of crystallographic occupancy.

    All total Fe is initially treated as undifferentiated Fe.
    """

    Si = c["Si"]
    Al = c["Al"]
    Ti = c["Ti"]
    Fe = c["Fe_total"]
    Mn = c["Mn"]
    Mg = c["Mg"]
    Ca = c["Ca"]
    Na = c["Na"]
    K = c["K"]

    # Fixed site assignments.

    Si_T = Si
    Ti_C = Ti
    Ca_B = Ca
    K_A = K

    # Tetrahedral Al.

    if Si_T + Al > 8:

        Al_T = 8 - Si_T

        Al_C = Si_T + Al - 8

    else:

        Al_T = Al
        Al_C = 0.0

    # Octahedral Mg.

    occupied_C = Al_C + Ti_C

    if occupied_C + Mg > 5:

        Mg_C = 5 - occupied_C

        Mg_B = occupied_C + Mg - 5

    else:

        Mg_C = Mg
        Mg_B = 0.0

    # Remaining octahedral capacity is assigned to Fe.

    occupied_C = (
        Al_C + Ti_C + Mg_C
    )

    if occupied_C + Fe > 5:

        Fe_C = 5 - occupied_C

        Fe_B = Fe - Fe_C

    else:

        Fe_C = Fe
        Fe_B = 0.0

    # Remaining octahedral capacity is assigned to Mn.

    occupied_C = (
        Al_C + Ti_C + Mg_C + Fe_C
    )

    if occupied_C + Mn > 5:

        Mn_C = 5 - occupied_C

        Mn_B = Mn - Mn_C

    else:

        Mn_C = Mn
        Mn_B = 0.0

    # Remaining B-site capacity.

    B_without_Na = (
        Mg_B + Fe_B + Mn_B + Ca_B
    )

    if B_without_Na > 2 + TOL:

        raise ValueError(
            "B-site occupancy exceeds 2 before Na allocation."
        )

    Na_B = min(
        Na,
        max(0.0, 2.0 - B_without_Na)
    )

    Na_A = Na - Na_B

    sites = {

        "Si_T": Si_T,
        "Al_T": Al_T,
        "Al_C": Al_C,
        "Ti_C": Ti_C,
        "Mg_C": Mg_C,
        "Fe_C": Fe_C,
        "Mn_C": Mn_C,
        "Mg_B": Mg_B,
        "Fe_B": Fe_B,
        "Mn_B": Mn_B,
        "Ca_B": Ca_B,
        "Na_B": Na_B,
        "Na_A": Na_A,
        "K_A": K_A

    }

    # Sum of unique site-assigned cations.
    # Do not add site subtotals a second time.

    sites["Site_Cation_Sum"] = sum(
        sites[name]
        for name in sites
        if name != "Site_Cation_Sum"
    )

    sites["T_Sum"] = Si_T + Al_T

    sites["C_Sum"] = (
        Al_C + Ti_C + Mg_C + Fe_C + Mn_C
    )

    sites["B_Sum"] = (
        Mg_B + Fe_B + Mn_B + Ca_B + Na_B
    )

    sites["A_Sum"] = Na_A + K_A

    sites["B_Overfilled"] = (
        sites["B_Sum"] > 2.0 + TOL
    )

    sites["C_Overfilled"] = (
        sites["C_Sum"] > 5.0 + TOL
    )

    sites["T_Overfilled"] = (
        sites["T_Sum"] > 8.0 + TOL
    )

    return sites


# =============================================================================
# F. FACTOR CALCULATION
# =============================================================================

def calculate_factors(c, sites):

    Si_T = sites["Si_T"]
    Al_T = sites["Al_T"]
    Al_C = sites["Al_C"]
    Ti_C = sites["Ti_C"]

    Mn_B = sites["Mn_B"]
    Mn_C = sites["Mn_C"]

    Ca_B = sites["Ca_B"]
    Na_B = sites["Na_B"]
    Na_A = sites["Na_A"]
    K_A = sites["K_A"]

    Fe_total = c["Fe_total"]

    Z = sites["Site_Cation_Sum"]

    denominators = {

        "f1": Z,
        "f2": Si_T,
        "f3": Z - Na_A - Na_B - K_A + Mn_B,
        "f4": Ca_B,
        "f6": Si_T + Al_T + Al_C,
        "f7": Z - K_A,
        "f8": Z - Na_A - Na_B - K_A - Ca_B - Mn_B - Mn_C,
        "f9": 46.0 - Al_T - Al_C - Si_T - Ti_C,
        "f10": 46.0 + Fe_total

    }

    for name, denominator in denominators.items():

        if denominator <= 0:

            raise ValueError(
                f"Invalid denominator for {name}: {denominator}"
            )

    # Factor equations based on the documented
    # site-specific Mutch implementation.

    factors = {

        "f1": 16.0 / Z,

        "f2": 8.0 / Si_T,

        "f3": 15.0 / (
            Z - Na_A - Na_B - K_A + Mn_B
        ),

        "f4": 2.0 / Ca_B,

        "f5": 1.0,

        "f6": 8.0 / (
            Si_T + Al_T + Al_C
        ),

        "f7": 15.0 / (
            Z - K_A
        ),

        # Both numerator conventions retained.
        "f8_12_9": 12.9 / (
            Z - Na_A - Na_B - K_A
            - Ca_B - Mn_B - Mn_C
        ),

        "f8_13_0": 13.0 / (
            Z - Na_A - Na_B - K_A
            - Ca_B - Mn_B - Mn_C
        ),

        "f9": 36.0 / (
            46.0 - Al_T - Al_C - Si_T - Ti_C
        ),

        "f10": 46.0 / (
            46.0 + Fe_total
        )

    }

    # Group A.

    group_A = [
        min(factors[f"f{i}"], 1.0)
        for i in range(1, 6)
    ]

    fA = min(group_A)

    # Group B with both f8 variants.

    group_B_12_9 = [
        factors["f6"],
        factors["f7"],
        factors["f8_12_9"],
        factors["f9"],
        factors["f10"]
    ]

    group_B_13 = [
        factors["f6"],
        factors["f7"],
        factors["f8_13_0"],
        factors["f9"],
        factors["f10"]
    ]

    fB_12_9 = max(group_B_12_9)
    fB_13 = max(group_B_13)

    # Two averaging conventions.

    fav_arithmetic_12_9 = (
        fA + fB_12_9
    ) / 2.0

    fav_arithmetic_13 = (
        fA + fB_13
    ) / 2.0

    fav_weighted_12_9 = (
        (2.0 / 3.0) * fA
        + (1.0 / 3.0) * fB_12_9
    )

    fav_weighted_13 = (
        (2.0 / 3.0) * fA
        + (1.0 / 3.0) * fB_13
    )

    def valid_interval(fA_value, fB_value):

        return (
            fB_value <= fA_value + TOL
            and fB_value <= 1.0 + TOL
        )

    return {

        **factors,

        "Z": Z,

        "fA": fA,

        "fB_12_9": fB_12_9,

        "fB_13": fB_13,

        "fav_arithmetic_12_9": fav_arithmetic_12_9,

        "fav_arithmetic_13": fav_arithmetic_13,

        "fav_weighted_12_9": fav_weighted_12_9,

        "fav_weighted_13": fav_weighted_13,

        "Constraint_Valid_12_9": valid_interval(
            fA, fB_12_9
        ),

        "Constraint_Valid_13": valid_interval(
            fA, fB_13
        )

    }


# =============================================================================
# G. IRON PARTITION
# =============================================================================

def calculate_iron(c, fav):

    Fe_total = c["Fe_total"]

    Fe_total_revised = fav * Fe_total

    Fe3_raw = 46.0 * (1.0 - fav)

    Fe3_final = max(0.0, Fe3_raw)

    Fe2_final = (
        Fe_total_revised - Fe3_final
    )

    return {

        "Fe_Total_Initial": Fe_total,

        "Fe_Total_Revised": Fe_total_revised,

        "Fe3_Raw": Fe3_raw,

        "Fe3_Final": Fe3_final,

        "Fe2_Final": Fe2_final,

        "Fe2_Negative": Fe2_final < -TOL,

        "Fe3_Exceeds_Revised_Fe": (
            Fe3_final > Fe_total_revised + TOL
        )

    }


# =============================================================================
# H. ORIGINAL WORKBOOK VALUES
# =============================================================================

def read_original_workbook():

    book = xlrd.open_workbook(
        str(ORIGINAL_XLS),
        on_demand=True
    )

    ws = book.sheet_by_name(
        "Hbl P-T (1997)"
    )

    records = {}

    for row in range(9, 22):

        sample = ws.cell_value(row, 1)

        if not sample:
            continue

        records[str(sample)] = {

            "Excel_Row": row + 1,

            "Fe3_H": number(
                ws.cell_value(row, 7)
            ),

            "Fe3_AR": number(
                ws.cell_value(row, 43)
            ),

            "Fe2_AW": number(
                ws.cell_value(row, 48)
            ),

            "Fe3_AX": number(
                ws.cell_value(row, 49)
            ),

            "f1": number(ws.cell_value(row, 30)),
            "f2": number(ws.cell_value(row, 31)),
            "f3": number(ws.cell_value(row, 32)),
            "f4": number(ws.cell_value(row, 33)),
            "f5": number(ws.cell_value(row, 34)),

            "Original_Minimum": number(
                ws.cell_value(row, 35)
            ),

            "f6": number(ws.cell_value(row, 36)),
            "f7": number(ws.cell_value(row, 37)),
            "f8": number(ws.cell_value(row, 38)),
            "f9": number(ws.cell_value(row, 39)),
            "f10": number(ws.cell_value(row, 40)),

            "Original_Maximum": number(
                ws.cell_value(row, 41)
            ),

            "Original_fav": number(
                ws.cell_value(row, 42)
            )

        }

    book.release_resources()

    return records


# =============================================================================
# I. EXCEL REPORT UTILITIES
# =============================================================================

def write_records(wb, name, records):

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

        for col in range(1, ws.max_column + 1):

            maximum = 0

            for row in range(
                1,
                min(ws.max_row, 250) + 1
            ):

                value = ws.cell(row, col).value

                if value is not None:

                    maximum = max(
                        maximum,
                        len(str(value))
                    )

            ws.column_dimensions[
                get_column_letter(col)
            ].width = min(
                max(14, maximum + 2),
                45
            )

        ws.row_dimensions[1].height = 35


# =============================================================================
# J. MAIN EXECUTION
# =============================================================================

def main():

    logger.info("=" * 79)

    logger.info(
        "STEP 06.3.1 — CORRECTED PUBLISHED Fe3+ VERIFICATION"
    )

    logger.info("=" * 79)

    for path in [
        ORIGINAL_XLS,
        CHEM_FILE,
        OPT_FILE
    ]:

        if not path.exists():
            raise FileNotFoundError(path)

        logger.info(
            "Verified input: %s",
            path
        )

    source_hash = sha256_file(
        ORIGINAL_XLS
    )

    # Load original chemical cation blocks.

    blocks = read_sheet(
        CHEM_FILE,
        "Original_Calculation_Blocks"
    )

    initial_records = {

        r["Sample"]: r

        for r in blocks

        if r["Calculation_Block"] == "Initial_Cations"

    }

    # Load extracted optimization factors.

    opt_records = read_sheet(
        CHEM_FILE,
        "Optimization_Factors"
    )

    opt_map = {
        r["Sample"]: r
        for r in opt_records
    }

    # Load original XLS values.

    original_map = read_original_workbook()

    samples = sorted(
        set(initial_records)
        & set(opt_map)
        & set(original_map)
    )

    if len(samples) != EXPECTED_SAMPLES:

        raise ValueError(
            f"Expected {EXPECTED_SAMPLES} samples; "
            f"found {len(samples)}."
        )

    logger.info(
        "Samples recovered: %d",
        len(samples)
    )

    factor_output = []
    site_output = []
    iron_output = []
    summary_output = []

    for sample in samples:

        initial = initial_records[sample]

        original = original_map[sample]

        extracted = opt_map[sample]

        c = {
            element: number(initial[element])
            for element in ELEMENTS
        }

        if any(value is None for value in c.values()):

            raise ValueError(
                f"Missing cation values for {sample}"
            )

        sites = allocate_initial_sites(c)

        factors = calculate_factors(
            c,
            sites
        )

        # -----------------------------------------------------
        # Site audit
        # -----------------------------------------------------

        site_output.append({

            "Sample": sample,

            **sites

        })

        # -----------------------------------------------------
        # Factor audit
        # -----------------------------------------------------

        factor_names = [
            "f1", "f2", "f3", "f4", "f5",
            "f6", "f7", "f9", "f10"
        ]

        for factor in factor_names:

            independent = factors[factor]

            original_value = original[factor]

            extracted_value = number(
                extracted.get(factor)
            )

            factor_output.append({

                "Sample": sample,

                "Factor": factor,

                "Independent_Value": independent,

                "Original_XLS_Value": original_value,

                "Extracted_Report_Value": extracted_value,

                "Original_Difference": (
                    independent - original_value
                    if original_value is not None
                    else None
                ),

                "Report_Difference": (
                    independent - extracted_value
                    if extracted_value is not None
                    else None
                ),

                "Original_Check": (
                    "PASS"
                    if compare(independent, original_value)
                    else "REVIEW"
                ),

                "Report_Check": (
                    "PASS"
                    if compare(independent, extracted_value)
                    else "REVIEW"
                )

            })

        # f8 variants.

        for name in ["f8_12_9", "f8_13_0"]:

            value = factors[name]

            factor_output.append({

                "Sample": sample,

                "Factor": name,

                "Independent_Value": value,

                "Original_XLS_Value": original["f8"],

                "Extracted_Report_Value": number(
                    extracted.get("f8")
                ),

                "Original_Difference": (
                    value - original["f8"]
                    if original["f8"] is not None
                    else None
                ),

                "Report_Difference": None,

                "Original_Check": (
                    "PASS"
                    if compare(value, original["f8"])
                    else "REVIEW"
                ),

                "Report_Check": "DIAGNOSTIC"

            })

        # -----------------------------------------------------
        # Averaging audit
        # -----------------------------------------------------

        averaging = {

            "Arithmetic_12_9":
                factors["fav_arithmetic_12_9"],

            "Arithmetic_13":
                factors["fav_arithmetic_13"],

            "Weighted_12_9":
                factors["fav_weighted_12_9"],

            "Weighted_13":
                factors["fav_weighted_13"]

        }

        for method, fav in averaging.items():

            iron = calculate_iron(
                c,
                fav
            )

            iron_output.append({

                "Sample": sample,

                "Averaging_Method": method,

                "fA": factors["fA"],

                "fB_12_9": factors["fB_12_9"],

                "fB_13": factors["fB_13"],

                "fav": fav,

                "Original_fav": original["Original_fav"],

                "fav_Difference": (
                    fav - original["Original_fav"]
                    if original["Original_fav"] is not None
                    else None
                ),

                "Fe3_Raw": iron["Fe3_Raw"],

                "Original_AR_Fe3": original["Fe3_AR"],

                "Fe3_Final": iron["Fe3_Final"],

                "Original_H_Fe3": original["Fe3_H"],

                "Fe2_Final": iron["Fe2_Final"],

                "Original_AW_Fe2": original["Fe2_AW"],

                "Constraint_Valid_12_9":
                    factors["Constraint_Valid_12_9"],

                "Constraint_Valid_13":
                    factors["Constraint_Valid_13"],

                "Fe2_Negative": iron["Fe2_Negative"],

                "Fe3_Exceeds_Revised_Fe":
                    iron["Fe3_Exceeds_Revised_Fe"]

            })

        # -----------------------------------------------------
        # Sample summary
        # -----------------------------------------------------

        site_valid = not any([
            sites["T_Overfilled"],
            sites["C_Overfilled"],
            sites["B_Overfilled"]
        ])

        original_factors = [
            factors[name]
            for name in factor_names
        ]

        summary_output.append({

            "Sample": sample,

            "Initial_Cation_Sum": sum(c.values()),

            "Site_Allocation_Status": (
                "PASS" if site_valid else "REVIEW"
            ),

            "fA": factors["fA"],

            "fB_12_9": factors["fB_12_9"],

            "fB_13": factors["fB_13"],

            "Original_fav": original["Original_fav"],

            "Arithmetic_fav_12_9":
                factors["fav_arithmetic_12_9"],

            "Weighted_fav_12_9":
                factors["fav_weighted_12_9"],

            "Constraint_Valid_12_9":
                factors["Constraint_Valid_12_9"],

            "Constraint_Valid_13":
                factors["Constraint_Valid_13"],

            "Original_H_Fe3": original["Fe3_H"],

            "Original_AR_Fe3": original["Fe3_AR"],

            "Original_AX_Fe3": original["Fe3_AX"],

            "Original_AW_Fe2": original["Fe2_AW"],

            "Scientific_Status": "REQUIRES REVIEW"

        })

        logger.info(
            "Processed %s",
            sample
        )

    # =============================================================================
    # K. REPORT
    # =============================================================================

    wb = openpyxl.Workbook()

    readme = wb.active
    readme.title = "Read_Me"

    for row in [

        ["Item", "Description"],

        ["Step", "06.3.1 corrected independent verification"],

        ["Samples", len(samples)],

        ["Source workbook SHA256", source_hash],

        ["Input cation basis", "Original independent reconstruction"],

        ["Site allocation", "Explicit computational reconstruction"],

        ["f8 variants", "12.9 and 13.0 both reported"],

        ["Averaging variants", "Arithmetic and weighted both reported"],

        ["Tolerance", TOL],

        ["Certification", "NOT ESTABLISHED"],

        ["Important limitation",
         "Site allocation is not independent crystallographic evidence."],

        ["Execution", datetime.now().isoformat()]

    ]:
        readme.append(row)

    write_records(
        wb,
        "Site_Allocation",
        site_output
    )

    write_records(
        wb,
        "Factor_Reconciliation",
        factor_output
    )

    write_records(
        wb,
        "Iron_and_Averaging",
        iron_output
    )

    write_records(
        wb,
        "Sample_Summary",
        summary_output
    )

    # ---------------------------------------------------------
    # Overall report
    # ---------------------------------------------------------

    ws = wb.create_sheet(
        "Overall_Summary",
        1
    )

    ws.append(["Metric", "Value"])

    ws.append(["Samples", len(samples)])

    ws.append([
        "Factor records",
        len(factor_output)
    ])

    ws.append([
        "Site allocation records",
        len(site_output)
    ])

    ws.append([
        "Iron/averaging records",
        len(iron_output)
    ])

    ws.append([
        "All samples independently certified",
        "NO"
    ])

    ws.append([
        "Scientific closure",
        "PENDING"
    ])

    ws.append([
        "ML training approval",
        "NOT GRANTED BY THIS AUDIT"
    ])

    # ---------------------------------------------------------
    # Limitations
    # ---------------------------------------------------------

    limitations = [

        {
            "Issue": "Site allocation",

            "Finding": "Site terms were reconstructed computationally.",

            "Limitation": "No direct crystallographic occupancy evidence.",

            "Required_Evidence": "Independent structural or published site-allocation evidence."
        },

        {
            "Issue": "f8 numerator",

            "Finding": "Both 12.9 and 13.0 variants are retained.",

            "Limitation": "The exact original workbook convention must be confirmed.",

            "Required_Evidence": "Original formula cell and published equation."
        },

        {
            "Issue": "Averaging",

            "Finding": "Arithmetic and weighted means are both calculated.",

            "Limitation": "The appropriate version must be established from the cited method and workbook.",

            "Required_Evidence": "Original formula and methodological source."
        },

        {
            "Issue": "Fe3+ provenance",

            "Finding": "Fe3+ is calculated from stoichiometric constraints.",

            "Limitation": "Not an independent spectroscopic determination.",

            "Required_Evidence": "Independent oxidation-state measurement where available."
        },

        {
            "Issue": "Scientific calibration",

            "Finding": "Arithmetic reproduction is assessed.",

            "Limitation": "Does not establish geological accuracy.",

            "Required_Evidence": "Independent experimental and geological validation."
        }

    ]

    write_records(
        wb,
        "Scientific_Limitations",
        limitations
    )

    format_workbook(wb)

    wb.save(OUTPUT_FILE)

    # =============================================================================
    # L. CONSOLE SUMMARY
    # =============================================================================

    logger.info("=" * 79)

    logger.info(
        "STEP 06.3.1 COMPLETED"
    )

    logger.info(
        "Samples: %d",
        len(samples)
    )

    logger.info(
        "Factor records: %d",
        len(factor_output)
    )

    logger.info(
        "Site records: %d",
        len(site_output)
    )

    logger.info(
        "Iron and averaging records: %d",
        len(iron_output)
    )

    logger.info(
        "Scientific certification: NOT ESTABLISHED"
    )

    logger.info(
        "Report: %s",
        OUTPUT_FILE
    )

    logger.info(
        "Log: %s",
        LOG_FILE
    )

    logger.info("=" * 79)


# =============================================================================
# M. RUN
# =============================================================================

if __name__ == "__main__":

    try:
        main()

    except Exception:

        logger.exception(
            "STEP 06.3.1 FAILED"
        )

        raise
