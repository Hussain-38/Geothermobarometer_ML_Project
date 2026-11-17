
# =============================================================================
# STEP 06 — HBL97 Fe3+ PROVENANCE AND SITE-TERM AUDIT
#
# Project: Geothermobarometer ML Project
# Purpose:
#   1. Reconstruct original Hbl97 workbook data provenance.
#   2. Compare original and independent chemical reconstructions.
#   3. Audit f1-f10 optimization factors and fav.
#   4. Audit Fe2+/Fe3+ partitioning.
#   5. Audit revised and Sigma13-normalized cations.
#   6. Identify outstanding mineralogical site-allocation evidence.
#
# IMPORTANT:
# Arithmetic agreement is not equivalent to scientific validation.
# No unverified mineralogical assumptions are introduced.
# =============================================================================

from pathlib import Path
from datetime import datetime
import logging
import math
import hashlib
import xlrd
import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. PROJECT CONFIGURATION
# =============================================================================

ROOT = Path(__file__).resolve().parent

SOURCE_XLS = (
    ROOT /
    "Geothermobarometers Calibrations" /
    "Hbl P-T (1997).xls"
)

INPUT_CHEM = (
    ROOT /
    "05C4_1_Hbl97_original_chemical_reconstruction" /
    "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

INPUT_AUDIT = (
    ROOT /
    "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit" /
    "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

OUTPUT_DIR = (
    ROOT /
    "06_Hbl97_Fe3_Provenance_and_SiteTerm_Audit"
)

OUTPUT_DIR.mkdir(
    parents=True,
    exist_ok=True
)

OUTPUT_FILE = (
    OUTPUT_DIR /
    "Step06_Hbl97_Fe3_Provenance_and_SiteTerm_Audit.xlsx"
)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

LOG_FILE = LOG_DIR / "step06_execution_log.txt"

TOLERANCE = 1e-5


# =============================================================================
# 2. LOGGING
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
    ],
    force=True
)

logger = logging.getLogger("STEP06")


# =============================================================================
# 3. GENERAL FUNCTIONS
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


def safe_float(value):

    if value is None:
        return None

    if isinstance(value, str):

        value = value.strip()

        if value == "":
            return None

    try:

        result = float(value)

        if math.isfinite(result):
            return result

    except (ValueError, TypeError):

        pass

    return None


def difference(a, b):

    a = safe_float(a)
    b = safe_float(b)

    if a is None or b is None:
        return None

    return a - b


def check_difference(a, b, tolerance=TOLERANCE):

    d = difference(a, b)

    if d is None:
        return "NOT EVALUATED"

    return (
        "PASS"
        if abs(d) <= tolerance
        else "REVIEW"
    )


def read_xlsx_sheet(path, sheet_name):

    if not path.exists():

        raise FileNotFoundError(
            f"Missing input workbook: {path}"
        )

    wb = openpyxl.load_workbook(
        path,
        data_only=True,
        read_only=True
    )

    if sheet_name not in wb.sheetnames:

        raise ValueError(
            f"Worksheet '{sheet_name}' not found in {path.name}. "
            f"Available sheets: {wb.sheetnames}"
        )

    ws = wb[sheet_name]

    rows = list(
        ws.iter_rows(values_only=True)
    )

    if not rows:

        wb.close()

        return []

    headers = [
        str(x).strip() if x is not None else ""
        for x in rows[0]
    ]

    records = []

    for row in rows[1:]:

        record = {}

        for i, header in enumerate(headers):

            if header:

                record[header] = (
                    row[i]
                    if i < len(row)
                    else None
                )

        if any(v is not None for v in record.values()):

            records.append(record)

    wb.close()

    return records


def make_index(records):

    result = {}

    for record in records:

        sample = record.get("Sample")

        if sample is None:
            continue

        key = str(sample).strip()

        if key:
            result[key] = record

    return result


def xl_column(col):

    return xlrd.formula.colname(col).upper()


# =============================================================================
# 4. VALIDATE INPUT FILES
# =============================================================================

logger.info("=" * 79)
logger.info("STEP 06 — HBL97 Fe3+ PROVENANCE AND SITE-TERM AUDIT")
logger.info("=" * 79)

for path in [
    SOURCE_XLS,
    INPUT_CHEM,
    INPUT_AUDIT
]:

    if not path.exists():

        raise FileNotFoundError(
            f"Required source not found:\n{path}"
        )

    logger.info("Verified input: %s", path)

source_hash = sha256_file(SOURCE_XLS)

logger.info("Original XLS SHA256: %s", source_hash)


# =============================================================================
# 5. READ INDEPENDENT RECONSTRUCTION
# =============================================================================

logger.info("Reading independent chemical reconstruction...")

chem_records = read_xlsx_sheet(
    INPUT_CHEM,
    "Independent_Reconstruction"
)

chem_index = make_index(chem_records)

logger.info(
    "Independent reconstruction records: %d",
    len(chem_records)
)


# =============================================================================
# 6. READ ORIGINAL OPTIMIZATION FACTORS
# =============================================================================

orig_factor_records = read_xlsx_sheet(
    INPUT_CHEM,
    "Optimization_Factors"
)

factor_index = make_index(
    orig_factor_records
)


# =============================================================================
# 7. READ STEP 05C.4.2 AUDIT
# =============================================================================

logger.info("Reading Fe3+ optimization audit...")

optimization_records = read_xlsx_sheet(
    INPUT_AUDIT,
    "Optimization_Audit"
)

optimization_index = make_index(
    optimization_records
)

iron_records = read_xlsx_sheet(
    INPUT_AUDIT,
    "Iron_Partition_Audit"
)

iron_index = make_index(
    iron_records
)

cation_records = read_xlsx_sheet(
    INPUT_AUDIT,
    "Revised_Cation_Audit"
)

sigma_records = read_xlsx_sheet(
    INPUT_AUDIT,
    "Sigma13_Normalization"
)

integrated_records = read_xlsx_sheet(
    INPUT_AUDIT,
    "Integrated_Sample_Audit"
)

limitation_records = read_xlsx_sheet(
    INPUT_AUDIT,
    "Scientific_Limitations"
)


# =============================================================================
# 8. READ ORIGINAL LEGACY XLS WORKBOOK
# =============================================================================

logger.info("Opening original Hbl97 XLS workbook...")

xls = xlrd.open_workbook(
    str(SOURCE_XLS),
    on_demand=True
)

ws = xls.sheet_by_name(
    "Hbl P-T (1997)"
)

# Column indices: zero-based.
ORIGINAL_COLUMNS = {

    "Publication": 0,
    "Sample": 1,

    "Input_T_C": 3,
    "Input_P_kbar": 4,

    "Calculated_T_C": 5,
    "Calculated_P_kbar": 6,

    "Fe3_H": 7,

    "SiO2": 9,
    "TiO2": 10,
    "Al2O3": 11,
    "FeO_total": 12,
    "MnO": 13,
    "MgO": 14,
    "CaO": 15,
    "Na2O": 16,
    "K2O": 17,

    "Ratio_Hbl": 19,

    "Si_initial": 20,
    "Ti_initial": 21,
    "Al_initial": 22,
    "Fe_total_initial": 23,
    "Mn_initial": 24,
    "Mg_initial": 25,
    "Ca_initial": 26,
    "Na_initial": 27,
    "K_initial": 28,
    "Initial_sum": 29,

    "f1": 30,
    "f2": 31,
    "f3": 32,
    "f4": 33,
    "f5": 34,

    "Minimum_f1_f5": 35,

    "f6": 36,
    "f7": 37,
    "f8": 38,
    "f9": 39,
    "f10": 40,

    "Maximum_f6_f10": 41,
    "fav": 42,

    "Fe3_AR": 43,

    "Si_revised": 45,
    "Ti_revised": 46,
    "Al_revised": 47,
    "Fe2_revised": 48,
    "Fe3_revised": 49,
    "Mn_revised": 50,
    "Mg_revised": 51,
    "Ca_revised": 52,
    "Na_revised": 53,
    "K_revised": 54,
    "Revised_sum": 55,

    "Sum_Mn": 57,
    "Ratio_Sigma13": 58,

    "Si_sigma13": 59,
    "Ti_sigma13": 60,
    "Al_sigma13": 61,
    "Fe2_sigma13": 62,
    "Fe3_sigma13": 63,
    "Mn_sigma13": 64,
    "Mg_sigma13": 65,

    "T97": 66,
    "P97": 67
}


original_records = []

for row in range(9, ws.nrows):

    sample = ws.cell_value(
        row,
        ORIGINAL_COLUMNS["Sample"]
    )

    if not sample:
        continue

    publication = ws.cell_value(
        row,
        ORIGINAL_COLUMNS["Publication"]
    )

    record = {

        "Excel_Row": row + 1,

        "Publication": publication,

        "Sample": str(sample).strip(),

    }

    for name, col in ORIGINAL_COLUMNS.items():

        if name in [
            "Publication",
            "Sample"
        ]:
            continue

        record[name] = safe_float(
            ws.cell_value(row, col)
        )

    original_records.append(record)
original_header_map = {

    name: (
        str(ws.cell_value(7, col)),
        str(ws.cell_value(8, col))
    )

    for name, col in ORIGINAL_COLUMNS.items()
}
xls.release_resources()

original_index = make_index(
    original_records
)

logger.info(
    "Original numerical sample records: %d",
    len(original_records)
)


# =============================================================================
# 9. ORIGINAL WORKBOOK COLUMN MAP
# =============================================================================

column_map = []

for name, col in ORIGINAL_COLUMNS.items():

    column_map.append({

        "Variable": name,

        "Excel_Column": xl_column(col),

        "Column_Number": col + 1,

      "Original_Header_Row8": original_header_map[name][0],
            str(ws.cell_value(7, col))
            if col < ws.ncols
            else ""
       

      "Original_Header_Row9": original_header_map[name][1],
            str(ws.cell_value(8, col))
            if col < ws.ncols
            else ""


        "Scientific_Status": (
            "Original workbook field; interpretation "
            "must follow documented methodology."
        )

    })


# =============================================================================
# 10. BUILD SAMPLE-LEVEL PROVENANCE AUDIT
# =============================================================================

logger.info("Building sample-level provenance audit...")

provenance_records = []

for sample, original in original_index.items():

    chem = chem_index.get(sample, {})
    factor = factor_index.get(sample, {})
    optimization = optimization_index.get(sample, {})
    iron = iron_index.get(sample, {})

    record = {

        "Sample": sample,

        "Publication": original["Publication"],

        "Original_Excel_Row": original["Excel_Row"],

        "Original_Fe3_H": original["Fe3_H"],

        "Original_Fe3_AR": original["Fe3_AR"],

        "Independent_Fe3_Input": chem.get(
            "Fe3_Hbl_Original"
        ),

        "Independent_Fe3_Separate": chem.get(
            "Fe3_Hbl_Original_Separate"
        ),

        "Original_Implied_Fe2": chem.get(
            "Fe2_Implied_23O"
        ),

        "Original_Implied_Fe3_Ratio": chem.get(
            "Fe3_FeTotal_Ratio_Implied"
        ),

        "Fe3_Origin_From_Audit": iron.get(
            "Fe3_Origin"
        ),

        "Original_Fe3_H_Negative": (
            "YES"
            if (
                original["Fe3_H"] is not None
                and original["Fe3_H"] < 0
            )
            else "NO"
        ),

        "Original_H_vs_Reconstruction": check_difference(
            original["Fe3_H"],
            chem.get("Fe3_Hbl_Original")
        ),

        "Original_AR_vs_Separate_Fe3": check_difference(
            original["Fe3_AR"],
            chem.get("Fe3_Hbl_Original_Separate")
        )

    }

    provenance_records.append(record)


# =============================================================================
# 11. OPTIMIZATION FACTOR AUDIT
# =============================================================================

logger.info("Auditing optimization factors...")

factor_audit = []

for sample, original in original_index.items():

    factor = factor_index.get(sample, {})
    optimization = optimization_index.get(sample, {})

    values = {
        f"f{i}": original.get(f"f{i}")
        for i in range(1, 11)
    }

    f1_f5 = [
        values[f"f{i}"]
        for i in range(1, 6)
    ]

    f6_f10 = [
        values[f"f{i}"]
        for i in range(6, 11)
    ]

    valid_first = [
        x for x in f1_f5
        if x is not None
    ]

    valid_second = [
        x for x in f6_f10
        if x is not None
    ]

    calculated_min = (
        min(valid_first)
        if len(valid_first) == 5
        else None
    )

    calculated_max = (
        max(valid_second)
        if len(valid_second) == 5
        else None
    )

    row = {

        "Sample": sample,

        **values,

        "Original_Minimum": original.get(
            "Minimum_f1_f5"
        ),

        "Reconstructed_Minimum": factor.get(
            "Minimum_f1_f5"
        ),

        "Calculated_Minimum": calculated_min,

        "Minimum_Arithmetic_Check": check_difference(
            calculated_min,
            original.get("Minimum_f1_f5")
        ),

        "Original_Maximum": original.get(
            "Maximum_f6_f10"
        ),

        "Reconstructed_Maximum": factor.get(
            "Maximum_f6_f10"
        ),

        "Calculated_Maximum": calculated_max,

        "Maximum_Arithmetic_Check": check_difference(
            calculated_max,
            original.get("Maximum_f6_f10")
        ),

        "Original_fav": original.get("fav"),

        "Reconstructed_fav": factor.get("fav"),

        "Audit_fav": optimization.get(
            "fav_Original"
        ),

        "fav_Reconstruction_Check": check_difference(
            original.get("fav"),
            optimization.get("fav_Original")
        )

    }

    factor_audit.append(row)


# =============================================================================
# 12. IRON PARTITION AUDIT
# =============================================================================

logger.info("Auditing Fe2+/Fe3+ partitioning...")

iron_audit = []

for sample, original in original_index.items():

    iron = iron_index.get(sample, {})

    initial_total = iron.get(
        "Initial_Fe_Total"
    )

    fav = iron.get("fav")

    predicted_fe2 = iron.get(
        "Predicted_Fe2"
    )

    original_fe2 = iron.get(
        "Original_Revised_Fe2"
    )

    original_fe3 = iron.get(
        "Original_Revised_Fe3"
    )

    row = {

        "Sample": sample,

        "Original_Fe3_H": original.get(
            "Fe3_H"
        ),

        "Original_Fe3_AR": original.get(
            "Fe3_AR"
        ),

        "Initial_Fe_Total": initial_total,

        "fav": fav,

        "Predicted_Revised_Fe_Total": iron.get(
            "Predicted_Revised_Fe_Total"
        ),

        "Original_H_Fe3": iron.get(
            "Original_H_Fe3"
        ),

        "Predicted_Fe2": predicted_fe2,

        "Original_Revised_Fe2": original_fe2,

        "Original_Revised_Fe3": original_fe3,

        "Fe2_Check": iron.get(
            "Fe2_Prediction_Check"
        ),

        "Fe3_H_vs_AX_Check": iron.get(
            "Fe3_H_vs_AX_Check"
        ),

        "Fe2_plus_Fe3": iron.get(
            "Fe2_plus_Fe3"
        ),

        "Fe_Total_Check": iron.get(
            "Fe_Total_Check"
        ),

        "Implied_Fe3_FeTotal_Ratio": iron.get(
            "Implied_Fe3_FeTotal_Ratio"
        ),

        "Fe3_Origin": iron.get(
            "Fe3_Origin"
        )

    }

    iron_audit.append(row)


# =============================================================================
# 13. REVISED CATION AUDIT
# =============================================================================

logger.info("Building revised cation comparison...")

cation_audit = []

revised_mapping = {

    "Si": "Si_revised",
    "Ti": "Ti_revised",
    "Al": "Al_revised",
    "Fe2": "Fe2_revised",
    "Fe3": "Fe3_revised",
    "Mn": "Mn_revised",
    "Mg": "Mg_revised",
    "Ca": "Ca_revised",
    "Na": "Na_revised",
    "K": "K_revised"

}

for sample, original in original_index.items():

    for element, original_field in revised_mapping.items():

        value = original.get(original_field)

        cation_audit.append({

            "Sample": sample,

            "Element": element,

            "Original_Workbook_Revised_Value": value,

            "Interpretation": (
                "Original workbook value; independently "
                "validated site assignment not established."
            )

        })


for record in cation_records:

    sample = record.get("Sample")
    element = record.get("Element")

    if sample is None:
        continue

    for row in cation_audit:

        if (
            row["Sample"] == str(sample).strip()
            and row["Element"] == element
        ):

            row["Independent_Initial_Cation"] = record.get(
                "Initial_Cation"
            )

            row["Independent_fav"] = record.get(
                "fav"
            )

            row["Predicted_Revised"] = record.get(
                "Predicted_Revised"
            )

            row["Audit_Original_Revised"] = record.get(
                "Original_Revised"
            )

            row["Difference"] = record.get(
                "Difference"
            )

            row["Check"] = record.get("Check")

            break


# =============================================================================
# 14. SIGMA13 NORMALIZATION AUDIT
# =============================================================================

logger.info("Building Sigma13 audit...")

sigma_audit = []

sigma_mapping = {

    "Si": "Si_sigma13",
    "Ti": "Ti_sigma13",
    "Al": "Al_sigma13",
    "Fe2": "Fe2_sigma13",
    "Fe3": "Fe3_sigma13",
    "Mn": "Mn_sigma13",
    "Mg": "Mg_sigma13"

}

for sample, original in original_index.items():

    matching = [
        x for x in sigma_records
        if str(x.get("Sample", "")).strip() == sample
    ]

    sigma = matching[0] if matching else {}

    for element, field in sigma_mapping.items():

        sigma_audit.append({

            "Sample": sample,

            "Element": element,

            "Original_Sigma13_Value": original.get(
                field
            ),

            "Independent_Predicted_Value": sigma.get(
                f"{element}_Predicted"
            ),

            "Independent_Original_Value": sigma.get(
                f"{element}_Original"
            ),

            "Check": sigma.get(
                f"{element}_Check"
            ),

            "Scientific_Interpretation": (
                "Numerical normalization audit only; "
                "not a certified crystallographic site assignment."
            )

        })


# =============================================================================
# 15. SITE-TERM EVIDENCE AUDIT
# =============================================================================

logger.info("Evaluating mineralogical site-term evidence...")

site_terms = [

    (
        "Tetrahedral site (T)",
        "Si, Al",
        "Independent site allocation and site-capacity verification"
    ),

    (
        "C1 and C2 sites",
        "Mg, Fe2+, Fe3+, Al, Ti, Mn",
        "Documented allocation scheme and charge-balance verification"
    ),

    (
        "B site",
        "Ca, Na",
        "Documented B-site occupancy and compositional constraints"
    ),

    (
        "A site",
        "Na, K and vacancies",
        "Documented A-site allocation and vacancy calculation"
    ),

    (
        "Total structural formula",
        "All relevant cations and anions",
        "Complete formula and charge-balance verification"
    ),

    (
        "Fe oxidation state",
        "Fe2+ and Fe3+",
        "Independent evidence or explicitly documented estimation procedure"
    )

]

site_evidence = []

for sample in original_index:

    for site, elements, requirement in site_terms:

        site_evidence.append({

            "Sample": sample,

            "Site_or_Property": site,

            "Relevant_Elements": elements,

            "Evidence_Available": (
                "Original workbook numerical values available"
            ),

            "Required_Evidence": requirement,

            "Scientific_Status": "NOT CERTIFIED",

            "Reason": (
                "Numerical cation reconstruction alone does not "
                "establish crystallographic site occupancy."
            )

        })


# =============================================================================
# 16. INTEGRATED AUDIT SUMMARY
# =============================================================================

integrated_audit = []

for record in integrated_records:

    integrated_audit.append({

        "Sample": record.get("Sample"),

        "Total_Checks": record.get("Total_Checks"),

        "PASS": record.get("PASS"),

        "FAIL": record.get("FAIL"),

        "NOT_EVALUATED": record.get("NOT_EVALUATED"),

        "Arithmetic_Audit": record.get(
            "Arithmetic_Audit"
        ),

        "Interpretation": (
            "Arithmetic agreement only; not independent "
            "mineralogical validation."
        )

    })


# =============================================================================
# 17. SCIENTIFIC LIMITATIONS
# =============================================================================

limitations = [

    {
        "Issue": "Fe3+ provenance",
        "Finding": "Original H and AR fields are separate.",
        "Limitation": "Their complete computational relationship must be established.",
        "Required_Evidence": "Original cell formulas and published method."
    },

    {
        "Issue": "Negative Fe3+",
        "Finding": "Original workbook instructs replacing negative H values with zero.",
        "Limitation": "This instruction does not independently validate oxidation state.",
        "Required_Evidence": "Original methodology and independent chemical evidence."
    },

    {
        "Issue": "Optimization",
        "Finding": "Ten factors and fav are explicitly recorded.",
        "Limitation": "Arithmetic reproduction does not prove geological validity.",
        "Required_Evidence": "Original optimization equations and constraints."
    },

    {
        "Issue": "Sigma13",
        "Finding": "Normalized cation values are recorded.",
        "Limitation": "Normalization is not equivalent to complete site allocation.",
        "Required_Evidence": "Published structural formula procedure."
    },

    {
        "Issue": "Crystallographic sites",
        "Finding": "Numerical cations are available.",
        "Limitation": "Site occupancy has not been certified.",
        "Required_Evidence": "Site allocation rules, anion assumptions and charge balance."
    },

    {
        "Issue": "ML readiness",
        "Finding": "Thirteen sample reconstructions are available.",
        "Limitation": "This is insufficient to certify a training dataset.",
        "Required_Evidence": "Publication-level validation and independent sample records."
    },

    {
        "Issue": "Publication provenance",
        "Finding": "All thirteen original sample rows identify Xiao et al. (2011).",
        "Limitation": "They must not automatically be attributed to Gerya et al. (1997).",
        "Required_Evidence": "Original publication and calibration provenance."
    }

]


# =============================================================================
# 18. CREATE EXCEL REPORT
# =============================================================================

logger.info("Creating consolidated Step 06 Excel report...")

output_wb = Workbook()

output_wb.remove(
    output_wb.active
)


def write_sheet(wb, name, records):

    ws = wb.create_sheet(name)

    if not records:

        ws.append(["No records available"])

        return ws

    headers = list(records[0].keys())

    ws.append(headers)

    for record in records:

        ws.append([
            record.get(header)
            for header in headers
        ])

    ws.freeze_panes = "A2"

    ws.auto_filter.ref = ws.dimensions

    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    for col in ws.columns:

        max_length = 0

        column_letter = get_column_letter(
            col[0].column
        )

        for cell in col:

            if cell.value is not None:

                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        ws.column_dimensions[
            column_letter
        ].width = min(
            max(max_length + 2, 12),
            45
        )

    return ws


# Read me
readme = [

    {
        "Item": "Project",
        "Description": "Hbl97 Fe3+ provenance and site-term audit"
    },

    {
        "Item": "Original workbook",
        "Description": str(SOURCE_XLS)
    },

    {
        "Item": "Original SHA256",
        "Description": source_hash
    },

    {
        "Item": "Sample count",
        "Description": len(original_records)
    },

    {
        "Item": "Independent reconstruction records",
        "Description": len(chem_records)
    },

    {
        "Item": "Optimization audit records",
        "Description": len(optimization_records)
    },

    {
        "Item": "Numerical tolerance",
        "Description": TOLERANCE
    },

    {
        "Item": "Scientific status",
        "Description": "AUDIT ONLY — NOT CERTIFIED FOR ML"
    },

    {
        "Item": "Important",
        "Description": (
            "No new Fe3+ determination, crystallographic "
            "site allocation or geological validation is claimed."
        )
    }

]

write_sheet(
    output_wb,
    "Read_Me",
    readme
)

write_sheet(
    output_wb,
    "Original_Workbook_Data",
    original_records
)

write_sheet(
    output_wb,
    "Original_Column_Map",
    column_map
)

write_sheet(
    output_wb,
    "Fe3_Provenance",
    provenance_records
)

write_sheet(
    output_wb,
    "Optimization_Factor_Audit",
    factor_audit
)

write_sheet(
    output_wb,
    "Iron_Partition_Audit",
    iron_audit
)

write_sheet(
    output_wb,
    "Revised_Cation_Audit",
    cation_audit
)

write_sheet(
    output_wb,
    "Sigma13_Audit",
    sigma_audit
)

write_sheet(
    output_wb,
    "Site_Term_Evidence",
    site_evidence
)

write_sheet(
    output_wb,
    "Integrated_Arithmetic_Audit",
    integrated_audit
)

write_sheet(
    output_wb,
    "Scientific_Limitations",
    limitations
)


# =============================================================================
# 19. SAVE OUTPUT
# =============================================================================

output_wb.save(
    OUTPUT_FILE
)

output_wb.close()


# =============================================================================
# 20. EXECUTION SUMMARY
# =============================================================================

logger.info("=" * 79)
logger.info("STEP 06 COMPLETED")
logger.info("=" * 79)

logger.info(
    "Original samples: %d",
    len(original_records)
)

logger.info(
    "Independent reconstruction records: %d",
    len(chem_records)
)

logger.info(
    "Optimization audit records: %d",
    len(optimization_records)
)

logger.info(
    "Fe3 provenance records: %d",
    len(provenance_records)
)

logger.info(
    "Optimization factor records: %d",
    len(factor_audit)
)

logger.info(
    "Iron partition records: %d",
    len(iron_audit)
)

logger.info(
    "Revised cation records: %d",
    len(cation_audit)
)

logger.info(
    "Sigma13 records: %d",
    len(sigma_audit)
)

logger.info(
    "Site-term evidence records: %d",
    len(site_evidence)
)

logger.info(
    "Output workbook: %s",
    OUTPUT_FILE
)

logger.info(
    "Execution log: %s",
    LOG_FILE
)

logger.info("=" * 79)

print("\nSTEP 06 REPORT GENERATED SUCCESSFULLY")

print("\nExcel report:")
print(OUTPUT_FILE)

print("\nExecution log:")
print(LOG_FILE)

print(
    "\nIMPORTANT: Numerical agreement does not establish "
    "independent Fe3+ determination or crystallographic "
    "site-allocation validity."
)
