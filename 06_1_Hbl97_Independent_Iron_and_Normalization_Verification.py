# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 17:21:19 2026

@author: hussa
"""


# =============================================================================
# STEP 06.1 — HBL97 INDEPENDENT IRON AND NORMALIZATION VERIFICATION
#
# Objectives:
# 1. Reconcile original H and AR Fe3+ fields.
# 2. Recalculate optimization summaries from original f1-f10.
# 3. Correct Fe2/Fe3 element-label matching.
# 4. Recalculate revised-cation arithmetic diagnostics.
# 5. Recalculate Sigma13 scaling diagnostics.
# 6. Identify unresolved scientific assumptions.
#
# No unverified mineralogical equations are certified.
# =============================================================================

from pathlib import Path
from collections import Counter
import logging
import math
import hashlib
import re

import xlrd
import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. CONFIGURATION
# =============================================================================

ROOT = Path(__file__).resolve().parent

SOURCE_XLS = (
    ROOT / "Geothermobarometers Calibrations" /
    "Hbl P-T (1997).xls"
)

INPUT_CHEM = (
    ROOT / "05C4_1_Hbl97_original_chemical_reconstruction" /
    "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

INPUT_AUDIT = (
    ROOT / "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit" /
    "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

OUTPUT_DIR = (
    ROOT / "06_1_Hbl97_Independent_Iron_Normalization_Verification"
)

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

LOG_DIR = OUTPUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = (
    OUTPUT_DIR /
    "Step06_1_Hbl97_Independent_Verification.xlsx"
)

LOG_FILE = LOG_DIR / "step06_1_execution_log.txt"

TOL = 1e-5


# =============================================================================
# 2. LOGGING
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

logger = logging.getLogger("STEP06_1")


# =============================================================================
# 3. UTILITIES
# =============================================================================

def number(value):

    if value is None:
        return None

    try:
        x = float(value)

        if math.isfinite(x):
            return x

    except (ValueError, TypeError):
        pass

    return None


def compare(a, b, tolerance=TOL):

    a = number(a)
    b = number(b)

    if a is None or b is None:
        return "NOT EVALUATED"

    if abs(a - b) <= tolerance:
        return "PASS"

    return "REVIEW"


def norm_element(value):

    if value is None:
        return ""

    s = str(value).strip().lower()

    s = s.replace("⁺", "+")
    s = s.replace("²", "2")
    s = s.replace("³", "3")
    s = s.replace(" ", "")
    s = s.replace("_", "")

    aliases = {
        "fe2": "Fe2",
        "fe2+": "Fe2",
        "fe(2+)": "Fe2",
        "fe3": "Fe3",
        "fe3+": "Fe3",
        "fe(3+)": "Fe3",
        "si": "Si",
        "ti": "Ti",
        "al": "Al",
        "mn": "Mn",
        "mg": "Mg",
        "ca": "Ca",
        "na": "Na",
        "k": "K"
    }

    return aliases.get(s, str(value).strip())


def sha256(path):

    h = hashlib.sha256()

    with open(path, "rb") as f:

        for block in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):
            h.update(block)

    return h.hexdigest()


def read_sheet(path, sheet_name):

    if not path.exists():
        raise FileNotFoundError(path)

    wb = openpyxl.load_workbook(
        path,
        data_only=True,
        read_only=True
    )

    if sheet_name not in wb.sheetnames:

        raise ValueError(
            f"Missing sheet {sheet_name} in {path.name}. "
            f"Available: {wb.sheetnames}"
        )

    ws = wb[sheet_name]

    iterator = ws.iter_rows(values_only=True)

    headers = next(iterator)

    headers = [
        str(x).strip() if x is not None else ""
        for x in headers
    ]

    records = []

    for row in iterator:

        record = {
            headers[i]: row[i]
            for i in range(min(len(headers), len(row)))
            if headers[i]
        }

        if any(v is not None for v in record.values()):
            records.append(record)

    wb.close()

    return records


def index_samples(records):

    result = {}

    for record in records:

        sample = record.get("Sample")

        if sample is not None:

            key = str(sample).strip()

            result[key] = record

    return result


def save_sheet(wb, name, records):

    ws = wb.create_sheet(name)

    if not records:

        ws.append(["No records"])
        return ws

    headers = list(records[0].keys())

    ws.append(headers)

    for record in records:

        ws.append([
            record.get(h)
            for h in headers
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
            wrap_text=True,
            vertical="center"
        )

    for col in ws.columns:

        letter = get_column_letter(col[0].column)

        maximum = max(
            (
                len(str(c.value))
                for c in col
                if c.value is not None
            ),
            default=12
        )

        ws.column_dimensions[letter].width = min(
            max(maximum + 2, 12),
            42
        )

    return ws


# =============================================================================
# 4. VALIDATE INPUTS
# =============================================================================

logger.info("=" * 79)
logger.info("STEP 06.1 — INDEPENDENT IRON AND NORMALIZATION VERIFICATION")
logger.info("=" * 79)

for path in [
    SOURCE_XLS,
    INPUT_CHEM,
    INPUT_AUDIT
]:

    if not path.exists():
        raise FileNotFoundError(
            f"Missing required file: {path}"
        )

    logger.info("Input verified: %s", path)

source_hash = sha256(SOURCE_XLS)

logger.info("Original XLS SHA256: %s", source_hash)


# =============================================================================
# 5. LOAD PREVIOUS REPORTS
# =============================================================================

chem_records = read_sheet(
    INPUT_CHEM,
    "Independent_Reconstruction"
)

factor_records = read_sheet(
    INPUT_CHEM,
    "Optimization_Factors"
)

cation_records = read_sheet(
    INPUT_AUDIT,
    "Revised_Cation_Audit"
)

iron_records = read_sheet(
    INPUT_AUDIT,
    "Iron_Partition_Audit"
)

sigma_records = read_sheet(
    INPUT_AUDIT,
    "Sigma13_Normalization"
)

chem_idx = index_samples(chem_records)
factor_idx = index_samples(factor_records)
iron_idx = index_samples(iron_records)


# =============================================================================
# 6. READ ORIGINAL XLS
# =============================================================================

logger.info("Reading original XLS workbook...")

xls = xlrd.open_workbook(
    str(SOURCE_XLS),
    on_demand=True
)

ws = xls.sheet_by_name("Hbl P-T (1997)")


COL = {

    "Sample": 1,
    "Publication": 0,

    "Input_T": 3,
    "Input_P": 4,
    "Calculated_T": 5,
    "Calculated_P": 6,

    "Fe3_H": 7,

    "Si": 20,
    "Ti": 21,
    "Al": 22,
    "Fe_total": 23,
    "Mn": 24,
    "Mg": 25,
    "Ca": 26,
    "Na": 27,
    "K": 28,

    "Initial_sum": 29,

    "f1": 30,
    "f2": 31,
    "f3": 32,
    "f4": 33,
    "f5": 34,

    "Minimum": 35,

    "f6": 36,
    "f7": 37,
    "f8": 38,
    "f9": 39,
    "f10": 40,

    "Maximum": 41,
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
    "Ratio_13": 58,

    "Si_13": 59,
    "Ti_13": 60,
    "Al_13": 61,
    "Fe2_13": 62,
    "Fe3_13": 63,
    "Mn_13": 64,
    "Mg_13": 65,

    "T97": 66,
    "P97": 67
}


original = []

for r in range(9, ws.nrows):

    sample = ws.cell_value(r, COL["Sample"])

    if not sample:
        continue

    rec = {
        "Sample": str(sample).strip(),
        "Excel_Row": r + 1,
        "Publication": ws.cell_value(r, COL["Publication"])
    }

    for name, col in COL.items():

        if name in ["Sample", "Publication"]:
            continue

        rec[name] = number(
            ws.cell_value(r, col)
        )

    original.append(rec)

xls.release_resources()

original_idx = index_samples(original)

logger.info("Original samples recovered: %d", len(original))


# =============================================================================
# 7. FE3 PROVENANCE RECONCILIATION
# =============================================================================

logger.info("Reconciling H and AR Fe3 fields...")

fe3_audit = []

for sample, orig in original_idx.items():

    chem = chem_idx.get(sample, {})

    H = orig.get("Fe3_H")
    AR = orig.get("Fe3_AR")

    separate = number(
        chem.get("Fe3_Hbl_Original_Separate")
    )

    independent = number(
        chem.get("Fe3_Hbl_Original")
    )

    rec = {

        "Sample": sample,

        "Original_H": H,

        "Original_AR": AR,

        "H_minus_AR": (
            H - AR
            if H is not None and AR is not None
            else None
        ),

        "Independent_Fe3_Input": independent,

        "Independent_Fe3_Separate": separate,

        "H_vs_Independent": compare(
            H, independent
        ),

        "AR_vs_Independent_Separate": compare(
            AR, separate
        ),

        "Negative_H": (
            "YES" if H is not None and H < 0 else "NO"
        ),

        "Negative_AR": (
            "YES" if AR is not None and AR < 0 else "NO"
        ),

        "Provenance_Status": (
            "UNRESOLVED — analytical origin not established"
        )

    }

    if (
        H is not None
        and AR is not None
        and abs(H - AR) > TOL
    ):

        rec["Discrepancy_Flag"] = "REVIEW"

    else:

        rec["Discrepancy_Flag"] = "NO MATERIAL DIFFERENCE"

    fe3_audit.append(rec)


# =============================================================================
# 8. OPTIMIZATION FACTOR RECONSTRUCTION
# =============================================================================

logger.info("Recalculating optimization summaries...")

optimization_audit = []

for sample, orig in original_idx.items():

    first = [
        orig.get(f"f{i}")
        for i in range(1, 6)
    ]

    second = [
        orig.get(f"f{i}")
        for i in range(6, 11)
    ]

    complete_first = all(x is not None for x in first)
    complete_second = all(x is not None for x in second)

    calc_min = min(first) if complete_first else None
    calc_max = max(second) if complete_second else None

    # The midpoint relation is tested as an observed arithmetic
    # hypothesis, not asserted as the published scientific equation.
    calc_fav = (
        (calc_min + calc_max) / 2
        if calc_min is not None and calc_max is not None
        else None
    )

    previous = factor_idx.get(sample, {})

    optimization_audit.append({

        "Sample": sample,

        "f1": orig.get("f1"),
        "f2": orig.get("f2"),
        "f3": orig.get("f3"),
        "f4": orig.get("f4"),
        "f5": orig.get("f5"),

        "Original_Minimum": orig.get("Minimum"),

        "Recalculated_Minimum": calc_min,

        "Minimum_Check": compare(
            orig.get("Minimum"), calc_min
        ),

        "f6": orig.get("f6"),
        "f7": orig.get("f7"),
        "f8": orig.get("f8"),
        "f9": orig.get("f9"),
        "f10": orig.get("f10"),

        "Original_Maximum": orig.get("Maximum"),

        "Recalculated_Maximum": calc_max,

        "Maximum_Check": compare(
            orig.get("Maximum"), calc_max
        ),

        "Original_fav": orig.get("fav"),

        "Recalculated_Midpoint": calc_fav,

        "Midpoint_Check": compare(
            orig.get("fav"), calc_fav
        ),

        "Previous_Reconstruction_fav": previous.get("fav"),

        "Scientific_Status": (
            "ARITHMETIC VERIFIED; published factor definitions "
            "still require documentary confirmation."
        )

    })


# =============================================================================
# 9. REVISED CATION MATCHING AND AUDIT
# =============================================================================

logger.info("Correcting cation element matching...")

cation_index = {}

for record in cation_records:

    sample = str(record.get("Sample", "")).strip()

    element = norm_element(
        record.get("Element")
    )

    cation_index[(sample, element)] = record


ELEMENTS = [
    "Si", "Ti", "Al", "Fe2",
    "Fe3", "Mn", "Mg", "Ca", "Na", "K"
]

ORIGINAL_REVISED_FIELDS = {

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

revised_audit = []

for sample, orig in original_idx.items():

    chem = chem_idx.get(sample, {})

    for element in ELEMENTS:

        source = cation_index.get(
            (sample, element),
            {}
        )

        original_value = orig.get(
            ORIGINAL_REVISED_FIELDS[element]
        )

        predicted = number(
            source.get("Predicted_Revised")
        )

        if predicted is None:

            predicted = number(
                source.get("Original_Revised")
            )

        previous_difference = number(
            source.get("Difference")
        )

        revised_audit.append({

            "Sample": sample,

            "Element": element,

            "Original_Workbook_Revised": original_value,

            "Previous_Predicted_Revised": predicted,

            "Previous_Reported_Difference": previous_difference,

            "Independent_Comparison_Difference": (
                original_value - predicted
                if original_value is not None and predicted is not None
                else None
            ),

            "Comparison_Status": compare(
                original_value,
                predicted
            ),

            "Source_Record_Found": (
                "YES" if source else "NO"
            ),

            "Scientific_Status": (
                "NUMERICAL COMPARISON ONLY; site assignment "
                "and equation validity not certified."
            )

        })


# =============================================================================
# 10. INDEPENDENT IRON ARITHMETIC
# =============================================================================

logger.info("Recalculating iron partition diagnostics...")

iron_audit = []

for sample, orig in original_idx.items():

    previous = iron_idx.get(sample, {})

    initial_fe = orig.get("Fe_total")
    fav = orig.get("fav")
    fe3 = orig.get("Fe3_H")

    predicted_total = (
        initial_fe * fav
        if initial_fe is not None and fav is not None
        else None
    )

    predicted_fe2 = (
        predicted_total - fe3
        if predicted_total is not None and fe3 is not None
        else None
    )

    original_fe2 = orig.get("Fe2_revised")
    original_fe3 = orig.get("Fe3_revised")

    iron_sum = (
        original_fe2 + original_fe3
        if original_fe2 is not None and original_fe3 is not None
        else None
    )

    iron_audit.append({

        "Sample": sample,

        "Original_Fe3_H": fe3,

        "Original_Fe3_AR": orig.get("Fe3_AR"),

        "Initial_Fe_Total": initial_fe,

        "fav": fav,

        "Recalculated_Fe_Total_candidate": predicted_total,

        "Original_Revised_Fe2": original_fe2,

        "Recalculated_Fe2_candidate": predicted_fe2,

        "Original_Revised_Fe3": original_fe3,

        "Fe2_Arithmetic_Check": compare(
            original_fe2, predicted_fe2
        ),

        "Fe3_H_vs_Revised_Fe3": compare(
            fe3, original_fe3
        ),

        "Revised_Fe2_plus_Fe3": iron_sum,

        "Revised_FeTotal_Consistency": compare(
            iron_sum, predicted_total
        ),

        "Interpretation": (
            "Candidate arithmetic relation tested against workbook values. "
            "Not certified as a published equation."
        )

    })


# =============================================================================
# 11. SIGMA13 NORMALIZATION DIAGNOSTIC
# =============================================================================

logger.info("Recalculating Sigma13 scaling diagnostics...")

SIGMA_FIELDS = {

    "Si": ("Si_revised", "Si_13"),
    "Ti": ("Ti_revised", "Ti_13"),
    "Al": ("Al_revised", "Al_13"),
    "Fe2": ("Fe2_revised", "Fe2_13"),
    "Fe3": ("Fe3_revised", "Fe3_13"),
    "Mn": ("Mn_revised", "Mn_13"),
    "Mg": ("Mg_revised", "Mg_13")

}

sigma_audit = []

for sample, orig in original_idx.items():

    ratio = orig.get("Ratio_13")

    for element, fields in SIGMA_FIELDS.items():

        revised_field, sigma_field = fields

        revised_value = orig.get(revised_field)
        original_sigma = orig.get(sigma_field)

        predicted = (
            revised_value * ratio
            if revised_value is not None and ratio is not None
            else None
        )

        sigma_audit.append({

            "Sample": sample,

            "Element": element,

            "Original_Revised_Cation": revised_value,

            "Original_Ratio_Sigma13": ratio,

            "Original_Sigma13_Value": original_sigma,

            "Recalculated_Scaling_Candidate": predicted,

            "Scaling_Difference": (
                original_sigma - predicted
                if original_sigma is not None and predicted is not None
                else None
            ),

            "Scaling_Check": compare(
                original_sigma, predicted
            ),

            "Interpretation": (
                "Scaling identity diagnostic only. The scientific "
                "basis and denominator of the normalization factor "
                "remain to be verified."
            )

        })


# =============================================================================
# 12. HB32B TARGETED INVESTIGATION
# =============================================================================

logger.info("Preparing HB32B targeted investigation...")

targeted = []

for sample in original_idx:

    if sample != "HB32B(M2)":
        continue

    orig = original_idx[sample]
    chem = chem_idx.get(sample, {})
    previous = iron_idx.get(sample, {})

    targeted.append({

        "Sample": sample,

        "Original_Excel_Row": orig.get("Excel_Row"),

        "H_Fe3": orig.get("Fe3_H"),

        "AR_Fe3": orig.get("Fe3_AR"),

        "Independent_Fe3": chem.get(
            "Fe3_Hbl_Original"
        ),

        "Separate_Independent_Fe3": chem.get(
            "Fe3_Hbl_Original_Separate"
        ),

        "Revised_Fe3": orig.get("Fe3_revised"),

        "Revised_Fe2": orig.get("Fe2_revised"),

        "Original_Implied_Fe2": chem.get(
            "Fe2_Implied_23O"
        ),

        "Previous_Fe3_Origin": previous.get(
            "Fe3_Origin"
        ),

        "Required_Action": (
            "Inspect original Excel formula/value provenance and "
            "publication method. Do not replace AR with zero "
            "without documentary justification."
        )

    })


# =============================================================================
# 13. QUALITY CONTROL SUMMARY
# =============================================================================

def count_status(records, field):

    return Counter(
        r.get(field)
        for r in records
    )


summary = []

for name, records, field in [

    ("Fe3 provenance", fe3_audit, "Discrepancy_Flag"),

    ("Minimum", optimization_audit, "Minimum_Check"),

    ("Maximum", optimization_audit, "Maximum_Check"),

    ("fav midpoint", optimization_audit, "Midpoint_Check"),

    ("Revised cations", revised_audit, "Comparison_Status"),

    ("Iron Fe2", iron_audit, "Fe2_Arithmetic_Check"),

    ("Iron Fe3", iron_audit, "Fe3_H_vs_Revised_Fe3"),

    ("Sigma13 scaling", sigma_audit, "Scaling_Check")

]:

    counts = count_status(records, field)

    summary.append({

        "Audit": name,

        "Total_Records": len(records),

        "PASS": counts.get("PASS", 0),

        "REVIEW": counts.get("REVIEW", 0),

        "NO_MATERIAL_DIFFERENCE": counts.get(
            "NO MATERIAL DIFFERENCE", 0
        ),

        "NOT_EVALUATED": counts.get(
            "NOT EVALUATED", 0
        )

    })


# =============================================================================
# 14. SCIENTIFIC DECISION REGISTER
# =============================================================================

decisions = [

    {
        "Issue": "Xiao et al. (2011) sample provenance",
        "Status": "UNRESOLVED",
        "Reason": "Exact HB sample-to-publication match not established.",
        "Action": "Obtain original publication and supplementary tables."
    },

    {
        "Issue": "H versus AR Fe3",
        "Status": "REQUIRES REVIEW",
        "Reason": "HB32B has different H and AR values.",
        "Action": "Trace original cell equations and input provenance."
    },

    {
        "Issue": "f1-f10",
        "Status": "ARITHMETIC AUDIT",
        "Reason": "Factor summaries can be recalculated, but definitions remain unverified.",
        "Action": "Verify equations against original methodology."
    },

    {
        "Issue": "fav",
        "Status": "ARITHMETIC AUDIT",
        "Reason": "Midpoint identity is tested, not assumed as published method.",
        "Action": "Verify optimization definition."
    },

    {
        "Issue": "Iron partition",
        "Status": "CANDIDATE RELATION TESTED",
        "Reason": "Numerical identity is not proof of Fe oxidation-state determination.",
        "Action": "Verify Fe3 source and Fe2 derivation."
    },

    {
        "Issue": "Sigma13",
        "Status": "SCALING AUDIT",
        "Reason": "Scaling is checked without certifying normalization basis.",
        "Action": "Verify the published denominator and site formula."
    },

    {
        "Issue": "Crystallographic site allocation",
        "Status": "NOT CERTIFIED",
        "Reason": "Site occupancy rules and charge balance are not established.",
        "Action": "Complete mineralogical method verification."
    },

    {
        "Issue": "ML readiness",
        "Status": "NOT APPROVED",
        "Reason": "Publication provenance and methodological verification remain open.",
        "Action": "Do not admit these records as validated training labels."
    }

]


# =============================================================================
# 15. SAVE REPORT
# =============================================================================

logger.info("Writing Step 06.1 workbook...")

out = Workbook()
out.remove(out.active)

readme = [

    {
        "Item": "Step",
        "Value": "06.1 Independent iron and normalization verification"
    },

    {
        "Item": "Original workbook SHA256",
        "Value": source_hash
    },

    {
        "Item": "Original samples",
        "Value": len(original)
    },

    {
        "Item": "Independent reconstruction records",
        "Value": len(chem_records)
    },

    {
        "Item": "Revised cation source records",
        "Value": len(cation_records)
    },

    {
        "Item": "Tolerance",
        "Value": TOL
    },

    {
        "Item": "Scientific status",
        "Value": "AUDIT ONLY — NOT APPROVED FOR ML"
    },

    {
        "Item": "Important",
        "Value": (
            "No unverified published equation, Fe3 determination "
            "or crystallographic site assignment is certified."
        )
    }

]

save_sheet(out, "Read_Me", readme)
save_sheet(out, "QC_Summary", summary)
save_sheet(out, "Fe3_Reconciliation", fe3_audit)
save_sheet(out, "Optimization_Recalculation", optimization_audit)
save_sheet(out, "Corrected_Cation_Matching", revised_audit)
save_sheet(out, "Independent_Iron_Audit", iron_audit)
save_sheet(out, "Sigma13_Scaling_Audit", sigma_audit)
save_sheet(out, "HB32B_Investigation", targeted)
save_sheet(out, "Scientific_Decisions", decisions)

out.save(OUTPUT_FILE)
out.close()


# =============================================================================
# 16. EXECUTION SUMMARY
# =============================================================================

logger.info("=" * 79)
logger.info("STEP 06.1 COMPLETED")
logger.info("=" * 79)

logger.info("Samples: %d", len(original))
logger.info("Fe3 provenance records: %d", len(fe3_audit))
logger.info("Optimization records: %d", len(optimization_audit))
logger.info("Revised cation comparisons: %d", len(revised_audit))
logger.info("Iron audit records: %d", len(iron_audit))
logger.info("Sigma13 scaling records: %d", len(sigma_audit))

logger.info("Output: %s", OUTPUT_FILE)
logger.info("Log: %s", LOG_FILE)

logger.info("=" * 79)

print("\nSTEP 06.1 COMPLETED SUCCESSFULLY")
print("\nExcel report:")
print(OUTPUT_FILE)
print("\nExecution log:")
print(LOG_FILE)

print(
    "\nIMPORTANT: Successful arithmetic verification does not "
    "certify Fe3 provenance, published methodology or site allocation."
)
