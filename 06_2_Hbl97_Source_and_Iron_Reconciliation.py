# -*- coding: utf-8 -*-
"""
Created on Mon Nov 16 17:31:36 2026

@author: hussa
"""


# =============================================================================
# STEP 06.2 — HBL97 SOURCE AND IRON RECONCILIATION
# =============================================================================
# Purpose:
# Integrate the eight non-iron cation records and the separate Fe2+/Fe3+
# audit into one 130-record reconciliation.
#
# IMPORTANT:
# This is an audit of existing calculations and source provenance.
# It does not certify Fe3 determination, site allocation or calibration.
# =============================================================================

from pathlib import Path
from collections import Counter
import logging
import hashlib
import math

import xlrd
import openpyxl

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


# =============================================================================
# 1. PATHS
# =============================================================================

ROOT = Path(__file__).resolve().parent

SOURCE_XLS = (
    ROOT / "Geothermobarometers Calibrations" /
    "Hbl P-T (1997).xls"
)

CHEM_FILE = (
    ROOT / "05C4_1_Hbl97_original_chemical_reconstruction" /
    "Step05C4_1_Hbl97_Original_Chemical_Reconstruction.xlsx"
)

AUDIT_FILE = (
    ROOT / "05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit" /
    "Step05C4_2_Hbl97_Fe3_Optimization_Normalization_Audit.xlsx"
)

STEP06_FILE = (
    ROOT / "06_Hbl97_Fe3_Provenance_and_SiteTerm_Audit" /
    "Step06_Hbl97_Fe3_Provenance_and_SiteTerm_Audit.xlsx"
)

STEP061_FILE = (
    ROOT / "06_1_Hbl97_Independent_Iron_Normalization_Verification" /
    "Step06_1_Hbl97_Independent_Verification.xlsx"
)

OUT_DIR = (
    ROOT / "06_2_Hbl97_Source_and_Iron_Reconciliation"
)

OUT_DIR.mkdir(exist_ok=True)

LOG_DIR = OUT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

OUTPUT = (
    OUT_DIR /
    "Step06_2_Hbl97_Source_and_Iron_Reconciliation.xlsx"
)

LOG_FILE = LOG_DIR / "step06_2_execution_log.txt"

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

logger = logging.getLogger("STEP06_2")


# =============================================================================
# 3. UTILITIES
# =============================================================================

def val(x):

    if x is None:
        return None

    try:
        x = float(x)

        if math.isfinite(x):
            return x

    except (ValueError, TypeError):
        pass

    return None


def check(a, b):

    a = val(a)
    b = val(b)

    if a is None or b is None:
        return "NOT EVALUATED"

    return (
        "PASS"
        if abs(a - b) <= TOL
        else "REVIEW"
    )


def element_name(x):

    if x is None:
        return ""

    s = str(x).strip().lower()

    s = s.replace("⁺", "+")
    s = s.replace("²", "2")
    s = s.replace("³", "3")
    s = s.replace(" ", "")
    s = s.replace("_", "")

    mapping = {

        "si": "Si",
        "ti": "Ti",
        "al": "Al",
        "mn": "Mn",
        "mg": "Mg",
        "ca": "Ca",
        "na": "Na",
        "k": "K",

        "fe2": "Fe2",
        "fe2+": "Fe2",
        "fe3": "Fe3",
        "fe3+": "Fe3"

    }

    return mapping.get(s, str(x).strip())


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
            f"{sheet_name} not found in {path.name}"
        )

    ws = wb[sheet_name]

    rows = ws.iter_rows(values_only=True)

    headers = next(rows)

    headers = [
        str(x).strip() if x is not None else ""
        for x in headers
    ]

    records = []

    for row in rows:

        rec = {
            headers[i]: row[i]
            for i in range(min(len(headers), len(row)))
            if headers[i]
        }

        if any(x is not None for x in rec.values()):
            records.append(rec)

    wb.close()

    return records


def index_sample(records):

    result = {}

    for r in records:

        sample = r.get("Sample")

        if sample is not None:
            result[str(sample).strip()] = r

    return result


def digest(path):

    h = hashlib.sha256()

    with open(path, "rb") as f:

        for block in iter(
            lambda: f.read(1024 * 1024),
            b""
        ):
            h.update(block)

    return h.hexdigest()


def save_sheet(wb, name, records):

    ws = wb.create_sheet(name)

    if not records:
        ws.append(["No records"])
        return

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
            wrap_text=True
        )

    for column in ws.columns:

        letter = get_column_letter(
            column[0].column
        )

        width = max(
            (
                len(str(c.value))
                for c in column
                if c.value is not None
            ),
            default=12
        )

        ws.column_dimensions[letter].width = min(
            max(width + 2, 12),
            42
        )


# =============================================================================
# 4. VALIDATE FILES
# =============================================================================

logger.info("=" * 78)
logger.info("STEP 06.2 — SOURCE AND IRON RECONCILIATION")
logger.info("=" * 78)

REQUIRED = [
    SOURCE_XLS,
    CHEM_FILE,
    AUDIT_FILE,
    STEP06_FILE,
    STEP061_FILE
]

for path in REQUIRED:

    if not path.exists():
        raise FileNotFoundError(path)

    logger.info("Verified: %s", path.name)

source_hash = digest(SOURCE_XLS)


# =============================================================================
# 5. LOAD AUDIT RECORDS
# =============================================================================

logger.info("Loading existing audit records...")

noniron = read_sheet(
    AUDIT_FILE,
    "Revised_Cation_Audit"
)

iron = read_sheet(
    AUDIT_FILE,
    "Iron_Partition_Audit"
)

optimization = read_sheet(
    AUDIT_FILE,
    "Optimization_Audit"
)

sigma = read_sheet(
    AUDIT_FILE,
    "Sigma13_Normalization"
)

chemistry = read_sheet(
    CHEM_FILE,
    "Independent_Reconstruction"
)

iron_step061 = read_sheet(
    STEP061_FILE,
    "Independent_Iron_Audit"
)

fe3_step061 = read_sheet(
    STEP061_FILE,
    "Fe3_Reconciliation"
)

chem_idx = index_sample(chemistry)
iron_idx = index_sample(iron)
optimization_idx = index_sample(optimization)

iron061_idx = index_sample(iron_step061)
fe3_idx = index_sample(fe3_step061)


# =============================================================================
# 6. READ ORIGINAL XLS
# =============================================================================

logger.info("Reading original workbook...")

xls = xlrd.open_workbook(
    str(SOURCE_XLS),
    on_demand=True
)

ws = xls.sheet_by_name("Hbl P-T (1997)")

COL = {

    "Sample": 1,

    "Fe3_H": 7,

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

    "fav": 42,

    "Fe3_AR": 43,

    "Ratio_13": 58,

    "Si_13": 59,
    "Ti_13": 60,
    "Al_13": 61,
    "Fe2_13": 62,
    "Fe3_13": 63,
    "Mn_13": 64,
    "Mg_13": 65

}

original = []

for r in range(9, ws.nrows):

    sample = ws.cell_value(r, 1)

    if not sample:
        continue

    rec = {
        "Sample": str(sample).strip(),
        "Excel_Row": r + 1
    }

    for name, col in COL.items():

        if name == "Sample":
            continue

        rec[name] = val(
            ws.cell_value(r, col)
        )

    original.append(rec)

xls.release_resources()

original_idx = index_sample(original)

logger.info("Original samples: %d", len(original))


# =============================================================================
# 7. INTEGRATE ALL TEN CATIONS
# =============================================================================

logger.info("Integrating non-iron and iron records...")

noniron_idx = {}

for record in noniron:

    key = (
        str(record.get("Sample", "")).strip(),
        element_name(record.get("Element"))
    )

    noniron_idx[key] = record


FIELDS = {

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

integrated = []

for sample, orig in original_idx.items():

    iron_rec = iron_idx.get(sample, {})

    for element, field in FIELDS.items():

        if element in ["Fe2", "Fe3"]:

            if element == "Fe2":

                previous_value = val(
                    iron_rec.get("Original_Revised_Fe2")
                )

                predicted = val(
                    iron_rec.get("Predicted_Fe2")
                )

            else:

                previous_value = val(
                    iron_rec.get("Original_Revised_Fe3")
                )

                predicted = val(
                    iron_rec.get("Original_H_Fe3")
                )

            source_sheet = "Iron_Partition_Audit"

        else:

            source = noniron_idx.get(
                (sample, element),
                {}
            )

            previous_value = val(
                source.get("Original_Revised")
            )

            predicted = val(
                source.get("Predicted_Revised")
            )

            source_sheet = "Revised_Cation_Audit"

        original_value = orig.get(field)

        integrated.append({

            "Sample": sample,

            "Element": element,

            "Original_Workbook_Value": original_value,

            "Previous_Audit_Value": previous_value,

            "Predicted_Value": predicted,

            "Original_vs_Audit": check(
                original_value,
                previous_value
            ),

            "Prediction_vs_Original": check(
                original_value,
                predicted
            ),

            "Source_Worksheet": source_sheet,

            "Source_Record_Status": (
                "FOUND" if previous_value is not None
                else "NOT FOUND"
            )

        })


# =============================================================================
# 8. FE3 PROVENANCE AND HB32B
# =============================================================================

logger.info("Reconciling Fe3 provenance...")

fe3_reconciliation = []

for sample, orig in original_idx.items():

    previous = iron_idx.get(sample, {})
    step061 = fe3_idx.get(sample, {})

    H = orig.get("Fe3_H")
    AR = orig.get("Fe3_AR")

    revised = orig.get("Fe3_revised")

    fe3_reconciliation.append({

        "Sample": sample,

        "Excel_Row": orig.get("Excel_Row"),

        "Original_H": H,

        "Original_AR": AR,

        "Original_Revised_AX": revised,

        "Iron_Audit_H": val(
            previous.get("Original_H_Fe3")
        ),

        "Iron_Audit_AX": val(
            previous.get("Original_Revised_Fe3")
        ),

        "Step061_H": val(
            step061.get("Original_H")
        ),

        "H_vs_AX": check(H, revised),

        "H_vs_AR": check(H, AR),

        "HB32B_Flag": (
            "SPECIAL REVIEW"
            if sample == "HB32B(M2)"
            else "STANDARD REVIEW"
        ),

        "Scientific_Status": (
            "PRESERVE ORIGINAL VALUES; analytical provenance unresolved."
        )

    })


# =============================================================================
# 9. OPTIMIZATION AND NORMALIZATION RECONCILIATION
# =============================================================================

optimization_reconciliation = []

for sample, record in optimization_idx.items():

    optimization_reconciliation.append({

        "Sample": sample,

        "Minimum_Check": record.get("Minimum_Check"),

        "Maximum_Check": record.get("Maximum_Check"),

        "fav_Check": record.get("fav_Check"),

        "fav_Original": record.get("fav_Original"),

        "fav_Reconstructed": record.get("fav_Reconstructed"),

        "Interpretation": (
            "Arithmetic agreement only; published factor definitions "
            "require independent documentary verification."
        )

    })


sigma_reconciliation = []

for record in sigma:

    sigma_reconciliation.append({

        "Sample": record.get("Sample"),

        "Revised_Sigma13_Cation_Sum": record.get(
            "Revised_Sigma13_Cation_Sum"
        ),

        "Predicted_MN13_Factor": record.get(
            "Predicted_MN13_Factor"
        ),

        "Predicted_Sigma13": record.get(
            "Predicted_Sigma13"
        ),

        "Sigma13_Check": record.get(
            "Sigma13_Check"
        ),

        "Scientific_Status": (
            "Arithmetic normalization agreement; mineralogical "
            "basis remains to be verified."
        )

    })


# =============================================================================
# 10. SAMPLE PUBLICATION CROSSWALK
# =============================================================================

logger.info("Creating publication crosswalk...")

publication_records = []

for sample in original_idx:

    publication_records.append({

        "Sample": sample,

        "Candidate_Publication": (
            "Xiao et al. (2011), International Journal of Earth Sciences"
        ),

        "DOI": "10.1007/s00531-010-0522-5",

        "Geological_Context": (
            "Zanhuang Complex, Trans-North China Orogen"
        ),

        "Workbook_Assemblage_Label": "M2",

        "Publication_M2_Interpretation": (
            "Peak metamorphic matrix assemblage"
        ),

        "Publication_M3_Interpretation": (
            "Amphibole + plagioclase symplectite around garnet"
        ),

        "Exact_Sample_Match": "NOT YET VERIFIED",

        "Chemical_Analysis_Match": "NOT YET VERIFIED",

        "Publication_Status": (
            "CANDIDATE REFERENCE — sample-level confirmation required"
        )

    })


# =============================================================================
# 11. QC SUMMARY
# =============================================================================

status_counts = Counter(
    r["Prediction_vs_Original"]
    for r in integrated
)

source_counts = Counter(
    r["Source_Record_Status"]
    for r in integrated
)

summary = [

    {
        "Metric": "Original samples",
        "Value": len(original)
    },

    {
        "Metric": "Non-iron source records",
        "Value": len(noniron)
    },

    {
        "Metric": "Iron source records",
        "Value": len(iron)
    },

    {
        "Metric": "Integrated cation records",
        "Value": len(integrated)
    },

    {
        "Metric": "Expected cation records",
        "Value": len(original) * 10
    },

    {
        "Metric": "Prediction comparisons PASS",
        "Value": status_counts.get("PASS", 0)
    },

    {
        "Metric": "Prediction comparisons REVIEW",
        "Value": status_counts.get("REVIEW", 0)
    },

    {
        "Metric": "Prediction comparisons NOT EVALUATED",
        "Value": status_counts.get("NOT EVALUATED", 0)
    },

    {
        "Metric": "Source records FOUND",
        "Value": source_counts.get("FOUND", 0)
    },

    {
        "Metric": "Source records NOT FOUND",
        "Value": source_counts.get("NOT FOUND", 0)
    },

    {
        "Metric": "Original workbook SHA256",
        "Value": source_hash
    }

]


# =============================================================================
# 12. SCIENTIFIC DECISION REGISTER
# =============================================================================

decisions = [

    {
        "Issue": "Xiao et al. publication identity",
        "Status": "IDENTIFIED",
        "Finding": (
            "Xiao et al. (2011), DOI 10.1007/s00531-010-0522-5."
        ),
        "Remaining_Work": (
            "Verify exact sample identifiers and oxide analyses."
        )
    },

    {
        "Issue": "M2 versus M3 interpretation",
        "Status": "REQUIRES REVIEW",
        "Finding": (
            "Publication describes M2 as peak matrix assemblage and "
            "M3 as amphibole-plagioclase symplectite."
        ),
        "Remaining_Work": (
            "Determine whether the workbook's M2 label refers to "
            "the same sample generation and mineral domain."
        )
    },

    {
        "Issue": "Fe2 and Fe3 records",
        "Status": "INTEGRATED",
        "Finding": (
            "Iron records are sourced from Iron_Partition_Audit, "
            "not Revised_Cation_Audit."
        ),
        "Remaining_Work": (
            "Verify independent derivation and Fe3 provenance."
        )
    },

    {
        "Issue": "HB32B H versus AR",
        "Status": "UNRESOLVED",
        "Finding": (
            "H = 0.000000; AR = -0.003466."
        ),
        "Remaining_Work": (
            "Trace original workbook formulas and scientific instructions."
        )
    },

    {
        "Issue": "Optimization factors",
        "Status": "ARITHMETIC RECONCILIATION",
        "Finding": (
            "Existing minimum, maximum and fav checks are retained."
        ),
        "Remaining_Work": (
            "Verify equations and physical meaning against the source."
        )
    },

    {
        "Issue": "Sigma13 normalization",
        "Status": "ARITHMETIC RECONCILIATION",
        "Finding": (
            "Existing normalization comparisons are retained."
        ),
        "Remaining_Work": (
            "Verify normalization basis and cation selection."
        )
    },

    {
        "Issue": "Crystallographic site allocation",
        "Status": "NOT CERTIFIED",
        "Finding": (
            "Arithmetic reconstruction is insufficient."
        ),
        "Remaining_Work": (
            "Complete structural formula and charge-balance verification."
        )
    },

    {
        "Issue": "ML training eligibility",
        "Status": "NOT APPROVED",
        "Finding": (
            "Scientific provenance and methodological validation remain open."
        ),
        "Remaining_Work": (
            "Complete scientific acceptance gates."
        )
    }

]


# =============================================================================
# 13. WRITE EXCEL REPORT
# =============================================================================

logger.info("Writing consolidated report...")

out = Workbook()
out.remove(out.active)

readme = [

    {
        "Item": "Step",
        "Value": "06.2 Source and Iron Reconciliation"
    },

    {
        "Item": "Source workbook SHA256",
        "Value": source_hash
    },

    {
        "Item": "Purpose",
        "Value": (
            "Integrate eight non-iron cations and two iron cations "
            "from their respective source worksheets."
        )
    },

    {
        "Item": "Publication",
        "Value": (
            "Xiao et al. (2011), DOI 10.1007/s00531-010-0522-5"
        )
    },

    {
        "Item": "Publication matching",
        "Value": (
            "Bibliographic and geological candidate identified; "
            "exact sample matching not yet certified."
        )
    },

    {
        "Item": "Scientific status",
        "Value": "AUDIT ONLY — NOT APPROVED FOR ML"
    }

]

save_sheet(out, "Read_Me", readme)
save_sheet(out, "QC_Summary", summary)
save_sheet(out, "Integrated_130_Cations", integrated)
save_sheet(out, "Fe3_Provenance", fe3_reconciliation)
save_sheet(out, "Optimization_Reconciliation", optimization_reconciliation)
save_sheet(out, "Sigma13_Reconciliation", sigma_reconciliation)
save_sheet(out, "Publication_Crosswalk", publication_records)
save_sheet(out, "Scientific_Decisions", decisions)

out.save(OUTPUT)
out.close()


# =============================================================================
# 14. FINAL LOG
# =============================================================================

logger.info("=" * 78)
logger.info("STEP 06.2 COMPLETED")
logger.info("=" * 78)

logger.info("Samples: %d", len(original))
logger.info("Integrated cations: %d", len(integrated))

logger.info(
    "PASS comparisons: %d",
    status_counts.get("PASS", 0)
)

logger.info(
    "REVIEW comparisons: %d",
    status_counts.get("REVIEW", 0)
)

logger.info(
    "NOT EVALUATED comparisons: %d",
    status_counts.get("NOT EVALUATED", 0)
)

logger.info("Excel report: %s", OUTPUT)
logger.info("Execution log: %s", LOG_FILE)

logger.info("=" * 78)

print("\nSTEP 06.2 COMPLETED SUCCESSFULLY")
print("\nIntegrated cation records:", len(integrated))
print("\nExcel report:")
print(OUTPUT)
print("\nExecution log:")
print(LOG_FILE)

print(
    "\nIMPORTANT: Publication identity does not establish "
    "sample-level matching or scientific calibration validity."
)
