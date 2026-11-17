# -*- coding: utf-8 -*-
"""
Created on Tue Nov 17 11:57:51 2026

@author: hussa
"""

    limitation_rows = records(step64, "Scientific_Limitations") if step64.exists() else None

    if sample_rows is None:
        raise FileNotFoundError(
            "Step 06.4 workbook or Sample_Summary worksheet is missing. "
            "Update INPUTS['Step06_4_Structural_Formula_Verification']."
        )

    # Preserve the existing sample-level results; do not recalculate them.
    sample_integrated = []
    for r in sample_rows:
        sample_integrated.append({
            "Sample": r.get("Sample"),
            "Fe_Partition_Arithmetic": r.get("Fe_Partition_Status"),
            "T_Site_Diagnostic": r.get("T_Status"),
            "C_Site_Diagnostic": r.get("C_Status"),
            "B_Site_Diagnostic": r.get("B_Status"),
            "A_Site_Diagnostic": r.get("A_Status"),
            "Charge_Balance_Diagnostic": r.get("Charge_Status"),
            "Site_Allocation_Diagnostic": r.get("Site_Allocation_Status"),
            "Scientific_Certification": "NOT ESTABLISHED",
            "ML_Feature_Eligibility": "HOLD — do not use derived site occupancies as validated features",
            "Interpretation": (
                "Fe partition arithmetic agreement is not independent Fe-valence measurement. "
                "Site allocation and charge-balance outputs are diagnostic only."
            )
        })

    # Quantitative counts are descriptive counts from the source workbook.
    def count_value(rows, field, value):
        return sum(1 for r in (rows or []) if str(r.get(field, "")).upper() == value)

    n = len(sample_rows)
    fe_pass = count_value(sample_rows, "Fe_Partition_Status", "PASS")
    c_excess = count_value(sample_rows, "C_Status", "EXCEEDS_CAPACITY")
    charge_pass = count_value(sample_rows, "Charge_Status", "PASS")
    charge_review = count_value(sample_rows, "Charge_Status", "REVIEW")
    site_review = count_value(sample_rows, "Site_Allocation_Status", "REVIEW")

    executive = [
        {"Metric": "Audit step", "Result": "06.5 — Consolidated Hbl97 scientific audit"},
        {"Metric": "Samples represented in Step 06.4", "Result": n},
        {"Metric": "Fe partition arithmetic PASS", "Result": f"{fe_pass}/{n}"},
        {"Metric": "C-site diagnostic exceeds capacity", "Result": f"{c_excess}/{n}"},
        {"Metric": "Charge diagnostic PASS", "Result": f"{charge_pass}/{n}"},
        {"Metric": "Charge diagnostic REVIEW", "Result": f"{charge_review}/{n}"},
        {"Metric": "Site-allocation diagnostic REVIEW", "Result": f"{site_review}/{n}"},
        {"Metric": "Scientific certification", "Result": "NOT ESTABLISHED"},
        {"Metric": "ML use of derived site occupancies", "Result": "HOLD pending validated published allocation method"},
        {"Metric": "Fe3+ interpretation", "Result": "Stoichiometric estimate / workbook-derived value; not independent spectroscopy"},
        {"Metric": "Step 06.5 scope", "Result": "Consolidates existing evidence; does not recalculate or certify source methods"},
    ]

    decisions = [
        {"ID": "D01", "Topic": "Fe partition arithmetic", "Current_evidence": f"{fe_pass}/{n} arithmetic checks pass in Step 06.4.", "Decision": "Retain as arithmetic reconciliation only.", "Not_claimed": "Independent determination of Fe oxidation state."},
        {"ID": "D02", "Topic": "Fe3+ provenance", "Current_evidence": "Step 06.4 uses original workbook Fe3+ values.", "Decision": "Preserve provenance and distinguish calculated Fe3+ from independently measured Fe3+.", "Not_claimed": "Spectroscopic or independent analytical validation."},
        {"ID": "D03", "Topic": "Site occupancy", "Current_evidence": f"{c_excess}/{n} C-site diagnostic records exceed nominal capacity under the implemented allocation.", "Decision": "Treat all T-C-B-A allocations as diagnostic pending reproduction of the applicable published method.", "Not_claimed": "Crystallographic site occupancy or mineral classification."},
        {"ID": "D04", "Topic": "Charge balance", "Current_evidence": f"{charge_pass}/{n} PASS and {charge_review}/{n} REVIEW in source report.", "Decision": "Retain residuals as diagnostics; verify valence, normalization, anions and charge conventions before interpretation.", "Not_claimed": "A geological charge-balance failure based solely on this simplified calculation."},
        {"ID": "D05", "Topic": "ML feature eligibility", "Current_evidence": "Derived site occupancy depends on unresolved allocation assumptions.", "Decision": "Do not treat derived T/C/B/A values as validated predictors or labels yet.", "Not_claimed": "That the raw oxide/cation data are unusable."},
        {"ID": "D06", "Topic": "Next evidence required", "Current_evidence": "Current workbook documents methodological limitations.", "Decision": "Reconstruct amphibole formula using the exact cited method and verify source table/sample mapping.", "Not_claimed": "Automatic resolution by forcing site totals to nominal capacities."},
    ]

    source_register = []
    for item in inventory:
        source_register.append(item)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Executive_Summary"
    ws.append(["Metric", "Result"])
    for r in executive:
        ws.append([r["Metric"], r["Result"]])

    add_table(wb, "Sample_Integrated_Status", sample_integrated)
    add_table(wb, "Scientific_Decisions", decisions)
    add_table(wb, "Source_Register", source_register)

    if formula_rows: add_table(wb, "Source_Structural_Formula", formula_rows)
    if site_rows: add_table(wb, "Source_Site_Occupancy", site_rows)
    if charge_rows: add_table(wb, "Source_Charge_Balance", charge_rows)
    if iron_rows: add_table(wb, "Source_Iron_Crosscheck", iron_rows)
    if limitation_rows: add_table(wb, "Source_Limitations", limitation_rows)

    readme = [
        ["Item", "Description"],
        ["Purpose", "Evidence consolidation of Hbl97 steps; not a new thermodynamic or crystallographic calculation."],
        ["Source of sample-level diagnostic data", str(step64)],
        ["Sample count", n],
        ["Important limitation", "The Step 06.4 sequential site-allocation approach is retained as reported but is not endorsed as a validated amphibole formula method."],
        ["Fe3+ limitation", "Arithmetic reconciliation does not independently establish Fe oxidation state."],
        ["Charge limitation", "Charge residuals depend on oxidation-state and anion assumptions; source diagnostics are not independently certified."],
        ["ML safeguard", "Do not use derived site occupancies as scientifically validated inputs/targets until method verification is completed."],
        ["Missing files", "Any missing source listed in Source_Register was not integrated. No missing result was inferred."],
        ["Reproducibility", "SHA256 hashes are recorded for every source workbook found at its configured path."],
        ["Generated", datetime.now().isoformat(timespec="seconds")],
    ]
    ws = wb.create_sheet("Read_Me", 0)
    for row in readme:
        ws.append(row)

    style_workbook(wb)
    wb.save(OUTPUT_FILE)
    log.info("Step 06.5 completed. Samples: %d", n)
    log.info("Fe partition arithmetic PASS: %d/%d", fe_pass, n)
    log.info("Scientific certification: NOT ESTABLISHED")
    log.info("Report: %s", OUTPUT_FILE)
    log.info("=" * 85)

if __name__ == "__main__":
    main()
