# Geothermobarometer ML Project — Available Script Archive

This repository contains the Python scripts that were recoverable from the uploaded project archive, plus the integrated Step 08.5.1–08.8 script prepared in the current conversation.

## Important completeness statement

This is **not yet a verified complete history of every script used in the project**. The uploaded `Geothermobarometer_ML_Project.zip` contained 24 `.py` entries:
- 17 contained substantive Python code and are included under `scripts/01_calibration_and_Hbl97/`.
- 7 were only 92-byte Spyder template stubs (empty apart from encoding/date/author comments); these are listed in `SCRIPT_INVENTORY.csv` but are not presented as working scripts.
- The uploaded archive did not contain the full source code for several later scripts discussed in the conversation, including Steps 07.3.1, 08.1, 08.1.2, 08.1.3, 08.2, 08.3, 08.3.1, 08.3 experimentation, 08.4, and 08.5. Their output workbooks/logs or execution traces were available for some steps, but those are not substitutes for the actual `.py` source files. Do not treat this archive as containing those missing scripts.

The integrated script in `scripts/08_ML_and_Explainability/` was generated as a combined analysis workflow. It is not a replacement for the original Step 08.2–08.5 scripts.

## Contents

- `scripts/01_calibration_and_Hbl97/`: substantive scripts recovered from the uploaded project ZIP.
- `scripts/08_ML_and_Explainability/`: integrated analysis script created in the current conversation.
- `SCRIPT_INVENTORY.csv`: exact script inventory, including empty template stubs and missing later scripts.
- `requirements.txt`: Python dependencies for the integrated analysis. Earlier scripts may require additional packages and older library versions.
- `AI_Use_Record.md`: disclosure and verification checklist.

## Data governance

Raw datasets, calibration workbooks, experimental input tables, publication PDFs, generated Excel outputs, and logs have intentionally been excluded from this script-only package. The supplied publication describes the underlying data as confidential. Obtain permission before committing raw or derived data to a public repository.

## Recommended GitHub use

1. Review every script and remove local absolute paths, personal information, and credentials.
2. Verify package requirements for each historical script.
3. Run scripts in dependency order using the corresponding input data in an authorized local environment.
4. Add missing Step 07 and Step 08 source scripts from the original working directory before describing this as the complete project repository.
5. Keep the repository private until data permissions and source completeness are confirmed.
