# Universal Precision Automation – V2

Streamlit application for date-wise inter-run precision extraction from instrument QC PDF reports.

## Universal analyte handling
- **No hardcoded analyte list.** The parser detects assay/analyte section headings from the report structure (the section heading followed by the report's result-table column headers).
- Extracts numeric result + unit pairs dynamically. It does not depend on the names Alb, ALT, AST, or any other predefined list.
- Newly introduced analytes are included automatically when the PDF contains a text-extractable assay section and result rows.
- Every extracted row is retained in `Raw_Data_Audit`, including rows that cannot yet be assigned to Level 1 or Level 2.
- `Upload_Summary` reports the detected analyte/date/QC-level groups.
- Downloads ONE combined Excel workbook containing a separate worksheet for every detected analyte, plus shared Raw_Data_Audit and Upload_Summary worksheets.

## Output workbook
All detected analytes are included in a single Excel file. Each analyte gets its own worksheet using the uploaded lab-specific template layout. The same workbook also contains `Raw_Data_Audit` and `Upload_Summary`. No separate analyte ZIP is generated.

## Precision workflow
1. Upload the lab-specific Excel template.
2. Upload one or more instrument PDF reports.
3. The application reads actual Result Date/Time values. It does not infer dates from labels such as DAY1/DAY2.
4. Results are grouped by analyte + actual date + QC level.
5. Up to five observations closest to the group median are selected, with original report order retained for ties.
6. The original extracted observations are preserved in the audit sheet.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy using GitHub and Streamlit Community Cloud
1. Extract this ZIP on your computer.
2. Create a new GitHub repository (for example `universal-precision-automation`).
3. In the repository, choose **Add file → Upload files**.
4. Upload `app.py`, `requirements.txt`, `DUMMY_TEMPLATE.xlsx`, `README.md`, and `.gitignore` from this folder. Do not upload `__pycache__`.
5. Commit the files to the `main` branch.
6. Open https://share.streamlit.io/ and sign in using the same GitHub account.
7. Select **Create app** / **Deploy an app**.
8. Select your repository, branch `main`, and main file path `app.py`.
9. Click **Deploy**. Streamlit Cloud installs the libraries listed in `requirements.txt`.
10. Open the deployed app, upload your actual lab template and QC PDF reports, then process and download the workbook(s).

## QC-level mapping
Starter identifiers are configured in `infer_level()`:
- Level 1: 101, 301, AB1, K5972, 6304961
- Level 2: 102, 302, AB2, K5973, 6304962

Other identifiers remain `Unassigned`; review them in `Raw_Data_Audit` and extend `infer_level()` only after confirming the lab's QC identity mapping. Analyte detection is separate from QC-level mapping: an analyte is not ignored just because its QC level is unknown.

## Important validation notes
- The template provides five date rows per QC level; only the first five chronological dates can be populated in this template layout.
- A text-extractable PDF is required. Scanned/image-only reports need OCR support, and substantially different report layouts may require parser tuning.
- The generic result parser extracts number + unit pairs. Review the audit sheet for unusual units, flags, and report formats.
- The “five closest to median” rule is a selection convention, not evidence that other results are invalid. Review all raw values and your approved precision protocol before using results in a regulated validation report.
- Verify analyte names, actual dates, QC levels, units, and selected replicates against the source PDF.
