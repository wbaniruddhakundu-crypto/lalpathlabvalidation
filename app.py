import io, re, copy, hashlib, zipfile
from datetime import datetime, date
import fitz
import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Universal Inter-Run Precision", page_icon="🧪", layout="wide")
st.title("🧪 Universal 5-Day Inter-Run Precision")
st.caption("Dynamically detects assay headings and result rows from the uploaded PDF; no predefined analyte list is used. Results are grouped by actual report date, not SID day labels.")

# ------------------------- parsing helpers -------------------------
NUM_UNIT = re.compile(r"^\s*([-+]?\d[\d,]*(?:\.\d+)?)\s*([A-Za-zµμ/%]+(?:\s*/\s*[A-Za-z]+)?|Index)?\s*$")
NUM_ANY = re.compile(r"([-+]?\d[\d,]*(?:\.\d+)?)\s*([A-Za-zµμ/%]+(?:\s*/\s*[A-Za-z]+)?|Index)?")
SKIP_HEADINGS = {
    "ASSAY REPORT", "PATIENT NAME", "SID", "RESULT", "RLU OR MAU", "FLAGS",
    "RESULT DATE/TIME", "FINAL RESULT", "LAB EVALUATIONS", "CALIBRATOR NAME OR",
    "QC MATERIAL NAME", "CALIBRATOR ID OR QC", "MATERIAL ID", "ORDER TIME",
    "SOFTWARE VERSION", "SAMPLE TYPE", "ANALYZER", "PRINTED"
}
KNOWN_NON_ANALYTE = re.compile(
    r"^(?:PAGE\s+\d+|PRINTED:|IRP\d+|SOFTWARE VERSION|ORDER TIME|SAMPLE TYPE|ANALYZER)\b", re.I
)

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def is_analyte_heading(line):
    """Recognize assay section headings from report structure, not a parameter whitelist."""
    x = clean(line)
    if not x or x.upper().rstrip(":") in SKIP_HEADINGS or KNOWN_NON_ANALYTE.match(x):
        return False
    if len(x) > 45 or re.search(r"\d{1,2}/\d{1,2}/\d{4}", x):
        return False
    if re.search(r"\b(?:PATIENT NAME|CALIBRATOR NAME|QC MATERIAL|RESULT DATE|RLU OR MAU)\b", x, re.I):
        return False
    # Instrument reports place the column headings directly below each analyte heading.
    return bool(re.fullmatch(r"[A-Za-z][A-Za-z0-9_.+ /()\-]{0,44}", x))

def extract_pdf_text(pdf_bytes):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    return "\n".join(page.get_text("text", sort=True) for page in doc)

def date_from_context(context_lines, fallback_date):
    """Return the actual result date nearest to a result row, as YYYY-MM-DD."""
    date_re = re.compile(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b")
    candidates = []
    for offset, line in enumerate(context_lines):
        m = date_re.search(line)
        if m:
            mm, dd, yyyy = map(int, m.groups())
            try:
                candidates.append((abs(offset - (len(context_lines)//2)), date(yyyy, mm, dd).isoformat()))
            except ValueError:
                pass
    if candidates:
        return sorted(candidates)[0][1]
    return fallback_date.isoformat() if hasattr(fallback_date, "isoformat") else str(fallback_date)

def infer_level(sid, qc_name="", qc_id=""):
    text = f"{sid} {qc_name} {qc_id}".upper()
    # More specific tokens first. These are defaults and can be overridden in the UI.
    if re.search(r"\b(?:102|302|DAY\s*[1-5]0?2|LINE[-_ ]?102|AB[-_ ]?2|LEVEL[-_ ]?2|L[-_ ]?2|K5973|6304962)\b", text):
        return "Level 2"
    if re.search(r"\b(?:101|301|DAY\s*[1-5]0?1|LINE[-_ ]?101|AB[-_ ]?1|LEVEL[-_ ]?1|L[-_ ]?1|K5972|6304961)\b", text):
        return "Level 1"
    # Common QC labels such as CHIVAB1-362 / CHIVAB2-363
    if re.search(r"AB[12]", text):
        return "Level 2" if re.search(r"AB2", text) else "Level 1"
    return "Unassigned"


def parse_pdf(pdf_bytes, filename, fallback_date):
    """Parse instrument report rows where SID, result and RLU may share or span lines."""
    text = extract_pdf_text(pdf_bytes)
    lines = [clean(x) for x in text.splitlines()]
    records, current_analyte = [], None
    # Generic value + unit matcher: no analyte names or fixed unit list.
    # Captures the assay result (first number followed by a unit), not the later bare RLU/MAU.
    unit_token = r"(?:Index|Ratio|[A-Za-zµμ%]+(?:\s*/\s*[A-Za-zµμ%]+)?)"
    result_pattern = re.compile(
        r"(?<![\w.])([-+]?\d[\d,]*(?:\.\d+)?)\s*(" + unit_token + r")(?![A-Za-zµμ])",
        re.I
    )
    sid_pattern = re.compile(r"\b(?:INTER\s*DAY\s*[1-9]\s*[A-Z0-9+_-]*|INTER\s*\d+\s+LINE|"
                             r"INTER\s*DAY\s*\d+|INTER\d+\s+LINE|DAY\s*\d{3}|"
                             r"LINE[-_ ]?\d+|DAY\s*[1-9]\d*)\b", re.I)
    for i, line in enumerate(lines):
        lookahead = " ".join(lines[i+1:i+7]).upper()
        if is_analyte_heading(line) and (
            "PATIENT NAME" in lookahead or "CALIBRATOR NAME" in lookahead or "QC MATERIAL NAME" in lookahead
        ):
            current_analyte = line
            continue
        if not current_analyte:
            continue

        # A result is the first numeric value followed by a measurement unit.
        # RLU/MAU is a later bare integer and is intentionally not matched.
        non_measurement_tokens = {
            "AM", "PM", "RLU", "MAU", "AB", "AG", "N", "HOURS", "HOUR",
            "OF", "RULE", "WITHIN", "CHECK", "LAB", "EVALUATIONS",
            "FINAL", "RESULT", "FLAGS"
        }
        matches = [m for m in result_pattern.finditer(line)
                   if m.group(2).strip().upper() not in non_measurement_tokens]
        if not matches:
            continue
        # Prefer explicit measurement units (Index, %, or compound units such as ng/mL).
        # This avoids mistaking a trailing flag word (e.g. "Rule") for the result unit.
        preferred = [m for m in matches if (
            m.group(2).strip().lower() in {"index", "ratio", "%"} or "/" in m.group(2)
        )]
        m = preferred[0] if preferred else matches[0]
        value = float(m.group(1).replace(",", ""))
        unit = clean(m.group(2)).replace(" ", "")

        context = lines[max(0, i-3):min(len(lines), i+4)]
        context_text = " ".join(context)
        sid = ""
        for c in context:
            if sid_pattern.search(c):
                # Prefer the SID carrying the inter-day/line label, not the date/time.
                sid = c
                break

        # QC-material reports may not have a SID column; use material name and QC ID.
        qc_id = ""
        for c in context:
            q = re.search(r"\b(?:K\d{4}|630496\d|\d{6,10})\b", c, re.I)
            if q:
                qc_id = q.group(0)
                break
        qc_name = ""
        for c in context:
            if re.search(r"[A-Za-z]{2,}", c) and not re.search(
                r"INTER|DAY|LINE|FINAL RESULT|LAB EVALUATIONS|RULE|CHECK|PRINTED|ASSAY REPORT", c, re.I
            ):
                # Avoid using the analyte heading or generic column labels.
                if c.upper() != current_analyte.upper() and not c.upper().startswith(
                    ("PATIENT NAME", "CALIBRATOR NAME", "QC MATERIAL NAME", "RESULT DATE")
                ):
                    qc_name = c
                    break

        # For patient/SID reports, use the SID; for QC-material reports use material ID/name.
        identity = sid or qc_name or qc_id
        if not identity:
            continue

        result_date = date_from_context(context, fallback_date)
        level = infer_level(" ".join(context), qc_name, qc_id)

        records.append({
            "Analyte": current_analyte, "Date": result_date, "Level": level,
            "Result": value, "Unit": unit, "SID / QC": sid or identity,
            "QC Material": qc_name, "QC ID": qc_id,
            "Source PDF": filename, "Page line": i+1
        })
    return records

def best_five(values):
    vals = [float(v) for v in values if pd.notna(v)]
    if len(vals) <= 5:
        return vals
    med = pd.Series(vals).median()
    # Select closest to median; stable sort preserves report order for ties.
    indexed = list(enumerate(vals))
    chosen = sorted(indexed, key=lambda z: (abs(z[1]-med), z[0]))[:5]
    chosen.sort(key=lambda z: z[0])
    return [v for _, v in chosen]

def selected_record_keys(records):
    """Return exact source-row keys for the five observations closest to each group median."""
    groups = {}
    for r in records:
        key = (r["Analyte"], r.get("Date", ""), r["Level"])
        groups.setdefault(key, []).append(r)
    selected = set()
    for group in groups.values():
        if len(group) <= 5:
            chosen = group
        else:
            med = pd.Series([float(x["Result"]) for x in group]).median()
            ranked = sorted(enumerate(group), key=lambda z: (abs(float(z[1]["Result"]) - med), z[0]))[:5]
            chosen = [x for _, x in ranked]
        for x in chosen:
            selected.add((x["Analyte"], x.get("Date", ""), x["Level"], x["Source PDF"], str(x["Page line"])))
    return selected


def safe_sheet_name(name, used):
    base = re.sub(r'[\[\]:*?/\\\\]', '_', str(name)).strip()[:31] or "Analyte"
    candidate, n = base, 1
    while candidate.lower() in used:
        suffix = f"_{n}"
        candidate = base[:31-len(suffix)] + suffix
        n += 1
    used.add(candidate.lower())
    return candidate

def make_analyte_sheet(wb, template_ws, analyte, used):
    ws = wb.copy_worksheet(template_ws)
    ws.title = safe_sheet_name(analyte, used)
    ws["C7"] = analyte
    # Keep formulas/layout but clear data cells from the copied template.
    for r in list(range(15, 20)) + list(range(32, 37)):
        ws.cell(r, 2).value = None
        for c in range(3, 8):
            ws.cell(r, c).value = None
    return ws

def update_workbook(template_bytes, records, previous_bytes=None):
    source = previous_bytes if previous_bytes else template_bytes
    wb = load_workbook(io.BytesIO(source))
    # Use the first worksheet as the layout template. New analytes are cloned from it.
    template_ws = wb[wb.sheetnames[0]]
    used = {s.title.lower() for s in wb.worksheets}
    analyte_to_ws = {}
    for ws in wb.worksheets:
        label = ws["C7"].value or ws["B7"].value or ws.title
        analyte_to_ws[str(label).strip().lower()] = ws
    # Keep only recognized data; append raw audit tab.
    for rec in records:
        key = rec["Analyte"].strip().lower()
        if key not in analyte_to_ws:
            ws = make_analyte_sheet(wb, template_ws, rec["Analyte"], used)
            analyte_to_ws[key] = ws
    for analyte, ws in analyte_to_ws.items():
        display_name = str(ws["C7"].value or ws["B7"].value or ws.title).strip()
        subset = [r for r in records if r["Analyte"].strip().lower() == analyte]
        # Populate one row per actual calendar date, independently for each QC level.
        # The workbook template provides five date rows for a five-day precision study.
        for level, start_row in [("Level 1", 15), ("Level 2", 32)]:
            dates = sorted({r["Date"] for r in subset if r["Level"] == level and r.get("Date")})
            for date_index, result_date in enumerate(dates[:5]):
                group = [r for r in subset if r["Level"] == level and r["Date"] == result_date]
                vals = best_five([r["Result"] for r in group])
                row = start_row + date_index
                ws.cell(row, 2).value = datetime.strptime(result_date, "%Y-%m-%d").date()
                ws.cell(row, 2).number_format = "dd-mmm-yyyy"
                for j in range(5):
                    ws.cell(row, 3+j).value = vals[j] if j < len(vals) else None
                ws.cell(row, 8).value = f"=IFERROR(AVERAGE(C{row}:G{row}),0)"
                ws.cell(row, 9).value = f"=IFERROR(STDEV(C{row}:G{row}),0)"
                ws.cell(row,10).value = f"=IFERROR((I{row}/H{row})*100,0)"
            if len(dates) > 5:
                st.warning(f"{display_name} / {level}: {len(dates)} distinct dates found; the template has space for five. Only the first five chronological dates are populated.")
            end_row = start_row + 4
            mean_row, sd_row, cv_row = (22,23,24) if level == "Level 1" else (39,40,41)
            ws.cell(mean_row,10).value = f"=IFERROR(AVERAGE(C{start_row}:G{end_row}),0)"
            ws.cell(sd_row,10).value = f"=IFERROR(STDEV(C{start_row}:G{end_row}),0)"
            ws.cell(cv_row,10).value = f"=IFERROR((J{sd_row}/J{mean_row})*100,0)"
        units = [r["Unit"] for r in subset if r.get("Unit")]
        if units:
            ws["C11"] = units[0]
            ws["C28"] = units[0]
    # Raw audit: preserve the extracted rows used for selection.
    audit_name = "Raw_Data_Audit"
    if audit_name in wb.sheetnames:
        del wb[audit_name]
    audit = wb.create_sheet(audit_name)
    headers = ["Analyte","Date","Level","Result","Unit","SID / QC","QC Material","QC ID","Source PDF","Page line","Selected for workbook"]
    audit.append(headers)
    selected_keys = selected_record_keys(records)
    for r in records:
        row_key = (r["Analyte"], r.get("Date", ""), r["Level"], r["Source PDF"], str(r["Page line"]))
        selected = row_key in selected_keys
        audit.append([r["Analyte"],r["Date"],r["Level"],r["Result"],r["Unit"],r["SID / QC"],r["QC Material"],r["QC ID"],r["Source PDF"],r["Page line"],"Yes" if selected else "No"])
    audit.freeze_panes = "A2"
    audit.auto_filter.ref = audit.dimensions
    for cell in audit[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for col in audit.columns:
        letter = get_column_letter(col[0].column)
        audit.column_dimensions[letter].width = min(max(max(len(str(c.value or "")) for c in col)+2, 12), 32)
    # Selection log
    if "Upload_Summary" in wb.sheetnames:
        del wb["Upload_Summary"]
    summary = wb.create_sheet("Upload_Summary")
    summary.append(["Analyte","Date","Level","Raw results found","Replicates populated","Selection method"])
    combos = sorted({(r["Analyte"],r["Date"],r["Level"]) for r in records})
    for a,d,l in combos:
        g=[r for r in records if (r["Analyte"],r["Date"],r["Level"])==(a,d,l)]
        summary.append([a,d,l,len(g),min(5,len(g)),"5 closest to median" if len(g)>5 else "All available (≤5)"])
    for c in summary[1]:
        c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F4E78")
    summary.freeze_panes="A2"; summary.auto_filter.ref=summary.dimensions
    out=io.BytesIO(); wb.save(out); return out.getvalue()


# ------------------------- UI -------------------------
st.info("Workflow: first upload your original template. On later sessions, upload the latest populated workbook so earlier days are preserved.")
c1,c2=st.columns(2)
with c1:
    template_file=st.file_uploader("1. Excel template (first run)", type=["xlsx"], key="template")
with c2:
    previous_file=st.file_uploader("Or latest populated workbook (to append new days)", type=["xlsx"], key="previous")
pdf_files=st.file_uploader("2. Upload one or more instrument PDF reports", type=["pdf"], accept_multiple_files=True)

if pdf_files:
    st.subheader("Date fallback (only used when the PDF has no result date)")
    date_map = {}
    cols = st.columns(min(3, len(pdf_files)))
    for idx, pdf in enumerate(pdf_files):
        with cols[idx % len(cols)]:
            date_map[pdf.name] = st.date_input(
                f"{pdf.name}: fallback date", value=date.today(), key=f"date_{idx}"
            )
    st.caption("The app reads the actual Result Date/Time from each PDF. This date selector is only a fallback if a result row has no date.")

st.subheader("Selection and QC-level rules")
st.write("The app groups results by analyte + actual calendar date + QC level, then selects the five results closest to that group median. All extracted results remain in the Raw_Data_Audit sheet.")
with st.expander("Default Level 1 / Level 2 identifier rules"):
    st.code("Level 1: 101, 301, AB1, K5972, 6304961\\nLevel 2: 102, 302, AB2, K5973, 6304962")
    st.warning("These are starter QC-level identifier rules, not an analyte list. Any assay heading found in the report is eligible for extraction. Unknown QC identifiers remain Unassigned for review; add verified identifier patterns to infer_level() if your lab uses a different convention.")

if st.button("🚀 Process PDFs and generate updated Excel", type="primary", disabled=not pdf_files or not (template_file or previous_file)):
    try:
        all_records=[]
        for pdf in pdf_files:
            fallback_date=date_map[pdf.name]
            recs=parse_pdf(pdf.getvalue(),pdf.name,fallback_date)
            all_records.extend(recs)
        if not all_records:
            st.error("No supported result rows were recognized. Check that the PDF contains selectable report text and review the uploaded report layout. The parser is analyte-independent, but materially different PDF layouts or scanned/image-only PDFs may require OCR or a layout-specific parser.")
        else:
            # Preserve prior raw data when updating an existing workbook.
            prior_records=[]
            if previous_file:
                # Existing audit sheet is re-read so earlier data remain available.
                old=load_workbook(io.BytesIO(previous_file.getvalue()),data_only=True)
                if "Raw_Data_Audit" in old.sheetnames:
                    df=pd.DataFrame(list(old["Raw_Data_Audit"].values))
                    if len(df)>1:
                        df.columns=df.iloc[0]
                        for _,row in df.iloc[1:].iterrows():
                            prior_records.append({
                                "Analyte":str(row.get("Analyte","")), "Date":str(row.get("Date", ""))[:10],
                                "Level":str(row.get("Level","Unassigned")), "Result":float(row.get("Result")),
                                "Unit":str(row.get("Unit","")), "SID / QC":str(row.get("SID / QC","")),
                                "QC Material":str(row.get("QC Material","")), "QC ID":str(row.get("QC ID","")),
                                "Source PDF":str(row.get("Source PDF","")), "Page line":row.get("Page line","")
                            })
            # De-duplicate identical records across repeated uploads.
            combined=prior_records+all_records
            seen=set(); unique=[]
            for r in combined:
                key=(r["Analyte"],r["Date"],r["Level"],r["Result"],r["SID / QC"],r["Source PDF"],str(r["Page line"]))
                if key not in seen:
                    seen.add(key); unique.append(r)
            base_bytes=template_file.getvalue() if template_file else previous_file.getvalue()
            prev_bytes=previous_file.getvalue() if previous_file else None
            result=update_workbook(base_bytes,unique,prev_bytes)
            st.success(f"Extracted {len(all_records)} new result rows across {len(set(r['Analyte'] for r in all_records))} analytes.")
            st.download_button("⬇️ Download combined precision workbook", data=result,
                file_name="Universal_5_Day_InterRun_Precision.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            preview=pd.DataFrame(all_records)
            st.dataframe(preview, use_container_width=True)
            st.warning("Review the Raw_Data_Audit and Upload_Summary sheets before using results for a regulated validation report. Confirm the QC-level mapping, actual dates and selected replicates.")
    except Exception as e:
        st.exception(e)

st.divider()
st.caption("Important: this is an assisted extraction tool, not a validated laboratory information system. Verify every analyte, QC level, day and selected replicate against the original instrument report.")
