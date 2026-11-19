import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUTPUT = r"c:\Users\Aayush Mathur\OneDrive\Desktop\Employee Performance\docs\04_Traceability_Matrix.xlsx"

# --- RTM DATA ---
# Columns: OQ, HMW, BR_ID, Requirement, Stakeholder, Priority, Status
# (Feature, DB, US, UI, TC-F, TC-T left blank)
RTM_ROWS = [
    # Direct 1-to-1 links
    ("OQ-01", "HMW-01", "BR-01", "Dept-wise avg performance rating dashboard",        "SH-01",       "Must",   "Identified"),
    ("OQ-02", "HMW-02", "BR-02", "Employee ranking – top/bottom 10%",                  "SH-01",       "Must",   "Identified"),
    ("OQ-03", "HMW-03", "BR-03", "Salary-vs-rating correlation per dept",              "SH-02",       "Should", "Identified"),
    ("OQ-04", "HMW-04", "BR-04", "Manager ranking by team avg rating",                 "SH-02",       "Must",   "Identified"),
    ("OQ-05", "HMW-05", "BR-05", "Rating distribution by gender/age + bias check",     "SH-03",       "Must",   "Identified"),
    ("OQ-10", "HMW-06", "BR-06", "Promotion-readiness score (0-100)",                  "SH-04",       "Must",   "Identified"),
    ("OQ-10", "HMW-06", "BR-14", "Flag employees overdue for promotion (≥4, >3 yrs)",  "SH-01+SH-04", "Must",   "Identified"),
    ("OQ-07", "HMW-07", "BR-07", "Attrition-risk score per department",                "SH-03",       "Must",   "Identified"),
    ("OQ-08", "HMW-09", "BR-08", "3-year company performance trend chart",             "SH-05",       "Must",   "Identified"),
    ("OQ-08", "HMW-12", "BR-08", "3-year performance trend (secondary mapping)",       "SH-05",       "Must",   "Identified"),
    ("OQ-09", "HMW-08", "BR-09", "Regression model – performance predictors",         "SH-05",       "Should", "Identified"),
    # HMW-10 covers all OQs → BR-11, BR-12
    ("OQ-01", "HMW-10", "BR-11", "Data-quality report for all OQs",                   "SH-06",       "Must",   "Identified"),
    ("OQ-02", "HMW-10", "BR-11", "Data-quality report – performer history",            "SH-06",       "Must",   "Identified"),
    ("OQ-03", "HMW-10", "BR-11", "Data-quality report – salary fields",                "SH-06",       "Must",   "Identified"),
    ("OQ-04", "HMW-10", "BR-11", "Data-quality report – manager mapping",              "SH-06",       "Must",   "Identified"),
    ("OQ-05", "HMW-10", "BR-11", "Data-quality report – demographic fields",           "SH-06",       "Must",   "Identified"),
    ("OQ-06", "HMW-10", "BR-11", "Data-quality report – job role fields",              "SH-06",       "Must",   "Identified"),
    ("OQ-07", "HMW-10", "BR-11", "Data-quality report – attrition/tenure fields",     "SH-06",       "Must",   "Identified"),
    ("OQ-08", "HMW-10", "BR-11", "Data-quality report – date/year fields",             "SH-06",       "Must",   "Identified"),
    ("OQ-09", "HMW-10", "BR-11", "Data-quality report – training/education fields",   "SH-06",       "Must",   "Identified"),
    ("OQ-10", "HMW-10", "BR-11", "Data-quality report – promotion/tenure fields",     "SH-06",       "Must",   "Identified"),
    ("OQ-01", "HMW-10", "BR-12", "CSV/Excel download for all OQ reports",              "SH-05",       "Could",  "Identified"),
    ("OQ-08", "HMW-10", "BR-12", "CSV/Excel download – trend data export",             "SH-05",       "Could",  "Identified"),
    # HMW-11 → BR-10, BR-13
    ("OQ-06", "HMW-11", "BR-10", "Interactive filters on all dashboard screens",       "SH-01+SH-02", "Must",   "Identified"),
    ("OQ-06", "HMW-11", "BR-13", "ANOVA/Kruskal-Wallis test across job roles",        "SH-02",       "Should", "Identified"),
]

HEADERS = [
    "OQ", "HMW", "BR ID", "Requirement", "Stakeholder",
    "Priority", "Status", "Feature", "DB table.column",
    "User Story", "UI Screen", "TC-F", "TC-T"
]

COL_WIDTHS = [8, 8, 8, 35, 12, 10, 12, 15, 20, 12, 12, 10, 10]

# Fills
HEADER_FILL   = PatternFill("solid", fgColor="4472C4")
ALT_FILL      = PatternFill("solid", fgColor="E8E8E8")
MUST_FILL     = PatternFill("solid", fgColor="FFCCCC")   # light red
SHOULD_FILL   = PatternFill("solid", fgColor="FFFF99")   # light yellow
COULD_FILL    = PatternFill("solid", fgColor="CCFFCC")   # light green
WHITE_FONT    = Font(color="FFFFFF", bold=True)
BOLD_FONT     = Font(bold=True)
HEADER_ALIGN  = Alignment(horizontal="center", vertical="center", wrap_text=True)
CENTER_ALIGN  = Alignment(horizontal="center", vertical="center")
WRAP_ALIGN    = Alignment(horizontal="left", vertical="center", wrap_text=True)

PRIORITY_COL_IDX = 6   # 1-based: Priority is column 6

def priority_fill(p):
    if p == "Must":   return MUST_FILL
    if p == "Should": return SHOULD_FILL
    if p == "Could":  return COULD_FILL
    return None

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

wb = openpyxl.Workbook()

# ===== SHEET 1: RTM =====
ws_rtm = wb.active
ws_rtm.title = "RTM"
ws_rtm.sheet_properties.tabColor = "4472C4"

# Header
for col_idx, (hdr, width) in enumerate(zip(HEADERS, COL_WIDTHS), start=1):
    cell = ws_rtm.cell(row=1, column=col_idx, value=hdr)
    cell.fill       = HEADER_FILL
    cell.font       = WHITE_FONT
    cell.alignment  = HEADER_ALIGN
    cell.border     = thin_border()
    ws_rtm.column_dimensions[get_column_letter(col_idx)].width = width

ws_rtm.row_dimensions[1].height = 30

# Data rows
for r_idx, row_data in enumerate(RTM_ROWS, start=2):
    oq, hmw, br_id, req, sh, pri, status = row_data
    full_row = [oq, hmw, br_id, req, sh, pri, status, "", "", "", "", "", ""]
    is_odd = (r_idx % 2 == 1)
    row_fill = ALT_FILL if is_odd else None

    for c_idx, val in enumerate(full_row, start=1):
        cell = ws_rtm.cell(row=r_idx, column=c_idx, value=val)
        cell.border = thin_border()
        # Alignment
        if c_idx == 4:          # Requirement
            cell.alignment = WRAP_ALIGN
        elif c_idx in (1, 2, 3, 5, 6, 7, 11, 12, 13):
            cell.alignment = CENTER_ALIGN
        else:
            cell.alignment = WRAP_ALIGN

        # Row shading (applied before priority override)
        if row_fill and c_idx != PRIORITY_COL_IDX:
            cell.fill = row_fill

        # Priority colour in priority column
        if c_idx == PRIORITY_COL_IDX:
            pf = priority_fill(pri)
            if pf:
                cell.fill = pf
            cell.font = BOLD_FONT

# Auto-filter
ws_rtm.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"

# Freeze row 1
ws_rtm.freeze_panes = "A2"

# ===== SHEET 2: Coverage =====
ws_cov = wb.create_sheet("Coverage")
ws_cov.sheet_properties.tabColor = "70AD47"   # green

# -- Helpers --
def cov_header(ws, row, text):
    cell = ws.cell(row=row, column=1, value=text)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="70AD47")
    cell.alignment = HEADER_ALIGN
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=3)

def cov_ok(ws, row, text):
    ws.cell(row=row, column=1, value=text).font = Font(italic=True, color="808080")

# All HMWs present in RTM
all_hmws  = {r[1] for r in RTM_ROWS}  # HMW-01..HMW-12 (HMW-10 split into rows per OQ)
all_brs   = {r[2] for r in RTM_ROWS}
all_oqs   = {r[0] for r in RTM_ROWS}

expected_hmws = {f"HMW-{i:02d}" for i in range(1, 13)}
expected_brs  = {f"BR-{i:02d}"  for i in range(1, 15)}
expected_oqs  = {f"OQ-{i:02d}"  for i in range(1, 11)}

missing_hmws = sorted(expected_hmws - all_hmws)
missing_brs  = sorted(expected_brs  - all_brs)
missing_oqs  = sorted(expected_oqs  - all_oqs)

row = 1
cov_header(ws_cov, row, "1. HMWs without any BR mapping")
row += 1
if missing_hmws:
    for h in missing_hmws:
        ws_cov.cell(row=row, column=1, value=h); row += 1
else:
    cov_ok(ws_cov, row, "✓ All HMW-01 to HMW-12 are mapped to at least one BR"); row += 1

row += 1
cov_header(ws_cov, row, "2. BRs without any HMW mapping")
row += 1
if missing_brs:
    for b in missing_brs:
        ws_cov.cell(row=row, column=1, value=b); row += 1
else:
    cov_ok(ws_cov, row, "✓ All BR-01 to BR-14 are mapped to at least one HMW"); row += 1

row += 1
cov_header(ws_cov, row, "3. OQs not covered by any HMW")
row += 1
if missing_oqs:
    for o in missing_oqs:
        ws_cov.cell(row=row, column=1, value=o); row += 1
else:
    cov_ok(ws_cov, row, "✓ All OQ-01 to OQ-10 are covered by at least one HMW"); row += 1

row += 2
cov_header(ws_cov, row, "4. Summary")
row += 1
summaries = [
    ("Total RTM rows",         len(RTM_ROWS)),
    ("Unique HMWs mapped",     len(all_hmws)),
    ("Unique BRs covered",     len(all_brs)),
    ("Unique OQs covered",     len(all_oqs)),
    ("Must requirements",      sum(1 for r in RTM_ROWS if r[5] == "Must")),
    ("Should requirements",    sum(1 for r in RTM_ROWS if r[5] == "Should")),
    ("Could requirements",     sum(1 for r in RTM_ROWS if r[5] == "Could")),
    ("Status = Identified",    sum(1 for r in RTM_ROWS if r[6] == "Identified")),
]
for label, val in summaries:
    ws_cov.cell(row=row, column=1, value=label).font = Font(bold=True)
    ws_cov.cell(row=row, column=2, value=val)
    row += 1

ws_cov.column_dimensions["A"].width = 45
ws_cov.column_dimensions["B"].width = 15
ws_cov.column_dimensions["C"].width = 15

wb.save(OUTPUT)

import os
assert os.path.exists(OUTPUT), "File not found!"
# Verify sheets
import openpyxl as _op
_wb = _op.load_workbook(OUTPUT)
assert "RTM"      in _wb.sheetnames, "RTM sheet missing"
assert "Coverage" in _wb.sheetnames, "Coverage sheet missing"
print(f"SUCCESS: {OUTPUT}")
print(f"RTM rows (data): {len(RTM_ROWS)}")
print(f"Sheets: {_wb.sheetnames}")
