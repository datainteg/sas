import pandas as pd
from datetime import datetime, timedelta

# User settings
SRC_FILE = "ProjectPlan.xlsx"        # your source file
SRC_SHEET = "Project Plan"           # sheet with your table
OUT_FILE = "Gantt.xlsx"
PHASES = [
    ("Analysis Start Date", "Analysis End Date", "Analysis"),
    ("Conversion Start Date", "Conversion End Date", "Conversion"),
    ("Execution Start Date", "Execution End Date", "Execution"),
    ("Recon Start Date", "Recon End Date", "Recon"),
    ("UAT Start Date", "UAT End Date", "UAT"),
]
SCRIPT_COL = "Script Name"           # adjust to exact header
PCT_COL = None                       # e.g., "% Complete" if exists

# 1) Read source
df = pd.read_excel(SRC_FILE, sheet_name=SRC_SHEET)

# 2) Normalize to long format
rows = []
for _, r in df.iterrows():
    script = r.get(SCRIPT_COL)
    for s_col, e_col, phase in PHASES:
        start = r.get(s_col)
        end = r.get(e_col)
        if pd.notna(start) and pd.notna(end):
            start = pd.to_datetime(start).date()
            end = pd.to_datetime(end).date()
            dur = (end - start).days + 1
            pct = float(r.get(PCT_COL)) if PCT_COL and pd.notna(r.get(PCT_COL)) else None
            rows.append({
                "Script Name": script,
                "Phase": phase,
                "Start": start,
                "End": end,
                "Duration": dur,
                "% Complete": pct
            })
tasks = pd.DataFrame(rows).sort_values(["Start","Script Name","Phase"])

if tasks.empty:
    raise SystemExit("No tasks with valid start/end found. Check column headers/phases.")

# 3) Compute timeline range
timeline_start = tasks["Start"].min()
timeline_end = tasks["End"].max()
days = (timeline_end - timeline_start).days + 1
date_seq = [timeline_start + timedelta(days=i) for i in range(days)]

# 4) Write Excel with Gantt grid
with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as xl:
    tasks.to_excel(xl, sheet_name="Gantt", index=False, startrow=2)
    wb  = xl.book
    ws  = xl.sheets["Gantt"]

    # Headers
    ws.write(0, 0, "Gantt Chart")
    ws.write(1, 0, "Note: Edit phases and columns in script if names differ.")

    # Place timeline in row 3, from column F (col index 5)
    start_col = 5
    for i, d in enumerate(date_seq):
        ws.write_datetime(2, start_col + i, pd.to_datetime(d))
        # narrow columns
        ws.set_column(start_col + i, start_col + i, 2.2)

    # Format
    date_fmt = wb.add_format({"num_format": "dd-mmm", "align": "center", "bold": True})
    ws.set_row(2, 18, date_fmt)

    # Freeze panes below headers and before grid
    ws.freeze_panes(3, start_col)

    # Weekend shading
    weekend_fmt = wb.add_format({"bg_color":"#EEEEEE"})
    for i, d in enumerate(date_seq):
        if d.weekday() >= 5:
            ws.conditional_format(3, start_col + i, 3 + len(tasks), start_col + i, {
                "type": "no_blanks", "format": weekend_fmt
            })

    # Gantt bar formatting by phase color
    phase_colors = {
        "Analysis":"#4F81BD",
        "Conversion":"#9BBB59",
        "Execution":"#C0504D",
        "Recon":"#8064A2",
        "UAT":"#F79646",
    }
    # Write a legend
    ws.write(0, 5, "Legend:")
    ccol = 6
    for ph, color in phase_colors.items():
        fmt = wb.add_format({"bg_color": color, "font_color":"white"})
        ws.write(0, ccol, ph, fmt); ccol += 1

    # Build conditional formats: one rule per row for span, colored by phase
    header_cols = ["Script Name","Phase","Start","End","Duration","% Complete"]
    nrows = len(tasks)
    # Find column offsets
    col_index = {name:i for i,name in enumerate(header_cols)}
    row0 = 3  # first data row (0-based index)

    for r in range(nrows):
        # Excel ranges are 0-based here for xlsxwriter
        start_cell = xl._xlsxwriter.sheets["Gantt"].table  # just to satisfy linter; not used

        # Row coordinates
        r_top = row0 + r
        # Build addresses for Start, End, Phase, Duration, %Complete
        start_addr = xl.book._xlsxwriter_conversion.rowcol_to_cell(r_top, 2, row_abs=False, col_abs=False)  # column "Start"
        end_addr   = xl.book._xlsxwriter_conversion.rowcol_to_cell(r_top, 3, row_abs=False, col_abs=False)
        phase_addr = xl.book._xlsxwriter_conversion.rowcol_to_cell(r_top, 1, row_abs=False, col_abs=False)
        dur_addr   = xl.book._xlsxwriter_conversion.rowcol_to_cell(r_top, 4, row_abs=False, col_abs=False)
        pct_addr   = xl.book._xlsxwriter_conversion.rowcol_to_cell(r_top, 5, row_abs=False, col_abs=False)

        # Span rule: fill if date header between Start/End
        # For each timeline column, we use a single row-level rule over full date range
        # Formula refers to top-left cell in each applied range column by column reference
        for i, d in enumerate(date_seq):
            col = start_col + i
            # We apply a simple "no formula" row-wide by using a per-column formula rule
            color = phase_colors.get(tasks.iloc[r]["Phase"], "#7F7F7F")
            fmt = wb.add_format({"bg_color": color})
            formula = f"=AND({xl.book._xlsxwriter_conversion.rowcol_to_cell(2, col)}>={start_addr},{xl.book._xlsxwriter_conversion.rowcol_to_cell(2, col)}<={end_addr})"
            ws.conditional_format(r_top, col, r_top, col, {
                "type":"formula",
                "criteria": formula,
                "format": fmt
            })

            # Optional % complete darker overlay
            if "% Complete" in tasks.columns and pd.notna(tasks.iloc[r]["% Complete"]):
                pct_fmt = wb.add_format({"bg_color": "#3C3C3C"})
                # Completed until Start + round(Duration*pct)-1
                formula_done = f"=AND({xl.book._xlsxwriter_conversion.rowcol_to_cell(2, col)}>={start_addr},{xl.book._xlsxwriter_conversion.rowcol_to_cell(2, col)}<={start_addr}+ROUND({dur_addr}*{pct_addr},0)-1)"
                ws.conditional_format(r_top, col, r_top, col, {
                    "type":"formula",
                    "criteria": formula_done,
                    "format": pct_fmt
                })

    # Today line
    today_fmt = wb.add_format({"bg_color":"#FF0000"})
    for i, d in enumerate(date_seq):
        if d == datetime.today().date():
            ws.conditional_format(3, start_col + i, 3 + nrows, start_col + i, {
                "type": "no_blanks", "format": today_fmt
            })
            break

print(f"Created {OUT_FILE}")
