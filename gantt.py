import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ============================ USER SETTINGS ============================
INPUT_FILE = "Project_Scripts_Conversion_Tracking.xlsx"  # your source Excel
SHEET_NAME = "Sheet1"                                    # sheet name
OUTPUT_FILE = "Professional_Gantt_Grouped.xlsx"          # output Excel
# ======================================================================

def build_gantt_from_excel():
    # Read user-provided Excel file
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

    # Required columns check (minimal set)
    required = [
        "Script Name", "Priority",
        "Analysis Start Date", "Analysis End Date",
        "Conversion Start Date", "Conversion End Date",
        "Execution Start Date", "Execution End Date",
        "Recon Start Date", "Recon End Date",
        "Issue Fix Start Date", "Issue Fix End Date",
        "UAT Start Date", "UAT End Date"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Optional helpful columns
    if "Resource" not in df.columns:
        df["Resource"] = ""
    if "Percent Complete" not in df.columns:
        df["Percent Complete"] = ""
    if "Status" not in df.columns:
        df["Status"] = ""
    if "Notes" not in df.columns:
        df["Notes"] = ""

    # Colors with full ARGB (FF = fully opaque)
    phase_colors = {
        "Analysis":    "FF5B9BD5",
        "Conversion":  "FF70AD47",
        "Execution":   "FFED7D31",
        "Recon":       "FFA5A5A5",
        "Issue Fix":   "FFFFC000",
        "UAT":         "FF4472C4",
    }

    # Phases mapping
    phases_map = [
        ("Analysis",   "Analysis Start Date",   "Analysis End Date"),
        ("Conversion", "Conversion Start Date", "Conversion End Date"),
        ("Execution",  "Execution Start Date",  "Execution End Date"),
        ("Recon",      "Recon Start Date",      "Recon End Date"),
        ("Issue Fix",  "Issue Fix Start Date",  "Issue Fix End Date"),
        ("UAT",        "UAT Start Date",        "UAT End Date"),
    ]

    # Group into per-script structure with phases array
    grouped = []
    for script_name, g in df.groupby("Script Name", sort=False):
        row0 = g.iloc[0]
        priority = row0.get("Priority", "")
        resource = row0.get("Resource", "")
        pct = row0.get("Percent Complete", "")
        status = row0.get("Status", "")
        notes = row0.get("Notes", "")

        phases = []
        for phase_name, s_col, e_col in phases_map:
            s_val = row0.get(s_col)
            e_val = row0.get(e_col)
            if pd.notna(s_val) and pd.notna(e_val) and s_val != "########" and e_val != "########":
                try:
                    s_dt = pd.to_datetime(s_val).date()
                    e_dt = pd.to_datetime(e_val).date()
                    if e_dt >= s_dt:
                        phases.append({
                            "Phase": phase_name,
                            "Start": s_dt,
                            "End": e_dt,
                            "Color": phase_colors[phase_name]
                        })
                except:
                    pass

        if phases:
            grouped.append({
                "Script Name": script_name,
                "Priority": priority,
                "Resource": resource,
                "Percent Complete": pct,
                "Status": status,
                "Notes": notes,
                "Phases": phases
            })

    if not grouped:
        raise ValueError("No valid phases with start/end dates found in the input.")

    # Timeline bounds
    all_starts = [p["Start"] for item in grouped for p in item["Phases"]]
    all_ends   = [p["End"]   for item in grouped for p in item["Phases"]]
    min_date = min(all_starts)
    max_date = max(all_ends)
    total_days = (max_date - min_date).days + 1

    # Create workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Gantt_Grouped"

    # Title
    ws.merge_cells("A1:L1")
    title = ws.cell(1, 1, "PROJECT GANTT CHART — GROUPED BY SCRIPT")
    title.font = Font(name="Calibri", size=18, bold=True, color="FFFFFFFF")
    title.fill = PatternFill(start_color="FF2F5597", end_color="FF2F5597", fill_type="solid")
    title.alignment = Alignment(horizontal="center")

    # Headers (task panel)
    headers = ["Script Name", "Priority", "Resource", "Phase", "Start", "End", "% Complete", "Status", "Notes"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(3, i, h)
        c.font = Font(bold=True, color="FFFFFFFF")
        c.fill = PatternFill(start_color="FF4472C4", end_color="FF4472C4", fill_type="solid")
        ws.column_dimensions[get_column_letter(i)].width = 22 if i in (1, 9) else 14

    # Timeline header
    timeline_start_col = len(headers) + 1
    for d in range(total_days):
        dt = min_date + timedelta(days=d)
        col = timeline_start_col + d
        cell = ws.cell(3, col, dt.strftime("%d\n%b"))
        cell.font = Font(bold=True, color="FFFFFFFF", size=9)
        cell.fill = PatternFill(start_color="FF5B9BD5", end_color="FF5B9BD5", fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col)].width = 4

    # Weekend shading
    for d in range(total_days):
        dt = min_date + timedelta(days=d)
        if dt.weekday() >= 5:  # Sat/Sun
            col = timeline_start_col + d
            for r in range(4, 10000):  # applied later effectively
                # we'll limit with actual max_row after drawing
                pass

    # Fill rows: merged Script/Priority/Resource; phases per row
    current_row = 4
    thin = Side(style="thin", color="FF000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for item in grouped:
        phases = item["Phases"]
        n = len(phases)
        start_of_block = current_row
        end_of_block = current_row + n - 1

        # Merge and write Script/Priority/Resource
        ws.merge_cells(start_row=start_of_block, start_column=1, end_row=end_of_block, end_column=1)
        ws.merge_cells(start_row=start_of_block, start_column=2, end_row=end_of_block, end_column=2)
        ws.merge_cells(start_row=start_of_block, start_column=3, end_row=end_of_block, end_column=3)

        ws.cell(start_of_block, 1, item["Script Name"]).alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
        ws.cell(start_of_block, 2, item["Priority"]).alignment = Alignment(vertical="center", horizontal="center")
        ws.cell(start_of_block, 3, item["Resource"]).alignment = Alignment(vertical="center", horizontal="center")

        # Draw each phase row
        for p in phases:
            ws.cell(current_row, 4, p["Phase"])
            ws.cell(current_row, 5, p["Start"])
            ws.cell(current_row, 6, p["End"])
            ws.cell(current_row, 7, item["Percent Complete"])
            ws.cell(current_row, 8, item["Status"])
            ws.cell(current_row, 9, item["Notes"])

            # Bars across timeline
            for d in range(total_days):
                dt = min_date + timedelta(days=d)
                if p["Start"] <= dt <= p["End"]:
                    col = timeline_start_col + d
                    cell = ws.cell(current_row, col, " ")
                    cell.fill = PatternFill(start_color=p["Color"], end_color=p["Color"], fill_type="solid")

            current_row += 1

        # Add a light separator row after each script block
        for c in range(1, timeline_start_col + total_days):
            ws.cell(current_row, c).fill = PatternFill(start_color="FFF3F3F3", end_color="FFF3F3F3", fill_type="solid")
        current_row += 1  # spacer

    max_row = ws.max_row

    # Apply borders to the whole used area
    max_col = timeline_start_col + total_days - 1
    for r in range(3, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).border = border

    # Weekend shading (after knowing max_row)
    for d in range(total_days):
        dt = min_date + timedelta(days=d)
        if dt.weekday() >= 5:
            col = timeline_start_col + d
            for r in range(4, max_row + 1):
                if ws.cell(r, col).fill.start_color.index == "00000000":  # only fill blanks
                    ws.cell(r, col).fill = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")

    # Today marker
    today = datetime.now().date()
    if min_date <= today <= max_date:
        tcol = timeline_start_col + (today - min_date).days
        for r in range(3, max_row + 1):
            ws.cell(r, tcol).border = Border(left=Side(style="thick", color="FFFF0000"))

    # Make the main grid an Excel Table for filtering (esp. Phase/Status)
    table_ref = f"A3:{get_column_letter(max_col)}{max_row}"
    table = Table(displayName="GanttTable", ref=table_ref)
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False, showRowStripes=True, showColumnStripes=False)
    table.tableStyleInfo = style
    ws.add_table(table)

    # Legend
    legend_row = max_row + 2
    ws.cell(legend_row, 1, "Legend").font = Font(size=14, bold=True)
    legend_items = [
        ("FF5B9BD5", "Analysis"),
        ("FF70AD47", "Conversion"),
        ("FFED7D31", "Execution"),
        ("FFA5A5A5", "Recon"),
        ("FFFFC000", "Issue Fix"),
        ("FF4472C4", "UAT"),
        ("FFFF0000", "Today marker (red line)"),
        ("FFF2F2F2", "Weekend shading")
    ]
    for i, (color, label) in enumerate(legend_items, start=1):
        cell = ws.cell(legend_row + i, 1, label)
        cell.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFFFF" if color.startswith("FF") else "FF000000")

    # Freeze panes: keep headers and task columns visible
    ws.freeze_panes = "D4"  # freezes rows above 4 and columns left of D

    # Save file
    wb.save(OUTPUT_FILE)
    print(f"Created {OUTPUT_FILE}")

if __name__ == "__main__":
    build_gantt_from_excel()
