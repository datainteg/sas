import pandas as pd
from datetime import datetime, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ============== USER SETTINGS ==============
INPUT_FILE  = "Project_Scripts_Conversion_Tracking.xlsx"
SHEET_NAME  = "Sheet1"
OUTPUT_FILE = "Professional_Gantt_Grouped.xlsx"
GROUP_PRIORITY = False  # set True if you also want Priority merged like the other identifiers
# ===========================================

def build_gantt_from_excel():
    df = pd.read_excel(INPUT_FILE, sheet_name=SHEET_NAME)

    # Required core columns
    required = [
        "Script Name", "Priority",
        "Analysis Start Date", "Analysis End Date",
        "Conversion Start Date", "Conversion End Date",
        "Execution Start Date", "Execution End Date",
        "Recon Start Date", "Recon End Date",
        "Issue Fix Start Date", "Issue Fix End Date",
        "UAT Start Date", "UAT End Date",
        "Business Function Owner", "Path"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Optional helper columns
    for opt in ["Resource", "Percent Complete", "Status", "Notes"]:
        if opt not in df.columns:
            df[opt] = ""

    # Phase colors (ARGB)
    phase_colors = {
        "Analysis":   "FF5B9BD5",
        "Conversion": "FF70AD47",
        "Execution":  "FFED7D31",
        "Recon":      "FFA5A5A5",
        "Issue Fix":  "FFFFC000",
        "UAT":        "FF4472C4",
    }

    phases_map = [
        ("Analysis",   "Analysis Start Date",   "Analysis End Date"),
        ("Conversion", "Conversion Start Date", "Conversion End Date"),
        ("Execution",  "Execution Start Date",  "Execution End Date"),
        ("Recon",      "Recon Start Date",      "Recon End Date"),
        ("Issue Fix",  "Issue Fix Start Date",  "Issue Fix End Date"),
        ("UAT",        "UAT Start Date",        "UAT End Date"),
    ]

    # Group rows by the three identifiers you specified
    group_keys = ["Business Function Owner", "Path", "Script Name"]
    grouped_items = []
    for keys, g in df.groupby(group_keys, sort=False):
        bfo, path, script = keys
        row0 = g.iloc[0]
        item = {
            "Business Function Owner": bfo,
            "Path": path,
            "Script Name": script,
            "Priority": row0.get("Priority", ""),
            "Resource": row0.get("Resource", ""),
            "Percent Complete": row0.get("Percent Complete", ""),
            "Status": row0.get("Status", ""),
            "Notes": row0.get("Notes", ""),
            "Phases": []
        }
        for phase_name, s_col, e_col in phases_map:
            s_val = row0.get(s_col)
            e_val = row0.get(e_col)
            if pd.notna(s_val) and pd.notna(e_val) and s_val != "########" and e_val != "########":
                try:
                    s_dt = pd.to_datetime(s_val).date()
                    e_dt = pd.to_datetime(e_val).date()
                    if e_dt >= s_dt:
                        item["Phases"].append({
                            "Phase": phase_name,
                            "Start": s_dt,
                            "End": e_dt,
                            "Color": phase_colors[phase_name]
                        })
                except:
                    pass
        if item["Phases"]:
            grouped_items.append(item)

    if not grouped_items:
        raise ValueError("No valid phases with start/end dates found.")

    # Timeline bounds
    all_starts = [p["Start"] for item in grouped_items for p in item["Phases"]]
    all_ends   = [p["End"]   for item in grouped_items for p in item["Phases"]]
    min_date = min(all_starts)
    max_date = max(all_ends)
    total_days = (max_date - min_date).days + 1

    # Workbook and sheet
    wb = Workbook()
    ws = wb.active
    ws.title = "Gantt_Grouped"

    # Title
    ws.merge_cells("A1:R1")
    title = ws.cell(1, 1, "PROJECT GANTT CHART — GROUPED BY BUSINESS FUNCTION OWNER / PATH / SCRIPT NAME")
    title.font = Font(name="Calibri", size=18, bold=True, color="FFFFFFFF")
    title.fill = PatternFill(start_color="FF2F5597", end_color="FF2F5597", fill_type="solid")
    title.alignment = Alignment(horizontal="center")

    # Headers
    headers = [
        "Business Function Owner", "Path", "Script Name",
        "Priority", "Resource", "Phase", "Start", "End",
        "% Complete", "Status", "Notes"
    ]
    for i, h in enumerate(headers, 1):
        c = ws.cell(3, i, h)
        c.font = Font(bold=True, color="FFFFFFFF")
        c.fill = PatternFill(start_color="FF4472C4", end_color="FF4472C4", fill_type="solid")
        ws.column_dimensions[get_column_letter(i)].width = 26 if i in (1,2,3,11) else 14

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

    # Borders
    thin = Side(style="thin", color="FF000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Draw grouped blocks
    current_row = 4
    for item in grouped_items:
        phases = item["Phases"]
        n = len(phases)
        start_block = current_row
        end_block = current_row + n - 1

        # Merge the three identifiers: BFO, Path, Script Name
        merge_map = {
            1: item["Business Function Owner"],
            2: item["Path"],
            3: item["Script Name"],
        }
        # Optionally merge Priority too
        if GROUP_PRIORITY:
            merge_map[4] = item["Priority"]
        else:
            ws.cell(start_block, 4, item["Priority"]).alignment = Alignment(vertical="center", horizontal="center")

        # Resource column: merge or leave per phase? Keep merged, similar to identifiers for clean look
        merge_map[1] = item["Resource"]

        for col_idx, value in merge_map.items():
            ws.merge_cells(start_row=start_block, start_column=col_idx, end_row=end_block, end_column=col_idx)
            cell = ws.cell(start_block, col_idx, value)
            align_h = "left" if col_idx in (1,2,3) else "center"
            cell.alignment = Alignment(vertical="center", horizontal=align_h, wrap_text=True)

        # Phase rows
        for p in phases:
            ws.cell(current_row, 6, p["Phase"])
            ws.cell(current_row, 7, p["Start"])
            ws.cell(current_row, 8, p["End"])
            ws.cell(current_row, 9, item["Percent Complete"])
            status_cell = ws.cell(current_row, 10, item["Status"])
            ws.cell(current_row, 11, item["Notes"])

            # Status coloring
            status_color_map = {
                "Completed": "FF70AD47",
                "In Progress": "FFED7D31",
                "Delayed": "FFFF0000"
            }
            status_fill = status_color_map.get(str(item["Status"]), "FFD9D9D9")
            status_cell.fill = PatternFill(start_color=status_fill, end_color=status_fill, fill_type="solid")
            if status_fill in ("FF70AD47", "FFED7D31", "FFFF0000"):
                status_cell.font = Font(bold=True, color="FFFFFFFF")

            # Bars across timeline
            for d in range(total_days):
                dt = min_date + timedelta(days=d)
                if p["Start"] <= dt <= p["End"]:
                    col = timeline_start_col + d
                    bar = ws.cell(current_row, col, " ")
                    bar.fill = PatternFill(start_color=p["Color"], end_color=p["Color"], fill_type="solid")

            current_row += 1

        # Spacer row
        for c in range(1, timeline_start_col + total_days):
            ws.cell(current_row, c).fill = PatternFill(start_color="FFF3F3F3", end_color="FFF3F3F3", fill_type="solid")
        current_row += 1

    max_row = ws.max_row
    max_col = timeline_start_col + total_days - 1

    # Apply borders
    for r in range(3, max_row + 1):
        for c in range(1, max_col + 1):
            ws.cell(r, c).border = border

    # Weekend shading (light gray where no bar)
    for d in range(total_days):
        dt = min_date + timedelta(days=d)
        if dt.weekday() >= 5:
            col = timeline_start_col + d
            for r in range(4, max_row + 1):
                cell = ws.cell(r, col)
                if cell.fill.start_color is None or cell.fill.start_color.index in ("00000000", None):
                    cell.fill = PatternFill(start_color="FFF2F2F2", end_color="FFF2F2F2", fill_type="solid")

    # Today marker
    today = datetime.now().date()
    if min_date <= today <= max_date:
        tcol = timeline_start_col + (today - min_date).days
        for r in range(3, max_row + 1):
            ws.cell(r, tcol).border = Border(left=Side(style="thick", color="FFFF0000"))

    # Excel Table for filters
    table_ref = f"A3:{get_column_letter(max_col)}{max_row}"
    table = Table(displayName="GanttTable", ref=table_ref)
    style = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True, showColumnStripes=False)
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
        ws.column_dimensions["A"].width = 40

    # Freeze panes: keep identifiers visible
    ws.freeze_panes = "F4"  # keeps left identifiers and header visible

    wb.save(OUTPUT_FILE)
    print(f"Created {OUTPUT_FILE}")

if __name__ == "__main__":
    build_gantt_from_excel()
