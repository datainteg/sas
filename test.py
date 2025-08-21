# build_gantt.py
import pandas as pd
from datetime import timedelta, datetime
from xlsxwriter.utility import rowcol_to_cell

# ---------------- User settings ----------------
SRC_FILE  = "ProjectPlan.xlsx"      # source workbook
SRC_SHEET = "Project Plan"          # sheet name with your table
OUT_FILE  = "Gantt.xlsx"            # output workbook

# If auto-detect fails, hard-set your identifier column here (e.g., "Script Name", "Object", "Control", etc.)
SCRIPT_COL = None  # e.g., "Script Name"

# Highlight today's date column?
HIGHLIGHT_TODAY = True

# Optional: map colors for each phase
PHASE_COLORS = {
    "Analysis":   "#4F81BD",
    "Conversion": "#9BBB59",
    "Execution":  "#C0504D",
    "Recon":      "#8064A2",
    "UAT":        "#F79646",
}
DEFAULT_COLOR = "#7F7F7F"
# ------------------------------------------------


def canonicalize(col: str) -> str:
    """Lowercase, remove spaces and punctuation-like chars for loose matching."""
    if not isinstance(col, str):
        return ""
    return "".join(ch for ch in col.lower() if ch.isalnum())


def autodetect_identifier(cols):
    # Try common names in priority order
    candidates = ["script name", "script", "object", "task", "item", "control", "name", "title"]
    canon_map = {canonicalize(c): c for c in cols}
    for c in candidates:
        if canonicalize(c) in canon_map:
            return canon_map[canonicalize(c)]
    # Fall back to the first non-date column if any
    return cols[0] if cols else None


def find_phase_columns(cols):
    """
    From messy headers like 'Analysis Start D', 'Analysis End D', 'Conversion End Date',
    build a list of (start_col, end_col, phase_name).
    """
    # Define the phases and aliases that might appear in headers
    phases = {
        "Analysis":  ["analysis"],
        "Conversion":["conversion"],
        "Execution": ["execution"],
        "Recon":     ["recon","reconciliation","reconcil"],
        "UAT":       ["uat","useracceptance"],
    }
    # Start/end keywords commonly seen
    start_keys = ["startd", "startdate", "start", "startdt"]
    end_keys   = ["endd", "enddate", "end", "enddt"]

    canon_cols = {canonicalize(c): c for c in cols}
    results = []
    used = set()

    for phase, aliases in phases.items():
        # build a set of canonical tokens to search for
        phase_tokens = [canonicalize(a) for a in aliases]
        start_match, end_match = None, None

        for col in cols:
            ccol = canonicalize(col)
            # must include phase token
            if not any(tok in ccol for tok in phase_tokens):
                continue
            # classify start/end
            if any(k in ccol for k in start_keys) and start_match is None:
                start_match = col
            if any(k in ccol for k in end_keys) and end_match is None:
                end_match = col

        if start_match and end_match:
            results.append((start_match, end_match, phase))
            used.add(start_match); used.add(end_match)

    return results


def read_and_normalize():
    df = pd.read_excel(SRC_FILE, sheet_name=SRC_SHEET, dtype=object)
    # Strip whitespace from headers
    df.columns = [str(c).strip() for c in df.columns]

    # Identify the key column
    id_col = SCRIPT_COL or autodetect_identifier(df.columns.tolist())
    if id_col not in df.columns:
        raise ValueError(f"Could not find identifier column. Set SCRIPT_COL to an exact header. "
                         f"Available columns: {list(df.columns)}")

    phase_defs = find_phase_columns(df.columns.tolist())
    if not phase_defs:
        raise ValueError("No phase start/end columns detected. Ensure headers include words like "
                         "'Analysis Start', 'Analysis End', 'Conversion Start', 'Conversion End', etc.")

    rows = []
    for _, r in df.iterrows():
        ident = r.get(id_col)
        # Skip completely blank identifiers
        if pd.isna(ident):
            continue
        for s_col, e_col, phase in phase_defs:
            start = r.get(s_col)
            end   = r.get(e_col)
            if pd.notna(start) and pd.notna(end):
                try:
                    start_d = pd.to_datetime(start).date()
                    end_d   = pd.to_datetime(end).date()
                except Exception:
                    continue
                if end_d >= start_d:
                    dur = (end_d - start_d).days + 1
                    rows.append({
                        "Item": str(ident),
                        "Phase": phase,
                        "Start": start_d,
                        "End": end_d,
                        "Duration": dur
                    })

    tasks = pd.DataFrame(rows)
    if tasks.empty:
        raise ValueError("No tasks with valid Start/End dates found after parsing. "
                         "Check column names and date values.")
    tasks = tasks.sort_values(["Start", "Item", "Phase"]).reset_index(drop=True)
    return tasks, id_col, phase_defs


def write_gantt(tasks: pd.DataFrame):
    # Timeline
    t0 = tasks["Start"].min()
    t1 = tasks["End"].max()
    days = (t1 - t0).days + 1
    date_seq = [t0 + timedelta(days=i) for i in range(days)]

    with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as xl:
        # Write normalized tasks to its own sheet
        tasks.to_excel(xl, sheet_name="Tasks", index=False)

        # Create Gantt sheet and place tasks table starting row 3 (0-index row=2)
        gantt = xl.book.add_worksheet("Gantt")
        # Write headers for the task table
        headers = ["Item", "Phase", "Start", "End", "Duration"]
        for j, h in enumerate(headers):
            gantt.write(2, j, h)
        # Write task rows
        for i, row in tasks.iterrows():
            gantt.write(3 + i, 0, row["Item"])
            gantt.write(3 + i, 1, row["Phase"])
            gantt.write_datetime(3 + i, 2, pd.to_datetime(row["Start"]))
            gantt.write_datetime(3 + i, 3, pd.to_datetime(row["End"]))
            gantt.write(3 + i, 4, int(row["Duration"]))

        # Timeline header row
        start_col = 6  # start timeline at column G (0-index 6)
        header_row = 2
        date_fmt = xl.book.add_format({"num_format": "dd-mmm", "align": "center", "bold": True})
        for i, d in enumerate(date_seq):
            gantt.write_datetime(header_row, start_col + i, pd.to_datetime(d))
            gantt.set_column(start_col + i, start_col + i, 2.2)
        gantt.set_row(header_row, 18, date_fmt)

        # Freeze panes below header and before grid
        first_data_row = 3
        gantt.freeze_panes(first_data_row, start_col)

        # Weekend shading
        weekend_fmt = xl.book.add_format({"bg_color": "#EEEEEE"})
        nrows = len(tasks)
        for i, d in enumerate(date_seq):
            if d.weekday() >= 5:
                gantt.conditional_format(first_data_row, start_col + i,
                                         first_data_row + nrows, start_col + i,
                                         {"type": "no_blanks", "format": weekend_fmt})

        # Legend
        gantt.write(0, start_col, "Legend:")
        ccol = start_col + 1
        for ph, color in PHASE_COLORS.items():
            fmt = xl.book.add_format({"bg_color": color, "font_color": "white"})
            gantt.write(0, ccol, ph, fmt)
            ccol += 1

        # Gantt bars via conditional formatting per cell
        # Table columns: Item(0) Phase(1) Start(2) End(3) Duration(4)
        for r in range(nrows):
            excel_r = first_data_row + r
            start_addr = rowcol_to_cell(excel_r, 2)
            end_addr   = rowcol_to_cell(excel_r, 3)
            # Resolve color for this row's phase
            phase = tasks.iloc[r]["Phase"]
            color = PHASE_COLORS.get(phase, DEFAULT_COLOR)
            span_fmt = xl.book.add_format({"bg_color": color})

            for i in enumerate(date_seq):
                c = start_col + i[0]
                header_addr = rowcol_to_cell(header_row, c)
                formula_span = f"=AND({header_addr}>={start_addr},{header_addr}<={end_addr})"
                gantt.conditional_format(excel_r, c, excel_r, c, {
                    "type": "formula",
                    "criteria": formula_span,
                    "format": span_fmt
                })

        # Today column (optional)
        if HIGHLIGHT_TODAY:
            today = datetime.today().date()
            for i, d in enumerate(date_seq):
                if d == today:
                    today_fmt = xl.book.add_format({"bg_color": "#FF0000"})
                    gantt.conditional_format(first_data_row, start_col + i,
                                             first_data_row + nrows, start_col + i,
                                             {"type": "no_blanks", "format": today_fmt})
                    break

        # Basic formatting widths
        gantt.set_column(0, 0, 28)  # Item
        gantt.set_column(1, 1, 14)  # Phase
        gantt.set_column(2, 3, 12)  # Start/End
        gantt.set_column(4, 4, 10)  # Duration

    print(f"Created {OUT_FILE} with Tasks and Gantt sheets.")


if __name__ == "__main__":
    tasks, id_col, phase_defs = read_and_normalize()
    print("Detected identifier column:", id_col)
    print("Detected phases (start/end/phase):")
    for s, e, p in phase_defs:
        print(f" - {p}: {s} | {e}")
    write_gantt(tasks)
