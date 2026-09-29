#!/usr/bin/env python3
"""
Generate Master IO List Sorted by Chassis and Slot for Project Jet Cooker.
Ingredion (Thailand) Co., Ltd. | AEC Industrial Engineering

Deliverables:
1. IO_List_Sorted_by_Chassis_Slot.xlsx
   - Chassis_Slot_Summary: Executive hardware KPI cards & complete slot allocation table
   - IO_List_Master_Sorted: Clean engineering presentation sorted by Chassis, Slot, Point
   - IO_List_Raw_Columns_Sorted: All 42 original columns sorted by Chassis, Slot, Point
   - Chassis_C1 .. Chassis_C7: Individual chassis wiring & commissioning tabs
2. IO_List_xDev-R01-Tag35-6.xlsx
   - Official next-revision master workbook with 'IO List' sheet sorted by Chassis & Slot
"""

import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_EXCEL = os.path.join(WORKSPACE_DIR, "IO_List_xDev-R01-Tag35-5.xlsx")
OUTPUT_STANDALONE = os.path.join(WORKSPACE_DIR, "IO_List_Sorted_by_Chassis_Slot.xlsx")
OUTPUT_NEXT_REV = os.path.join(WORKSPACE_DIR, "IO_List_xDev-R01-Tag35-6.xlsx")

def clean_val(val):
    if pd.isna(val):
        return "-"
    s = str(val).strip()
    if s.lower() in ["nan", "none", "", "#n/a"]:
        return "-"
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s

def chassis_sort_key(ch):
    s = str(ch).strip().upper()
    if s.startswith("C") and s[1:].isdigit():
        return int(s[1:])
    return 999

def slot_sort_key(sl):
    try:
        return float(sl)
    except:
        return 999.0

def point_sort_key(pt):
    try:
        return float(pt)
    except:
        return 999.0

def get_base_card(card_val):
    if pd.isna(card_val):
        return "-"
    s = str(card_val).strip()
    parts = s.split("-")
    if len(parts) >= 3 and parts[-1].isdigit():
        return "-".join(parts[:-1])
    return s

def set_cell(cell, value, font=None, fill=None, alignment=None, border=None, is_text=False):
    """Safely set cell value, ensuring strings starting with '=' are stored as text."""
    if value is not None and isinstance(value, str) and (value.startswith("=") or is_text):
        cell.value = value
        cell.data_type = 's'
    else:
        cell.value = value
        
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border

def build_sorted_dataset():
    print(f"Loading source dataset: {SOURCE_EXCEL}")
    df_raw = pd.read_excel(SOURCE_EXCEL, sheet_name="IO List")
    
    # Exclude summary/total rows (e.g. Destination == '112')
    df_clean = df_raw[df_raw["Destination"].astype(str).str.strip() != "112"].copy()
    df_clean = df_clean[df_clean["Destination"].notna()].copy()

    # Assign sort keys: Chassis -> Slot -> Point -> Instrument Tag -> PLC Tag
    df_clean["_ch_key"] = df_clean["Chassis"].apply(chassis_sort_key)
    df_clean["_sl_key"] = df_clean["Slot"].apply(slot_sort_key)
    df_clean["_pt_key"] = df_clean["Point"].apply(point_sort_key)
    df_clean["_tag_key"] = df_clean["Instruement Tag"].astype(str).str.upper()
    df_clean["_plc_key"] = df_clean["PLC_Tag"].astype(str).str.upper()

    df_sorted = df_clean.sort_values(
        by=["_ch_key", "_sl_key", "_pt_key", "_tag_key", "_plc_key"]
    ).reset_index(drop=True)

    print(f"Total sorted valid I/O channels: {len(df_sorted)}")
    return df_sorted

# Styling definitions
NAVY_HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SUB_HEADER_FILL = PatternFill(start_color="2D4A77", end_color="2D4A77", fill_type="solid")
SUB_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

KPI_TITLE_FILL = PatternFill(start_color="0B2545", end_color="0B2545", fill_type="solid")
KPI_TITLE_FONT = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
CARD_HEADER_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

HEADER_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='medium', color='0B2545'),
    bottom=Side(style='medium', color='0B2545')
)

ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')

# Color fills for IO Type badges
TYPE_STYLES = {
    "DI": {
        "fill": PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="2B6CB0")
    },
    "DO": {
        "fill": PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="B7791F")
    },
    "AI": {
        "fill": PatternFill(start_color="F0FFF4", end_color="F0FFF4", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="276749")
    },
    "AO": {
        "fill": PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="234E52")
    },
    "BUS": {
        "fill": PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="553C9A")
    }
}

STATUS_STYLES = {
    "ACTIVE": {
        "fill": PatternFill(start_color="DEF7EC", end_color="DEF7EC", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=True, color="03543F")
    },
    "SPARE": {
        "fill": PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid"),
        "font": Font(name="Calibri", size=10, bold=False, color="718096")
    }
}

CARD_FUNCTION_MAP = {
    "1756-IB32": "32-Pt 24VDC Digital Input Module",
    "1756-OB32": "32-Pt 24VDC Digital Output Module",
    "1756-IF16": "16-Pt 4-20mA HART Analog Input Module",
    "1756-OF8": "8-Pt 4-20mA HART Analog Output Module",
    "-": "MCC Draw-out Terminal / Bus Interface"
}

def create_kpi_card(ws, start_row, start_col, title, value, subtitle, accent_color="1B365D"):
    accent_fill = PatternFill(start_color=accent_color, end_color=accent_color, fill_type="solid")
    
    # Title bar
    set_cell(ws.cell(start_row, start_col), title, font=Font(name="Calibri", size=10, bold=True, color="FFFFFF"), fill=accent_fill, alignment=ALIGN_CENTER)
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col+1)
    
    # Value box
    set_cell(ws.cell(start_row+1, start_col), value, font=Font(name="Calibri", size=18, bold=True, color="1E293B"), alignment=ALIGN_CENTER)
    ws.merge_cells(start_row=start_row+1, start_column=start_col, end_row=start_row+2, end_column=start_col+1)
    
    # Subtitle box
    set_cell(ws.cell(start_row+3, start_col), subtitle, font=Font(name="Calibri", size=9, italic=True, color="64748B"), alignment=ALIGN_CENTER)
    ws.merge_cells(start_row=start_row+3, start_column=start_col, end_row=start_row+3, end_column=start_col+1)
    
    # Borders
    for r in range(start_row, start_row+4):
        for c in range(start_col, start_col+2):
            ws.cell(r, c).border = THIN_BORDER

def populate_summary_sheet(ws, df_sorted):
    ws.views.sheetView[0].showGridLines = True
    
    # Banner
    ws.cell(1, 1, "INGREDION (THAILAND) CO., LTD.  |  PROJECT JET COOKER").font = Font(name="Calibri", size=14, bold=True, color="1B365D")
    ws.cell(2, 1, "PLC I/O SYSTEM ARCHITECTURE & CHASSIS / SLOT SUMMARY REPORT").font = Font(name="Calibri", size=12, bold=True, color="475569")
    ws.cell(3, 1, "Master I/O Allocation Sorted by Chassis & Slot | Engineering Deliverable").font = Font(name="Calibri", size=10, italic=True, color="64748B")
    
    # Compute Metrics
    total_pts = len(df_sorted)
    is_spare = df_sorted["Instruement Tag"].astype(str).str.strip().str.lower().str.contains("spare")
    spare_pts = int(is_spare.sum())
    act_pts = int(total_pts - spare_pts)
    spare_pct = (spare_pts / total_pts * 100) if total_pts > 0 else 0
    
    num_chassis = int(df_sorted["Chassis"].nunique())
    num_slots = int(df_sorted.groupby(["Chassis", "Slot"]).ngroups)
    
    # KPI Cards row 5
    create_kpi_card(ws, 5, 1, "TOTAL I/O POINTS", f"{total_pts:,}", "Allocated across all chassis", "1B365D")
    create_kpi_card(ws, 5, 3, "ACTIVE CHANNELS", f"{act_pts:,}", f"{act_pts/total_pts*100:.1f}% operational load", "03543F")
    create_kpi_card(ws, 5, 5, "SPARE CHANNELS", f"{spare_pts:,}", f"{spare_pct:.1f}% spare capacity", "4A5568")
    create_kpi_card(ws, 5, 7, "TOTAL CHASSIS", f"{num_chassis}", "C1, C2, C3, C4, C5, C6, C7", "2B6CB0")
    create_kpi_card(ws, 5, 9, "ACTIVE I/O SLOTS", f"{num_slots}", "Populated module slots", "7C3AED")

    # Table 1: I/O Type Breakdown (Row 10)
    ws.cell(10, 1, "1. I/O Signal Type Distribution").font = Font(name="Calibri", size=11, bold=True, color="1B365D")
    type_headers = ["I/O Type", "Description", "Total Points", "Active Points", "Spare Points", "Spare %", "Signal Level / Standard"]
    for c_idx, h in enumerate(type_headers, 1):
        set_cell(ws.cell(11, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx not in [2, 7] else ALIGN_LEFT)

    type_desc = {
        "DI": ("Digital Input", "24 VDC Sink/Source (Dry Contact / Proximity)"),
        "DO": ("Digital Output", "24 VDC Sourcing (Solenoid Valve / Interlock)"),
        "AI": ("Analog Input", "4 - 20 mA Current Loop (HART / Transmitter)"),
        "AO": ("Analog Output", "4 - 20 mA Current Loop (Control Valve / I-P)"),
        "BUS": ("Communication Bus", "Industrial Ethernet / Serial Bus Link")
    }

    cur_r = 12
    for io_t in ["DI", "DO", "AI", "AO", "BUS"]:
        sub_t = df_sorted[df_sorted["I/O Type"] == io_t]
        t_tot = len(sub_t)
        t_spr = int(sub_t["Instruement Tag"].astype(str).str.strip().str.lower().str.contains("spare").sum())
        t_act = int(t_tot - t_spr)
        t_pct = (t_spr / t_tot * 100) if t_tot > 0 else 0
        desc_info = type_desc.get(io_t, (io_t, "-"))

        row_vals = [io_t, desc_info[0], t_tot, t_act, t_spr, f"{t_pct:.1f}%", desc_info[1]]
        fill = ZEBRA_ODD if (cur_r % 2) == 1 else ZEBRA_EVEN
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(cur_r, c_idx)
            if c_idx == 1 and io_t in TYPE_STYLES:
                set_cell(cell, val, font=TYPE_STYLES[io_t]["font"], fill=TYPE_STYLES[io_t]["fill"],
                         border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx in [3, 4, 5, 6]:
                set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill,
                         border=THIN_BORDER, alignment=ALIGN_RIGHT)
            else:
                set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill,
                         border=THIN_BORDER, alignment=ALIGN_LEFT)
        cur_r += 1

    # Total row for Type Breakdown
    set_cell(ws.cell(cur_r, 1), "TOTAL", font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)
    set_cell(ws.cell(cur_r, 2), "All Signal Types Combined", font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_LEFT)
    set_cell(ws.cell(cur_r, 3), total_pts, font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_RIGHT)
    set_cell(ws.cell(cur_r, 4), act_pts, font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_RIGHT)
    set_cell(ws.cell(cur_r, 5), spare_pts, font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_RIGHT)
    set_cell(ws.cell(cur_r, 6), f"{spare_pct:.1f}%", font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_RIGHT)
    set_cell(ws.cell(cur_r, 7), "-", font=Font(name="Calibri", size=10, bold=True), fill=CARD_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_LEFT)

    # Table 2: Complete Chassis & Slot Hardware Allocation Matrix (Row cur_r + 3)
    cur_r += 3
    ws.cell(cur_r, 1, "2. Chassis & Slot Hardware Allocation Matrix (Sorted)").font = Font(name="Calibri", size=11, bold=True, color="1B365D")
    cur_r += 1
    slot_headers = [
        "Chassis", "Slot", "Enclosure / Panel", "Slot Address", "Module / Card Model",
        "Module Description", "I/O Type", "Total Pts", "Active Pts", "Spare Pts", "Spare %",
        "Connected Junction Boxes", "Channel Range"
    ]
    for c_idx, h in enumerate(slot_headers, 1):
        set_cell(ws.cell(cur_r, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [1, 2, 4, 7] else (ALIGN_RIGHT if c_idx in [8, 9, 10, 11] else ALIGN_LEFT))

    cur_r += 1
    # Group by chassis and slot
    df_sorted["_card_base"] = df_sorted["Card"].apply(get_base_card)
    
    for _, g in df_sorted.groupby(["_ch_key", "Chassis", "_sl_key"]):
        ch = str(g["Chassis"].iloc[0]).strip()
        sl = int(g["Slot"].iloc[0])
        ref_des = ", ".join(sorted(set(str(v).strip() for v in g[" Reference Designation"].dropna() if str(v).strip())))
        slot_addr = f"{ch}S{sl:02d}"
        
        cards = [c for c in g["_card_base"].unique() if c != "-"]
        card_name = cards[0] if cards else "-"
        card_desc = CARD_FUNCTION_MAP.get(card_name, "ControlLogix I/O Module")
        io_type = "/".join(sorted(g["I/O Type"].unique()))
        
        pts_tot = len(g)
        is_sp = g["Instruement Tag"].astype(str).str.strip().str.lower().str.contains("spare")
        pts_spr = int(is_sp.sum())
        pts_act = int(pts_tot - pts_spr)
        sp_pct = (pts_spr / pts_tot * 100) if pts_tot > 0 else 0
        
        jbs = ", ".join(sorted(set(str(v).strip() for v in g["Destination"].dropna() if str(v).strip())))
        min_pt = int(g["Point"].min())
        max_pt = int(g["Point"].max())
        pt_range = f"Pt {min_pt:02d} - Pt {max_pt:02d}"
        
        fill = ZEBRA_ODD if (cur_r % 2) == 1 else ZEBRA_EVEN
        row_vals = [
            ch, sl, ref_des, slot_addr, card_name, card_desc,
            io_type, pts_tot, pts_act, pts_spr, f"{sp_pct:.1f}%",
            jbs, pt_range
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(cur_r, c_idx)
            if c_idx in [1, 2, 4]:
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=(c_idx in [1, 4])),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 7:
                if io_type in TYPE_STYLES:
                    set_cell(cell, val, font=TYPE_STYLES[io_type]["font"], fill=TYPE_STYLES[io_type]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill,
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx in [8, 9, 10, 11]:
                set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill,
                         border=THIN_BORDER, alignment=ALIGN_RIGHT)
            elif c_idx == 3: # Reference Designation
                set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill,
                         border=THIN_BORDER, alignment=ALIGN_LEFT, is_text=True)
            else:
                set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill,
                         border=THIN_BORDER, alignment=ALIGN_LEFT)
        cur_r += 1

    # Adjust column widths
    col_widths = {
        "A": 10, "B": 8, "C": 20, "D": 14, "E": 18, "F": 34,
        "G": 12, "H": 11, "I": 11, "J": 11, "K": 11, "L": 30, "M": 16
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w


def populate_master_sorted_sheet(ws, df_sorted):
    ws.views.sheetView[0].showGridLines = True
    
    # Title Block (Rows 1-2)
    ws.merge_cells("A1:U1")
    title_cell = ws["A1"]
    set_cell(title_cell, "INGREDION (THAILAND) CO., LTD.  |  PROJECT JET COOKER  |  MASTER INSTRUMENT I/O LIST",
             font=Font(name="Calibri", size=13, bold=True, color="FFFFFF"), fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:U2")
    sub_cell = ws["A2"]
    set_cell(sub_cell, "I/O Points Sorted by Chassis, Slot, Point  |  Rockwell ControlLogix 1756 Platform  |  AEC Industrial Engineering",
             font=Font(name="Calibri", size=10, italic=True, color="FFFFFF"), fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Seq No.",
        "Chassis",
        "Slot",
        "Point",
        "Channel Code",
        "Terminal Block",
        "Module / Card",
        "I/O Type",
        "Signal Standard",
        "Instrument Tag",
        "Instrument Description",
        "Control Description",
        "PLC Tag Name",
        "P&ID No.",
        "Destination JB",
        "JB Location",
        "Zone",
        "Floor",
        "Panel Enclosure",
        "Status",
        "Required I/O Spec"
    ]
    
    header_row = 4
    ws.row_dimensions[header_row].height = 28
    for c_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(header_row, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [1, 2, 3, 4, 5, 6, 8, 18, 20] else ALIGN_LEFT)

    # Data Rows
    row_idx = 5
    for seq, (_, row) in enumerate(df_sorted.iterrows(), 1):
        chassis = clean_val(row.get("Chassis"))
        slot = int(row.get("Slot")) if pd.notna(row.get("Slot")) else 0
        point = int(row.get("Point")) if pd.notna(row.get("Point")) else 0
        channel = clean_val(row.get("Column7"))
        terminal = clean_val(row.get("Terminal"))
        card = clean_val(row.get("Card"))
        io_type = clean_val(row.get("I/O Type"))
        signal = clean_val(row.get("Signal Type2"))
        inst_tag = clean_val(row.get("Instruement Tag"))
        is_spare = inst_tag.lower() == "spare" or "spare" in inst_tag.lower()
        
        # Resolve description cleanly
        desc = clean_val(row.get("Instruement_Description"))
        ctrl_desc = clean_val(row.get("Control Description"))
        if desc == "-" or not desc:
            desc = ctrl_desc if ctrl_desc != "-" else ("Spare Channel" if is_spare else "-")
        if is_spare and (ctrl_desc == "-" or not ctrl_desc):
            ctrl_desc = "Unused Spare Point for Future Expansion"
            
        plc_tag = clean_val(row.get("PLC_Tag"))
        pid = clean_val(row.get("P&ID No. 3.5"))
        if pid == "-":
            pid = clean_val(row.get("P&ID No. 3.3"))
            
        dest = clean_val(row.get("Destination"))
        jb_loc = clean_val(row.get("Junction Box Location"))
        zone = clean_val(row.get("Zone"))
        floor = clean_val(row.get("Floor"))
        panel = clean_val(row.get(" Reference Designation"))
        status = "SPARE" if is_spare else "ACTIVE"
        req_io = clean_val(row.get("Require IO"))

        row_data = [
            seq, chassis, slot, point, channel, terminal, card, io_type, signal,
            inst_tag, desc, ctrl_desc, plc_tag, pid, dest, jb_loc, zone, floor,
            panel, status, req_io
        ]

        fill = ZEBRA_ODD if (row_idx % 2) == 1 else ZEBRA_EVEN
        for c_idx, val in enumerate(row_data, 1):
            cell = ws.cell(row_idx, c_idx)
            cell_font = Font(name="Calibri", size=10, color="718096" if is_spare and c_idx not in [8, 20] else "1E293B")
            
            if c_idx in [1, 2, 3, 4, 5, 6, 18]:
                f_bold = (c_idx in [2, 3, 5] and not is_spare)
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=f_bold, color="1B365D" if f_bold else cell_font.color),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 8: # I/O Type badge
                if io_type in TYPE_STYLES:
                    set_cell(cell, val, font=TYPE_STYLES[io_type]["font"], fill=TYPE_STYLES[io_type]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 10: # Instrument Tag
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=(not is_spare), color="1E293B" if not is_spare else "718096"),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)
            elif c_idx == 19: # Panel Reference Designation
                set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT, is_text=True)
            elif c_idx == 20: # Status badge
                if status in STATUS_STYLES:
                    set_cell(cell, val, font=STATUS_STYLES[status]["font"], fill=STATUS_STYLES[status]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            else:
                set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)

        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

    # Freeze header & filters
    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:U{row_idx-1}"

    # Column widths
    col_widths = {
        "A": 9,   # Seq
        "B": 9,   # Chassis
        "C": 8,   # Slot
        "D": 8,   # Point
        "E": 14,  # Channel Code
        "F": 15,  # Terminal Block
        "G": 16,  # Module / Card
        "H": 10,  # I/O Type
        "I": 16,  # Signal Standard
        "J": 18,  # Instrument Tag
        "K": 34,  # Instrument Description
        "L": 38,  # Control Description
        "M": 22,  # PLC Tag
        "N": 20,  # P&ID No.
        "O": 16,  # Destination JB
        "P": 28,  # JB Location
        "Q": 18,  # Zone
        "R": 10,  # Floor
        "S": 18,  # Panel Enclosure
        "T": 11,  # Status
        "U": 28   # Required IO
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w


def populate_raw_sorted_sheet(ws, df_sorted):
    """Preserves all 42 original columns sorted by Chassis, Slot, Point."""
    ws.views.sheetView[0].showGridLines = True
    
    # Original column order (excluding temporary sort keys)
    raw_cols = [c for c in df_sorted.columns if not c.startswith("_")]
    
    ws.row_dimensions[1].height = 26
    for c_idx, col_name in enumerate(raw_cols, 1):
        set_cell(ws.cell(1, c_idx), col_name, font=HEADER_FONT, fill=NAVY_HEADER_FILL,
                 border=HEADER_BORDER, alignment=ALIGN_LEFT, is_text=True)

    for r_idx, (_, row) in enumerate(df_sorted.iterrows(), 2):
        fill = ZEBRA_ODD if (r_idx % 2) == 1 else ZEBRA_EVEN
        for c_idx, col_name in enumerate(raw_cols, 1):
            val = row[col_name]
            if pd.isna(val):
                cell_val = ""
            elif isinstance(val, float) and val.is_integer():
                cell_val = int(val)
            else:
                cell_val = val
            set_cell(ws.cell(r_idx, c_idx), cell_val, font=Font(name="Calibri", size=10),
                     fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT,
                     is_text=(str(cell_val).startswith("=")))
        ws.row_dimensions[r_idx].height = 19

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(raw_cols))}{len(df_sorted)+1}"
    
    for c_idx in range(1, len(raw_cols) + 1):
        col_let = get_column_letter(c_idx)
        ws.column_dimensions[col_let].width = 16


def populate_chassis_sheet(ws, df_chassis, chassis_name):
    """Dedicated sheet for each individual chassis."""
    ws.views.sheetView[0].showGridLines = True
    
    # Header Banner
    ws.merge_cells("A1:N1")
    title_cell = ws["A1"]
    ref_des = df_chassis[" Reference Designation"].dropna().iloc[0] if len(df_chassis) > 0 else "-"
    set_cell(title_cell, f"CHASSIS {chassis_name}  |  PANEL / ENCLOSURE: {ref_des}  |  PROJECT JET COOKER",
             font=Font(name="Calibri", size=12, bold=True, color="FFFFFF"), fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    headers = [
        "Item", "Slot", "Point", "Channel Code", "Terminal", "Card Model",
        "I/O Type", "Signal", "Instrument Tag", "Instrument Description",
        "PLC Tag", "P&ID No.", "Destination JB", "Status"
    ]
    header_row = 3
    ws.row_dimensions[header_row].height = 26
    for c_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(header_row, c_idx), h, font=HEADER_FONT, fill=SUB_HEADER_FILL,
                 border=HEADER_BORDER, alignment=ALIGN_CENTER if c_idx in [1, 2, 3, 4, 5, 7, 14] else ALIGN_LEFT)

    row_idx = 4
    for seq, (_, row) in enumerate(df_chassis.iterrows(), 1):
        slot = int(row.get("Slot")) if pd.notna(row.get("Slot")) else 0
        point = int(row.get("Point")) if pd.notna(row.get("Point")) else 0
        channel = clean_val(row.get("Column7"))
        terminal = clean_val(row.get("Terminal"))
        card = clean_val(row.get("Card"))
        io_type = clean_val(row.get("I/O Type"))
        signal = clean_val(row.get("Signal Type2"))
        inst_tag = clean_val(row.get("Instruement Tag"))
        is_spare = inst_tag.lower() == "spare" or "spare" in inst_tag.lower()
        
        desc = clean_val(row.get("Instruement_Description"))
        if desc == "-" or not desc:
            desc = clean_val(row.get("Control Description"))
        if is_spare and desc == "-":
            desc = "Spare I/O Point"
            
        plc_tag = clean_val(row.get("PLC_Tag"))
        pid = clean_val(row.get("P&ID No. 3.5"))
        if pid == "-":
            pid = clean_val(row.get("P&ID No. 3.3"))
        dest = clean_val(row.get("Destination"))
        status = "SPARE" if is_spare else "ACTIVE"

        row_vals = [
            seq, slot, point, channel, terminal, card, io_type, signal,
            inst_tag, desc, plc_tag, pid, dest, status
        ]
        
        fill = ZEBRA_ODD if (row_idx % 2) == 1 else ZEBRA_EVEN
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row_idx, c_idx)
            cell_font = Font(name="Calibri", size=10, color="718096" if is_spare and c_idx not in [7, 14] else "1E293B")
            
            if c_idx in [1, 2, 3, 4, 5]:
                f_bold = (c_idx in [2, 4] and not is_spare)
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=f_bold, color="1B365D" if f_bold else cell_font.color),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 7: # IO Type
                if io_type in TYPE_STYLES:
                    set_cell(cell, val, font=TYPE_STYLES[io_type]["font"], fill=TYPE_STYLES[io_type]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 9: # Instrument Tag
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=(not is_spare), color="1E293B" if not is_spare else "718096"),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)
            elif c_idx == 14: # Status
                if status in STATUS_STYLES:
                    set_cell(cell, val, font=STATUS_STYLES[status]["font"], fill=STATUS_STYLES[status]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            else:
                set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)
                
        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:N{row_idx-1}"
    
    col_widths = {
        "A": 8, "B": 8, "C": 8, "D": 14, "E": 14, "F": 16,
        "G": 10, "H": 16, "I": 18, "J": 36, "K": 22, "L": 20, "M": 16, "N": 11
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w


def generate_standalone_workbook(df_sorted):
    print(f"Creating standalone formatted workbook: {OUTPUT_STANDALONE}")
    wb = openpyxl.Workbook()
    
    # Sheet 1: Summary Dashboard
    ws_sum = wb.active
    ws_sum.title = "Chassis_Slot_Summary"
    populate_summary_sheet(ws_sum, df_sorted)

    # Sheet 2: Master Sorted IO List
    ws_master = wb.create_sheet(title="IO_List_Master_Sorted")
    populate_master_sorted_sheet(ws_master, df_sorted)

    # Sheet 3: Original 42 Columns Sorted
    ws_raw = wb.create_sheet(title="IO_List_Raw_Columns_Sorted")
    populate_raw_sorted_sheet(ws_raw, df_sorted)

    # Sheets 4-10: Per Chassis Tabs
    for ch_name in ["C1", "C2", "C3", "C4", "C5", "C6", "C7"]:
        df_ch = df_sorted[df_sorted["Chassis"] == ch_name].copy()
        if len(df_ch) > 0:
            ws_ch = wb.create_sheet(title=f"Chassis_{ch_name}")
            populate_chassis_sheet(ws_ch, df_ch, ch_name)

    wb.save(OUTPUT_STANDALONE)
    print(f"Successfully saved: {OUTPUT_STANDALONE} (Sheets: {wb.sheetnames})")


def generate_next_revision_master(df_sorted):
    """Creates IO_List_xDev-R01-Tag35-6.xlsx with sorted 'IO List' sheet."""
    print(f"Creating next revision master workbook: {OUTPUT_NEXT_REV}")
    
    wb_src = openpyxl.load_workbook(SOURCE_EXCEL)
    
    if "IO List" in wb_src.sheetnames:
        idx = wb_src.sheetnames.index("IO List")
        wb_src.remove(wb_src["IO List"])
        ws_io = wb_src.create_sheet(title="IO List", index=idx)
    else:
        ws_io = wb_src.create_sheet(title="IO List")

    populate_raw_sorted_sheet(ws_io, df_sorted)
    
    wb_src.save(OUTPUT_NEXT_REV)
    print(f"Successfully saved next revision: {OUTPUT_NEXT_REV}")


def main():
    df_sorted = build_sorted_dataset()
    generate_standalone_workbook(df_sorted)
    generate_next_revision_master(df_sorted)
    print("All tasks completed successfully!")

if __name__ == "__main__":
    main()
