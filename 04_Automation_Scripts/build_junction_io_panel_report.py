#!/usr/bin/env python3
"""
build_junction_io_panel_report.py
--------------------------------------------------------------------------------
Generates the comprehensive Junction I/O List and Panel [P1-P4] Allocation
Excel Report for Kalasin Spray Dryer Automation (xCIP-1545 / Ref: x2608003).

Features:
  1. Master Executive Summary Sheet with interactive KPI cards and hyperlinked directory.
  2. Panel [P1-P4] Allocation Matrix cross-tabulating signals from each panel to each JB.
  3. 16 Dedicated sheets for each field Junction Box & Enclosure (JB-401..JB-618, IS-JBs, CA1, RIO-200, MCC).
  4. Per-junction Panel Source Allocation Box detailing exact points and signal types taken from P1-P4.
  5. Full point-by-point schedule with source panel, marshaling terminal (Px-TBxxx-m), PLC chassis/slot/pin,
     wire tags, instrument tag/desc, I/O type, signal spec, and active/spare status.
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
PROCESS_ROOT = "/Users/x92120/xApp-001/x260911-002-Kalasin-Process"

SOURCE_IO_LIST = os.path.join(PROJECT_ROOT, "03_IO_Lists_and_Schedules", "IO_List_xDev-R02.xlsx")
SOURCE_SLOT_CFG = os.path.join(PROJECT_ROOT, "03_IO_Lists_and_Schedules", "IO_List-By_SlotConfig.xlsx")

OUT_DIR_PROCESS = os.path.join(PROCESS_ROOT, "261005-IO_Config")
OUT_DIR_ENG = os.path.join(PROJECT_ROOT, "261005-IO_Config")

os.makedirs(OUT_DIR_PROCESS, exist_ok=True)
os.makedirs(OUT_DIR_ENG, exist_ok=True)

MASTER_EXCEL_PATH = os.path.join(OUT_DIR_ENG, "Junction_IO_List_Panel_P1_P4_Allocation_Report.xlsx")
SUMMARY_EXCEL_PATH = os.path.join(OUT_DIR_ENG, "Junction_to_Panel_P1_P4_Summary_Matrix.xlsx")

# Theme Palette (Light Executive Corporate Standard)
NAVY_HEADER = "1E293B"       # Slate 800
NAVY_SUBHEADER = "334155"    # Slate 700
WHITE = "FFFFFF"
BORDER_COLOR = "CBD5E1"      # Slate 300
LIGHT_BG_1 = "FFFFFF"
LIGHT_BG_2 = "F8FAFC"        # Slate 50
ACCENT_BLUE = "0284C7"       # Sky 600

# Panel Distinct Colors
PANEL_COLORS = {
    'P1': {'fill': 'DBEAFE', 'font': '1E3A8A', 'desc': 'Panel P1 (Chassis C1 - Main Controller)'},
    'P2': {'fill': 'FEF3C7', 'font': '78350F', 'desc': 'Panel P2 (Chassis C2 - High-Density Expansion)'},
    'P3': {'fill': 'EDE9FE', 'font': '5B21B6', 'desc': 'Panel P3 (Chassis C3 - Remote I/O Field Drop)'},
    'P4': {'fill': 'FFE4E6', 'font': '9F1239', 'desc': 'Panel P4 (Chassis C4 - Field Drop & IS Barriers)'},
    'RIO-200': {'fill': 'E0E7FF', 'font': '3730A3', 'desc': 'RIO-200 (Chassis C5 - Slurry Outbuilding)'},
    'MCC': {'fill': 'F1F5F9', 'font': '334155', 'desc': 'MCC Room (Chassis C6/C7 Switchgear)'},
    'UNKNOWN': {'fill': 'F1F5F9', 'font': '64748B', 'desc': 'Direct / Auxiliary'}
}

# Status Fills
ACTIVE_FILL = "DCFCE7"       # Emerald 100
ACTIVE_FONT = "14532D"       # Emerald 900
SPARE_FILL = "FEF9C3"        # Amber 100
SPARE_FONT = "854D0E"        # Amber 900

DEST_METADATA = {
    'JB-401': {
        'desc': 'Infeed Area 2nd Floor Junction Box',
        'location': 'Infeed Processing 2nd Floor',
        'size': '800 x 600 x 250 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Slurry Infeed & Pre-Treatment',
        'primary_panel': 'Panel P1 (100%)',
        'trunk_cable': '1x 24C Control + 1x 12Pr Analog to Panel M1/P1'
    },
    'JB-402': {
        'desc': 'Jet Cooker 2nd Floor Junction Box',
        'location': 'Jet Cooker Processing 2F',
        'size': '450 x 300 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Thermal Cooking Skid',
        'primary_panel': 'Panel P2 (98.8%) & Panel P3 (1.2%)',
        'trunk_cable': '1x 18C Control to M2/P2 + 1x 2Pr Analog to M3/P3'
    },
    'JB-601': {
        'desc': 'Spray Dryer 1st Floor Junction Box',
        'location': 'Spray Dryer Floor 1F (Base)',
        'size': '500 x 400 x 200 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Spray Dryer Base Area & Product Discharge',
        'primary_panel': 'Panel P1 (95.0%) & MCC (5.0%)',
        'trunk_cable': '1x 18C Control + 1x 12Pr Analog to M1/P1'
    },
    'JB-602': {
        'desc': 'Spray Dryer 3rd Floor Junction Box',
        'location': 'Spray Dryer Floor 3F',
        'size': '800 x 600 x 250 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Chamber Middle Section & Air Inlets',
        'primary_panel': 'Panel P2 (73.5%) & Panel P3 (26.5%)',
        'trunk_cable': '1x 24C Control to M2/P2 + 1x 16Pr Analog to M3/P3'
    },
    'JB-606': {
        'desc': 'Spray Dryer 6th Floor Junction Box',
        'location': 'Spray Dryer 6F (Air Heater Section)',
        'size': '300 x 200 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Direct Gas Burner Section & Air Heating',
        'primary_panel': 'Panel P3 (100%)',
        'trunk_cable': '1x 18C Control + 1x 4Pr Analog to Panel M3/P3'
    },
    'JB-607': {
        'desc': 'Spray Dryer 7th Floor Junction Box',
        'location': 'Spray Dryer Tower 7th Floor',
        'size': '800 x 600 x 250 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Spray Dryer High Elevation Process Deck',
        'primary_panel': 'Panel P2 (69.8%) & Panel P3 (29.8%)',
        'trunk_cable': '2x 36C Control to M2/P2 + 2x 16Pr Analog to M3/P3'
    },
    'JB-608': {
        'desc': 'Spray Dryer 8th Floor Temperature JB',
        'location': 'Spray Dryer 8F (Top Plenum)',
        'size': '300 x 200 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Top Roof Sensors & Air Exhaust',
        'primary_panel': 'Panel P3 (100%)',
        'trunk_cable': '1x 4Pr RTD Shielded Cable to Panel M3/P3'
    },
    'JB-612': {
        'desc': 'Packing Tower 2nd Floor Junction Box',
        'location': 'Packing Tower 2F (Intermediate)',
        'size': '450 x 300 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Packaging & Intermediate Exhaust Line',
        'primary_panel': 'Panel P3 (98.6%) & BUS (1.4%)',
        'trunk_cable': '1x 18C Control + 1x 12Pr Analog to Panel M3/P3'
    },
    'JB-618': {
        'desc': 'Packing Tower 8th Floor Junction Box',
        'location': 'Packing Tower 8F (Top Head)',
        'size': '450 x 300 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Packing Exhaust Cyclone & Emission Monitoring',
        'primary_panel': 'Panel P3 (90.0%) & Panel P4 (10.0%)',
        'trunk_cable': '1x 24C Control to M3/P3 + 1x 12C Control to M4/P4'
    },
    'IS-JB-603': {
        'desc': 'Spray Dryer 3rd Floor Intrinsically Safe JB',
        'location': 'Spray Dryer 3F (Chamber Access Zone)',
        'size': '300 x 200 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'Dryer Chamber Monitoring (Zone 1 / 21)',
        'primary_panel': 'Panel P4 (100% via P&F IS Barriers)',
        'trunk_cable': '1x 12Pr Ex-i Blue Shielded to Panel M4/P4 Barriers'
    },
    'IS-JB-608': {
        'desc': 'Spray Dryer 8th Floor Intrinsically Safe JB',
        'location': 'Spray Dryer 8F (Roof Atomizer Deck)',
        'size': '300 x 200 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'High Shear Atomizer Deck (Zone 0 / 20)',
        'primary_panel': 'Panel P4 (100% via P&F IS Barriers)',
        'trunk_cable': '1x 16Pr Ex-i Blue Shielded to Panel M4/P4 Barriers'
    },
    'IS-JB-612': {
        'desc': 'Packing Tower 2nd Floor Intrinsically Safe JB',
        'location': 'Packing Tower 2F (Dust Hazard Area)',
        'size': '300 x 200 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'Product Discharge Hopper (Combustible Dust)',
        'primary_panel': 'Panel P4 (100% via P&F IS Barriers)',
        'trunk_cable': '1x 12Pr Ex-i Blue Shielded to Panel M4/P4 Barriers'
    },
    'IS-JB-618': {
        'desc': 'Packing Tower 8th Floor Intrinsically Safe JB',
        'location': 'Packing Tower 8F (Hazardous Vapor Area)',
        'size': '450 x 300 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'Explosion Vent / Vapor Zone (Hazardous Area)',
        'primary_panel': 'Panel P4 (100% via P&F IS Barriers)',
        'trunk_cable': '1x 16Pr Ex-i Blue Shielded to Panel M4/P4 Barriers'
    },
    'CA1': {
        'desc': 'Main Control Panel CA1 Local Terminals',
        'location': 'Control Room Cabinet CA1',
        'size': '1600 x 2000 x 800 mm',
        'type': 'Main PLC Enclosure Internal Marshaling',
        'area': 'Central Control Room (MCP)',
        'primary_panel': 'Panel P1 (90.0%), P2/P3/P4/RIO (10.0%)',
        'trunk_cable': 'Internal Wiring Ducts / Inter-Panel Bus'
    },
    'RIO-200': {
        'desc': 'Remote I/O Enclosure RIO-200',
        'location': 'Slurry Outbuilding 2nd Floor',
        'size': '800 x 1000 x 400 mm',
        'type': 'Remote I/O Skid Cabinet (DLR Fiber Node)',
        'area': 'Remote Slurry Skid Preparation',
        'primary_panel': 'RIO-200 (Chassis C5 Fiber Drop)',
        'trunk_cable': '165m 6-Core Steel Wire Armoured Fiber Optic Trunk'
    },
    'MCC': {
        'desc': 'Motor Control Center (MCC Room)',
        'location': 'MCC Room Ground Floor',
        'size': 'Custom Switchgear Lineup',
        'type': 'MCC Switchgear & Bus Interface',
        'area': 'MCC Substation',
        'primary_panel': 'MCC Interface (Chassis C6/C7)',
        'trunk_cable': 'Hardwired Control Multicores to Starter Buckets'
    }
}

ORDERED_JBS = [
    'JB-401', 'JB-402', 'JB-601', 'JB-602', 'JB-606', 'JB-607', 'JB-608',
    'JB-612', 'JB-618', 'IS-JB-603', 'IS-JB-608', 'IS-JB-612', 'IS-JB-618',
    'CA1', 'RIO-200', 'MCC'
]

def load_data():
    wb = openpyxl.load_workbook(SOURCE_IO_LIST, data_only=True)
    ws = wb['IO List']
    rows = list(ws.iter_rows(values_only=True))

    data_by_jb = {jb: [] for jb in ORDERED_JBS}

    for r in rows[1:]:
        dest = str(r[1]).strip() if r[1] else ''
        matched_jb = None
        # Match IS-JBs first to avoid collision with JB-6xx
        for jb in ['IS-JB-603', 'IS-JB-608', 'IS-JB-612', 'IS-JB-618']:
            if jb in dest:
                matched_jb = jb
                break
        if not matched_jb:
            for jb in ORDERED_JBS:
                if jb in dest:
                    matched_jb = jb
                    break
        if not matched_jb and ('CA1' in dest or 'Control Room' in dest):
            matched_jb = 'CA1'

        if matched_jb:
            term_item = str(r[44]).strip() if r[44] else ''
            chassis = str(r[30]).strip() if r[30] else ''

            panel = 'UNKNOWN'
            if term_item.startswith('P1'): panel = 'P1'
            elif term_item.startswith('P2'): panel = 'P2'
            elif term_item.startswith('P3'): panel = 'P3'
            elif term_item.startswith('P4'): panel = 'P4'
            elif chassis == 'C1': panel = 'P1'
            elif chassis == 'C2': panel = 'P2'
            elif chassis == 'C3': panel = 'P3'
            elif chassis == 'C4': panel = 'P4'
            elif chassis == 'C5': panel = 'RIO-200'
            elif chassis in ['C6', 'C7']: panel = 'MCC'

            inst_tag = str(r[11]).strip() if r[11] is not None else ''
            inst_desc_raw = str(r[12]).strip() if r[12] is not None else ''
            ctrl_desc_raw = str(r[16]).strip() if r[16] is not None else ''

            if inst_desc_raw and inst_desc_raw not in ['0', 'None', '-']:
                final_desc = inst_desc_raw
            elif ctrl_desc_raw and ctrl_desc_raw not in ['0', 'None', '-']:
                final_desc = ctrl_desc_raw
            else:
                final_desc = ''

            is_spare = (
                'spare' in inst_tag.lower() or 
                'reserve' in inst_desc_raw.lower() or 
                'reserve' in ctrl_desc_raw.lower() or 
                inst_tag in ['', '-', 'None', '0'] or
                'spare' in str(r[19] or '').lower()
            )
            status = 'SPARE' if is_spare else 'ACTIVE'

            if is_spare:
                if not final_desc or final_desc in ['0', 'None', '-']:
                    final_desc = 'Spare Reserve Channel'
                if not inst_tag or inst_tag in ['-', '0']:
                    inst_tag = 'Spare Reserve'
            else:
                if not final_desc:
                    final_desc = 'Field Signal Channel'

            data_by_jb[matched_jb].append({
                'panel': panel,
                'term_item': term_item,
                'chassis': chassis,
                'slot': r[31],
                'point': r[32],
                'card': r[26],
                'plc_side': r[45],
                'term_side': r[46],
                'inst_tag': inst_tag,
                'inst_desc': final_desc,
                'io_type': r[20] if r[20] else 'OTHER',
                'sig_type': r[27] if r[27] else '24VDC / 4-20mA',
                'plc_tag': r[19] if r[19] else ('Spare' if is_spare else '-'),
                'pid': r[14] if r[14] else '-',
                'status': status
            })

    return data_by_jb

def apply_thin_border(cell, color=BORDER_COLOR):
    side = Side(style='thin', color=color)
    cell.border = Border(left=side, right=side, top=side, bottom=side)

def build_master_workbook(data_by_jb):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 1. SHEET: 00_Executive_Summary
    ws_sum = wb.create_sheet(title="00_Executive_Summary")
    ws_sum.views.sheetView[0].showGridLines = True

    # Title Banner
    ws_sum.merge_cells("A1:N1")
    title_cell = ws_sum["A1"]
    title_cell.value = "KALASIN SPRAY DRYER AUTOMATION — JUNCTION BOX I/O & PANEL [P1-P4] CONFIGURATION"
    title_cell.font = Font(name="Arial", size=15, bold=True, color=WHITE)
    title_cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:N2")
    sub_cell = ws_sum["A2"]
    sub_cell.value = "Comprehensive Cross-Reference Directory: Field Junction Boxes vs Control Marshaling Panels P1, P2, P3, and P4"
    sub_cell.font = Font(name="Arial", size=10, italic=True, color=WHITE)
    sub_cell.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    sub_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 20

    # Summary Stats
    total_pts = sum(len(pts) for pts in data_by_jb.values())
    total_act = sum(sum(1 for p in pts if p['status'] == 'ACTIVE') for pts in data_by_jb.values())
    total_spr = total_pts - total_act
    p1_pts = sum(sum(1 for p in pts if p['panel'] == 'P1') for pts in data_by_jb.values())
    p2_pts = sum(sum(1 for p in pts if p['panel'] == 'P2') for pts in data_by_jb.values())
    p3_pts = sum(sum(1 for p in pts if p['panel'] == 'P3') for pts in data_by_jb.values())
    p4_pts = sum(sum(1 for p in pts if p['panel'] == 'P4') for pts in data_by_jb.values())

    kpis = [
        ("TOTAL FIELD DESTINATIONS", "16 Enclosures", "12 Standard + 4 IS Ex-i", "DBEAFE", "1E3A8A"),
        ("TOTAL PLANT I/O CHANNELS", f"{total_pts:,} Points", f"Active: {total_act:,} | Spare: {total_spr:,}", "FEF3C7", "78350F"),
        ("PANEL P1 ALLOCATION", f"{p1_pts} Points ({p1_pts/total_pts*100:.1f}%)", "Chassis C1 (Main Controller)", "DBEAFE", "1E3A8A"),
        ("PANEL P2 ALLOCATION", f"{p2_pts} Points ({p2_pts/total_pts*100:.1f}%)", "Chassis C2 (High-Density Expansion)", "FEF3C7", "78350F"),
        ("PANEL P3 ALLOCATION", f"{p3_pts} Points ({p3_pts/total_pts*100:.1f}%)", "Chassis C3 (Remote I/O Field Drop)", "EDE9FE", "5B21B6"),
        ("PANEL P4 ALLOCATION", f"{p4_pts} Points ({p4_pts/total_pts*100:.1f}%)", "Chassis C4 (Field Drop & IS Barriers)", "FFE4E6", "9F1239"),
    ]

    card_ranges = [("A4:B5"), ("C4:D5"), ("E4:F5"), ("G4:H5"), ("I4:J5"), ("K4:N5")]
    for (title, val, sub, bg, fg), crange in zip(kpis, card_ranges):
        start_c, end_c = crange.split(":")
        ws_sum.merge_cells(crange)
        c = ws_sum[start_c]
        c.value = f"{title}\n{val}\n{sub}"
        c.font = Font(name="Arial", size=9, bold=True, color=fg)
        c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        apply_thin_border(c, color=BORDER_COLOR)

    ws_sum.row_dimensions[4].height = 24
    ws_sum.row_dimensions[5].height = 24

    headers = [
        "Dest ID", "Enclosure Description", "Plant Location / Floor", "Dimensions & Enclosure Spec",
        "From P1", "From P2", "From P3", "From P4", "Total Pts", "Active", "Spare", "% Spare",
        "Primary Panel Source", "Direct Sheet Link"
    ]
    ws_sum.row_dimensions[7].height = 28
    for col_idx, h in enumerate(headers, start=1):
        cell = ws_sum.cell(row=7, column=col_idx, value=h)
        cell.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        apply_thin_border(cell)

    for row_idx, jb in enumerate(ORDERED_JBS, start=8):
        pts = data_by_jb[jb]
        meta = DEST_METADATA.get(jb, {})
        tot = len(pts)
        act = sum(1 for p in pts if p['status'] == 'ACTIVE')
        spr = tot - act
        pct_spr = (spr / tot * 100) if tot > 0 else 0

        p1_cnt = sum(1 for p in pts if p['panel'] == 'P1')
        p2_cnt = sum(1 for p in pts if p['panel'] == 'P2')
        p3_cnt = sum(1 for p in pts if p['panel'] == 'P3')
        p4_cnt = sum(1 for p in pts if p['panel'] == 'P4')

        row_data = [
            jb, meta.get('desc', jb), meta.get('location', '-'), f"{meta.get('size', '-')} ({meta.get('type', '-')})",
            p1_cnt, p2_cnt, p3_cnt, p4_cnt, tot, act, spr, f"{pct_spr:.1f}%",
            meta.get('primary_panel', '-'), f"View {jb} Schedule"
        ]

        bg_color = LIGHT_BG_1 if row_idx % 2 == 0 else LIGHT_BG_2
        ws_sum.row_dimensions[row_idx].height = 20

        for col_idx, val in enumerate(row_data, start=1):
            cell = ws_sum.cell(row=row_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            apply_thin_border(cell)

            if col_idx == 1:
                cell.font = Font(name="Arial", size=9, bold=True, color=ACCENT_BLUE)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [5, 6, 7, 8, 9, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx in [5, 6, 7, 8] and val > 0:
                    p_key = f"P{col_idx-4}"
                    cell.fill = PatternFill(start_color=PANEL_COLORS[p_key]['fill'], end_color=PANEL_COLORS[p_key]['fill'], fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=True, color=PANEL_COLORS[p_key]['font'])
            elif col_idx == 12:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 14:
                cell.font = Font(name="Arial", size=9, bold=True, underline="single", color=ACCENT_BLUE)
                cell.hyperlink = f"#'{jb}'!A1"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # Total Row
    tot_row = 8 + len(ORDERED_JBS)
    ws_sum.row_dimensions[tot_row].height = 24
    ws_sum.cell(row=tot_row, column=1, value="TOTALS").font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=2, value="All 16 Field Junction Boxes & Enclosures").font = Font(name="Arial", size=9, italic=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=3, value="-").font = Font(name="Arial", size=9, color=WHITE)
    ws_sum.cell(row=tot_row, column=4, value="-").font = Font(name="Arial", size=9, color=WHITE)
    ws_sum.cell(row=tot_row, column=5, value=p1_pts).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=6, value=p2_pts).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=7, value=p3_pts).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=8, value=p4_pts).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=9, value=total_pts).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=10, value=total_act).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=11, value=total_spr).font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=12, value=f"{total_spr/total_pts*100:.1f}%").font = Font(name="Arial", size=10, bold=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=13, value="Complete Plant Scope").font = Font(name="Arial", size=9, italic=True, color=WHITE)
    ws_sum.cell(row=tot_row, column=14, value="-").font = Font(name="Arial", size=9, color=WHITE)

    for c_idx in range(1, 15):
        cell = ws_sum.cell(row=tot_row, column=c_idx)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        apply_thin_border(cell)
        if c_idx in [5, 6, 7, 8, 9, 10, 11, 12]:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    for col in ws_sum.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws_sum.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws_sum.column_dimensions['A'].width = 14
    ws_sum.column_dimensions['B'].width = 38
    ws_sum.column_dimensions['C'].width = 32
    ws_sum.column_dimensions['D'].width = 44
    ws_sum.column_dimensions['M'].width = 35
    ws_sum.column_dimensions['N'].width = 18

    # 2. SHEET: 01_Panel_Allocation_Matrix
    ws_mat = wb.create_sheet(title="01_Panel_Allocation_Matrix")
    ws_mat.views.sheetView[0].showGridLines = True

    ws_mat.merge_cells("A1:R1")
    m_title = ws_mat["A1"]
    m_title.value = "DETAILED JUNCTION BOX TO PANEL [P1-P4] SIGNAL ALLOCATION MATRIX"
    m_title.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    m_title.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    m_title.alignment = Alignment(horizontal="center", vertical="center")
    ws_mat.row_dimensions[1].height = 32

    ws_mat.merge_cells("A2:R2")
    m_sub = ws_mat["A2"]
    m_sub.value = "Granular Signal Type Breakdown (DI, DO, AI, AO) Delivered by Panels P1, P2, P3, and P4 to each Field Junction Box"
    m_sub.font = Font(name="Arial", size=10, italic=True, color=WHITE)
    m_sub.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    m_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws_mat.row_dimensions[2].height = 20

    super_hdrs = [
        ("A4:C4", "DESTINATION JUNCTION BOX", NAVY_HEADER),
        ("D4:G4", "PANEL P1 (CHASSIS C1)", "1E3A8A"),
        ("H4:K4", "PANEL P2 (CHASSIS C2)", "78350F"),
        ("L4:O4", "PANEL P3 (CHASSIS C3)", "5B21B6"),
        ("P4:R4", "PANEL P4 (CHASSIS C4 / IS)", "9F1239"),
    ]
    ws_mat.row_dimensions[4].height = 22
    for rng, htext, color in super_hdrs:
        ws_mat.merge_cells(rng)
        sc = ws_mat[rng.split(":")[0]]
        sc.value = htext
        sc.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        sc.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        sc.alignment = Alignment(horizontal="center", vertical="center")

    mat_cols = [
        "Dest ID", "Location", "Area Description",
        "P1 DI", "P1 DO", "P1 AI", "P1 Total",
        "P2 DI", "P2 DO", "P2 AI", "P2 Total",
        "P3 DI", "P3 DO", "P3 AI", "P3 Total",
        "P4 DI", "P4 DO/AI", "P4 Total"
    ]
    ws_mat.row_dimensions[5].height = 24
    for c_idx, h in enumerate(mat_cols, start=1):
        cell = ws_mat.cell(row=5, column=c_idx, value=h)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        apply_thin_border(cell)

    for r_idx, jb in enumerate(ORDERED_JBS, start=6):
        pts = data_by_jb[jb]
        meta = DEST_METADATA.get(jb, {})

        def count_sig(p_name, s_type=None):
            if s_type:
                return sum(1 for p in pts if p['panel'] == p_name and p['io_type'] == s_type)
            return sum(1 for p in pts if p['panel'] == p_name)

        p1_di = count_sig('P1', 'DI')
        p1_do = count_sig('P1', 'DO')
        p1_ai = count_sig('P1', 'AI')
        p1_tot = count_sig('P1')

        p2_di = count_sig('P2', 'DI')
        p2_do = count_sig('P2', 'DO')
        p2_ai = count_sig('P2', 'AI')
        p2_tot = count_sig('P2')

        p3_di = count_sig('P3', 'DI')
        p3_do = count_sig('P3', 'DO')
        p3_ai = count_sig('P3', 'AI') + count_sig('P3', 'AO')
        p3_tot = count_sig('P3')

        p4_di = count_sig('P4', 'DI')
        p4_other = count_sig('P4', 'DO') + count_sig('P4', 'AI') + count_sig('P4', 'AO')
        p4_tot = count_sig('P4')

        row_vals = [
            jb, meta.get('location', '-'), meta.get('area', '-'),
            p1_di or '-', p1_do or '-', p1_ai or '-', p1_tot or '-',
            p2_di or '-', p2_do or '-', p2_ai or '-', p2_tot or '-',
            p3_di or '-', p3_do or '-', p3_ai or '-', p3_tot or '-',
            p4_di or '-', p4_other or '-', p4_tot or '-'
        ]

        bg_color = LIGHT_BG_1 if r_idx % 2 == 0 else LIGHT_BG_2
        ws_mat.row_dimensions[r_idx].height = 20

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws_mat.cell(row=r_idx, column=col_idx, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
            apply_thin_border(cell)

            if col_idx == 1:
                cell.font = Font(name="Arial", size=9, bold=True, color=ACCENT_BLUE)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [2, 3]:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx in [7, 11, 15, 18] and val != '-':
                    cell.font = Font(name="Arial", size=9, bold=True)

    for col in ws_mat.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_mat.column_dimensions[col_letter].width = max(max_len + 3, 11)
    ws_mat.column_dimensions['A'].width = 14
    ws_mat.column_dimensions['B'].width = 30
    ws_mat.column_dimensions['C'].width = 36

    # 3. DEDICATED SHEETS FOR EACH JUNCTION BOX
    for jb in ORDERED_JBS:
        pts = data_by_jb[jb]
        meta = DEST_METADATA.get(jb, {})

        ws_jb = wb.create_sheet(title=jb)
        ws_jb.views.sheetView[0].showGridLines = True

        ws_jb.merge_cells("A1:N1")
        h1 = ws_jb["A1"]
        h1.value = f"JUNCTION BOX I/O TERMINATION SCHEDULE --- {jb} ({meta.get('desc', jb).upper()})"
        h1.font = Font(name="Arial", size=13, bold=True, color=WHITE)
        h1.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        h1.alignment = Alignment(horizontal="left", vertical="center")
        ws_jb.row_dimensions[1].height = 30

        ws_jb.merge_cells("A2:N2")
        h2 = ws_jb["A2"]
        h2.value = f"Location: {meta.get('location', '-')}  |  Enclosure: {meta.get('size', '-')} ({meta.get('type', '-')})  |  Trunk Cable: {meta.get('trunk_cable', '-')}"
        h2.font = Font(name="Arial", size=9, italic=True, color=WHITE)
        h2.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        h2.alignment = Alignment(horizontal="left", vertical="center")
        ws_jb.row_dimensions[2].height = 20

        ws_jb["O1"].value = "⬅ Back to Index"
        ws_jb["O1"].hyperlink = "#'00_Executive_Summary'!A1"
        ws_jb["O1"].font = Font(name="Arial", size=9, bold=True, color=WHITE, underline="single")
        ws_jb["O1"].fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        ws_jb["O1"].alignment = Alignment(horizontal="center", vertical="center")

        ws_jb["O2"].value = "⬅ Panel Matrix"
        ws_jb["O2"].hyperlink = "#'01_Panel_Allocation_Matrix'!A1"
        ws_jb["O2"].font = Font(name="Arial", size=9, bold=True, color=WHITE, underline="single")
        ws_jb["O2"].fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        ws_jb["O2"].alignment = Alignment(horizontal="center", vertical="center")

        p1_c = sum(1 for p in pts if p['panel'] == 'P1')
        p2_c = sum(1 for p in pts if p['panel'] == 'P2')
        p3_c = sum(1 for p in pts if p['panel'] == 'P3')
        p4_c = sum(1 for p in pts if p['panel'] == 'P4')
        act_c = sum(1 for p in pts if p['status'] == 'ACTIVE')
        spr_c = len(pts) - act_c

        p1_di = sum(1 for p in pts if p['panel']=='P1' and p['io_type']=='DI')
        p1_do = sum(1 for p in pts if p['panel']=='P1' and p['io_type']=='DO')
        p1_ai = sum(1 for p in pts if p['panel']=='P1' and p['io_type']=='AI')
        kpi_p1 = f"FROM PANEL P1\n{p1_c} Points\nDI:{p1_di} | DO:{p1_do} | AI:{p1_ai}"

        p2_di = sum(1 for p in pts if p['panel']=='P2' and p['io_type']=='DI')
        p2_do = sum(1 for p in pts if p['panel']=='P2' and p['io_type']=='DO')
        p2_ai = sum(1 for p in pts if p['panel']=='P2' and p['io_type']=='AI')
        kpi_p2 = f"FROM PANEL P2\n{p2_c} Points\nDI:{p2_di} | DO:{p2_do} | AI:{p2_ai}"

        p3_di = sum(1 for p in pts if p['panel']=='P3' and p['io_type']=='DI')
        p3_do = sum(1 for p in pts if p['panel']=='P3' and p['io_type']=='DO')
        p3_ai = sum(1 for p in pts if p['panel']=='P3' and p['io_type']=='AI')
        p3_ao = sum(1 for p in pts if p['panel']=='P3' and p['io_type']=='AO')
        p3_sub = f"DI:{p3_di} | DO:{p3_do} | AI:{p3_ai}" + (f" | AO:{p3_ao}" if p3_ao > 0 else "")
        kpi_p3 = f"FROM PANEL P3\n{p3_c} Points\n{p3_sub}"

        p4_di = sum(1 for p in pts if p['panel']=='P4' and p['io_type']=='DI')
        p4_do = sum(1 for p in pts if p['panel']=='P4' and p['io_type']=='DO')
        p4_ai = sum(1 for p in pts if p['panel']=='P4' and p['io_type']=='AI')
        p4_ao = sum(1 for p in pts if p['panel']=='P4' and p['io_type']=='AO')
        p4_sub = f"DI:{p4_di} | DO:{p4_do} | AI:{p4_ai}" + (f" | AO:{p4_ao}" if p4_ao > 0 else "")
        kpi_p4 = f"FROM PANEL P4\n{p4_c} Points\n{p4_sub}"

        kpi_tot = f"TOTAL CAPACITY\n{len(pts)} Channels\nActive: {act_c} | Spare: {spr_c} ({spr_c/len(pts)*100:.1f}%)" if len(pts)>0 else "0 Channels"

        jb_kpis = [
            ("A4:C5", kpi_p1, PANEL_COLORS['P1']['fill'], PANEL_COLORS['P1']['font']),
            ("D4:F5", kpi_p2, PANEL_COLORS['P2']['fill'], PANEL_COLORS['P2']['font']),
            ("G4:I5", kpi_p3, PANEL_COLORS['P3']['fill'], PANEL_COLORS['P3']['font']),
            ("J4:L5", kpi_p4, PANEL_COLORS['P4']['fill'], PANEL_COLORS['P4']['font']),
            ("M4:O5", kpi_tot, "DCFCE7", "14532D")
        ]
        ws_jb.row_dimensions[4].height = 20
        ws_jb.row_dimensions[5].height = 20
        for rng, txt, bg, fg in jb_kpis:
            ws_jb.merge_cells(rng)
            c = ws_jb[rng.split(":")[0]]
            c.value = txt
            c.font = Font(name="Arial", size=9, bold=True, color=fg)
            c.fill = PatternFill(start_color=bg, end_color=bg, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            apply_thin_border(c)

        jb_hdrs = [
            "Point #", "Source Panel", "Panel Terminal Block", "Chassis", "Slot",
            "Module Model", "Wire Tag PLC Side", "Wire Tag Terminal Side",
            "Instrument Tag", "Instrument Description", "I/O Type", "Signal Specification",
            "PLC Tag Name", "Channel Status", "P&ID Ref"
        ]
        ws_jb.row_dimensions[7].height = 26
        for c_idx, h in enumerate(jb_hdrs, start=1):
            cell = ws_jb.cell(row=7, column=c_idx, value=h)
            cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            apply_thin_border(cell)

        for pt_idx, pt in enumerate(pts, start=1):
            cur_r = 7 + pt_idx
            ws_jb.row_dimensions[cur_r].height = 19
            bg_col = LIGHT_BG_1 if pt_idx % 2 == 0 else LIGHT_BG_2

            row_data = [
                pt_idx,
                pt['panel'],
                pt['term_item'] if pt['term_item'] else '-',
                pt['chassis'],
                pt['slot'],
                pt['card'],
                pt['plc_side'] if pt['plc_side'] else '-',
                pt['term_side'] if pt['term_side'] else '-',
                pt['inst_tag'],
                pt['inst_desc'],
                pt['io_type'],
                pt['sig_type'],
                pt['plc_tag'],
                pt['status'],
                pt['pid']
            ]

            for c_idx, val in enumerate(row_data, start=1):
                cell = ws_jb.cell(row=cur_r, column=c_idx, value=val)
                cell.font = Font(name="Arial", size=9)
                cell.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type="solid")
                apply_thin_border(cell)

                if c_idx == 1:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Arial", size=8, color="64748B")
                elif c_idx == 2:
                    p_info = PANEL_COLORS.get(val, PANEL_COLORS['UNKNOWN'])
                    cell.fill = PatternFill(start_color=p_info['fill'], end_color=p_info['fill'], fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=True, color=p_info['font'])
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 3:
                    cell.font = Font(name="Consolas", size=9, bold=True, color="0369A1")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [4, 5]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx in [7, 8]:
                    cell.font = Font(name="Consolas", size=8, color="334155")
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c_idx == 9:
                    cell.font = Font(name="Arial", size=9, bold=True, color="0F172A")
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c_idx == 11:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Arial", size=9, bold=True)
                elif c_idx == 14:
                    if val == 'ACTIVE':
                        cell.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                        cell.font = Font(name="Arial", size=8, bold=True, color=ACTIVE_FONT)
                    else:
                        cell.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                        cell.font = Font(name="Arial", size=8, bold=True, color=SPARE_FONT)
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == 15:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.font = Font(name="Arial", size=8, color="64748B")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

        for col in ws_jb.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = str(cell.value or '')
                if '\n' in val_str:
                    val_str = max(val_str.split('\n'), key=len)
                max_len = max(max_len, len(val_str))
            ws_jb.column_dimensions[col_letter].width = max(max_len + 3, 10)
        ws_jb.column_dimensions['A'].width = 8
        ws_jb.column_dimensions['B'].width = 14
        ws_jb.column_dimensions['C'].width = 20
        ws_jb.column_dimensions['G'].width = 24
        ws_jb.column_dimensions['H'].width = 24
        ws_jb.column_dimensions['I'].width = 18
        ws_jb.column_dimensions['J'].width = 36
        ws_jb.column_dimensions['L'].width = 24
        ws_jb.column_dimensions['M'].width = 26

    wb.save(MASTER_EXCEL_PATH)
    print(f"[OK] Master Excel Created: {MASTER_EXCEL_PATH} ({os.path.getsize(MASTER_EXCEL_PATH):,} bytes)")
    if os.path.exists(PROCESS_ROOT):
        process_target = os.path.join(OUT_DIR_PROCESS, "Junction_IO_List_Panel_P1_P4_Allocation_Report.xlsx")
        wb.save(process_target)
        print(f"[OK] Mirrored to Process: {process_target}")

def build_summary_matrix_workbook(data_by_jb):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Junction_to_Panel_Matrix"
    ws.views.sheetView[0].showGridLines = True

    ws.merge_cells("A1:K1")
    t = ws["A1"]
    t.value = "KALASIN SPRAY DRYER AUTOMATION — JUNCTION BOX TO PANEL [P1-P4] SUMMARY MATRIX"
    t.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    t.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    hdrs = [
        "Junction Box ID", "Enclosure Description", "Plant Location / Floor",
        "From Panel P1", "From Panel P2", "From Panel P3", "From Panel P4",
        "Total I/O Points", "Active Points", "Installed Spares", "% Spare Margin"
    ]
    ws.row_dimensions[3].height = 26
    for c_idx, h in enumerate(hdrs, start=1):
        c = ws.cell(row=3, column=c_idx, value=h)
        c.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center")
        apply_thin_border(c)

    for r_idx, jb in enumerate(ORDERED_JBS, start=4):
        pts = data_by_jb[jb]
        meta = DEST_METADATA.get(jb, {})
        tot = len(pts)
        act = sum(1 for p in pts if p['status'] == 'ACTIVE')
        spr = tot - act
        pct_spr = (spr / tot * 100) if tot > 0 else 0

        p1_cnt = sum(1 for p in pts if p['panel'] == 'P1')
        p2_cnt = sum(1 for p in pts if p['panel'] == 'P2')
        p3_cnt = sum(1 for p in pts if p['panel'] == 'P3')
        p4_cnt = sum(1 for p in pts if p['panel'] == 'P4')

        row_vals = [
            jb, meta.get('desc', jb), meta.get('location', '-'),
            p1_cnt, p2_cnt, p3_cnt, p4_cnt,
            tot, act, spr, f"{pct_spr:.1f}%"
        ]
        ws.row_dimensions[r_idx].height = 20
        bg_col = LIGHT_BG_1 if r_idx % 2 == 0 else LIGHT_BG_2

        for c_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type="solid")
            apply_thin_border(cell)

            if c_idx == 1:
                cell.font = Font(name="Arial", size=9, bold=True, color=ACCENT_BLUE)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c_idx in [4, 5, 6, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if val > 0:
                    p_key = f"P{c_idx-3}"
                    cell.fill = PatternFill(start_color=PANEL_COLORS[p_key]['fill'], end_color=PANEL_COLORS[p_key]['fill'], fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=True, color=PANEL_COLORS[p_key]['font'])
            elif c_idx in [8, 9, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
                if c_idx == 8:
                    cell.font = Font(name="Arial", size=9, bold=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['B'].width = 38
    ws.column_dimensions['C'].width = 32

    wb.save(SUMMARY_EXCEL_PATH)
    print(f"[OK] Summary Matrix Created: {SUMMARY_EXCEL_PATH} ({os.path.getsize(SUMMARY_EXCEL_PATH):,} bytes)")
    if os.path.exists(PROCESS_ROOT):
        process_summary = os.path.join(OUT_DIR_PROCESS, "Junction_to_Panel_P1_P4_Summary_Matrix.xlsx")
        wb.save(process_summary)
        print(f"[OK] Mirrored to Process: {process_summary}")

if __name__ == "__main__":
    print("=" * 80)
    print("BUILDING JUNCTION BOX I/O & PANEL [P1-P4] CONFIGURATION EXCEL DELIVERABLES")
    print("=" * 80)
    data = load_data()
    build_master_workbook(data)
    build_summary_matrix_workbook(data)
    print("=" * 80)
    print("SUCCESS: ALL DELIVERABLES GENERATED IN 261005-IO_Config!")
    print("=" * 80)
