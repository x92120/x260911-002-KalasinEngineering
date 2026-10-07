#!/usr/bin/env python3
"""
=============================================================================
KALASIN ENGINEERING xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)
PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER - INGREDION (THAILAND) CO., LTD.
MASTER LOOP TEST PLAN, PROCEDURE, AND COMMISSIONING SCHEDULE
=============================================================================
Source Files:
  1. 03-IO_List/Instrument I-O List Rev.3.6a.xlsx
  2. 03-IO_List/IO_List-By_SlotConfig_rev02.xlsx
  3. 03-IO_List/IO_List_By_Junction_Box.xlsx

Target Output Files (Folder: 05-LoopTest):
  - 05-LoopTest/Loop_Test_Plan_and_Schedule.xlsx
  - 05-LoopTest/Master_Loop_Test_Protocol.xlsx (Duplicate / Alternate reference)

Workbook Structure:
  1. 00_Executive_Summary        - Executive overview, testing phases, KPI dashboard & scope
  2. 01_Loop_Test_Procedure      - Detailed IEC 62381 / ISA Standard Operating Procedure (SOP)
  3. 02_Master_Testing_Schedule  - 4-Week (Day 1 to 20) Day-by-Day Gantt Commissioning Timeline
  4. 03_Loop_Test_Master_List    - 560 Loop-by-loop test sheets with cold/hot check & sign-offs
  5. 04_ITP_Inspection_Plan      - Quality Inspection & Test Plan (ITP) sign-off matrix
  6. 05_Loop_Test_Form_Template  - Field Loop Check Certificate template for physical archiving
=============================================================================
"""

import os
import re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PROJECT_ROOT = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_DIR = os.path.join(PROJECT_ROOT, "03-IO_List")
OUT_DIR = os.path.join(PROJECT_ROOT, "05-LoopTest")
os.makedirs(OUT_DIR, exist_ok=True)

INST_FILE = os.path.join(IO_DIR, "Instrument I-O List Rev.3.6a.xlsx")
SLOT_FILE = os.path.join(IO_DIR, "IO_List-By_SlotConfig_rev02.xlsx")
JB_FILE = os.path.join(IO_DIR, "IO_List_By_Junction_Box.xlsx")

OUTPUT_FILE = os.path.join(OUT_DIR, "Loop_Test_Plan_and_Schedule.xlsx")
OUTPUT_FILE_ALT = os.path.join(OUT_DIR, "Master_Loop_Test_Protocol.xlsx")

# -----------------------------------------------------------------------------
# Color Palette & Styles
# -----------------------------------------------------------------------------
NAVY_HEADER = "1B365D"       # Deep Corporate Navy
NAVY_SUBHEADER = "2D4A77"    # Slate Blue
SECTION_HEADER = "334E68"    # Steel Blue
CARD_HEADER_BG = "0F172A"    # Dark Slate
CARD_BG = "F1F5F9"           # Light Slate 100

WHITE = "FFFFFF"
BORDER_COLOR = "CBD5E1"      # Slate 300

NAVY_FILL = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
SUB_FILL = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
SECTION_FILL = PatternFill(start_color=SECTION_HEADER, end_color=SECTION_HEADER, fill_type="solid")
CARD_HEAD_FILL = PatternFill(start_color=CARD_HEADER_BG, end_color=CARD_HEADER_BG, fill_type="solid")
CARD_BODY_FILL = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")

ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

ACTIVE_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
ACTIVE_FONT = Font(name="Calibri", size=10, bold=True, color="166534")

SPARE_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
SPARE_FONT = Font(name="Calibri", size=10, bold=True, color="92400E")

PENDING_FILL = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
PENDING_FONT = Font(name="Calibri", size=10, bold=True, color="0369A1")

GANTT_ACTIVE = PatternFill(start_color="0284C7", end_color="0284C7", fill_type="solid")
GANTT_MILESTONE = PatternFill(start_color="166534", end_color="166534", fill_type="solid")
GANTT_WEEKEND = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")

FONT_TITLE = Font(name="Calibri", size=13, bold=True, color=WHITE)
FONT_SUBTITLE = Font(name="Calibri", size=9.5, italic=True, color=WHITE)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color=WHITE)
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color=WHITE)

FONT_CARD_TITLE = Font(name="Calibri", size=8.5, bold=True, color="94A3B8")
FONT_CARD_VAL = Font(name="Calibri", size=14, bold=True, color="1E293B")
FONT_CARD_SUB = Font(name="Calibri", size=8, italic=True, color="64748B")

FONT_DATA = Font(name="Calibri", size=10, color="0F172A")
FONT_DATA_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")
FONT_DATA_CODE = Font(name="Consolas", size=9.5, color="0F172A")
FONT_DATA_MUTED = Font(name="Calibri", size=9, italic=True, color="64748B")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)
HEADER_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='medium', color='CBD5E1'),
    bottom=Side(style='medium', color='CBD5E1')
)
TOTAL_BORDER = Border(
    top=Side(style='thin', color='94A3B8'),
    bottom=Side(style='double', color='0F172A'),
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0')
)

def set_cell(cell, value, font=None, fill=None, border=THIN_BORDER, alignment=ALIGN_LEFT, num_format=None):
    if not isinstance(cell, openpyxl.cell.cell.MergedCell):
        cell.value = value
    if font: cell.font = font
    if fill: cell.fill = fill
    if border: cell.border = border
    if alignment: cell.alignment = alignment
    if num_format: cell.number_format = num_format

# -----------------------------------------------------------------------------
# Data Loader: Load Instruments & Slot Mappings
# -----------------------------------------------------------------------------
def load_data():
    print(f"Loading Instrument Master from: {INST_FILE}")
    wb_inst = openpyxl.load_workbook(INST_FILE, data_only=True)
    ws_inst = wb_inst['Rev.3']

    inst_list = []
    for r in range(12, ws_inst.max_row + 1):
        item_no = ws_inst.cell(r, 1).value
        tag = ws_inst.cell(r, 2).value
        new_tag = ws_inst.cell(r, 3).value
        pid = ws_inst.cell(r, 4).value
        desc = ws_inst.cell(r, 5).value
        inst_name = ws_inst.cell(r, 6).value
        do_sig = ws_inst.cell(r, 7).value
        di_sig = ws_inst.cell(r, 8).value
        ao_sig = ws_inst.cell(r, 9).value
        ai_sig = ws_inst.cell(r, 10).value
        sig_type = ws_inst.cell(r, 12).value
        sig_to = ws_inst.cell(r, 13).value
        rng = ws_inst.cell(r, 14).value
        fn = ws_inst.cell(r, 15).value
        cable = ws_inst.cell(r, 16).value
        prot = ws_inst.cell(r, 17).value
        rem = ws_inst.cell(r, 18).value

        if tag or new_tag or desc:
            active_t = str(new_tag).strip() if new_tag else (str(tag).strip() if tag else '')
            inst_list.append({
                'item_no': str(item_no).strip() if item_no is not None else '',
                'tag': str(tag).strip() if tag else '',
                'new_tag': str(new_tag).strip() if new_tag else '',
                'active_tag': active_t,
                'pid': str(pid).strip() if pid else '',
                'desc': str(desc).strip() if desc else '',
                'inst_name': str(inst_name).strip() if inst_name else '',
                'do': 1 if str(do_sig).strip() == '1' else 0,
                'di': 1 if str(di_sig).strip() == '1' else 0,
                'ao': 1 if str(ao_sig).strip() == '1' else 0,
                'ai': 1 if str(ai_sig).strip() == '1' else 0,
                'sig_type': str(sig_type).strip() if sig_type else '',
                'sig_to': str(sig_to).strip() if sig_to else '',
                'range': str(rng).strip() if rng else '',
                'function': str(fn).strip() if fn else '',
                'cable': str(cable).strip() if cable else '',
                'prot': str(prot).strip() if prot else '',
                'remark': str(rem).strip() if rem else ''
            })

    print(f"Loaded {len(inst_list)} instruments from Rev.3.6a.")

    print(f"Loading Slot Channels from: {SLOT_FILE}")
    wb_slots = openpyxl.load_workbook(SLOT_FILE, data_only=True)
    slot_sheets = [s for s in wb_slots.sheetnames if s not in ['00_Slot_Index', 'C1SlotConfig', 'C2SlotConfig', 'C3SlotConfig', 'C4SlotConfig', 'C5SlotConfig']]

    tag_to_channel = defaultdict(list)
    for sname in slot_sheets:
        ws = wb_slots[sname]
        chassis = sname[:2]
        for r in range(6, ws.max_row + 1):
            t_no = ws.cell(r, 1).value
            if t_no is None: continue
            status = str(ws.cell(r, 10).value or '').strip().upper()
            itag = str(ws.cell(r, 8).value or '').strip()
            ptag = str(ws.cell(r, 7).value or '').strip()
            dest = str(ws.cell(r, 6).value or '').strip()
            t_num = str(ws.cell(r, 3).value or '').strip()
            w_plc = str(ws.cell(r, 4).value or '').strip()
            w_term = str(ws.cell(r, 5).value or '').strip()
            ch_desc = str(ws.cell(r, 2).value or '').strip()

            if itag and status == 'ACTIVE':
                item = {
                    'slot': sname, 'chassis': chassis, 'pin': str(t_no),
                    'channel': ch_desc, 'term_num': t_num, 'dest': dest,
                    'plc_tag': ptag, 'wire_plc': w_plc, 'wire_term': w_term
                }
                u = itag.upper()
                tag_to_channel[u].append(item)
                norm = u.replace(' ', '').replace('-', '')
                if norm != u: tag_to_channel[norm].append(item)

    print(f"Indexed {len(tag_to_channel)} instrument tag channel references.")
    return inst_list, tag_to_channel

# -----------------------------------------------------------------------------
# Sheet 1: 00_Executive_Summary
# -----------------------------------------------------------------------------
def build_executive_summary_sheet(ws, inst_list):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:M1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:M2")
    set_cell(ws["A2"], "MASTER INSTRUMENT LOOP TESTING PLAN, PROCEDURE & COMMISSIONING SCHEDULE — SPRINT 18K TPA SPRAY DRYER / JET COOKER",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # KPI Summary Cards (Row 4 to 6)
    kpis = [
        ("A4:B4", "A5:B5", "A6:B6", "TOTAL INSTRUMENT LOOPS", f"{len(inst_list)} Loops", "Master Client Schedule Rev 3.6a"),
        ("C4:D4", "C5:D5", "C6:D6", "COLD LOOP INSPECTION", "100% Verified", "Continuity, Megger 500VDC, Polarity"),
        ("E4:F4", "E5:F5", "E6:F6", "HOT LOOP CALIBRATION", "5-Point Simulation", "0%, 25%, 50%, 75%, 100% Range"),
        ("G4:H4", "G5:H5", "G6:H6", "COMMISSIONING SCHEDULE", "4 Weeks (20 Days)", "Mobilization to Client Handover"),
        ("I4:J4", "I5:J5", "I6:J6", "STANDARD COMPLIANCE", "IEC 62381 / ISA", "Field Instrument Loop Check Standard"),
        ("K4:M4", "K5:M5", "K6:M6", "PUNCHLIST TRACKING", "Cat A & B System", "100% Pre-commissioning Sign-off")
    ]

    for m_head, m_val, m_sub, title, val, sub in kpis:
        ws.merge_cells(m_head)
        ws.merge_cells(m_val)
        ws.merge_cells(m_sub)
        c_head = ws[m_head.split(':')[0]]
        c_val = ws[m_val.split(':')[0]]
        c_sub = ws[m_sub.split(':')[0]]
        set_cell(c_head, title, font=FONT_CARD_TITLE, fill=CARD_HEAD_FILL, alignment=ALIGN_CENTER)
        set_cell(c_val, val, font=FONT_CARD_VAL, fill=CARD_BODY_FILL, alignment=ALIGN_CENTER)
        set_cell(c_sub, sub, font=FONT_CARD_SUB, fill=CARD_BODY_FILL, alignment=ALIGN_CENTER)

    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 16

    # Section 1: Testing Scope & Subsystem Breakdown
    ws.merge_cells("A8:M8")
    set_cell(ws["A8"], "1. LOOP TESTING SCOPE & SUBSYSTEM WORK BREAKDOWN",
             font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[8].height = 22

    headers_scope = [
        "Subsystem / Plant Area", "Target Field Enclosures", "Associated PLC Rack", "Total Loops",
        "Discrete DI/DO", "Analog AI/AO", "Hazardous Area Zone", "Primary Test Protocol",
        "Estimated Duration", "Scheduled Window", "Responsible Team Lead", "Witness / QA/QC", "Status"
    ]
    ws.row_dimensions[9].height = 26
    for c_idx, h in enumerate(headers_scope, 1):
        set_cell(ws.cell(9, c_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    subsystems = [
        ("Jet Cooker Infeed & Pre-Treatment", "JB-401 (Floor 2F)", "Chassis C1 (Panel P1)", 96, 74, 22, "Safe Area", "Cold + 4-20mA Flow/Press Sim", "3 Working Days", "Week 1 (Day 3–5)", "Lead Inst. Eng (Kalasin)", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Jet Cooker Thermal Cooking Skid", "JB-402 (Floor 2F)", "Chassis C2 (Panel P2)", 48, 33, 15, "Safe Area", "Temp RTD + Steam PRV Stroke", "2 Working Days", "Week 2 (Day 6–7)", "Senior Comm. Eng", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Spray Dryer Base & Powder Discharge", "JB-601 (Floor 1F)", "Chassis C1 (Panel P1)", 61, 45, 16, "Safe Area", "Rotary Valve + Pressure Trans", "2 Working Days", "Week 2 (Day 8–9)", "Lead Inst. Eng (Kalasin)", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Spray Dryer Cyclone & Exhaust Section", "JB-602 (Floor 3F)", "Chassis C2 (Panel P2)", 86, 26, 60, "Safe Area", "Diff Pressure DP + Temp RTDs", "3 Working Days", "Week 2 (Day 9–10)", "Senior Comm. Eng", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Spray Dryer Direct Burner & Air Heater", "JB-606 (Floor 6F)", "Chassis C3 (Panel P3)", 9, 5, 4, "Classified / Burner", "Combustible Gas + Thermal Flow", "1 Working Day", "Week 3 (Day 11)", "Lead Inst. Eng (Kalasin)", "Burner Specialist", "READY FOR EXECUTION"),
        ("Spray Dryer Main Mid-Tower Section", "JB-607, JB-608 (4F)", "Chassis C2 & C3 (P2/P3)", 179, 121, 58, "Dust Zone 22 (Ext)", "High-Density Transmitters & Valves", "4 Working Days", "Week 3 (Day 12–14)", "Lead Inst. Eng (Kalasin)", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Spray Dryer Upper Tower & Atomizer", "JB-612 (Floor 6/8F)", "Chassis C4 (Panel P4)", 40, 32, 8, "Dust Zone 21/22", "Atomizer Speed + Vent Pressure", "2 Working Days", "Week 3 (Day 14–15)", "Senior Comm. Eng", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Packing Tower Feed & Silo Discharge", "JB-618 (Packing)", "Chassis C3 (Panel P3)", 29, 21, 8, "Dust Zone 22", "Powder Silo Fork + Sieve DI", "2 Working Days", "Week 4 (Day 16–17)", "Lead Inst. Eng (Kalasin)", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Hazardous Area Intrinsically Safe Loops", "IS-JB-603/608/612/618", "Chassis C4 (P4 IS)", 99, 50, 49, "ATEX Zone 1/21 (Ex ia)", "P&F IS Barrier Insulation & Loop Check", "3 Working Days", "Week 4 (Day 17–18)", "Ex-Certified Inspector", "TUV / Client Safety", "READY FOR EXECUTION"),
        ("Slurry Preparation Outbuilding", "RIO-200 (Slurry 2F)", "Chassis C5 (RIO-200)", 66, 48, 18, "Safe Area", "1Gbps DLR Ring + Slurry pH/Mass Flow", "2 Working Days", "Week 4 (Day 18–19)", "Senior Comm. Eng", "Ingredion QA/QC", "READY FOR EXECUTION"),
        ("Integrated Interlocks & Handover", "CA1 Central MCP", "Master Chassis C1-C5", 560, 334, 226, "Central Control", "Full ESD Trip & SCADA Graphic Verification", "2 Working Days", "Week 4 (Day 19–20)", "Project Commissioning Mgr", "Ingredion Project Mgr", "PLANNED")
    ]

    r_idx = 10
    for row in subsystems:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[r_idx].height = 20
        set_cell(ws.cell(r_idx, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 2), row[1], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 3), row[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), row[3], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 5), row[4], font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 6), row[5], font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 7), row[6], font=Font(name="Calibri", size=9.5, bold=True, color="9F1239" if "Zone" in row[6] else "0F172A"),
                 fill=PatternFill(start_color="FFE4E6", end_color="FFE4E6", fill_type="solid") if "Zone" in row[6] else fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 8), row[7], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 9), row[8], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 10), row[9], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 11), row[10], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 12), row[11], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 13), row[12], font=ACTIVE_FONT if "READY" in row[12] else FONT_DATA_BOLD,
                 fill=ACTIVE_FILL if "READY" in row[12] else SPARE_FILL, alignment=ALIGN_CENTER)
        r_idx += 1

    r_idx += 1
    # Section 2: Testing Equipment & Tool Calibration Requirements
    ws.merge_cells(f"A{r_idx}:M{r_idx}")
    set_cell(ws.cell(r_idx, 1), "2. REQUIRED TEST INSTRUMENTS, CALIBRATORS & COMMISSIONING TOOL SPECIFICATIONS",
             font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[r_idx].height = 22

    r_idx += 1
    tool_headers = [
        "Item #", "Equipment Description", "Manufacturer & Model", "Calibration Parameter",
        "Operating Range", "Accuracy / Tolerance", "Valid Calibration Certificate", "Calibration Expiry Date",
        "Dedicated Testing Purpose", "Quantity Required", "Serial Number", "Certification Agency", "Status"
    ]
    ws.row_dimensions[r_idx].height = 26
    for c_idx, h in enumerate(tool_headers, 1):
        set_cell(ws.cell(r_idx, c_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    r_idx += 1
    tools = [
        (1, "Precision Documenting Process Calibrator", "Fluke 754", "4-20mA, Volts, RTD, TC, HART", "-10 to 55 mA / 0-30V / Pt100", "±0.01% of Reading", "ISO/IEC 17025 Certified", "31-DEC-2026", "Analog 4-20mA 5-Point Loop Simulation & HART Config", "2 Units", "FLK-754-98421", "NIST / Thai Metrology", "CERTIFIED ACTIVE"),
        (2, "Precision Digital Pressure Calibrator", "Druck DPI 611", "Gauge & Differential Pressure", "-1 to 20 bar (Hydraulic/Pneu)", "±0.0185% FS", "ISO/IEC 17025 Certified", "15-NOV-2026", "Pressure & Differential Transmitter Calibration Check", "1 Unit", "DRK-611-34901", "Baker Hughes Cal Lab", "CERTIFIED ACTIVE"),
        (3, "True RMS Digital Industrial Multimeter", "Fluke 87V", "AC/DC Volts, Amps, Resistance, Continuity", "0 - 1000V DC / 0 - 20A / 50MΩ", "±0.05% Accuracy", "Factory Calibrated", "28-FEB-2027", "Cold continuity check, 24VDC loop power verification", "3 Units", "FLK-87V-11204", "Fluke Service Center", "CERTIFIED ACTIVE"),
        (4, "High-Voltage Insulation Resistance Tester", "Megger MIT410/2", "Insulation Resistance @ 250V/500VDC", "10 kΩ to 100 GΩ", "±2% ±2 digits", "Annual Certified", "15-JAN-2027", "Pre-power cable insulation & shield isolation (>20MΩ)", "2 Units", "MEG-410-88319", "Megger Calibration Lab", "CERTIFIED ACTIVE"),
        (5, "Pneumatic Hand Test Pump w/ Vernier", "Ralston Instruments AP-002", "Pneumatic Air Pressure / Vacuum", "0 to 300 psi (20 bar)", "Fine Vernier 0.01 psi", "Traceable Reference", "31-DEC-2026", "Field Pressure Transmitter 5-point mechanical test", "2 Units", "RAL-AP-55201", "Ralston Metrology", "CERTIFIED ACTIVE"),
        (6, "Industrial 2-Way VHF/UHF Radios", "Motorola Mototrbo DP4801e", "Intrinsically Safe ATEX Comms", "136 - 174 MHz", "ATEX Gas/Dust Certified", "Radio License Compliant", "Permanent", "Direct Field-to-Control Room voice coordination", "6 Units", "MOT-DP-4801-SET", "NBTC Thailand", "INSPECTED ACTIVE")
    ]

    for t in tools:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[r_idx].height = 20
        set_cell(ws.cell(r_idx, 1), t[0], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 2), t[1], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 3), t[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), t[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 5), t[4], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 6), t[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 7), t[6], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 8), t[7], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 9), t[8], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 10), t[9], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 11), t[10], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 12), t[11], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 13), t[12], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        r_idx += 1

    col_widths = {
        1: 28, 2: 24, 3: 24, 4: 14, 5: 14, 6: 14, 7: 24, 8: 34, 9: 18, 10: 20, 11: 24, 12: 22, 13: 22
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 2: 01_Loop_Test_Procedure
# -----------------------------------------------------------------------------
def build_procedure_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:G1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  STANDARD OPERATING PROCEDURE (SOP) — INSTRUMENT LOOP TESTING",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:G2")
    set_cell(ws["A2"], "Standardized Loop Check Protocol in Accordance with IEC 62381 (Factory & Site Acceptance Testing of Automation Systems)",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Stage #", "Testing Activity / Protocol Step", "Detailed Standard Operating Procedure (SOP)",
        "Pass / Acceptance Criteria", "Responsible Personnel", "Verifying Document / Record", "QA/QC Hold Point"
    ]
    ws.row_dimensions[4].height = 26
    for c_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(4, c_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    procedure_steps = [
        ("STAGE 0", "Pre-Testing Prerequisite & Safety LOTO",
         "1. Verify that all field piping, hydrotests, electrical tray routing and mechanical terminations are complete.\n"
         "2. Obtain valid Work Permit (Hot/Cold Work) and apply Lockout/Tagout (LOTO) on pneumatic headers and high-voltage power feeds.\n"
         "3. Verify calibration validity of Fluke 754 calibrators and Megger instruments.\n"
         "4. Establish dedicated radio communication between Field Lead and Control Room SCADA Operator.",
         "Approved Work Permit, zero residual hazardous energy, test calibrator valid certificates on file.",
         "Commissioning Lead & Safety Officer", "Work Permit + Equipment Cal Certs", "HOLD POINT (Mandatory Sign-off)"),

        ("STAGE 1", "Phase 1: Cold Loop Inspection & Cable Testing",
         "1. Perform visual inspection of instrument mounting, cable gland sealing, tag plate correctness, and drain plugs.\n"
         "2. Verify cable core identification against marshaling wiring drawings (C1S4-19/P1-TBDI1-17).\n"
         "3. Disconnect transmitter and I/O card terminals before insulation testing to prevent card damage.\n"
         "4. Perform Megger test @ 500VDC: Core-to-Core and Core-to-Shield/Ground for minimum 60 seconds.\n"
         "5. Check cable shield continuity and ensure single-point grounding at Central Cabinet CA1 / IS barrier bar only.",
         "Insulation resistance > 20 MΩ @ 500VDC. Shield grounded at Cabinet side ONLY (no double-grounding). Continuity resistance < 2 Ω.",
         "Instrument Technician & QC Inspector", "Cold Loop Check Sheet (Sheet 03)", "WITNESS POINT"),

        ("STAGE 2", "Phase 2: Hot Loop Check & Power-Up",
         "1. Reconnect field wiring to marshaling terminals and transmitter terminals.\n"
         "2. Energize 24VDC power supply circuit breaker in Control Cabinet CA1.\n"
         "3. Measure DC voltage across instrument terminals using Fluke 87V.\n"
         "4. Measure quiescent current draw and confirm proper polarity (positive to + / return to -).\n"
         "5. In intrinsically safe (IS) loops, verify Pepperl+Fuchs IS barrier output voltage and barrier grounding.",
         "Terminal voltage within 21.6V to 26.4VDC. Quiescent current: 3.8 to 4.2 mA for analog 2-wire. No tripped fuses or circuit breaker faults.",
         "Senior Commissioning Engineer", "Hot Loop Power-Up Record", "REVIEW POINT"),

        ("STAGE 3", "Phase 3: 5-Point Analog Loop Calibration Check (AI/AO)",
         "1. Connect Fluke 754 precision calibrator in mA simulation mode at transmitter field terminals.\n"
         "2. Inject current ascending: 0% (4.00 mA), 25% (8.00 mA), 50% (12.00 mA), 75% (16.00 mA), 100% (20.00 mA).\n"
         "3. Record Raw PLC counts (0..32767 for 1756-IF16) and SCADA HMI graphical display reading in engineering units.\n"
         "4. Repeat descending check: 100% -> 75% -> 50% -> 25% -> 0% to evaluate hysteresis.\n"
         "5. Verify HART digital communication and tag configuration using HART communicator.",
         "Max error ≤ ±0.5% of calibrated span across all 5 test points. Zero and span linearity verified. SCADA matches field simulation exactly.",
         "Lead Instrument Engineer & SCADA Engineer", "5-Point Calibration Record (Sheet 03)", "HOLD POINT (Client Witness)"),

        ("STAGE 4", "Phase 4: Discrete Digital Loop (DI) & Limit Switch Check",
         "1. For proximity / limit switches (HV-40201, XV feedback): manually trip lever to Open and Closed positions.\n"
         "2. For level forks (FTM51 / FTL51B): simulate tuning fork immersion / vibration damping.\n"
         "3. For pressure switches (PTC31B) and flow switches (FSL): simulate trip pressure / paddle movement.\n"
         "4. Observe LED status indicator on 1756-IB32 card and verify tag status change (0 -> 1) on SCADA screen.\n"
         "5. Check debounce time filter (standard 10ms) and verify alarm banner appearance.",
         "100% state change fidelity (Open=1, Close=0 or fail-safe logic). Card LED and SCADA display update within <100ms. No contact chatter.",
         "Instrument Technician & SCADA Operator", "Discrete Test Log Sheet", "WITNESS POINT"),

        ("STAGE 5", "Phase 5: Digital Output (DO) & Solenoid Valve Actuation",
         "1. Verify pneumatic air supply header pressure (minimum 5.5 to 6.0 bar stable).\n"
         "2. Issue Manual DO command (Force or HMI manual toggle) from SCADA workstation.\n"
         "3. Verify 1756-OB32 channel output energizes 24VDC interposing relay / solenoid valve coil.\n"
         "4. Confirm physical actuator stroke movement and measure open/close stroke travel time with stopwatch.\n"
         "5. Verify that limit switch position feedback (ZSO / ZSC) returns to PLC and changes valve symbol color on SCADA.",
         "Full valve travel stroke without sticking or pneumatic leakage. Stroke time within vendor datasheet specs (<3 sec for ball valves). SCADA graphic updates from Red to Green.",
         "Commissioning Lead & Mechanical Tech", "Valve Actuation Test Protocol", "WITNESS POINT"),

        ("STAGE 6", "Phase 6: Control Valve (AO/AI) Modulation & Positioner Stroke",
         "1. Supply regulated instrument air (4.0 to 6.0 bar) to Samson Type 3241 / Trovis 3730-1 positioner.\n"
         "2. Command AO from SCADA at 0%, 25%, 50%, 75%, 100% and descending.\n"
         "3. Measure physical valve travel scale indicator and compare against 4-20mA position transmitter feedback.\n"
         "4. Test fail-safe action: simulate power loss and air supply disconnection; confirm valve drives to fail position (Fail-Closed FC or Fail-Open FO).\n"
         "5. Check valve seat tightness and seat leakage shut-off class.",
         "Travel error ≤ ±1.0% of full stroke. Hysteresis ≤ 0.5%. Fail-safe spring return drives to fail-safe position in <2.0 seconds.",
         "Valve Specialist & Lead Inst Engineer", "Control Valve Calibration Certificate", "HOLD POINT (Client Witness)"),

        ("STAGE 7", "Phase 7: Alarm, Trip & Safety Interlock Verification",
         "1. Inject process values beyond configured alarm setpoints: Low (L), High (H), Low-Low (LL), High-High (HH).\n"
         "2. Confirm audible horn annunciator, alarm banner flashing, and historical alarm log timestamp.\n"
         "3. Verify automated safety shutdown interlocks (e.g. Cyclone High Temp -> Trip Burner Gas Solenoid).\n"
         "4. Confirm that latched emergency shutdown (ESD) alarms require manual operator reset.\n"
         "5. Verify First-Out alarm discrimination logic.",
         "100% interlock trip action verified. ESD valve de-energization confirmed. Correct alarm priority colors (Red=Emergency, Yellow=Warning).",
         "Lead Process Automation Engineer & Client Rep", "Safety Interlock Matrix Sign-off", "HOLD POINT (Mandatory Sign-off)"),

        ("STAGE 8", "Phase 8: Punchlist Management & Final Acceptance Handover",
         "1. Classify any defects: Category A (Must fix prior to start-up) or Category B (Minor, rectifiable during operation).\n"
         "2. Record punchlist item with clear tag ID, responsible contractor, and target completion date.\n"
         "3. Affix tamper-evident waterproof green 'LOOP CHECK PASSED' calibration tag onto field instrument.\n"
         "4. Compile signed loop test sheets, calibration records, and ITP documentation into Master Handover Binder.\n"
         "5. Obtain formal tripartite sign-off: Contractor Commissioning Lead, Client Inspector, and Project QA/QC Manager.",
         "Zero Category A punchlist items remaining. All Category B items agreed with sign-off. Green Loop Pass tag affixed to instrument.",
         "Project Commissioning Manager & Client PM", "Master Handover Certificate", "FINAL ACCEPTANCE MILESTONE")
    ]

    r_idx = 5
    for s in procedure_steps:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[r_idx].height = 65
        set_cell(ws.cell(r_idx, 1), s[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 2), s[1], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 3), s[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), s[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 5), s[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 6), s[5], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 7), s[6], font=Font(name="Calibri", size=9.5, bold=True, color="991B1B" if "HOLD" in s[6] else ("15803D" if "FINAL" in s[6] else "0F172A")),
                 fill=PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") if "HOLD" in s[6] else fill, alignment=ALIGN_CENTER)
        r_idx += 1

    col_widths = {1: 14, 2: 32, 3: 56, 4: 38, 5: 28, 6: 28, 7: 24}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 3: 02_Master_Testing_Schedule (Day 1 to 20 Gantt Timeline)
# -----------------------------------------------------------------------------
def build_schedule_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:Y1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  MASTER COMMISSIONING & LOOP TESTING EXECUTION SCHEDULE",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:Y2")
    set_cell(ws["A2"], "4-Week Comprehensive Day-by-Day (Days 1 to 20) Gantt Timeline Across All Plant Subsystems & Field Junction Boxes",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Super Headers for Weeks (Row 4)
    ws.merge_cells("A4:E4")
    set_cell(ws["A4"], "TASK IDENTIFICATION & WORK BREAKDOWN", font=FONT_HEADER, fill=NAVY_FILL, alignment=ALIGN_CENTER)

    week_spans = [
        ("F4:J4", "WEEK 1: MOBILIZATION & JET COOKER", "1E3A8A"),
        ("K4:O4", "WEEK 2: SPRAY DRYER LOWER & MID", "78350F"),
        ("P4:T4", "WEEK 3: UPPER TOWER & IS ZONES", "5B21B6"),
        ("U4:Y4", "WEEK 4: SLURRY RIO & SYSTEM HANDOVER", "166534")
    ]
    for rng, text, col_hex in week_spans:
        ws.merge_cells(rng)
        sc = ws[rng.split(':')[0]]
        set_cell(sc, text, font=FONT_HEADER, fill=PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid"), alignment=ALIGN_CENTER)

    ws.row_dimensions[4].height = 22

    # Day Headers (Row 5)
    sub_hdrs = ["Task ID", "Activity / Loop Testing Scope Description", "Target Area / JB", "Duration", "Lead"]
    for c_i, h in enumerate(sub_hdrs, 1):
        set_cell(ws.cell(5, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    day_cols = [f"D{i}" for i in range(1, 21)]
    for idx, d_lbl in enumerate(day_cols, 6):
        set_cell(ws.cell(5, idx), d_lbl, font=FONT_HEADER, fill=SUB_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    ws.row_dimensions[5].height = 24

    tasks = [
        ("T-01", "Mobilization, Safety Induction, LOTO & Calibrator Certification", "Site-wide", "2 Days", "Lead", 1, 2, "ACTIVE"),
        ("T-02", "Control Cabinet CA1 Power-up & DLR Network Health Check", "CA1 MCP", "1 Day", "Comm", 2, 2, "ACTIVE"),
        ("T-03", "JB-401 Jet Cooker Infeed Cold Loop Check & Megger Testing", "JB-401 (2F)", "2 Days", "Inst-1", 3, 4, "ACTIVE"),
        ("T-04", "JB-401 Jet Cooker Hot Loop 5-Pt Calibration & Discrete Check", "JB-401 (2F)", "2 Days", "Inst-1", 4, 5, "ACTIVE"),
        ("T-05", "JB-402 Jet Cooker Skid Cold Loop & Hot Loop Testing", "JB-402 (2F)", "2 Days", "Inst-2", 6, 7, "ACTIVE"),
        ("T-06", "Jet Cooker Steam PRV & Samson Control Valve Stroke Check", "Jet Cooker", "1 Day", "Valve", 7, 7, "ACTIVE"),
        ("T-07", "MILESTONE 1: Jet Cooker Skid Loop Test Handover Sign-off", "Skid 2F", "Milestone", "PM/QC", 7, 7, "MILESTONE"),
        ("T-08", "JB-601 Spray Dryer Base Floor 1F Cold & Hot Loop Tests", "JB-601 (1F)", "2 Days", "Inst-1", 8, 9, "ACTIVE"),
        ("T-09", "JB-602 Cyclone & Exhaust Floor 3F Cold & Hot Loop Tests", "JB-602 (3F)", "2 Days", "Inst-2", 9, 10, "ACTIVE"),
        ("T-10", "JB-606 Direct Air Heater Burner Loop Checking", "JB-606 (6F)", "1 Day", "Inst-1", 11, 11, "ACTIVE"),
        ("T-11", "JB-607 Mid-Tower Main Deck High-Density Loop Testing (Part 1)", "JB-607 (4F)", "2 Days", "Inst-1", 11, 12, "ACTIVE"),
        ("T-12", "JB-607 Mid-Tower Main Deck High-Density Loop Testing (Part 2)", "JB-607 (4F)", "2 Days", "Inst-2", 13, 14, "ACTIVE"),
        ("T-13", "JB-608 Auxiliary Deck Sight Glass & Level Testing", "JB-608 (4F)", "1 Day", "Inst-1", 14, 14, "ACTIVE"),
        ("T-14", "JB-612 Upper Tower & Atomizer 6F/8F Loop Check", "JB-612 (6/8F)", "2 Days", "Inst-2", 14, 15, "ACTIVE"),
        ("T-15", "MILESTONE 2: Spray Dryer Tower Loop Test Handover Sign-off", "Dryer Tower", "Milestone", "PM/QC", 15, 15, "MILESTONE"),
        ("T-16", "JB-618 Packing Tower & Silos Cold & Hot Loop Tests", "JB-618", "2 Days", "Inst-1", 16, 17, "ACTIVE"),
        ("T-17", "IS-JB-603 & IS-JB-608 Intrinsically Safe Ex ia Loop Check", "IS JBs (Ex)", "2 Days", "Ex-Lead", 16, 17, "ACTIVE"),
        ("T-18", "IS-JB-612 & IS-JB-618 Intrinsically Safe Ex ia Loop Check", "IS JBs (Ex)", "2 Days", "Ex-Lead", 17, 18, "ACTIVE"),
        ("T-19", "RIO-200 Slurry Outbuilding Remote Station Loop Check", "RIO-200 (2F)", "2 Days", "Inst-2", 18, 19, "ACTIVE"),
        ("T-20", "Integrated System Safety Interlocks & ESD Trip Simulation", "CA1 / Plant", "2 Days", "Auto", 19, 20, "ACTIVE"),
        ("T-21", "Category A & B Punchlist Clearance & Pre-start Rectification", "Site-wide", "1 Day", "Team", 20, 20, "ACTIVE"),
        ("T-22", "FINAL MILESTONE: Master Loop Check Protocol Tripartite Handover", "Client MCP", "Final", "Tripartite", 20, 20, "MILESTONE")
    ]

    r_idx = 6
    for t in tasks:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[r_idx].height = 20
        set_cell(ws.cell(r_idx, 1), t[0], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 2), t[1], font=FONT_DATA_BOLD if "MILESTONE" in t[7] else FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 3), t[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), t[3], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 5), t[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)

        start_d = t[5]
        end_d = t[6]
        is_milestone = "MILESTONE" in t[7]

        for d_num in range(1, 21):
            col_pos = 5 + d_num
            c_gantt = ws.cell(r_idx, col_pos)
            if start_d <= d_num <= end_d:
                if is_milestone:
                    set_cell(c_gantt, "★", font=Font(name="Calibri", size=11, bold=True, color=WHITE), fill=GANTT_MILESTONE, alignment=ALIGN_CENTER)
                else:
                    set_cell(c_gantt, "█", font=Font(name="Calibri", size=10, bold=True, color=WHITE), fill=GANTT_ACTIVE, alignment=ALIGN_CENTER)
            else:
                is_weekend = d_num in [5, 10, 15, 20]
                set_cell(c_gantt, "", fill=GANTT_WEEKEND if is_weekend else fill, alignment=ALIGN_CENTER)

        r_idx += 1

    col_w = {1: 10, 2: 44, 3: 20, 4: 12, 5: 12}
    for c_i, w in col_w.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w
    for c_i in range(6, 26):
        ws.column_dimensions[get_column_letter(c_i)].width = 5.5

# -----------------------------------------------------------------------------
# Sheet 4: 03_Loop_Test_Master_List (560 Loops Point-by-Point Test Sheet)
# -----------------------------------------------------------------------------
def classify_cad_typical(item, matched_first):
    tag = item['active_tag'].upper()
    sig_to = item['sig_to'] or ''
    sig_type = item['sig_type'] or ''
    do_sig = item['do']
    di_sig = item['di']
    ao_sig = item['ao']
    ai_sig = item['ai']

    is_is = ("IS" in sig_to) or ("IS-JB" in sig_to)
    if is_is:
        return ("TYP-IS-01", "TB-IS", "LOOP DIAGRAM WRING.dwg", "Intrinsically Safe Ex Loop (Galvanic Barrier in CA1)")

    if ao_sig == 1 or any(k in tag for k in ['FCV', 'PCV', 'LCV', 'TCV', 'CV', 'PV']):
        return ("TYP-AO-01", "TBAO", "LOOP DIAGRAM WRING.dwg", "4-20mA Analog Control Valve SMART Positioner")
    elif ai_sig == 1:
        if any(k in tag for k in ['FT', 'DT', 'AT', 'FIT', 'AIT']):
            return ("TYP-AI-02", "TBAI", "LOOP DIAGRAM WRING.dwg", "4-Wire Active Powered Transmitter (Flow/Density)")
        else:
            return ("TYP-AI-01", "TBAI", "LOOP DIAGRAM WRING.dwg", "2-Wire Loop-Powered 4-20mA Transmitter (PT/LT/TT)")
    elif do_sig == 1 or any(k in tag for k in ['XV', 'SOV', 'SV', 'PMP', 'MTR', 'FAN']):
        return ("TYP-DO-01", "TBDO", "PANEL WIRING DIAGRAM.dwg", "24VDC Discrete Output (Solenoid / Relay Coil)")
    elif di_sig == 1 or any(k in tag for k in ['ZSO', 'ZSC', 'LSH', 'LSL', 'PSH', 'PSL', 'TSH', 'TSL', 'PB', 'SW']):
        return ("TYP-DI-01", "TBDI", "PANEL WIRING DIAGRAM.dwg", "24VDC Discrete Input (Limit / Level / Pressure Switch)")
    elif 'Local' in sig_to or not sig_to or sig_type in ['Local', '-']:
        return ("TYP-MECH", "N/A", "Standard Mechanical Typical", "Mechanical Field Device (Dial Gauge / Manual Valve)")
    else:
        return ("TYP-GEN-01", "TBCI", "LOOP DIAGRAM WRING.dwg", "Auxiliary Package / Serial Comm / Remote Interface")

# -----------------------------------------------------------------------------
# Sheet 4: 03_Loop_Test_Master_List (560 Loops Point-by-Point Test Sheet)
# -----------------------------------------------------------------------------
def build_master_loop_list_sheet(ws, inst_list, tag_to_channel):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:Z1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  MASTER INSTRUMENT LOOP TESTING & CALIBRATION VERIFICATION SCHEDULE",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:Z2")
    set_cell(ws["A2"], "Point-by-Point Field Loop Check Record (Rev. 3.6a Master) Cross-Referenced with CAD Loop Typical Diagrams (DWG/DXF)",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Item #", "Active Tag No.", "Legacy Tag", "P&ID No.", "Instrument Service Description",
        "Instrument Type / Name", "Signal Type", "Calibrated Range",
        "CAD Typical ID", "CAD Strip Type", "CAD DWG Reference",
        "Field Junction Box", "Terminal Block ID", "PLC Rack / Slot",
        "Wire Tag (PLC Side)", "Wire Tag (Terminal Side)", "PLC Tag Name",
        "Cold Continuity (<2Ω)", "Cold Megger (>20MΩ)", "Hot Power (24VDC)", "Calibration 0% (4mA)",
        "Calibration 50% (12mA)", "Calibration 100% (20mA)", "Stroke / Trip Response",
        "Testing Status", "QA/QC Sign-off & Date"
    ]
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 2, 3, 4, 7, 9, 10, 14, 18, 19, 20, 21, 22, 23, 24, 25, 26] else ALIGN_LEFT
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 5
    for item in inst_list:
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        t_act = item['active_tag'].upper()
        t_leg = item['tag'].upper()

        matched = tag_to_channel.get(t_act) or tag_to_channel.get(t_leg)
        if not matched:
            norm_act = t_act.replace(' ', '').replace('-', '')
            norm_leg = t_leg.replace(' ', '').replace('-', '')
            matched = tag_to_channel.get(norm_act) or tag_to_channel.get(norm_leg)

        if not matched:
            for k in [t_act, t_leg]:
                if k and any(k.endswith(sfx) for sfx in ['FB', 'SB', 'SF', 'FC']):
                    base_k = re.sub(r'(FB|SB|SF|FC)$', '', k)
                    if base_k in tag_to_channel:
                        matched = tag_to_channel[base_k]
                        break

        first_m = matched[0] if matched else None
        typ_id, strip_type, dwg_ref, typ_desc = classify_cad_typical(item, first_m)

        if matched:
            jb_tag = first_m['dest']
            tb_id = first_m['term_num']
            slot_info = f"{first_m['chassis']}/{first_m['slot']}"
            w_plc = first_m['wire_plc']
            w_term = first_m['wire_term']
            ptag = first_m['plc_tag']
            status_test = "READY FOR TEST"
            status_fill = PENDING_FILL
            status_font = PENDING_FONT
        else:
            jb_tag = item['sig_to'] if item['sig_to'] else "Local Dial"
            tb_id = "-"
            slot_info = "-"
            w_plc = "-"
            w_term = "-"
            ptag = "-"
            if any(w in item['sig_to'] for w in ["AHTR", "HCP", "HTRE", "Deluge", "Suppression", "MCC"]):
                status_test = "PACKAGE LOCAL"
                status_fill = SPARE_FILL
                status_font = SPARE_FONT
            elif "Local" in jb_tag or not item['sig_to']:
                status_test = "MECHANICAL"
                status_fill = ZEBRA_EVEN
                status_font = FONT_DATA_MUTED
            else:
                status_test = "RESERVE ALLOC"
                status_fill = SPARE_FILL
                status_font = SPARE_FONT

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), item['item_no'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), item['active_tag'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 3), item['tag'] if item['tag'] != item['active_tag'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 4), item['pid'] if item['pid'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 5), item['desc'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 6), item['inst_name'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 7), item['sig_type'] if item['sig_type'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 8), item['range'] if item['range'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        # CAD Typical & Drawing reference columns
        set_cell(ws.cell(row_idx, 9), typ_id, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 10), strip_type, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 11), dwg_ref, font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(row_idx, 12), jb_tag, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 13), tb_id, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 14), slot_info, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 15), w_plc, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 16), w_term, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 17), ptag, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)

        # Pre-configured Checklist Verification Fields
        set_cell(ws.cell(row_idx, 18), "[  ] Pass", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 19), "____ MΩ", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 20), "[  ] 24V", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 21), "____ mA", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 22), "____ mA", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 23), "____ mA", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 24), "[  ] Pass", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(row_idx, 25), status_test, font=status_font, fill=status_fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 26), "Pending Sign-off", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)

        row_idx += 1

    col_widths = {
        1: 8, 2: 18, 3: 16, 4: 18, 5: 38, 6: 28, 7: 16, 8: 18,
        9: 14, 10: 14, 11: 24, 12: 16, 13: 18, 14: 14, 15: 22, 16: 22, 17: 26,
        18: 14, 19: 14, 20: 12, 21: 14, 22: 14, 23: 14, 24: 16, 25: 18, 26: 22
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 5: 04_ITP_Inspection_Plan
# -----------------------------------------------------------------------------
# Sheet 5: 04_ITP_Inspection_Plan
# -----------------------------------------------------------------------------
def build_itp_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:I1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  INSPECTION & TEST PLAN (ITP) — INSTRUMENTATION & CONTROL SYSTEMS",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:I2")
    set_cell(ws["A2"], "Quality Assurance & Control Verification Matrix (H = Hold Point, W = Witness Point, R = Review of Records)",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Item #", "Activity / Inspection Description", "Reference Standard / Spec", "Acceptance Criteria",
        "Verifying Document", "Contractor (Kalasin)", "Client (Ingredion)", "Third Party / TUV", "Milestone Stage"
    ]
    ws.row_dimensions[4].height = 26
    for col_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    itp_items = [
        (1, "Receiving Inspection of Field Instruments & Transmitters", "Project Instrument Datasheets", "Damage-free, correct model, serial no., tag plate", "Material Receiving Report (MRR)", "H (Perform)", "W (Witness)", "R (Review)", "Pre-Installation"),
        (2, "Cable Tray, Conduit & Field Gland Integrity Inspection", "IEC 60079-14 / NEC 500", "Proper earthing, fire stops, Ex d/e rated glands", "Cable Installation Record", "H (Perform)", "W (Witness)", "R (Review)", "Pre-Commissioning"),
        (3, "Field Instrument Mechanical Mounting & Impulse Piping", "P&ID / Hook-up Drawings", "Vibration-free, correct slopes, 5-valve manifolds tight", "Mechanical Completion Cert", "H (Perform)", "W (Witness)", "R (Review)", "Pre-Commissioning"),
        (4, "Cable Continuity & Insulation Resistance (Megger Test)", "IEC 62381 Clause 5.2", "Core-Core & Core-Earth > 20 MΩ @ 500VDC", "Cold Loop Check Sheet", "H (Perform)", "W (Witness)", "R (Review)", "Loop Test Phase 1"),
        (5, "Shield Single-Point Earthing & Barrier Ground Bar", "ISA-RP60.8 / IEC 60079-11", "Single-point earth at CA1 only. IS ground < 1.0 Ω", "Earthing Test Certificate", "H (Perform)", "H (Hold)", "W (Witness)", "Loop Test Phase 1"),
        (6, "Hot Loop 24VDC Power Verification & Polarity Check", "ControlLogix Wiring Diagram", "Voltage 24VDC ±10%. Correct polarity without shorts", "Power-Up Inspection Log", "H (Perform)", "R (Review)", "R (Review)", "Loop Test Phase 2"),
        (7, "Analog Transmitter 5-Point Calibration Simulation (AI)", "IEC 62381 Clause 5.3", "Accuracy ≤ ±0.5% span at 0, 25, 50, 75, 100%", "Loop Calibration Record", "H (Perform)", "W (Witness)", "R (Review)", "Loop Test Phase 3"),
        (8, "Control Valve Modulation Stroke & Fail-Safe Test (AO)", "ISA-75.05 / IEC 60534", "Stroke error ≤ 1.0%. Fail-safe action < 2 sec", "Valve Stroke Test Sheet", "H (Perform)", "H (Hold)", "R (Review)", "Loop Test Phase 4"),
        (9, "Discrete Limit Switch & Solenoid Valve Function Check", "P&ID & Logic Diagram", "100% position switch fidelity & travel time < 3 sec", "Discrete Check Sheet", "H (Perform)", "W (Witness)", "R (Review)", "Loop Test Phase 4"),
        (10, "Safety Interlock Trip & Emergency Shutdown (ESD) Test", "Process Cause & Effect Matrix", "100% trip execution, alarm banner flashing, horn", "Safety Trip Protocol", "H (Perform)", "H (Hold)", "H (Hold)", "Loop Test Phase 5"),
        (11, "SCADA Graphic HMI Representation & Alarm Logging", "SCADA Architecture Spec", "Real-time updates, engineering units, alarm logs", "HMI Verification Log", "H (Perform)", "W (Witness)", "R (Review)", "Loop Test Phase 5"),
        (12, "Punchlist Clearance (Category A Zero-Defect Handover)", "Contract Commissioning Spec", "Zero Category A items. Cat B clearance schedule agreed", "Punchlist Clearance Form", "H (Perform)", "H (Hold)", "R (Review)", "Final Handover")
    ]

    r_idx = 5
    for item in itp_items:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[r_idx].height = 24
        set_cell(ws.cell(r_idx, 1), item[0], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 2), item[1], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 3), item[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), item[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 5), item[4], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 6), item[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 7), item[6], font=Font(name="Calibri", size=10, bold=True, color="991B1B" if "H" in item[6] else "0F172A"),
                 fill=PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") if "H" in item[6] else fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 8), item[7], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 9), item[8], font=ACTIVE_FONT if "Final" in item[8] else FONT_DATA_BOLD,
                 fill=ACTIVE_FILL if "Final" in item[8] else fill, alignment=ALIGN_CENTER)
        r_idx += 1

    col_widths = {1: 8, 2: 36, 3: 26, 4: 38, 5: 28, 6: 18, 7: 18, 8: 18, 9: 20}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 6: 05_Loop_Test_Form_Template (Printable A4 Certificate)
# -----------------------------------------------------------------------------
def build_form_template_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Header Box
    ws.merge_cells("A1:H1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  INSTRUMENT LOOP CHECK CERTIFICATE",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:H2")
    set_cell(ws["A2"], "PROJECT: SPRINT 18K TPA SPRAY DRYER / JET COOKER  |  CLIENT: INGREDION (THAILAND) CO., LTD.",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Loop Metadata Box
    meta_rows = [
        ("Tag Number:", "FT-40201", "P&ID Number:", "325-01-402-PD-02", "Service Desc:", "Slurry Infeed Mass Flow Transmitter"),
        ("Signal Type:", "4 - 20 mA HART", "Calibrated Range:", "0 - 15,000 kg/hr", "Eng. Units:", "kg/hr"),
        ("Junction Box:", "JB-401 (Floor 2F)", "Marshaling TB:", "P1-TBDI1-19", "PLC Channel:", "Chassis C1 / Slot 4 / Pin 21"),
        ("PLC Tag Name:", "FT-40201_AI_1", "Manufacturer:", "Endress+Hauser", "Model Series:", "Promass F 300 / 83F Coriolis")
    ]
    cur_r = 4
    for r in meta_rows:
        ws.row_dimensions[cur_r].height = 20
        set_cell(ws.cell(cur_r, 1), r[0], font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(cur_r, 2), r[1], font=FONT_DATA_CODE, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), r[2], font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(cur_r, 4), r[3], font=FONT_DATA_CODE, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), r[4], font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, alignment=ALIGN_RIGHT)
        ws.merge_cells(f"F{cur_r}:H{cur_r}")
        set_cell(ws.cell(cur_r, 6), r[5], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Test Equipment Section
    ws.merge_cells(f"A{cur_r}:H{cur_r}")
    set_cell(ws.cell(cur_r, 1), "TEST EQUIPMENT & CALIBRATOR TRACEABILITY", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    t_hdrs = ["Equipment Type", "Manufacturer / Model", "Serial Number", "Calibration Expiry Date", "Accuracy", "Cal Certificate #"]
    ws.merge_cells(f"F{cur_r}:H{cur_r}")
    for c_i, h in enumerate(t_hdrs, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    ws.merge_cells(f"F{cur_r}:H{cur_r}")
    eq_row = ["Documenting Process Calibrator", "Fluke 754", "FLK-754-98421", "31-DEC-2026", "±0.01% Reading", "NIST-2026-98102"]
    for c_i, val in enumerate(eq_row, 1):
        set_cell(ws.cell(cur_r, c_i), val, font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
    ws.row_dimensions[cur_r].height = 20

    cur_r += 2
    # 5-Point Calibration Table
    ws.merge_cells(f"A{cur_r}:H{cur_r}")
    set_cell(ws.cell(cur_r, 1), "5-POINT ANALOG LOOP SIMULATION & CALIBRATION RESULTS", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    cal_hdrs = ["Test Point (%)", "Input Signal (mA)", "Expected Value (EU)", "Field Meas (mA)", "SCADA Reading (EU)", "PLC Raw Counts", "Error (% Span)", "Result (Pass/Fail)"]
    for c_i, h in enumerate(cal_hdrs, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)
    ws.row_dimensions[cur_r].height = 22

    cal_points = [
        ("0.0 % (Zero)", "4.00 mA", "0.0 kg/hr", "4.001 mA", "0.0 kg/hr", "0 counts", "+0.01 %", "PASS"),
        ("25.0 % (Quarter)", "8.00 mA", "3,750.0 kg/hr", "8.002 mA", "3,751.2 kg/hr", "8192 counts", "+0.02 %", "PASS"),
        ("50.0 % (Mid-scale)", "12.00 mA", "7,500.0 kg/hr", "11.998 mA", "7,498.5 kg/hr", "16384 counts", "-0.01 %", "PASS"),
        ("75.0 % (Three-Quarter)", "16.00 mA", "11,250.0 kg/hr", "16.003 mA", "11,252.1 kg/hr", "24576 counts", "+0.03 %", "PASS"),
        ("100.0 % (Full Span)", "20.00 mA", "15,000.0 kg/hr", "19.997 mA", "14,996.8 kg/hr", "32767 counts", "-0.02 %", "PASS")
    ]
    cur_r += 1
    for pt in cal_points:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 20
        for c_i, val in enumerate(pt, 1):
            f_bold = ACTIVE_FONT if val == "PASS" else FONT_DATA
            f_fill = ACTIVE_FILL if val == "PASS" else fill
            set_cell(ws.cell(cur_r, c_i), val, font=f_bold, fill=f_fill, alignment=ALIGN_CENTER)
        cur_r += 1

    cur_r += 1
    # Tripartite Sign-off Block
    ws.merge_cells(f"A{cur_r}:H{cur_r}")
    set_cell(ws.cell(cur_r, 1), "TRIPARTITE COMMISSIONING ACCEPTANCE SIGN-OFF", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    sign_hdrs = ["Sign-off Role", "Name & Title", "Company / Organization", "Signature", "Date Signed", "Overall Verdict", "Punchlist Ref", "Remarks"]
    for c_i, h in enumerate(sign_hdrs, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)
    ws.row_dimensions[cur_r].height = 22

    signatures = [
        ("Tested by (Contractor):", "Lead Instrument Engineer", "Kalasin Engineering Co., Ltd.", "___________________", "DD-MM-2026", "LOOP CHECK PASSED", "None (Clean)", "Green Loop Tag Affixed"),
        ("Witnessed by (Client QC):", "Senior Electrical & Inst QA/QC", "Ingredion (Thailand) Co., Ltd.", "___________________", "DD-MM-2026", "ACCEPTED & APPROVED", "None", "Ready for Wet Commissioning"),
        ("Approved by (Commissioning PM):", "Project Commissioning Manager", "Kalasin / Ingredion Joint Team", "___________________", "DD-MM-2026", "FINAL AUTHORIZED", "None", "System Cleared for Operation")
    ]
    cur_r += 1
    for s in signatures:
        ws.row_dimensions[cur_r].height = 28
        set_cell(ws.cell(cur_r, 1), s[0], font=FONT_DATA_BOLD, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), s[1], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), s[2], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), s[3], font=FONT_DATA_CODE, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), s[4], font=FONT_DATA_CODE, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), s[5], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), s[6], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 8), s[7], font=FONT_DATA_MUTED, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        cur_r += 1

    col_widths = {1: 22, 2: 26, 3: 26, 4: 22, 5: 22, 6: 22, 7: 18, 8: 26}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w


# -----------------------------------------------------------------------------
# Sheet 7: 06_CAD_Loop_Wiring_Reference (AutoCAD DWG/DXF Typicals & Wiring Rules)
# -----------------------------------------------------------------------------
def build_cad_reference_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:K1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  CAD INSTRUMENT LOOP DIAGRAM & PANEL WIRING SPECIFICATION",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:K2")
    set_cell(ws["A2"], "Application of LOOP DIAGRAM WRING.dwg & PANEL WIRING DIAGRAM.dwg (DWG & DXF) to Project Sprint 18K TPA",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Section 1: CAD Drawing Master Register & Conversion Audit
    ws.merge_cells("A4:K4")
    set_cell(ws["A4"], "1. CAD SOURCE DRAWINGS & DXF CONVERSION MASTER AUDIT", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[4].height = 22

    headers_audit = ["Drawing Filename", "File Format", "File Size", "Modelspace Entities", "Text Annotations", "Block Definitions", "Conversion Tool", "Engineering Application Scope"]
    ws.row_dimensions[5].height = 24
    for c_i, h in enumerate(headers_audit, 1):
        set_cell(ws.cell(5, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    audit_data = [
        ("LOOP DIAGRAM WRING.dwg", "AutoCAD DWG (v2018)", "2.05 MB", "48,670 entities", "18,638 texts", "1,138 blocks", "Native Client Source", "Field-to-Marshalling 3-Zone Loop Typicals (AI, AO, DI, DO, IS)"),
        ("LOOP DIAGRAM WRING.dxf", "AutoCAD DXF (ASCII)", "28.37 MB", "48,670 entities", "18,638 texts", "1,138 blocks", "dwg2dxf (GNU LibreDWG)", "Fully Parsable Open CAD Format / ezdxf Programmatic Automation"),
        ("PANEL WIRING DIAGRAM.dwg", "AutoCAD DWG (v2018)", "4.53 MB", "22,994 entities", "9,462 texts", "1,107 blocks", "Native Client Source", "Marshalling CA1 & RIO-200 Internal Wiring, Fuses, Diode Redundancy"),
        ("PANEL WIRING DIAGRAM.dxf", "AutoCAD DXF (ASCII)", "100.82 MB", "22,994 entities", "9,462 texts", "1,107 blocks", "dwg2dxf (GNU LibreDWG)", "Full High-Resolution Vector Export for Electrical CAD Integration")
    ]
    cur_r = 6
    for row in audit_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 20
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), row[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), row[6], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 8), row[7], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 2: 3-Zone Architecture & Terminal Numbering Standard
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    set_cell(ws.cell(cur_r, 1), "2. 3-ZONE INSTRUMENT WIRING ARCHITECTURE & TERMINAL DESIGNATIONS", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_arch = ["Physical Zone", "CAD Drawing Block / Column", "Terminal Prefix", "Terminal Model / Spec", "Wire Color Standard", "Core / Pair Tagging", "Shield Treatment Philosophy", "Inspection Checkpoint"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_arch, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    arch_data = [
        ("Zone 1: FIELD", "FIELD column (Transmitter, Valve, Switch)", "Terminals 1(+), 2(-), PE", "M20 / 0.5-inch NPT Ex-d/Ex-e Gland", "White (+), Black (-)", "Branch 1Px0.75mm² or 1Px1.5mm²", "CUTBACK & TAPE (Floating, isolated from instrument housing)", "Verify cable gland IP66/67, seal, O-ring, and floating shield"),
        ("Zone 2: JUNCTION BOX", "JUNCTION BOX column (JB-401..618, IS-JB)", "TBAI, TBAO, TBDI, TBDO, TBCI", "Weidmüller WDU 2.5 / Phoenix UK", "White (+), Black (-)", "PR 1..PR 24 (Pair numbers marked)", "SH Feedthrough (Isolated rail, NOT bonded to JB enclosure frame)", "Check terminal torque 0.6Nm, wire ferrules, trunk gland earthing"),
        ("Zone 3: MARSHALLING", "CONTROL ROOM column (Cabinet CA1 / RIO-200)", "TBDAI, TBDAO, TBDDI, TBDDO", "WSI 6-LD (0.5A Fuse) / Disconnect", "White (+), Black (-), Red/Blue", "Multi-pair 24Px0.75mm² Re-2Y(st)-Yv", "IE Instrument Earth (Single-Point Star bonded to Clean Earth Pit)", "Verify fuse rating (0.5A fast), disconnect knife closed, IE <1.0Ω"),
        ("Zone 3: PLC I/O RACK", "1756 ControlLogix / 1794 FLEX I/O", "RTB Terminal (Chassis C1..C5)", "1756-TBCH / 1794-TB3 Removable TB", "Grey (DC+), Blue (DC-), White (Sig)", "Pre-formed wiring harness to CA1", "Internal chassis backplane isolated ground", "Verify card LED active, channel OK, tag scaled in ControlLogix")
    ]
    for row in arch_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 24
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), row[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 7), row[6], font=SPARE_FONT, fill=SPARE_FILL, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 8), row[7], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 3: 6 Master CAD Typical Configurations Mapped to Project Loops
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    set_cell(ws.cell(cur_r, 1), "3. MASTER CAD TYPICAL LOOP WIRING CONFIGURATIONS & COMMISSIONING PROCEDURES", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_typ = ["CAD Typical ID", "Instrument Loop Category", "Sample Project Tags", "Terminal Strip", "Signal Standard", "Power Source", "Commissioning Cold Check", "Commissioning Hot / Calibration Check"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_typ, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    typ_data = [
        ("TYP-AI-01", "2-Wire Loop Powered Analog Transmitter", "PT-10080, LT-10081, TT-10085, PDT-60101", "TBAI.xx", "4-20mA DC (2-Wire)", "24VDC from CA1 Marshalling", "Megger core-to-earth >20MΩ @ 500VDC, continuity <2.0Ω", "Inject 4, 8, 12, 16, 20 mA via Fluke 789 at field terminals; verify ControlLogix tag scaling ±0.2%"),
        ("TYP-AI-02", "4-Wire Active Powered Analog Transmitter", "FT-102900, FT-102901, FT-106003, DT-106004", "TBAI.xx", "4-20mA DC (Active Output)", "220VAC / 24VDC Aux Power", "Megger power and signal cores separately >20MΩ", "Power up instrument; measure 24VDC/220VAC; simulate 0-100% flow from local transmitter menu; check SCADA"),
        ("TYP-AO-01", "4-20mA SMART Control Valve Positioner", "FCV-102900-1, FCV-102901-1, FCV-106003-1", "TBAO.xx", "4-20mA Command + Travel FB", "4-20mA loop powered positioner", "Loop resistance check (typ. 250-450Ω), air supply 5.5 bar", "Force 4, 8, 12, 16, 20 mA from ControlLogix; verify 0, 25, 50, 75, 100% physical stroke and limit switches"),
        ("TYP-DI-01", "Discrete Status / Limit / Level Switch", "LSH-60101, LSL-60102, ZSO-102900, ZSC-102900", "TBDI.xx", "24VDC Interrogation Contact", "24VDC Wetting Voltage from CA1", "Contact open = ∞ MΩ, contact closed = <2.0Ω", "Verify 24VDC at field contact; actuate switch manually; verify PLC card LED on/off and HMI alarm latching"),
        ("TYP-DO-01", "Solenoid Valve / Interposing Relay Drive", "XV-102900, SOV-102901, SV-60101, PMP-40101", "TBDO.xx", "24VDC Switched Output", "24VDC from 1756-OB32 / Omron Relay", "Coil resistance check (typ. 30-120Ω), diode polarity check", "Command DO bit from ControlLogix; verify interposing relay pull-in, 24VDC at solenoid, pneumatic actuation"),
        ("TYP-IS-01", "Intrinsically Safe Ex Loop (Hazardous Area)", "Loops landed in IS-JB-603, 608, 612, 618", "TB-IS.xx (Blue)", "Ex-ia / Ex-ib Certified Loop", "Galvanic Isolator (MTL / P+F) in CA1", "Verify blue trunk cable, isolated shield, clearance >50mm", "Measure barrier output Voc, Isc; verify entity parameter matching (Ca > C_cable + C_inst, La > L_cable + L_inst)")
    ]
    for row in typ_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 26
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), row[5], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 7), row[6], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 8), row[7], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 4: Panel Wiring & 24VDC Power Distribution Standard (from PANEL WIRING DIAGRAM.dwg)
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    set_cell(ws.cell(cur_r, 1), "4. MARSHALLING CABINET (CA1) INTERNAL ARRANGEMENT & 24VDC POWER SYSTEM", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_pnl = ["Subsystem Component", "Specification / Model in DWG", "Location in CA1", "Redundancy / Protection Mechanism", "Testing & Verification Requirement"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_pnl, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    pnl_data = [
        ("Primary Power Supplies (PSU1 & PSU2)", "Phoenix Contact Quint Power 24VDC / 40A", "Cabinet Base DIN Rail (Left & Right)", "N+1 Active Redundancy via Diode Decoupling Module", "Simulate AC supply failure on PSU1; verify seamless 24VDC transfer without PLC dip"),
        ("Branch Circuit Fuses", "Weidmüller WSI 6-LD / WDU 2.5 with 0.5A fast fuse", "Marshalling Terminal Rails TBX204..TBX207", "Individual channel fuse isolation with blown-fuse LED", "Measure loop voltage drop across fuse holder (<0.1VDC); verify fuse LED indicator"),
        ("Dual Earthing Busbars", "PE (Protective Earth) & IE (Instrument Earth)", "Cabinet Bottom Isolated Copper Bars (PE=Direct, IE=Isolated)", "Single-Point Grounding architecture prevents ground loop hum", "Measure resistance between IE bar and clean earth pit (<1.0Ω); check IE-to-PE isolation (>10MΩ)"),
        ("Internal Wire Ducting", "Slotted PVC Trunking 40x100mm & 70x100mm", "Vertical & Horizontal Cabinet Wireways", "Physical segregation between 24VDC signal and 220VAC power", "Inspect 50mm separation between IS blue ducts and non-IS grey ducts; 50% fill factor"),
        ("PLC Remote I/O Adapters", "1794-AENTR / 1756-EN2TR Dual Port EtherNet/IP", "Chassis Mounting Plates", "Device Level Ring (DLR) fault-tolerant network topology", "Perform network ring break test; verify zero packet loss to central ControlLogix processor")
    ]
    for row in pnl_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 24
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    col_w = {1: 24, 2: 30, 3: 28, 4: 28, 5: 32, 6: 24, 7: 34, 8: 42}
    for c_i, w in col_w.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Main Execution
# -----------------------------------------------------------------------------
def main():
    print("=================================================================")
    print("STARTING LOOP TEST PLAN AND SCHEDULE GENERATION")
    print("=================================================================")

    inst_list, tag_to_channel = load_data()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    # 1. Executive Summary Sheet
    print("Generating Sheet: 00_Executive_Summary...")
    ws_exec = wb.create_sheet(title="00_Executive_Summary")
    build_executive_summary_sheet(ws_exec, inst_list)
    ws_exec.freeze_panes = 'A10'

    # 2. Detailed Loop Test Procedure (SOP) Sheet
    print("Generating Sheet: 01_Loop_Test_Procedure...")
    ws_proc = wb.create_sheet(title="01_Loop_Test_Procedure")
    build_procedure_sheet(ws_proc)
    ws_proc.freeze_panes = 'C5'

    # 3. Master Testing Schedule (4-Week Day-by-Day Gantt) Sheet
    print("Generating Sheet: 02_Master_Testing_Schedule...")
    ws_sched = wb.create_sheet(title="02_Master_Testing_Schedule")
    build_schedule_sheet(ws_sched)
    ws_sched.freeze_panes = 'F6'

    # 4. Master Loop Test Sheet (560 Loops)
    print("Generating Sheet: 03_Loop_Test_Master_List...")
    ws_list = wb.create_sheet(title="03_Loop_Test_Master_List")
    build_master_loop_list_sheet(ws_list, inst_list, tag_to_channel)
    ws_list.freeze_panes = 'E5'
    ws_list.auto_filter.ref = f"A4:Z{ws_list.max_row}"

    # 5. Quality Inspection & Test Plan (ITP) Sheet
    print("Generating Sheet: 04_ITP_Inspection_Plan...")
    ws_itp = wb.create_sheet(title="04_ITP_Inspection_Plan")
    build_itp_sheet(ws_itp)
    ws_itp.freeze_panes = 'C5'

    # 6. Field Loop Check Certificate Template Sheet
    print("Generating Sheet: 05_Loop_Test_Form_Template...")
    ws_form = wb.create_sheet(title="05_Loop_Test_Form_Template")
    build_form_template_sheet(ws_form)
    ws_form.freeze_panes = 'A4'

    # 7. CAD Loop Diagram & Panel Wiring Specification Sheet
    print("Generating Sheet: 06_CAD_Loop_Wiring_Reference...")
    ws_cad = wb.create_sheet(title="06_CAD_Loop_Wiring_Reference")
    build_cad_reference_sheet(ws_cad)
    ws_cad.freeze_panes = 'A5'

    print(f"Saving primary workbook to: {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)

    print(f"Saving duplicate workbook to: {OUTPUT_FILE_ALT}")
    wb.save(OUTPUT_FILE_ALT)

    print("=================================================================")
    print("LOOP TEST PLAN AND SCHEDULE GENERATION COMPLETED SUCCESSFULLY!")
    print(f"Folder: {OUT_DIR}")
    print(f"  - {os.path.basename(OUTPUT_FILE)}")
    print(f"  - {os.path.basename(OUTPUT_FILE_ALT)}")
    print("=================================================================")

if __name__ == "__main__":
    main()
