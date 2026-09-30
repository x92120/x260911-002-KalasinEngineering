#!/usr/bin/env python3
"""
=============================================================================
KALASIN ENGINEERING xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)
IO LIST EXTRACTION BY CHASSIS SLOT CONFIGURATION (C1S0 through C5S4)
ALL CHASSIS C1 - C4 STANDARDIZED TO 13-SLOT ARCHITECTURE (SLOTS 0 TO 12)
WITH DEDICATED CxSlotConfig CHASSIS SUMMARY SHEETS
AND TERMINAL NUMBER COLUMN (Px-TBDIx-m) ON ALL SLOT SCHEDULES
=============================================================================
Source File: 03_IO_Lists_and_Schedules/IO_List_xDev-R02.xlsx (Sheet: 'IO List')
Output File: 03_IO_Lists_and_Schedules/IO_List-By_SlotConfig.xlsx

Chassis Configuration (13-Slot Standard for C1 to C4):
  - C1: 13 Slots (Slots 0 to 12) -> C1S0 (CPU) .. C1S12 (1756-N2 Empty)
  - C2: 13 Slots (Slots 0 to 12) -> C2S0 (Comm) .. C2S10 (AI), C2S11 (AI Spare), C2S12 (AI Spare)
  - C3: 13 Slots (Slots 0 to 12) -> C3S0 (Comm) .. C3S11 (Empty), C3S12 (AO 1756-OF8)
  - C4: 13 Slots (Slots 0 to 12) -> C4S0 (Comm) .. C4S6 (AI), C4S7..C4S12 (1756-N2 Empty)
  - C5: 5 Slots  (Slots 0 to 4)  -> C5S0 (Comm) .. C5S4 (Slurry Building RIO-200)

Required Columns per Slot Sheet (10 Columns):
  Col 1:  Terminal_No                  (Ordered numerically: 1..36, 1..20, RJ45-1/2, USB-1, etc.)
  Col 2:  Terminal_Description         (Channel / function: IN-0, OUT-0, GND-0, VOUT-0, etc.)
  Col 3:  Terminal_Number (Px-TBDIx-m) (Marshaling Terminal Block item: P1-TBDI1-1, P2-TBDI1-8, etc.)
  Col 4:  PLC_Tag_Side                 (Wire mark on PLC side: C1S4-1/P1-TBDI1-1)
  Col 5:  Terminal_Tag_Side            (Wire mark on terminal side: P1-TBDI1-1/C1S4-1)
  Col 6:  Destination_Area             (Target JB / Field Enclosure: CA1, JB-401, etc.)
  Col 7:  PLC_Tag                      (Program tag name: Spare_DI_0, FSL-40201_DI_1)
  Col 8:  Instrument_Tag               (Field instrument tag)
  Col 9:  Instrument_Description       (Instrument process service description)
  Col 10: Channel_Status               (ACTIVE / SPARE / COMMON / EMPTY)
=============================================================================
"""

import os
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
SOURCE_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R02.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List-By_SlotConfig.xlsx")

# Color Palette Definitions
NAVY_HEADER = "1E293B"      # Slate 800 - Main Table Header
NAVY_SUBHEADER = "334155"   # Slate 700 - Subtitle Banner
CHASSIS_HDR = "0F172A"      # Slate 900 - Chassis Header
WHITE = "FFFFFF"
BORDER_COLOR = "CBD5E1"     # Slate 300
LIGHT_BG_1 = "FFFFFF"
LIGHT_BG_2 = "F8FAFC"       # Slate 50

# Group Header Colors
SRC_HDR_FILL = "DBEAFE"     # Light Ice Blue
SRC_HDR_FONT = "1E3A8A"     # Deep Blue
DEST_HDR_FILL = "FEF3C7"    # Light Amber/Sand
DEST_HDR_FONT = "78350F"    # Deep Amber/Brown
SYS_HDR_FILL = "E2E8F0"     # Pale Slate
SYS_HDR_FONT = "0F172A"     # Slate 900

# Status Fills
ACTIVE_FILL = "DCFCE7"      # Light Emerald
ACTIVE_FONT = "14532D"      # Dark Forest Green
SPARE_FILL = "FEF9C3"       # Light Yellow/Amber
SPARE_FONT = "854D0E"       # Dark Amber
COMMON_FILL = "F1F5F9"     # Slate 100
COMMON_FONT = "475569"     # Slate 600

# 13-Slot Architecture for all main ControlLogix Chassis C1 to C4
CHASSIS_META = {
    'C1': {
        'name': 'Main Controller Rack (1756-L950TPSXT CPU)',
        'location': 'Control Cabinet CA1 (Main Control Room MCP)',
        'model': '1756-A13 (13-Slot ControlLogix Chassis)',
        'psu': '1756-PA75 (85-265V AC Power Supply)',
        'slots': list(range(0, 13))  # 13 Slots: S0 to S12
    },
    'C2': {
        'name': 'Main Expansion Rack (High-Density Field I/O)',
        'location': 'Control Cabinet CA1 (Main Control Room MCP)',
        'model': '1756-A13 (13-Slot ControlLogix Chassis)',
        'psu': '1756-PA75 (85-265V AC Power Supply)',
        'slots': list(range(0, 13))  # 13 Slots: S0 to S12 (includes C2S11 & C2S12)
    },
    'C3': {
        'name': 'Remote I/O Rack 1 (Spray Dryer / Packing Tower)',
        'location': 'Field RIO Cabinet (Spray Dryer / Packing Tower)',
        'model': '1756-A13 (13-Slot ControlLogix Chassis)',
        'psu': '1756-PA75 (85-265V AC Power Supply)',
        'slots': list(range(0, 13))  # 13 Slots: S0 to S12
    },
    'C4': {
        'name': 'Remote I/O Rack 2 (Upper Spray Dryer 6th/8th Floor)',
        'location': 'Field RIO Cabinet (Spray Dryer 6th/8th Floor)',
        'model': '1756-A13 (13-Slot ControlLogix Chassis)',
        'psu': '1756-PA75 (85-265V AC Power Supply)',
        'slots': list(range(0, 13))  # 13 Slots: S0 to S12 (includes C4S12)
    },
    'C5': {
        'name': 'Remote I/O Station 5 (Slurry Building RIO-200)',
        'location': 'Slurry Building 2nd Floor (RIO-200)',
        'model': '1756-A7 (7-Slot ControlLogix Chassis)',
        'psu': '1756-PA75 (85-265V AC Power Supply)',
        'slots': list(range(0, 5))   # 5 Slots: S0 to S4
    }
}

MODULE_INFO = {
    '1756-L950TPSXT': {
        'desc': 'ControlLogix 5580 Controller (40MB User Memory, Conformal Coated)',
        'sig': 'CPU (Controller)',
        'rtb': 'Integrated (RJ45 / USB / SD Slot)',
        'total_chan': 5
    },
    '1756-EN4TR': {
        'desc': 'EtherNet/IP Communication Adapter (1Gbps, 2-Port DLR)',
        'sig': 'COMM (EtherNet/IP)',
        'rtb': '2x RJ45 (Gigabit) + 1x USB-B',
        'total_chan': 4
    },
    '1756-IB32': {
        'desc': '32-Point 24VDC Sink/Source Digital Input Module',
        'sig': 'DI (24VDC Digital Input)',
        'rtb': '1756-TBCH (36-Pin RTB)',
        'total_chan': 32
    },
    '1756-OB32': {
        'desc': '32-Point 24VDC Sourcing Digital Output Module',
        'sig': 'DO (24VDC Digital Output)',
        'rtb': '1756-TBCH (36-Pin RTB)',
        'total_chan': 32
    },
    '1756-IF16': {
        'desc': '16-Point Isolated/Non-Isolated Analog Input Module',
        'sig': 'AI (4-20mA / Voltage / RTD)',
        'rtb': '1756-TBCH (36-Pin RTB)',
        'total_chan': 16
    },
    '1756-OF8': {
        'desc': '8-Point High-Resolution Analog Current/Voltage Output Module',
        'sig': 'AO (4-20mA Analog Output)',
        'rtb': '1756-TBNH (20-Pin RTB)',
        'total_chan': 8
    },
    '1756-N2': {
        'desc': 'Slot Filler / Reserve Unpopulated Position',
        'sig': 'EMPTY (Reserve Position)',
        'rtb': 'N/A (Blanking Plate)',
        'total_chan': 0
    }
}

SLOT_TB_MAP = {
    'C1S4': 'P1-TBDI1',
    'C1S5': 'P1-TBDI2',
    'C1S6': 'P1-TBDI3',
    'C1S7': 'P1-TBDI4',
    'C1S8': 'P1-TBRL1',
    'C1S9': 'P1-TBRL2',
    'C1S10': 'P1-TBAI1',
    'C1S11': 'P1-TBAI2',
    
    'C2S1': 'P2-TBDI1',
    'C2S2': 'P2-TBDI2',
    'C2S3': 'P2-TBDI3',
    'C2S4': 'P2-TBDI4',
    'C2S5': 'P2-TBDI5',
    'C2S6': 'P2-TBRL1',
    'C2S7': 'P2-TBRL2',
    'C2S8': 'P2-TBRL3',
    'C2S9': 'P2-TBAI1',
    'C2S10': 'P2-TBAI2',
    'C2S11': 'P2-TBAI3',
    'C2S12': 'P2-TBAI4',
    
    'C3S1': 'P3-TBDI1',
    'C3S2': 'P3-TBDI2',
    'C3S3': 'P3-TBDI3',
    'C3S4': 'P3-TBRL1',
    'C3S5': 'P3-TBRL2',
    'C3S6': 'P3-TBAI1',
    'C3S7': 'P3-TBAI2',
    'C3S8': 'P3-TBAI3',
    'C3S9': 'P3-TBAI4',
    'C3S12': 'P3-TBAO1',
    
    'C4S1': 'P4-TBDI1',
    'C4S2': 'P4-TBDI2',
    'C4S3': 'P4-TBDI3',
    'C4S4': 'P4-TBRL1',
    'C4S5': 'P4-TBAI1',
    'C4S6': 'P4-TBAI2',
}

def thin_border():
    s = Side(style='thin', color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s_light = Side(style='thin', color="475569")
    s_bottom = Side(style='medium', color="0F172A")
    return Border(left=s_light, right=s_light, top=s_light, bottom=s_bottom)

def summary_total_border():
    s_top = Side(style='thin', color="0F172A")
    s_double = Side(style='double', color="0F172A")
    s_side = Side(style='thin', color="CBD5E1")
    return Border(left=s_side, right=s_side, top=s_top, bottom=s_double)

def load_source_data():
    print(f"Loading source workbook: {SOURCE_FILE}")
    wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)
    
    # 1. Load Card_Terminal definitions: model -> {terminal_no: sig_chan}
    card_term_map = defaultdict(dict)
    if 'Card_Terminal' in wb.sheetnames:
        ws_card = wb['Card_Terminal']
        for r in list(ws_card.iter_rows(values_only=True))[1:]:
            model = r[2]
            sig_chan = r[4]
            term = r[5]
            if model and term is not None:
                try:
                    t_int = int(term)
                    card_term_map[str(model).strip()][t_int] = str(sig_chan).strip() if sig_chan else f"PIN-{t_int}"
                except:
                    pass

    # 2. Load Slot_Description mapping
    slot_desc_map = {}
    if 'Slot_Description' in wb.sheetnames:
        ws_slot = wb['Slot_Description']
        for r in list(ws_slot.iter_rows(values_only=True))[1:]:
            code = r[0]
            rack = r[1]
            slot_no = r[2]
            card_short = r[3]
            card_desc = r[4]
            io_t = r[5]
            if rack and slot_no is not None:
                try:
                    s_key = f"{rack}S{int(slot_no)}"
                    slot_desc_map[s_key] = {
                        'card_short': card_short,
                        'card_desc': card_desc,
                        'io_type': io_t
                    }
                except:
                    pass

    # 3. Load IO List rows
    ws_io = wb['IO List']
    slot_rows = defaultdict(list)
    slot_term_lookup = defaultdict(dict)
    
    for r in list(ws_io.iter_rows(values_only=True))[1:]:
        col7 = str(r[2]).strip() if r[2] is not None else ''
        ch = str(r[30]).strip() if r[30] is not None else ''
        sl = str(r[31]).strip() if r[31] is not None else ''
        if not ch and col7.startswith('C') and 'S' in col7:
            ch = col7.split('S')[0]
            sl = col7.split('S')[1].split(',')[0]
            
        if ch in ['C1', 'C2', 'C3', 'C4', 'C5']:
            try:
                s_key = f"{ch}S{int(sl)}"
            except:
                s_key = f"{ch}S{sl}"
                
            slot_rows[s_key].append(r)
            front_term = r[34]
            if front_term is not None:
                try:
                    slot_term_lookup[s_key][int(front_term)] = r
                except:
                    pass

    return card_term_map, slot_desc_map, slot_rows, slot_term_lookup

def get_slot_summary_info(s_key, ch, sl_num, slot_desc_map, slot_rows):
    s_desc_info = slot_desc_map.get(s_key, {})
    rows = slot_rows.get(s_key, [])
    
    if rows:
        card_model = rows[0][3] or rows[0][26]
        io_type = rows[0][20] or ""
        dests = sorted(list(set(str(r[1]).strip() for r in rows if r[1])))
        dest_str = ", ".join(dests)
        spares = sum(1 for r in rows if 'SPARE' in str(r[11]).upper() or not r[11] or r[11] == '#N/A')
        active = len(rows) - spares
    else:
        card_model = s_desc_info.get('card_desc', '1756-N2')
        io_type = s_desc_info.get('io_type', 'EMPTY')
        dest_str = "CA1" if ch in ['C1', 'C2'] else "Field RIO"
        active = 0
        spares = 0

    if not card_model:
        card_model = s_desc_info.get('card_desc', '1756-N2')
        
    if '-' in str(card_model):
        parts = str(card_model).split('-')
        if len(parts) >= 2 and not parts[1].isdigit():
            clean_card = f"{parts[0]}-{parts[1]}"
        else:
            clean_card = str(card_model)
    else:
        clean_card = str(card_model)

    mod_meta = MODULE_INFO.get(clean_card, {
        'desc': f'Module {clean_card}',
        'sig': io_type or 'General',
        'rtb': 'Standard RTB',
        'total_chan': len(rows)
    })
    total_chan = mod_meta['total_chan']

    if 'L950' in clean_card or 'CPU' in str(io_type).upper() or s_key == 'C1S0':
        clean_card = '1756-L950TPSXT'
        active = 5
        spares = 0
        dest_str = "CA1 (Main Control Room)"
        status = "ACTIVE"
    elif 'EN4TR' in clean_card or 'COMM' in str(io_type).upper() or 'ETH' in str(io_type).upper() or s_key.endswith('S0'):
        clean_card = '1756-EN4TR'
        active = 4
        spares = 0
        dest_str = "CA1 / Field RIO (DLR Ring)"
        status = "ACTIVE"
    elif clean_card in ['1756-IB32', '1756-OB32', '1756-IF16', '1756-OF8']:
        if active > 0:
            status = "ACTIVE"
        else:
            # Installed spare card (e.g. C2S11, C2S12)
            status = "SPARE RESERVE"
            spares = total_chan
    else:
        clean_card = '1756-N2'
        status = "EMPTY"
        dest_str = "Unpopulated / Reserve"

    return {
        'clean_card': clean_card,
        'desc': mod_meta['desc'],
        'sig': mod_meta['sig'],
        'rtb': mod_meta['rtb'],
        'dest_str': dest_str,
        'total_chan': total_chan,
        'active': active,
        'spares': spares,
        'status': status
    }

def build_chassis_config_sheet(ws, ch, ch_meta, slot_desc_map, slot_rows):
    """
    Builds the dedicated CxSlotConfig sheet listing all mounted slots for Chassis cx.
    """
    ws.views.sheetView[0].showGridLines = True
    
    # 1. Main Header Title
    ws.merge_cells("A1:K1")
    t1 = ws["A1"]
    t1.value = f"CHASSIS {ch} HARDWARE & SLOT CONFIGURATION SCHEDULE"
    t1.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    t1.fill = PatternFill(start_color=CHASSIS_HDR, end_color=CHASSIS_HDR, fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # 2. Chassis Title & Location
    ws.merge_cells("A2:K2")
    t2 = ws["A2"]
    t2.value = f"{ch_meta['name'].upper()}   |   LOCATION: {ch_meta['location'].upper()}"
    t2.font = Font(name="Arial", size=9.5, bold=True, color=WHITE)
    t2.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 20

    # 3. Chassis Meta Banner
    ws.merge_cells("A3:K3")
    t3 = ws["A3"]
    t3.value = f"CHASSIS MODEL: {ch_meta['model']}   |   POWER SUPPLY: {ch_meta['psu']}   |   NETWORK: ETHERNET/IP 1GBPS DLR RING"
    t3.font = Font(name="Arial", size=8.5, bold=True, color="334155")
    t3.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    t3.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18

    # 4. Navigation Links
    ws["A4"].value = "← Back to Master Slot Directory (00_Slot_Index)"
    ws["A4"].hyperlink = "#'00_Slot_Index'!A1"
    ws["A4"].font = Font(name="Arial", size=8.5, italic=True, color="2563EB", underline="single")
    ws.row_dimensions[4].height = 17

    # 5. Table Headers - Row 6
    headers = [
        "Slot_No",
        "Slot_Tag",
        "Module_Catalog",
        "Module_Description",
        "Signal_Family",
        "Terminal_Block_Type",
        "Destination_Area(s)",
        "Total_Channels",
        "Active_Channels",
        "Spare_Channels",
        "Slot_Status"
    ]
    ws.append([]) # Row 5 empty
    ws.append(headers) # Row 6 headers
    ws.row_dimensions[6].height = 22
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=6, column=col_idx)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border()

    r_num = 7
    tot_slots = len(ch_meta['slots'])
    tot_chan = 0
    tot_active = 0
    tot_spare = 0

    for s_num in ch_meta['slots']:
        s_key = f"{ch}S{s_num}"
        info = get_slot_summary_info(s_key, ch, s_num, slot_desc_map, slot_rows)
        
        tot_chan += info['total_chan']
        tot_active += info['active']
        tot_spare += info['spares']

        row_vals = [
            s_num,
            s_key,
            info['clean_card'],
            info['desc'],
            info['sig'],
            info['rtb'],
            info['dest_str'],
            info['total_chan'] if info['total_chan'] > 0 else "-",
            info['active'] if info['active'] > 0 else "-",
            info['spares'] if info['spares'] > 0 else "-",
            info['status']
        ]
        ws.append(row_vals)
        ws.row_dimensions[r_num].height = 20
        fill_col = LIGHT_BG_1 if r_num % 2 == 1 else LIGHT_BG_2

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_num, column=col_idx)
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill(start_color=fill_col, end_color=fill_col, fill_type="solid")
            cell.border = thin_border()
            
            # Alignments
            if col_idx in [1, 2, 8, 9, 10, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [3, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

            # Slot_Tag hyperlink to individual slot sheet
            if col_idx == 2:
                cell.hyperlink = f"#'{s_key}'!A1"
                cell.font = Font(name="Arial", size=9, bold=True, color="1E40AF", underline="single")

            # Status styling
            if col_idx == 11:
                if info['status'] == "ACTIVE":
                    cell.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
                elif info['status'] == "SPARE RESERVE":
                    cell.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=False, color=SPARE_FONT)
                else:
                    cell.fill = PatternFill(start_color=COMMON_FILL, end_color=COMMON_FILL, fill_type="solid")
                    cell.font = Font(name="Arial", size=8.5, color=COMMON_FONT)

        r_num += 1

    # Total Summary Row
    ws.append([]) # Blank spacer
    r_total = r_num + 1
    total_row_vals = [
        "TOTAL",
        f"{tot_slots} Slots",
        "-",
        f"Chassis {ch} Total Configured Capacity",
        "-",
        "-",
        "-",
        tot_chan,
        tot_active,
        tot_spare,
        f"{round(tot_active/tot_chan*100 if tot_chan else 0)}% Active"
    ]
    ws.append(total_row_vals)
    ws.row_dimensions[r_total].height = 22
    for col_idx, val in enumerate(total_row_vals, start=1):
        cell = ws.cell(row=r_total, column=col_idx)
        cell.font = Font(name="Arial", size=9, bold=True, color="0F172A")
        cell.fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
        cell.border = summary_total_border()
        if col_idx in [1, 2, 8, 9, 10, 11]:
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2, 3, 4]: continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

def build_slot_sheet(ws, s_key, ch, sl_num, ch_meta, card_term_map, slot_desc_map, slot_rows, slot_term_lookup):
    """
    Builds an individual slot terminal wiring schedule sheet (e.g. C1S0, C1S4, C2S11, etc.)
    with dedicated Terminal_Number (Px-TBDIx-m) column.
    """
    ws.views.sheetView[0].showGridLines = True
    s_desc_info = slot_desc_map.get(s_key, {})
    rows = slot_rows.get(s_key, [])
    t_lookup = slot_term_lookup.get(s_key, {})
    
    if rows:
        raw_card = rows[0][3] or rows[0][26]
        raw_io = rows[0][20] or ""
    else:
        raw_card = s_desc_info.get('card_desc', '1756-N2')
        raw_io = s_desc_info.get('io_type', 'EMPTY')
        
    if not raw_card:
        raw_card = s_desc_info.get('card_desc', '1756-N2')

    if '-' in str(raw_card):
        p = str(raw_card).split('-')
        clean_card = f"{p[0]}-{p[1]}" if (len(p) >= 2 and not p[1].isdigit()) else str(raw_card)
    else:
        clean_card = str(raw_card)
        
    enclosure = ch_meta['location']
    default_tb_prefix = SLOT_TB_MAP.get(s_key, f"P{ch[1]}-TB")
    
    # Determine Card Type and Terminal Count
    if 'L950' in clean_card or 'CPU' in str(raw_io).upper() or s_key == 'C1S0':
        card_type_title = "CENTRAL PROCESSING UNIT (1756-L950TPSXT CONTROLLOGIX 5580 CPU)"
        max_terminals = 5
        card_lookup_key = None
        clean_card = "1756-L950TPSXT"
    elif 'EN4TR' in clean_card or 'COMM' in str(raw_io).upper() or 'ETH' in str(raw_io).upper() or s_key.endswith('S0'):
        card_type_title = "ETHERNET COMMUNICATION ADAPTER (1756-EN4TR DLR 2-PORT GIGABIT)"
        max_terminals = 4
        card_lookup_key = None
        clean_card = "1756-EN4TR"
    elif clean_card == '1756-IB32':
        card_type_title = "DIGITAL INPUT CARD (1756-IB32 32-CH 24VDC SINK/SOURCE)"
        max_terminals = 36
        card_lookup_key = '1756-IB32'
    elif clean_card == '1756-OB32':
        card_type_title = "DIGITAL OUTPUT CARD (1756-OB32 32-CH 24VDC SOURCE)"
        max_terminals = 36
        card_lookup_key = '1756-OB32'
    elif clean_card == '1756-IF16':
        card_type_title = "ANALOG INPUT CARD (1756-IF16 16-CH CURRENT/VOLTAGE)"
        max_terminals = 36
        card_lookup_key = '1756-IF16'
    elif clean_card == '1756-OF8':
        card_type_title = "ANALOG OUTPUT CARD (1756-OF8 8-CH CURRENT/VOLTAGE)"
        max_terminals = 20
        card_lookup_key = '1756-OF8'
    else:
        card_type_title = f"RESERVE / EMPTY SLOT ({clean_card})"
        max_terminals = 0
        card_lookup_key = None
        
    # Sheet Header Block (Columns A to J)
    ws.merge_cells("A1:J1")
    h1 = ws["A1"]
    h1.value = f"SLOT CONFIGURATION & TERMINAL SCHEDULE --- {s_key}"
    h1.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    h1.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    h1.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 25
    
    ws.merge_cells("A2:J2")
    h2 = ws["A2"]
    h2.value = f"{card_type_title}  |  LOCATION: {enclosure.upper()}"
    h2.font = Font(name="Arial", size=9, bold=True, color=WHITE)
    h2.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    h2.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 19
    
    # Meta row: Chassis, Slot, Module, Destination
    if rows:
        dests = sorted(list(set(str(r[1]).strip() for r in rows if r[1])))
        dest_str = ", ".join(dests)
    else:
        dest_str = "CA1" if ch in ['C1', 'C2'] else "Field RIO"
        
    if s_key == 'C1S0':
        dest_str = "CA1 (Main Control Room)"
    elif s_key.endswith('S0'):
        dest_str = "CA1 / Field RIO (EtherNet/IP DLR Trunk)"
        
    ws.merge_cells("A3:J3")
    h3 = ws["A3"]
    h3.value = f"CHASSIS: {ch}   |   SLOT: {sl_num}   |   MODULE: {clean_card}   |   DESTINATION AREA: {dest_str}"
    h3.font = Font(name="Arial", size=8.5, bold=True, color="334155")
    h3.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    h3.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[3].height = 18
    
    # Return Links: to CxSlotConfig AND to 00_Slot_Index
    ws.merge_cells("A4:D4")
    ws["A4"].value = f"← Back to Chassis Summary ({ch}SlotConfig)"
    ws["A4"].hyperlink = f"#'{ch}SlotConfig'!A1"
    ws["A4"].font = Font(name="Arial", size=8.5, bold=True, color="1E40AF", underline="single")
    
    ws.merge_cells("E4:J4")
    ws["E4"].value = "← Back to Master Slot Directory (00_Slot_Index)"
    ws["E4"].hyperlink = "#'00_Slot_Index'!A1"
    ws["E4"].font = Font(name="Arial", size=8.5, italic=True, color="475569", underline="single")
    ws["E4"].alignment = Alignment(horizontal="right", vertical="center")
    ws.row_dimensions[4].height = 16

    # Column Headers - Row 5 (10 Columns)
    sheet_headers = [
        "Terminal_No",
        "Terminal_Description",
        "Terminal_Number (Px-TBDIx-m)",
        "PLC_Tag_Side",
        "Terminal_Tag_Side",
        "Destination_Area",
        "PLC_Tag",
        "Instrument_Tag",
        "Instrument_Description",
        "Channel_Status"
    ]
    ws.append(sheet_headers)
    ws.row_dimensions[5].height = 22
    for col_idx, h in enumerate(sheet_headers, start=1):
        cell = ws.cell(row=5, column=col_idx)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border()

    r_data_start = 6
    
    # Branch 1: CPU Module (1756-L950TPSXT)
    if 'L950' in clean_card or 'CPU' in str(raw_io).upper() or s_key == 'C1S0':
        target_dest = dest_str if dest_str != "Reserve / Unused" else "CA1"
        cpu_rows = [
            ("RJ45-1", "Embedded 1Gbps EtherNet/IP Port (SCADA / Enterprise Network)", 
             "ETH-CORE-01", f"{s_key}-ETH1/SW-CORE-01", f"SW-CORE-01/{s_key}-ETH1", 
             target_dest, "CPU_Embedded_ETH", "PLC-01", 
             "Main Process Controller EtherNet/IP Uplink Port", "ACTIVE"),
            ("USB-1", "USB 2.0 Type-B Programming Port (Maintenance / Diagnostics)", 
             "USB-CONSOLE", f"{s_key}-USB/CONSOLE", f"CONSOLE/{s_key}-USB", 
             target_dest, "CPU_USB_Console", "USB-PGM", 
             "Engineering Workstation Maintenance Access Port", "ACTIVE"),
            ("SD-1", "1784-SD2 Industrial SD Memory Card Slot (Firmware & Program Backup)", 
             "SD-SLOT", f"{s_key}-SD/INTERNAL", f"INTERNAL/{s_key}-SD", 
             target_dest, "CPU_SD_Card", "1784-SD2", 
             "Non-volatile System Image & User Program Storage", "ACTIVE"),
            ("ESM-1", "Capacitor Energy Storage Module (Maintenance-free, Battery-free)", 
             "ESM-CAP", f"{s_key}-ESM/INTERNAL", f"INTERNAL/{s_key}-ESM", 
             target_dest, "CPU_ESM_Module", "1756-ESMCAP", 
             "Energy Storage Module for Power Loss Data Retention", "ACTIVE"),
            ("DISP-1", "4-Character Alphanumeric Diagnostic Display & Status LEDs (RUN, FORCE, SD, OK)", 
             "DIAG-DISP", f"{s_key}-DISP/DIAG", f"DIAG/{s_key}-DISP", 
             target_dest, "CPU_Status_Display", "DIAG-LED", 
             "System Controller Status & Fault Diagnostic Indicators", "ACTIVE")
        ]
        for row_vals in cpu_rows:
            ws.append(row_vals)
            ws.row_dimensions[r_data_start].height = 19
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_data_start, column=col_idx)
                cell.font = Font(name="Arial", size=9)
                cell.fill = PatternFill(start_color=LIGHT_BG_1, end_color=LIGHT_BG_1, fill_type="solid")
                cell.border = thin_border()
                if col_idx in [1, 2, 3, 6, 8, 10]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                if col_idx == 10:
                    cell.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
            r_data_start += 1

    # Branch 2: Communication Adapters (1756-EN4TR)
    elif 'EN4TR' in clean_card or 'COMM' in str(raw_io).upper() or 'ETH' in str(raw_io).upper() or s_key.endswith('S0'):
        target_dest = dest_str if dest_str != "Reserve / Unused" else ("CA1" if ch in ['C1', 'C2'] else "Field RIO")
        
        # Check if source IO List has specific tagging for this slot (like C2S0)
        tag_plc_p1 = f"{s_key}-P1/ETH"
        tag_term_p1 = f"ETH/{s_key}-P1"
        tag_tb_p1 = "ETH-DLR-P1"
        if rows and rows[0][44] and str(rows[0][44]).strip() not in ['-', '#N/A', '']:
            tag_tb_p1 = str(rows[0][44]).strip()
        if rows and rows[0][45] and str(rows[0][45]).strip() not in ['-', '#N/A', '']:
            tag_plc_p1 = str(rows[0][45]).strip()
        if rows and rows[0][46] and str(rows[0][46]).strip() not in ['-', '#N/A', '']:
            tag_term_p1 = str(rows[0][46]).strip()
            
        comm_rows = [
            ("RJ45-1", "EtherNet/IP Port 1 (DLR Ring Node 1 / Primary Adapter Link)", 
             tag_tb_p1, tag_plc_p1, tag_term_p1, target_dest, f"{s_key}_ETH_Port1", "ETH-01", 
             "10/100/1000 Mbps RJ45 Device Level Ring Port 1", "ACTIVE"),
            ("RJ45-2", "EtherNet/IP Port 2 (DLR Ring Node 2 / Redundant Loop)", 
             "ETH-DLR-P2", f"{s_key}-P2/ETH", f"ETH/{s_key}-P2", target_dest, f"{s_key}_ETH_Port2", "ETH-02", 
             "10/100/1000 Mbps RJ45 Device Level Ring Port 2", "ACTIVE"),
            ("USB-1", "USB 2.0 Type-B Device Port (Local Configuration & Diagnostics)", 
             "USB-CONSOLE", f"{s_key}-USB/CONSOLE", f"CONSOLE/{s_key}-USB", target_dest, f"{s_key}_USB_Config", "USB-CFG", 
             "Local IP Configuration and Firmware Flashing Port", "ACTIVE"),
            ("DISP-1", "4-Character Diagnostic Display & Status LEDs (OK, NET A/B, LINK 1/2)", 
             "DIAG-DISP", f"{s_key}-DISP/DIAG", f"DIAG/{s_key}-DISP", target_dest, f"{s_key}_Status_Display", "DIAG-LED", 
             "Hardware Status & Network Link Diagnostic Indicators", "ACTIVE")
        ]
        for row_vals in comm_rows:
            ws.append(row_vals)
            ws.row_dimensions[r_data_start].height = 19
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_data_start, column=col_idx)
                cell.font = Font(name="Arial", size=9)
                cell.fill = PatternFill(start_color=LIGHT_BG_1, end_color=LIGHT_BG_1, fill_type="solid")
                cell.border = thin_border()
                if col_idx in [1, 2, 3, 6, 8, 10]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                if col_idx == 10:
                    cell.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                    cell.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
            r_data_start += 1

    # Branch 3: Standard Terminal-Based Cards (IB32, OB32, IF16, OF8)
    elif max_terminals > 0:
        term_map = card_term_map.get(card_lookup_key, {})
        
        for t_num in range(1, max_terminals + 1):
            t_desc = term_map.get(t_num, f"PIN-{t_num}")
            
            # Check if we have active/spare signal assigned in IO List
            if t_num in t_lookup:
                r = t_lookup[t_num]
                # Col 45 in Excel is Terminal Tagging-Item (e.g. P1-TBDI1-1)
                tb_item = str(r[44]).strip() if (r[44] is not None and str(r[44]).strip() not in ['-', '#N/A', '']) else f"{default_tb_prefix}-{t_num}"
                tag_plc = r[45] if r[45] is not None else f"{s_key}-{t_num}/{tb_item}"
                tag_term = r[46] if r[46] is not None else f"{tb_item}/{s_key}-{t_num}"
                dest = r[1] if r[1] is not None else dest_str
                plc_tag = r[19] if r[19] is not None else (r[6] if r[6] is not None else "-")
                inst_tag = r[11] if r[11] is not None else "-"
                inst_desc = r[12] if r[12] is not None else "-"
                
                # Status determination
                is_spare = ('SPARE' in str(inst_tag).upper() or not inst_tag or inst_tag == '#N/A' or inst_tag == '0')
                status = "SPARE" if is_spare else "ACTIVE"
                if is_spare:
                    inst_tag = "Spare"
                    inst_desc = "Spare Reserve Channel"
            else:
                # Pin not in IO List: Power Common, RTN, Ground or Unwired Pin / Installed Spare Card
                dest = dest_str.split(',')[0] if dest_str else ("CA1" if ch in ['C1', 'C2'] else "Field RIO")
                is_power = any(p in t_desc.upper() for p in ['GND', 'DC', 'RTN', 'COM', 'VOUT'])
                if is_power:
                    status = "COMMON"
                    tb_item = "0VDC" if ('GND' in t_desc.upper() or 'RTN' in t_desc.upper()) else "+24VDC"
                    tag_plc = tb_item
                    tag_term = tb_item
                    plc_tag = "Internal Power Common"
                    inst_tag = "-"
                    inst_desc = "Module Power / Common Return Terminal"
                else:
                    # Spare channel on an installed card (e.g. C2S11, C2S12 IF16 spare card)
                    status = "SPARE"
                    tb_item = f"{default_tb_prefix}-{t_num}"
                    tag_plc = f"{s_key}-{t_num}/{tb_item}"
                    tag_term = f"{tb_item}/{s_key}-{t_num}"
                    if 'IF' in clean_card:
                        plc_tag = f"Spare_AI_{t_num}"
                    elif 'IB' in clean_card:
                        plc_tag = f"Spare_DI_{t_num}"
                    elif 'OB' in clean_card:
                        plc_tag = f"Spare_DO_{t_num}"
                    else:
                        plc_tag = f"Spare_AO_{t_num}"
                    inst_tag = "Spare"
                    inst_desc = "Spare Reserve Channel"
                    
            row_vals = [
                t_num,
                t_desc,
                tb_item,
                tag_plc,
                tag_term,
                dest,
                plc_tag,
                inst_tag,
                inst_desc,
                status
            ]
            ws.append(row_vals)
            ws.row_dimensions[r_data_start].height = 19
            
            # Styling row
            bg_col = LIGHT_BG_1 if t_num % 2 == 1 else LIGHT_BG_2
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws.cell(row=r_data_start, column=col_idx)
                cell.font = Font(name="Arial", size=9)
                cell.fill = PatternFill(start_color=bg_col, end_color=bg_col, fill_type="solid")
                cell.border = thin_border()
                
                # Alignments
                if col_idx in [1, 2, 3, 6, 8, 10]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                    
                # Status styling
                if col_idx == 10:
                    if status == "ACTIVE":
                        cell.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                        cell.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
                    elif status == "SPARE":
                        cell.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                        cell.font = Font(name="Arial", size=9, bold=False, color=SPARE_FONT)
                    else:
                        cell.fill = PatternFill(start_color=COMMON_FILL, end_color=COMMON_FILL, fill_type="solid")
                        cell.font = Font(name="Arial", size=8.5, color=COMMON_FONT)
            r_data_start += 1

    # Branch 4: Empty / Reserve Slot (1756-N2)
    else:
        row_vals = [
            "-", "Empty Slot / Reserve (1756-N2)", "-", "-", "-", "-", 
            "-", "-", "Unpopulated Slot Position", "EMPTY"
        ]
        ws.append(row_vals)
        ws.row_dimensions[r_data_start].height = 19
        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=r_data_start, column=col_idx)
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill(start_color=LIGHT_BG_1, end_color=LIGHT_BG_1, fill_type="solid")
            cell.border = thin_border()
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2, 3, 4]: continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 11)

def build_slot_workbook():
    card_term_map, slot_desc_map, slot_rows, slot_term_lookup = load_source_data()
    
    wb_out = openpyxl.Workbook()
    wb_out.remove(wb_out.active)  # remove default sheet
    
    # -------------------------------------------------------------
    # Sheet 0: 00_Slot_Index (Master Navigation Directory)
    # -------------------------------------------------------------
    ws_idx = wb_out.create_sheet(title="00_Slot_Index")
    ws_idx.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws_idx.merge_cells("A1:I1")
    t1 = ws_idx["A1"]
    t1.value = "KALASIN ENGINEERING xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)"
    t1.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    t1.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws_idx.row_dimensions[1].height = 26
    
    ws_idx.merge_cells("A2:I2")
    t2 = ws_idx["A2"]
    t2.value = "CHASSIS C1 - C4 (13-SLOT STANDARD) & C5: MASTER SLOT CONFIGURATION DIRECTORY"
    t2.font = Font(name="Arial", size=10, bold=True, color=WHITE)
    t2.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws_idx.row_dimensions[2].height = 20

    # Section 1: Chassis Summary Jump Table
    ws_idx.merge_cells("A4:I4")
    s1 = ws_idx["A4"]
    s1.value = "1. CHASSIS HARDWARE OVERVIEW DIRECTORY (CxSlotConfig)"
    s1.font = Font(name="Arial", size=10, bold=True, color="1E3A8A")
    s1.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
    s1.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_idx.row_dimensions[4].height = 22

    ch_headers = [
        "Chassis_Code", "Chassis_Overview_Sheet", "Chassis_Rack_Description", 
        "Physical_Location", "Hardware_Chassis_Model", "Mounted_Slots", 
        "Total_I/O_Channels", "Active_Channels", "Spare_Channels"
    ]
    ws_idx.append(ch_headers) # Row 5
    ws_idx.row_dimensions[5].height = 21
    for col_idx, h in enumerate(ch_headers, start=1):
        cell = ws_idx.cell(row=5, column=col_idx)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border()

    r_ch_num = 6
    for ch in ['C1', 'C2', 'C3', 'C4', 'C5']:
        ch_meta = CHASSIS_META[ch]
        cfg_sheet_name = f"{ch}SlotConfig"
        
        # Calculate chassis channel stats
        c_tot = 0
        c_act = 0
        c_spr = 0
        for s_num in ch_meta['slots']:
            s_key = f"{ch}S{s_num}"
            s_info = get_slot_summary_info(s_key, ch, s_num, slot_desc_map, slot_rows)
            c_tot += s_info['total_chan']
            c_act += s_info['active']
            c_spr += s_info['spares']

        ch_vals = [
            ch,
            cfg_sheet_name,
            ch_meta['name'],
            ch_meta['location'],
            ch_meta['model'],
            f"{len(ch_meta['slots'])} Slots (S0..S{max(ch_meta['slots'])})",
            c_tot,
            c_act,
            c_spr
        ]
        ws_idx.append(ch_vals)
        ws_idx.row_dimensions[r_ch_num].height = 20
        fill_col = LIGHT_BG_1 if r_ch_num % 2 == 1 else LIGHT_BG_2
        for col_idx, val in enumerate(ch_vals, start=1):
            cell = ws_idx.cell(row=r_ch_num, column=col_idx)
            cell.font = Font(name="Arial", size=9)
            cell.fill = PatternFill(start_color=fill_col, end_color=fill_col, fill_type="solid")
            cell.border = thin_border()
            if col_idx in [1, 2, 5, 6, 7, 8, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)

            # Hyperlink to CxSlotConfig sheet
            if col_idx == 2:
                cell.hyperlink = f"#'{cfg_sheet_name}'!A1"
                cell.font = Font(name="Arial", size=9, bold=True, color="1E40AF", underline="single")
        r_ch_num += 1

    # Section 2: Detailed Slot Directory
    ws_idx.append([]) # Spacer row
    r_s2_hdr = r_ch_num + 1
    ws_idx.merge_cells(f"A{r_s2_hdr}:I{r_s2_hdr}")
    s2 = ws_idx[f"A{r_s2_hdr}"]
    s2.value = "2. MASTER SLOT SCHEDULE & TERMINAL WIRING DIRECTORY (Click Slot Tag to open Sheet)"
    s2.font = Font(name="Arial", size=10, bold=True, color=SRC_HDR_FONT)
    s2.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
    s2.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws_idx.row_dimensions[r_s2_hdr].height = 22

    idx_headers = [
        "Slot_Tag", "Chassis", "Slot_No", "Hardware_Card", 
        "Signal_Family", "Terminals", "Destination_Area(s)", "Active_Points", "Spare_Points"
    ]
    ws_idx.append(idx_headers)
    r_hdr2 = r_s2_hdr + 1
    ws_idx.row_dimensions[r_hdr2].height = 22
    for col_idx, h in enumerate(idx_headers, start=1):
        cell = ws_idx.cell(row=r_hdr2, column=col_idx)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = header_border()
        
    r_idx_num = r_hdr2 + 1
    for ch in ['C1', 'C2', 'C3', 'C4', 'C5']:
        for s_num in CHASSIS_META[ch]['slots']:
            s_key = f"{ch}S{s_num}"
            info = get_slot_summary_info(s_key, ch, s_num, slot_desc_map, slot_rows)
            
            row_vals = [
                s_key, ch, s_num, info['clean_card'], info['sig'], 
                info['total_chan'] if info['total_chan'] > 0 else "-", 
                info['dest_str'], 
                info['active'] if info['active'] > 0 else "-", 
                info['spares'] if info['spares'] > 0 else "-"
            ]
            ws_idx.append(row_vals)
            ws_idx.row_dimensions[r_idx_num].height = 19
            fill_col = LIGHT_BG_1 if r_idx_num % 2 == 0 else LIGHT_BG_2
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws_idx.cell(row=r_idx_num, column=col_idx)
                cell.font = Font(name="Arial", size=9)
                cell.fill = PatternFill(start_color=fill_col, end_color=fill_col, fill_type="solid")
                cell.border = thin_border()
                if col_idx in [1, 2, 3, 6, 8, 9]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx in [4, 5]:
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                    
                # Add hyperlink from Index to target Slot Sheet
                if col_idx == 1:
                    cell.hyperlink = f"#'{s_key}'!A1"
                    cell.font = Font(name="Arial", size=9, bold=True, color="1E40AF", underline="single")
            r_idx_num += 1

    # Auto-adjust column widths for Index sheet
    for col in ws_idx.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.row in [1, 2, 4, r_s2_hdr]: continue
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws_idx.column_dimensions[col_letter].width = max(max_len + 4, 12)

    # -------------------------------------------------------------
    # Chassis Groups: CxSlotConfig followed by individual slot sheets
    # -------------------------------------------------------------
    total_generated_sheets = 1
    for ch in ['C1', 'C2', 'C3', 'C4', 'C5']:
        ch_meta = CHASSIS_META[ch]
        
        # Step A: Create CxSlotConfig sheet immediately before CxS0
        cfg_sheet_name = f"{ch}SlotConfig"
        ws_cfg = wb_out.create_sheet(title=cfg_sheet_name)
        build_chassis_config_sheet(ws_cfg, ch, ch_meta, slot_desc_map, slot_rows)
        total_generated_sheets += 1
        
        # Step B: Create individual slot sheets for this chassis (CxS0, CxS1, ...)
        for s_num in ch_meta['slots']:
            s_key = f"{ch}S{s_num}"
            ws_slot = wb_out.create_sheet(title=s_key)
            build_slot_sheet(ws_slot, s_key, ch, s_num, ch_meta, card_term_map, slot_desc_map, slot_rows, slot_term_lookup)
            total_generated_sheets += 1

    wb_out.save(OUTPUT_FILE)
    print(f"Total sheets generated in workbook: {total_generated_sheets}")
    print(f"Successfully generated: {OUTPUT_FILE}")

if __name__ == "__main__":
    build_slot_workbook()
