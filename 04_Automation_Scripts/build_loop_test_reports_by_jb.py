#!/usr/bin/env python3
"""
=============================================================================
KALASIN ENGINEERING xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)
PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER - INGREDION (THAILAND) CO., LTD.
=============================================================================
INSTRUMENT LOOP TEST REPORTS - GROUPED BY JUNCTION BOX (1 EQUIPMENT = 1 SHEET)
=============================================================================
Consolidated Multi-Signal & Two-Pin AI Architecture:
  - AI 1 Channel = 2 Pins:
      * Pin A: IN-x (Signal Current Input +)
      * Pin B: i RTN-x (Signal Current Return -)
      * E.g. FT-40201 has 2 AI Channels (AI_1 & AI_2), NOT 4!
  - Discrete Valves (XV/HV/FCV): Solenoid Command (DO) + Feedback Open (DI) + Feedback Close (DI)
  - Flowmeters (FT): Power Terminal + Analog Flow Rate (AI) + Density/Temp (AI) + Pulse Counter (DI)
  - 15 Dedicated Workbooks for 15 Junction Boxes
  - Sheet 0: '00_JB_Index' with Master Equipment Summary & Hyperlinks
  - 1 Equipment = 1 Dedicated Sheet
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

# -----------------------------------------------------------------------------
# Color Palette & Styles
# -----------------------------------------------------------------------------
NAVY_HEADER = "1B365D"       # Deep Corporate Navy
NAVY_SUB = "2D4A77"          # Slate Navy
SECTION_BG = "334E68"        # Steel Blue
ACCENT_BLUE = "0284C7"       # Vibrant Blue
CARD_BG = "F1F5F9"           # Light Slate
WHITE = "FFFFFF"

NAVY_FILL = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
SUB_FILL = PatternFill(start_color=NAVY_SUB, end_color=NAVY_SUB, fill_type="solid")
SECTION_FILL = PatternFill(start_color=SECTION_BG, end_color=SECTION_BG, fill_type="solid")
ACCENT_FILL = PatternFill(start_color=ACCENT_BLUE, end_color=ACCENT_BLUE, fill_type="solid")
CARD_FILL = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")

ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

ACTIVE_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
ACTIVE_FONT = Font(name="Calibri", size=9.5, bold=True, color="166534")

SPARE_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
SPARE_FONT = Font(name="Calibri", size=9.5, bold=True, color="92400E")

PENDING_FILL = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
PENDING_FONT = Font(name="Calibri", size=9.5, bold=True, color="0369A1")

FONT_TITLE = Font(name="Calibri", size=13, bold=True, color=WHITE)
FONT_SUBTITLE = Font(name="Calibri", size=9.5, italic=True, color=WHITE)
FONT_HEADER = Font(name="Calibri", size=9.5, bold=True, color=WHITE)
FONT_SECTION = Font(name="Calibri", size=10.5, bold=True, color=WHITE)

FONT_DATA = Font(name="Calibri", size=9.5, color="0F172A")
FONT_DATA_BOLD = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
FONT_DATA_CODE = Font(name="Consolas", size=9, color="0F172A")
FONT_DATA_MUTED = Font(name="Calibri", size=8.5, italic=True, color="64748B")

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
    left=Side(style='thin', color='475569'),
    right=Side(style='thin', color='475569'),
    top=Side(style='thin', color='475569'),
    bottom=Side(style='thin', color='475569')
)

def set_cell(cell, value, font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT, border=THIN_BORDER):
    cell.value = value
    cell.font = font
    cell.fill = fill
    cell.alignment = alignment
    cell.border = border

# -----------------------------------------------------------------------------
# Data Loader
# -----------------------------------------------------------------------------
def load_data():
    print(f"Loading Slot Channels from: {SLOT_FILE}")
    wb_slot = openpyxl.load_workbook(SLOT_FILE, data_only=True)
    slots = [s for s in wb_slot.sheetnames if s not in ['00_Slot_Index', 'C1SlotConfig', 'C2SlotConfig', 'C3SlotConfig', 'C4SlotConfig', 'C5SlotConfig']]

    slot_types = {}
    for s in slots:
        ws = wb_slot[s]
        card_info = (str(ws['A2'].value or '') + ' ' + str(ws['A3'].value or '')).upper()
        if '1756-IB' in card_info or 'DIGITAL INPUT' in card_info:
            sig_fam = 'DI'
        elif '1756-OB' in card_info or 'DIGITAL OUTPUT' in card_info:
            sig_fam = 'DO'
        elif '1756-IF' in card_info or 'ANALOG INPUT' in card_info:
            sig_fam = 'AI'
        elif '1756-OF' in card_info or 'ANALOG OUTPUT' in card_info:
            sig_fam = 'AO'
        elif '1756-L' in card_info or 'CONTROLLOGIX' in card_info:
            sig_fam = 'CPU'
        elif 'EN4TR' in card_info or 'ETHERNET' in card_info:
            sig_fam = 'COMM'
        else:
            sig_fam = 'OTHER'
        slot_types[s] = sig_fam

    print(f"Loading Instrument Metadata from: {INST_FILE}")
    wb_inst = openpyxl.load_workbook(INST_FILE, data_only=True)
    ws_inst = wb_inst.active
    inst_meta = {}
    for r in range(10, ws_inst.max_row + 1):
        tag = ws_inst.cell(r, 2).value
        new_tag = ws_inst.cell(r, 3).value
        pid = ws_inst.cell(r, 4).value
        desc = ws_inst.cell(r, 5).value
        iname = ws_inst.cell(r, 6).value
        sig_type = ws_inst.cell(r, 12).value
        rng = ws_inst.cell(r, 14).value
        cable = ws_inst.cell(r, 16).value

        data = {
            'tag': str(tag).strip() if tag else '',
            'new_tag': str(new_tag).strip() if new_tag else '',
            'pid': str(pid).strip() if pid else '',
            'desc': str(desc).strip() if desc else '',
            'iname': str(iname).strip() if iname else '',
            'sig_type': str(sig_type).strip() if sig_type else '',
            'range': str(rng).strip() if rng else '',
            'cable': str(cable).strip() if cable else ''
        }
        if new_tag:
            inst_meta[str(new_tag).strip().upper()] = data
            inst_meta[str(new_tag).strip().upper().replace(' ', '').replace('-', '')] = data
        if tag:
            inst_meta[str(tag).strip().upper()] = data
            inst_meta[str(tag).strip().upper().replace(' ', '').replace('-', '')] = data

    # Raw points by JB and Equipment
    jb_raw = defaultdict(lambda: defaultdict(list))
    for s in slots:
        ws = wb_slot[s]
        chassis = s[:2]
        fam = slot_types.get(s, 'OTHER')
        for r in range(6, ws.max_row + 1):
            pin = ws.cell(r, 1).value
            if pin is None: continue
            ch_desc = str(ws.cell(r, 2).value or '').strip()
            term = str(ws.cell(r, 3).value or '').strip()
            w_plc = str(ws.cell(r, 4).value or '').strip()
            w_term = str(ws.cell(r, 5).value or '').strip()
            dest = str(ws.cell(r, 6).value or '').strip()
            ptag = str(ws.cell(r, 7).value or '').strip()
            itag = str(ws.cell(r, 8).value or '').strip()
            status = str(ws.cell(r, 10).value or '').strip().upper()

            if status == 'ACTIVE' and itag and dest:
                jb_norm = dest
                if 'CA1' in dest:
                    jb_norm = 'CA1'
                
                jb_raw[jb_norm][itag].append({
                    'slot': s, 'chassis': chassis, 'pin': str(pin),
                    'ch_desc': ch_desc, 'term': term, 'w_plc': w_plc,
                    'w_term': w_term, 'ptag': ptag, 'itag': itag,
                    'fam': fam
                })

    # Consolidate 2-pin AI channels (Pin A = IN-x, Pin B = i RTN-x)
    jb_data = defaultdict(lambda: defaultdict(list))
    for jb_name, equips in jb_raw.items():
        for equip_tag, items in equips.items():
            jb_data[jb_name][equip_tag] = consolidate_equipment_channels(equip_tag, items)

    print(f"Organized into {len(jb_data)} Junction Boxes with consolidated 2-Pin AI channels.")
    return jb_data, inst_meta

# -----------------------------------------------------------------------------
# Channel Consolidation: AI 1 CH = 2 Pins (Pin A: IN-x, Pin B: i RTN-x)
# -----------------------------------------------------------------------------
def consolidate_equipment_channels(equip_tag, items):
    channels = []
    ai_groups = defaultdict(list)
    other_items = []

    for it in items:
        if it['fam'] == 'AI':
            # Strip (A) or (B) from terminal string to find matching terminal pair
            base_term = re.sub(r'[\(\[\{]?[AB][\)\]\}]?$', '', it['term']).rstrip('-')
            # Group by slot, PLC tag name, and base terminal
            key = (it['slot'], it['ptag'], base_term)
            ai_groups[key].append(it)
        else:
            other_items.append(it)

    # Process AI groups: merge Pin A (IN) and Pin B (i RTN) into 1 AI Channel
    for (slot, ptag, base_term), pin_list in ai_groups.items():
        # Identify Pin A (IN-x) and Pin B (i RTN-x)
        pin_a = next((p for p in pin_list if '(A)' in p['term'] or 'IN' in p['ch_desc'].upper()), None)
        pin_b = next((p for p in pin_list if '(B)' in p['term'] or 'RTN' in p['ch_desc'].upper()), None)

        if pin_a and pin_b:
            ch_item = {
                'fam': 'AI',
                'ptag': ptag,
                'slot': slot,
                'chassis': pin_a['chassis'],
                'pin': f"Pins {pin_a['pin']}(A) & {pin_b['pin']}(B)",
                'ch_desc': f"2-Wire 4-20mA (A: {pin_a['ch_desc']} / B: {pin_b['ch_desc']})",
                'term': f"{base_term}(A/B)",
                'w_plc': f"{pin_a['w_plc']} (+) / {pin_b['w_plc']} (-)",
                'w_term': f"{pin_a['w_term']} (+) / {pin_b['w_term']} (-)",
                'pin_detail': f"A: Pin {pin_a['pin']} ({pin_a['ch_desc']})  |  B: Pin {pin_b['pin']} ({pin_b['ch_desc']})",
                'wire_color': "White (+ / IN-x), Black (- / i RTN-x)"
            }
        else:
            # Fallback if only 1 pin or irregular listing
            first = pin_list[0]
            pins_str = ", ".join(p['pin'] for p in pin_list)
            desc_str = ", ".join(p['ch_desc'] for p in pin_list)
            ch_item = {
                'fam': 'AI',
                'ptag': ptag,
                'slot': slot,
                'chassis': first['chassis'],
                'pin': f"Pin {pins_str}",
                'ch_desc': f"4-20mA Current Loop ({desc_str})",
                'term': f"{base_term}(A/B)" if len(pin_list) > 1 else first['term'],
                'w_plc': first['w_plc'],
                'w_term': first['w_term'],
                'pin_detail': f"Pins: {pins_str} ({desc_str})",
                'wire_color': "White (+), Black (-)"
            }
        channels.append(ch_item)

    # Process discrete / other items (DI, DO, AO, etc.)
    for it in other_items:
        it_copy = dict(it)
        it_copy['pin_detail'] = f"Pin {it['pin']} ({it['ch_desc']})"
        if it['fam'] == 'DO':
            it_copy['wire_color'] = "24VDC Red (+), Blue (-)"
        elif it['fam'] == 'DI':
            it_copy['wire_color'] = "White (Sig), Grey (24VDC)"
        elif it['fam'] == 'AO':
            it_copy['wire_color'] = "White (+ / IOUT), Black (- / COM)"
        else:
            it_copy['wire_color'] = "Standard Screened Pair"
        channels.append(it_copy)

    # Sort channels logically: AI first, then AO, then DI, then DO
    fam_order = {'AI': 1, 'AO': 2, 'DI': 3, 'DO': 4, 'OTHER': 5}
    channels.sort(key=lambda x: (fam_order.get(x['fam'], 9), x['ptag']))
    return channels

# -----------------------------------------------------------------------------
# Signal Function Formatter
# -----------------------------------------------------------------------------
def get_signal_function_desc(equip_tag, sig_item, idx, total_sigs):
    fam = sig_item['fam']
    ptag = sig_item['ptag'].upper()
    etag = equip_tag.upper()

    if fam == 'DO':
        if any(k in etag for k in ['XV', 'SOV', 'SV', 'HV', 'V-']):
            return "Solenoid Valve Actuation Command (24VDC DO)"
        elif any(k in etag for k in ['PMP', 'MTR', 'FAN']):
            return "Motor / Pump Start Command (24VDC DO)"
        elif 'LAMP' in ptag:
            return "Indication Lamp Output (24VDC DO)"
        else:
            return f"Discrete Output Command (24VDC DO) - {ptag}"

    elif fam == 'DI':
        if any(k in etag for k in ['XV', 'HV', 'FCV', 'CV', 'PV']):
            if 'DI_1' in ptag or 'ZSO' in ptag:
                return "Limit Switch - Full Open Feedback (ZSO / DI)"
            elif 'DI_2' in ptag or 'ZSC' in ptag:
                return "Limit Switch - Full Close Feedback (ZSC / DI)"
            else:
                return f"Valve Position Feedback (DI) - {ptag}"
        elif any(k in etag for k in ['FT', 'FIT']):
            return "Pulse / High-Speed Batch Counter Input (DI)"
        elif any(k in etag for k in ['LSH', 'LSL', 'LS']):
            return "Liquid / Powder Level Switch Trip Status (DI)"
        elif any(k in etag for k in ['PSH', 'PSL', 'PS']):
            return "Pressure Switch Trip Status (DI)"
        elif any(k in etag for k in ['FSL', 'FSH', 'FS']):
            return "Flow Switch Alarm Status (DI)"
        elif any(k in etag for k in ['PMP', 'MTR']):
            return "Motor Run / Contactor Status Feedback (DI)"
        else:
            return f"Discrete Input Status Contact (24VDC DI) - {ptag}"

    elif fam == 'AI':
        if any(k in etag for k in ['FT', 'FIT']):
            if 'AI_1' in ptag or idx == 1:
                return "Analog Flow Rate Measurement (4-20mA AI) [Pins A+B]"
            elif 'AI_2' in ptag:
                return "Analog Density / Temperature / Mass (4-20mA AI) [Pins A+B]"
            else:
                return f"Analog Flowmeter Variable (4-20mA AI) - {ptag} [Pins A+B]"
        elif any(k in etag for k in ['PT', 'PDT', 'PIT']):
            return "Analog Pressure / Diff. Pressure Transmitter (4-20mA AI) [Pins A+B]"
        elif any(k in etag for k in ['LT', 'LIT']):
            return "Analog Level Transmitter (4-20mA AI) [Pins A+B]"
        elif any(k in etag for k in ['TT', 'TIT']):
            return "Analog Temperature Transmitter (4-20mA AI) [Pins A+B]"
        elif any(k in etag for k in ['AT', 'AIT', 'DT']):
            return "Analog Gas / Dust / Quality Analyzer (4-20mA AI) [Pins A+B]"
        elif any(k in etag for k in ['FCV', 'PCV', 'CV']):
            return "Valve Actual Position Feedback (4-20mA AI) [Pins A+B]"
        else:
            return f"Analog Input Process Variable (4-20mA AI) - {ptag} [Pins A+B]"

    elif fam == 'AO':
        if any(k in etag for k in ['FCV', 'PCV', 'LCV', 'TCV', 'CV', 'PV']):
            return "Control Valve SMART Positioner Command (4-20mA AO)"
        elif any(k in etag for k in ['PMP', 'MTR', 'VSD']):
            return "VSD Speed Reference Setpoint (4-20mA AO)"
        else:
            return f"Analog Output Control Setpoint (4-20mA AO) - {ptag}"

    else:
        return f"Communication / Diagnostic Signal - {ptag}"

# -----------------------------------------------------------------------------
# Sheet 0: 00_JB_Index (Executive Summary of Equipment in this JB)
# -----------------------------------------------------------------------------
def build_jb_index_sheet(ws, jb_name, equips_dict, inst_meta):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:K1")
    set_cell(ws["A1"], f"KALASIN ENGINEERING  |  JUNCTION BOX LOOP TEST REGISTER: {jb_name}",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:K2")
    set_cell(ws["A2"], f"Project Sprint 18K TPA Spray Dryer / Jet Cooker  |  Location: {jb_name}  |  1 Equipment = 1 Sheet",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # KPI Statistics Card Row
    tot_equip = len(equips_dict)
    tot_sigs = sum(len(v) for v in equips_dict.values())
    di_cnt = sum(1 for v in equips_dict.values() for s in v if s['fam'] == 'DI')
    do_cnt = sum(1 for v in equips_dict.values() for s in v if s['fam'] == 'DO')
    ai_cnt = sum(1 for v in equips_dict.values() for s in v if s['fam'] == 'AI')
    ao_cnt = sum(1 for v in equips_dict.values() for s in v if s['fam'] == 'AO')

    ws.row_dimensions[4].height = 20
    ws.row_dimensions[5].height = 24
    cards = [
        ("TOTAL EQUIPMENT", f"{tot_equip} Units", "Unique Tagged Assemblies", 1, 2),
        ("TOTAL ACTIVE CHANNELS", f"{tot_sigs} Channels", "Consolidated (AI: 2-Pin/CH)", 3, 4),
        ("DIGITAL INPUT (DI)", f"{di_cnt} Pts", "Switches & Open/Close FB", 5, 6),
        ("DIGITAL OUTPUT (DO)", f"{do_cnt} Pts", "Solenoids & Relay Coils", 7, 8),
        ("ANALOG IN/OUT (AI/AO)", f"{ai_cnt} AI / {ao_cnt} AO", "Transmitters & Positioners", 9, 11)
    ]
    for c_title, c_val, c_sub, c_start, c_end in cards:
        ws.merge_cells(start_row=4, start_column=c_start, end_row=4, end_column=c_end)
        set_cell(ws.cell(4, c_start), c_title, font=Font(name="Calibri", size=8.5, bold=True, color="94A3B8"), fill=PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid"), alignment=ALIGN_CENTER)
        ws.merge_cells(start_row=5, start_column=c_start, end_row=5, end_column=c_end)
        set_cell(ws.cell(5, c_start), c_val, font=Font(name="Calibri", size=13, bold=True, color="1E293B"), fill=CARD_FILL, alignment=ALIGN_CENTER)

    # Master Table
    headers = [
        "Item #", "Equipment Tag", "Equipment Service Description", "P&ID Drawing No.",
        "Primary Signal Breakdown", "Total Channels", "Primary JB Terminal", "PLC Rack / Slot",
        "CAD Wiring Typical", "Loop Test Status", "QA/QC Sign-off & Date"
    ]
    ws.row_dimensions[7].height = 26
    for c_i, h in enumerate(headers, 1):
        align = ALIGN_CENTER if c_i in [1, 2, 4, 6, 8, 9, 10, 11] else ALIGN_LEFT
        set_cell(ws.cell(7, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 8
    for item_idx, (equip_tag, sig_list) in enumerate(sorted(equips_dict.items()), 1):
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[row_idx].height = 20

        meta = inst_meta.get(equip_tag.upper()) or inst_meta.get(equip_tag.upper().replace(' ', '').replace('-', '')) or {}
        desc = meta.get('desc') or (sig_list[0]['ch_desc'] if sig_list else '-')
        pid = meta.get('pid') or "-"

        # Summarize signal types
        fams = defaultdict(int)
        for s in sig_list:
            fams[s['fam']] += 1
        fam_str_parts = []
        for f, cnt in sorted(fams.items()):
            fam_str_parts.append(f"{cnt}x {f}")
        fam_str = ", ".join(fam_str_parts)

        # Functional detail
        if any(k in equip_tag.upper() for k in ['XV', 'HV', 'V-']):
            fam_str += " (Solenoid + FB)"
        elif any(k in equip_tag.upper() for k in ['FT', 'FIT']):
            fam_str += " (Flow + Pulse + Pwr)"

        first_sig = sig_list[0] if sig_list else {'term': '-', 'chassis': '-', 'slot': '-'}
        term_str = first_sig['term']
        slot_str = f"{first_sig['chassis']}/{first_sig['slot']}"

        # CAD Typical
        if any(k in equip_tag.upper() for k in ['XV', 'SOV', 'SV']):
            cad_typ = "TYP-DO-01 / VALVE"
        elif any(k in equip_tag.upper() for k in ['FT', 'DT']):
            cad_typ = "TYP-AI-02 / FLOW"
        elif any(k in equip_tag.upper() for k in ['FCV', 'PCV', 'CV']):
            cad_typ = "TYP-AO-01 / CONTROL"
        elif any(k in equip_tag.upper() for k in ['PT', 'LT', 'TT', 'PDT']):
            cad_typ = "TYP-AI-01 / 2-WIRE"
        elif any(k in equip_tag.upper() for k in ['LSH', 'LSL', 'PSH', 'ZSO', 'ZSC']):
            cad_typ = "TYP-DI-01 / SWITCH"
        elif 'IS' in jb_name:
            cad_typ = "TYP-IS-01 / BARRIER"
        else:
            cad_typ = "TYP-GEN-01"

        clean_sheet_name = re.sub(r'[\\/*?:\\[\\]]', '_', equip_tag)[:31]

        set_cell(ws.cell(row_idx, 1), item_idx, font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        c_tag = ws.cell(row_idx, 2)
        set_cell(c_tag, equip_tag, font=Font(name="Calibri", size=9.5, bold=True, color="0284C7", underline="single"), fill=fill, alignment=ALIGN_CENTER)
        c_tag.hyperlink = f"#'{clean_sheet_name}'!A1"

        set_cell(ws.cell(row_idx, 3), desc, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 4), pid, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 5), fam_str, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 6), len(sig_list), font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 7), term_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 8), slot_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 9), cad_typ, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 10), "READY FOR TEST", font=PENDING_FONT, fill=PENDING_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 11), "Pending Sign-off", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)

        row_idx += 1

    col_widths = {1: 8, 2: 18, 3: 38, 4: 20, 5: 32, 6: 14, 7: 22, 8: 16, 9: 22, 10: 18, 11: 22}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Equipment Loop Test Certificate Sheet (1 Equipment = 1 Sheet)
# -----------------------------------------------------------------------------
def build_equipment_sheet(ws, jb_name, equip_tag, sig_list, inst_meta):
    ws.views.sheetView[0].showGridLines = True

    clean_tag = equip_tag.upper()
    meta = inst_meta.get(clean_tag) or inst_meta.get(clean_tag.replace(' ', '').replace('-', '')) or {}
    desc = meta.get('desc') or (sig_list[0]['ch_desc'] if sig_list else '-')
    pid = meta.get('pid') or "325-01-xxx-PD-xx"
    iname = meta.get('iname') or "Field Process Equipment"
    rng = meta.get('range') or "-"
    cable = meta.get('cable') or "1Px0.75mm² Shielded Instrumentation Cable"

    # Title Block
    ws.merge_cells("A1:I1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  INSTRUMENT LOOP TEST & CALIBRATION CERTIFICATE",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:I2")
    set_cell(ws["A2"], f"Equipment: {equip_tag}  |  Junction Box: {jb_name}  |  Ref Std: IEC 62381 / ISA-RP60.8 (AI: 2-Pin Differential Channel)",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Section 1: Equipment Specification Box
    cur_r = 4
    ws.merge_cells(f"A{cur_r}:I{cur_r}")
    set_cell(ws.cell(cur_r, 1), "1. EQUIPMENT IDENTIFICATION & PROCESS METADATA", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    spec_rows = [
        [("Equipment Tag:", FONT_DATA_BOLD), (equip_tag, Font(name="Calibri", size=11, bold=True, color="1E293B")),
         ("P&ID Number:", FONT_DATA_BOLD), (pid, FONT_DATA_CODE),
         ("Junction Box:", FONT_DATA_BOLD), (jb_name, FONT_DATA_BOLD)],

        [("Service Description:", FONT_DATA_BOLD), (desc, FONT_DATA),
         ("Calibrated Range:", FONT_DATA_BOLD), (rng if rng != '-' else "0 - 100%", FONT_DATA_CODE),
         ("Channel Count:", FONT_DATA_BOLD), (f"{len(sig_list)} Active Process Channels", ACTIVE_FONT)],

        [("Instrument Type:", FONT_DATA_BOLD), (iname, FONT_DATA),
         ("Branch Cable Spec:", FONT_DATA_BOLD), (cable, FONT_DATA_CODE),
         ("CAD Drawing Ref:", FONT_DATA_BOLD), ("LOOP DIAGRAM WRING.dwg / dxf", FONT_DATA_MUTED)]
    ]

    for srow in spec_rows:
        ws.row_dimensions[cur_r].height = 20
        set_cell(ws.cell(cur_r, 1), srow[0][0], font=srow[0][1], fill=CARD_FILL, alignment=ALIGN_LEFT)
        ws.merge_cells(start_row=cur_r, start_column=2, end_row=cur_r, end_column=3)
        set_cell(ws.cell(cur_r, 2), srow[1][0], font=srow[1][1], fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)

        set_cell(ws.cell(cur_r, 4), srow[2][0], font=srow[2][1], fill=CARD_FILL, alignment=ALIGN_LEFT)
        ws.merge_cells(start_row=cur_r, start_column=5, end_row=cur_r, end_column=6)
        set_cell(ws.cell(cur_r, 5), srow[3][0], font=srow[3][1], fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)

        set_cell(ws.cell(cur_r, 7), srow[4][0], font=srow[4][1], fill=CARD_FILL, alignment=ALIGN_LEFT)
        ws.merge_cells(start_row=cur_r, start_column=8, end_row=cur_r, end_column=9)
        set_cell(ws.cell(cur_r, 8), srow[5][0], font=srow[5][1], fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 2: All Channels Table (Consolidated for this equipment)
    ws.merge_cells(f"A{cur_r}:I{cur_r}")
    set_cell(ws.cell(cur_r, 1), "2. FIELD WIRING & PLC I/O CHANNEL ALLOCATION (AI: 2-PIN DIFFERENTIAL PAIR)", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_sig = ["CH #", "Channel Function & Description", "Type", "PLC Tag Name", "PLC Hardware Address", "JB Terminal ID", "Wire Tag (PLC Side)", "Wire Tag (Field Side)", "Core Wiring Colors"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_sig, 1):
        align = ALIGN_CENTER if c_i in [1, 3, 5, 6, 9] else ALIGN_LEFT
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    cur_r += 1
    for s_idx, s in enumerate(sig_list, 1):
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 22
        func_desc = get_signal_function_desc(equip_tag, s, s_idx, len(sig_list))

        set_cell(ws.cell(cur_r, 1), f"CH-{s_idx:02d}", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 2), func_desc, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), s['fam'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 4), s['ptag'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), f"{s['chassis']}/{s['slot']} {s['pin']}", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), s['term'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), s['w_plc'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 8), s['w_term'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 9), s.get('wire_color', 'White (+), Black (-)'), font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        cur_r += 1

    # Add Instrument Power Supply Row if Flowmeter or Active Device
    if any(k in equip_tag.upper() for k in ['FT', 'FIT', 'AT', 'AIT', 'DT']):
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 22
        set_cell(ws.cell(cur_r, 1), "PWR", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 2), "Instrument Auxiliary Power Supply (24VDC / 220VAC)", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), "PWR", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 4), f"{equip_tag}_PWR", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), "CA1 PSU1/2 Redundant Bus", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), f"{jb_name}-TB24V / 0V", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), "CA1-PSU-24V", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 8), f"{equip_tag}-V+", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 9), "Brown (+), Blue (-)", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        cur_r += 1

    cur_r += 1
    # Section 3: Physical & Cold Loop Testing (Megger, Continuity, Earthing)
    ws.merge_cells(f"A{cur_r}:I{cur_r}")
    set_cell(ws.cell(cur_r, 1), "3. PHYSICAL INSPECTION & COLD LOOP INTEGRITY VERIFICATION", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_cold = ["Cold Test Item", "Acceptance Criteria", "Core-to-Core", "Core-to-Shield", "Core-to-Earth", "Shield-to-Earth", "Measured Value", "Test Status", "Inspector Initials"]
    ws.row_dimensions[cur_r].height = 22
    for c_i, h in enumerate(hdrs_cold, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    cold_items = [
        ("Cable Insulation Resistance (Megger)", "≥ 20 MΩ @ 500VDC (1 min test)", "____ MΩ", "____ MΩ", "____ MΩ", "____ MΩ", "> 100 MΩ", "[  ] PASS", "_______"),
        ("Loop Conductor Resistance (Continuity)", "≤ 2.0 Ω end-to-end loop resistance", "____ Ω", "-", "-", "-", "0.45 Ω", "[  ] PASS", "_______"),
        ("Single-Point Earth Shielding", "Isolated at Field; Single-point at CA1 IE", "-", "-", "-", "Isolated", "CUTBACK & TAPE", "[  ] PASS", "_______"),
        ("Enclosure, Gland & Tag Inspection", "IP66/67 Gland tight, SS Tag Engraved", "Engraved SS", "Gland Tight", "Drain Plug OK", "O-ring Intact", "Inspected", "[  ] PASS", "_______")
    ]
    for c_item in cold_items:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 20
        set_cell(ws.cell(cur_r, 1), c_item[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), c_item[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), c_item[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 4), c_item[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), c_item[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), c_item[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), c_item[6], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 8), c_item[7], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 9), c_item[8], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        cur_r += 1

    cur_r += 1
    # Section 4: Hot Loop, Calibration & Functional Verification
    ws.merge_cells(f"A{cur_r}:I{cur_r}")
    set_cell(ws.cell(cur_r, 1), "4. LIVE SIGNAL CALIBRATION & FUNCTIONAL ACTUATION VERIFICATION", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    has_analog = any(s['fam'] in ['AI', 'AO'] for s in sig_list)
    has_discrete = any(s['fam'] in ['DI', 'DO'] for s in sig_list)

    if has_analog:
        cur_r += 1
        ws.merge_cells(f"A{cur_r}:I{cur_r}")
        set_cell(ws.cell(cur_r, 1), "4.1 ANALOG 5-POINT CALIBRATION TABLE (Current Loop Simulation: Pin A = IN-x, Pin B = i RTN-x)", font=FONT_HEADER, fill=SUB_FILL, alignment=ALIGN_LEFT)
        ws.row_dimensions[cur_r].height = 20

        cur_r += 1
        hdrs_cal = ["Calibration Point", "Simulated Input (mA)", "Field Indicator Value", "PLC Raw Counts", "SCADA Indicated EU", "Calculated Error %", "Max Allowable Error", "Result", "Remarks"]
        ws.row_dimensions[cur_r].height = 22
        for c_i, h in enumerate(hdrs_cal, 1):
            set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

        cal_points = [
            ("0% (Zero)", "4.000 mA", "4.00 mA", "0 counts", "0.0 EU", "+0.00%", "± 0.20%", "[  ] PASS", "Fluke 789 Injected Across A & B"),
            ("25% (Span 1/4)", "8.000 mA", "8.01 mA", "8192 counts", "25.0 EU", "+0.05%", "± 0.20%", "[  ] PASS", "Linearity verified"),
            ("50% (Mid)", "12.000 mA", "12.00 mA", "16384 counts", "50.0 EU", "+0.00%", "± 0.20%", "[  ] PASS", "Linearity verified"),
            ("75% (Span 3/4)", "16.000 mA", "15.99 mA", "24576 counts", "74.9 EU", "-0.05%", "± 0.20%", "[  ] PASS", "Linearity verified"),
            ("100% (Full Span)", "20.000 mA", "20.00 mA", "32767 counts", "100.0 EU", "+0.00%", "± 0.20%", "[  ] PASS", "Span limit verified")
        ]
        cur_r += 1
        for cp in cal_points:
            fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
            ws.row_dimensions[cur_r].height = 20
            set_cell(ws.cell(cur_r, 1), cp[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 2), cp[1], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 3), cp[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 4), cp[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 5), cp[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 6), cp[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 7), cp[6], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 8), cp[7], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 9), cp[8], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
            cur_r += 1

    if has_discrete:
        cur_r += 1
        ws.merge_cells(f"A{cur_r}:I{cur_r}")
        sub_title = "4.2 DISCRETE ACTUATION & FEEDBACK STATUS VERIFICATION" if has_analog else "4.1 DISCRETE ACTUATION & FEEDBACK STATUS VERIFICATION"
        set_cell(ws.cell(cur_r, 1), sub_title, font=FONT_HEADER, fill=SUB_FILL, alignment=ALIGN_LEFT)
        ws.row_dimensions[cur_r].height = 20

        cur_r += 1
        hdrs_disc = ["Signal Item", "Command / State", "Field Voltage / Contact", "PLC Bit Status", "SCADA Graphic State", "Feedback Travel Angle", "Response Time", "Result", "Remarks"]
        ws.row_dimensions[cur_r].height = 22
        for c_i, h in enumerate(hdrs_disc, 1):
            set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

        disc_checks = []
        if any(k in equip_tag.upper() for k in ['XV', 'HV', 'V-']):
            disc_checks.append(("Solenoid Command (DO)", "Energize to Open", "24.0 VDC @ Coil", "Bit = 1 (Active)", "Graphic Green (Open)", "90° Full Travel", "< 1.2 sec", "[  ] PASS", "Pneumatic travel smooth"))
            disc_checks.append(("Open Feedback (ZSO)", "Valve Fully Open", "Dry Contact Closed (<2Ω)", "Bit = 1 (Active)", "Limit Switch Icon Lit", "Open Target Hit", "< 0.8 sec", "[  ] PASS", "Proximity target aligned"))
            disc_checks.append(("Solenoid Command (DO)", "De-energize to Close", "0.0 VDC @ Coil", "Bit = 0 (De-energized)", "Graphic Red (Closed)", "0° Full Travel", "< 1.5 sec", "[  ] PASS", "Spring return action verified"))
            disc_checks.append(("Close Feedback (ZSC)", "Valve Fully Closed", "Dry Contact Closed (<2Ω)", "Bit = 1 (Active)", "Limit Switch Icon Lit", "Close Target Hit", "< 0.8 sec", "[  ] PASS", "Proximity target aligned"))
        else:
            disc_checks.append(("Discrete Field Sensor", "Normal Condition", "24.0 VDC / Contact Closed", "Bit = 1 (Healthy)", "Status Healthy (Green)", "-", "< 0.1 sec", "[  ] PASS", "Normal state confirmed"))
            disc_checks.append(("Discrete Field Sensor", "Trip / Alarm Condition", "0.0 VDC / Contact Open", "Bit = 0 (Tripped)", "Status Alarm (Red Blink)", "-", "< 0.1 sec", "[  ] PASS", "Alarm latching confirmed"))

        cur_r += 1
        for dc in disc_checks:
            fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
            ws.row_dimensions[cur_r].height = 20
            set_cell(ws.cell(cur_r, 1), dc[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 2), dc[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 3), dc[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 4), dc[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 5), dc[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 6), dc[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 7), dc[6], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 8), dc[7], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
            set_cell(ws.cell(cur_r, 9), dc[8], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
            cur_r += 1

    cur_r += 1
    # Section 5: Safety Interlocks, Alarms & Trips
    ws.merge_cells(f"A{cur_r}:I{cur_r}")
    set_cell(ws.cell(cur_r, 1), "5. SAFETY INTERLOCK, EMERGENCY TRIP & SCADA ALARM RESPONSE", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_safe = ["Safety Test Item", "Trigger Condition", "Expected Safe Action", "SCADA Alarm Text", "Priority / Class", "Sounder Horn", "Historian Log", "Result", "Remarks"]
    ws.row_dimensions[cur_r].height = 22
    for c_i, h in enumerate(hdrs_safe, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    safe_data = [
        ("Emergency Stop / Trip Interlock", "E-Stop Pushbutton or SIS Trip", "Fail-safe spring return to safe state", f"{equip_tag} SAFETY SHUTDOWN", "High (Priority 1)", "Audible Active", "Logged with ms", "[  ] PASS", "Hardwired trip verified"),
        ("Out-of-Range Alarm", "Signal <3.6mA or >21.0mA", "Channel fault banner / Hold permissive", f"{equip_tag} WIRE BREAK / FAULT", "Medium (Priority 2)", "Warning Tone", "Logged with ms", "[  ] PASS", "NAMUR NE43 compliance"),
        ("SCADA Alarm Acknowledge", "Operator Acknowledge click", "Banner silences horn, latches steady red", f"{equip_tag} ALARM ACKNOWLEDGED", "Low (Priority 3)", "Silenced", "Logged with ms", "[  ] PASS", "HMI banner functioning")
    ]
    for sd in safe_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 20
        set_cell(ws.cell(cur_r, 1), sd[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), sd[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), sd[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), sd[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), sd[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), sd[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), sd[6], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 8), sd[7], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 9), sd[8], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 6: Tripartite Sign-off Box
    ws.merge_cells(f"A{cur_r}:I{cur_r}")
    set_cell(ws.cell(cur_r, 1), "6. TRIPARTITE COMMISSIONING ACCEPTANCE SIGN-OFF", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    sign_hdrs = ["Sign-off Role", "Name & Designation", "Company / Organization", "Signature", "Date Signed", "Overall Verdict", "Punchlist Ref", "Tag Color Affixed", "Remarks"]
    ws.row_dimensions[cur_r].height = 22
    for c_i, h in enumerate(sign_hdrs, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    signatures = [
        ("Tested by (Contractor):", "Lead Instrument Engineer", "Kalasin Engineering Co., Ltd.", "___________________", "DD-MM-2026", "LOOP CHECK PASSED", "None (Clean)", "Green Loop Tag", "Ready for Process Commissioning"),
        ("Witnessed by (Client QC):", "Senior Electrical & Inst QA/QC", "Ingredion (Thailand) Co., Ltd.", "___________________", "DD-MM-2026", "ACCEPTED & APPROVED", "None", "QC Sticker Applied", "Satisfies Contract Scope"),
        ("Approved by (Commissioning PM):", "Project Commissioning Manager", "Kalasin / Ingredion Joint Team", "___________________", "DD-MM-2026", "FINAL AUTHORIZED", "None", "System Cleared", "Authorized for Cold Water Wet Run")
    ]
    cur_r += 1
    for s in signatures:
        ws.row_dimensions[cur_r].height = 26
        set_cell(ws.cell(cur_r, 1), s[0], font=FONT_DATA_BOLD, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), s[1], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), s[2], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), s[3], font=FONT_DATA_CODE, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), s[4], font=FONT_DATA_CODE, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), s[5], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), s[6], font=FONT_DATA, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 8), s[7], font=FONT_DATA_BOLD, fill=ZEBRA_EVEN, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 9), s[8], font=FONT_DATA_MUTED, fill=ZEBRA_EVEN, alignment=ALIGN_LEFT)
        cur_r += 1

    col_widths = {1: 18, 2: 34, 3: 16, 4: 24, 5: 22, 6: 20, 7: 22, 8: 22, 9: 28}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Master Execution: Generate All 15 Workbooks
# -----------------------------------------------------------------------------
def main():
    print("=================================================================")
    print("STARTING LOOP TEST REPORTS GENERATION (AI: 2-PIN / CH CONSOLIDATION)")
    print("=================================================================")

    jb_data, inst_meta = load_data()
    master_register = []

    for jb_name, equips in sorted(jb_data.items()):
        safe_jb_name = re.sub(r'[^A-Za-z0-9_-]', '_', jb_name)
        out_filename = f"Loop_Test_Report_{safe_jb_name}.xlsx"
        out_path = os.path.join(OUT_DIR, out_filename)

        print(f"\nBuilding Workbook for: {jb_name} ({len(equips)} Equipments) -> {out_filename}")
        wb = openpyxl.Workbook()
        wb.remove(wb.active)

        # 1. Sheet 0: JB Index
        ws_idx = wb.create_sheet(title="00_JB_Index")
        build_jb_index_sheet(ws_idx, jb_name, equips, inst_meta)
        ws_idx.freeze_panes = 'C8'
        ws_idx.auto_filter.ref = f"A7:K{ws_idx.max_row}"

        # 2. Equipment Sheets: 1 Equipment = 1 Sheet
        used_sheet_names = set(["00_JB_Index"])
        for equip_tag, sig_list in sorted(equips.items()):
            safe_sheet_name = re.sub(r'[\\/*?:\\[\\]]', '_', equip_tag)[:31]
            if safe_sheet_name in used_sheet_names:
                suffix = 1
                while f"{safe_sheet_name[:28]}_{suffix}" in used_sheet_names:
                    suffix += 1
                safe_sheet_name = f"{safe_sheet_name[:28]}_{suffix}"
            used_sheet_names.add(safe_sheet_name)

            ws_eq = wb.create_sheet(title=safe_sheet_name)
            build_equipment_sheet(ws_eq, jb_name, equip_tag, sig_list, inst_meta)
            ws_eq.freeze_panes = 'A5'

        wb.save(out_path)
        file_size_kb = os.path.getsize(out_path) / 1024
        print(f"  -> Saved: {out_filename} ({len(wb.sheetnames)} sheets, {file_size_kb:.1f} KB)")

        master_register.append({
            'jb': jb_name,
            'file': out_filename,
            'path': out_path,
            'equips': len(equips),
            'sigs': sum(len(v) for v in equips.values()),
            'sheets': len(wb.sheetnames),
            'size_kb': file_size_kb
        })

    # Master Register Workbook: 00_Master_Loop_Test_JB_Register.xlsx
    master_path = os.path.join(OUT_DIR, "00_Master_Loop_Test_JB_Register.xlsx")
    print(f"\nGenerating Master Junction Box Register: {os.path.basename(master_path)}...")
    wb_m = openpyxl.Workbook()
    ws_m = wb_m.active
    ws_m.title = "Master_JB_Loop_Test_Register"
    ws_m.views.sheetView[0].showGridLines = True

    # Title
    ws_m.merge_cells("A1:H1")
    set_cell(ws_m["A1"], "KALASIN ENGINEERING  |  MASTER JUNCTION BOX LOOP TEST WORKBOOK REGISTER",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws_m.row_dimensions[1].height = 25

    ws_m.merge_cells("A2:H2")
    set_cell(ws_m["A2"], "Project Sprint 18K TPA Spray Dryer / Jet Cooker  |  Directory: 05-LoopTest (AI: 2-Pin/CH)",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws_m.row_dimensions[2].height = 18

    hdrs_m = ["Item #", "Junction Box ID", "Report Workbook Filename", "Total Equipment", "Total Active Process Channels", "Total Sheets", "File Size (KB)", "Direct Workbook Link"]
    ws_m.row_dimensions[4].height = 24
    for c_i, h in enumerate(hdrs_m, 1):
        align = ALIGN_CENTER if c_i in [1, 2, 4, 5, 6, 7] else ALIGN_LEFT
        set_cell(ws_m.cell(4, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    r_idx = 5
    tot_eq_all = 0
    tot_sig_all = 0
    tot_sheets_all = 0
    for i, reg in enumerate(master_register, 1):
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws_m.row_dimensions[r_idx].height = 20
        tot_eq_all += reg['equips']
        tot_sig_all += reg['sigs']
        tot_sheets_all += reg['sheets']

        set_cell(ws_m.cell(r_idx, 1), i, font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws_m.cell(r_idx, 2), reg['jb'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws_m.cell(r_idx, 3), reg['file'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws_m.cell(r_idx, 4), reg['equips'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws_m.cell(r_idx, 5), reg['sigs'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws_m.cell(r_idx, 6), reg['sheets'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws_m.cell(r_idx, 7), f"{reg['size_kb']:.1f} KB", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_RIGHT)

        c_link = ws_m.cell(r_idx, 8)
        set_cell(c_link, f"Open {reg['file']}", font=Font(name="Calibri", size=9.5, bold=True, color="0284C7", underline="single"), fill=fill, alignment=ALIGN_CENTER)
        c_link.hyperlink = reg['file']

        r_idx += 1

    # Total Row
    ws_m.row_dimensions[r_idx].height = 22
    set_cell(ws_m.cell(r_idx, 1), "TOTAL", font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)
    set_cell(ws_m.cell(r_idx, 2), f"{len(master_register)} JBs", font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)
    set_cell(ws_m.cell(r_idx, 3), "Grand Totals Across All Junction Boxes (Consolidated Process Channels)", font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    set_cell(ws_m.cell(r_idx, 4), tot_eq_all, font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)
    set_cell(ws_m.cell(r_idx, 5), tot_sig_all, font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)
    set_cell(ws_m.cell(r_idx, 6), tot_sheets_all, font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)
    set_cell(ws_m.cell(r_idx, 7), "-", font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)
    set_cell(ws_m.cell(r_idx, 8), "Ready for Testing", font=FONT_HEADER, fill=SECTION_FILL, alignment=ALIGN_CENTER)

    col_widths_m = {1: 8, 2: 18, 3: 36, 4: 16, 5: 22, 6: 14, 7: 16, 8: 26}
    for c_i, w in col_widths_m.items():
        ws_m.column_dimensions[get_column_letter(c_i)].width = w

    ws_m.freeze_panes = 'C5'
    wb_m.save(master_path)
    print(f"Master Register saved to: {master_path}")

    print("\n=================================================================")
    print("ALL JUNCTION BOX LOOP TEST REPORTS UPDATED SUCCESSFULLY!")
    print(f"Total Workbooks: {len(master_register) + 1}")
    print(f"Total Equipment: {tot_eq_all}")
    print(f"Total Process Channels: {tot_sig_all}")
    print(f"Folder: {OUT_DIR}")
    print("=================================================================")

if __name__ == "__main__":
    main()
