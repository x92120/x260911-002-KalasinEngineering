#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
COMPILATION OF COMPLETE PLC WIRE MARKS ACCORDING TO CHASSIS & SLOT CONFIG
Outputs:
  - 02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC-r02.xlsx
  - 03_IO_Lists_and_Schedules/Tags list PLC-r02.xlsx
================================================================================
Standard ePlan Wire Marking Pair:
  - At PLC Swing Arm:        TBxxx:Pin/CxxSxx:Pin  (e.g., TBDI1:1A/C1S4:1)
  - At Terminal Block/Relay: CxxSxx:Pin/TBxxx:Pin  (e.g., C1S4:1/TBDI1:1A)
================================================================================
"""

import os
import sys
import re
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_LIST_PATH = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-6.xlsx")
ORIG_TAGS_PATH = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC.xlsx")

OUTPUT_PATH_1 = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC-r02.xlsx")
OUTPUT_PATH_2 = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Tags list PLC-r02.xlsx")

# Styling Definitions
FONT_FAMILY = "Segoe UI"
NAVY_HEADER = "1E293B"      # Deep Slate Navy
NAVY_SUBHEADER = "334155"   # Slate Subheader
WHITE = "FFFFFF"
LIGHT_BG_1 = "FFFFFF"
LIGHT_BG_2 = "F8FAFC"
BORDER_COLOR = "CBD5E1"

SRC_HDR_FILL = "DBEAFE"     # Light Ice Blue
SRC_HDR_FONT = "1E3A8A"
DEST_HDR_FILL = "FEF3C7"    # Light Amber
DEST_HDR_FONT = "78350F"

ACTIVE_FILL = "DCFCE7"      # Light Emerald
ACTIVE_FONT = "14532D"
SPARE_FILL = "FEF9C3"       # Light Yellow
SPARE_FONT = "854D0E"

def thin_border():
    s = Side(style='thin', color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style='medium', color="94A3B8")
    return Border(left=s, right=s, top=s, bottom=s)

def load_original_tags():
    """Load original tags from Tags list PLC.xlsx to guarantee 100% fidelity"""
    orig_map = {}
    if os.path.exists(ORIG_TAGS_PATH):
        wb = openpyxl.load_workbook(ORIG_TAGS_PATH, data_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            for r in range(2, ws.max_row + 1):
                p_tag = ws.cell(r, 2).value
                t_tag = ws.cell(r, 3).value
                rem = ws.cell(r, 4).value
                if p_tag or t_tag:
                    # Match C{x}S{y}:{pin}
                    combined = f"{p_tag or ''} {t_tag or ''}"
                    m = re.search(r'C(\d+)S(\d+):(\d+)', combined)
                    if m:
                        k = (f"C{m.group(1)}", int(m.group(2)), int(m.group(3)))
                        orig_map[k] = {
                            'tag_plc': str(p_tag).strip() if p_tag else '',
                            'tag_term': str(t_tag).strip() if t_tag else '',
                            'remarks': str(rem).strip() if rem else ''
                        }
    return orig_map

def main():
    print(f"Loading master I/O database: {IO_LIST_PATH}")
    wb_io = openpyxl.load_workbook(IO_LIST_PATH, data_only=True)
    ws_io = wb_io['IO List']
    
    headers = [ws_io.cell(1, c).value for c in range(1, ws_io.max_column + 1)]
    
    ch_idx = headers.index('Chassis') + 1
    sl_idx = headers.index('Slot') + 1
    pt_idx = headers.index('Point') + 1
    io_idx = headers.index('I/O Type') + 1
    card_idx = headers.index('Card') + 1
    plc_idx = headers.index('PLC_Tag') + 1
    inst_idx = headers.index('Instruement Tag') + 1
    desc_idx = headers.index('Instruement_Description') + 1
    dest_idx = headers.index('Destination') + 1
    term_idx = headers.index('Terminal') + 1
    pid_idx = headers.index('P&ID No. 3.3') + 1
    pid2_idx = headers.index('P&ID No. 3.5') + 1
    sig_idx = headers.index('Signal Type') + 1

    orig_tags = load_original_tags()
    print(f"Loaded {len(orig_tags)} legacy wire tags from original file.")

    # Load all rows
    raw_records = []
    for r in range(2, ws_io.max_row + 1):
        ch = ws_io.cell(r, ch_idx).value
        sl = ws_io.cell(r, sl_idx).value
        pt = ws_io.cell(r, pt_idx).value
        if ch is None or sl is None:
            continue
            
        c_str = str(ch).strip()
        s_int = int(sl) if str(sl).isdigit() else str(sl)
        pt_int = int(pt) if pt is not None and str(pt).isdigit() else 0
        
        io_t = str(ws_io.cell(r, io_idx).value or '').strip().upper()
        card = str(ws_io.cell(r, card_idx).value or '').strip()
        plc_tag = str(ws_io.cell(r, plc_idx).value or '').strip()
        inst_tag = str(ws_io.cell(r, inst_idx).value or '').strip()
        desc = str(ws_io.cell(r, desc_idx).value or '').strip()
        dest = str(ws_io.cell(r, dest_idx).value or '').strip()
        term = str(ws_io.cell(r, term_idx).value or '').strip()
        pid = str(ws_io.cell(r, pid_idx).value or ws_io.cell(r, pid2_idx).value or '').strip()
        sig = str(ws_io.cell(r, sig_idx).value or '').strip()

        is_spare = False
        if not inst_tag or inst_tag.lower() in ['spare', 'none', '-'] or 'spare' in plc_tag.lower():
            is_spare = True

        raw_records.append({
            'source_row': r,
            'chassis': c_str,
            'slot': s_int,
            'point': pt_int,
            'channel': pt_int + 1,
            'io_type': io_t,
            'card': card,
            'plc_tag': plc_tag,
            'inst_tag': inst_tag if not is_spare else 'Spare',
            'desc': desc if not is_spare else 'Spare Channel',
            'dest': dest,
            'terminal_raw': term,
            'pid': pid,
            'signal_type': sig,
            'is_spare': is_spare,
            'status': 'SPARE' if is_spare else 'ACTIVE'
        })

    print(f"Loaded {len(raw_records)} active channel records.")

    # Sort records logically by Chassis, Slot, Point
    def sort_key(rec):
        c_num = int(rec['chassis'][1:]) if rec['chassis'].startswith('C') and rec['chassis'][1:].isdigit() else 999
        s_num = int(rec['slot']) if isinstance(rec['slot'], int) or str(rec['slot']).isdigit() else 999
        p_num = int(rec['point']) if isinstance(rec['point'], int) or str(rec['point']).isdigit() else 999
        return (c_num, s_num, p_num)

    raw_records.sort(key=sort_key)

    # Calculate per-chassis slot card sequences and assign deterministic wire marks
    # Group by chassis and card type to determine TB sequence (TBDI1, TBDI2, etc. and RL.1..RL.32, etc.)
    # We maintain slot order for each chassis
    chassis_groups = defaultdict(lambda: defaultdict(list))
    for rec in raw_records:
        chassis_groups[rec['chassis']][rec['slot']].append(rec)

    compiled_records = []
    
    # Global Running Sequential Numbering across Control System (No Overlaps):
    di_card_idx = 0
    relay_global_idx = 0
    ai_card_idx = 0
    ao_card_idx = 0

    # Dedicated sequence for MCC Room (Chassis C6 & C7)
    mcc_di_idx = 0
    mcc_relay_idx = 0

    # Process each chassis with Chassis-based Terminal Numbering:
    # C1: TBDI1..4, TBDO1..2, TBAI1..2
    # C2: TBDI21..25, TBDO21..23, TBAI21..25
    # C3: TBDI31..32, TBDO31, TBAI31..32, TBAO31
    # C4: TBDI41..43, TBDO41, TBAI41..42
    # C5: TBDI51..53, TBDO51..52, TBAI51
    # C6/C7: TBMCC-DI / TBMCC-DO / TBMCC-BUS
    for c_str in sorted(chassis_groups.keys(), key=lambda x: int(x[1:]) if x[1:].isdigit() else 999):
        slots_in_ch = sorted(chassis_groups[c_str].keys(), key=lambda x: int(x) if str(x).isdigit() else 999)

        c_num = int(c_str[1:]) if c_str.startswith('C') and c_str[1:].isdigit() else 0
        base_idx = c_num * 10 if c_num > 1 else 0

        ch_di_idx = 0
        ch_do_idx = 0
        ch_ai_idx = 0
        ch_ao_idx = 0

        for raw_sl in slots_in_ch:
            slot_records = chassis_groups[c_str][raw_sl]
            # determine card I/O type
            io_types = [r['io_type'] for r in slot_records]
            main_io = max(set(io_types), key=io_types.count) if io_types else 'DI'
            
            if c_str in ['C6', 'C7']:
                if main_io == 'DI':
                    mcc_di_idx += 1
                    tb_strip_name = f"TBMCC-DI{mcc_di_idx}"
                elif main_io == 'DO':
                    mcc_relay_idx += 1
                    tb_strip_name = f"TBMCC-DO{mcc_relay_idx}"
                elif main_io == 'BUS':
                    tb_strip_name = "TBMCC-BUS"
                else:
                    tb_strip_name = f"TBMCC-{main_io}"
            else:
                if main_io == 'DI':
                    ch_di_idx += 1
                    tb_num = base_idx + ch_di_idx if c_num > 1 else ch_di_idx
                    tb_strip_name = f"TBDI{tb_num}"
                elif main_io == 'DO':
                    ch_do_idx += 1
                    tb_num = base_idx + ch_do_idx if c_num > 1 else ch_do_idx
                    tb_strip_name = f"TBDO{tb_num}"
                elif main_io == 'AI':
                    ch_ai_idx += 1
                    tb_num = base_idx + ch_ai_idx if c_num > 1 else ch_ai_idx
                    tb_strip_name = f"TBAI{tb_num}"
                elif main_io == 'AO':
                    ch_ao_idx += 1
                    tb_num = base_idx + ch_ao_idx if c_num > 1 else ch_ao_idx
                    tb_strip_name = f"TBAO{tb_num}"
                else:
                    tb_strip_name = f"TB{main_io}"

            # Physical Rack Slot mapping (0-indexed rack slots):
            # C1: S0=CPU, S1..S3=ETH4 -> raw slots 3..10 map to physical rack slots 4..11 (C1S4..C1S11)
            # C2: S0=ETH -> raw slots 1..13 map to physical rack slots 1..13 (C2S1..C2S13)
            # C3..C5: S0=ETH -> raw slots 1..N map to physical rack slots 1..N
            # C6, C7: MCC slots
            sl_int = int(raw_sl) if str(raw_sl).isdigit() else 0
            if c_str == 'C1':
                phys_slot = sl_int + 1
            else:
                phys_slot = sl_int

            sx_str = f"S{phys_slot}"
            cx_sx = f"{c_str}{sx_str}"

            for rec in slot_records:
                pt = rec['point']
                ch = rec['channel']
                
                # Check swing-arm terminal pin formula
                # For 32-point cards (1756-IB32 / 1756-OB32):
                # Points 0..15 -> Pins 1..16
                # Points 16..31 -> Pins 19..34 (Pins 17, 18, 35, 36 are 24VDC/0VDC commons)
                if pt < 16:
                    plc_pin = pt + 1
                else:
                    plc_pin = pt + 3

                # Determine destination terminal mark based on card type
                if main_io in ['DI', 'DO']:
                    # 32-point cards use 1A..16A for lower byte, 1B..16B for upper byte
                    if pt < 16:
                        tb_pin = f"{tb_strip_name}:{ch}A"
                    else:
                        tb_pin = f"{tb_strip_name}:{ch - 16}B"
                elif main_io in ['AI', 'AO']:
                    tb_pin = f"{tb_strip_name}:{ch}(+)"
                elif main_io == 'BUS':
                    tb_pin = f"{tb_strip_name}:{ch:02d}"
                else:
                    tb_pin = f"{tb_strip_name}:{ch}"

                # Slot starts from 0 with format CxSx (no leading zero)
                plc_term = f"{cx_sx}:{plc_pin}"
                tag_term = f"{plc_term}/{tb_pin}"
                tag_plc = f"{tb_pin}/{plc_term}"
                rem = "Compiled Point" if not rec['is_spare'] else "Spare Channel"

                rec['slot_0'] = phys_slot
                rec['raw_slot'] = raw_sl
                rec['sx_str'] = sx_str
                rec['cx_sx'] = cx_sx
                rec['sxx_str'] = sx_str
                rec['cxx_sxx'] = cx_sx
                rec['plc_term'] = plc_term
                rec['tb_pin'] = tb_pin
                rec['tag_term'] = tag_term
                rec['tag_plc'] = tag_plc
                rec['wire_mark'] = tag_term  # CxSx/TBxxx
                rec['remarks'] = rem
                compiled_records.append(rec)

    print(f"Successfully compiled {len(compiled_records)} total wire mark tags.")

    # Create Workbook
    wb_out = openpyxl.Workbook()
    # remove default sheet
    wb_out.remove(wb_out.active)

    # Sort records strictly by Chassis (Cxx), Slot (Sxx starting from 0), Point
    def cxx_sxx_sort_key(r):
        ch_str = str(r['chassis']).strip().upper()
        c_num = int(ch_str[1:]) if ch_str.startswith('C') and ch_str[1:].isdigit() else 999
        s_num = r['slot_0']
        p_val = r['point']
        p_num = int(p_val) if isinstance(p_val, int) or str(p_val).isdigit() else 999
        return (c_num, s_num, p_num)

    sorted_by_cxx_sxx = sorted(compiled_records, key=cxx_sxx_sort_key)

    # =========================================================================
    # SHEET 1: All_Marks_Sorted_CxSx (Dedicated Cx/Sx Sorted Wire Marks)
    # =========================================================================
    ws_cxx = wb_out.create_sheet("All_Marks_Sorted_CxSx")
    ws_cxx.views.sheetView[0].showGridLines = True

    # Title Block
    ws_cxx.merge_cells("A1:R1")
    tcell = ws_cxx["A1"]
    tcell.value = "ALL PLC WIRE MARKS - SORTED BY CHASSIS & SLOT (FORMAT CxSx, NO LEADING ZERO)"
    tcell.font = Font(name=FONT_FAMILY, size=13.5, bold=True, color=WHITE)
    tcell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    ws_cxx.row_dimensions[1].height = 36

    ws_cxx.merge_cells("A2:R2")
    subcell = ws_cxx["A2"]
    subcell.value = "Master Rockwell ControlLogix 1756 PLC System | 1,210 Points Ordered by Chassis (C1..C7) & Slot (S0..S12) with CxSx/TBxxx Wire Marks"
    subcell.font = Font(name=FONT_FAMILY, size=9.5, italic=True, color="E2E8F0")
    subcell.fill = PatternFill("solid", fgColor=NAVY_SUBHEADER)
    subcell.alignment = Alignment(horizontal="center", vertical="center")
    ws_cxx.row_dimensions[2].height = 20

    cxx_headers = [
        ("No.", 6),
        ("Chassis\n(Cx)", 9),
        ("Slot\n(Sx)", 8),
        ("Chassis/Slot\n(CxSx)", 14),
        ("Ch.", 5),
        ("Pt.", 5),
        ("Wire Mark Tag\n(CxSx/TBxxx)", 24),
        ("PLC Tag Mark\n(TBxxx/CxSx)", 24),
        ("Terminal / Relay\n(TBxxx)", 18),
        ("PLC Tag Name", 24),
        ("Instrument Tag", 18),
        ("Description / Service", 34),
        ("Destination\n(JB / MCC)", 14),
        ("Signal Type", 12),
        ("Status", 10),
        ("Raw Slot Ref\n(IO List)", 13),
        ("P&ID Ref", 14),
        ("Remarks", 18)
    ]

    ws_cxx.row_dimensions[4].height = 28
    for col_idx, (h_name, width) in enumerate(cxx_headers, 1):
        cell = ws_cxx.cell(4, col_idx, h_name)
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border()
        col_letter = get_column_letter(col_idx)
        ws_cxx.column_dimensions[col_letter].width = width

    for r_i, rec in enumerate(sorted_by_cxx_sxx, 5):
        ws_cxx.row_dimensions[r_i].height = 18
        bg_color = LIGHT_BG_1 if r_i % 2 == 1 else LIGHT_BG_2

        cxx_row_data = [
            (r_i - 4, "center", False, None, None),
            (rec['chassis'], "center", True, "0F172A", None),
            (rec['sx_str'], "center", True, "0F172A", None),
            (rec['cx_sx'], "center", True, "1E3A8A", SRC_HDR_FILL),
            (rec['channel'], "center", False, None, None),
            (rec['point'], "center", False, None, None),
            (rec['tag_term'], "left", True, "0F172A", None),               # CxSx/TBxxx
            (rec['tag_plc'], "left", False, "1E3A8A", None),               # TBxxx/CxSx
            (rec['tb_pin'], "center", False, None, None),
            (rec['plc_tag'], "left", False, None, None),
            (rec['inst_tag'], "left", True, None, None),
            (rec['desc'], "left", False, None, None),
            (rec['dest'], "center", False, None, None),
            (rec['io_type'], "center", False, None, None),
            (rec['status'], "center", True, ACTIVE_FONT if rec['status'] == 'ACTIVE' else SPARE_FONT, ACTIVE_FILL if rec['status'] == 'ACTIVE' else SPARE_FILL),
            (f"Slot {rec['raw_slot']}", "center", False, None, None),
            (rec['pid'], "center", False, None, None),
            (rec['remarks'], "left", False, None, None)
        ]

        for col_idx, (val, align, is_bold, text_color, custom_fill) in enumerate(cxx_row_data, 1):
            cell = ws_cxx.cell(r_i, col_idx, val)
            cell.font = Font(name=FONT_FAMILY, size=9, bold=is_bold, color=text_color if text_color else "0F172A")
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.border = thin_border()
            if custom_fill:
                cell.fill = PatternFill("solid", fgColor=custom_fill)
            else:
                cell.fill = PatternFill("solid", fgColor=bg_color)

    ws_cxx.freeze_panes = "A5"
    ws_cxx.auto_filter.ref = f"A4:R{len(sorted_by_cxx_sxx) + 4}"

    # =========================================================================
    # SHEET 2: 00_Master_Wire_Marks
    # =========================================================================
    ws_master = wb_out.create_sheet("00_Master_Wire_Marks")
    ws_master.views.sheetView[0].showGridLines = True

    # Title Block
    ws_master.merge_cells("A1:N1")
    tcell = ws_master["A1"]
    tcell.value = "KALASIN STARCH PLANT - MASTER PLC WIRE MARKS SCHEDULE (COMPILED CxSx/TBxxx)"
    tcell.font = Font(name=FONT_FAMILY, size=14, bold=True, color=WHITE)
    tcell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    ws_master.row_dimensions[1].height = 36

    ws_master.merge_cells("A2:N2")
    subcell = ws_master["A2"]
    subcell.value = "Rockwell ControlLogix 1756 PLC System | Complete ePlan Wire Marking Database (1,210 Tags Compiled Across Chassis C1 to C7)"
    subcell.font = Font(name=FONT_FAMILY, size=9.5, italic=True, color="E2E8F0")
    subcell.fill = PatternFill("solid", fgColor=NAVY_SUBHEADER)
    subcell.alignment = Alignment(horizontal="center", vertical="center")
    ws_master.row_dimensions[2].height = 20

    # Header Row
    master_headers = [
        ("No.", 6),
        ("Tag Wire Mark\n(CxSx/TBxxx)", 24),
        ("Tag PLC\n(TBxxx/CxSx)", 24),
        ("Chassis", 8),
        ("Slot (Sx)", 8),
        ("Ch.", 5),
        ("Pt.", 5),
        ("Terminal / Pin", 18),
        ("PLC Tag Name", 24),
        ("Instrument Tag", 18),
        ("Description", 34),
        ("Destination", 14),
        ("Signal Type", 12),
        ("Status", 10),
        ("Remarks", 18)
    ]

    ws_master.row_dimensions[4].height = 28
    for col_idx, (h_name, width) in enumerate(master_headers, 1):
        cell = ws_master.cell(4, col_idx, h_name)
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border()
        col_letter = get_column_letter(col_idx)
        ws_master.column_dimensions[col_letter].width = width

    # Data Rows
    for r_i, rec in enumerate(compiled_records, 5):
        ws_master.row_dimensions[r_i].height = 18
        bg_color = LIGHT_BG_1 if r_i % 2 == 1 else LIGHT_BG_2
        
        row_data = [
            (r_i - 4, "center", False, None, None),
            (rec['tag_term'], "left", True, "0F172A", SRC_HDR_FILL),      # CxSx/TBxxx
            (rec['tag_plc'], "left", False, "1E3A8A", None),              # TBxxx/CxSx
            (rec['chassis'], "center", True, None, None),
            (rec['sx_str'], "center", True, None, None),
            (rec['channel'], "center", False, None, None),
            (rec['point'], "center", False, None, None),
            (rec['tb_pin'], "center", False, None, None),
            (rec['plc_tag'], "left", False, None, None),
            (rec['inst_tag'], "left", True, None, None),
            (rec['desc'], "left", False, None, None),
            (rec['dest'], "center", False, None, None),
            (rec['io_type'], "center", False, None, None),
            (rec['status'], "center", True, ACTIVE_FONT if rec['status'] == 'ACTIVE' else SPARE_FONT, ACTIVE_FILL if rec['status'] == 'ACTIVE' else SPARE_FILL),
            (rec['remarks'], "left", False, None, None)
        ]

        for col_idx, (val, align, is_bold, text_color, custom_fill) in enumerate(row_data, 1):
            cell = ws_master.cell(r_i, col_idx, val)
            cell.font = Font(name=FONT_FAMILY, size=9, bold=is_bold, color=text_color if text_color else "0F172A")
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.border = thin_border()
            if custom_fill:
                cell.fill = PatternFill("solid", fgColor=custom_fill)
            else:
                cell.fill = PatternFill("solid", fgColor=bg_color)

    # Freeze Panes on Master Sheet
    ws_master.freeze_panes = "A5"

    # =========================================================================
    # SHEET 3: 00_Summary_By_Slot
    # =========================================================================
    ws_sum = wb_out.create_sheet("00_Summary_By_Slot")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("A1:M1")
    tcell = ws_sum["A1"]
    tcell.value = "CHASSIS & SLOT CONFIGURATION - WIRE MARKS COMPILATION SUMMARY (FORMAT CxSx)"
    tcell.font = Font(name=FONT_FAMILY, size=13, bold=True, color=WHITE)
    tcell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    tcell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 32

    ws_sum.merge_cells("A2:M2")
    subcell = ws_sum["A2"]
    subcell.value = "Hardware Card Cross-Reference, Terminal Block ID Sequences, Active/Spare Counts, and Wire Mark Ranges"
    subcell.font = Font(name=FONT_FAMILY, size=9, italic=True, color="E2E8F0")
    subcell.fill = PatternFill("solid", fgColor=NAVY_SUBHEADER)
    subcell.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 18

    sum_headers = [
        ("Chassis", 9),
        ("Slot (Sx)", 10),
        ("Chassis/Slot (CxSx)", 14),
        ("Signal Type", 11),
        ("Hardware Card", 15),
        ("Channels", 9),
        ("Terminal Block / Relay Range", 24),
        ("Wire Mark Range (CxSx/TBxxx)", 36),
        ("PLC Tag Wire Mark Range (TBxxx/CxSx)", 36),
        ("Raw Slot Ref", 14),
        ("Active", 8),
        ("Spare", 8),
        ("Total Tags", 10)
    ]

    ws_sum.row_dimensions[4].height = 26
    for col_idx, (h_name, width) in enumerate(sum_headers, 1):
        cell = ws_sum.cell(4, col_idx, h_name)
        cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border()
        col_letter = get_column_letter(col_idx)
        ws_sum.column_dimensions[col_letter].width = width

    # Aggregate by (chassis, slot_0)
    slot_summary = defaultdict(list)
    for rec in compiled_records:
        slot_summary[(rec['chassis'], rec['slot_0'])].append(rec)

    def slot_sort_key(k):
        c_num = int(k[0][1:]) if k[0].startswith('C') and k[0][1:].isdigit() else 999
        s_num = k[1]
        return (c_num, s_num)

    s_row = 5
    for (c_str, sl_0) in sorted(slot_summary.keys(), key=slot_sort_key):
        recs = slot_summary[(c_str, sl_0)]
        io_t = recs[0]['io_type']
        card = recs[0]['card']
        total_pts = len(recs)
        active_pts = sum(1 for r in recs if r['status'] == 'ACTIVE')
        spare_pts = total_pts - active_pts
        sxx_lbl = recs[0]['sxx_str']
        cxx_sxx_lbl = recs[0]['cxx_sxx']
        raw_sl_lbl = f"Slot {recs[0]['raw_slot']}"
        
        first_wm = recs[0]['tag_term']
        last_wm = recs[-1]['tag_term']
        wm_range = f"{first_wm} .. {last_wm}"
        
        first_plc = recs[0]['tag_plc']
        last_plc = recs[-1]['tag_plc']
        plc_range = f"{first_plc} .. {last_plc}"
        
        first_tb = recs[0]['tb_pin']
        last_tb = recs[-1]['tb_pin']
        tb_range = f"{first_tb} .. {last_tb}"

        ws_sum.row_dimensions[s_row].height = 20
        bg_color = LIGHT_BG_1 if s_row % 2 == 1 else LIGHT_BG_2

        s_vals = [
            (c_str, "center", True),
            (sxx_lbl, "center", True),
            (cxx_sxx_lbl, "center", True),
            (io_t, "center", True),
            (card, "center", False),
            (total_pts, "center", False),
            (tb_range, "center", False),
            (wm_range, "left", True),
            (plc_range, "left", False),
            (raw_sl_lbl, "center", False),
            (active_pts, "center", False),
            (spare_pts, "center", False),
            (total_pts, "center", True)
        ]

        for col_idx, (val, align, is_b) in enumerate(s_vals, 1):
            cell = ws_sum.cell(s_row, col_idx, val)
            cell.font = Font(name=FONT_FAMILY, size=9, bold=is_b, color="0F172A")
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.border = thin_border()
            cell.fill = PatternFill("solid", fgColor=bg_color)
        
        s_row += 1

    # Freeze panes
    ws_sum.freeze_panes = "A5"

    # =========================================================================
    # INDIVIDUAL CHASSIS & SIGNAL SHEETS (Full ePlan Standard Compatibility)
    # =========================================================================
    # Group compiled records into dedicated tabs
    tab_definitions = [
        ("DI Chassis 1", lambda r: r['chassis'] == 'C1' and r['io_type'] == 'DI'),
        ("DO Chassis 1", lambda r: r['chassis'] == 'C1' and r['io_type'] == 'DO'),
        ("AI Chassis 1", lambda r: r['chassis'] == 'C1' and r['io_type'] == 'AI'),
        ("DI Chassis 2", lambda r: r['chassis'] == 'C2' and r['io_type'] == 'DI'),
        ("DO Chassis 2", lambda r: r['chassis'] == 'C2' and r['io_type'] == 'DO'),
        ("AI Chassis 2", lambda r: r['chassis'] == 'C2' and r['io_type'] == 'AI'),
        ("Chassis 3 (Skid Rack)", lambda r: r['chassis'] == 'C3'),
        ("Chassis 4 (Spray Dryer)", lambda r: r['chassis'] == 'C4'),
        ("Chassis 5 (RIO-200 Skid)", lambda r: r['chassis'] == 'C5'),
        ("Chassis 6 & 7 (MCC)", lambda r: r['chassis'] in ['C6', 'C7'])
    ]

    for title, filter_fn in tab_definitions:
        tab_recs = [r for r in compiled_records if filter_fn(r)]
        if not tab_recs:
            continue
            
        ws = wb_out.create_sheet(title)
        ws.views.sheetView[0].showGridLines = True

        is_di = 'DI' in title
        is_do = 'DO' in title
        is_ai = 'AI' in title
        sig_label = 'DI' if is_di else ('DO' if is_do else ('AI' if is_ai else 'IO'))

        col_b_name = f"Tag {sig_label} PLC\n(TBxxx/CxSx)"
        col_c_name = f"Tag {sig_label} Terminal Fuse\n(CxSx/TBxxx)"

        tab_headers = [
            ("No.", 7),
            (col_b_name, 23),
            (col_c_name, 23),
            ("หมายเหตุ / Remarks", 18),
            ("PLC Tag Name", 22),
            ("Instrument Tag", 16),
            ("Description", 30),
            ("Destination", 14)
        ]

        ws.row_dimensions[1].height = 26
        for col_idx, (h_name, width) in enumerate(tab_headers, 1):
            cell = ws.cell(1, col_idx, h_name)
            cell.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
            cell.fill = PatternFill("solid", fgColor=NAVY_HEADER)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = header_border()
            col_letter = get_column_letter(col_idx)
            ws.column_dimensions[col_letter].width = width

        # Group tab_recs by slot_0 (0-indexed)
        cur_row = 2
        card_grouped = defaultdict(list)
        for r in tab_recs:
            card_grouped[r['slot_0']].append(r)

        card_counter = 1
        for sl_0 in sorted(card_grouped.keys()):
            s_recs = card_grouped[sl_0]
            c_type = s_recs[0]['io_type']
            sx_lbl = s_recs[0]['sx_str']
            raw_sl = s_recs[0]['raw_slot']
            
            # If not first card, insert card section break
            if card_counter > 1:
                ws.row_dimensions[cur_row].height = 14
                cur_row += 1
                # Card label row
                ws.row_dimensions[cur_row].height = 20
                c_lbl = ws.cell(cur_row, 1, f"{c_type} Card {card_counter} (Slot {sx_lbl} / HW Slot {raw_sl})")
                c_lbl.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=SRC_HDR_FONT)
                c_lbl.fill = PatternFill("solid", fgColor=SRC_HDR_FILL)
                cur_row += 1

            # Insert pins
            no_counter = 1
            for rec_idx, r in enumerate(s_recs, 1):
                # For 32-point cards, insert power/common pins at Pin 17/18 and 35/36 if matching original
                if len(s_recs) == 32 and rec_idx == 17:
                    # Insert Pin 17 & 18 commons
                    pwr_v = "24VDC" if c_type == 'DO' else "0VDC"
                    for p_val in [pwr_v, "0VDC"]:
                        ws.row_dimensions[cur_row].height = 17
                        ws.cell(cur_row, 1, float(no_counter)).alignment = Alignment(horizontal="center", vertical="center")
                        ws.cell(cur_row, 2, p_val).alignment = Alignment(horizontal="center", vertical="center")
                        ws.cell(cur_row, 3, p_val).alignment = Alignment(horizontal="center", vertical="center")
                        ws.cell(cur_row, 4, "ยังไม่ใส่มาร์ก (Internal Power Common)").alignment = Alignment(horizontal="left", vertical="center")
                        for ci in range(1, 9):
                            ws.cell(cur_row, ci).border = thin_border()
                            ws.cell(cur_row, ci).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                        cur_row += 1
                        no_counter += 1

                ws.row_dimensions[cur_row].height = 18
                bg = LIGHT_BG_1 if cur_row % 2 == 1 else LIGHT_BG_2

                ws.cell(cur_row, 1, float(no_counter)).alignment = Alignment(horizontal="center", vertical="center")
                ws.cell(cur_row, 2, r['tag_plc']).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(cur_row, 3, r['tag_term']).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(cur_row, 4, r['remarks']).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(cur_row, 5, r['plc_tag']).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(cur_row, 6, r['inst_tag']).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(cur_row, 7, r['desc']).alignment = Alignment(horizontal="left", vertical="center")
                ws.cell(cur_row, 8, r['dest']).alignment = Alignment(horizontal="center", vertical="center")

                ws.cell(cur_row, 1).font = Font(name=FONT_FAMILY, size=9)
                ws.cell(cur_row, 2).font = Font(name=FONT_FAMILY, size=9, color="1E3A8A")
                ws.cell(cur_row, 3).font = Font(name=FONT_FAMILY, size=9, bold=True, color="0F172A")
                ws.cell(cur_row, 4).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                ws.cell(cur_row, 5).font = Font(name=FONT_FAMILY, size=9)
                ws.cell(cur_row, 6).font = Font(name=FONT_FAMILY, size=9, bold=True)
                ws.cell(cur_row, 7).font = Font(name=FONT_FAMILY, size=9)
                ws.cell(cur_row, 8).font = Font(name=FONT_FAMILY, size=9)

                for ci in range(1, 9):
                    ws.cell(cur_row, ci).border = thin_border()
                    ws.cell(cur_row, ci).fill = PatternFill("solid", fgColor=bg)

                cur_row += 1
                no_counter += 1

            # End of 32-pt card commons
            if len(s_recs) == 32:
                pwr_v = "24VDC" if c_type == 'DO' else "0VDC"
                for p_val in [pwr_v, "0VDC"]:
                    ws.row_dimensions[cur_row].height = 17
                    ws.cell(cur_row, 1, float(no_counter)).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(cur_row, 2, p_val).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(cur_row, 3, p_val).alignment = Alignment(horizontal="center", vertical="center")
                    ws.cell(cur_row, 4, "ยังไม่ใส่มาร์ก (Internal Power Common)").alignment = Alignment(horizontal="left", vertical="center")
                    for ci in range(1, 9):
                        ws.cell(cur_row, ci).border = thin_border()
                        ws.cell(cur_row, ci).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                    cur_row += 1
                    no_counter += 1

            card_counter += 1

        ws.freeze_panes = "A2"

    # Save to both target locations
    print(f"Saving to {OUTPUT_PATH_1}...")
    wb_out.save(OUTPUT_PATH_1)
    print(f"Saving to {OUTPUT_PATH_2}...")
    wb_out.save(OUTPUT_PATH_2)
    print("Generation completed successfully!")

if __name__ == '__main__':
    main()
