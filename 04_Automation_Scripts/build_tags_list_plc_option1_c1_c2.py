#!/usr/bin/env python3
"""
build_tags_list_plc_option1_c1_c2.py
====================================
Generates a dedicated workbook for Option 1: Direct Slot Correlation (TB-CxSx).
Applies strictly to Chassis C1 and C2 as requested by the user.

Terminal Strip Naming Convention:
  - TB-C1S4  (for DI Card 1 at C1S4)
  - TB-C1S5  (for DI Card 2 at C1S5)
  - TB-C1S6  (for DI Card 3 at C1S6)
  - TB-C1S7  (for DI Card 4 at C1S7)
  - TB-C1S8  (for DO Card 1 at C1S8)
  - TB-C1S9  (for DO Card 2 at C1S9)
  - TB-C1S10 (for AI Card 1 at C1S10)
  - TB-C1S11 (for AI Card 2 at C1S11)
  
  - TB-C2S1  (for DI Card 1 at C2S1)
  - TB-C2S2  (for DI Card 2 at C2S2)
  - TB-C2S3  (for DI Card 3 at C2S3)
  - TB-C2S4  (for DI Card 4 at C2S4)
  - TB-C2S5  (for DI Card 5 at C2S5)
  - TB-C2S6  (for DO Card 1 at C2S6)
  - TB-C2S7  (for DO Card 2 at C2S7)
  - TB-C2S8  (for DO Card 3 at C2S8)
  - TB-C2S9  (for AI Card 1 at C2S9)
  - TB-C2S10 (for AI Card 2 at C2S10)
  - TB-C2S11 (for AI Card 3 at C2S11)
  - TB-C2S12 (for AI Card 4 at C2S12)
  - TB-C2S13 (for AI Card 5 at C2S13)

Pins:
  - 32-point DI/DO: Pins 1A..16A (channels 1..16), 1B..16B (channels 17..32)
  - 16-point AI: Pins 1(+)..16(+)

Wire Marks:
  - Field End: CxSx:Pin/TB-CxSx:Pin (e.g. C2S1:1/TB-C2S1:1A)
  - PLC End:   TB-CxSx:Pin/CxSx:Pin (e.g. TB-C2S1:1A/C2S1:1)

Outputs:
  - 02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC-Option1-DirectSlot-C1-C2.xlsx
  - 03_IO_Lists_and_Schedules/Tags list PLC-Option1-DirectSlot-C1-C2.xlsx
"""

import os
import sys
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from collections import defaultdict

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
IO_LIST_PATH = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules", "IO_List_xDev-R01-Tag35-6.xlsx")
OUT_EPLAN_PATH = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing", "02_ePlan_Exports", "Tags list PLC-Option1-DirectSlot-C1-C2.xlsx")
OUT_SCHEDULE_PATH = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules", "Tags list PLC-Option1-DirectSlot-C1-C2.xlsx")

# Palette - Light Executive / Steel Navy Theme
FONT_FAMILY = "Segoe UI"
NAVY_HEADER = "1E3A8A"        # Deep Slate Blue / Navy
NAVY_SUBHEADER = "2563EB"     # Royal Blue accent
SRC_HDR_FILL = "DBEAFE"       # Light Ice Blue
SRC_HDR_FONT = "1E40AF"       # Dark Navy
LIGHT_BG_1 = "FFFFFF"         # Pure White
LIGHT_BG_2 = "F8FAFC"         # Soft Off-White / Slate 50
BORDER_COLOR = "CBD5E1"       # Slate 300
BORDER_HEADER = "94A3B8"      # Slate 400
TEXT_MAIN = "0F172A"          # Slate 900
TEXT_MUTED = "64748B"         # Slate 500
WHITE = "FFFFFF"
GREEN_ACTIVE = "15803D"       # Forest Green
AMBER_SPARE = "B45309"        # Amber 700

def thin_border():
    s = Side(style='thin', color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style='medium', color=BORDER_HEADER)
    return Border(left=s, right=s, top=s, bottom=s)

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

    records_c1_c2 = []
    for r in range(2, ws_io.max_row + 1):
        ch = ws_io.cell(r, ch_idx).value
        sl = ws_io.cell(r, sl_idx).value
        pt = ws_io.cell(r, pt_idx).value
        if ch is None or sl is None:
            continue
            
        c_str = str(ch).strip()
        if c_str not in ['C1', 'C2']:
            continue
            
        sl_int = int(sl) if str(sl).isdigit() else 0
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

        # Physical Rack Slot mapping (0-indexed backplane slots):
        # C1: S0=CPU, S1..S3=ETH4 -> raw slots 3..10 map to physical rack slots 4..11 (C1S4..C1S11)
        # C2: S0=ETH -> raw slots 1..13 map to physical rack slots 1..13 (C2S1..C2S13)
        if c_str == 'C1':
            phys_slot = sl_int + 1
        else:
            phys_slot = sl_int

        is_spare = False
        if not inst_tag or inst_tag.lower() in ['spare', 'none', '-'] or 'spare' in plc_tag.lower():
            is_spare = True

        records_c1_c2.append({
            'source_row': r,
            'chassis': c_str,
            'raw_slot': sl_int,
            'phys_slot': phys_slot,
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

    print(f"Loaded {len(records_c1_c2)} channel records for C1 & C2.")

    # Sort records logically by Chassis, Physical Slot, Point
    def sort_key(rec):
        c_num = int(rec['chassis'][1:])
        return (c_num, rec['phys_slot'], rec['point'])

    records_c1_c2.sort(key=sort_key)

    # Process Option 1: Direct Slot Correlation (TB-CxSx)
    compiled_records = []
    
    # Group by (chassis, phys_slot)
    slot_groups = defaultdict(list)
    for rec in records_c1_c2:
        slot_groups[(rec['chassis'], rec['phys_slot'])].append(rec)

    for (c_str, phys_slot) in sorted(slot_groups.keys()):
        slot_records = slot_groups[(c_str, phys_slot)]
        io_types = [r['io_type'] for r in slot_records]
        main_io = max(set(io_types), key=io_types.count) if io_types else 'DI'
        
        sx_str = f"S{phys_slot}"
        cx_sx = f"{c_str}{sx_str}"
        tb_strip_name = f"TB-{cx_sx}"   # Option 1: TB-C1S4, TB-C2S1, TB-C2S6...

        for rec in slot_records:
            pt = rec['point']
            ch = rec['channel']
            
            # Swing-arm terminal pin formula (Rockwell 1756-TBCH / 36-pin):
            # Points 0..15 -> Pins 1..16
            # Points 16..31 -> Pins 19..34 (Pins 17, 18, 35, 36 are 24VDC/0VDC commons)
            if pt < 16:
                plc_pin = pt + 1
            else:
                plc_pin = pt + 3

            # Determine destination terminal pin (TB-CxSx:Pin)
            if main_io in ['DI', 'DO']:
                # 32-point 2-level terminal blocks: 1A..16A (lower level), 1B..16B (upper level)
                if pt < 16:
                    tb_pin = f"{tb_strip_name}:{ch}A"
                else:
                    tb_pin = f"{tb_strip_name}:{ch - 16}B"
            elif main_io in ['AI', 'AO']:
                tb_pin = f"{tb_strip_name}:{ch}(+)"
            else:
                tb_pin = f"{tb_strip_name}:{ch}"

            plc_term = f"{cx_sx}:{plc_pin}"
            tag_term = f"{plc_term}/{tb_pin}"  # Field End Wire Mark
            tag_plc = f"{tb_pin}/{plc_term}"   # PLC Swing-arm Wire Mark
            rem = "Compiled Point" if not rec['is_spare'] else "Spare Channel"

            rec['sx_str'] = sx_str
            rec['cx_sx'] = cx_sx
            rec['tb_strip'] = tb_strip_name
            rec['tb_pin'] = tb_pin
            rec['plc_term'] = plc_term
            rec['tag_term'] = tag_term
            rec['tag_plc'] = tag_plc
            rec['wire_mark'] = tag_term
            rec['remarks'] = rem
            compiled_records.append(rec)

    print(f"Compiled {len(compiled_records)} total wire mark tags under Option 1 (TB-CxSx).")

    # =========================================================================
    # BUILD WORKBOOK
    # =========================================================================
    wb_out = openpyxl.Workbook()
    # Remove default sheet
    wb_out.remove(wb_out.active)

    # -------------------------------------------------------------------------
    # SHEET 1: Summary_By_Slot (Overview of C1 & C2 Slots)
    # -------------------------------------------------------------------------
    ws_sum = wb_out.create_sheet("Summary_By_Slot")
    ws_sum.views.sheetView[0].showGridLines = True

    ws_sum.merge_cells("A1:K1")
    s_title = ws_sum["A1"]
    s_title.value = "CHASSIS C1 & C2 CARD TERMINAL SUMMARY (OPTION 1: DIRECT SLOT CORRELATION TB-CxSx)"
    s_title.font = Font(name=FONT_FAMILY, size=13, bold=True, color=WHITE)
    s_title.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    s_title.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 36

    ws_sum.merge_cells("A2:K2")
    s_sub = ws_sum["A2"]
    s_sub.value = "Rockwell ControlLogix 1756 PLC | Terminal Strips Directly Matched 1:1 with Card Slot (e.g. TB-C1S4, TB-C2S1, TB-C2S6)"
    s_sub.font = Font(name=FONT_FAMILY, size=9.5, italic=True, color="E2E8F0")
    s_sub.fill = PatternFill("solid", fgColor=NAVY_SUBHEADER)
    s_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[2].height = 20

    sum_headers = [
        ("No.", 6),
        ("Chassis\n(Cx)", 10),
        ("Physical Slot\n(Sx)", 14),
        ("Slot Code\n(CxSx)", 14),
        ("I/O Type", 10),
        ("Module Card", 16),
        ("Total Points", 12),
        ("Active", 9),
        ("Spare", 9),
        ("Terminal Strip\n(TB-CxSx)", 18),
        ("Terminal Pin Range", 26)
    ]

    ws_sum.row_dimensions[4].height = 28
    for c_i, (h_name, w) in enumerate(sum_headers, 1):
        c = ws_sum.cell(4, c_i, h_name)
        c.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border()
        ws_sum.column_dimensions[get_column_letter(c_i)].width = w

    # Populate summary rows
    sum_row = 5
    card_no = 1
    for (c_str, p_sl) in sorted(slot_groups.keys()):
        s_recs = slot_groups[(c_str, p_sl)]
        io_t = s_recs[0]['io_type']
        card_m = s_recs[0]['card']
        tb_strip = s_recs[0]['tb_strip']
        total_pts = len(s_recs)
        active_pts = sum(1 for r in s_recs if not r['is_spare'])
        spare_pts = total_pts - active_pts
        
        first_tb = s_recs[0]['tb_pin']
        last_tb = s_recs[-1]['tb_pin']
        tb_range_str = f"{first_tb.split(':')[1]} .. {last_tb.split(':')[1]}"

        ws_sum.row_dimensions[sum_row].height = 20
        bg = LIGHT_BG_1 if card_no % 2 == 1 else LIGHT_BG_2

        row_vals = [
            (card_no, "center", False, None),
            (c_str, "center", True, "0F172A"),
            (f"Slot {p_sl}", "center", True, "0F172A"),
            (f"{c_str}S{p_sl}", "center", True, "1E3A8A"),
            (io_t, "center", True, "15803D" if io_t == 'DI' else ("B45309" if io_t == 'DO' else "0369A1")),
            (card_m, "center", False, None),
            (total_pts, "center", False, None),
            (active_pts, "center", True, GREEN_ACTIVE),
            (spare_pts, "center", True, AMBER_SPARE),
            (tb_strip, "center", True, "1E3A8A"),
            (tb_range_str, "center", False, None)
        ]

        for col_idx, (val, align, bold, f_color) in enumerate(row_vals, 1):
            cell = ws_sum.cell(sum_row, col_idx, val)
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.font = Font(name=FONT_FAMILY, size=9.5, bold=bold, color=f_color if f_color else TEXT_MAIN)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.border = thin_border()

        sum_row += 1
        card_no += 1

    # Freeze panes at A5
    ws_sum.freeze_panes = "A5"

    # -------------------------------------------------------------------------
    # SHEET 2: All_Marks_Sorted_CxSx (Master Schedule for C1 & C2)
    # -------------------------------------------------------------------------
    ws_all = wb_out.create_sheet("All_Marks_Sorted_CxSx")
    ws_all.views.sheetView[0].showGridLines = True

    ws_all.merge_cells("A1:Q1")
    a_title = ws_all["A1"]
    a_title.value = "C1 & C2 MASTER WIRE MARKS - OPTION 1: DIRECT SLOT CORRELATION (TB-CxSx)"
    a_title.font = Font(name=FONT_FAMILY, size=13, bold=True, color=WHITE)
    a_title.fill = PatternFill("solid", fgColor=NAVY_HEADER)
    a_title.alignment = Alignment(horizontal="center", vertical="center")
    ws_all.row_dimensions[1].height = 36

    ws_all.merge_cells("A2:Q2")
    a_sub = ws_all["A2"]
    a_sub.value = "Self-Documenting Terminal Strip Naming | Destination Terminal is Directly Named After PLC Slot (TB-C1S4..S11 & TB-C2S1..S13)"
    a_sub.font = Font(name=FONT_FAMILY, size=9.5, italic=True, color="E2E8F0")
    a_sub.fill = PatternFill("solid", fgColor=NAVY_SUBHEADER)
    a_sub.alignment = Alignment(horizontal="center", vertical="center")
    ws_all.row_dimensions[2].height = 20

    all_headers = [
        ("No.", 6),
        ("Chassis", 9),
        ("Slot (Sx)", 9),
        ("Chassis/Slot", 13),
        ("Ch.", 5),
        ("Pt.", 5),
        ("Wire Mark Tag\n(Field End)", 26),
        ("PLC Tag Mark\n(Swing Arm End)", 26),
        ("Terminal Strip\n(TB-CxSx:Pin)", 20),
        ("PLC Tag Name", 24),
        ("Instrument Tag", 18),
        ("Description / Service", 36),
        ("Destination\n(JB / Room)", 14),
        ("Signal Type", 12),
        ("Status", 10),
        ("P&ID Ref", 14),
        ("Remarks", 18)
    ]

    ws_all.row_dimensions[4].height = 28
    for c_i, (h_name, w) in enumerate(all_headers, 1):
        c = ws_all.cell(4, c_i, h_name)
        c.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
        c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border()
        ws_all.column_dimensions[get_column_letter(c_i)].width = w

    for r_i, rec in enumerate(compiled_records, 5):
        ws_all.row_dimensions[r_i].height = 18
        bg = LIGHT_BG_1 if r_i % 2 == 1 else LIGHT_BG_2

        row_data = [
            (r_i - 4, "center", False, None, None),
            (rec['chassis'], "center", True, "0F172A", None),
            (rec['sx_str'], "center", True, "0F172A", None),
            (rec['cx_sx'], "center", True, "1E3A8A", SRC_HDR_FILL),
            (rec['channel'], "center", False, None, None),
            (rec['point'], "center", False, None, None),
            (rec['tag_term'], "center", True, "1E3A8A", None),
            (rec['tag_plc'], "center", True, "0F172A", None),
            (rec['tb_pin'], "center", True, "15803D", None),
            (rec['plc_tag'], "left", False, None, None),
            (rec['inst_tag'], "left", True, None, None),
            (rec['desc'], "left", False, None, None),
            (rec['dest'], "center", False, None, None),
            (rec['signal_type'], "center", False, None, None),
            (rec['status'], "center", True, GREEN_ACTIVE if rec['status'] == 'ACTIVE' else AMBER_SPARE, None),
            (rec['pid'], "center", False, None, None),
            (rec['remarks'], "left", False, TEXT_MUTED, None)
        ]

        for col_idx, (val, align, bold, f_color, custom_fill) in enumerate(row_data, 1):
            cell = ws_all.cell(r_i, col_idx, val)
            cell.alignment = Alignment(horizontal=align, vertical="center")
            cell.font = Font(name=FONT_FAMILY, size=9.5, bold=bold, color=f_color if f_color else TEXT_MAIN)
            cell.fill = PatternFill("solid", fgColor=custom_fill if custom_fill else bg)
            cell.border = thin_border()

    ws_all.auto_filter.ref = f"A4:Q{len(compiled_records) + 4}"
    ws_all.freeze_panes = "A5"

    # -------------------------------------------------------------------------
    # SHEET 3: Chassis 1 (MCP-M1) Full Schedule
    # SHEET 4: Chassis 2 (MCP-M2) Full Schedule
    # -------------------------------------------------------------------------
    for ch_name, sheet_title in [('C1', 'Chassis 1 (MCP-M1)'), ('C2', 'Chassis 2 (MCP-M2)')]:
        ws_ch = wb_out.create_sheet(ch_name)
        ws_ch.views.sheetView[0].showGridLines = True

        ws_ch.merge_cells("A1:K1")
        ch_t = ws_ch["A1"]
        ch_t.value = f"CONTROL CABINET {ch_name} WIRE SCHEDULE - OPTION 1 (TB-{ch_name}Sx)"
        ch_t.font = Font(name=FONT_FAMILY, size=13, bold=True, color=WHITE)
        ch_t.fill = PatternFill("solid", fgColor=NAVY_HEADER)
        ch_t.alignment = Alignment(horizontal="center", vertical="center")
        ws_ch.row_dimensions[1].height = 36

        ch_recs = [r for r in compiled_records if r['chassis'] == ch_name]

        ch_headers = [
            ("No.", 6),
            ("Slot", 9),
            ("Ch.", 5),
            ("Pt.", 5),
            ("Field Wire Mark Tag\n(CxSx/TB-CxSx)", 26),
            ("PLC Wire Mark Tag\n(TB-CxSx/CxSx)", 26),
            ("Terminal Strip & Pin\n(TB-CxSx:Pin)", 20),
            ("PLC Tag Name", 24),
            ("Instrument Tag", 18),
            ("Description", 36),
            ("Destination", 14)
        ]

        ws_ch.row_dimensions[3].height = 26
        for c_i, (h_name, w) in enumerate(ch_headers, 1):
            c = ws_ch.cell(3, c_i, h_name)
            c.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = header_border()
            ws_ch.column_dimensions[get_column_letter(c_i)].width = w

        c_row = 4
        # Group by physical slot
        grouped_by_slot = defaultdict(list)
        for r in ch_recs:
            grouped_by_slot[r['phys_slot']].append(r)

        card_k = 1
        for p_sl in sorted(grouped_by_slot.keys()):
            s_recs = grouped_by_slot[p_sl]
            c_type = s_recs[0]['io_type']
            sx_lbl = s_recs[0]['sx_str']
            tb_strip = s_recs[0]['tb_strip']
            
            # Card Section Break Header
            ws_ch.row_dimensions[c_row].height = 22
            ws_ch.merge_cells(start_row=c_row, start_column=1, end_row=c_row, end_column=11)
            hdr_cell = ws_ch.cell(c_row, 1, f"CARD {card_k}: {c_type} (Slot {sx_lbl})  |  Terminal Block: {tb_strip}  ({len(s_recs)} Channels)")
            hdr_cell.font = Font(name=FONT_FAMILY, size=10, bold=True, color=SRC_HDR_FONT)
            hdr_cell.fill = PatternFill("solid", fgColor=SRC_HDR_FILL)
            hdr_cell.alignment = Alignment(horizontal="left", vertical="center")
            for ci in range(1, 12):
                ws_ch.cell(c_row, ci).border = thin_border()
            c_row += 1

            for rec_idx, r in enumerate(s_recs, 1):
                # For 32-point cards, insert internal commons at Pin 17/18 and 35/36
                if len(s_recs) == 32 and rec_idx == 17:
                    pwr_val = "24VDC" if c_type == 'DO' else "0VDC"
                    for p_v in [pwr_val, "0VDC"]:
                        ws_ch.row_dimensions[c_row].height = 17
                        ws_ch.cell(c_row, 1, "-").alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 2, sx_lbl).alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 5, p_v).alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 6, p_v).alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 7, f"{tb_strip}:{p_v}").alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 10, "Internal Power Common (ยังไม่ใส่มาร์ก)").alignment = Alignment(horizontal="left", vertical="center")
                        for ci in range(1, 12):
                            ws_ch.cell(c_row, ci).border = thin_border()
                            ws_ch.cell(c_row, ci).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                        c_row += 1

                ws_ch.row_dimensions[c_row].height = 18
                bg = LIGHT_BG_1 if rec_idx % 2 == 1 else LIGHT_BG_2

                ch_row_vals = [
                    (rec_idx, "center", False, None),
                    (sx_lbl, "center", True, "0F172A"),
                    (r['channel'], "center", False, None),
                    (r['point'], "center", False, None),
                    (r['tag_term'], "center", True, "1E3A8A"),
                    (r['tag_plc'], "center", True, "0F172A"),
                    (r['tb_pin'], "center", True, "15803D"),
                    (r['plc_tag'], "left", False, None),
                    (r['inst_tag'], "left", True, None),
                    (r['desc'], "left", False, None),
                    (r['dest'], "center", False, None)
                ]

                for col_idx, (val, align, bold, f_color) in enumerate(ch_row_vals, 1):
                    cell = ws_ch.cell(c_row, col_idx, val)
                    cell.alignment = Alignment(horizontal=align, vertical="center")
                    cell.font = Font(name=FONT_FAMILY, size=9, bold=bold, color=f_color if f_color else TEXT_MAIN)
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.border = thin_border()
                c_row += 1

                if len(s_recs) == 32 and rec_idx == 32:
                    pwr_val = "24VDC" if c_type == 'DO' else "0VDC"
                    for p_v in [pwr_val, "0VDC"]:
                        ws_ch.row_dimensions[c_row].height = 17
                        ws_ch.cell(c_row, 1, "-").alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 2, sx_lbl).alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 5, p_v).alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 6, p_v).alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 7, f"{tb_strip}:{p_v}").alignment = Alignment(horizontal="center", vertical="center")
                        ws_ch.cell(c_row, 10, "Internal Power Common (ยังไม่ใส่มาร์ก)").alignment = Alignment(horizontal="left", vertical="center")
                        for ci in range(1, 12):
                            ws_ch.cell(c_row, ci).border = thin_border()
                            ws_ch.cell(c_row, ci).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                        c_row += 1

            card_k += 1

        ws_ch.freeze_panes = "A4"

    # -------------------------------------------------------------------------
    # SHEET 5..8: Classic ePlan Sheets (DI Chassis 1, DO Chassis 1, DI Chassis 2, DO Chassis 2)
    # -------------------------------------------------------------------------
    classic_tabs = [
        ("DI Chassis 1", "C1", "DI", "Tag DI PLC", "Tag DI Terminal"),
        ("DO Chassis 1", "C1", "DO", "Tag DO PLC", "Tag DO Terminal"),
        ("DI Chassis 2", "C2", "DI", "Tag DI PLC", "Tag DI Terminal"),
        ("DO Chassis 2", "C2", "DO", "Tag DO PLC", "Tag DO Terminal")
    ]

    for (t_name, ch_code, io_code, col_b_title, col_c_title) in classic_tabs:
        ws_t = wb_out.create_sheet(t_name)
        ws_t.views.sheetView[0].showGridLines = True

        t_headers = [
            ("No.", 7),
            (col_b_title, 26),
            (col_c_title, 26),
            ("หมายเหตุ / Remarks", 20),
            ("PLC Tag Name", 24),
            ("Instrument Tag", 18),
            ("Description", 36),
            ("Destination", 14)
        ]

        ws_t.row_dimensions[1].height = 28
        for c_i, (h_name, w) in enumerate(t_headers, 1):
            c = ws_t.cell(1, c_i, h_name)
            c.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=WHITE)
            c.fill = PatternFill("solid", fgColor=NAVY_HEADER)
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = header_border()
            ws_t.column_dimensions[get_column_letter(c_i)].width = w

        tab_recs = [r for r in compiled_records if r['chassis'] == ch_code and r['io_type'] == io_code]
        grouped_tab = defaultdict(list)
        for r in tab_recs:
            grouped_tab[r['phys_slot']].append(r)

        cur_r = 2
        card_counter = 1
        for p_sl in sorted(grouped_tab.keys()):
            card_recs = grouped_tab[p_sl]
            sx_lbl = card_recs[0]['sx_str']
            tb_strip = card_recs[0]['tb_strip']

            if card_counter > 1:
                # Spacer
                ws_t.row_dimensions[cur_r].height = 14
                cur_r += 1
                # Card Section Break
                ws_t.row_dimensions[cur_r].height = 20
                ws_t.merge_cells(start_row=cur_r, start_column=1, end_row=cur_r, end_column=8)
                lbl = ws_t.cell(cur_r, 1, f"{io_code} Card {card_counter} (Slot {sx_lbl})  |  Terminal Block: {tb_strip}")
                lbl.font = Font(name=FONT_FAMILY, size=9.5, bold=True, color=SRC_HDR_FONT)
                lbl.fill = PatternFill("solid", fgColor=SRC_HDR_FILL)
                lbl.alignment = Alignment(horizontal="left", vertical="center")
                for ci in range(1, 9):
                    ws_t.cell(cur_r, ci).border = thin_border()
                cur_r += 1

            no_counter = 1
            for rec_idx, r in enumerate(card_recs, 1):
                # Common rows
                if len(card_recs) == 32 and rec_idx == 17:
                    pwr_v = "24VDC" if io_code == 'DO' else "0VDC"
                    for pv in [pwr_v, "0VDC"]:
                        ws_t.row_dimensions[cur_r].height = 17
                        ws_t.cell(cur_r, 1, float(no_counter)).alignment = Alignment(horizontal="center", vertical="center")
                        ws_t.cell(cur_r, 2, pv).alignment = Alignment(horizontal="center", vertical="center")
                        ws_t.cell(cur_r, 3, pv).alignment = Alignment(horizontal="center", vertical="center")
                        ws_t.cell(cur_r, 4, "ยังไม่ใส่มาร์ก (Internal Power Common)").alignment = Alignment(horizontal="left", vertical="center")
                        for ci in range(1, 9):
                            ws_t.cell(cur_r, ci).border = thin_border()
                            ws_t.cell(cur_r, ci).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                        cur_r += 1
                        no_counter += 1

                ws_t.row_dimensions[cur_r].height = 18
                bg = LIGHT_BG_1 if rec_idx % 2 == 1 else LIGHT_BG_2

                tab_row = [
                    (float(no_counter), "center", False, None),
                    (r['tag_plc'], "center", True, "0F172A"),
                    (r['tag_term'], "center", True, "1E3A8A"),
                    (r['remarks'], "left", False, TEXT_MUTED),
                    (r['plc_tag'], "left", False, None),
                    (r['inst_tag'], "left", True, None),
                    (r['desc'], "left", False, None),
                    (r['dest'], "center", False, None)
                ]

                for ci, (val, align, bold, f_color) in enumerate(tab_row, 1):
                    cell = ws_t.cell(cur_r, ci, val)
                    cell.alignment = Alignment(horizontal=align, vertical="center")
                    cell.font = Font(name=FONT_FAMILY, size=9, bold=bold, color=f_color if f_color else TEXT_MAIN)
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.border = thin_border()
                cur_r += 1
                no_counter += 1

                if len(card_recs) == 32 and rec_idx == 32:
                    pwr_v = "24VDC" if io_code == 'DO' else "0VDC"
                    for pv in [pwr_v, "0VDC"]:
                        ws_t.row_dimensions[cur_r].height = 17
                        ws_t.cell(cur_r, 1, float(no_counter)).alignment = Alignment(horizontal="center", vertical="center")
                        ws_t.cell(cur_r, 2, pv).alignment = Alignment(horizontal="center", vertical="center")
                        ws_t.cell(cur_r, 3, pv).alignment = Alignment(horizontal="center", vertical="center")
                        ws_t.cell(cur_r, 4, "ยังไม่ใส่มาร์ก (Internal Power Common)").alignment = Alignment(horizontal="left", vertical="center")
                        for ci in range(1, 9):
                            ws_t.cell(cur_r, ci).border = thin_border()
                            ws_t.cell(cur_r, ci).font = Font(name=FONT_FAMILY, size=8.5, italic=True, color="64748B")
                        cur_r += 1
                        no_counter += 1

            card_counter += 1

        ws_t.freeze_panes = "A2"

    # Save to both target directories
    os.makedirs(os.path.dirname(OUT_EPLAN_PATH), exist_ok=True)
    os.makedirs(os.path.dirname(OUT_SCHEDULE_PATH), exist_ok=True)

    print(f"Saving to {OUT_EPLAN_PATH}...")
    wb_out.save(OUT_EPLAN_PATH)
    print(f"Saving to {OUT_SCHEDULE_PATH}...")
    wb_out.save(OUT_SCHEDULE_PATH)
    print("Generation of Option 1 Workbook (C1, C2) completed successfully!")

if __name__ == "__main__":
    main()
