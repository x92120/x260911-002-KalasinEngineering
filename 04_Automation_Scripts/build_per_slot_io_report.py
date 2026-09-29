#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
CHASSIS I/O REPORT: ONE SHEET PER SLOT (SOURCE & DESTINATION)
================================================================================
Generates:
  1. Master Workbook: 03_IO_Lists_and_Schedules/Chassis_IO_List_Per_Slot_Report.xlsx
     - Sheet '00_Slot_Index': Master Navigation Hub with Hyperlinks to every Slot sheet.
     - 47 Dedicated Slot Sheets (e.g., 'C1_Slot_03', 'C1_Slot_04', ..., 'C7_Slot_07').
     - Each slot sheet includes 'Return to Index' hyperlink, full Source & Destination specs.
  2. Per-Chassis Workbooks in folder 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/:
     - Chassis_C1_Slots.xlsx
     - Chassis_C2_Slots.xlsx
     - Chassis_C3_Slots.xlsx
     - Chassis_C4_Slots.xlsx
     - Chassis_C5_Slots.xlsx
     - Chassis_C6_Slots.xlsx
     - Chassis_C7_Slots.xlsx

Styling: Light theme, engineering standard, openpyxl professional formatting.
================================================================================
"""

import os
import sys
import re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_LIST_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-6.xlsx")
TAGS_PLC_FILE = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC.xlsx")
OUTPUT_MASTER = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Chassis_IO_List_Per_Slot_Report.xlsx")
OUTPUT_DIR_CHASSIS = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Slot_by_Slot_Reports")

os.makedirs(OUTPUT_DIR_CHASSIS, exist_ok=True)

# Theme Palette (Light Executive Standard)
NAVY_HEADER = "1E293B"      # Deep Slate Navy
NAVY_SUBHEADER = "334155"   # Slate Subheader
WHITE = "FFFFFF"
LIGHT_BG_1 = "FFFFFF"
LIGHT_BG_2 = "F8FAFC"       # Very light slate/gray
BORDER_COLOR = "CBD5E1"     # Light slate border

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

CHASSIS_METADATA = {
    'C1': {
        'name': 'Main Controller Rack (Chassis 1)',
        'panel': 'Control Cabinet CA1 (Main Control Room)',
        'role': '1756-L950TPSXT Controller + Local Fast I/O',
        'network': 'Stratix 5400 / EN4TR Ring Node 1 (192.168.1.10)',
    },
    'C2': {
        'name': 'I/O Expansion Rack 1 (Chassis 2)',
        'panel': 'Control Cabinet CA1 (Bay 2)',
        'role': 'Digital & Analog Process I/O Expansion',
        'network': 'EN4TR DLR Ring Node 2 (192.168.1.11)',
    },
    'C3': {
        'name': 'I/O Expansion Rack 2 (Chassis 3)',
        'panel': 'Control Cabinet CA1 (Bay 3)',
        'role': 'Evaporator, Concentrator & Dosing Skid I/O',
        'network': 'EN4TR DLR Ring Node 3 (192.168.1.12)',
    },
    'C4': {
        'name': 'I/O Expansion Rack 3 (Chassis 4)',
        'panel': 'Control Cabinet CA1 (Bay 4)',
        'role': 'Spray Dryer Tower & Exhaust Auxiliary I/O',
        'network': 'EN4TR DLR Ring Node 4 (192.168.1.13)',
    },
    'C5': {
        'name': 'Remote I/O Skid RIO-200 (Chassis 5)',
        'panel': 'Remote I/O Enclosure RIO-200 (Process Floor 1F)',
        'role': 'Remote Field I/O via Fiber Optic Trunk (DLR Ring Node 5)',
        'network': 'EN4TR Fiber DLR Ring Node 5 (192.168.1.14)',
    },
    'C6': {
        'name': 'MCC Auxiliary I/O (Chassis 6)',
        'panel': 'Motor Control Center Room (MCC Panel)',
        'role': 'Direct Hardwired Motor Starter Feedback & Permissives',
        'network': 'MCC Hardwired Trunk Line',
    },
    'C7': {
        'name': 'MCC Fieldbus & Network Interface (Chassis 7)',
        'panel': 'Motor Control Center Room (MCC Panel & Field Bus)',
        'role': 'Modbus/Fieldbus Gateways & MCC Feeder Interlocks',
        'network': 'Modbus TCP / EtherNet Trunk',
    }
}

def thin_border():
    s = Side(style='thin', color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style='medium', color="94A3B8")
    return Border(left=s, right=s, top=s, bottom=s)

def load_eplan_tags(tags_file):
    wb = openpyxl.load_workbook(tags_file, data_only=True)
    tags_map = {}
    for sname in wb.sheetnames:
        ws = wb[sname]
        for r in range(2, ws.max_row + 1):
            v_plc = ws.cell(row=r, column=2).value
            v_term = ws.cell(row=r, column=3).value
            note = ws.cell(row=r, column=4).value
            if not v_plc and not v_term:
                continue
            combined = f"{v_plc or ''} {v_term or ''}"
            m = re.search(r'C(\d+)S(\d+):(\d+)', combined)
            if m:
                ch = f"C{m.group(1)}"
                slot = int(m.group(2))
                chan = int(m.group(3))
                tags_map[(ch, slot, chan)] = {
                    'plc_tag': str(v_plc).strip() if v_plc else '',
                    'term_tag': str(v_term).strip() if v_term else '',
                    'note': str(note).strip() if note else ''
                }
    return tags_map

def load_data():
    eplan_tags = load_eplan_tags(TAGS_PLC_FILE)
    wb = openpyxl.load_workbook(IO_LIST_FILE, data_only=True)
    ws = wb['IO List']
    
    rows = []
    for r in range(2, ws.max_row + 1):
        chassis = ws.cell(row=r, column=32).value
        slot = ws.cell(row=r, column=33).value
        point = ws.cell(row=r, column=34).value
        if chassis is None or slot is None:
            continue
            
        c_str = str(chassis).strip()
        s_int = int(slot) if str(slot).isdigit() else str(slot).strip()
        pt_int = int(point) if point is not None and str(point).isdigit() else 0
        chan_int = pt_int + 1
        
        card = ws.cell(row=r, column=28).value
        iotype = ws.cell(row=r, column=22).value
        plc_tag = ws.cell(row=r, column=21).value
        inst_tag = ws.cell(row=r, column=12).value
        desc = ws.cell(row=r, column=13).value
        dest = ws.cell(row=r, column=2).value
        loc = ws.cell(row=r, column=9).value
        floor = ws.cell(row=r, column=14).value
        zone = ws.cell(row=r, column=11).value
        term = ws.cell(row=r, column=6).value
        termlbl = ws.cell(row=r, column=8).value
        term2 = ws.cell(row=r, column=35).value
        pid = ws.cell(row=r, column=16).value or ws.cell(row=r, column=15).value
        sig = ws.cell(row=r, column=29).value
        ctrl_desc = ws.cell(row=r, column=17).value
        
        is_spare = False
        inst_str = str(inst_tag or '').strip().lower()
        tag_str = str(plc_tag or '').strip().lower()
        if 'spare' in inst_str or 'spare' in tag_str or not inst_tag or inst_str == 'none':
            is_spare = True
            
        eplan_key = (c_str, s_int, chan_int)
        if eplan_key in eplan_tags:
            eplan_plc = eplan_tags[eplan_key]['plc_tag']
            eplan_term = eplan_tags[eplan_key]['term_tag']
        else:
            eplan_plc = f"{term}:{chan_int}/{c_str}S{s_int}:{chan_int}" if term else f"{c_str}S{s_int}:{chan_int}"
            eplan_term = f"{c_str}S{s_int}:{chan_int}/{term}:{chan_int}" if term else f"{c_str}S{s_int}:{chan_int}"
            
        rows.append({
            'chassis': c_str,
            'slot': s_int,
            'point': pt_int,
            'channel': chan_int,
            'card': str(card).strip() if card else '',
            'iotype': str(iotype).strip().upper() if iotype else '',
            'plc_tag': str(plc_tag).strip() if plc_tag else '',
            'term2': str(term2).strip() if term2 else '',
            'is_spare': is_spare,
            'status': 'SPARE' if is_spare else 'ACTIVE',
            'destination': str(dest).strip() if dest else '',
            'location': str(loc).strip() if loc else '',
            'floor': str(floor).strip() if floor else '',
            'zone': str(zone).strip() if zone else '',
            'terminal': str(term).strip() if term else '',
            'terminal_label': str(termlbl).strip() if termlbl else '',
            'instrument_tag': str(inst_tag).strip() if inst_tag else ('Spare' if is_spare else ''),
            'description': str(desc).strip() if desc else ('Spare Channel' if is_spare else ''),
            'pid': str(pid).strip() if pid else '',
            'signal_type': str(sig).strip() if sig else '',
            'control_desc': str(ctrl_desc).strip() if ctrl_desc else '',
            'eplan_plc_tag': eplan_plc,
            'eplan_term_tag': eplan_term
        })
        
    return rows

def populate_slot_sheet(ws, ch, sl, pts, is_master=True):
    ws.views.sheetView[0].showGridLines = True
    meta = CHASSIS_METADATA.get(ch, {})
    ch_name = meta.get('name', f"Chassis {ch}")
    ch_panel = meta.get('panel', 'Main Control Panel')
    ch_role = meta.get('role', 'Process Control & I/O')
    ch_net = meta.get('network', 'EtherNet/IP DLR Ring')
    
    act_cnt = sum(1 for p in pts if not p['is_spare'])
    spr_cnt = len(pts) - act_cnt
    spr_pct = (spr_cnt / len(pts) * 100) if pts else 0
    card_model = re.sub(r'-\d+$', '', pts[0]['card']) if pts and pts[0]['card'] else 'I/O Module'
    iotype = pts[0]['iotype'] if pts else 'I/O'
    dests = sorted(list({p['destination'] for p in pts if p['destination']}))
    dest_str = ", ".join(dests) if dests else "-"
    
    # Title Block
    ws.merge_cells("A1:S1")
    t = ws["A1"]
    t.value = f"KALASIN ENGINEERING — {ch.upper()} SLOT {sl:02d} ({card_model} • {iotype})"
    t.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    t.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    
    # Subheader Banner
    ws.merge_cells("A2:Q2")
    st = ws["A2"]
    st.value = f"Chassis: {ch} ({ch_name}) | Panel: {ch_panel} | Net: {ch_net} | Total Points: {len(pts)} (Active: {act_cnt}, Spare: {spr_cnt} / {spr_pct:.1f}%)"
    st.font = Font(name="Arial", size=9, bold=True, color=WHITE)
    st.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    
    # Navigation Link Back to Index
    if is_master:
        ws.merge_cells("R2:S2")
        nav = ws["R2"]
        nav.value = "⬅ Return to Slot Index"
        nav.hyperlink = "#'00_Slot_Index'!A1"
        nav.font = Font(name="Arial", size=9, bold=True, color="1E3A8A", underline="single")
        nav.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
        nav.alignment = Alignment(horizontal="center", vertical="center")
    else:
        ws.merge_cells("R2:S2")
        nav = ws["R2"]
        nav.value = f"Destination: {dest_str[:20]}"
        nav.font = Font(name="Arial", size=8, bold=True, color=WHITE)
        nav.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        nav.alignment = Alignment(horizontal="center", vertical="center")
        
    ws.row_dimensions[2].height = 22
    
    # Dual Level Headers (Row 4 & 5)
    ws.merge_cells("A4:H4")
    g1 = ws["A4"]
    g1.value = "SOURCE SPECIFICATION (PLC RACK / SLOT / TERMINAL / ePLAN WIRE TAG)"
    g1.font = Font(name="Arial", size=9, bold=True, color=SRC_HDR_FONT)
    g1.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
    g1.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("I4:K4")
    g2 = ws["I4"]
    g2.value = "SIGNAL SPECIFICATION"
    g2.font = Font(name="Arial", size=9, bold=True, color=SYS_HDR_FONT)
    g2.fill = PatternFill(start_color=SYS_HDR_FILL, end_color=SYS_HDR_FILL, fill_type="solid")
    g2.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("L4:S4")
    g3 = ws["L4"]
    g3.value = "DESTINATION SPECIFICATION (FIELD JB / MCC / TERMINAL PIN / FIELD DEVICE)"
    g3.font = Font(name="Arial", size=9, bold=True, color=DEST_HDR_FONT)
    g3.fill = PatternFill(start_color=DEST_HDR_FILL, end_color=DEST_HDR_FILL, fill_type="solid")
    g3.alignment = Alignment(horizontal="center", vertical="center")
    
    for c_k in range(1, 20):
        ws.cell(row=4, column=c_k).border = header_border()
    ws.row_dimensions[4].height = 22
    
    headers = [
        "Point Ref ID", "Slot", "Pt", "Ch", "Module Model", "PLC Tag Name",
        "PLC Card Pin", "ePlan Wire Tag",
        "I/O Type", "Signal Spec", "Point Status",
        "Dest Enclosure", "Location / Area", "Floor / Zone", "Terminal Block",
        "Term Pin", "Instrument Tag", "Instrument Description", "P&ID Drawing No."
    ]
    
    ws.row_dimensions[5].height = 24
    for col_idx, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=5, column=col_idx, value=h_text)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border()
        
    cur_row = 6
    for idx, r in enumerate(pts, start=1):
        pt_ref = f"{r['chassis']}-S{r['slot']:02d}-P{r['point']:02d}"
        flr_zn = f"{r['floor']} / {r['zone']}" if r['floor'] and r['zone'] else (r['floor'] or r['zone'] or '-')
        
        row_vals = [
            pt_ref,
            r['slot'],
            r['point'],
            r['channel'],
            re.sub(r'-\d+$', '', r['card']),
            r['plc_tag'],
            r['term2'] or f"Pin {r['channel']}",
            r['eplan_term_tag'] or r['eplan_plc_tag'],
            r['iotype'],
            r['signal_type'] or ("24VDC Dry Contact" if r['iotype'] == 'DI' else ("24VDC Sourcing" if r['iotype'] == 'DO' else "4-20mA HART")),
            r['status'],
            r['destination'] or "-",
            r['location'] or "-",
            flr_zn,
            r['terminal'] or "-",
            r['terminal_label'] or "-",
            r['instrument_tag'],
            r['description'],
            r['pid'] or "-"
        ]
        
        is_even = (idx % 2 == 0)
        row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
        
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws.cell(row=cur_row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=9, bold=(col_idx in (1, 6, 11, 17)))
            c.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
            c.border = thin_border()
            
            if col_idx in (1, 2, 3, 4, 5, 7, 9, 10, 11, 12, 14, 16, 19):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
            if col_idx == 11:
                if r['status'] == "ACTIVE":
                    c.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                    c.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
                else:
                    c.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                    c.font = Font(name="Arial", size=9, bold=True, color=SPARE_FONT)
                    
        ws.row_dimensions[cur_row].height = 19
        cur_row += 1
        
    ws.freeze_panes = "F6"
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in (1, 2, 3, 4):
                continue
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
        
    ws.column_dimensions['A'].width = 16
    ws.column_dimensions['F'].width = 24
    ws.column_dimensions['H'].width = 22
    ws.column_dimensions['L'].width = 18
    ws.column_dimensions['Q'].width = 20
    ws.column_dimensions['R'].width = 36

def build_reports():
    print("Loading data...")
    rows = load_data()
    print(f"Loaded {len(rows)} data rows.")
    
    # Group by (chassis, slot)
    def slot_sort_key(item):
        ch, sl = item
        ch_num = int(ch[1:]) if ch.startswith('C') and ch[1:].isdigit() else 99
        sl_num = int(sl) if isinstance(sl, int) or str(sl).isdigit() else 99
        return (ch_num, sl_num)
        
    slot_groups = defaultdict(list)
    for r in rows:
        slot_groups[(r['chassis'], r['slot'])].append(r)
        
    sorted_slots = sorted(slot_groups.keys(), key=slot_sort_key)
    print(f"Total slots to generate: {len(sorted_slots)}")
    
    # -------------------------------------------------------------------------
    # 1. Master Workbook (00_Slot_Index + 47 Slot Sheets)
    # -------------------------------------------------------------------------
    print("\nCreating Master Workbook with 1 Sheet per Slot...")
    wb_master = openpyxl.Workbook()
    wb_master.remove(wb_master.active) # Remove default sheet
    
    # Index Sheet
    ws_idx = wb_master.create_sheet(title="00_Slot_Index")
    ws_idx.views.sheetView[0].showGridLines = True
    
    # Title
    ws_idx.merge_cells("A1:K1")
    t1 = ws_idx["A1"]
    t1.value = "KALASIN ENGINEERING — xCIP AUTOMATION SYSTEM (Ref: x2608003)"
    t1.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    t1.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws_idx.row_dimensions[1].height = 32
    
    ws_idx.merge_cells("A2:K2")
    st1 = ws_idx["A2"]
    st1.value = "INTERACTIVE I/O SLOT DIRECTORY — CLICK ANY SLOT TO OPEN ITS DETAILED I/O SCHEDULE"
    st1.font = Font(name="Arial", size=11, bold=True, color=WHITE)
    st1.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    st1.alignment = Alignment(horizontal="center", vertical="center")
    ws_idx.row_dimensions[2].height = 24
    
    # Quick KPI row
    total_pts = len(rows)
    total_act = sum(1 for r in rows if not r['is_spare'])
    total_spr = total_pts - total_act
    
    kpis = [
        ("Total Managed Slots", f"{len(sorted_slots)} Active Slots", "A", "B", "1E3A8A", "DBEAFE"),
        ("Total I/O Capacity", f"{total_pts:,} Channels", "C", "D", "0F172A", "E2E8F0"),
        ("Active Assigned Points", f"{total_act:,} Points", "E", "F", "14532D", "DCFCE7"),
        ("Engineering Spare Points", f"{total_spr:,} Points ({(total_spr/total_pts*100):.1f}%)", "G", "H", "854D0E", "FEF9C3"),
        ("Chassis Spread", "7 Racks (C1 to C7)", "I", "K", "581C87", "F3E8FF")
    ]
    for title, val, c_s, c_e, f_col, b_col in kpis:
        ws_idx.merge_cells(f"{c_s}4:{c_e}4")
        ws_idx.merge_cells(f"{c_s}5:{c_e}5")
        c4 = ws_idx[f"{c_s}4"]
        c4.value = title.upper()
        c4.font = Font(name="Arial", size=8, bold=True, color=f_col)
        c4.fill = PatternFill(start_color=b_col, end_color=b_col, fill_type="solid")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        
        c5 = ws_idx[f"{c_s}5"]
        c5.value = val
        c5.font = Font(name="Arial", size=11, bold=True, color=f_col)
        c5.fill = PatternFill(start_color=b_col, end_color=b_col, fill_type="solid")
        c5.alignment = Alignment(horizontal="center", vertical="center")
        
        col_s_idx = openpyxl.utils.column_index_from_string(c_s)
        col_e_idx = openpyxl.utils.column_index_from_string(c_e)
        for r_k in (4, 5):
            for c_k in range(col_s_idx, col_e_idx + 1):
                ws_idx.cell(row=r_k, column=c_k).border = thin_border()
                
    ws_idx.row_dimensions[4].height = 18
    ws_idx.row_dimensions[5].height = 24
    
    # Table Header
    headers_idx = [
        "Chassis", "Slot No.", "Module Model", "I/O Type", "Capacity",
        "Active Points", "Spare Points", "Spare %", "Primary Destinations",
        "Enclosure & Location", "Interactive Sheet Link"
    ]
    ws_idx.row_dimensions[7].height = 26
    for col_idx, h_text in enumerate(headers_idx, start=1):
        c = ws_idx.cell(row=7, column=col_idx, value=h_text)
        c.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border()
        
    cur_idx_row = 8
    last_ch = None
    
    for (ch, sl) in sorted_slots:
        pts = slot_groups[(ch, sl)]
        # Sort points within slot
        pts.sort(key=lambda x: x['point'])
        
        sheet_title = f"{ch}_Slot_{sl:02d}"
        
        # Insert Chassis divider row if chassis changed
        if ch != last_ch:
            last_ch = ch
            meta = CHASSIS_METADATA.get(ch, {})
            ws_idx.merge_cells(start_row=cur_idx_row, start_column=1, end_row=cur_idx_row, end_column=11)
            gh = ws_idx.cell(row=cur_idx_row, column=1)
            gh.value = f"▶ {ch} — {meta.get('name', ch).upper()} | Panel: {meta.get('panel', '-')} | Net: {meta.get('network', '-')}"
            gh.font = Font(name="Arial", size=9, bold=True, color=SRC_HDR_FONT)
            gh.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
            gh.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c_k in range(1, 12):
                ws_idx.cell(row=cur_idx_row, column=c_k).border = thin_border()
            ws_idx.row_dimensions[cur_idx_row].height = 22
            cur_idx_row += 1
            
        act_cnt = sum(1 for p in pts if not p['is_spare'])
        spr_cnt = len(pts) - act_cnt
        spr_pct = (spr_cnt / len(pts) * 100) if pts else 0
        card_model = re.sub(r'-\d+$', '', pts[0]['card']) if pts and pts[0]['card'] else 'I/O Module'
        iotype = pts[0]['iotype'] if pts else 'I/O'
        dests = sorted(list({p['destination'] for p in pts if p['destination']}))
        dest_str = ", ".join(dests) if dests else "-"
        loc_str = pts[0]['location'] or CHASSIS_METADATA.get(ch, {}).get('panel', '-')
        
        idx_vals = [
            ch,
            sl,
            card_model,
            iotype,
            len(pts),
            act_cnt,
            spr_cnt,
            f"{spr_pct:.1f}%",
            dest_str,
            loc_str,
            f"Open {sheet_title} ➔"
        ]
        
        is_even = (cur_idx_row % 2 == 0)
        row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
        
        for col_idx, val in enumerate(idx_vals, start=1):
            c = ws_idx.cell(row=cur_idx_row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=9, bold=(col_idx in (1, 2, 3, 11)))
            c.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
            c.border = thin_border()
            
            if col_idx in (1, 2, 4, 5, 6, 7, 8):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 11:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.hyperlink = f"#'{sheet_title}'!A1"
                c.font = Font(name="Arial", size=9, bold=True, color="1E3A8A", underline="single")
                c.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
        ws_idx.row_dimensions[cur_idx_row].height = 20
        cur_idx_row += 1
        
        # Create the dedicated slot sheet in Master Workbook
        ws_slot = wb_master.create_sheet(title=sheet_title)
        populate_slot_sheet(ws_slot, ch, sl, pts, is_master=True)
        
    ws_idx.freeze_panes = "C8"
    for col in ws_idx.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in (1, 2, 4, 5):
                continue
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws_idx.column_dimensions[col_letter].width = max(max_len + 3, 11)
        
    ws_idx.column_dimensions['A'].width = 10
    ws_idx.column_dimensions['B'].width = 10
    ws_idx.column_dimensions['C'].width = 16
    ws_idx.column_dimensions['D'].width = 12
    ws_idx.column_dimensions['I'].width = 24
    ws_idx.column_dimensions['J'].width = 30
    ws_idx.column_dimensions['K'].width = 22
    
    print(f"Saving Master Workbook: {OUTPUT_MASTER}...")
    wb_master.save(OUTPUT_MASTER)
    print("Master Workbook saved successfully!")
    
    # -------------------------------------------------------------------------
    # 2. Per-Chassis Workbooks (One file per Chassis, each with 1 sheet per slot)
    # -------------------------------------------------------------------------
    print("\nGenerating Per-Chassis Workbooks in Slot_by_Slot_Reports/...")
    all_chassis = sorted(list({ch for (ch, sl) in sorted_slots}), key=lambda x: int(x[1:]) if x[1:].isdigit() else 99)
    for ch in all_chassis:
        ch_file = os.path.join(OUTPUT_DIR_CHASSIS, f"Chassis_{ch}_Slots_IO_Report.xlsx")
        wb_ch = openpyxl.Workbook()
        wb_ch.remove(wb_ch.active)
        
        # Filter slots for this chassis
        ch_slots = [(c, s) for (c, s) in sorted_slots if c == ch]
        
        # Create Index for this chassis
        ws_ch_idx = wb_ch.create_sheet(title=f"{ch}_Slot_Summary")
        ws_ch_idx.views.sheetView[0].showGridLines = True
        
        meta = CHASSIS_METADATA.get(ch, {})
        ws_ch_idx.merge_cells("A1:I1")
        t_ch = ws_ch_idx["A1"]
        t_ch.value = f"KALASIN ENGINEERING — {ch.upper()} ({meta.get('name', ch).upper()})"
        t_ch.font = Font(name="Arial", size=13, bold=True, color=WHITE)
        t_ch.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        t_ch.alignment = Alignment(horizontal="center", vertical="center")
        ws_ch_idx.row_dimensions[1].height = 28
        
        ws_ch_idx.merge_cells("A2:I2")
        st_ch = ws_ch_idx["A2"]
        st_ch.value = f"Location: {meta.get('panel', '-')} | Net: {meta.get('network', '-')} | Total Active Slots: {len(ch_slots)}"
        st_ch.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        st_ch.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        st_ch.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        ws_ch_idx.row_dimensions[2].height = 22
        
        ch_headers = [
            "Slot No.", "Module Model", "I/O Type", "Capacity",
            "Active Points", "Spare Points", "Spare %", "Primary Destinations",
            "Sheet Link"
        ]
        ws_ch_idx.row_dimensions[4].height = 24
        for col_idx, h_text in enumerate(ch_headers, start=1):
            c = ws_ch_idx.cell(row=4, column=col_idx, value=h_text)
            c.font = Font(name="Arial", size=9, bold=True, color=WHITE)
            c.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = header_border()
            
        r_c = 5
        for (c_id, sl) in ch_slots:
            pts = slot_groups[(c_id, sl)]
            pts.sort(key=lambda x: x['point'])
            sheet_title = f"Slot_{sl:02d}"
            
            act_cnt = sum(1 for p in pts if not p['is_spare'])
            spr_cnt = len(pts) - act_cnt
            spr_pct = (spr_cnt / len(pts) * 100) if pts else 0
            card_model = re.sub(r'-\d+$', '', pts[0]['card']) if pts and pts[0]['card'] else 'I/O Module'
            iotype = pts[0]['iotype'] if pts else 'I/O'
            dests = sorted(list({p['destination'] for p in pts if p['destination']}))
            dest_str = ", ".join(dests) if dests else "-"
            
            row_vals = [
                f"Slot {sl:02d}",
                card_model,
                iotype,
                len(pts),
                act_cnt,
                spr_cnt,
                f"{spr_pct:.1f}%",
                dest_str,
                f"Go to Slot {sl:02d} ➔"
            ]
            
            is_even = (r_c % 2 == 0)
            row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
            for col_idx, val in enumerate(row_vals, start=1):
                cell = ws_ch_idx.cell(row=r_c, column=col_idx, value=val)
                cell.font = Font(name="Arial", size=9, bold=(col_idx in (1, 2, 9)))
                cell.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
                cell.border = thin_border()
                
                if col_idx in (1, 3, 4, 5, 6, 7):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif col_idx == 9:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.hyperlink = f"#'{sheet_title}'!A1"
                    cell.font = Font(name="Arial", size=9, bold=True, color="1E3A8A", underline="single")
                    cell.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")
            ws_ch_idx.row_dimensions[r_c].height = 20
            r_c += 1
            
            # Create slot sheet
            ws_slot_indiv = wb_ch.create_sheet(title=sheet_title)
            populate_slot_sheet(ws_slot_indiv, c_id, sl, pts, is_master=False)
            
        for col in ws_ch_idx.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                if cell.row in (1, 2):
                    continue
                val_str = str(cell.value or '')
                max_len = max(max_len, len(val_str))
            ws_ch_idx.column_dimensions[col_letter].width = max(max_len + 3, 11)
            
        ws_ch_idx.column_dimensions['A'].width = 12
        ws_ch_idx.column_dimensions['B'].width = 18
        ws_ch_idx.column_dimensions['H'].width = 24
        ws_ch_idx.column_dimensions['I'].width = 18
        
        wb_ch.save(ch_file)
        print(f"Saved: {ch_file}")
        
    print("\nAll Reports Generated Successfully!")

if __name__ == "__main__":
    build_reports()
