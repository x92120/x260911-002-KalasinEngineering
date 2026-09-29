#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
CHASSIS & SLOT I/O MASTER REPORT: SOURCE & DESTINATION SCHEDULE
================================================================================
Cross-references:
  - 03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-6.xlsx (Sheets: 'IO List', 'Slot_Description')
  - 02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC.xlsx (Sheets: DI/DO Chassis 1 & 2)
Outputs:
  - 03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.xlsx

Styling: Light theme, engineering executive standard, openpyxl professional format.
================================================================================
"""

import os
import sys
import re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Define Paths
BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_LIST_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-6.xlsx")
TAGS_PLC_FILE = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.xlsx")

# Theme Palette (Light Executive Styling)
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
        'role': '1756-L950TPSXT Controller + DLR Ring Supervisor + Local Fast I/O',
        'network': 'Stratix 5400 / EN4TR Ring Node 1 (192.168.1.10)',
        'slots_total': 13
    },
    'C2': {
        'name': 'I/O Expansion Rack 1 (Chassis 2)',
        'panel': 'Control Cabinet CA1 (Bay 2)',
        'role': 'Digital & Analog Process I/O Expansion',
        'network': 'EN4TR DLR Ring Node 2 (192.168.1.11)',
        'slots_total': 13
    },
    'C3': {
        'name': 'I/O Expansion Rack 2 (Chassis 3)',
        'panel': 'Control Cabinet CA1 (Bay 3)',
        'role': 'Evaporator, Concentrator & Dosing Skid I/O',
        'network': 'EN4TR DLR Ring Node 3 (192.168.1.12)',
        'slots_total': 13
    },
    'C4': {
        'name': 'I/O Expansion Rack 3 (Chassis 4)',
        'panel': 'Control Cabinet CA1 (Bay 4)',
        'role': 'Spray Dryer Tower & Exhaust Auxiliary I/O',
        'network': 'EN4TR DLR Ring Node 4 (192.168.1.13)',
        'slots_total': 13
    },
    'C5': {
        'name': 'Remote I/O Skid RIO-200 (Chassis 5)',
        'panel': 'Remote I/O Enclosure RIO-200 (Process Floor 1F)',
        'role': 'Remote Field I/O via Fiber Optic Trunk (DLR Ring Node 5)',
        'network': 'EN4TR Fiber DLR Ring Node 5 (192.168.1.14)',
        'slots_total': 13
    },
    'C6': {
        'name': 'MCC Auxiliary I/O (Chassis 6)',
        'panel': 'Motor Control Center Room (MCC Panel)',
        'role': 'Direct Hardwired Motor Starter Feedback & Permissives',
        'network': 'MCC Hardwired Trunk Line',
        'slots_total': 8
    },
    'C7': {
        'name': 'MCC Fieldbus & Network Interface (Chassis 7)',
        'panel': 'Motor Control Center Room (MCC Panel & Field Bus)',
        'role': 'Modbus/Fieldbus Gateways & MCC Feeder Interlocks',
        'network': 'Modbus TCP / EtherNet Trunk',
        'slots_total': 8
    }
}

def thin_border():
    s = Side(style='thin', color=BORDER_COLOR)
    return Border(left=s, right=s, top=s, bottom=s)

def header_border():
    s = Side(style='medium', color="94A3B8")
    return Border(left=s, right=s, top=s, bottom=s)

def load_eplan_tags(tags_file):
    """Load ePlan wire tags keyed by (Chassis, Slot, Channel)"""
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
                    'note': str(note).strip() if note else '',
                    'sheet': sname
                }
    return tags_map

def load_hardware_slot_definitions(io_file):
    """Load hardware slot definitions from Slot_Description sheet"""
    wb = openpyxl.load_workbook(io_file, data_only=True)
    hw_slots = {}
    if 'Slot_Description' in wb.sheetnames:
        ws = wb['Slot_Description']
        for r in range(18, ws.max_row + 1):
            c_code = ws.cell(row=r, column=1).value
            rack = ws.cell(row=r, column=2).value
            slot = ws.cell(row=r, column=3).value
            card_type = ws.cell(row=r, column=4).value
            card_model = ws.cell(row=r, column=5).value
            sig_type = ws.cell(row=r, column=6).value
            if rack is not None and slot is not None:
                r_str = str(rack).strip()
                s_int = int(slot) if str(slot).isdigit() else str(slot)
                hw_slots[(r_str, s_int)] = {
                    'card_type': str(card_type).strip() if card_type else '',
                    'card_model': str(card_model).strip() if card_model else '',
                    'sig_type': str(sig_type).strip() if sig_type else ''
                }
    return hw_slots

def load_io_list(io_file):
    """Load all operational rows from IO List sheet"""
    wb = openpyxl.load_workbook(io_file, data_only=True)
    ws = wb['IO List']
    
    rows = []
    for r in range(2, ws.max_row + 1):
        chassis = ws.cell(row=r, column=32).value
        slot = ws.cell(row=r, column=33).value
        point = ws.cell(row=r, column=34).value
        
        if chassis is None:
            continue
            
        c_str = str(chassis).strip()
        s_int = int(slot) if slot is not None and str(slot).isdigit() else (str(slot).strip() if slot else 0)
        pt_int = int(point) if point is not None and str(point).isdigit() else 0
        
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
        ref_des = ws.cell(row=r, column=1).value
        
        is_spare = False
        inst_str = str(inst_tag or '').strip().lower()
        tag_str = str(plc_tag or '').strip().lower()
        if 'spare' in inst_str or 'spare' in tag_str or not inst_tag or inst_str == 'none':
            is_spare = True
            
        rows.append({
            'source_row': r,
            'chassis': c_str,
            'slot': s_int,
            'point': pt_int,
            'channel': pt_int + 1,
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
            'ref_des': str(ref_des).strip() if ref_des else ''
        })
        
    return rows

def build_report():
    print("Loading source files...")
    eplan_tags = load_eplan_tags(TAGS_PLC_FILE)
    print(f"Loaded {len(eplan_tags)} ePlan tag mappings.")
    
    hw_slots = load_hardware_slot_definitions(IO_LIST_FILE)
    print(f"Loaded {len(hw_slots)} hardware slot definitions.")
    
    io_rows = load_io_list(IO_LIST_FILE)
    print(f"Loaded {len(io_rows)} active rows from IO List.")
    
    # Cross-reference ePlan wire tags into io_rows
    matched_eplan = 0
    for r in io_rows:
        key = (r['chassis'], r['slot'], r['channel'])
        if key in eplan_tags:
            r['eplan_plc_tag'] = eplan_tags[key]['plc_tag']
            r['eplan_term_tag'] = eplan_tags[key]['term_tag']
            r['eplan_note'] = eplan_tags[key]['note']
            matched_eplan += 1
        else:
            # Fallback or synthetic standard ePlan wire tagging
            r['eplan_plc_tag'] = f"{r['terminal']}:{r['channel']}/{r['chassis']}S{r['slot']}:{r['channel']}" if r['terminal'] else f"{r['chassis']}S{r['slot']}:{r['channel']}"
            r['eplan_term_tag'] = f"{r['chassis']}S{r['slot']}:{r['channel']}/{r['terminal']}:{r['channel']}" if r['terminal'] else f"{r['chassis']}S{r['slot']}:{r['channel']}"
            r['eplan_note'] = ''
            
    print(f"Directly matched {matched_eplan} ePlan tags to I/O points.")
    
    # Sort IO rows: Chassis (C1..C7), Slot, Point
    def sort_key(row):
        ch = row['chassis']
        ch_num = int(ch[1:]) if ch.startswith('C') and ch[1:].isdigit() else 99
        sl = row['slot'] if isinstance(row['slot'], int) else 99
        pt = row['point'] if isinstance(row['point'], int) else 0
        return (ch_num, sl, pt)
        
    io_rows.sort(key=sort_key)
    
    # Initialize Workbook
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)
    
    # =========================================================================
    # TAB 1: Chassis and Slot Summary & Hardware Architecture
    # =========================================================================
    print("Building Tab 1: Chassis_and_Slot_List...")
    ws1 = wb.create_sheet(title="Chassis_and_Slot_List")
    ws1.views.sheetView[0].showGridLines = True
    
    # Title Block
    ws1.merge_cells("A1:M1")
    t_cell = ws1["A1"]
    t_cell.value = "KALASIN ENGINEERING - xCIP AUTOMATION SYSTEM (Ref: x2608003)"
    t_cell.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    t_cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 32
    
    ws1.merge_cells("A2:M2")
    st_cell = ws1["A2"]
    st_cell.value = "MASTER CHASSIS & SLOT ALLOCATION SCHEDULE (CONTROLLOGIX 1756 & EXPANSION RACKS)"
    st_cell.font = Font(name="Arial", size=11, bold=True, color=WHITE)
    st_cell.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    st_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 24
    
    # KPI Summary Cards Block (Rows 4 - 6)
    total_pts = len(io_rows)
    active_pts = sum(1 for r in io_rows if not r['is_spare'])
    spare_pts = total_pts - active_pts
    spare_pct = (spare_pts / total_pts * 100) if total_pts > 0 else 0
    
    kpi_defs = [
        ("Total Chassis Racks", "7 Chassis (C1-C7)", "A", "B", "1E3A8A", "DBEAFE"),
        ("Total I/O Capacity", f"{total_pts:,} Channels", "C", "D", "0F172A", "E2E8F0"),
        ("Active Field I/O", f"{active_pts:,} Points", "E", "F", "14532D", "DCFCE7"),
        ("Spare Engineering Capacity", f"{spare_pts:,} Points ({spare_pct:.1f}%)", "G", "H", "854D0E", "FEF9C3"),
        ("Controller CPU", "1756-L950TPSXT", "I", "J", "581C87", "F3E8FF"),
        ("DLR Network Standard", "EtherNet/IP Dual Ring", "K", "M", "065F46", "D1FAE5"),
    ]
    
    for title, val, c_start, c_end, f_col, b_col in kpi_defs:
        r4 = f"{c_start}4:{c_end}4"
        r5 = f"{c_start}5:{c_end}5"
        ws1.merge_cells(r4)
        ws1.merge_cells(r5)
        
        c4 = ws1[f"{c_start}4"]
        c4.value = title.upper()
        c4.font = Font(name="Arial", size=8, bold=True, color=f_col)
        c4.fill = PatternFill(start_color=b_col, end_color=b_col, fill_type="solid")
        c4.alignment = Alignment(horizontal="center", vertical="center")
        
        c5 = ws1[f"{c_start}5"]
        c5.value = val
        c5.font = Font(name="Arial", size=11, bold=True, color=f_col)
        c5.fill = PatternFill(start_color=b_col, end_color=b_col, fill_type="solid")
        c5.alignment = Alignment(horizontal="center", vertical="center")
        
        # Apply borders to merged KPI card
        col_s_idx = openpyxl.utils.column_index_from_string(c_start)
        col_e_idx = openpyxl.utils.column_index_from_string(c_end)
        for r_k in (4, 5):
            for c_k in range(col_s_idx, col_e_idx + 1):
                ws1.cell(row=r_k, column=c_k).border = thin_border()
                
    ws1.row_dimensions[4].height = 18
    ws1.row_dimensions[5].height = 24
    
    # Table Headers for Tab 1
    t1_headers = [
        "Chassis ID", "Chassis Description & Location", "Slot No.", "Module Model",
        "Module Description / Functional Role", "I/O Type", "Capacity (Pts)",
        "Active Points", "Spare Points", "Spare %", "Primary Field Destinations",
        "Network Node / Protocol", "Engineering Status"
    ]
    
    header_row = 7
    ws1.row_dimensions[header_row].height = 28
    for col_idx, h_text in enumerate(t1_headers, start=1):
        cell = ws1.cell(row=header_row, column=col_idx, value=h_text)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border()
        
    # Group points by Chassis and Slot
    slot_stats = defaultdict(lambda: {
        'total': 0, 'active': 0, 'spare': 0, 'types': Counter(),
        'cards': set(), 'dests': set()
    })
    
    for r in io_rows:
        key = (r['chassis'], r['slot'])
        slot_stats[key]['total'] += 1
        if r['is_spare']:
            slot_stats[key]['spare'] += 1
        else:
            slot_stats[key]['active'] += 1
        if r['iotype']:
            slot_stats[key]['types'][r['iotype']] += 1
        if r['card']:
            # Strip point suffix if present e.g. 1756-IB32-0 -> 1756-IB32
            base_card = re.sub(r'-\d+$', '', r['card'])
            slot_stats[key]['cards'].add(base_card)
        if r['destination']:
            slot_stats[key]['dests'].add(r['destination'])
            
    # Iterate through each chassis and each slot (0 to 12)
    cur_row = 8
    all_chassis_keys = ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7']
    
    for ch in all_chassis_keys:
        meta = CHASSIS_METADATA.get(ch, {})
        num_slots = meta.get('slots_total', 13)
        ch_name = meta.get('name', f'Chassis {ch}')
        ch_panel = meta.get('panel', 'Field / Control Panel')
        ch_net = meta.get('network', 'EtherNet/IP')
        
        # Chassis Group Header
        ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=13)
        gh_cell = ws1.cell(row=cur_row, column=1)
        gh_cell.value = f"▶ {ch} — {ch_name.upper()} | Location: {ch_panel} | Net: {ch_net}"
        gh_cell.font = Font(name="Arial", size=10, bold=True, color=SRC_HDR_FONT)
        gh_cell.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
        gh_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
        for c_k in range(1, 14):
            ws1.cell(row=cur_row, column=c_k).border = thin_border()
        ws1.row_dimensions[cur_row].height = 24
        cur_row += 1
        
        # Slots
        for sl in range(num_slots):
            st = slot_stats.get((ch, sl))
            hw_def = hw_slots.get((ch, sl), {})
            
            # Module resolution
            card_model = ""
            card_desc = ""
            iotype = ""
            cap = 0
            
            if hw_def:
                card_model = hw_def.get('card_model', '')
                card_type = hw_def.get('card_type', '')
                sig_type = hw_def.get('sig_type', '')
                
                if card_model == '1756-L950TPSXT':
                    card_desc = "ControlLogix 5590 Redundant Process Controller"
                    iotype = "CPU"
                    cap = 0
                elif card_model == '1756-EN4TR':
                    card_desc = "Dual-Port Gigabit EtherNet/IP DLR Comm Adapter"
                    iotype = "COMM"
                    cap = 0
                elif card_model == '1756-IB32':
                    card_desc = "32-Channel 24VDC Sink/Source Digital Input Module"
                    iotype = "DI"
                    cap = 32
                elif card_model == '1756-OB32':
                    card_desc = "32-Channel 24VDC Sourcing Digital Output Module"
                    iotype = "DO"
                    cap = 32
                elif card_model == '1756-IF16':
                    card_desc = "16-Channel High-Speed Isolated Analog Input Module"
                    iotype = "AI"
                    cap = 16
                elif card_model == '1756-OF8':
                    card_desc = "8-Channel Isolated Analog Current/Voltage Output Module"
                    iotype = "AO"
                    cap = 8
                elif card_model == '1756-N2':
                    card_desc = "Slot Filler / Future Expansion Space"
                    iotype = "EMPTY"
                    cap = 0
            
            # If st exists with actual IO data, override/enrich
            if st and st['total'] > 0:
                if not card_model and st['cards']:
                    card_model = ", ".join(sorted(st['cards']))
                if not iotype and st['types']:
                    iotype = "/".join(sorted(st['types'].keys()))
                cap = max(cap, st['total'])
                act_cnt = st['active']
                spr_cnt = st['spare']
                dests_str = ", ".join(sorted(st['dests'])) if st['dests'] else "-"
            else:
                act_cnt = 0
                spr_cnt = cap
                dests_str = "None (CPU/Comm/Spare)" if cap == 0 else "Spare Channel Bank"
                
            spare_p = (spr_cnt / cap * 100) if cap > 0 else 0
            
            # Status resolution
            if iotype in ("CPU", "COMM"):
                eng_status = "OPERATIONAL"
            elif iotype == "EMPTY":
                eng_status = "EMPTY SLOT"
            elif act_cnt > 0:
                eng_status = "ASSIGNED"
            else:
                eng_status = "RESERVED"
                
            row_data = [
                ch,
                ch_name,
                sl,
                card_model or "-",
                card_desc or "-",
                iotype or "-",
                cap if cap > 0 else "-",
                act_cnt if cap > 0 else "-",
                spr_cnt if cap > 0 else "-",
                f"{spare_p:.1f}%" if cap > 0 else "-",
                dests_str,
                ch_net,
                eng_status
            ]
            
            is_even = (sl % 2 == 0)
            row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
            
            for col_idx, val in enumerate(row_data, start=1):
                c = ws1.cell(row=cur_row, column=col_idx, value=val)
                c.font = Font(name="Arial", size=9, bold=(col_idx in (1, 3, 4, 13)))
                c.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
                c.border = thin_border()
                
                # Alignments
                if col_idx in (1, 3, 6, 7, 8, 9, 10, 13):
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="left", vertical="center")
                    
                # Badges for status
                if col_idx == 13:
                    if eng_status == "ASSIGNED":
                        c.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                        c.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
                    elif eng_status in ("RESERVED", "EMPTY SLOT"):
                        c.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                        c.font = Font(name="Arial", size=9, bold=True, color=SPARE_FONT)
                    elif eng_status == "OPERATIONAL":
                        c.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
                        c.font = Font(name="Arial", size=9, bold=True, color=SRC_HDR_FONT)
                        
            ws1.row_dimensions[cur_row].height = 20
            cur_row += 1
            
        cur_row += 1 # Spacing between chassis groups
        
    # Auto-adjust column widths
    ws1.freeze_panes = "D8"
    for col in ws1.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in (1, 2, 4, 5):
                continue
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws1.column_dimensions[col_letter].width = max(max_len + 4, 11)
        
    ws1.column_dimensions['A'].width = 12
    ws1.column_dimensions['B'].width = 32
    ws1.column_dimensions['C'].width = 10
    ws1.column_dimensions['D'].width = 18
    ws1.column_dimensions['E'].width = 38
    ws1.column_dimensions['K'].width = 30
    ws1.column_dimensions['L'].width = 28
    
    # =========================================================================
    # TAB 2: IO_Master_Source_Destination (1,210 Points Full Schedule)
    # =========================================================================
    print("Building Tab 2: IO_Master_Source_Destination...")
    ws2 = wb.create_sheet(title="IO_Master_Source_Destination")
    ws2.views.sheetView[0].showGridLines = True
    
    # Header Banner
    ws2.merge_cells("A1:U1")
    t2 = ws2["A1"]
    t2.value = "KALASIN ENGINEERING - xCIP TURNKEY AUTOMATION (Ref: x2608003)"
    t2.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    t2.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 30
    
    ws2.merge_cells("A2:U2")
    st2 = ws2["A2"]
    st2.value = "COMPLETE I/O SOURCE AND DESTINATION WIRING SCHEDULE (TOTAL 1,210 CHANNELS)"
    st2.font = Font(name="Arial", size=11, bold=True, color=WHITE)
    st2.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    st2.alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[2].height = 22
    
    # Level 1 Group Headers (Row 4)
    # Col A-I: SOURCE (PLC Cabinet)
    # Col J-L: SIGNAL ATTRIBUTES
    # Col M-U: DESTINATION (Field Enclosure & Instrument)
    ws2.merge_cells("A4:I4")
    g_src = ws2["A4"]
    g_src.value = "SOURCE SPECIFICATION (PLC CABINET / RACK / MODULE TERMINAL / ePLAN WIRE TAG)"
    g_src.font = Font(name="Arial", size=10, bold=True, color=SRC_HDR_FONT)
    g_src.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
    g_src.alignment = Alignment(horizontal="center", vertical="center")
    
    ws2.merge_cells("J4:L4")
    g_sys = ws2["J4"]
    g_sys.value = "SIGNAL CLASSIFICATION"
    g_sys.font = Font(name="Arial", size=10, bold=True, color=SYS_HDR_FONT)
    g_sys.fill = PatternFill(start_color=SYS_HDR_FILL, end_color=SYS_HDR_FILL, fill_type="solid")
    g_sys.alignment = Alignment(horizontal="center", vertical="center")
    
    ws2.merge_cells("M4:U4")
    g_dst = ws2["M4"]
    g_dst.value = "DESTINATION SPECIFICATION (FIELD JB / MCC / TERMINAL PIN / FIELD INSTRUMENT TAG)"
    g_dst.font = Font(name="Arial", size=10, bold=True, color=DEST_HDR_FONT)
    g_dst.fill = PatternFill(start_color=DEST_HDR_FILL, end_color=DEST_HDR_FILL, fill_type="solid")
    g_dst.alignment = Alignment(horizontal="center", vertical="center")
    
    for c_idx in range(1, 22):
        ws2.cell(row=4, column=c_idx).border = header_border()
    ws2.row_dimensions[4].height = 24
    
    # Level 2 Detailed Column Headers (Row 5)
    t2_headers = [
        # SOURCE (Cols 1-9)
        "Point Ref ID", "Chassis", "Slot", "Point", "Chan",
        "Module Model", "PLC Tag Name", "PLC Card Pin", "ePlan Wire Tag",
        # SIGNAL (Cols 10-12)
        "I/O Type", "Signal Spec", "Point Status",
        # DESTINATION (Cols 13-21)
        "Dest Enclosure", "Location / Area", "Floor / Zone", "Field Terminal Block",
        "Field Terminal Pin", "Instrument Tag", "Instrument Description",
        "P&ID Drawing No.", "Control / Interlock Logic"
    ]
    
    ws2.row_dimensions[5].height = 26
    for col_idx, h_text in enumerate(t2_headers, start=1):
        cell = ws2.cell(row=5, column=col_idx, value=h_text)
        cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = header_border()
        
    # Populate Tab 2 Data
    cur_row = 6
    for idx, r in enumerate(io_rows, start=1):
        pt_ref = f"{r['chassis']}-S{r['slot']:02d}-P{r['point']:02d}"
        
        # Floor / Zone formatting
        flr_zn = f"{r['floor']} / {r['zone']}" if r['floor'] and r['zone'] else (r['floor'] or r['zone'] or '-')
        
        row_vals = [
            pt_ref,
            r['chassis'],
            r['slot'],
            r['point'],
            r['channel'],
            re.sub(r'-\d+$', '', r['card']),
            r['plc_tag'],
            r['term2'] or f"Pin {r['channel']}",
            r['eplan_term_tag'] or r['eplan_plc_tag'],
            r['iotype'],
            r['signal_type'] or "24VDC Dry Contact" if r['iotype'] == 'DI' else ("24VDC Sourcing" if r['iotype'] == 'DO' else "4-20mA HART"),
            r['status'],
            r['destination'] or "-",
            r['location'] or "-",
            flr_zn,
            r['terminal'] or "-",
            r['terminal_label'] or "-",
            r['instrument_tag'],
            r['description'],
            r['pid'] or "-",
            r['control_desc'] or "-"
        ]
        
        is_even = (idx % 2 == 0)
        row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
        
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws2.cell(row=cur_row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=9, bold=(col_idx in (1, 7, 12, 18)))
            c.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
            c.border = thin_border()
            
            # Alignments
            if col_idx in (1, 2, 3, 4, 5, 6, 8, 10, 11, 12, 13, 15, 17, 20):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
            # Status badge
            if col_idx == 12:
                if r['status'] == "ACTIVE":
                    c.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                    c.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
                else:
                    c.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                    c.font = Font(name="Arial", size=9, bold=True, color=SPARE_FONT)
                    
        ws2.row_dimensions[cur_row].height = 19
        cur_row += 1
        
    ws2.freeze_panes = "F6"
    for col in ws2.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = 0
        for cell in col:
            if cell.row in (1, 2, 3, 4):
                continue
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 10)
        
    ws2.column_dimensions['A'].width = 16
    ws2.column_dimensions['G'].width = 24
    ws2.column_dimensions['I'].width = 22
    ws2.column_dimensions['M'].width = 18
    ws2.column_dimensions['R'].width = 20
    ws2.column_dimensions['S'].width = 36
    ws2.column_dimensions['U'].width = 30
    
    # =========================================================================
    # TABS 3+: Dedicated Sheets for Each Chassis (C1 through C7)
    # =========================================================================
    for ch in all_chassis_keys:
        ch_rows = [r for r in io_rows if r['chassis'] == ch]
        meta = CHASSIS_METADATA.get(ch, {})
        ch_name = meta.get('name', f"Chassis {ch}")
        ch_panel = meta.get('panel', 'Main Control Panel')
        ch_role = meta.get('role', 'Control & I/O')
        ch_net = meta.get('network', 'DLR Ring')
        
        sheet_title = f"Chassis_{ch}"
        print(f"Building Sheet: {sheet_title} ({len(ch_rows)} points)...")
        ws = wb.create_sheet(title=sheet_title)
        ws.views.sheetView[0].showGridLines = True
        
        # Header Banner
        ws.merge_cells("A1:S1")
        tb = ws["A1"]
        tb.value = f"KALASIN ENGINEERING — {ch.upper()} ({ch_name.upper()})"
        tb.font = Font(name="Arial", size=13, bold=True, color=WHITE)
        tb.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        tb.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28
        
        # Chassis Info Subheader
        ws.merge_cells("A2:S2")
        tb2 = ws["A2"]
        act_c = sum(1 for r in ch_rows if not r['is_spare'])
        spr_c = len(ch_rows) - act_c
        tb2.value = f"Location: {ch_panel} | Role: {ch_role} | Net: {ch_net} | Total Points: {len(ch_rows)} (Active: {act_c}, Spare: {spr_c})"
        tb2.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        tb2.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        tb2.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 22
        
        # Dual Group Header (Row 4)
        ws.merge_cells("A4:H4")
        g1 = ws["A4"]
        g1.value = "SOURCE (PLC RACK / SLOT / TERMINAL / ePLAN TAG)"
        g1.font = Font(name="Arial", size=10, bold=True, color=SRC_HDR_FONT)
        g1.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
        g1.alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells("I4:K4")
        g2 = ws["I4"]
        g2.value = "SIGNAL SPECIFICATION"
        g2.font = Font(name="Arial", size=10, bold=True, color=SYS_HDR_FONT)
        g2.fill = PatternFill(start_color=SYS_HDR_FILL, end_color=SYS_HDR_FILL, fill_type="solid")
        g2.alignment = Alignment(horizontal="center", vertical="center")
        
        ws.merge_cells("L4:S4")
        g3 = ws["L4"]
        g3.value = "DESTINATION (FIELD JB / MCC / TERMINAL PIN / FIELD DEVICE)"
        g3.font = Font(name="Arial", size=10, bold=True, color=DEST_HDR_FONT)
        g3.fill = PatternFill(start_color=DEST_HDR_FILL, end_color=DEST_HDR_FILL, fill_type="solid")
        g3.alignment = Alignment(horizontal="center", vertical="center")
        
        for c_k in range(1, 20):
            ws.cell(row=4, column=c_k).border = header_border()
        ws.row_dimensions[4].height = 22
        
        # Detailed Columns (Row 5)
        ch_headers = [
            "Point ID", "Slot", "Pt", "Ch", "Module Model", "PLC Tag Name",
            "PLC Card Pin", "ePlan Wire Tag",
            "I/O Type", "Signal Type", "Status",
            "Destination", "Field Location", "Floor/Zone", "Terminal Block",
            "Term Pin", "Instrument Tag", "Instrument Description", "P&ID Drawing"
        ]
        
        ws.row_dimensions[5].height = 24
        for col_idx, h_text in enumerate(ch_headers, start=1):
            cell = ws.cell(row=5, column=col_idx, value=h_text)
            cell.font = Font(name="Arial", size=9, bold=True, color=WHITE)
            cell.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = header_border()
            
        # Group by slot
        slots_in_ch = sorted(list({r['slot'] for r in ch_rows}), key=lambda x: int(x) if str(x).isdigit() else 99)
        c_row = 6
        
        for sl in slots_in_ch:
            slot_pts = [r for r in ch_rows if r['slot'] == sl]
            slot_act = sum(1 for r in slot_pts if not r['is_spare'])
            slot_spr = len(slot_pts) - slot_act
            slot_card = slot_pts[0]['card'] if slot_pts else ''
            slot_iotype = slot_pts[0]['iotype'] if slot_pts else ''
            
            # Slot Subheader
            ws.merge_cells(start_row=c_row, start_column=1, end_row=c_row, end_column=19)
            s_cell = ws.cell(row=c_row, column=1)
            s_cell.value = f"Slot {sl:02d} — Module: {re.sub(r'-\d+$', '', slot_card)} ({slot_iotype}) | Total: {len(slot_pts)} Channels (Active: {slot_act}, Spare: {slot_spr})"
            s_cell.font = Font(name="Arial", size=9, bold=True, color="0F172A")
            s_cell.fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
            s_cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
            for c_k in range(1, 20):
                ws.cell(row=c_row, column=c_k).border = thin_border()
            ws.row_dimensions[c_row].height = 20
            c_row += 1
            
            for pt_idx, r in enumerate(slot_pts, start=1):
                pt_ref = f"{r['chassis']}-S{r['slot']:02d}-P{r['point']:02d}"
                flr_zn = f"{r['floor']} / {r['zone']}" if r['floor'] and r['zone'] else (r['floor'] or r['zone'] or '-')
                
                vals = [
                    pt_ref,
                    r['slot'],
                    r['point'],
                    r['channel'],
                    re.sub(r'-\d+$', '', r['card']),
                    r['plc_tag'],
                    r['term2'] or f"Pin {r['channel']}",
                    r['eplan_term_tag'] or r['eplan_plc_tag'],
                    r['iotype'],
                    r['signal_type'] or "24VDC",
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
                
                is_even = (pt_idx % 2 == 0)
                row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
                
                for col_idx, val in enumerate(vals, start=1):
                    c = ws.cell(row=c_row, column=col_idx, value=val)
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
                            
                ws.row_dimensions[c_row].height = 19
                c_row += 1
                
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
        
    print(f"Saving report to: {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)
    print("Report successfully generated!")

if __name__ == "__main__":
    build_report()
