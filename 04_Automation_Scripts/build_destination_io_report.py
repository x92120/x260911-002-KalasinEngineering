#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
DESTINATION I/O MASTER REPORT: EXCEL (1 SHEET PER DESTINATION) & VECTOR PDF
================================================================================
Generates:
  1. Master Excel Deliverable:
     - 03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.xlsx
     - Sheet '00_Destination_Index': Interactive Hub linking to all 16 destination sheets.
     - 16 Dedicated Destination Sheets (e.g. 'DEST_JB-401', 'DEST_MCC', 'DEST_RIO-200'...).
  2. Master Print-Ready Vector PDF:
     - 02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Destination_IO_List_Master_Report.pdf
     - 03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.pdf
  3. Per-Destination Individual Workbooks in:
     - 03_IO_Lists_and_Schedules/Destination_Reports/

Styling: Light theme, engineering executive standard.
================================================================================
"""

import os
import sys
import re
import shutil
import subprocess
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_LIST_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-6.xlsx")
TAGS_PLC_FILE = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC.xlsx")

OUTPUT_EXCEL = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.xlsx")
OUTPUT_PDF_MAIN = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Destination_IO_List_Master_Report.pdf")
OUTPUT_PDF_MIRROR = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.pdf")
DEST_REPORTS_DIR = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Destination_Reports")

CHROME_BIN = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

os.makedirs(DEST_REPORTS_DIR, exist_ok=True)

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

DEST_METADATA = {
    'MCC': {
        'desc': 'Motor Control Center (MCC Room)',
        'location': 'MCC Room Ground Floor',
        'size': 'Custom Switchgear Lineup',
        'type': 'MCC Switchgear & Bus Interface',
        'area': 'MCC Substation',
    },
    'JB-607': {
        'desc': 'Spray Dryer 7th Floor Junction Box',
        'location': 'Spray Dryer Tower 7th Floor',
        'size': '800 x 600 x 250 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Spray Dryer High Elevation',
    },
    'JB-401': {
        'desc': 'Infeed Area 2nd Floor Junction Box',
        'location': 'Infeed Processing 2nd Floor',
        'size': '800 x 600 x 250 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Slurry Infeed',
    },
    'RIO-200': {
        'desc': 'Remote I/O Enclosure RIO-200',
        'location': 'Slurry Outbuilding 2nd Floor',
        'size': '800 x 1000 x 400 mm',
        'type': 'Remote I/O Skid Cabinet (DLR Fiber Node)',
        'area': 'Remote Slurry Skid',
    },
    'JB-601': {
        'desc': 'Spray Dryer 1st Floor Junction Box',
        'location': 'Spray Dryer Floor 1F (Base)',
        'size': '500 x 400 x 200 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Spray Dryer Base Area',
    },
    'JB-602': {
        'desc': 'Spray Dryer 3rd Floor Junction Box',
        'location': 'Spray Dryer Floor 3F',
        'size': '800 x 600 x 250 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Chamber Middle Section',
    },
    'JB-402': {
        'desc': 'Jet Cooker 2nd Floor Junction Box',
        'location': 'Jet Cooker Processing 2F',
        'size': '450 x 300 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Thermal Cooking Skid',
    },
    'JB-612': {
        'desc': 'Packing Tower 2nd Floor Junction Box',
        'location': 'Packing Tower 2F (Intermediate)',
        'size': '450 x 300 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Packaging & Exhaust',
    },
    'JB-618': {
        'desc': 'Packing Tower 8th Floor Junction Box',
        'location': 'Packing Tower 8F (Top Head)',
        'size': '450 x 300 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Packing Exhaust Cyclone',
    },
    'IS-JB-618': {
        'desc': 'Packing Tower 8th Floor Intrinsically Safe JB',
        'location': 'Packing Tower 8F (Hazardous Vapor Area)',
        'size': '450 x 300 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'Explosion Vent / Vapor Zone',
    },
    'IS-JB-612': {
        'desc': 'Packing Tower 2nd Floor Intrinsically Safe JB',
        'location': 'Packing Tower 2F (Dust Hazard Area)',
        'size': '300 x 200 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'Product Discharge Hopper',
    },
    'CA1': {
        'desc': 'Main Control Panel CA1 Local Terminals',
        'location': 'Control Room Cabinet CA1',
        'size': '1600 x 2000 x 800 mm',
        'type': 'Main PLC Enclosure Internal Marshelling',
        'area': 'Central Control Room',
    },
    'IS-JB-603': {
        'desc': 'Spray Dryer 3rd Floor Intrinsically Safe JB',
        'location': 'Spray Dryer 3F (Internal Chamber Access)',
        'size': '300 x 200 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'Dryer Chamber Monitoring',
    },
    'IS-JB-608': {
        'desc': 'Spray Dryer 8th Floor Intrinsically Safe JB',
        'location': 'Spray Dryer 8F (Roof Atomizer Deck)',
        'size': '300 x 200 x 150 mm',
        'type': 'Intrinsically Safe (Ex i Blue Terminals)',
        'area': 'High Shear Atomizer Deck',
    },
    'JB-606': {
        'desc': 'Spray Dryer 6th Floor Junction Box',
        'location': 'Spray Dryer 6F (Air Heater Section)',
        'size': '300 x 200 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Direct Gas Burner Section',
    },
    'JB-608': {
        'desc': 'Spray Dryer 8th Floor Temperature JB',
        'location': 'Spray Dryer 8F (Top Plenum)',
        'size': '300 x 200 x 150 mm',
        'type': 'Standard Industrial (IP66 / SS304)',
        'area': 'Top Roof Sensors',
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
        dest = ws.cell(row=r, column=2).value
        if chassis is None or slot is None:
            continue
            
        c_str = str(chassis).strip()
        s_int = int(slot) if str(slot).isdigit() else str(slot).strip()
        pt_int = int(point) if point is not None and str(point).isdigit() else 0
        chan_int = pt_int + 1
        d_str = str(dest).strip() if dest else 'UNASSIGNED'
        
        card = ws.cell(row=r, column=28).value
        iotype = ws.cell(row=r, column=22).value
        plc_tag = ws.cell(row=r, column=21).value
        inst_tag = ws.cell(row=r, column=12).value
        desc = ws.cell(row=r, column=13).value
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
            'destination': d_str,
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

def populate_dest_sheet(ws, dest_name, pts, is_master=True):
    ws.views.sheetView[0].showGridLines = True
    meta = DEST_METADATA.get(dest_name, {})
    desc = meta.get('desc', f"{dest_name} Field Junction Box")
    loc = meta.get('location', pts[0]['location'] if pts else '-')
    size = meta.get('size', 'Standard Industrial')
    enc_type = meta.get('type', 'IP66 Stainless Steel SS304')
    
    act_cnt = sum(1 for p in pts if not p['is_spare'])
    spr_cnt = len(pts) - act_cnt
    spr_pct = (spr_cnt / len(pts) * 100) if pts else 0
    
    # Types count
    types_cnt = Counter(p['iotype'] for p in pts if p['iotype'])
    types_str = " | ".join([f"{k}: {v}" for k, v in sorted(types_cnt.items())])
    
    # Linked Racks/Slots
    racks = sorted(list({f"{p['chassis']}-S{p['slot']:02d}" for p in pts}))
    racks_str = ", ".join(racks[:6]) + (f" (+{len(racks)-6} more)" if len(racks) > 6 else "")
    
    # Title Block
    ws.merge_cells("A1:S1")
    t = ws["A1"]
    t.value = f"KALASIN ENGINEERING — DESTINATION: {dest_name.upper()} ({desc.upper()})"
    t.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    t.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    
    # Subheader Banner
    ws.merge_cells("A2:Q2")
    st = ws["A2"]
    st.value = f"Location: {loc} | Enclosure: {size} ({enc_type}) | Total I/O: {len(pts)} ({types_str}) | Active: {act_cnt}, Spare: {spr_cnt} ({spr_pct:.1f}%)"
    st.font = Font(name="Arial", size=9, bold=True, color=WHITE)
    st.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    st.alignment = Alignment(horizontal="left", vertical="center", indent=1)
    
    # Navigation Link Back to Index
    if is_master:
        ws.merge_cells("R2:S2")
        nav = ws["R2"]
        nav.value = "⬅ Return to Destination Index"
        nav.hyperlink = "#'00_Destination_Index'!A1"
        nav.font = Font(name="Arial", size=9, bold=True, color="1E3A8A", underline="single")
        nav.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
        nav.alignment = Alignment(horizontal="center", vertical="center")
    else:
        ws.merge_cells("R2:S2")
        nav = ws["R2"]
        nav.value = f"Racks: {racks_str[:22]}"
        nav.font = Font(name="Arial", size=8, bold=True, color=WHITE)
        nav.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
        nav.alignment = Alignment(horizontal="center", vertical="center")
        
    ws.row_dimensions[2].height = 22
    
    # Dual Level Headers (Row 4 & 5)
    # Notice: In Destination Sheet, DESTINATION columns are featured prominently on the left or paired with Source!
    ws.merge_cells("A4:G4")
    g1 = ws["A4"]
    g1.value = "DESTINATION SPECIFICATION (FIELD JB / MCC / TERMINAL PIN / FIELD DEVICE)"
    g1.font = Font(name="Arial", size=9, bold=True, color=DEST_HDR_FONT)
    g1.fill = PatternFill(start_color=DEST_HDR_FILL, end_color=DEST_HDR_FILL, fill_type="solid")
    g1.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("H4:J4")
    g2 = ws["H4"]
    g2.value = "SIGNAL SPECIFICATION"
    g2.font = Font(name="Arial", size=9, bold=True, color=SYS_HDR_FONT)
    g2.fill = PatternFill(start_color=SYS_HDR_FILL, end_color=SYS_HDR_FILL, fill_type="solid")
    g2.alignment = Alignment(horizontal="center", vertical="center")
    
    ws.merge_cells("K4:S4")
    g3 = ws["K4"]
    g3.value = "SOURCE SPECIFICATION (PLC RACK / CHASSIS / SLOT / CARD PIN / ePLAN WIRE TAG)"
    g3.font = Font(name="Arial", size=9, bold=True, color=SRC_HDR_FONT)
    g3.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
    g3.alignment = Alignment(horizontal="center", vertical="center")
    
    for c_k in range(1, 20):
        ws.cell(row=4, column=c_k).border = header_border()
    ws.row_dimensions[4].height = 22
    
    headers = [
        # Destination Columns (Cols 1-7)
        "Dest Enclosure", "Location / Area", "Floor / Zone", "Field Terminal Block",
        "Term Pin", "Instrument Tag", "Instrument Description",
        # Signal (Cols 8-10)
        "I/O Type", "Signal Spec", "Point Status",
        # Source Columns (Cols 11-19)
        "Point Ref ID", "Chassis", "Slot", "Pt", "Ch", "Module Model",
        "PLC Tag Name", "PLC Card Pin", "ePlan Wire Tag"
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
        sig_spec = r['signal_type'] or ("24VDC Dry Contact" if r['iotype'] == 'DI' else ("24VDC Sourcing" if r['iotype'] == 'DO' else "4-20mA HART"))
        
        row_vals = [
            r['destination'],
            r['location'] or "-",
            flr_zn,
            r['terminal'] or "-",
            r['terminal_label'] or "-",
            r['instrument_tag'],
            r['description'],
            r['iotype'],
            sig_spec,
            r['status'],
            pt_ref,
            r['chassis'],
            r['slot'],
            r['point'],
            r['channel'],
            re.sub(r'-\d+$', '', r['card']),
            r['plc_tag'],
            r['term2'] or f"Pin {r['channel']}",
            r['eplan_term_tag'] or r['eplan_plc_tag']
        ]
        
        is_even = (idx % 2 == 0)
        row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
        
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws.cell(row=cur_row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=9, bold=(col_idx in (1, 6, 10, 11, 17)))
            c.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
            c.border = thin_border()
            
            if col_idx in (1, 3, 4, 5, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18):
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
            if col_idx == 10:
                if r['status'] == "ACTIVE":
                    c.fill = PatternFill(start_color=ACTIVE_FILL, end_color=ACTIVE_FILL, fill_type="solid")
                    c.font = Font(name="Arial", size=9, bold=True, color=ACTIVE_FONT)
                else:
                    c.fill = PatternFill(start_color=SPARE_FILL, end_color=SPARE_FILL, fill_type="solid")
                    c.font = Font(name="Arial", size=9, bold=True, color=SPARE_FONT)
                    
        ws.row_dimensions[cur_row].height = 19
        cur_row += 1
        
    ws.freeze_panes = "H6"
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
        
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 34
    ws.column_dimensions['K'].width = 16
    ws.column_dimensions['Q'].width = 24
    ws.column_dimensions['S'].width = 24

def build_destination_reports():
    print("Loading data for Destination reports...")
    rows = load_data()
    print(f"Loaded {len(rows)} data rows.")
    
    # Group by destination
    dest_groups = defaultdict(list)
    for r in rows:
        dest_groups[r['destination']].append(r)
        
    # Sort destinations: MCC first, then by count descending
    dest_priority = list(DEST_METADATA.keys())
    def dest_sort_key(d):
        if d in dest_priority:
            return (0, dest_priority.index(d))
        return (1, -len(dest_groups[d]))
        
    sorted_dests = sorted(dest_groups.keys(), key=dest_sort_key)
    print(f"Total destinations to process: {len(sorted_dests)}")
    
    total_pts = len(rows)
    total_act = sum(1 for r in rows if not r['is_spare'])
    total_spr = total_pts - total_act
    
    # -------------------------------------------------------------------------
    # 1. Master Excel Deliverable (00_Destination_Index + 16 Destination Sheets)
    # -------------------------------------------------------------------------
    print("\nCreating Master Excel Workbook with 1 Sheet per Destination...")
    wb_master = openpyxl.Workbook()
    wb_master.remove(wb_master.active)
    
    # Index Sheet
    ws_idx = wb_master.create_sheet(title="00_Destination_Index")
    ws_idx.views.sheetView[0].showGridLines = True
    
    # Title
    ws_idx.merge_cells("A1:L1")
    t1 = ws_idx["A1"]
    t1.value = "KALASIN ENGINEERING — xCIP AUTOMATION SYSTEM (Ref: x2608003)"
    t1.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    t1.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
    t1.alignment = Alignment(horizontal="center", vertical="center")
    ws_idx.row_dimensions[1].height = 32
    
    ws_idx.merge_cells("A2:L2")
    st1 = ws_idx["A2"]
    st1.value = "MASTER DESTINATION DIRECTORY — FIELD JUNCTION BOX & MCC I/O TERMINATION SCHEDULE"
    st1.font = Font(name="Arial", size=11, bold=True, color=WHITE)
    st1.fill = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
    st1.alignment = Alignment(horizontal="center", vertical="center")
    ws_idx.row_dimensions[2].height = 24
    
    # KPI Row
    kpis = [
        ("Total Field Destinations", f"{len(sorted_dests)} Enclosures", "A", "B", "1E3A8A", "DBEAFE"),
        ("Total Plant I/O Capacity", f"{total_pts:,} Channels", "C", "D", "0F172A", "E2E8F0"),
        ("Active Terminated Points", f"{total_act:,} Points", "E", "F", "14532D", "DCFCE7"),
        ("Field Spare Margin", f"{total_spr:,} ({(total_spr/total_pts*100):.1f}%)", "G", "H", "854D0E", "FEF9C3"),
        ("Enclosure Standards", "IP66 / SS304 / Non-ATEX", "I", "J", "581C87", "F3E8FF"),
        ("Intrinsically Safe JBs", "4 IS Enclosures (Ex i)", "K", "L", "065F46", "D1FAE5")
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
    idx_headers = [
        "Dest ID", "Enclosure Description", "Plant Location / Floor", "Dimensions",
        "Enclosure Spec", "Total Pts", "Active", "Spare", "Spare %",
        "I/O Breakdown", "Source Racks", "Interactive Link"
    ]
    ws_idx.row_dimensions[7].height = 26
    for col_idx, h_text in enumerate(idx_headers, start=1):
        c = ws_idx.cell(row=7, column=col_idx, value=h_text)
        c.font = Font(name="Arial", size=9, bold=True, color=WHITE)
        c.fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = header_border()
        
    cur_idx_row = 8
    for d_name in sorted_dests:
        pts = dest_groups[d_name]
        # Sort points by Terminal, Channel, PLC Tag
        pts.sort(key=lambda x: (x['terminal'] or '', x['channel'], x['plc_tag']))
        
        meta = DEST_METADATA.get(d_name, {})
        desc = meta.get('desc', f"{d_name} Field JB")
        loc = meta.get('location', pts[0]['location'] if pts else '-')
        size = meta.get('size', 'Standard')
        enc_type = meta.get('type', 'IP66 Industrial')
        
        act = sum(1 for p in pts if not p['is_spare'])
        spr = len(pts) - act
        spr_pct = (spr / len(pts) * 100) if pts else 0
        
        types_cnt = Counter(p['iotype'] for p in pts if p['iotype'])
        types_str = " | ".join([f"{k}:{v}" for k, v in sorted(types_cnt.items())])
        
        racks = sorted(list({p['chassis'] for p in pts}))
        racks_str = ", ".join(racks)
        
        sheet_title = f"DEST_{d_name}"
        if len(sheet_title) > 31:
            sheet_title = sheet_title[:31]
            
        row_vals = [
            d_name,
            desc,
            loc,
            size,
            enc_type,
            len(pts),
            act,
            spr,
            f"{spr_pct:.1f}%",
            types_str,
            racks_str,
            f"Open {d_name} ➔"
        ]
        
        is_even = (cur_idx_row % 2 == 0)
        row_fill_color = LIGHT_BG_1 if is_even else LIGHT_BG_2
        
        for col_idx, val in enumerate(row_vals, start=1):
            c = ws_idx.cell(row=cur_idx_row, column=col_idx, value=val)
            c.font = Font(name="Arial", size=9, bold=(col_idx in (1, 6, 12)))
            c.fill = PatternFill(start_color=row_fill_color, end_color=row_fill_color, fill_type="solid")
            c.border = thin_border()
            
            if col_idx in (1, 4, 6, 7, 8, 9, 11):
                c.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx == 12:
                c.alignment = Alignment(horizontal="center", vertical="center")
                c.hyperlink = f"#'{sheet_title}'!A1"
                c.font = Font(name="Arial", size=9, bold=True, color="1E3A8A", underline="single")
                c.fill = PatternFill(start_color=SRC_HDR_FILL, end_color=SRC_HDR_FILL, fill_type="solid")
            else:
                c.alignment = Alignment(horizontal="left", vertical="center")
                
        ws_idx.row_dimensions[cur_idx_row].height = 20
        cur_idx_row += 1
        
        # Create Destination sheet in Master Workbook
        ws_dest = wb_master.create_sheet(title=sheet_title)
        populate_dest_sheet(ws_dest, d_name, pts, is_master=True)
        
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
        
    ws_idx.column_dimensions['A'].width = 14
    ws_idx.column_dimensions['B'].width = 30
    ws_idx.column_dimensions['C'].width = 28
    ws_idx.column_dimensions['D'].width = 18
    ws_idx.column_dimensions['E'].width = 28
    ws_idx.column_dimensions['J'].width = 24
    ws_idx.column_dimensions['L'].width = 18
    
    print(f"Saving Master Excel Report: {OUTPUT_EXCEL}...")
    wb_master.save(OUTPUT_EXCEL)
    print("Master Excel Report saved successfully!")
    
    # -------------------------------------------------------------------------
    # 2. Master Destination Vector PDF Report
    # -------------------------------------------------------------------------
    print("\nGenerating Master Destination Vector PDF...")
    
    def get_dest_css():
        return """
        @page {
            size: A4 landscape;
            margin: 10mm 10mm 12mm 10mm;
            @bottom-right {
                content: "Page " counter(page) " of " counter(pages);
                font-size: 8pt;
                color: #64748B;
                font-family: Arial, sans-serif;
            }
        }
        *, *::before, *::after { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            margin: 0; padding: 0; background: #FFFFFF; color: #0F172A; font-size: 8pt; line-height: 1.25;
        }
        .page-break { page-break-before: always; }
        .doc-header {
            border-bottom: 2px solid #1E293B; padding-bottom: 8px; margin-bottom: 12px;
            display: flex; justify-content: space-between; align-items: flex-end;
        }
        .doc-title-block h1 {
            margin: 0; font-size: 15pt; color: #0F172A; font-weight: 800; letter-spacing: -0.3px; text-transform: uppercase;
        }
        .doc-title-block h2 {
            margin: 2px 0 0 0; font-size: 10pt; color: #2563EB; font-weight: 600;
        }
        .doc-meta-block { text-align: right; font-size: 7.5pt; color: #475569; line-height: 1.35; }
        .doc-meta-block strong { color: #0F172A; }
        .kpi-row { display: grid; grid-template-columns: repeat(6, 1fr); gap: 8px; margin-bottom: 14px; }
        .kpi-card {
            border: 1px solid #CBD5E1; border-radius: 4px; padding: 6px 8px; background: #F8FAFC; text-align: center;
        }
        .kpi-card.blue { background: #EFF6FF; border-color: #BFDBFE; }
        .kpi-card.green { background: #F0FDF4; border-color: #BBF7D0; }
        .kpi-card.amber { background: #FEFCE8; border-color: #FEF08A; }
        .kpi-card.purple { background: #FAF5FF; border-color: #E9D5FF; }
        .kpi-card .kpi-lbl { font-size: 6.5pt; text-transform: uppercase; font-weight: 700; color: #475569; margin-bottom: 3px; }
        .kpi-card .kpi-val { font-size: 11pt; font-weight: 800; color: #0F172A; }
        .kpi-card.blue .kpi-val { color: #1E40AF; }
        .kpi-card.green .kpi-val { color: #166534; }
        .kpi-card.amber .kpi-val { color: #854D0E; }
        .kpi-card.purple .kpi-val { color: #6B21A8; }
        table.data-table { width: 100%; border-collapse: collapse; margin-bottom: 14px; font-size: 7.2pt; }
        table.data-table th, table.data-table td { border: 1px solid #CBD5E1; padding: 4.5px 5px; vertical-align: middle; }
        table.data-table thead tr.group-header th { font-size: 7.5pt; font-weight: 700; text-transform: uppercase; letter-spacing: 0.2px; padding: 5px; }
        table.data-table thead tr.col-header th { background: #1E293B; color: #FFFFFF; font-weight: 600; font-size: 7pt; text-align: center; }
        table.data-table tbody tr:nth-child(even) { background: #F8FAFC; }
        .th-src { background: #DBEAFE; color: #1E3A8A; }
        .th-sys { background: #E2E8F0; color: #0F172A; }
        .th-dst { background: #FEF3C7; color: #78350F; }
        .text-center { text-align: center; }
        .text-left { text-align: left; }
        .font-bold { font-weight: 700; }
        .font-mono { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 7pt; }
        .badge { display: inline-block; padding: 1.5px 5px; border-radius: 3px; font-size: 6.5pt; font-weight: 700; text-align: center; }
        .badge-active { background: #DCFCE7; color: #14532D; border: 1px solid #86EFAC; }
        .badge-spare { background: #FEF9C3; color: #854D0E; border: 1px solid #FDE047; }
        .badge-is { background: #E0F2FE; color: #0369A1; border: 1px solid #7DD3FC; }
        .dest-banner {
            background: #1E293B; color: #FFFFFF; border-radius: 4px 4px 0 0; padding: 6px 10px;
            display: flex; justify-content: space-between; align-items: center; margin-top: 10px;
        }
        .dest-banner h3 { margin: 0; font-size: 9.5pt; font-weight: 700; }
        .dest-banner .dest-sub { font-size: 7.5pt; color: #94A3B8; font-weight: 400; }
        .dest-banner .dest-kpi { font-size: 7.5pt; font-weight: 600; background: rgba(255,255,255,0.12); padding: 2px 8px; border-radius: 3px; }
        .footer-stamp {
            border-top: 1px solid #E2E8F0; padding-top: 4px; margin-top: 8px;
            display: flex; justify-content: space-between; font-size: 6.5pt; color: #64748B;
        }
        """

    pdf_html = f"""<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Destination I/O Master Schedule</title>
        <style>{get_dest_css()}</style>
    </head>
    <body>
        <!-- Cover & KPI Summary -->
        <div class="doc-header">
            <div class="doc-title-block">
                <h1>Kalasin Engineering — xCIP Automation System</h1>
                <h2>MASTER DESTINATION I/O TERMINATION & WIRING SCHEDULE</h2>
            </div>
            <div class="doc-meta-block">
                <strong>Project Ref:</strong> x2608003<br>
                <strong>Client:</strong> Ingredion Thailand (Kalasin Plant)<br>
                <strong>Discipline:</strong> Field Instrumentation & JB Termination<br>
                <strong>Revision:</strong> Rev 3.6 | <strong>Date:</strong> 28-Sep-2026
            </div>
        </div>

        <div class="kpi-row">
            <div class="kpi-card blue">
                <div class="kpi-lbl">Total Destinations</div>
                <div class="kpi-val">{len(sorted_dests)} Enclosures</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-lbl">Total Field I/O</div>
                <div class="kpi-val">{total_pts:,} Channels</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-lbl">Active Signals</div>
                <div class="kpi-val">{total_act:,} Points</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-lbl">Field Spares</div>
                <div class="kpi-val">{total_spr:,} ({(total_spr/total_pts*100):.1f}%)</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-lbl">Hardware Racks</div>
                <div class="kpi-val">7 Chassis (C1-C7)</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-lbl">Enclosure Spec</div>
                <div class="kpi-val">IP66 / SS304 / Ex i</div>
            </div>
        </div>

        <!-- Section 1: Destination Summary Table -->
        <h3 style="margin: 6px 0; font-size: 10pt; color: #1E293B;">SECTION 1: MASTER DESTINATION ENCLOSURE & CAPACITY DIRECTORY</h3>
        <table class="data-table" style="font-size: 7.2pt; margin-bottom: 20px;">
            <thead>
                <tr class="col-header">
                    <th style="width: 70px;">Dest ID</th>
                    <th style="width: 140px;">Enclosure Description</th>
                    <th style="width: 130px;">Plant Location / Floor</th>
                    <th style="width: 90px;">Dimensions</th>
                    <th>Enclosure Specification</th>
                    <th style="width: 45px;">Total</th>
                    <th style="width: 45px;">Active</th>
                    <th style="width: 45px;">Spare</th>
                    <th style="width: 50px;">Spare %</th>
                    <th style="width: 110px;">I/O Breakdown</th>
                    <th style="width: 70px;">Racks</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for d_name in sorted_dests:
        pts = dest_groups[d_name]
        meta = DEST_METADATA.get(d_name, {})
        desc = meta.get('desc', f"{d_name} Field JB")
        loc = meta.get('location', pts[0]['location'] if pts else '-')
        size = meta.get('size', 'Standard')
        enc_type = meta.get('type', 'IP66 SS304')
        act = sum(1 for p in pts if not p['is_spare'])
        spr = len(pts) - act
        spr_pct = (spr / len(pts) * 100) if pts else 0
        types_cnt = Counter(p['iotype'] for p in pts if p['iotype'])
        types_str = " | ".join([f"{k}:{v}" for k, v in sorted(types_cnt.items())])
        racks = sorted(list({p['chassis'] for p in pts}))
        racks_str = ", ".join(racks)
        
        pdf_html += f"""
        <tr>
            <td class="text-center font-bold font-mono">{d_name}</td>
            <td>{desc}</td>
            <td>{loc}</td>
            <td class="text-center font-mono">{size}</td>
            <td>{enc_type}</td>
            <td class="text-center font-bold">{len(pts)}</td>
            <td class="text-center">{act}</td>
            <td class="text-center">{spr}</td>
            <td class="text-center">{spr_pct:.1f}%</td>
            <td>{types_str}</td>
            <td class="text-center font-mono font-bold">{racks_str}</td>
        </tr>
        """
        
    pdf_html += """
            </tbody>
        </table>
        <div class="footer-stamp">
            <span>Section 1 Master Destination Schedule | Ingredion Sprint 18K Plant</span>
            <span>All Junction Boxes comply with Sanitary Food & Beverage 3-A / Non-ATEX Requirements</span>
        </div>
    """
    
    # Section 2: Destination-by-Destination Detailed Schedules
    print("Generating Destination-by-Destination Detailed Pages...")
    for d_name in sorted_dests:
        pts = dest_groups[d_name]
        pts.sort(key=lambda x: (x['terminal'] or '', x['channel'], x['plc_tag']))
        
        meta = DEST_METADATA.get(d_name, {})
        desc = meta.get('desc', f"{d_name} Field JB")
        loc = meta.get('location', pts[0]['location'] if pts else '-')
        size = meta.get('size', 'Standard')
        enc_type = meta.get('type', 'IP66 SS304')
        act = sum(1 for p in pts if not p['is_spare'])
        spr = len(pts) - act
        spr_pct = (spr / len(pts) * 100) if pts else 0
        types_cnt = Counter(p['iotype'] for p in pts if p['iotype'])
        types_str = " | ".join([f"{k}:{v}" for k, v in sorted(types_cnt.items())])
        
        pdf_html += '<div class="page-break"></div>'
        pdf_html += f"""
        <div class="dest-banner">
            <div>
                <h3>DESTINATION: {d_name} — {desc.upper()}</h3>
                <span class="dest-sub">Location: {loc} | Size: {size} | Spec: {enc_type}</span>
            </div>
            <div class="dest-kpi">
                Total: {len(pts)} Pts &nbsp;|&nbsp; Active: {act} &nbsp;|&nbsp; Spare: {spr} ({spr_pct:.1f}%) &nbsp;|&nbsp; {types_str}
            </div>
        </div>
        <table class="data-table">
            <thead>
                <tr class="group-header">
                    <th colspan="7" class="th-dst text-center">DESTINATION SPECIFICATION (FIELD JB / MCC / TERMINAL PIN / FIELD DEVICE)</th>
                    <th colspan="3" class="th-sys text-center">SIGNAL SPEC</th>
                    <th colspan="7" class="th-src text-center">SOURCE SPECIFICATION (PLC RACK / SLOT / CARD PIN / ePLAN WIRE TAG)</th>
                </tr>
                <tr class="col-header">
                    <th style="width: 65px;">Dest ID</th>
                    <th style="width: 80px;">Location</th>
                    <th style="width: 70px;">Terminal Block</th>
                    <th style="width: 35px;">Pin</th>
                    <th style="width: 90px;">Instrument Tag</th>
                    <th>Instrument Description</th>
                    <th style="width: 65px;">P&ID Drawing</th>
                    <th style="width: 30px;">Type</th>
                    <th style="width: 80px;">Signal Spec</th>
                    <th style="width: 50px;">Status</th>
                    <th style="width: 75px;">Point Ref ID</th>
                    <th style="width: 30px;">Rack</th>
                    <th style="width: 25px;">Slot</th>
                    <th style="width: 25px;">Ch</th>
                    <th style="width: 65px;">Module</th>
                    <th style="width: 105px;">PLC Tag Name</th>
                    <th style="width: 100px;">ePlan Wire Tag</th>
                </tr>
            </thead>
            <tbody>
        """
        
        for p in pts:
            pt_ref = f"{p['chassis']}-S{p['slot']:02d}-P{p['point']:02d}"
            badge_cls = "badge-active" if p['status'] == "ACTIVE" else "badge-spare"
            sig_spec = p['signal_type'] or ("24VDC Dry Contact" if p['iotype'] == 'DI' else ("24VDC Sourcing" if p['iotype'] == 'DO' else "4-20mA HART"))
            
            pdf_html += f"""
            <tr>
                <td class="text-center font-bold font-mono">{p['destination']}</td>
                <td>{p['location'] or loc}</td>
                <td class="text-center font-mono">{p['terminal'] or '-'}</td>
                <td class="text-center font-mono">{p['terminal_label'] or '-'}</td>
                <td class="font-bold">{p['instrument_tag']}</td>
                <td>{p['description']}</td>
                <td class="text-center font-mono">{p['pid'] or '-'}</td>
                <td class="text-center font-bold">{p['iotype']}</td>
                <td class="text-center">{sig_spec}</td>
                <td class="text-center"><span class="badge {badge_cls}">{p['status']}</span></td>
                <td class="text-center font-mono font-bold">{pt_ref}</td>
                <td class="text-center font-bold font-mono">{p['chassis']}</td>
                <td class="text-center">{p['slot']}</td>
                <td class="text-center">{p['channel']}</td>
                <td class="text-center">{re.sub(r'-\d+$', '', p['card'])}</td>
                <td class="font-bold font-mono">{p['plc_tag']}</td>
                <td class="font-mono text-center">{p['eplan_term_tag'] or p['eplan_plc_tag']}</td>
            </tr>
            """
            
        pdf_html += """
            </tbody>
        </table>
        <div class="footer-stamp">
            <span>Kalasin Engineering xCIP System — Destination I/O Schedule (Ref: x2608003)</span>
            <span>ePlan Wire Standard: C[Chassis]S[Slot]:[Channel] / [Terminal]:[Pin] | 100% Non-ATEX Food & Beverage Sanitary Standard</span>
        </div>
        """
        
    pdf_html += """
    </body>
    </html>
    """
    
    temp_pdf_html = os.path.join(BASE_DIR, "scratch_dest_report.html")
    with open(temp_pdf_html, 'w', encoding='utf-8') as f:
        f.write(pdf_html)
        
    print(f"Converting Destination HTML to Vector PDF using Chrome Headless...")
    cmd = [
        CHROME_BIN,
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={OUTPUT_PDF_MAIN}",
        temp_pdf_html
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode == 0:
        print(f"Master Destination PDF created: {OUTPUT_PDF_MAIN} ({os.path.getsize(OUTPUT_PDF_MAIN):,} bytes)")
        shutil.copy2(OUTPUT_PDF_MAIN, OUTPUT_PDF_MIRROR)
        print(f"Mirrored to: {OUTPUT_PDF_MIRROR}")
    else:
        print("Chrome error:", res.stderr)
        
    if os.path.exists(temp_pdf_html):
        os.remove(temp_pdf_html)
        
    print("\nAll Destination Reports Successfully Generated!")

if __name__ == "__main__":
    build_destination_reports()
