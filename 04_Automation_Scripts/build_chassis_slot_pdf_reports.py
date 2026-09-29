#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
CHASSIS & SLOT I/O REPORT: VECTOR PDF GENERATOR (LIGHT EXECUTIVE STYLE)
================================================================================
Generates:
  1. Master Consolidated PDF:
     - 02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Chassis_and_Slot_IO_Source_Destination_Report.pdf
     - 03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.pdf
  2. Per-Chassis PDFs (One PDF per Rack with 1 Section/Page per Slot):
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C1_Slots_IO_Report.pdf
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C2_Slots_IO_Report.pdf
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C3_Slots_IO_Report.pdf
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C4_Slots_IO_Report.pdf
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C5_Slots_IO_Report.pdf
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C6_Slots_IO_Report.pdf
     - 03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/Chassis_C7_Slots_IO_Report.pdf
================================================================================
"""

import os
import sys
import re
import shutil
import subprocess
from collections import defaultdict, Counter
import openpyxl

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_LIST_FILE = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-6.xlsx")
TAGS_PLC_FILE = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/02_ePlan_Exports/Tags list PLC.xlsx")

PDF_DEST_MAIN = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Chassis_and_Slot_IO_Source_Destination_Report.pdf")
PDF_DEST_IO_DIR = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.pdf")
SLOT_REPORTS_DIR = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/Slot_by_Slot_Reports")

CHROME_BIN = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"

CHASSIS_METADATA = {
    'C1': {
        'name': 'Main Controller Rack (Chassis 1)',
        'panel': 'Control Cabinet CA1 (Main Control Room)',
        'role': '1756-L950TPSXT Controller + Local Fast I/O',
        'network': 'Stratix 5400 / EN4TR Ring Node 1 (192.168.1.10)',
        'total_slots': 13
    },
    'C2': {
        'name': 'I/O Expansion Rack 1 (Chassis 2)',
        'panel': 'Control Cabinet CA1 (Bay 2)',
        'role': 'Digital & Analog Process I/O Expansion',
        'network': 'EN4TR DLR Ring Node 2 (192.168.1.11)',
        'total_slots': 13
    },
    'C3': {
        'name': 'I/O Expansion Rack 2 (Chassis 3)',
        'panel': 'Control Cabinet CA1 (Bay 3)',
        'role': 'Evaporator, Concentrator & Dosing Skid I/O',
        'network': 'EN4TR DLR Ring Node 3 (192.168.1.12)',
        'total_slots': 13
    },
    'C4': {
        'name': 'I/O Expansion Rack 3 (Chassis 4)',
        'panel': 'Control Cabinet CA1 (Bay 4)',
        'role': 'Spray Dryer Tower & Exhaust Auxiliary I/O',
        'network': 'EN4TR DLR Ring Node 4 (192.168.1.13)',
        'total_slots': 13
    },
    'C5': {
        'name': 'Remote I/O Skid RIO-200 (Chassis 5)',
        'panel': 'Remote I/O Enclosure RIO-200 (Process Floor 1F)',
        'role': 'Remote Field I/O via Fiber Optic Trunk (DLR Ring Node 5)',
        'network': 'EN4TR Fiber DLR Ring Node 5 (192.168.1.14)',
        'total_slots': 13
    },
    'C6': {
        'name': 'MCC Auxiliary I/O (Chassis 6)',
        'panel': 'Motor Control Center Room (MCC Panel)',
        'role': 'Direct Hardwired Motor Starter Feedback & Permissives',
        'network': 'MCC Hardwired Trunk Line',
        'total_slots': 8
    },
    'C7': {
        'name': 'MCC Fieldbus & Network Interface (Chassis 7)',
        'panel': 'Motor Control Center Room (MCC Panel & Field Bus)',
        'role': 'Modbus/Fieldbus Gateways & MCC Feeder Interlocks',
        'network': 'Modbus TCP / EtherNet Trunk',
        'total_slots': 8
    }
}

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

def load_hardware_slot_definitions(io_file):
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

def get_base_css():
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
    *, *::before, *::after {
        box-sizing: border-box;
    }
    body {
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
        margin: 0;
        padding: 0;
        background: #FFFFFF;
        color: #0F172A;
        font-size: 8pt;
        line-height: 1.25;
    }
    .page-break {
        page-break-before: always;
    }
    .no-break {
        page-break-inside: avoid;
    }
    
    /* Cover & Header */
    .doc-header {
        border-bottom: 2px solid #1E293B;
        padding-bottom: 8px;
        margin-bottom: 12px;
        display: flex;
        justify-content: space-between;
        align-items: flex-end;
    }
    .doc-title-block h1 {
        margin: 0;
        font-size: 15pt;
        color: #0F172A;
        font-weight: 800;
        letter-spacing: -0.3px;
        text-transform: uppercase;
    }
    .doc-title-block h2 {
        margin: 2px 0 0 0;
        font-size: 10pt;
        color: #2563EB;
        font-weight: 600;
    }
    .doc-meta-block {
        text-align: right;
        font-size: 7.5pt;
        color: #475569;
        line-height: 1.35;
    }
    .doc-meta-block strong {
        color: #0F172A;
    }

    /* KPI Summary Cards */
    .kpi-row {
        display: grid;
        grid-template-columns: repeat(6, 1fr);
        gap: 8px;
        margin-bottom: 14px;
    }
    .kpi-card {
        border: 1px solid #CBD5E1;
        border-radius: 4px;
        padding: 6px 8px;
        background: #F8FAFC;
        text-align: center;
    }
    .kpi-card.blue { background: #EFF6FF; border-color: #BFDBFE; }
    .kpi-card.green { background: #F0FDF4; border-color: #BBF7D0; }
    .kpi-card.amber { background: #FEFCE8; border-color: #FEF08A; }
    .kpi-card.purple { background: #FAF5FF; border-color: #E9D5FF; }
    .kpi-card .kpi-lbl {
        font-size: 6.5pt;
        text-transform: uppercase;
        font-weight: 700;
        color: #475569;
        margin-bottom: 3px;
    }
    .kpi-card .kpi-val {
        font-size: 11pt;
        font-weight: 800;
        color: #0F172A;
    }
    .kpi-card.blue .kpi-val { color: #1E40AF; }
    .kpi-card.green .kpi-val { color: #166534; }
    .kpi-card.amber .kpi-val { color: #854D0E; }
    .kpi-card.purple .kpi-val { color: #6B21A8; }

    /* Tables */
    table.data-table {
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 14px;
        font-size: 7.2pt;
    }
    table.data-table th, table.data-table td {
        border: 1px solid #CBD5E1;
        padding: 4.5px 5px;
        vertical-align: middle;
    }
    table.data-table thead tr.group-header th {
        font-size: 7.5pt;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.2px;
        padding: 5px;
    }
    table.data-table thead tr.col-header th {
        background: #1E293B;
        color: #FFFFFF;
        font-weight: 600;
        font-size: 7pt;
        text-align: center;
    }
    table.data-table tbody tr:nth-child(even) {
        background: #F8FAFC;
    }
    table.data-table tbody tr:hover {
        background: #F1F5F9;
    }
    
    .th-src { background: #DBEAFE; color: #1E3A8A; }
    .th-sys { background: #E2E8F0; color: #0F172A; }
    .th-dst { background: #FEF3C7; color: #78350F; }

    .text-center { text-align: center; }
    .text-left { text-align: left; }
    .text-right { text-align: right; }
    .font-bold { font-weight: 700; }
    .font-mono { font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace; font-size: 7pt; }

    /* Status Badges */
    .badge {
        display: inline-block;
        padding: 1.5px 5px;
        border-radius: 3px;
        font-size: 6.5pt;
        font-weight: 700;
        text-align: center;
    }
    .badge-active { background: #DCFCE7; color: #14532D; border: 1px solid #86EFAC; }
    .badge-spare { background: #FEF9C3; color: #854D0E; border: 1px solid #FDE047; }
    .badge-cpu { background: #F3E8FF; color: #581C87; border: 1px solid #D8B4FE; }
    .badge-comm { background: #DBEAFE; color: #1E40AF; border: 1px solid #93C5FD; }

    /* Slot Banner */
    .slot-banner {
        background: #1E293B;
        color: #FFFFFF;
        border-radius: 4px 4px 0 0;
        padding: 6px 10px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-top: 10px;
    }
    .slot-banner h3 {
        margin: 0;
        font-size: 9.5pt;
        font-weight: 700;
    }
    .slot-banner .slot-sub {
        font-size: 7.5pt;
        color: #94A3B8;
        font-weight: 400;
    }
    .slot-banner .slot-kpi {
        font-size: 7.5pt;
        font-weight: 600;
        background: rgba(255,255,255,0.12);
        padding: 2px 8px;
        border-radius: 3px;
    }
    
    .footer-stamp {
        border-top: 1px solid #E2E8F0;
        padding-top: 4px;
        margin-top: 8px;
        display: flex;
        justify-content: space-between;
        font-size: 6.5pt;
        color: #64748B;
    }
    """

def generate_slot_table_html(ch, sl, pts):
    meta = CHASSIS_METADATA.get(ch, {})
    ch_name = meta.get('name', f"Chassis {ch}")
    ch_panel = meta.get('panel', 'Main Control Panel')
    ch_net = meta.get('network', 'DLR Ring')
    
    act_cnt = sum(1 for p in pts if not p['is_spare'])
    spr_cnt = len(pts) - act_cnt
    spr_pct = (spr_cnt / len(pts) * 100) if pts else 0
    card_model = re.sub(r'-\d+$', '', pts[0]['card']) if pts and pts[0]['card'] else '1756 I/O'
    iotype = pts[0]['iotype'] if pts else 'I/O'
    
    html = f"""
    <div class="slot-banner">
        <div>
            <h3>{ch} SLOT {sl:02d} — {card_model} ({iotype})</h3>
            <span class="slot-sub">Rack: {ch_name} | Location: {ch_panel} | Protocol: {ch_net}</span>
        </div>
        <div class="slot-kpi">
            Capacity: {len(pts)} Pts &nbsp;|&nbsp; Active: {act_cnt} &nbsp;|&nbsp; Spare: {spr_cnt} ({spr_pct:.1f}%)
        </div>
    </div>
    <table class="data-table">
        <thead>
            <tr class="group-header">
                <th colspan="7" class="th-src text-center">SOURCE SPECIFICATION (PLC RACK / CARD PIN / ePLAN WIRE TAG)</th>
                <th colspan="3" class="th-sys text-center">SIGNAL SPEC</th>
                <th colspan="7" class="th-dst text-center">DESTINATION SPECIFICATION (FIELD JB / MCC / TERMINAL PIN / INSTRUMENT TAG)</th>
            </tr>
            <tr class="col-header">
                <th style="width: 70px;">Point Ref ID</th>
                <th style="width: 25px;">Pt</th>
                <th style="width: 25px;">Ch</th>
                <th style="width: 65px;">Module</th>
                <th style="width: 105px;">PLC Tag Name</th>
                <th style="width: 45px;">Card Pin</th>
                <th style="width: 100px;">ePlan Wire Tag</th>
                <th style="width: 30px;">Type</th>
                <th style="width: 80px;">Signal Spec</th>
                <th style="width: 50px;">Status</th>
                <th style="width: 65px;">Destination</th>
                <th style="width: 80px;">Location / Floor</th>
                <th style="width: 65px;">TB / Fuse</th>
                <th style="width: 35px;">Pin</th>
                <th style="width: 85px;">Instrument Tag</th>
                <th>Instrument Description</th>
                <th style="width: 65px;">P&ID Drawing</th>
            </tr>
        </thead>
        <tbody>
    """
    
    for p in pts:
        pt_ref = f"{p['chassis']}-S{p['slot']:02d}-P{p['point']:02d}"
        badge_cls = "badge-active" if p['status'] == "ACTIVE" else "badge-spare"
        flr_zn = f"{p['floor']} / {p['zone']}" if p['floor'] and p['zone'] else (p['floor'] or p['zone'] or '-')
        sig_spec = p['signal_type'] or ("24VDC Dry Contact" if p['iotype'] == 'DI' else ("24VDC Sourcing" if p['iotype'] == 'DO' else "4-20mA HART"))
        
        html += f"""
        <tr>
            <td class="text-center font-mono font-bold">{pt_ref}</td>
            <td class="text-center">{p['point']}</td>
            <td class="text-center">{p['channel']}</td>
            <td class="text-center">{re.sub(r'-\d+$', '', p['card'])}</td>
            <td class="font-bold font-mono">{p['plc_tag']}</td>
            <td class="text-center font-mono">{p['term2'] or f'Pin {p["channel"]}'}</td>
            <td class="font-mono text-center">{p['eplan_term_tag'] or p['eplan_plc_tag']}</td>
            <td class="text-center font-bold">{p['iotype']}</td>
            <td class="text-center">{sig_spec}</td>
            <td class="text-center"><span class="badge {badge_cls}">{p['status']}</span></td>
            <td class="text-center font-bold">{p['destination'] or '-'}</td>
            <td>{p['location'] or flr_zn}</td>
            <td class="text-center font-mono">{p['terminal'] or '-'}</td>
            <td class="text-center font-mono">{p['terminal_label'] or '-'}</td>
            <td class="font-bold">{p['instrument_tag']}</td>
            <td>{p['description']}</td>
            <td class="text-center font-mono">{p['pid'] or '-'}</td>
        </tr>
        """
        
    html += """
        </tbody>
    </table>
    <div class="footer-stamp">
        <span>Kalasin Engineering xCIP System — Turnkey Automated CIP Skid (Ref: x2608003)</span>
        <span>ePlan Wire Standard: C[Chassis]S[Slot]:[Channel] / [Terminal]:[Pin] | 100% Non-ATEX Food & Beverage Sanitary Standard</span>
    </div>
    """
    return html

def build_pdf_reports():
    print("Loading data for PDF reports...")
    rows = load_data()
    hw_slots = load_hardware_slot_definitions(IO_LIST_FILE)
    print(f"Loaded {len(rows)} data rows and {len(hw_slots)} hardware slot definitions.")
    
    # Group by chassis and slot
    slot_groups = defaultdict(list)
    for r in rows:
        slot_groups[(r['chassis'], r['slot'])].append(r)
        
    def slot_sort_key(item):
        ch, sl = item
        ch_num = int(ch[1:]) if ch.startswith('C') and ch[1:].isdigit() else 99
        sl_num = int(sl) if isinstance(sl, int) or str(sl).isdigit() else 99
        return (ch_num, sl_num)
        
    sorted_slots = sorted(slot_groups.keys(), key=slot_sort_key)
    
    total_pts = len(rows)
    total_act = sum(1 for r in rows if not r['is_spare'])
    total_spr = total_pts - total_act
    
    # -------------------------------------------------------------------------
    # 1. Master Consolidated Report HTML
    # -------------------------------------------------------------------------
    print("Generating Master Consolidated Report HTML...")
    master_html = f"""<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <title>Master Chassis & Slot I/O Source-Destination Report</title>
        <style>{get_base_css()}</style>
    </head>
    <body>
        <!-- Cover / Header Page -->
        <div class="doc-header">
            <div class="doc-title-block">
                <h1>Kalasin Engineering — xCIP Automation System</h1>
                <h2>MASTER CHASSIS & SLOT I/O SOURCE-TO-DESTINATION SCHEDULE</h2>
            </div>
            <div class="doc-meta-block">
                <strong>Project Ref:</strong> x2608003<br>
                <strong>Client:</strong> Ingredion Thailand (Kalasin Plant)<br>
                <strong>Plant Area:</strong> Sprint 18K TPA Spray Dryer Plant<br>
                <strong>Revision:</strong> Rev 3.6 | <strong>Date:</strong> 28-Sep-2026
            </div>
        </div>

        <div class="kpi-row">
            <div class="kpi-card blue">
                <div class="kpi-lbl">Total Chassis</div>
                <div class="kpi-val">7 Racks (C1-C7)</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-lbl">Total Capacity</div>
                <div class="kpi-val">{total_pts:,} Pts</div>
            </div>
            <div class="kpi-card green">
                <div class="kpi-lbl">Active Field I/O</div>
                <div class="kpi-val">{total_act:,} Pts</div>
            </div>
            <div class="kpi-card amber">
                <div class="kpi-lbl">Spare Margin</div>
                <div class="kpi-val">{total_spr:,} ({(total_spr/total_pts*100):.1f}%)</div>
            </div>
            <div class="kpi-card purple">
                <div class="kpi-lbl">Main Controller</div>
                <div class="kpi-val">1756-L950TPSXT</div>
            </div>
            <div class="kpi-card blue">
                <div class="kpi-lbl">DLR Ring Network</div>
                <div class="kpi-val">Stratix 5400 / EN4TR</div>
            </div>
        </div>

        <!-- Section 1: Master Hardware Allocation Matrix -->
        <h3 style="margin: 6px 0; font-size: 10pt; color: #1E293B;">SECTION 1: MASTER CHASSIS & HARDWARE ALLOCATION MATRIX</h3>
        <table class="data-table" style="font-size: 7pt; margin-bottom: 20px;">
            <thead>
                <tr class="col-header">
                    <th style="width: 50px;">Chassis</th>
                    <th style="width: 140px;">Chassis Name & Location</th>
                    <th style="width: 35px;">Slot</th>
                    <th style="width: 90px;">Module Model</th>
                    <th>Module Description & Functional Role</th>
                    <th style="width: 45px;">Type</th>
                    <th style="width: 45px;">Capacity</th>
                    <th style="width: 45px;">Active</th>
                    <th style="width: 45px;">Spare</th>
                    <th style="width: 50px;">Spare %</th>
                    <th style="width: 120px;">Primary Destinations</th>
                    <th style="width: 70px;">Status</th>
                </tr>
            </thead>
            <tbody>
    """
    
    for ch_key in ['C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7']:
        meta = CHASSIS_METADATA.get(ch_key, {})
        num_slots = meta.get('total_slots', 13)
        ch_name = meta.get('name', ch_key)
        
        for sl in range(num_slots):
            pts = slot_groups.get((ch_key, sl), [])
            hw_def = hw_slots.get((ch_key, sl), {})
            
            card_model = hw_def.get('card_model', '')
            if not card_model and pts:
                card_model = re.sub(r'-\d+$', '', pts[0]['card'])
                
            iotype = hw_def.get('card_type', '')
            if not iotype and pts:
                iotype = pts[0]['iotype']
                
            cap = len(pts)
            act = sum(1 for p in pts if not p['is_spare'])
            spr = cap - act
            spr_pct = (spr / cap * 100) if cap > 0 else 0
            
            dests = sorted(list({p['destination'] for p in pts if p['destination']}))
            dest_str = ", ".join(dests) if dests else ("CPU/Comm" if cap == 0 else "-")
            
            desc_str = ""
            status_badge = ""
            if card_model == '1756-L950TPSXT':
                desc_str = "ControlLogix 5590 Redundant Process Controller"
                status_badge = '<span class="badge badge-cpu">CPU</span>'
            elif card_model == '1756-EN4TR':
                desc_str = "Dual-Port Gigabit EtherNet/IP DLR Comm Adapter"
                status_badge = '<span class="badge badge-comm">COMM</span>'
            elif card_model == '1756-IB32':
                desc_str = "32-Channel 24VDC Sink/Source Digital Input Module"
                status_badge = '<span class="badge badge-active">ASSIGNED</span>' if act > 0 else '<span class="badge badge-spare">RESERVED</span>'
            elif card_model == '1756-OB32':
                desc_str = "32-Channel 24VDC Sourcing Digital Output Module"
                status_badge = '<span class="badge badge-active">ASSIGNED</span>' if act > 0 else '<span class="badge badge-spare">RESERVED</span>'
            elif card_model == '1756-IF16':
                desc_str = "16-Channel High-Speed Isolated Analog Input Module"
                status_badge = '<span class="badge badge-active">ASSIGNED</span>' if act > 0 else '<span class="badge badge-spare">RESERVED</span>'
            elif card_model == '1756-OF8':
                desc_str = "8-Channel Isolated Analog Output Module"
                status_badge = '<span class="badge badge-active">ASSIGNED</span>' if act > 0 else '<span class="badge badge-spare">RESERVED</span>'
            elif card_model == '1756-N2':
                desc_str = "Slot Filler / Future Expansion Space"
                status_badge = '<span class="badge">EMPTY</span>'
            else:
                desc_str = f"Direct Hardware Interface / Bus"
                status_badge = '<span class="badge badge-active">ASSIGNED</span>' if act > 0 else '<span class="badge badge-spare">RESERVED</span>'
                
            master_html += f"""
            <tr>
                <td class="text-center font-bold font-mono">{ch_key}</td>
                <td>{ch_name}</td>
                <td class="text-center font-mono">{sl}</td>
                <td class="text-center font-bold font-mono">{card_model or '-'}</td>
                <td>{desc_str}</td>
                <td class="text-center font-bold">{iotype or '-'}</td>
                <td class="text-center">{cap if cap > 0 else '-'}</td>
                <td class="text-center">{act if cap > 0 else '-'}</td>
                <td class="text-center">{spr if cap > 0 else '-'}</td>
                <td class="text-center">{f'{spr_pct:.1f}%' if cap > 0 else '-'}</td>
                <td>{dest_str}</td>
                <td class="text-center">{status_badge}</td>
            </tr>
            """
            
    master_html += """
            </tbody>
        </table>
        
        <div class="footer-stamp">
            <span>Section 1 Master Hardware Schedule | Ingredion Sprint 18K Plant</span>
            <span>Generated from Verified IO List Tag35-6 and ePlan PLC Tags Schedule</span>
        </div>
    """
    
    # Section 2: Every single slot on its own page!
    print("Generating Slot-by-Slot Detail Pages...")
    for (ch, sl) in sorted_slots:
        pts = slot_groups[(ch, sl)]
        pts.sort(key=lambda x: x['point'])
        master_html += '<div class="page-break"></div>'
        master_html += generate_slot_table_html(ch, sl, pts)
        
    master_html += """
    </body>
    </html>
    """
    
    # Write temporary Master HTML
    temp_master_html = os.path.join(BASE_DIR, "scratch_master_report.html")
    with open(temp_master_html, 'w', encoding='utf-8') as f:
        f.write(master_html)
        
    print(f"Converting Master HTML to PDF using Chrome Headless...")
    cmd = [
        CHROME_BIN,
        "--headless",
        "--disable-gpu",
        "--no-pdf-header-footer",
        f"--print-to-pdf={PDF_DEST_MAIN}",
        temp_master_html
    ]
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print("Chrome error:", res.stderr)
    else:
        print(f"Master PDF created: {PDF_DEST_MAIN} ({os.path.getsize(PDF_DEST_MAIN):,} bytes)")
        # Mirror to 03_IO_Lists_and_Schedules
        shutil.copy2(PDF_DEST_MAIN, PDF_DEST_IO_DIR)
        print(f"Mirrored to: {PDF_DEST_IO_DIR}")
        
    # Clean up temp HTML
    if os.path.exists(temp_master_html):
        os.remove(temp_master_html)
        
    # -------------------------------------------------------------------------
    # 2. Per-Chassis PDFs (Chassis_C1_Slots_IO_Report.pdf to C7)
    # -------------------------------------------------------------------------
    print("\nGenerating Per-Chassis PDFs...")
    all_chassis = sorted(list({ch for (ch, sl) in sorted_slots}), key=lambda x: int(x[1:]) if x[1:].isdigit() else 99)
    for ch in all_chassis:
        ch_meta = CHASSIS_METADATA.get(ch, {})
        ch_slots = [(c, s) for (c, s) in sorted_slots if c == ch]
        ch_pts = [p for (c, s) in ch_slots for p in slot_groups[(c, s)]]
        
        c_act = sum(1 for p in ch_pts if not p['is_spare'])
        c_spr = len(ch_pts) - c_act
        
        ch_html = f"""<!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>Chassis {ch} Slot-by-Slot I/O Report</title>
            <style>{get_base_css()}</style>
        </head>
        <body>
            <div class="doc-header">
                <div class="doc-title-block">
                    <h1>Kalasin Engineering — {ch} ({ch_meta.get('name', ch)})</h1>
                    <h2>SLOT-BY-SLOT I/O SOURCE AND DESTINATION WIRING SCHEDULE</h2>
                </div>
                <div class="doc-meta-block">
                    <strong>Project Ref:</strong> x2608003<br>
                    <strong>Rack Location:</strong> {ch_meta.get('panel', '-')}<br>
                    <strong>Network:</strong> {ch_meta.get('network', '-')}<br>
                    <strong>Revision:</strong> Rev 3.6 | <strong>Date:</strong> 28-Sep-2026
                </div>
            </div>

            <div class="kpi-row">
                <div class="kpi-card blue">
                    <div class="kpi-lbl">Chassis ID</div>
                    <div class="kpi-val">{ch}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-lbl">Active Slots</div>
                    <div class="kpi-val">{len(ch_slots)} Slots</div>
                </div>
                <div class="kpi-card green">
                    <div class="kpi-lbl">Total Capacity</div>
                    <div class="kpi-val">{len(ch_pts)} Channels</div>
                </div>
                <div class="kpi-card blue">
                    <div class="kpi-lbl">Active Points</div>
                    <div class="kpi-val">{c_act} Points</div>
                </div>
                <div class="kpi-card amber">
                    <div class="kpi-lbl">Spare Points</div>
                    <div class="kpi-val">{c_spr} ({(c_spr/len(ch_pts)*100):.1f}%)</div>
                </div>
                <div class="kpi-card purple">
                    <div class="kpi-lbl">Panel Location</div>
                    <div class="kpi-val">{ch_meta.get('panel', '-')[:18]}</div>
                </div>
            </div>
        """
        
        # Add slot tables
        for idx, (c_id, sl) in enumerate(ch_slots):
            pts = slot_groups[(c_id, sl)]
            pts.sort(key=lambda x: x['point'])
            if idx > 0:
                ch_html += '<div class="page-break"></div>'
            ch_html += generate_slot_table_html(c_id, sl, pts)
            
        ch_html += """
        </body>
        </html>
        """
        
        temp_ch_html = os.path.join(BASE_DIR, f"scratch_{ch}.html")
        with open(temp_ch_html, 'w', encoding='utf-8') as f:
            f.write(ch_html)
            
        out_pdf = os.path.join(SLOT_REPORTS_DIR, f"Chassis_{ch}_Slots_IO_Report.pdf")
        cmd = [
            CHROME_BIN,
            "--headless",
            "--disable-gpu",
            "--no-pdf-header-footer",
            f"--print-to-pdf={out_pdf}",
            temp_ch_html
        ]
        res = subprocess.run(cmd, capture_output=True, text=True)
        if res.returncode == 0:
            print(f"Saved: {out_pdf} ({os.path.getsize(out_pdf):,} bytes)")
        if os.path.exists(temp_ch_html):
            os.remove(temp_ch_html)
            
    print("\nAll PDF Reports Generated Successfully!")

if __name__ == "__main__":
    build_pdf_reports()
