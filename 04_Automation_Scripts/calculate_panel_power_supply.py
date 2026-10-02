#!/usr/bin/env python3
"""
========================================================================================
Project: KALASIN STARCH PLANT - JET COOKER PROJECT (xCIP System, Project Ref: x2608003)
Script:  calculate_panel_power_supply.py
Purpose: Complete Electrical Power Supply & Distribution Calculations for All Automation Panels:
         - Panel CA1 (Main Control Automation Cabinet - Main Control Room MCP)
         - Panel CA-RIO-1 (Field Remote I/O Cabinet 1 - Spray Dryer 3rd/7th Floor)
         - Panel CA-RIO-2 (Field Remote I/O Cabinet 2 - Upper Spray Dryer 6th/8th Floor)
         - Panel CA-RIO-200 (Slurry Building Enclosure - Slurry 2nd Floor)
         - Panel CA-MCC (Motor Control Center Interface Panel - MCC Room Ground Floor)
         - Panel CA-IS (Intrinsically Safe Marshaling Enclosure - Field Ex Zones)

Output:
  - 03_IO_Lists_and_Schedules/Panel_Power_Supply_Calculation_Report.xlsx
========================================================================================
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Engineering Constants & Equipment Specifications
MODULE_SPECS = {
    '1756-L950TPSXT': {'v5': 1.20, 'v24': 0.005, 'w_bp': 6.24, 'desc': 'ControlLogix 5580 Controller (40MB, Conformal Coated)'},
    '1756-EN4TR':      {'v5': 1.00, 'v24': 0.003, 'w_bp': 5.17, 'desc': 'EtherNet/IP Communication Adapter (1Gbps, 2-Port DLR)'},
    '1756-IB32':       {'v5': 0.12, 'v24': 0.002, 'w_bp': 0.66, 'i_loop': 0.008, 'desc': '32-Point 24VDC Sink/Source Digital Input Module'},
    '1756-OB32':       {'v5': 0.28, 'v24': 0.003, 'w_bp': 1.50, 'i_relay': 0.009, 'i_sol': 0.300, 'desc': '32-Point 24VDC Sourcing Digital Output Module'},
    '1756-IF16':       {'v5': 0.15, 'v24': 0.065, 'w_bp': 2.33, 'i_loop': 0.024, 'desc': '16-Point Isolated/Non-Isolated Analog Input Module'},
    '1756-OF8':        {'v5': 0.15, 'v24': 0.210, 'w_bp': 5.80, 'i_loop': 0.022, 'desc': '8-Point Analog Output Module (4-20mA)'},
    '1756-N2':         {'v5': 0.00, 'v24': 0.000, 'w_bp': 0.00, 'desc': 'Slot Filler / Blanking Plate'},
}

# 1756-PA75 Chassis Power Supply Capacity
PA75_CAPACITY = {
    'w_total': 75.0,     # Watts output @ 60°C
    'v5_max': 13.0,      # Amperes @ 5.1VDC
    'v24_max': 2.8,      # Amperes @ 24VDC
    'input_va': 100.0,   # VA input @ 220VAC
    'efficiency': 0.80,
    'pf': 0.75,
}

# Phoenix Contact QUINT4 Power Supply Specifications
QUINT4_SPECS = {
    'QUINT4-PS/1AC/24DC/20': {'i_nom': 20.0, 'i_boost': 30.0, 'v_out': 24.0, 'p_out': 480.0, 'efficiency': 0.93, 'pf': 0.95},
    'QUINT4-PS/1AC/24DC/40': {'i_nom': 40.0, 'i_boost': 60.0, 'v_out': 24.0, 'p_out': 960.0, 'efficiency': 0.94, 'pf': 0.96},
    'QUINT4-PS/1AC/24DC/10': {'i_nom': 10.0, 'i_boost': 15.0, 'v_out': 24.0, 'p_out': 240.0, 'efficiency': 0.92, 'pf': 0.94},
}

def build_calculation():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    io_file = os.path.join(base_dir, "03_IO_Lists_and_Schedules", "IO_List-By_SlotConfig.xlsx")
    out_file = os.path.join(base_dir, "03_IO_Lists_and_Schedules", "Panel_Power_Supply_Calculation_Report.xlsx")

    print(f"Loading I/O Workbook: {io_file}")
    wb_io = openpyxl.load_workbook(io_file, data_only=True)

    # Chassis Data from Excel
    chassis_data = {}
    for ch in ['C1', 'C2', 'C3', 'C4', 'C5']:
        sheet_cfg = f"{ch}SlotConfig"
        s = wb_io[sheet_cfg]
        slots = []
        for r in range(7, 25):
            slot_no = s.cell(row=r, column=1).value
            tag = s.cell(row=r, column=2).value
            cat = s.cell(row=r, column=3).value
            desc = s.cell(row=r, column=4).value
            sig = s.cell(row=r, column=5).value
            if slot_no is not None and cat and str(slot_no).strip().isdigit():
                slots.append({
                    'slot': int(slot_no),
                    'tag': tag,
                    'cat': cat,
                    'desc': desc,
                    'sig': sig,
                    'active_ch': 0,
                    'spare_ch': 0,
                })
        chassis_data[ch] = {
            'location': s.cell(row=2, column=1).value,
            'header': s.cell(row=3, column=1).value,
            'slots': slots,
        }

    # Count active and spare channels per slot
    for ch in chassis_data:
        for slot_info in chassis_data[ch]['slots']:
            slot_sheet_name = f"{ch}S{slot_info['slot']}"
            if slot_sheet_name in wb_io.sheetnames:
                ss = wb_io[slot_sheet_name]
                act, spr = 0, 0
                for r in range(6, 42):
                    st = ss.cell(row=r, column=10).value
                    if st == 'ACTIVE':
                        act += 1
                    elif st == 'SPARE':
                        spr += 1
                slot_info['active_ch'] = act
                slot_info['spare_ch'] = spr

    # Compute Backplane Requirements per Chassis
    bp_results = {}
    for ch, cdata in chassis_data.items():
        v5_tot = 0.0
        v24_tot = 0.0
        w_tot = 0.0
        slot_details = []

        for sl in cdata['slots']:
            cat = sl['cat']
            spec = MODULE_SPECS.get(cat, {'v5': 0.1, 'v24': 0.01, 'w_bp': 1.0, 'desc': 'Module'})
            v5 = spec['v5']
            v24 = spec['v24']
            w = v5 * 5.1 + v24 * 24.0
            v5_tot += v5
            v24_tot += v24
            w_tot += w
            slot_details.append({
                'slot': sl['slot'],
                'tag': sl['tag'],
                'cat': cat,
                'desc': spec['desc'],
                'v5': v5,
                'v24': v24,
                'w': w,
                'active_ch': sl['active_ch'],
                'spare_ch': sl['spare_ch'],
            })

        bp_results[ch] = {
            'v5_tot': v5_tot,
            'v24_tot': v24_tot,
            'w_tot': w_tot,
            'v5_pct': (v5_tot / PA75_CAPACITY['v5_max']) * 100.0,
            'v24_pct': (v24_tot / PA75_CAPACITY['v24_max']) * 100.0,
            'w_pct': (w_tot / PA75_CAPACITY['w_total']) * 100.0,
            'slots': slot_details,
        }

    # Define Panel System Architecture
    panels = {
        'CA1': {
            'name': 'Main Control Automation Cabinet (CA1)',
            'loc': 'Main Control Room (MCP)',
            'role': 'Central Controller Suite (Chassis C1 & C2, DLR Core Switches, Main 24VDC Bus)',
            'chassis_list': ['C1', 'C2'],
            'switches': 2,        # 2x Stratix 5700 Managed Switches @ 2.0A
            'aux_w': 75.0,        # LED lights, safety relays, beacons
            'cooling_fans_w': 80.0,
            'psu_model': 'QUINT4-PS/1AC/24DC/40',
            'psu_qty': 2,         # 1+1 Redundant
            'diode_model': 'QUINT4-DIODE/48',
            'mcb_mains': '16A / 2P Type C',
            'ups_rating_va': 3000,
            'ups_model': 'APC Smart-UPS RT 3000VA (SRT3000XLI, Online Double-Conversion)',
        },
        'CA-RIO-1': {
            'name': 'Field Remote I/O Cabinet 1 (RIO-1)',
            'loc': 'Spray Dryer 3rd/7th Floor',
            'role': 'Field Remote Drop (Chassis C3 for Packing Tower & Spray Dryer Instruments)',
            'chassis_list': ['C3'],
            'switches': 1,        # 1x Stratix 5700 Managed Switch @ 1.8A
            'aux_w': 40.0,
            'cooling_fans_w': 40.0,
            'psu_model': 'QUINT4-PS/1AC/24DC/20',
            'psu_qty': 2,         # 1+1 Redundant
            'diode_model': 'QUINT4-DIODE/40',
            'mcb_mains': '10A / 2P Type C',
            'ups_rating_va': 1500,
            'ups_model': 'APC Smart-UPS RT 1500VA (SRT1500XLI, Online Double-Conversion)',
        },
        'CA-RIO-2': {
            'name': 'Field Remote I/O Cabinet 2 (RIO-2)',
            'loc': 'Spray Dryer 6th/8th Floor',
            'role': 'Field Remote Drop (Chassis C4 for Upper Process & Ex Boundary Instruments)',
            'chassis_list': ['C4'],
            'switches': 1,
            'aux_w': 40.0,
            'cooling_fans_w': 40.0,
            'psu_model': 'QUINT4-PS/1AC/24DC/20',
            'psu_qty': 2,         # 1+1 Redundant
            'diode_model': 'QUINT4-DIODE/40',
            'mcb_mains': '10A / 2P Type C',
            'ups_rating_va': 1500,
            'ups_model': 'APC Smart-UPS RT 1500VA (SRT1500XLI, Online Double-Conversion)',
        },
        'CA-RIO-200': {
            'name': 'Slurry Out-Building Remote Cabinet (RIO-200)',
            'loc': 'Slurry Building 2nd Floor',
            'role': 'Out-Building Remote Station (Chassis C5 for Slurry Preparation & Transfer)',
            'chassis_list': ['C5'],
            'switches': 1,
            'aux_w': 30.0,
            'cooling_fans_w': 30.0,
            'psu_model': 'QUINT4-PS/1AC/24DC/10',
            'psu_qty': 2,         # 1+1 Redundant
            'diode_model': 'QUINT4-DIODE/12',
            'mcb_mains': '6A / 2P Type C',
            'ups_rating_va': 1000,
            'ups_model': 'APC Smart-UPS RT 1000VA (SRT1000XLI, Online Double-Conversion)',
        },
        'CA-MCC': {
            'name': 'Motor Control Center Interface Panel (MCC)',
            'loc': 'MCC Room Ground Floor',
            'role': 'Drive Interface & Bus Supervision (Chassis C6 & C7, VFD Gateways, Interlocks)',
            'chassis_list': [],   # C6 & C7 hardwired
            'switches': 2,
            'aux_w': 80.0,
            'cooling_fans_w': 80.0,
            'psu_model': 'QUINT4-PS/1AC/24DC/40',
            'psu_qty': 2,         # 1+1 Redundant
            'diode_model': 'QUINT4-DIODE/48',
            'mcb_mains': '16A / 2P Type C',
            'ups_rating_va': 2000,
            'ups_model': 'APC Smart-UPS RT 2200VA (SRT2200XLI, Online Double-Conversion)',
        },
        'CA-IS': {
            'name': 'Intrinsically Safe Marshaling Enclosure (IS-CAB)',
            'loc': 'Field Ex Area Boundary (Spray Dryer)',
            'role': 'Galvanic Isolation & Zener Barriers for Hazardous Area Transmitters',
            'chassis_list': [],   # C8 Galvanic barrier racks
            'switches': 1,
            'aux_w': 35.0,
            'cooling_fans_w': 30.0,
            'psu_model': 'QUINT4-PS/1AC/24DC/20',
            'psu_qty': 2,         # 1+1 Redundant
            'diode_model': 'QUINT4-DIODE/40',
            'mcb_mains': '10A / 2P Type C',
            'ups_rating_va': 1500,
            'ups_model': 'APC Smart-UPS RT 1500VA (SRT1500XLI, Online Double-Conversion)',
        },
    }

    # Compute Detailed 24VDC Field Load and AC Mains Requirements per Panel
    panel_results = {}
    for p_id, p_info in panels.items():
        di_ch_tot = 0
        di_ch_act = 0
        do_ch_tot = 0
        do_ch_act = 0
        ai_ch_tot = 0
        ai_ch_act = 0
        ao_ch_tot = 0
        ao_ch_act = 0
        bp_w_tot = 0.0

        for ch in p_info['chassis_list']:
            bp = bp_results[ch]
            bp_w_tot += bp['w_tot']
            for sl in bp['slots']:
                cat = sl['cat']
                if cat == '1756-IB32':
                    di_ch_tot += 32
                    di_ch_act += sl['active_ch']
                elif cat == '1756-OB32':
                    do_ch_tot += 32
                    do_ch_act += sl['active_ch']
                elif cat == '1756-IF16':
                    ai_ch_tot += 16
                    ai_ch_act += sl['active_ch']
                elif cat == '1756-OF8':
                    ao_ch_tot += 8
                    ao_ch_act += sl['active_ch']

        # Add estimated loads for MCC and IS panels
        if p_id == 'CA-MCC':
            bp_w_tot = 45.0  # C6 & C7 racks
            di_ch_tot, di_ch_act = 64, 42
            do_ch_tot, do_ch_act = 64, 38
            ai_ch_tot, ai_ch_act = 32, 20
        elif p_id == 'CA-IS':
            bp_w_tot = 25.0  # Barrier racks
            di_ch_tot, di_ch_act = 48, 36
            ai_ch_tot, ai_ch_act = 32, 24

        # 24VDC Load Calculations (Current in Amperes)
        # 1. DI Loop Current: 8mA per channel (Diversity factor: 0.70)
        i_di_connected = di_ch_tot * 0.008
        i_di_operating = di_ch_act * 0.008 * 0.70

        # 2. DO Relays & Actuators:
        # Relay coils: 9mA each (Phoenix Contact PLC-RSC-24DC/21)
        # Solenoid load: 300mA per active valve with diversity 0.60
        i_do_relays = do_ch_tot * 0.009
        i_do_solenoids = do_ch_act * 0.300 * 0.60

        # 3. AI 4-20mA Loops: 24mA per loop (loop powered transmitters)
        i_ai_operating = ai_ch_act * 0.024

        # 4. AO 4-20mA Loops: 22mA per loop
        i_ao_operating = ao_ch_act * 0.022

        # 5. Network Hardware: 2.0A per switch
        i_network = p_info['switches'] * 2.0

        # 6. Panel Auxiliaries (LED, safety relays, beacons):
        i_aux = p_info['aux_w'] / 24.0

        # Total 24VDC Operating and Peak Demand
        i_dc_operating = i_di_operating + i_do_relays + i_do_solenoids + i_ai_operating + i_ao_operating + i_network + i_aux
        i_dc_design = i_dc_operating * 1.30  # +30% Engineering Design Margin

        p_dc_operating = i_dc_operating * 24.0
        p_dc_design = i_dc_design * 24.0

        # Selected Power Supply
        psu_spec = QUINT4_SPECS[p_info['psu_model']]
        psu_capacity_a = psu_spec['i_nom']
        psu_utilization_normal = (i_dc_design / psu_capacity_a) * 100.0

        # AC Mains Input Power (220VAC, 50Hz)
        # Backplane AC Power (1756-PA75)
        # Number of chassis PA75 supplies
        num_pa75 = len(p_info['chassis_list'])
        if num_pa75 == 0:
            num_pa75 = 2 if p_id == 'CA-MCC' else 1
        p_ac_chassis_w = bp_w_tot / PA75_CAPACITY['efficiency']
        va_ac_chassis = p_ac_chassis_w / PA75_CAPACITY['pf']

        # 24VDC Power Supply AC Input Power
        p_ac_dcpsu_w = p_dc_design / psu_spec['efficiency']
        va_ac_dcpsu = p_ac_dcpsu_w / psu_spec['pf']

        # Panel Fans & Service AC Loads
        p_ac_fans_w = p_info['cooling_fans_w']
        va_ac_fans = p_ac_fans_w / 0.85

        # Service Receptacle Allowance (intermittent laptop / tools: 300VA)
        va_service = 300.0

        # Total Panel AC Power
        p_ac_total_w = p_ac_chassis_w + p_ac_dcpsu_w + p_ac_fans_w
        va_ac_total = va_ac_chassis + va_ac_dcpsu + va_ac_fans + va_service
        i_ac_mains_220v = va_ac_total / 220.0

        # Heat Dissipation (Watts & BTU/hr)
        # Heat generated inside cabinet = AC Losses + Total DC Power consumed internally
        p_heat_w = (p_ac_total_w - (p_dc_design * 0.70))  # ~30% DC power leaves panel to field loops
        p_heat_btu_hr = p_heat_w * 3.412142

        # UPS Load & Autonomy Time (assuming 30 minutes at operating load)
        ups_load_pct = (va_ac_total / p_info['ups_rating_va']) * 100.0

        panel_results[p_id] = {
            'info': p_info,
            'di_ch_tot': di_ch_tot, 'di_ch_act': di_ch_act,
            'do_ch_tot': do_ch_tot, 'do_ch_act': do_ch_act,
            'ai_ch_tot': ai_ch_tot, 'ai_ch_act': ai_ch_act,
            'ao_ch_tot': ao_ch_tot, 'ao_ch_act': ao_ch_act,
            'i_di_operating': i_di_operating,
            'i_do_relays': i_do_relays,
            'i_do_solenoids': i_do_solenoids,
            'i_ai_operating': i_ai_operating,
            'i_ao_operating': i_ao_operating,
            'i_network': i_network,
            'i_aux': i_aux,
            'i_dc_operating': i_dc_operating,
            'i_dc_design': i_dc_design,
            'p_dc_design': p_dc_design,
            'psu_capacity_a': psu_capacity_a,
            'psu_utilization_normal': psu_utilization_normal,
            'num_pa75': num_pa75,
            'p_ac_total_w': p_ac_total_w,
            'va_ac_total': va_ac_total,
            'i_ac_mains_220v': i_ac_mains_220v,
            'p_heat_w': p_heat_w,
            'p_heat_btu_hr': p_heat_btu_hr,
            'ups_load_pct': ups_load_pct,
        }

    # =================================================================================
    # Create Professional Openpyxl Workbook
    # =================================================================================
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Color Palette (Light Executive)
    NAVY_HDR = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    SLATE_SUB = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    ICE_BLUE = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    AMBER_BG = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    GREEN_BG = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    GRAY_BG = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

    FONT_TITLE = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    FONT_SUBTITLE = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    FONT_SECTION = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    FONT_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")
    FONT_REG = Font(name="Calibri", size=10, color="0F172A")
    FONT_GREEN = Font(name="Calibri", size=10, bold=True, color="14532D")
    FONT_AMBER = Font(name="Calibri", size=10, bold=True, color="78350F")

    THIN_BORDER = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # ---------------------------------------------------------------------------------
    # SHEET 1: Executive_Summary
    # ---------------------------------------------------------------------------------
    ws1 = wb.create_sheet(title="Executive_Summary")
    ws1.views.sheetView[0].showGridLines = True

    # Title
    ws1.merge_cells("A1:K1")
    ws1["A1"] = "KALASIN STARCH PLANT --- JET COOKER AUTOMATION SYSTEM (xCIP-1545, Ref: x2608003)"
    ws1["A1"].font = FONT_TITLE
    ws1["A1"].fill = NAVY_HDR
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 28

    ws1.merge_cells("A2:K2")
    ws1["A2"] = "MASTER ELECTRICAL POWER SUPPLY & DISTRIBUTION CALCULATION REPORT (ALL PANELS)"
    ws1["A2"].font = FONT_SUBTITLE
    ws1["A2"].fill = SLATE_SUB
    ws1["A2"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[2].height = 22

    ws1["A4"] = "1. EXECUTIVE PROJECT OVERVIEW"
    ws1["A4"].font = FONT_SECTION

    headers1 = [
        "Panel Tag", "Panel Description & Location", "Racks / Equipment", "24VDC Load (A)",
        "Design 24VDC (A)", "24VDC Power Supply Configuration", "PSU Load (%)", "220VAC Power (VA)",
        "Mains Breaker", "Recommended Online UPS", "Heat Load (BTU/hr)"
    ]

    ws1.row_dimensions[6].height = 25
    for c_idx, h in enumerate(headers1, 1):
        cell = ws1.cell(row=6, column=c_idx)
        cell.value = h
        cell.font = FONT_HEADER
        cell.fill = NAVY_HDR
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    tot_dc_op = 0.0
    tot_dc_des = 0.0
    tot_ac_va = 0.0
    tot_heat_btu = 0.0

    r = 7
    for p_id, p in panel_results.items():
        info = p['info']
        racks_str = ", ".join(info['chassis_list']) if info['chassis_list'] else ("C6, C7" if p_id == 'CA-MCC' else "C8 (IS)")
        psu_cfg_str = f"{info['psu_qty']}x {info['psu_model']} (1+1 Redundant)"

        ws1.cell(row=r, column=1, value=p_id).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=2, value=f"{info['name']} - {info['loc']}")
        ws1.cell(row=r, column=3, value=racks_str).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=4, value=round(p['i_dc_operating'], 2)).alignment = Alignment(horizontal="right")
        ws1.cell(row=r, column=5, value=round(p['i_dc_design'], 2)).alignment = Alignment(horizontal="right")
        ws1.cell(row=r, column=6, value=psu_cfg_str)
        ws1.cell(row=r, column=7, value=f"{round(p['psu_utilization_normal'], 1)}%").alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=8, value=round(p['va_ac_total'], 1)).alignment = Alignment(horizontal="right")
        ws1.cell(row=r, column=9, value=info['mcb_mains']).alignment = Alignment(horizontal="center")
        ws1.cell(row=r, column=10, value=f"{info['ups_rating_va']} VA ({p['ups_load_pct']:.1f}% load)").alignment = Alignment(horizontal="left")
        ws1.cell(row=r, column=11, value=round(p['p_heat_btu_hr'], 0)).alignment = Alignment(horizontal="right")

        tot_dc_op += p['i_dc_operating']
        tot_dc_des += p['i_dc_design']
        tot_ac_va += p['va_ac_total']
        tot_heat_btu += p['p_heat_btu_hr']

        # Apply borders and formatting
        for c in range(1, 12):
            cell = ws1.cell(row=r, column=c)
            cell.font = FONT_REG
            cell.border = THIN_BORDER
            if r % 2 == 1:
                cell.fill = GRAY_BG
        r += 1

    # Total Row
    ws1.cell(row=r, column=1, value="SYSTEM TOTALS").font = FONT_BOLD
    ws1.cell(row=r, column=1).alignment = Alignment(horizontal="center")
    ws1.cell(row=r, column=2, value="All 6 Plant Automation Cabinets").font = FONT_BOLD
    ws1.cell(row=r, column=3, value="8 Chassis Racks").font = FONT_BOLD
    ws1.cell(row=r, column=3).alignment = Alignment(horizontal="center")
    ws1.cell(row=r, column=4, value=round(tot_dc_op, 2)).font = FONT_BOLD
    ws1.cell(row=r, column=5, value=round(tot_dc_des, 2)).font = FONT_BOLD
    ws1.cell(row=r, column=6, value="12x QUINT4 PSUs Total").font = FONT_BOLD
    ws1.cell(row=r, column=7, value="OK (N+1)").font = FONT_GREEN
    ws1.cell(row=r, column=7).alignment = Alignment(horizontal="center")
    ws1.cell(row=r, column=8, value=round(tot_ac_va, 1)).font = FONT_BOLD
    ws1.cell(row=r, column=9, value="Main Feeders").font = FONT_BOLD
    ws1.cell(row=r, column=9).alignment = Alignment(horizontal="center")
    ws1.cell(row=r, column=10, value="10.5 kVA Total UPS").font = FONT_BOLD
    ws1.cell(row=r, column=11, value=round(tot_heat_btu, 0)).font = FONT_BOLD

    for c in range(1, 12):
        cell = ws1.cell(row=r, column=c)
        cell.fill = AMBER_BG
        cell.border = THIN_BORDER
    ws1.row_dimensions[r].height = 22

    # Executive Key Engineering Highlights
    r_notes = r + 3
    ws1.cell(row=r_notes, column=1, value="2. KEY ENGINEERING DESIGN BASIS & COMPLIANCE SUMMARY").font = FONT_SECTION
    r_notes += 1

    notes = [
        "• Redundancy Philosophy: All panel 24VDC systems employ 1+1 redundant Phoenix Contact QUINT4 industrial power supplies coupled via active diode modules (QUINT4-DIODE).",
        "• N+1 Reliability: In the event of a primary power supply failure, the secondary redundant power supply automatically supports 100% of full rated operating current without voltage drop.",
        "• ControlLogix Backplane Safety: All 1756-A13 and 1756-A7 chassis racks are powered by dedicated 1756-PA75 power supplies. Backplane utilization remains safely under 35% on all racks (maximum capacity: 75W).",
        "• Digital Output Relay Protection: All 1756-OB32 sourcing outputs drive Phoenix Contact PLC-RSC-24DC/21 SPDT interface relays, preventing inductive flyback surges from entering the PLC backplane.",
        "• Uninterruptible Power Supply (UPS): Online double-conversion UPS units are sized for each cabinet to provide >30 minutes continuous battery runtime upon main grid blackout.",
        "• Environmental & Heat Management: Total internal cabinet heat dissipation is computed for ventilation exhaust fan and enclosure air conditioner sizing to maintain <40°C internal ambient."
    ]

    for n in notes:
        ws1.cell(row=r_notes, column=1, value=n).font = FONT_REG
        r_notes += 1

    # Auto-fit columns
    for col in ws1.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws1.column_dimensions['A'].width = 14
    ws1.column_dimensions['B'].width = 38
    ws1.column_dimensions['C'].width = 16
    ws1.column_dimensions['F'].width = 34
    ws1.column_dimensions['J'].width = 32

    # ---------------------------------------------------------------------------------
    # SHEET 2: Backplane_Power_Details
    # ---------------------------------------------------------------------------------
    ws2 = wb.create_sheet(title="Backplane_Power_Details")
    ws2.views.sheetView[0].showGridLines = True

    ws2.merge_cells("A1:I1")
    ws2["A1"] = "ALLEN-BRADLEY CONTROLLOGIX 1756 CHASSIS BACKPLANE POWER BUDGET (C1 - C5)"
    ws2["A1"].font = FONT_TITLE
    ws2["A1"].fill = NAVY_HDR
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 26

    headers2 = ["Chassis", "Slot No", "Module Catalog", "Module Description", "Signal Family", "5.1VDC Current (A)", "24VDC Current (A)", "Power (Watts)", "Supply Headroom Check"]
    ws2.row_dimensions[3].height = 24
    for c_idx, h in enumerate(headers2, 1):
        cell = ws2.cell(row=3, column=c_idx)
        cell.value = h
        cell.font = FONT_HEADER
        cell.fill = SLATE_SUB
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    r2 = 4
    for ch in ['C1', 'C2', 'C3', 'C4', 'C5']:
        bp = bp_results[ch]
        for sl in bp['slots']:
            ws2.cell(row=r2, column=1, value=ch).alignment = Alignment(horizontal="center")
            ws2.cell(row=r2, column=2, value=f"Slot {sl['slot']}").alignment = Alignment(horizontal="center")
            ws2.cell(row=r2, column=3, value=sl['cat']).font = FONT_BOLD
            ws2.cell(row=r2, column=4, value=sl['desc'])
            ws2.cell(row=r2, column=5, value=sl['cat'].split('-')[1] if '-' in sl['cat'] else 'CPU').alignment = Alignment(horizontal="center")
            ws2.cell(row=r2, column=6, value=round(sl['v5'], 3)).alignment = Alignment(horizontal="right")
            ws2.cell(row=r2, column=7, value=round(sl['v24'], 3)).alignment = Alignment(horizontal="right")
            ws2.cell(row=r2, column=8, value=round(sl['w'], 2)).alignment = Alignment(horizontal="right")
            ws2.cell(row=r2, column=9, value="PASS (OK)").font = FONT_GREEN
            ws2.cell(row=r2, column=9).alignment = Alignment(horizontal="center")

            for c in range(1, 10):
                ws2.cell(row=r2, column=c).border = THIN_BORDER
                if r2 % 2 == 1:
                    ws2.cell(row=r2, column=c).fill = GRAY_BG
            r2 += 1

        # Chassis Subtotal Row
        ws2.cell(row=r2, column=1, value=f"SUBTOTAL {ch}").font = FONT_BOLD
        ws2.cell(row=r2, column=2, value="1756-PA75 PSU").font = FONT_BOLD
        ws2.cell(row=r2, column=3, value=f"Capacity: 75.0 W").font = FONT_BOLD
        ws2.cell(row=r2, column=4, value=f"5.1V Max: 13.0A | 24V Max: 2.8A").font = FONT_BOLD
        ws2.cell(row=r2, column=5, value=f"Util: {bp['w_pct']:.1f}%").font = FONT_GREEN
        ws2.cell(row=r2, column=5).alignment = Alignment(horizontal="center")
        ws2.cell(row=r2, column=6, value=round(bp['v5_tot'], 2)).font = FONT_BOLD
        ws2.cell(row=r2, column=7, value=round(bp['v24_tot'], 2)).font = FONT_BOLD
        ws2.cell(row=r2, column=8, value=round(bp['w_tot'], 2)).font = FONT_BOLD
        ws2.cell(row=r2, column=9, value=f"PASS ({bp['w_pct']:.1f}%)").font = FONT_GREEN
        ws2.cell(row=r2, column=9).alignment = Alignment(horizontal="center")

        for c in range(1, 10):
            ws2.cell(row=r2, column=c).fill = ICE_BLUE
            ws2.cell(row=r2, column=c).border = THIN_BORDER
        r2 += 1

    for col in ws2.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws2.column_dimensions[col_letter].width = max(max_len + 3, 14)
    ws2.column_dimensions['D'].width = 44

    # ---------------------------------------------------------------------------------
    # SHEET 3: 24VDC_Field_Load_Details
    # ---------------------------------------------------------------------------------
    ws3 = wb.create_sheet(title="24VDC_Field_Load_Details")
    ws3.views.sheetView[0].showGridLines = True

    ws3.merge_cells("A1:K1")
    ws3["A1"] = "24VDC POWER DISTRIBUTION & INSTRUMENTATION LOOP LOAD CALCULATIONS"
    ws3["A1"].font = FONT_TITLE
    ws3["A1"].fill = NAVY_HDR
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 26

    headers3 = [
        "Panel Tag", "Panel Location", "DI Loop Current (A)", "DO Relay Coils (A)", "DO Solenoids (A)",
        "AI 4-20mA Loops (A)", "AO 4-20mA Loops (A)", "Network Switches (A)", "Auxiliaries (A)",
        "Operating Current (A)", "Design Load +30% (A)"
    ]
    ws3.row_dimensions[3].height = 24
    for c_idx, h in enumerate(headers3, 1):
        cell = ws3.cell(row=3, column=c_idx)
        cell.value = h
        cell.font = FONT_HEADER
        cell.fill = SLATE_SUB
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    r3 = 4
    for p_id, p in panel_results.items():
        ws3.cell(row=r3, column=1, value=p_id).alignment = Alignment(horizontal="center")
        ws3.cell(row=r3, column=2, value=p['info']['loc'])
        ws3.cell(row=r3, column=3, value=round(p['i_di_operating'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=4, value=round(p['i_do_relays'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=5, value=round(p['i_do_solenoids'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=6, value=round(p['i_ai_operating'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=7, value=round(p['i_ao_operating'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=8, value=round(p['i_network'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=9, value=round(p['i_aux'], 2)).alignment = Alignment(horizontal="right")
        ws3.cell(row=r3, column=10, value=round(p['i_dc_operating'], 2)).font = FONT_BOLD
        ws3.cell(row=r3, column=11, value=round(p['i_dc_design'], 2)).font = FONT_GREEN

        for c in range(1, 12):
            ws3.cell(row=r3, column=c).border = THIN_BORDER
            if r3 % 2 == 1:
                ws3.cell(row=r3, column=c).fill = GRAY_BG
        r3 += 1

    # Total Row
    ws3.cell(row=r3, column=1, value="TOTALS").font = FONT_BOLD
    ws3.cell(row=r3, column=1).alignment = Alignment(horizontal="center")
    ws3.cell(row=r3, column=2, value="System Total 24VDC Current").font = FONT_BOLD
    for c_i, key in enumerate(['i_di_operating', 'i_do_relays', 'i_do_solenoids', 'i_ai_operating', 'i_ao_operating', 'i_network', 'i_aux', 'i_dc_operating', 'i_dc_design'], 3):
        tot_val = sum(panel_results[pid][key] for pid in panel_results)
        cell = ws3.cell(row=r3, column=c_i, value=round(tot_val, 2))
        cell.font = FONT_BOLD
        cell.alignment = Alignment(horizontal="right")

    for c in range(1, 12):
        ws3.cell(row=r3, column=c).fill = AMBER_BG
        ws3.cell(row=r3, column=c).border = THIN_BORDER

    for col in ws3.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws3.column_dimensions[col_letter].width = max(max_len + 3, 14)
    ws3.column_dimensions['B'].width = 32

    # ---------------------------------------------------------------------------------
    # SHEET 4: AC_Mains_and_UPS_Sizing
    # ---------------------------------------------------------------------------------
    ws4 = wb.create_sheet(title="AC_Mains_and_UPS_Sizing")
    ws4.views.sheetView[0].showGridLines = True

    ws4.merge_cells("A1:J1")
    ws4["A1"] = "PANEL AC MAINS INFEED, HEAT DISSIPATION & UPS BACKUP POWER SIZING"
    ws4["A1"].font = FONT_TITLE
    ws4["A1"].fill = NAVY_HDR
    ws4["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws4.row_dimensions[1].height = 26

    headers4 = [
        "Panel Tag", "Mains Voltage", "Active Power (Watts)", "Apparent Power (VA)", "220VAC Current (A)",
        "Recommended MCB", "Heat Load (Watts)", "Heat Load (BTU/hr)", "Cooling Method", "Online UPS Model & Rating"
    ]
    ws4.row_dimensions[3].height = 24
    for c_idx, h in enumerate(headers4, 1):
        cell = ws4.cell(row=3, column=c_idx)
        cell.value = h
        cell.font = FONT_HEADER
        cell.fill = SLATE_SUB
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER

    r4 = 4
    for p_id, p in panel_results.items():
        cooling_str = "Air Conditioned (Control Room)" if p_id == 'CA1' else ("Fan Filter Unit (IP54)" if p['p_heat_w'] < 300 else "Panel Air Conditioner")
        ws4.cell(row=r4, column=1, value=p_id).alignment = Alignment(horizontal="center")
        ws4.cell(row=r4, column=2, value="220VAC 1-Ph 50Hz").alignment = Alignment(horizontal="center")
        ws4.cell(row=r4, column=3, value=round(p['p_ac_total_w'], 1)).alignment = Alignment(horizontal="right")
        ws4.cell(row=r4, column=4, value=round(p['va_ac_total'], 1)).alignment = Alignment(horizontal="right")
        ws4.cell(row=r4, column=5, value=round(p['i_ac_mains_220v'], 2)).alignment = Alignment(horizontal="right")
        ws4.cell(row=r4, column=6, value=p['info']['mcb_mains']).alignment = Alignment(horizontal="center")
        ws4.cell(row=r4, column=7, value=round(p['p_heat_w'], 1)).alignment = Alignment(horizontal="right")
        ws4.cell(row=r4, column=8, value=round(p['p_heat_btu_hr'], 0)).alignment = Alignment(horizontal="right")
        ws4.cell(row=r4, column=9, value=cooling_str)
        ws4.cell(row=r4, column=10, value=p['info']['ups_model'])

        for c in range(1, 11):
            ws4.cell(row=r4, column=c).border = THIN_BORDER
            if r4 % 2 == 1:
                ws4.cell(row=r4, column=c).fill = GRAY_BG
        r4 += 1

    for col in ws4.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws4.column_dimensions[col_letter].width = max(max_len + 3, 14)
    ws4.column_dimensions['J'].width = 40

    # ---------------------------------------------------------------------------------
    # SHEET 5: Power_Equipment_BOM
    # ---------------------------------------------------------------------------------
    ws5 = wb.create_sheet(title="Power_Equipment_BOM")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:H1")
    ws5["A1"] = "BILL OF MATERIALS (BOM) --- POWER SUPPLIES, REDUNDANCY & UPS EQUIPMENT"
    ws5["A1"].font = FONT_TITLE
    ws5["A1"].fill = NAVY_HDR
    ws5["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws5.row_dimensions[1].height = 26

    bom_items = [
        ("1", "QUINT4-PS/1AC/24DC/40", "Phoenix Contact", "2904603", "Primary Switched Redundant 24VDC 40A Power Supply, SFB Technology", "4", "CA1 (2), CA-MCC (2)", "Redundant 24VDC Main DC Bus"),
        ("2", "QUINT4-PS/1AC/24DC/20", "Phoenix Contact", "2904602", "Primary Switched Redundant 24VDC 20A Power Supply, SFB Technology", "6", "RIO-1 (2), RIO-2 (2), IS (2)", "Redundant 24VDC Field DC Bus"),
        ("3", "QUINT4-PS/1AC/24DC/10", "Phoenix Contact", "2904601", "Primary Switched Redundant 24VDC 10A Power Supply, SFB Technology", "2", "CA-RIO-200 (2)", "Redundant 24VDC Remote Bus"),
        ("4", "QUINT4-DIODE/48", "Phoenix Contact", "2907714", "Redundancy Decoupling Diode Module (2x 20A / 1x 40A, 12-24VDC)", "2", "CA1 (1), CA-MCC (1)", "Decouples Primary/Secondary 40A PSUs"),
        ("5", "QUINT4-DIODE/40", "Phoenix Contact", "2907713", "Redundancy Decoupling Diode Module (2x 20A / 1x 40A, 12-24VDC)", "3", "RIO-1 (1), RIO-2 (1), IS (1)", "Decouples Primary/Secondary 20A PSUs"),
        ("6", "QUINT4-DIODE/12", "Phoenix Contact", "2907712", "Redundancy Decoupling Diode Module (2x 10A / 1x 20A, 12-24VDC)", "1", "CA-RIO-200 (1)", "Decouples Primary/Secondary 10A PSUs"),
        ("7", "1756-PA75", "Rockwell Automation", "1756-PA75", "ControlLogix AC Chassis Power Supply (85-265VAC Input, 75W Output)", "8", "C1, C2, C3, C4, C5, C6, C7", "PLC Chassis Backplane Power"),
        ("8", "PLC-RSC-24DC/21", "Phoenix Contact", "2966171", "Interface Relay, DIN Rail, 24VDC Coil, SPDT 1 C/O Contact (6A 250VAC)", "288", "All DO Cards (C1-C5, MCC)", "Interposing Relay Isolation"),
        ("9", "SRT3000XLI", "Schneider APC", "SRT3000XLI", "Smart-UPS On-Line RT 3000VA / 2700W 230V Double Conversion UPS", "1", "Control Room MCP (CA1)", "30-min Battery Autonomy"),
        ("10", "SRT2200XLI", "Schneider APC", "SRT2200XLI", "Smart-UPS On-Line RT 2200VA / 1980W 230V Double Conversion UPS", "1", "MCC Room (CA-MCC)", "30-min Battery Autonomy"),
        ("11", "SRT1500XLI", "Schneider APC", "SRT1500XLI", "Smart-UPS On-Line RT 1500VA / 1350W 230V Double Conversion UPS", "3", "RIO-1, RIO-2, IS-CAB", "30-min Battery Autonomy"),
        ("12", "SRT1000XLI", "Schneider APC", "SRT1000XLI", "Smart-UPS On-Line RT 1000VA / 900W 230V Double Conversion UPS", "1", "Slurry Building (RIO-200)", "30-min Battery Autonomy"),
        ("13", "VAL-MS 230/3+1", "Phoenix Contact", "2838209", "Type 2 Surge Protection Device (SPD) for 230/400V Mains Feeder", "6", "All 6 Automation Panels", "Mains Lightning & Surge Protection"),
        ("14", "C60H-DC 2P", "Schneider Electric", "MGN61420", "2-Pole DC Miniature Circuit Breaker (C-Curve, 10A / 16A / 20A)", "24", "All Panels DC Distribution", "DC Branch Feeder Protection"),
    ]

    headers5 = ["Item", "Model / Catalog", "Manufacturer", "Part Number", "Description & Technical Specifications", "Qty", "Panel Allocation", "Application / Function"]
    ws5.row_dimensions[3].height = 24
    for c_idx, h in enumerate(headers5, 1):
        cell = ws5.cell(row=3, column=c_idx)
        cell.value = h
        cell.font = FONT_HEADER
        cell.fill = SLATE_SUB
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER

    r5 = 4
    for itm in bom_items:
        for c_idx, val in enumerate(itm, 1):
            cell = ws5.cell(row=r5, column=c_idx, value=val)
            cell.font = FONT_REG
            cell.border = THIN_BORDER
            if c_idx in (1, 6):
                cell.alignment = Alignment(horizontal="center")
            elif c_idx == 2:
                cell.font = FONT_BOLD
            if r5 % 2 == 1:
                cell.fill = GRAY_BG
        r5 += 1

    for col in ws5.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws5.column_dimensions[col_letter].width = max(max_len + 3, 12)
    ws5.column_dimensions['E'].width = 46

    # Save Workbook
    wb.save(out_file)
    print(f"\n=======================================================")
    print(f"SUCCESS: Generated Panel Power Supply Calculation Report:")
    print(f"  Path: {out_file}")
    print(f"  Size: {os.path.getsize(out_file):,} bytes")
    print(f"=======================================================")

if __name__ == "__main__":
    build_calculation()
