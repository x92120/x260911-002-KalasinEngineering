#!/usr/bin/env python3
"""
=============================================================================
KALASIN ENGINEERING xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)
PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER - INGREDION (THAILAND) CO., LTD.
MASTER I/O MAPPING REPORT GENERATOR (EXCEL WORKBOOK)
=============================================================================
Input Files (Folder: 03-IO_List):
  1. Instrument I-O List Rev.3.6a.xlsx (Sheet: 'Rev.3')
  2. IO_List-By_SlotConfig_rev02.xlsx (Sheets: '00_Slot_Index', 'C1SlotConfig'..'C5SlotConfig', 'C1S0'..'C5S4')

Output File:
  03-IO_List/IO_Mapping_Report.xlsx (and Master_IO_Mapping_Report.xlsx)

Workbook Structure:
  1. 00_Executive_Summary        - High-level KPI dashboard, chassis utilization, signal breakdown
  2. 01_Instrument_IO_Mapping    - Master instrument-centric view (Rev.3.6a mapped to PLC hardware)
  3. 02_PLC_Channel_IO_Mapping   - Complete point-by-point hardware schedule (1,466 channels)
  4. 03_Chassis_Slot_Directory   - All 5 chassis and 57 slots hardware overview
  5. 04_Destination_JB_Matrix    - Field junction box / enclosure I/O allocation matrix
  6. 05_Instrument_Catalog_Ref   - Master catalog reference (manufacturers, models, cables)
=============================================================================
"""

import os
import re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
IO_DIR = os.path.join(BASE_DIR, "03-IO_List")
INST_FILE = os.path.join(IO_DIR, "Instrument I-O List Rev.3.6a.xlsx")
SLOT_FILE = os.path.join(IO_DIR, "IO_List-By_SlotConfig_rev02.xlsx")
OUTPUT_FILE = os.path.join(IO_DIR, "IO_Mapping_Report.xlsx")
OUTPUT_FILE_ALT = os.path.join(IO_DIR, "Master_IO_Mapping_Report.xlsx")

# -----------------------------------------------------------------------------
# Color Palette & Styles
# -----------------------------------------------------------------------------
NAVY_HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
SUB_HEADER_FILL = PatternFill(start_color="2D4A77", end_color="2D4A77", fill_type="solid")
SECTION_HEADER_FILL = PatternFill(start_color="334E68", end_color="334E68", fill_type="solid")
CARD_HEADER_FILL = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
CARD_BG_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

STATUS_ACTIVE_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
STATUS_ACTIVE_FONT = Font(name="Calibri", size=10, bold=True, color="166534")

STATUS_SPARE_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
STATUS_SPARE_FONT = Font(name="Calibri", size=10, bold=True, color="92400E")

STATUS_COMMON_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
STATUS_COMMON_FONT = Font(name="Calibri", size=10, italic=True, color="475569")

STATUS_RIO_FILL = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
STATUS_RIO_FONT = Font(name="Calibri", size=10, bold=True, color="0369A1")

STATUS_PKG_FILL = PatternFill(start_color="EDE9FE", end_color="EDE9FE", fill_type="solid")
STATUS_PKG_FONT = Font(name="Calibri", size=10, bold=True, color="5B21B6")

STATUS_LOCAL_FILL = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
STATUS_LOCAL_FONT = Font(name="Calibri", size=10, italic=True, color="374151")

FONT_TITLE = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
FONT_SUBTITLE = Font(name="Calibri", size=10, italic=True, color="FFFFFF")
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_CARD_TITLE = Font(name="Calibri", size=9, bold=True, color="94A3B8")
FONT_CARD_VAL = Font(name="Calibri", size=16, bold=True, color="1E293B")
FONT_CARD_SUB = Font(name="Calibri", size=8, italic=True, color="64748B")

FONT_DATA = Font(name="Calibri", size=10, color="0F172A")
FONT_DATA_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")
FONT_DATA_CODE = Font(name="Consolas", size=9.5, color="0F172A")
FONT_DATA_MUTED = Font(name="Calibri", size=9.5, italic=True, color="64748B")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
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

def infer_instrument_model(tag, desc, inst_name, signal_type, prot_type):
    t_upper = str(tag).strip().upper()
    d_lower = str(desc).strip().lower()
    in_lower = str(inst_name).strip().lower()
    p_upper = str(prot_type).strip().upper()
    prefix = t_upper.split('-')[0] if '-' in t_upper else re.split(r'(\d+)', t_upper)[0]

    if prefix == "FT" and ("coriolis" in d_lower or "mass" in d_lower or "promass" in d_lower or "slurry" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Promass F 300 / 83F (Coriolis Mass Flow)", "order_code": "83F50 / 83F80 Series",
                "conn": "Flange ASME B16.5 Cl.150 RF", "power": "4-20mA HART / 24VDC", "ex_class": "Ex ia / Safe Area"}
    if prefix == "FT" and ("magnetic" in d_lower or "mag" in d_lower or "promag" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Promag W 400 / 50W (Electromagnetic Flow)", "order_code": "50W Series",
                "conn": "Flange ASME B16.5 Cl.150", "power": "4-20mA HART / 24VDC", "ex_class": "Safe Area"}
    if prefix == "FT" and ("thermal" in d_lower or "gas" in d_lower or "air" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "t-mass F 300 / I 300 (Thermal Mass Flow)", "order_code": "6F3B / 6I3B Series",
                "conn": "Flange Cl.150 or 1\" NPT", "power": "4-20mA HART / 24VDC", "ex_class": "Safe Area"}
    if prefix in ["FT", "FI", "FE"]:
        return {"mfg": "Endress+Hauser", "model": "Proline Series Flowmeter", "order_code": "Standard Flow Package",
                "conn": "Flange Cl.150 RF", "power": "4-20mA HART / 24VDC", "ex_class": "Safe Area"}
    if prefix in ["DPT", "DPIT"]:
        return {"mfg": "Endress+Hauser", "model": "Deltabar PMD55B (Diff. Pressure)", "order_code": "PMD55B w/ 5-Valve Manifold DA63M",
                "conn": "1/4\" NPT Female w/ 5-Valve Manifold", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex ta/tb IIIC / Safe"}
    if prefix in ["PT", "PIT", "PI"]:
        return {"mfg": "Endress+Hauser", "model": "Cerabar PMP51B / PMP43", "order_code": "PMP51B / PMP43 Hygienic",
                "conn": "1/2\" NPT Male / Tri-Clamp", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex d IIC / Safe Area"}
    if prefix in ["PS", "PSL", "PSH"]:
        return {"mfg": "Endress+Hauser", "model": "Ceraphant PTC31B (Pressure Switch)", "order_code": "PTC31B Series",
                "conn": "1/2\" NPT Male", "power": "24VDC / Dry Contact Relay / PNP", "ex_class": "Safe Area"}
    if prefix in ["DPG", "DPI"]:
        return {"mfg": "Ashcroft", "model": "Model 1132 Diff. Pressure Gauge", "order_code": "1132 w/ V03 Manifold",
                "conn": "1/4\" NPT Female", "power": "Local Mechanical Dial", "ex_class": "Safe Area"}
    if prefix in ["PG"]:
        return {"mfg": "Ashcroft", "model": "Model T5500 All Stainless Pressure Gauge", "order_code": "T5500 w/ Siphon",
                "conn": "1/2\" NPT Male", "power": "Local Mechanical Dial", "ex_class": "Safe Area"}
    if prefix in ["LT", "LIT"]:
        if "radar" in d_lower or "silo" in d_lower or "dryer" in d_lower:
            return {"mfg": "Endress+Hauser", "model": "Micropilot FMR67B (80GHz Radar Level)", "order_code": "FMR67B Series",
                    "conn": "Flange DN80 3\" / Cl.150", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex ia / Ex ta/tb IIIC"}
        return {"mfg": "Endress+Hauser", "model": "Cerabar PMP43 / PMP51B (Hydrostatic)", "order_code": "PMP43 Flush Diaphragm",
                "conn": "Universal Clamp / Flush Mount", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Safe Area"}
    if prefix in ["LS", "LSL", "LSH"]:
        if any(w in d_lower for w in ["powder", "solid", "dryer", "cyclone", "silo"]):
            return {"mfg": "Endress+Hauser", "model": "Soliphant FTM51 (Vibrating Fork Solids)", "order_code": "FTM51 Series",
                    "conn": "Thread 1-1/2\" NPT / Flange", "power": "24VDC / Relay Contact Output", "ex_class": "Ex ta/tb IIIC / Safe"}
        return {"mfg": "Endress+Hauser", "model": "Liquiphant FTL51B (Vibrating Fork Liquids)", "order_code": "FTL51B Series",
                "conn": "Thread 3/4\" NPT / Tri-Clamp", "power": "24VDC / Relay Contact Output", "ex_class": "Ex ia / Safe Area"}
    if prefix in ["TT", "TIT"]:
        return {"mfg": "Endress+Hauser", "model": "iTEMP TMT71 / TM411 (Pt100 RTD)", "order_code": "TMT71 + TM411 Head Assembly",
                "conn": "Threaded 1/2\" NPT w/ Thermowell", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex ia / Safe Area"}
    if prefix in ["TG", "TI"]:
        return {"mfg": "Ashcroft", "model": "Model FI Bimetal Dial Thermometer", "order_code": "FI Series w/ Thermowell",
                "conn": "1/2\" NPT Threaded Thermowell", "power": "Local Mechanical Dial", "ex_class": "Safe Area"}
    if prefix in ["TS", "TSH", "TSL"]:
        return {"mfg": "Endress+Hauser", "model": "Thermophant TTR35 (Temperature Switch)", "order_code": "TTR35 Series",
                "conn": "1/2\" NPT Hygienic Adapter", "power": "24VDC / Relay / PNP", "ex_class": "Safe Area"}
    if prefix in ["AT", "AI", "PHT"]:
        if "ph" in d_lower:
            return {"mfg": "Mettler Toledo", "model": "M300G2 + InPro3250i (pH System)", "order_code": "M300G2 4-Wire + InTrac 777P",
                    "conn": "DN25 Weld-in Socket w/ InTrac 777P", "power": "4-20mA HART + Relay / 24VDC", "ex_class": "Safe Area"}
        elif any(w in d_lower for w in ["gas", "flame", "lel"]):
            return {"mfg": "Sensidyne / SmartGas", "model": "925FGD Explosionproof Gas Detector", "order_code": "925FGD Series",
                    "conn": "3/4\" NPT Conduit Entry", "power": "4-20mA + Relays / 24VDC", "ex_class": "ATEX/IECEx Ex d IIC"}
        return {"mfg": "ENVEA / PCME", "model": "PCME QAL 991 (Dust Analyzer)", "order_code": "6626013-33-201-001",
                "conn": "Flange DN100 w/ Purge Air", "power": "4-20mA Output + Relay / 24VDC", "ex_class": "Safe Area"}
    if prefix in ["TCV", "PCV", "FCV", "CV", "TV"]:
        return {"mfg": "Samson", "model": "Type 3241 Globe Valve + Trovis 3730-1", "order_code": "3241-ANSI + 3730-1 Positioner",
                "conn": "Flange ASME B16.5 Cl.150/300 RF", "power": "4-20mA Setpoint + 4-20mA Feedback", "ex_class": "Ex ia IIC / Safe Area"}
    if prefix in ["XV", "SV", "BV", "ISV"]:
        return {"mfg": "TeroFox / El-O-Matic", "model": "TF-20DFS Ball Valve + F-Actuator + ESV", "order_code": "TF-20DFS + EL-O-Matic F + ESV",
                "conn": "Flange Cl.150 RF / 1/4\" Air Port", "power": "24VDC Solenoid (DO) + Dry Contact LS (DI)", "ex_class": "Ex d IIC / Safe Area"}
    if prefix in ["HV"]:
        return {"mfg": "TeroFox / APL-HKC", "model": "TF-20DFS Manual Ball Valve + APL-210N Box", "order_code": "TF-20DFS + APL-210N Switch Box",
                "conn": "Flange Cl.150 RF / Manual Lever", "power": "Dry Contact Limit Switches (2x DI)", "ex_class": "Ex d IIC / Safe Area"}
    if prefix in ["ZS"]:
        return {"mfg": "APL-HKC", "model": "APL-210N / APL-510N Valve Monitor", "order_code": "APL-210N / APL-510N",
                "conn": "NAMUR VDI/VDE 3845 Direct Mount", "power": "Dry Contact Micro-switches (2x SPDT)", "ex_class": "Ex d IIC / Safe Area"}
    if prefix in ["PSE"]:
        return {"mfg": "BS&B / Fike", "model": "Burst-Alert Sensor / Rupture Disk Monitor", "order_code": "Burst Disk Sensor",
                "conn": "Cl.150 Flange Holder", "power": "Dry Contact NC Loop (DI)", "ex_class": "Ex ia IIC"}
    if prefix in ["VIB", "NCT"]:
        return {"mfg": "Netter Vibration", "model": "NCT 5 Pneumatic Turbine Vibrator", "order_code": "NCT 5 Series",
                "conn": "G 1/8\" BSP Pneumatic Port", "power": "Pneumatic 2 - 6 bar / 24VDC Solenoid", "ex_class": "ATEX Zone 1/21 Safe"}
    if prefix in ["ATY", "PCY", "BFY", "RVM", "PCM", "BLM", "P", "M"]:
        return {"mfg": "Rockwell Automation", "model": "PowerFlex VSD / E300 Relay in MCC", "order_code": "Centerline 2500 Draw-out Bucket",
                "conn": "Terminal Block in MCC Compartment", "power": "24VDC Interlock / 380VAC 3-Phase", "ex_class": "MCC Room"}
    return {"mfg": "Industrial Process Equipment", "model": "Standard Process Field Device", "order_code": "-",
            "conn": "Standard Process Connection", "power": "24VDC / Dry Contact", "ex_class": "Safe Area"}

def suggest_cable(prefix, io_type, signal_type, cable_existing, prot_type):
    c_exist = str(cable_existing).strip() if cable_existing and str(cable_existing).strip() not in ["-", ""] else ""
    p_upper = str(prot_type).upper()
    is_is = "EX I" in p_upper or "EX IA" in p_upper or "EX IC" in p_upper

    if prefix in ["XV", "SV", "BV", "ISV"]:
        return ("CVV 3Cx1.5 (Solenoid) + CVV 4Cx1.0 (Limit Switches)", "3Cx1.5 + 4Cx1.0", "M20x1.5 Ex d Double Compression", "CVV 30Cx1.5 Trunk to CA1")
    elif prefix in ["TCV", "PCV", "FCV", "CV", "TV"]:
        return ("LiY-CY TP 2Px1.0 (Setpoint AO + Feedback AI)", "2Px1.0 Shielded Pairs", "M20x1.5 Brass Nickel-Plated Ex d/e", "LiY-CY(IS/OS) 16Px0.75 Trunk")
    elif prefix in ["AT", "AI", "PHT"]:
        return ("LiY-CY TP 2Px1.0 (Signal) + CVV 3Cx1.5 (24VDC Power)", "2Px1.0 + 3Cx1.5", "M20x1.5 Ex d/e Gland", "LiY-CY 16Px0.75 + CVV Trunk")
    elif prefix == "FT" and ("3p" in c_exist.lower() or "mass" in c_exist.lower() or "cvv 3c" in c_exist.lower()):
        return ("LiY-CY TP 3Px1.0 (Signal/Density) + CVV 3Cx1.5 (Power)", "3Px1.0 + 3Cx1.5", "M20x1.5 Ex d/e Gland", "LiY-CY(IS/OS) 16Px0.75 Trunk")
    elif io_type in ["AI", "AO"] or "4 - 20" in str(signal_type):
        c_spec = "LiY-CY(EB) 1Px1.0 (Blue)" if is_is else "LiY-CY TP 1Px1.0 (Shielded Pair)"
        return (c_spec, "1Px1.0 Shielded Pair", "M20x1.5 Polyamide Ex i / Brass Ex d", "LiY-CY(IS/OS) 16Px0.75 Trunk")
    elif prefix in ["HV", "ZS"]:
        return ("CVV 4Cx1.0 (Open/Close Limit Switches)", "4Cx1.0 Multi-core", "M20x1.5 / 1/2\" NPT Ex d/e", "CVV 30Cx1.5 Trunk to CA1")
    elif prefix in ["LS", "LSH", "LSL", "PS", "PSH", "PSL", "TS", "TSH", "TSL"]:
        return ("CVV 3Cx1.5 (24VDC Power + Switch Output)", "3Cx1.5 Multi-core", "M20x1.5 Ex d/e Gland", "CVV 30Cx1.5 Trunk to CA1")
    elif io_type == "BUS" or "bus" in str(signal_type).lower():
        return ("CAT 6 S/FTP Industrial Ethernet / Belden 3105A", "4 Pairs 24 AWG / 1 Pair 22 AWG", "M20x1.5 Industrial RJ45 Gland", "Fiber Optic / Industrial Ethernet Trunk")
    elif prefix in ["ATY", "PCY", "BFY", "RVM", "PCM", "BLM", "P", "M"]:
        return ("CVV 7Cx1.5 Control Cable (MCC Start/Stop/Fault/Run)", "7Cx1.5 Multi-core", "M25x1.5 Heavy Industrial Gland", "CVV Multi-Core Control Trunk to MCC")
    return ("CVV 2Cx1.5 / LiY-CY 1Px1.0 General Control", "2Cx1.5 or 1Px1.0", "M20x1.5 Industrial Gland", "Standard Trunk Cable Tray")

# -----------------------------------------------------------------------------
# Data Loading & Processing
# -----------------------------------------------------------------------------
def load_all_data():
    print(f"Loading Instrument Master: {INST_FILE}")
    wb_inst = openpyxl.load_workbook(INST_FILE, data_only=True)
    ws_inst = wb_inst['Rev.3']

    inst_list = []
    inst_lookup = {}
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
        bus_sig = ws_inst.cell(r, 11).value
        sig_type = ws_inst.cell(r, 12).value
        sig_to = ws_inst.cell(r, 13).value
        rng = ws_inst.cell(r, 14).value
        fn = ws_inst.cell(r, 15).value
        cable = ws_inst.cell(r, 16).value
        prot = ws_inst.cell(r, 17).value
        remark = ws_inst.cell(r, 18).value

        if tag or new_tag or desc:
            item = {
                'row': r,
                'item_no': str(item_no).strip() if item_no is not None else '',
                'tag': str(tag).strip() if tag else '',
                'new_tag': str(new_tag).strip() if new_tag else '',
                'pid': str(pid).strip() if pid else '',
                'desc': str(desc).strip() if desc else '',
                'inst_name': str(inst_name).strip() if inst_name else '',
                'do': int(do_sig) if isinstance(do_sig, (int, float)) and do_sig > 0 else (1 if str(do_sig).strip() == '1' else 0),
                'di': int(di_sig) if isinstance(di_sig, (int, float)) and di_sig > 0 else (1 if str(di_sig).strip() == '1' else 0),
                'ao': int(ao_sig) if isinstance(ao_sig, (int, float)) and ao_sig > 0 else (1 if str(ao_sig).strip() == '1' else 0),
                'ai': int(ai_sig) if isinstance(ai_sig, (int, float)) and ai_sig > 0 else (1 if str(ai_sig).strip() == '1' else 0),
                'bus': int(bus_sig) if isinstance(bus_sig, (int, float)) and bus_sig > 0 else (1 if str(bus_sig).strip() == '1' else 0),
                'sig_type': str(sig_type).strip() if sig_type else '',
                'sig_to': str(sig_to).strip() if sig_to else '',
                'range': str(rng).strip() if rng else '',
                'function': str(fn).strip() if fn else '',
                'cable': str(cable).strip() if cable else '',
                'prot': str(prot).strip() if prot else '',
                'remark': str(remark).strip() if remark else ''
            }
            # Active working tag: new_tag preferred, fallback to tag
            item['active_tag'] = item['new_tag'] if item['new_tag'] else item['tag']
            inst_list.append(item)

            # Store in lookup by all variations
            for k in [item['active_tag'], item['new_tag'], item['tag']]:
                if k:
                    u = k.upper()
                    inst_lookup[u] = item
                    inst_lookup[u.replace(' ', '').replace('-', '')] = item

    print(f"Loaded {len(inst_list)} instruments from Rev.3.6a.")

    print(f"Loading Slot Configuration: {SLOT_FILE}")
    wb_slots = openpyxl.load_workbook(SLOT_FILE, data_only=True)
    ws_idx = wb_slots['00_Slot_Index']

    # Read chassis overview
    chassis_info = []
    for r in range(6, 11):
        c_code = ws_idx.cell(r, 1).value
        c_sheet = ws_idx.cell(r, 2).value
        c_desc = ws_idx.cell(r, 3).value
        c_loc = ws_idx.cell(r, 4).value
        c_model = ws_idx.cell(r, 5).value
        c_slots = ws_idx.cell(r, 6).value
        c_tot = ws_idx.cell(r, 7).value
        c_act = ws_idx.cell(r, 8).value
        c_spr = ws_idx.cell(r, 9).value
        if c_code:
            chassis_info.append({
                'code': str(c_code).strip(),
                'sheet': str(c_sheet).strip(),
                'desc': str(c_desc).strip(),
                'location': str(c_loc).strip(),
                'model': str(c_model).strip(),
                'slots_mounted': str(c_slots).strip(),
                'total_io': c_tot,
                'active_io': c_act,
                'spare_io': c_spr
            })

    # Read slot schedule directory
    slot_dir = []
    slot_hw_lookup = {}
    for r in range(14, ws_idx.max_row + 1):
        s_tag = ws_idx.cell(r, 1).value
        s_chassis = ws_idx.cell(r, 2).value
        s_slot = ws_idx.cell(r, 3).value
        s_card = ws_idx.cell(r, 4).value
        s_family = ws_idx.cell(r, 5).value
        s_terms = ws_idx.cell(r, 6).value
        s_dest = ws_idx.cell(r, 7).value
        s_act = ws_idx.cell(r, 8).value
        s_spr = ws_idx.cell(r, 9).value
        if s_tag:
            item = {
                'slot_tag': str(s_tag).strip(),
                'chassis': str(s_chassis).strip(),
                'slot_no': s_slot,
                'card': str(s_card).strip(),
                'family': str(s_family).strip(),
                'terminals': s_terms,
                'destination': str(s_dest).strip(),
                'active': s_act,
                'spare': s_spr
            }
            slot_dir.append(item)
            slot_hw_lookup[item['slot_tag']] = item

    # Read all slot channel details
    slot_sheets = [s for s in wb_slots.sheetnames if s not in ['00_Slot_Index', 'C1SlotConfig', 'C2SlotConfig', 'C3SlotConfig', 'C4SlotConfig', 'C5SlotConfig']]
    channels = []
    tag_to_channels = defaultdict(list)

    seq = 0
    for sname in slot_sheets:
        ws = wb_slots[sname]
        chassis = sname[:2]
        slot_info = slot_hw_lookup.get(sname, {})
        card_model = slot_info.get('card', '')
        family = slot_info.get('family', '')

        for r in range(6, ws.max_row + 1):
            t_no = ws.cell(r, 1).value
            if t_no is None: continue
            seq += 1
            t_desc = ws.cell(r, 2).value
            t_num = ws.cell(r, 3).value
            w_plc = ws.cell(r, 4).value
            w_term = ws.cell(r, 5).value
            dest = ws.cell(r, 6).value
            p_tag = ws.cell(r, 7).value
            i_tag = ws.cell(r, 8).value
            i_desc = ws.cell(r, 9).value
            status = ws.cell(r, 10).value

            ch_item = {
                'seq': seq,
                'sheet': sname,
                'chassis': chassis,
                'slot_tag': sname,
                'card_model': card_model,
                'family': family,
                'term_no': str(t_no).strip(),
                'term_desc': str(t_desc).strip() if t_desc else '',
                'term_num': str(t_num).strip() if t_num else '',
                'wire_plc': str(w_plc).strip() if w_plc else '',
                'wire_term': str(w_term).strip() if w_term else '',
                'dest_area': str(dest).strip() if dest else '',
                'plc_tag': str(p_tag).strip() if p_tag else '',
                'inst_tag': str(i_tag).strip() if i_tag else '',
                'inst_desc': str(i_desc).strip() if i_desc else '',
                'status': str(status).strip().upper() if status else 'ACTIVE'
            }
            channels.append(ch_item)

            if ch_item['inst_tag'] and ch_item['status'] == 'ACTIVE':
                itag_u = ch_item['inst_tag'].upper()
                tag_to_channels[itag_u].append(ch_item)
                norm = itag_u.replace(' ', '').replace('-', '')
                if norm != itag_u:
                    tag_to_channels[norm].append(ch_item)

    print(f"Loaded {len(channels)} physical channels across {len(slot_sheets)} slot sheets.")
    return inst_list, inst_lookup, chassis_info, slot_dir, slot_hw_lookup, channels, tag_to_channels

# -----------------------------------------------------------------------------
# Sheet 1: 00_Executive_Summary
# -----------------------------------------------------------------------------
def build_executive_summary_sheet(ws, inst_list, chassis_info, slot_dir, channels):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:M1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)",
             font=FONT_TITLE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:M2")
    set_cell(ws["A2"], "MASTER PLC I/O MAPPING & HARDWARE ALLOCATION REPORT — SPRINT 18K TPA SPRAY DRYER / JET COOKER",
             font=FONT_SUBTITLE, fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # KPI Summary Cards (Row 4 to 6)
    kpis = [
        ("A4:B4", "A5:B5", "A6:B6", "A", "TOTAL INSTRUMENTS", str(len(inst_list)), "Rev 3.6a Master Client Schedule"),
        ("C4:D4", "C5:D5", "C6:D6", "C", "PHYSICAL PLC CHANNELS", f"{len(channels):,}", "5 Chassis (C1-C5), 57 Slots"),
        ("E4:F4", "E5:F5", "E6:F6", "E", "ACTIVE PROCESS I/O", f"{sum(1 for c in channels if c['status'] == 'ACTIVE'):,}", "Assigned to Process Instruments"),
        ("G4:H4", "G5:H5", "G6:H6", "G", "SPARE CAPACITY", f"{sum(1 for c in channels if c['status'] == 'SPARE'):,}", "38.3% Expansion Margin (>20% Req)"),
        ("I4:J4", "I5:J5", "I6:J6", "I", "COMMON / INTERNAL", f"{sum(1 for c in channels if c['status'] not in ['ACTIVE', 'SPARE']):,}", "Power Commons, CPU & Comm Ports"),
        ("K4:M4", "K5:M5", "K6:M6", "K", "NETWORK REDUNDANCY", "1Gbps DLR RING", "EtherNet/IP Device Level Ring CA1-RIO")
    ]

    for m_head, m_val, m_sub, col_start, title, val, sub in kpis:
        ws.merge_cells(m_head)
        ws.merge_cells(m_val)
        ws.merge_cells(m_sub)
        c_head = ws[m_head.split(':')[0]]
        c_val = ws[m_val.split(':')[0]]
        c_sub = ws[m_sub.split(':')[0]]
        set_cell(c_head, title, font=FONT_CARD_TITLE, fill=CARD_HEADER_FILL, alignment=ALIGN_CENTER)
        set_cell(c_val, val, font=FONT_CARD_VAL, fill=CARD_BG_FILL, alignment=ALIGN_CENTER)
        set_cell(c_sub, sub, font=FONT_CARD_SUB, fill=CARD_BG_FILL, alignment=ALIGN_CENTER)

    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 16

    # Section 1: Chassis Overview Table (Row 8)
    ws.merge_cells("A8:M8")
    set_cell(ws["A8"], "1. CONTROLLOGIX 1756 CHASSIS HARDWARE & UTILIZATION DIRECTORY",
             font=FONT_SECTION, fill=SECTION_HEADER_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[8].height = 22

    ch_headers = [
        "Chassis Code", "Overview Sheet", "Chassis & Function Description", "Physical Location",
        "Hardware Model", "Mounted Slots", "Total Channels", "Active Channels", "Spare Channels",
        "Spare Margin %", "Main Power Supply", "Network Link", "Cabinet / Panel"
    ]
    ws.row_dimensions[9].height = 24
    for col_idx, h in enumerate(ch_headers, 1):
        set_cell(ws.cell(9, col_idx), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    # Populate Chassis Rows
    r_idx = 10
    total_tot, total_act, total_spr = 0, 0, 0
    ps_models = {
        'C1': '1756-PA75 (85-265VAC)',
        'C2': '1756-PA75 (85-265VAC)',
        'C3': '1756-PA75 (85-265VAC)',
        'C4': '1756-PA75 (85-265VAC)',
        'C5': '1756-PB75 (24VDC Redundant)'
    }
    net_links = {
        'C1': '1Gbps DLR Ring + SCADA Uplink',
        'C2': 'Internal Backplane Bridge',
        'C3': '1Gbps DLR Fiber Ring Ring Port 1/2',
        'C4': '1Gbps DLR Copper Ring Port 1/2',
        'C5': '1Gbps DLR Copper Ring (Slurry Bldg)'
    }
    cab_panels = {
        'C1': 'CA1 (Main Control Room MCP)',
        'C2': 'CA1 (Main Control Room MCP)',
        'C3': 'Field RIO Cabinet (Spray Dryer)',
        'C4': 'Field RIO Cabinet (6th/8th Fl)',
        'C5': 'Slurry Bldg 2nd Fl (RIO-200)'
    }

    # Recalculate exact channel metrics from loaded channels
    ch_metrics = defaultdict(lambda: {'total': 0, 'active': 0, 'spare': 0})
    for c in channels:
        ch_metrics[c['chassis']]['total'] += 1
        if c['status'] == 'ACTIVE':
            ch_metrics[c['chassis']]['active'] += 1
        elif c['status'] == 'SPARE':
            ch_metrics[c['chassis']]['spare'] += 1

    for ch in chassis_info:
        code = ch['code']
        m = ch_metrics[code]
        tot = m['total']
        act = m['active']
        spr = m['spare']
        spare_pct = (spr / tot * 100.0) if tot > 0 else 0.0

        total_tot += tot
        total_act += act
        total_spr += spr

        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[r_idx].height = 20

        set_cell(ws.cell(r_idx, 1), code, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 2), ch['sheet'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 3), ch['desc'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), ch['location'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 5), ch['model'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 6), ch['slots_mounted'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 7), tot, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 8), act, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 9), spr, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 10), f"{spare_pct:.1f}%", font=FONT_DATA_BOLD, fill=STATUS_SPARE_FILL if spare_pct >= 20 else fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 11), ps_models.get(code, "-"), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 12), net_links.get(code, "-"), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 13), cab_panels.get(code, "-"), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        r_idx += 1

    # Total Row for Chassis
    ws.row_dimensions[r_idx].height = 22
    ws.merge_cells(f"A{r_idx}:F{r_idx}")
    set_cell(ws.cell(r_idx, 1), "TOTAL SYSTEM CAPACITY (5 CHASSIS)", font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    for c_i in range(2, 7):
        set_cell(ws.cell(r_idx, c_i), "", fill=CARD_BG_FILL, border=TOTAL_BORDER)

    avg_spare_pct = (total_spr / total_tot * 100.0) if total_tot > 0 else 0.0
    set_cell(ws.cell(r_idx, 7), total_tot, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 8), total_act, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 9), total_spr, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 10), f"{avg_spare_pct:.1f}%", font=FONT_DATA_BOLD, fill=STATUS_SPARE_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)
    for c_i in range(11, 14):
        set_cell(ws.cell(r_idx, c_i), "-", font=FONT_DATA, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)

    r_idx += 2

    # Section 2: Signal Family & Module Breakdown (Two tables side by side)
    ws.merge_cells(f"A{r_idx}:F{r_idx}")
    set_cell(ws.cell(r_idx, 1), "2. I/O SIGNAL FAMILY & MODULE SCHEDULE", font=FONT_SECTION, fill=SECTION_HEADER_FILL, alignment=ALIGN_LEFT)

    ws.merge_cells(f"H{r_idx}:M{r_idx}")
    set_cell(ws.cell(r_idx, 8), "3. CLIENT INSTRUMENT SCOPE & DESTINATION SUMMARY", font=FONT_SECTION, fill=SECTION_HEADER_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[r_idx].height = 22

    r_idx += 1
    # Table 2 Header
    sig_headers = ["Signal Family", "Module Hardware", "Signal Type", "Total Pts", "Active Pts", "Spare Pts"]
    for c_i, h in enumerate(sig_headers, 1):
        set_cell(ws.cell(r_idx, c_i), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    # Table 3 Header
    dest_headers = ["Scope Destination", "Primary Wiring Target", "Loop Count", "Active Pts", "Typical Cable", "Interface Remark"]
    for c_i, h in enumerate(dest_headers, 8):
        set_cell(ws.cell(r_idx, c_i), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)
    ws.row_dimensions[r_idx].height = 24

    # Calculate Signal Family counts
    family_stats = defaultdict(lambda: {'total': 0, 'active': 0, 'spare': 0})
    for c in channels:
        f = c['family']
        family_stats[f]['total'] += 1
        if c['status'] == 'ACTIVE': family_stats[f]['active'] += 1
        elif c['status'] == 'SPARE': family_stats[f]['spare'] += 1

    family_rows = [
        ("DI (24VDC Digital Input)", "1756-IB32 (32-Pt Sink/Source)", "24VDC Dry Contact / PNP", family_stats['DI (24VDC Digital Input)']),
        ("DO (24VDC Digital Output)", "1756-OB32 (32-Pt Source Transistor)", "24VDC Energized Solenoids", family_stats['DO (24VDC Digital Output)']),
        ("AI (4-20mA / RTD / Volts)", "1756-IF16 (16-Pt Single-Ended / Diff)", "4-20mA HART Transmitter", family_stats['AI (4-20mA / Voltage / RTD)']),
        ("AO (4-20mA Analog Output)", "1756-OF8 (8-Pt Voltage / Current)", "4-20mA Control Valves (CV)", family_stats['AO (4-20mA Analog Output)']),
        ("CPU (Main Controller)", "1756-L950TPSXT (ControlLogix 5580)", "Process Controller & SD/USB", family_stats['CPU (Controller)']),
        ("COMM (EtherNet/IP)", "1756-EN4TR (1Gbps 2-Port DLR)", "DLR Device Level Ring", family_stats['COMM (EtherNet/IP)']),
        ("EMPTY (Chassis Slot Reserve)", "1756-N2 (Slot Filler Blank Plate)", "Future Hardware Expansion", family_stats['EMPTY (Reserve Position)'])
    ]

    # Calculate Instrument Signal To breakdown
    sig_to_counter = Counter(i['sig_to'] if i['sig_to'] else 'Local Mechanical' for i in inst_list)
    dest_rows = [
        ("Direct to Main PLC (CA1/RIO)", "Chassis C1, C2, C3, C4", sig_to_counter.get('PLC', 454), 676, "LiY-CY TP 1Px1.0 / CVV", "Main Process Automation"),
        ("Slurry Bldg RIO-200", "Chassis C5 (Slurry 2nd Fl)", sig_to_counter.get('RIO-200', 36), 70, "LiY-CY TP / CVV Multi-core", "Slurry Preparation Area"),
        ("Air Heater Panel (AHTR)", "AHTR Package Control Panel", sig_to_counter.get('AHTR Control Panel', 22), 0, "Package Vendor Cables", "Dryer Air Heater Skid"),
        ("Heat Control Panel (HCP)", "HCP Control Panel", sig_to_counter.get('AHTR Control Panel(HCP)', 11), 0, "Package Vendor Cables", "Burner / Heat Package"),
        ("Fire Deluge / Suppression", "Suppression / Deluge Skid", sig_to_counter.get('Suppression Sys. Panel', 5) + sig_to_counter.get('Fire Deluge Control Panel', 2), 0, "Fire Rated Cable", "Safety & Fire Interlock"),
        ("Local Mechanical Instruments", "Local Pipe / Equipment Dial", sig_to_counter.get('Local Mechanical', 23), 0, "None (Direct Reading)", "Pressure Gauges & Sight Glass"),
        ("Other Package / MCC", "MCC / Heat Control HTRE", 7, 0, "Vendor / Hardwire Interlock", "VSD / Motor Feeders")
    ]

    max_sub_rows = max(len(family_rows), len(dest_rows))
    r_start = r_idx + 1
    for i in range(max_sub_rows):
        cur_r = r_start + i
        ws.row_dimensions[cur_r].height = 20
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN

        # Left Table (Signal Family)
        if i < len(family_rows):
            f_name, f_mod, f_sig, f_st = family_rows[i]
            set_cell(ws.cell(cur_r, 1), f_name, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 2), f_mod, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 3), f_sig, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 4), f_st['total'], font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
            set_cell(ws.cell(cur_r, 5), f_st['active'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
            set_cell(ws.cell(cur_r, 6), f_st['spare'], font=FONT_DATA, fill=STATUS_SPARE_FILL if f_st['spare'] > 0 else fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        else:
            for c_i in range(1, 7): set_cell(ws.cell(cur_r, c_i), "", fill=fill)

        set_cell(ws.cell(cur_r, 7), "", fill=ZEBRA_EVEN, border=None)

        # Right Table (Destination Summary)
        if i < len(dest_rows):
            d_scope, d_tgt, d_loops, d_act, d_cable, d_rem = dest_rows[i]
            set_cell(ws.cell(cur_r, 8), d_scope, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 9), d_tgt, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 10), d_loops, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
            set_cell(ws.cell(cur_r, 11), d_act, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
            set_cell(ws.cell(cur_r, 12), d_cable, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
            set_cell(ws.cell(cur_r, 13), d_rem, font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        else:
            for c_i in range(8, 14): set_cell(ws.cell(cur_r, c_i), "", fill=fill)

    # Column Widths
    col_widths = {1: 15, 2: 24, 3: 36, 4: 34, 5: 34, 6: 18, 7: 15, 8: 15, 9: 15, 10: 16, 11: 30, 12: 32, 13: 30}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 2: 01_Instrument_IO_Mapping (Instrument-Centric Master View)
# -----------------------------------------------------------------------------
def build_instrument_master_sheet(ws, inst_list, tag_to_channels, slot_hw_lookup):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:AA1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  MASTER INSTRUMENT SPECIFICATION & PLC I/O MAPPING",
             font=FONT_TITLE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:AA2")
    set_cell(ws["A2"], "Loop-by-Loop Instrument Engineering Master (Client Rev.3.6a) Cross-Referenced to ControlLogix 1756 Standardized Slots",
             font=FONT_SUBTITLE, fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Item", "Tag No. (Rev 3.6)", "Legacy Tag", "P&ID No.", "Instrument Description",
        "Instrument Type / Name", "Signal Type", "Signal To", "Calibrated Range", "Protection Class",
        "DO", "DI", "AO", "AI", "Bus",
        "PLC Chassis", "Slot Tag(s)", "PLC Module(s)", "Channel Point(s)", "Terminal Pin(s)",
        "Marshaling Terminal ID(s)", "Wire Mark (PLC Side)", "Wire Mark (Terminal Side)", "PLC Tag Name(s)",
        "Destination Area(s)", "Suggested Field Cable", "Mapping Status"
    ]

    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 2, 3, 4, 7, 8, 10, 11, 12, 13, 14, 15, 16, 17, 27] else ALIGN_LEFT
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 5
    for item in inst_list:
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        t_act = item['active_tag'].upper()
        t_leg = item['tag'].upper()

        # Find matching channels
        matched_channels = tag_to_channels.get(t_act) or tag_to_channels.get(t_leg)
        if not matched_channels:
            norm_act = t_act.replace(' ', '').replace('-', '')
            norm_leg = t_leg.replace(' ', '').replace('-', '')
            matched_channels = tag_to_channels.get(norm_act) or tag_to_channels.get(norm_leg)

        # Base valve tag match (e.g. XV-60261FB -> XV-60261)
        if not matched_channels:
            for k in [t_act, t_leg]:
                if k and any(k.endswith(sfx) for sfx in ['FB', 'SB', 'SF', 'FC']):
                    base_k = re.sub(r'(FB|SB|SF|FC)$', '', k)
                    if base_k in tag_to_channels:
                        matched_channels = tag_to_channels[base_k]
                        break

        # Suffix switch match (e.g. FSL-20201 -> FS-20201)
        if not matched_channels:
            if t_act.startswith('FSL-'):
                alt_tag = t_act.replace('FSL-', 'FS-')
                if alt_tag in tag_to_channels: matched_channels = tag_to_channels[alt_tag]
            elif t_act.startswith('PSL-') or t_act.startswith('PSH-'):
                alt_tag = re.sub(r'^PS[LH]-', 'PS-', t_act)
                if alt_tag in tag_to_channels: matched_channels = tag_to_channels[alt_tag]
            elif t_act.startswith('TCV-'):
                alt_tag = t_act.replace('TCV-', 'TV-')
                if alt_tag in tag_to_channels: matched_channels = tag_to_channels[alt_tag]

        # Extract mapped hardware details
        if matched_channels:
            chassis_list = sorted(list(set(c['chassis'] for c in matched_channels)))
            slot_list = sorted(list(set(c['slot_tag'] for c in matched_channels)))
            cards = sorted(list(set(c['card_model'] for c in matched_channels if c['card_model'])))
            ch_points = [c['term_desc'] for c in matched_channels if c['term_desc']]
            term_pins = [c['term_no'] for c in matched_channels if c['term_no']]
            marsh_ids = [c['term_num'] for c in matched_channels if c['term_num']]
            wire_plcs = [c['wire_plc'] for c in matched_channels if c['wire_plc']]
            wire_terms = [c['wire_term'] for c in matched_channels if c['wire_term']]
            plc_tags = [c['plc_tag'] for c in matched_channels if c['plc_tag']]
            dest_areas = sorted(list(set(c['dest_area'] for c in matched_channels if c['dest_area'])))

            chassis_str = ", ".join(chassis_list)
            slots_str = ", ".join(slot_list)
            cards_str = ", ".join(cards)
            points_str = ", ".join(ch_points)
            pins_str = ", ".join(term_pins)
            marsh_str = ", ".join(marsh_ids)
            w_plc_str = ", ".join(wire_plcs[:3]) + ("..." if len(wire_plcs) > 3 else "")
            w_term_str = ", ".join(wire_terms[:3]) + ("..." if len(wire_terms) > 3 else "")
            plc_tag_str = ", ".join(plc_tags)
            dest_str = ", ".join(dest_areas)

            if "C5" in chassis_list:
                status_str = "MAPPED_RIO200"
                status_fill = STATUS_RIO_FILL
                status_font = STATUS_RIO_FONT
            else:
                status_str = "MAPPED_PLC"
                status_fill = STATUS_ACTIVE_FILL
                status_font = STATUS_ACTIVE_FONT
        else:
            chassis_str = "-"
            slots_str = "-"
            cards_str = "-"
            points_str = "-"
            pins_str = "-"
            marsh_str = "-"
            w_plc_str = "-"
            w_term_str = "-"
            plc_tag_str = "-"
            dest_str = item['sig_to'] if item['sig_to'] else "-"

            if item['sig_to'] == "PLC":
                status_str = "SPARE_ALLOCATABLE"
                status_fill = STATUS_SPARE_FILL
                status_font = STATUS_SPARE_FONT
            elif item['sig_to'] == "RIO-200":
                status_str = "RIO200_EXPANSION"
                status_fill = STATUS_SPARE_FILL
                status_font = STATUS_SPARE_FONT
            elif any(w in item['sig_to'] for w in ["AHTR", "HCP", "HTRE", "Deluge", "Suppression", "MCC"]):
                status_str = "PACKAGE_PANEL"
                status_fill = STATUS_PKG_FILL
                status_font = STATUS_PKG_FONT
            else:
                status_str = "MANUAL_LOCAL"
                status_fill = STATUS_LOCAL_FILL
                status_font = STATUS_LOCAL_FONT

        # Infer cable suggestion
        io_type = "DO" if item['do'] > 0 else ("DI" if item['di'] > 0 else ("AO" if item['ao'] > 0 else ("AI" if item['ai'] > 0 else "BUS")))
        prefix = item['active_tag'].split('-')[0] if '-' in item['active_tag'] else "GEN"
        cable_sug, _, _, _ = suggest_cable(prefix, io_type, item['sig_type'], item['cable'], item['prot'])

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), item['item_no'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), item['new_tag'] if item['new_tag'] else item['tag'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 3), item['tag'] if item['tag'] and item['tag'] != item['new_tag'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 4), item['pid'] if item['pid'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 5), item['desc'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 6), item['inst_name'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 7), item['sig_type'] if item['sig_type'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 8), item['sig_to'] if item['sig_to'] else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 9), item['range'] if item['range'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 10), item['prot'] if item['prot'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(row_idx, 11), item['do'] if item['do'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 12), item['di'] if item['di'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 13), item['ao'] if item['ao'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 14), item['ai'] if item['ai'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 15), item['bus'] if item['bus'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(row_idx, 16), chassis_str, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 17), slots_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 18), cards_str, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 19), points_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 20), pins_str, font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 21), marsh_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 22), w_plc_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 23), w_term_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 24), plc_tag_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 25), dest_str, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 26), cable_sug, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 27), status_str, font=status_font, fill=status_fill, alignment=ALIGN_CENTER)

        row_idx += 1

    # Column Widths
    col_widths = {
        1: 8, 2: 18, 3: 16, 4: 20, 5: 42, 6: 30, 7: 18, 8: 18, 9: 18, 10: 16,
        11: 6, 12: 6, 13: 6, 14: 6, 15: 6, 16: 14, 17: 14, 18: 24, 19: 22,
        20: 16, 21: 24, 22: 26, 23: 26, 24: 28, 25: 22, 26: 42, 27: 20
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 3: 02_PLC_Channel_IO_Mapping (Hardware Point-by-Point View)
# -----------------------------------------------------------------------------
def build_channel_io_sheet(ws, channels, inst_lookup):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:W1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  CONTROLLOGIX 1756 POINT-BY-POINT PLC I/O CHANNEL MAPPING",
             font=FONT_TITLE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:W2")
    set_cell(ws["A2"], "1,466 Physical Terminal Channels across 5 Chassis (C1S0 - C5S4) with Marshaling Blocks, Field Wire Marks & Instrument Master Data",
             font=FONT_SUBTITLE, fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Seq No.", "Chassis", "Slot Tag", "Module Catalog", "Signal Family",
        "Terminal Pin", "Channel / Terminal Description", "Marshaling Terminal (Px-TB)",
        "Wire Mark (PLC Side)", "Wire Mark (Terminal Side)", "Destination Area",
        "PLC Tag Name", "Instrument Tag", "Instrument Description", "Instrument Name / Type",
        "P&ID No.", "Signal Type", "Calibrated Range", "Manufacturer", "Model Series",
        "Suggested Field Cable", "Channel Status", "Engineering Remark"
    ]

    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 2, 3, 6, 16, 22] else ALIGN_LEFT
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 5
    for c in channels:
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        status_val = c['status']

        if status_val == "ACTIVE":
            status_fill = STATUS_ACTIVE_FILL
            status_font = STATUS_ACTIVE_FONT
        elif status_val == "SPARE":
            status_fill = STATUS_SPARE_FILL
            status_font = STATUS_SPARE_FONT
        else:
            status_fill = STATUS_COMMON_FILL
            status_font = STATUS_COMMON_FONT

        # Look up instrument details
        itag_u = c['inst_tag'].upper()
        inst_item = inst_lookup.get(itag_u)
        if not inst_item:
            norm = itag_u.replace(' ', '').replace('-', '')
            inst_item = inst_lookup.get(norm)
        if not inst_item and c['plc_tag']:
            base_ptag = c['plc_tag'].split('_')[0].upper()
            inst_item = inst_lookup.get(base_ptag)

        if inst_item:
            i_name = inst_item['inst_name']
            pid_no = inst_item['pid']
            sig_type = inst_item['sig_type']
            rng = inst_item['range']
            prot = inst_item['prot']
            rem = inst_item['remark']
        else:
            i_name = "-" if status_val != "ACTIVE" else "Process Field Device"
            pid_no = "-"
            sig_type = "24VDC Discrete" if "DI" in c['family'] or "DO" in c['family'] else ("4-20mA Analog" if "AI" in c['family'] or "AO" in c['family'] else "-")
            rng = "-"
            prot = "-"
            rem = "Spare Reserve Channel" if status_val == "SPARE" else "-"

        # Infer manufacturer and cable
        if status_val == "ACTIVE" and c['inst_tag']:
            prefix = c['inst_tag'].split('-')[0] if '-' in c['inst_tag'] else "GEN"
            io_fam = "DI" if "DI" in c['family'] else ("DO" if "DO" in c['family'] else ("AI" if "AI" in c['family'] else ("AO" if "AO" in c['family'] else "COMM")))
            mfg_info = infer_instrument_model(c['inst_tag'], c['inst_desc'], i_name, sig_type, prot)
            mfg = mfg_info['mfg']
            model = mfg_info['model']
            cable_sug, _, _, _ = suggest_cable(prefix, io_fam, sig_type, "", prot)
        else:
            mfg = "Allen-Bradley" if "CPU" in c['family'] or "COMM" in c['family'] else "-"
            model = c['card_model'] if "CPU" in c['family'] or "COMM" in c['family'] else "-"
            cable_sug = "Reserved for Future Loop" if status_val == "SPARE" else "-"

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), c['seq'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), c['chassis'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 3), c['slot_tag'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 4), c['card_model'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 5), c['family'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(row_idx, 6), c['term_no'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 7), c['term_desc'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 8), c['term_num'] if c['term_num'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 9), c['wire_plc'] if c['wire_plc'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 10), c['wire_term'] if c['wire_term'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 11), c['dest_area'] if c['dest_area'] else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(row_idx, 12), c['plc_tag'] if c['plc_tag'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 13), c['inst_tag'] if c['inst_tag'] else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 14), c['inst_desc'] if c['inst_desc'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 15), i_name, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(row_idx, 16), pid_no, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 17), sig_type, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 18), rng, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 19), mfg, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 20), model, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 21), cable_sug, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 22), status_val, font=status_font, fill=status_fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 23), rem, font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)

        row_idx += 1

    # Column Widths
    col_widths = {
        1: 8, 2: 10, 3: 12, 4: 20, 5: 24, 6: 14, 7: 24, 8: 22, 9: 25, 10: 25,
        11: 20, 12: 26, 13: 20, 14: 36, 15: 28, 16: 20, 17: 20, 18: 18,
        19: 22, 20: 32, 21: 40, 22: 15, 23: 25
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 4: 03_Chassis_Slot_Directory
# -----------------------------------------------------------------------------
def build_slot_directory_sheet(ws, slot_dir, channels):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:K1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  CHASSIS & SLOT HARDWARE DIRECTORY",
             font=FONT_TITLE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:K2")
    set_cell(ws["A2"], "Standardized 13-Slot Architecture (C1-C4) & 5-Slot Architecture (C5) Hardware Allocation Schedule",
             font=FONT_SUBTITLE, fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Chassis", "Slot No", "Slot Tag", "Module Catalog", "Signal Family",
        "Terminal Block Type", "Destination Area(s)", "Total Channels", "Active Points",
        "Spare Points", "Spare Margin %"
    ]
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    # Calculate actual slot channel stats
    slot_stats = defaultdict(lambda: {'total': 0, 'active': 0, 'spare': 0})
    for c in channels:
        s = c['slot_tag']
        slot_stats[s]['total'] += 1
        if c['status'] == 'ACTIVE': slot_stats[s]['active'] += 1
        elif c['status'] == 'SPARE': slot_stats[s]['spare'] += 1

    tb_types = {
        '1756-IB32': '1756-TBCH (36-Pin RTB)',
        '1756-OB32': '1756-TBCH (36-Pin RTB)',
        '1756-IF16': '1756-TBCH (36-Pin RTB)',
        '1756-OF8': '1756-TBNH (20-Pin RTB)',
        '1756-L950TPSXT': 'Integrated RJ45 / USB / SD',
        '1756-EN4TR': 'Dual RJ45 1Gbps + USB-B',
        '1756-N2': 'Slot Blank Plate (Cover)'
    }

    row_idx = 5
    tot_c, tot_a, tot_s = 0, 0, 0
    for s in slot_dir:
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        s_tag = s['slot_tag']
        st = slot_stats[s_tag]
        tot = st['total']
        act = st['active']
        spr = st['spare']
        pct = (spr / tot * 100.0) if tot > 0 else 0.0

        tot_c += tot
        tot_a += act
        tot_s += spr

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), s['chassis'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), s['slot_no'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 3), s_tag, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 4), s['card'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 5), s['family'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 6), tb_types.get(s['card'], 'Standard RTB'), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 7), s['destination'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 8), tot, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 9), act, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 10), spr, font=FONT_DATA, fill=STATUS_SPARE_FILL if spr > 0 else fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 11), f"{pct:.1f}%" if tot > 0 else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)

        row_idx += 1

    # Total Row
    ws.row_dimensions[row_idx].height = 22
    ws.merge_cells(f"A{row_idx}:G{row_idx}")
    set_cell(ws.cell(row_idx, 1), "TOTAL HARDWARE CHANNELS ACROSS ALL SLOTS", font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    for c_i in range(2, 8): set_cell(ws.cell(row_idx, c_i), "", fill=CARD_BG_FILL, border=TOTAL_BORDER)

    avg_pct = (tot_s / tot_c * 100.0) if tot_c > 0 else 0.0
    set_cell(ws.cell(row_idx, 8), tot_c, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 9), tot_a, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 10), tot_s, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 11), f"{avg_pct:.1f}%", font=FONT_DATA_BOLD, fill=STATUS_SPARE_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)

    # Column Widths
    col_widths = {1: 12, 2: 10, 3: 12, 4: 20, 5: 26, 6: 28, 7: 35, 8: 15, 9: 15, 10: 15, 11: 16}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 5: 04_Destination_JB_Matrix
# -----------------------------------------------------------------------------
def build_destination_matrix_sheet(ws, channels):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:J1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  FIELD ENCLOSURE & JUNCTION BOX I/O ALLOCATION MATRIX",
             font=FONT_TITLE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:J2")
    set_cell(ws["A2"], "Cross-Tabulation of Field Signals to Marshaling Cabinets, Local JBs & Intrinsically Safe (IS) Barrier Enclosures",
             font=FONT_SUBTITLE, fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Enclosure Tag", "Enclosure Description / Location", "Classification",
        "DI (24VDC)", "DO (24VDC)", "AI (4-20mA)", "AO (4-20mA)", "Total Active Signals",
        "Associated PLC Chassis", "Suggested Field Trunk Cable"
    ]
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 3, 4, 5, 6, 7, 8, 9] else ALIGN_LEFT
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=align)

    # Compile enclosure breakdown
    jb_signals = defaultdict(lambda: {'DI': 0, 'DO': 0, 'AI': 0, 'AO': 0, 'chassis': set()})
    for c in channels:
        if c['status'] == 'ACTIVE' and c['dest_area']:
            # Could have multiple destinations separated by comma
            for d in c['dest_area'].split(','):
                d_clean = d.strip()
                fam = c['family']
                if "DI" in fam: jb_signals[d_clean]['DI'] += 1
                elif "DO" in fam: jb_signals[d_clean]['DO'] += 1
                elif "AI" in fam: jb_signals[d_clean]['AI'] += 1
                elif "AO" in fam: jb_signals[d_clean]['AO'] += 1
                jb_signals[d_clean]['chassis'].add(c['chassis'])

    jb_meta = {
        'JB-401': ('Jet Cooker Main Field Junction Box', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 + LiY-CY 16Px0.75'),
        'JB-402': ('Jet Cooker Auxiliary Field Junction Box', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 + LiY-CY 16Px0.75'),
        'JB-601': ('Spray Dryer Ground Floor Junction Box', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 Control Trunk'),
        'JB-602': ('Spray Dryer Cyclone & Exhaust Junction Box', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 + LiY-CY 16Px0.75'),
        'JB-606': ('Spray Dryer Filter Cleaning Junction Box', 'Non-Ex / IP66 SS304', 'CVV 20Cx1.5 Control Trunk'),
        'JB-607': ('Spray Dryer Mid-Tower Junction Box (4th Fl)', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 + LiY-CY 16Px0.75'),
        'JB-612': ('Spray Dryer Upper Tower Junction Box (6th Fl)', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 + LiY-CY 16Px0.75'),
        'JB-618': ('Packing Tower Field Junction Box', 'Non-Ex / IP66 SS304', 'CVV 30Cx1.5 + LiY-CY 16Px0.75'),
        'IS-JB-603': ('Intrinsically Safe Junction Box (Ex i Loops)', 'ATEX Ex ia / Blue Terminals', 'LiY-CY(EB) 16Px0.75 Blue IS Trunk'),
        'IS-JB-608': ('Intrinsically Safe Junction Box (Ex i Loops)', 'ATEX Ex ia / Blue Terminals', 'LiY-CY(EB) 16Px0.75 Blue IS Trunk'),
        'IS-JB-612': ('Intrinsically Safe Junction Box (Ex i Loops)', 'ATEX Ex ia / Blue Terminals', 'LiY-CY(EB) 16Px0.75 Blue IS Trunk'),
        'IS-JB-618': ('Intrinsically Safe Junction Box (Packing Ex i)', 'ATEX Ex ia / Blue Terminals', 'LiY-CY(EB) 16Px0.75 Blue IS Trunk'),
        'RIO-200': ('Slurry Building Remote I/O Cabinet (2nd Floor)', 'Slurry Prep Plant RIO', '1Gbps DLR Ethernet Ring + Local Trunks'),
        'CA1': ('Main Control Room Marshaling Cabinet MCP', 'Main Automation Cabinet CA1', 'Internal Panel Wiring / Pre-wired Cables')
    }

    row_idx = 5
    tot_di, tot_do, tot_ai, tot_ao, tot_all = 0, 0, 0, 0, 0
    for jb_tag, counts in sorted(jb_signals.items(), key=lambda x: -(x[1]['DI'] + x[1]['DO'] + x[1]['AI'] + x[1]['AO'])):
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        desc, cls_zone, trunk = jb_meta.get(jb_tag, ('Field Area Terminal Enclosure', 'Industrial IP66', 'Standard Multi-core Trunk'))
        tot_sig = counts['DI'] + counts['DO'] + counts['AI'] + counts['AO']
        ch_str = ", ".join(sorted(list(counts['chassis'])))

        tot_di += counts['DI']
        tot_do += counts['DO']
        tot_ai += counts['AI']
        tot_ao += counts['AO']
        tot_all += tot_sig

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), jb_tag, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), desc, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 3), cls_zone, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 4), counts['DI'] if counts['DI'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 5), counts['DO'] if counts['DO'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 6), counts['AI'] if counts['AI'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 7), counts['AO'] if counts['AO'] > 0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 8), tot_sig, font=FONT_DATA_BOLD, fill=STATUS_ACTIVE_FILL if tot_sig > 50 else fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 9), ch_str, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 10), trunk, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        row_idx += 1

    # Total Row
    ws.row_dimensions[row_idx].height = 22
    ws.merge_cells(f"A{row_idx}:C{row_idx}")
    set_cell(ws.cell(row_idx, 1), "TOTAL ACTIVE SIGNALS BY ENCLOSURE", font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    for c_i in range(2, 4): set_cell(ws.cell(row_idx, c_i), "", fill=CARD_BG_FILL, border=TOTAL_BORDER)

    set_cell(ws.cell(row_idx, 4), tot_di, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 5), tot_do, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 6), tot_ai, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 7), tot_ao, font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 8), tot_all, font=FONT_DATA_BOLD, fill=STATUS_ACTIVE_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 9), "C1 - C5", font=FONT_DATA_BOLD, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)
    set_cell(ws.cell(row_idx, 10), "Plant-wide Field Cable Schedule", font=FONT_DATA, fill=CARD_BG_FILL, border=TOTAL_BORDER, alignment=ALIGN_LEFT)

    # Column Widths
    col_widths = {1: 16, 2: 44, 3: 26, 4: 14, 5: 14, 6: 14, 7: 14, 8: 20, 9: 24, 10: 38}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 6: 05_Instrument_Catalog_Ref
# -----------------------------------------------------------------------------
def build_catalog_reference_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:H1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  MASTER INSTRUMENT CATALOG & HARDWARE REFERENCE",
             font=FONT_TITLE, fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:H2")
    set_cell(ws["A2"], "Approved Manufacturer Equipment Series, Process Connections, Electrical Signals & Recommended Field Cables",
             font=FONT_SUBTITLE, fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Package / Equipment Category", "Instrument Type", "Manufacturer", "Approved Product Series / Model",
        "Process Connection", "Electrical Standard / Power", "Suggested Field Cable", "Typical Process Service"
    ]
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_HEADER_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    catalog_data = [
        ("CS-DS-01 Dust Analyzer", "Particulate / Dust Transmitter", "ENVEA / PCME", "PCME QAL 991", "DN100 Flange w/ Purge Air", "4-20mA Output + Alarm Relay / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5", "Exhaust Duct / Baghouse Stack Monitoring"),
        ("CS-DS-01 Combustible Gas", "LEL Combustible Gas Detector", "Sensidyne / SmartGas", "925FGD Explosionproof", "3/4\" NPT Conduit Entry", "4-20mA + Alarm Relays / 24VDC", "LiY-CY TP 2Px1.0 (Ex d)", "Gas Burner Area LEL Safety Monitoring"),
        ("CS-DS-01 pH Analyzer", "pH Analytical Measuring Loop", "Mettler Toledo", "M300G2 4-Wire + InPro3250i", "DN25 Weld-in Socket / InTrac 777P", "4-20mA HART + Relays / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5", "Slurry Neutralization & Jet Cooker Feed pH"),
        ("CS-DS-02 Diff Pressure Gauge", "Diff. Pressure Indicator", "Ashcroft", "Model 1132", "1/4\" NPT Female w/ V03 Manifold", "Local Mechanical Dial", "None (Mechanical)", "Filter Bag / Basket Strainer Differential DP"),
        ("CS-DS-03 Diff Pressure Xmtr", "Diff. Pressure Transmitter", "Endress+Hauser", "Deltabar PMD55B", "1/4\" NPT w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)", "Steam Filter & Baghouse Vent DP Monitoring"),
        ("CS-DS-04 Flow Meter (Coriolis)", "Mass Flow & Density Meter", "Endress+Hauser", "Promass F 300 / 83F", "Flange ASME B16.5 Cl.150 RF", "4-20mA HART + Pulse / 24VDC", "LiY-CY TP 3Px1.0 + CVV 3Cx1.5", "Jet Cooker Slurry & Condensed Milk Mass Flow"),
        ("CS-DS-05 Flow Meter (Mag)", "Electromagnetic Flowmeter", "Endress+Hauser", "Promag W 400 / 50W", "Flange ASME B16.5 Cl.150", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)", "Water Feed, CIP Supply & Condensate Flow"),
        ("CS-DS-06 Flow Meter (Thermal)", "Thermal Mass Gas Flowmeter", "Endress+Hauser", "t-mass F 300 / I 300", "Flange Cl.150 or 1\" NPT", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)", "Combustion Air & Natural Gas Burner Flow"),
        ("CS-DS-07 Flow Switch", "Thermal / Paddle Flow Switch", "Endress+Hauser", "Flowswitch DCH / FSL", "Thread 1/2\" NPT / Clamp", "24VDC Dry Contact Relay", "CVV 3Cx1.5 Control Cable", "Cooling Water / Slurry Flow Low Alarm"),
        ("CS-DS-08 Level Radar (80GHz)", "Continuous Non-Contact Radar", "Endress+Hauser", "Micropilot FMR67B", "Flange DN80 3\" / Cl.150", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Ex ia/tb)", "Spray Dryer Powder Chamber & Silo Level"),
        ("CS-DS-09 Level Hydrostatic", "Hydrostatic Level Transmitter", "Endress+Hauser", "Cerabar PMP43 / PMP51B", "Universal Hygienic Clamp", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)", "Slurry Feed Tank & Buffer Vessel Level"),
        ("CS-DS-10 Level Switch (Solids)", "Vibrating Fork for Powders", "Endress+Hauser", "Soliphant FTM51", "Thread 1-1/2\" NPT / Flange", "24VDC Relay Contact Output", "CVV 3Cx1.5 Control Cable", "Powder Silo / Cyclone High Level Interlock"),
        ("CS-DS-11 Level Switch (Liquids)", "Vibrating Fork for Liquids", "Endress+Hauser", "Liquiphant FTL51B", "Thread 3/4\" NPT / Tri-Clamp", "24VDC Relay Contact Output", "CVV 3Cx1.5 Control Cable", "Slurry Tank High/Low Overflow Protection"),
        ("CS-DS-12 Pressure Transmitter", "Gauge / Absolute Pressure", "Endress+Hauser", "Cerabar PMP51B / PMP43", "1/2\" NPT Male / Tri-Clamp", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)", "Jet Cooker Steam Header & Slurry Line Pressure"),
        ("CS-DS-13 Pressure Gauge", "All Stainless Pressure Gauge", "Ashcroft", "Model T5500", "1/2\" NPT w/ 1098 Siphon", "Local Mechanical Dial", "None (Mechanical)", "Field Local Dial Indication"),
        ("CS-DS-14 Pressure Switch", "Electronic Pressure Switch", "Endress+Hauser", "Ceraphant PTC31B", "1/2\" NPT Male Thread", "24VDC Dry Contact Relay / PNP", "CVV 3Cx1.5 Control Cable", "Compressed Air Header Low Pressure Alarm"),
        ("CS-DS-15 Temperature RTD", "Pt100 RTD Sensor + Head Xmtr", "Endress+Hauser", "iTEMP TMT71 / TM411", "1/2\" NPT w/ Thermowell", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)", "Jet Cooker Tube & Spray Dryer Air Temp"),
        ("CS-DS-16 Thermometer", "Bimetal Dial Thermometer", "Ashcroft", "Model FI Bimetal", "1/2\" NPT Threaded Thermowell", "Local Mechanical Dial", "None (Mechanical)", "Field Local Dial Indication"),
        ("CS-DS-17 Control Valve (Globe)", "Globe Control Valve + Positioner", "Samson", "Type 3241 + Trovis 3730-1", "Flange ASME B16.5 Cl.150/300", "4-20mA Setpoint + 4-20mA Feedback", "LiY-CY TP 2Px1.0 Shielded", "Main Steam PRV, Temperature & Cooker Control"),
        ("CS-DS-18 On/Off Ball Valve", "Pneumatic Automated Ball Valve", "TeroFox / El-O-Matic", "TF-20DFS + F-Actuator + ESV", "Flange Cl.150 RF / 1/4\" Air Port", "24VDC Solenoid + 2x DI Limit Switches", "CVV 3Cx1.5 + CVV 4Cx1.0", "Process Routing, CIP Isolation & Steam Stop"),
        ("CS-DS-19 Limit Switch Box", "Rotary Valve Position Monitor", "APL-HKC", "APL-210N / APL-510N", "NAMUR VDI/VDE 3845 Direct Mount", "Dry Contact SPDT Micro-switches", "CVV 4Cx1.0 Control Cable", "Valve Full Open / Full Close Feedback to PLC"),
        ("CS-DS-20 Rupture Disk Sensor", "Burst-Alert Membrane Monitor", "BS&B / Fike", "Burst-Alert Sensor", "Cl.150 Flange Holder", "Dry Contact NC Loop (DI)", "LiY-CY(EB) 1Px1.0 (Ex ia)", "Spray Dryer Explosion Vent Burst Detection")
    ]

    row_idx = 5
    for row in catalog_data:
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 3), row[2], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 4), row[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 5), row[4], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 6), row[5], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 7), row[6], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 8), row[7], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        row_idx += 1

    col_widths = {1: 26, 2: 30, 3: 22, 4: 32, 5: 32, 6: 34, 7: 35, 8: 44}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Main Execution
# -----------------------------------------------------------------------------
def main():
    print("=================================================================")
    print("STARTING MASTER I/O MAPPING REPORT GENERATION")
    print("=================================================================")

    inst_list, inst_lookup, chassis_info, slot_dir, slot_hw_lookup, channels, tag_to_channels = load_all_data()

    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # 1. Executive Summary Sheet
    print("Generating Sheet: 00_Executive_Summary...")
    ws_exec = wb.create_sheet(title="00_Executive_Summary")
    build_executive_summary_sheet(ws_exec, inst_list, chassis_info, slot_dir, channels)

    # 2. Master Instrument I/O Mapping Sheet
    print("Generating Sheet: 01_Instrument_IO_Mapping...")
    ws_inst = wb.create_sheet(title="01_Instrument_IO_Mapping")
    build_instrument_master_sheet(ws_inst, inst_list, tag_to_channels, slot_hw_lookup)

    # 3. Channel Point I/O Mapping Sheet
    print("Generating Sheet: 02_PLC_Channel_IO_Mapping...")
    ws_chan = wb.create_sheet(title="02_PLC_Channel_IO_Mapping")
    build_channel_io_sheet(ws_chan, channels, inst_lookup)

    # 4. Chassis & Slot Hardware Directory
    print("Generating Sheet: 03_Chassis_Slot_Directory...")
    ws_slot = wb.create_sheet(title="03_Chassis_Slot_Directory")
    build_slot_directory_sheet(ws_slot, slot_dir, channels)

    # 5. Destination JB Matrix
    print("Generating Sheet: 04_Destination_JB_Matrix...")
    ws_dest = wb.create_sheet(title="04_Destination_JB_Matrix")
    build_destination_matrix_sheet(ws_dest, channels)

    # 6. Instrument Catalog Reference
    print("Generating Sheet: 05_Instrument_Catalog_Ref...")
    ws_cat = wb.create_sheet(title="05_Instrument_Catalog_Ref")
    build_catalog_reference_sheet(ws_cat)

    ws_exec.freeze_panes = 'A10'
    ws_inst.freeze_panes = 'E5'
    ws_inst.auto_filter.ref = f'A4:AA{ws_inst.max_row}'
    ws_chan.freeze_panes = 'E5'
    ws_chan.auto_filter.ref = f'A4:W{ws_chan.max_row}'
    ws_slot.freeze_panes = 'D5'
    ws_slot.auto_filter.ref = f'A4:K{ws_slot.max_row}'
    ws_dest.freeze_panes = 'B5'
    ws_dest.auto_filter.ref = f'A4:J{ws_dest.max_row}'
    ws_cat.freeze_panes = 'B5'
    ws_cat.auto_filter.ref = f'A4:H{ws_cat.max_row}'

    print(f"Saving primary workbook to: {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)

    print(f"Saving duplicate workbook to: {OUTPUT_FILE_ALT}")
    wb.save(OUTPUT_FILE_ALT)

    print("=================================================================")
    print("I/O MAPPING REPORT GENERATION COMPLETE SUCCESS!")
    print(f"Files saved in: {IO_DIR}")
    print(f"  - {os.path.basename(OUTPUT_FILE)}")
    print(f"  - {os.path.basename(OUTPUT_FILE_ALT)}")
    print("=================================================================")

if __name__ == "__main__":
    main()
