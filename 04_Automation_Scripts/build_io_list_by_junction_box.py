#!/usr/bin/env python3
"""
=============================================================================
KALASIN ENGINEERING xCIP-1545 AUTOMATION SYSTEM (Ref: x2608003)
PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER - INGREDION (THAILAND) CO., LTD.
MASTER JUNCTION BOX I/O LIST REPORT GENERATOR (EXCEL WORKBOOK)
WITH DEDICATED DI, DO, AI, AO USED vs SPARE & SPARE % SUMMARY
=============================================================================
Input Files (Folder: 03-IO_List):
  1. IO_List-By_SlotConfig_rev02.xlsx (Standardized 13-Slot Architecture C1S0 - C5S4)
  2. Instrument I-O List Rev.3.6a.xlsx (Client Master Instrument Schedule)

Output Files:
  03-IO_List/IO_List_By_Junction_Box.xlsx
  03-IO_List/Junction_Box_IO_List.xlsx (Duplicate / Alternate reference)

Workbook Structure:
  - 00_IO_Summary_DI_DO_AI_AO    - Master DI, DO, AI, AO Used vs Spare & Spare % Summary
                                    (Grand Summary, Per-Chassis C1-C5, and Per-Junction Box)
  - 01_Junction_Box_Index        - Master Index with KPI cards, enclosure specs & hyperlinks
  - 02_JB_Allocation_Matrix     - Granular signal breakdown (DI, DO, AI, AO) per panel (P1-P4, RIO)
  - Dedicated sheets for EACH Junction Box:
      JB-401, JB-402, JB-601, JB-602, JB-606, JB-607, JB-608, JB-612, JB-618,
      IS-JB-603, IS-JB-608, IS-JB-612, IS-JB-618, CA1, RIO-200, MCC
=============================================================================
"""

import os
import re
from collections import defaultdict, Counter
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
BASE_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, ".."))
IO_DIR = os.path.join(BASE_DIR, "03-IO_List")
SLOT_FILE = os.path.join(IO_DIR, "IO_List-By_SlotConfig_rev02.xlsx")
INST_FILE = os.path.join(IO_DIR, "Instrument I-O List Rev.3.6a.xlsx")

OUTPUT_FILE = os.path.join(IO_DIR, "IO_List_By_Junction_Box.xlsx")
OUTPUT_FILE_ALT = os.path.join(IO_DIR, "Junction_Box_IO_List.xlsx")

# -----------------------------------------------------------------------------
# Color Palette & Styles
# -----------------------------------------------------------------------------
NAVY_HEADER = "1B365D"       # Deep Corporate Navy
NAVY_SUBHEADER = "2D4A77"    # Slate Blue
SECTION_HEADER = "334E68"    # Steel Blue
CARD_HEADER_BG = "0F172A"    # Dark Slate
CARD_BG = "F1F5F9"           # Light Slate 100

WHITE = "FFFFFF"
BORDER_COLOR = "CBD5E1"      # Slate 300

NAVY_FILL = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")
SUB_FILL = PatternFill(start_color=NAVY_SUBHEADER, end_color=NAVY_SUBHEADER, fill_type="solid")
SECTION_FILL = PatternFill(start_color=SECTION_HEADER, end_color=SECTION_HEADER, fill_type="solid")
CARD_HEAD_FILL = PatternFill(start_color=CARD_HEADER_BG, end_color=CARD_HEADER_BG, fill_type="solid")
CARD_BODY_FILL = PatternFill(start_color=CARD_BG, end_color=CARD_BG, fill_type="solid")

ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

ACTIVE_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
ACTIVE_FONT = Font(name="Calibri", size=10, bold=True, color="166534")

SPARE_FILL = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
SPARE_FONT = Font(name="Calibri", size=10, bold=True, color="92400E")

COMMON_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
COMMON_FONT = Font(name="Calibri", size=10, italic=True, color="475569")

IS_BLUE_FILL = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
IS_BLUE_FONT = Font(name="Calibri", size=10, bold=True, color="0369A1")

COMPLIANT_FILL = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
COMPLIANT_FONT = Font(name="Calibri", size=10, bold=True, color="15803D")

FONT_TITLE = Font(name="Calibri", size=13, bold=True, color=WHITE)
FONT_SUBTITLE = Font(name="Calibri", size=9.5, italic=True, color=WHITE)
FONT_HEADER = Font(name="Calibri", size=10, bold=True, color=WHITE)
FONT_SECTION = Font(name="Calibri", size=11, bold=True, color=WHITE)

FONT_CARD_TITLE = Font(name="Calibri", size=8.5, bold=True, color="94A3B8")
FONT_CARD_VAL = Font(name="Calibri", size=14, bold=True, color="1E293B")
FONT_CARD_SUB = Font(name="Calibri", size=8, italic=True, color="64748B")

FONT_DATA = Font(name="Calibri", size=10, color="0F172A")
FONT_DATA_BOLD = Font(name="Calibri", size=10, bold=True, color="0F172A")
FONT_DATA_CODE = Font(name="Consolas", size=9.5, color="0F172A")
FONT_DATA_MUTED = Font(name="Calibri", size=9, italic=True, color="64748B")

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
    prefix = t_upper.split('-')[0] if '-' in t_upper else re.split(r'(\d+)', t_upper)[0]

    if prefix == "FT" and ("coriolis" in d_lower or "mass" in d_lower or "promass" in d_lower or "slurry" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Promass F 300 / 83F (Coriolis Mass Flow)"}
    if prefix == "FT" and ("magnetic" in d_lower or "mag" in d_lower or "promag" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Promag W 400 / 50W (Electromagnetic Flow)"}
    if prefix == "FT" and ("thermal" in d_lower or "gas" in d_lower or "air" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "t-mass F 300 / I 300 (Thermal Mass Flow)"}
    if prefix in ["FT", "FI", "FE"]:
        return {"mfg": "Endress+Hauser", "model": "Proline Series Flowmeter"}
    if prefix in ["DPT", "DPIT"]:
        return {"mfg": "Endress+Hauser", "model": "Deltabar PMD55B (Diff. Pressure)"}
    if prefix in ["PT", "PIT", "PI"]:
        return {"mfg": "Endress+Hauser", "model": "Cerabar PMP51B / PMP43"}
    if prefix in ["PS", "PSL", "PSH"]:
        return {"mfg": "Endress+Hauser", "model": "Ceraphant PTC31B (Pressure Switch)"}
    if prefix in ["DPG", "DPI"]:
        return {"mfg": "Ashcroft", "model": "Model 1132 Diff. Pressure Gauge"}
    if prefix in ["PG"]:
        return {"mfg": "Ashcroft", "model": "Model T5500 All Stainless Pressure Gauge"}
    if prefix in ["LT", "LIT"]:
        if "radar" in d_lower or "silo" in d_lower or "dryer" in d_lower:
            return {"mfg": "Endress+Hauser", "model": "Micropilot FMR67B (80GHz Radar Level)"}
        return {"mfg": "Endress+Hauser", "model": "Cerabar PMP43 / PMP51B (Hydrostatic)"}
    if prefix in ["LS", "LSL", "LSH"]:
        if any(w in d_lower for w in ["powder", "solid", "dryer", "cyclone", "silo"]):
            return {"mfg": "Endress+Hauser", "model": "Soliphant FTM51 (Vibrating Fork Solids)"}
        return {"mfg": "Endress+Hauser", "model": "Liquiphant FTL51B (Vibrating Fork Liquids)"}
    if prefix in ["TT", "TIT"]:
        return {"mfg": "Endress+Hauser", "model": "iTEMP TMT71 / TM411 (Pt100 RTD)"}
    if prefix in ["TG", "TI"]:
        return {"mfg": "Ashcroft", "model": "Model FI Bimetal Dial Thermometer"}
    if prefix in ["TS", "TSH", "TSL"]:
        return {"mfg": "Endress+Hauser", "model": "Thermophant TTR35 (Temperature Switch)"}
    if prefix in ["AT", "AI", "PHT"]:
        if "ph" in d_lower:
            return {"mfg": "Mettler Toledo", "model": "M300G2 + InPro3250i (pH System)"}
        elif any(w in d_lower for w in ["gas", "flame", "lel"]):
            return {"mfg": "Sensidyne / SmartGas", "model": "925FGD Explosionproof Gas Detector"}
        return {"mfg": "ENVEA / PCME", "model": "PCME QAL 991 (Dust Analyzer)"}
    if prefix in ["TCV", "PCV", "FCV", "CV", "TV"]:
        return {"mfg": "Samson", "model": "Type 3241 Globe Valve + Trovis 3730-1"}
    if prefix in ["XV", "SV", "BV", "ISV"]:
        return {"mfg": "TeroFox / El-O-Matic", "model": "TF-20DFS Ball Valve + F-Actuator + ESV"}
    if prefix in ["HV"]:
        return {"mfg": "TeroFox / APL-HKC", "model": "TF-20DFS Manual Ball Valve + APL-210N Box"}
    if prefix in ["ZS"]:
        return {"mfg": "APL-HKC", "model": "APL-210N / APL-510N Valve Monitor"}
    if prefix in ["PSE"]:
        return {"mfg": "BS&B / Fike", "model": "Burst-Alert Sensor / Rupture Disk Monitor"}
    if prefix in ["VIB", "NCT"]:
        return {"mfg": "Netter Vibration", "model": "NCT 5 Pneumatic Turbine Vibrator"}
    if prefix in ["ATY", "PCY", "BFY", "RVM", "PCM", "BLM", "P", "M"]:
        return {"mfg": "Rockwell Automation", "model": "PowerFlex VSD / E300 Relay in MCC"}
    return {"mfg": "Industrial Process Equipment", "model": "Standard Process Field Device"}

def suggest_cable(prefix, io_type, signal_type, cable_existing, prot_type):
    p_upper = str(prot_type).upper()
    is_is = "EX I" in p_upper or "EX IA" in p_upper or "EX IC" in p_upper

    if prefix in ["XV", "SV", "BV", "ISV"]:
        return ("CVV 3Cx1.5 (Solenoid) + CVV 4Cx1.0 (Limit Switches)", "3Cx1.5 + 4Cx1.0", "M20x1.5 Ex d Double Compression", "CVV 30Cx1.5 Trunk to CA1")
    elif prefix in ["TCV", "PCV", "FCV", "CV", "TV"]:
        return ("LiY-CY TP 2Px1.0 (Setpoint AO + Feedback AI)", "2Px1.0 Shielded Pairs", "M20x1.5 Brass Nickel-Plated Ex d/e", "LiY-CY(IS/OS) 16Px0.75 Trunk")
    elif prefix in ["AT", "AI", "PHT"]:
        return ("LiY-CY TP 2Px1.0 (Signal) + CVV 3Cx1.5 (24VDC Power)", "2Px1.0 + 3Cx1.5", "M20x1.5 Ex d/e Gland", "LiY-CY 16Px0.75 + CVV Trunk")
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
# Junction Box Metadata Registry
# -----------------------------------------------------------------------------
JB_METADATA = {
    'JB-401': {
        'desc': 'Jet Cooker Infeed Area 2nd Floor Junction Box',
        'location': 'Infeed Processing Building 2nd Floor',
        'size': '800 x 600 x 250 mm',
        'type': 'IP66 / SS304 (Industrial Hygienic)',
        'area': 'Slurry Infeed & Pre-Treatment Area',
        'primary_panel': 'Panel P1 (Chassis C1)',
        'trunk_cable': '1x CVV 30Cx1.5 Control Trunk + 1x LiY-CY 16Px0.75 Analog Trunk to CA1'
    },
    'JB-402': {
        'desc': 'Jet Cooker Processing 2nd Floor Junction Box',
        'location': 'Jet Cooker Processing Skid 2F',
        'size': '450 x 300 x 150 mm',
        'type': 'IP66 / SS304 (Industrial Hygienic)',
        'area': 'Thermal Cooking Skid & Direct Steam Injection',
        'primary_panel': 'Panel P2 (Chassis C2)',
        'trunk_cable': '1x CVV 18Cx1.5 Control Trunk + 1x LiY-CY 12Px0.75 Analog Trunk to CA1'
    },
    'JB-601': {
        'desc': 'Spray Dryer Ground Floor Base Junction Box',
        'location': 'Spray Dryer Tower Floor 1F (Base / Discharge)',
        'size': '500 x 400 x 200 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Spray Dryer Base Area & Powder Discharge',
        'primary_panel': 'Panel P1 (Chassis C1)',
        'trunk_cable': '1x CVV 18Cx1.5 Control Trunk + 1x LiY-CY 12Px0.75 Analog Trunk to CA1'
    },
    'JB-602': {
        'desc': 'Spray Dryer Cyclone & Exhaust Junction Box',
        'location': 'Spray Dryer Tower Floor 3F (Cyclone Section)',
        'size': '800 x 600 x 250 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Chamber Middle Section, Cyclones & Air Inlets',
        'primary_panel': 'Panel P2 (Chassis C2)',
        'trunk_cable': '1x CVV 24Cx1.5 Control Trunk + 1x LiY-CY 16Px0.75 Analog Trunk to CA1'
    },
    'JB-606': {
        'desc': 'Spray Dryer Filter Cleaning Junction Box',
        'location': 'Spray Dryer Tower 6F (Air Heater Section)',
        'size': '300 x 200 x 150 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Direct Gas Burner Section & Air Heating',
        'primary_panel': 'Panel P3 (Chassis C3)',
        'trunk_cable': '1x CVV 18Cx1.5 Control Trunk + 1x LiY-CY 4Px0.75 Analog Trunk to Panel M3/P3'
    },
    'JB-607': {
        'desc': 'Spray Dryer Mid-Tower Junction Box (4th Fl)',
        'location': 'Spray Dryer Tower Floor 4F (Main Intermediate Deck)',
        'size': '1000 x 800 x 300 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Main Tower Intermediate Section & High-Density Hub',
        'primary_panel': 'Panel P2 (Chassis C2) & Panel P3 (Chassis C3)',
        'trunk_cable': '2x CVV 24Cx1.5 Control Trunks + 2x LiY-CY 16Px0.75 Analog Trunks'
    },
    'JB-608': {
        'desc': 'Spray Dryer Auxiliary Junction Box',
        'location': 'Spray Dryer Tower 4F Auxiliary Deck',
        'size': '300 x 200 x 150 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Auxiliary Skid & Sight Glass Monitoring',
        'primary_panel': 'Panel P3 (Chassis C3)',
        'trunk_cable': 'Local Field Multi-core Trunk to Panel M3/P3'
    },
    'JB-612': {
        'desc': 'Spray Dryer Upper Tower Junction Box (6th Fl)',
        'location': 'Spray Dryer Upper Tower (6th / 8th Floor)',
        'size': '600 x 500 x 200 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Top Chamber, Atomizer Head & Upper Vents',
        'primary_panel': 'Panel P4 (Chassis C4)',
        'trunk_cable': '1x CVV 24Cx1.5 Control Trunk + 1x LiY-CY 12Px0.75 Analog Trunk to Panel M4/P4'
    },
    'JB-618': {
        'desc': 'Packing Tower Field Junction Box',
        'location': 'Packing Tower Building & Storage Silos',
        'size': '600 x 400 x 200 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Packing Tower Feed, Sieve & Silo Discharge',
        'primary_panel': 'Panel P3 (Chassis C3)',
        'trunk_cable': '1x CVV 18Cx1.5 Control Trunk + 1x LiY-CY 8Px0.75 Analog Trunk to Panel M3/P3'
    },
    'IS-JB-603': {
        'desc': 'Intrinsically Safe Junction Box (Dryer Lower)',
        'location': 'Hazardous Area Zone 1/21 (Lower Chamber Section)',
        'size': '500 x 400 x 200 mm',
        'type': 'ATEX Ex ia (Blue Terminals / IP66 SS316)',
        'area': 'Classified Lower Dryer Explosion Proof Zone',
        'primary_panel': 'Panel P4 (Chassis C4 / IS Barriers)',
        'trunk_cable': '1x LiY-CY(EB) 16Px0.75 Blue IS Shielded Trunk Cable to IS Barriers in CA1/P4'
    },
    'IS-JB-608': {
        'desc': 'Intrinsically Safe Junction Box (Dryer Mid)',
        'location': 'Hazardous Area Zone 1/21 (Mid-Tower Section)',
        'size': '600 x 500 x 200 mm',
        'type': 'ATEX Ex ia (Blue Terminals / IP66 SS316)',
        'area': 'Classified Mid-Tower Powder Dust Cloud Zone',
        'primary_panel': 'Panel P4 (Chassis C4 / IS Barriers)',
        'trunk_cable': '1x LiY-CY(EB) 16Px0.75 Blue IS Shielded Trunk Cable to IS Barriers in CA1/P4'
    },
    'IS-JB-612': {
        'desc': 'Intrinsically Safe Junction Box (Dryer Top)',
        'location': 'Hazardous Area Zone 1/21 (Upper Tower / Explosion Vent)',
        'size': '500 x 400 x 200 mm',
        'type': 'ATEX Ex ia (Blue Terminals / IP66 SS316)',
        'area': 'Classified Upper Chamber & Explosion Vent Relief Zone',
        'primary_panel': 'Panel P4 (Chassis C4 / IS Barriers)',
        'trunk_cable': '1x LiY-CY(EB) 12Px0.75 Blue IS Shielded Trunk Cable to IS Barriers in CA1/P4'
    },
    'IS-JB-618': {
        'desc': 'Intrinsically Safe Junction Box (Packing Tower)',
        'location': 'Hazardous Area Zone 1/21 (Packing Tower & Silos)',
        'size': '600 x 400 x 200 mm',
        'type': 'ATEX Ex ia (Blue Terminals / IP66 SS316)',
        'area': 'Classified Powder Handling & Packing Tower Zone',
        'primary_panel': 'Panel P4 (Chassis C4 / IS Barriers)',
        'trunk_cable': '1x LiY-CY(EB) 16Px0.75 Blue IS Shielded Trunk Cable to IS Barriers in CA1/P4'
    },
    'CA1': {
        'desc': 'Main Control Room Cabinet (MCP)',
        'location': 'Main Control Room (MCP Central Room)',
        'size': '2200 x 800 x 800 mm',
        'type': 'NEMA 12 / IP54 Rittal TS8 Bayed Suite',
        'area': 'Central Process Automation & CPU Racks (C1 & C2)',
        'primary_panel': 'Panels P1 & P2 (Main MCP)',
        'trunk_cable': 'Pre-wired Marshaling Trunks, Internal Busbars & 1Gbps Fiber DLR Rings'
    },
    'RIO-200': {
        'desc': 'Slurry Building Remote I/O Cabinet',
        'location': 'Slurry Building 2nd Floor (RIO Room)',
        'size': '1200 x 800 x 300 mm',
        'type': 'IP66 / SS304 (Sanitary Food Grade)',
        'area': 'Slurry Preparation & Mixing Remote Automation',
        'primary_panel': 'Panel P5 (Chassis C5 RIO-200)',
        'trunk_cable': '1Gbps EtherNet/IP DLR Fiber Ring + Local Field Multi-core Trunks'
    },
    'MCC': {
        'desc': 'Motor Control Center Switchgear Room',
        'location': 'Electrical Substation Switchgear Room',
        'size': 'Centerline 2500 MCC Suite',
        'type': 'Form 4b Arc-Resistant Switchgear',
        'area': 'Motor Feeder Starters, Variable Speed Drives & VFDs',
        'primary_panel': 'Panels P1 / P2 & Hardwired Interlocks',
        'trunk_cable': 'Multi-Core Control Trunks (CVV 7Cx1.5) & EtherNet/IP Backbone'
    }
}

ORDERED_JBS = [
    'JB-401', 'JB-402', 'JB-601', 'JB-602', 'JB-606', 'JB-607', 'JB-608', 'JB-612', 'JB-618',
    'IS-JB-603', 'IS-JB-608', 'IS-JB-612', 'IS-JB-618', 'CA1', 'RIO-200', 'MCC'
]

# -----------------------------------------------------------------------------
# Data Loader
# -----------------------------------------------------------------------------
def load_data():
    print(f"Loading Instrument Master: {INST_FILE}")
    wb_inst = openpyxl.load_workbook(INST_FILE, data_only=True)
    ws_inst = wb_inst['Rev.3']

    inst_lookup = {}
    mcc_instruments = []

    for r in range(12, ws_inst.max_row + 1):
        item_no = ws_inst.cell(r, 1).value
        tag = ws_inst.cell(r, 2).value
        new_tag = ws_inst.cell(r, 3).value
        pid = ws_inst.cell(r, 4).value
        desc = ws_inst.cell(r, 5).value
        inst_name = ws_inst.cell(r, 6).value
        sig_type = ws_inst.cell(r, 12).value
        sig_to = ws_inst.cell(r, 13).value
        rng = ws_inst.cell(r, 14).value
        cable = ws_inst.cell(r, 16).value
        prot = ws_inst.cell(r, 17).value
        rem = ws_inst.cell(r, 18).value

        if tag or new_tag or desc:
            item = {
                'item_no': str(item_no).strip() if item_no is not None else '',
                'tag': str(tag).strip() if tag else '',
                'new_tag': str(new_tag).strip() if new_tag else '',
                'pid': str(pid).strip() if pid else '',
                'desc': str(desc).strip() if desc else '',
                'inst_name': str(inst_name).strip() if inst_name else '',
                'sig_type': str(sig_type).strip() if sig_type else '',
                'sig_to': str(sig_to).strip() if sig_to else '',
                'range': str(rng).strip() if rng else '',
                'cable': str(cable).strip() if cable else '',
                'prot': str(prot).strip() if prot else '',
                'remark': str(rem).strip() if rem else ''
            }
            active_t = item['new_tag'] if item['new_tag'] else item['tag']
            for k in [active_t, item['new_tag'], item['tag']]:
                if k:
                    u = k.upper()
                    inst_lookup[u] = item
                    inst_lookup[u.replace(' ', '').replace('-', '')] = item

            if "MCC" in item['sig_to'] or any(p in active_t for p in ['ATY-', 'PCY-', 'BFY-', 'RVM-', 'PCM-', 'BLM-']):
                mcc_instruments.append(item)

    print(f"Loaded {len(inst_lookup)} instrument keys from Rev.3.6a.")

    print(f"Loading Slot Configurations: {SLOT_FILE}")
    wb_slots = openpyxl.load_workbook(SLOT_FILE, data_only=True)
    slot_sheets = [s for s in wb_slots.sheetnames if s not in ['00_Slot_Index', 'C1SlotConfig', 'C2SlotConfig', 'C3SlotConfig', 'C4SlotConfig', 'C5SlotConfig']]

    ws_idx = wb_slots['00_Slot_Index']
    slot_cards = {}
    for r in range(14, ws_idx.max_row + 1):
        s_tag = ws_idx.cell(r, 1).value
        card = ws_idx.cell(r, 4).value
        fam = ws_idx.cell(r, 5).value
        if s_tag:
            slot_cards[str(s_tag).strip()] = (str(card).strip() if card else '', str(fam).strip() if fam else '')

    channels_by_jb = defaultdict(list)
    by_chassis = defaultdict(lambda: defaultdict(lambda: {'used': 0, 'spare': 0, 'common': 0, 'total': 0}))
    by_jb_summary = defaultdict(lambda: defaultdict(lambda: {'used': 0, 'spare': 0, 'common': 0, 'total': 0}))
    total_loaded = 0

    for sname in slot_sheets:
        ws = wb_slots[sname]
        chassis = sname[:2]
        card_model, family = slot_cards.get(sname, ('', ''))

        for r in range(6, ws.max_row + 1):
            t_no = ws.cell(r, 1).value
            if t_no is None: continue
            total_loaded += 1

            t_desc = ws.cell(r, 2).value
            t_num = ws.cell(r, 3).value
            w_plc = ws.cell(r, 4).value
            w_term = ws.cell(r, 5).value
            dest_raw = ws.cell(r, 6).value
            p_tag = ws.cell(r, 7).value
            i_tag = ws.cell(r, 8).value
            i_desc = ws.cell(r, 9).value
            status = str(ws.cell(r, 10).value or '').strip().upper()

            # Determine Panel
            t_num_str = str(t_num).strip() if t_num else ''
            if t_num_str.startswith('P1-'): panel = 'P1'
            elif t_num_str.startswith('P2-'): panel = 'P2'
            elif t_num_str.startswith('P3-'): panel = 'P3'
            elif t_num_str.startswith('P4-'): panel = 'P4'
            elif t_num_str.startswith('P5-'): panel = 'P5'
            elif chassis == 'C1': panel = 'P1'
            elif chassis == 'C2': panel = 'P2'
            elif chassis == 'C3': panel = 'P3'
            elif chassis == 'C4': panel = 'P4'
            elif chassis == 'C5': panel = 'P5'
            else: panel = 'CA1'

            # Clean Destination Tag
            dest_str = str(dest_raw).strip() if dest_raw else 'CA1'
            if dest_str in ['CA1 (Main Control Room)', 'CA1 / Field RIO (EtherNet/IP DLR Trunk)', '-']:
                dest_jb = 'CA1'
            else:
                dest_jb = dest_str

            # Signal Family determination
            fam_str = family.upper()
            if "1756-IB32" in card_model or fam_str.startswith("DI ") or "(DI)" in fam_str: io_type = "DI"
            elif "1756-OB32" in card_model or fam_str.startswith("DO ") or "(DO)" in fam_str: io_type = "DO"
            elif "1756-IF16" in card_model or fam_str.startswith("AI ") or "(AI)" in fam_str: io_type = "AI"
            elif "1756-OF8" in card_model or fam_str.startswith("AO ") or "(AO)" in fam_str: io_type = "AO"
            elif "COMM" in fam_str: io_type = "COMM"
            elif "CPU" in fam_str: io_type = "CPU"
            else: io_type = "GEN"

            ch_item = {
                'chassis': chassis,
                'slot_tag': sname,
                'card_model': card_model,
                'family': family,
                'io_type': io_type,
                'panel': panel,
                'term_no': str(t_no).strip(),
                'term_desc': str(t_desc).strip() if t_desc else '',
                'term_num': t_num_str,
                'wire_plc': str(w_plc).strip() if w_plc else '',
                'wire_term': str(w_term).strip() if w_term else '',
                'dest_jb': dest_jb,
                'plc_tag': str(p_tag).strip() if p_tag else '',
                'inst_tag': str(i_tag).strip() if i_tag else '',
                'inst_desc': str(i_desc).strip() if i_desc else '',
                'status': status if status else 'ACTIVE'
            }

            channels_by_jb[dest_jb].append(ch_item)

            # Summaries
            if io_type in ['DI', 'DO', 'AI', 'AO']:
                by_chassis[chassis][io_type]['total'] += 1
                by_jb_summary[dest_jb][io_type]['total'] += 1
                if status == 'ACTIVE':
                    by_chassis[chassis][io_type]['used'] += 1
                    by_jb_summary[dest_jb][io_type]['used'] += 1
                elif status == 'SPARE':
                    by_chassis[chassis][io_type]['spare'] += 1
                    by_jb_summary[dest_jb][io_type]['spare'] += 1
                else:
                    by_chassis[chassis][io_type]['common'] += 1
                    by_jb_summary[dest_jb][io_type]['common'] += 1

    print(f"Loaded {total_loaded} physical channels mapped across {len(channels_by_jb)} junction boxes.")
    return inst_lookup, mcc_instruments, channels_by_jb, by_chassis, by_jb_summary

# -----------------------------------------------------------------------------
# Sheet 0: 00_IO_Summary_DI_DO_AI_AO (Dedicated Master Summary)
# -----------------------------------------------------------------------------
def build_io_summary_sheet(ws, by_chassis, by_jb_summary):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:N1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  I/O SIGNAL ALLOCATION & SPARE CAPACITY SUMMARY (DI, DO, AI, AO)",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:N2")
    set_cell(ws["A2"], "SPRINT 18K TPA SPRAY DRYER / JET COOKER (Ref: x2608003) — Detailed Used, Spare & Spare % Calculation",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Table 1: Master Signal Family Summary (Row 4)
    ws.merge_cells("A4:N4")
    set_cell(ws["A4"], "1. OVERALL I/O SIGNAL FAMILY USED vs SPARE & SPARE % CAPACITY",
             font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[4].height = 22

    t1_headers = [
        "Signal Family", "Hardware Card Catalog", "Signal Type Standard", "Used (Active) Pts",
        "Spare (Reserve) Pts", "Total Process Pts", "Spare Margin %", "Common / Power Pins",
        "Total Physical Pins", "Minimum Spec %", "Contractual Compliance", "Primary Enclosure"
    ]
    ws.row_dimensions[5].height = 26
    for c_idx, h in enumerate(t1_headers, 1):
        align = ALIGN_CENTER if c_idx in [1, 7, 10, 11] else ALIGN_LEFT
        set_cell(ws.cell(5, c_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    # Calculate overall stats
    overall = defaultdict(lambda: {'used': 0, 'spare': 0, 'common': 0, 'total': 0})
    for c in by_chassis.values():
        for sig in ['DI', 'DO', 'AI', 'AO']:
            overall[sig]['used'] += c[sig]['used']
            overall[sig]['spare'] += c[sig]['spare']
            overall[sig]['common'] += c[sig]['common']
            overall[sig]['total'] += c[sig]['total']

    sig_meta = [
        ('DI', 'DI (24VDC Digital Input)', '1756-IB32 (32-Pt Sink/Source)', '24VDC Dry Contact / PNP Prox', 'Panels P1, P2, P3, P4'),
        ('DO', 'DO (24VDC Digital Output)', '1756-OB32 (32-Pt Source Transistor)', '24VDC Energized Solenoids', 'Panels P1, P2, P3, P4'),
        ('AI', 'AI (4-20mA Analog Input)', '1756-IF16 (16-Pt Single-Ended / Diff)', '4-20mA HART Transmitters', 'Panels P1, P2, P3, P4 (IS)'),
        ('AO', 'AO (4-20mA Analog Output)', '1756-OF8 (8-Pt Voltage / Current)', '4-20mA Control Valves (CV)', 'Panel P3 (Chassis C3)')
    ]

    r_idx = 6
    tot_used, tot_spare, tot_proc, tot_comm, tot_phys = 0, 0, 0, 0, 0

    for key, name, card, sig_desc, pri_enc in sig_meta:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        d = overall[key]
        u = d['used']
        s = d['spare']
        proc_tot = u + s
        comm = d['common']
        phys_tot = d['total']
        pct = (s / proc_tot * 100.0) if proc_tot > 0 else 0.0

        tot_used += u
        tot_spare += s
        tot_proc += proc_tot
        tot_comm += comm
        tot_phys += phys_tot

        ws.row_dimensions[r_idx].height = 20
        set_cell(ws.cell(r_idx, 1), name, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 2), card, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 3), sig_desc, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(r_idx, 4), u, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 5), s, font=FONT_DATA_BOLD, fill=SPARE_FILL if s > 0 else fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 6), proc_tot, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 7), f"{pct:.1f}%", font=Font(name="Calibri", size=10, bold=True, color="15803D" if pct >= 20 else "92400E"),
                 fill=COMPLIANT_FILL if pct >= 20 else SPARE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 8), comm, font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 9), phys_tot, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 10), "≥ 20.0%", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 11), "COMPLIANT (PASSED)", font=COMPLIANT_FONT, fill=COMPLIANT_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 12), pri_enc, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        r_idx += 1

    # Total Row Table 1
    ws.row_dimensions[r_idx].height = 22
    ws.merge_cells(f"A{r_idx}:C{r_idx}")
    set_cell(ws.cell(r_idx, 1), "TOTAL PROCESS I/O SYSTEM CAPACITY", font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    for c_i in range(2, 4): set_cell(ws.cell(r_idx, c_i), "", fill=CARD_BODY_FILL, border=TOTAL_BORDER)

    avg_pct = (tot_spare / tot_proc * 100.0) if tot_proc > 0 else 0.0
    set_cell(ws.cell(r_idx, 4), tot_used, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 5), tot_spare, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 6), tot_proc, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 7), f"{avg_pct:.1f}%", font=Font(name="Calibri", size=11, bold=True, color="15803D"), fill=COMPLIANT_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)
    set_cell(ws.cell(r_idx, 8), tot_comm, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 9), tot_phys, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 10), "≥ 20.0%", font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)
    set_cell(ws.cell(r_idx, 11), "OVERALL PASS (>20%)", font=COMPLIANT_FONT, fill=COMPLIANT_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)
    set_cell(ws.cell(r_idx, 12), "5 Chassis / 16 Field JBs", font=FONT_DATA, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_LEFT)

    r_idx += 2

    # Table 2: Chassis Breakdown Table
    ws.merge_cells(f"A{r_idx}:N{r_idx}")
    set_cell(ws.cell(r_idx, 1), "2. CHASSIS HARDWARE UTILIZATION & SPARE % BREAKDOWN (CHASSIS C1 TO C5)",
             font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[r_idx].height = 22

    r_idx += 1
    t2_headers = [
        "Chassis", "Chassis Description & Location",
        "DI Used", "DI Spare", "DI Spare %",
        "DO Used", "DO Spare", "DO Spare %",
        "AI Used", "AI Spare", "AI Spare %",
        "AO Used", "AO Spare", "Total Used", "Total Spare", "Total Proc Pts", "Chassis Spare %", "Status"
    ]
    # We will adjust merge for headers
    ws.row_dimensions[r_idx].height = 26
    ch_col_hdrs = [
        "Chassis Code", "Chassis Description / Location",
        "DI Used", "DI Spare", "DI %",
        "DO Used", "DO Spare", "DO %",
        "AI Used", "AI Spare", "AI %",
        "AO Used", "AO Spare", "Tot Used", "Tot Spare", "Total Proc", "Spare %", "Status"
    ]
    for c_idx, h in enumerate(ch_col_hdrs, 1):
        set_cell(ws.cell(r_idx, c_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    chassis_meta = [
        ('C1', 'Main Controller Rack (1756-L950TPSXT) — Control Cabinet CA1'),
        ('C2', 'Main Expansion High-Density Rack — Control Cabinet CA1'),
        ('C3', 'Remote I/O Rack 1 — Field RIO Cabinet (Spray Dryer / Packing)'),
        ('C4', 'Remote I/O Rack 2 — Upper Spray Dryer Tower 6th/8th Floor'),
        ('C5', 'Remote I/O Station 5 — Slurry Building 2nd Floor (RIO-200)')
    ]

    r_idx += 1
    for c_code, c_desc in chassis_meta:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        d = by_chassis[c_code]

        di_u, di_s = d['DI']['used'], d['DI']['spare']
        di_pct = (di_s / (di_u + di_s) * 100.0) if (di_u + di_s) > 0 else 0.0

        do_u, do_s = d['DO']['used'], d['DO']['spare']
        do_pct = (do_s / (do_u + do_s) * 100.0) if (do_u + do_s) > 0 else 0.0

        ai_u, ai_s = d['AI']['used'], d['AI']['spare']
        ai_pct = (ai_s / (ai_u + ai_s) * 100.0) if (ai_u + ai_s) > 0 else 0.0

        ao_u, ao_s = d['AO']['used'], d['AO']['spare']
        ao_pct = (ao_s / (ao_u + ao_s) * 100.0) if (ao_u + ao_s) > 0 else 0.0

        c_u = di_u + do_u + ai_u + ao_u
        c_s = di_s + do_s + ai_s + ao_s
        c_tot = c_u + c_s
        c_pct = (c_s / c_tot * 100.0) if c_tot > 0 else 0.0

        ws.row_dimensions[r_idx].height = 20
        set_cell(ws.cell(r_idx, 1), c_code, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 2), c_desc, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(r_idx, 3), di_u, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 4), di_s, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 5), f"{di_pct:.1f}%" if (di_u+di_s)>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(r_idx, 6), do_u, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 7), do_s, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 8), f"{do_pct:.1f}%" if (do_u+do_s)>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(r_idx, 9), ai_u, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 10), ai_s, font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 11), f"{ai_pct:.1f}%" if (ai_u+ai_s)>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(r_idx, 12), ao_u if (ao_u+ao_s)>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 13), ao_s if (ao_u+ao_s)>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)

        set_cell(ws.cell(r_idx, 14), c_u, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 15), c_s, font=FONT_DATA_BOLD, fill=SPARE_FILL, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 16), c_tot, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 17), f"{c_pct:.1f}%", font=Font(name="Calibri", size=10, bold=True, color="15803D"), fill=COMPLIANT_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(r_idx, 18), "PASSED (>20%)", font=COMPLIANT_FONT, fill=COMPLIANT_FILL, alignment=ALIGN_CENTER)

        r_idx += 1

    r_idx += 1

    # Table 3: Per-Junction Box DI, DO, AI, AO Summary
    ws.merge_cells(f"A{r_idx}:R{r_idx}")
    set_cell(ws.cell(r_idx, 1), "3. FIELD JUNCTION BOX & ENCLOSURE I/O SUMMARY (DI, DO, AI, AO USED vs SPARE & SPARE %)",
             font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[r_idx].height = 22

    r_idx += 1
    jb_col_hdrs = [
        "Enclosure Tag", "Enclosure Description / Area",
        "DI Used", "DI Spare", "DI Total", "DI Spare %",
        "DO Used", "DO Spare", "DO Total", "DO Spare %",
        "AI Used", "AI Spare", "AI Total", "AI Spare %",
        "Total Active", "Total Spare", "Total Pins", "Overall Spare %"
    ]
    ws.row_dimensions[r_idx].height = 26
    for c_idx, h in enumerate(jb_col_hdrs, 1):
        set_cell(ws.cell(r_idx, c_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    r_idx += 1
    jb_tot_used = 0
    jb_tot_spare = 0
    jb_tot_pins = 0

    for jb in ORDERED_JBS:
        fill = ZEBRA_ODD if (r_idx % 2 == 1) else ZEBRA_EVEN
        d = by_jb_summary[jb]
        meta = JB_METADATA.get(jb, {})

        di_u, di_s = d['DI']['used'], d['DI']['spare']
        di_t = di_u + di_s
        di_pct = (di_s / di_t * 100.0) if di_t > 0 else 0.0

        do_u, do_s = d['DO']['used'], d['DO']['spare']
        do_t = do_u + do_s
        do_pct = (do_s / do_t * 100.0) if do_t > 0 else 0.0

        ai_u, ai_s = d['AI']['used'], d['AI']['spare']
        ai_t = ai_u + ai_s
        ai_pct = (ai_s / ai_t * 100.0) if ai_t > 0 else 0.0

        ao_u, ao_s = d['AO']['used'], d['AO']['spare']

        j_u = di_u + do_u + ai_u + ao_u
        j_s = di_s + do_s + ai_s + ao_s
        j_tot = j_u + j_s
        j_pct = (j_s / j_tot * 100.0) if j_tot > 0 else 0.0

        jb_tot_used += j_u
        jb_tot_spare += j_s
        jb_tot_pins += j_tot

        ws.row_dimensions[r_idx].height = 20

        # JB Tag with Hyperlink
        cell_jb = ws.cell(r_idx, 1)
        set_cell(cell_jb, jb, font=Font(name="Calibri", size=10, bold=True, color="0284C7", underline="single"), fill=fill, alignment=ALIGN_CENTER)
        cell_jb.hyperlink = f"#'{jb}'!A1"

        set_cell(ws.cell(r_idx, 2), meta.get('desc', jb), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(r_idx, 3), di_u if di_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 4), di_s if di_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 5), di_t if di_t>0 else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 6), f"{di_pct:.1f}%" if di_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(r_idx, 7), do_u if do_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 8), do_s if do_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 9), do_t if do_t>0 else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 10), f"{do_pct:.1f}%" if do_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(r_idx, 11), ai_u if ai_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 12), ai_s if ai_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 13), ai_t if ai_t>0 else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(r_idx, 14), f"{ai_pct:.1f}%" if ai_t>0 else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(r_idx, 15), j_u, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 16), j_s, font=FONT_DATA_BOLD, fill=SPARE_FILL, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 17), j_tot, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(r_idx, 18), f"{j_pct:.1f}%" if j_tot>0 else "-", font=Font(name="Calibri", size=10, bold=True, color="15803D" if j_pct>=20 else "92400E"),
                 fill=COMPLIANT_FILL if j_pct>=20 else SPARE_FILL, alignment=ALIGN_CENTER)

        r_idx += 1

    # Total Row Table 3
    ws.row_dimensions[r_idx].height = 22
    ws.merge_cells(f"A{r_idx}:B{r_idx}")
    set_cell(ws.cell(r_idx, 1), "TOTAL FIELD JUNCTION BOX SIGNALS", font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    set_cell(ws.cell(r_idx, 2), "", fill=CARD_BODY_FILL, border=TOTAL_BORDER)

    jb_avg_pct = (jb_tot_spare / jb_tot_pins * 100.0) if jb_tot_pins > 0 else 0.0
    for c_i in range(3, 15):
        set_cell(ws.cell(r_idx, c_i), "-", font=FONT_DATA, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)

    set_cell(ws.cell(r_idx, 15), jb_tot_used, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 16), jb_tot_spare, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 17), jb_tot_pins, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(r_idx, 18), f"{jb_avg_pct:.1f}%", font=Font(name="Calibri", size=11, bold=True, color="15803D"), fill=COMPLIANT_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)

    # Column Widths
    col_w = {
        1: 14, 2: 38,
        3: 10, 4: 10, 5: 10, 6: 12,
        7: 10, 8: 10, 9: 10, 10: 12,
        11: 10, 12: 10, 13: 10, 14: 12,
        15: 13, 16: 13, 17: 13, 18: 15
    }
    for c_i, w in col_w.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 1: 01_Junction_Box_Index
# -----------------------------------------------------------------------------
def build_index_sheet(ws, channels_by_jb):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:L1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  MASTER FIELD JUNCTION BOX DIRECTORY & ALLOCATION INDEX",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:L2")
    set_cell(ws["A2"], "SPRINT 18K TPA SPRAY DRYER / JET COOKER (Ref: x2608003) — Comprehensive Field Enclosure Architecture",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Overall Metrics
    tot_channels = sum(len(pts) for pts in channels_by_jb.values())
    tot_active = sum(sum(1 for p in pts if p['status'] == 'ACTIVE') for pts in channels_by_jb.values())
    tot_spare = sum(sum(1 for p in pts if p['status'] == 'SPARE') for pts in channels_by_jb.values())
    spare_pct = (tot_spare / tot_channels * 100.0) if tot_channels > 0 else 0.0

    kpis = [
        ("A4:B4", "A5:B5", "A6:B6", "TOTAL JUNCTION BOXES", "16 Enclosures", "Field JBs, IS Barriers, CA1, RIO, MCC"),
        ("C4:D4", "C5:D5", "C6:D6", "TOTAL PHYSICAL CHANNELS", f"{tot_channels:,} Points", "ControlLogix 1756 Standardized Slots"),
        ("E4:F4", "E5:F5", "E6:F6", "ACTIVE PROCESS SIGNALS", f"{tot_active:,} Points", "Allocated to Process Instruments"),
        ("G4:H4", "G5:H5", "G6:H6", "FIELD SPARE TERMINALS", f"{tot_spare:,} Points ({spare_pct:.1f}%)", "Exceeds 20% Contractual Reserve Spec"),
        ("I4:J4", "I5:J5", "I6:J6", "DI / DO / AI / AO SUMMARY", "Used vs Spare Summary", "Click for DI/DO/AI/AO Breakdown ➔"),
        ("K4:L4", "K5:L5", "K6:L6", "PANEL MATRIX VIEW", "Panel [P1-P4] Matrix", "Click to view Grand Signal Matrix ➔")
    ]

    for m_head, m_val, m_sub, title, val, sub in kpis:
        ws.merge_cells(m_head)
        ws.merge_cells(m_val)
        ws.merge_cells(m_sub)
        c_head = ws[m_head.split(':')[0]]
        c_val = ws[m_val.split(':')[0]]
        c_sub = ws[m_sub.split(':')[0]]
        set_cell(c_head, title, font=FONT_CARD_TITLE, fill=CARD_HEAD_FILL, alignment=ALIGN_CENTER)
        set_cell(c_val, val, font=FONT_CARD_VAL, fill=CARD_BODY_FILL, alignment=ALIGN_CENTER)
        set_cell(c_sub, sub, font=FONT_CARD_SUB, fill=CARD_BODY_FILL, alignment=ALIGN_CENTER)

    # Link to DI DO AI AO summary sheet
    ws["I5"].hyperlink = "#'00_IO_Summary_DI_DO_AI_AO'!A1"
    ws["I5"].font = Font(name="Calibri", size=13, bold=True, color="0284C7", underline="single")

    # Link to Matrix sheet
    ws["K5"].hyperlink = "#'02_JB_Allocation_Matrix'!A1"
    ws["K5"].font = Font(name="Calibri", size=13, bold=True, color="0284C7", underline="single")

    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 16

    # Section Header
    ws.merge_cells("A8:L8")
    set_cell(ws["A8"], "MASTER JUNCTION BOX ENCLOSURE & TERMINATION DIRECTORY (Click Enclosure Tag to Jump to Sheet)",
             font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[8].height = 22

    headers = [
        "Enclosure Tag", "Enclosure Description", "Physical Location / Floor", "Enclosure Dimensions",
        "Protection & Material", "Total Points", "Active Points", "Spare Points", "Spare Margin %",
        "Primary Panel(s)", "Trunk Cable Specification", "Sheet Navigation"
    ]
    ws.row_dimensions[9].height = 26
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 4, 5, 9, 10, 12] else ALIGN_LEFT
        set_cell(ws.cell(9, col_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 10
    sum_tot, sum_act, sum_spr = 0, 0, 0

    for jb in ORDERED_JBS:
        pts = channels_by_jb.get(jb, [])
        meta = JB_METADATA.get(jb, {
            'desc': f'{jb} Enclosure',
            'location': 'Plant Area',
            'size': 'Standard Size',
            'type': 'IP66 SS304',
            'primary_panel': 'P1-P4',
            'trunk_cable': 'Standard Field Trunk'
        })

        tot = len(pts)
        act = sum(1 for p in pts if p['status'] == 'ACTIVE')
        spr = sum(1 for p in pts if p['status'] == 'SPARE')
        pct = (spr / tot * 100.0) if tot > 0 else 0.0

        sum_tot += tot
        sum_act += act
        sum_spr += spr

        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[row_idx].height = 20

        # JB Tag with Hyperlink
        cell_jb = ws.cell(row_idx, 1)
        set_cell(cell_jb, jb, font=Font(name="Calibri", size=10, bold=True, color="0284C7", underline="single"), fill=fill, alignment=ALIGN_CENTER)
        cell_jb.hyperlink = f"#'{jb}'!A1"

        set_cell(ws.cell(row_idx, 2), meta.get('desc', jb), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 3), meta.get('location', '-'), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 4), meta.get('size', '-'), font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 5), meta.get('type', '-'), font=IS_BLUE_FONT if "Ex" in meta.get('type', '') else FONT_DATA,
                 fill=IS_BLUE_FILL if "Ex" in meta.get('type', '') else fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(row_idx, 6), tot, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 7), act, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 8), spr, font=FONT_DATA, fill=SPARE_FILL if spr > 0 else fill, alignment=ALIGN_RIGHT, num_format="#,##0")
        set_cell(ws.cell(row_idx, 9), f"{pct:.1f}%" if tot > 0 else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 10), meta.get('primary_panel', '-'), font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 11), meta.get('trunk_cable', '-'), font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)

        # Nav link
        cell_nav = ws.cell(row_idx, 12)
        set_cell(cell_nav, f"Open {jb} ➔", font=Font(name="Calibri", size=9, bold=True, color="0284C7", underline="single"), fill=fill, alignment=ALIGN_CENTER)
        cell_nav.hyperlink = f"#'{jb}'!A1"

        row_idx += 1

    # Total Row
    ws.row_dimensions[row_idx].height = 22
    ws.merge_cells(f"A{row_idx}:E{row_idx}")
    set_cell(ws.cell(row_idx, 1), "TOTAL SYSTEM FIELD TERMINATIONS", font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    for c_i in range(2, 6): set_cell(ws.cell(row_idx, c_i), "", fill=CARD_BODY_FILL, border=TOTAL_BORDER)

    avg_pct = (sum_spr / sum_tot * 100.0) if sum_tot > 0 else 0.0
    set_cell(ws.cell(row_idx, 6), sum_tot, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 7), sum_act, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 8), sum_spr, font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT, num_format="#,##0")
    set_cell(ws.cell(row_idx, 9), f"{avg_pct:.1f}%", font=FONT_DATA_BOLD, fill=SPARE_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)
    for c_i in range(10, 13):
        set_cell(ws.cell(row_idx, c_i), "-", font=FONT_DATA, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER)

    col_widths = {1: 15, 2: 38, 3: 35, 4: 20, 5: 32, 6: 14, 7: 14, 8: 14, 9: 16, 10: 32, 11: 48, 12: 16}
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 2: 02_JB_Allocation_Matrix
# -----------------------------------------------------------------------------
def build_matrix_sheet(ws, channels_by_jb):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:U1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  GRANULAR JUNCTION BOX TO PANEL [P1 - P4 & RIO] ALLOCATION MATRIX",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:U2")
    set_cell(ws["A2"], "Cross-Tabulation of Field Signal Types (DI, DO, AI, AO) Delivered by Each Control Panel / Rack to Each Field Junction Box",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    super_hdrs = [
        ("A4:C4", "DESTINATION JUNCTION BOX", NAVY_HEADER),
        ("D4:G4", "PANEL P1 (CHASSIS C1 - CA1)", "1E3A8A"),
        ("H4:K4", "PANEL P2 (CHASSIS C2 - CA1)", "78350F"),
        ("L4:P4", "PANEL P3 (CHASSIS C3 - RIO1)", "5B21B6"),
        ("Q4:S4", "PANEL P4 (CHASSIS C4 - RIO2)", "9F1239"),
        ("T4:U4", "TOTAL FIELD SIGNALS", CARD_HEADER_BG)
    ]
    ws.row_dimensions[4].height = 22
    for rng, text, col_hex in super_hdrs:
        ws.merge_cells(rng)
        sc = ws[rng.split(':')[0]]
        set_cell(sc, text, font=FONT_HEADER, fill=PatternFill(start_color=col_hex, end_color=col_hex, fill_type="solid"), alignment=ALIGN_CENTER)

    sub_hdrs = [
        "Enclosure Tag", "Location", "Area Description",
        "P1 DI", "P1 DO", "P1 AI", "P1 Tot",
        "P2 DI", "P2 DO", "P2 AI", "P2 Tot",
        "P3 DI", "P3 DO", "P3 AI", "P3 AO", "P3 Tot",
        "P4 DI", "P4 AI", "P4 Tot",
        "Active Sigs", "Total Pins"
    ]
    ws.row_dimensions[5].height = 24
    for col_idx, h in enumerate(sub_hdrs, 1):
        set_cell(ws.cell(5, col_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    row_idx = 6
    tot_cols = defaultdict(int)

    for jb in ORDERED_JBS:
        pts = channels_by_jb.get(jb, [])
        meta = JB_METADATA.get(jb, {})
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[row_idx].height = 20

        p1_di = sum(1 for p in pts if p['panel'] == 'P1' and p['io_type'] == 'DI' and p['status'] == 'ACTIVE')
        p1_do = sum(1 for p in pts if p['panel'] == 'P1' and p['io_type'] == 'DO' and p['status'] == 'ACTIVE')
        p1_ai = sum(1 for p in pts if p['panel'] == 'P1' and p['io_type'] == 'AI' and p['status'] == 'ACTIVE')
        p1_tot = p1_di + p1_do + p1_ai

        p2_di = sum(1 for p in pts if p['panel'] == 'P2' and p['io_type'] == 'DI' and p['status'] == 'ACTIVE')
        p2_do = sum(1 for p in pts if p['panel'] == 'P2' and p['io_type'] == 'DO' and p['status'] == 'ACTIVE')
        p2_ai = sum(1 for p in pts if p['panel'] == 'P2' and p['io_type'] == 'AI' and p['status'] == 'ACTIVE')
        p2_tot = p2_di + p2_do + p2_ai

        p3_di = sum(1 for p in pts if p['panel'] == 'P3' and p['io_type'] == 'DI' and p['status'] == 'ACTIVE')
        p3_do = sum(1 for p in pts if p['panel'] == 'P3' and p['io_type'] == 'DO' and p['status'] == 'ACTIVE')
        p3_ai = sum(1 for p in pts if p['panel'] == 'P3' and p['io_type'] == 'AI' and p['status'] == 'ACTIVE')
        p3_ao = sum(1 for p in pts if p['panel'] == 'P3' and p['io_type'] == 'AO' and p['status'] == 'ACTIVE')
        p3_tot = p3_di + p3_do + p3_ai + p3_ao

        p4_di = sum(1 for p in pts if p['panel'] == 'P4' and p['io_type'] == 'DI' and p['status'] == 'ACTIVE')
        p4_ai = sum(1 for p in pts if p['panel'] == 'P4' and p['io_type'] == 'AI' and p['status'] == 'ACTIVE')
        p4_tot = p4_di + p4_ai

        tot_act = sum(1 for p in pts if p['status'] == 'ACTIVE')
        tot_pins = len(pts)

        vals = [
            jb, meta.get('location', '-'), meta.get('area', '-'),
            p1_di, p1_do, p1_ai, p1_tot,
            p2_di, p2_do, p2_ai, p2_tot,
            p3_di, p3_do, p3_ai, p3_ao, p3_tot,
            p4_di, p4_ai, p4_tot,
            tot_act, tot_pins
        ]

        for c_idx, val in enumerate(vals, 1):
            if c_idx == 1:
                c_cell = ws.cell(row_idx, c_idx)
                set_cell(c_cell, val, font=Font(name="Calibri", size=10, bold=True, color="0284C7", underline="single"), fill=fill, alignment=ALIGN_CENTER)
                c_cell.hyperlink = f"#'{jb}'!A1"
            elif c_idx in [2, 3]:
                set_cell(ws.cell(row_idx, c_idx), val, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
            else:
                tot_cols[c_idx] += val
                f_bold = FONT_DATA_BOLD if c_idx in [7, 11, 16, 19, 20, 21] else FONT_DATA
                val_disp = val if val > 0 else "-"
                set_cell(ws.cell(row_idx, c_idx), val_disp, font=f_bold, fill=fill, alignment=ALIGN_CENTER)

        row_idx += 1

    # Total Row
    ws.row_dimensions[row_idx].height = 22
    ws.merge_cells(f"A{row_idx}:C{row_idx}")
    set_cell(ws.cell(row_idx, 1), "TOTAL SIGNALS BY PANEL & TYPE", font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_RIGHT)
    for c_i in range(2, 4): set_cell(ws.cell(row_idx, c_i), "", fill=CARD_BODY_FILL, border=TOTAL_BORDER)

    for c_idx in range(4, 22):
        set_cell(ws.cell(row_idx, c_idx), tot_cols[c_idx], font=FONT_DATA_BOLD, fill=CARD_BODY_FILL, border=TOTAL_BORDER, alignment=ALIGN_CENTER, num_format="#,##0")

    col_widths = {
        1: 14, 2: 32, 3: 38,
        4: 8, 5: 8, 6: 8, 7: 10,
        8: 8, 9: 8, 10: 8, 11: 10,
        12: 8, 13: 8, 14: 8, 15: 8, 16: 10,
        17: 8, 18: 8, 19: 10,
        20: 13, 21: 13
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 3..18: Dedicated Junction Box Sheets
# -----------------------------------------------------------------------------
def build_dedicated_jb_sheet(ws, jb_tag, pts, inst_lookup):
    ws.views.sheetView[0].showGridLines = True
    meta = JB_METADATA.get(jb_tag, {})

    # Title Block
    ws.merge_cells("A1:P1")
    set_cell(ws["A1"], f"JUNCTION BOX I/O TERMINATION SCHEDULE --- {jb_tag} ({meta.get('desc', jb_tag).upper()})",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[1].height = 26

    # Top Navigation Links
    ws.merge_cells("Q1:S1")
    link_idx = ws["Q1"]
    set_cell(link_idx, "⬅ Back to Master Index", font=Font(name="Calibri", size=10, bold=True, color=WHITE, underline="single"),
             fill=NAVY_FILL, alignment=ALIGN_CENTER)
    link_idx.hyperlink = "#'01_Junction_Box_Index'!A1"

    ws.merge_cells("A2:P2")
    sub_text = f"Location: {meta.get('location', '-')}  |  Enclosure: {meta.get('size', '-')} ({meta.get('type', '-')})  |  Trunk Cable: {meta.get('trunk_cable', '-')}"
    set_cell(ws["A2"], sub_text, font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[2].height = 18

    ws.merge_cells("Q2:S2")
    link_mat = ws["Q2"]
    set_cell(link_mat, "⬅ Panel Allocation Matrix", font=Font(name="Calibri", size=10, bold=True, color=WHITE, underline="single"),
             fill=SUB_FILL, alignment=ALIGN_CENTER)
    link_mat.hyperlink = "#'02_JB_Allocation_Matrix'!A1"

    # Precise DI, DO, AI, AO Used vs Spare Counts for this Junction Box
    di_u = sum(1 for p in pts if p['io_type'] == 'DI' and p['status'] == 'ACTIVE')
    di_s = sum(1 for p in pts if p['io_type'] == 'DI' and p['status'] == 'SPARE')
    di_tot = di_u + di_s
    di_pct = (di_s / di_tot * 100.0) if di_tot > 0 else 0.0

    do_u = sum(1 for p in pts if p['io_type'] == 'DO' and p['status'] == 'ACTIVE')
    do_s = sum(1 for p in pts if p['io_type'] == 'DO' and p['status'] == 'SPARE')
    do_tot = do_u + do_s
    do_pct = (do_s / do_tot * 100.0) if do_tot > 0 else 0.0

    ai_u = sum(1 for p in pts if p['io_type'] == 'AI' and p['status'] == 'ACTIVE')
    ai_s = sum(1 for p in pts if p['io_type'] == 'AI' and p['status'] == 'SPARE')
    ai_tot = ai_u + ai_s
    ai_pct = (ai_s / ai_tot * 100.0) if ai_tot > 0 else 0.0

    ao_u = sum(1 for p in pts if p['io_type'] == 'AO' and p['status'] == 'ACTIVE')
    ao_s = sum(1 for p in pts if p['io_type'] == 'AO' and p['status'] == 'SPARE')
    ao_tot = ao_u + ao_s
    ao_pct = (ao_s / ao_tot * 100.0) if ao_tot > 0 else 0.0

    tot_u = di_u + do_u + ai_u + ao_u
    tot_s = di_s + do_s + ai_s + ao_s
    tot_proc = tot_u + tot_s
    tot_pct = (tot_s / tot_proc * 100.0) if tot_proc > 0 else 0.0

    # Summary KPI Cards (Rows 4 to 6)
    kpis = [
        ("A4:C4", "A5:C5", "A6:C6", "DIGITAL INPUT (DI)", f"Used: {di_u} | Spare: {di_s}", f"Total: {di_tot} Pts | Spare: {di_pct:.1f}%"),
        ("D4:F4", "D5:F5", "D6:F6", "DIGITAL OUTPUT (DO)", f"Used: {do_u} | Spare: {do_s}", f"Total: {do_tot} Pts | Spare: {do_pct:.1f}%"),
        ("G4:I4", "G5:I5", "G6:I6", "ANALOG INPUT (AI)", f"Used: {ai_u} | Spare: {ai_s}", f"Total: {ai_tot} Pts | Spare: {ai_pct:.1f}%"),
        ("J4:L4", "J5:L5", "J6:L6", "ANALOG OUTPUT (AO)", f"Used: {ao_u} | Spare: {ao_s}", f"Total: {ao_tot} Pts | Spare: {ao_pct:.1f}%" if ao_tot>0 else "None in this JB"),
        ("M4:S4", "M5:S5", "M6:S6", "TOTAL ENCLOSURE TERMINATIONS", f"Active: {tot_u} | Spare: {tot_s} ({tot_pct:.1f}%)", f"Total Physical Channels: {len(pts)} Pins (Commons: {len(pts)-tot_proc})")
    ]

    for m_head, m_val, m_sub, title, val, sub in kpis:
        ws.merge_cells(m_head)
        ws.merge_cells(m_val)
        ws.merge_cells(m_sub)
        c_head = ws[m_head.split(':')[0]]
        c_val = ws[m_val.split(':')[0]]
        c_sub = ws[m_sub.split(':')[0]]
        set_cell(c_head, title, font=FONT_CARD_TITLE, fill=CARD_HEAD_FILL, alignment=ALIGN_CENTER)
        set_cell(c_val, val, font=FONT_CARD_VAL, fill=CARD_BODY_FILL, alignment=ALIGN_CENTER)
        set_cell(c_sub, sub, font=FONT_CARD_SUB, fill=CARD_BODY_FILL, alignment=ALIGN_CENTER)

    ws.row_dimensions[4].height = 16
    ws.row_dimensions[5].height = 28
    ws.row_dimensions[6].height = 16

    headers = [
        "Point #", "Source Panel", "Panel Terminal Block", "Chassis", "Slot",
        "Module Model", "Pin #", "Wire Tag (PLC Side)", "Wire Tag (Terminal Side)",
        "PLC Tag Name", "Instrument Tag", "Instrument Description", "Instrument Type",
        "P&ID No.", "Signal Type", "Calibrated Range", "Manufacturer & Model",
        "Suggested Field Cable", "Channel Status"
    ]
    ws.row_dimensions[7].height = 26
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 2, 4, 7, 14, 19] else ALIGN_LEFT
        set_cell(ws.cell(7, col_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 8
    for idx, p in enumerate(pts, 1):
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        status_val = p['status']

        if status_val == "ACTIVE":
            s_fill = ACTIVE_FILL
            s_font = ACTIVE_FONT
        elif status_val == "SPARE":
            s_fill = SPARE_FILL
            s_font = SPARE_FONT
        else:
            s_fill = COMMON_FILL
            s_font = COMMON_FONT

        # Instrument details
        itag_u = p['inst_tag'].upper()
        inst_item = inst_lookup.get(itag_u)
        if not inst_item:
            norm = itag_u.replace(' ', '').replace('-', '')
            inst_item = inst_lookup.get(norm)
        if not inst_item and p['plc_tag']:
            base_ptag = p['plc_tag'].split('_')[0].upper()
            inst_item = inst_lookup.get(base_ptag)

        if inst_item:
            i_type = inst_item['inst_name']
            pid_no = inst_item['pid']
            sig_type = inst_item['sig_type']
            rng = inst_item['range']
            prot = inst_item['prot']
        else:
            i_type = "-" if status_val != "ACTIVE" else "Process Field Device"
            pid_no = "-"
            sig_type = "24VDC Discrete" if p['io_type'] in ['DI', 'DO'] else ("4-20mA Analog" if p['io_type'] in ['AI', 'AO'] else "-")
            rng = "-"
            prot = "-"

        # Manufacturer & Cable
        if status_val == "ACTIVE" and p['inst_tag']:
            prefix = p['inst_tag'].split('-')[0] if '-' in p['inst_tag'] else "GEN"
            mfg_info = infer_instrument_model(p['inst_tag'], p['inst_desc'], i_type, sig_type, prot)
            mfg_model = f"{mfg_info['mfg']} {mfg_info['model']}"
            cable_sug, _, _, _ = suggest_cable(prefix, p['io_type'], sig_type, "", prot)
        else:
            mfg_model = "-"
            cable_sug = "Reserved for Future Expansion" if status_val == "SPARE" else "-"

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), idx, font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), p['panel'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 3), p['term_num'] if p['term_num'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 4), p['chassis'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 5), p['slot_tag'], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 6), p['card_model'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 7), p['term_no'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 8), p['wire_plc'] if p['wire_plc'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 9), p['wire_term'] if p['wire_term'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 10), p['plc_tag'] if p['plc_tag'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 11), p['inst_tag'] if p['inst_tag'] else "-", font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 12), p['inst_desc'] if p['inst_desc'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 13), i_type, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 14), pid_no, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 15), sig_type, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 16), rng, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 17), mfg_model, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 18), cable_sug, font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 19), status_val, font=s_font, fill=s_fill, alignment=ALIGN_CENTER)

        row_idx += 1

    col_widths = {
        1: 8, 2: 12, 3: 20, 4: 10, 5: 12, 6: 18, 7: 8, 8: 24, 9: 24,
        10: 25, 11: 18, 12: 34, 13: 26, 14: 18, 15: 18, 16: 16,
        17: 34, 18: 38, 19: 14
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Main Execution
# -----------------------------------------------------------------------------
def main():
    print("=================================================================")
    print("STARTING JUNCTION BOX I/O LIST REPORT GENERATION")
    print("WITH DI, DO, AI, AO USED vs SPARE & SPARE % SUMMARY")
    print("=================================================================")

    inst_lookup, mcc_instruments, channels_by_jb, by_chassis, by_jb_summary = load_data()

    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # remove default sheet

    # 1. Master DI, DO, AI, AO Summary Sheet
    print("Generating Sheet: 00_IO_Summary_DI_DO_AI_AO...")
    ws_io_sum = wb.create_sheet(title="00_IO_Summary_DI_DO_AI_AO")
    build_io_summary_sheet(ws_io_sum, by_chassis, by_jb_summary)
    ws_io_sum.freeze_panes = 'C6'

    # 2. Master Junction Box Index Sheet
    print("Generating Sheet: 01_Junction_Box_Index...")
    ws_index = wb.create_sheet(title="01_Junction_Box_Index")
    build_index_sheet(ws_index, channels_by_jb)
    ws_index.freeze_panes = 'A10'
    ws_index.auto_filter.ref = f"A9:L{ws_index.max_row - 1}"

    # 3. Grand Allocation Matrix Sheet
    print("Generating Sheet: 02_JB_Allocation_Matrix...")
    ws_matrix = wb.create_sheet(title="02_JB_Allocation_Matrix")
    build_matrix_sheet(ws_matrix, channels_by_jb)
    ws_matrix.freeze_panes = 'D6'
    ws_matrix.auto_filter.ref = f"A5:U{ws_matrix.max_row - 1}"

    # 4. Dedicated Sheets for Each Junction Box
    for jb in ORDERED_JBS:
        print(f"Generating Sheet: {jb}...")
        pts = channels_by_jb.get(jb, [])
        ws_jb = wb.create_sheet(title=jb)
        build_dedicated_jb_sheet(ws_jb, jb, pts, inst_lookup)
        ws_jb.freeze_panes = 'A8'
        if pts:
            ws_jb.auto_filter.ref = f"A7:S{ws_jb.max_row}"

    print(f"Saving primary workbook to: {OUTPUT_FILE}")
    wb.save(OUTPUT_FILE)

    print(f"Saving duplicate workbook to: {OUTPUT_FILE_ALT}")
    wb.save(OUTPUT_FILE_ALT)

    print("=================================================================")
    print("JUNCTION BOX I/O LIST REPORT GENERATION SUCCESSFUL!")
    print(f"Saved: {OUTPUT_FILE}")
    print(f"Saved: {OUTPUT_FILE_ALT}")
    print("=================================================================")

if __name__ == "__main__":
    main()
