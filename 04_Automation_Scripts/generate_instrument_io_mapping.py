#!/usr/bin/env python3
"""
Generate Master Instrument Specification, PLC I/O Mapping, and Cable Schedule.
Ingredion (Thailand) Co., Ltd. | Project Jet Cooker | AEC Industrial Engineering

Input Files:
- 11-instrument Manual/Instrument I-O List Rev.3.6.xlsx
- IO_List_xDev-R01-Tag35-6.xlsx
- Technical Manuals and Datasheets in 11-instrument Manual/

Output File:
- Instrument_IO_Mapping_with_Models_and_Cables.xlsx
"""

import os
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
REV36_FILE = os.path.join(WORKSPACE_DIR, "11-instrument Manual", "Instrument I-O List Rev.3.6.xlsx")
DEV_FILE = os.path.join(WORKSPACE_DIR, "IO_List_xDev-R01-Tag35-6.xlsx")
OUTPUT_FILE = os.path.join(WORKSPACE_DIR, "Instrument_IO_Mapping_with_Models_and_Cables.xlsx")

# Styling Definitions
NAVY_HEADER_FILL = PatternFill(start_color="1B365D", end_color="1B365D", fill_type="solid")
SUB_HEADER_FILL = PatternFill(start_color="2D4A77", end_color="2D4A77", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
SUB_HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")

CARD_HEADER_FILL = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
ZEBRA_EVEN = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
ZEBRA_ODD = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0')
)

HEADER_BORDER = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='medium', color='0B2545'),
    bottom=Side(style='medium', color='0B2545')
)

ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')

TYPE_STYLES = {
    "DI": {"fill": PatternFill(start_color="EBF8FF", end_color="EBF8FF", fill_type="solid"),
           "font": Font(name="Calibri", size=10, bold=True, color="2B6CB0")},
    "DO": {"fill": PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid"),
           "font": Font(name="Calibri", size=10, bold=True, color="B7791F")},
    "AI": {"fill": PatternFill(start_color="F0FFF4", end_color="F0FFF4", fill_type="solid"),
           "font": Font(name="Calibri", size=10, bold=True, color="276749")},
    "AO": {"fill": PatternFill(start_color="E6FFFA", end_color="E6FFFA", fill_type="solid"),
           "font": Font(name="Calibri", size=10, bold=True, color="234E52")},
    "BUS": {"fill": PatternFill(start_color="FAF5FF", end_color="FAF5FF", fill_type="solid"),
            "font": Font(name="Calibri", size=10, bold=True, color="553C9A")}
}

STATUS_STYLES = {
    "ACTIVE": {"fill": PatternFill(start_color="DEF7EC", end_color="DEF7EC", fill_type="solid"),
               "font": Font(name="Calibri", size=10, bold=True, color="03543F")},
    "SPARE": {"fill": PatternFill(start_color="EDF2F7", end_color="EDF2F7", fill_type="solid"),
              "font": Font(name="Calibri", size=10, bold=False, color="718096")},
    "WIRED": {"fill": PatternFill(start_color="DEF7EC", end_color="DEF7EC", fill_type="solid"),
              "font": Font(name="Calibri", size=10, bold=True, color="03543F")},
    "MCC": {"fill": PatternFill(start_color="FEFCBF", end_color="FEFCBF", fill_type="solid"),
            "font": Font(name="Calibri", size=10, bold=True, color="B7791F")},
    "LOCAL": {"fill": PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid"),
              "font": Font(name="Calibri", size=10, bold=False, color="6B7280")}
}

def set_cell(cell, value, font=None, fill=None, alignment=None, border=None, is_text=False):
    if value is not None and isinstance(value, str) and (value.startswith("=") or is_text):
        cell.value = value
        cell.data_type = 's'
    else:
        cell.value = value
    if font:
        cell.font = font
    if fill:
        cell.fill = fill
    if alignment:
        cell.alignment = alignment
    if border:
        cell.border = border

def clean_val(val):
    if pd.isna(val):
        return "-"
    s = str(val).strip()
    if s.lower() in ["nan", "none", "", "#n/a"]:
        return "-"
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    return s

SPECIFIC_MODELS = {
    "FT-40201": ("Endress+Hauser", "Promass E 300 (Coriolis Mass Flow)", "8E3B50-14000/0", "DN50 2\" Cl.150 RF Flange", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-40202": ("Endress+Hauser", "Promag P 300 (Electromagnetic Flow)", "5P3B50-355V8/0", "DN50 2\" Cl.150 RF Flange, PTFE Liner", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-60201": ("Endress+Hauser", "t-mass I 300 (Thermal Mass Flow)", "6I3BL3-83L4/0", "1\" NPT Insertion Sensor", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-60210": ("Endress+Hauser", "t-mass F 300 (Thermal Mass Flow)", "6F3B1H-3ND2/0", "DN100 4\" Cl.150 RF Flange", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-60211": ("Endress+Hauser", "t-mass F 300 (Thermal Mass Flow)", "6F3B1H-3ND2/0", "DN100 4\" Cl.150 RF Flange", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-60221": ("Endress+Hauser", "t-mass I 300 (Thermal Mass Flow)", "6I3BL3-83L4/0", "1\" NPT Insertion Sensor", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-60307": ("Endress+Hauser", "t-mass F 300 (Thermal Mass Flow)", "6F3B1H-3ND2/0", "DN100 4\" Cl.150 RF Flange", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-60901": ("Endress+Hauser", "Prowirl F 200 (Vortex Steam Flow)", "7F2C1H-2EPA8/0", "DN100 4\" Cl.300 RF Flange", "4-20mA HART + Pulse / 24VDC", "Non-Ex / Safe Area"),
    "FT-60902": ("Endress+Hauser", "Prowirl F 200 (Vortex Steam Flow)", "7F2C80-2JQJ7/0", "DN80 3\" Cl.300 RF Flange", "4-20mA HART + Pulse / 24VDC", "Non-Ex / Safe Area"),
    "FT-61101": ("Endress+Hauser", "t-mass F 300 (Thermal Mass Flow)", "6F3B1H-3ND2/0", "DN100 4\" Cl.150 RF Flange", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-90301": ("Endress+Hauser", "Promag P 300 (Electromagnetic Flow)", "5P3B1H-286C7/0", "DN100 4\" Cl.150 RF Flange, PTFE Liner", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FT-91001": ("Endress+Hauser", "Promass E 300 (Coriolis Mass Flow)", "8E3B50-1RP84/0", "DN50 2\" Cl.150 RF Flange", "4-20mA HART / 24VDC", "ATEX/IECEx Ex d IIC/IIIC"),
    "FT-93001": ("Endress+Hauser", "Promag P 300 (Electromagnetic Flow)", "5P3B80-2KAM5/0", "DN80 3\" Cl.150 RF Flange, PTFE Liner", "4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "FI-80301": ("Endress+Hauser", "Promass E 300 (Coriolis Mass Flow)", "8E3B50-Series", "DN50 2\" Cl.150 RF Flange (LPG Line)", "4-20mA HART / 24VDC", "ATEX/IECEx Ex d IIC"),
    "FT-80301": ("Endress+Hauser", "Promass E 300 (Coriolis Mass Flow)", "8E3B50-Series", "DN50 2\" Cl.150 RF Flange (LPG Line)", "4-20mA HART / 24VDC", "ATEX/IECEx Ex d IIC"),
    
    "DPIT-60901": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2Q8L7/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "DPT-60901": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2Q8L7/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "DPIT-61310": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2NTN7/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "ATEX/IECEx II 1/2D Ex ta/tb IIIC"),
    "DPIT-61312": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2NTN7/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "ATEX/IECEx II 1/2D Ex ta/tb IIIC"),
    "DPT-61301": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2NTN7/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "ATEX/IECEx II 1/2D Ex ta/tb IIIC"),
    "DPT-61302": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2NTN7/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "ATEX/IECEx II 1/2D Ex ta/tb IIIC"),
    "DPT-60202": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2Q8M4/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "DPT-60203": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-2Q8M4/0", "NPT 1/4\" w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    "DPT-40201": ("Endress+Hauser", "Deltabar PMD55B (Diff. Pressure)", "PMD55B-Series", "NPT 1/4\" w/ 5-Valve Manifold", "2-wire 4-20mA HART / 24VDC", "Non-Ex / Safe Area"),
    
    "AT-60241": ("ENVEA / PCME", "PCME QAL 991 (Dust Analyzer)", "6626013-33-201-001", "Flange DN100 w/ Purge Air", "4-20mA Output + Alarm Relay / 24VDC", "Ex t/d / Safe Area"),
    "pHT-40201": ("Mettler Toledo", "M300G2 + InPro3250i (pH Analyzer)", "M300G2 4-Wire + InTrac 777P", "DN25 Weld-in Socket / InTrac 777P", "4-20mA HART + Relay / 24VDC", "Non-Ex / Safe Area"),
    "pHT-40202": ("Mettler Toledo", "M300G2 + InPro3250i (pH Analyzer)", "M300G2 4-Wire + InTrac 777P", "DN25 Weld-in Socket / InTrac 777P", "4-20mA HART + Relay / 24VDC", "Non-Ex / Safe Area"),
    "AT-40101": ("Mettler Toledo", "M300G2 + InPro3250i (pH Analyzer)", "M300G2 4-Wire + InTrac 777P", "DN25 Weld-in Socket / InTrac 777P", "4-20mA HART + Relay / 24VDC", "Non-Ex / Safe Area"),
    "AT-40102": ("Mettler Toledo", "M300G2 + InPro3250i (pH Analyzer)", "M300G2 4-Wire + InTrac 777P", "DN25 Weld-in Socket / InTrac 777P", "4-20mA HART + Relay / 24VDC", "Non-Ex / Safe Area"),
    
    "PT-80302": ("Endress+Hauser", "Cerabar PMP51B (Pressure Transmitter)", "PMP51B-Series", "1/2\" MNPT Threaded", "2-wire 4-20mA HART / 24VDC", "ATEX/IECEx Ex d IIC"),
}

def infer_instrument_model(tag, desc, inst_name, signal_type, cable_type, prot_type):
    tag_clean = str(tag).strip().upper()
    if tag_clean in SPECIFIC_MODELS:
        mfg, model, code, conn, pwr, ex_cls = SPECIFIC_MODELS[tag_clean]
        return {"mfg": mfg, "model": model, "order_code": code, "conn": conn, "power": pwr, "ex_class": ex_cls}

    prefix = tag_clean.split("-")[0] if "-" in tag_clean else tag_clean[:3]
    d_lower = str(desc).lower() + " " + str(inst_name).lower()
    
    if "ultra-sonic" in d_lower or "ultrasonic" in d_lower or "60261f" in tag_clean.lower():
        return {"mfg": "Keyence", "model": "FD-H Series (Clamp-on Ultrasonic Flowmeter)", "order_code": "FD-H20 / FD-H32 / FD-HF",
                "conn": "Clamp-on Sensor Head (External Pipe Mount)", "power": "4-20mA + Pulse / 24VDC", "ex_class": "IP67 Industrial / Safe Area"}
    
    if prefix == "FT" and ("coriolis" in d_lower or "mass" in d_lower or "slurry" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Promass E 300 (Coriolis Mass Flowmeter)", "order_code": "8E3B50 Series",
                "conn": "Flange ASME B16.5 Cl.150 RF, 316L", "power": "4-20mA HART / 24VDC Loop/Aux", "ex_class": "Ex d IIC / Safe Area"}
        
    if prefix == "FT" and ("magnetic" in d_lower or "mag" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Promag P 300 (Electromagnetic Flowmeter)", "order_code": "5P3B Series",
                "conn": "Flange ASME B16.5 Cl.150 RF, PTFE Liner", "power": "4-20mA HART / 24VDC", "ex_class": "Non-Ex / Safe Area"}

    if prefix == "FT" and ("vortex" in d_lower or "steam" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "Prowirl F 200 (Vortex Flowmeter)", "order_code": "7F2C Series",
                "conn": "Flange ASME B16.5 Cl.300 RF, 316L", "power": "4-20mA HART + Pulse / 24VDC", "ex_class": "Non-Ex / Safe Area"}

    if prefix == "FT" and ("thermal" in d_lower or "gas" in d_lower or "air" in d_lower):
        return {"mfg": "Endress+Hauser", "model": "t-mass F 300 / I 300 (Thermal Mass Flowmeter)", "order_code": "6F3B / 6I3B Series",
                "conn": "Flange Cl.150 or 1\" NPT Compression", "power": "4-20mA HART / 24VDC", "ex_class": "Non-Ex / Safe Area"}
        
    if prefix in ["FT", "FI", "FE"]:
        return {"mfg": "Endress+Hauser", "model": "Proline Series Flowmeter", "order_code": "Standard Flow Package",
                "conn": "Flange Cl.150 RF", "power": "4-20mA HART / 24VDC", "ex_class": "Safe Area"}

    if prefix in ["DPT", "DPIT"]:
        return {"mfg": "Endress+Hauser", "model": "Deltabar PMD55B (Diff. Pressure Transmitter)", "order_code": "PMD55B Series w/ 5-Valve Manifold",
                "conn": "1/4\" NPT Female w/ 5-Valve Manifold DA63M", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex ta/tb IIIC / Safe Area"}

    if prefix in ["PT", "PIT", "PI"]:
        return {"mfg": "Endress+Hauser", "model": "Cerabar PMP51B / PMC51B / PMP43", "order_code": "PMP51B / PMP43 Hygienic",
                "conn": "1/2\" NPT Male / Tri-Clamp Clamp", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex d IIC / Safe Area"}

    if prefix in ["PS", "PSL", "PSH"]:
        return {"mfg": "Endress+Hauser", "model": "Ceraphant PTC31B (Pressure Switch)", "order_code": "PTC31B Series",
                "conn": "1/2\" NPT Male Thread", "power": "24VDC / Dry Contact Relay / PNP", "ex_class": "Safe Area"}

    if prefix in ["DPG", "DPI"]:
        return {"mfg": "Ashcroft", "model": "Model 1132 Differential Pressure Gauge", "order_code": "1132 w/ V03 Manifold",
                "conn": "1/4\" NPT Female", "power": "Local Mechanical Indication", "ex_class": "Safe Area"}

    if prefix in ["PG"]:
        return {"mfg": "Ashcroft", "model": "Model T5500 All Stainless Steel Pressure Gauge", "order_code": "T5500 w/ 1098 Siphon",
                "conn": "1/2\" NPT Male", "power": "Local Mechanical Indication", "ex_class": "Safe Area"}

    if prefix in ["LT", "LIT"]:
        if "radar" in d_lower or "silo" in d_lower or "dryer" in d_lower:
            return {"mfg": "Endress+Hauser", "model": "Micropilot FMR67B (80GHz Radar Level)", "order_code": "FMR67B Series",
                    "conn": "Flange DN80 3\" / Cl.150", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex ia / Ex ta/tb IIIC"}
        else:
            return {"mfg": "Endress+Hauser", "model": "Cerabar PMP43 / PMP51B (Hydrostatic Level)", "order_code": "PMP43 Flush Diaphragm",
                    "conn": "Universal Clamp / Flush Mount", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Safe Area"}

    if prefix in ["LS", "LSL", "LSH"]:
        if "powder" in d_lower or "solid" in d_lower or "dryer" in d_lower or "cyclone" in d_lower or "silo" in d_lower:
            return {"mfg": "Endress+Hauser", "model": "Soliphant FTM51 (Vibrating Fork for Solids)", "order_code": "FTM51 Series",
                    "conn": "Thread 1-1/2\" NPT / Flange", "power": "24VDC / Relay Contact Output", "ex_class": "Ex ta/tb IIIC / Safe Area"}
        else:
            return {"mfg": "Endress+Hauser", "model": "Liquiphant FTL51B (Vibrating Fork for Liquids)", "order_code": "FTL51B Series",
                    "conn": "Thread 3/4\" NPT / Tri-Clamp", "power": "24VDC / Relay Contact Output", "ex_class": "Ex ia / Safe Area"}

    if prefix in ["TT", "TIT"]:
        return {"mfg": "Endress+Hauser", "model": "iTEMP TMT71 / TM411 (Pt100 RTD Transmitter)", "order_code": "TMT71 + TM411 Head Assembly",
                "conn": "Threaded 1/2\" NPT w/ Thermowell", "power": "2-wire 4-20mA HART / 24VDC", "ex_class": "Ex ia / Safe Area"}

    if prefix in ["TG", "TI"]:
        return {"mfg": "Ashcroft", "model": "Model FI Bimetal Dial Thermometer", "order_code": "FI Series w/ Threaded Thermowell",
                "conn": "1/2\" NPT Threaded Thermowell", "power": "Local Mechanical Indication", "ex_class": "Safe Area"}

    if prefix in ["TS", "TSH", "TSL"]:
        return {"mfg": "Endress+Hauser", "model": "Thermophant TTR35 (Temperature Switch)", "order_code": "TTR35 Series",
                "conn": "1/2\" NPT w/ Hygienic Adapter", "power": "24VDC / Dry Contact Relay / PNP", "ex_class": "Safe Area"}

    if prefix in ["AT", "AI", "PHT"]:
        if "ph" in d_lower:
            return {"mfg": "Mettler Toledo", "model": "M300G2 + InPro3250i (pH System)", "order_code": "M300G2 4-Wire + InTrac 777P",
                    "conn": "DN25 Weld-in Socket w/ InTrac 777P", "power": "4-20mA HART + Relay / 24VDC", "ex_class": "Safe Area"}
        elif "gas" in d_lower or "flame" in d_lower or "lel" in d_lower:
            return {"mfg": "Sensidyne / SmartGas", "model": "925FGD Explosionproof Combustible Gas Detector", "order_code": "925FGD Series",
                    "conn": "3/4\" NPT Conduit Entry", "power": "4-20mA + Relays / 24VDC", "ex_class": "ATEX/IECEx Ex d IIC"}
        else:
            return {"mfg": "ENVEA / PCME", "model": "PCME QAL 991 (Dust Analyzer)", "order_code": "6626013-33-201-001",
                    "conn": "Flange DN100 w/ Air Purge", "power": "4-20mA Output + Relay / 24VDC", "ex_class": "Ex t/d / Safe Area"}

    if prefix in ["TCV", "PCV", "FCV", "CV", "TV"]:
        return {"mfg": "Samson", "model": "Type 3241 Globe Valve + Trovis 3730-1 Positioner", "order_code": "3241-ANSI + 3730-1 Positioner",
                "conn": "Flange ASME B16.5 Cl.150/300 RF", "power": "4-20mA Setpoint + 4-20mA Feedback", "ex_class": "Ex ia IIC / Safe Area"}

    if prefix in ["XV", "SV", "BV", "ISV"]:
        return {"mfg": "TeroFox / El-O-Matic / Power-Genex", "model": "TF-20DFS Valve + F-Series Actuator + ESV Solenoid",
                "order_code": "TF-20DFS + EL-O-Matic F + ESV Solenoid", "conn": "Flange Cl.150 RF / 1/4\" Air Supply",
                "power": "24VDC Solenoid (DO) + Dry Contact LS (DI)", "ex_class": "Ex d IIC / Safe Area"}

    if prefix in ["HV"]:
        return {"mfg": "TeroFox / APL-HKC", "model": "TF-20DFS Manual Ball Valve + APL-210N/510N Switch Box",
                "order_code": "TF-20DFS + APL-210N Limit Switch", "conn": "Flange Cl.150 RF / Manual Lever",
                "power": "Dry Contact Limit Switches (2x DI)", "ex_class": "Ex d IIC / Safe Area"}

    if prefix in ["ZS"]:
        return {"mfg": "APL-HKC", "model": "APL-210N / APL-510N Valve Position Monitor", "order_code": "APL-210N / APL-510N",
                "conn": "NAMUR VDI/VDE 3845 Direct Mount", "power": "Dry Contact Micro-switches (2x SPDT)", "ex_class": "Ex d IIC / Safe Area"}

    if prefix in ["PSE"]:
        return {"mfg": "BS&B / Fike", "model": "Burst-Alert Sensor / Rupture Disk Monitor", "order_code": "Burst Disk Sensor",
                "conn": "Cl.150 Flange Holder", "power": "Dry Contact NC Loop (DI)", "ex_class": "Ex ia IIC"}

    if prefix in ["VIB", "NCT"]:
        return {"mfg": "Netter Vibration", "model": "NCT 5 Pneumatic Turbine Vibrator", "order_code": "NCT 5 Series",
                "conn": "G 1/8\" BSP Pneumatic Port", "power": "Pneumatic 2 - 6 bar / 24VDC Solenoid", "ex_class": "ATEX Zone 1/21 Safe"}

    if prefix in ["ATY", "PCY", "BFY", "RVM", "PCM", "BLM", "P", "M"]:
        return {"mfg": "Rockwell Automation", "model": "Allen-Bradley PowerFlex VSD / E300 Relay in MCC",
                "order_code": "Centerline 2500 Draw-out Bucket", "conn": "Terminal Block in MCC Compartment",
                "power": "24VDC Interlock / 380VAC 3-Phase", "ex_class": "MCC Room (Non-Hazardous)"}

    return {"mfg": "Industrial Process Equipment", "model": "Standard Process Field Device", "order_code": "-",
            "conn": "Standard Process Connection", "power": "24VDC / Dry Contact", "ex_class": "Safe Area"}

def suggest_cable(prefix, io_type, signal_type, cable_existing, prot_type):
    c_exist = str(cable_existing).strip() if pd.notna(cable_existing) and str(cable_existing).strip() not in ["-", ""] else ""
    p_upper = str(prot_type).upper()
    is_is = "EX I" in p_upper or "EX IA" in p_upper or "EX IC" in p_upper
    
    if prefix in ["XV", "SV", "BV", "ISV"]:
        rec_cable = "CVV 3Cx1.5 (Solenoid) + CVV 4Cx1.0 (Limit Switches) or CVV 5Cx1.5"
        cores_size = "3 Cores x 1.5 mm² + 4 Cores x 1.0 mm²"
        shield = "Heavy-Duty PVC Outer Sheath / Annealed Copper Conductors"
        color = "Black UV-Resistant Outer Sheath"
        gland = "M20x1.5 / 1/2\" NPT Ex d Double Compression Gland"
        trunk = "CVV 30Cx1.5 Multi-Core Control Trunk Cable to CA1"
        desc = "Independent routing for 24VDC solenoid valve and Open/Close limit switch feedback"

    elif prefix in ["TCV", "PCV", "FCV", "CV", "TV"]:
        rec_cable = "LiY-CY TP 2Px1.0 (Setpoint AO + Position Feedback AI)"
        cores_size = "2 Pairs x 1.0 mm² (Individually & Overall Shielded)"
        shield = "Individual Foil Screen per Pair + Overall Tinned Copper Braid"
        color = "Blue (RAL 5015)" if is_is else "Gray / Black"
        gland = "M20x1.5 Brass Nickel-Plated Ex d/e"
        trunk = "LiY-CY(IS/OS) 16Px0.75 Multi-Pair Trunk Cable to CA1"
        desc = "Dedicated twisted pair for 4-20mA positioner input + pair for position feedback"

    elif prefix in ["AT", "AI", "PHT"]:
        rec_cable = "LiY-CY TP 2Px1.0 (Signal + Alarm) + CVV 3Cx1.5 (24VDC Power)"
        cores_size = "2 Pairs x 1.0 mm² + 3 Cores x 1.5 mm²"
        shield = "Overall Tinned Copper Braid + Al-Mylar Screen"
        color = "Gray / Black Industrial"
        gland = "M20x1.5 Ex d/e Double Compression Gland"
        trunk = "LiY-CY 16Px0.75 (Signals) + CVV Power Trunk to CA1"
        desc = "Separate shielded cable for 4-20mA/status and 3-core power supply cable (L, N/-, PE)"

    elif prefix == "FT" and ("3p" in c_exist.lower() or "mass" in c_exist.lower() or "cvv 3c" in c_exist.lower()):
        rec_cable = "LiY-CY TP 3Px1.0 (Mass Flow, Density, Status) + CVV 3Cx1.5 (Power)"
        cores_size = "3 Pairs x 1.0 mm² + 3 Cores x 1.5 mm²"
        shield = "Individually & Overall Shielded Twisted Pairs"
        color = "Gray (RAL 7001) / Black"
        gland = "M20x1.5 Brass Nickel-Plated Ex d/e Gland"
        trunk = "LiY-CY(IS/OS) 16Px0.75 Multi-Pair Trunk Cable to CA1"
        desc = "Multi-channel Coriolis/Magmeter interface for multiple process variables + power"

    elif io_type in ["AI", "AO"] or "4 - 20" in str(signal_type):
        rec_cable = "LiY-CY(EB) 1Px1.0" if is_is else "LiY-CY TP 1Px1.0"
        cores_size = "1 Pair x 1.0 mm²"
        shield = "Al-Mylar Tape Screen + Tinned Copper Braid Shield (100% Coverage)"
        color = "Blue (RAL 5015) for Ex i" if is_is else "Gray (RAL 7001) / Black"
        gland = "M20x1.5 Polyamide Ex i (Blue)" if is_is else "M20x1.5 Brass Nickel-Plated Ex d/e"
        trunk = "LiY-CY(IS/OS) 16Px0.75 Multi-Pair Shielded Trunk Cable to CA1"
        desc = "Shielded Twisted Pair for noise immunity & HART digital protocol signaling"

    elif prefix in ["HV", "ZS"]:
        rec_cable = "CVV 4Cx1.0"
        cores_size = "4 Cores x 1.0 mm²"
        shield = "Flame-Retardant Heavy-Duty PVC Outer Sheath"
        color = "Black UV-Resistant / Gray"
        gland = "M20x1.5 / 1/2\" NPT Ex d/e Gland"
        trunk = "CVV 30Cx1.5 Multi-Core Control Trunk Cable to CA1"
        desc = "Direct connection to Open/Close SPDT limit switches inside APL switch box"

    elif prefix in ["LS", "LSH", "LSL", "PS", "PSH", "PSL", "TS", "TSH", "TSL"]:
        rec_cable = "CVV 3Cx1.5"
        cores_size = "3 Cores x 1.5 mm²"
        shield = "Heavy-Duty PVC Insulation & Outer Jacket"
        color = "Black / Blue (if IS loop)"
        gland = "M20x1.5 Ex d/e Gland"
        trunk = "CVV 30Cx1.5 Multi-Core Control Trunk Cable to CA1"
        desc = "24VDC sensor power supply + dry contact relay / PNP switching output"

    elif io_type == "BUS" or "bus" in str(signal_type).lower():
        rec_cable = "CAT 6 S/FTP Industrial Ethernet Cable / Belden 3105A RS-485"
        cores_size = "4 Pairs 24 AWG (Ethernet) / 1 Pair 22 AWG (RS-485)"
        shield = "Braided Shield + Foil Screen per Pair (S/FTP)"
        color = "Teal / Green Industrial Jacket"
        gland = "M20x1.5 Industrial RJ45 / Cable Gland"
        trunk = "Fiber Optic Backbone / Industrial Ethernet Ring to CA1"
        desc = "High-speed industrial digital communication link with noise suppression"

    elif prefix in ["ATY", "PCY", "BFY", "RVM", "PCM", "BLM", "P", "M"]:
        rec_cable = "CVV 8Cx1.5 / CVV 12Cx1.5 Control Cable"
        cores_size = "8 - 12 Cores x 1.5 mm²"
        shield = "Heavy-Duty PVC Control Cable"
        color = "Black Industrial Sheath"
        gland = "M25 / M32 Brass Industrial Gland"
        trunk = "MCC Internal Wiring Duct / Multi-Core Interlock Trunk"
        desc = "Start/Stop command (DO), Run/Trip status (DI), and VSD speed reference"

    else:
        rec_cable = c_exist if c_exist else "CVV 3Cx1.5"
        cores_size = "3 Cores x 1.5 mm²"
        shield = "Standard PVC Control Cable"
        color = "Black / Gray"
        gland = "M20x1.5 Cable Gland"
        trunk = "Multi-Core Trunk Cable to CA1 / MCC"
        desc = "Standard field instrumentation connection"

    if c_exist and c_exist != rec_cable:
        rec_cable_final = f"{rec_cable} (Ref: {c_exist})"
    else:
        rec_cable_final = rec_cable

    return {
        "cable_code": rec_cable_final,
        "cores_size": cores_size,
        "shield": shield,
        "color": color,
        "gland": gland,
        "trunk": trunk,
        "desc": desc
    }

def load_and_merge_data():
    print(f"Loading Rev 3.6 Instrument Master: {REV36_FILE}")
    df_rev36_raw = pd.read_excel(REV36_FILE, sheet_name="Rev.3", skiprows=7)
    cols = list(df_rev36_raw.columns)
    cols[6] = "DO"
    cols[7] = "DI"
    cols[8] = "AO"
    cols[9] = "AI"
    cols[10] = "Bus"
    df_rev36_raw.columns = [str(c).strip() for c in cols]
    df_rev36 = df_rev36_raw[df_rev36_raw["Item."].notna() & (df_rev36_raw["Item."].astype(str) != "Item.")].copy()
    print(f"Loaded {len(df_rev36)} instruments from Rev 3.6")

    print(f"Loading Dev PLC I/O List: {DEV_FILE}")
    df_dev_raw = pd.read_excel(DEV_FILE, sheet_name="IO List")
    df_dev = df_dev_raw[df_dev_raw["Destination"].astype(str).str.strip() != "112"].copy()
    df_dev = df_dev[df_dev["Destination"].notna()].copy()
    print(f"Loaded {len(df_dev)} PLC I/O channels from Dev master")

    dev_by_tag = {}
    for _, r in df_dev.iterrows():
        t = str(r["Instruement Tag"]).strip().upper()
        if t not in dev_by_tag:
            dev_by_tag[t] = []
        dev_by_tag[t].append(r)

    mapped_instruments = []
    for _, r in df_rev36.iterrows():
        item_no = clean_val(r.get("Item."))
        t_old = clean_val(r.get("Tag. No."))
        t_new = clean_val(r.get("New Tag. No."))
        
        active_tag = t_new if t_new != "-" and t_new else t_old
        legacy_tag = t_old if t_new != "-" and t_new else "-"
        
        pid = clean_val(r.get("P&ID No."))
        desc = clean_val(r.get("Description"))
        inst_name = clean_val(r.get("Instrument"))
        
        c_do = clean_val(r.get("DO"))
        c_di = clean_val(r.get("DI"))
        c_ao = clean_val(r.get("AO"))
        c_ai = clean_val(r.get("AI"))
        c_bus = clean_val(r.get("Bus"))
        
        sig_type = clean_val(r.get("Signal type"))
        sig_to = clean_val(r.get("Signal to"))
        range_val = clean_val(r.get("Range"))
        cable_exist = clean_val(r.get("Cable Type"))
        prot_type = clean_val(r.get("Type of Protection"))
        
        matched_pts = []
        matched_dev_tag = "-"
        
        for candidate in [t_new, t_old]:
            if candidate != "-" and candidate:
                c_up = candidate.upper()
                if c_up in dev_by_tag:
                    matched_pts = dev_by_tag[c_up]
                    matched_dev_tag = candidate
                    break
        
        if not matched_pts:
            for candidate in [t_new, t_old]:
                if candidate != "-" and candidate:
                    c_norm = candidate.upper().replace("-", "")
                    for k, v in dev_by_tag.items():
                        if k.replace("-", "") == c_norm:
                            matched_pts = v
                            matched_dev_tag = k
                            break
                    if matched_pts:
                        break

        if matched_pts:
            num_channels = len(matched_pts)
            chassis_set = sorted(set(str(p["Chassis"]).strip() for p in matched_pts if pd.notna(p["Chassis"])))
            slots_set = sorted(set(f"S{int(p['Slot']):02d}" for p in matched_pts if pd.notna(p["Slot"])))
            cards_set = sorted(set(str(p["Card"]).strip() for p in matched_pts if pd.notna(p["Card"])))
            panels_set = sorted(set(str(p[" Reference Designation"]).strip() for p in matched_pts if pd.notna(p[" Reference Designation"])))
            jbs_set = sorted(set(str(p["Destination"]).strip() for p in matched_pts if pd.notna(p["Destination"])))
            locs_set = sorted(set(str(p["Junction Box Location"]).strip() for p in matched_pts if pd.notna(p["Junction Box Location"])))
            zones_set = sorted(set(str(p["Zone"]).strip() for p in matched_pts if pd.notna(p["Zone"])))
            floors_set = sorted(set(str(p["Floor"]).strip() for p in matched_pts if pd.notna(p["Floor"])))
            
            addr_list = [f"{p['Column7']} ({p['I/O Type']})" for p in matched_pts if pd.notna(p.get("Column7"))]
            plc_tag_list = [str(p['PLC_Tag']).strip() for p in matched_pts if pd.notna(p.get("PLC_Tag"))]
            
            chassis_str = ", ".join(chassis_set)
            slots_str = ", ".join(slots_set)
            cards_str = ", ".join(cards_set)
            panels_str = ", ".join(panels_set)
            jbs_str = ", ".join(jbs_set)
            locs_str = ", ".join(locs_set)
            zones_str = ", ".join(zones_set)
            floors_str = ", ".join(floors_set)
            addr_str = "; ".join(addr_list[:4]) + ("..." if len(addr_list) > 4 else "")
            plc_tags_str = "; ".join(plc_tag_list[:3]) + ("..." if len(plc_tag_list) > 3 else "")
            status = "WIRED" if "MCC" not in panels_str else "MCC"
        else:
            num_channels = 0
            chassis_str = "-"
            slots_str = "-"
            cards_str = "-"
            panels_str = "-"
            jbs_str = "-"
            locs_str = "-"
            zones_str = "-"
            floors_str = "-"
            addr_str = "-"
            plc_tags_str = "-"
            status = "LOCAL" if sig_to.lower() in ["local", "-", ""] else "UNMAPPED"

        model_info = infer_instrument_model(active_tag, desc, inst_name, sig_type, cable_exist, prot_type)
        primary_io = "AI" if c_ai not in ["-", "0", ""] else ("DI" if c_di not in ["-", "0", ""] else ("DO" if c_do not in ["-", "0", ""] else ("AO" if c_ao not in ["-", "0", ""] else ("BUS" if c_bus not in ["-", "0", ""] else "-"))))
        prefix = active_tag.split("-")[0] if "-" in active_tag else active_tag[:3]
        cable_sug = suggest_cable(prefix, primary_io, sig_type, cable_exist, prot_type)

        mapped_instruments.append({
            "item_no": item_no,
            "active_tag": active_tag,
            "legacy_tag": legacy_tag,
            "pid": pid,
            "desc": desc,
            "inst_name": inst_name,
            "sig_type": sig_type,
            "range": range_val,
            "prot_type": prot_type,
            "mfg": model_info["mfg"],
            "model": model_info["model"],
            "order_code": model_info["order_code"],
            "conn": model_info["conn"],
            "power": model_info["power"],
            "ex_class": model_info["ex_class"],
            "c_do": c_do, "c_di": c_di, "c_ao": c_ao, "c_ai": c_ai, "c_bus": c_bus,
            "num_channels": num_channels,
            "matched_dev_tag": matched_dev_tag,
            "chassis": chassis_str,
            "slots": slots_str,
            "channel_addrs": addr_str,
            "plc_tags": plc_tags_str,
            "cards": cards_str,
            "panels": panels_str,
            "jbs": jbs_str,
            "jb_loc": locs_str,
            "zone": zones_str,
            "floor": floors_str,
            "status": status,
            "cable_sug": cable_sug["cable_code"],
            "cores_size": cable_sug["cores_size"],
            "shield": cable_sug["shield"],
            "gland": cable_sug["gland"],
            "trunk": cable_sug["trunk"],
            "cable_desc": cable_sug["desc"]
        })

    df_mapped = pd.DataFrame(mapped_instruments)
    return df_mapped, df_dev

def create_kpi_card(ws, start_row, start_col, title, value, subtitle, accent_color="1B365D"):
    accent_fill = PatternFill(start_color=accent_color, end_color=accent_color, fill_type="solid")
    set_cell(ws.cell(start_row, start_col), title, font=Font(name="Calibri", size=10, bold=True, color="FFFFFF"), fill=accent_fill, alignment=ALIGN_CENTER)
    ws.merge_cells(start_row=start_row, start_column=start_col, end_row=start_row, end_column=start_col+1)
    
    set_cell(ws.cell(start_row+1, start_col), value, font=Font(name="Calibri", size=18, bold=True, color="1E293B"), alignment=ALIGN_CENTER)
    ws.merge_cells(start_row=start_row+1, start_column=start_col, end_row=start_row+2, end_column=start_col+1)
    
    set_cell(ws.cell(start_row+3, start_col), subtitle, font=Font(name="Calibri", size=9, italic=True, color="64748B"), alignment=ALIGN_CENTER)
    ws.merge_cells(start_row=start_row+3, start_column=start_col, end_row=start_row+3, end_column=start_col+1)
    
    for r in range(start_row, start_row+4):
        for c in range(start_col, start_col+2):
            ws.cell(r, c).border = THIN_BORDER

def populate_dashboard(ws, df_mapped, df_dev):
    ws.views.sheetView[0].showGridLines = True
    
    ws.cell(1, 1, "INGREDION (THAILAND) CO., LTD.  |  PROJECT JET COOKER").font = Font(name="Calibri", size=14, bold=True, color="1B365D")
    ws.cell(2, 1, "INSTRUMENT MODEL SPECIFICATION, PLC I/O MAP & CABLE SCHEDULE").font = Font(name="Calibri", size=12, bold=True, color="475569")
    ws.cell(3, 1, "Reconciliation between Rev 3.6 Master & Dev-Tag35-6 Hardware Architecture | AEC Industrial Engineering").font = Font(name="Calibri", size=10, italic=True, color="64748B")

    tot_inst = len(df_mapped)
    wired_inst = len(df_mapped[df_mapped["status"].isin(["WIRED", "MCC"])])
    tot_io_pts = len(df_dev)
    spare_io = len(df_dev[df_dev["Instruement Tag"].astype(str).str.lower().str.contains("spare")])
    act_io = tot_io_pts - spare_io
    
    create_kpi_card(ws, 5, 1, "TOTAL INSTRUMENTS", f"{tot_inst:,}", "Rev 3.6 Engineering Master", "1B365D")
    create_kpi_card(ws, 5, 3, "WIRED TO PLC / MCC", f"{wired_inst:,}", f"{wired_inst/tot_inst*100:.1f}% electrical loops", "03543F")
    create_kpi_card(ws, 5, 5, "PLC I/O CHANNELS", f"{tot_io_pts:,}", f"{act_io:,} Active / {spare_io:,} Spare", "2B6CB0")
    create_kpi_card(ws, 5, 7, "CONTROLLOGIX RACKS", "7 Racks", "Chassis C1 to C7 (47 Slots)", "7C3AED")
    create_kpi_card(ws, 5, 9, "VENDORS MAPPED", "10+ Brands", "E+H, Samson, Mettler, etc.", "B7791F")

    cur_r = 10
    ws.cell(cur_r, 1, "1. Major Instrument Equipment Packages & Manufacturer Model Mapping").font = Font(name="Calibri", size=11, bold=True, color="1B365D")
    cur_r += 1
    
    pkg_headers = [
        "Package / Equipment Category", "Tag Prefix", "Primary Manufacturer",
        "Standard Model Series", "Process Connection", "Standard Signal", "Recommended Cable Spec"
    ]
    for c_idx, h in enumerate(pkg_headers, 1):
        set_cell(ws.cell(cur_r, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [2, 6] else ALIGN_LEFT)
                 
    cur_r += 1
    pkg_rows = [
        ("Coriolis Mass Flowmeter", "FT", "Endress+Hauser", "Promass E 300 (8E3B50)", "DN50 Cl.150 RF Flange", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Electromagnetic Flowmeter", "FT", "Endress+Hauser", "Promag P 300 (5P3B50/80/1H)", "DN50/DN80/DN100 Cl.150 Flange", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Vortex Steam Flowmeter", "FT", "Endress+Hauser", "Prowirl F 200 (7F2C80/1H)", "DN80/DN100 Cl.300 RF Flange", "4-20mA HART + Pulse", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Thermal Mass Gas Flowmeter", "FT", "Endress+Hauser", "t-mass F 300 / I 300 (6F3B/6I3B)", "DN100 Flange / 1\" NPT Insertion", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Ultrasonic Flowmeter (Clamp-on)", "FT", "Keyence", "FD-H Series (FD-H20/H32)", "Non-invasive External Clamp-on", "4-20mA + Pulse / 24VDC", "LiY-CY TP 2Px1.0 (Individually Shielded)"),
        ("Differential Pressure Transmitter", "DPT / DPIT", "Endress+Hauser", "Deltabar PMD55B Series", "1/4\" NPT w/ 5-Valve Manifold DA63M", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Gauge / Absolute Pressure Transmitter", "PT", "Endress+Hauser", "Cerabar PMP51B / PMC51B / PMP43", "1/2\" NPT / Hygienic Clamp", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Radar Level Transmitter", "LT", "Endress+Hauser", "Micropilot FMR67B (80 GHz Radar)", "DN80 3\" Cl.150 Flange", "4-20mA HART / 24VDC", "LiY-CY(EB) 1Px1.0 (Blue IS Sheath)"),
        ("Hydrostatic Level Transmitter", "LT", "Endress+Hauser", "Cerabar PMP43 (Flush Diaphragm)", "Hygienic Universal Clamp", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("Vibrating Level Switch (Liquids)", "LSH / LSL", "Endress+Hauser", "Liquiphant FTL51B", "3/4\" NPT Threaded / Clamp", "24VDC / Relay Contact", "CVV 3Cx1.5 (Control Cable)"),
        ("Vibrating Level Switch (Solids)", "LSH", "Endress+Hauser", "Soliphant FTM51", "1-1/2\" NPT Threaded / Flange", "24VDC / Relay Contact", "CVV 3Cx1.5 (Control Cable)"),
        ("Temperature Transmitter", "TT", "Endress+Hauser", "iTEMP TMT71 + TM411 (Pt100 RTD)", "1/2\" NPT w/ Threaded Thermowell", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded Twisted Pair)"),
        ("pH Analytical Transmitter", "pHT / AT", "Mettler Toledo", "M300G2 4-Wire + InPro3250i + InTrac 777P", "DN25 Weld-in Socket", "4-20mA HART / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5 (Pwr)"),
        ("Dust Analyzer Transmitter", "AT", "ENVEA / PCME", "PCME QAL 991 (Particulate Monitor)", "DN100 Flange w/ Purge Air Unit", "4-20mA + Relays / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5 (Pwr)"),
        ("Combustible / LEL Gas Detector", "AT", "Sensidyne / SmartGas", "925FGD Explosionproof Detector", "3/4\" NPT Conduit Entry", "4-20mA + Relays / 24VDC", "LiY-CY TP 2Px1.0 (Ex d Explosionproof)"),
        ("Globe Control Valve (Modulating)", "TCV / PCV", "Samson", "Type 3241 + Trovis 3730-1 Positioner", "Flange ASME B16.5 Cl.150/300 RF", "4-20mA Setpoint + Feedback", "LiY-CY TP 2Px1.0 (Dual Shielded Pairs)"),
        ("Pneumatic On-Off Valve (Automated)", "XV / SV", "TeroFox / El-O-Matic", "TF-20DFS Ball Valve + F-Series + ESV Sol.", "Flange Cl.150 RF, SS316", "24VDC DO + 2x DI Limit Sw.", "CVV 3Cx1.5 (Sol) + CVV 4Cx1.0 (Limit)"),
        ("Manual Hand Valve w/ Position Sw.", "HV", "TeroFox / APL-HKC", "TF-20DFS Valve + APL-210N/510N Box", "Flange Cl.150 RF, SS316", "2x DI Dry Contact Feedback", "CVV 4Cx1.0 (Control Cable)"),
        ("Pneumatic Turbine Vibrator", "VIB", "Netter Vibration", "NCT 5 Pneumatic Turbine Vibrator", "G 1/8\" BSP Pneumatic Port", "24VDC Solenoid DO", "CVV 3Cx1.5 (Solenoid Control)")
    ]
    
    for row_data in pkg_rows:
        fill = ZEBRA_ODD if (cur_r % 2) == 1 else ZEBRA_EVEN
        for c_idx, val in enumerate(row_data, 1):
            cell = ws.cell(cur_r, c_idx)
            set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill, border=THIN_BORDER,
                     alignment=ALIGN_CENTER if c_idx in [2, 6] else ALIGN_LEFT)
        cur_r += 1

    cur_r += 2
    ws.cell(cur_r, 1, "2. Standard Instrumentation & Control Cable Selection Guide").font = Font(name="Calibri", size=11, bold=True, color="1B365D")
    cur_r += 1
    
    cable_headers = ["Cable Type / Code", "Conductor Spec", "Screen / Shielding Construction", "Voltage / Temp Rating", "Jacket Color", "Standard Project Application"]
    for c_idx, h in enumerate(cable_headers, 1):
        set_cell(ws.cell(cur_r, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [2, 4, 5] else ALIGN_LEFT)
                 
    cur_r += 1
    cable_data = [
        ("LiY-CY TP 1Px1.0", "1 Twisted Pair x 1.0 mm² Stranded Cu", "Al-mylar tape screen + Tinned copper braid", "300/500 V, -30°C to +80°C", "Gray (RAL 7001) / Black", "Standard 2-wire 4-20mA HART transmitters (AI/AO)"),
        ("LiY-CY(EB) 1Px1.0", "1 Twisted Pair x 1.0 mm² Stranded Cu", "Al-mylar tape screen + Tinned copper braid", "300/500 V, -30°C to +80°C", "Blue (RAL 5015) for Ex i", "Intrinsically Safe Ex i loops connected to IS-JB-603/608/612/618"),
        ("LiY-CY TP 2Px1.0", "2 Twisted Pairs x 1.0 mm² Stranded Cu", "Individual pair screen + Overall copper braid", "300/500 V, -30°C to +80°C", "Gray / Black", "Control valves (AO setpoint + AI feedback), Analyzers (Signal + Alarm)"),
        ("CVV 3Cx1.5", "3 Cores x 1.5 mm² Annealed Solid Cu", "Heavy-duty PVC jacket, unshielded", "600 V, 70°C", "Black UV-Resistant", "24VDC Solenoid valves (DO), Level switches, 24VDC instrument power"),
        ("CVV 4Cx1.0", "4 Cores x 1.0 mm² Annealed Solid Cu", "Heavy-duty PVC jacket, unshielded", "600 V, 70°C", "Black UV-Resistant", "Valve position limit switches (ZSO/ZSC Open & Close dry contact DI)"),
        ("CVV 5Cx1.5", "5 Cores x 1.5 mm² Annealed Solid Cu", "Heavy-duty PVC jacket, unshielded", "600 V, 70°C", "Black UV-Resistant", "Combined on-off valve feed (Solenoid 24VDC + Limit switch common/signals)"),
        ("CAT 6 S/FTP", "4 Pairs 24 AWG Solid Cu", "Braided copper shield + Foil per pair (S/FTP)", "Industrial High-Flex, 80°C", "Teal / Green", "EtherNet/IP digital communications, MCC gateway, remote field interfaces"),
        ("LiY-CY(IS/OS) 16Px0.75", "16 Twisted Pairs x 0.75 mm² Cu", "Individual Foil + Overall Copper Braid Trunk", "300/500 V, 80°C", "Black / Blue Trunk", "Main analog trunk cable from field junction boxes to PLC Cabinet CA1"),
        ("CVV 30Cx1.5", "30 Cores x 1.5 mm² Multi-Core Cu", "Overall PVC heavy-duty armored trunk", "600 V, 70°C", "Black Industrial", "Main discrete digital trunk cable from field junction boxes to CA1")
    ]
    
    for c_row in cable_data:
        fill = ZEBRA_ODD if (cur_r % 2) == 1 else ZEBRA_EVEN
        for c_idx, val in enumerate(c_row, 1):
            cell = ws.cell(cur_r, c_idx)
            set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill, border=THIN_BORDER,
                     alignment=ALIGN_CENTER if c_idx in [2, 4, 5] else ALIGN_LEFT)
        cur_r += 1

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 36
    ws.column_dimensions["E"].width = 32
    ws.column_dimensions["F"].width = 24
    ws.column_dimensions["G"].width = 44

def populate_instrument_master_sheet(ws, df_mapped):
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:AE1")
    title_cell = ws["A1"]
    set_cell(title_cell, "INGREDION (THAILAND) CO., LTD.  |  PROJECT JET COOKER  |  MASTER INSTRUMENT SPECIFICATION & I/O MAP",
             font=Font(name="Calibri", size=13, bold=True, color="FFFFFF"), fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:AE2")
    sub_cell = ws["A2"]
    set_cell(sub_cell, "Loop-by-Loop Instrument Model, Electrical Parameter, PLC I/O Allocation, and Cable Schedule | Rev 3.6 Master",
             font=Font(name="Calibri", size=10, italic=True, color="FFFFFF"), fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Item",
        "Tag No. (Rev 3.6)",
        "Legacy Tag (3.3)",
        "P&ID No.",
        "Description",
        "Instrument Name",
        "Signal Type",
        "Calibrated Range",
        "Protection Class",
        
        "Manufacturer",
        "Model Series",
        "Specific Order Code",
        "Process Connection",
        "Electrical / Power Rating",
        
        "DO", "DI", "AO", "AI", "Bus",
        "PLC Chassis",
        "PLC Slot(s)",
        "Channel Address(es)",
        "PLC Tag Name(s)",
        "Module Model(s)",
        "Panel Enclosure",
        "Destination JB",
        "Status",
        
        "Suggested Field Cable",
        "Conductor Cores & Size",
        "Shielding Construction",
        "Cable Gland Spec",
        "Trunk Cable Route"
    ]
    
    header_row = 4
    ws.row_dimensions[header_row].height = 28
    for c_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(header_row, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [1, 2, 3, 4, 7, 9, 15, 16, 17, 18, 19, 20, 21, 27] else ALIGN_LEFT)

    row_idx = 5
    for _, r in df_mapped.iterrows():
        fill = ZEBRA_ODD if (row_idx % 2) == 1 else ZEBRA_EVEN
        status_val = r["status"]
        
        row_vals = [
            r["item_no"],
            r["active_tag"],
            r["legacy_tag"],
            r["pid"],
            r["desc"],
            r["inst_name"],
            r["sig_type"],
            r["range"],
            r["prot_type"],
            
            r["mfg"],
            r["model"],
            r["order_code"],
            r["conn"],
            r["power"],
            
            r["c_do"], r["c_di"], r["c_ao"], r["c_ai"], r["c_bus"],
            r["chassis"],
            r["slots"],
            r["channel_addrs"],
            r["plc_tags"],
            r["cards"],
            r["panels"],
            r["jbs"],
            status_val,
            
            r["cable_sug"],
            r["cores_size"],
            r["shield"],
            r["gland"],
            r["trunk"]
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row_idx, c_idx)
            c_font = Font(name="Calibri", size=10)
            
            if c_idx == 2:
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=True, color="1B365D"),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)
            elif c_idx in [1, 3, 4, 7, 9, 15, 16, 17, 18, 19, 20, 21]:
                set_cell(cell, val, font=c_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 25:
                set_cell(cell, val, font=c_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT, is_text=True)
            elif c_idx == 27:
                if status_val in STATUS_STYLES:
                    set_cell(cell, val, font=STATUS_STYLES[status_val]["font"], fill=STATUS_STYLES[status_val]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=c_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            else:
                set_cell(cell, val, font=c_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)

        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:{get_column_letter(len(headers))}{row_idx-1}"

    col_widths = {
        "A": 8, "B": 16, "C": 15, "D": 18, "E": 34, "F": 24, "G": 16, "H": 14, "I": 14,
        "J": 22, "K": 32, "L": 22, "M": 28, "N": 26,
        "O": 6, "P": 6, "Q": 6, "R": 6, "S": 6,
        "T": 12, "U": 14, "V": 26, "W": 26, "X": 18, "Y": 18, "Z": 14, "AA": 11,
        "AB": 36, "AC": 24, "AD": 32, "AE": 26, "AF": 32
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

def populate_channel_io_sheet(ws, df_dev, df_mapped):
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:S1")
    title_cell = ws["A1"]
    set_cell(title_cell, "INGREDION (THAILAND) CO., LTD.  |  PROJECT JET COOKER  |  CHANNEL-BY-CHANNEL PLC I/O MAP",
             font=Font(name="Calibri", size=13, bold=True, color="FFFFFF"), fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:S2")
    sub_cell = ws["A2"]
    set_cell(sub_cell, "1,210 I/O Channel Wiring Points Sorted by Chassis, Slot, Point with Instrument Models & Cable Suggestions",
             font=Font(name="Calibri", size=10, italic=True, color="FFFFFF"), fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    inst_lookup = {}
    for _, r in df_mapped.iterrows():
        t_act = str(r["active_tag"]).strip().upper()
        inst_lookup[t_act] = r
        t_mat = str(r["matched_dev_tag"]).strip().upper()
        if t_mat != "-":
            inst_lookup[t_mat] = r

    headers = [
        "Seq No.", "Chassis", "Slot", "Point", "Channel Code", "Terminal Block",
        "PLC Card Model", "I/O Type", "Signal Level", "PLC Tag Name", "Instrument Tag",
        "Instrument Description", "Manufacturer", "Model Series", "Field Cable Suggestion",
        "P&ID No.", "Destination JB", "Panel Enclosure", "Status"
    ]
    
    header_row = 4
    ws.row_dimensions[header_row].height = 28
    for c_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(header_row, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [1, 2, 3, 4, 5, 6, 8, 19] else ALIGN_LEFT)

    row_idx = 5
    for seq, (_, row) in enumerate(df_dev.iterrows(), 1):
        chassis = clean_val(row.get("Chassis"))
        slot = int(row.get("Slot")) if pd.notna(row.get("Slot")) else 0
        point = int(row.get("Point")) if pd.notna(row.get("Point")) else 0
        channel = clean_val(row.get("Column7"))
        terminal = clean_val(row.get("Terminal"))
        card = clean_val(row.get("Card"))
        io_type = clean_val(row.get("I/O Type"))
        signal = clean_val(row.get("Signal Type2"))
        plc_tag = clean_val(row.get("PLC_Tag"))
        inst_tag = clean_val(row.get("Instruement Tag"))
        is_spare = inst_tag.lower() == "spare" or "spare" in inst_tag.lower()
        
        matched_info = inst_lookup.get(inst_tag.upper(), None)
        if matched_info is not None:
            desc = matched_info["desc"]
            mfg = matched_info["mfg"]
            model = matched_info["model"]
            cable = matched_info["cable_sug"]
            pid = matched_info["pid"]
        else:
            desc = clean_val(row.get("Instruement_Description"))
            if desc == "-" or not desc:
                desc = clean_val(row.get("Control Description"))
            if is_spare and (desc == "-" or not desc):
                desc = "Spare I/O Terminal Point"
            mfg = "-" if is_spare else "Rockwell / Automation Field Equipment"
            model = "-" if is_spare else "Standard Loop Device"
            cable = "-" if is_spare else "LiY-CY TP 1Px1.0 / CVV"
            pid = clean_val(row.get("P&ID No. 3.5"))
            if pid == "-":
                pid = clean_val(row.get("P&ID No. 3.3"))

        dest = clean_val(row.get("Destination"))
        panel = clean_val(row.get(" Reference Designation"))
        status = "SPARE" if is_spare else "ACTIVE"

        fill = ZEBRA_ODD if (row_idx % 2) == 1 else ZEBRA_EVEN
        row_vals = [
            seq, chassis, slot, point, channel, terminal, card, io_type, signal,
            plc_tag, inst_tag, desc, mfg, model, cable, pid, dest, panel, status
        ]
        
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws.cell(row_idx, c_idx)
            cell_font = Font(name="Calibri", size=10, color="718096" if is_spare and c_idx not in [8, 19] else "1E293B")
            
            if c_idx in [1, 2, 3, 4, 5, 6]:
                f_bold = (c_idx in [2, 3, 5] and not is_spare)
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=f_bold, color="1B365D" if f_bold else cell_font.color),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 8:
                if io_type in TYPE_STYLES:
                    set_cell(cell, val, font=TYPE_STYLES[io_type]["font"], fill=TYPE_STYLES[io_type]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            elif c_idx == 11:
                set_cell(cell, val, font=Font(name="Calibri", size=10, bold=(not is_spare), color="1B365D" if not is_spare else "718096"),
                         fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)
            elif c_idx == 18:
                set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT, is_text=True)
            elif c_idx == 19:
                if status in STATUS_STYLES:
                    set_cell(cell, val, font=STATUS_STYLES[status]["font"], fill=STATUS_STYLES[status]["fill"],
                             border=THIN_BORDER, alignment=ALIGN_CENTER)
                else:
                    set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_CENTER)
            else:
                set_cell(cell, val, font=cell_font, fill=fill, border=THIN_BORDER, alignment=ALIGN_LEFT)

        ws.row_dimensions[row_idx].height = 20
        row_idx += 1

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:S{row_idx-1}"

    col_widths = {
        "A": 9, "B": 9, "C": 8, "D": 8, "E": 14, "F": 15, "G": 16, "H": 10, "I": 14,
        "J": 22, "K": 18, "L": 34, "M": 20, "N": 30, "O": 28, "P": 18, "Q": 16, "R": 18, "S": 11
    }
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

def populate_catalog_reference(ws):
    ws.views.sheetView[0].showGridLines = True
    
    ws.merge_cells("A1:H1")
    title_cell = ws["A1"]
    set_cell(title_cell, "INGREDION (THAILAND) CO., LTD.  |  PROJECT JET COOKER  |  INSTRUMENT CATALOG REFERENCE",
             font=Font(name="Calibri", size=13, bold=True, color="FFFFFF"), fill=NAVY_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 24

    ws.merge_cells("A2:H2")
    sub_cell = ws["A2"]
    set_cell(sub_cell, "Master Index of Equipment Packages, Manufacturer Manuals, Process Connections & Cable Standards in Folder",
             font=Font(name="Calibri", size=10, italic=True, color="FFFFFF"), fill=SUB_HEADER_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Package Directory", "Equipment Type", "Manufacturer", "Product Series / Model",
        "Key Technical Data Sheet", "Process Connection", "Electrical Standard", "Suggested Field Cable"
    ]
    header_row = 4
    ws.row_dimensions[header_row].height = 28
    for c_idx, h in enumerate(headers, 1):
        set_cell(ws.cell(header_row, c_idx), h, font=HEADER_FONT, fill=NAVY_HEADER_FILL, border=HEADER_BORDER,
                 alignment=ALIGN_CENTER if c_idx in [2, 7] else ALIGN_LEFT)

    cat_rows = [
        ("CS-DS-01 Dust Analyzer", "Particulate / Dust Transmitter", "ENVEA / PCME", "PCME QAL 991", "6626013-33-201-001_Dust Analyzer Data Sheets Rev.0.pdf", "DN100 Flange w/ Purge Air", "4-20mA Output + Alarm Relay / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5 (Pwr)"),
        ("CS-DS-01 LEL gas Analyzer", "Combustible Gas Detector", "Sensidyne / SmartGas", "925FGD Explosionproof", "6626014-33-204-001_General Arrangement Drawing_Rev.0.pdf", "3/4\" NPT Conduit Entry", "4-20mA + Alarm Relays / 24VDC", "LiY-CY TP 2Px1.0 (Ex d)"),
        ("CS-DS-01 pH Analyzer", "pH Analytical Loop", "Mettler Toledo", "M300G2 4-Wire + InPro3250i", "DS_pH_InPro3250i_PA2097EN_RevA.pdf / TD_THO_Meter_M300G2.pdf", "DN25 Weld-in Socket / InTrac 777P", "4-20mA HART + Relays / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5 (Pwr)"),
        ("CS-DS-02 Diff Pressure Gauge", "Diff. Pressure Indicator", "Ashcroft", "Model 1132", "datasheet-differential-1132-pressure-gauge.pdf", "1/4\" NPT Female w/ V03 Manifold", "Local Mechanical Dial", "None (Mechanical)"),
        ("CS-DS-03Diff Pressure Transmitter", "Diff. Pressure Transmitter", "Endress+Hauser", "Deltabar PMD55B", "Spec. Dif. Pressure_R3 19122025.pdf / TI PMD55B.pdf", "1/4\" NPT w/ 5-Valve Manifold DA63M", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)"),
        ("CS-DS-05 Flow meter (Coriolis)", "Coriolis Mass Flowmeter", "Endress+Hauser", "Promass E 300 (8E3B50)", "Spec. Flow_R4 12192025.pdf / TI E300.pdf", "DN50 Cl.150 RF Flange", "4-20mA HART (Mass/Dens) / 24VDC", "LiY-CY TP 3Px1.0 + CVV 3Cx1.5 (Pwr)"),
        ("CS-DS-05 Flow meter (Magmeter)", "Electromagnetic Flowmeter", "Endress+Hauser", "Promag P 300 (5P3B50/80/1H)", "Spec. Flow_R4 12192025.pdf / TI P300.pdf", "DN50/80/100 Cl.150 Flange", "4-20mA HART / 24VDC", "LiY-CY TP 2Px1.0 + CVV 3Cx1.5 (Pwr)"),
        ("CS-DS-05 Flow meter (Vortex)", "Vortex Steam Flowmeter", "Endress+Hauser", "Prowirl F 200 (7F2C80/1H)", "Spec. Flow_R4 12192025.pdf / TI F200.pdf", "DN80/100 Cl.300 RF Flange", "4-20mA HART + Pulse / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)"),
        ("CS-DS-05 Flow meter (Thermal)", "Thermal Mass Gas Flowmeter", "Endress+Hauser", "t-mass F 300 / I 300 (6F3B/6I3B)", "Spec. Flow_R4 12192025.pdf / TI T-Mass F300.pdf", "DN100 Flange / 1\" NPT Insertion", "4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)"),
        ("CS-DS-05 Ultra Sonic Flow meter", "Clamp-on Ultrasonic Flow", "Keyence", "FD-H Series (FD-H20/H32)", "FD-H Instruction manual TH.pdf", "External Clamp-on Sensor", "4-20mA + Pulse / 24VDC", "LiY-CY TP 2Px1.0 (Shielded)"),
        ("CS-DS-07 Hand valve w/ Switch", "Manual Valve Position Box", "TeroFox / APL-HKC", "TF-20DFS + APL-210N", "APL-Series IOM.pdf / General Ball Valves Catalogue.pdf", "Flange Cl.150 RF, SS316", "2x DI Dry Contact Limit Switches", "CVV 4Cx1.0 (Control Cable)"),
        ("CS-DS-09 Level switch (Liquid)", "Vibrating Fork Level Switch", "Endress+Hauser", "Liquiphant FTL51B", "TI FTL51B.pdf", "3/4\" NPT Threaded / Tri-Clamp", "24VDC / Relay Contact Output", "CVV 3Cx1.5 (Control Cable)"),
        ("CS-DS-09 Level switch (Solid)", "Vibrating Fork Powder Switch", "Endress+Hauser", "Soliphant FTM51", "TI FTM51.pdf", "1-1/2\" NPT Threaded / Flange", "24VDC / Relay Contact Output", "CVV 3Cx1.5 (Control Cable)"),
        ("CS-DS-10 Level transmitter", "Radar Level Transmitter", "Endress+Hauser", "Micropilot FMR67B (80GHz)", "TI FMR67B.pdf / TI PMD78B.pdf", "DN80 3\" Cl.150 Flange", "2-wire 4-20mA HART / 24VDC", "LiY-CY(EB) 1Px1.0 (Blue IS Sheath)"),
        ("CS-DS-11 Pressure Gauge", "Pressure Indicator", "Ashcroft", "Model T5500", "DS-T5500-EN-2.pdf / datasheet-1098-1100-siphon.pdf", "1/2\" NPT Male w/ Siphon", "Local Mechanical Dial", "None (Mechanical)"),
        ("CS-DS-12 Pressure Switch", "Electronic Pressure Switch", "Endress+Hauser", "Ceraphant PTC31B", "TI PTC31B.pdf", "1/2\" NPT Male Thread", "24VDC / Dry Contact Relay / PNP", "CVV 3Cx1.5 (Control Cable)"),
        ("CS-DS-13 Pressure Transmitter", "Pressure Transmitter", "Endress+Hauser", "Cerabar PMP51B / PMC51B / PMP43", "TI PMP51B.pdf / TI PMC51B.pdf / TI PMP43.pdf", "1/2\" NPT Male / Tri-Clamp", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)"),
        ("CS-DS-15 Temp gauge", "Bimetal Thermometer", "Ashcroft", "Model FI + Thermowell", "Datasheet_FI_EN_RevB.pdf / datasheet-threaded-thermowell.pdf", "1/2\" NPT Threaded Thermowell", "Local Mechanical Dial", "None (Mechanical)"),
        ("CS-DS-16 Temp transmitter", "RTD Temperature Transmitter", "Endress+Hauser", "iTEMP TMT71 + TM411 (Pt100)", "TI TM411.pdf / TI TMT71.pdf", "1/2\" NPT w/ Threaded Thermowell", "2-wire 4-20mA HART / 24VDC", "LiY-CY TP 1Px1.0 (Shielded)"),
        ("CS-DS-17 Pneumatic on-off valve", "Pneumatic On-Off Valve", "TeroFox / El-O-Matic", "TF-20DFS + F-Series + ESV", "EL-O-MATIC complete data sheet.pdf / 2023ESV.pdf", "Flange Cl.150 RF, SS316", "24VDC DO + 2x DI Limit Sw.", "CVV 3Cx1.5 (Sol) + CVV 4Cx1.0 (Limit)"),
        ("CS-DS-20 Control valve Sanmax", "Globe Control Valve", "Samson", "Type 3241 + Trovis 3730-1", "3241-ANSI.pdf / TI-Trovis 3730-1.pdf", "Flange ASME B16.5 Cl.150/300 RF", "4-20mA Setpoint + Position Feedback", "LiY-CY TP 2Px1.0 (Dual Shielded)"),
        ("Pneumatic Vibrating", "Pneumatic Turbine Vibrator", "Netter Vibration", "NCT 5 Vibrator", "KBA_NCT_NCB_NCR-1882EN.pdf / NCT5 Technical data.jpg", "G 1/8\" BSP Port", "24VDC Solenoid Valve DO", "CVV 3Cx1.5 (Solenoid Control)")
    ]

    cur_r = 5
    for cat in cat_rows:
        fill = ZEBRA_ODD if (cur_r % 2) == 1 else ZEBRA_EVEN
        for c_idx, val in enumerate(cat, 1):
            cell = ws.cell(cur_r, c_idx)
            set_cell(cell, val, font=Font(name="Calibri", size=10), fill=fill, border=THIN_BORDER,
                     alignment=ALIGN_CENTER if c_idx in [2, 7] else ALIGN_LEFT)
        ws.row_dimensions[cur_r].height = 20
        cur_r += 1

    ws.freeze_panes = "A5"
    ws.auto_filter.ref = f"A4:H{cur_r-1}"

    col_widths = {"A": 30, "B": 26, "C": 22, "D": 32, "E": 44, "F": 30, "G": 28, "H": 36}
    for col_let, w in col_widths.items():
        ws.column_dimensions[col_let].width = w

def main():
    df_mapped, df_dev = load_and_merge_data()

    print(f"Creating comprehensive workbook: {OUTPUT_FILE}")
    wb = openpyxl.Workbook()

    # Sheet 1: Executive Dashboard
    ws_dash = wb.active
    ws_dash.title = "Executive_Summary"
    populate_dashboard(ws_dash, df_mapped, df_dev)

    # Sheet 2: Master Instrument Specification & IO Map
    ws_master = wb.create_sheet(title="Instrument_Master_IO_Map")
    populate_instrument_master_sheet(ws_master, df_mapped)

    # Sheet 3: Channel-by-Channel IO Map (1,210 points)
    ws_channel = wb.create_sheet(title="Channel_Point_IO_Map")
    populate_channel_io_sheet(ws_channel, df_dev, df_mapped)

    # Sheet 4: Instrument Catalog Reference
    ws_cat = wb.create_sheet(title="Instrument_Catalog_Reference")
    populate_catalog_reference(ws_cat)

    wb.save(OUTPUT_FILE)
    print(f"Successfully generated: {OUTPUT_FILE} ({os.path.getsize(OUTPUT_FILE):,} bytes)")

if __name__ == "__main__":
    main()
