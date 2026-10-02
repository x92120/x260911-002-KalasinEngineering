#!/usr/bin/env python3
"""
========================================================================================
Project: KALASIN STARCH PLANT - JET COOKER AUTOMATION SYSTEM (xCIP-1545)
Client:  INGREDION (THAILAND) CO., LTD.
Script:  generate_pf_is_barrier_dxf.py
Purpose: Generate detailed publication-grade AutoCAD DXF (1:1 Metric mm) electrical 
         schematic, terminal pinout, and loop wiring drawings for Pepperl+Fuchs 
         Intrinsically Safe (IS) equipment:
         1. 216711_eng.pdf -> HiC2821 Switch Amplifier (DI Barrier, SIL 2)
         2. 233883_eng.pdf -> HiC2871 Solenoid Driver (DO Barrier, SIL 3)
         3. 321423_eng.pdf -> HiC2025 SMART Transmitter Power Supply (AI Barrier, SIL 2)
         4. 260436_eng.pdf -> HiCTB16-SCT-44C-SC-RA 16-Slot Universal Termination Board
         5. Master Loop Hookup -> Panel CA-IS & ControlLogix 1756 Interface Schematic

Outputs:
  - 02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/IS_Barriers_CAD/*.dxf
  - 02_Electrical_and_eDrawing/05_Instrument_Manuals/Pepperl+Fuchs IS Barriers/*.dxf
========================================================================================
"""

import os
import shutil
import ezdxf
from ezdxf import colors
from ezdxf.enums import TextEntityAlignment

def setup_dxf_document():
    doc = ezdxf.new("R2010", setup=True)
    doc.header["$INSUNITS"] = 4  # Millimeters
    msp = doc.modelspace()
    
    # Standard IEC / ISA CAD Layers
    layers = [
        ("0_BORDER", 7, 70),             # White, 0.70mm
        ("0_TITLE_BLOCK", 4, 35),        # Cyan, 0.35mm
        ("IE-EQUIP", 2, 35),             # Yellow, 0.35mm
        ("IE-WIRE", 1, 25),              # Red, 0.25mm
        ("IE-WIRE-BLUE", 5, 25),         # Blue, 0.25mm (IS Hazardous Wire)
        ("IE-WIRE-BLACK", 7, 25),        # White/Black, 0.25mm (Safe Area Wire)
        ("IE-CABLE", 6, 25),             # Magenta, 0.25mm
        ("TER-BLUE", 5, 35),             # Blue (Ex Hazardous Terminals)
        ("TER-BLACK", 7, 35),            # White/Black (Safe Control Terminals)
        ("IS_BARRIERS", 140, 35),        # Sky Blue, 0.35mm
        ("RELAYS", 30, 35),              # Orange, 0.35mm
        ("NOTATIONS", 8, 18),            # Gray, 0.18mm
        ("TEXTS", 7, 25),                # White, 0.25mm
        ("TEXTS-CYAN", 4, 25),           # Cyan, 0.25mm
        ("TEXTS-YELLOW", 2, 25),         # Yellow, 0.25mm
        ("TEXTS-GREEN", 3, 25),          # Green, 0.25mm
        ("DIMENSIONS", 1, 18),           # Red, 0.18mm
        ("HAZARDOUS_ZONE", 1, 18),       # Red dashed zone
    ]
    for name, col, lw in layers:
        if name not in doc.layers:
            doc.layers.add(name, color=col, lineweight=lw)
            
    return doc, msp

def add_rect(msp, x, y, w, h, layer="0_BORDER"):
    pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
    return msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})

def add_text(msp, text, x, y, height=2.5, layer="TEXTS", align=TextEntityAlignment.LEFT):
    txt = msp.add_text(text, dxfattribs={"height": height, "layer": layer})
    txt.set_placement((x, y), align=align)
    return txt

def draw_title_block(msp, title, dwg_no, sheet_no=1, total_sheets=5, part_no=""):
    ox, oy = 0.0, 0.0
    w_sheet, h_sheet = 420.0, 297.0  # ISO A3 Landscape
    
    # Outer Border & Margin
    add_rect(msp, ox + 5.0, oy + 5.0, w_sheet - 10.0, h_sheet - 10.0, layer="0_BORDER")
    add_rect(msp, ox + 7.0, oy + 7.0, w_sheet - 14.0, h_sheet - 14.0, layer="NOTATIONS")
    
    # Title Block (Bottom Right)
    tb_w = 185.0
    tb_h = 40.0
    tx = ox + w_sheet - 7.0 - tb_w
    ty = oy + 7.0
    
    add_rect(msp, tx, ty, tb_w, tb_h, layer="0_BORDER")
    msp.add_line((tx, ty + 12.0), (tx + tb_w, ty + 12.0), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tx, ty + 24.0), (tx + tb_w, ty + 24.0), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tx + 115.0, ty), (tx + 115.0, ty + 24.0), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tx + 60.0, ty + 24.0), (tx + 60.0, ty + tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tx + 120.0, ty + 24.0), (tx + 120.0, ty + tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tx + 155.0, ty + 24.0), (tx + 155.0, ty + tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})
    
    # Logos and Text
    add_text(msp, "INGREDION (THAILAND) CO., LTD.", tx + 3.0, ty + 34.0, height=3.2, layer="0_TITLE_BLOCK")
    add_text(msp, "KALASIN STARCH PLANT - SPRINT 18K JET COOKER", tx + 3.0, ty + 28.5, height=2.2, layer="TEXTS")
    add_text(msp, "AEC INDUSTRIAL ENGINEERING", tx + 63.0, ty + 34.0, height=2.8, layer="TEXTS-CYAN")
    add_text(msp, "PROCESS CONTROL & SAFETY SYSTEMS", tx + 63.0, ty + 28.5, height=2.0, layer="NOTATIONS")
    
    add_text(msp, f"DOC NO: {dwg_no}", tx + 3.0, ty + 17.5, height=2.8, layer="TEXTS-YELLOW")
    add_text(msp, f"DWG TITLE: {title}", tx + 3.0, ty + 14.0, height=2.2, layer="TEXTS")
    add_text(msp, f"EQUIPMENT REF: PEPPERL+FUCHS {part_no}", tx + 3.0, ty + 7.0, height=2.0, layer="NOTATIONS")
    add_text(msp, "PANEL: CA-IS (MARSHALLING / EX BOUNDARY)", tx + 3.0, ty + 3.0, height=2.0, layer="TEXTS-GREEN")
    
    add_text(msp, f"SHEET: {sheet_no:02d} OF {total_sheets:02d}", tx + 120.0, ty + 17.5, height=2.5, layer="TEXTS")
    add_text(msp, "REV: 03 (APPROVED)", tx + 120.0, ty + 13.5, height=2.2, layer="TEXTS-GREEN")
    add_text(msp, "SCALE: 1:1 METRIC mm", tx + 120.0, ty + 7.0, height=2.0, layer="NOTATIONS")
    add_text(msp, "DATE: OCT 2026", tx + 120.0, ty + 3.0, height=2.0, layer="NOTATIONS")

# --------------------------------------------------------------------------------------
# 1. DRAWING 1: HiC2821 Switch Amplifier (DI Isolated Barrier)
# --------------------------------------------------------------------------------------
def draw_hic2821(out_path):
    doc, msp = setup_dxf_document()
    draw_title_block(msp, "HiC2821 SWITCH AMPLIFIER (DI ISOLATED BARRIER)", "KAL-JC-IS-001", sheet_no=1, total_sheets=5, part_no="HiC2821 / 216711")
    
    # Main Drawing Frame Areas
    # Hazardous Area (Left), Safe Area (Right)
    # Boundary Line at X = 200.0
    msp.add_line((200.0, 50.0), (200.0, 275.0), dxfattribs={"layer": "HAZARDOUS_ZONE", "color": 1})
    add_text(msp, "══════════ HAZARDOUS AREA (EX ZONE 0, 1, 20, 21) ══════════", 25.0, 280.0, height=3.5, layer="TEXTS-YELLOW")
    add_text(msp, "══════════ SAFE AREA (PANEL CA-IS MARSHALLING) ══════════", 210.0, 280.0, height=3.5, layer="TEXTS-CYAN")
    
    # 1. HiC2821 Module Enclosure (Center)
    mx, my, mw, mh = 170.0, 90.0, 60.0, 170.0
    add_rect(msp, mx, my, mw, mh, layer="IS_BARRIERS")
    add_text(msp, "PEPPERL+FUCHS", mx + 12.0, my + mh - 8.0, height=3.0, layer="TEXTS-YELLOW")
    add_text(msp, "HiC2821", mx + 18.0, my + mh - 16.0, height=4.5, layer="TEXTS-CYAN")
    add_text(msp, "1-CH SWITCH AMPLIFIER", mx + 11.0, my + mh - 22.0, height=2.2, layer="TEXTS")
    add_text(msp, "SIL 2 / SC 3 (IEC 61508)", mx + 13.0, my + mh - 27.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "ATEX: II (1)G [Ex ia Ga] IIC", mx + 10.0, my + mh - 32.0, height=1.8, layer="NOTATIONS")
    
    # Optical Isolation Barrier inside module
    msp.add_line((mx + 28.0, my + 15.0), (mx + 28.0, my + 130.0), dxfattribs={"layer": "NOTATIONS"})
    add_text(msp, "GALVANIC ISOLATION", mx + 20.0, my + 75.0, height=1.8, layer="NOTATIONS")
    
    # Status LEDs
    add_rect(msp, mx + 5.0, my + mh - 45.0, 8.0, 6.0, layer="0_BORDER")
    add_text(msp, "PWR (Green)", mx + 15.0, my + mh - 43.0, height=2.0, layer="TEXTS-GREEN")
    add_rect(msp, mx + 5.0, my + mh - 55.0, 8.0, 6.0, layer="0_BORDER")
    add_text(msp, "OUT1 (Yellow)", mx + 15.0, my + mh - 53.0, height=2.0, layer="TEXTS-YELLOW")
    add_rect(msp, mx + 5.0, my + mh - 65.0, 8.0, 6.0, layer="0_BORDER")
    add_text(msp, "FAULT (Red/LFD)", mx + 15.0, my + mh - 63.0, height=2.0, layer="IE-WIRE")
    
    # Terminals on HiC2821:
    # Hazardous Field Side: SL2 5a (+), 5b (-)
    # Terminals on left edge
    t_y1 = my + 120.0
    t_y2 = my + 100.0
    add_rect(msp, mx - 6.0, t_y1 - 4.0, 6.0, 8.0, layer="TER-BLUE")
    add_text(msp, "5a (+)", mx - 18.0, t_y1 - 1.5, height=2.5, layer="TEXTS-YELLOW")
    add_rect(msp, mx - 6.0, t_y2 - 4.0, 6.0, 8.0, layer="TER-BLUE")
    add_text(msp, "5b (-)", mx - 18.0, t_y2 - 1.5, height=2.5, layer="TEXTS-YELLOW")
    
    # Safe Area Side: SL1 8a, 7a (Relay Out 1), 10a, 9a (Relay Out 2 / Fault), 1a/1b, 2a/2b (Power)
    s_y1 = my + 140.0
    s_y2 = my + 125.0
    s_y3 = my + 110.0
    s_y4 = my + 95.0
    s_y5 = my + 60.0
    s_y6 = my + 45.0
    s_y7 = my + 25.0
    
    # Output Relay 1 (8a, 7a)
    add_rect(msp, mx + mw, s_y1 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "8a (NO)", mx + mw + 8.0, s_y1 - 1.5, height=2.2, layer="TEXTS")
    add_rect(msp, mx + mw, s_y2 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "7a (COM)", mx + mw + 8.0, s_y2 - 1.5, height=2.2, layer="TEXTS")
    
    # Output Relay 2 (10a, 9a)
    add_rect(msp, mx + mw, s_y3 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "10a (NO)", mx + mw + 8.0, s_y3 - 1.5, height=2.2, layer="TEXTS")
    add_rect(msp, mx + mw, s_y4 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "9a (COM)", mx + mw + 8.0, s_y4 - 1.5, height=2.2, layer="TEXTS")
    
    # Fault Bus (6b)
    add_rect(msp, mx + mw, s_y5 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "6b (FAULT BUS)", mx + mw + 8.0, s_y5 - 1.5, height=2.2, layer="IE-WIRE")
    
    # Power Supply (2a/2b +24V, 1a/1b 0V)
    add_rect(msp, mx + mw, s_y6 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "2a, 2b (+24VDC Bus)", mx + mw + 8.0, s_y6 - 1.5, height=2.2, layer="IE-WIRE")
    add_rect(msp, mx + mw, s_y7 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "1a, 1b (0VDC Bus)", mx + mw + 8.0, s_y7 - 1.5, height=2.2, layer="NOTATIONS")
    
    # Field Wiring (Hazardous Side - Left)
    # Field Device: Ex NAMUR Proximity Sensor / Dry Contact Limit Switch
    fx, fy = 40.0, 100.0
    add_rect(msp, fx, fy, 45.0, 45.0, layer="IE-EQUIP")
    add_text(msp, "EX FIELD DEVICE", fx + 5.0, fy + 38.0, height=2.5, layer="TEXTS-YELLOW")
    add_text(msp, "NAMUR Proximity Switch", fx + 3.0, fy + 30.0, height=2.0, layer="TEXTS")
    add_text(msp, "or Dry Contact Switch", fx + 4.0, fy + 24.0, height=2.0, layer="TEXTS")
    add_text(msp, "(e.g. ZS-60201 Limit)", fx + 4.0, fy + 17.0, height=1.8, layer="TEXTS-CYAN")
    add_text(msp, "Zone 0 / 1 / 20 / 21", fx + 6.0, fy + 8.0, height=2.0, layer="NOTATIONS")
    
    # Blue Field Cables (SL2 5a, 5b)
    msp.add_line((fx + 45.0, fy + 30.0), (mx - 6.0, t_y1), dxfattribs={"layer": "IE-WIRE-BLUE"})
    add_text(msp, "Cable: 1Px1.5mm² IS Blue (Belden 8760)", fx + 50.0, t_y1 + 4.0, height=1.8, layer="TEXTS-CYAN")
    msp.add_line((fx + 45.0, fy + 15.0), (mx - 6.0, t_y2), dxfattribs={"layer": "IE-WIRE-BLUE"})
    
    # Safe Area Control Wiring (Right)
    # PLC 1756-IB32 Digital Input Card
    px, py = 320.0, 90.0
    add_rect(msp, px, py, 60.0, 80.0, layer="IE-EQUIP")
    add_text(msp, "ALLEN-BRADLEY", px + 12.0, py + 72.0, height=2.8, layer="TEXTS-YELLOW")
    add_text(msp, "1756-IB32", px + 18.0, py + 64.0, height=4.0, layer="TEXTS-CYAN")
    add_text(msp, "32-PT 24VDC DI MODULE", px + 5.0, py + 56.0, height=2.0, layer="TEXTS")
    add_text(msp, "Panel CA1 / CA-RIO Drop", px + 8.0, py + 48.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "Pin: IN-x (Channel Input)", px + 5.0, py + 30.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "Terminal: P4-TBDIx-m", px + 5.0, py + 20.0, height=2.0, layer="TEXTS-CYAN")
    
    # Connection line to PLC
    msp.add_line((mx + mw + 6.0, s_y1), (px, s_y1), dxfattribs={"layer": "IE-WIRE"})
    add_text(msp, "Signal to PLC DI Channel", mx + mw + 25.0, s_y1 + 3.0, height=2.0, layer="TEXTS")
    
    # 2. DIP Switch Configuration Table (Lower Left)
    bx, by = 20.0, 15.0
    add_rect(msp, bx, by, 135.0, 32.0, layer="0_BORDER")
    add_text(msp, "DIP SWITCH CONFIGURATION (HiC2821)", bx + 3.0, by + 26.0, height=2.5, layer="TEXTS-YELLOW")
    add_text(msp, "S1: Mode of Operation (I: Normal / II: Inverted)  -> Factory: I (Normal)", bx + 3.0, by + 20.0, height=2.0, layer="TEXTS")
    add_text(msp, "S2: Output II Function (I: Signal / II: Fault Alarm) -> Factory: II (Fault)", bx + 3.0, by + 15.0, height=2.0, layer="TEXTS")
    add_text(msp, "S3: Line Fault Detection (LFD) (I: ON / II: OFF) -> Factory: I (ON)", bx + 3.0, by + 10.0, height=2.0, layer="TEXTS")
    add_text(msp, "S4: No function / Spare", bx + 3.0, by + 5.0, height=2.0, layer="NOTATIONS")
    
    doc.saveas(out_path)

# --------------------------------------------------------------------------------------
# 2. DRAWING 2: HiC2871 Solenoid Driver (DO Isolated Barrier)
# --------------------------------------------------------------------------------------
def draw_hic2871(out_path):
    doc, msp = setup_dxf_document()
    draw_title_block(msp, "HiC2871 SOLENOID DRIVER (DO ISOLATED BARRIER)", "KAL-JC-IS-002", sheet_no=2, total_sheets=5, part_no="HiC2871 / 233883")
    
    # Zones
    msp.add_line((200.0, 50.0), (200.0, 275.0), dxfattribs={"layer": "HAZARDOUS_ZONE", "color": 1})
    add_text(msp, "══════════ HAZARDOUS AREA (EX ZONE 0, 1, 20, 21) ══════════", 25.0, 280.0, height=3.5, layer="TEXTS-YELLOW")
    add_text(msp, "══════════ SAFE AREA (PANEL CA-IS MARSHALLING) ══════════", 210.0, 280.0, height=3.5, layer="TEXTS-CYAN")
    
    # HiC2871 Module Enclosure
    mx, my, mw, mh = 170.0, 90.0, 60.0, 170.0
    add_rect(msp, mx, my, mw, mh, layer="IS_BARRIERS")
    add_text(msp, "PEPPERL+FUCHS", mx + 12.0, my + mh - 8.0, height=3.0, layer="TEXTS-YELLOW")
    add_text(msp, "HiC2871", mx + 18.0, my + mh - 16.0, height=4.5, layer="TEXTS-CYAN")
    add_text(msp, "1-CH SOLENOID DRIVER", mx + 10.0, my + mh - 22.0, height=2.2, layer="TEXTS")
    add_text(msp, "SIL 3 (IEC 61508) LOOP POWERED", mx + 5.0, my + mh - 27.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "OUTPUT: 45 mA @ 12V DC", mx + 12.0, my + mh - 32.0, height=2.0, layer="TEXTS-YELLOW")
    add_text(msp, "ATEX: II (1)G [Ex ia Ga] IIC", mx + 10.0, my + mh - 37.0, height=1.8, layer="NOTATIONS")
    
    # Internal Zener Barrier & Current Limiter Block
    add_rect(msp, mx + 15.0, my + 50.0, 30.0, 60.0, layer="0_BORDER")
    add_text(msp, "SAFETY BARRIER", mx + 16.0, my + 95.0, height=2.0, layer="TEXTS-YELLOW")
    add_text(msp, "& CURRENT LIMITER", mx + 16.0, my + 88.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "Ri <= 238 Ohm", mx + 20.0, my + 78.0, height=2.0, layer="TEXTS-CYAN")
    add_text(msp, "Uo = 25.2V / Io = 110mA", mx + 16.0, my + 68.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "Po = 570 mW", mx + 22.0, my + 58.0, height=1.8, layer="NOTATIONS")
    
    # Yellow Status LED
    add_rect(msp, mx + 5.0, my + mh - 50.0, 8.0, 6.0, layer="0_BORDER")
    add_text(msp, "STATUS (Yellow: Output ON)", mx + 15.0, my + mh - 48.0, height=2.0, layer="TEXTS-YELLOW")
    
    # Terminals:
    # Hazardous Field Side: SL2 5a (+), 5b (-)
    t_y1 = my + 100.0
    t_y2 = my + 70.0
    add_rect(msp, mx - 6.0, t_y1 - 4.0, 6.0, 8.0, layer="TER-BLUE")
    add_text(msp, "5a (+) [Ex Output]", mx - 28.0, t_y1 - 1.5, height=2.2, layer="TEXTS-YELLOW")
    add_rect(msp, mx - 6.0, t_y2 - 4.0, 6.0, 8.0, layer="TER-BLUE")
    add_text(msp, "5b (-) [Ex Output]", mx - 28.0, t_y2 - 1.5, height=2.2, layer="TEXTS-YELLOW")
    
    # Safe Control Side: SL1 8a (+), 7a (-) (Loop Powered Control Input)
    s_y1 = my + 100.0
    s_y2 = my + 70.0
    add_rect(msp, mx + mw, s_y1 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "8a (+) [Control In 19-30V]", mx + mw + 8.0, s_y1 - 1.5, height=2.2, layer="TEXTS")
    add_rect(msp, mx + mw, s_y2 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "7a (-) [Control In 0V]", mx + mw + 8.0, s_y2 - 1.5, height=2.2, layer="TEXTS")
    
    # Field Ex Solenoid Valve
    fx, fy = 40.0, 75.0
    add_rect(msp, fx, fy, 45.0, 45.0, layer="IE-EQUIP")
    add_text(msp, "EX SOLENOID VALVE", fx + 3.0, fy + 38.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "Pneumatic On-Off Valve", fx + 3.0, fy + 30.0, height=2.0, layer="TEXTS")
    add_text(msp, "Coil: 12VDC / 45mA IS", fx + 4.0, fy + 24.0, height=2.0, layer="TEXTS-CYAN")
    add_text(msp, "(e.g. XV-61301 / ISV)", fx + 5.0, fy + 17.0, height=1.8, layer="TEXTS-GREEN")
    add_text(msp, "Zone 0 / 1 / 20 / 21", fx + 7.0, fy + 8.0, height=2.0, layer="NOTATIONS")
    
    # Blue Field Wires
    msp.add_line((fx + 45.0, fy + 32.0), (mx - 6.0, t_y1), dxfattribs={"layer": "IE-WIRE-BLUE"})
    add_text(msp, "Blue IS Cable: 2Cx1.5mm² (Belden 8471)", fx + 50.0, t_y1 + 4.0, height=1.8, layer="TEXTS-CYAN")
    msp.add_line((fx + 45.0, fy + 12.0), (mx - 6.0, t_y2), dxfattribs={"layer": "IE-WIRE-BLUE"})
    
    # Safe Control Side Driving Circuit:
    # PLC 1756-OB32 -> Phoenix Contact PLC-RSC-24DC/21 Relay -> HiC2871 Input
    rx, ry = 270.0, 65.0
    add_rect(msp, rx, ry, 40.0, 60.0, layer="RELAYS")
    add_text(msp, "PHOENIX CONTACT", rx + 4.0, ry + 52.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "PLC-RSC-24DC/21", rx + 4.0, ry + 44.0, height=2.2, layer="TEXTS-CYAN")
    add_text(msp, "SPDT Relay Contact", rx + 5.0, ry + 36.0, height=2.0, layer="TEXTS")
    add_text(msp, "Driven by 1756-OB32", rx + 4.0, ry + 28.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "24VDC Panel Supply", rx + 4.0, ry + 12.0, height=1.8, layer="IE-WIRE")
    
    msp.add_line((rx + 40.0, ry + 35.0), (mx + mw, s_y1), dxfattribs={"layer": "IE-WIRE"})
    msp.add_line((rx + 40.0, ry + 20.0), (mx + mw, s_y2), dxfattribs={"layer": "NOTATIONS"})
    
    # Notes Block
    bx, by = 20.0, 15.0
    add_rect(msp, bx, by, 135.0, 32.0, layer="0_BORDER")
    add_text(msp, "ENGINEERING APPLICATION NOTES (HiC2871)", bx + 3.0, by + 26.0, height=2.5, layer="TEXTS-YELLOW")
    add_text(msp, "1. Loop-powered design: No auxiliary 24V bus power required for module operation.", bx + 3.0, by + 20.0, height=2.0, layer="TEXTS")
    add_text(msp, "2. Compatible with standard low-power intrinsically safe solenoid coils (12V / 45mA).", bx + 3.0, by + 15.0, height=2.0, layer="TEXTS")
    add_text(msp, "3. Highest functional safety rating: Certified for SIL 3 safety instrumented functions.", bx + 3.0, by + 10.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "4. Pin 1 and Pin 4 are trimmed for keying / polarizing safety on the termination board.", bx + 3.0, by + 5.0, height=2.0, layer="NOTATIONS")
    
    doc.saveas(out_path)

# --------------------------------------------------------------------------------------
# 3. DRAWING 3: HiC2025 SMART Transmitter Power Supply (AI Isolated Barrier)
# --------------------------------------------------------------------------------------
def draw_hic2025(out_path):
    doc, msp = setup_dxf_document()
    draw_title_block(msp, "HiC2025 SMART TRANSMITTER POWER SUPPLY (AI BARRIER)", "KAL-JC-IS-003", sheet_no=3, total_sheets=5, part_no="HiC2025 / 321423")
    
    # Zones
    msp.add_line((200.0, 50.0), (200.0, 275.0), dxfattribs={"layer": "HAZARDOUS_ZONE", "color": 1})
    add_text(msp, "══════════ HAZARDOUS AREA (EX ZONE 0, 1, 20, 21) ══════════", 25.0, 280.0, height=3.5, layer="TEXTS-YELLOW")
    add_text(msp, "══════════ SAFE AREA (PANEL CA-IS MARSHALLING) ══════════", 210.0, 280.0, height=3.5, layer="TEXTS-CYAN")
    
    # HiC2025 Module Enclosure
    mx, my, mw, mh = 170.0, 90.0, 60.0, 170.0
    add_rect(msp, mx, my, mw, mh, layer="IS_BARRIERS")
    add_text(msp, "PEPPERL+FUCHS", mx + 12.0, my + mh - 8.0, height=3.0, layer="TEXTS-YELLOW")
    add_text(msp, "HiC2025", mx + 18.0, my + mh - 16.0, height=4.5, layer="TEXTS-CYAN")
    add_text(msp, "SMART TRANSMITTER REPEATER", mx + 4.0, my + mh - 22.0, height=2.0, layer="TEXTS")
    add_text(msp, "SIL 2 / SC 3 (IEC 61508)", mx + 13.0, my + mh - 27.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "HART PROTOCOL TRANSPARENT", mx + 6.0, my + mh - 32.0, height=2.0, layer="TEXTS-YELLOW")
    add_text(msp, "ATEX: II (1)G [Ex ia Ga] IIC", mx + 10.0, my + mh - 37.0, height=1.8, layer="NOTATIONS")
    
    # Internal Circuit: DC/DC Converter + HART Modulator
    add_rect(msp, mx + 12.0, my + 50.0, 36.0, 65.0, layer="0_BORDER")
    add_text(msp, "GALVANIC ISOLATION", mx + 14.0, my + 105.0, height=2.0, layer="NOTATIONS")
    add_text(msp, "DC/DC CONVERTER", mx + 15.0, my + 95.0, height=2.0, layer="TEXTS-CYAN")
    add_text(msp, "HART FSK COUPLER", mx + 15.0, my + 85.0, height=2.0, layer="TEXTS-YELLOW")
    add_text(msp, "CURRENT REPEATER", mx + 15.0, my + 75.0, height=2.0, layer="TEXTS")
    add_text(msp, "Accuracy < 0.1%", mx + 18.0, my + 65.0, height=1.8, layer="TEXTS-GREEN")
    add_text(msp, "Power: <= 800 mW", mx + 16.0, my + 55.0, height=1.8, layer="NOTATIONS")
    
    # Green PWR LED
    add_rect(msp, mx + 5.0, my + mh - 50.0, 8.0, 6.0, layer="0_BORDER")
    add_text(msp, "PWR (Green: 24VDC Normal)", mx + 15.0, my + mh - 48.0, height=2.0, layer="TEXTS-GREEN")
    
    # Terminals:
    # Hazardous Field Side: SL2 5a (+), 5b (-) (2-Wire SMART Transmitter)
    t_y1 = my + 115.0
    t_y2 = my + 85.0
    add_rect(msp, mx - 6.0, t_y1 - 4.0, 6.0, 8.0, layer="TER-BLUE")
    add_text(msp, "5a (+) [24V Supply/Signal]", mx - 38.0, t_y1 - 1.5, height=2.2, layer="TEXTS-YELLOW")
    add_rect(msp, mx - 6.0, t_y2 - 4.0, 6.0, 8.0, layer="TER-BLUE")
    add_text(msp, "5b (-) [Current Return]", mx - 38.0, t_y2 - 1.5, height=2.2, layer="TEXTS-YELLOW")
    
    # Safe Control Side: SL1 8a (+), 7a (-) (4-20mA Output to 1756-IF16), 1a/1b, 2a/2b (Power)
    s_y1 = my + 115.0
    s_y2 = my + 85.0
    s_y3 = my + 45.0
    s_y4 = my + 25.0
    
    add_rect(msp, mx + mw, s_y1 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "8a (+) [4-20mA Out]", mx + mw + 8.0, s_y1 - 1.5, height=2.2, layer="TEXTS")
    add_rect(msp, mx + mw, s_y2 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "7a (-) [4-20mA Return]", mx + mw + 8.0, s_y2 - 1.5, height=2.2, layer="TEXTS")
    add_rect(msp, mx + mw, s_y3 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "2a, 2b (+24VDC Bus)", mx + mw + 8.0, s_y3 - 1.5, height=2.2, layer="IE-WIRE")
    add_rect(msp, mx + mw, s_y4 - 4.0, 6.0, 8.0, layer="TER-BLACK")
    add_text(msp, "1a, 1b (0VDC Bus)", mx + mw + 8.0, s_y4 - 1.5, height=2.2, layer="NOTATIONS")
    
    # Field 2-Wire SMART Transmitter (e.g. Endress+Hauser / Rosemount Pressure/Temp)
    fx, fy = 35.0, 80.0
    add_rect(msp, fx, fy, 48.0, 50.0, layer="IE-EQUIP")
    add_text(msp, "2-WIRE TRANSMITTER", fx + 3.0, fy + 42.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "HART 4-20mA Ex ia", fx + 5.0, fy + 34.0, height=2.0, layer="TEXTS")
    add_text(msp, "Temp: TT-60201 (RTD)", fx + 4.0, fy + 26.0, height=1.8, layer="TEXTS-CYAN")
    add_text(msp, "Press: PT-60201 (HART)", fx + 4.0, fy + 18.0, height=1.8, layer="TEXTS-CYAN")
    add_text(msp, "Zone 0 / 1 / 20 / 21", fx + 7.0, fy + 8.0, height=2.0, layer="NOTATIONS")
    
    # Blue Field Cables
    msp.add_line((fx + 48.0, fy + 35.0), (mx - 6.0, t_y1), dxfattribs={"layer": "IE-WIRE-BLUE"})
    add_text(msp, "Shielded Twisted Pair: 1Px1.5mm² IS Blue (Belden 8777)", fx + 50.0, t_y1 + 4.0, height=1.8, layer="TEXTS-CYAN")
    msp.add_line((fx + 48.0, fy + 15.0), (mx - 6.0, t_y2), dxfattribs={"layer": "IE-WIRE-BLUE"})
    
    # Control Side: ControlLogix 1756-IF16 Analog Input Module
    px, py = 320.0, 75.0
    add_rect(msp, px, py, 60.0, 75.0, layer="IE-EQUIP")
    add_text(msp, "ALLEN-BRADLEY", px + 12.0, py + 67.0, height=2.8, layer="TEXTS-YELLOW")
    add_text(msp, "1756-IF16", px + 18.0, py + 59.0, height=4.0, layer="TEXTS-CYAN")
    add_text(msp, "16-CH ANALOG INPUT", px + 8.0, py + 51.0, height=2.0, layer="TEXTS")
    add_text(msp, "Panel CA1 / CA-RIO", px + 10.0, py + 43.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "Pin: IN-x (+ Current Input)", px + 4.0, py + 26.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "Pin: iRTN-x (- Return)", px + 4.0, py + 16.0, height=2.0, layer="TEXTS-CYAN")
    
    msp.add_line((mx + mw + 6.0, s_y1), (px, s_y1), dxfattribs={"layer": "IE-WIRE"})
    msp.add_line((mx + mw + 6.0, s_y2), (px, s_y2), dxfattribs={"layer": "IE-WIRE"})
    add_text(msp, "Isolated 4-20mA Analog Signal", mx + mw + 20.0, s_y1 + 3.0, height=2.0, layer="TEXTS")
    
    # DIP Switch Table
    bx, by = 20.0, 15.0
    add_rect(msp, bx, by, 135.0, 32.0, layer="0_BORDER")
    add_text(msp, "OUTPUT MODE CONFIGURATION (DIP SWITCH S1-S4)", bx + 3.0, by + 26.0, height=2.5, layer="TEXTS-YELLOW")
    add_text(msp, "Current Source 4 ... 20 mA: S1=OFF, S2=OFF, S3=ON, S4=OFF (Factory Setting -> To 1756-IF16)", bx + 3.0, by + 20.0, height=1.9, layer="TEXTS-GREEN")
    add_text(msp, "Voltage Source 1 ... 5 V:   S1=OFF, S2=OFF, S3=ON, S4=ON  (250 Ohm internal shunt resistor)", bx + 3.0, by + 15.0, height=1.9, layer="TEXTS")
    add_text(msp, "Current Sink 4 ... 20 mA:   S1=OFF, S2=ON,  S3=OFF, S4=OFF (External powered loop)", bx + 3.0, by + 10.0, height=1.9, layer="TEXTS")
    add_text(msp, "HART Handheld Communicator connects across terminals 8a/7a or field terminals 5a/5b.", bx + 3.0, by + 5.0, height=1.8, layer="TEXTS-CYAN")
    
    doc.saveas(out_path)

# --------------------------------------------------------------------------------------
# 4. DRAWING 4: HiCTB16-SCT-44C-SC-RA 16-Slot Termination Board Layout
# --------------------------------------------------------------------------------------
def draw_hictb16(out_path):
    doc, msp = setup_dxf_document()
    draw_title_block(msp, "HiCTB16-SCT-44C-SC-RA 16-SLOT TERMINATION BOARD", "KAL-JC-IS-004", sheet_no=4, total_sheets=5, part_no="HiCTB16 / 260436")
    
    # Outer Baseplate Outline (1:1 mm scale: 266 mm wide x 170 mm high)
    # Centered on drawing:
    bx, by = 50.0, 75.0
    bw, bh = 290.0, 160.0
    
    add_rect(msp, bx, by, bw, bh, layer="0_BORDER")
    add_text(msp, "PEPPERL+FUCHS  HiCTB16-SCT-44C-SC-RA (UNIVERSAL TERMINATION BOARD FOR 16 HiC MODULES)", bx + 10.0, by + bh + 4.0, height=3.2, layer="TEXTS-CYAN")
    add_text(msp, "DIMENSIONS: 266 x 170 x 143 mm (W x H x D with modules)  |  DIN RAIL 35mm MOUNT  |  PART NO: 260436", bx + 10.0, by + bh - 6.0, height=2.2, layer="NOTATIONS")
    
    # Hazardous Field Area (Top Half) - Blue Screw Terminals
    add_rect(msp, bx + 10.0, by + bh - 32.0, bw - 20.0, 22.0, layer="TER-BLUE")
    add_text(msp, "HAZARDOUS AREA FIELD CONNECTIONS (BLUE SCREW TERMINALS) --- ZONE 0 / 1 / 20 / 21", bx + 15.0, by + bh - 16.0, height=2.4, layer="TEXTS-YELLOW")
    
    # Safe Control Area (Bottom Half) - Black Screw Terminals
    add_rect(msp, bx + 10.0, by + 10.0, bw - 20.0, 22.0, layer="TER-BLACK")
    add_text(msp, "SAFE AREA CONTROLLER CONNECTIONS (BLACK SCREW TERMINALS) --- TO CONTROLLOGIX 1756", bx + 15.0, by + 14.0, height=2.4, layer="TEXTS-CYAN")
    
    # 16 Module Slots (M1 to M16)
    slot_w = 15.0
    slot_gap = 1.5
    start_sx = bx + 15.0
    sy = by + 36.0
    sh = bh - 72.0
    
    for i in range(16):
        sx = start_sx + i * (slot_w + slot_gap)
        # Slot Outline
        add_rect(msp, sx, sy, slot_w, sh, layer="IS_BARRIERS")
        add_text(msp, f"M{i+1}", sx + 3.0, sy + sh - 8.0, height=2.2, layer="TEXTS-YELLOW")
        
        # Blue terminal pins (top)
        add_rect(msp, sx + 1.0, by + bh - 28.0, 6.0, 6.0, layer="TER-BLUE")
        add_rect(msp, sx + 8.0, by + bh - 28.0, 6.0, 6.0, layer="TER-BLUE")
        add_rect(msp, sx + 1.0, by + bh - 20.0, 6.0, 6.0, layer="TER-BLUE")
        add_rect(msp, sx + 8.0, by + bh - 20.0, 6.0, 6.0, layer="TER-BLUE")
        
        # Black terminal pins (bottom)
        add_rect(msp, sx + 1.0, by + 14.0, 6.0, 6.0, layer="TER-BLACK")
        add_rect(msp, sx + 8.0, by + 14.0, 6.0, 6.0, layer="TER-BLACK")
        add_rect(msp, sx + 1.0, by + 22.0, 6.0, 6.0, layer="TER-BLACK")
        add_rect(msp, sx + 8.0, by + 22.0, 6.0, 6.0, layer="TER-BLACK")
        
        # Module assignment sample
        if i in (0, 1, 2, 3):
            add_text(msp, "HiC2821", sx + 1.0, sy + 35.0, height=1.6, layer="TEXTS-GREEN")
            add_text(msp, "[DI]", sx + 4.0, sy + 25.0, height=1.8, layer="TEXTS")
        elif i in (4, 5, 6, 7):
            add_text(msp, "HiC2871", sx + 1.0, sy + 35.0, height=1.6, layer="TEXTS-YELLOW")
            add_text(msp, "[DO]", sx + 4.0, sy + 25.0, height=1.8, layer="TEXTS")
        elif i in (8, 9, 10, 11, 12, 13):
            add_text(msp, "HiC2025", sx + 1.0, sy + 35.0, height=1.6, layer="TEXTS-CYAN")
            add_text(msp, "[AI]", sx + 4.0, sy + 25.0, height=1.8, layer="TEXTS")
        else:
            add_text(msp, "SPARE", sx + 2.0, sy + 30.0, height=1.6, layer="NOTATIONS")
            
    # Redundant Power Feed Block X20 (Right side of board)
    pwx = bx + bw - 38.0
    pwy = sy + 15.0
    add_rect(msp, pwx, pwy, 32.0, 45.0, layer="0_BORDER")
    add_text(msp, "X20 POWER & FAULT", pwx + 2.0, pwy + 38.0, height=2.0, layer="TEXTS-YELLOW")
    add_text(msp, "PWR 1 (+) 24VDC", pwx + 3.0, pwy + 30.0, height=1.8, layer="IE-WIRE")
    add_text(msp, "PWR 1 (-) 0VDC", pwx + 3.0, pwy + 24.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "PWR 2 (+) 24VDC", pwx + 3.0, pwy + 18.0, height=1.8, layer="IE-WIRE")
    add_text(msp, "PWR 2 (-) 0VDC", pwx + 3.0, pwy + 12.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "FAULT RELAY (1/2)", pwx + 3.0, pwy + 5.0, height=1.8, layer="TEXTS-GREEN")
    
    # Grounding Rail (Bottom)
    msp.add_line((bx + 10.0, by + 4.0), (bx + bw - 10.0, by + 4.0), dxfattribs={"layer": "0_BORDER", "lineweight": 50})
    add_text(msp, "EQUIPOTENTIAL BONDING / INSTRUMENT GROUND BUS (IS GND)", bx + 60.0, by + 6.0, height=2.0, layer="NOTATIONS")
    
    # Specifications Block
    bx2, by2 = 20.0, 15.0
    add_rect(msp, bx2, by2, 180.0, 32.0, layer="0_BORDER")
    add_text(msp, "HiCTB16 TERMINATION BOARD TECHNICAL SPECIFICATIONS", bx2 + 3.0, by2 + 26.0, height=2.5, layer="TEXTS-YELLOW")
    add_text(msp, "• Capacity: 16 plug-in slots for any mix of HiC isolated barriers (DI, DO, AI, AO, Temperature, Frequency).", bx2 + 3.0, by2 + 20.0, height=2.0, layer="TEXTS")
    add_text(msp, "• Redundant Power: Dual diode-decoupled 24VDC feeds (PWR1 / PWR2) with 2x 2A slow-blow fuses.", bx2 + 3.0, by2 + 15.0, height=2.0, layer="TEXTS")
    add_text(msp, "• Integrated Fault Bus: Master fault alarm contact triggers upon power loss or field line fault (breakage/short).", bx2 + 3.0, by2 + 10.0, height=2.0, layer="TEXTS-GREEN")
    add_text(msp, "• Mechanical Coding: Unique pinout keys prevent insertion of wrong barrier type into safety slots.", bx2 + 3.0, by2 + 5.0, height=2.0, layer="NOTATIONS")
    
    doc.saveas(out_path)

# --------------------------------------------------------------------------------------
# 5. DRAWING 5: Master Intrinsic Safety System Loop Diagram (Panel CA-IS)
# --------------------------------------------------------------------------------------
def draw_master_is_system(out_path):
    doc, msp = setup_dxf_document()
    draw_title_block(msp, "PANEL CA-IS MASTER INTRINSIC SAFETY LOOP & MARSHALLING SCHEMATIC", "KAL-JC-IS-000", sheet_no=5, total_sheets=5, part_no="SYSTEM ARCHITECTURE")
    
    # 3 Distinct Plant Zones
    # Zone 1: Field Hazardous (X: 10 to 120)
    # Zone 2: Marshalling Panel CA-IS with HiCTB16 (X: 140 to 270)
    # Zone 3: Main Automation Cabinet CA1 / ControlLogix (X: 290 to 410)
    
    add_rect(msp, 15.0, 50.0, 110.0, 220.0, layer="HAZARDOUS_ZONE")
    add_text(msp, "ZONE 1: HAZARDOUS AREA (EX ZONE 0, 1, 20, 21)", 20.0, 263.0, height=2.8, layer="TEXTS-YELLOW")
    add_text(msp, "Spray Dryer Tower, Baghouse & Slurry Vessels", 20.0, 256.0, height=2.0, layer="NOTATIONS")
    
    add_rect(msp, 135.0, 50.0, 140.0, 220.0, layer="0_BORDER")
    add_text(msp, "ZONE 2: MARSHALLING PANEL CA-IS", 140.0, 263.0, height=2.8, layer="TEXTS-CYAN")
    add_text(msp, "Location: Field Ex Boundary (Spray Dryer 3rd Floor)", 140.0, 256.0, height=2.0, layer="NOTATIONS")
    
    add_rect(msp, 285.0, 50.0, 120.0, 220.0, layer="0_BORDER")
    add_text(msp, "ZONE 3: AUTOMATION CABINET (CA1 / RIO)", 290.0, 263.0, height=2.8, layer="TEXTS-GREEN")
    add_text(msp, "Rockwell ControlLogix 5580 / 1756 Chassis Racks", 290.0, 256.0, height=2.0, layer="NOTATIONS")
    
    # LOOP 1: DIGITAL INPUT (NAMUR / Dry Contact -> HiC2821 -> 1756-IB32)
    # Field
    add_rect(msp, 25.0, 200.0, 90.0, 40.0, layer="IE-EQUIP")
    add_text(msp, "ZS-60201A / B (Ex Limit Switch)", 30.0, 230.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "Spray Chamber Explosion Hatch Limit", 30.0, 223.0, height=1.8, layer="TEXTS")
    add_text(msp, "Zone 21 Combustible Dust (Ex tb IIIC)", 30.0, 215.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "Contact: Dry Contact with 1k/10k Resistors", 30.0, 207.0, height=1.8, layer="TEXTS-CYAN")
    
    # Barrier on HiCTB16
    add_rect(msp, 145.0, 200.0, 120.0, 40.0, layer="IS_BARRIERS")
    add_text(msp, "HiC2821 SWITCH AMPLIFIER (Slot M1 on HiCTB16)", 150.0, 230.0, height=2.2, layer="TEXTS-CYAN")
    add_text(msp, "Blue Term: SL2 5a/5b  |  Black Term: SL1 8a/7a (Relay)", 150.0, 223.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "SIL 2 Safety Barrier  |  Line Fault Detection Active", 150.0, 215.0, height=1.8, layer="TEXTS-GREEN")
    add_text(msp, "Power: 24VDC Bus from QUINT4 Redundant Bus", 150.0, 207.0, height=1.8, layer="IE-WIRE")
    
    # PLC Module
    add_rect(msp, 295.0, 200.0, 100.0, 40.0, layer="IE-EQUIP")
    add_text(msp, "1756-IB32 (DIGITAL INPUT)", 300.0, 230.0, height=2.2, layer="TEXTS-GREEN")
    add_text(msp, "Chassis C4 Slot 1  |  Tag: ZS-60201A_DI_1", 300.0, 223.0, height=1.8, layer="TEXTS-YELLOW")
    add_text(msp, "Terminal: P4-TBDI1-17  |  Pin: IN-16", 300.0, 215.0, height=1.8, layer="TEXTS-CYAN")
    add_text(msp, "Status: ACTIVE (Interlock to SCADA)", 300.0, 207.0, height=1.8, layer="NOTATIONS")
    
    # Connecting Wires Loop 1
    msp.add_line((115.0, 220.0), (145.0, 220.0), dxfattribs={"layer": "IE-WIRE-BLUE"})
    msp.add_line((265.0, 220.0), (295.0, 220.0), dxfattribs={"layer": "IE-WIRE"})
    
    # LOOP 2: DIGITAL OUTPUT (1756-OB32 -> HiC2871 -> Ex Solenoid Valve)
    # Field
    add_rect(msp, 25.0, 135.0, 90.0, 40.0, layer="IE-EQUIP")
    add_text(msp, "XV-61301 (Ex SOLENOID VALVE)", 30.0, 165.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "Spray Dryer Explosion Isolation Valve", 30.0, 158.0, height=1.8, layer="TEXTS")
    add_text(msp, "Zone 1 / 21 Hazardous Location", 30.0, 150.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "Coil: 12VDC / 45mA Intrinsically Safe", 30.0, 142.0, height=1.8, layer="TEXTS-CYAN")
    
    # Barrier on HiCTB16
    add_rect(msp, 145.0, 135.0, 120.0, 40.0, layer="IS_BARRIERS")
    add_text(msp, "HiC2871 SOLENOID DRIVER (Slot M5 on HiCTB16)", 150.0, 165.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "Blue Term: SL2 5a/5b (Ex) | Black Term: SL1 8a/7a (Ctrl)", 150.0, 158.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "SIL 3 Safety Rated  |  Loop Powered (No Bus Draw)", 150.0, 150.0, height=1.8, layer="TEXTS-GREEN")
    add_text(msp, "Controlled via PLC-RSC-24DC/21 Relay Contact", 150.0, 142.0, height=1.8, layer="RELAYS")
    
    # PLC Module
    add_rect(msp, 295.0, 135.0, 100.0, 40.0, layer="IE-EQUIP")
    add_text(msp, "1756-OB32 (DIGITAL OUTPUT)", 300.0, 165.0, height=2.2, layer="TEXTS-GREEN")
    add_text(msp, "Chassis C4 Slot 4  |  Tag: XV-61301_DO_1", 300.0, 158.0, height=1.8, layer="TEXTS-YELLOW")
    add_text(msp, "Interposing Relay: PLC-RSC-24DC/21", 300.0, 150.0, height=1.8, layer="RELAYS")
    add_text(msp, "Action: De-energize to Trip (Failsafe)", 300.0, 142.0, height=1.8, layer="NOTATIONS")
    
    # Connecting Wires Loop 2
    msp.add_line((115.0, 155.0), (145.0, 155.0), dxfattribs={"layer": "IE-WIRE-BLUE"})
    msp.add_line((265.0, 155.0), (295.0, 155.0), dxfattribs={"layer": "IE-WIRE"})
    
    # LOOP 3: ANALOG INPUT (4-20mA HART Transmitter -> HiC2025 -> 1756-IF16)
    # Field
    add_rect(msp, 25.0, 70.0, 90.0, 40.0, layer="IE-EQUIP")
    add_text(msp, "TT-60201 / PT-60201 (SMART Tx)", 30.0, 100.0, height=2.2, layer="TEXTS-YELLOW")
    add_text(msp, "Spray Chamber Exhaust Temperature", 30.0, 93.0, height=1.8, layer="TEXTS")
    add_text(msp, "2-Wire 4-20mA HART Communication", 30.0, 85.0, height=1.8, layer="TEXTS-CYAN")
    add_text(msp, "ATEX Ex ia IIC T6 Ga (Intrinsic Safety)", 30.0, 77.0, height=1.8, layer="NOTATIONS")
    
    # Barrier on HiCTB16
    add_rect(msp, 145.0, 70.0, 120.0, 40.0, layer="IS_BARRIERS")
    add_text(msp, "HiC2025 SMART Tx REPEATER (Slot M9 on HiCTB16)", 150.0, 100.0, height=2.2, layer="TEXTS-CYAN")
    add_text(msp, "Blue Term: SL2 5a/5b (Tx Pwr) | Black: SL1 8a/7a (4-20mA)", 150.0, 93.0, height=1.8, layer="NOTATIONS")
    add_text(msp, "SIL 2 (SC 3) Certified  |  Bi-directional HART Pass", 150.0, 85.0, height=1.8, layer="TEXTS-GREEN")
    add_text(msp, "Output Mode: 4-20mA Current Source to 1756-IF16", 150.0, 77.0, height=1.8, layer="TEXTS-YELLOW")
    
    # PLC Module
    add_rect(msp, 295.0, 70.0, 100.0, 40.0, layer="IE-EQUIP")
    add_text(msp, "1756-IF16 (ANALOG INPUT)", 300.0, 100.0, height=2.2, layer="TEXTS-GREEN")
    add_text(msp, "Chassis C4 Slot 5  |  Tag: TT-60201_AI_1", 300.0, 93.0, height=1.8, layer="TEXTS-YELLOW")
    add_text(msp, "Input: 4-20mA Single-Ended / Differential", 300.0, 85.0, height=1.8, layer="TEXTS-CYAN")
    add_text(msp, "Range: 0 - 250 °C (Calibrated Scale)", 300.0, 77.0, height=1.8, layer="NOTATIONS")
    
    # Connecting Wires Loop 3
    msp.add_line((115.0, 90.0), (145.0, 90.0), dxfattribs={"layer": "IE-WIRE-BLUE"})
    msp.add_line((265.0, 90.0), (295.0, 90.0), dxfattribs={"layer": "IE-WIRE"})
    
    doc.saveas(out_path)

def main():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    out_dir_cad = os.path.join(base_dir, "02_Electrical_and_eDrawing", "03_CAD_Exports_DXF_SVG", "IS_Barriers_CAD")
    out_dir_manuals = os.path.join(base_dir, "02_Electrical_and_eDrawing", "05_Instrument_Manuals", "Pepperl+Fuchs IS Barriers")
    
    os.makedirs(out_dir_cad, exist_ok=True)
    os.makedirs(out_dir_manuals, exist_ok=True)
    
    drawings = [
        ("KAL-JC-IS-001_HiC2821_Switch_Amplifier_DI.dxf", draw_hic2821),
        ("KAL-JC-IS-002_HiC2871_Solenoid_Driver_DO.dxf", draw_hic2871),
        ("KAL-JC-IS-003_HiC2025_SMART_Transmitter_Power_Supply_AI.dxf", draw_hic2025),
        ("KAL-JC-IS-004_HiCTB16_16Slot_Termination_Board.dxf", draw_hictb16),
        ("KAL-JC-IS-000_IS_Marshaling_Panel_CA-IS_Loop_Diagram.dxf", draw_master_is_system),
    ]
    
    print("==================================================================")
    print("GENERATING PEPPERL+FUCHS INTRINSICALLY SAFE CAD DXF DRAWINGS")
    print("==================================================================")
    
    for fname, func in drawings:
        p_cad = os.path.join(out_dir_cad, fname)
        p_man = os.path.join(out_dir_manuals, fname)
        
        func(p_cad)
        shutil.copy2(p_cad, p_man)
        print(f" -> Generated DXF: {fname} ({os.path.getsize(p_cad):,} bytes)")

    print("\n------------------------------------------------------------------")
    print("EXPORTING VECTOR SVG PREVIEWS FOR CAD / WEB VIEWING")
    print("------------------------------------------------------------------")
    from ezdxf.addons.drawing import RenderContext, Frontend
    from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
    import matplotlib.pyplot as plt

    for fname, _ in drawings:
        p_cad = os.path.join(out_dir_cad, fname)
        svg_name = fname.replace(".dxf", ".svg")
        p_svg_cad = os.path.join(out_dir_cad, svg_name)
        p_svg_man = os.path.join(out_dir_manuals, svg_name)
        
        doc = ezdxf.readfile(p_cad)
        msp = doc.modelspace()
        fig = plt.figure(figsize=(16, 11), dpi=150)
        ax = fig.add_axes([0, 0, 1, 1])
        ctx = RenderContext(doc)
        out = MatplotlibBackend(ax)
        Frontend(ctx, out).draw_layout(msp, finalize=True)
        fig.savefig(p_svg_cad, format='svg')
        plt.close(fig)
        
        shutil.copy2(p_svg_cad, p_svg_man)
        print(f" -> Exported SVG: {svg_name} ({os.path.getsize(p_svg_cad):,} bytes)")

    print("\nSUCCESS: All 5 CAD DXF & SVG Drawings generated in both:")
    print(f"  1. {out_dir_cad}")
    print(f"  2. {out_dir_manuals}")
    print("==================================================================")

if __name__ == "__main__":
    main()

