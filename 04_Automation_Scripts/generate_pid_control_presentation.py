#!/usr/bin/env python3
"""
Generate PID Loop Control Specification & Architecture Presentation Slides
Customer: Ingredion (Thailand) Co., Ltd. - Kalasin Plant
Project: SPRINT 18K TPA Spray Dryer & Jet Cooker Facility
Platform: Allen-Bradley ControlLogix 5580 (PIDE Enhanced PID) & FactoryTalk View SE
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, ImageFont

# Workspace Output Paths
WORKSPACE_DIR = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List"
SCADA_DIR = os.path.join(WORKSPACE_DIR, "SCADA_Demo_Video")
os.makedirs(SCADA_DIR, exist_ok=True)

PPTX_PATH_1 = os.path.join(WORKSPACE_DIR, "PID_Loop_Control_Specification_Kalasin.pptx")
PPTX_PATH_2 = os.path.join(SCADA_DIR, "PID_Loop_Control_Specification_Kalasin.pptx")
PNG_PATH_1 = os.path.join(WORKSPACE_DIR, "PID_Loop_Control_Architecture_Overview.png")
PNG_PATH_2 = os.path.join(SCADA_DIR, "PID_Loop_Control_Architecture_Overview.png")

# Industrial High-Tech Color Palette
DARK_BG = RGBColor(15, 23, 42)        # Slate 900
PANEL_BG = RGBColor(24, 34, 53)       # Slate 850
CARD_BG = RGBColor(30, 41, 59)        # Slate 800
CARD_BG_LIGHT = RGBColor(38, 52, 75)  # Slate 750
CARD_BORDER = RGBColor(51, 65, 85)    # Slate 700
ACCENT_BLUE = RGBColor(56, 189, 248)  # Cyan/Sky 400
ACCENT_GREEN = RGBColor(34, 197, 94)  # Emerald 500
ACCENT_AMBER = RGBColor(245, 158, 11) # Amber 500
ACCENT_ROSE = RGBColor(244, 63, 94)   # Rose 500
ACCENT_PURPLE = RGBColor(168, 85, 247)# Purple 500
ACCENT_TEAL = RGBColor(20, 184, 166)  # Teal 500
ACCENT_CYAN = RGBColor(6, 182, 212)   # Cyan 500
WHITE = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(203, 213, 225)  # Slate 300
TEXT_MUTED = RGBColor(148, 163, 184)  # Slate 400


def add_slide_header(slide, title_text, category_text="PROCESS AUTOMATION & INSTRUMENTATION SPECIFICATION"):
    """Helper to add consistent top header ribbon to slides"""
    header_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.75))
    header_box.fill.solid()
    header_box.fill.fore_color.rgb = CARD_BG
    header_box.line.color.rgb = CARD_BORDER
    header_box.line.width = Pt(1)

    tf = header_box.text_frame
    tf.word_wrap = True
    p1 = tf.paragraphs[0]
    p1.text = f"{category_text}  |  INGREDION KALASIN — PROJECT SPRINT"
    p1.font.size = Pt(9)
    p1.font.bold = True
    p1.font.color.rgb = ACCENT_BLUE
    p1.alignment = PP_ALIGN.LEFT

    p2 = tf.add_paragraph()
    p2.text = title_text
    p2.font.size = Pt(15)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    p2.space_before = Pt(2)
    p2.alignment = PP_ALIGN.LEFT


def create_presentation():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # =========================================================================
    # SLIDE 1: Title & Executive Summary
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BG
    bg1.line.fill.background()

    # Title Card
    tb = s1.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.733), Inches(3.6))
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "INGREDION (THAILAND) CO., LTD.  —  KALASIN PLANT"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE

    p2 = tf.add_paragraph()
    p2.text = "Master PID Loop Control Specification"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    p2.space_before = Pt(8)

    p3 = tf.add_paragraph()
    p3.text = "Project SPRINT: 18,000 TPA Spray Dryer Plant & Continuous Jet Cooker Unit"
    p3.font.size = Pt(20)
    p3.font.color.rgb = LIGHT_GRAY
    p3.space_before = Pt(6)

    p4 = tf.add_paragraph()
    p4.text = "Complete Process Control Loop Architecture: 16 PID Loops | 5 Hardwired 4-20mA AO Valves & Burner Demand | 10 EtherNet/IP VFD Speed Controls | 1 Remote Utility AO | ControlLogix 5580 PIDE Platform"
    p4.font.size = Pt(13)
    p4.font.bold = True
    p4.font.color.rgb = ACCENT_GREEN
    p4.space_before = Pt(14)

    # Executive Overview Cards (4 Grid Cards)
    cards_data = [
        ("🎯 CORE THERMAL LOOPS", "5 Critical PID Loops", "Direct Steam Jet Cooking (TCV-40222), Drying Air Inlet Temp (UY-FH-60210B), Pre-Heat Exchanger (TCV-60214), Steam Header (PCV-60201) & Camera Chilled Water (TCV-60202).", ACCENT_AMBER),
        ("🛡️ DRAFT & AIR FLOW", "4 Fast-Acting Loops", "Drying Chamber Draft Pressure (PIC-60220 / FANE-21 VFD), Main Supply Air Flow (FIC-60210 / FANS-21), Burner Combustion Air Ratio (FIC-60211) & Conveying Fan (FIC-60201).", ACCENT_BLUE),
        ("💧 QUALITY & MOISTURE", "4 Process Loops", "Chamber Outlet Temp / Moisture Cascade (TIC-60226 to PUPD-11 Slurry Pump VFD), Jet Cooker Slurry Mass Flow (FIC-40203) & pH Conditioning (pHTC-40201/40202).", ACCENT_TEAL),
        ("⚡ HARDWARE ARCHITECTURE", "ControlLogix 5580", "1756-OF8 HART Module (Chassis C3 Slot 13), 20 PowerFlex VFDs over EtherNet/IP DLR (Chassis C7 Bus), 1756-IF16 HART Inputs & FT View SE Faceplates.", ACCENT_PURPLE)
    ]

    left_positions = [Inches(0.8), Inches(3.8), Inches(6.8), Inches(9.8)]
    for i, (title, sub, body, color) in enumerate(cards_data):
        card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_positions[i], Inches(4.5), Inches(2.733), Inches(2.0))
        card.fill.solid()
        card.fill.fore_color.rgb = CARD_BG
        card.line.color.rgb = color
        card.line.width = Pt(1.5)

        ctf = card.text_frame
        ctf.word_wrap = True
        ctf.margin_top = Inches(0.12)
        ctf.margin_left = Inches(0.12)
        ctf.margin_right = Inches(0.12)

        cp1 = ctf.paragraphs[0]
        cp1.text = title
        cp1.font.size = Pt(11)
        cp1.font.bold = True
        cp1.font.color.rgb = color

        cp2 = ctf.add_paragraph()
        cp2.text = sub
        cp2.font.size = Pt(10)
        cp2.font.bold = True
        cp2.font.color.rgb = WHITE
        cp2.space_before = Pt(3)

        cp3 = ctf.add_paragraph()
        cp3.text = body
        cp3.font.size = Pt(8.5)
        cp3.font.color.rgb = LIGHT_GRAY
        cp3.space_before = Pt(4)

    # Footer
    meta_box = s1.shapes.add_textbox(Inches(0.8), Inches(6.75), Inches(11.733), Inches(0.5))
    meta_tf = meta_box.text_frame
    mp = meta_tf.paragraphs[0]
    mp.text = "Document Ref: PRJ-2603001-PID-SPEC-R01  |  Engineering Consultant: AEC Industrial Engineering  |  Rev 3.6 Master Aligned"
    mp.font.size = Pt(10)
    mp.font.color.rgb = TEXT_MUTED

    # =========================================================================
    # SLIDE 2: Process Control Overview & Master Loop Mapping
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    bg2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = DARK_BG
    bg2.line.fill.background()

    add_slide_header(s2, "PLANT PROCESS SCHEMATIC & PID CONTROL LOOP ALLOCATION MATRIX")

    # 4 Main Process Zones (Horizontal Layout)
    zones = [
        ("ZONE 1 : SLURRY PREP & JET COOKER", Inches(0.5), Inches(1.2), Inches(3.0), Inches(5.8), ACCENT_TEAL, [
            ("pHTC-40201 / 40202", "Slurry pH Adjustment", "PV: pHT-40201/40202\nCV: Dosing Pumps PD-40204/05\nTarget: 5.8 - 6.2 pH\nP&ID: 325-01-402-PD-01"),
            ("LIC-40203", "Slurry Infeed Tank Level", "PV: LT-40203 (0-100%)\nCV: Upstream Feed Valve/Pump\nTarget: 60 - 75% Buffer\nP&ID: 325-01-402-PD-02"),
            ("FIC-40203", "Jet Cooker Slurry Mass Flow", "PV: FT-40221/40222 Coriolis\nCV: PUCF-11 Feed Pump VFD\nTarget: 4.5 - 7.5 t/h\nP&ID: 325-01-402-PD-02"),
            ("TIC-40222 [CRITICAL]", "Jet Cooker Steam Temp", "PV: TT-40221/40222 Duplex RTD\nCV: TCV-40222 (Samson 3241)\nTarget: 120 - 140 °C (Cooking)\nOutput: 4-20mA (C3S13,00)")
        ]),
        ("ZONE 2 : DRYER AIR HEATING & STEAM", Inches(3.6), Inches(1.2), Inches(3.0), Inches(5.8), ACCENT_AMBER, [
            ("PIC-60201 [MASTER]", "Steam Header Pressure", "PV: PT-60201 (0-25 bar)\nCV: PCV-60201 (Samson NPS 4\")\nTarget: 16.0 bar(g) Steady\nOutput: 4-20mA (C3S13,03)"),
            ("TIC-60214", "Air Pre-Heater Temp (HXCH)", "PV: TT-60214 (0-150°C)\nCV: TCV-60214 (Samson NPS 2\")\nTarget: 75 - 85 °C Pre-heat\nOutput: 4-20mA (C3S13,02)"),
            ("TIC-60210 [THERMAL]", "Dryer Inlet Air Temperature", "PV: TT-60210/60211 (Hot Duct)\nCV: UY-FH-60210B (BMS Demand)\nTarget: 185 - 205 °C Air Supply\nOutput: 4-20mA (C3S13,04)"),
            ("FIC-60210 / 60211", "Dryer Supply Air & Comb. Flow", "PV: FT-60210/11 & PT-60216\nCV: FANS-21 & FANC-21 VFDs\nTarget: 32,000 Nm³/h & Ratio\nInterface: EtherNet/IP Bus")
        ]),
        ("ZONE 3 : DRYING CHAMBER & DRAFT", Inches(6.7), Inches(1.2), Inches(3.0), Inches(5.8), ACCENT_ROSE, [
            ("PIC-60220 [SAFETY]", "Chamber Draft Pressure", "PV: PT-60220 (-10..+10 mbar)\nCV: FANE-21 Exhaust Fan VFD\nTarget: -1.2 mbar (Negative Draft)\nFast-acting anti-puff loop"),
            ("TIC-60226 / FIC-60205", "Exhaust Temp / Moisture Cascade", "PV: TT-60226 / TT-60240\nCV: PUPD-11 Slurry Pump VFD\nTarget: 88 - 92 °C Outlet Air\nControls powder moisture ~6%"),
            ("TIC-60202", "Camera Purge Air Chilled Water", "PV: TT-60202 (0-60°C)\nCV: TCV-60202 (3-Way Samson)\nTarget: 22 - 25 °C Lens Cool\nOutput: 4-20mA (C3S13,01)"),
            ("TIC-60215", "Glycol Heat Recovery Loop", "PV: TT-60215A/B Glycol Temp\nCV: PUCF-21 Circulation Pump\nTarget: 60 °C Energy Recapture\nExhaust Dewpoint Protection")
        ]),
        ("ZONE 4 : CONVEYING, LEV & UTILITY", Inches(9.8), Inches(1.2), Inches(3.0), Inches(5.8), ACCENT_PURPLE, [
            ("FIC-60201", "Powder Cooling Air Flow", "PV: FT-60201 & PT-60203\nCV: FANS-31 Conveying Fan VFD\nTarget: 18 - 22 m/s Velocity\nPneumatic Cyclone Discharge"),
            ("PIC-61302", "Silo & LEV Dust Extraction", "PV: PT-61302 Vacuum Line\nCV: BF-61302 LEV Blower VFD\nTarget: -15 mbar Negative Head\nATEX Dust Containment"),
            ("LIC-61301", "Packing Silo Level Balance", "PV: LT-61301 Radar Level\nCV: Cyclone Rotary Valves\nTarget: 20 - 80% Capacity\nContinuous Bagging Feed"),
            ("PIC-81001", "Boiler Remote Setpoint Tracking", "PV: Total Steam Demand (FT-60202)\nCV: External Remote Setpoint\nTarget: Dynamic Steam Sync\nOutput: 4-20mA (RIO-200)")
        ])
    ]

    for z_title, z_left, z_top, z_width, z_height, z_color, loop_list in zones:
        z_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, z_left, z_top, z_width, z_height)
        z_box.fill.solid()
        z_box.fill.fore_color.rgb = PANEL_BG
        z_box.line.color.rgb = z_color
        z_box.line.width = Pt(1.5)

        zh_tf = z_box.text_frame
        zh_tf.word_wrap = True
        zh_tf.margin_top = Inches(0.08)
        zh_tf.margin_left = Inches(0.1)
        zh_tf.margin_right = Inches(0.1)
        zh_p = zh_tf.paragraphs[0]
        zh_p.text = z_title
        zh_p.font.size = Pt(9.5)
        zh_p.font.bold = True
        zh_p.font.color.rgb = z_color

        card_top = z_top + Inches(0.4)
        c_height = Inches(1.25)
        for l_tag, l_name, l_details in loop_list:
            l_card = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, z_left + Inches(0.1), card_top, z_width - Inches(0.2), c_height)
            l_card.fill.solid()
            l_card.fill.fore_color.rgb = CARD_BG
            l_card.line.color.rgb = CARD_BORDER
            l_card.line.width = Pt(1)

            ltf = l_card.text_frame
            ltf.word_wrap = True
            ltf.margin_top = Inches(0.06)
            ltf.margin_left = Inches(0.08)
            ltf.margin_right = Inches(0.08)
            ltf.margin_bottom = Inches(0.04)

            lp1 = ltf.paragraphs[0]
            lp1.text = l_tag
            lp1.font.size = Pt(9.5)
            lp1.font.bold = True
            lp1.font.color.rgb = WHITE

            lp2 = ltf.add_paragraph()
            lp2.text = l_name
            lp2.font.size = Pt(8)
            lp2.font.bold = True
            lp2.font.color.rgb = z_color

            lp3 = ltf.add_paragraph()
            lp3.text = l_details
            lp3.font.size = Pt(7.5)
            lp3.font.color.rgb = LIGHT_GRAY

            card_top += c_height + Inches(0.08)

    # =========================================================================
    # SLIDE 3: Area 402 Deep Dive (Jet Cooker & Slurry Preparation)
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    bg3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg3.fill.solid()
    bg3.fill.fore_color.rgb = DARK_BG
    bg3.line.fill.background()

    add_slide_header(s3, "AREA 402 : CONTINUOUS JET COOKER & SLURRY PREPARATION CONTROL LOOPS")

    card_l3 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.2), Inches(6.0), Inches(5.8))
    card_l3.fill.solid()
    card_l3.fill.fore_color.rgb = PANEL_BG
    card_l3.line.color.rgb = ACCENT_AMBER
    card_l3.line.width = Pt(1.5)

    tf_l3 = card_l3.text_frame
    tf_l3.word_wrap = True
    tf_l3.margin_left = Inches(0.2)
    tf_l3.margin_right = Inches(0.2)
    tf_l3.margin_top = Inches(0.15)

    p_l3_h = tf_l3.paragraphs[0]
    p_l3_h.text = "JET COOKER THERMAL & HYDRAULIC CONTROL (P&ID: 325-01-402-PD-02)"
    p_l3_h.font.size = Pt(11)
    p_l3_h.font.bold = True
    p_l3_h.font.color.rgb = ACCENT_AMBER

    inner_boxes_l3 = [
        ("TIC-40222 : Jet Cooker Direct Steam Cooking Temperature", [
            ("Process Variable (PV)", "Dual High-Precision RTDs TT-40221 & TT-40222 (Duplex Pt100, 0–150°C)"),
            ("Set Point (SP)", "Nominal 120.0 °C to 140.0 °C (Recipe selectable for complete starch gelatinization)"),
            ("Control Variable (CV)", "TCV-40222 (Samson Type 3241 Globe Valve w/ Trovis 3730-1 Positioner)"),
            ("Hardware Interface", "4–20 mA Analog Output @ Chassis C3 Slot 13 Channel 00 (TCV-40222_AO_1)"),
            ("Fail-Safe Action", "Fail Closed (FC) / Spring Closes on power or air loss. High-speed steam shutoff"),
            ("Control Philosophy", "Fast-acting reverse-acting PID (Increase steam increases temp). PIDE executed at 100ms periodic task with derivative filter. High/Low deviation alarms (+/- 2.0°C) with auto-interlock to slurry feed pump to prevent unburst starch discharge.")
        ]),
        ("FIC-40203 : Jet Cooker Slurry Mass Feed Flow Rate Control", [
            ("Process Variable (PV)", "Coriolis Mass Flowmeter FT-40221 / FT-40222 (Dual channel mass flow + density)"),
            ("Set Point (SP)", "4.50 to 7.50 t/h (Calibrated for exact reactor residence time)"),
            ("Control Variable (CV)", "PCY-40203 (Centrifugal Slurry Feed Pump PUCF-11 VFD Speed Reference)"),
            ("Hardware Interface", "EtherNet/IP Industrial Bus link to MCC PowerFlex Drive (Chassis C7 Slot 00)"),
            ("Control Philosophy", "Maintains constant mass flow regardless of slurry viscosity fluctuations. Tightly cascaded with Jet Cooker downstream infeed level (LT-40203) and interlocked with minimum flow switch FS-40205 to prevent pump cavitation.")
        ])
    ]

    top_l3 = Inches(1.6)
    for title, items in inner_boxes_l3:
        b = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.7), top_l3, Inches(5.6), Inches(2.55))
        b.fill.solid()
        b.fill.fore_color.rgb = CARD_BG
        b.line.color.rgb = CARD_BORDER
        b.line.width = Pt(1)

        btf = b.text_frame
        btf.word_wrap = True
        btf.margin_left = Inches(0.12)
        btf.margin_right = Inches(0.12)
        btf.margin_top = Inches(0.08)

        bp = btf.paragraphs[0]
        bp.text = title
        bp.font.size = Pt(10)
        bp.font.bold = True
        bp.font.color.rgb = WHITE

        for k, v in items:
            p_item = btf.add_paragraph()
            p_item.text = f"• {k}: {v}"
            p_item.font.size = Pt(8)
            p_item.font.color.rgb = LIGHT_GRAY
            p_item.space_before = Pt(2)

        top_l3 += Inches(2.65)

    card_r3 = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.2), Inches(6.0), Inches(5.8))
    card_r3.fill.solid()
    card_r3.fill.fore_color.rgb = PANEL_BG
    card_r3.line.color.rgb = ACCENT_TEAL
    card_r3.line.width = Pt(1.5)

    tf_r3 = card_r3.text_frame
    tf_r3.word_wrap = True
    tf_r3.margin_left = Inches(0.2)
    tf_r3.margin_right = Inches(0.2)
    tf_r3.margin_top = Inches(0.15)

    p_r3_h = tf_r3.paragraphs[0]
    p_r3_h.text = "SLURRY CONDITIONING & LEVEL BUFFER (P&ID: 325-01-402-PD-01)"
    p_r3_h.font.size = Pt(11)
    p_r3_h.font.bold = True
    p_r3_h.font.color.rgb = ACCENT_TEAL

    inner_boxes_r3 = [
        ("pHTC-40201 & pHTC-40202 : pH Adjustment Tank 1 & Tank 2 Loops", [
            ("Process Variable (PV)", "pHT-40201 (Tank 1) & pHT-40202 (Tank 2) — Mettler Toledo InPro 3250i + M300"),
            ("Set Point (SP)", "5.80 to 6.20 pH (Strict enzymatic / processing window for modified starch)"),
            ("Control Variable (CV)", "Acid / Caustic Dosing Pumps PD-40204 & PD-40205 (Stroke frequency / VSD)"),
            ("Hardware Interface", "Modulated discrete pulse dosing / 4-20mA speed over RIO/MCC bus"),
            ("Control Philosophy", "Dual-directional split-range neutralization PID. Non-linear gain scheduling applied around neutralization point (pH 7.0) to prevent severe reagent overshoot. Continuous automatic sensor cleaning cycle via SV-40201/02 water purge.")
        ]),
        ("LIC-40201 / 40202 / 40203 : Buffer Tank Level Regulation", [
            ("Process Variable (PV)", "Hydrostatic / Radar Level Transmitters LT-40201, LT-40202, LT-40203 (0–100%)"),
            ("Set Point (SP)", "65.0% Working Level (Provides 30-minute hydraulic buffer between units)"),
            ("Control Variable (CV)", "Slurry transfer valves XV-40201E / XV-40202F & transfer pumps PC-40201/02"),
            ("Safety Interlocks", "Low-level cutoff (LSL) prevents pump dry-run; High-level cutoff (LSH) trips upstream feed and sounds audible plant horn.")
        ])
    ]

    top_r3 = Inches(1.6)
    for title, items in inner_boxes_r3:
        b = s3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(7.0), top_r3, Inches(5.6), Inches(2.55))
        b.fill.solid()
        b.fill.fore_color.rgb = CARD_BG
        b.line.color.rgb = CARD_BORDER
        b.line.width = Pt(1)

        btf = b.text_frame
        btf.word_wrap = True
        btf.margin_left = Inches(0.12)
        btf.margin_right = Inches(0.12)
        btf.margin_top = Inches(0.08)

        bp = btf.paragraphs[0]
        bp.text = title
        bp.font.size = Pt(10)
        bp.font.bold = True
        bp.font.color.rgb = WHITE

        for k, v in items:
            p_item = btf.add_paragraph()
            p_item.text = f"• {k}: {v}"
            p_item.font.size = Pt(8)
            p_item.font.color.rgb = LIGHT_GRAY
            p_item.space_before = Pt(2)

        top_r3 += Inches(2.65)

    # =========================================================================
    # SLIDE 4: Area 602 Deep Dive (Spray Dryer Thermal & Combustion Loops)
    # =========================================================================
    s4 = prs.slides.add_slide(blank_layout)
    bg4 = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg4.fill.solid()
    bg4.fill.fore_color.rgb = DARK_BG
    bg4.line.fill.background()

    add_slide_header(s4, "AREA 602 : SPRAY DRYER THERMAL & COMBUSTION CONTROL LOOPS")

    card_w = Inches(3.9)
    card_gap = Inches(0.3)
    card_h = Inches(5.8)

    c4_data = [
        ("TIC-60210 : DRYER INLET AIR TEMP", ACCENT_ROSE, [
            ("Process Function", "Master Thermal Drying Energy Loop"),
            ("Process Variable (PV)", "TT-60210 & TT-60211 (Dual Duplex Thermocouples in Hot Air Duct DTHA-21, 0–300°C)"),
            ("Set Point (SP)", "185.0 °C to 205.0 °C (Product recipe dependent)"),
            ("Control Variable (CV)", "UY-FH-60210B (BMS Burner Demand Firing Rate 0–100%)"),
            ("Hardware Channel", "4–20 mA AO @ Chassis C3 Slot 13 Channel 04"),
            ("Actuator / Skid", "Direct-Fired Gas Air Heater AHTR-21 + Burner Controller"),
            ("Fail-Safe State", "0% Firing Demand / Immediate Burner Trip on E-Stop"),
            ("Tuning / Strategy", "High thermal inertia loop. Feedforward compensation from Main Airflow (FT-60210) dynamically adjusts fuel demand before duct temp sags. Interlocked with Burner Flame Safeguard.")
        ]),
        ("TIC-60214 : AIR PRE-HEATER STEAM", ACCENT_AMBER, [
            ("Process Function", "Fresh Air Steam Pre-Heating (Energy Recovery)"),
            ("Process Variable (PV)", "TT-60214 (PT100 Transmitter downstream of Heat Exchanger HXCH-21, 0–150°C)"),
            ("Set Point (SP)", "75.0 °C to 85.0 °C (Pre-heat base temperature)"),
            ("Control Variable (CV)", "TCV-60214 (Samson Type 3241 Globe Valve, NPS 2\", Class 300, Cv 47 + Trovis 3730-1)"),
            ("Hardware Channel", "4–20 mA AO @ Chassis C3 Slot 13 Channel 02"),
            ("Actuator / Skid", "Shell & Tube Steam Heat Exchanger HXCH-21"),
            ("Fail-Safe State", "Fail Closed (FC) via internal mechanical spring"),
            ("Tuning / Strategy", "Maximizes steam condensate energy before air enters gas burner, reducing LPG fuel consumption by up to 28%. Smooth PI control prevents thermal shock to exchanger tubes.")
        ]),
        ("PIC-60201 : MAIN STEAM HEADER", ACCENT_BLUE, [
            ("Process Function", "Plant High-Pressure Steam Pressure Reducing Station"),
            ("Process Variable (PV)", "PT-60201 (Endress+Hauser Cerabar PMP51B, calibrated 0–25 bar(g)) + TT-60201 Temp"),
            ("Set Point (SP)", "16.0 bar(g) (Stable header supply pressure)"),
            ("Control Variable (CV)", "PCV-60201 (Samson Type 3241 Globe Valve, NPS 4\", Class 300, Cv 190 + Trovis 3730-1)"),
            ("Hardware Channel", "4–20 mA AO @ Chassis C3 Slot 13 Channel 03"),
            ("Actuator / Skid", "Main Steam Header PRV Station (Area 602-PD-02)"),
            ("Fail-Safe State", "Fail Closed (FC) / Tight Shut-off Class IV"),
            ("Tuning / Strategy", "Fast-response pressure loop (50ms task). Mitigates boiler pressure surges (16–20 bar) down to stable 16.0 bar for the dryer coils and jet cooker, eliminating thermal hunting.")
        ])
    ]

    left_c4 = Inches(0.5)
    for title, color, items in c4_data:
        box = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_c4, Inches(1.2), card_w, card_h)
        box.fill.solid()
        box.fill.fore_color.rgb = PANEL_BG
        box.line.color.rgb = color
        box.line.width = Pt(1.5)

        btf = box.text_frame
        btf.word_wrap = True
        btf.margin_left = Inches(0.15)
        btf.margin_right = Inches(0.15)
        btf.margin_top = Inches(0.12)

        bp = btf.paragraphs[0]
        bp.text = title
        bp.font.size = Pt(11)
        bp.font.bold = True
        bp.font.color.rgb = color

        for k, v in items:
            p_item = btf.add_paragraph()
            p_item.text = f"{k}:"
            p_item.font.size = Pt(8.5)
            p_item.font.bold = True
            p_item.font.color.rgb = WHITE
            p_item.space_before = Pt(3)

            p_desc = btf.add_paragraph()
            p_desc.text = v
            p_desc.font.size = Pt(8)
            p_desc.font.color.rgb = LIGHT_GRAY

        left_c4 += card_w + card_gap

    # =========================================================================
    # SLIDE 5: Area 602 Deep Dive (Chamber Draft, Quality & Moisture Cascade)
    # =========================================================================
    s5 = prs.slides.add_slide(blank_layout)
    bg5 = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg5.fill.solid()
    bg5.fill.fore_color.rgb = DARK_BG
    bg5.line.fill.background()

    add_slide_header(s5, "AREA 602 : CHAMBER DRAFT, PRODUCT MOISTURE & OPTICAL COOLING LOOPS")

    c5_data = [
        ("PIC-60220 : CHAMBER DRAFT PRESSURE", ACCENT_ROSE, [
            ("Process Function", "Dryer Explosion & Dust Containment Safety Loop"),
            ("Process Variable (PV)", "PT-60220 (Chamber Differential Draft, -10.0 to +10.0 mbar range)"),
            ("Set Point (SP)", "-1.20 mbar (-120 Pa slight negative pressure)"),
            ("Control Variable (CV)", "BFY-60240 (Main Exhaust Fan FANE-21 VFD Speed Reference 0–50 Hz)"),
            ("Hardware Channel", "EtherNet/IP Communication Bus to MCC (Chassis C7 Slot 00)"),
            ("Safety Interlocks", "PSL-60220 Low Pressure Switch & PSH-60220 Chamber Overpressure Trip"),
            ("Fail-Safe State", "Auto-Ramp to safe draft speed; immediate interlock trip on duct blockage"),
            ("Control Strategy", "Highest safety priority in spray drying. If chamber goes positive (+mbar), hot flammable starch dust leaks into tower structure. High-speed anti-windup PIDE executes at 50ms with non-linear error filtering.")
        ]),
        ("TIC-60226 : OUTLET TEMP / MOISTURE CASCADE", ACCENT_GREEN, [
            ("Process Function", "Final Starch Powder Moisture Regulation (~5.5-6.5%)"),
            ("Process Variable (PV)", "TT-60226 (Chamber Discharge Temp) & TT-60240 (Baghouse Inlet, 0–150°C)"),
            ("Set Point (SP)", "88.0 °C to 92.0 °C (Directly correlates to powder residual moisture)"),
            ("Control Variable (CV)", "PCY-60205 (High-Pressure Slurry Feed Pump PUPD-11 VFD Speed Reference)"),
            ("Secondary Feedback", "16 Ultrasonic Flow Transmitters FT-60261FA..FT-60264FR (individual nozzles)"),
            ("Hardware Channel", "EtherNet/IP Bus link to High Pressure Pump VFD (Area 602-PD-01)"),
            ("Fail-Safe State", "Auto-cutback slurry feed upon high chamber moisture or flame loss"),
            ("Control Strategy", "Master/Slave Cascade. Master temp loop adjusts slurry feed flow rate: If outlet temp drops (powder too wet), feed rate is automatically reduced; if outlet temp climbs, feed rate is increased.")
        ]),
        ("TIC-60202 : CAMERA AIR CHILLED WATER", ACCENT_CYAN, [
            ("Process Function", "In-Chamber Optical Inspection Protection Loop"),
            ("Process Variable (PV)", "TT-60202 (Purge Air Temperature to High-Temp Camera ACAM-21, 0–60°C)"),
            ("Set Point (SP)", "22.0 °C to 25.0 °C (Maintains camera housing below electronics limit)"),
            ("Control Variable (CV)", "TCV-60202 / TV-60202 (Samson / ITQ40 3-Way Modulating Water Valve)"),
            ("Hardware Channel", "4–20 mA AO @ Chassis C3 Slot 13 Channel 01 (TV-60202_AO_1)"),
            ("Hardware Support", "FT-60207 (Purge Air Flow) & PT-60207 (Purge Air Pressure)"),
            ("Fail-Safe State", "Fail-to-Port A (100% Chilled Water bypass to heat exchanger on power loss)"),
            ("Control Strategy", "Standard PI loop modulating chilled water bypass. Prevents condensation fogging on optical lenses while protecting sensitive camera sensors from 200°C chamber ambient heat.")
        ])
    ]

    left_c5 = Inches(0.5)
    for title, color, items in c5_data:
        box = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_c5, Inches(1.2), card_w, card_h)
        box.fill.solid()
        box.fill.fore_color.rgb = PANEL_BG
        box.line.color.rgb = color
        box.line.width = Pt(1.5)

        btf = box.text_frame
        btf.word_wrap = True
        btf.margin_left = Inches(0.15)
        btf.margin_right = Inches(0.15)
        btf.margin_top = Inches(0.12)

        bp = btf.paragraphs[0]
        bp.text = title
        bp.font.size = Pt(11)
        bp.font.bold = True
        bp.font.color.rgb = color

        for k, v in items:
            p_item = btf.add_paragraph()
            p_item.text = f"{k}:"
            p_item.font.size = Pt(8.5)
            p_item.font.bold = True
            p_item.font.color.rgb = WHITE
            p_item.space_before = Pt(3)

            p_desc = btf.add_paragraph()
            p_desc.text = v
            p_desc.font.size = Pt(8)
            p_desc.font.color.rgb = LIGHT_GRAY

        left_c5 += card_w + card_gap

    # =========================================================================
    # SLIDE 6: Area 602/613 & Utilities (Heat Recovery, Powder Handling & Boiler)
    # =========================================================================
    s6 = prs.slides.add_slide(blank_layout)
    bg6 = s6.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg6.fill.solid()
    bg6.fill.fore_color.rgb = DARK_BG
    bg6.line.fill.background()

    add_slide_header(s6, "AREA 602/613 & UTILITIES : HEAT RECOVERY, CONVEYING & BOILER SYNC")

    grid_w = Inches(5.95)
    grid_h = Inches(2.7)
    positions_2x2 = [
        (Inches(0.5), Inches(1.2)),
        (Inches(6.8), Inches(1.2)),
        (Inches(0.5), Inches(4.1)),
        (Inches(6.8), Inches(4.1))
    ]

    c6_data = [
        ("TIC-60215 : GLYCOL HEAT RECOVERY RUN-AROUND LOOP", ACCENT_TEAL, [
            ("Process Service", "Exhaust Waste Heat Recapture Skid (PUCF-21 / HXCR-21, P&ID: 325-01-602-PD-03)"),
            ("Process Variable (PV)", "TT-60215A & TT-60215B (Glycol Supply & Return Temp) + PT-60215 Pressure"),
            ("Set Point (SP)", "55.0 °C to 65.0 °C (Maintains exhaust above acid/moisture condensation dew point)"),
            ("Control Variable (CV)", "PCY-60215 (Heat Recovery Centrifugal Circulation Pump PUCF-21 VFD Speed Reference)"),
            ("Control Strategy", "Modulates 15% polypropylene glycol circulation rate between exhaust heat coil and fresh air preheater. Recovers up to 1.8 MW of thermal energy without fouling exhaust stack.")
        ]),
        ("FIC-60201 : POWDER COOLING & PNEUMATIC CONVEYING", ACCENT_BLUE, [
            ("Process Service", "Cyclone Discharge Powder Chilling & Conveying (P&ID: 325-01-602-PD-08)"),
            ("Process Variable (PV)", "FT-60201 (Cold Conveying Air Mass Flow) & PT-60203 (Duct Transport Pressure)"),
            ("Set Point (SP)", "18.0 to 22.0 m/s Saltation Velocity (Prevents starch powder settling / line plugging)"),
            ("Control Variable (CV)", "BFY-60206 (Cooling / Conveying Fan FANS-31 VFD Speed Reference)"),
            ("Control Strategy", "Rapidly cools warm powder exiting cyclone below starch glass transition point (45°C) to prevent caking and agglomeration during transfer to storage silos.")
        ]),
        ("PIC-61302 : SILO VENTING & LEV DUST EXTRACTION VACUUM", ACCENT_AMBER, [
            ("Process Service", "Packaging Silo & Bagging Station Dust Capture (P&ID: 325-01-613-PD-01 to PD-05)"),
            ("Process Variable (PV)", "PT-61302 (Vacuum Duct Pressure) & DPT-61301/02 (Vent Filter Differential Pressure)"),
            ("Set Point (SP)", "-15.0 mbar Negative Pressure at LEV Pick-up Hoods"),
            ("Control Variable (CV)", "BFY-61302 (LEV Blower Fan BF-61302 VFD) & BFY-61303 (Fugitive Dust Fan VFD)"),
            ("Control Strategy", "Maintains negative capture velocity across packing heads. Automatically triggers pulse-jet reverse air filter cleaning (SV-61310) when DPT exceeds 12.0 mbar.")
        ]),
        ("PIC-81001 : BOILER STEAM DEMAND REMOTE SETPOINT SYNC", ACCENT_PURPLE, [
            ("Process Service", "Utility Boiler Steam Synchronization Skid (P&ID: 325-01-814-PD-01 / RIO-200)"),
            ("Process Variable (PV)", "Instantaneous Plant Steam Flow FT-60202 + FT-60203 & Steam Header Pressure PT-60201"),
            ("Set Point (SP)", "Dynamic Steam Demand Tracking (Feedforward compensation)"),
            ("Control Variable (CV)", "External Remote Setpoint (4–20 mA AO from RIO-200 to Boiler Modulation Skid)"),
            ("Control Strategy", "Feedforward boiler firing setpoint sent before large steam valves (PCV-60201 / TCV-40222) ramp up, eliminating boiler pressure sags during plant startup and batch switching.")
        ])
    ]

    for idx, (title, color, items) in enumerate(c6_data):
        pos_l, pos_t = positions_2x2[idx]
        box = s6.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, pos_l, pos_t, grid_w, grid_h)
        box.fill.solid()
        box.fill.fore_color.rgb = PANEL_BG
        box.line.color.rgb = color
        box.line.width = Pt(1.5)

        btf = box.text_frame
        btf.word_wrap = True
        btf.margin_left = Inches(0.12)
        btf.margin_right = Inches(0.12)
        btf.margin_top = Inches(0.08)

        bp = btf.paragraphs[0]
        bp.text = title
        bp.font.size = Pt(10)
        bp.font.bold = True
        bp.font.color.rgb = color

        for k, v in items:
            p_item = btf.add_paragraph()
            p_item.text = f"• {k}: {v}"
            p_item.font.size = Pt(7.8)
            p_item.font.color.rgb = LIGHT_GRAY
            p_item.space_before = Pt(1.5)

    # =========================================================================
    # SLIDE 7: Master PID Loop Engineering Schedule & Matrix Table
    # =========================================================================
    s7 = prs.slides.add_slide(blank_layout)
    bg7 = s7.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg7.fill.solid()
    bg7.fill.fore_color.rgb = DARK_BG
    bg7.line.fill.background()

    add_slide_header(s7, "MASTER PID LOOP ENGINEERING SCHEDULE & I/O ALLOCATION TABLE")

    table_shape = s7.shapes.add_table(17, 9, Inches(0.4), Inches(1.2), Inches(12.533), Inches(5.8))
    table = table_shape.table

    col_widths = [Inches(1.1), Inches(2.3), Inches(1.3), Inches(1.5), Inches(1.1), Inches(1.9), Inches(1.3), Inches(0.8), Inches(1.2)]
    for ci, w in enumerate(col_widths):
        table.columns[ci].width = w

    headers = ["Loop Tag", "Process Service Description", "P&ID Drawing", "PV Sensor Tag(s)", "Nominal SP", "Final Control Element", "I/O Interface", "Action", "Fail State"]
    for ci, h_text in enumerate(headers):
        cell = table.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = CARD_BG_LIGHT
        p = cell.text_frame.paragraphs[0]
        p.text = h_text
        p.font.size = Pt(8.5)
        p.font.bold = True
        p.font.color.rgb = ACCENT_BLUE
        p.alignment = PP_ALIGN.CENTER

    table_data = [
        ("TIC-40222", "Jet Cooker Direct Steam Cooking Temp", "325-01-402-PD-02", "TT-40221 / TT-40222", "120–140 °C", "TCV-40222 (Samson 3241)", "4-20mA (C3S13,00)", "Rev", "Fail Closed"),
        ("FIC-40203", "Jet Cooker Slurry Mass Flow Control", "325-01-402-PD-02", "FT-40221 / FT-40222", "4.5–7.5 t/h", "PUCF-11 Slurry Pump VFD", "EtherNet/IP Bus", "Rev", "Stop on Fail"),
        ("pHTC-40201", "pH Adjustment Tank 1 Conditioning", "325-01-402-PD-01", "pHT-40201 (0–14 pH)", "5.8–6.2 pH", "PD-40204 Dosing Pump", "Discrete/Speed", "Split", "Stop on Fail"),
        ("pHTC-40202", "pH Adjustment Tank 2 Conditioning", "325-01-402-PD-01", "pHT-40202 (0–14 pH)", "5.8–6.2 pH", "PD-40205 Dosing Pump", "Discrete/Speed", "Split", "Stop on Fail"),
        ("PIC-60201", "Spray Dryer Main Steam Header Press.", "325-01-602-PD-02", "PT-60201 (0–25 bar)", "16.0 bar(g)", "PCV-60201 (Samson NPS 4\")", "4-20mA (C3S13,03)", "Rev", "Fail Closed"),
        ("TIC-60214", "Air Pre-Heater Steam Exchanger Temp", "325-01-602-PD-03", "TT-60214 (0–150°C)", "75–85 °C", "TCV-60214 (Samson NPS 2\")", "4-20mA (C3S13,02)", "Rev", "Fail Closed"),
        ("TIC-60210", "Dryer Inlet Air Heating Temperature", "325-01-602-PD-06", "TT-60210 / TT-60211", "185–205 °C", "UY-FH-60210B (BMS Demand)", "4-20mA (C3S13,04)", "Rev", "0% Demand"),
        ("FIC-60210", "Main Supply Air Flow Rate Control", "325-01-602-PD-03", "FT-60210 / FT-60211", "32,000 Nm³/h", "FANS-21 Supply Fan VFD", "EtherNet/IP Bus", "Rev", "Coast / Stop"),
        ("FIC-60211", "Burner Combustion Air Flow / Ratio", "325-01-602-PD-03", "PT-60216 / TT-60216", "Ratio to Gas", "FANC-21 Comb. Fan VFD", "EtherNet/IP Bus", "Rev", "Coast / Purge"),
        ("PIC-60220", "Dryer Chamber Draft Pressure (Safety)", "325-01-602-PD-04", "PT-60220 (-10..+10 mbar)", "-1.20 mbar", "FANE-21 Exhaust Fan VFD", "EtherNet/IP Bus", "Direct", "Safe Ramp"),
        ("TIC-60226", "Exhaust Temp / Moisture Cascade Loop", "325-01-602-PD-04", "TT-60226 / TT-60240", "88–92 °C", "PUPD-11 Slurry Pump VFD", "EtherNet/IP Bus", "Direct", "Auto Cutback"),
        ("TIC-60202", "Camera Purge Air Chilled Water Temp", "325-01-602-PD-07", "TT-60202 (0–60°C)", "22–25 °C", "TCV-60202 (3-Way Valve)", "4-20mA (C3S13,01)", "Direct", "Fail Port A"),
        ("TIC-60215", "Glycol Heat Recovery Circulation", "325-01-602-PD-03", "TT-60215A / TT-60215B", "55–65 °C", "PUCF-21 Glycol Pump VFD", "EtherNet/IP Bus", "Rev", "Modulate"),
        ("FIC-60201", "Powder Cooling & Conveying Air Flow", "325-01-602-PD-08", "FT-60201 / PT-60203", "18–22 m/s", "FANS-31 Conveying Fan VFD", "EtherNet/IP Bus", "Rev", "Coast / Stop"),
        ("PIC-61302", "Packing Silo & LEV Extraction Vacuum", "325-01-613-PD-03", "PT-61302 (-50..0 mbar)", "-15 mbar", "BF-61302 LEV Blower VFD", "EtherNet/IP Bus", "Direct", "Modulate"),
        ("PIC-81001", "Boiler Steam Demand Remote Sync", "325-01-814-PD-01", "Total Steam (FT-60202)", "Feedforward", "Remote Setpoint to Boiler", "4-20mA (RIO-200)", "Direct", "Demand Safe")
    ]

    for ri, row_vals in enumerate(table_data, start=1):
        bg_col = CARD_BG if (ri % 2 == 1) else PANEL_BG
        for ci, val in enumerate(row_vals):
            cell = table.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg_col
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(7.5)
            if ci == 0:
                p.font.bold = True
                p.font.color.rgb = ACCENT_AMBER if "TIC" in val else (ACCENT_ROSE if "PIC" in val else ACCENT_TEAL)
            elif ci in (7, 8):
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER
            else:
                p.font.color.rgb = LIGHT_GRAY

    # =========================================================================
    # SLIDE 8: Rockwell ControlLogix PIDE Implementation & HMI Integration
    # =========================================================================
    s8 = prs.slides.add_slide(blank_layout)
    bg8 = s8.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg8.fill.solid()
    bg8.fill.fore_color.rgb = DARK_BG
    bg8.line.fill.background()

    add_slide_header(s8, "ROCKWELL CONTROLLOGIX PIDE ARCHITECTURE & FACTORYTALK VIEW SE INTEGRATION")

    c8_w = Inches(3.9)
    c8_gap = Inches(0.3)
    c8_h = Inches(5.8)

    c8_data = [
        ("CONTROLLOGIX PIDE INSTRUCTION", ACCENT_BLUE, [
            ("Algorithm Type", "Enhanced PID (PIDE) in Function Block Diagram (FBD) / Periodic Tasks"),
            ("Task Execution Rates", "• Pressure & Draft Loops (PIC-60220, PIC-60201): 50 ms Periodic Task\n• Flow & Speed Loops (FIC-40203, FIC-60210): 100 ms Periodic Task\n• Temperature & pH Loops (TIC-40222, TIC-60210): 250 ms Periodic Task"),
            ("Operating Modes", "• Hand (Operator Manual CV override via HMI)\n• Auto (Automatic regulation to Local Setpoint)\n• Cascade (Remote Setpoint from supervisory cascade)"),
            ("Anti-Windup & Clamping", "Built-in dynamic CV upper/lower limits (CVHighLimit = 100%, CVLowLimit = 0%) and programmable slew rate clamping (%/sec) to prevent actuator hammering."),
            ("Derivative Filtering", "Low-pass derivative filter applied to PV (not error) to eliminate electrical sensor noise spikes from kicking control output.")
        ]),
        ("ADVANCED PROCESS STRATEGIES", ACCENT_AMBER, [
            ("Chamber Draft Feedforward", "Combines Exhaust Fan speed with Main Air Supply Fan speed feedforward (FANE-21 tracks FANS-21) to cancel transient draft surges during air startup."),
            ("Burner Firing Heat Decoupling", "Ratio decoupling between combustion airflow and gas demand (UY-FH-60210B) ensures stoichiometric flame profile during load ramps."),
            ("Slurry Moisture Cascade", "Outer Temperature Loop (TIC-60226) outputs Remote SP to Inner Slurry Mass Flow Loop (FIC-40205) for high-fidelity moisture containment."),
            ("Dual Duplex Sensor Validation", "High-temp RTDs/TCs (TT-40221/22 and TT-60210/11) utilize 2-out-of-2 deviation monitoring. Sensor drift (>3°C) generates preventive maintenance alarm.")
        ]),
        ("FACTORYTALK VIEW SE HMI STANDARD", ACCENT_GREEN, [
            ("PlantPAx Faceplates", "Standardized ISA-101 high-performance HMI faceplates deployed across 3 Operator Workstations (OWS) and 1 EWS."),
            ("Operator Interaction", "• One-click Auto / Manual / Cascade toggling\n• Visual SP / PV / CV real-time bar graphs\n• Dedicated Tuning tab with P, I, D gains, Deadband & Filter coefficients\n• Historical trending with 1-second sample rate"),
            ("Alarm Management", "EEMUA 191 / ISA-18.2 compliant alarming:\n• PV High-High (Emergency Interlock Trip)\n• PV High / Low (Process Warning)\n• Deviation Alarm (Control Loop Degradation)\n• Transmitter Out-of-Range (Bad PV Quality)")
        ])
    ]

    left_c8 = Inches(0.5)
    for title, color, items in c8_data:
        box = s8.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left_c8, Inches(1.2), c8_w, c8_h)
        box.fill.solid()
        box.fill.fore_color.rgb = PANEL_BG
        box.line.color.rgb = color
        box.line.width = Pt(1.5)

        btf = box.text_frame
        btf.word_wrap = True
        btf.margin_left = Inches(0.15)
        btf.margin_right = Inches(0.15)
        btf.margin_top = Inches(0.12)

        bp = btf.paragraphs[0]
        bp.text = title
        bp.font.size = Pt(11)
        bp.font.bold = True
        bp.font.color.rgb = color

        for k, v in items:
            p_item = btf.add_paragraph()
            p_item.text = f"{k}:"
            p_item.font.size = Pt(8.5)
            p_item.font.bold = True
            p_item.font.color.rgb = WHITE
            p_item.space_before = Pt(3)

            p_desc = btf.add_paragraph()
            p_desc.text = v
            p_desc.font.size = Pt(8)
            p_desc.font.color.rgb = LIGHT_GRAY

        left_c8 += c8_w + card_gap

    # Save presentations
    prs.save(PPTX_PATH_1)
    shutil.copyfile(PPTX_PATH_1, PPTX_PATH_2)
    print(f"Successfully generated PowerPoint slides:\n  - {PPTX_PATH_1}\n  - {PPTX_PATH_2}")


def render_pid_architecture_png():
    """Render high-resolution 1920x1080 visual diagram of the PID loop architecture"""
    W, H = 1920, 1080
    im = Image.new("RGB", (W, H), (15, 23, 42)) # Slate 900
    d = ImageDraw.Draw(im)

    def get_font(size, bold=False):
        try:
            if bold:
                return ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial Bold.ttf", size)
            return ImageFont.truetype("/System/Library/Fonts/Supplemental/Arial.ttf", size)
        except Exception:
            try:
                return ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", size)
            except Exception:
                return ImageFont.load_default()

    f_title = get_font(28, bold=True)
    f_sub = get_font(14, bold=True)
    f_sec = get_font(16, bold=True)
    f_card_h = get_font(13, bold=True)
    f_body = get_font(11)
    f_body_bold = get_font(11, bold=True)
    f_meta = get_font(11)

    # Top Header Ribbon
    d.rectangle([(40, 30), (1880, 110)], fill=(30, 41, 59), outline=(71, 85, 105), width=2)
    d.text((60, 40), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT | PROJECT SPRINT (18,000 TPA SPRAY DRYER)", fill=(56, 189, 248), font=f_sub)
    d.text((60, 65), "MASTER PROCESS PID LOOP CONTROL ARCHITECTURE & I/O TOPOLOGY", fill=(255, 255, 255), font=f_title)
    d.text((1500, 70), "16 PID CONTROL LOOPS", fill=(34, 197, 94), font=f_sub)

    # 4 Main Zone Blocks
    zone_defs = [
        ("ZONE 1 : SLURRY PREP & JET COOKER", 40, 130, 440, 890, (20, 184, 166), [
            ("TIC-40222 [CRITICAL]", "Jet Cooker Steam Temp Control", "PV: TT-40221 / TT-40222 (Duplex RTD)\nSP: 120.0–140.0 °C Starch Gelatinization\nCV: TCV-40222 (Samson 3241 Globe)\nOut: 4-20 mA AO (Chassis C3 Slot 13,00)\nFail State: Fail Closed (FC)"),
            ("FIC-40203", "Jet Cooker Slurry Mass Flow", "PV: FT-40221 / FT-40222 (Coriolis)\nSP: 4.5–7.5 t/h Controlled Feed\nCV: PUCF-11 Feed Pump VFD Speed\nInterface: EtherNet/IP Bus (C7S00)\nAction: Reverse / Anti-Cavitation"),
            ("pHTC-40201 / 02", "Slurry pH Adjustment Tanks 1 & 2", "PV: pHT-40201 / pHT-40202 (InPro 3250i)\nSP: 5.8–6.2 pH Reaction Buffer\nCV: Dosing Pumps PD-40204 / PD-40205\nAction: Split-Range Dual Neutralization\nSensors: Auto Water Purge Wash"),
            ("LIC-40203", "Jet Cooker Infeed Buffer Level", "PV: LT-40203 Level Transmitter\nSP: 65.0% Operational Buffer\nCV: Upstream Slurry Transfer Feed\nSafety: Low-level dry-run pump cutoff")
        ]),
        ("ZONE 2 : AIR HEATING & STEAM", 500, 130, 440, 890, (245, 158, 11), [
            ("TIC-60210 [MASTER]", "Dryer Drying Air Inlet Temperature", "PV: TT-60210 / TT-60211 (Hot Air Duct)\nSP: 185.0–205.0 °C Thermal Setpoint\nCV: UY-FH-60210B (Burner Demand)\nOut: 4-20 mA AO (Chassis C3 Slot 13,04)\nSkid: AHTR-21 Gas Air Heater (BMS)"),
            ("TIC-60214", "Air Pre-Heater Steam Exchanger", "PV: TT-60214 (Downstream HXCH-21)\nSP: 75.0–85.0 °C Air Pre-heat\nCV: TCV-60214 (Samson NPS 2\" Cv 47)\nOut: 4-20 mA AO (Chassis C3 Slot 13,02)\nStrategy: Steam Condensate Recapture"),
            ("PIC-60201", "Dryer Steam Header Pressure", "PV: PT-60201 (0–25 bar(g) Cerabar)\nSP: 16.0 bar(g) Stable Supply Header\nCV: PCV-60201 (Samson NPS 4\" Cv 190)\nOut: 4-20 mA AO (Chassis C3 Slot 13,03)\nAction: Fast Surge Mitigation (50ms)"),
            ("FIC-60210 / 11", "Main Supply Air & Combustion Flow", "PV: FT-60210/11 Mass Air & PT-60216\nSP: 32,000 Nm³/h & Fuel-Air Ratio\nCV: FANS-21 & FANC-21 VFD Drives\nInterface: EtherNet/IP DLR Bus Link\nInterlock: NFPA 86 Burner Purge")
        ]),
        ("ZONE 3 : DRYING CHAMBER & EXHAUST", 960, 130, 440, 890, (244, 63, 94), [
            ("PIC-60220 [SAFETY]", "Drying Chamber Draft Pressure", "PV: PT-60220 (-10.0..+10.0 mbar)\nSP: -1.20 mbar Continuous Negative Draft\nCV: FANE-21 Main Exhaust Fan VFD\nInterface: EtherNet/IP Bus (C7S00)\nCritical: Eliminates dust leakage & puff"),
            ("TIC-60226 / FIC", "Exhaust Temp / Moisture Cascade", "PV: TT-60226 Outlet & TT-60240 Baghouse\nSP: 88.0–92.0 °C (Powder Moisture ~6%)\nCV: PUPD-11 Slurry High-Press Pump VFD\nFeedback: 16 Atomizing Flowmeters\nAction: Dynamic Product Quality Control"),
            ("TIC-60202", "Optical Camera Air Chilled Water", "PV: TT-60202 Camera Purge Temp\nSP: 22.0–25.0 °C Lens Temperature\nCV: TCV-60202 (3-Way Modulating Valve)\nOut: 4-20 mA AO (Chassis C3 Slot 13,01)\nProtects: In-Chamber ACAM-21 Camera"),
            ("TIC-60215", "Glycol Run-Around Heat Recovery", "PV: TT-60215A/B Exhaust Recovery Temp\nSP: 55.0–65.0 °C Energy Recapture\nCV: PUCF-21 Glycol Pump VFD Speed\nInterface: EtherNet/IP Bus (C7S00)\nProtection: Exhaust acid dewpoint check")
        ]),
        ("ZONE 4 : CONVEYING, LEV & UTILITY", 1420, 130, 460, 890, (168, 85, 247), [
            ("FIC-60201", "Powder Cooling Air Conveying Flow", "PV: FT-60201 Air Flow & PT-60203 Pressure\nSP: 18.0–22.0 m/s Saltation Velocity\nCV: FANS-31 Conveying Fan VFD Speed\nFunction: Rapid powder chilling < 45°C\nInterface: EtherNet/IP Bus (C7S00)"),
            ("PIC-61302", "Silo & LEV Dust Extraction Vacuum", "PV: PT-61302 & DPT-61301/02 Differential\nSP: -15.0 mbar Extraction Vacuum\nCV: BF-61302 LEV Blower VFD Speed\nSafety: ATEX Flammable Dust Control\nPulse Jet: Auto filter reverse air clean"),
            ("LIC-61301", "Packing Storage Silo Level Balance", "PV: LT-61301 Radar Level Transmitter\nSP: 20–80% Working Buffer Capacity\nCV: Cyclone Rotary Valves RVAL-21/22\nFunction: Continuous feed to bagging"),
            ("PIC-81001", "Boiler Remote Steam Demand Sync", "PV: FT-60202 Total Steam Flow + Press.\nSP: Dynamic Feedforward Demand Track\nCV: Remote Setpoint to Boiler Modulation\nOut: 4-20 mA AO (RIO-200 Skid)\nMitigates: Boiler header pressure collapse")
        ])
    ]

    for z_title, z_x, z_y, z_w, z_h, z_color, loop_cards in zone_defs:
        d.rounded_rectangle([(z_x, z_y), (z_x + z_w, z_y + z_h)], radius=12, fill=(24, 34, 53), outline=z_color, width=2)
        d.text((z_x + 15, z_y + 12), z_title, fill=z_color, font=f_sec)

        c_y = z_y + 45
        c_h = 195
        c_w = z_w - 30
        c_x = z_x + 15
        for l_tag, l_name, l_body in loop_cards:
            d.rounded_rectangle([(c_x, c_y), (c_x + c_w, c_y + c_h)], radius=8, fill=(30, 41, 59), outline=(51, 65, 85), width=1)
            d.text((c_x + 12, c_y + 10), l_tag, fill=(255, 255, 255), font=f_card_h)
            d.text((c_x + 12, c_y + 30), l_name, fill=z_color, font=f_body_bold)

            line_y = c_y + 52
            for b_line in l_body.split("\n"):
                d.text((c_x + 12, line_y), b_line, fill=(203, 213, 225), font=f_body)
                line_y += 18

            c_y += c_h + 15

    d.rectangle([(40, 1030), (1880, 1060)], fill=(30, 41, 59), outline=(51, 65, 85), width=1)
    d.text((60, 1038), "Rockwell ControlLogix 5580 (1756-L83E Redundant) | PIDE Enhanced PID Instruction | Periodic Tasks (50ms / 100ms / 250ms) | FactoryTalk View SE Distributed", fill=(148, 163, 184), font=f_meta)
    d.text((1580, 1038), "AEC Industrial Engineering | Rev 3.6", fill=(56, 189, 248), font=f_meta)

    im.save(PNG_PATH_1, quality=95)
    shutil.copyfile(PNG_PATH_1, PNG_PATH_2)
    print(f"Successfully generated architecture overview image:\n  - {PNG_PATH_1}\n  - {PNG_PATH_2}")


if __name__ == "__main__":
    create_presentation()
    render_pid_architecture_png()
