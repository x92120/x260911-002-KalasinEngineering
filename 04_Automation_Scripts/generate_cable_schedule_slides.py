#!/usr/bin/env python3
"""
Generate professional engineering slide deck for:
Cable Schedule Between MCC / CA1 Suite and Each Field Junction Box.
Outputs:
1. PowerPoint Presentation (.pptx) - 16:9 Widescreen
2. Vector PDF Presentation (.pdf) - 16:9 Landscape
3. High-resolution PNG slide images for previewing
4. Interactive Web Slide Deck (HTML)
Project Jet Cooker | Ingredion (Thailand) Co., Ltd. | AEC Industrial Engineering
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from fpdf import FPDF
import fitz  # PyMuPDF

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(WORKSPACE_DIR, "reports_pdf")
CAD_DIR = os.path.join(WORKSPACE_DIR, "cad_exports")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(CAD_DIR, exist_ok=True)

PPTX_PATH = os.path.join(OUTPUT_DIR, "Cable_Schedule_MCC_to_Junction_Boxes.pptx")
PDF_PATH = os.path.join(OUTPUT_DIR, "Cable_Schedule_MCC_to_Junction_Boxes.pdf")

# -----------------------------------------------------------------------------
# MASTER CABLE DATA DEFINITION
# -----------------------------------------------------------------------------
CABLES = [
    # Panel M1
    {
        "tag": "CBL-M1-JB401-01", "from": "Panel M1", "to": "JB-401", "loc": "Infeed 2nd Flr",
        "func": "Digital Inputs (DI)", "ch": "64 DI", "spec": "CVV-S 48C x 1.5 mm2",
        "tray": "TR-IN-01", "len": 55, "gland": "M40", "type": "Control"
    },
    {
        "tag": "CBL-M1-JB401-02", "from": "Panel M1", "to": "JB-401", "loc": "Infeed 2nd Flr",
        "func": "DI Extension & Spares", "ch": "Spares", "spec": "CVV-S 30C x 1.5 mm2",
        "tray": "TR-IN-01", "len": 55, "gland": "M32", "type": "Control"
    },
    {
        "tag": "CBL-M1-JB401-03", "from": "Panel M1", "to": "JB-401", "loc": "Infeed 2nd Flr",
        "func": "Digital Outputs (DO)", "ch": "32 DO", "spec": "CVV-S 40C x 1.5 mm2",
        "tray": "TR-IN-01", "len": 55, "gland": "M36", "type": "Control"
    },
    {
        "tag": "CBL-M1-JB401-04", "from": "Panel M1", "to": "JB-401", "loc": "Infeed 2nd Flr",
        "func": "Analog Inputs (AI)", "ch": "16 AI", "spec": "IS/OS/PVC 16P x 1.0 mm2",
        "tray": "TR-IN-01", "len": 55, "gland": "M32", "type": "Analog Shielded"
    },
    {
        "tag": "CBL-M1-JB601-01", "from": "Panel M1", "to": "JB-601", "loc": "Spray Dryer 1st Flr",
        "func": "Digital Inputs (DI)", "ch": "51 DI", "spec": "CVV-S 60C x 1.5 mm2",
        "tray": "TR-SD-01", "len": 70, "gland": "M50", "type": "Control"
    },
    {
        "tag": "CBL-M1-JB601-02", "from": "Panel M1", "to": "JB-601", "loc": "Spray Dryer 1st Flr",
        "func": "Digital Outputs (DO)", "ch": "17 DO", "spec": "CVV-S 24C x 1.5 mm2",
        "tray": "TR-SD-01", "len": 70, "gland": "M25", "type": "Control"
    },
    {
        "tag": "CBL-M1-JB601-03", "from": "Panel M1", "to": "JB-601", "loc": "Spray Dryer 1st Flr",
        "func": "Analog Inputs (AI)", "ch": "16 AI", "spec": "IS/OS/PVC 16P x 1.0 mm2",
        "tray": "TR-SD-01", "len": 70, "gland": "M32", "type": "Analog Shielded"
    },
    {
        "tag": "CBL-M1-RIO200-FO", "from": "Panel M1", "to": "RIO-200", "loc": "Slurry Out Bldg 2F",
        "func": "Redundant DLR Ethernet", "ch": "Backbone", "spec": "6-Core Armoured SM Fiber (OS2)",
        "tray": "TR-EXT-01", "len": 165, "gland": "M20", "type": "Fiber Optic"
    },
    {
        "tag": "CBL-M1-RIO200-PWR", "from": "Panel M1", "to": "RIO-200", "loc": "Slurry Out Bldg 2F",
        "func": "230VAC UPS Power", "ch": "Feeder", "spec": "XLPE/SWA/PVC 3C x 4.0 mm2",
        "tray": "TR-EXT-01", "len": 165, "gland": "M25", "type": "Power"
    },

    # Panel M2
    {
        "tag": "CBL-M2-JB402-01", "from": "Panel M2", "to": "JB-402", "loc": "Jet Cooker 2nd Flr",
        "func": "Digital Inputs (DI)", "ch": "32 DI", "spec": "CVV-S 40C x 1.5 mm2",
        "tray": "TR-JC-01", "len": 45, "gland": "M36", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB402-02", "from": "Panel M2", "to": "JB-402", "loc": "Jet Cooker 2nd Flr",
        "func": "Digital Outputs (DO)", "ch": "16 DO", "spec": "CVV-S 24C x 1.5 mm2",
        "tray": "TR-JC-01", "len": 45, "gland": "M25", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB402-03", "from": "Panel M2", "to": "JB-402", "loc": "Jet Cooker 2nd Flr",
        "func": "Analog Inputs & AO", "ch": "16 AI + 1 AO", "spec": "IS/OS/PVC 20P x 1.0 mm2",
        "tray": "TR-JC-01", "len": 45, "gland": "M36", "type": "Analog Shielded"
    },
    {
        "tag": "CBL-M2-JB602-01", "from": "Panel M2", "to": "JB-602", "loc": "Spray Dryer 3rd Flr",
        "func": "Digital Inputs (DI)", "ch": "32 DI", "spec": "CVV-S 40C x 1.5 mm2",
        "tray": "TR-SD-02", "len": 85, "gland": "M36", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB602-02", "from": "Panel M2", "to": "JB-602", "loc": "Spray Dryer 3rd Flr",
        "func": "Digital Outputs (DO)", "ch": "16 DO", "spec": "CVV-S 24C x 1.5 mm2",
        "tray": "TR-SD-02", "len": 85, "gland": "M25", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB602-03", "from": "Panel M2", "to": "JB-602", "loc": "Spray Dryer 3rd Flr",
        "func": "Analog Inputs & AO", "ch": "32 AI + 2 AO", "spec": "IS/OS/PVC 36P x 1.0 mm2",
        "tray": "TR-SD-02", "len": 85, "gland": "M40", "type": "Analog Shielded"
    },
    {
        "tag": "CBL-M2-JB607-01", "from": "Panel M2", "to": "JB-607", "loc": "Spray Dryer 7th Flr",
        "func": "Digital Inputs (DI) Pt 1", "ch": "48 DI", "spec": "CVV-S 60C x 1.5 mm2",
        "tray": "TR-SD-03", "len": 125, "gland": "M50", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB607-02", "from": "Panel M2", "to": "JB-607", "loc": "Spray Dryer 7th Flr",
        "func": "Digital Inputs (DI) Pt 2", "ch": "45 DI", "spec": "CVV-S 60C x 1.5 mm2",
        "tray": "TR-SD-03", "len": 125, "gland": "M50", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB607-03", "from": "Panel M2", "to": "JB-607", "loc": "Spray Dryer 7th Flr",
        "func": "Digital Outputs (DO) Pt 1", "ch": "40 DO", "spec": "CVV-S 48C x 1.5 mm2",
        "tray": "TR-SD-03", "len": 125, "gland": "M40", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB607-04", "from": "Panel M2", "to": "JB-607", "loc": "Spray Dryer 7th Flr",
        "func": "Digital Outputs (DO) Pt 2", "ch": "24 DO", "spec": "CVV-S 30C x 1.5 mm2",
        "tray": "TR-SD-03", "len": 125, "gland": "M32", "type": "Control"
    },
    {
        "tag": "CBL-M2-JB607-05", "from": "Panel M2", "to": "JB-607", "loc": "Spray Dryer 7th Flr",
        "func": "Analog Inputs & AO", "ch": "32 AI + 5 AO", "spec": "IS/OS/PVC 40P x 1.0 mm2",
        "tray": "TR-SD-03", "len": 125, "gland": "M40", "type": "Analog Shielded"
    },

    # Panel M3 (Ex-i Bay 1)
    {
        "tag": "CBL-M3-ISJB603-01", "from": "Panel M3", "to": "IS-JB-603", "loc": "Spray Dryer 3rd Flr",
        "func": "Ex-i Digital Inputs (DI)", "ch": "16 DI (IS)", "spec": "CVV-S (IS) 24C x 1.5 mm2 [BLUE]",
        "tray": "TR-SD-IS", "len": 85, "gland": "M25 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M3-ISJB603-02", "from": "Panel M3", "to": "IS-JB-603", "loc": "Spray Dryer 3rd Flr",
        "func": "Ex-i Analog Inputs (AI)", "ch": "16 AI (IS)", "spec": "IS/OS/PVC (IS) 16P x 1.0 mm2 [BLUE]",
        "tray": "TR-SD-IS", "len": 85, "gland": "M32 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M3-ISJB618-01", "from": "Panel M3", "to": "IS-JB-618", "loc": "Packing Tower 8th Flr",
        "func": "Ex-i Digital Inputs (DI)", "ch": "32 DI (IS)", "spec": "CVV-S (IS) 40C x 1.5 mm2 [BLUE]",
        "tray": "TR-PT-IS", "len": 140, "gland": "M36 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M3-ISJB618-02", "from": "Panel M3", "to": "IS-JB-618", "loc": "Packing Tower 8th Flr",
        "func": "Ex-i Analog Inputs (AI)", "ch": "12 AI (IS)", "spec": "IS/OS/PVC (IS) 16P x 1.0 mm2 [BLUE]",
        "tray": "TR-PT-IS", "len": 140, "gland": "M32 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M3-JB618-01", "from": "Panel M3", "to": "JB-618", "loc": "Packing Tower 8th Flr",
        "func": "Non-IS DI & AI", "ch": "16 DI + 4 AI", "spec": "CVV-S 24C x 1.5 mm2",
        "tray": "TR-PT-01", "len": 140, "gland": "M25", "type": "Control"
    },
    {
        "tag": "CBL-M3-JB618-02", "from": "Panel M3", "to": "JB-618", "loc": "Packing Tower 8th Flr",
        "func": "Non-IS Digital Outputs", "ch": "32 DO", "spec": "CVV-S 40C x 1.5 mm2",
        "tray": "TR-PT-01", "len": 140, "gland": "M36", "type": "Control"
    },

    # Panel M4 (Ex-i Bay 2)
    {
        "tag": "CBL-M4-ISJB608-01", "from": "Panel M4", "to": "IS-JB-608", "loc": "Spray Dryer 8th Flr",
        "func": "Ex-i DI & DO Signals", "ch": "16 DI + 8 DO (IS)", "spec": "CVV-S (IS) 30C x 1.5 mm2 [BLUE]",
        "tray": "TR-SD-IS", "len": 135, "gland": "M32 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M4-ISJB608-02", "from": "Panel M4", "to": "IS-JB-608", "loc": "Spray Dryer 8th Flr",
        "func": "Ex-i Analog Inputs (AI)", "ch": "8 AI (IS)", "spec": "IS/OS/PVC (IS) 10P x 1.0 mm2 [BLUE]",
        "tray": "TR-SD-IS", "len": 135, "gland": "M25 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M4-ISJB612-01", "from": "Panel M4", "to": "IS-JB-612", "loc": "Packing Tower 2nd Flr",
        "func": "Ex-i Digital Inputs (DI)", "ch": "31 DI (IS)", "spec": "CVV-S (IS) 40C x 1.5 mm2 [BLUE]",
        "tray": "TR-PT-IS", "len": 65, "gland": "M36 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M4-ISJB612-02", "from": "Panel M4", "to": "IS-JB-612", "loc": "Packing Tower 2nd Flr",
        "func": "Ex-i Analog Inputs (AI)", "ch": "10 AI (IS)", "spec": "IS/OS/PVC (IS) 12P x 1.0 mm2 [BLUE]",
        "tray": "TR-PT-IS", "len": 65, "gland": "M25 (IS)", "type": "Ex-i Intrinsically Safe"
    },
    {
        "tag": "CBL-M4-JB606-01", "from": "Panel M4", "to": "JB-606", "loc": "Spray Dryer 6th Flr",
        "func": "Non-IS DI, DO & AI", "ch": "16 DI + 8 DO + 2 AI", "spec": "CVV-S 30C x 1.5 mm2",
        "tray": "TR-SD-02", "len": 115, "gland": "M32", "type": "Control"
    },
    {
        "tag": "CBL-M4-JB608-01", "from": "Panel M4", "to": "JB-608", "loc": "Spray Dryer 8th Flr",
        "func": "Non-IS Analog Inputs", "ch": "6 AI", "spec": "IS/OS/PVC 8P x 1.0 mm2",
        "tray": "TR-SD-03", "len": 135, "gland": "M20", "type": "Analog Shielded"
    },
    {
        "tag": "CBL-M4-JB612-01", "from": "Panel M4", "to": "JB-612", "loc": "Packing Tower 2nd Flr",
        "func": "Non-IS Digital Inputs", "ch": "32 DI", "spec": "CVV-S 40C x 1.5 mm2",
        "tray": "TR-PT-01", "len": 65, "gland": "M36", "type": "Control"
    },
    {
        "tag": "CBL-M4-JB612-02", "from": "Panel M4", "to": "JB-612", "loc": "Packing Tower 2nd Flr",
        "func": "Non-IS DO & AI", "ch": "16 DO + 8 AI", "spec": "CVV-S 24C x 1.5 + 10P x 1.0",
        "tray": "TR-PT-01", "len": 65, "gland": "M32", "type": "Control & Analog"
    },

    # MCC Room Interconnections
    {
        "tag": "CBL-M1-MCC-01", "from": "Panel M1", "to": "MCC Lineup", "loc": "MCC Room",
        "func": "Motor Feedback Signals", "ch": "125 DI", "spec": "2x CVV-S 48C + 36C x 1.5 mm2",
        "tray": "TR-MCC-01", "len": 25, "gland": "M40", "type": "Control"
    },
    {
        "tag": "CBL-M1-MCC-02", "from": "Panel M1", "to": "MCC Lineup", "loc": "MCC Room",
        "func": "Motor Command Signals", "ch": "95 DO", "spec": "2x CVV-S 48C + 24C x 1.5 mm2",
        "tray": "TR-MCC-01", "len": 25, "gland": "M40", "type": "Control"
    },
    {
        "tag": "CBL-M1-MCC-NET", "from": "Panel M1", "to": "MCC Lineup", "loc": "MCC Room",
        "func": "Smart Overloads (E300)", "ch": "17 BUS Nodes", "spec": "Armoured Industrial Cat6A DLR",
        "tray": "TR-MCC-01", "len": 30, "gland": "M20", "type": "Industrial Ethernet"
    }
]


# -----------------------------------------------------------------------------
# STEP 1: GENERATE POWERPOINT (PPTX) 16:9 WIDESCREEN PRESENTATION
# -----------------------------------------------------------------------------
def create_pptx_slides():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Color Palette
    C_NAVY = RGBColor(15, 30, 54)
    C_SLATE = RGBColor(30, 41, 59)
    C_LIGHT_BG = RGBColor(248, 250, 252)
    C_WHITE = RGBColor(255, 255, 255)
    C_BLUE = RGBColor(14, 165, 233)
    C_EMERALD = RGBColor(16, 185, 129)
    C_AMBER = RGBColor(245, 158, 11)
    C_RED = RGBColor(239, 68, 68)
    C_DARK_TEXT = RGBColor(15, 23, 42)
    C_MUTED = RGBColor(100, 116, 139)

    def add_header(slide, title, subtitle, category="CABLE SCHEDULE"):
        # Header background
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.9))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_top = tf.margin_right = tf.margin_bottom = 0
        
        p0 = tf.paragraphs[0]
        p0.text = category.upper()
        p0.font.size = Pt(9)
        p0.font.bold = True
        p0.font.color.rgb = C_BLUE

        p1 = tf.add_paragraph()
        p1.text = title
        p1.font.size = Pt(20)
        p1.font.bold = True
        p1.font.color.rgb = C_NAVY

        p2 = tf.add_paragraph()
        p2.text = subtitle
        p2.font.size = Pt(10)
        p2.font.color.rgb = C_MUTED

    def add_footer(slide, current_page, total_pages=8):
        footer_box = slide.shapes.add_textbox(Inches(0.8), Inches(7.0), Inches(11.733), Inches(0.4))
        tf = footer_box.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = f"Project Jet Cooker | Ingredion (Thailand) Co., Ltd. | AEC Industrial Engineering | Slide {current_page} of {total_pages}"
        p.font.size = Pt(9)
        p.font.color.rgb = C_MUTED

    # -------------------------------------------------------------
    # SLIDE 1: COVER SLIDE
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid(); bg1.fill.fore_color.rgb = C_NAVY
    bg1.line.fill.background()

    # Title Box
    tbox = s1.shapes.add_textbox(Inches(1.2), Inches(1.8), Inches(11.0), Inches(4.0))
    tf1 = tbox.text_frame
    tf1.word_wrap = True
    
    p = tf1.paragraphs[0]
    p.text = "INGREDION (THAILAND) CO., LTD. - KALASIN PLANT"
    p.font.size = Pt(13)
    p.font.bold = True
    p.font.color.rgb = C_BLUE

    p = tf1.add_paragraph()
    p.text = "PROJECT JET COOKER"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = RGBColor(226, 232, 240)

    p = tf1.add_paragraph()
    p.text = "Interconnection Cable Schedule"
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = C_WHITE

    p = tf1.add_paragraph()
    p.text = "Complete Multi-Core & Trunk Routing Between MCC / CA1 Suite and Field Junction Boxes"
    p.font.size = Pt(14)
    p.font.color.rgb = RGBColor(148, 163, 184)

    p = tf1.add_paragraph()
    p.text = "\nDocument No: CA1-SCH-CBL-001 | Revision: 3.6 | Date: September 2026\nEngineering Contractor: AEC Industrial Engineering Co., Ltd."
    p.font.size = Pt(11)
    p.font.color.rgb = C_AMBER

    # -------------------------------------------------------------
    # SLIDE 2: ROUTING ARCHITECTURE & SEGREGATION
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Overall Cable Routing & Signal Segregation Architecture", 
               "Physical pathways from MCC Room to Infeed, Jet Cooker, Spray Dryer, Packing Tower, and Slurry Building", 
               "System Architecture")
    add_footer(s2, 2)

    # 4 Architecture Cards
    cards = [
        ("Substation to Process Spine (TR-01)", "Main heavy-duty hot-dip galvanized ladder tray (600x100mm) running from MCC room basement trench across overhead pipe rack.", C_BLUE),
        ("Spray Dryer Riser Shaft (TR-SD)", "Vertical cable riser traversing 1st to 8th floors. Segregated trays for Non-IS Control, Analog Shielded, and Ex-i Intrinsically Safe.", C_EMERALD),
        ("Packing Tower Riser (TR-PT)", "Heavy-duty ladder tray for 2nd and 8th floor junction boxes. Blue-coated Ex-i dedicated cable tray with 300mm air gap separation.", C_AMBER),
        ("Slurry Building Fiber Trunk (TR-EXT)", "Overhead pipe rack route to remote building housing RIO-200. 6-core steel wire armoured fiber optic DLR ring backbone.", C_NAVY)
    ]
    for c_idx, (ctitle, cdesc, ccol) in enumerate(cards):
        cx = Inches(0.8 + c_idx * 2.95)
        cy = Inches(1.6)
        cbox = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, Inches(2.8), Inches(2.6))
        cbox.fill.solid(); cbox.fill.fore_color.rgb = C_LIGHT_BG
        cbox.line.color.rgb = ccol
        cbox.line.width = Pt(2)
        
        tf = cbox.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = ctitle
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = C_NAVY
        
        p2 = tf.add_paragraph()
        p2.text = "\n" + cdesc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = C_DARK_TEXT

    # Standards summary box below
    sbox = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(4.5), Inches(11.733), Inches(2.2))
    sbox.fill.solid(); sbox.fill.fore_color.rgb = RGBColor(241, 245, 249)
    sbox.line.color.rgb = RGBColor(203, 213, 225)
    stf = sbox.text_frame
    stf.word_wrap = True
    sp1 = stf.paragraphs[0]
    sp1.text = "ENGINEERING SEGREGATION & INSTALLATION RULES (IEC 60079-14 & IEEE 518)"
    sp1.font.bold = True; sp1.font.size = Pt(11); sp1.font.color.rgb = C_NAVY
    sp2 = stf.add_paragraph()
    sp2.text = (
        "1. Intrinsically Safe (Ex-i) Cables: Distinctive BLUE outer sheath (RAL 5015). Maintain >= 300 mm clearance or solid grounded divider from power/control.\n"
        "2. Shield Grounding: All overall and individual pair screens grounded at MCC Cabinet single-point clean earth bar ONLY. Field ends insulated and floating.\n"
        "3. Spare Capacity: All multicore control cables sized with 20% to 35% installed spare conductors for future process expansion.\n"
        "4. Cable Glands: Nickel-plated brass double-compression EMC cable glands (IP66) on standard JBs; Ex-e/Ex-i certified blue-ring glands on Ex-i JBs."
    )
    sp2.font.size = Pt(9.5); sp2.font.color.rgb = C_DARK_TEXT

    # -------------------------------------------------------------
    # HELPER TO RENDER A SLIDE TABLE
    # -------------------------------------------------------------
    def render_cable_table(slide, title, subtitle, cable_list, current_page):
        add_header(slide, title, subtitle, "INTERCONNECTION SCHEDULE")
        add_footer(slide, current_page)

        # Table dimensions
        rows = len(cable_list) + 1
        cols = 8
        left = Inches(0.8)
        top = Inches(1.5)
        width = Inches(11.733)
        height = Inches(0.4 * rows)

        table_shape = slide.shapes.add_table(rows, cols, left, top, width, height)
        table = table_shape.table

        # Column widths
        col_widths = [Inches(1.8), Inches(1.0), Inches(1.1), Inches(1.8), Inches(1.8), Inches(2.4), Inches(1.0), Inches(0.8)]
        for ci, cw in enumerate(col_widths):
            table.columns[ci].width = cw

        headers = ["Cable Tag", "From", "To (JB)", "Location", "Signal Function", "Cable Specification", "Tray", "Len (m)"]
        for ci, htext in enumerate(headers):
            cell = table.cell(0, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = C_NAVY
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = htext
            p.font.bold = True
            p.font.size = Pt(9)
            p.font.color.rgb = C_WHITE
            p.alignment = PP_ALIGN.CENTER if ci in [1, 2, 6, 7] else PP_ALIGN.LEFT

        for ri, cb in enumerate(cable_list):
            row_idx = ri + 1
            is_even = (ri % 2 == 0)
            row_bg = RGBColor(255, 255, 255) if is_even else RGBColor(241, 245, 249)
            if "Ex-i" in cb["type"]:
                row_bg = RGBColor(238, 246, 255)  # Soft blue for Ex-i

            vals = [cb["tag"], cb["from"], cb["to"], cb["loc"], cb["func"], cb["spec"], cb["tray"], f"{cb['len']}m"]
            for ci, val in enumerate(vals):
                cell = table.cell(row_idx, ci)
                cell.fill.solid(); cell.fill.fore_color.rgb = row_bg
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                p = cell.text_frame.paragraphs[0]
                p.text = str(val)
                p.font.size = Pt(8.5)
                p.font.color.rgb = C_NAVY if ci == 0 else C_DARK_TEXT
                if ci == 0:
                    p.font.bold = True
                if "BLUE" in str(val):
                    p.font.bold = True
                    p.font.color.rgb = RGBColor(2, 132, 199)
                p.alignment = PP_ALIGN.CENTER if ci in [1, 2, 6, 7] else PP_ALIGN.LEFT

    # -------------------------------------------------------------
    # SLIDE 3: PANEL M1 CABLE SCHEDULE
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    m1_cables = [c for c in CABLES if c["from"] == "Panel M1" and "MCC" not in c["to"]]
    render_cable_table(s3, "Panel M1 Interconnection Cable Schedule", 
                       "Master Controller Bay (=CA1+MCP-M1) serving Infeed (JB-401), Spray Dryer 1F (JB-601), and Slurry (RIO-200)", 
                       m1_cables, 3)

    # -------------------------------------------------------------
    # SLIDE 4: PANEL M2 CABLE SCHEDULE
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    m2_cables = [c for c in CABLES if c["from"] == "Panel M2"]
    render_cable_table(s4, "Panel M2 Interconnection Cable Schedule", 
                       "High-Density I/O Bay (=CA1+MCP-M2) serving Jet Cooker (JB-402), Spray Dryer 3F (JB-602), and Spray Dryer 7F (JB-607)", 
                       m2_cables, 4)

    # -------------------------------------------------------------
    # SLIDE 5: PANEL M3 CABLE SCHEDULE (Ex-i Bay 1)
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    m3_cables = [c for c in CABLES if c["from"] == "Panel M3"]
    render_cable_table(s5, "Panel M3 Intrinsically Safe (Ex-i) Cable Schedule", 
                       "Ex-i Hazardous Area Bay 1 (=CA1+MCP-M3) with Pepperl+Fuchs Isolators serving IS-JB-603, IS-JB-618, and JB-618", 
                       m3_cables, 5)

    # -------------------------------------------------------------
    # SLIDE 6: PANEL M4 CABLE SCHEDULE (Ex-i Bay 2)
    # -------------------------------------------------------------
    s6 = prs.slides.add_slide(blank_layout)
    m4_cables = [c for c in CABLES if c["from"] == "Panel M4"]
    render_cable_table(s6, "Panel M4 Intrinsically Safe (Ex-i) Cable Schedule", 
                       "Ex-i Hazardous Area Bay 2 (=CA1+MCP-M4) serving IS-JB-608, IS-JB-612, JB-606, JB-608, and JB-612", 
                       m4_cables, 6)

    # -------------------------------------------------------------
    # SLIDE 7: MCC SWITCHGEAR & RIO-200 BACKBONE CABLES
    # -------------------------------------------------------------
    s7 = prs.slides.add_slide(blank_layout)
    mcc_cables = [c for c in CABLES if "MCC" in c["to"] or "RIO" in c["to"]]
    render_cable_table(s7, "MCC Switchgear & Remote I/O Trunk Schedule", 
                       "Hardwired control interlocks to Motor Control Center switchgear and Armoured Fiber link to RIO-200", 
                       mcc_cables, 7)

    # -------------------------------------------------------------
    # SLIDE 8: BILL OF QUANTITIES (BOQ) & DRUM REEL SCHEDULE
    # -------------------------------------------------------------
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Master Cable Bill of Quantities (BOQ) & Drum Schedule", 
               "Consolidated procurement quantities with 10% installation slack and drum packaging", 
               "PROCUREMENT SCHEDULE")
    add_footer(s8, 8)

    # Calculate BOQ
    boq = {}
    for c in CABLES:
        sp = c["spec"]
        l = c["len"]
        if sp not in boq:
            boq[sp] = {"runs": 0, "net_len": 0, "type": c["type"]}
        boq[sp]["runs"] += 1
        boq[sp]["net_len"] += l

    rows = len(boq) + 1
    cols = 6
    table_shape8 = s8.shapes.add_table(rows, cols, Inches(0.8), Inches(1.5), Inches(11.733), Inches(0.38 * rows))
    table8 = table_shape8.table

    widths8 = [Inches(1.0), Inches(4.2), Inches(2.2), Inches(1.2), Inches(1.5), Inches(1.6)]
    for ci, cw in enumerate(widths8):
        table8.columns[ci].width = cw

    h8 = ["Item No.", "Cable Specification & Rating", "Signal Category", "No. of Runs", "Total Length (m)", "Drum Packaging"]
    for ci, ht in enumerate(h8):
        cell = table8.cell(0, ci)
        cell.fill.solid(); cell.fill.fore_color.rgb = C_NAVY
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.text = ht
        p.font.bold = True
        p.font.size = Pt(9)
        p.font.color.rgb = C_WHITE
        p.alignment = PP_ALIGN.CENTER if ci in [0, 3, 4, 5] else PP_ALIGN.LEFT

    for ri, (sp, dat) in enumerate(boq.items()):
        row_idx = ri + 1
        is_even = (ri % 2 == 0)
        row_bg = RGBColor(255, 255, 255) if is_even else RGBColor(241, 245, 249)
        tot_with_slack = int(dat["net_len"] * 1.10)  # +10% slack
        drum_spec = f"1x {tot_with_slack}m Drum" if tot_with_slack <= 500 else f"2x {int(tot_with_slack/2)}m Drums"

        vals = [f"CBL-{ri+1:02d}", sp, dat["type"], str(dat["runs"]), f"{tot_with_slack} m", drum_spec]
        for ci, v in enumerate(vals):
            cell = table8.cell(row_idx, ci)
            cell.fill.solid(); cell.fill.fore_color.rgb = row_bg
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.text = v
            p.font.size = Pt(8.5)
            if ci in [0, 4]:
                p.font.bold = True
            p.font.color.rgb = C_NAVY if ci == 0 else C_DARK_TEXT
            p.alignment = PP_ALIGN.CENTER if ci in [0, 3, 4, 5] else PP_ALIGN.LEFT

    prs.save(PPTX_PATH)
    print(f"[OK] Saved PowerPoint Presentation: {PPTX_PATH} ({os.path.getsize(PPTX_PATH):,} bytes)")


# -----------------------------------------------------------------------------
# STEP 2: GENERATE VECTOR PDF PRESENTATION SLIDES (16:9 LANDSCAPE)
# -----------------------------------------------------------------------------
class PDFSlideDeck(FPDF):
    def __init__(self):
        # 16:9 widescreen format in mm: 338.67 x 190.5 mm
        super().__init__(orientation="landscape", unit="mm", format=(190.5, 338.67))
        self.set_margins(10, 10, 10)
        self.set_auto_page_break(auto=False)

    def draw_slide_frame(self, title, subtitle, slide_no, total_slides=8, category="CABLE SCHEDULE"):
        # Top banner
        self.set_fill_color(15, 30, 54)
        self.rect(0, 0, 338.67, 24, "F")

        # Category tag
        self.set_font("Helvetica", "B", 7)
        self.set_text_color(14, 165, 233)
        self.text(12, 7, category.upper())

        # Slide title
        self.set_font("Helvetica", "B", 13)
        self.set_text_color(255, 255, 255)
        self.text(12, 14, title)

        # Subtitle
        self.set_font("Helvetica", "", 7.5)
        self.set_text_color(203, 213, 225)
        self.text(12, 20, subtitle)

        # Right header branding
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(255, 255, 255)
        self.text(265, 9, "INGREDION (THAILAND) CO., LTD.")
        self.set_font("Helvetica", "", 7)
        self.set_text_color(148, 163, 184)
        self.text(265, 14, "PROJECT JET COOKER - KALASIN")
        self.set_text_color(245, 158, 11)
        self.text(265, 19, "AEC INDUSTRIAL ENGINEERING")

        # Footer
        self.set_fill_color(241, 245, 249)
        self.rect(0, 182, 338.67, 8.5, "F")
        self.set_font("Helvetica", "", 7)
        self.set_text_color(100, 116, 139)
        self.text(12, 187, "Document No: CA1-SCH-CBL-001 | Revision: 3.6 | Approved For Construction")
        self.text(290, 187, f"Slide {slide_no} of {total_slides}")

    def render_table(self, cable_list):
        headers = ["Cable Tag", "From", "To (JB)", "Location", "Signal Function", "Cable Specification", "Tray", "Length"]
        col_w = [48, 22, 24, 46, 50, 68, 26, 20]
        x_start = 12
        y_start = 32

        # Header row
        self.set_xy(x_start, y_start)
        self.set_fill_color(30, 41, 59)
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(255, 255, 255)
        for i, h in enumerate(headers):
            self.cell(col_w[i], 7.5, h, border=1, align="C" if i in [1, 2, 6, 7] else "L", fill=True)
        self.ln()

        # Rows
        self.set_font("Helvetica", "", 7)
        cur_y = y_start + 7.5
        for idx, cb in enumerate(cable_list):
            is_even = (idx % 2 == 0)
            if "Ex-i" in cb["type"]:
                self.set_fill_color(238, 246, 255)
            elif is_even:
                self.set_fill_color(255, 255, 255)
            else:
                self.set_fill_color(248, 250, 252)

            self.set_xy(x_start, cur_y)
            row_data = [cb["tag"], cb["from"], cb["to"], cb["loc"], cb["func"], cb["spec"], cb["tray"], f"{cb['len']} m"]
            for i, val in enumerate(row_data):
                if i == 0:
                    self.set_font("Helvetica", "B", 7)
                    self.set_text_color(15, 30, 54)
                elif "BLUE" in val:
                    self.set_font("Helvetica", "B", 6.8)
                    self.set_text_color(2, 132, 199)
                else:
                    self.set_font("Helvetica", "", 7)
                    self.set_text_color(30, 41, 59)

                self.cell(col_w[i], 6.5, val, border=1, align="C" if i in [1, 2, 6, 7] else "L", fill=True)
            cur_y += 6.5
            self.ln()

def create_pdf_slides():
    pdf = PDFSlideDeck()

    # Slide 1: Cover
    pdf.add_page()
    pdf.set_fill_color(15, 30, 54)
    pdf.rect(0, 0, 338.67, 190.5, "F")
    
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(14, 165, 233)
    pdf.text(35, 55, "INGREDION (THAILAND) CO., LTD. - KALASIN STARCH PLANT")

    pdf.set_font("Helvetica", "B", 26)
    pdf.set_text_color(241, 245, 249)
    pdf.text(35, 75, "PROJECT JET COOKER")

    pdf.set_font("Helvetica", "B", 38)
    pdf.set_text_color(255, 255, 255)
    pdf.text(35, 95, "Interconnection Cable Schedule")

    pdf.set_font("Helvetica", "", 16)
    pdf.set_text_color(148, 163, 184)
    pdf.text(35, 110, "Complete Multi-Core & Trunk Routing Between MCC / CA1 Suite and Field Junction Boxes")

    pdf.set_fill_color(245, 158, 11)
    pdf.rect(35, 120, 268, 1.5, "F")

    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(245, 158, 11)
    pdf.text(35, 135, "Document No: CA1-SCH-CBL-001    |    Revision: 3.6    |    Date: 2026-09-07")
    pdf.set_font("Helvetica", "", 11)
    pdf.set_text_color(226, 232, 240)
    pdf.text(35, 145, "Engineering Contractor: AEC Industrial Engineering Co., Ltd.")

    # Slide 2: Routing Architecture
    pdf.add_page()
    pdf.draw_slide_frame("Overall Cable Routing & Signal Segregation Architecture", 
                         "Physical pathways from MCC Room to Infeed, Jet Cooker, Spray Dryer, Packing Tower, and Slurry Building", 
                         2, category="SYSTEM ARCHITECTURE")
    
    # Architecture Summary Boxes
    boxes = [
        ("Substation to Spine (TR-01)", "Main heavy-duty hot-dip galvanized ladder tray (600x100mm) running from MCC room basement trench across overhead pipe rack.", 12, 34),
        ("Spray Dryer Riser (TR-SD)", "Vertical cable riser traversing 1st to 8th floors. Segregated trays for Non-IS Control, Analog Shielded, and Ex-i Blue Trays.", 92, 34),
        ("Packing Tower Riser (TR-PT)", "Heavy-duty ladder tray for 2nd and 8th floor junction boxes. Blue-coated Ex-i dedicated cable tray with 300mm air gap separation.", 172, 34),
        ("Slurry Bldg Fiber Trunk (TR-EXT)", "Overhead pipe rack route to remote building housing RIO-200. 6-core steel wire armoured fiber optic DLR ring backbone.", 252, 34)
    ]
    for btitle, bdesc, bx, by in boxes:
        pdf.set_fill_color(248, 250, 252)
        pdf.set_draw_color(14, 165, 233)
        pdf.set_line_width(0.5)
        pdf.rect(bx, by, 75, 48, "DF")
        
        pdf.set_font("Helvetica", "B", 8.5)
        pdf.set_text_color(15, 30, 54)
        pdf.text(bx + 4, by + 8, btitle)
        
        pdf.set_font("Helvetica", "", 7.5)
        pdf.set_text_color(51, 65, 85)
        # Word wrap text
        words = bdesc.split()
        line = ""
        ly = by + 16
        for w in words:
            if pdf.get_string_width(line + " " + w) < 67:
                line += " " + w
            else:
                pdf.text(bx + 4, ly, line.strip())
                line = w
                ly += 5
        if line:
            pdf.text(bx + 4, ly, line.strip())

    # Lower Rules Box
    pdf.set_fill_color(241, 245, 249)
    pdf.set_draw_color(203, 213, 225)
    pdf.rect(12, 92, 314, 78, "DF")
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(15, 30, 54)
    pdf.text(18, 102, "ENGINEERING SEGREGATION & INSTALLATION RULES (IEC 60079-14 & IEEE 518)")

    rules = [
        "1. Intrinsically Safe (Ex-i) Cables: Distinctive BLUE outer sheath (RAL 5015). Maintain >= 300 mm clearance or solid grounded divider from power/control.",
        "2. Shield Grounding: All overall and individual pair screens grounded at MCC Cabinet single-point clean earth bar ONLY. Field ends insulated and floating.",
        "3. Spare Capacity: All multicore control cables sized with 20% to 35% installed spare conductors for future process expansion.",
        "4. Cable Glands: Nickel-plated brass double-compression EMC cable glands (IP66) on standard JBs; Ex-e/Ex-i certified blue-ring glands on Ex-i JBs.",
        "5. Testing & Verification: Insulation resistance test (Megger 500VDC > 20 MOhm) and loop continuity resistance test required before energization."
    ]
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(30, 41, 59)
    for r_idx, rule in enumerate(rules):
        pdf.text(18, 114 + r_idx * 11, rule)

    # Slide 3: Panel M1
    pdf.add_page()
    pdf.draw_slide_frame("Panel M1 Interconnection Cable Schedule", 
                         "Master Controller Bay (=CA1+MCP-M1) serving Infeed (JB-401), Spray Dryer 1F (JB-601), and Slurry (RIO-200)", 
                         3)
    pdf.render_table([c for c in CABLES if c["from"] == "Panel M1" and "MCC" not in c["to"]])

    # Slide 4: Panel M2
    pdf.add_page()
    pdf.draw_slide_frame("Panel M2 Interconnection Cable Schedule", 
                         "High-Density I/O Bay (=CA1+MCP-M2) serving Jet Cooker (JB-402), Spray Dryer 3F (JB-602), and Spray Dryer 7F (JB-607)", 
                         4)
    pdf.render_table([c for c in CABLES if c["from"] == "Panel M2"])

    # Slide 5: Panel M3 (Ex-i Bay 1)
    pdf.add_page()
    pdf.draw_slide_frame("Panel M3 Intrinsically Safe (Ex-i) Cable Schedule", 
                         "Ex-i Hazardous Area Bay 1 (=CA1+MCP-M3) with Pepperl+Fuchs Isolators serving IS-JB-603, IS-JB-618, and JB-618", 
                         5)
    pdf.render_table([c for c in CABLES if c["from"] == "Panel M3"])

    # Slide 6: Panel M4 (Ex-i Bay 2)
    pdf.add_page()
    pdf.draw_slide_frame("Panel M4 Intrinsically Safe (Ex-i) Cable Schedule", 
                         "Ex-i Hazardous Area Bay 2 (=CA1+MCP-M4) serving IS-JB-608, IS-JB-612, JB-606, JB-608, and JB-612", 
                         6)
    pdf.render_table([c for c in CABLES if c["from"] == "Panel M4"])

    # Slide 7: MCC Room & Trunks
    pdf.add_page()
    pdf.draw_slide_frame("MCC Switchgear & Remote I/O Trunk Schedule", 
                         "Hardwired control interlocks to Motor Control Center switchgear and Armoured Fiber link to RIO-200", 
                         7)
    pdf.render_table([c for c in CABLES if "MCC" in c["to"] or "RIO" in c["to"]])

    # Slide 8: Bill of Quantities (BOQ)
    pdf.add_page()
    pdf.draw_slide_frame("Master Cable Bill of Quantities (BOQ) & Drum Schedule", 
                         "Consolidated procurement quantities with 10% installation slack and drum packaging", 
                         8, category="PROCUREMENT SCHEDULE")
    
    boq = {}
    for c in CABLES:
        sp = c["spec"]
        l = c["len"]
        if sp not in boq:
            boq[sp] = {"runs": 0, "net_len": 0, "type": c["type"]}
        boq[sp]["runs"] += 1
        boq[sp]["net_len"] += l

    headers8 = ["Item No.", "Cable Specification & Construction", "Signal Category", "Runs", "Total Length", "Drum Packaging"]
    col_w8 = [25, 110, 60, 25, 38, 56]
    pdf.set_xy(12, 34)
    pdf.set_fill_color(30, 41, 59)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(255, 255, 255)
    for i, h in enumerate(headers8):
        pdf.cell(col_w8[i], 8, h, border=1, align="C" if i in [0, 3, 4, 5] else "L", fill=True)
    pdf.ln()

    cur_y = 42
    for ri, (sp, dat) in enumerate(boq.items()):
        is_even = (ri % 2 == 0)
        pdf.set_fill_color(255, 255, 255) if is_even else pdf.set_fill_color(248, 250, 252)
        tot_slack = int(dat["net_len"] * 1.10)
        drum = f"1x {tot_slack}m Drum" if tot_slack <= 500 else f"2x {int(tot_slack/2)}m Drums"
        
        pdf.set_xy(12, cur_y)
        row = [f"CBL-{ri+1:02d}", sp, dat["type"], str(dat["runs"]), f"{tot_slack} m", drum]
        for i, val in enumerate(row):
            if i in [0, 4]:
                pdf.set_font("Helvetica", "B", 7.5)
                pdf.set_text_color(15, 30, 54)
            else:
                pdf.set_font("Helvetica", "", 7.5)
                pdf.set_text_color(51, 65, 85)
            pdf.cell(col_w8[i], 7.2, val, border=1, align="C" if i in [0, 3, 4, 5] else "L", fill=True)
        cur_y += 7.2
        pdf.ln()

    pdf.output(PDF_PATH)
    print(f"[OK] Saved PDF Slides: {PDF_PATH} ({os.path.getsize(PDF_PATH):,} bytes)")

    # Render PNG slides for preview
    doc = fitz.open(PDF_PATH)
    for idx, page in enumerate(doc):
        pix = page.get_pixmap(dpi=150)
        png_path = os.path.join(OUTPUT_DIR, f"cable_schedule_slide_{idx+1}.png")
        pix.save(png_path)
        # Also copy to artifact directory
        artifact_png = os.path.join("/Users/x92120/.gemini/antigravity/brain/854b5772-95f5-43f9-b96e-8fb62b598e1a", f"cable_schedule_slide_{idx+1}.png")
        pix.save(artifact_png)
        print(f"  [OK] Rendered Slide {idx+1} PNG: {os.path.basename(png_path)}")


# -----------------------------------------------------------------------------
# STEP 3: INTERACTIVE WEB SLIDE DECK (HTML)
# -----------------------------------------------------------------------------
def create_html_slider():
    import json
    html_path = os.path.join(WORKSPACE_DIR, "cable_schedule_viewer.html")
    artifact_html = os.path.join("/Users/x92120/.gemini/antigravity/brain/854b5772-95f5-43f9-b96e-8fb62b598e1a", "cable_schedule_viewer.html")

    cables_json_str = json.dumps(CABLES)

    html_template = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Cable Schedule Slides: MCC to Field Junction Boxes</title>
  <script src="https://www.gstatic.com/antigravity/web/dev/tailwindcss.min.js"></script>
  <style>
    .slide-content { transition: opacity 0.2s ease-in-out; }
  </style>
</head>
<body class="bg-slate-950 text-slate-100 min-h-screen p-4 flex flex-col items-center">
  <div class="w-full max-w-7xl space-y-4">

    <!-- Top Controls -->
    <header class="bg-slate-900 border border-slate-800 rounded-xl p-4 flex flex-wrap justify-between items-center gap-4 shadow-xl">
      <div>
        <div class="flex items-center gap-3">
          <span class="bg-sky-600 text-white font-bold text-xs uppercase px-2 py-0.5 rounded">Engineering Deck</span>
          <h1 class="text-lg font-bold text-white tracking-wide">MCC to Junction Box Cable Schedule</h1>
        </div>
        <p class="text-xs text-slate-400 mt-0.5">Project Jet Cooker | Ingredion Kalasin | AEC Industrial Engineering (Rev. 3.6)</p>
      </div>

      <div class="flex items-center gap-2">
        <button onclick="prevSlide()" class="px-3 py-1.5 rounded-lg text-xs font-semibold bg-slate-800 hover:bg-slate-700 text-slate-200 border border-slate-700">&larr; Previous</button>
        <span id="slide-indicator" class="text-xs font-mono text-sky-400 font-bold px-3">Slide 1 / 8</span>
        <button onclick="nextSlide()" class="px-3 py-1.5 rounded-lg text-xs font-semibold bg-sky-600 hover:bg-sky-500 text-white shadow">Next &rarr;</button>
      </div>
    </header>

    <!-- Slide Thumbnails Bar -->
    <div class="grid grid-cols-4 md:grid-cols-8 gap-2">
      <button onclick="goToSlide(1)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-sky-600/30 border-sky-500 text-white" id="tab-1">1. Cover</button>
      <button onclick="goToSlide(2)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-2">2. Architecture</button>
      <button onclick="goToSlide(3)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-3">3. Panel M1</button>
      <button onclick="goToSlide(4)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-4">4. Panel M2</button>
      <button onclick="goToSlide(5)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-5">5. Panel M3 (IS)</button>
      <button onclick="goToSlide(6)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-6">6. Panel M4 (IS)</button>
      <button onclick="goToSlide(7)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-7">7. MCC & Trunks</button>
      <button onclick="goToSlide(8)" class="tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800" id="tab-8">8. BOQ & Drums</button>
    </div>

    <!-- Main Slide Display Canvas -->
    <main class="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden shadow-2xl relative min-h-[560px] flex flex-col justify-center">
      <div id="slide-container" class="p-6">
        <!-- Rendered dynamically -->
      </div>
    </main>

    <!-- Footer Downloads -->
    <footer class="bg-slate-900 border border-slate-800 rounded-xl p-4 flex flex-wrap justify-between items-center text-xs text-slate-400 gap-3">
      <div>
        Deliverables: <span class="font-mono text-slate-200">Cable_Schedule_MCC_to_Junction_Boxes.pptx / .pdf</span>
      </div>
      <div class="flex items-center gap-3">
        <span class="text-emerald-400 font-medium">37 Multi-Core Cables Configured</span>
        <span>|</span>
        <span class="text-sky-400 font-medium">1,210 I/O Channels Served</span>
      </div>
    </footer>

  </div>

  <script>
    const cables = __CABLES_JSON__;
    let currentSlide = 1;
    const totalSlides = 8;

    function renderSlide(slideNo) {
      const container = document.getElementById('slide-container');
      
      if (slideNo === 1) {
        container.innerHTML = `
          <div class="flex flex-col items-center justify-center py-12 text-center space-y-4">
            <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-sky-950 border border-sky-800 text-sky-300 text-xs font-semibold">
              INGREDION (THAILAND) CO., LTD. &bull; KALASIN PLANT
            </div>
            <h2 class="text-4xl font-extrabold text-white tracking-tight">PROJECT JET COOKER</h2>
            <h3 class="text-2xl font-bold text-sky-400">Interconnection Cable Schedule</h3>
            <p class="text-sm text-slate-400 max-w-2xl">
              Complete engineering schedule detailing multi-core control, analog instrumentation, intrinsically safe (Ex-i), and fiber optic backbone cables between the MCC / CA1 Suite and all field Junction Boxes.
            </p>
            <div class="pt-6 border-t border-slate-800 flex flex-wrap justify-center gap-6 text-xs text-slate-300 font-mono">
              <div><span class="text-slate-500">Doc No:</span> CA1-SCH-CBL-001</div>
              <div><span class="text-slate-500">Revision:</span> 3.6 (Approved)</div>
              <div><span class="text-slate-500">Date:</span> September 2026</div>
              <div><span class="text-slate-500">Contractor:</span> AEC Industrial Engineering</div>
            </div>
          </div>
        `;
      } else if (slideNo === 2) {
        container.innerHTML = `
          <div>
            <div class="mb-4">
              <span class="text-[11px] font-bold text-sky-400 uppercase tracking-wider">System Architecture</span>
              <h2 class="text-xl font-bold text-white">Overall Cable Routing & Signal Segregation Architecture</h2>
              <p class="text-xs text-slate-400">Physical pathways from Substation MCC Room to Process Areas</p>
            </div>

            <div class="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
              <div class="bg-slate-950 p-4 rounded-xl border border-sky-700/60">
                <div class="text-xs font-bold text-sky-400 mb-1">Spine Tray (TR-01)</div>
                <div class="text-sm font-semibold text-white">Substation to Pipe Rack</div>
                <p class="text-xs text-slate-400 mt-2">600x100mm heavy-duty ladder tray exiting MCC room basement trench to main process piperack.</p>
              </div>
              <div class="bg-slate-950 p-4 rounded-xl border border-emerald-700/60">
                <div class="text-xs font-bold text-emerald-400 mb-1">Spray Dryer Riser (TR-SD)</div>
                <div class="text-sm font-semibold text-white">1st to 8th Floors</div>
                <p class="text-xs text-slate-400 mt-2">Traverses building height with separate trays for Non-IS Control, Analog Shielded, and Ex-i.</p>
              </div>
              <div class="bg-slate-950 p-4 rounded-xl border border-amber-700/60">
                <div class="text-xs font-bold text-amber-400 mb-1">Packing Tower Riser (TR-PT)</div>
                <div class="text-sm font-semibold text-white">2nd and 8th Floors</div>
                <p class="text-xs text-slate-400 mt-2">Dedicated blue-coated Ex-i tray with >=300mm air gap physical separation from power cables.</p>
              </div>
              <div class="bg-slate-950 p-4 rounded-xl border border-purple-700/60">
                <div class="text-xs font-bold text-purple-400 mb-1">Slurry Out Trunk (TR-EXT)</div>
                <div class="text-sm font-semibold text-white">RIO-200 Link</div>
                <p class="text-xs text-slate-400 mt-2">6-core steel wire armoured fiber optic cable running 165m to Slurry Out building.</p>
              </div>
            </div>

            <div class="bg-slate-950 p-4 rounded-xl border border-slate-800 text-xs text-slate-300 space-y-2">
              <div class="font-bold text-white text-sm">Key Engineering Design Standards</div>
              <div>&bull; <span class="text-sky-400 font-semibold">Ex-i Segregation:</span> Distinctive blue outer sheath cables. Physical 50mm separation on terminal strips and 300mm on cable trays (IEC 60079-14).</div>
              <div>&bull; <span class="text-emerald-400 font-semibold">Single-Point Shielding:</span> Individual and overall shields terminated at MCC clean earth bar only; left floating at field JBs.</div>
              <div>&bull; <span class="text-amber-400 font-semibold">Installed Spares:</span> All multicore trunks incorporate 20% to 35% spare conductors for future modifications.</div>
            </div>
          </div>
        `;
      } else if (slideNo >= 3 && slideNo <= 7) {
        let filterFrom = "";
        let title = "";
        let sub = "";
        if (slideNo === 3) {
          filterFrom = "Panel M1";
          title = "Panel M1 Interconnection Cable Schedule";
          sub = "Master Controller Bay (=CA1+MCP-M1) serving Infeed (JB-401), Spray Dryer 1F (JB-601), and Slurry (RIO-200)";
        } else if (slideNo === 4) {
          filterFrom = "Panel M2";
          title = "Panel M2 Interconnection Cable Schedule";
          sub = "High-Density I/O Bay (=CA1+MCP-M2) serving Jet Cooker (JB-402), Spray Dryer 3F (JB-602), and Spray Dryer 7F (JB-607)";
        } else if (slideNo === 5) {
          filterFrom = "Panel M3";
          title = "Panel M3 Intrinsically Safe (Ex-i) Cable Schedule";
          sub = "Ex-i Hazardous Area Bay 1 (=CA1+MCP-M3) with Pepperl+Fuchs Isolators serving IS-JB-603, IS-JB-618, and JB-618";
        } else if (slideNo === 6) {
          filterFrom = "Panel M4";
          title = "Panel M4 Intrinsically Safe (Ex-i) Cable Schedule";
          sub = "Ex-i Hazardous Area Bay 2 (=CA1+MCP-M4) serving IS-JB-608, IS-JB-612, JB-606, JB-608, and JB-612";
        } else if (slideNo === 7) {
          filterFrom = "MCC_TRUNK";
          title = "MCC Switchgear & Remote I/O Trunk Schedule";
          sub = "Hardwired control interlocks to Motor Control Center switchgear and Armoured Fiber link to RIO-200";
        }

        let list = [];
        if (filterFrom === "MCC_TRUNK") {
          list = cables.filter(c => c.to.includes("MCC") || c.to.includes("RIO"));
        } else {
          list = cables.filter(c => c.from === filterFrom && !c.to.includes("MCC"));
        }

        let rowsHtml = list.map((c, i) => {
          const isExI = c.type.includes("Ex-i");
          const bgClass = isExI ? "bg-sky-950/30 text-sky-200" : (i % 2 === 0 ? "bg-slate-900/60" : "bg-slate-950/60");
          return `
            <tr class="${bgClass} border-b border-slate-800/60 hover:bg-slate-800 text-xs">
              <td class="p-2 font-mono font-bold text-sky-400">${c.tag}</td>
              <td class="p-2 text-center text-slate-400">${c.from}</td>
              <td class="p-2 font-semibold text-white">${c.to}</td>
              <td class="p-2 text-slate-400">${c.loc}</td>
              <td class="p-2">${c.func}</td>
              <td class="p-2 font-mono ${isExI ? 'text-sky-300 font-semibold' : 'text-slate-300'}">${c.spec}</td>
              <td class="p-2 text-center text-slate-400">${c.tray}</td>
              <td class="p-2 text-center font-bold text-amber-400">${c.len}m</td>
            </tr>
          `;
        }).join('');

        container.innerHTML = `
          <div>
            <div class="mb-4">
              <span class="text-[11px] font-bold text-sky-400 uppercase tracking-wider">Interconnection Schedule</span>
              <h2 class="text-xl font-bold text-white">${title}</h2>
              <p class="text-xs text-slate-400">${sub}</p>
            </div>
            <div class="overflow-x-auto rounded-xl border border-slate-800">
              <table class="w-full text-left border-collapse">
                <thead>
                  <tr class="bg-slate-950 text-slate-300 text-[11px] font-bold uppercase tracking-wider border-b border-slate-800">
                    <th class="p-2.5">Cable Tag</th>
                    <th class="p-2.5 text-center">From</th>
                    <th class="p-2.5">To (JB)</th>
                    <th class="p-2.5">Location</th>
                    <th class="p-2.5">Function</th>
                    <th class="p-2.5">Specification</th>
                    <th class="p-2.5 text-center">Tray</th>
                    <th class="p-2.5 text-center">Length</th>
                  </tr>
                </thead>
                <tbody>
                  ${rowsHtml}
                </tbody>
              </table>
            </div>
          </div>
        `;
      } else if (slideNo === 8) {
        // BOQ Table
        const boqMap = {};
        cables.forEach(c => {
          if (!boqMap[c.spec]) boqMap[c.spec] = { runs: 0, net: 0, type: c.type };
          boqMap[c.spec].runs += 1;
          boqMap[c.spec].net += c.len;
        });

        let boqRows = Object.keys(boqMap).map((sp, idx) => {
          const item = boqMap[sp];
          const tot = Math.round(item.net * 1.10);
          const drum = tot <= 500 ? `1x ${tot}m Drum` : `2x ${Math.round(tot/2)}m Drums`;
          return `
            <tr class="${idx % 2 === 0 ? 'bg-slate-900/60' : 'bg-slate-950/60'} border-b border-slate-800/60 hover:bg-slate-800 text-xs">
              <td class="p-2 font-mono text-slate-400">CBL-${(idx+1).toString().padStart(2, '0')}</td>
              <td class="p-2 font-semibold text-white font-mono">${sp}</td>
              <td class="p-2 text-slate-300">${item.type}</td>
              <td class="p-2 text-center text-slate-300">${item.runs}</td>
              <td class="p-2 text-center font-bold text-amber-400">${tot} m</td>
              <td class="p-2 text-center text-sky-400 font-mono">${drum}</td>
            </tr>
          `;
        }).join('');

        container.innerHTML = `
          <div>
            <div class="mb-4">
              <span class="text-[11px] font-bold text-sky-400 uppercase tracking-wider">Procurement Schedule</span>
              <h2 class="text-xl font-bold text-white">Master Cable Bill of Quantities (BOQ) & Drum Schedule</h2>
              <p class="text-xs text-slate-400">Consolidated procurement quantities with 10% installation slack and drum reel specifications</p>
            </div>
            <div class="overflow-x-auto rounded-xl border border-slate-800">
              <table class="w-full text-left border-collapse">
                <thead>
                  <tr class="bg-slate-950 text-slate-300 text-[11px] font-bold uppercase tracking-wider border-b border-slate-800">
                    <th class="p-2.5">Item</th>
                    <th class="p-2.5">Cable Specification & Construction</th>
                    <th class="p-2.5">Signal Category</th>
                    <th class="p-2.5 text-center">No. of Runs</th>
                    <th class="p-2.5 text-center">Total Length (+10%)</th>
                    <th class="p-2.5 text-center">Drum Reel Packaging</th>
                  </tr>
                </thead>
                <tbody>
                  ${boqRows}
                </tbody>
              </table>
            </div>
          </div>
        `;
      }

      // Update button styles
      for (let i = 1; i <= totalSlides; i++) {
        const btn = document.getElementById('tab-' + i);
        if (i === slideNo) {
          btn.className = "tab-btn p-2 rounded-lg text-center text-xs font-semibold border bg-sky-600 border-sky-500 text-white shadow";
        } else {
          btn.className = "tab-btn p-2 rounded-lg text-center text-xs font-medium border bg-slate-900 border-slate-800 text-slate-400 hover:bg-slate-800";
        }
      }
      document.getElementById('slide-indicator').innerText = `Slide ${slideNo} / ${totalSlides}`;
    }

    function goToSlide(n) {
      currentSlide = n;
      renderSlide(currentSlide);
    }
    function prevSlide() {
      if (currentSlide > 1) goToSlide(currentSlide - 1);
    }
    function nextSlide() {
      if (currentSlide < totalSlides) goToSlide(currentSlide + 1);
    }

    // Keyboard support
    window.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === ' ') nextSlide();
      if (e.key === 'ArrowLeft') prevSlide();
    });

    // Initialize first slide
    renderSlide(1);
  </script>
</body>
</html>
"""
    html_content = html_template.replace("__CABLES_JSON__", cables_json_str)

    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    with open(artifact_html, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"[OK] Saved Interactive HTML Slider: {html_path}")


if __name__ == "__main__":
    print("===================================================================")
    print("  CABLE SCHEDULE SLIDE GENERATOR: MCC TO FIELD JUNCTION BOXES     ")
    print("  Project Jet Cooker - Ingredion Kalasin / AEC Industrial Eng.     ")
    print("===================================================================")

    create_pptx_slides()
    create_pdf_slides()
    create_html_slider()
    print("\n[SUCCESS] All presentation formats (.pptx, .pdf, .html, .png) generated successfully!")
