#!/usr/bin/env python3
"""
Generate Automation & Control System Architecture Presentation Slides
Customer: Ingredion (Thailand) Co., Ltd. - Kalasin Plant
Project: SPRINT 18K TPA Spray Dryer (Jet Cooker)

System Architecture Elements:
- 2 Servers (Redundant): FactoryTalk View SE Primary & Secondary Servers (Dell PowerEdge R550)
- 1 EWS: Engineering Workstation (Studio 5000 Logix Designer & FT View Studio)
- 3 OWS: Operator Workstations (Dual Display clients for Process, Dryer & Utilities)
- Redundant Controller: Allen-Bradley ControlLogix 5580 (1756-L83E / 1756-RM2)
- Industrial Network: Stratix 5700/5400 Managed Fiber Ring + Device Level Ring (DLR)
- Field Remote I/O: RIO-200, RIO-400, RIO-600, RIO-MCC (Flex I/O 1794 w/ HART)
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, ImageFont

# Directories
OUT_DIR_PROJECT = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUT_DIR_WORKSPACE = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"

PPTX_PATH_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin.pptx")
PPTX_PATH_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin.pptx")
PNG_PATH_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Slide_Diagram.png")
PNG_PATH_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Slide_Diagram.png")

os.makedirs(OUT_DIR_PROJECT, exist_ok=True)
os.makedirs(OUT_DIR_WORKSPACE, exist_ok=True)

# Color Palette (Industrial Modern Tech Theme)
DARK_BG = RGBColor(15, 23, 42)       # Slate 900
CARD_BG = RGBColor(30, 41, 59)       # Slate 800
CARD_BORDER = RGBColor(71, 85, 105)  # Slate 600
PRIMARY_BLUE = RGBColor(14, 116, 144)# Cyan 700
ACCENT_BLUE = RGBColor(56, 189, 248) # Sky 400
ACCENT_GREEN = RGBColor(34, 197, 94) # Green 500
ACCENT_AMBER = RGBColor(245, 158, 11)# Amber 500
ACCENT_RED = RGBColor(239, 68, 68)   # Red 500
WHITE = RGBColor(255, 255, 255)
LIGHT_GRAY = RGBColor(203, 213, 225) # Slate 300
TEXT_MUTED = RGBColor(148, 163, 184) # Slate 400

def create_presentation():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # -------------------------------------------------------------
    # SLIDE 1: Title & Executive Summary
    # -------------------------------------------------------------
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = DARK_BG
    bg1.line.fill.background()

    # Title Card
    tb = s1.shapes.add_textbox(Inches(1.0), Inches(1.8), Inches(11.333), Inches(3.5))
    tf = tb.text_frame
    tf.word_wrap = True

    p = tf.paragraphs[0]
    p.text = "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT"
    p.font.size = Pt(20)
    p.font.bold = True
    p.font.color.rgb = ACCENT_BLUE

    p2 = tf.add_paragraph()
    p2.text = "Control & Automation Architecture Specification"
    p2.font.size = Pt(36)
    p2.font.bold = True
    p2.font.color.rgb = WHITE
    p2.space_before = Pt(12)

    p3 = tf.add_paragraph()
    p3.text = "Project SPRINT: 18,000 TPA Spray Dryer Plant & Jet Cooker System"
    p3.font.size = Pt(22)
    p3.font.color.rgb = LIGHT_GRAY
    p3.space_before = Pt(8)

    p4 = tf.add_paragraph()
    p4.text = "High-Availability Architecture: Dual Redundant SCADA Servers | 1 EWS | 3 OWS | Redundant ControlLogix 5580"
    p4.font.size = Pt(14)
    p4.font.color.rgb = ACCENT_GREEN
    p4.space_before = Pt(18)

    # Footer Metadata
    tb_meta = s1.shapes.add_textbox(Inches(1.0), Inches(5.8), Inches(11.333), Inches(1.0))
    tf_meta = tb_meta.text_frame
    pm = tf_meta.paragraphs[0]
    pm.text = "Revision: Rev 3.5 Aligned  |  Engineering Consultant: AEC Industrial Engineering  |  Date: September 2026"
    pm.font.size = Pt(12)
    pm.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 2: Core Automation System Architecture Diagram
    # -------------------------------------------------------------
    s2 = prs.slides.add_slide(blank_layout)
    bg2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = DARK_BG
    bg2.line.fill.background()

    # Header Ribbon
    header = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(0.3), Inches(12.333), Inches(0.7))
    header.fill.solid()
    header.fill.fore_color.rgb = CARD_BG
    header.line.color.rgb = CARD_BORDER
    tf_h = header.text_frame
    p_h = tf_h.paragraphs[0]
    p_h.text = "AUTOMATION & SCADA NETWORK ARCHITECTURE — PROJECT SPRINT (KALASIN)"
    p_h.font.size = Pt(16)
    p_h.font.bold = True
    p_h.font.color.rgb = WHITE
    p_h.alignment = PP_ALIGN.LEFT

    # LEVEL 3: SERVER TIER (2 Redundant Servers + Storage)
    # Background Group Box
    box_l3 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(1.15), Inches(12.333), Inches(1.4))
    box_l3.fill.solid()
    box_l3.fill.fore_color.rgb = RGBColor(20, 30, 48)
    box_l3.line.color.rgb = RGBColor(56, 189, 248)
    box_l3.line.width = Pt(1.5)

    lbl_l3 = s2.shapes.add_textbox(Inches(0.6), Inches(1.15), Inches(3.0), Inches(0.3))
    lbl_l3.text_frame.paragraphs[0].text = "LEVEL 3 : SERVER ROOM / DATA CENTER (FAULT-TOLERANT PAIR)"
    lbl_l3.text_frame.paragraphs[0].font.size = Pt(10)
    lbl_l3.text_frame.paragraphs[0].font.bold = True
    lbl_l3.text_frame.paragraphs[0].font.color.rgb = ACCENT_BLUE

    # Server A (Primary)
    srv_a = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.45), Inches(3.4), Inches(0.95))
    srv_a.fill.solid()
    srv_a.fill.fore_color.rgb = CARD_BG
    srv_a.line.color.rgb = ACCENT_GREEN
    srv_a.line.width = Pt(1.5)
    tf_sa = srv_a.text_frame
    tf_sa.word_wrap = True
    psa1 = tf_sa.paragraphs[0]
    psa1.text = "🖥️ SCADA SERVER A (PRIMARY)"
    psa1.font.size = Pt(12)
    psa1.font.bold = True
    psa1.font.color.rgb = ACCENT_GREEN
    psa2 = tf_sa.add_paragraph()
    psa2.text = "Dell PowerEdge R550 | Dual Xeon, 64GB, RAID-10\nFactoryTalk View SE Server A | Historian & FTD"
    psa2.font.size = Pt(9)
    psa2.font.color.rgb = LIGHT_GRAY

    # Redundancy Sync Channel between Servers
    sync_box = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(4.3), Inches(1.7), Inches(1.2), Inches(0.45))
    sync_box.fill.solid()
    sync_box.fill.fore_color.rgb = RGBColor(14, 116, 144)
    sync_box.line.fill.background()
    sync_tf = sync_box.text_frame
    psync = sync_tf.paragraphs[0]
    psync.text = "⇄ SYNC LINK\nHeartbeat / RAID"
    psync.font.size = Pt(8)
    psync.font.bold = True
    psync.font.color.rgb = WHITE
    psync.alignment = PP_ALIGN.CENTER

    # Server B (Secondary Standby)
    srv_b = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(1.45), Inches(3.4), Inches(0.95))
    srv_b.fill.solid()
    srv_b.fill.fore_color.rgb = CARD_BG
    srv_b.line.color.rgb = ACCENT_AMBER
    srv_b.line.width = Pt(1.5)
    tf_sb = srv_b.text_frame
    tf_sb.word_wrap = True
    psb1 = tf_sb.paragraphs[0]
    psb1.text = "🖥️ SCADA SERVER B (STANDBY REDUNDANT)"
    psb1.font.size = Pt(12)
    psb1.font.bold = True
    psb1.font.color.rgb = ACCENT_AMBER
    psb2 = tf_sb.add_paragraph()
    psb2.text = "Dell PowerEdge R550 | Dual Xeon, 64GB, RAID-10\nFactoryTalk View SE Server B (Hot Standby Failover)"
    psb2.font.size = Pt(9)
    psb2.font.color.rgb = LIGHT_GRAY

    # Network Storage & Domain Controller
    nas_box = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.2), Inches(1.45), Inches(3.4), Inches(0.95))
    nas_box.fill.solid()
    nas_box.fill.fore_color.rgb = CARD_BG
    nas_box.line.color.rgb = CARD_BORDER
    tf_nas = nas_box.text_frame
    pnas1 = tf_nas.paragraphs[0]
    pnas1.text = "💾 NAS STORAGE & DOMAIN CONTROLLER"
    pnas1.font.size = Pt(12)
    pnas1.font.bold = True
    pnas1.font.color.rgb = WHITE
    pnas2 = tf_nas.add_paragraph()
    pnas2.text = "Centralized Backup / AssetCentre / Active Directory\nRAID-6 Archive | UPS Redundant Power"
    pnas2.font.size = Pt(9)
    pnas2.font.color.rgb = LIGHT_GRAY

    # LEVEL 2: CONTROL ROOM TIER (1 EWS + 3 OWS)
    box_l2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(2.7), Inches(12.333), Inches(1.5))
    box_l2.fill.solid()
    box_l2.fill.fore_color.rgb = RGBColor(22, 33, 54)
    box_l2.line.color.rgb = RGBColor(14, 165, 233)
    box_l2.line.width = Pt(1.5)

    lbl_l2 = s2.shapes.add_textbox(Inches(0.6), Inches(2.7), Inches(4.0), Inches(0.3))
    lbl_l2.text_frame.paragraphs[0].text = "LEVEL 2 : MAIN CONTROL ROOM (MCR) CLIENT STATIONS"
    lbl_l2.text_frame.paragraphs[0].font.size = Pt(10)
    lbl_l2.text_frame.paragraphs[0].font.bold = True
    lbl_l2.text_frame.paragraphs[0].font.color.rgb = ACCENT_BLUE

    # 1 EWS (Engineering Workstation)
    ews = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(3.0), Inches(2.8), Inches(1.05))
    ews.fill.solid()
    ews.fill.fore_color.rgb = CARD_BG
    ews.line.color.rgb = RGBColor(168, 85, 247) # Purple
    ews.line.width = Pt(1.5)
    tf_ews = ews.text_frame
    tf_ews.word_wrap = True
    pews1 = tf_ews.paragraphs[0]
    pews1.text = "💻 1x EWS (ENGINEERING)"
    pews1.font.size = Pt(11)
    pews1.font.bold = True
    pews1.font.color.rgb = RGBColor(216, 180, 254)
    pews2 = tf_ews.add_paragraph()
    pews2.text = "Studio 5000 Logix Designer (v33+)\nFT View Studio Enterprise\nDual 27\" Displays | Stratix Mgmt"
    pews2.font.size = Pt(8.5)
    pews2.font.color.rgb = LIGHT_GRAY

    # 3 OWS (Operator Workstations)
    # OWS 1
    ows1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(3.8), Inches(3.0), Inches(2.8), Inches(1.05))
    ows1.fill.solid()
    ows1.fill.fore_color.rgb = CARD_BG
    ows1.line.color.rgb = ACCENT_BLUE
    ows1.line.width = Pt(1.5)
    tf_o1 = ows1.text_frame
    tf_o1.word_wrap = True
    po1 = tf_o1.paragraphs[0]
    po1.text = "🖥️ OWS-01 (OPERATOR 1)"
    po1.font.size = Pt(11)
    po1.font.bold = True
    po1.font.color.rgb = ACCENT_BLUE
    po1_2 = tf_o1.add_paragraph()
    po1_2.text = "Area 202: Slurry Preparation\nArea 402: pH Adjust & Jet Cooker\nFactoryTalk View SE Client"
    po1_2.font.size = Pt(8.5)
    po1_2.font.color.rgb = LIGHT_GRAY

    # OWS 2
    ows2 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(3.0), Inches(2.8), Inches(1.05))
    ows2.fill.solid()
    ows2.fill.fore_color.rgb = CARD_BG
    ows2.line.color.rgb = ACCENT_BLUE
    ows2.line.width = Pt(1.5)
    tf_o2 = ows2.text_frame
    tf_o2.word_wrap = True
    po2 = tf_o2.paragraphs[0]
    po2.text = "🖥️ OWS-02 (OPERATOR 2)"
    po2.font.size = Pt(11)
    po2.font.bold = True
    po2.font.color.rgb = ACCENT_BLUE
    po2_2 = tf_o2.add_paragraph()
    po2_2.text = "Area 602: Spray Dryer Tower\nHP Pump, BMS Burner, Baghouse\nFactoryTalk View SE Client"
    po2_2.font.size = Pt(8.5)
    po2_2.font.color.rgb = LIGHT_GRAY

    # OWS 3
    ows3 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(9.8), Inches(3.0), Inches(2.8), Inches(1.05))
    ows3.fill.solid()
    ows3.fill.fore_color.rgb = CARD_BG
    ows3.line.color.rgb = ACCENT_BLUE
    ows3.line.width = Pt(1.5)
    tf_o3 = ows3.text_frame
    tf_o3.word_wrap = True
    po3 = tf_o3.paragraphs[0]
    po3.text = "🖥️ OWS-03 (OPERATOR 3)"
    po3.font.size = Pt(11)
    po3.font.bold = True
    po3.font.color.rgb = ACCENT_BLUE
    po3_2 = tf_o3.add_paragraph()
    po3_2.text = "Area 613: Silo & Packing Tower\nUtilities: Steam, Air, LPG, WTP\nFactoryTalk View SE Client"
    po3_2.font.size = Pt(8.5)
    po3_2.font.color.rgb = LIGHT_GRAY

    # NETWORK SWITCHES BAR (Plant Ethernet Backbone)
    net_bar = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(4.35), Inches(12.333), Inches(0.45))
    net_bar.fill.solid()
    net_bar.fill.fore_color.rgb = RGBColor(15, 76, 129)
    net_bar.line.color.rgb = RGBColor(56, 189, 248)
    tf_net = net_bar.text_frame
    pnet = tf_net.paragraphs[0]
    pnet.text = "⚡ PLANT ETHERNET BACKBONE RING: Stratix 5700/5400 Managed Industrial Gigabit Switches (Redundant Fiber Ring)"
    pnet.font.size = Pt(10)
    pnet.font.bold = True
    pnet.font.color.rgb = WHITE
    pnet.alignment = PP_ALIGN.CENTER

    # LEVEL 1: CONTROLLER TIER (Redundant ControlLogix)
    box_l1 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.5), Inches(4.95), Inches(5.8), Inches(1.35))
    box_l1.fill.solid()
    box_l1.fill.fore_color.rgb = CARD_BG
    box_l1.line.color.rgb = ACCENT_GREEN
    box_l1.line.width = Pt(1.5)
    tf_l1 = box_l1.text_frame
    tf_l1.word_wrap = True
    pl1 = tf_l1.paragraphs[0]
    pl1.text = "🎛️ REDUNDANT CONTROLLER CHASSIS (PLC-01A & PLC-01B)"
    pl1.font.size = Pt(11)
    pl1.font.bold = True
    pl1.font.color.rgb = ACCENT_GREEN
    pl1_2 = tf_l1.add_paragraph()
    pl1_2.text = "• Allen-Bradley 1756-L83E GuardLogix (10MB Memory, High Perf)\n• 1756-RM2 Redundancy Modules (Fiber-Optic Cross-Sync)\n• Dual 1756-EN4TR EtherNet/IP DLR Modules & 1756-PA75 PSU\n• Bumpless controller switchover (<20ms)"
    pl1_2.font.size = Pt(8.5)
    pl1_2.font.color.rgb = LIGHT_GRAY

    # LEVEL 0: DISTRIBUTED REMOTE I/O & MCC (Right Side)
    box_l0 = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.5), Inches(4.95), Inches(6.333), Inches(1.35))
    box_l0.fill.solid()
    box_l0.fill.fore_color.rgb = CARD_BG
    box_l0.line.color.rgb = ACCENT_AMBER
    box_l0.line.width = Pt(1.5)
    tf_l0 = box_l0.text_frame
    tf_l0.word_wrap = True
    pl0 = tf_l0.paragraphs[0]
    pl0.text = "📡 DISTRIBUTED REMOTE I/O & MOTOR CONTROL (DLR RING)"
    pl0.font.size = Pt(11)
    pl0.font.bold = True
    pl0.font.color.rgb = ACCENT_AMBER
    pl0_2 = tf_l0.add_paragraph()
    pl0_2.text = "• RIO-200 (Slurry Prep): 1794-AENTR Flex I/O w/ HART (DI, DO, AI, AO)\n• RIO-400 (pH & Jet Cooker): 1794-AENTR Flex I/O w/ HART\n• RIO-600 (Spray Dryer Tower): 1794-AENTR Flex I/O w/ HART\n• RIO-MCC: Intelligent Motor Starters & PowerFlex 755/525 VFDs (EtherNet/IP)\n• Skid Integration: BMS Burner, IEP Explosion Suppression, Getabec Boiler"
    pl0_2.font.size = Pt(8.5)
    pl0_2.font.color.rgb = LIGHT_GRAY

    # Bottom Architecture Summary Note
    bot_box = s2.shapes.add_textbox(Inches(0.5), Inches(6.45), Inches(12.333), Inches(0.7))
    tf_bot = bot_box.text_frame
    pbot = tf_bot.paragraphs[0]
    pbot.text = "High-Availability Guarantee: Zero Single Point of Failure (SPOF) across Servers, PLC CPU, Network, and Power Supplies.\nTotal I/O Signals: 472 Physical Channels | Cycle Scan Time: < 20 ms | High Performance HMI (ANSI/ISA-101) & Alarm (ANSI/ISA-18.2)"
    pbot.font.size = Pt(9.5)
    pbot.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 3: Hardware & Specification Matrix Table
    # -------------------------------------------------------------
    s3 = prs.slides.add_slide(blank_layout)
    bg3 = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg3.fill.solid()
    bg3.fill.fore_color.rgb = DARK_BG
    bg3.line.fill.background()

    # Title
    t3 = s3.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.8))
    p3 = t3.text_frame.paragraphs[0]
    p3.text = "HARDWARE & SOFTWARE BILL OF MATERIALS (BOM) SPECIFICATION"
    p3.font.size = Pt(18)
    p3.font.bold = True
    p3.font.color.rgb = WHITE

    # Table
    rows, cols = 7, 5
    table_shape = s3.shapes.add_table(rows, cols, Inches(0.8), Inches(1.3), Inches(11.733), Inches(5.2))
    table = table_shape.table
    table.columns[0].width = Inches(2.2) # Role / Node
    table.columns[1].width = Inches(1.0) # Qty
    table.columns[2].width = Inches(3.2) # Hardware Specification
    table.columns[3].width = Inches(3.2) # Software & Licenses
    table.columns[4].width = Inches(2.133) # Redundancy / Location

    table_data = [
        ["System Component", "Qty", "Hardware Model / Spec", "Installed Software / OS", "Redundancy Role"],
        ["SCADA Servers (Redundant)", "2 Units", "Dell PowerEdge R550 (Dual Xeon Silver, 64GB RAM, RAID-10 SSD, Dual 800W PSU)", "Windows Server 2022 Datacenter\nFactoryTalk View SE Server v13\nFactoryTalk Historian ME / SE\nFactoryTalk Directory Server", "Primary / Secondary\nAuto-Failover (< 1 sec)\nServer Room Rack"],
        ["Engineering Workstation (EWS)", "1 Unit", "Dell Precision 3660 Tower\nCore i7, 32GB RAM, 1TB NVMe SSD\nDual 27\" 4K IPS Displays", "Windows 11 Pro for Workstations\nStudio 5000 Logix Designer (v33+)\nFactoryTalk View Studio Enterprise\nFactoryTalk Network Manager", "Non-redundant\nDedicated Dev/Diag\nMCR Engineering Desk"],
        ["Operator Workstations (OWS)", "3 Units", "Dell OptiPlex 7000 Micro / Tower\nCore i5, 16GB RAM, 512GB SSD\nDual 27\" Industrial Monitors", "Windows 11 Enterprise LTSC\nFactoryTalk View SE Client v13\nAuto-login Operator Shell Mode", "OWS-01: Slurry & Jet Cooker\nOWS-02: Spray Dryer & BMS\nOWS-03: Silo, Packing & Utility"],
        ["Large Wall Display Monitor", "2 Units", "50\" 4K Industrial Commercial Display\nContinuous 24/7 Rating, HDMI/DP", "Connected to OWS / Matrix Splitter\nPlant Overview & Alarm Marquee", "MCR Wall Mounted\n(Per IGD Kalasin Spec)"],
        ["ControlLogix Controller", "2 Racks", "1756-A4 Chassis (x2)\n1756-L83E CPU (x2)\n1756-RM2 Redundancy (x2)\n1756-EN4TR DLR Bridge (x2)", "ControlLogix Firmware v33+\nRedundancy Firmware Bundle\nEmbedded Safety Logic", "Bumpless Redundant Pair\nCross-Fiber Optical Sync\nMain PLC Cabinet"]
    ]

    for r_idx, row in enumerate(table_data):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if r_idx == 0:
                cell.fill.fore_color.rgb = RGBColor(14, 116, 144) # Header cyan
                for p in cell.text_frame.paragraphs:
                    p.font.bold = True
                    p.font.size = Pt(11)
                    p.font.color.rgb = WHITE
            else:
                cell.fill.fore_color.rgb = CARD_BG if r_idx % 2 == 1 else RGBColor(38, 52, 74)
                for p in cell.text_frame.paragraphs:
                    p.font.size = Pt(9.5)
                    p.font.color.rgb = LIGHT_GRAY

    # -------------------------------------------------------------
    # SLIDE 4: Workstation Functional Allocation & Operational Layout
    # -------------------------------------------------------------
    s4 = prs.slides.add_slide(blank_layout)
    bg4 = s4.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg4.fill.solid()
    bg4.fill.fore_color.rgb = DARK_BG
    bg4.line.fill.background()

    t4 = s4.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.8))
    p4 = t4.text_frame.paragraphs[0]
    p4.text = "OPERATOR & ENGINEERING WORKSTATION ALLOCATION (MCR LAYOUT)"
    p4.font.size = Pt(18)
    p4.font.bold = True
    p4.font.color.rgb = WHITE

    cards = [
        ("OWS-01 : PREPARATION & COOKING", Inches(0.8), Inches(1.3), Inches(3.6), Inches(4.8), ACCENT_BLUE, [
            "Primary Responsibility:",
            "• Area 202: Re-Slurry Tanks TS-20201/02 & Agitators",
            "• Slurry Transfer Pumps PC-20201 / PC-20202",
            "• Area 402: pH Adjustment TS-40201/02 Skids",
            "• Continuous Jet Cooker Chamber & Steam Infeed",
            "• Acid/Base Chemical Dosing Flow Controls",
            "",
            "Screen Configuration:",
            "• Screen 1: Re-Slurry & pH Mimic (Slide 2)",
            "• Screen 2: Batch Recipe, pH Trends & Jet Cooker Temp"
        ]),
        ("OWS-02 : SPRAY DRYER & SAFETY", Inches(4.866), Inches(1.3), Inches(3.6), Inches(4.8), ACCENT_GREEN, [
            "Primary Responsibility:",
            "• Area 602: Spray Drying Chamber CHAM-21",
            "• Triplex High-Pressure Pump PD-60202 (300 barg)",
            "• Direct Gas Air Heater FH-60210A & BMS Firing",
            "• Primary Cyclone CYCL-21 & Rotary Valves",
            "• Exhaust Fan BF-60240 & Baghouse DTEX-21",
            "• IEP Explosion Suppression & Chamber Deluge",
            "",
            "Screen Configuration:",
            "• Screen 1: Spray Dryer Tower Mimic (Slide 4)",
            "• Screen 2: High-Speed Trends & Critical Interlocks"
        ]),
        ("OWS-03 : PACKING & UTILITIES", Inches(8.933), Inches(1.3), Inches(3.6), Inches(4.8), ACCENT_AMBER, [
            "Primary Responsibility:",
            "• Area 613: Starch Powder Silo TS-61301 & Bin Vent",
            "• Vibratory Sifters SC-61301 & Big-Bag Packer ME-61306",
            "• Central Vacuum Cleaning System FA-61311",
            "• Area 922: SPRINT Steam Boiler (Getabec Package)",
            "• Area 913: LPG Bulk Bullet Tank & Vaporizer Station",
            "• Area 930: Water Treatment Clarifiers & Sand Filters",
            "• Area 950: Plant Compressed Air System (6.8 barg)",
            "",
            "Screen Configuration:",
            "• Screen 1: Packing & Bagging Lines",
            "• Screen 2: Plant Utilities & Boiler Overview"
        ])
    ]

    for title, cx, cy, cw, ch, col, lines in cards:
        c_shape = s4.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, cx, cy, cw, ch)
        c_shape.fill.solid()
        c_shape.fill.fore_color.rgb = CARD_BG
        c_shape.line.color.rgb = col
        c_shape.line.width = Pt(2)
        
        tf = c_shape.text_frame
        tf.word_wrap = True
        pt = tf.paragraphs[0]
        pt.text = title
        pt.font.size = Pt(12)
        pt.font.bold = True
        pt.font.color.rgb = col
        
        for l in lines:
            pl = tf.add_paragraph()
            pl.text = l
            if l.endswith(":"):
                pl.font.size = Pt(10)
                pl.font.bold = True
                pl.font.color.rgb = WHITE
                pl.space_before = Pt(6)
            else:
                pl.font.size = Pt(8.5)
                pl.font.color.rgb = LIGHT_GRAY

    # Footer note
    fn = s4.shapes.add_textbox(Inches(0.8), Inches(6.3), Inches(11.733), Inches(0.8))
    pfn = fn.text_frame.paragraphs[0]
    pfn.text = "EWS Station Role: 1 Dedicated Engineering Workstation located in MCR for online logic debugging, tag modifications, HMI screen engineering, and network maintenance. Restricted operator access with Windows Domain Active Directory."
    pfn.font.size = Pt(10)
    pfn.font.color.rgb = TEXT_MUTED

    # -------------------------------------------------------------
    # SLIDE 5: Network Topology & High Availability Guarantee
    # -------------------------------------------------------------
    s5 = prs.slides.add_slide(blank_layout)
    bg5 = s5.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg5.fill.solid()
    bg5.fill.fore_color.rgb = DARK_BG
    bg5.line.fill.background()

    t5 = s5.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(0.8))
    p5 = t5.text_frame.paragraphs[0]
    p5.text = "NETWORK TOPOLOGY & FAULT-TOLERANT REDUNDANCY STRATEGY"
    p5.font.size = Pt(18)
    p5.font.bold = True
    p5.font.color.rgb = WHITE

    boxes_s5 = [
        ("SCADA Server Redundancy (Active / Standby)", [
            "• FactoryTalk View SE Redundant Server Architecture maintains real-time active state synchronization.",
            "• In event of Server A failure (hardware, OS, or network), Server B takes over seamlessly without operator disruption.",
            "• Automatic client redirection: All 3 OWS and EWS automatically reconnect to Secondary Server in < 1 second.",
            "• Dual NIC Teaming on Dell R550 servers eliminates single NIC or patch cable vulnerability."
        ], Inches(0.8), Inches(1.3), Inches(5.6), Inches(2.4), ACCENT_BLUE),

        ("PLC Controller Redundancy (1756-RM2)", [
            "• ControlLogix 5580 dual chassis with high-speed 1 Gbps optical fiber redundancy link (1756-RM2 modules).",
            "• Bumpless transfer: Secondary controller assumes mastership within 1 program scan (< 20 ms).",
            "• Dual 1756-PA75 redundant power supplies prevent shutdown from utility feed dropouts.",
            "• Both primary and secondary racks interface to redundant EtherNet/IP DLR rings."
        ], Inches(6.933), Inches(1.3), Inches(5.6), Inches(2.4), ACCENT_GREEN),

        ("Ethernet Ring & DLR Topology (Stratix 5700)", [
            "• Resilient Ethernet Protocol (REP) backbone ring connecting Server Room, MCR, and Field Electrical Rooms.",
            "• Device Level Ring (DLR) provides sub-3ms recovery time upon any single fiber or cable break.",
            "• Segregated VLANs for SCADA client-server traffic, I/O multicast traffic, and Plant IT corporate connection.",
            "• Industrial Hirschmann / Stratix firewalls safeguard OT network from unauthorized external access."
        ], Inches(0.8), Inches(4.0), Inches(5.6), Inches(2.5), ACCENT_AMBER),

        ("Maintenance & Scalability Readiness", [
            "• 20% Spare I/O capacity built into RIO-200, RIO-400, and RIO-600 Flex I/O chassis.",
            "• Hot-Swappable RIUP (Removal and Insertion Under Power) for all 1794 Flex I/O modules.",
            "• FactoryTalk AssetCentre provides automated version control, audit trails, and scheduled disaster backups.",
            "• Full compliance with IEC 62443 Industrial Cybersecurity and ISA-101 HMI standard guidelines."
        ], Inches(6.933), Inches(4.0), Inches(5.6), Inches(2.5), RGBColor(168, 85, 247))
    ]

    for b_title, b_items, bx, by, bw, bh, bcol in boxes_s5:
        b_shp = s5.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, bx, by, bw, bh)
        b_shp.fill.solid()
        b_shp.fill.fore_color.rgb = CARD_BG
        b_shp.line.color.rgb = bcol
        b_shp.line.width = Pt(1.5)
        
        tf = b_shp.text_frame
        tf.word_wrap = True
        pt = tf.paragraphs[0]
        pt.text = b_title
        pt.font.size = Pt(11)
        pt.font.bold = True
        pt.font.color.rgb = bcol
        
        for item in b_items:
            pi = tf.add_paragraph()
            pi.text = item
            pi.font.size = Pt(8.5)
            pi.font.color.rgb = LIGHT_GRAY
            pi.space_before = Pt(3)

    prs.save(PPTX_PATH_1)
    shutil.copyfile(PPTX_PATH_1, PPTX_PATH_2)
    print(f"Saved PPTX to:\n  - {PPTX_PATH_1}\n  - {PPTX_PATH_2}")

# -------------------------------------------------------------
# Standalone High-Resolution Image Generator (1920x1080)
# -------------------------------------------------------------
def render_architecture_png():
    W, H = 1920, 1080
    im = Image.new("RGB", (W, H), (15, 23, 42))
    d = ImageDraw.Draw(im)

    def get_font(size, bold=False):
        for p in [
            "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
            "/System/Library/Fonts/Helvetica.ttc"
        ]:
            if os.path.exists(p):
                try:
                    return ImageFont.truetype(p, size)
                except Exception:
                    pass
        return ImageFont.load_default()

    f_title = get_font(24, bold=True)
    f_sub = get_font(15, bold=True)
    f_body = get_font(12, bold=False)
    f_body_bold = get_font(12, bold=True)
    f_tag = get_font(11, bold=True)

    # Top Header
    d.rectangle([40, 30, 1880, 110], fill=(30, 41, 59), outline=(71, 85, 105), width=2)
    d.text((60, 42), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT | PROJECT SPRINT", font=f_sub, fill=(56, 189, 248))
    d.text((60, 68), "Control & Automation System Architecture (2 Redundant Servers, 1 EWS, 3 OWS)", font=f_title, fill=(255, 255, 255))
    d.text((1500, 50), "P&ID Rev 3.5 Baseline\nAllen-Bradley PlantPAx 5.0 DCS", font=f_body_bold, fill=(34, 197, 94))

    # LEVEL 3: SERVER TIER
    d.rectangle([40, 130, 1880, 310], fill=(20, 30, 48), outline=(56, 189, 248), width=2)
    d.text((60, 140), "LEVEL 3: SERVER ROOM & DATA CENTER (FAULT-TOLERANT SCADA CLUSTER)", font=f_sub, fill=(56, 189, 248))

    # Server A
    d.rectangle([80, 175, 580, 290], fill=(30, 41, 59), outline=(34, 197, 94), width=2)
    d.text((100, 188), "🖥️ SCADA SERVER A (PRIMARY ACTIVE)", font=f_sub, fill=(34, 197, 94))
    d.text((100, 215), "Dell PowerEdge R550 Rackmount (Dual Xeon, 64GB, RAID-10)\nFactoryTalk View SE Server A | FactoryTalk Historian ME\nFactoryTalk Directory Server | Dual 800W Redundant PSU", font=f_body, fill=(203, 213, 225))

    # Sync Link
    d.rectangle([610, 210, 750, 260], fill=(14, 116, 144))
    d.text((630, 222), "⇄ SYNC LINK\nHeartbeat / Data", font=f_tag, fill=(255, 255, 255))

    # Server B
    d.rectangle([780, 175, 1280, 290], fill=(30, 41, 59), outline=(245, 158, 11), width=2)
    d.text((800, 188), "🖥️ SCADA SERVER B (STANDBY REDUNDANT)", font=f_sub, fill=(245, 158, 11))
    d.text((800, 215), "Dell PowerEdge R550 Rackmount (Dual Xeon, 64GB, RAID-10)\nFactoryTalk View SE Server B (Hot Standby Failover < 1s)\nReal-time Data Replica | Dual 800W Redundant PSU", font=f_body, fill=(203, 213, 225))

    # NAS / Domain Controller
    d.rectangle([1320, 175, 1840, 290], fill=(30, 41, 59), outline=(71, 85, 105), width=2)
    d.text((1340, 188), "💾 CENTRAL NAS & DOMAIN CONTROLLER", font=f_sub, fill=(255, 255, 255))
    d.text((1340, 215), "AssetCentre Version Control & Automated Disaster Recovery\nWindows Active Directory Domain Controller / DNS / NTP Server\nCentralized Engineering Project Archive", font=f_body, fill=(203, 213, 225))

    # LEVEL 2: CLIENT TIER (1 EWS + 3 OWS)
    d.rectangle([40, 330, 1880, 520], fill=(22, 33, 54), outline=(14, 165, 233), width=2)
    d.text((60, 340), "LEVEL 2: MAIN CONTROL ROOM (MCR) CLIENT WORKSTATIONS", font=f_sub, fill=(14, 165, 233))

    clients = [
        ("💻 1x EWS (ENGINEERING)", "Dell Precision Workstation (32GB, 1TB SSD)\nStudio 5000 Logix Designer (v33+)\nFactoryTalk View Studio Enterprise\nDual 27\" 4K Displays | Stratix Mgmt", 80, (168, 85, 247)),
        ("🖥️ OWS-01 (OPERATOR 1)", "Area 202: Slurry Prep TS-20201/02\nArea 402: pH Adjust & Jet Cooker Skid\nFactoryTalk View SE Client\nDual 27\" Industrial Displays", 530, (56, 189, 248)),
        ("🖥️ OWS-02 (OPERATOR 2)", "Area 602: Spray Dryer Tower CHAM-21\nTriplex HP Pump PD-60202 & Air Heater BMS\nBaghouse DTEX-21 & IEP Suppression\nDual 27\" Industrial Displays", 980, (56, 189, 248)),
        ("🖥️ OWS-03 (OPERATOR 3)", "Area 613: Starch Powder Silo & Packing\nArea 922: SPRINT Steam Boiler (Getabec)\nArea 913/930/950: LPG, WTP & Compressed Air\nDual 27\" Industrial Displays", 1430, (56, 189, 248))
    ]

    for ctitle, cdesc, cx, ccol in clients:
        d.rectangle([cx, 370, cx + 410, 495], fill=(30, 41, 59), outline=ccol, width=2)
        d.text((cx + 15, 385), ctitle, font=f_sub, fill=ccol)
        d.text((cx + 15, 412), cdesc, font=f_body, fill=(203, 213, 225))

    # ETHERNET SWITCHES BAR
    d.rectangle([40, 540, 1880, 600], fill=(15, 76, 129), outline=(56, 189, 248), width=2)
    d.text((320, 558), "⚡ PLANT ETHERNET BACKBONE RING: Stratix 5700/5400 Managed Industrial Gigabit Switches (Redundant Fiber Ring)", font=f_sub, fill=(255, 255, 255))

    # LEVEL 1: CONTROLLER TIER
    d.rectangle([40, 620, 920, 840], fill=(30, 41, 59), outline=(34, 197, 94), width=2)
    d.text((60, 635), "🎛️ REDUNDANT CONTROLLER CHASSIS (PLC-01A & PLC-01B)", font=f_sub, fill=(34, 197, 94))
    d.text((60, 670), 
        "• Dual Allen-Bradley 1756-A4 Chassis in Main PLC Cabinet\n"
        "• Dual 1756-L83E GuardLogix 5580 Controller (10 MB Memory, 1 GHz Scan Engine)\n"
        "• 1756-RM2 Redundancy Modules: High-Speed Optical Cross-Chassis Synchronization\n"
        "• 1756-EN4TR Dual-Port EtherNet/IP Bridge Modules for Device Level Ring (DLR)\n"
        "• Dual 1756-PA75 Redundant Power Supply Bundles with Separate AC Feeders\n"
        "• Bumpless Failover Time: < 20 ms (Zero Process Disruption)", font=f_body, fill=(203, 213, 225))

    # LEVEL 0: DISTRIBUTED REMOTE I/O & MCC
    d.rectangle([960, 620, 1880, 840], fill=(30, 41, 59), outline=(245, 158, 11), width=2)
    d.text((980, 635), "📡 DISTRIBUTED REMOTE I/O PANELS (DLR RING TOPOLOGY)", font=f_sub, fill=(245, 158, 11))
    d.text((980, 670),
        "• RIO-200 (Slurry Prep Area): 1794-AENTR Flex I/O w/ HART (1794-IB32, OB16, IF8IH)\n"
        "• RIO-400 (Conditioning & Jet Cooker): 1794-AENTR Flex I/O w/ HART\n"
        "• RIO-600 (Spray Dryer Tower): 1794-AENTR Flex I/O w/ HART\n"
        "• RIO-MCC (Motor Control Center): Intelligent Starters & PowerFlex 755/525 VFDs\n"
        "• Vendor Skid Integration: BMS Burner Panel, Getabec Boiler, IEP Explosion Suppression\n"
        "• Total 472 Physical I/O Points aligned with P&ID Rev 3.5 Specification", font=f_body, fill=(203, 213, 225))

    # Bottom Callout Summary
    d.rectangle([40, 860, 1880, 1030], fill=(20, 26, 36), outline=(71, 85, 105), width=2)
    d.text((60, 875), "HIGH AVAILABILITY & SYSTEM RELIABILITY HIGHLIGHTS:", font=f_sub, fill=(255, 215, 0))
    d.text((60, 905),
        "1. Complete Server Redundancy: 2 Dell PowerEdge R550 servers running FactoryTalk View SE active/standby pair with sub-second failover.\n"
        "2. Multi-Operator Coverage: 3 dedicated OWS client workstations + 1 Engineering Workstation (EWS) covering all plant areas.\n"
        "3. Deterministic Safety & Control: Redundant ControlLogix 5580 CPU with fiber cross-sync and Device Level Ring (DLR) network (< 3ms self-healing).\n"
        "4. Seamless P&ID Rev 3.5 Baseline: Harmonized with 38 P&ID drawing sheets and 472 field tags across Kalasin Plant.",
        font=f_body, fill=(203, 213, 225))

    im.save(PNG_PATH_1, quality=95)
    shutil.copyfile(PNG_PATH_1, PNG_PATH_2)
    print(f"Saved Architecture PNG to:\n  - {PNG_PATH_1}\n  - {PNG_PATH_2}")

if __name__ == "__main__":
    create_presentation()
    render_architecture_png()
