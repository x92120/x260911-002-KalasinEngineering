#!/usr/bin/env python3
"""
Generate Clean Light-Themed Automation Architecture Slide Deck & Export to PDF
Customer: Ingredion (Thailand) Co., Ltd. - Kalasin Plant
Project: SPRINT 18K TPA Spray Dryer (Jet Cooker)

Includes:
1. Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx
2. Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf (5 Widescreen Pages)
3. Automation_Architecture_Slide_Diagram_Light.png (1920x1080 High-Res)
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image, ImageDraw, ImageFont
import fitz

# Target Folders
OUT_DIR_PROJECT = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUT_DIR_WORKSPACE = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"

PPTX_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx")
PPTX_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx")
PDF_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf")
PDF_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf")
PNG_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Slide_Diagram_Light.png")
PNG_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Slide_Diagram_Light.png")

# Light Theme Color Palette
L_BG = RGBColor(248, 250, 252)         # Slate 50
L_CARD_BG = RGBColor(255, 255, 255)    # Pure White
L_CARD_BORDER = RGBColor(203, 213, 225)# Slate 300
L_TEXT_DARK = RGBColor(15, 23, 42)     # Slate 900
L_TEXT_MUTED = RGBColor(100, 116, 139) # Slate 500
L_PRIMARY_BLUE = RGBColor(2, 132, 199) # Sky 600
L_DARK_BLUE = RGBColor(14, 76, 146)    # Corporate Navy
L_ACCENT_GREEN = RGBColor(22, 163, 74) # Emerald 600
L_ACCENT_AMBER = RGBColor(217, 119, 6) # Amber 600
L_ACCENT_PURPLE = RGBColor(147, 51, 234)# Purple 600

def get_font(size, bold=False):
    for p in [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc"
    ] :
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()

# -------------------------------------------------------------
# 1. RENDER 5 HIGH-RESOLUTION SLIDES AS IMAGES FOR PDF EXPORT
# -------------------------------------------------------------
def render_slide_images():
    slides_img = []
    W, H = 1920, 1080

    f_title = get_font(28, bold=True)
    f_h1 = get_font(22, bold=True)
    f_h2 = get_font(16, bold=True)
    f_sub = get_font(14, bold=True)
    f_body = get_font(12, bold=False)
    f_body_bold = get_font(12, bold=True)
    f_small = get_font(10, bold=False)

    # --- SLIDE 1: Title Card (Light) ---
    im1 = Image.new("RGB", (W, H), (248, 250, 252))
    d1 = ImageDraw.Draw(im1)
    # Background accent band
    d1.rectangle([0, 0, W, 20], fill=(14, 76, 146))
    d1.rectangle([0, H - 20, W, H], fill=(14, 76, 146))
    
    # Center Card
    d1.rectangle([140, 140, 1780, 940], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    # Header tag
    d1.rectangle([200, 200, 560, 240], fill=(224, 242, 254), outline=(2, 132, 199), width=1)
    d1.text((215, 212), "CONTROL SYSTEM ARCHITECTURE SPECIFICATION", font=f_sub, fill=(2, 132, 199))
    
    d1.text((200, 270), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT", font=f_h1, fill=(14, 76, 146))
    d1.text((200, 315), "Project SPRINT: 18K TPA Spray Dryer & Jet Cooker", font=f_title, fill=(15, 23, 42))
    d1.text((200, 370), "High-Availability Distributed Control System (Rockwell Automation PlantPAx 5.0 Baseline)", font=f_h2, fill=(100, 116, 139))
    
    # Key Metrics Boxes
    boxes = [
        ("2 SCADA Servers", "Fault-Tolerant Redundant Cluster\nPrimary & Secondary Hot Standby\nDell PowerEdge R550 Rackmount", (2, 132, 199), (240, 249, 255)),
        ("1 EWS Workstation", "Engineering & Diagnostics Station\nStudio 5000 Logix (v33+) & FT View\nDual 27\" 4K IPS Displays", (147, 51, 234), (250, 245, 255)),
        ("3 OWS Workstations", "Multi-Area Operator Clients\nOWS-01 (Slurry), OWS-02 (Dryer), OWS-03\nDual 27\" Industrial Monitors", (14, 76, 146), (241, 245, 249)),
        ("PLC Redundancy", "Redundant ControlLogix 5580\n1756-L83E + 1756-RM2 Fiber Sync\nBumpless Switchover < 20 ms", (22, 163, 74), (240, 253, 244))
    ]
    bx = 200
    for b_title, b_desc, b_color, b_bg in boxes:
        d1.rectangle([bx, 450, bx + 340, 680], fill=b_bg, outline=b_color, width=2)
        d1.text((bx + 20, 480), b_title, font=f_h1, fill=b_color)
        d1.text((bx + 20, 530), b_desc, font=f_body, fill=(30, 41, 59))
        bx += 380

    # Meta Footer
    d1.line([200, 780, 1720, 780], fill=(226, 232, 240), width=2)
    d1.text((200, 810), "Revision: P&ID Rev 3.5 Aligned  |  Engineering Consultant: AEC Industrial Engineering", font=f_body_bold, fill=(71, 85, 105))
    d1.text((200, 835), "Date of Approval: September 2026  |  Total Process Signals: 472 I/O Channels", font=f_body, fill=(100, 116, 139))
    slides_img.append(im1)

    # --- SLIDE 2: Core Architecture Diagram (Light) ---
    im2 = Image.new("RGB", (W, H), (248, 250, 252))
    d2 = ImageDraw.Draw(im2)
    
    # Top Header
    d2.rectangle([40, 30, 1880, 110], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d2.text((60, 42), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT | PROJECT SPRINT", font=f_sub, fill=(2, 132, 199))
    d2.text((60, 68), "Control & Automation System Architecture (2 Redundant Servers, 1 EWS, 3 OWS)", font=f_title, fill=(14, 76, 146))
    d2.text((1500, 50), "P&ID Rev 3.5 Baseline\nAllen-Bradley PlantPAx 5.0 DCS", font=f_body_bold, fill=(22, 163, 74))

    # LEVEL 3: SERVER TIER
    d2.rectangle([40, 130, 1880, 310], fill=(255, 255, 255), outline=(2, 132, 199), width=2)
    d2.text((60, 140), "LEVEL 3: SERVER ROOM & DATA CENTER (FAULT-TOLERANT SCADA CLUSTER)", font=f_sub, fill=(2, 132, 199))

    # Server A
    d2.rectangle([80, 175, 580, 290], fill=(240, 253, 244), outline=(22, 163, 74), width=2)
    d2.text((100, 188), "🖥️ SCADA SERVER A (PRIMARY ACTIVE)", font=f_sub, fill=(22, 163, 74))
    d2.text((100, 215), "Dell PowerEdge R550 Rackmount (Dual Xeon, 64GB, RAID-10)\nFactoryTalk View SE Server A | FactoryTalk Historian ME\nFactoryTalk Directory Server | Dual 800W Redundant PSU", font=f_body, fill=(15, 23, 42))

    # Sync Link
    d2.rectangle([610, 210, 750, 260], fill=(2, 132, 199))
    d2.text((630, 222), "⇄ SYNC LINK\nHeartbeat / Data", font=f_body_bold, fill=(255, 255, 255))

    # Server B
    d2.rectangle([780, 175, 1280, 290], fill=(254, 243, 199), outline=(217, 119, 6), width=2)
    d2.text((800, 188), "🖥️ SCADA SERVER B (STANDBY REDUNDANT)", font=f_sub, fill=(217, 119, 6))
    d2.text((800, 215), "Dell PowerEdge R550 Rackmount (Dual Xeon, 64GB, RAID-10)\nFactoryTalk View SE Server B (Hot Standby Failover < 1s)\nReal-time Data Replica | Dual 800W Redundant PSU", font=f_body, fill=(15, 23, 42))

    # NAS / Domain Controller
    d2.rectangle([1320, 175, 1840, 290], fill=(241, 245, 249), outline=(148, 163, 184), width=2)
    d2.text((1340, 188), "💾 CENTRAL NAS & DOMAIN CONTROLLER", font=f_sub, fill=(14, 76, 146))
    d2.text((1340, 215), "AssetCentre Version Control & Automated Disaster Recovery\nWindows Active Directory Domain Controller / DNS / NTP Server\nCentralized Engineering Project Archive", font=f_body, fill=(15, 23, 42))

    # LEVEL 2: CLIENT TIER (1 EWS + 3 OWS)
    d2.rectangle([40, 330, 1880, 520], fill=(255, 255, 255), outline=(14, 76, 146), width=2)
    d2.text((60, 340), "LEVEL 2: MAIN CONTROL ROOM (MCR) CLIENT WORKSTATIONS", font=f_sub, fill=(14, 76, 146))

    clients = [
        ("💻 1x EWS (ENGINEERING)", "Dell Precision Workstation (32GB, 1TB SSD)\nStudio 5000 Logix Designer (v33+)\nFactoryTalk View Studio Enterprise\nDual 27\" 4K Displays | Stratix Mgmt", 80, (147, 51, 234), (250, 245, 255)),
        ("🖥️ OWS-01 (OPERATOR 1)", "Area 202: Slurry Prep TS-20201/02\nArea 402: pH Adjust & Jet Cooker Skid\nFactoryTalk View SE Client\nDual 27\" Industrial Displays", 530, (2, 132, 199), (240, 249, 255)),
        ("🖥️ OWS-02 (OPERATOR 2)", "Area 602: Spray Dryer Tower CHAM-21\nTriplex HP Pump PD-60202 & Air Heater BMS\nBaghouse DTEX-21 & IEP Suppression\nDual 27\" Industrial Displays", 980, (2, 132, 199), (240, 249, 255)),
        ("🖥️ OWS-03 (OPERATOR 3)", "Area 613: Starch Powder Silo & Packing\nArea 922: SPRINT Steam Boiler (Getabec)\nArea 913/930/950: LPG, WTP & Compressed Air\nDual 27\" Industrial Displays", 1430, (2, 132, 199), (240, 249, 255))
    ]

    for ctitle, cdesc, cx, ccol, cbg in clients:
        d2.rectangle([cx, 370, cx + 410, 495], fill=cbg, outline=ccol, width=2)
        d2.text((cx + 15, 385), ctitle, font=f_sub, fill=ccol)
        d2.text((cx + 15, 412), cdesc, font=f_body, fill=(15, 23, 42))

    # ETHERNET SWITCHES BAR
    d2.rectangle([40, 540, 1880, 600], fill=(14, 76, 146), outline=(2, 132, 199), width=2)
    d2.text((320, 558), "⚡ PLANT ETHERNET BACKBONE RING: Stratix 5700/5400 Managed Industrial Gigabit Switches (Redundant Fiber Ring)", font=f_sub, fill=(255, 255, 255))

    # LEVEL 1: CONTROLLER TIER
    d2.rectangle([40, 620, 920, 840], fill=(255, 255, 255), outline=(22, 163, 74), width=2)
    d2.text((60, 635), "🎛️ REDUNDANT CONTROLLER CHASSIS (PLC-01A & PLC-01B)", font=f_sub, fill=(22, 163, 74))
    d2.text((60, 670), 
        "• Dual Allen-Bradley 1756-A4 Chassis in Main PLC Cabinet\n"
        "• Dual 1756-L83E GuardLogix 5580 Controller (10 MB Memory, 1 GHz Scan Engine)\n"
        "• 1756-RM2 Redundancy Modules: High-Speed Optical Cross-Chassis Synchronization\n"
        "• 1756-EN2TR Dual-Port EtherNet/IP Bridge Modules for Device Level Ring (DLR)\n"
        "• Dual 1756-PA75 Redundant Power Supply Bundles with Separate AC Feeders\n"
        "• Bumpless Failover Time: < 20 ms (Zero Process Disruption)", font=f_body, fill=(30, 41, 59))

    # LEVEL 0: DISTRIBUTED REMOTE I/O & MCC
    d2.rectangle([960, 620, 1880, 840], fill=(255, 255, 255), outline=(217, 119, 6), width=2)
    d2.text((980, 635), "📡 DISTRIBUTED REMOTE I/O PANELS (DLR RING TOPOLOGY)", font=f_sub, fill=(217, 119, 6))
    d2.text((980, 670),
        "• RIO-200 (Slurry Prep Area): 1794-AENTR Flex I/O w/ HART (1794-IB32, OB16, IF8IH)\n"
        "• RIO-400 (Conditioning & Jet Cooker): 1794-AENTR Flex I/O w/ HART\n"
        "• RIO-600 (Spray Dryer Tower): 1794-AENTR Flex I/O w/ HART\n"
        "• RIO-MCC (Motor Control Center): Intelligent Starters & PowerFlex 755/525 VFDs\n"
        "• Vendor Skid Integration: BMS Burner Panel, Getabec Boiler, IEP Explosion Suppression\n"
        "• Total 472 Physical I/O Points aligned with P&ID Rev 3.5 Specification", font=f_body, fill=(30, 41, 59))

    # Bottom Callout Summary
    d2.rectangle([40, 860, 1880, 1030], fill=(241, 245, 249), outline=(203, 213, 225), width=2)
    d2.text((60, 875), "HIGH AVAILABILITY & SYSTEM RELIABILITY HIGHLIGHTS:", font=f_sub, fill=(14, 76, 146))
    d2.text((60, 905),
        "1. Complete Server Redundancy: 2 Dell PowerEdge R550 servers running FactoryTalk View SE active/standby pair with sub-second failover.\n"
        "2. Multi-Operator Coverage: 3 dedicated OWS client workstations + 1 Engineering Workstation (EWS) covering all plant areas.\n"
        "3. Deterministic Safety & Control: Redundant ControlLogix 5580 CPU with fiber cross-sync and Device Level Ring (DLR) network (< 3ms self-healing).\n"
        "4. Seamless P&ID Rev 3.5 Baseline: Harmonized with 38 P&ID drawing sheets and 472 field tags across Kalasin Plant.",
        font=f_body, fill=(51, 65, 85))

    slides_img.append(im2)
    im2.save(PNG_LIGHT_1, quality=95)
    shutil.copyfile(PNG_LIGHT_1, PNG_LIGHT_2)

    # --- SLIDE 3: Hardware BOM Table (Light) ---
    im3 = Image.new("RGB", (W, H), (248, 250, 252))
    d3 = ImageDraw.Draw(im3)
    d3.rectangle([40, 30, 1880, 100], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d3.text((60, 45), "HARDWARE & SOFTWARE BILL OF MATERIALS (BOM) SPECIFICATION", font=f_h1, fill=(14, 76, 146))
    d3.text((60, 75), "Supporting 2 Redundant SCADA Servers, 1 EWS, and 3 OWS (P&ID Rev 3.5 Aligned)", font=f_body, fill=(100, 116, 139))

    # Table drawing
    tx, ty = 40, 120
    tw, th = 1840, 880
    d3.rectangle([tx, ty, tx + tw, ty + th], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    
    # Header row
    d3.rectangle([tx, ty, tx + tw, ty + 50], fill=(14, 76, 146))
    cols = [
        ("SYSTEM NODE", 60, 260),
        ("QTY", 340, 80),
        ("HARDWARE SPECIFICATION", 440, 520),
        ("SOFTWARE & LICENSES", 980, 500),
        ("REDUNDANCY ROLE", 1500, 360)
    ]
    for cname, cx, _ in cols:
        d3.text((cx, ty + 15), cname, font=f_sub, fill=(255, 255, 255))

    rows_data = [
        ("SCADA Servers (Redundant)", "2 Units", "Dell PowerEdge R550 Rackmount (Dual Intel Xeon Silver, 64GB RAM, 4x 960GB SSD RAID-10, Dual 800W PSU)", "Windows Server 2022 Datacenter\nFactoryTalk View SE Server v13\nFactoryTalk Historian ME & FTD", "Primary / Standby Pair\nAuto-Failover (< 1.0 sec)\nServer Room Rack"),
        ("Engineering Workstation (EWS)", "1 Unit", "Dell Precision 3660 Tower (Core i7, 32GB RAM, 1TB NVMe SSD, Dual 27\" 4K IPS Displays)", "Windows 11 Pro for Workstations\nStudio 5000 Logix Designer (v33+)\nFactoryTalk View Studio Enterprise", "Dedicated Dev/Diagnostics\nNon-Redundant\nMCR Engineering Desk"),
        ("Operator Workstations (OWS)", "3 Units", "Dell OptiPlex 7000 Series (Core i5, 16GB RAM, 512GB SSD, Dual 27\" Industrial Displays per desk)", "Windows 11 Enterprise LTSC\nFactoryTalk View SE Client v13\nAuto-Login Operator Shell Mode", "OWS-01: Slurry & Jet Cooker\nOWS-02: Spray Dryer & BMS\nOWS-03: Packing & Utilities"),
        ("Wall Overview Displays", "2 Units", "50\" 4K Industrial Commercial Display (24/7 Rating, HDMI Matrix Switcher)", "Live Plant Overview & Master Alarm\nPer Ingredion Kalasin Standard Spec", "MCR Wall Mounted\nMatrix Shared Video Feed"),
        ("ControlLogix Controller Pair", "2 Racks", "Dual 1756-A4 Chassis, Dual 1756-L83E CPU (10MB),\n1756-RM2 Redundancy, 1756-EN2TR DLR, 1756-PA75 PSU", "ControlLogix Firmware v33+\nEmbedded GuardLogix Safety Tasks", "Bumpless Optical Cross-Sync\nSwitchover < 20 ms"),
        ("Distributed Remote I/O", "4 Skids", "1794 Flex I/O w/ HART (1794-AENTR, IB32, OB16, IF8IH)\nRIO-200, RIO-400, RIO-600, RIO-MCC", "FactoryTalk Network Manager Config\nHART Instrument Device DTMs", "DLR Self-Healing Ring\nRecovery < 3 ms")
    ]

    ry = ty + 50
    for r_idx, r in enumerate(rows_data):
        r_bg = (248, 250, 252) if r_idx % 2 == 1 else (255, 255, 255)
        d3.rectangle([tx, ry, tx + tw, ry + 130], fill=r_bg)
        d3.line([tx, ry + 130, tx + tw, ry + 130], fill=(226, 232, 240), width=1)
        
        d3.text((60, ry + 20), r[0], font=f_body_bold, fill=(14, 76, 146))
        d3.text((340, ry + 20), r[1], font=f_body_bold, fill=(2, 132, 199))
        d3.text((440, ry + 15), r[2], font=f_body, fill=(30, 41, 59))
        d3.text((980, ry + 15), r[3], font=f_body, fill=(30, 41, 59))
        d3.text((1500, ry + 15), r[4], font=f_body_bold if "<" in r[4] else f_body, fill=(22, 163, 74) if "<" in r[4] else (71, 85, 105))
        ry += 130
    slides_img.append(im3)

    # --- SLIDE 4: Workstation Allocation (Light) ---
    im4 = Image.new("RGB", (W, H), (248, 250, 252))
    d4 = ImageDraw.Draw(im4)
    d4.rectangle([40, 30, 1880, 100], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d4.text((60, 45), "OPERATOR & ENGINEERING WORKSTATION ALLOCATION (MCR LAYOUT)", font=f_h1, fill=(14, 76, 146))
    d4.text((60, 75), "Functional duties and dual-display mapping for OWS-01, OWS-02, OWS-03, and EWS", font=f_body, fill=(100, 116, 139))

    cards_l = [
        ("OWS-01 : SLURRY PREP & JET COOKER", 40, (2, 132, 199), (240, 249, 255), [
            "Primary Responsibility:",
            "• Area 202: Re-Slurry Tanks TS-20201/02 & Agitators",
            "• Transfer Pump PC-20201/02 Flow Management",
            "• Area 402: pH Adjustment TS-40201/02 Skids",
            "• Continuous Jet Cooker Chamber & Steam Infeed",
            "• Acid, Caustic & Enzyme Dosing Loops",
            "",
            "Screen Configuration (Dual 27\" IPS):",
            "• Screen 1: Slurry Prep & pH Adjustment Mimic",
            "• Screen 2: Jet Cooker Real-Time Trends & Recipes"
        ]),
        ("OWS-02 : SPRAY DRYER & SAFETY BMS", 660, (22, 163, 74), (240, 253, 244), [
            "Primary Responsibility:",
            "• Area 602: Spray Drying Chamber CHAM-21",
            "• Dedert Triplex HP Pump PD-60202 (300 barg)",
            "• Direct Gas Air Heater FH-60210A & BMS Firing",
            "• Primary Cyclone CYCL-21 & Rotary Valves",
            "• Exhaust Fan BF-60240 & Baghouse DTEX-21",
            "• IEP Explosion Suppression & Chamber Deluge",
            "",
            "Screen Configuration (Dual 27\" IPS):",
            "• Screen 1: Spray Dryer Tower & Atomizer Mimic",
            "• Screen 2: High-Speed Trends & BMS Safety Interlocks"
        ]),
        ("OWS-03 : SILO, PACKING & UTILITIES", 1280, (217, 119, 6), (254, 243, 199), [
            "Primary Responsibility:",
            "• Area 613: Starch Powder Silo TS-61301 & Sifters",
            "• Big-Bag Packer ME-61306 & Robotic Palletizer",
            "• Area 922: SPRINT Steam Boiler (Getabec Package)",
            "• Area 913: LPG Bullet Tank & Vaporizer Station",
            "• Area 930: Water Treatment Plant Clarifier & Filters",
            "• Area 950: Plant Compressed Air System (6.8 barg)",
            "",
            "Screen Configuration (Dual 27\" IPS):",
            "• Screen 1: Packing & Bagging Lines",
            "• Screen 2: Steam Boiler & Plant Utilities Overview"
        ])
    ]

    for title, cx, ccol, cbg, lines in cards_l:
        d4.rectangle([cx, 130, cx + 580, 840], fill=cbg, outline=ccol, width=2)
        d4.rectangle([cx, 130, cx + 580, 180], fill=ccol)
        d4.text((cx + 20, 145), title, font=f_sub, fill=(255, 255, 255))
        
        ly = 210
        for line in lines:
            if line.endswith(":"):
                d4.text((cx + 25, ly), line, font=f_body_bold, fill=(15, 23, 42))
                ly += 25
            else:
                d4.text((cx + 25, ly), line, font=f_body, fill=(51, 65, 85))
                ly += 22

    # EWS Bottom Bar
    d4.rectangle([40, 870, 1880, 1030], fill=(250, 245, 255), outline=(147, 51, 234), width=2)
    d4.text((60, 885), "💻 1x EWS (ENGINEERING WORKSTATION) ROLE & MCR DESK SPECIFICATION:", font=f_sub, fill=(147, 51, 234))
    d4.text((60, 915),
        "• Dedicated development terminal for plant automation engineers: Online ControlLogix logic modification, I/O commissioning, and HMI design.\n"
        "• Equipped with Studio 5000 Logix Designer (v33+), FactoryTalk View Studio Enterprise, FactoryTalk Network Manager, and AssetCentre Client.\n"
        "• High-security Windows Active Directory domain role: Restricts unauthorized operator tampering while allowing full diagnostic visibility.",
        font=f_body, fill=(15, 23, 42))
    slides_img.append(im4)

    # --- SLIDE 5: High Availability Strategy (Light) ---
    im5 = Image.new("RGB", (W, H), (248, 250, 252))
    d5 = ImageDraw.Draw(im5)
    d5.rectangle([40, 30, 1880, 100], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d5.text((60, 45), "HIGH AVAILABILITY, REDUNDANCY STRATEGY & FAULT TOLERANCE", font=f_h1, fill=(14, 76, 146))
    d5.text((60, 75), "Comprehensive protection against hardware, power, network, and disk drive failures", font=f_body, fill=(100, 116, 139))

    boxes_s5_l = [
        ("SCADA Server Redundancy (< 1.0s Failover)", (2, 132, 199), (240, 249, 255), 40, 130, 900, 380, [
            "• Active/Standby FactoryTalk View SE Server architecture maintains real-time tag synchronization.",
            "• If Server A encounters hardware, OS, or network failure, Server B assumes mastership in < 1.0 second.",
            "• Zero operator interruption: All 3 OWS and EWS automatically redirect to Server B without session restart.",
            "• Dual 10GbE bonded NICs on Dell R550 servers eliminate network interface card dropouts."
        ]),
        ("PLC Controller Redundancy (< 20ms Bumpless)", (22, 163, 74), (240, 253, 244), 980, 130, 900, 380, [
            "• Dual ControlLogix 5580 chassis connected via 1756-RM2 high-speed optical fiber cross-sync modules.",
            "• Bumpless transfer: Secondary controller assumes mastership within 1 program scan (< 20 ms).",
            "• PID loop states, motor run statuses, and safety permissives are retained identically.",
            "• Dual 1756-PA75 power supplies on separate AC feeds prevent power outage disruption."
        ]),
        ("Ethernet Ring & DLR Self-Healing (< 3ms)", (217, 119, 6), (254, 243, 199), 40, 540, 900, 380, [
            "• Device Level Ring (DLR) network provides sub-3ms recovery upon any physical fiber or cable break.",
            "• Stratix 5700/5400 managed industrial switches provide VLAN segmentation for SCADA, I/O, and plant IT.",
            "• Resilient Ethernet Protocol (REP) backbone ring connects Server Room, MCR, and Electrical Rooms.",
            "• Continuous network monitoring through FactoryTalk Network Manager."
        ]),
        ("Cybersecurity & Disaster Governance (IEC 62443)", (147, 51, 234), (250, 245, 255), 980, 540, 900, 380, [
            "• FactoryTalk AssetCentre provides automated project version control, audit trails, and scheduled backups.",
            "• FactoryTalk Security integrated with Active Directory domain policy ensures strict role-based access.",
            "• Industrial firewalls isolate process OT network from enterprise IT and internet threats.",
            "• 20% Spare I/O capacity built into RIO-200, RIO-400, and RIO-600 Flex I/O chassis for future expansion."
        ])
    ]

    for title, bcol, bbg, bx, by, bw, bh, bullets in boxes_s5_l:
        d5.rectangle([bx, by, bx + bw, by + bh], fill=bbg, outline=bcol, width=2)
        d5.rectangle([bx, by, bx + bw, by + 45], fill=bcol)
        d5.text((bx + 20, by + 12), title, font=f_sub, fill=(255, 255, 255))
        
        ly = by + 65
        for b in bullets:
            d5.text((bx + 20, ly), b, font=f_body, fill=(15, 23, 42))
            ly += 30

    # Summary Footer
    d5.rectangle([40, 940, 1880, 1030], fill=(241, 245, 249), outline=(203, 213, 225), width=2)
    d5.text((60, 955), "EXECUTIVE ENGINEERING VERIFICATION:", font=f_sub, fill=(14, 76, 146))
    d5.text((60, 980), "This automation architecture fully satisfies the engineering requirements for Ingredion Kalasin Project SPRINT, ensuring zero single point of failure, seamless operator usability across all 3 OWS, and 100% compliance with P&ID Rev 3.5.", font=f_body_bold, fill=(51, 65, 85))

    slides_img.append(im5)
    return slides_img

# -------------------------------------------------------------
# 2. COMPILE TO MULTI-PAGE PDF USING PYMUPDF (FITZ)
# -------------------------------------------------------------
def compile_pdf(slides_img):
    doc = fitz.open()
    # 16:9 Widescreen dimensions in PDF points: 1920x1080 scaled to 960x540
    page_w, page_h = 960, 540
    
    for idx, im in enumerate(slides_img):
        page = doc.new_page(width=page_w, height=page_h)
        # Convert PIL Image to PNG bytes
        import io
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        png_bytes = buf.getvalue()
        
        rect = fitz.Rect(0, 0, page_w, page_h)
        page.insert_image(rect, stream=png_bytes)
        print(f"Added Page {idx + 1} to PDF")

    doc.save(PDF_LIGHT_1)
    doc.close()
    shutil.copyfile(PDF_LIGHT_1, PDF_LIGHT_2)
    print(f"Saved PDF to:\n  - {PDF_LIGHT_1}\n  - {PDF_LIGHT_2}")

# -------------------------------------------------------------
# 3. CREATE LIGHT PPTX PRESENTATION DECK
# -------------------------------------------------------------
def create_light_pptx():
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Slide 1
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = L_BG
    bg1.line.fill.background()

    card1 = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.8), Inches(11.733), Inches(5.9))
    card1.fill.solid()
    card1.fill.fore_color.rgb = L_CARD_BG
    card1.line.color.rgb = L_CARD_BORDER

    tf1 = card1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT"
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = L_PRIMARY_BLUE

    p2 = tf1.add_paragraph()
    p2.text = "Project SPRINT: 18K TPA Spray Dryer (Jet Cooker)"
    p2.font.size = Pt(32)
    p2.font.bold = True
    p2.font.color.rgb = L_DARK_BLUE
    p2.space_before = Pt(8)

    p3 = tf1.add_paragraph()
    p3.text = "Automation & Control Architecture Specification (2 Redundant Servers, 1 EWS, 3 OWS)"
    p3.font.size = Pt(18)
    p3.font.bold = True
    p3.font.color.rgb = L_TEXT_DARK
    p3.space_before = Pt(8)

    p4 = tf1.add_paragraph()
    p4.text = "Rockwell Automation PlantPAx 5.0 Baseline  |  P&ID Rev 3.5 Aligned  |  Zero Single Point of Failure"
    p4.font.size = Pt(13)
    p4.font.color.rgb = L_ACCENT_GREEN
    p4.space_before = Pt(14)

    # Slide 2
    s2 = prs.slides.add_slide(blank_layout)
    bg2 = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg2.fill.solid()
    bg2.fill.fore_color.rgb = L_BG
    bg2.line.fill.background()

    # Insert the high-res light diagram directly onto slide 2
    s2.shapes.add_picture(PNG_LIGHT_1, Inches(0.2), Inches(0.2), width=Inches(12.933))

    prs.save(PPTX_LIGHT_1)
    shutil.copyfile(PPTX_LIGHT_1, PPTX_LIGHT_2)
    print(f"Saved Light PPTX to:\n  - {PPTX_LIGHT_1}\n  - {PPTX_LIGHT_2}")

def main():
    print("=== Generating Light Architecture Slides & PDF Export ===")
    slides_img = render_slide_images()
    compile_pdf(slides_img)
    create_light_pptx()
    print("=== All Light Version Deliverables Generated Successfully! ===")

if __name__ == "__main__":
    main()
