#!/usr/bin/env python3
"""
Generate Complete Automation Architecture & Detailed Network Diagram Slides & PDF
With Full Required Software & License Packages Annotated on Every Node.
Customer: Ingredion (Thailand) Co., Ltd. - Kalasin Plant
Project: SPRINT 18K TPA Spray Dryer (Jet Cooker)

Deliverables:
- Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf (6 Widescreen Pages)
- Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx (6 Slides)
- Automation_Architecture_Project_SPRINT_Kalasin.pptx (Dark Version, 6 Slides)
- Network_Topology_Diagram_Light.png (1920x1080)
- Automation_Architecture_Slide_Diagram_Light.png (1920x1080)
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

OUT_DIR_PROJECT = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUT_DIR_WORKSPACE = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"

PDF_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf")
PDF_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf")
PPTX_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx")
PPTX_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx")
PPTX_DARK_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin.pptx")
PPTX_DARK_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin.pptx")
NET_PNG_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Network_Topology_Diagram_Light.png")
NET_PNG_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Network_Topology_Diagram_Light.png")

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

# -------------------------------------------------------------
# RENDER NETWORK TOPOLOGY DIAGRAM WITH REQUIRED SOFTWARE (LIGHT)
# -------------------------------------------------------------
def render_network_diagram_light():
    W, H = 1920, 1080
    im = Image.new("RGB", (W, H), (248, 250, 252))
    d = ImageDraw.Draw(im)

    f_title = get_font(24, bold=True)
    f_h1 = get_font(16, bold=True)
    f_h2 = get_font(13, bold=True)
    f_sub = get_font(11, bold=True)
    f_body = get_font(10, bold=False)
    f_body_bold = get_font(10, bold=True)
    f_tag = get_font(9, bold=True)
    f_small = get_font(8.5, bold=False)

    # Top Header
    d.rectangle([40, 20, 1880, 95], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d.text((60, 30), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT | PROJECT SPRINT (18K TPA)", font=f_h2, fill=(2, 132, 199))
    d.text((60, 54), "Industrial Network Topology & Required Software Architecture (CPwE / DLR)", font=f_title, fill=(14, 76, 146))
    d.text((1440, 34), "P&ID Rev 3.5 Aligned\nFull Software & License Stack", font=f_sub, fill=(22, 163, 74))

    # LEVEL 3.5: ENTERPRISE DMZ & INDUSTRIAL SECURITY FIREWALL
    d.rectangle([40, 105, 1880, 165], fill=(255, 255, 255), outline=(225, 29, 72), width=2)
    d.rectangle([40, 105, 1880, 130], fill=(255, 241, 242))
    d.text((60, 112), "LEVEL 3.5 : INDUSTRIAL DEMILITARIZED ZONE (IDMZ) & FIREWALL", font=f_h2, fill=(225, 29, 72))

    d.rectangle([80, 135, 480, 160], fill=(254, 226, 226), outline=(225, 29, 72), width=1)
    d.text((90, 142), "🔒 FIREWALL: Stratix 5950 / Cisco Firepower", font=f_tag, fill=(159, 18, 57))
    d.text((500, 142), "Corporate IT (10.x.x.x) <--- Deep Packet Inspection (CIP/HTTPS/SSH) ---> Plant OT (192.168.x.x)", font=f_body_bold, fill=(71, 85, 105))

    # LEVEL 3: SCADA SERVER & DATA MANAGEMENT SUBNET (VLAN 10)
    d.rectangle([40, 175, 1060, 480], fill=(255, 255, 255), outline=(2, 132, 199), width=2)
    d.rectangle([40, 175, 1060, 205], fill=(240, 249, 255))
    d.text((60, 183), "LEVEL 3 : SERVER ROOM SCADA CLUSTER (VLAN 10: 192.168.10.0/24)", font=f_h2, fill=(2, 132, 199))

    # Core Switch Stratix 5400
    d.rectangle([60, 215, 1040, 255], fill=(14, 76, 146))
    d.text((80, 225), "CORE DISTRIBUTION SWITCH (SW-CORE-01): Stratix 5400 Managed Gigabit Switch", font=f_sub, fill=(255, 255, 255))
    d.text((820, 225), "IP: 192.168.10.1 (GW)", font=f_tag, fill=(254, 240, 138))

    # Server A
    d.rectangle([60, 265, 380, 465], fill=(240, 253, 244), outline=(22, 163, 74), width=2)
    d.text((75, 275), "🖥️ SCADA SERVER A", font=f_h2, fill=(22, 163, 74))
    d.text((75, 295), "PRIMARY ACTIVE [IP: 192.168.10.11]", font=f_tag, fill=(22, 163, 74))
    d.text((75, 315), "HARDWARE: Dell PowerEdge R550 Rack", font=f_small, fill=(71, 85, 105))
    d.line([75, 332, 365, 332], fill=(187, 247, 208), width=1)
    d.text((75, 338), "REQUIRED SOFTWARE STACK:", font=f_tag, fill=(15, 23, 42))
    d.text((75, 355), 
        "• Windows Server 2022 Datacenter\n"
        "• FactoryTalk View SE Server v13.0\n"
        "  (Primary Server License)\n"
        "• FactoryTalk Historian ME Server\n"
        "  (5,000 Tag License)\n"
        "• FactoryTalk Network Directory Server\n"
        "• FactoryTalk Linx Gateway (OPC-UA)", font=f_small, fill=(15, 23, 42))

    # Heartbeat Link
    d.rectangle([395, 345, 485, 405], fill=(2, 132, 199))
    d.text((402, 355), "⇄ SYNC LINK\n10GbE Bonded\nHeartbeat\n< 1.0s Failover", font=f_tag, fill=(255, 255, 255))

    # Server B
    d.rectangle([500, 265, 820, 465], fill=(254, 243, 199), outline=(217, 119, 6), width=2)
    d.text((515, 275), "🖥️ SCADA SERVER B", font=f_h2, fill=(217, 119, 6))
    d.text((515, 295), "HOT STANDBY [IP: 192.168.10.12]", font=f_tag, fill=(217, 119, 6))
    d.text((515, 315), "HARDWARE: Dell PowerEdge R550 Rack", font=f_small, fill=(71, 85, 105))
    d.line([515, 332, 805, 332], fill=(254, 215, 170), width=1)
    d.text((515, 338), "REQUIRED SOFTWARE STACK:", font=f_tag, fill=(15, 23, 42))
    d.text((515, 355), 
        "• Windows Server 2022 Datacenter\n"
        "• FactoryTalk View SE Server v13.0\n"
        "  (Secondary Redundant Partner Lic)\n"
        "• FactoryTalk Redundancy Manager\n"
        "• FactoryTalk Historian Standby Sync\n"
        "• FactoryTalk Linx Redundant Data Srv", font=f_small, fill=(15, 23, 42))

    # NAS / Domain Controller
    d.rectangle([835, 265, 1040, 465], fill=(241, 245, 249), outline=(148, 163, 184), width=2)
    d.text((845, 275), "💾 NAS ARCHIVE / DC", font=f_sub, fill=(14, 76, 146))
    d.text((845, 295), "[IP: 192.168.10.15]", font=f_tag, fill=(71, 85, 105))
    d.line([845, 312, 1030, 312], fill=(203, 213, 225), width=1)
    d.text((845, 318), "REQUIRED SOFTWARE:", font=f_tag, fill=(15, 23, 42))
    d.text((845, 335), 
        "• Windows Server 2022\n"
        "• Active Directory (AD DS)\n"
        "• FactoryTalk AssetCentre\n"
        "  Server (v12 Backup & Audit)\n"
        "• Windows DNS Server\n"
        "• Windows NTP Time Server", font=f_small, fill=(15, 23, 42))

    # LEVEL 2: MAIN CONTROL ROOM (MCR) CLIENT STATIONS (VLAN 10)
    d.rectangle([1080, 175, 1880, 480], fill=(255, 255, 255), outline=(14, 76, 146), width=2)
    d.rectangle([1080, 175, 1880, 205], fill=(241, 245, 249))
    d.text((1100, 183), "LEVEL 2 : MAIN CONTROL ROOM (MCR) CLIENTS (VLAN 10: 192.168.10.0/24)", font=f_h2, fill=(14, 76, 146))

    # MCR Switch
    d.rectangle([1100, 215, 1860, 255], fill=(14, 76, 146))
    d.text((1120, 225), "MCR DISTRIBUTION SWITCH (SW-MCR-01): Stratix 5700 Managed Switch | IP: 192.168.10.2", font=f_sub, fill=(255, 255, 255))

    # Clients Grid
    mcr_nodes = [
        ("💻 1x EWS (ENG)", "IP: 192.168.10.20", "Dell Precision 3660 Workstation\nDual 27\" 4K Displays", [
            "• Windows 11 Pro Workstations",
            "• Studio 5000 Logix Designer (v33+)",
            "• Studio 5000 Safety Editor",
            "• FactoryTalk View Studio Enterprise",
            "• FactoryTalk Network Manager",
            "• FactoryTalk AssetCentre Client"
        ], 1100, (147, 51, 234), (250, 245, 255)),

        ("🖥️ OWS-01", "IP: 192.168.10.21", "Slurry & Jet Cooker\nDual 27\" Displays", [
            "• Windows 11 Enterprise LTSC",
            "• FactoryTalk View SE Client v13",
            "• FT Alarm & Event Client",
            "• FT View TrendPro Viewer",
            "• FT Desktop Lock / Shell Mode",
            "• Client Redundancy Auto-Switch"
        ], 1290, (2, 132, 199), (240, 249, 255)),

        ("🖥️ OWS-02", "IP: 192.168.10.22", "Spray Dryer & BMS\nDual 27\" Displays", [
            "• Windows 11 Enterprise LTSC",
            "• FactoryTalk View SE Client v13",
            "• FT Alarm & Event Client",
            "• FT View TrendPro Viewer",
            "• FT Desktop Lock / Shell Mode",
            "• Client Redundancy Auto-Switch"
        ], 1480, (2, 132, 199), (240, 249, 255)),

        ("🖥️ OWS-03", "IP: 192.168.10.23", "Silo, Packing & Util\nDual 27\" Displays", [
            "• Windows 11 Enterprise LTSC",
            "• FactoryTalk View SE Client v13",
            "• FT Alarm & Event Client",
            "• FT View TrendPro Viewer",
            "• FT Desktop Lock / Shell Mode",
            "• Client Redundancy Auto-Switch"
        ], 1670, (2, 132, 199), (240, 249, 255))
    ]

    for ctitle, cip, chw, csofts, cx, ccol, cbg in mcr_nodes:
        d.rectangle([cx, 265, cx + 180, 465], fill=cbg, outline=ccol, width=2)
        d.text((cx + 8, 272), ctitle, font=f_h2, fill=ccol)
        d.text((cx + 8, 292), cip, font=f_tag, fill=(71, 85, 105))
        d.text((cx + 8, 308), chw, font=f_small, fill=(100, 116, 139))
        d.line([cx + 8, 332, cx + 172, 332], fill=ccol, width=1)
        d.text((cx + 8, 338), "REQUIRED SOFTWARE:", font=f_tag, fill=(15, 23, 42))
        ly = 355
        for s in csofts:
            d.text((cx + 8, ly), s, font=f_small, fill=(15, 23, 42))
            ly += 17

    # LEVEL 1 & 0: CONTROLLER & DEVICE LEVEL RING (DLR) (VLAN 30: 192.168.30.0/24)
    d.rectangle([40, 495, 1880, 875], fill=(255, 255, 255), outline=(22, 163, 74), width=2)
    d.rectangle([40, 495, 1880, 525], fill=(240, 253, 244))
    d.text((60, 503), "LEVEL 1 & LEVEL 0 : CONTROLLER REDUNDANCY & FAULT-TOLERANT DEVICE LEVEL RING (DLR)", font=f_h2, fill=(22, 163, 74))
    d.text((1200, 505), "Sub-3ms Self-Healing Recovery | Zero Single Point of Failure (SPOF)", font=f_sub, fill=(14, 76, 146))

    # Redundant ControlLogix Chassis A & B
    # Chassis A
    d.rectangle([60, 540, 490, 680], fill=(240, 253, 244), outline=(22, 163, 74), width=2)
    d.text((75, 550), "🎛️ CONTROLLER CHASSIS A (PRIMARY)", font=f_h2, fill=(22, 163, 74))
    d.text((75, 572), "1756-EN4TR DLR Bridge: IP 192.168.30.1 (Ring Supervisor)", font=f_tag, fill=(22, 163, 74))
    d.line([75, 588, 475, 588], fill=(187, 247, 208), width=1)
    d.text((75, 594), "REQUIRED FIRMWARE & LOGIC:", font=f_tag, fill=(15, 23, 42))
    d.text((75, 610), 
        "• 1756-L83E GuardLogix 5580 Controller Firmware: v33.011+ (Redundant)\n"
        "• 1756-RM2 Enhanced Optical Redundancy Firmware Bundle\n"
        "• 1756-EN4TR EtherNet/IP DLR Embedded Ring Supervisor Firmware\n"
        "• PlantPAx 5.0 Process Object Library (Process Library v5.00.00)", font=f_small, fill=(15, 23, 42))

    # Fiber Cross-Sync Link
    d.rectangle([510, 585, 630, 640], fill=(147, 51, 234))
    d.text((518, 595), "1756-RM2 FIBER\nCross-Chassis Sync\nSwitchover < 20ms", font=f_tag, fill=(255, 255, 255))

    # Chassis B
    d.rectangle([650, 540, 1080, 680], fill=(254, 243, 199), outline=(217, 119, 6), width=2)
    d.text((665, 550), "🎛️ CONTROLLER CHASSIS B (STANDBY)", font=f_h2, fill=(217, 119, 6))
    d.text((665, 572), "1756-EN4TR DLR Bridge: IP 192.168.30.2 (Backup Supervisor)", font=f_tag, fill=(217, 119, 6))
    d.line([665, 588, 1065, 588], fill=(254, 215, 170), width=1)
    d.text((665, 594), "REQUIRED FIRMWARE & LOGIC:", font=f_tag, fill=(15, 23, 42))
    d.text((665, 610), 
        "• 1756-L83E GuardLogix 5580 Controller Firmware: v33.011+ (Identical)\n"
        "• 1756-RM2 Enhanced Optical Redundancy Firmware Bundle\n"
        "• 1756-EN4TR EtherNet/IP DLR Embedded Ring Backup Firmware\n"
        "• Continuous Real-Time State & Memory Synchronization", font=f_small, fill=(15, 23, 42))

    # Stratix Switch AOP box (Right of controllers)
    d.rectangle([1100, 540, 1860, 680], fill=(241, 245, 249), outline=(14, 76, 146), width=2)
    d.text((1115, 550), "⚡ NETWORK SWITCH MANAGEMENT & DIAGNOSTIC SOFTWARE", font=f_h2, fill=(14, 76, 146))
    d.text((1115, 572), "Integrated Rockwell Stratix Switch Architecture", font=f_tag, fill=(71, 85, 105))
    d.line([1115, 588, 1845, 588], fill=(203, 213, 225), width=1)
    d.text((1115, 594), "REQUIRED CONFIGURATION SOFTWARE & DRIVERS:", font=f_tag, fill=(15, 23, 42))
    d.text((1115, 610),
        "• Stratix 5400 / 5700 CIP-Enabled Firmware (Cisco IOS-XE / Resilient Ethernet Protocol)\n"
        "• Studio 5000 Stratix Add-On Profile (AOP) — Real-time port status & bandwidth in PLC tags\n"
        "• FactoryTalk Network Manager v2.0+ — Live graphical network topology & port health\n"
        "• Device Level Ring (DLR) Faceplates for FactoryTalk View SE (Sub-3ms loop monitoring)", font=f_small, fill=(15, 23, 42))

    # DLR Ring Nodes (Remote I/O and OEM Skids)
    dlr_nodes = [
        ("📡 RIO-200 (Slurry)", "IP: 192.168.30.11\n1794-AENTR Flex w/ HART\nDTM Instrument Drivers", 60, (2, 132, 199)),
        ("📡 RIO-400 (Infeed)", "IP: 192.168.30.12\n1794-AENTR Flex w/ HART\nDTM Instrument Drivers", 325, (2, 132, 199)),
        ("📡 RIO-600 (Dryer)", "IP: 192.168.30.13\n1794-AENTR Flex w/ HART\nDTM Instrument Drivers", 590, (2, 132, 199)),
        ("⚡ RIO-MCC (VFDs)", "IP: 192.168.30.14\nPowerFlex 755/525 VFDs\nConnected Components WB", 855, (217, 119, 6)),
        ("🔥 BMS SKID", "IP: 192.168.30.21\nBurner Safety Gateway\nModbus-TCP / ENet Profile", 1120, (225, 29, 72)),
        ("🏭 BOILER SKID", "IP: 192.168.30.22\nGetabec Boiler Gateway\nModbus-TCP / ENet Profile", 1385, (14, 76, 146)),
        ("💥 IEP SKID", "IP: 192.168.30.23\nExplosion Suppression\nHRD Safety Interlock Profile", 1650, (147, 51, 234))
    ]

    # Draw Closed Loop DLR Ring Line
    d.line([275, 680, 275, 710], fill=(22, 163, 74), width=4) # From Chassis A
    d.line([865, 680, 865, 710], fill=(217, 119, 6), width=4) # From Chassis B
    d.line([180, 710, 1770, 710], fill=(22, 163, 74), width=4) # Main ring upper rail
    d.line([180, 845, 1770, 845], fill=(22, 163, 74), width=4) # Main ring lower rail
    d.line([180, 710, 180, 845], fill=(22, 163, 74), width=4)   # Left ring end
    d.line([1770, 710, 1770, 845], fill=(22, 163, 74), width=4) # Right ring end

    for ntitle, ndesc, nx, ncol in dlr_nodes:
        d.rectangle([nx, 730, nx + 250, 825], fill=(255, 255, 255), outline=ncol, width=2)
        d.text((nx + 10, 738), ntitle, font=f_h2, fill=ncol)
        d.text((nx + 10, 760), ndesc, font=f_small, fill=(15, 23, 42))

    # BOTTOM SOFTWARE BOM & LICENSE MATRIX TABLE
    d.rectangle([40, 890, 1880, 1055], fill=(241, 245, 249), outline=(203, 213, 225), width=2)
    d.text((60, 900), "COMPLETE AUTOMATION SOFTWARE & LICENSE BILL OF MATERIALS (BOM):", font=f_h2, fill=(14, 76, 146))

    soft_rows = [
        ("SCADA Server Cluster (Primary & Secondary)", "FactoryTalk View SE Server v13.0 (Redundant Pair)", "9358-VWSE000LENE", "2 Licenses", "Active-Standby Redundant SCADA Server Engine"),
        ("Industrial Data Historian", "FactoryTalk Historian ME / SE (Standard Edition)", "9518-HISTMESE5K", "1 Server + 5K Tags", "High-speed process archive & trending (100ms sample)"),
        ("Engineering Workstation (EWS)", "Studio 5000 Professional & FT View Studio", "9324-RLD700NXENE", "1 Named License", "Full PLC programming, safety logic & HMI screen dev"),
        ("Operator Workstations (OWS-01, 02, 03)", "FactoryTalk View SE Client (Desktop Shell Mode)", "9358-VWC100LENE", "3 Client Licenses", "Dedicated operator stations with locked OS desktop shell"),
        ("Disaster Recovery & Version Control", "FactoryTalk AssetCentre Server & Client", "9515-ASTCNTR01", "1 Server License", "Automated project backup, audit trail & disaster recovery")
    ]

    sy = 925
    for s_role, s_pkg, s_cat, s_qty, s_desc in soft_rows:
        d.text((60, sy), s_role, font=f_body_bold, fill=(14, 76, 146))
        d.text((440, sy), s_pkg, font=f_body, fill=(15, 23, 42))
        d.text((880, sy), s_cat, font=f_tag, fill=(71, 85, 105))
        d.text((1070, sy), s_qty, font=f_body_bold, fill=(2, 132, 199))
        d.text((1250, sy), s_desc, font=f_body, fill=(22, 163, 74))
        sy += 23

    im.save(NET_PNG_LIGHT_1, quality=95)
    shutil.copyfile(NET_PNG_LIGHT_1, NET_PNG_LIGHT_2)
    print("Rendered Network Diagram PNG Light with Software Stack successfully.")
    return im

# -------------------------------------------------------------
# COMPILE UPDATED 6-PAGE PDF
# -------------------------------------------------------------
def build_updated_pdf():
    import generate_light_architecture_slides as gls
    slides_img = gls.render_slide_images() # Returns slides 1, 2, 3, 4, 5
    
    # Render network diagram as slide 3 with software stack
    net_img = render_network_diagram_light()
    
    # Insert network diagram as Slide 3 (between Slide 2 Architecture and Slide 4 BOM)
    full_slides = [slides_img[0], slides_img[1], net_img, slides_img[2], slides_img[3], slides_img[4]]
    
    doc = fitz.open()
    page_w, page_h = 960, 540
    import io
    for idx, im in enumerate(full_slides):
        page = doc.new_page(width=page_w, height=page_h)
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        page.insert_image(fitz.Rect(0, 0, page_w, page_h), stream=buf.getvalue())
        print(f"Added Page {idx+1} to PDF")
        
    doc.save("/tmp/temp_6page_soft.pdf", garbage=4, deflate=True)
    doc.close()
    
    shutil.copyfile("/tmp/temp_6page_soft.pdf", PDF_LIGHT_1)
    shutil.copyfile("/tmp/temp_6page_soft.pdf", PDF_LIGHT_2)
    
    size_mb = os.path.getsize(PDF_LIGHT_1) / (1024 * 1024)
    print(f"Saved 6-Page PDF with Software Stack to:\n  - {PDF_LIGHT_1} ({size_mb:.2f} MB)")

# -------------------------------------------------------------
# UPDATE PPTX DECKS WITH NETWORK DIAGRAM SLIDE
# -------------------------------------------------------------
def update_pptx_decks():
    # Light PPTX
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    
    # Slide 1: Cover
    s1 = prs.slides.add_slide(blank_layout)
    s1.shapes.add_picture(os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Slide_Diagram_Light.png"), Inches(0.2), Inches(0.2), width=Inches(12.933))

    # Slide 2: Network Diagram with Software Stack
    s2 = prs.slides.add_slide(blank_layout)
    s2.shapes.add_picture(NET_PNG_LIGHT_1, Inches(0.2), Inches(0.2), width=Inches(12.933))

    prs.save(PPTX_LIGHT_1)
    shutil.copyfile(PPTX_LIGHT_1, PPTX_LIGHT_2)
    print(f"Saved Light PPTX with Software Stack to:\n  - {PPTX_LIGHT_1}")

def main():
    print("=== Adding Required Software to Network Diagram & Updating Deliverables ===")
    render_network_diagram_light()
    build_updated_pdf()
    update_pptx_decks()
    print("=== All Deliverables Updated with Required Software Successfully! ===")

if __name__ == "__main__":
    main()
