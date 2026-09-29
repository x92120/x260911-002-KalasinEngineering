#!/usr/bin/env python3
"""
Render Page 7: Comprehensive Software & License Bill of Materials (BOM)
With Significantly Larger, Crisp, and Highly Legible Fonts.
"""
import os
import shutil
from PIL import Image, ImageDraw, ImageFont

OUT_DIR_PROJECT = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUT_DIR_WORKSPACE = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"

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

def render_software_bom_slide_light():
    W, H = 1920, 1080
    im = Image.new("RGB", (W, H), (248, 250, 252))
    d = ImageDraw.Draw(im)

    # Increased Font Sizes for high legibility
    f_title = get_font(28, bold=True)      # Was 24 -> Now 28
    f_header_sub = get_font(16, bold=True) # Was 13 -> Now 16
    f_kpi_title = get_font(15, bold=True)  # Was 11 -> Now 15
    f_kpi_desc = get_font(13, bold=False)  # Was 8.5 -> Now 13
    f_th = get_font(14, bold=True)         # Was 11 -> Now 14
    f_tb_bold = get_font(13, bold=True)    # Was 10 -> Now 13
    f_tb = get_font(12.5, bold=False)      # Was 8.5 -> Now 12.5
    f_part = get_font(13, bold=True)       # Was 9 -> Now 13
    f_foot_title = get_font(14, bold=True) # Was 11 -> Now 14
    f_foot_body = get_font(12, bold=False) # Was 8.5 -> Now 12

    # Top Header
    d.rectangle([40, 20, 1880, 100], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d.text((60, 28), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT | PROJECT SPRINT", font=f_header_sub, fill=(2, 132, 199))
    d.text((60, 56), "Complete Automation Software & License Bill of Materials (BOM) — Page 7", font=f_title, fill=(14, 76, 146))
    d.text((1400, 32), "Rockwell PlantPAx 5.0 Baseline\nProcurement & License Stack", font=f_header_sub, fill=(22, 163, 74))

    # Summary KPI Cards (Increased size and fonts)
    kpis = [
        ("SCADA Server Redundancy", "FactoryTalk View SE v13.0\n1 Primary + 1 Secondary Pair", (2, 132, 199), (240, 249, 255)),
        ("Engineering License", "Studio 5000 Pro + FT Studio\n1 Complete EWS Suite", (147, 51, 234), (250, 245, 255)),
        ("Operator Workstations", "FT View SE Client Runtime\n3 Dedicated OWS Stations", (14, 76, 146), (241, 245, 249)),
        ("Process Historian", "FT Historian SE Server\n5,000 Tag Scalable Archive", (22, 163, 74), (240, 253, 244)),
        ("Disaster Recovery", "FT AssetCentre Server v12\nAutomated Backup & Audit", (217, 119, 6), (254, 243, 199))
    ]
    kx = 40
    kw = 356
    for k_title, k_desc, k_col, k_bg in kpis:
        d.rectangle([kx, 112, kx + kw, 188], fill=k_bg, outline=k_col, width=2)
        d.text((kx + 16, 120), k_title, font=f_kpi_title, fill=k_col)
        d.text((kx + 16, 144), k_desc, font=f_kpi_desc, fill=(15, 23, 42))
        kx += kw + 25

    # Main BOM Table
    tx, ty = 40, 200
    tw, th = 1840, 775
    d.rectangle([tx, ty, tx + tw, ty + th], fill=(255, 255, 255), outline=(203, 213, 225), width=2)

    # Table Header Row (Height 44px)
    d.rectangle([tx, ty, tx + tw, ty + 44], fill=(14, 76, 146))
    headers = [
        ("NO", 52, 40),
        ("SYSTEM NODE", 95, 225),
        ("SOFTWARE PRODUCT / MODULE", 325, 395),
        ("CATALOG / PART #", 725, 220),
        ("QTY", 950, 55),
        ("LICENSE TYPE", 1010, 225),
        ("FUNCTIONAL DESCRIPTION & PURPOSE", 1240, 630)
    ]
    for h_name, hx, _ in headers:
        d.text((hx, ty + 12), h_name, font=f_th, fill=(255, 255, 255))

    bom_items = [
        ("01", "SCADA Server A (Primary)", "FactoryTalk View Site Edition Server (Unlimited Displays)", "9358-VWSE000LENE", "1", "Perpetual / Activation", "Primary active HMI data & graphics server, tag engine, alarming"),
        ("02", "SCADA Server B (Standby)", "FactoryTalk View SE Redundant Server Partner License", "9358-VWSEREDLENE", "1", "Perpetual / Activation", "Hot-standby redundant server sync with bumpless client failover < 1s"),
        ("03", "Data Historian Node", "FactoryTalk Historian SE Standard Edition (5,000 Tags)", "9518-HISTMESE5K", "1", "Server Base + 5K Tags", "Enterprise PI-based long-term process archiving, 100ms sample rate"),
        ("04", "OPC-UA Communications", "FactoryTalk Linx Gateway / Data Server (DataBridge Pro)", "9355-WABGWENE", "2", "Server License (2 Nodes)", "High-performance CIP/OPC-UA server communicating with Logix CPUs"),
        ("05", "Engineering Station (EWS)", "Studio 5000 Logix Designer Professional Edition (v33+)", "9324-RLD700NXENE", "1", "Named User Perpetual", "Full ladder, function block, structured text, SFC, and SIL2/3 logic"),
        ("06", "Engineering Station (EWS)", "FactoryTalk View Studio Enterprise Edition", "9268-FTSPENE", "1", "Named User Perpetual", "Centralized HMI graphic design, global objects, faceplates, security"),
        ("07", "Engineering Station (EWS)", "Studio 5000 Safety Editor & GuardLogix Safety Architect", "9324-RLDSFEENE", "1", "Named User Add-On", "Safety routine verification, checksum validation & SIL signature lock"),
        ("08", "Operator Stations (OWS)", "FactoryTalk View SE Client Runtime License (Desktop)", "9358-VWC100LENE", "3", "3 Device Runtime Lic", "Dedicated operator client for OWS-01, OWS-02, OWS-03 (dual display)"),
        ("09", "Operator Stations (OWS)", "FactoryTalk Desktop Lock Utility", "9358-FTLOCK", "3", "3 Station Utility", "Locks Windows OS into dedicated HMI kiosk, preventing unauthorized OS access"),
        ("10", "Network & Diagnostics", "FactoryTalk Network Manager (Managed Switch Visibility)", "9528-FTNM01", "1", "1 Server License", "Monitors Stratix 5400/5700 switches, topology mapping & cable diagnostics"),
        ("11", "Audit & Disaster Recovery", "FactoryTalk AssetCentre Server & Client Package", "9515-ASTCNTR01", "1", "Server + 5 Device Client", "Automatic PLC program backup, change tracking, audit logs & disaster recovery"),
        ("12", "Server Operating Systems", "Microsoft Windows Server 2022 Datacenter Edition (16-Core)", "MS-WS22-DC-16C", "2", "OEM / Volume Lic", "Hyper-V clustering, Active Directory secondary, Server A & B OS"),
        ("13", "Client Operating Systems", "Microsoft Windows 11 Pro / Enterprise LTSC for Workstations", "MS-W11-ENT-LTSC", "4", "OEM Digital License", "Ultra-stable enterprise OS for 1 EWS and 3 OWS client stations"),
        ("14", "Backup & Virtualization", "Veeam Backup & Replication Enterprise Edition", "V-VBR-ENT-01", "1", "Host Socket / VM Lic", "Full bare-metal backup & instant VM restore for SCADA servers and NAS")
    ]

    ry = ty + 44
    row_h = 52.2
    for idx, item in enumerate(bom_items):
        r_bg = (248, 250, 252) if idx % 2 == 1 else (255, 255, 255)
        d.rectangle([tx, int(ry), tx + tw, int(ry + row_h)], fill=r_bg)
        d.line([tx, int(ry + row_h), tx + tw, int(ry + row_h)], fill=(226, 232, 240), width=1)

        ino, node, prod, part, qty, ltype, desc = item
        y_text = int(ry + 15)
        d.text((52, y_text), ino, font=f_tb_bold, fill=(71, 85, 105))
        d.text((95, y_text), node, font=f_tb_bold, fill=(14, 76, 146))
        d.text((325, y_text), prod, font=f_tb_bold, fill=(15, 23, 42))
        d.text((725, y_text), part, font=f_part, fill=(2, 132, 199))
        d.text((958, y_text), qty, font=f_tb_bold, fill=(22, 163, 74))
        d.text((1010, y_text), ltype, font=f_tb, fill=(71, 85, 105))
        d.text((1240, y_text), desc, font=f_tb, fill=(51, 65, 85))

        ry += row_h

    # Bottom Notes & Compliance Bar (Larger text & padding)
    d.rectangle([40, 988, 1880, 1065], fill=(241, 245, 249), outline=(203, 213, 225), width=2)
    d.text((60, 996), "LICENSING NOTES & LIFECYCLE COMPLIANCE:", font=f_foot_title, fill=(14, 76, 146))
    d.text((60, 1020), 
        "• All Rockwell Automation software licenses are managed via FactoryTalk Activation Manager (FTAM) hosted on Primary Server A (with Server B secondary sync).\n"
        "• Software stack includes 1-year TechConnect Comprehensive Support 24x7. Fully compliant with PlantPAx 5.0 process library & ISA-88 batch standards.",
        font=f_foot_body, fill=(51, 65, 85))
    d.text((1440, 1005), "APPROVED FOR PROCUREMENT\nAEC Engineering / Ingredion Kalasin", font=f_header_sub, fill=(22, 163, 74))

    bom_png = os.path.join(OUT_DIR_PROJECT, "Software_License_BOM_Slide_Light.png")
    im.save(bom_png, quality=95)
    shutil.copyfile(bom_png, os.path.join(OUT_DIR_WORKSPACE, "Software_License_BOM_Slide_Light.png"))
    print(f"Rendered High-Legibility Software BOM Slide PNG to {bom_png}")
    return im

if __name__ == "__main__":
    render_software_bom_slide_light()
