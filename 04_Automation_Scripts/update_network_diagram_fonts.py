#!/usr/bin/env python3
"""
Update Network Diagram generator with increased, high-visibility fonts.
"""
import os
import shutil
from PIL import Image, ImageDraw, ImageFont

OUT_DIR_PROJECT = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUT_DIR_WORKSPACE = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"
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

def render_network_diagram_light():
    W, H = 1920, 1080
    im = Image.new("RGB", (W, H), (248, 250, 252))
    d = ImageDraw.Draw(im)

    # Increased Font Sizes for Network Topology Diagram
    f_title = get_font(26, bold=True)      # Was 24
    f_h1 = get_font(17, bold=True)         # Was 16
    f_h2 = get_font(14, bold=True)         # Was 13
    f_sub = get_font(12, bold=True)        # Was 11
    f_body = get_font(11, bold=False)      # Was 10
    f_body_bold = get_font(11, bold=True)  # Was 10
    f_tag = get_font(10.5, bold=True)      # Was 9
    f_small = get_font(10, bold=False)     # Was 8.5
    f_small_bold = get_font(10, bold=True)

    # Top Header
    d.rectangle([40, 18, 1880, 95], fill=(255, 255, 255), outline=(203, 213, 225), width=2)
    d.text((60, 26), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT | PROJECT SPRINT (18K TPA)", font=f_h2, fill=(2, 132, 199))
    d.text((60, 52), "Industrial Network Topology & Required Software Architecture (CPwE / DLR)", font=f_title, fill=(14, 76, 146))
    d.text((1420, 30), "P&ID Rev 3.5 Aligned\nFull Software & License Stack", font=f_sub, fill=(22, 163, 74))

    # LEVEL 3.5: ENTERPRISE DMZ & INDUSTRIAL SECURITY FIREWALL
    d.rectangle([40, 102, 1880, 162], fill=(255, 255, 255), outline=(225, 29, 72), width=2)
    d.rectangle([40, 102, 1880, 128], fill=(255, 241, 242))
    d.text((60, 108), "LEVEL 3.5 : INDUSTRIAL DEMILITARIZED ZONE (IDMZ) & FIREWALL", font=f_h2, fill=(225, 29, 72))

    d.rectangle([80, 132, 500, 158], fill=(254, 226, 226), outline=(225, 29, 72), width=1)
    d.text((90, 138), "🔒 FIREWALL: Stratix 5950 / Cisco Firepower", font=f_tag, fill=(159, 18, 57))
    d.text((520, 138), "Corporate IT (10.x.x.x) <--- Deep Packet Inspection (CIP/HTTPS/SSH) ---> Plant OT (192.168.x.x)", font=f_body_bold, fill=(71, 85, 105))

    # LEVEL 3: SCADA SERVER & DATA MANAGEMENT SUBNET (VLAN 10)
    d.rectangle([40, 170, 1060, 480], fill=(255, 255, 255), outline=(2, 132, 199), width=2)
    d.rectangle([40, 170, 1060, 202], fill=(240, 249, 255))
    d.text((60, 178), "LEVEL 3 : SERVER ROOM SCADA CLUSTER (VLAN 10: 192.168.10.0/24)", font=f_h2, fill=(2, 132, 199))

    # Core Switch Stratix 5400
    d.rectangle([60, 212, 1040, 255], fill=(14, 76, 146))
    d.text((80, 223), "CORE DISTRIBUTION SWITCH (SW-CORE-01): Stratix 5400 Managed Gigabit Switch", font=f_sub, fill=(255, 255, 255))
    d.text((810, 223), "IP: 192.168.10.1 (GW)", font=f_tag, fill=(254, 240, 138))

    # Server A
    d.rectangle([60, 265, 385, 468], fill=(240, 253, 244), outline=(22, 163, 74), width=2)
    d.text((75, 273), "🖥️ SCADA SERVER A", font=f_h2, fill=(22, 163, 74))
    d.text((75, 294), "PRIMARY ACTIVE [192.168.10.11]", font=f_tag, fill=(22, 163, 74))
    d.text((75, 313), "Dell PowerEdge R550 Rackmount", font=f_small_bold, fill=(71, 85, 105))
    d.line([75, 330, 370, 330], fill=(187, 247, 208), width=1)
    d.text((75, 336), "REQUIRED SOFTWARE STACK:", font=f_tag, fill=(15, 23, 42))
    d.text((75, 353), 
        "• Windows Server 2022 Datacenter\n"
        "• FactoryTalk View SE Server v13\n"
        "  (Primary Active License)\n"
        "• FactoryTalk Historian SE\n"
        "  (5,000 Tag Base License)\n"
        "• FactoryTalk Directory Server\n"
        "• FactoryTalk Linx Gateway (OPC-UA)", font=f_small, fill=(15, 23, 42))

    # Heartbeat Link
    d.rectangle([398, 345, 488, 408], fill=(2, 132, 199))
    d.text((404, 353), "⇄ SYNC LINK\n10GbE Bonded\nHeartbeat\n< 1.0s Failover", font=f_tag, fill=(255, 255, 255))

    # Server B
    d.rectangle([500, 265, 825, 468], fill=(254, 243, 199), outline=(217, 119, 6), width=2)
    d.text((515, 273), "🖥️ SCADA SERVER B", font=f_h2, fill=(217, 119, 6))
    d.text((515, 294), "HOT STANDBY [192.168.10.12]", font=f_tag, fill=(217, 119, 6))
    d.text((515, 313), "Dell PowerEdge R550 Rackmount", font=f_small_bold, fill=(71, 85, 105))
    d.line([515, 330, 810, 330], fill=(254, 215, 170), width=1)
    d.text((515, 336), "REQUIRED SOFTWARE STACK:", font=f_tag, fill=(15, 23, 42))
    d.text((515, 353), 
        "• Windows Server 2022 Datacenter\n"
        "• FactoryTalk View SE Server v13\n"
        "  (Secondary Partner License)\n"
        "• FactoryTalk Redundancy Manager\n"
        "• FT Historian Standby Sync\n"
        "• FactoryTalk Linx Redundant OPC", font=f_small, fill=(15, 23, 42))

    # NAS / Domain Controller
    d.rectangle([838, 265, 1040, 468], fill=(241, 245, 249), outline=(148, 163, 184), width=2)
    d.text((848, 273), "💾 NAS / DC", font=f_sub, fill=(14, 76, 146))
    d.text((848, 294), "[192.168.10.15]", font=f_tag, fill=(71, 85, 105))
    d.line([848, 312, 1030, 312], fill=(203, 213, 225), width=1)
    d.text((848, 320), "REQUIRED SOFTWARE:", font=f_tag, fill=(15, 23, 42))
    d.text((848, 338), 
        "• Windows Server 2022\n"
        "• Active Directory (AD DS)\n"
        "• FactoryTalk AssetCentre\n"
        "  Server (v12 Backup & Audit)\n"
        "• Windows DNS & NTP Srv\n"
        "• Veeam Backup Agent", font=f_small, fill=(15, 23, 42))

    # LEVEL 2: MAIN CONTROL ROOM (MCR) CLIENT STATIONS (VLAN 10)
    d.rectangle([1075, 170, 1880, 480], fill=(255, 255, 255), outline=(14, 76, 146), width=2)
    d.rectangle([1075, 170, 1880, 202], fill=(241, 245, 249))
    d.text((1095, 178), "LEVEL 2 : MAIN CONTROL ROOM (MCR) CLIENTS (VLAN 10: 192.168.10.0/24)", font=f_h2, fill=(14, 76, 146))

    # MCR Switch
    d.rectangle([1095, 212, 1860, 255], fill=(14, 76, 146))
    d.text((1115, 223), "MCR DISTRIBUTION SWITCH (SW-MCR-01): Stratix 5700 Managed Switch | IP: 192.168.10.2", font=f_sub, fill=(255, 255, 255))

    # Clients Grid
    mcr_nodes = [
        ("💻 1x EWS", "IP: 192.168.10.20", "Dell Precision 3660", [
            "• Windows 11 Pro Workstation",
            "• Studio 5000 Logix (v33+)",
            "• Studio 5000 Safety Editor",
            "• FT View Studio Enterprise",
            "• FT Network Manager",
            "• FT AssetCentre Client"
        ], 1095, (147, 51, 234), (250, 245, 255)),

        ("🖥️ OWS-01", "IP: 192.168.10.21", "Slurry & Jet Cooker", [
            "• Windows 11 LTSC",
            "• FT View SE Client v13",
            "• FT Alarm & Event Client",
            "• FT View TrendPro",
            "• FT Desktop Lock Shell",
            "• Auto-Switch to Srv B"
        ], 1290, (2, 132, 199), (240, 249, 255)),

        ("🖥️ OWS-02", "IP: 192.168.10.22", "Dryer & Safety BMS", [
            "• Windows 11 LTSC",
            "• FT View SE Client v13",
            "• FT Alarm & Event Client",
            "• FT View TrendPro",
            "• FT Desktop Lock Shell",
            "• Auto-Switch to Srv B"
        ], 1485, (2, 132, 199), (240, 249, 255)),

        ("🖥️ OWS-03", "IP: 192.168.10.23", "Silo & Utilities", [
            "• Windows 11 LTSC",
            "• FT View SE Client v13",
            "• FT Alarm & Event Client",
            "• FT View TrendPro",
            "• FT Desktop Lock Shell",
            "• Auto-Switch to Srv B"
        ], 1680, (2, 132, 199), (240, 249, 255))
    ]

    for ctitle, cip, chw, csofts, cx, ccol, cbg in mcr_nodes:
        d.rectangle([cx, 265, cx + 180, 468], fill=cbg, outline=ccol, width=2)
        d.text((cx + 8, 273), ctitle, font=f_h2, fill=ccol)
        d.text((cx + 8, 294), cip, font=f_tag, fill=(71, 85, 105))
        d.text((cx + 8, 312), chw, font=f_small_bold, fill=(100, 116, 139))
        d.line([cx + 8, 330, cx + 172, 330], fill=ccol, width=1)
        d.text((cx + 8, 336), "REQUIRED SOFTWARE:", font=f_tag, fill=(15, 23, 42))
        ly = 353
        for s in csofts:
            d.text((cx + 8, ly), s, font=f_small, fill=(15, 23, 42))
            ly += 18

    # LEVEL 1 & 0: CONTROLLER & DEVICE LEVEL RING (DLR) (VLAN 30: 192.168.30.0/24)
    d.rectangle([40, 492, 1880, 875], fill=(255, 255, 255), outline=(22, 163, 74), width=2)
    d.rectangle([40, 492, 1880, 524], fill=(240, 253, 244))
    d.text((60, 500), "LEVEL 1 & LEVEL 0 : CONTROLLER REDUNDANCY & FAULT-TOLERANT DEVICE LEVEL RING (DLR)", font=f_h2, fill=(22, 163, 74))
    d.text((1200, 502), "Sub-3ms Self-Healing Recovery | Zero Single Point of Failure (SPOF)", font=f_sub, fill=(14, 76, 146))

    # Redundant ControlLogix Chassis A & B
    # Chassis A
    d.rectangle([60, 538, 490, 680], fill=(240, 253, 244), outline=(22, 163, 74), width=2)
    d.text((75, 548), "🎛️ CONTROLLER CHASSIS A (PRIMARY)", font=f_h2, fill=(22, 163, 74))
    d.text((75, 570), "Slot 0: 1756-L83E GuardLogix 5580 CPU [IP: 192.168.20.11]\nSlot 1: 1756-RM2 High-Speed Optical Redundancy Module\nSlot 2: 1756-EN4TR Dual EtherNet/IP Bridge (DLR Supervisor A)\nSlot 3: 1756-EN2T Uplink to Server VLAN 10 [192.168.10.31]", font=f_body, fill=(15, 23, 42))
    d.text((75, 650), "Firmware: v33.011+ Redundant Bundle | Bumpless Transfer < 20 ms", font=f_small_bold, fill=(22, 163, 74))

    # Optical Redundancy Sync
    d.rectangle([510, 580, 640, 640], fill=(22, 163, 74))
    d.text((518, 592), "⇄ 1756-RM2\nOptical Fiber\nCross-Sync\n< 20 ms", font=f_tag, fill=(255, 255, 255))

    # Chassis B
    d.rectangle([660, 538, 1090, 680], fill=(254, 243, 199), outline=(217, 119, 6), width=2)
    d.text((675, 548), "🎛️ CONTROLLER CHASSIS B (SECONDARY)", font=f_h2, fill=(217, 119, 6))
    d.text((675, 570), "Slot 0: 1756-L83E GuardLogix 5580 CPU [IP: 192.168.20.12]\nSlot 1: 1756-RM2 High-Speed Optical Redundancy Module\nSlot 2: 1756-EN4TR Dual EtherNet/IP Bridge (DLR Supervisor B)\nSlot 3: 1756-EN2T Uplink to Server VLAN 10 [192.168.10.32]", font=f_body, fill=(15, 23, 42))
    d.text((675, 650), "Hot Standby Synchronized Execution | Zero Process Perturbation", font=f_small_bold, fill=(217, 119, 6))

    # DLR Explanation Card
    d.rectangle([1110, 538, 1860, 680], fill=(241, 245, 249), outline=(148, 163, 184), width=2)
    d.text((1125, 548), "DEVICE LEVEL RING (DLR) ARCHITECTURE HIGHLIGHTS (VLAN 30):", font=f_h2, fill=(14, 76, 146))
    d.text((1125, 572), 
        "• Self-Healing Ring Topology: Embedded hardware switching enables < 3 ms recovery on cable break.\n"
        "• Dual Ring Supervisors: 1756-EN4TR on Chassis A acts as Active Supervisor; Chassis B is Backup.\n"
        "• Zero Packet Loss: Remote I/O drops RIO-200, 400, 600 maintain continuous communication.\n"
        "• Direct Skid Integration: BMS Burner, Getabec Boiler, and IEP Systems interconnected seamlessly.", font=f_body, fill=(30, 41, 59))

    # DLR Ring Nodes (Bottom)
    dlr_nodes = [
        ("📡 RIO-200 (Slurry)", "IP: 192.168.30.11\n1794-AENTR Flex w/ HART\nDTM Drivers", 60, (2, 132, 199)),
        ("📡 RIO-400 (Infeed)", "IP: 192.168.30.12\n1794-AENTR Flex w/ HART\nDTM Drivers", 325, (2, 132, 199)),
        ("📡 RIO-600 (Dryer)", "IP: 192.168.30.13\n1794-AENTR Flex w/ HART\nDTM Drivers", 590, (2, 132, 199)),
        ("⚡ RIO-MCC (VFDs)", "IP: 192.168.30.14\nPowerFlex 755/525 VFDs\nConnected Components", 855, (217, 119, 6)),
        ("🔥 BMS SKID", "IP: 192.168.30.21\nBurner Safety Gateway\nModbus-TCP / ENet", 1120, (225, 29, 72)),
        ("🏭 BOILER SKID", "IP: 192.168.30.22\nGetabec Boiler Gateway\nModbus-TCP / ENet", 1385, (14, 76, 146)),
        ("💥 IEP SKID", "IP: 192.168.30.23\nExplosion Suppression\nHRD Safety Profile", 1650, (147, 51, 234))
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
        d.text((nx + 10, 762), ndesc, font=f_body, fill=(15, 23, 42))

    # BOTTOM SOFTWARE BOM & LICENSE MATRIX TABLE (Larger text)
    d.rectangle([40, 890, 1880, 1060], fill=(241, 245, 249), outline=(203, 213, 225), width=2)
    d.text((60, 898), "SUMMARY AUTOMATION SOFTWARE & LICENSE BILL OF MATERIALS (BOM) — REFER TO PAGE 7 FOR COMPLETE DETAILS:", font=f_h2, fill=(14, 76, 146))

    soft_rows = [
        ("SCADA Servers (Primary & Secondary)", "FactoryTalk View SE Server v13.0 (Redundant Pair)", "9358-VWSE000LENE", "2 Lic", "Active-Standby Redundant SCADA Server Engine (< 1.0s Failover)"),
        ("Industrial Process Historian", "FactoryTalk Historian ME / SE (Standard Edition)", "9518-HISTMESE5K", "1 Srv + 5K Tags", "High-speed process archive & trending (100ms sample rate)"),
        ("Engineering Workstation (EWS)", "Studio 5000 Professional & FT View Studio Enterprise", "9324-RLD700NXENE", "1 Named Lic", "Full PLC programming, safety logic & HMI graphic screen development"),
        ("Operator Workstations (OWS-01, 02, 03)", "FactoryTalk View SE Client (Desktop Shell Mode)", "9358-VWC100LENE", "3 Client Lic", "Dedicated operator stations with locked OS desktop kiosk mode"),
        ("Disaster Recovery & Version Control", "FactoryTalk AssetCentre Server & Client Package", "9515-ASTCNTR01", "1 Srv Lic", "Automated project backup, change tracking, audit logs & recovery")
    ]

    sy = 926
    for s_role, s_pkg, s_cat, s_qty, s_desc in soft_rows:
        d.text((60, sy), s_role, font=f_body_bold, fill=(14, 76, 146))
        d.text((430, sy), s_pkg, font=f_body, fill=(15, 23, 42))
        d.text((880, sy), s_cat, font=f_tag, fill=(2, 132, 199))
        d.text((1065, sy), s_qty, font=f_body_bold, fill=(22, 163, 74))
        d.text((1220, sy), s_desc, font=f_body, fill=(51, 65, 85))
        sy += 24

    im.save(NET_PNG_LIGHT_1, quality=95)
    shutil.copyfile(NET_PNG_LIGHT_1, NET_PNG_LIGHT_2)
    print("Rendered Network Diagram PNG Light with Increased Font Sizes successfully.")
    return im

if __name__ == "__main__":
    render_network_diagram_light()
