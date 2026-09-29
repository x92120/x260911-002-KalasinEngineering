#!/usr/bin/env python3
"""
Generate Complete Industrial CAD DXF & SVG Drawings for:
SPRINT 18K Panel Layout — System Architecture
Client: Ingredion_001 | Engineering: Ecomax | Drawing No: 600-600-001
Features:
- EXISTING SERVER ROOM (Server Rack, Console, Fiber Uplink)
- OPERATION 1 ROOM (Network Cabinet, EWS, Op1, Op2, 2x 55" Displays, Printer)
- OPERATION STATION 2 PANEL (Remote Operator Station)
- OEM PACKAGES (Dedert Burner, B+B Packer, Palletizer, Rework, Boiler)
- MAIN CONTROL PANEL (MCP-01: Racks 1-4 with 1756-L950TPSXT, DLR Ring Supervisor, Redundant EN4TR, I/O)
- RIO AREA-200 (Rack 5 Remote I/O via Fiber Optic DLR)
- MCC PANEL (Incoming Breaker, VFDs, Busbars, Plant Area Feeds)
- COMPLETE CABLE NETWORKS (Control LAN - Red, I/O DLR - Green, MCC - Blue, Fiber - Pink/Orange)
- OFFICIAL TITLE BLOCK & LEGEND matching sample drawing S__103251990.jpg
"""

import os
import ezdxf
from ezdxf import colors
from ezdxf.enums import TextEntityAlignment
import fitz

# Coordinate Extents (Standard A1 Landscape: 2400 x 1700 CAD Units)
DWG_W = 2400
DWG_H = 1700

def create_dxf():
    doc = ezdxf.new('R2010', setup=True)
    msp = doc.modelspace()

    # Define CAD Layers & ACI Colors
    # 1=Red, 2=Yellow, 3=Green, 4=Cyan, 5=Blue, 6=Magenta, 7=White, 8=DarkGray, 9=LightGray
    layers = [
        ('TITLE_BLOCK', 7, 'Continuous'),
        ('BORDER_GRID', 8, 'Continuous'),
        ('ROOM_BOUNDARY', 4, 'DASHED'),
        ('PANEL_ENCLOSURE', 7, 'Continuous'),
        ('EQUIPMENT_FRAME', 8, 'Continuous'),
        ('EQUIPMENT_DETAIL', 7, 'Continuous'),
        ('PLC_CHASSIS', 5, 'Continuous'),
        ('PLC_MODULE', 4, 'Continuous'),
        ('CBL_CONTROL_LAN', 1, 'Continuous'),     # Red: Ethernet LAN for Control
        ('CBL_IO_NETWORK', 3, 'Continuous'),      # Green: Ethernet for I/O (DLR)
        ('CBL_MCC_NETWORK', 5, 'Continuous'),     # Blue: Ethernet for MCC
        ('CBL_FIBER_OPTIC', 6, 'DASHED'),         # Magenta/Pink: Fiber Optic
        ('TEXT_TITLE', 7, 'Continuous'),
        ('TEXT_SUBTITLE', 4, 'Continuous'),
        ('TEXT_EQUIPMENT', 7, 'Continuous'),
        ('TEXT_TAG', 3, 'Continuous'),
        ('TEXT_ANNOTATION', 2, 'Continuous'),
        ('LEGEND', 7, 'Continuous'),
    ]

    for l_name, l_col, l_type in layers:
        doc.layers.add(name=l_name, color=l_col, linetype=l_type)

    # ----------------------------------------------------
    # Helper Functions
    # ----------------------------------------------------
    def rect(x, y, w, h, layer='PANEL_ENCLOSURE', linetype='ByLayer', color=None):
        pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h), (x, y)]
        attr = {'layer': layer, 'linetype': linetype}
        if color is not None:
            attr['color'] = color
        return msp.add_lwpolyline(pts, dxfattribs=attr)

    def line(p1, p2, layer='EQUIPMENT_DETAIL', linetype='ByLayer', color=None):
        attr = {'layer': layer, 'linetype': linetype}
        if color is not None:
            attr['color'] = color
        return msp.add_line(p1, p2, dxfattribs=attr)

    def polyline(pts, layer='EQUIPMENT_DETAIL', linetype='ByLayer', color=None):
        attr = {'layer': layer, 'linetype': linetype}
        if color is not None:
            attr['color'] = color
        return msp.add_lwpolyline(pts, dxfattribs=attr)

    def text(txt, x, y, h=10, layer='TEXT_EQUIPMENT', color=None, align='LEFT'):
        attr = {'layer': layer, 'height': h}
        if color is not None:
            attr['color'] = color
        t = msp.add_text(txt, dxfattribs=attr)
        if align == 'CENTER':
            t.set_placement((x, y), align=TextEntityAlignment.CENTER)
        elif align == 'RIGHT':
            t.set_placement((x, y), align=TextEntityAlignment.RIGHT)
        else:
            t.set_placement((x, y), align=TextEntityAlignment.LEFT)
        return t

    # ----------------------------------------------------
    # 1. DRAWING BORDER & REFERENCE GRID
    # ----------------------------------------------------
    # Outer sheet border
    rect(20, 20, DWG_W - 40, DWG_H - 40, layer='TITLE_BLOCK', color=7)
    # Inner drawing margin
    rect(35, 35, DWG_W - 70, DWG_H - 70, layer='TITLE_BLOCK', color=7)

    # Reference grid divisions along top/bottom (1 to 9) and left/right (A to F)
    x_steps = [35 + i * ((DWG_W - 70) / 9) for i in range(10)]
    for i in range(1, 10):
        gx = (x_steps[i - 1] + x_steps[i]) / 2
        text(str(i), gx, DWG_H - 30, h=8, layer='BORDER_GRID', align='CENTER')
        text(str(i), gx, 25, h=8, layer='BORDER_GRID', align='CENTER')

    y_steps = [35 + i * ((DWG_H - 70) / 6) for i in range(7)]
    y_labels = ['F', 'E', 'D', 'C', 'B', 'A']
    for i in range(6):
        gy = (y_steps[i] + y_steps[i + 1]) / 2
        text(y_labels[i], 27, gy, h=8, layer='BORDER_GRID', align='CENTER')
        text(y_labels[i], DWG_W - 27, gy, h=8, layer='BORDER_GRID', align='CENTER')

    # ----------------------------------------------------
    # 2. TITLE BLOCK (Matching S__103251990.jpg)
    # ----------------------------------------------------
    tb_x = 35
    tb_y = 35
    tb_w = DWG_W - 70
    tb_h = 75

    rect(tb_x, tb_y, tb_w, tb_h, layer='TITLE_BLOCK')
    # Vertical dividers in title block
    line((tb_x + 400, tb_y), (tb_x + 400, tb_y + tb_h), layer='TITLE_BLOCK')
    line((tb_x + 900, tb_y), (tb_x + 900, tb_y + tb_h), layer='TITLE_BLOCK')
    line((tb_x + 1200, tb_y), (tb_x + 1200, tb_y + tb_h), layer='TITLE_BLOCK')
    line((tb_x + 1750, tb_y), (tb_x + 1750, tb_y + tb_h), layer='TITLE_BLOCK')

    # Revisions sub-table (Column 1)
    line((tb_x, tb_y + 50), (tb_x + 400, tb_y + 50), layer='TITLE_BLOCK')
    line((tb_x, tb_y + 25), (tb_x + 400, tb_y + 25), layer='TITLE_BLOCK')
    line((tb_x + 100, tb_y), (tb_x + 100, tb_y + 50), layer='TITLE_BLOCK')
    line((tb_x + 200, tb_y), (tb_x + 200, tb_y + 50), layer='TITLE_BLOCK')
    line((tb_x + 300, tb_y), (tb_x + 300, tb_y + 50), layer='TITLE_BLOCK')

    text('Name: THUMWH', tb_x + 10, tb_y + 58, h=7, layer='TEXT_TITLE')
    text('Date: 09/10/26', tb_x + 110, tb_y + 58, h=7, layer='TEXT_TITLE')
    text('Ed.: THUMWH', tb_x + 210, tb_y + 58, h=7, layer='TEXT_TITLE')
    text('Appr.: Original', tb_x + 310, tb_y + 58, h=7, layer='TEXT_TITLE')
    text('Replacement of:', tb_x + 10, tb_y + 35, h=6, layer='TEXT_TITLE')
    text('Replaced by:', tb_x + 210, tb_y + 35, h=6, layer='TEXT_TITLE')
    text('THUM', tb_x + 10, tb_y + 10, h=10, layer='TEXT_TITLE')

    # Project Title (Column 2)
    text('SPRINT 18K Panel Layout', tb_x + 420, tb_y + 45, h=14, layer='TEXT_TITLE')
    text('Rockwell ControlLogix 5580 & DLR Architecture', tb_x + 420, tb_y + 20, h=9, layer='TEXT_SUBTITLE')

    # Company Name (Column 3)
    text('Ecomax', tb_x + 920, tb_y + 45, h=14, layer='TEXT_TITLE')
    text('Engineering & Automation', tb_x + 920, tb_y + 20, h=8, layer='TEXT_TITLE')

    # System Architecture (Column 4)
    text('System architecture', tb_x + 1220, tb_y + 45, h=14, layer='TEXT_TITLE')
    text('Plant-wide Automation Network', tb_x + 1220, tb_y + 20, h=8, layer='TEXT_TITLE')

    # Client & Drawing Number (Column 5)
    text('Client: Ingredion_001', tb_x + 1770, tb_y + 50, h=9, layer='TEXT_TITLE')
    text('Drawing No: 600-600-001', tb_x + 1770, tb_y + 30, h=11, layer='TEXT_TITLE', color=3)
    text('Rev: B01 | Sheet: 1 of 1', tb_x + 1770, tb_y + 12, h=7, layer='TEXT_TITLE')

    # ----------------------------------------------------
    # 3. LEGEND BOX (Lower Right above Title Block)
    # ----------------------------------------------------
    lg_x = 1820
    lg_y = 120
    lg_w = 515
    lg_h = 100

    rect(lg_x, lg_y, lg_w, lg_h, layer='LEGEND')
    text('CABLE LEGEND & NETWORK SPECIFICATION', lg_x + 15, lg_y + 82, h=8, layer='LEGEND', color=7)
    
    # Red: Ethernet LAN Cable for Control
    line((lg_x + 15, lg_y + 65), (lg_x + 75, lg_y + 65), layer='CBL_CONTROL_LAN', color=1)
    text('ETHERNET LAN CABLE FOR CONTROL (Cat6A)', lg_x + 90, lg_y + 62, h=7, layer='LEGEND', color=1)

    # Green: Ethernet for I/O Network (DLR)
    line((lg_x + 15, lg_y + 45), (lg_x + 75, lg_y + 45), layer='CBL_IO_NETWORK', color=3)
    text('ETHERNET FOR I/O NETWORK (DLR RING)', lg_x + 90, lg_y + 42, h=7, layer='LEGEND', color=3)

    # Blue: Ethernet for MCC Network
    line((lg_x + 15, lg_y + 25), (lg_x + 75, lg_y + 25), layer='CBL_MCC_NETWORK', color=5)
    text('ETHERNET FOR MCC NETWORK (PowerFlex VFDs)', lg_x + 90, lg_y + 22, h=7, layer='LEGEND', color=5)

    # Magenta/Pink: Fiber Optic Cable
    line((lg_x + 15, lg_y + 10), (lg_x + 75, lg_y + 10), layer='CBL_FIBER_OPTIC', linetype='DASHED', color=6)
    text('FIBER OPTIC BACKBONE (OM3/OS2 Armored)', lg_x + 90, lg_y + 7, h=7, layer='LEGEND', color=6)

    # ----------------------------------------------------
    # 4. SECTION: EXISTING SERVER ROOM (Top Left)
    # ----------------------------------------------------
    sr_x = 55
    sr_y = 1060
    sr_w = 420
    sr_h = 550

    rect(sr_x, sr_y, sr_w, sr_h, layer='ROOM_BOUNDARY', linetype='DASHED')
    text('EXISTING SERVER ROOM', sr_x + 20, sr_y + sr_h - 25, h=11, layer='TEXT_TITLE')

    # 19" Server Rack Cabinet (42U)
    rack_x = sr_x + 50
    rack_y = sr_y + 50
    rack_w = 170
    rack_h = 420
    rect(rack_x, rack_y, rack_w, rack_h, layer='EQUIPMENT_FRAME')
    rect(rack_x + 10, rack_y + 10, rack_w - 20, rack_h - 20, layer='EQUIPMENT_FRAME')
    text('SERVER RACK (42U)', rack_x + rack_w / 2, rack_y + rack_h - 22, h=8, layer='TEXT_EQUIPMENT', align='CENTER')

    # Draw rack shelves/servers
    for idx, s_name in enumerate(['Patch Panel 1 & 2', 'Core Switch (Stratix)', 'FT View SE Server', 'FT Historian Server', 'Domain / Batch Server', 'Online UPS Unit (5kVA)']):
        sy = rack_y + rack_h - 60 - idx * 55
        rect(rack_x + 15, sy, rack_w - 30, 42, layer='EQUIPMENT_DETAIL')
        text(s_name, rack_x + rack_w / 2, sy + 18, h=7, layer='TEXT_EQUIPMENT', align='CENTER')

    # Console Monitor & Terminal
    c_x = sr_x + 260
    c_y = sr_y + 180
    # Monitor
    rect(c_x, c_y + 40, 80, 60, layer='EQUIPMENT_DETAIL')
    rect(c_x + 5, c_y + 45, 70, 50, layer='EQUIPMENT_DETAIL')
    line((c_x + 35, c_y + 40), (c_x + 35, c_y + 25), layer='EQUIPMENT_DETAIL')
    line((c_x + 25, c_y + 25), (c_x + 55, c_y + 25), layer='EQUIPMENT_DETAIL')
    # Keyboard & PC
    rect(c_x + 15, c_y + 5, 50, 15, layer='EQUIPMENT_DETAIL')
    rect(c_x + 95, c_y + 5, 30, 95, layer='EQUIPMENT_DETAIL')
    text('SERVER CONSOLE', c_x + 40, c_y + 115, h=7, layer='TEXT_EQUIPMENT', align='CENTER')

    # ----------------------------------------------------
    # 5. SECTION: OPERATION 1 ROOM (Top Center)
    # ----------------------------------------------------
    op_x = 510
    op_y = 1060
    op_w = 1180
    op_h = 550

    rect(op_x, op_y, op_w, op_h, layer='ROOM_BOUNDARY', linetype='DASHED')
    text('OPERATION 1 ROOM (MAIN CONTROL ROOM)', op_x + 20, op_y + op_h - 25, h=11, layer='TEXT_TITLE')

    # Network / Server Cabinet in Operation Room
    net_x = op_x + 40
    net_y = op_y + 50
    net_w = 190
    net_h = 420
    rect(net_x, net_y, net_w, net_h, layer='EQUIPMENT_FRAME')
    rect(net_x + 10, net_y + 10, net_w - 20, net_h - 20, layer='EQUIPMENT_FRAME')
    text('NETWORK CABINET', net_x + net_w / 2, net_y + net_h - 22, h=8, layer='TEXT_EQUIPMENT', align='CENTER')

    # Network switches & Patch panel
    rect(net_x + 15, net_y + 300, net_w - 30, 40, layer='EQUIPMENT_DETAIL')
    text('Switch 1: 192.168.10.x', net_x + net_w / 2, net_y + 316, h=7, layer='TEXT_IP_TAG', color=4, align='CENTER')

    rect(net_x + 15, net_y + 240, net_w - 30, 40, layer='EQUIPMENT_DETAIL')
    text('Switch 2: 192.168.20.x', net_x + net_w / 2, net_y + 256, h=7, layer='TEXT_IP_TAG', color=4, align='CENTER')

    rect(net_x + 15, net_y + 180, net_w - 30, 40, layer='EQUIPMENT_DETAIL')
    text('Cat6A Patch Panel 48P', net_x + net_w / 2, net_y + 196, h=7, layer='TEXT_EQUIPMENT', align='CENTER')

    rect(net_x + 15, net_y + 120, net_w - 30, 40, layer='EQUIPMENT_DETAIL')
    text('Fiber Patch Panel (FDF)', net_x + net_w / 2, net_y + 136, h=7, layer='TEXT_EQUIPMENT', color=6, align='CENTER')

    rect(net_x + 15, net_y + 50, net_w - 30, 50, layer='EQUIPMENT_DETAIL')
    text('Control Room UPS 3kVA', net_x + net_w / 2, net_y + 72, h=7, layer='TEXT_EQUIPMENT', align='CENTER')

    # Helper to draw Workstation Desk + PC + Monitor
    def draw_workstation(wx, wy, title_str, ip_str):
        # Desk
        rect(wx - 10, wy - 5, 170, 110, layer='EQUIPMENT_FRAME', color=8)
        # Monitor
        rect(wx + 20, wy + 35, 90, 60, layer='EQUIPMENT_DETAIL')
        rect(wx + 25, wy + 40, 80, 50, layer='EQUIPMENT_DETAIL')
        line((wx + 65, wy + 35), (wx + 65, wy + 20), layer='EQUIPMENT_DETAIL')
        line((wx + 50, wy + 20), (wx + 80, wy + 20), layer='EQUIPMENT_DETAIL')
        # Keyboard & Mouse
        rect(wx + 30, wy + 5, 55, 12, layer='EQUIPMENT_DETAIL')
        rect(wx + 92, wy + 5, 10, 12, layer='EQUIPMENT_DETAIL')
        # PC Tower
        rect(wx + 120, wy + 10, 30, 85, layer='EQUIPMENT_DETAIL')
        circle_y = wy + 75
        msp.add_circle((wx + 135, circle_y), 4, dxfattribs={'layer': 'EQUIPMENT_DETAIL'})
        # Text
        text(title_str, wx + 65, wy + 115, h=8, layer='TEXT_EQUIPMENT', align='CENTER')
        text(ip_str, wx + 65, wy - 16, h=7, layer='TEXT_IP_TAG', color=4, align='CENTER')

    # Workstation 1: Engineering Workstation (EWS)
    draw_workstation(op_x + 280, op_y + 80, 'ENGINEERING WORKSTATION', 'IP: 192.168.10.150')

    # Workstation 2: Operator 1
    draw_workstation(op_x + 490, op_y + 80, 'OPERATOR 1', 'IP: 192.168.20.101')

    # Workstation 3: Operator 2
    draw_workstation(op_x + 700, op_y + 80, 'OPERATOR 2', 'IP: 192.168.20.102')

    # 2x Wall Mounted 55" Large Displays (Above Operator Desks)
    disp1_x = op_x + 400
    disp_y = op_y + 260
    disp_w = 220
    disp_h = 130
    rect(disp1_x, disp_y, disp_w, disp_h, layer='EQUIPMENT_FRAME')
    rect(disp1_x + 10, disp_y + 10, disp_w - 20, disp_h - 20, layer='EQUIPMENT_DETAIL')
    text('55" Display Monitor', disp1_x + disp_w / 2, disp_y + 75, h=9, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Display Process Overview', disp1_x + disp_w / 2, disp_y + 50, h=8, layer='TEXT_SUBTITLE', color=4, align='CENTER')

    disp2_x = op_x + 660
    rect(disp2_x, disp_y, disp_w, disp_h, layer='EQUIPMENT_FRAME')
    rect(disp2_x + 10, disp_y + 10, disp_w - 20, disp_h - 20, layer='EQUIPMENT_DETAIL')
    text('55" Display Monitor', disp2_x + disp_w / 2, disp_y + 75, h=9, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Display Trending', disp2_x + disp_w / 2, disp_y + 50, h=8, layer='TEXT_SUBTITLE', color=4, align='CENTER')

    # Alarm & Event Logging Printer
    prt_x = op_x + 980
    prt_y = op_y + 90
    rect(prt_x, prt_y, 110, 85, layer='EQUIPMENT_FRAME')
    rect(prt_x + 15, prt_y + 35, 80, 40, layer='EQUIPMENT_DETAIL')
    rect(prt_x + 25, prt_y + 15, 60, 15, layer='EQUIPMENT_DETAIL')
    text('Printer for', prt_x + 55, prt_y + 102, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Logging & Alarm Report', prt_x + 55, prt_y + 90, h=6, layer='TEXT_EQUIPMENT', align='CENTER')
    text('IP: 192.168.20.200', prt_x + 55, prt_y - 12, h=6, layer='TEXT_IP_TAG', color=4, align='CENTER')

    # ----------------------------------------------------
    # 6. SECTION: OPERATION STATION 2 PANEL (Top Right)
    # ----------------------------------------------------
    op2_x = 1730
    op2_y = 1060
    op2_w = 605
    op2_h = 550

    rect(op2_x, op2_y, op2_w, op2_h, layer='ROOM_BOUNDARY', linetype='DASHED')
    text('OPERATION STATION 2 PANEL (REMOTE CONTROL)', op2_x + 20, op2_y + op2_h - 25, h=11, layer='TEXT_TITLE')

    # Remote Operator Desk & PC
    draw_workstation(op2_x + 180, op2_y + 180, 'OPERATOR STATION 2', 'IP: 192.168.20.103')

    # ----------------------------------------------------
    # 7. SECTION: OEM & THIRD-PARTY PACKAGES (Middle Row)
    # ----------------------------------------------------
    packages_data = [
        ('Dedert Package', 'Burner (10.81.x.x)', 720, 870, 190, 80, 'Burner BMS PLC'),
        ('B+B Packer', 'Bagging System', 950, 870, 180, 80, 'Bag Packing PLC'),
        ('Palletizer', 'Bag Palletizer Robot', 1170, 870, 180, 80, 'Palletizer Controller'),
        ('Rework', 'Starch Rework Skid', 1390, 870, 180, 80, 'Rework Skid I/O'),
        ('Boiler', 'Steam Boiler Package', 1610, 870, 180, 80, 'Boiler Control Panel')
    ]

    for p_name, p_sub, px, py, pw, ph, p_role in packages_data:
        rect(px, py, pw, ph, layer='EQUIPMENT_FRAME', linetype='DASHED')
        text(p_name, px + pw / 2, py + ph - 22, h=8, layer='TEXT_EQUIPMENT', align='CENTER')
        text(p_sub, px + pw / 2, py + ph - 40, h=7, layer='TEXT_SUBTITLE', color=4, align='CENTER')
        text(p_role, px + pw / 2, py + 15, h=6, layer='TEXT_ANNOTATION', color=2, align='CENTER')

    # ----------------------------------------------------
    # 8. SECTION: MAIN CONTROL PANEL (MCP-01: CS1 - CS4)
    # ----------------------------------------------------
    mcp_x = 55
    mcp_y = 120
    mcp_w = 780
    mcp_h = 890

    rect(mcp_x, mcp_y, mcp_w, mcp_h, layer='PANEL_ENCLOSURE')
    text('MAIN CONTROL PANEL (MCP-01) — AREA 602 DRYER BUILDING', mcp_x + 20, mcp_y + mcp_h - 25, h=10, layer='TEXT_TITLE')

    # Top Utility Zone in MCP
    rect(mcp_x + 20, mcp_y + mcp_h - 130, mcp_w - 40, 90, layer='EQUIPMENT_FRAME')
    text('POWER DISTRIBUTION & UTILITY ZONE', mcp_x + 30, mcp_y + mcp_h - 55, h=8, layer='TEXT_EQUIPMENT')
    # 24VDC Power Supplies
    rect(mcp_x + 40, mcp_y + mcp_h - 120, 100, 50, layer='EQUIPMENT_DETAIL')
    text('24VDC 20A PSU 1', mcp_x + 90, mcp_y + mcp_h - 90, h=6, layer='TEXT_EQUIPMENT', align='CENTER')
    rect(mcp_x + 160, mcp_y + mcp_h - 120, 100, 50, layer='EQUIPMENT_DETAIL')
    text('24VDC 20A PSU 2', mcp_x + 210, mcp_y + mcp_h - 90, h=6, layer='TEXT_EQUIPMENT', align='CENTER')
    # Fiber DLR Tap / Converter (MC-01)
    rect(mcp_x + 300, mcp_y + mcp_h - 120, 160, 50, layer='EQUIPMENT_DETAIL')
    text('1783-ETAP2F / MC-01', mcp_x + 380, mcp_y + mcp_h - 85, h=7, layer='TEXT_EQUIPMENT', color=3, align='CENTER')
    text('DLR Fiber Tap (MCP side)', mcp_x + 380, mcp_y + mcp_h - 102, h=6, layer='TEXT_SUBTITLE', color=6, align='CENTER')

    # Function to draw 13-Slot Chassis in DXF
    def draw_chassis_cad(cx, cy, cw, ch, rack_tag, role_str, slots_summary):
        rect(cx, cy, cw, ch, layer='PLC_CHASSIS', color=5)
        text(f'{rack_tag} — {role_str}', cx + 15, cy + ch - 18, h=8, layer='TEXT_EQUIPMENT', color=7)
        # Power supply
        rect(cx + 10, cy + 8, 35, ch - 32, layer='PLC_MODULE', color=8)
        text('PWR', cx + 27, cy + ch / 2, h=6, layer='TEXT_EQUIPMENT', align='CENTER')
        
        # 13 slots
        slot_w = (cw - 65) / 13
        for s_idx in range(13):
            sx = cx + 52 + s_idx * slot_w
            sy = cy + 8
            rect(sx, sy, slot_w - 3, ch - 32, layer='PLC_MODULE', color=4)
            # Slot number
            text(f'{s_idx:02d}', sx + slot_w / 2 - 1.5, sy + ch - 42, h=5, layer='TEXT_EQUIPMENT', align='CENTER')
            # Module name from summary
            mod_text = slots_summary[s_idx]
            text(mod_text, sx + slot_w / 2 - 1.5, sy + 15, h=5, layer='TEXT_MODULE', color=3 if 'EN4' in mod_text or 'CPU' in mod_text else 7, align='CENTER')

    # Chassis Slot summaries based on our engineering redesign
    cs1_slots = ['CPU', 'EN4T', 'EN4T', '32DI', '32DI', '32DI', '32DI', '32DO', '32DO', '16AI', '16AI', '8AO', '8AO']
    cs2_slots = ['EN4T', '32DI', '32DI', '32DI', '32DI', '32DO', '32DO', '16AI', '16AI', '8AO', '8AO', 'N2', 'N2']
    cs3_slots = ['EN4T', '32DI', '32DI', '32DI', '32DI', '32DO', '32DO', '16AI', '16AI', '8AO', '8AO', 'N2', 'N2']
    cs4_slots = ['EN4T', '32DI', '32DI', '32DI', '32DI', '32DO', '32DO', '16AI', '16AI', '8AO', '8AO', 'N2', 'N2']
    cs5_slots = ['EN4T', '32DI', '32DI', '32DI', '32DI', '32DO', '32DO', '16AI', '16AI', '8AO', '8AO', 'N2', 'N2']

    # Draw RACK 1 (CS1)
    draw_chassis_cad(mcp_x + 20, mcp_y + 570, mcp_w - 40, 130, 'RACK 1 (CS1)', '1756-L950TPSXT Master Controller & DLR Supervisor', cs1_slots)
    # Draw RACK 2 (CS2)
    draw_chassis_cad(mcp_x + 20, mcp_y + 400, mcp_w - 40, 130, 'RACK 2 (CS2)', 'Expansion Rack 1 (1756-EN4TR in Slot 00)', cs2_slots)
    # Draw RACK 3 (CS3)
    draw_chassis_cad(mcp_x + 20, mcp_y + 230, mcp_w - 40, 130, 'RACK 3 (CS3)', 'Expansion Rack 2 (1756-EN4TR in Slot 00)', cs3_slots)
    # Draw RACK 4 (CS4)
    draw_chassis_cad(mcp_x + 20, mcp_y + 60, mcp_w - 40, 130, 'RACK 4 (CS4)', 'Expansion Rack 3 (1756-EN4TR in Slot 00)', cs4_slots)

    # ----------------------------------------------------
    # 9. SECTION: RIO AREA-200 (Bottom Center)
    # ----------------------------------------------------
    rio_x = 880
    rio_y = 380
    rio_w = 540
    rio_h = 360

    rect(rio_x, rio_y, rio_w, rio_h, layer='PANEL_ENCLOSURE')
    text('RIO Area-200 (PRE-SLURRY EXPANSION PANEL)', rio_x + 20, rio_y + rio_h - 25, h=10, layer='TEXT_TITLE')

    # Fiber Tap / Converter (MC-02) in RIO200
    rect(rio_x + 30, rio_y + rio_h - 110, 180, 50, layer='EQUIPMENT_DETAIL')
    text('1783-ETAP2F / MC-02', rio_x + 120, rio_y + rio_h - 78, h=7, layer='TEXT_EQUIPMENT', color=3, align='CENTER')
    text('DLR Fiber Tap (RIO side)', rio_x + 120, rio_y + rio_h - 95, h=6, layer='TEXT_SUBTITLE', color=6, align='CENTER')

    # Draw RACK 5 (CS5) in RIO200
    draw_chassis_cad(rio_x + 20, rio_y + 70, rio_w - 40, 140, 'RACK 5 (CS5)', 'Remote I/O Chassis via Fiber Optic', cs5_slots)

    # ----------------------------------------------------
    # 10. SECTION: MCC PANEL (Bottom Right)
    # ----------------------------------------------------
    mcc_x = 1460
    mcc_y = 240
    mcc_w = 875
    mcc_h = 500

    rect(mcc_x, mcc_y, mcc_w, mcc_h, layer='PANEL_ENCLOSURE', linetype='DASHED')
    text('MCC Panel (MOTOR CONTROL CENTER) — POWER & DRIVES', mcc_x + 20, mcc_y + mcc_h - 25, h=10, layer='TEXT_TITLE')

    # MCC Cubicles 1 to 4
    cub_w = (mcc_w - 40) / 4
    for c_idx in range(4):
        cx = mcc_x + 20 + c_idx * cub_w
        rect(cx, mcc_y + 30, cub_w - 5, mcc_h - 70, layer='EQUIPMENT_FRAME')
        text(f'CUBICLE {c_idx + 1}', cx + cub_w / 2 - 2.5, mcc_y + mcc_h - 55, h=7, layer='TEXT_EQUIPMENT', align='CENTER')

    # Cubicle 1: Incoming Breaker
    rect(mcc_x + 30, mcc_y + 240, cub_w - 25, 140, layer='EQUIPMENT_DETAIL')
    text('MAIN INCOMING ACB', mcc_x + 20 + cub_w / 2, mcc_y + 350, h=8, layer='TEXT_EQUIPMENT', align='CENTER')
    text('1600A 400V 50Hz 50kA', mcc_x + 20 + cub_w / 2, mcc_y + 330, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Power Meter 1408-BC3A', mcc_x + 20 + cub_w / 2, mcc_y + 310, h=7, layer='TEXT_IP_TAG', color=4, align='CENTER')

    # Busbar line across cubicles
    line((mcc_x + 40, mcc_y + 210), (mcc_x + mcc_w - 40, mcc_y + 210), layer='EQUIPMENT_DETAIL', color=2)
    text('400VAC MAIN BUSBAR (COPPER)', mcc_x + 250, mcc_y + 216, h=6, layer='TEXT_ANNOTATION', color=2)

    # Cubicle 2: VFD Section 1
    vfd_w = cub_w - 30
    rect(mcc_x + 20 + cub_w + 10, mcc_y + 240, vfd_w, 140, layer='EQUIPMENT_DETAIL')
    text('PowerFlex 755 (110kW)', mcc_x + 20 + cub_w + cub_w / 2, mcc_y + 345, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Main Supply Air Fan FN-60201', mcc_x + 20 + cub_w + cub_w / 2, mcc_y + 325, h=6, layer='TEXT_SUBTITLE', align='CENTER')
    text('IP: 192.168.60.21', mcc_x + 20 + cub_w + cub_w / 2, mcc_y + 305, h=6, layer='TEXT_IP_TAG', color=4, align='CENTER')

    rect(mcc_x + 20 + cub_w + 10, mcc_y + 60, vfd_w, 130, layer='EQUIPMENT_DETAIL')
    text('PowerFlex 755 (132kW)', mcc_x + 20 + cub_w + cub_w / 2, mcc_y + 160, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Exhaust Air Fan FN-60202', mcc_x + 20 + cub_w + cub_w / 2, mcc_y + 140, h=6, layer='TEXT_SUBTITLE', align='CENTER')
    text('IP: 192.168.60.22', mcc_x + 20 + cub_w + cub_w / 2, mcc_y + 120, h=6, layer='TEXT_IP_TAG', color=4, align='CENTER')

    # Cubicle 3: VFD Section 2
    rect(mcc_x + 20 + cub_w * 2 + 10, mcc_y + 240, vfd_w, 140, layer='EQUIPMENT_DETAIL')
    text('PowerFlex 525 (18.5kW)', mcc_x + 20 + cub_w * 2 + cub_w / 2, mcc_y + 345, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Slurry Transfer Pump PC-20201', mcc_x + 20 + cub_w * 2 + cub_w / 2, mcc_y + 325, h=6, layer='TEXT_SUBTITLE', align='CENTER')
    text('IP: 192.168.40.21', mcc_x + 20 + cub_w * 2 + cub_w / 2, mcc_y + 305, h=6, layer='TEXT_IP_TAG', color=4, align='CENTER')

    rect(mcc_x + 20 + cub_w * 2 + 10, mcc_y + 60, vfd_w, 130, layer='EQUIPMENT_DETAIL')
    text('PowerFlex 525 (18.5kW)', mcc_x + 20 + cub_w * 2 + cub_w / 2, mcc_y + 160, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('Slurry Transfer Pump PC-20211', mcc_x + 20 + cub_w * 2 + cub_w / 2, mcc_y + 140, h=6, layer='TEXT_SUBTITLE', align='CENTER')
    text('IP: 192.168.40.22', mcc_x + 20 + cub_w * 2 + cub_w / 2, mcc_y + 120, h=6, layer='TEXT_IP_TAG', color=4, align='CENTER')

    # Cubicle 4: Auxiliary Feeders & Notes
    rect(mcc_x + 20 + cub_w * 3 + 10, mcc_y + 240, vfd_w, 140, layer='EQUIPMENT_DETAIL')
    text('Agitator Starters', mcc_x + 20 + cub_w * 3 + cub_w / 2, mcc_y + 345, h=7, layer='TEXT_EQUIPMENT', align='CENTER')
    text('ME-20201 / ME-20211', mcc_x + 20 + cub_w * 3 + cub_w / 2, mcc_y + 325, h=6, layer='TEXT_SUBTITLE', align='CENTER')

    # Handwritten Plant Feed Notes from Drawing
    text('1st SPD Plant Feeder', mcc_x + 20 + cub_w * 3 + cub_w / 2, mcc_y + 160, h=8, layer='TEXT_ANNOTATION', color=2, align='CENTER')
    text('4th Floor Distribution', mcc_x + 20 + cub_w * 3 + cub_w / 2, mcc_y + 140, h=8, layer='TEXT_ANNOTATION', color=2, align='CENTER')
    text('Boiler Support Power', mcc_x + 20 + cub_w * 3 + cub_w / 2, mcc_y + 120, h=8, layer='TEXT_ANNOTATION', color=2, align='CENTER')

    # ----------------------------------------------------
    # 11. CABLE RUNS & NETWORK CONNECTIONS
    # ----------------------------------------------------
    # 11A. RED LINES: ETHERNET LAN CABLE FOR CONTROL
    # Server Room -> Network Cabinet in Op1 Room
    polyline([(rack_x + rack_w, rack_y + 360), (net_x, rack_y + 360)], layer='CBL_CONTROL_LAN', color=1)
    text('Fiber optic (Control LAN)', (rack_x + rack_w + net_x) / 2, rack_y + 372, h=7, layer='TEXT_ANNOTATION', color=2, align='CENTER')

    # Network Cabinet -> Workstations (EWS, Op1, Op2, Displays, Printer)
    polyline([(net_x + net_w, net_y + 320), (op_x + 350, net_y + 320), (op_x + 350, op_y + 190)], layer='CBL_CONTROL_LAN', color=1)
    polyline([(net_x + net_w, net_y + 320), (op_x + 560, net_y + 320), (op_x + 560, op_y + 190)], layer='CBL_CONTROL_LAN', color=1)
    polyline([(net_x + net_w, net_y + 320), (op_x + 770, net_y + 320), (op_x + 770, op_y + 190)], layer='CBL_CONTROL_LAN', color=1)
    polyline([(net_x + net_w, net_y + 320), (disp1_x + disp_w / 2, net_y + 320), (disp1_x + disp_w / 2, disp_y)], layer='CBL_CONTROL_LAN', color=1)
    polyline([(net_x + net_w, net_y + 320), (disp2_x + disp_w / 2, net_y + 320), (disp2_x + disp_w / 2, disp_y)], layer='CBL_CONTROL_LAN', color=1)
    polyline([(net_x + net_w, net_y + 320), (prt_x + 55, net_y + 320), (prt_x + 55, prt_y + 85)], layer='CBL_CONTROL_LAN', color=1)

    # Network Cabinet -> Operation Station 2 Panel
    polyline([(net_x + net_w, net_y + 300), (op2_x + 100, net_y + 300), (op2_x + 100, op2_y + 250), (op2_x + 220, op2_y + 250)], layer='CBL_CONTROL_LAN', color=1)

    # Network Cabinet -> Vendor Packages (Dedert, B+B Packer, Palletizer, Rework, Boiler)
    pkg_trunk_y = 1005
    polyline([(net_x + 100, net_y), (net_x + 100, pkg_trunk_y), (1700, pkg_trunk_y)], layer='CBL_CONTROL_LAN', color=1)
    text('Ethernet (10.81.x.x)', 720, pkg_trunk_y + 12, h=7, layer='TEXT_ANNOTATION', color=2)

    for px in [815, 1040, 1260, 1480, 1700]:
        line((px, pkg_trunk_y), (px, 950), layer='CBL_CONTROL_LAN', color=1)
    text('Fiber optic (Remote Skids)', 1550, pkg_trunk_y + 12, h=7, layer='TEXT_ANNOTATION', color=6)

    # Network Cabinet -> Main Control Panel (Rack 1 Slot 02 SCADA Uplink)
    polyline([(net_x + 60, net_y), (net_x + 60, mcp_y + 735), (mcp_x + 210, mcp_y + 735), (mcp_x + 210, mcp_y + 700)], layer='CBL_CONTROL_LAN', color=1)
    text('CBL-UPL-01 (To SCADA)', net_x - 10, 800, h=7, layer='TEXT_ANNOTATION', color=1)

    # 11B. GREEN LINES: ETHERNET FOR I/O NETWORK (DLR RING)
    # RACK 1 (Supervisor) -> RACK 2
    r1_p1 = (mcp_x + 130, mcp_y + 570)
    r2_p1 = (mcp_x + 130, mcp_y + 530)
    polyline([(r1_p1[0], r1_p1[1]), (mcp_x + 15, r1_p1[1]), (mcp_x + 15, r2_p1[1]), (r2_p1[0], r2_p1[1])], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-01', mcp_x - 10, (r1_p1[1] + r2_p1[1]) / 2, h=6, layer='TEXT_TAG', color=3)

    # RACK 2 -> RACK 3
    r2_p2 = (mcp_x + 130, mcp_y + 400)
    r3_p1 = (mcp_x + 130, mcp_y + 360)
    polyline([(r2_p2[0], r2_p2[1]), (mcp_x + 15, r2_p2[1]), (mcp_x + 15, r3_p1[1]), (r3_p1[0], r3_p1[1])], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-02', mcp_x - 10, (r2_p2[1] + r3_p1[1]) / 2, h=6, layer='TEXT_TAG', color=3)

    # RACK 3 -> RACK 4
    r3_p2 = (mcp_x + 130, mcp_y + 230)
    r4_p1 = (mcp_x + 130, mcp_y + 190)
    polyline([(r3_p2[0], r3_p2[1]), (mcp_x + 15, r3_p2[1]), (mcp_x + 15, r4_p1[1]), (r4_p1[0], r4_p1[1])], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-03', mcp_x - 10, (r3_p2[1] + r4_p1[1]) / 2, h=6, layer='TEXT_TAG', color=3)

    # RACK 4 -> Fiber Tap (MC-01)
    r4_p2 = (mcp_x + 130, mcp_y + 60)
    tap_mcp = (mcp_x + 350, mcp_y + mcp_h - 120)
    polyline([(r4_p2[0], r4_p2[1]), (mcp_x + 5, r4_p2[1]), (mcp_x + 5, mcp_y + mcp_h - 145), (tap_mcp[0], mcp_y + mcp_h - 145), (tap_mcp[0], tap_mcp[1])], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-04', mcp_x - 20, 300, h=6, layer='TEXT_TAG', color=3)

    # Fiber Tap MC-01 -> RIO Area-200 Fiber Tap MC-02 (FIBER OUTBOUND RUN)
    # Routed cleanly below packages at Y=810, avoiding Dedert package box!
    tap_rio = (rio_x + 100, rio_y + rio_h - 110)
    polyline([(mcp_x + 460, mcp_y + mcp_h - 95), (mcp_x + 500, mcp_y + mcp_h - 95), (mcp_x + 500, 810), (rio_x + 60, 810), (rio_x + 60, tap_rio[1] + 25), (tap_rio[0], tap_rio[1] + 25), (tap_rio[0], tap_rio[1])], layer='CBL_FIBER_OPTIC', linetype='DASHED', color=6)
    text('Fiber optic (FOB-RIO-01 Outbound)', 650, 820, h=7, layer='TEXT_ANNOTATION', color=6)

    # RIO200 Tap MC-02 -> RACK 5 (CS5)
    polyline([(tap_rio[0] + 40, tap_rio[1]), (tap_rio[0] + 40, rio_y + 210), (rio_x + 85, rio_y + 210)], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-05', tap_rio[0] + 48, rio_y + 225, h=6, layer='TEXT_TAG', color=3)

    # RACK 5 (CS5) -> RIO200 Tap MC-02 (Return path)
    polyline([(rio_x + 115, rio_y + 70), (rio_x + 115, rio_y + 40), (rio_x + 250, rio_y + 40), (rio_x + 250, rio_y + rio_h - 85), (tap_rio[0] + 60, rio_y + rio_h - 85), (tap_rio[0] + 60, tap_rio[1])], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-06', rio_x + 140, rio_y + 48, h=6, layer='TEXT_TAG', color=3)

    # Fiber Tap MC-02 -> Main Control Panel MC-01 (FIBER RETURN RUN)
    # Routed cleanly below RIO200 at Y=320!
    polyline([(tap_rio[0] + 20, tap_rio[1]), (tap_rio[0] + 20, 320), (mcp_x + 720, 320), (mcp_x + 720, mcp_y + mcp_h - 110), (mcp_x + 420, mcp_y + mcp_h - 110)], layer='CBL_FIBER_OPTIC', linetype='DASHED', color=6)
    text('Fiber optic (FOB-RIO-02 Return)', 680, 328, h=7, layer='TEXT_ANNOTATION', color=6)

    # MC-01 -> RACK 1 Slot 01 Port 2 (CLOSES RING TO BEACON PORT)
    # Routed cleanly through the top margin of Rack 1
    polyline([(mcp_x + 320, tap_mcp[1]), (mcp_x + 320, mcp_y + 715), (mcp_x + 160, mcp_y + 715), (mcp_x + 160, mcp_y + 700)], layer='CBL_IO_NETWORK', color=3)
    text('CBL-DLR-07 (Ring Close to P2 Beacon)', mcp_x + 175, mcp_y + 722, h=6, layer='TEXT_TAG', color=3)

    # 11C. BLUE LINES: ETHERNET FOR MCC NETWORK
    # Main Control Panel -> MCC Panel (routed cleanly around bottom at Y=170)
    polyline([(mcp_x + mcp_w, mcp_y + 400), (mcp_x + mcp_w + 30, mcp_y + 400), (mcp_x + mcp_w + 30, 170), (mcc_x + 30, 170), (mcc_x + 30, mcc_y + 120)], layer='CBL_MCC_NETWORK', color=5)
    text('ETHERNET FOR MCC NETWORK (Blue)', 1050, 178, h=7, layer='TEXT_TAG', color=5)

    return doc

def export_all_formats(dxf_path, out_dir):
    import re
    from ezdxf.addons.drawing import Frontend, RenderContext, layout
    from ezdxf.addons.drawing.svg import SVGBackend
    from ezdxf.addons.drawing.config import Configuration, BackgroundPolicy, ColorPolicy

    doc = ezdxf.readfile(dxf_path)
    msp = doc.modelspace()
    ctx = RenderContext(doc)
    ctx.set_current_layout(msp)
    ctx.current_layout_properties.set_colors(bg='#ffffff', fg='#000000')

    backend = SVGBackend()
    config = Configuration(
        background_policy=BackgroundPolicy.WHITE,
        color_policy=ColorPolicy.COLOR,
        lineweight_scaling=1.5
    )

    Frontend(ctx, backend, config=config).draw_layout(msp)

    page_layout = layout.Page.from_dxf_layout(msp)
    svg_raw = backend.get_string(page_layout)

    # Inline styles into SVG paths so PyMuPDF, browsers, and vector apps render identically
    class_styles = {}
    for m in re.finditer(r'\.([A-Z0-9]+)\s*\{([^}]+)\}', svg_raw):
        cls = m.group(1)
        props = m.group(2)
        class_styles[cls] = props

    def repl_path(match):
        full = match.group(0)
        cls_match = re.search(r'class=[\"\']([A-Z0-9]+)[\"\']', full)
        if cls_match:
            cls = cls_match.group(1)
            style_val = class_styles.get(cls, '')
            fill_m = re.search(r'fill:\s*([^;]+);', style_val)
            stroke_m = re.search(r'stroke:\s*([^;]+);', style_val)
            sw_m = re.search(r'stroke-width:\s*([^;]+);', style_val)
            fill = fill_m.group(1) if fill_m else 'none'
            stroke = stroke_m.group(1) if stroke_m else 'none'
            sw = sw_m.group(1) if sw_m else '1'
            res = re.sub(r'class=[\"\'][A-Z0-9]+[\"\']', f'fill="{fill}" stroke="{stroke}" stroke-width="{sw}"', full)
            return res
        return full

    inlined_svg = re.sub(r'<path[^>]+>', repl_path, svg_raw)

    svg_path = os.path.join(out_dir, "SPRINT_18K_Control_Layout_System_Architecture.svg")
    with open(svg_path, "w", encoding="utf-8") as f:
        f.write(inlined_svg)
    print(f"Saved SVG: {svg_path}")

    # Generate PDF
    pdf_doc = fitz.open(svg_path)
    pdf_bytes = pdf_doc.convert_to_pdf()
    pdf_final = fitz.open('pdf', pdf_bytes)
    pdf_path = os.path.join(out_dir, "SPRINT_18K_Control_Layout_System_Architecture.pdf")
    pdf_final.save(pdf_path)
    print(f"Saved PDF: {pdf_path}")

    # Generate high-res PNG (150 DPI for crisp readability)
    page = pdf_final[0]
    pix = page.get_pixmap(dpi=150)
    png_path = os.path.join(out_dir, "SPRINT_18K_Control_Layout_System_Architecture.png")
    pix.save(png_path)
    print(f"Saved PNG: {png_path} ({pix.width}x{pix.height})")

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(out_dir, exist_ok=True)
    dxf_path = os.path.join(out_dir, "SPRINT_18K_Control_Layout_System_Architecture.dxf")

    print("Building AutoCAD DXF...")
    doc = create_dxf()
    doc.saveas(dxf_path)
    print(f"Successfully generated DXF: {dxf_path}")

    print("Exporting SVG, PDF, and PNG companion deliverables...")
    export_all_formats(dxf_path, out_dir)
    print("ALL DELIVERABLES COMPLETED!")
