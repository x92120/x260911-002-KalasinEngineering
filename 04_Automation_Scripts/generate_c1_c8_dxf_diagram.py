#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
CHASSIS C1-C8 SLOT TO DESTINATION JUNCTION BOX ARCHITECTURE DIAGRAM (DXF + SVG)
================================================================================
Generates:
  1. High-Precision AutoCAD DXF (1:1 Metric mm model space with standard CAD layers):
     02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/C1_C8_Slot_to_Destination_JB_Architecture_Diagram.dxf
  2. Scalable Vector Graphics (SVG):
     02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/C1_C8_Slot_to_Destination_JB_Architecture_Diagram.svg
================================================================================
"""

import os
import sys
from collections import defaultdict, Counter
import ezdxf
from ezdxf import colors
from ezdxf.enums import TextEntityAlignment

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
CAD_DIR = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG")
os.makedirs(CAD_DIR, exist_ok=True)

DXF_OUTPUT = os.path.join(CAD_DIR, "C1_C8_Slot_to_Destination_JB_Architecture_Diagram.dxf")
SVG_OUTPUT = os.path.join(CAD_DIR, "C1_C8_Slot_to_Destination_JB_Architecture_Diagram.svg")

# Data definitions
CHASSIS_DEFS = [
    {
        'id': 'C1',
        'title': 'CHASSIS C1: MAIN CONTROLLER RACK',
        'loc': 'CA1 Main Cabinet (Control Room)',
        'slots': [
            ('00', '1756-L950TPSXT', 'CPU', [('CA1', 1)]),
            ('01', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('02', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('03', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('04', '1756-IB32', 'DI', [('CA1', 16), ('JB-401', 16)]),
            ('05', '1756-IB32', 'DI', [('JB-401', 32)]),
            ('06', '1756-IB32', 'DI', [('JB-401', 16), ('JB-601', 16)]),
            ('07', '1756-IB32', 'DI', [('JB-601', 32)]),
            ('08', '1756-OB32', 'DO', [('CA1', 16), ('JB-401', 16)]),
            ('09', '1756-OB32', 'DO', [('JB-401', 16), ('JB-601', 16)]),
            ('10', '1756-IF16', 'AI', [('JB-401', 16)]),
            ('11', '1756-IF16', 'AI', [('JB-601', 16)]),
        ]
    },
    {
        'id': 'C2',
        'title': 'CHASSIS C2: MAIN EXPANSION RACK',
        'loc': 'CA1 Main Cabinet (Control Room)',
        'slots': [
            ('00', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('01', '1756-IB32', 'DI', [('JB-402', 32)]),
            ('02', '1756-IB32', 'DI', [('JB-602', 32)]),
            ('03', '1756-IB32', 'DI', [('JB-607', 29)]),
            ('04', '1756-IB32', 'DI', [('JB-607', 32)]),
            ('05', '1756-IB32', 'DI', [('JB-607', 32)]),
            ('06', '1756-OB32', 'DO', [('JB-402', 16), ('JB-602', 16)]),
            ('07', '1756-OB32', 'DO', [('JB-607', 32)]),
            ('08', '1756-OB32', 'DO', [('JB-607', 32)]),
            ('09', '1756-IF16', 'AI', [('JB-402', 16)]),
            ('10', '1756-IF16', 'AI', [('JB-602', 16)]),
        ]
    },
    {
        'id': 'C3',
        'title': 'CHASSIS C3: REMOTE I/O RACK 1',
        'loc': 'Field Enclosure (Spray Dryer 3rd/7th Fl)',
        'slots': [
            ('00', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('01', '1756-IB32', 'DI', [('JB-618', 11)]),
            ('02', '1756-IB32', 'DI', [('JB-618', 5)]),
            ('03', '1756-OB32', 'DO', [('JB-618', 32)]),
            ('05', '1756-IF16', 'AI', [('JB-618', 4)]),
            ('06', '1756-IF16/RTD', 'AI', [('JB-602', 16)]),
            ('07', '1756-IF16/RTD', 'AI', [('JB-607', 16)]),
            ('08', '1756-IF16/RTD', 'AI', [('JB-607', 16)]),
            ('11', '1756-OF8', 'AO', [('JB-402', 1), ('JB-602', 2), ('JB-607', 5)]),
        ]
    },
    {
        'id': 'C4',
        'title': 'CHASSIS C4: REMOTE I/O RACK 2',
        'loc': 'Field Enclosure (Spray Dryer 6th/8th Fl)',
        'slots': [
            ('00', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('01', '1756-IB32', 'DI', [('JB-606', 16)]),
            ('02', '1756-IB32', 'DI', [('JB-612', 32)]),
            ('04', '1756-OB32', 'DO', [('JB-606', 8), ('JB-612', 16)]),
            ('05', '1756-IF16', 'AI', [('JB-606', 2), ('JB-608', 6), ('JB-612', 8)]),
        ]
    },
    {
        'id': 'C5',
        'title': 'CHASSIS C5: RIO-200 SLURRY PROCESS',
        'loc': 'Slurry Out-Building 2nd Floor',
        'slots': [
            ('00', '1756-EN4TR', 'ETH', [('CA1', 1)]),
            ('01', '1756-IB32', 'DI', [('RIO-200', 32)]),
            ('02', '1756-IB32', 'DI', [('RIO-200', 32)]),
            ('03', '1756-OB32', 'DO', [('RIO-200', 32)]),
            ('04', '1756-IF16', 'AI', [('RIO-200', 16)]),
        ]
    },
    {
        'id': 'C6',
        'title': 'CHASSIS C6: MOTOR CONTROL INTERFACE',
        'loc': 'MCC Room Ground Floor',
        'slots': [
            ('05', '1756-OB16', 'DO', [('MCC', 8)]),
            ('06', '1756-IB16', 'DI', [('MCC', 10)]),
        ]
    },
    {
        'id': 'C7',
        'title': 'CHASSIS C7: MCC & BUS SUPERVISORY RACK',
        'loc': 'MCC Room Ground Floor',
        'slots': [
            ('00', 'Fieldbus-GW', 'BUS', [('JB-601', 1), ('JB-607', 1), ('JB-612', 1), ('MCC', 17)]),
            ('01', '1756-IB32', 'DI', [('MCC', 32)]),
            ('02', '1756-IB32', 'DI', [('JB-601', 2), ('MCC', 30)]),
            ('03', '1756-IB32', 'DI', [('JB-601', 1), ('MCC', 25)]),
            ('04', '1756-IB32', 'DI', [('MCC', 28)]),
            ('05', '1756-OB32', 'DO', [('MCC', 32)]),
            ('06', '1756-OB32', 'DO', [('JB-601', 1), ('MCC', 31)]),
            ('07', '1756-OB32', 'DO', [('MCC', 24)]),
        ]
    },
    {
        'id': 'C8',
        'title': 'CHASSIS C8: INTRINSICALLY SAFE (IS) SUB-SYSTEM',
        'loc': 'Field Galvanic Isolation Cabinets (Ex Zones)',
        'slots': [
            ('01', 'IS-IB32', 'DI', [('IS-JB-603', 16), ('IS-JB-608', 16)]),
            ('02', 'IS-IB32', 'DI', [('IS-JB-612', 31)]),
            ('03', 'IS-IB32', 'DI', [('IS-JB-618', 32)]),
            ('04', 'IS-OB08', 'DO', [('IS-JB-608', 8)]),
            ('05', 'IS-IF16', 'AI', [('IS-JB-603', 16)]),
            ('06', 'IS-IF16', 'AI', [('IS-JB-608', 8), ('IS-JB-612', 10)]),
            ('07', 'IS-IF16', 'AI', [('IS-JB-618', 12)]),
        ]
    }
]

DEST_DEFS = [
    ('CA1', 'Control Cabinet CA1', 'Main Control Room', 37),
    ('JB-401', 'Infeed Junction Box', 'Infeed 2nd Floor', 144),
    ('JB-402', 'Jet Cooker Junction Box', 'Jet Cooker 2nd Floor', 65),
    ('JB-601', 'Spray Dryer Bottom JB', 'Spray Dryer 1st Floor', 85),
    ('JB-602', 'Spray Dryer Chamber JB', 'Spray Dryer 3rd Floor', 66),
    ('JB-606', 'Spray Dryer Filter JB', 'Spray Dryer 6th Floor', 26),
    ('JB-607', 'Spray Dryer Exhaust JB', 'Spray Dryer 7th Floor', 211),
    ('JB-608', 'Spray Dryer Air Top JB', 'Spray Dryer 8th Floor', 6),
    ('JB-612', 'Packing Tower Lower JB', 'Packing Tower 2nd Floor', 57),
    ('JB-618', 'Packing Tower Upper JB', 'Packing Tower 8th Floor', 52),
    ('RIO-200', 'Remote Slurry Station', 'Slurry Out-Building 2nd Fl', 112),
    ('MCC', 'Motor Control Center', 'MCC Room Ground Floor', 237),
    ('IS-JB-603', 'IS Chamber Ex JB', 'Spray Dryer 3rd Floor (IS)', 32),
    ('IS-JB-608', 'IS Top Level Ex JB', 'Spray Dryer 8th Floor (IS)', 32),
    ('IS-JB-612', 'IS Lower Tower Ex JB', 'Packing Tower 2nd Fl (IS)', 41),
    ('IS-JB-618', 'IS Upper Tower Ex JB', 'Packing Tower 8th Fl (IS)', 44),
]

def generate_cad():
    print("Initializing AutoCAD DXF Document (R2010 Metric mm)...")
    doc = ezdxf.new("R2010", setup=True)
    doc.header["$INSUNITS"] = 4  # 4 = mm
    msp = doc.modelspace()
    
    # Setup Layers with standard ACI colors and lineweights
    layers = [
        ("0_BORDER", 7, 50),             # White/Black (0.50mm)
        ("0_TITLE_BLOCK", 4, 35),        # Cyan (0.35mm)
        ("CHASSIS_FRAME", 7, 40),        # White (0.40mm)
        ("SLOT_BOX", 8, 25),             # Light Gray (0.25mm)
        ("DEST_FRAME", 3, 40),           # Green (0.40mm)
        ("ROUTING_DI", 4, 25),           # Cyan (0.25mm)
        ("ROUTING_DO", 2, 25),           # Yellow (0.25mm)
        ("ROUTING_AI", 1, 25),           # Red (0.25mm)
        ("ROUTING_AO", 6, 25),           # Magenta (0.25mm)
        ("ROUTING_BUS", 5, 25),          # Blue (0.25mm)
        ("ROUTING_ETH", 8, 20),          # Gray (0.20mm)
        ("TEXT_MAIN", 7, 30),            # White (0.30mm)
        ("TEXT_SUB", 4, 20),             # Cyan (0.20mm)
        ("TEXT_ANNOT", 8, 15),           # Gray (0.15mm)
        ("TITLE_TEXT", 2, 40)            # Yellow (0.40mm)
    ]
    for name, col, lw in layers:
        layer = doc.layers.add(name)
        layer.color = col
        layer.lineweight = lw

    # -------------------------------------------------------------
    # DRAWING DIMENSIONS & BORDER
    # Canvas Size: 1800mm x 1150mm (A0+ Standard Proportion)
    # -------------------------------------------------------------
    WIDTH = 1800.0
    HEIGHT = 1150.0
    MARGIN = 25.0
    
    # Outer Border
    msp.add_lwpolyline([(MARGIN, MARGIN), (WIDTH - MARGIN, MARGIN), 
                        (WIDTH - MARGIN, HEIGHT - MARGIN), (MARGIN, HEIGHT - MARGIN)],
                       close=True, dxfattribs={"layer": "0_BORDER"})
    # Inner Border
    msp.add_lwpolyline([(MARGIN + 5, MARGIN + 5), (WIDTH - MARGIN - 5, MARGIN + 5), 
                        (WIDTH - MARGIN - 5, HEIGHT - MARGIN - 5), (MARGIN + 5, HEIGHT - MARGIN - 5)],
                       close=True, dxfattribs={"layer": "0_BORDER"})

    # Title Banner at Top
    top_y = HEIGHT - MARGIN - 12
    msp.add_text("KALASIN ENGINEERING xCIP-1545 SYSTEM (PROJECT REF: x2608003)",
                 dxfattribs={"layer": "TITLE_TEXT", "height": 9.0, "style": "Standard"}).set_placement((MARGIN + 20, top_y - 12))
    msp.add_text("CHASSIS C1 - C8 TO DESTINATION JUNCTION BOX ARCHITECTURE & INTERCONNECTION DIAGRAM",
                 dxfattribs={"layer": "TEXT_MAIN", "height": 7.0, "style": "Standard"}).set_placement((MARGIN + 20, top_y - 24))
    msp.add_line((MARGIN + 10, top_y - 30), (WIDTH - MARGIN - 10, top_y - 30), dxfattribs={"layer": "0_BORDER"})

    # -------------------------------------------------------------
    # LEFT REGION: 8 CHASSIS BLOCKS (2 Columns x 4 Rows)
    # -------------------------------------------------------------
    # Column 1 (C1, C3, C5, C7): X=50 to X=430
    # Column 2 (C2, C4, C6, C8): X=450 to X=830
    # -------------------------------------------------------------
    chassis_ports = {}  # key: (chassis_id, slot_id, dest_name) -> (x, y)
    
    row_y_starts = [top_y - 45, top_y - 315, top_y - 585, top_y - 855]
    
    col_x_offsets = {
        'C1': (50.0, row_y_starts[0]),
        'C2': (440.0, row_y_starts[0]),
        'C3': (50.0, row_y_starts[1]),
        'C4': (440.0, row_y_starts[1]),
        'C5': (50.0, row_y_starts[2]),
        'C6': (440.0, row_y_starts[2]),
        'C7': (50.0, row_y_starts[3]),
        'C8': (440.0, row_y_starts[3])
    }
    
    CH_WIDTH = 370.0
    CH_HEADER_H = 26.0
    SLOT_H = 17.0
    
    for ch_def in CHASSIS_DEFS:
        cid = ch_def['id']
        x0, y0 = col_x_offsets[cid]
        num_slots = len(ch_def['slots'])
        total_ch_h = CH_HEADER_H + num_slots * SLOT_H + 6.0
        
        # Chassis Outer Enclosure Box
        msp.add_lwpolyline([(x0, y0), (x0 + CH_WIDTH, y0), 
                            (x0 + CH_WIDTH, y0 - total_ch_h), (x0, y0 - total_ch_h)],
                           close=True, dxfattribs={"layer": "CHASSIS_FRAME"})
        
        # Header Box
        msp.add_lwpolyline([(x0, y0), (x0 + CH_WIDTH, y0), 
                            (x0 + CH_WIDTH, y0 - CH_HEADER_H), (x0, y0 - CH_HEADER_H)],
                           close=True, dxfattribs={"layer": "0_TITLE_BLOCK"})
        
        msp.add_text(ch_def['title'], dxfattribs={"layer": "TITLE_TEXT", "height": 4.2}).set_placement((x0 + 8.0, y0 - 10.0))
        msp.add_text(f"LOC: {ch_def['loc']}", dxfattribs={"layer": "TEXT_SUB", "height": 3.2}).set_placement((x0 + 8.0, y0 - 20.0))
        
        # Slots
        curr_y = y0 - CH_HEADER_H - 4.0
        for slot_no, card, io_type, dest_list in ch_def['slots']:
            slot_box_y = curr_y - SLOT_H
            msp.add_lwpolyline([(x0 + 4, curr_y), (x0 + CH_WIDTH - 4, curr_y),
                                (x0 + CH_WIDTH - 4, slot_box_y), (x0 + 4, slot_box_y)],
                               close=True, dxfattribs={"layer": "SLOT_BOX"})
            
            # Slot Label & Card
            msp.add_text(f"SLOT {slot_no}", dxfattribs={"layer": "TEXT_MAIN", "height": 3.4}).set_placement((x0 + 8.0, slot_box_y + 9.5))
            msp.add_text(card, dxfattribs={"layer": "TEXT_SUB", "height": 3.2}).set_placement((x0 + 46.0, slot_box_y + 9.5))
            
            # Badge for Signal Type
            badge_color_layer = f"ROUTING_{io_type}" if f"ROUTING_{io_type}" in [l[0] for l in layers] else "TEXT_MAIN"
            msp.add_text(f"[{io_type}]", dxfattribs={"layer": badge_color_layer, "height": 3.2}).set_placement((x0 + 130.0, slot_box_y + 9.5))
            
            # Destination summary text inside slot
            d_summary = ", ".join([f"{d} ({pts}p)" for d, pts in dest_list])
            if len(d_summary) > 28:
                d_summary = d_summary[:26] + ".."
            msp.add_text(f"-> {d_summary}", dxfattribs={"layer": "TEXT_ANNOT", "height": 2.8}).set_placement((x0 + 175.0, slot_box_y + 4.5))
            
            # Store connection port coordinates (right edge of slot box)
            port_x = x0 + CH_WIDTH - 4.0
            port_y = slot_box_y + (SLOT_H / 2.0)
            for d_name, pts in dest_list:
                chassis_ports[(cid, slot_no, d_name)] = (port_x, port_y, io_type, pts)
                
            curr_y -= SLOT_H + 1.5

    # -------------------------------------------------------------
    # RIGHT REGION: 16 DESTINATION JUNCTION BOXES & CABINETS
    # Located from X=1380 to X=1750
    # Arranged vertically from top to bottom
    # -------------------------------------------------------------
    DEST_X0 = 1380.0
    DEST_W = 370.0
    DEST_H = 46.0
    DEST_GAP = 13.0
    
    dest_ports = {}
    dest_curr_y = top_y - 45.0
    
    for d_name, full_name, loc_str, total_pts in DEST_DEFS:
        d_box_y = dest_curr_y - DEST_H
        msp.add_lwpolyline([(DEST_X0, dest_curr_y), (DEST_X0 + DEST_W, dest_curr_y),
                            (DEST_X0 + DEST_W, d_box_y), (DEST_X0, d_box_y)],
                           close=True, dxfattribs={"layer": "DEST_FRAME"})
        
        # Top banner within box
        msp.add_lwpolyline([(DEST_X0, dest_curr_y), (DEST_X0 + DEST_W, dest_curr_y),
                            (DEST_X0 + DEST_W, dest_curr_y - 18.0), (DEST_X0, dest_curr_y - 18.0)],
                           close=True, dxfattribs={"layer": "0_TITLE_BLOCK"})
        
        msp.add_text(d_name, dxfattribs={"layer": "TITLE_TEXT", "height": 5.0}).set_placement((DEST_X0 + 10.0, dest_curr_y - 13.0))
        msp.add_text(f"{total_pts} PTS", dxfattribs={"layer": "TEXT_MAIN", "height": 4.5}).set_placement((DEST_X0 + DEST_W - 55.0, dest_curr_y - 13.0))
        
        msp.add_text(full_name, dxfattribs={"layer": "TEXT_MAIN", "height": 3.4}).set_placement((DEST_X0 + 10.0, d_box_y + 17.0))
        msp.add_text(f"LOC: {loc_str}", dxfattribs={"layer": "TEXT_SUB", "height": 3.0}).set_placement((DEST_X0 + 10.0, d_box_y + 7.0))
        
        dest_ports[d_name] = (DEST_X0, d_box_y + (DEST_H / 2.0))
        dest_curr_y -= DEST_H + DEST_GAP

    # -------------------------------------------------------------
    # MIDDLE REGION: INTERCONNECTION TRUNKS & ROUTING CORRIDORS
    # Connecting Chassis Ports to Destination JB Ports
    # Corridor X ranges: X=840 to X=1370
    # -------------------------------------------------------------
    # Group connections by (chassis_id, dest_name)
    ch_dest_groups = defaultdict(list)
    for (cid, slot_no, d_name), (px, py, io_type, pts) in chassis_ports.items():
        ch_dest_groups[(cid, d_name)].append((slot_no, px, py, io_type, pts))
        
    # Draw bundled engineering trunks with orthogonal routing
    trunk_x_step = (1350.0 - 870.0) / (len(DEST_DEFS) + 1)
    dest_trunk_x = {}
    for idx, (d_name, _, _, _) in enumerate(DEST_DEFS):
        dest_trunk_x[d_name] = 870.0 + idx * trunk_x_step
        
    for (cid, d_name), slot_conns in ch_dest_groups.items():
        dx, dy = dest_ports[d_name]
        trx = dest_trunk_x[d_name]
        
        # Pick dominant signal type layer
        type_counts = Counter(s[3] for s in slot_conns)
        dom_type = type_counts.most_common(1)[0][0]
        layer_name = f"ROUTING_{dom_type}" if f"ROUTING_{dom_type}" in [l[0] for l in layers] else "ROUTING_DI"
        
        tot_trunk_pts = sum(s[4] for s in slot_conns)
        
        # Connect each slot to the vertical trunk line
        for slot_no, px, py, io_t, pts in slot_conns:
            msp.add_lwpolyline([(px, py), (trx, py)], dxfattribs={"layer": layer_name})
            # Draw junction dot at slot port
            msp.add_circle((px, py), radius=1.2, dxfattribs={"layer": layer_name})
            
        # Draw vertical trunk segment from min_y to max_y (and into dy)
        all_ys = [s[2] for s in slot_conns] + [dy]
        min_y, max_y = min(all_ys), max(all_ys)
        msp.add_lwpolyline([(trx, min_y), (trx, max_y)], dxfattribs={"layer": layer_name})
        
        # Draw feed from trunk into Destination JB
        msp.add_lwpolyline([(trx, dy), (dx, dy)], dxfattribs={"layer": layer_name})
        msp.add_circle((dx, dy), radius=1.5, dxfattribs={"layer": layer_name})
        
        # Trunk Annotation Text
        slots_str = ", ".join([s[0] for s in slot_conns])
        ann_text = f"{cid} [S:{slots_str}] -> {tot_trunk_pts}p"
        ann_y = dy + 2.5 if dy > 500 else dy - 4.5
        msp.add_text(ann_text, dxfattribs={"layer": "TEXT_ANNOT", "height": 2.4}).set_placement((trx + 2.0, ann_y))

    # -------------------------------------------------------------
    # LEGEND & COLOR CODE BLOCK (X=860 to X=1350, Y=45 to Y=125)
    # -------------------------------------------------------------
    leg_x0 = 860.0
    leg_y0 = 125.0
    leg_w = 480.0
    leg_h = 75.0
    msp.add_lwpolyline([(leg_x0, leg_y0), (leg_x0 + leg_w, leg_y0),
                        (leg_x0 + leg_w, leg_y0 - leg_h), (leg_x0, leg_y0 - leg_h)],
                       close=True, dxfattribs={"layer": "0_BORDER"})
    msp.add_text("SIGNAL TYPE & WIRING CONVENTION LEGEND",
                 dxfattribs={"layer": "TITLE_TEXT", "height": 3.8}).set_placement((leg_x0 + 10.0, leg_y0 - 12.0))
    
    legend_items = [
        ("DI (Discrete Input - 24VDC)", "ROUTING_DI", "Cyan (Layer ROUTING_DI)"),
        ("DO (Discrete Output - 24VDC)", "ROUTING_DO", "Yellow (Layer ROUTING_DO)"),
        ("AI (Analog Input - 4-20mA / RTD)", "ROUTING_AI", "Red (Layer ROUTING_AI)"),
        ("AO (Analog Output - 4-20mA)", "ROUTING_AO", "Magenta (Layer ROUTING_AO)"),
        ("BUS (Fieldbus / VFD Comm)", "ROUTING_BUS", "Blue (Layer ROUTING_BUS)"),
        ("ETH (EtherNet/IP DLR Links)", "ROUTING_ETH", "Gray (Layer ROUTING_ETH)")
    ]
    for idx, (lbl, l_name, desc) in enumerate(legend_items):
        col_i = idx % 2
        row_i = idx // 2
        lx = leg_x0 + 15.0 + col_i * 230.0
        ly = leg_y0 - 24.0 - row_i * 15.0
        msp.add_line((lx, ly), (lx + 30.0, ly), dxfattribs={"layer": l_name})
        msp.add_circle((lx + 15.0, ly), radius=1.5, dxfattribs={"layer": l_name})
        msp.add_text(lbl, dxfattribs={"layer": "TEXT_MAIN", "height": 3.0}).set_placement((lx + 36.0, ly - 1.2))

    # -------------------------------------------------------------
    # TITLE BLOCK (Bottom Right Corner: X=1380 to X=1775, Y=30 to Y=125)
    # -------------------------------------------------------------
    tb_x0 = 1380.0
    tb_y0 = 125.0
    tb_w = 370.0
    tb_h = 75.0
    msp.add_lwpolyline([(tb_x0, tb_y0), (tb_x0 + tb_w, tb_y0),
                        (tb_x0 + tb_w, tb_y0 - tb_h), (tb_x0, tb_y0 - tb_h)],
                       close=True, dxfattribs={"layer": "0_BORDER"})
    
    # Internal Dividers
    msp.add_line((tb_x0, tb_y0 - 25.0), (tb_x0 + tb_w, tb_y0 - 25.0), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tb_x0, tb_y0 - 50.0), (tb_x0 + tb_w, tb_y0 - 50.0), dxfattribs={"layer": "0_TITLE_BLOCK"})
    msp.add_line((tb_x0 + 220.0, tb_y0 - 25.0), (tb_x0 + 220.0, tb_y0 - tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})
    
    msp.add_text("KALASIN ENGINEERING xCIP TURNKEY PROJECT",
                 dxfattribs={"layer": "TITLE_TEXT", "height": 4.2}).set_placement((tb_x0 + 10.0, tb_y0 - 10.0))
    msp.add_text("CLIENT: INGREDION (THAILAND) CO., LTD. / ISHITON",
                 dxfattribs={"layer": "TEXT_SUB", "height": 3.0}).set_placement((tb_x0 + 10.0, tb_y0 - 18.0))
    
    msp.add_text("DRAWING TITLE:", dxfattribs={"layer": "TEXT_ANNOT", "height": 2.5}).set_placement((tb_x0 + 10.0, tb_y0 - 32.0))
    msp.add_text("C1-C8 CHASSIS & SLOT TO DESTINATION JB ARCHITECTURE",
                 dxfattribs={"layer": "TEXT_MAIN", "height": 3.4}).set_placement((tb_x0 + 10.0, tb_y0 - 42.0))
    
    msp.add_text("DWG NO: KAL-SYS-DIA-C1C8-001", dxfattribs={"layer": "TITLE_TEXT", "height": 3.6}).set_placement((tb_x0 + 10.0, tb_y0 - 62.0))
    msp.add_text("PROJECT REF: x2608003", dxfattribs={"layer": "TEXT_SUB", "height": 3.0}).set_placement((tb_x0 + 10.0, tb_y0 - 70.0))
    
    msp.add_text("REV: 1.0 (APPROVED)", dxfattribs={"layer": "TITLE_TEXT", "height": 3.6}).set_placement((tb_x0 + 230.0, tb_y0 - 35.0))
    msp.add_text("SCALE: 1:1 mm MODEL", dxfattribs={"layer": "TEXT_MAIN", "height": 3.0}).set_placement((tb_x0 + 230.0, tb_y0 - 44.0))
    msp.add_text("DATE: 2026-09-29", dxfattribs={"layer": "TEXT_SUB", "height": 3.0}).set_placement((tb_x0 + 230.0, tb_y0 - 62.0))

    # Save DXF
    print(f"Saving AutoCAD DXF to: {DXF_OUTPUT}")
    doc.saveas(DXF_OUTPUT)
    print("AutoCAD DXF generated successfully.")

def export_svg():
    print("Generating High-Fidelity SVG Vector Graphic...")
    # Read DXF or create direct SVG representation matching DXF coordinates
    # Dimensions: 1800 x 1150
    svg_lines = [
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1800 1150" width="100%" height="100%" style="background:#0F172A; font-family: -apple-system, BlinkMacSystemFont, Segoe UI, Roboto, Helvetica, Arial, sans-serif;">',
        '<defs>',
        '  <style>',
        '    .title { fill: #FACC15; font-weight: bold; }',
        '    .main-text { fill: #F8FAFC; }',
        '    .sub-text { fill: #38BDF8; }',
        '    .annot-text { fill: #94A3B8; font-size: 8px; }',
        '    .border-line { stroke: #475569; stroke-width: 1.5; fill: none; }',
        '    .chassis-box { fill: #1E293B; stroke: #64748B; stroke-width: 1.5; rx: 4; }',
        '    .slot-box { fill: #0F172A; stroke: #334155; stroke-width: 1; rx: 2; }',
        '    .dest-box { fill: #1E293B; stroke: #10B981; stroke-width: 1.5; rx: 4; }',
        '    .dest-header { fill: #064E3B; stroke: #10B981; stroke-width: 1; }',
        '    .route-di { stroke: #38BDF8; stroke-width: 1.2; fill: none; opacity: 0.85; }',
        '    .route-do { stroke: #FACC15; stroke-width: 1.2; fill: none; opacity: 0.85; }',
        '    .route-ai { stroke: #F87171; stroke-width: 1.2; fill: none; opacity: 0.85; }',
        '    .route-ao { stroke: #E879F9; stroke-width: 1.2; fill: none; opacity: 0.85; }',
        '    .route-bus { stroke: #818CF8; stroke-width: 1.2; fill: none; opacity: 0.85; }',
        '    .dot-di { fill: #38BDF8; }',
        '    .dot-do { fill: #FACC15; }',
        '    .dot-ai { fill: #F87171; }',
        '    .dot-ao { fill: #E879F9; }',
        '    .dot-bus { fill: #818CF8; }',
        '  </style>',
        '</defs>'
    ]
    
    # Outer & Inner Border
    svg_lines.append('<rect x="25" y="25" width="1750" height="1100" fill="none" stroke="#64748B" stroke-width="2"/>')
    svg_lines.append('<rect x="30" y="30" width="1740" height="1090" fill="none" stroke="#334155" stroke-width="1"/>')
    
    # Top Banner
    svg_lines.append('<rect x="30" y="30" width="1740" height="50" fill="#1E293B"/>')
    svg_lines.append('<text x="50" y="52" class="title" font-size="16">KALASIN ENGINEERING xCIP-1545 SYSTEM (PROJECT REF: x2608003)</text>')
    svg_lines.append('<text x="50" y="70" class="main-text" font-size="12">CHASSIS C1 - C8 TO DESTINATION JUNCTION BOX ARCHITECTURE &amp; INTERCONNECTION DIAGRAM</text>')
    
    # Coordinate converter for SVG (invert Y because SVG (0,0) is top-left)
    def to_svg(x, y):
        return x, 1150.0 - y

    # Chassis Layout in SVG
    row_y_starts = [1150 - 45 - 25, 1150 - 315 - 25, 1150 - 585 - 25, 1150 - 855 - 25]
    col_x_offsets = {
        'C1': (50.0, 95.0),
        'C2': (440.0, 95.0),
        'C3': (50.0, 365.0),
        'C4': (440.0, 365.0),
        'C5': (50.0, 635.0),
        'C6': (440.0, 635.0),
        'C7': (50.0, 905.0),
        'C8': (440.0, 905.0)
    }
    
    CH_W = 370.0
    CH_HDR_H = 26.0
    SLOT_H = 17.0
    
    svg_ch_ports = {}
    
    for ch in CHASSIS_DEFS:
        cid = ch['id']
        x0, y0 = col_x_offsets[cid]
        num_s = len(ch['slots'])
        tot_h = CH_HDR_H + num_s * (SLOT_H + 1.5) + 6.0
        
        # Chassis box
        svg_lines.append(f'<rect x="{x0}" y="{y0}" width="{CH_W}" height="{tot_h}" class="chassis-box"/>')
        svg_lines.append(f'<rect x="{x0}" y="{y0}" width="{CH_W}" height="{CH_HDR_H}" fill="#334155" rx="3"/>')
        svg_lines.append(f'<text x="{x0+10}" y="{y0+14}" class="title" font-size="10">{ch["title"]}</text>')
        svg_lines.append(f'<text x="{x0+10}" y="{y0+23}" class="sub-text" font-size="8">LOC: {ch["loc"]}</text>')
        
        sy = y0 + CH_HDR_H + 4.0
        for s_no, card, io_t, d_list in ch['slots']:
            svg_lines.append(f'<rect x="{x0+4}" y="{sy}" width="{CH_W-8}" height="{SLOT_H}" class="slot-box"/>')
            svg_lines.append(f'<text x="{x0+8}" y="{sy+12}" class="main-text" font-size="9" font-weight="bold">SLOT {s_no}</text>')
            svg_lines.append(f'<text x="{x0+52}" y="{sy+12}" class="sub-text" font-size="8.5">{card}</text>')
            
            # Badge
            badge_cls = f"route-{io_t.lower()}" if io_t.lower() in ['di', 'do', 'ai', 'ao', 'bus'] else "main-text"
            svg_lines.append(f'<text x="{x0+135}" y="{sy+12}" class="{badge_cls}" font-size="8.5" font-weight="bold">[{io_t}]</text>')
            
            d_sum = ", ".join([f"{d} ({p}p)" for d, p in d_list])
            if len(d_sum) > 28: d_sum = d_sum[:26] + ".."
            svg_lines.append(f'<text x="{x0+180}" y="{sy+12}" class="annot-text">-> {d_sum}</text>')
            
            port_x = x0 + CH_W - 4.0
            port_y = sy + (SLOT_H / 2.0)
            for d_name, pts in d_list:
                svg_ch_ports[(cid, s_no, d_name)] = (port_x, port_y, io_t, pts)
                
            sy += SLOT_H + 1.5

    # Destination JBs in SVG
    DEST_X = 1380.0
    DEST_W = 370.0
    DEST_H = 46.0
    DEST_GAP = 13.0
    
    svg_dest_ports = {}
    dy_curr = 95.0
    
    for d_name, full_name, loc_str, total_pts in DEST_DEFS:
        svg_lines.append(f'<rect x="{DEST_X}" y="{dy_curr}" width="{DEST_W}" height="{DEST_H}" class="dest-box"/>')
        svg_lines.append(f'<rect x="{DEST_X}" y="{dy_curr}" width="{DEST_W}" height="18" class="dest-header"/>')
        svg_lines.append(f'<text x="{DEST_X+10}" y="{dy_curr+13}" class="title" font-size="11">{d_name}</text>')
        svg_lines.append(f'<text x="{DEST_X+DEST_W-60}" y="{dy_curr+13}" class="main-text" font-size="10" font-weight="bold">{total_pts} PTS</text>')
        svg_lines.append(f'<text x="{DEST_X+10}" y="{dy_curr+30}" class="main-text" font-size="9">{full_name}</text>')
        svg_lines.append(f'<text x="{DEST_X+10}" y="{dy_curr+41}" class="sub-text" font-size="8">LOC: {loc_str}</text>')
        
        svg_dest_ports[d_name] = (DEST_X, dy_curr + (DEST_H / 2.0))
        dy_curr += DEST_H + DEST_GAP

    # Interconnection lines in SVG
    ch_d_groups = defaultdict(list)
    for (cid, s_no, d_name), (px, py, io_t, pts) in svg_ch_ports.items():
        ch_d_groups[(cid, d_name)].append((s_no, px, py, io_t, pts))
        
    trunk_step = (1350.0 - 870.0) / (len(DEST_DEFS) + 1)
    d_trunk_x = {d[0]: (870.0 + idx * trunk_step) for idx, d in enumerate(DEST_DEFS)}
    
    for (cid, d_name), conns in ch_d_groups.items():
        dx, dy = svg_dest_ports[d_name]
        trx = d_trunk_x[d_name]
        
        t_counts = Counter(c[3] for c in conns)
        dom_t = t_counts.most_common(1)[0][0].lower()
        cls_name = f"route-{dom_t}" if dom_t in ['di', 'do', 'ai', 'ao', 'bus'] else "route-di"
        dot_cls = f"dot-{dom_t}" if dom_t in ['di', 'do', 'ai', 'ao', 'bus'] else "dot-di"
        
        tot_pts = sum(c[4] for c in conns)
        
        for s_no, px, py, io_t, pts in conns:
            svg_lines.append(f'<path d="M {px} {py} L {trx} {py}" class="{cls_name}"/>')
            svg_lines.append(f'<circle cx="{px}" cy="{py}" r="2" class="{dot_cls}"/>')
            
        all_ys = [c[2] for c in conns] + [dy]
        min_y, max_y = min(all_ys), max(all_ys)
        svg_lines.append(f'<path d="M {trx} {min_y} L {trx} {max_y}" class="{cls_name}"/>')
        svg_lines.append(f'<path d="M {trx} {dy} L {dx} {dy}" class="{cls_name}"/>')
        svg_lines.append(f'<circle cx="{dx}" cy="{dy}" r="2.5" class="{dot_cls}"/>')
        
        slots_str = ", ".join([c[0] for c in conns])
        ann_text = f"{cid} [S:{slots_str}] -> {tot_pts}p"
        ann_y = dy - 3 if dy < 600 else dy + 8
        svg_lines.append(f'<text x="{trx+3}" y="{ann_y}" class="annot-text">{ann_text}</text>')

    # Legend in SVG
    svg_lines.append('<rect x="860" y="1035" width="480" height="75" fill="#1E293B" stroke="#475569" stroke-width="1.2" rx="4"/>')
    svg_lines.append('<text x="875" y="1052" class="title" font-size="10">SIGNAL TYPE &amp; WIRING CONVENTION LEGEND</text>')
    
    leg_items = [
        ("DI (Discrete Input - 24VDC)", "route-di", "dot-di"),
        ("DO (Discrete Output - 24VDC)", "route-do", "dot-do"),
        ("AI (Analog Input - 4-20mA / RTD)", "route-ai", "dot-ai"),
        ("AO (Analog Output - 4-20mA)", "route-ao", "dot-ao"),
        ("BUS (Fieldbus / VFD Comm)", "route-bus", "dot-bus"),
        ("ETH (EtherNet/IP DLR Links)", "route-eth", "dot-eth")
    ]
    for idx, (lbl, l_cls, dot_c) in enumerate(leg_items):
        ci = idx % 2
        ri = idx // 2
        lx = 875.0 + ci * 240.0
        ly = 1068.0 + ri * 15.0
        svg_lines.append(f'<path d="M {lx} {ly} L {lx+30} {ly}" class="{l_cls}"/>')
        svg_lines.append(f'<circle cx="{lx+15}" cy="{ly}" r="2" class="{dot_c}"/>')
        svg_lines.append(f'<text x="{lx+36}" y="{ly+3.5}" class="main-text" font-size="8.5">{lbl}</text>')

    # Title Block in SVG
    svg_lines.append('<rect x="1380" y="1035" width="370" height="75" fill="#1E293B" stroke="#64748B" stroke-width="1.2" rx="4"/>')
    svg_lines.append('<line x1="1380" y1="1060" x2="1750" y2="1060" stroke="#334155" stroke-width="1"/>')
    svg_lines.append('<line x1="1380" y1="1085" x2="1750" y2="1085" stroke="#334155" stroke-width="1"/>')
    svg_lines.append('<line x1="1600" y1="1060" x2="1600" y2="1110" stroke="#334155" stroke-width="1"/>')
    
    svg_lines.append('<text x="1390" y="1050" class="title" font-size="10">KALASIN ENGINEERING xCIP TURNKEY PROJECT</text>')
    svg_lines.append('<text x="1390" y="1057" class="sub-text" font-size="7.5">CLIENT: INGREDION (THAILAND) CO., LTD. / ISHITON</text>')
    
    svg_lines.append('<text x="1390" y="1070" class="annot-text">DRAWING TITLE:</text>')
    svg_lines.append('<text x="1390" y="1080" class="main-text" font-size="9" font-weight="bold">C1-C8 CHASSIS &amp; SLOT TO DESTINATION JB ARCHITECTURE</text>')
    
    svg_lines.append('<text x="1390" y="1098" class="title" font-size="9.5">DWG NO: KAL-SYS-DIA-C1C8-001</text>')
    svg_lines.append('<text x="1390" y="1106" class="sub-text" font-size="8">PROJECT REF: x2608003</text>')
    
    svg_lines.append('<text x="1610" y="1074" class="title" font-size="9">REV: 1.0 (APPROVED)</text>')
    svg_lines.append('<text x="1610" y="1095" class="sub-text" font-size="8">SCALE: 1:1 MODEL (A0)</text>')
    svg_lines.append('<text x="1610" y="1105" class="sub-text" font-size="8">DATE: 2026-09-29</text>')

    svg_lines.append('</svg>')
    
    print(f"Saving Scalable Vector Graphics to: {SVG_OUTPUT}")
    with open(SVG_OUTPUT, "w", encoding="utf-8") as f:
        f.write("\n".join(svg_lines))
    print("SVG generated successfully.")

if __name__ == "__main__":
    generate_cad()
    export_svg()
