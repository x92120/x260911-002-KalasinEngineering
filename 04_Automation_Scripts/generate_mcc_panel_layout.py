#!/usr/bin/env python3
"""
Generate complete engineering panel layout and mounting plan drawings for
MCC / CA1 4-Panel Suite (Panel M1, M2, M3, M4) for Project Jet Cooker.
Produces a publication-ready A3 Landscape engineering PDF drawing package and high-res PNG images.
"""

import os
import sys
import pandas as pd
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import fitz

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(WORKSPACE_DIR, "reports_pdf")
os.makedirs(OUTPUT_DIR, exist_ok=True)

PDF_FILENAME = os.path.join(OUTPUT_DIR, "MCC_Panel_Layout_Mounting_Plan_M1_M2_M3_M4.pdf")

class EngineeringDrawing(FPDF):
    def __init__(self):
        super().__init__(orientation="landscape", unit="mm", format="A3")
        self.set_margins(10, 10, 10)
        self.set_auto_page_break(auto=False)
        self.alias_nb_pages()

    def draw_drawing_frame(self, sheet_no, total_sheets, sheet_title, scale="1:12"):
        margin = 10
        w = 420 - 2 * margin # 400 mm
        h = 297 - 2 * margin # 277 mm
        
        # Outer border
        self.set_draw_color(30, 40, 60)
        self.set_line_width(0.7)
        self.rect(margin, margin, w, h)
        
        # Grid references
        self.set_font("Helvetica", "", 6)
        self.set_text_color(100, 110, 130)
        for i, label in enumerate(["1", "2", "3", "4", "5", "6", "7", "8"]):
            x = margin + (i + 0.5) * (w / 8)
            self.text(x, margin - 2, label)
            self.text(x, margin + h + 4, label)
        for i, label in enumerate(["A", "B", "C", "D", "E", "F"]):
            y = margin + (i + 0.5) * (h / 6)
            self.text(margin - 4, y, label)
            self.text(margin + w + 2, y, label)

        # Standard Engineering Title Block
        tb_w = 165
        tb_h = 36
        tb_x = margin + w - tb_w
        tb_y = margin + h - tb_h
        
        # Title block background
        self.set_fill_color(255, 255, 255)
        self.set_draw_color(30, 40, 60)
        self.set_line_width(0.5)
        self.rect(tb_x, tb_y, tb_w, tb_h, "DF")
        
        # Divisions
        self.line(tb_x, tb_y + 10, tb_x + tb_w, tb_y + 10)
        self.line(tb_x, tb_y + 22, tb_x + tb_w, tb_y + 22)
        self.line(tb_x + 105, tb_y, tb_x + 105, tb_y + 22)
        self.line(tb_x + 50, tb_y + 22, tb_x + 50, tb_y + tb_h)
        self.line(tb_x + 105, tb_y + 22, tb_x + 105, tb_y + tb_h)
        self.line(tb_x + 135, tb_y + 22, tb_x + 135, tb_y + tb_h)

        # Logo / Client block
        self.set_text_color(24, 43, 73)
        self.set_font("Helvetica", "B", 8)
        self.text(tb_x + 3, tb_y + 4.5, "INGREDION (THAILAND) CO., LTD.")
        self.set_font("Helvetica", "", 7)
        self.text(tb_x + 3, tb_y + 8, "KALASIN STARCH PLANT - JET COOKER PROJECT")

        # Contractor block
        self.set_font("Helvetica", "B", 8)
        self.text(tb_x + 108, tb_y + 4.5, "AEC INDUSTRIAL ENGINEERING")
        self.set_font("Helvetica", "", 6.5)
        self.text(tb_x + 108, tb_y + 8, "SYSTEMS INTEGRATION & AUTOMATION")

        # Document Title (Truncate or scale to fit 100mm)
        self.set_font("Helvetica", "B", 8)
        self.set_text_color(15, 30, 60)
        
        # Word wrap title if too long
        if self.get_string_width(sheet_title) > 100:
            self.set_font("Helvetica", "B", 7.2)
        self.text(tb_x + 3, tb_y + 15, sheet_title)
        
        self.set_font("Helvetica", "", 6.5)
        self.set_text_color(80, 90, 100)
        self.text(tb_x + 3, tb_y + 19.5, "MAIN CONTROL CABINET CA1 / MOTOR CONTROL PANELS M1..M4")

        # Status / Doc No / Revision
        self.set_font("Helvetica", "B", 7)
        self.set_text_color(30, 40, 60)
        self.text(tb_x + 108, tb_y + 14, "DOC NO:")
        self.set_font("Helvetica", "", 7)
        self.text(tb_x + 125, tb_y + 14, "CA1-DWG-PL-00" + str(sheet_no))
        self.set_font("Helvetica", "B", 7)
        self.text(tb_x + 108, tb_y + 19, "REV: 3.6")
        self.text(tb_x + 128, tb_y + 19, "DATE: 2026-09-07")

        # Sub-cells
        self.set_font("Helvetica", "B", 6.5)
        self.text(tb_x + 3, tb_y + 26, "SCALE:")
        self.set_font("Helvetica", "", 7)
        self.text(tb_x + 18, tb_y + 26, scale)

        self.set_font("Helvetica", "B", 6.5)
        self.text(tb_x + 3, tb_y + 32, "PROJECT:")
        self.set_font("Helvetica", "", 7)
        self.text(tb_x + 18, tb_y + 32, "JET COOKER")

        self.set_font("Helvetica", "B", 6.5)
        self.text(tb_x + 53, tb_y + 26, "DRAWN:")
        self.set_font("Helvetica", "", 7)
        self.text(tb_x + 72, tb_y + 26, "AEC-ENG")

        self.set_font("Helvetica", "B", 6.5)
        self.text(tb_x + 53, tb_y + 32, "CHECKED:")
        self.set_font("Helvetica", "", 7)
        self.text(tb_x + 72, tb_y + 32, "LEAD-PE")

        self.set_font("Helvetica", "B", 6.5)
        self.text(tb_x + 108, tb_y + 26, "STATUS:")
        self.set_font("Helvetica", "B", 7)
        self.set_text_color(20, 120, 40)
        self.text(tb_x + 108, tb_y + 32, "APPROVED")
        self.set_text_color(30, 40, 60)

        self.set_font("Helvetica", "B", 7)
        self.text(tb_x + 138, tb_y + 26, "SHEET:")
        self.set_font("Helvetica", "B", 10)
        self.set_text_color(24, 43, 73)
        self.text(tb_x + 142, tb_y + 33, f"{sheet_no} / {total_sheets}")

def draw_legend_box(pdf, x, y, w, h, items):
    pdf.set_fill_color(250, 252, 255)
    pdf.set_draw_color(180, 195, 215)
    pdf.set_line_width(0.3)
    pdf.rect(x, y, w, h, "DF")
    
    pdf.set_fill_color(35, 55, 85)
    pdf.rect(x, y, w, 5.5, "F")
    pdf.set_font("Helvetica", "B", 7)
    pdf.set_text_color(255, 255, 255)
    pdf.text(x + 3, y + 4, "NOTES & DESIGN CRITERIA")
    
    pdf.set_font("Helvetica", "", 6.5)
    pdf.set_text_color(40, 50, 65)
    cur_y = y + 9.5
    for item in items:
        pdf.text(x + 3, cur_y, item)
        cur_y += 4.5

# -------------------------------------------------------------
# SHEET 1: FRONT ELEVATION & GENERAL ARRANGEMENT (OUTER DOORS)
# -------------------------------------------------------------
def build_sheet_1(pdf):
    pdf.add_page()
    pdf.draw_drawing_frame(1, 7, "4-PANEL SUITE - FRONT ELEVATION (ENCLOSURE EXTERNAL VIEW)", scale="1:15")
    
    start_x = 35
    start_y = 50
    bay_w = 800 / 15 # 53.33 mm
    body_h = 2000 / 15 # 133.33 mm
    plinth_h = 100 / 15 # 6.67 mm
    total_w = bay_w * 4
    
    # Plinth (100mm)
    pdf.set_fill_color(50, 55, 65)
    pdf.set_draw_color(30, 35, 45)
    pdf.set_line_width(0.5)
    pdf.rect(start_x, start_y + body_h, total_w, plinth_h, "DF")
    
    for i in range(4):
        px = start_x + i * bay_w
        pdf.set_fill_color(70, 75, 85)
        pdf.rect(px + 6, start_y + body_h + 1.5, bay_w - 12, plinth_h - 3, "DF")
        pdf.set_font("Helvetica", "", 5)
        pdf.set_text_color(200, 210, 220)
        pdf.text(px + bay_w/2 - 7, start_y + body_h + 4.8, "PLINTH 100mm")
    
    bay_names = ["PANEL M1\n(MASTER CONTROLLER)", "PANEL M2\n(I/O EXPANSION 1)", "PANEL M3\n(I/O & IS BARRIERS)", "PANEL M4\n(I/O & IS BARRIERS)"]
    
    for i in range(4):
        bx = start_x + i * bay_w
        pdf.set_fill_color(228, 232, 238)
        pdf.set_draw_color(60, 70, 85)
        pdf.set_line_width(0.6)
        pdf.rect(bx, start_y, bay_w, body_h, "DF")
        
        pdf.set_fill_color(238, 241, 246)
        pdf.set_draw_color(120, 130, 145)
        pdf.set_line_width(0.3)
        door_x = bx + 2
        door_y = start_y + 2
        door_w = bay_w - 4
        door_h = body_h - 4
        pdf.rect(door_x, door_y, door_w, door_h, "DF")
        
        # Hinges
        pdf.set_fill_color(90, 100, 115)
        pdf.rect(door_x, door_y + 15, 1.2, 5, "F")
        pdf.rect(door_x, door_y + door_h/2 - 2.5, 1.2, 5, "F")
        pdf.rect(door_x, door_y + door_h - 20, 1.2, 5, "F")
        
        # Handle
        handle_x = door_x + door_w - 3.5
        handle_y = door_y + door_h/2 - 8
        pdf.set_fill_color(40, 45, 55)
        pdf.rect(handle_x, handle_y, 2.2, 16, "F")
        pdf.set_fill_color(180, 190, 200)
        pdf.circle(handle_x + 1.1, handle_y + 8, 0.7, "F")

        # Eyebolts
        pdf.set_draw_color(70, 80, 95)
        pdf.set_line_width(0.4)
        pdf.circle(bx + 8, start_y - 2.5, 1.8)
        pdf.circle(bx + bay_w - 8, start_y - 2.5, 1.8)

        # Roof Fan
        pdf.set_fill_color(190, 200, 212)
        pdf.rect(bx + bay_w/2 - 12, start_y + 4, 24, 7, "DF")
        pdf.set_font("Helvetica", "", 4.5)
        pdf.set_text_color(60, 70, 80)
        pdf.text(bx + bay_w/2 - 9, start_y + 8.5, "ROOF EXHAUST FAN")

        # Header Plate
        pdf.set_fill_color(24, 43, 73)
        pdf.rect(door_x + 3, door_y + 14, door_w - 6, 7, "F")
        pdf.set_font("Helvetica", "B", 6)
        pdf.set_text_color(255, 255, 255)
        pname = bay_names[i].split("\n")[0]
        pdf.text(door_x + door_w/2 - pdf.get_string_width(pname)/2, door_y + 18.5, pname)

        if i == 0:
            # HMI Touchscreen
            hmi_w = 26
            hmi_h = 20
            hmi_x = door_x + (door_w - hmi_w)/2
            hmi_y = door_y + 26
            pdf.set_fill_color(30, 35, 45)
            pdf.rect(hmi_x, hmi_y, hmi_w, hmi_h, "DF")
            pdf.set_fill_color(70, 130, 180)
            pdf.rect(hmi_x + 2, hmi_y + 2, hmi_w - 4, hmi_h - 4, "F")
            pdf.set_font("Helvetica", "B", 4.5)
            pdf.set_text_color(255, 255, 255)
            pdf.text(hmi_x + 4, hmi_y + 9, "PanelView Plus 7")
            pdf.text(hmi_x + 4, hmi_y + 13, "JET COOKER SCADA")

            # Pilot Lights
            pl_y = hmi_y + hmi_h + 8
            colors = [(220, 50, 50), (230, 200, 30), (50, 120, 220)]
            labels = ["R", "S", "T"]
            for c_idx, (r, g, b) in enumerate(colors):
                lx = door_x + 12 + c_idx * 9
                pdf.set_fill_color(r, g, b)
                pdf.circle(lx, pl_y, 2.2, "DF")
                pdf.set_font("Helvetica", "B", 4)
                pdf.set_text_color(40, 40, 40)
                pdf.text(lx - 1, pl_y + 4.5, labels[c_idx])

            # Push Buttons
            pb_y = pl_y + 12
            pb_items = [
                ((40, 160, 60), "START"),
                ((200, 40, 40), "STOP"),
                ((40, 100, 220), "RESET"),
                ((230, 230, 230), "TEST")
            ]
            for p_idx, ((r, g, b), lbl) in enumerate(pb_items):
                px = door_x + 8 + p_idx * 9
                pdf.set_fill_color(r, g, b)
                pdf.circle(px, pb_y, 2.2, "DF")
                pdf.set_font("Helvetica", "", 3.5)
                pdf.set_text_color(40, 40, 40)
                pdf.text(px - pdf.get_string_width(lbl)/2, pb_y + 4.5, lbl)

            # E-Stop Button
            es_x = door_x + door_w/2
            es_y = pb_y + 14
            pdf.set_fill_color(245, 215, 30)
            pdf.circle(es_x, es_y, 4.5, "F")
            pdf.set_fill_color(210, 30, 30)
            pdf.circle(es_x, es_y, 3, "DF")
            pdf.set_font("Helvetica", "B", 4)
            pdf.set_text_color(180, 20, 20)
            pdf.text(es_x - 5, es_y + 6.5, "E-STOP")

            # Main Disconnect
            sw_y = es_y + 16
            pdf.set_fill_color(245, 215, 30)
            pdf.rect(es_x - 4, sw_y - 4, 8, 8, "F")
            pdf.set_fill_color(190, 30, 30)
            pdf.rect(es_x - 1.2, sw_y - 4, 2.4, 8, "F")
            pdf.set_font("Helvetica", "B", 4)
            pdf.set_text_color(40, 40, 40)
            pdf.text(es_x - 7, sw_y + 7.5, "MAIN SWITCH")

        else:
            pl_y = door_y + 35
            stat_lights = [
                ((255, 255, 255), "POWER ON"),
                ((40, 160, 60), "HEALTHY"),
                ((230, 150, 20), "ALARM"),
                ((210, 30, 30), "FAULT")
            ]
            for s_idx, ((r, g, b), lbl) in enumerate(stat_lights):
                lx = door_x + 8 + s_idx * 9
                pdf.set_fill_color(r, g, b)
                pdf.circle(lx, pl_y, 2.2, "DF")
                pdf.set_font("Helvetica", "", 3.5)
                pdf.set_text_color(40, 40, 40)
                pdf.text(lx - pdf.get_string_width(lbl)/2, pl_y + 4.5, lbl)

            # Warning Hazard Triangle
            tri_x = door_x + door_w/2
            tri_y = pl_y + 25
            pdf.set_fill_color(245, 215, 30)
            pdf.polygon([(tri_x, tri_y - 6), (tri_x - 6, tri_y + 4), (tri_x + 6, tri_y + 4)], style="F")
            pdf.set_font("Helvetica", "B", 5.5)
            pdf.set_text_color(20, 20, 20)
            pdf.text(tri_x - 0.9, tri_y + 2.5, "!")
            pdf.set_font("Helvetica", "B", 4)
            pdf.text(tri_x - 10, tri_y + 7, "400V / 230V HAZARD")

            # Document Pocket
            doc_h_w = door_w - 14
            doc_h_h = 24
            doc_h_x = door_x + 7
            doc_h_y = door_y + door_h - 38
            pdf.set_fill_color(225, 230, 238)
            pdf.rect(doc_h_x, doc_h_y, doc_h_w, doc_h_h, "DF")
            pdf.set_font("Helvetica", "I", 4.5)
            pdf.set_text_color(100, 110, 120)
            pdf.text(doc_h_x + 3, doc_h_y + 12, "DOCUMENT POCKET A4")

        # Bottom Air Inlet Filter
        pdf.set_fill_color(190, 200, 212)
        pdf.rect(door_x + 6, door_y + door_h - 12, door_w - 12, 8, "DF")
        pdf.set_font("Helvetica", "", 4)
        pdf.set_text_color(70, 80, 90)
        pdf.text(door_x + door_w/2 - 8, door_y + door_h - 7.5, "AIR INLET FILTER")

    # Dimensions
    dim_color = (180, 30, 30)
    pdf.set_draw_color(*dim_color)
    pdf.set_text_color(*dim_color)
    pdf.set_line_width(0.3)
    
    # Bottom Width: 3200 mm
    dim_y = start_y + body_h + plinth_h + 10
    pdf.line(start_x, dim_y, start_x + total_w, dim_y)
    pdf.line(start_x, start_y + body_h + plinth_h + 2, start_x, dim_y + 3)
    pdf.line(start_x + total_w, start_y + body_h + plinth_h + 2, start_x + total_w, dim_y + 3)
    pdf.set_font("Helvetica", "B", 8)
    pdf.text(start_x + total_w/2 - 16, dim_y - 2, "3200 mm (TOTAL SUITE WIDTH)")

    # Per-bay 800 mm
    dim_sub_y = dim_y + 7
    for i in range(4):
        bx = start_x + i * bay_w
        pdf.line(bx, dim_sub_y, bx + bay_w, dim_sub_y)
        pdf.line(bx, dim_y + 2, bx, dim_sub_y + 2)
        pdf.line(bx + bay_w, dim_y + 2, bx + bay_w, dim_sub_y + 2)
        pdf.set_font("Helvetica", "", 6)
        pdf.text(bx + bay_w/2 - 7, dim_sub_y - 1.5, "800 mm")

    # Height: 2000 mm + 100 mm
    dim_x = start_x - 12
    pdf.line(dim_x, start_y, dim_x, start_y + body_h + plinth_h)
    pdf.line(start_x - 3, start_y, dim_x - 3, start_y)
    pdf.line(start_x - 3, start_y + body_h, dim_x - 3, start_y + body_h)
    pdf.line(start_x - 3, start_y + body_h + plinth_h, dim_x - 3, start_y + body_h + plinth_h)
    
    pdf.set_font("Helvetica", "B", 7)
    pdf.text(dim_x - 18, start_y + body_h/2, "2000 mm")
    pdf.text(dim_x - 16, start_y + body_h + plinth_h/2 + 1, "100 mm")

    notes = [
        "1. ENCLOSURE SYSTEM: Rittal VX25 / TS8 Modular Bayed Suite, IP54 rated.",
        "2. FINISH: Powder coated sheet steel RAL 7035 (Light Grey), Plinth RAL 7022 (Umbra Grey).",
        "3. SUPPLY VOLTAGE: 3x 400VAC + N + PE, 50Hz. Control voltage: 24VDC (redundant).",
        "4. CABLE ENTRY: Bottom entry via removable gland plates with EMC cable brass glands.",
        "5. COOLING: Roof-mounted forced ventilation with exhaust fans and washable filters.",
        "6. EARTHING: Continuous copper PE busbar (30x5mm) and isolated clean screen bar.",
        "7. ISOLATION: Complete galvanic segregation between Ex-i (IS) and Non-Ex field wiring."
    ]
    draw_legend_box(pdf, 245, 48, 150, 46, notes)

# -------------------------------------------------------------
# SHEET 2: INTERNAL GENERAL ARRANGEMENT (DOORS OPEN)
# -------------------------------------------------------------
def build_sheet_2(pdf):
    pdf.add_page()
    pdf.draw_drawing_frame(2, 7, "4-PANEL SUITE - INTERNAL GENERAL ARRANGEMENT (DOORS OPEN)", scale="1:15")
    
    start_x = 35
    start_y = 50
    bay_w = 800 / 15
    body_h = 2000 / 15
    plinth_h = 100 / 15
    total_w = bay_w * 4
    
    pdf.set_fill_color(50, 55, 65)
    pdf.rect(start_x, start_y + body_h, total_w, plinth_h, "DF")
    
    rack_models = ["C1: 1756-L950TPSXT (CPU)", "C2: High-Density I/O (13 Slots)", "C3: IS I/O + 8AO (Chassis C3)", "C4: IS I/O (Chassis C4)"]
    jb_destinations = ["JB-401, JB-601, CA1", "JB-602, JB-402, JB-607", "IS-JB-603, IS-JB-618, JB-618", "JB-606, IS-JB-612, JB-612, IS-JB-608"]
    
    for i in range(4):
        bx = start_x + i * bay_w
        pdf.set_fill_color(220, 225, 232)
        pdf.set_draw_color(70, 80, 95)
        pdf.set_line_width(0.5)
        pdf.rect(bx, start_y, bay_w, body_h, "DF")
        
        plate_x = bx + 2.5
        plate_y = start_y + 3
        plate_w = bay_w - 5
        plate_h = body_h - 6
        pdf.set_fill_color(242, 245, 248)
        pdf.rect(plate_x, plate_y, plate_w, plate_h, "DF")
        
        # Wire Ducts
        vd_w = 4
        pdf.set_fill_color(195, 200, 208)
        pdf.set_draw_color(140, 145, 155)
        pdf.set_line_width(0.2)
        pdf.rect(plate_x + 1, plate_y + 4, vd_w, plate_h - 8, "DF")
        pdf.rect(plate_x + plate_w - vd_w - 1, plate_y + 4, vd_w, plate_h - 8, "DF")
        
        hd_h = 3.5
        hd_ys = [plate_y + 5, plate_y + 36, plate_y + 70, plate_y + 98]
        for hy in hd_ys:
            pdf.rect(plate_x + vd_w + 1, hy, plate_w - 2*vd_w - 2, hd_h, "DF")

        # Row 1: Power & Comms
        r1_y = plate_y + 11
        r1_h = 22
        pdf.set_fill_color(210, 215, 220)
        pdf.rect(plate_x + vd_w + 2, r1_y + r1_h/2 - 0.8, plate_w - 2*vd_w - 4, 1.6, "F")
        
        if i == 0:
            pdf.set_fill_color(80, 90, 100)
            pdf.rect(plate_x + vd_w + 3, r1_y + 3, 6, 14, "DF")
            pdf.set_font("Helvetica", "", 3.5)
            pdf.set_text_color(255, 255, 255)
            pdf.text(plate_x + vd_w + 3.5, r1_y + 11, "MCB")
            
            pdf.set_fill_color(40, 70, 120)
            pdf.rect(plate_x + vd_w + 10.5, r1_y + 1, 9, 17, "DF")
            pdf.rect(plate_x + vd_w + 20.5, r1_y + 1, 9, 17, "DF")
            pdf.set_font("Helvetica", "B", 3.5)
            pdf.set_text_color(255, 255, 255)
            pdf.text(plate_x + vd_w + 11.5, r1_y + 10, "24V 20A")
            pdf.text(plate_x + vd_w + 21.5, r1_y + 10, "24V 20A")

            pdf.set_fill_color(50, 130, 80)
            pdf.rect(plate_x + vd_w + 31, r1_y + 2, 8, 15, "DF")
            pdf.set_font("Helvetica", "B", 3.2)
            pdf.text(plate_x + vd_w + 31.5, r1_y + 10, "ETH SW")
        else:
            pdf.set_fill_color(70, 80, 95)
            pdf.rect(plate_x + vd_w + 3, r1_y + 3, 14, 14, "DF")
            pdf.set_font("Helvetica", "", 3.5)
            pdf.set_text_color(255, 255, 255)
            pdf.text(plate_x + vd_w + 4, r1_y + 11, "DC FUSE/CB")
            
            pdf.set_fill_color(50, 130, 80)
            pdf.rect(plate_x + vd_w + 20, r1_y + 2, 10, 15, "DF")
            pdf.set_font("Helvetica", "B", 3.2)
            pdf.text(plate_x + vd_w + 20.5, r1_y + 10, "EN4TR ADPT")

        # Row 2: ControlLogix Chassis C1..C4
        r2_y = plate_y + 42
        r2_w = plate_w - 2*vd_w - 4
        r2_h = 24
        
        pdf.set_fill_color(30, 35, 42)
        pdf.set_draw_color(15, 20, 25)
        pdf.set_line_width(0.3)
        pdf.rect(plate_x + vd_w + 2, r2_y, r2_w, r2_h, "DF")
        
        ps_w = 6.5
        pdf.set_fill_color(180, 40, 40)
        pdf.rect(plate_x + vd_w + 3, r2_y + 1.5, ps_w, r2_h - 3, "DF")
        pdf.set_font("Helvetica", "B", 3)
        pdf.set_text_color(255, 255, 255)
        pdf.text(plate_x + vd_w + 3.5, r2_y + 13, "PA72")

        slot_w = (r2_w - ps_w - 3) / 13
        slot_start_x = plate_x + vd_w + 3.5 + ps_w
        for s in range(13):
            sx = slot_start_x + s * slot_w
            if i == 0 and s == 0:
                card_col = (20, 100, 180)
                card_lbl = "CPU"
            elif i == 0 and s == 1:
                card_col = (130, 135, 140)
                card_lbl = "RES"
            elif (i == 0 and s in [2,3,4,5]) or (i == 1 and s in [0,1,2,3,4]) or (i == 2 and s in [0,1]) or (i == 3 and s in [0,1,2]):
                card_col = (40, 140, 70)
                card_lbl = "DI"
            elif (i == 0 and s in [6,7]) or (i == 1 and s in [5,6,7]) or (i == 2 and s == 2) or (i == 3 and s == 3):
                card_col = (200, 120, 20)
                card_lbl = "DO"
            elif (i == 0 and s in [8,9]) or (i == 1 and s in [8,9,10,11,12]) or (i == 2 and s in [3,4]) or (i == 3 and s in [4,5]):
                card_col = (90, 60, 160)
                card_lbl = "AI"
            elif i == 2 and s == 12:
                card_col = (20, 160, 160)
                card_lbl = "AO"
            else:
                card_col = (90, 95, 105)
                card_lbl = "N2"

            pdf.set_fill_color(*card_col)
            pdf.rect(sx, r2_y + 1.5, slot_w - 0.3, r2_h - 3, "DF")
            pdf.set_font("Helvetica", "B", 2.6)
            pdf.set_text_color(255, 255, 255)
            pdf.text(sx + 0.3, r2_y + 13, card_lbl)

        pdf.set_font("Helvetica", "B", 4.5)
        pdf.set_text_color(20, 40, 70)
        pdf.text(plate_x + vd_w + 3, r2_y - 2, f"RACK {rack_models[i]}")

        # Row 3: Relays / Barriers
        r3_y = plate_y + 76
        r3_h = 19
        pdf.set_fill_color(210, 215, 220)
        pdf.rect(plate_x + vd_w + 2, r3_y + r3_h/2 - 0.8, plate_w - 2*vd_w - 4, 1.6, "F")
        
        if i in [2, 3]:
            pdf.set_fill_color(30, 110, 220)
            barrier_count = 18 if i == 2 else 14
            bar_w = (plate_w - 2*vd_w - 8) / barrier_count
            for b_idx in range(barrier_count):
                bx_pos = plate_x + vd_w + 4 + b_idx * bar_w
                pdf.rect(bx_pos, r3_y + 2, bar_w - 0.4, 15, "DF")
            pdf.set_font("Helvetica", "B", 4)
            pdf.set_text_color(15, 80, 180)
            pdf.text(plate_x + vd_w + 4, r3_y - 1.5, "INTRINSIC SAFETY BARRIERS (Ex-i)")
        else:
            pdf.set_fill_color(120, 125, 135)
            relay_count = 16
            rel_w = (plate_w - 2*vd_w - 8) / relay_count
            for r_idx in range(relay_count):
                rx_pos = plate_x + vd_w + 4 + r_idx * rel_w
                pdf.rect(rx_pos, r3_y + 3, rel_w - 0.4, 13, "DF")
            pdf.set_font("Helvetica", "B", 4)
            pdf.set_text_color(60, 70, 80)
            pdf.text(plate_x + vd_w + 4, r3_y - 1.5, "OUTPUT INTERPOSING RELAYS (24VDC)")

        # Row 4: Terminals
        r4_y = plate_y + 104
        r4_h = 20
        pdf.set_fill_color(210, 215, 220)
        pdf.rect(plate_x + vd_w + 2, r4_y + 9, plate_w - 2*vd_w - 4, 1.6, "F")
        
        if i in [2, 3]:
            term_w = plate_w - 2*vd_w - 6
            pdf.set_fill_color(160, 165, 175)
            pdf.rect(plate_x + vd_w + 3, r4_y + 3, term_w * 0.4, 14, "DF")
            pdf.set_fill_color(230, 80, 20)
            pdf.rect(plate_x + vd_w + 3 + term_w * 0.4, r4_y + 1, 1.5, 17, "DF")
            pdf.set_fill_color(40, 120, 220)
            pdf.rect(plate_x + vd_w + 5 + term_w * 0.4, r4_y + 3, term_w * 0.55, 14, "DF")
            
            pdf.set_font("Helvetica", "B", 3.8)
            pdf.set_text_color(40, 50, 65)
            pdf.text(plate_x + vd_w + 3, r4_y - 1.5, "NON-IS TB")
            pdf.set_text_color(20, 90, 200)
            pdf.text(plate_x + vd_w + 18, r4_y - 1.5, "IS BLUE TERMINALS (Ex-i)")
        else:
            pdf.set_fill_color(160, 165, 175)
            pdf.rect(plate_x + vd_w + 3, r4_y + 3, plate_w - 2*vd_w - 6, 14, "DF")
            pdf.set_font("Helvetica", "B", 4)
            pdf.set_text_color(50, 60, 75)
            pdf.text(plate_x + vd_w + 4, r4_y - 1.5, "FIELD TERMINAL STRIPS (X1, X2)")

        # PE Bar
        pe_y = plate_y + plate_h - 4.5
        pdf.set_fill_color(200, 150, 40)
        pdf.rect(plate_x + 3, pe_y, plate_w - 6, 1.8, "DF")
        pdf.set_font("Helvetica", "B", 3.5)
        pdf.set_text_color(40, 30, 10)
        pdf.text(plate_x + 4, pe_y + 1.4, "PE / EARTH COPPER BAR 30x5mm")

        pdf.set_font("Helvetica", "B", 4.2)
        pdf.set_text_color(24, 43, 73)
        pdf.text(plate_x + 2, plate_y + plate_h + 3.5, f"SERVES: {jb_destinations[i]}")

    notes_sheet2 = [
        "1. CHASSIS SYSTEM: Rockwell Automation Allen-Bradley ControlLogix 1756-A13 (13-slot racks).",
        "2. WIRING SEGREGATION: 50mm min. clearance between 230VAC, 24VDC, and analog signal paths.",
        "3. INTRINSIC SAFETY: Panels M3 & M4 feature dedicated blue IS terminal strips and 50mm physical",
        "   air gap partition plates complying with IEC 60079-11 / IEC 60079-14 Ex-i standards.",
        "4. SPARE CAPACITY: All 4 panels maintain >25% spare DIN rail space and spare card slots.",
        "5. WIRE DUCT: Halogen-free slotted PVC wire duct with snap-on covers (60x80mm & 80x100mm)."
    ]
    draw_legend_box(pdf, 245, 48, 150, 42, notes_sheet2)

# -------------------------------------------------------------
# SHEETS 3 TO 6: LARGE-SCALE DETAILED MOUNTING PLANS (M1..M4)
# -------------------------------------------------------------
def build_detailed_panel_sheet(pdf, sheet_no, panel_id, panel_title, chassis_id, jb_serves, slot_data, is_is_panel=False):
    pdf.add_page()
    pdf.draw_drawing_frame(sheet_no, 7, f"{panel_id} - SUBPANEL MOUNTING PLAN & COMPONENT SCHEDULE", scale="1:8")
    
    start_x = 22
    start_y = 35
    plate_w = 700 / 9
    plate_h = 1900 / 9
    
    frame_w = 800 / 9
    frame_h = 2000 / 9
    pdf.set_fill_color(225, 230, 238)
    pdf.set_draw_color(60, 70, 85)
    pdf.set_line_width(0.7)
    pdf.rect(start_x, start_y, frame_w, frame_h, "DF")
    
    mp_x = start_x + (frame_w - plate_w)/2
    mp_y = start_y + (frame_h - plate_h)/2
    pdf.set_fill_color(244, 247, 250)
    pdf.set_draw_color(130, 140, 155)
    pdf.set_line_width(0.4)
    pdf.rect(mp_x, mp_y, plate_w, plate_h, "DF")

    vd_w = 60 / 9
    pdf.set_fill_color(200, 205, 212)
    pdf.set_draw_color(150, 155, 165)
    pdf.set_line_width(0.2)
    pdf.rect(mp_x + 1.5, mp_y + 4, vd_w, plate_h - 8, "DF")
    pdf.rect(mp_x + plate_w - vd_w - 1.5, mp_y + 4, vd_w, plate_h - 8, "DF")

    hd_h = 60 / 9
    rw_x = mp_x + vd_w + 1.5
    rw_w = plate_w - 2*vd_w - 3
    
    hd_ys = [mp_y + 5, mp_y + 50, mp_y + 105, mp_y + 150]
    for hy in hd_ys:
        pdf.rect(rw_x, hy, rw_w, hd_h, "DF")
        for hx in range(int(rw_x + 2), int(rw_x + rw_w - 2), 4):
            pdf.line(hx, hy + 1, hx, hy + hd_h - 1)

    # Row 1: Power & Comms
    r1_y = mp_y + 15
    pdf.set_fill_color(210, 215, 220)
    pdf.rect(rw_x + 1, r1_y + 15, rw_w - 2, 2.5, "F")
    
    if panel_id == "PANEL M1":
        pdf.set_fill_color(70, 80, 95)
        pdf.rect(rw_x + 2, r1_y + 5, 8, 22, "DF")
        pdf.set_font("Helvetica", "B", 4.5)
        pdf.set_text_color(255, 255, 255)
        pdf.text(rw_x + 2.5, r1_y + 16, "Q1 MCB")
        
        pdf.set_fill_color(30, 65, 120)
        pdf.rect(rw_x + 12, r1_y + 2, 13, 27, "DF")
        pdf.rect(rw_x + 26, r1_y + 2, 13, 27, "DF")
        pdf.set_font("Helvetica", "B", 4.5)
        pdf.text(rw_x + 13, r1_y + 15, "PS1 24V")
        pdf.text(rw_x + 13, r1_y + 19, "20A QUINT")
        pdf.text(rw_x + 27, r1_y + 15, "PS2 24V")
        pdf.text(rw_x + 27, r1_y + 19, "20A QUINT")

        pdf.set_fill_color(45, 90, 160)
        pdf.rect(rw_x + 40, r1_y + 4, 7, 24, "DF")
        pdf.set_font("Helvetica", "B", 4)
        pdf.text(rw_x + 40.5, r1_y + 16, "ORING")

        pdf.set_fill_color(40, 130, 70)
        pdf.rect(rw_x + 49, r1_y + 3, 11, 25, "DF")
        pdf.set_font("Helvetica", "B", 4.5)
        pdf.text(rw_x + 49.5, r1_y + 15, "STRATIX")
        pdf.text(rw_x + 49.5, r1_y + 19, "5700 SW")
    else:
        pdf.set_fill_color(60, 75, 90)
        pdf.rect(rw_x + 3, r1_y + 5, 20, 22, "DF")
        pdf.set_font("Helvetica", "B", 4.5)
        pdf.set_text_color(255, 255, 255)
        pdf.text(rw_x + 4, r1_y + 16, "24VDC DISTRIB.")
        pdf.text(rw_x + 4, r1_y + 20, "CB1..CB8")

        pdf.set_fill_color(40, 130, 70)
        pdf.rect(rw_x + 27, r1_y + 3, 14, 25, "DF")
        pdf.set_font("Helvetica", "B", 4.5)
        pdf.text(rw_x + 28, r1_y + 15, "1756-EN4TR")
        pdf.text(rw_x + 28, r1_y + 19, "ETHERNET")

    # Row 2: ControlLogix 1756-A13 Rack
    r2_y = mp_y + 60
    r2_w = rw_w - 2
    r2_h = 40
    
    pdf.set_fill_color(25, 30, 38)
    pdf.set_draw_color(10, 15, 20)
    pdf.set_line_width(0.5)
    pdf.rect(rw_x + 1, r2_y, r2_w, r2_h, "DF")

    ps_w = 9
    pdf.set_fill_color(175, 35, 35)
    pdf.rect(rw_x + 2, r2_y + 2, ps_w, r2_h - 4, "DF")
    pdf.set_font("Helvetica", "B", 4.5)
    pdf.set_text_color(255, 255, 255)
    pdf.text(rw_x + 2.5, r2_y + 18, "1756")
    pdf.text(rw_x + 2.5, r2_y + 22, "PA72")

    slot_w = (r2_w - ps_w - 4) / 13
    slot_base_x = rw_x + 3 + ps_w
    
    for s_idx in range(13):
        slot_num = s_idx + 1
        sx = slot_base_x + s_idx * slot_w
        slot_info = slot_data.get(slot_num, {"model": "1756-N2", "type": "EMPTY", "color": (90, 95, 105)})
        
        pdf.set_fill_color(*slot_info["color"])
        pdf.rect(sx, r2_y + 2, slot_w - 0.4, r2_h - 4, "DF")
        
        pdf.set_font("Helvetica", "B", 3.2)
        pdf.set_text_color(255, 255, 255)
        pdf.text(sx + 0.3, r2_y + 5, f"S{slot_num}")
        
        pdf.set_font("Helvetica", "B", 3.4)
        m_code = slot_info["type"]
        pdf.text(sx + 0.3, r2_y + 20, m_code[:4])

    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(24, 43, 73)
    pdf.text(rw_x + 2, r2_y - 2, f"CHASSIS {chassis_id}: 1756-A13 (13 SLOTS)")

    # Row 3: Relays / Barriers
    r3_y = mp_y + 115
    pdf.set_fill_color(210, 215, 220)
    pdf.rect(rw_x + 1, r3_y + 12, rw_w - 2, 2.5, "F")

    if is_is_panel:
        pdf.set_fill_color(30, 115, 225)
        barrier_num = 20
        bw = (rw_w - 6) / barrier_num
        for bi in range(barrier_num):
            pdf.rect(rw_x + 3 + bi * bw, r3_y + 3, bw - 0.5, 22, "DF")
        pdf.set_font("Helvetica", "B", 5)
        pdf.set_text_color(20, 90, 200)
        pdf.text(rw_x + 3, r3_y - 2, "INTRINSIC SAFETY BARRIERS (Ex-i) - MTL / P+F ISOLATORS")
    else:
        pdf.set_fill_color(115, 120, 130)
        relay_num = 24
        rw_rel = (rw_w - 6) / relay_num
        for ri in range(relay_num):
            pdf.rect(rw_x + 3 + ri * rw_rel, r3_y + 4, rw_rel - 0.5, 20, "DF")
        pdf.set_font("Helvetica", "B", 5)
        pdf.set_text_color(50, 60, 75)
        pdf.text(rw_x + 3, r3_y - 2, "INTERPOSING RELAY MODULES (DO CONTACT ISOLATION)")

    # Row 4: Terminals
    r4_y = mp_y + 160
    pdf.set_fill_color(210, 215, 220)
    pdf.rect(rw_x + 1, r4_y + 12, rw_w - 2, 2.5, "F")

    if is_is_panel:
        term_avail = rw_w - 6
        pdf.set_fill_color(150, 155, 165)
        pdf.rect(rw_x + 3, r4_y + 3, term_avail * 0.35, 22, "DF")
        pdf.set_fill_color(235, 85, 25)
        pdf.rect(rw_x + 3 + term_avail * 0.35, r4_y + 1, 2, 26, "DF")
        pdf.set_fill_color(35, 115, 225)
        pdf.rect(rw_x + 6 + term_avail * 0.35, r4_y + 3, term_avail * 0.60, 22, "DF")
        
        pdf.set_font("Helvetica", "B", 5)
        pdf.set_text_color(50, 60, 75)
        pdf.text(rw_x + 3, r4_y - 2, "NON-IS TB")
        pdf.set_text_color(20, 90, 200)
        pdf.text(rw_x + 28, r4_y - 2, "BLUE IS TERMINALS (Ex-i WIRING)")
    else:
        pdf.set_fill_color(150, 155, 165)
        pdf.rect(rw_x + 3, r4_y + 3, rw_w - 6, 22, "DF")
        pdf.set_font("Helvetica", "B", 5)
        pdf.set_text_color(50, 60, 75)
        pdf.text(rw_x + 3, r4_y - 2, "FIELD CABLE TERMINAL STRIPS (X1 / X2)")

    # PE Bar
    pdf.set_fill_color(200, 150, 40)
    pdf.rect(rw_x + 2, mp_y + plate_h - 6, rw_w - 4, 3, "DF")
    pdf.set_font("Helvetica", "B", 4.5)
    pdf.set_text_color(40, 30, 10)
    pdf.text(rw_x + 4, mp_y + plate_h - 3.8, "PE / EARTH COPPER BAR 30x5mm & INSTRUMENT CLEAN SHIELD BAR")

    # Table on Right Side
    tbl_x = 122
    tbl_y = 35
    tbl_w = 280
    
    pdf.set_fill_color(24, 43, 73)
    pdf.rect(tbl_x, tbl_y, tbl_w, 7, "F")
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(255, 255, 255)
    pdf.text(tbl_x + 4, tbl_y + 5, f"{panel_id} - RACK {chassis_id} HARDWARE ALLOCATION & FIELD TERMINAL SCHEDULE")
    
    cols = [
        {"name": "Slot", "w": 12, "align": "C"},
        {"name": "Module Catalog", "w": 30, "align": "C"},
        {"name": "Description / Module Function", "w": 70, "align": "L"},
        {"name": "I/O Type", "w": 18, "align": "C"},
        {"name": "Channels", "w": 18, "align": "C"},
        {"name": "Terminal Strip", "w": 28, "align": "C"},
        {"name": "Connected Junction Boxes / Signals", "w": 68, "align": "L"},
        {"name": "Status", "w": 36, "align": "C"},
    ]
    
    th_y = tbl_y + 7
    pdf.set_fill_color(35, 55, 85)
    pdf.rect(tbl_x, th_y, tbl_w, 6, "F")
    pdf.set_font("Helvetica", "B", 6.8)
    pdf.set_text_color(255, 255, 255)
    
    cur_x = tbl_x
    for c in cols:
        pdf.rect(cur_x, th_y, c["w"], 6)
        # Position text exactly inside cell
        if c["align"] == "C":
            tw = pdf.get_string_width(c["name"])
            pdf.text(cur_x + (c["w"] - tw)/2, th_y + 4.2, c["name"])
        else:
            pdf.text(cur_x + 2, th_y + 4.2, c["name"])
        cur_x += c["w"]

    tr_y = th_y + 6
    row_h = 7.5
    
    # PS row
    pdf.set_fill_color(255, 255, 255)
    pdf.rect(tbl_x, tr_y, tbl_w, row_h, "F")
    pdf.set_font("Helvetica", "B", 7)
    pdf.set_text_color(180, 40, 40)
    pdf.text(tbl_x + 3, tr_y + 5, "PS")
    pdf.set_font("Helvetica", "B", 7)
    pdf.set_text_color(30, 30, 30)
    pdf.text(tbl_x + 14, tr_y + 5, "1756-PA72")
    pdf.set_font("Helvetica", "", 6.8)
    pdf.text(tbl_x + 44, tr_y + 5, "ControlLogix 85-265V AC Power Supply (72W @ 60C)")
    pdf.text(tbl_x + 115, tr_y + 5, "POWER")
    pdf.text(tbl_x + 134, tr_y + 5, "120/230VAC")
    pdf.text(tbl_x + 152, tr_y + 5, "Internal Rack")
    pdf.text(tbl_x + 180, tr_y + 5, "Chassis Backplane Power")
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.set_text_color(20, 120, 40)
    pdf.text(tbl_x + 248, tr_y + 5, "ACTIVE (RED)")
    tr_y += row_h

    for s_num in range(1, 14):
        s_data = slot_data.get(s_num, {
            "model": "1756-N2",
            "desc": "Slot Filler Cover (Spare Slot)",
            "type": "EMPTY",
            "ch": "-",
            "tb": "-",
            "serves": "Reserved for future system expansion",
            "status": "SPARE (AVAILABLE)"
        })
        
        if s_num % 2 == 0:
            pdf.set_fill_color(248, 250, 253)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.rect(tbl_x, tr_y, tbl_w, row_h, "F")
        pdf.rect(tbl_x, tr_y, tbl_w, row_h)
        
        pdf.set_font("Helvetica", "B", 7)
        pdf.set_text_color(20, 30, 45)
        pdf.text(tbl_x + 4, tr_y + 5, f"{s_num:02d}")
        
        pdf.set_font("Helvetica", "B" if s_data["type"] != "EMPTY" else "", 7)
        pdf.text(tbl_x + 14, tr_y + 5, s_data["model"])
        
        pdf.set_font("Helvetica", "", 6.8)
        pdf.text(tbl_x + 44, tr_y + 5, s_data["desc"][:45])
        
        pdf.set_font("Helvetica", "B", 7)
        pdf.text(tbl_x + 115, tr_y + 5, s_data["type"])
        
        pdf.set_font("Helvetica", "", 7)
        pdf.text(tbl_x + 134, tr_y + 5, str(s_data["ch"]))
        pdf.text(tbl_x + 152, tr_y + 5, s_data["tb"])
        pdf.text(tbl_x + 180, tr_y + 5, s_data["serves"][:42])
        
        if "SPARE" in s_data["status"]:
            pdf.set_font("Helvetica", "", 6.5)
            pdf.set_text_color(120, 120, 120)
        else:
            pdf.set_font("Helvetica", "B", 6.5)
            pdf.set_text_color(20, 120, 40)
        pdf.text(tbl_x + 248, tr_y + 5, s_data["status"])
        
        tr_y += row_h

    # Bottom notes box
    notes_x = tbl_x
    notes_y = tr_y + 8
    notes_w = tbl_w
    notes_h = 42
    
    pdf.set_fill_color(248, 250, 254)
    pdf.set_draw_color(180, 200, 220)
    pdf.rect(notes_x, notes_y, notes_w, notes_h, "DF")
    
    pdf.set_fill_color(24, 43, 73)
    pdf.rect(notes_x, notes_y, notes_w, 5.5, "F")
    pdf.set_font("Helvetica", "B", 7)
    pdf.set_text_color(255, 255, 255)
    pdf.text(notes_x + 3, notes_y + 4, f"{panel_id} INSTALLATION SPECIFICATIONS & FIELD CONNECTIONS")
    
    pdf.set_font("Helvetica", "", 7)
    pdf.set_text_color(30, 40, 50)
    lines = [
        f"- ENCLOSURE SIZE: 800 mm (W) x 2000 mm (H) x 800 mm (D) + 100 mm plinth base. Single door.",
        f"- FIELD DESTINATIONS: Connects to field junction boxes: {jb_serves}.",
        f"- POWER INFEED: Dual redundant 24VDC feeder lines from Panel M1 distribution bus.",
        f"- NETWORK ARCHITECTURE: Dual-ring EtherNet/IP Device Level Ring (DLR) connection.",
        "- CABLING RULES: All analog inputs wired with individually shielded twisted pairs (LiY-CY TP).",
        "- SHIELD GROUNDING: Cable screens grounded at panel entry copper earth bar only (single-point earthing)."
    ]
    cur_ny = notes_y + 10
    for l in lines:
        pdf.text(notes_x + 3, cur_ny, l)
        cur_ny += 5.2

# -------------------------------------------------------------
# SHEET 7: BILL OF MATERIALS & ALLOCATION MATRIX
# -------------------------------------------------------------
def build_sheet_7(pdf):
    pdf.add_page()
    pdf.draw_drawing_frame(7, 7, "4-PANEL SUITE - COMPLETE BILL OF MATERIALS (BOM) & SPECIFICATION", scale="N/A")
    
    start_x = 18
    start_y = 35
    tbl_w = 384
    
    pdf.set_fill_color(24, 43, 73)
    pdf.rect(start_x, start_y, tbl_w, 7, "F")
    pdf.set_font("Helvetica", "B", 8.5)
    pdf.set_text_color(255, 255, 255)
    pdf.text(start_x + 4, start_y + 5, "MAJOR HARDWARE BILL OF MATERIALS (BOM) - ENCLOSURES, PLC HARDWARE & ACCESSORIES")

    bom_cols = [
        {"name": "Item", "w": 12, "align": "C"},
        {"name": "Part Number / Catalog", "w": 40, "align": "L"},
        {"name": "Manufacturer", "w": 32, "align": "L"},
        {"name": "Component Description", "w": 85, "align": "L"},
        {"name": "Installed Location", "w": 38, "align": "C"},
        {"name": "Qty", "w": 14, "align": "C"},
        {"name": "Spare Qty", "w": 18, "align": "C"},
        {"name": "Technical Specifications / Ratings", "w": 95, "align": "L"},
        {"name": "Remarks", "w": 50, "align": "L"},
    ]

    th_y = start_y + 7
    pdf.set_fill_color(35, 55, 85)
    pdf.rect(start_x, th_y, tbl_w, 6, "F")
    pdf.set_font("Helvetica", "B", 6.8)
    pdf.set_text_color(255, 255, 255)
    
    cur_x = start_x
    for c in bom_cols:
        pdf.rect(cur_x, th_y, c["w"], 6)
        if c["align"] == "C":
            tw = pdf.get_string_width(c["name"])
            pdf.text(cur_x + (c["w"] - tw)/2, th_y + 4.2, c["name"])
        else:
            pdf.text(cur_x + 2, th_y + 4.2, c["name"])
        cur_x += c["w"]

    bom_items = [
        ("1", "VX 8806.500", "Rittal", "Industrial Bayed Enclosure Suite (4 bays)", "M1, M2, M3, M4", "4", "-", "800x2000x800 mm, IP54, Sheet Steel RAL 7035", "Complete with side panels & plinth"),
        ("2", "1756-A13", "Rockwell Automation", "ControlLogix 13-Slot Chassis", "M1, M2, M3, M4", "4", "-", "Standard backplane, horizontal mounting", "Chassis C1, C2, C3, C4"),
        ("3", "1756-PA72", "Rockwell Automation", "ControlLogix Power Supply", "M1, M2, M3, M4", "4", "1", "85-265V AC, 72W output @ 60 deg C", "One supply per chassis"),
        ("4", "1756-L950TPSXT", "Rockwell Automation", "ControlLogix 5580 Controller", "Panel M1 (C1S01)", "1", "1", "Extreme environment, 50MB memory, 1Gbps ETH", "Master Plant Controller"),
        ("5", "1756-EN4TR", "Rockwell Automation", "EtherNet/IP Dual-Port DLR Comm Module", "M2, M3, M4", "3", "1", "10/100 Mbps, 128 TCP connections, DLR ring", "I/O Remote Adapter modules"),
        ("6", "1756-IB32", "Rockwell Automation", "32-Point 24VDC Digital Input Module", "M1, M2, M3, M4", "14", "2", "24VDC Sink/Source, 10-31.2VDC, Isolated groups", "544 total DI channels available"),
        ("7", "1756-OB32", "Rockwell Automation", "32-Point 24VDC Digital Output Module", "M1, M2, M3, M4", "7", "1", "24VDC Source, 10-31.2VDC, 0.5A per point", "224 total DO channels available"),
        ("8", "1756-IF16", "Rockwell Automation", "16-Point Analog Input Module", "M1, M2, M3, M4", "11", "1", "Current / Voltage (4-20mA, 0-10V), 16-bit resolution", "176 total AI channels available"),
        ("9", "1756-OF8", "Rockwell Automation", "8-Point Analog Output Module", "Panel M3 (C3S13)", "1", "1", "Current / Voltage (4-20mA, 0-10V), 15-bit resolution", "8 total AO channels available"),
        ("10", "1756-N2", "Rockwell Automation", "ControlLogix Slot Filler Plate", "M1, M3, M4", "17", "-", "Chassis slot cover plate for dust/cooling protection", "Covers unpopulated slots"),
        ("11", "QUINT4-PS/1AC/24DC/20", "Phoenix Contact", "Primary Switched Redundant 24VDC Power Supply", "Panel M1", "2", "1", "Input 100-240VAC, Output 24VDC 20A, SFB Technology", "Primary/Secondary Redundant"),
        ("12", "QUINT4-DIODE/40", "Phoenix Contact", "Diode Redundancy Module", "Panel M1", "1", "-", "Redundancy decoupling module 2x 20A / 1x 40A", "Decouples PS1 & PS2"),
        ("13", "Stratix 5700", "Rockwell Automation", "Industrial Managed Ethernet Switch", "Panel M1", "1", "-", "16 Ports 10/100 Mbps, DLR, CIP Sync, VLAN support", "Panel Network Backbone"),
        ("14", "KFD2-STC4-Ex1", "Pepperl+Fuchs", "Intrinsically Safe SMART Transmitter Isolator", "Panel M3, M4", "38", "5", "Ex-ia IIC, 4-20mA current repeater with HART support", "For IS-JB-603, 608, 612, 618"),
        ("15", "KFD2-SL2-Ex2", "Pepperl+Fuchs", "Intrinsically Safe Digital Switch Isolator", "Panel M3, M4", "46", "6", "Ex-ia IIC, Dual-channel dry contact / NAMUR sensor", "For IS Level & Limit Switches"),
        ("16", "Finder 38.51", "Finder", "Interposing Relay Module (DO Isolation)", "M1, M2, M3, M4", "120", "15", "24VDC Coil, 1 CO 6A contact, LED indicator", "DIN Rail mount relay interface"),
        ("17", "WDU 2.5 / WDU 4", "Weidmuller", "Feed-Through Terminal Blocks (Gray)", "M1, M2, M3, M4", "950", "100", "800V, 24A, Screw connection, multi-tier", "Standard Field Wiring"),
        ("18", "WDU 2.5 BL", "Weidmuller", "Intrinsically Safe Terminal Blocks (Blue)", "Panel M3, M4", "320", "50", "Ex-i Blue colored terminal blocks for hazardous area", "IS Field Wiring Only"),
        ("19", "SK 3328.540", "Rittal", "Roof Mounted Cooling Unit / Fans", "M1, M2, M3, M4", "4", "-", "230V, 50/60Hz, 1500W cooling capacity / high flow", "Temperature maintenance < 35C"),
    ]

    tr_y = th_y + 6
    row_h = 7.2
    for item in bom_items:
        if int(item[0]) % 2 == 0:
            pdf.set_fill_color(248, 250, 253)
        else:
            pdf.set_fill_color(255, 255, 255)
        pdf.rect(start_x, tr_y, tbl_w, row_h, "F")
        pdf.rect(start_x, tr_y, tbl_w, row_h)
        
        pdf.set_font("Helvetica", "B", 6.8)
        pdf.set_text_color(20, 30, 45)
        pdf.text(start_x + 4, tr_y + 4.8, item[0])
        pdf.text(start_x + 14, tr_y + 4.8, item[1])
        pdf.set_font("Helvetica", "", 6.8)
        pdf.text(start_x + 54, tr_y + 4.8, item[2])
        pdf.text(start_x + 86, tr_y + 4.8, item[3][:45])
        pdf.text(start_x + 172, tr_y + 4.8, item[4])
        pdf.set_font("Helvetica", "B", 6.8)
        pdf.text(start_x + 215, tr_y + 4.8, item[5])
        pdf.set_font("Helvetica", "", 6.8)
        pdf.text(start_x + 230, tr_y + 4.8, item[6])
        pdf.text(start_x + 248, tr_y + 4.8, item[7][:52])
        pdf.text(start_x + 344, tr_y + 4.8, item[8][:28])
        
        tr_y += row_h

def generate_all_drawings():
    print("Generating Engineering Drawing Package...")
    pdf = EngineeringDrawing()
    
    print("Building Sheet 1: Front Elevation...")
    build_sheet_1(pdf)
    
    print("Building Sheet 2: Internal General Arrangement...")
    build_sheet_2(pdf)
    
    slot_data_m1 = {
        1: {"model": "1756-L950TPSXT", "desc": "ControlLogix 5580 Controller (50MB Memory)", "type": "CPU", "ch": "1", "tb": "ETH1/2", "serves": "Plant Master Controller / SCADA", "status": "ACTIVE (MASTER)", "color": (20, 100, 180)},
        2: {"model": "Reserve", "desc": "Reserved for 1756-EN4TR / Redundancy", "type": "RES", "ch": "-", "tb": "-", "serves": "Reserved for Controller Redundancy", "status": "SPARE (RESERVE)", "color": (120, 125, 135)},
        3: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C1S03-X1,X2", "serves": "JB-401 Field Signals & Spares", "status": "ACTIVE", "color": (40, 140, 70)},
        4: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C1S04-X1,X2", "serves": "JB-401 Field Signals", "status": "ACTIVE", "color": (40, 140, 70)},
        5: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C1S05-X1,X2", "serves": "JB-601 Field Signals", "status": "ACTIVE", "color": (40, 140, 70)},
        6: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C1S06-X1,X2", "serves": "JB-601 Field Signals & CA1 Spares", "status": "ACTIVE", "color": (40, 140, 70)},
        7: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C1S07-X1,X2", "serves": "JB-401 Solenoid Valves & DO", "status": "ACTIVE", "color": (200, 120, 20)},
        8: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C1S08-X1,X2", "serves": "JB-601 Solenoids & CA1 Spares", "status": "ACTIVE", "color": (200, 120, 20)},
        9: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C1S09-X1", "serves": "JB-401 Analog Transmitters (4-20mA)", "status": "ACTIVE", "color": (90, 60, 160)},
        10: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C1S10-X1", "serves": "JB-601 Analog Transmitters (4-20mA)", "status": "ACTIVE", "color": (90, 60, 160)},
        11: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        12: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        13: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
    }

    slot_data_m2 = {
        1: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C2S01-X1,X2", "serves": "JB-402 & JB-602 Flow & Limit Switches", "status": "ACTIVE", "color": (40, 140, 70)},
        2: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C2S02-X1,X2", "serves": "JB-602 & JB-607 Field Inputs", "status": "ACTIVE", "color": (40, 140, 70)},
        3: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C2S03-X1,X2", "serves": "JB-607 Spray Dryer 7th Flr DI", "status": "ACTIVE", "color": (40, 140, 70)},
        4: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C2S04-X1,X2", "serves": "JB-607 Spray Dryer 7th Flr DI", "status": "ACTIVE", "color": (40, 140, 70)},
        5: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C2S05-X1,X2", "serves": "JB-607 Spray Dryer 7th Flr DI", "status": "ACTIVE", "color": (40, 140, 70)},
        6: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C2S06-X1,X2", "serves": "JB-402 & JB-602 On/Off Valves", "status": "ACTIVE", "color": (200, 120, 20)},
        7: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C2S07-X1,X2", "serves": "JB-607 Solenoid Valves (DO)", "status": "ACTIVE", "color": (200, 120, 20)},
        8: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C2S08-X1,X2", "serves": "JB-607 Solenoid Valves (DO)", "status": "ACTIVE", "color": (200, 120, 20)},
        9: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C2S09-X1", "serves": "JB-402 Jet Cooker Flow/Press AI", "status": "ACTIVE", "color": (90, 60, 160)},
        10: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C2S10-X1", "serves": "JB-602 Spray Dryer 3rd Flr AI", "status": "ACTIVE", "color": (90, 60, 160)},
        11: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C2S11-X1", "serves": "JB-602 Spray Dryer 3rd Flr AI", "status": "ACTIVE", "color": (90, 60, 160)},
        12: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C2S12-X1", "serves": "JB-607 Spray Dryer 7th Flr AI", "status": "ACTIVE", "color": (90, 60, 160)},
        13: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C2S13-X1", "serves": "JB-607 Spray Dryer 7th Flr AI", "status": "ACTIVE", "color": (90, 60, 160)},
    }

    slot_data_m3 = {
        1: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C3S01-X1,X2", "serves": "IS-JB-603 & IS-JB-618 Ex-i DI", "status": "ACTIVE (IS)", "color": (40, 140, 70)},
        2: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C3S02-X1,X2", "serves": "IS-JB-618 & JB-618 Field DI", "status": "ACTIVE", "color": (40, 140, 70)},
        3: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C3S03-X1,X2", "serves": "JB-618 Packing Tower 8th Flr DO", "status": "ACTIVE", "color": (200, 120, 20)},
        4: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C3S04-X1", "serves": "IS-JB-603 Ex-i Transmitters (AI)", "status": "ACTIVE (IS)", "color": (90, 60, 160)},
        5: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C3S05-X1", "serves": "IS-JB-618 & JB-618 Transmitters", "status": "ACTIVE (IS)", "color": (90, 60, 160)},
        6: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        7: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        8: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        9: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        10: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        11: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        12: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        13: {"model": "1756-OF8", "desc": "8-Point High Resolution Analog Output", "type": "AO", "ch": "8", "tb": "C3S13-X1", "serves": "Control Valves / Speed References", "status": "ACTIVE", "color": (20, 160, 160)},
    }

    slot_data_m4 = {
        1: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C4S01-X1,X2", "serves": "JB-606 & IS-JB-612 Ex-i DI", "status": "ACTIVE", "color": (40, 140, 70)},
        2: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C4S02-X1,X2", "serves": "IS-JB-612 & JB-612 Field DI", "status": "ACTIVE (IS)", "color": (40, 140, 70)},
        3: {"model": "1756-IB32", "desc": "32-Point 24VDC Sink/Source Digital Input", "type": "DI", "ch": "32", "tb": "C4S03-X1,X2", "serves": "IS-JB-608 & JB-608 Ex-i DI", "status": "ACTIVE (IS)", "color": (40, 140, 70)},
        4: {"model": "1756-OB32", "desc": "32-Point 24VDC Source Digital Output", "type": "DO", "ch": "32", "tb": "C4S04-X1,X2", "serves": "JB-606, JB-612, IS-JB-608 DO", "status": "ACTIVE", "color": (200, 120, 20)},
        5: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C4S05-X1", "serves": "JB-606 & IS-JB-612 Ex-i AI", "status": "ACTIVE (IS)", "color": (90, 60, 160)},
        6: {"model": "1756-IF16", "desc": "16-Point High Resolution Analog Input", "type": "AI", "ch": "16", "tb": "C4S06-X1", "serves": "JB-612, IS-JB-608, JB-608 AI", "status": "ACTIVE (IS)", "color": (90, 60, 160)},
        7: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        8: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        9: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        10: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        11: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        12: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
        13: {"model": "1756-N2", "desc": "Slot Filler Plate (Empty Slot)", "type": "EMPTY", "ch": "-", "tb": "-", "serves": "Chassis Spare Expansion Slot", "status": "SPARE (AVAILABLE)", "color": (90, 95, 105)},
    }

    print("Building Sheet 3: Panel M1 Detailed Mounting Plan...")
    build_detailed_panel_sheet(pdf, 3, "PANEL M1", "MASTER CONTROLLER & POWER BAY", "C1", "JB-401, JB-601, CA1", slot_data_m1, is_is_panel=False)

    print("Building Sheet 4: Panel M2 Detailed Mounting Plan...")
    build_detailed_panel_sheet(pdf, 4, "PANEL M2", "HIGH-DENSITY I/O EXPANSION BAY 1", "C2", "JB-602, JB-402, JB-607", slot_data_m2, is_is_panel=False)

    print("Building Sheet 5: Panel M3 Detailed Mounting Plan...")
    build_detailed_panel_sheet(pdf, 5, "PANEL M3", "I/O & INTRINSIC SAFETY (Ex-i) BAY 2", "C3", "IS-JB-603, IS-JB-618, JB-618", slot_data_m3, is_is_panel=True)

    print("Building Sheet 6: Panel M4 Detailed Mounting Plan...")
    build_detailed_panel_sheet(pdf, 6, "PANEL M4", "I/O & INTRINSIC SAFETY (Ex-i) BAY 3", "C4", "JB-606, IS-JB-612, JB-612, IS-JB-608, JB-608", slot_data_m4, is_is_panel=True)

    print("Building Sheet 7: Bill of Materials...")
    build_sheet_7(pdf)

    pdf.output(PDF_FILENAME)
    print(f"PDF Successfully saved to {PDF_FILENAME}")
    
    print("Rendering high-res PNG images of drawing sheets...")
    doc = fitz.open(PDF_FILENAME)
    png_paths = []
    for p_idx in range(len(doc)):
        pix = doc[p_idx].get_pixmap(dpi=150)
        img_name = f"mcc_panel_layout_sheet_{p_idx+1}.png"
        img_path = os.path.join(OUTPUT_DIR, img_name)
        pix.save(img_path)
        png_paths.append(img_path)
        print(f"  Rendered: {img_name}")

    print("All panel layout drawings generated successfully!")
    return PDF_FILENAME, png_paths

if __name__ == "__main__":
    generate_all_drawings()
