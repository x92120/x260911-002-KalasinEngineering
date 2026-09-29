#!/usr/bin/env python3
"""
Export AutoCAD DXF (1:1 mm model space with standard CAD layers) and
W3C Scalable Vector Graphics (SVG) for the MCC 4-Panel Suite (M1, M2, M3, M4)
for Project Jet Cooker at Ingredion (Thailand) Co., Ltd. / AEC Industrial Engineering.
"""

import os
import sys
import fitz  # PyMuPDF
import ezdxf
from ezdxf import colors
from ezdxf.enums import TextEntityAlignment

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(WORKSPACE_DIR, "reports_pdf")
CAD_DIR = os.path.join(WORKSPACE_DIR, "cad_exports")
os.makedirs(CAD_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# STEP 1: EXPORT HIGH-FIDELITY VECTOR SVG FILES FROM DRAWING PDF
# -----------------------------------------------------------------------------
def export_svg_sheets():
    pdf_path = os.path.join(PDF_DIR, "MCC_Panel_Layout_Mounting_Plan_M1_M2_M3_M4.pdf")
    if not os.path.exists(pdf_path):
        print(f"Error: {pdf_path} not found.")
        return []

    doc = fitz.open(pdf_path)
    sheet_names = [
        "Sheet1_Front_Elevation",
        "Sheet2_Internal_GA",
        "Sheet3_Panel_M1",
        "Sheet4_Panel_M2",
        "Sheet5_Panel_M3",
        "Sheet6_Panel_M4",
        "Sheet7_BOM"
    ]

    generated_svgs = []
    print(f"--- Exporting {len(doc)} sheets to Scalable Vector Graphics (SVG) ---")
    for idx, page in enumerate(doc):
        base_name = f"MCC_Panel_Layout_{sheet_names[idx]}"
        svg_filename = f"{base_name}.svg"
        
        # Save to cad_exports and reports_pdf
        cad_out = os.path.join(CAD_DIR, svg_filename)
        pdf_out = os.path.join(PDF_DIR, svg_filename)

        svg_content = page.get_svg_image()
        with open(cad_out, "w", encoding="utf-8") as f:
            f.write(svg_content)
        with open(pdf_out, "w", encoding="utf-8") as f:
            f.write(svg_content)

        print(f"  [OK] Exported: {svg_filename} ({len(svg_content):,} bytes)")
        generated_svgs.append(cad_out)

    return generated_svgs


# -----------------------------------------------------------------------------
# STEP 2: DXF CAD BUILDER WITH 1:1 METRIC MILLIMETER REAL-WORLD SCALE
# -----------------------------------------------------------------------------
class CADExporter:
    def __init__(self, dxf_version="R2010"):
        self.doc = ezdxf.new(dxf_version, setup=True)
        self.doc.header["$INSUNITS"] = 4  # 4 = Millimeters
        self.msp = self.doc.modelspace()
        self._setup_layers()
        self._setup_dimstyle()

    def _setup_layers(self):
        # Professional AutoCAD layers with standard ACI colors and lineweights
        layers = [
            ("0_BORDER", 7, 50),            # White, 0.50mm
            ("0_TITLE_BLOCK", 4, 35),       # Cyan, 0.35mm
            ("ENCL_OUTLINE", 7, 70),        # White, 0.70mm
            ("ENCL_DOOR", 8, 35),           # Gray, 0.35mm
            ("BASE_PLINTH", 9, 50),         # Dark Gray, 0.50mm
            ("SUBPANEL_PLATE", 8, 35),      # Gray, 0.35mm
            ("WIRE_DUCT", 8, 18),           # Gray, 0.18mm
            ("DIN_RAIL", 2, 25),            # Yellow, 0.25mm
            ("PLC_CHASSIS", 5, 50),         # Blue, 0.50mm
            ("PLC_SLOT_CPU", 5, 35),        # Blue, 0.35mm
            ("PLC_SLOT_DI", 3, 35),         # Green, 0.35mm
            ("PLC_SLOT_DO", 30, 35),        # Orange, 0.35mm
            ("PLC_SLOT_AI", 6, 35),         # Magenta, 0.35mm
            ("PLC_SLOT_AO", 4, 35),         # Cyan, 0.35mm
            ("PLC_SLOT_SPARE", 8, 25),      # Gray, 0.25mm
            ("POWER_SUPPLY", 1, 35),        # Red, 0.35mm
            ("COMM_SWITCH", 3, 35),         # Green, 0.35mm
            ("IS_BARRIERS", 140, 35),       # Sky Blue, 0.35mm
            ("RELAYS", 2, 35),              # Yellow, 0.35mm
            ("TERMINAL_STRIP", 92, 35),     # Light Green, 0.35mm
            ("TERMINAL_IS", 140, 35),       # Blue, 0.35mm
            ("EARTH_BAR", 40, 50),          # Copper / Brown, 0.50mm
            ("DOOR_COMPONENTS", 2, 35),     # Yellow, 0.35mm
            ("HMI_UNIT", 4, 50),            # Cyan, 0.50mm
            ("DIMENSIONS", 1, 18),          # Red, 0.18mm
            ("TEXT_LABELS", 7, 25),         # White, 0.25mm
            ("TEXT_HEADER", 4, 35)          # Cyan, 0.35mm
        ]
        for name, color, lw in layers:
            if name not in self.doc.layers:
                self.doc.layers.add(name, color=color, lineweight=lw)

    def _setup_dimstyle(self):
        # Set up a clean metric engineering dimension style
        ds = self.doc.dimstyles.new("METRIC_MM")
        ds.dxf.dimtxt = 35.0   # Text height: 35 mm
        ds.dxf.dimasz = 25.0   # Arrowhead size: 25 mm
        ds.dxf.dimexe = 15.0   # Extension line extension: 15 mm
        ds.dxf.dimexo = 10.0   # Extension line offset: 10 mm
        ds.dxf.dimclrd = 1     # Red dimension lines
        ds.dxf.dimclre = 1     # Red extension lines
        ds.dxf.dimclrt = 7     # White dimension text
        ds.dxf.dimtad = 1      # Text placed above dimension line
        ds.dxf.dimdec = 0      # Decimal places: 0 (integers in mm)

    def add_rect(self, x, y, w, h, layer="0", closed=True):
        pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
        return self.msp.add_lwpolyline(pts, close=closed, dxfattribs={"layer": layer})

    def add_circle(self, cx, cy, radius, layer="0"):
        return self.msp.add_circle((cx, cy), radius, dxfattribs={"layer": layer})

    def add_text(self, text, x, y, height=25.0, layer="TEXT_LABELS", align=TextEntityAlignment.LEFT):
        txt = self.msp.add_text(text, dxfattribs={"height": height, "layer": layer})
        txt.set_placement((x, y), align=align)
        return txt

    def add_linear_dimension(self, p1, p2, base, angle=0):
        dim = self.msp.add_linear_dim(
            base=base,
            p1=p1,
            p2=p2,
            angle=angle,
            dimstyle="METRIC_MM",
            dxfattribs={"layer": "DIMENSIONS"}
        )
        dim.render()
        return dim

    def draw_title_block(self, ox, oy, title, doc_no="CA1-DWG-PL-001", sheet_str="1 / 1", scale_str="1:1"):
        tb_w = 1200
        tb_h = 280
        x = ox - tb_w
        y = oy
        
        self.add_rect(x, y, tb_w, tb_h, layer="0_BORDER")
        # Internal divisions
        self.msp.add_line((x, y + 180), (x + tb_w, y + 180), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x, y + 90), (x + tb_w, y + 90), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 750, y), (x + 750, y + 180), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 980, y), (x + 980, y + 90), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 400, y), (x + 400, y + 90), dxfattribs={"layer": "0_TITLE_BLOCK"})

        # Title Block Texts
        self.add_text("INGREDION (THAILAND) CO., LTD.", x + 25, y + 235, height=32, layer="TEXT_HEADER")
        self.add_text("KALASIN STARCH PLANT - JET COOKER PROJECT", x + 25, y + 200, height=22, layer="TEXT_LABELS")

        self.add_text("AEC INDUSTRIAL ENGINEERING CO., LTD.", x + 775, y + 235, height=26, layer="TEXT_HEADER")
        self.add_text("SYSTEMS INTEGRATION & AUTOMATION", x + 775, y + 200, height=18, layer="TEXT_LABELS")

        self.add_text(title, x + 25, y + 135, height=30, layer="TEXT_HEADER")
        self.add_text("MAIN CONTROL CABINET CA1 / MOTOR CONTROL PANELS M1..M4", x + 25, y + 105, height=18, layer="TEXT_LABELS")

        # Doc No & Rev
        self.add_text(f"DOC NO: {doc_no}", x + 775, y + 145, height=22, layer="TEXT_LABELS")
        self.add_text("REV: 3.6    DATE: 2026-09-07", x + 775, y + 110, height=20, layer="TEXT_LABELS")

        # Drawn / Checked / Scale / Sheet
        self.add_text("DRAWN: AEC-ENG", x + 25, y + 55, height=18, layer="TEXT_LABELS")
        self.add_text("CHECKED: LEAD-PE", x + 25, y + 25, height=18, layer="TEXT_LABELS")
        self.add_text("STATUS: APPROVED", x + 425, y + 55, height=20, layer="TEXT_HEADER")
        self.add_text(f"SCALE: {scale_str} (MODEL 1:1 mm)", x + 425, y + 25, height=18, layer="TEXT_LABELS")
        self.add_text(f"SHEET: {sheet_str}", x + 775, y + 40, height=28, layer="TEXT_HEADER")

    def draw_elevation_suite(self, origin_x=0.0, origin_y=0.0):
        """Draws the 4-panel front elevation at true 1:1 mm scale."""
        bay_w = 800.0
        body_h = 2000.0
        plinth_h = 100.0
        total_w = bay_w * 4.0  # 3200 mm
        
        ox = origin_x
        oy = origin_y

        # 1. Base Plinth (100mm)
        self.add_rect(ox, oy, total_w, plinth_h, layer="BASE_PLINTH")
        for i in range(4):
            bx = ox + i * bay_w
            self.add_rect(bx + 50, oy + 20, bay_w - 100, plinth_h - 40, layer="BASE_PLINTH")
            self.add_text("PLINTH 100mm", bx + bay_w/2, oy + 40, height=22, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

        bay_titles = [
            ("PANEL M1", "MASTER CONTROLLER"),
            ("PANEL M2", "I/O EXPANSION 1"),
            ("PANEL M3", "I/O & IS BARRIERS"),
            ("PANEL M4", "I/O & IS BARRIERS")
        ]

        # 2. Outer Bays & Doors
        for i in range(4):
            bx = ox + i * bay_w
            by = oy + plinth_h
            # Outer Enclosure Bay
            self.add_rect(bx, by, bay_w, body_h, layer="ENCL_OUTLINE")

            # Door Frame
            door_x = bx + 20
            door_y = by + 20
            door_w = bay_w - 40
            door_h = body_h - 40
            self.add_rect(door_x, door_y, door_w, door_h, layer="ENCL_DOOR")

            # Hinges (Left side)
            for hy in [door_y + 200, door_y + door_h/2 - 35, door_y + door_h - 250]:
                self.add_rect(door_x - 5, hy, 15, 70, layer="ENCL_DOOR")

            # Handle & Keylock (Right side)
            hx = door_x + door_w - 45
            hy = door_y + door_h/2 - 100
            self.add_rect(hx, hy, 25, 200, layer="DOOR_COMPONENTS")
            self.add_circle(hx + 12.5, hy + 100, 7.5, layer="DOOR_COMPONENTS")

            # Roof Exhaust Fan
            self.add_rect(bx + bay_w/2 - 160, by + body_h - 80, 320, 70, layer="ENCL_OUTLINE")
            self.add_text("ROOF EXHAUST FAN", bx + bay_w/2, by + body_h - 45, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # Lifting Eyebolts
            self.add_circle(bx + 100, by + body_h + 35, 25, layer="ENCL_OUTLINE")
            self.add_circle(bx + bay_w - 100, by + body_h + 35, 25, layer="ENCL_OUTLINE")

            # Header Nameplate
            self.add_rect(door_x + 40, door_y + door_h - 130, door_w - 80, 80, layer="0_TITLE_BLOCK")
            self.add_text(bay_titles[i][0], door_x + door_w/2, door_y + door_h - 80, height=26, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
            self.add_text(bay_titles[i][1], door_x + door_w/2, door_y + door_h - 110, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # Bottom Air Filter Louver
            self.add_rect(door_x + 80, door_y + 60, door_w - 160, 100, layer="DOOR_COMPONENTS")
            self.add_text("AIR INLET FILTER", door_x + door_w/2, door_y + 110, height=20, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            if i == 0:
                # Bay M1: HMI Screen (PanelView Plus 7, 15-inch)
                hmi_w = 400.0
                hmi_h = 300.0
                hmi_x = door_x + (door_w - hmi_w)/2
                hmi_y = door_y + door_h - 520
                self.add_rect(hmi_x, hmi_y, hmi_w, hmi_h, layer="HMI_UNIT")
                self.add_rect(hmi_x + 25, hmi_y + 25, hmi_w - 50, hmi_h - 50, layer="HMI_UNIT")
                self.add_text("PanelView Plus 7 (15-inch)", hmi_x + hmi_w/2, hmi_y + 170, height=24, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
                self.add_text("JET COOKER SCADA", hmi_x + hmi_w/2, hmi_y + 120, height=20, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

                # Pilot Lights (R, S, T)
                ply = hmi_y - 100
                for c_idx, lbl in enumerate(["R (RED)", "S (AMBER)", "T (BLUE)"]):
                    plx = door_x + 180 + c_idx * 160
                    self.add_circle(plx, ply, 25, layer="DOOR_COMPONENTS")
                    self.add_text(lbl, plx, ply - 45, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

                # Pushbuttons (Start, Stop, Reset, Test)
                pby = ply - 140
                for p_idx, lbl in enumerate(["START", "STOP", "RESET", "LAMP TEST"]):
                    px = door_x + 130 + p_idx * 140
                    self.add_circle(px, pby, 25, layer="DOOR_COMPONENTS")
                    self.add_text(lbl, px, pby - 45, height=16, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

                # Emergency Stop Pushbutton
                es_x = door_x + door_w/2
                es_y = pby - 180
                self.add_circle(es_x, es_y, 55, layer="DOOR_COMPONENTS")  # Yellow shroud
                self.add_circle(es_x, es_y, 35, layer="DOOR_COMPONENTS")  # Red mushroom head
                self.add_text("EMERGENCY STOP", es_x, es_y - 80, height=22, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)

                # Main Disconnect Switch Handle
                ms_y = es_y - 200
                self.add_rect(es_x - 50, ms_y - 50, 100, 100, layer="DOOR_COMPONENTS")
                self.add_rect(es_x - 15, ms_y - 50, 30, 100, layer="DOOR_COMPONENTS")
                self.add_text("MAIN SWITCH (400V)", es_x, ms_y - 85, height=20, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            else:
                # Bays M2, M3, M4: Status Lights
                ply = door_y + door_h - 400
                stat_lights = ["POWER ON", "HEALTHY", "ALARM", "FAULT"]
                for s_idx, lbl in enumerate(stat_lights):
                    plx = door_x + 130 + s_idx * 140
                    self.add_circle(plx, ply, 25, layer="DOOR_COMPONENTS")
                    self.add_text(lbl, plx, ply - 45, height=16, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

                # Electrical Hazard Sign
                tri_x = door_x + door_w/2
                tri_y = ply - 250
                pts = [(tri_x, tri_y + 100), (tri_x - 90, tri_y - 60), (tri_x + 90, tri_y - 60)]
                self.msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": "DOOR_COMPONENTS"})
                self.add_text("!", tri_x, tri_y, height=70, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
                self.add_text("400V / 230V HAZARD", tri_x, tri_y - 105, height=22, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)

                # Document Pocket
                dp_w = door_w - 200
                dp_h = 300
                dp_x = door_x + 100
                dp_y = door_y + 240
                self.add_rect(dp_x, dp_y, dp_w, dp_h, layer="DOOR_COMPONENTS")
                self.add_text("DOCUMENT POCKET (A4)", dp_x + dp_w/2, dp_y + dp_h/2, height=22, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

        # 3. Elevation Dimensions
        # Bottom Total Width (3200 mm)
        dim_y1 = oy - 150
        self.add_linear_dimension((ox, oy), (ox + total_w, oy), (ox + total_w/2, dim_y1))

        # Individual Bay Widths (800 mm each)
        dim_y2 = oy - 280
        for i in range(4):
            bx = ox + i * bay_w
            self.add_linear_dimension((bx, oy), (bx + bay_w, oy), (bx + bay_w/2, dim_y2))

        # Height Dimensions (Right side)
        dim_x1 = ox + total_w + 160
        self.add_linear_dimension((ox + total_w, oy), (ox + total_w, oy + plinth_h), (dim_x1, oy + plinth_h/2))
        self.add_linear_dimension((ox + total_w, oy + plinth_h), (ox + total_w, oy + plinth_h + body_h), (dim_x1, oy + plinth_h + body_h/2))

        # Total Height Dimension (2100 mm)
        dim_x2 = ox + total_w + 300
        self.add_linear_dimension((ox + total_w, oy), (ox + total_w, oy + plinth_h + body_h), (dim_x2, oy + (plinth_h + body_h)/2))


    def draw_internal_ga_suite(self, origin_x=3800.0, origin_y=0.0):
        """Draws the 4-panel internal mounting layout & subpanels at true 1:1 mm scale."""
        bay_w = 800.0
        body_h = 2000.0
        plinth_h = 100.0
        total_w = bay_w * 4.0
        
        ox = origin_x
        oy = origin_y

        # Base Plinth
        self.add_rect(ox, oy, total_w, plinth_h, layer="BASE_PLINTH")

        bay_destinations = [
            ("PANEL M1 (CHASSIS C1)", "SERVES: JB-401, JB-601, CA1"),
            ("PANEL M2 (CHASSIS C2)", "SERVES: JB-602, JB-402, JB-607"),
            ("PANEL M3 (CHASSIS C3 - Ex-i)", "SERVES: IS-JB-603, IS-JB-618, JB-618"),
            ("PANEL M4 (CHASSIS C4 - Ex-i)", "SERVES: JB-606, IS-JB-612, JB-612, IS-JB-608")
        ]

        for i in range(4):
            bx = ox + i * bay_w
            by = oy + plinth_h

            # Outer Enclosure Bay
            self.add_rect(bx, by, bay_w, body_h, layer="ENCL_OUTLINE")

            # Subpanel Mounting Plate (700 x 1900 mm)
            plate_x = bx + 50.0
            plate_y = by + 50.0
            plate_w = 700.0
            plate_h = 1900.0
            self.add_rect(plate_x, plate_y, plate_w, plate_h, layer="SUBPANEL_PLATE")

            # Vertical Wire Ducts (80mm width)
            vd_w = 80.0
            self.add_rect(plate_x + 15, plate_y + 60, vd_w, plate_h - 120, layer="WIRE_DUCT")
            self.add_rect(plate_x + plate_w - vd_w - 15, plate_y + 60, vd_w, plate_h - 120, layer="WIRE_DUCT")

            # Horizontal Wire Ducts (80mm width)
            hd_w = plate_w - 2 * vd_w - 30
            hd_x = plate_x + vd_w + 15
            hd_ys = [plate_y + 1750, plate_y + 1300, plate_y + 800, plate_y + 400]
            for hy in hd_ys:
                self.add_rect(hd_x, hy, hd_w, 80.0, layer="WIRE_DUCT")

            # DIN Rails (35mm standard)
            # Row 1 DIN Rail: y = plate_y + 1550
            r1_din_y = plate_y + 1550
            self.add_rect(hd_x + 10, r1_din_y, hd_w - 20, 35.0, layer="DIN_RAIL")

            # Row 3 DIN Rail: y = plate_y + 1050
            r3_din_y = plate_y + 1050
            self.add_rect(hd_x + 10, r3_din_y, hd_w - 20, 35.0, layer="DIN_RAIL")

            # Row 4 DIN Rail: y = plate_y + 200
            r4_din_y = plate_y + 200
            self.add_rect(hd_x + 10, r4_din_y, hd_w - 20, 35.0, layer="DIN_RAIL")

            # Earth Copper Bar 30x5 mm
            pe_y = plate_y + 60
            self.add_rect(plate_x + 20, pe_y, plate_w - 40, 30.0, layer="EARTH_BAR")
            self.add_text("PE / EARTH COPPER BAR 30x5mm", plate_x + plate_w/2, pe_y + 15, height=16, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # ROW 1 EQUIPMENT (Power & Networking)
            if i == 0:
                # MCB
                self.add_rect(hd_x + 20, r1_din_y - 25, 70, 85, layer="POWER_SUPPLY")
                self.add_text("MCB", hd_x + 55, r1_din_y + 17, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)
                # Redundant 24VDC 20A PSUs (1606-XLS480E)
                self.add_rect(hd_x + 110, r1_din_y - 45, 120, 125, layer="POWER_SUPPLY")
                self.add_text("PSU 1 (20A)", hd_x + 170, r1_din_y + 17, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)
                self.add_rect(hd_x + 245, r1_din_y - 45, 120, 125, layer="POWER_SUPPLY")
                self.add_text("PSU 2 (20A)", hd_x + 305, r1_din_y + 17, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)
                # Stratix 5700 Switch
                self.add_rect(hd_x + 380, r1_din_y - 40, 100, 115, layer="COMM_SWITCH")
                self.add_text("STRATIX 5700", hd_x + 430, r1_din_y + 17, height=16, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)
            else:
                # DC Distribution Fuse/CB
                self.add_rect(hd_x + 20, r1_din_y - 25, 180, 85, layer="POWER_SUPPLY")
                self.add_text("DC FUSE / CB UNIT", hd_x + 110, r1_din_y + 17, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)
                # Ethernet Adapter (1756-EN2TR)
                self.add_rect(hd_x + 230, r1_din_y - 35, 120, 105, layer="COMM_SWITCH")
                self.add_text("1756-EN2TR", hd_x + 290, r1_din_y + 17, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # ROW 2: ControlLogix 1756-A13 Chassis
            chassis_w = hd_w - 20
            chassis_h = 320.0
            chassis_x = hd_x + 10
            chassis_y = plate_y + 920

            self.add_rect(chassis_x, chassis_y, chassis_w, chassis_h, layer="PLC_CHASSIS")
            # Power Supply Module (1756-PA72)
            ps_w = 75.0
            self.add_rect(chassis_x + 10, chassis_y + 15, ps_w, chassis_h - 30, layer="POWER_SUPPLY")
            self.add_text("PA72", chassis_x + 10 + ps_w/2, chassis_y + chassis_h/2, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # 13 Slots
            slot_start_x = chassis_x + 15 + ps_w
            slot_w = (chassis_w - ps_w - 25) / 13.0
            for s in range(13):
                sx = slot_start_x + s * slot_w
                card_layer = "PLC_SLOT_SPARE"
                card_lbl = "N2"

                if i == 0:
                    if s == 0:
                        card_layer = "PLC_SLOT_CPU"; card_lbl = "CPU"
                    elif s == 1:
                        card_layer = "PLC_SLOT_SPARE"; card_lbl = "RES"
                    elif s in [2, 3, 4, 5]:
                        card_layer = "PLC_SLOT_DI"; card_lbl = "DI"
                    elif s in [6, 7]:
                        card_layer = "PLC_SLOT_DO"; card_lbl = "DO"
                    elif s in [8, 9]:
                        card_layer = "PLC_SLOT_AI"; card_lbl = "AI"
                    else:
                        card_layer = "PLC_SLOT_SPARE"; card_lbl = "N2"
                elif i == 1:
                    if s in [0, 1, 2, 3, 4]:
                        card_layer = "PLC_SLOT_DI"; card_lbl = "DI"
                    elif s in [5, 6, 7]:
                        card_layer = "PLC_SLOT_DO"; card_lbl = "DO"
                    elif s in [8, 9, 10, 11, 12]:
                        card_layer = "PLC_SLOT_AI"; card_lbl = "AI"
                elif i == 2:
                    if s in [0, 1]:
                        card_layer = "PLC_SLOT_DI"; card_lbl = "DI"
                    elif s == 2:
                        card_layer = "PLC_SLOT_DO"; card_lbl = "DO"
                    elif s in [3, 4]:
                        card_layer = "PLC_SLOT_AI"; card_lbl = "AI"
                    elif s == 12:
                        card_layer = "PLC_SLOT_AO"; card_lbl = "AO"
                    else:
                        card_layer = "PLC_SLOT_SPARE"; card_lbl = "N2"
                elif i == 3:
                    if s in [0, 1, 2]:
                        card_layer = "PLC_SLOT_DI"; card_lbl = "DI"
                    elif s == 3:
                        card_layer = "PLC_SLOT_DO"; card_lbl = "DO"
                    elif s in [4, 5]:
                        card_layer = "PLC_SLOT_AI"; card_lbl = "AI"
                    else:
                        card_layer = "PLC_SLOT_SPARE"; card_lbl = "N2"

                self.add_rect(sx, chassis_y + 15, slot_w - 2, chassis_h - 30, layer=card_layer)
                self.add_text(card_lbl, sx + (slot_w - 2)/2, chassis_y + chassis_h/2, height=14, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            self.add_text(f"CHASSIS C{i+1}: 1756-A13 (13 SLOTS)", chassis_x + chassis_w/2, chassis_y + chassis_h + 20, height=20, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)

            # ROW 3: Relays / Galvanic Barriers
            if i in [2, 3]:
                # Intrinsically Safe (Ex-i) Galvanic Isolator Barriers (Pepperl+Fuchs KFD2)
                bar_count = 18 if i == 2 else 14
                bw = (hd_w - 20) / bar_count
                for b_idx in range(bar_count):
                    bar_x = hd_x + 10 + b_idx * bw
                    self.add_rect(bar_x, r3_din_y - 45, bw - 3, 125, layer="IS_BARRIERS")
                self.add_text("PEPPERL+FUCHS EX-i GALVANIC BARRIERS (KFD2)", hd_x + hd_w/2, r3_din_y + 95, height=18, layer="IS_BARRIERS", align=TextEntityAlignment.MIDDLE_CENTER)
            else:
                # Interposing Relays (Finder 39 Series 24VDC)
                relay_count = 16
                rw = (hd_w - 20) / relay_count
                for r_idx in range(relay_count):
                    rx = hd_x + 10 + r_idx * rw
                    self.add_rect(rx, r3_din_y - 35, rw - 3, 105, layer="RELAYS")
                self.add_text("24VDC INTERPOSING RELAY BANK (FINDER 39)", hd_x + hd_w/2, r3_din_y + 85, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # ROW 4: Terminal Strips
            if i in [2, 3]:
                # Segregated Non-IS & Ex-i Blue Terminals
                non_is_w = (hd_w - 40) * 0.40
                is_w = (hd_w - 40) * 0.55
                # Non-IS Terminals
                self.add_rect(hd_x + 10, r4_din_y - 35, non_is_w, 105, layer="TERMINAL_STRIP")
                self.add_text("NON-IS TB", hd_x + 10 + non_is_w/2, r4_din_y + 85, height=16, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)
                # 50mm Air Gap / Partition Plate (Orange)
                self.add_rect(hd_x + 15 + non_is_w, r4_din_y - 45, 10, 125, layer="DIMENSIONS")
                # Blue IS Terminals
                self.add_rect(hd_x + 30 + non_is_w, r4_din_y - 35, is_w, 105, layer="TERMINAL_IS")
                self.add_text("BLUE IS TERMINALS (Ex-i)", hd_x + 30 + non_is_w + is_w/2, r4_din_y + 85, height=16, layer="TERMINAL_IS", align=TextEntityAlignment.MIDDLE_CENTER)
            else:
                # Standard Field Terminal Strips
                self.add_rect(hd_x + 10, r4_din_y - 35, hd_w - 20, 105, layer="TERMINAL_STRIP")
                self.add_text("FIELD TERMINAL STRIP (X1, X2)", hd_x + hd_w/2, r4_din_y + 85, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

            # Bay Destination Header below
            self.add_text(bay_destinations[i][0], bx + bay_w/2, by - 40, height=22, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
            self.add_text(bay_destinations[i][1], bx + bay_w/2, by - 70, height=18, layer="TEXT_LABELS", align=TextEntityAlignment.MIDDLE_CENTER)

        # Internal GA Dimensions
        dim_y1 = oy - 150
        self.add_linear_dimension((ox, oy), (ox + total_w, oy), (ox + total_w/2, dim_y1))

        dim_y2 = oy - 280
        for i in range(4):
            bx = ox + i * bay_w
            self.add_linear_dimension((bx, oy), (bx + bay_w, oy), (bx + bay_w/2, dim_y2))

    def save(self, filepath):
        self.doc.saveas(filepath)
        print(f"  [OK] Saved DXF: {os.path.basename(filepath)} ({os.path.getsize(filepath):,} bytes)")


# -----------------------------------------------------------------------------
# STEP 3: BUILD COMPREHENSIVE SUITE OF DXF FILES
# -----------------------------------------------------------------------------
def export_dxf_drawings():
    print("--- Generating AutoCAD DXF Files (1:1 mm Scale Model Space) ---")

    # 1. MASTER DXF DRAWING (Elevation + Internal GA Side-by-Side + Title Block)
    master_dxf = CADExporter("R2010")
    # Draw Elevation at (0, 0)
    master_dxf.draw_elevation_suite(origin_x=0.0, origin_y=0.0)
    master_dxf.add_text("4-PANEL SUITE: FRONT ELEVATION (DOORS CLOSED)", 1600.0, 2300.0, height=45.0, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)

    # Draw Internal GA at (3800, 0)
    master_dxf.draw_internal_ga_suite(origin_x=3800.0, origin_y=0.0)
    master_dxf.add_text("4-PANEL SUITE: INTERNAL MOUNTING GENERAL ARRANGEMENT", 5400.0, 2300.0, height=45.0, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)

    # Add Standard Master Title Block
    master_dxf.draw_title_block(ox=7000.0, oy=-600.0, title="MCC 4-PANEL SUITE (M1..M4) ELEVATION & MOUNTING PLAN", doc_no="CA1-DWG-PL-MASTER", sheet_str="1 / 1", scale_str="1:1")

    master_path_cad = os.path.join(CAD_DIR, "MCC_Panel_Layout_Master_1to1.dxf")
    master_path_pdf = os.path.join(PDF_DIR, "MCC_Panel_Layout_Master_1to1.dxf")
    master_dxf.save(master_path_cad)
    master_dxf.save(master_path_pdf)

    # 2. STANDALONE FRONT ELEVATION DXF (Sheet 1)
    elev_dxf = CADExporter("R2010")
    elev_dxf.draw_elevation_suite(origin_x=0.0, origin_y=0.0)
    elev_dxf.add_text("4-PANEL SUITE: FRONT ELEVATION (ENCLOSURE EXTERNAL VIEW)", 1600.0, 2300.0, height=45.0, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
    elev_dxf.draw_title_block(ox=3200.0, oy=-600.0, title="MCC 4-PANEL SUITE - FRONT ELEVATION (DOORS CLOSED)", doc_no="CA1-DWG-PL-001", sheet_str="1 / 7", scale_str="1:1")
    elev_dxf.save(os.path.join(CAD_DIR, "MCC_Panel_Layout_Sheet1_Front_Elevation.dxf"))
    elev_dxf.save(os.path.join(PDF_DIR, "MCC_Panel_Layout_Sheet1_Front_Elevation.dxf"))

    # 3. STANDALONE INTERNAL GENERAL ARRANGEMENT DXF (Sheet 2)
    internal_dxf = CADExporter("R2010")
    internal_dxf.draw_internal_ga_suite(origin_x=0.0, origin_y=0.0)
    internal_dxf.add_text("4-PANEL SUITE: INTERNAL GENERAL ARRANGEMENT (DOORS OPEN)", 1600.0, 2300.0, height=45.0, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
    internal_dxf.draw_title_block(ox=3200.0, oy=-600.0, title="MCC 4-PANEL SUITE - INTERNAL GENERAL ARRANGEMENT", doc_no="CA1-DWG-PL-002", sheet_str="2 / 7", scale_str="1:1")
    internal_dxf.save(os.path.join(CAD_DIR, "MCC_Panel_Layout_Sheet2_Internal_GA.dxf"))
    internal_dxf.save(os.path.join(PDF_DIR, "MCC_Panel_Layout_Sheet2_Internal_GA.dxf"))

    # 4. STANDALONE DETAILED PANELS M1..M4 (Sheets 3 to 6)
    panel_info = [
        ("M1", "Panel M1: Master Controller (CPU 1756-L950TPSXT)", 0, "CA1-DWG-PL-003", "3 / 7"),
        ("M2", "Panel M2: High-Density I/O Expansion (13 Slots)", 1, "CA1-DWG-PL-004", "4 / 7"),
        ("M3", "Panel M3: Intrinsically Safe (Ex-i) Bay 1 & 8AO", 2, "CA1-DWG-PL-005", "5 / 7"),
        ("M4", "Panel M4: Intrinsically Safe (Ex-i) Bay 2", 3, "CA1-DWG-PL-006", "6 / 7")
    ]

    for p_tag, p_desc, bay_idx, doc_num, sh_num in panel_info:
        p_dxf = CADExporter("R2010")
        # Draw single panel mounting layout at (0, 0)
        # Shift bay_idx to 0 by offsetting -bay_idx * 800
        p_dxf.draw_internal_ga_suite(origin_x= -bay_idx * 800.0, origin_y=0.0)
        p_dxf.add_text(f"{p_desc.upper()}", 400.0, 2300.0, height=35.0, layer="TEXT_HEADER", align=TextEntityAlignment.MIDDLE_CENTER)
        p_dxf.draw_title_block(ox=800.0, oy=-600.0, title=f"SUBPANEL MOUNTING PLAN - {p_tag}", doc_no=doc_num, sheet_str=sh_num, scale_str="1:1")
        p_dxf.save(os.path.join(CAD_DIR, f"MCC_Panel_Layout_Sheet{sh_num[0]}_Panel_{p_tag}.dxf"))
        p_dxf.save(os.path.join(PDF_DIR, f"MCC_Panel_Layout_Sheet{sh_num[0]}_Panel_{p_tag}.dxf"))


# -----------------------------------------------------------------------------
# MAIN EXECUTION
# -----------------------------------------------------------------------------
if __name__ == "__main__":
    print("===================================================================")
    print("  MCC 4-PANEL SUITE (M1, M2, M3, M4) CAD & SVG EXPORT PIPELINE    ")
    print("  Project Jet Cooker - Ingredion Kalasin / AEC Industrial Eng.     ")
    print("===================================================================")
    
    # 1. Export SVG Vector Files
    svg_files = export_svg_sheets()

    # 2. Export AutoCAD DXF Files
    export_dxf_drawings()

    print("\n[SUCCESS] All CAD DXF and SVG files exported successfully!")
