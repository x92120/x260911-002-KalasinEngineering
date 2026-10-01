#!/usr/bin/env python3
"""
========================================================================================
Project: KALASIN STARCH PLANT - JET COOKER PROJECT
Script:  generate_template_do.py
Purpose: Generate master CAD DXF Template for Digital Output modules (1756-OB32)
         replacing the fuse terminal with an SPDT 24VDC Interface Relay
         (Phoenix Contact PLC-RSC-24DC/21, Part No: 2966171).
Target:  eDrawingTemplate/template_DO.dxf
========================================================================================
"""

import os
import ezdxf

def update_spdt_relay_block(b):
    """
    Replaces fuse terminal entities inside block with Phoenix Contact
    PLC-RSC-24DC/21 Series Interface Relay (SPDT, 24VDC coil).
    Terminals:
      Coil: A1(+), A2(-) with freewheeling diode
      Contacts: 11 (COM), 14 (NO), 12 (NC)
    """
    # Delete old entities
    old_ents = list(b)
    for e in old_ents:
        b.delete_entity(e)

    # 1. Outer boundary / housing (6.2mm slim DIN rail interface relay)
    b.add_lwpolyline(
        [(75215.0, 32870.0), (79130.0, 32870.0), (79130.0, 33405.0), (75215.0, 33405.0), (75215.0, 32870.0)],
        dxfattribs={"layer": "0", "color": 7}
    )

    # 2. Terminal Connection Circles
    # Left side contacts: 11 (COM), 14 (NO), 12 (NC)
    b.add_circle(center=(75483.4, 33137.3), radius=110.0, dxfattribs={"layer": "0", "color": 7}) # 11 COM
    b.add_circle(center=(75483.4, 33310.0), radius=90.0, dxfattribs={"layer": "0", "color": 7})  # 14 NO
    b.add_circle(center=(75483.4, 32965.0), radius=90.0, dxfattribs={"layer": "0", "color": 7})  # 12 NC

    # Right side coil: A1(+), A2(-)
    b.add_circle(center=(78861.5, 33260.0), radius=100.0, dxfattribs={"layer": "power", "color": 1}) # A1(+) Red
    b.add_circle(center=(78861.5, 33015.0), radius=100.0, dxfattribs={"layer": "power", "color": 7}) # A2(-) White/Black

    # 3. Relay Coil Symbol with Diagonal Slash (centered at X ~ 78000, Y ~ 33137)
    b.add_lwpolyline(
        [(77750.0, 33020.0), (78350.0, 33020.0), (78350.0, 33255.0), (77750.0, 33255.0), (77750.0, 33020.0)],
        dxfattribs={"layer": "power", "color": 1}
    )
    b.add_line((77750.0, 33020.0), (78350.0, 33255.0), dxfattribs={"layer": "power", "color": 1}) # Diagonal slash

    # Coil connection leads to A1 and A2
    b.add_line((78350.0, 33255.0), (78861.5, 33260.0), dxfattribs={"layer": "power", "color": 1}) # to A1(+)
    b.add_line((78350.0, 33020.0), (78861.5, 33015.0), dxfattribs={"layer": "power", "color": 7}) # to A2(-)

    # Freewheeling Diode symbol (parallel to coil)
    b.add_lwpolyline(
        [(77550.0, 33050.0), (77550.0, 33225.0), (77450.0, 33137.5), (77550.0, 33050.0)],
        dxfattribs={"layer": "power", "color": 1}
    )
    b.add_line((77450.0, 33050.0), (77450.0, 33225.0), dxfattribs={"layer": "power", "color": 1}) # Cathode bar
    b.add_line((77500.0, 33255.0), (77750.0, 33255.0), dxfattribs={"layer": "power", "color": 1})
    b.add_line((77500.0, 33020.0), (77750.0, 33020.0), dxfattribs={"layer": "power", "color": 1})
    b.add_line((77500.0, 33255.0), (77500.0, 33225.0), dxfattribs={"layer": "power", "color": 1})
    b.add_line((77500.0, 33020.0), (77500.0, 33050.0), dxfattribs={"layer": "power", "color": 1})

    # 4. SPDT Contact Switch Symbol
    # COM (11) lead to pivot
    b.add_line((75483.4, 33137.3), (76250.0, 33137.3), dxfattribs={"layer": "0", "color": 7})
    b.add_circle(center=(76250.0, 33137.3), radius=35.0, dxfattribs={"layer": "0", "color": 7}) # Pivot point

    # Contact pads: NO (14) and NC (12)
    b.add_circle(center=(76800.0, 33290.0), radius=35.0, dxfattribs={"layer": "0", "color": 7}) # NO pad (14)
    b.add_circle(center=(76800.0, 32985.0), radius=35.0, dxfattribs={"layer": "0", "color": 7}) # NC pad (12)

    # Switch arm (normally open angled toward 14)
    b.add_line((76250.0, 33137.3), (76760.0, 33240.0), dxfattribs={"layer": "0", "color": 7})

    # Connections from pads to terminal circles
    b.add_line((76800.0, 33290.0), (75483.4, 33310.0), dxfattribs={"layer": "0", "color": 7})
    b.add_line((76800.0, 32985.0), (75483.4, 32965.0), dxfattribs={"layer": "0", "color": 7})

    # Dashed mechanical coupling link from coil to switch arm
    b.add_line((77750.0, 33137.3), (76900.0, 33137.3), dxfattribs={"layer": "0", "color": 3, "linetype": "DASHED"})

    # 5. Terminal Pin Labels
    b.add_text("11", dxfattribs={"layer": "0", "height": 80.0, "color": 7, "insert": (75620.0, 33115.0, 0)})
    b.add_text("14", dxfattribs={"layer": "0", "height": 80.0, "color": 7, "insert": (75620.0, 33285.0, 0)})
    b.add_text("12", dxfattribs={"layer": "0", "height": 80.0, "color": 7, "insert": (75620.0, 32945.0, 0)})
    b.add_text("A1+", dxfattribs={"layer": "power", "height": 80.0, "color": 1, "insert": (78480.0, 33240.0, 0)})
    b.add_text("A2-", dxfattribs={"layer": "power", "height": 80.0, "color": 7, "insert": (78480.0, 32995.0, 0)})

    # Model & Coil Rating Texts
    b.add_text("PLC-RSC-24DC/21", dxfattribs={"layer": "0", "height": 75.0, "color": 3, "insert": (76300.0, 33330.0, 0)})
    b.add_text("24VDC", dxfattribs={"layer": "power", "height": 70.0, "color": 2, "insert": (77830.0, 33075.0, 0)})

    # 6. Preserve Terminal Number Attribute
    b.add_attdef(
        "TERMINAL_A",
        dxfattribs={"layer": "0", "height": 212.5, "color": 7, "insert": (75876.4, 33018.8, 0), "text": "1"}
    )

def main():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    template_di_path = os.path.join(base_dir, "eDrawingTemplate", "template_DI.dxf")
    template_do_path = os.path.join(base_dir, "eDrawingTemplate", "template_DO.dxf")

    print(f"Reading base DI template: {template_di_path}")
    doc = ezdxf.readfile(template_di_path)
    msp = doc.modelspace()

    # ---------------------------------------------------------------------------------
    # 1. Replace Fuse Terminal with SPDT 24VDC Interface Relay in Block Definition
    # ---------------------------------------------------------------------------------
    if "xTermFuse" in doc.blocks:
        fuse_block = doc.blocks.get("xTermFuse")
        update_spdt_relay_block(fuse_block)
        print("Updated block 'xTermFuse' with SPDT 24VDC Interface Relay symbol (PLC-RSC-24DC/21).")

    # Also register explicit block name 'xRelay_SPDT_24V'
    if "xRelay_SPDT_24V" not in doc.blocks and "xTermFuse" in doc.blocks:
        fuse_b = doc.blocks.get("xTermFuse")
        relay_b = doc.blocks.new(name="xRelay_SPDT_24V", base_point=fuse_b.base_point)
        for e in fuse_b:
            relay_b.add_entity(e.copy())
        print("Created new block definition 'xRelay_SPDT_24V'.")

    # ---------------------------------------------------------------------------------
    # 2. Global Text Replacements (Headers, TB tags, Signal Names, Wire Tags)
    # ---------------------------------------------------------------------------------
    text_replacements = 0
    for t in msp:
        if t.dxftype() == "TEXT":
            orig = t.dxf.text
            new = orig
            new = new.replace("1756-IB32", "1756-OB32")
            new = new.replace("DIGITAL INPUT MODULE", "DIGITAL OUTPUT MODULE")
            new = new.replace("CHASSIS 1 SLOT 4", "CHASSIS 1 SLOT 8")
            new = new.replace("P1-TBDI-F1", "P1-TBRL1")
            new = new.replace("P1-TBDI1", "P1-TBRL1")
            new = new.replace("TBDI1A", "TBRL1A")
            new = new.replace("ISBDI", "ISBDO")
            new = new.replace("TBDI", "TBRL")
            new = new.replace("C2S3-", "C1S8-")
            new = new.replace("C1S4-", "C1S8-")

            # Signal names IN- -> OUT-
            if new.startswith("IN-"):
                new = "OUT-" + new[3:]
            elif new == "IN19":
                new = "OUT-19"

            if new != orig:
                t.dxf.text = new
                text_replacements += 1

    print(f"Applied general text replacements to {text_replacements} entities.")

    # ---------------------------------------------------------------------------------
    # 3. Update Power Pins & Wire Colors (1756-OB32 Power Layout)
    #    Pin 17: RTN-0 (0VDC return)       -> Color: BK
    #    Pin 18: DC-0(+) (+24VDC supply)   -> Color: RD
    #    Pin 35: RTN-1 (0VDC return)       -> Color: BK
    #    Pin 36: DC-1(+) (+24VDC supply)   -> Color: RD
    # ---------------------------------------------------------------------------------
    for t in msp:
        if t.dxftype() == "TEXT":
            p = t.dxf.insert
            txt = t.dxf.text
            # Power pin signal texts
            if abs(p[1] - 48951.7) < 50.0:
                if 15000 < p[0] < 16000:  # Pin 18 (left)
                    t.dxf.text = "DC-0(+)"
                elif 18500 < p[0] < 19500:  # Pin 17 (right)
                    t.dxf.text = "RTN-0"
            elif abs(p[1] - 34992.5) < 50.0:
                if 15000 < p[0] < 16000:  # Pin 36 (left)
                    t.dxf.text = "DC-1(+)"
                elif 18500 < p[0] < 19500:  # Pin 35 (right)
                    t.dxf.text = "RTN-1"

            # Power pin wire colors
            if abs(p[1] - 49330.4) < 50.0 and 11000 < p[0] < 12000:
                t.dxf.text = "RD"  # Pin 18: DC-0(+)
            elif abs(p[1] - 35371.2) < 50.0 and 11000 < p[0] < 12000:
                t.dxf.text = "RD"  # Pin 36: DC-1(+)

            # Power pin wire tags on PLC side
            if "C1S8-18/" in txt or "C1S8-36/" in txt:
                t.dxf.text = txt.split("/")[0] + "/+24VDC"
            elif "C1S8-17/" in txt or "C1S8-35/" in txt:
                t.dxf.text = txt.split("/")[0] + "/0VDC"

    # ---------------------------------------------------------------------------------
    # 4. Clone and Update Terminal Block Definition & Inserts (xDI_3_Term -> xDO_3_Term)
    # ---------------------------------------------------------------------------------
    if "xDI_3_Term" in doc.blocks:
        di_block = doc.blocks.get("xDI_3_Term")
        if "xDO_3_Term" not in doc.blocks:
            do_block = doc.blocks.new(name="xDO_3_Term", base_point=di_block.base_point)
            for e in di_block:
                do_block.add_entity(e.copy())
            print("Created new block definition 'xDO_3_Term'.")

    di_inserts = [e for e in msp if e.dxftype() == "INSERT" and e.dxf.name in ("xDI_3_Term", "xDO_3_Term")]
    di_inserts.sort(key=lambda e: -e.dxf.insert[1])
    print(f"Updating {len(di_inserts)} terminal block inserts to xDO_3_Term...")

    for i, ins in enumerate(di_inserts):
        ins.dxf.name = "xDO_3_Term"
        k = i  # Output channel 0 to 31
        pin = (k + 1) if k < 16 else (k + 3)

        for attrib in ins.attribs:
            if attrib.dxf.tag == "PLC_POINT":
                attrib.dxf.text = f"C1S8-{pin}"
            elif attrib.dxf.tag == "TAG":
                attrib.dxf.text = f"XV-{40201 + k}"
            elif attrib.dxf.tag == "Description":
                attrib.dxf.text = f"Digital Output Channel {k}"

    # ---------------------------------------------------------------------------------
    # 5. Save New Master DO Template
    # ---------------------------------------------------------------------------------
    doc.saveas(template_do_path)
    file_size = os.path.getsize(template_do_path)
    print(f"\n=======================================================")
    print(f"SUCCESS: Generated DO Template DXF (with SPDT Relay):")
    print(f"  Path: {template_do_path}")
    print(f"  Size: {file_size:,} bytes")
    print(f"=======================================================")

if __name__ == "__main__":
    main()
