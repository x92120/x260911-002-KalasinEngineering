#!/usr/bin/env python3
"""
========================================================================================
Project: KALASIN STARCH PLANT - JET COOKER PROJECT
Script:  generate_template_do.py
Purpose: Generate master CAD DXF Template for Digital Output modules (1756-OB32)
         directly from 'eDrawingTemplate/template_DI.dxf'.
Target:  eDrawingTemplate/template_DO.dxf
========================================================================================
"""

import os
import ezdxf

def main():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    template_di_path = os.path.join(base_dir, "eDrawingTemplate", "template_DI.dxf")
    template_do_path = os.path.join(base_dir, "eDrawingTemplate", "template_DO.dxf")

    print(f"Reading base DI template: {template_di_path}")
    doc = ezdxf.readfile(template_di_path)
    msp = doc.modelspace()

    # ---------------------------------------------------------------------------------
    # 1. Global String Replacements on All Modelspace Texts
    # ---------------------------------------------------------------------------------
    # Converts:
    #   '1756-IB32' -> '1756-OB32'
    #   'DIGITAL INPUT MODULE' -> 'DIGITAL OUTPUT MODULE'
    #   'CHASSIS 1 SLOT 4' -> 'CHASSIS 1 SLOT 8'
    #   'P1-TBDI1' -> 'P1-TBDO1'
    #   'P1-TBDI-F1' -> 'P1-TBDO-F1'
    #   'TBDI' -> 'TBDO'
    #   'ISBDI' -> 'ISBDO'
    #   'C2S3-' -> 'C1S8-'
    #   'C1S4-' -> 'C1S8-'
    text_replacements = 0
    for t in msp:
        if t.dxftype() == "TEXT":
            orig = t.dxf.text
            new = orig
            new = new.replace("1756-IB32", "1756-OB32")
            new = new.replace("DIGITAL INPUT MODULE", "DIGITAL OUTPUT MODULE")
            new = new.replace("CHASSIS 1 SLOT 4", "CHASSIS 1 SLOT 8")
            new = new.replace("ISBDI", "ISBDO")
            new = new.replace("TBDI", "TBDO")
            new = new.replace("C2S3-", "C1S8-")
            new = new.replace("C1S4-", "C1S8-")

            # Handle signal names IN- -> OUT-
            if new.startswith("IN-"):
                new = "OUT-" + new[3:]
            elif new == "IN19":
                new = "OUT-19"

            if new != orig:
                t.dxf.text = new
                text_replacements += 1

    print(f"Applied general text replacements to {text_replacements} entities.")

    # ---------------------------------------------------------------------------------
    # 2. Update Power Pins & Wire Colors (1756-OB32 Sourcing Output Power Layout)
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
    # 3. Clone and Update Terminal Block Definition & Inserts (xDI_3_Term -> xDO_3_Term)
    # ---------------------------------------------------------------------------------
    if "xDI_3_Term" in doc.blocks:
        di_block = doc.blocks.get("xDI_3_Term")
        if "xDO_3_Term" not in doc.blocks:
            do_block = doc.blocks.new(name="xDO_3_Term", base_point=di_block.base_point)
            for e in di_block:
                do_block.add_entity(e.copy())
            print("Created new block definition 'xDO_3_Term'.")

    # Update all 32 xDI_3_Term block inserts to xDO_3_Term
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
    # 4. Save New Master DO Template
    # ---------------------------------------------------------------------------------
    doc.saveas(template_do_path)
    file_size = os.path.getsize(template_do_path)
    print(f"\n=======================================================")
    print(f"SUCCESS: Generated DO Template DXF:")
    print(f"  Path: {template_do_path}")
    print(f"  Size: {file_size:,} bytes")
    print(f"=======================================================")

if __name__ == "__main__":
    main()
