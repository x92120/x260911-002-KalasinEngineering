#!/usr/bin/env python3
"""
========================================================================================
Project: KALASIN STARCH PLANT - JET COOKER PROJECT
Script:  generate_do_slot_dxf_drawings.py
Purpose: Generate individual CAD DXF Electrical Wiring Drawings (one slot per DXF)
         strictly for Digital Output modules (1756-OB32) across Chassis C1 to C4 (and C5)
         using the master template 'eDrawingTemplate/template_DO.dxf' (with SPDT Relay).
Source:  03_IO_Lists_and_Schedules/IO_List-By_SlotConfig.xlsx
Outputs: 02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/DO_Slot_Drawings/
========================================================================================
"""

import os
import openpyxl
import ezdxf

def main():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    template_path = os.path.join(base_dir, "eDrawingTemplate", "template_DO.dxf")
    excel_path = os.path.join(base_dir, "03_IO_Lists_and_Schedules", "IO_List-By_SlotConfig.xlsx")
    out_dir = os.path.join(base_dir, "02_Electrical_and_eDrawing", "03_CAD_Exports_DXF_SVG", "DO_Slot_Drawings")

    os.makedirs(out_dir, exist_ok=True)

    print(f"Loading Excel workbook: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # Digital Output slots across Chassis C1 to C5
    target_slots = [
        "C1S8", "C1S9",
        "C2S6", "C2S7", "C2S8",
        "C3S4", "C3S5",
        "C4S4",
        "C5S3"
    ]

    print(f"Target DO slots count: {len(target_slots)}")

    even_pins = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36]
    odd_pins = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35]

    generated_files = []

    for slot_key in target_slots:
        if slot_key not in wb.sheetnames:
            print(f"Warning: Sheet {slot_key} not found in workbook, skipping.")
            continue

        sheet = wb[slot_key]
        chassis_str = slot_key[:2]   # 'C1', 'C2', etc.
        slot_num_str = slot_key[3:]  # '8', '6', etc.
        chassis_num = slot_key[1]

        # Extract Destination Area
        rack_meta = str(sheet.cell(row=3, column=1).value or "")
        dest_area = "CA1 / FIELD"
        if "DESTINATION AREA:" in rack_meta:
            dest_area = rack_meta.split("DESTINATION AREA:")[1].strip()

        # Parse pins data from Excel (Rows 6 to 41, 36 pins total)
        pin_data = {}
        tb_name = ""
        active_count = 0
        spare_count = 0

        for r in range(6, 42):
            pin_num = sheet.cell(row=r, column=1).value
            if pin_num is None:
                continue
            pin_num = int(pin_num)
            term_desc = str(sheet.cell(row=r, column=2).value or "")
            tb_col = str(sheet.cell(row=r, column=3).value or "")
            plc_tag_side = str(sheet.cell(row=r, column=4).value or "")
            term_tag_side = str(sheet.cell(row=r, column=5).value or "")
            dest = str(sheet.cell(row=r, column=6).value or "")
            plc_tag = str(sheet.cell(row=r, column=7).value or "")
            inst_tag = str(sheet.cell(row=r, column=8).value or "")
            inst_desc = str(sheet.cell(row=r, column=9).value or "")
            status = str(sheet.cell(row=r, column=10).value or "").upper()

            if "-" in tb_col and tb_col not in ("0VDC", "+24VDC") and not tb_name:
                tb_name = tb_col.rsplit("-", 1)[0]

            if status == "ACTIVE":
                active_count += 1
            elif status == "SPARE":
                spare_count += 1

            pin_data[pin_num] = {
                "desc": term_desc,
                "tb_col": tb_col,
                "plc_tag_side": plc_tag_side,
                "term_tag_side": term_tag_side,
                "dest": dest,
                "plc_tag": plc_tag,
                "inst_tag": inst_tag,
                "inst_desc": inst_desc,
                "status": status,
            }

        if not tb_name:
            tb_map = {
                "C1S8": "P1-TBRL1", "C1S9": "P1-TBRL2",
                "C2S6": "P2-TBRL1", "C2S7": "P2-TBRL2", "C2S8": "P2-TBRL3",
                "C3S4": "P3-TBRL1", "C3S5": "P3-TBRL2",
                "C4S4": "P4-TBRL1",
                "C5S3": "P5-TBRL1",
            }
            tb_name = tb_map.get(slot_key, f"P{chassis_num}-TBRL1")

        print(f"Processing {slot_key} (TB: {tb_name}, Active: {active_count}, Spare: {spare_count})...")

        # Load fresh DO template (with SPDT Relay)
        doc = ezdxf.readfile(template_path)
        msp = doc.modelspace()

        # ---------------------------------------------------------------------------------
        # 1. Update Title Text
        # ---------------------------------------------------------------------------------
        for t in msp:
            if t.dxftype() == "TEXT" and "CHASSIS" in t.dxf.text and "SLOT" in t.dxf.text:
                t.dxf.text = f"CHASSIS {chassis_str} SLOT {slot_num_str}"
                break

        # ---------------------------------------------------------------------------------
        # 2. Update Relay Terminal Block Headers across Drawing Views
        # ---------------------------------------------------------------------------------
        for t in msp:
            if t.dxftype() == "TEXT":
                txt = t.dxf.text
                if "TBRL" in txt or "TBDO" in txt or "P1-TBRL1" in txt:
                    t.dxf.text = tb_name

        # ---------------------------------------------------------------------------------
        # 3. Update PLC Module Wire Tags (Left and Right columns)
        # ---------------------------------------------------------------------------------
        left_tags = [t for t in msp if t.dxftype() == "TEXT" and 4000 < t.dxf.insert[0] < 6000 and 34000 < t.dxf.insert[1] < 64000]
        left_tags.sort(key=lambda t: -t.dxf.insert[1])

        right_tags = [t for t in msp if t.dxftype() == "TEXT" and 21000 < t.dxf.insert[0] < 24000 and 34000 < t.dxf.insert[1] < 64000]
        right_tags.sort(key=lambda t: -t.dxf.insert[1])

        if len(left_tags) == 18 and len(right_tags) == 18:
            for pin, t in zip(even_pins, left_tags):
                if pin in pin_data:
                    p_info = pin_data[pin]
                    if pin in (18, 36):
                        t.dxf.text = f"{slot_key}-{pin}/+24VDC"
                    else:
                        t.dxf.text = p_info["plc_tag_side"]

            for pin, t in zip(odd_pins, right_tags):
                if pin in pin_data:
                    p_info = pin_data[pin]
                    if pin in (17, 35):
                        t.dxf.text = f"{slot_key}-{pin}/0VDC"
                    else:
                        t.dxf.text = p_info["plc_tag_side"]

        # ---------------------------------------------------------------------------------
        # 4. Update Terminal Block Inserts (xDO_3_Term)
        # ---------------------------------------------------------------------------------
        term_to_pin = {}
        for k in range(1, 17):
            term_to_pin[k] = k
        for k in range(17, 33):
            term_to_pin[k] = k + 2

        do_inserts = [e for e in msp if e.dxftype() == "INSERT" and e.dxf.name in ("xDO_3_Term", "xDI_3_Term")]
        do_inserts.sort(key=lambda e: -e.dxf.insert[1])

        for i, ins in enumerate(do_inserts):
            k = i + 1
            pin = term_to_pin.get(k)
            p_info = pin_data.get(pin, {})
            tag_name = p_info.get("inst_tag", f"DO_{k-1}")
            if not tag_name or tag_name in ("-", "Spare"):
                tag_name = f"SPARE_{k-1}"
            desc_text = p_info.get("inst_desc", f"Digital Output Channel {k-1}")
            if desc_text in ("0", "-", "None"):
                desc_text = f"DO Channel {k-1} Relay Circuit"

            for attrib in ins.attribs:
                if attrib.dxf.tag == "PLC_POINT":
                    attrib.dxf.text = f"{slot_key}-{pin}"
                elif attrib.dxf.tag == "TAG":
                    attrib.dxf.text = str(tag_name)
                elif attrib.dxf.tag == "Description":
                    attrib.dxf.text = str(desc_text)[:28]

        # ---------------------------------------------------------------------------------
        # 5. Add Professional Title Block & Channel Schedule in Lower-Left Quadrant
        #    Coordinates: X: 3200 to 32000, Y: 10500 to 34500
        # ---------------------------------------------------------------------------------
        bx1, bx2 = 3200.0, 32000.0
        by1, by2 = 10500.0, 34500.0

        # Outer Frame
        msp.add_lwpolyline(
            [(bx1, by1), (bx2, by1), (bx2, by2), (bx1, by2), (bx1, by1)],
            dxfattribs={"layer": "ASHADE", "color": 7}
        )
        msp.add_lwpolyline(
            [(bx1 + 100, by1 + 100), (bx2 - 100, by1 + 100), (bx2 - 100, by2 - 100), (bx1 + 100, by2 - 100), (bx1 + 100, by1 + 100)],
            dxfattribs={"layer": "ASHADE", "color": 7}
        )

        # Title Card Header Box (Y: 31200 to 34300)
        msp.add_line((bx1 + 100, 31200), (bx2 - 100, 31200), dxfattribs={"layer": "ASHADE", "color": 7})
        msp.add_text(
            "INGREDION (THAILAND) CO., LTD.  |  KALASIN STARCH PLANT - JET COOKER PROJECT",
            dxfattribs={"layer": "0", "height": 340.0, "style": "Tahoma", "color": 2, "insert": (bx1 + 400, 33600, 0)}
        )
        msp.add_text(
            f"PLC I/O WIRING SCHEMATIC --- CHASSIS {chassis_str} SLOT {slot_num_str} (1756-OB32 DIGITAL OUTPUT)",
            dxfattribs={"layer": "0", "height": 380.0, "style": "Tahoma", "color": 3, "insert": (bx1 + 400, 32800, 0)}
        )
        msp.add_text(
            f"RELAY BLOCK: {tb_name} (PLC-RSC-24DC/21 SPDT)  |  LOCATION: {dest_area}  |  CHANNELS: 32 (ACTIVE: {active_count}, SPARE: {spare_count})",
            dxfattribs={"layer": "0", "height": 280.0, "style": "Tahoma", "color": 7, "insert": (bx1 + 400, 32000, 0)}
        )
        msp.add_text(
            f"DRAWING NO: KAL-JC-DO-{slot_key}   |   REV: 03   |   SYSTEM: CONTROLLOGIX 1756   |   RELAY: 24VDC COIL, 1 C/O SPDT",
            dxfattribs={"layer": "0", "height": 240.0, "style": "Tahoma", "color": 4, "insert": (bx1 + 400, 31450, 0)}
        )

        # Schedule Table Header (Y: 30400 to 31200)
        col_x = [
            bx1 + 200,   # CH (index 0)
            bx1 + 1800,  # PIN (index 1)
            bx1 + 3600,  # RELAY TB NO (index 2)
            bx1 + 7500,  # PLC WIRE TAG (index 3)
            bx1 + 13000, # FIELD ACTUATOR TAG (index 4)
            bx1 + 18500, # SERVICE / DESCRIPTION (index 5)
            bx1 + 25800  # STATUS (index 6)
        ]
        msp.add_line((bx1 + 100, 30400), (bx2 - 100, 30400), dxfattribs={"layer": "ASHADE", "color": 7})

        headers = ["CH", "PIN", "RELAY TB NO", "PLC WIRE TAG", "ACTUATOR TAG", "SERVICE / DESCRIPTION", "STATUS"]
        for hx, htext in zip(col_x, headers):
            msp.add_text(
                htext,
                dxfattribs={"layer": "0", "height": 240.0, "style": "Tahoma", "color": 2, "insert": (hx, 30650, 0)}
            )

        # 32 Channel Rows (Y from 30400 down to 11200)
        row_h = 590.0
        y_cursor = 30400.0 - row_h

        for k in range(1, 33):
            pin = term_to_pin[k]
            ch_num = k - 1
            p_info = pin_data.get(pin, {})
            tb_no = p_info.get("tb_col", f"{tb_name}-{k}")
            wire_tag = p_info.get("plc_tag_side", f"{slot_key}-{pin}/{tb_no}")
            inst_tag = p_info.get("inst_tag", "-")
            inst_desc = p_info.get("inst_desc", "-")
            status = p_info.get("status", "SPARE")

            if inst_desc in ("0", "None", "-") and status == "ACTIVE":
                inst_desc = f"Actuator Loop {inst_tag}"
            elif len(inst_desc) > 36:
                inst_desc = inst_desc[:34] + ".."

            msp.add_line((bx1 + 100, y_cursor), (bx2 - 100, y_cursor), dxfattribs={"layer": "ASHADE", "color": 8})

            y_text = y_cursor + 180.0
            stat_color = 3 if status == "ACTIVE" else (1 if status == "COMMON" else 8)

            msp.add_text(f"OUT-{ch_num:02d}", dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 7, "insert": (col_x[0], y_text, 0)})
            msp.add_text(f"Pin {pin:02d}", dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 7, "insert": (col_x[1], y_text, 0)})
            msp.add_text(tb_no, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 4, "insert": (col_x[2], y_text, 0)})
            msp.add_text(wire_tag, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 4, "insert": (col_x[3], y_text, 0)})
            msp.add_text(inst_tag, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 7, "insert": (col_x[4], y_text, 0)})
            msp.add_text(inst_desc, dxfattribs={"layer": "0", "height": 200.0, "style": "Tahoma", "color": 7, "insert": (col_x[5], y_text, 0)})
            msp.add_text(status, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": stat_color, "insert": (col_x[6], y_text, 0)})

            y_cursor -= row_h

        # Save generated DXF
        out_filename = f"{slot_key}_DO_Wiring_Diagram.dxf"
        out_filepath = os.path.join(out_dir, out_filename)
        doc.saveas(out_filepath)
        generated_files.append(out_filepath)

        print(f"  -> Generated: {out_filepath} ({os.path.getsize(out_filepath):,} bytes)")

    print(f"\n=======================================================")
    print(f"Successfully generated {len(generated_files)} DO Wiring DXF Drawings!")
    print(f"Output directory: {out_dir}")
    print(f"=======================================================")

if __name__ == "__main__":
    main()
