#!/usr/bin/env python3
"""
========================================================================================
Project: KALASIN STARCH PLANT - JET COOKER PROJECT
Script:  generate_di_slot_dxf_drawings.py
Purpose: Generate 15 individual CAD DXF Electrical Wiring Drawings (one slot per DXF)
         strictly for Digital Input modules (1756-IB32) across Chassis C1 to C4
         using the master template 'eDrawingTemplate/template_DI.dxf'.
Source:  03_IO_Lists_and_Schedules/IO_List-By_SlotConfig.xlsx
Outputs: 02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/DI_Slot_Drawings/
========================================================================================
"""

import os
import shutil
import openpyxl
import ezdxf

def main():
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    template_path = os.path.join(base_dir, "eDrawingTemplate", "template_DI.dxf")
    excel_path = os.path.join(base_dir, "03_IO_Lists_and_Schedules", "IO_List-By_SlotConfig.xlsx")
    out_dir = os.path.join(base_dir, "02_Electrical_and_eDrawing", "03_CAD_Exports_DXF_SVG", "DI_Slot_Drawings")

    os.makedirs(out_dir, exist_ok=True)

    print(f"Loading Excel workbook: {excel_path}")
    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # 15 Digital Input slots across Chassis C1 to C4
    target_slots = [
        "C1S4", "C1S5", "C1S6", "C1S7",
        "C2S1", "C2S2", "C2S3", "C2S4", "C2S5",
        "C3S1", "C3S2", "C3S3",
        "C4S1", "C4S2", "C4S3"
    ]

    print(f"Target DI slots count: {len(target_slots)}")

    # Pin layout sequence in template_DI.dxf:
    # Left column: even pins 2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36 (decreasing Y)
    # Right column: odd pins 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35 (decreasing Y)
    even_pins = [2, 4, 6, 8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30, 32, 34, 36]
    odd_pins = [1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23, 25, 27, 29, 31, 33, 35]

    generated_files = []

    for slot_key in target_slots:
        if slot_key not in wb.sheetnames:
            print(f"Warning: Sheet {slot_key} not found in workbook, skipping.")
            continue

        sheet = wb[slot_key]
        chassis_str = slot_key[:2]  # 'C1', 'C2', etc.
        slot_num_str = slot_key[3:]  # '4', '1', etc.
        chassis_num = slot_key[1]    # '1', '2', etc.

        # Header info
        location_header = str(sheet.cell(row=2, column=1).value or "")
        rack_meta = str(sheet.cell(row=3, column=1).value or "")

        # Extract Destination Area
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

            if "-" in tb_col and tb_col != "0VDC" and not tb_name:
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
            # Fallback based on slot key
            tb_map = {
                "C1S4": "P1-TBDI1", "C1S5": "P1-TBDI2", "C1S6": "P1-TBDI3", "C1S7": "P1-TBDI4",
                "C2S1": "P2-TBDI1", "C2S2": "P2-TBDI2", "C2S3": "P2-TBDI3", "C2S4": "P2-TBDI4", "C2S5": "P2-TBDI5",
                "C3S1": "P3-TBDI1", "C3S2": "P3-TBDI2", "C3S3": "P3-TBDI3",
                "C4S1": "P4-TBDI1", "C4S2": "P4-TBDI2", "C4S3": "P4-TBDI3",
            }
            tb_name = tb_map.get(slot_key, f"P{chassis_num}-TBDI{slot_num_str}")

        print(f"Processing {slot_key} (TB: {tb_name}, Active: {active_count}, Spare: {spare_count})...")

        # Load fresh template
        doc = ezdxf.readfile(template_path)
        msp = doc.modelspace()

        # ---------------------------------------------------------------------------------
        # 1. Update Title Text
        # ---------------------------------------------------------------------------------
        for t in msp:
            if t.dxftype() == "TEXT" and "CHASSIS" in t.dxf.text:
                t.dxf.text = f"CHASSIS {chassis_str} SLOT {slot_num_str}"
                break

        # ---------------------------------------------------------------------------------
        # 2. Update Terminal Block Headers (3 occurrences in template)
        # ---------------------------------------------------------------------------------
        for t in msp:
            if t.dxftype() == "TEXT" and ("P1-TBDI1" in t.dxf.text or "TBDI" in t.dxf.text):
                t.dxf.text = tb_name

        # ---------------------------------------------------------------------------------
        # 3. Update PLC Module Wire Tags (Left and Right columns)
        # ---------------------------------------------------------------------------------
        left_tags = [t for t in msp if t.dxftype() == "TEXT" and 4000 < t.dxf.insert[0] < 6000 and 39000 < t.dxf.insert[1] < 64000]
        left_tags.sort(key=lambda t: -t.dxf.insert[1])

        right_tags = [t for t in msp if t.dxftype() == "TEXT" and 21000 < t.dxf.insert[0] < 24000 and 39000 < t.dxf.insert[1] < 64000]
        right_tags.sort(key=lambda t: -t.dxf.insert[1])

        if len(left_tags) == 18 and len(right_tags) == 18:
            for pin, t in zip(even_pins, left_tags):
                if pin in pin_data:
                    p_info = pin_data[pin]
                    if pin in (18, 36):
                        t.dxf.text = f"{slot_key}-{pin}/0VDC"
                    else:
                        t.dxf.text = p_info["plc_tag_side"]

            for pin, t in zip(odd_pins, right_tags):
                if pin in pin_data:
                    p_info = pin_data[pin]
                    if pin in (17, 35):
                        t.dxf.text = f"{slot_key}-{pin}/0VDC"
                    else:
                        t.dxf.text = p_info["plc_tag_side"]
        else:
            print(f"  Warning: Expected 18 left and 18 right tags, found {len(left_tags)} and {len(right_tags)}")

        # ---------------------------------------------------------------------------------
        # 4. Update Terminal Block Units (Term_Group3 block inserts)
        # ---------------------------------------------------------------------------------
        inserts = [e for e in msp if e.dxftype() == "INSERT" and e.dxf.name == "Term_Group3"]
        inserts.sort(key=lambda e: -e.dxf.insert[1])

        for i, ins in enumerate(inserts):
            k = i + 1  # Terminal 1 to 32
            for attrib in ins.attribs:
                if attrib.dxf.tag == "TERMINAL_A":
                    attrib.dxf.text = f"{k}A"
                elif attrib.dxf.tag == "TERMINAL_B":
                    attrib.dxf.text = f"{k}B"

        # ---------------------------------------------------------------------------------
        # 5. Update & Complete Terminal Block Wire Tags at X = 58050.0
        # ---------------------------------------------------------------------------------
        # Map terminal k to module pin:
        # Terminals 1..16 -> Pins 1..16 (Inputs 0..15)
        # Terminals 17..32 -> Pins 19..34 (Inputs 16..31)
        term_to_pin = {}
        for k in range(1, 17):
            term_to_pin[k] = k
        for k in range(17, 33):
            term_to_pin[k] = k + 2

        # First, remove old wire tags at X = 58050.0 to prevent duplicates
        old_wire_tags = [t for t in msp if t.dxftype() == "TEXT" and abs(t.dxf.insert[0] - 58050.0) < 50.0]
        for t in old_wire_tags:
            msp.delete_entity(t)

        # Re-add all 32 terminal block wire tags with precision alignment
        for i, ins in enumerate(inserts):
            k = i + 1
            pin = term_to_pin.get(k)
            if pin in pin_data:
                tag_text = pin_data[pin]["plc_tag_side"]
            else:
                tag_text = f"{slot_key}-{pin}/{tb_name}-{k}"

            y_pos = ins.dxf.insert[1] - 641.4
            msp.add_text(
                tag_text,
                dxfattribs={
                    "layer": "0",
                    "height": 259.08,
                    "style": "Tahoma",
                    "color": 4,  # Cyan
                    "insert": (58050.0, y_pos, 0.0),
                }
            )

        # ---------------------------------------------------------------------------------
        # 6. Add Professional Title Block & Channel Schedule in Lower-Left Quadrant
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
            f"PLC I/O WIRING SCHEMATIC --- CHASSIS {chassis_str} SLOT {slot_num_str} (1756-IB32 DIGITAL INPUT)",
            dxfattribs={"layer": "0", "height": 380.0, "style": "Tahoma", "color": 3, "insert": (bx1 + 400, 32800, 0)}
        )
        msp.add_text(
            f"TERMINAL BLOCK: {tb_name}   |   LOCATION: {dest_area}   |   CHANNELS: 32 (ACTIVE: {active_count}, SPARE: {spare_count})",
            dxfattribs={"layer": "0", "height": 280.0, "style": "Tahoma", "color": 7, "insert": (bx1 + 400, 32000, 0)}
        )
        msp.add_text(
            f"DRAWING NO: KAL-JC-DI-{slot_key}   |   REV: 03   |   SYSTEM: CONTROLLOGIX 1756   |   REF EXCEL: IO_List-By_SlotConfig.xlsx",
            dxfattribs={"layer": "0", "height": 240.0, "style": "Tahoma", "color": 4, "insert": (bx1 + 400, 31450, 0)}
        )

        # Schedule Table Header (Y: 30400 to 31200)
        col_x = [
            bx1 + 200,   # CH (index 0)
            bx1 + 1800,  # PIN (index 1)
            bx1 + 3600,  # TB NO (index 2)
            bx1 + 7500,  # PLC WIRE TAG (index 3)
            bx1 + 13000, # INSTRUMENT TAG (index 4)
            bx1 + 18500, # SERVICE / DESCRIPTION (index 5)
            bx1 + 25800  # STATUS (index 6)
        ]
        msp.add_line((bx1 + 100, 30400), (bx2 - 100, 30400), dxfattribs={"layer": "ASHADE", "color": 7})

        headers = ["CH", "PIN", "TB NUMBER", "PLC WIRE TAG", "INSTRUMENT TAG", "SERVICE / DESCRIPTION", "STATUS"]
        for hx, htext in zip(col_x, headers):
            msp.add_text(
                htext,
                dxfattribs={"layer": "0", "height": 240.0, "style": "Tahoma", "color": 2, "insert": (hx, 30650, 0)}
            )

        # 32 Channel Rows (Y from 30400 down to 11200, height = 19200 / 32 = 600 per row)
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

            # Truncate description if too long
            if len(inst_desc) > 36:
                inst_desc = inst_desc[:34] + ".."

            # Alternating light grid line
            msp.add_line((bx1 + 100, y_cursor), (bx2 - 100, y_cursor), dxfattribs={"layer": "ASHADE", "color": 8})

            y_text = y_cursor + 180.0
            stat_color = 3 if status == "ACTIVE" else (1 if status == "COMMON" else 8)

            msp.add_text(f"IN-{ch_num:02d}", dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 7, "insert": (col_x[0], y_text, 0)})
            msp.add_text(f"Pin {pin:02d}", dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 7, "insert": (col_x[1], y_text, 0)})
            msp.add_text(tb_no, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 4, "insert": (col_x[2], y_text, 0)})
            msp.add_text(wire_tag, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 4, "insert": (col_x[3], y_text, 0)})
            msp.add_text(inst_tag, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": 7, "insert": (col_x[4], y_text, 0)})
            msp.add_text(inst_desc, dxfattribs={"layer": "0", "height": 200.0, "style": "Tahoma", "color": 7, "insert": (col_x[5], y_text, 0)})
            msp.add_text(status, dxfattribs={"layer": "0", "height": 220.0, "style": "Tahoma", "color": stat_color, "insert": (col_x[6], y_text, 0)})

            y_cursor -= row_h

        # Save generated DXF
        out_filename = f"{slot_key}_DI_Wiring_Diagram.dxf"
        out_filepath = os.path.join(out_dir, out_filename)
        doc.saveas(out_filepath)
        generated_files.append(out_filepath)

        print(f"  -> Generated: {out_filepath} ({os.path.getsize(out_filepath):,} bytes)")

    print(f"\n=======================================================")
    print(f"Successfully generated {len(generated_files)} DI Wiring DXF Drawings!")
    print(f"Output directory: {out_dir}")
    print(f"=======================================================")

if __name__ == "__main__":
    main()
