#!/usr/bin/env python3
"""
========================================================================================
Script: convert_pf_datasheet_to_dxf.py
Purpose: Convert Pepperl+Fuchs Datasheet '260436_eng.pdf' (HiCTB16 Termination Board)
         to high-fidelity DXF drawing files (per page + master multi-page layout).
========================================================================================
"""

import os
import sys
import pymupdf
import ezdxf

sys.stdout.reconfigure(encoding='utf-8')

def bezier_sample(p1, p2, p3, p4, num_samples=8):
    points = []
    for i in range(num_samples + 1):
        t = i / float(num_samples)
        mt = 1.0 - t
        x = (mt**3)*p1[0] + 3*(mt**2)*t*p2[0] + 3*mt*(t**2)*p3[0] + (t**3)*p4[0]
        y = (mt**3)*p1[1] + 3*(mt**2)*t*p2[1] + 3*mt*(t**2)*p3[1] + (t**3)*p4[1]
        points.append((x, y))
    return points

def convert_pdf_page_to_dxf(page, doc_dxf, x_offset=0.0, y_offset=0.0, scale=1.0):
    msp = doc_dxf.modelspace()
    rect = page.rect
    H = rect.height
    W = rect.width

    for layer_name, color in [("OUTLINE", 7), ("SCHEMATICS", 7), ("TEXT", 7), ("BORDER", 2)]:
        if layer_name not in doc_dxf.layers:
            doc_dxf.layers.add(layer_name, color=color)

    # 1. Page Border Frame
    border_pts = [
        (0.0 + x_offset, 0.0 + y_offset),
        (W * scale + x_offset, 0.0 + y_offset),
        (W * scale + x_offset, H * scale + y_offset),
        (0.0 + x_offset, H * scale + y_offset),
        (0.0 + x_offset, 0.0 + y_offset)
    ]
    msp.add_lwpolyline(border_pts, dxfattribs={'layer': 'BORDER', 'color': 2})

    # 2. Vector Drawings
    drawings = page.get_drawings()
    for draw in drawings:
        color = draw.get("color")
        layer = "SCHEMATICS"
        dxf_color = 7
        if color:
            r, g, b = [int(c * 255) for c in color[:3]]
            if r > 200 and g < 100 and b < 100: dxf_color = 1
            elif r < 100 and g > 200 and b < 100: dxf_color = 3
            elif r < 100 and g < 100 and b > 200: dxf_color = 5
            elif r > 200 and g > 200 and b < 100: dxf_color = 2
            elif r > 200 and g < 100 and b > 200: dxf_color = 6

        items = draw.get("items", [])
        for item in items:
            itype = item[0]
            if itype == "l":
                p1, p2 = item[1], item[2]
                x1 = p1.x * scale + x_offset
                y1 = (H - p1.y) * scale + y_offset
                x2 = p2.x * scale + x_offset
                y2 = (H - p2.y) * scale + y_offset
                msp.add_line((x1, y1), (x2, y2), dxfattribs={'layer': layer, 'color': dxf_color})

            elif itype == "re":
                r = item[1]
                x0 = r.x0 * scale + x_offset
                y0 = (H - r.y1) * scale + y_offset
                x1 = r.x1 * scale + x_offset
                y1 = (H - r.y0) * scale + y_offset
                pts = [(x0, y0), (x1, y0), (x1, y1), (x0, y1), (x0, y0)]
                msp.add_lwpolyline(pts, dxfattribs={'layer': layer, 'color': dxf_color})

            elif itype == "c":
                p1, p2, p3, p4 = item[1], item[2], item[3], item[4]
                pts = bezier_sample((p1.x, p1.y), (p2.x, p2.y), (p3.x, p3.y), (p4.x, p4.y), num_samples=8)
                dxf_pts = [(pt[0] * scale + x_offset, (H - pt[1]) * scale + y_offset) for pt in pts]
                msp.add_lwpolyline(dxf_pts, dxfattribs={'layer': layer, 'color': dxf_color})

            elif itype == "qu":
                q = item[1]
                pts = [
                    (q.ul.x * scale + x_offset, (H - q.ul.y) * scale + y_offset),
                    (q.ur.x * scale + x_offset, (H - q.ur.y) * scale + y_offset),
                    (q.lr.x * scale + x_offset, (H - q.lr.y) * scale + y_offset),
                    (q.ll.x * scale + x_offset, (H - q.ll.y) * scale + y_offset),
                    (q.ul.x * scale + x_offset, (H - q.ul.y) * scale + y_offset)
                ]
                msp.add_lwpolyline(pts, dxfattribs={'layer': layer, 'color': dxf_color})

    # 3. Text Extraction
    text_page = page.get_text("rawdict")
    for block in text_page.get("blocks", []):
        if block.get("type") == 0:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    txt = span.get("text", "").strip()
                    if not txt: continue
                    
                    origin = span.get("origin")
                    size = max(span.get("size", 6.0) * scale, 2.0)
                    
                    tx = origin[0] * scale + x_offset
                    ty = (H - origin[1]) * scale + y_offset
                    
                    clean_txt = "".join(c for c in txt if ord(c) >= 32)
                    if clean_txt:
                        try:
                            msp.add_text(
                                clean_txt,
                                dxfattribs={
                                    'height': size * 0.7,
                                    'insert': (tx, ty),
                                    'layer': 'TEXT',
                                    'color': 7
                                }
                            )
                        except Exception:
                            pass

def main():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    pdf_path = os.path.join(base_dir, "Datasheet", "PF", "260436_eng.pdf")
    output_dir = os.path.join(base_dir, "Datasheet", "PF")

    if not os.path.exists(pdf_path):
        print(f"Error: PDF file not found at {pdf_path}")
        return

    print(f"Opening PDF document: {pdf_path}")
    doc_pdf = pymupdf.open(pdf_path)
    total_pages = len(doc_pdf)
    print(f"Total pages: {total_pages}")

    master_dxf = ezdxf.new("R2010")
    X_MARGIN = 150.0

    for i in range(total_pages):
        page = doc_pdf[i]
        page_num = i + 1

        page_dxf = ezdxf.new("R2010")
        convert_pdf_page_to_dxf(page, page_dxf, x_offset=0.0, y_offset=0.0, scale=1.0)
        
        single_filename = f"PF_260436_HiCTB16_Page_{page_num:02d}.dxf"
        single_filepath = os.path.join(output_dir, single_filename)
        page_dxf.saveas(single_filepath)
        print(f"  -> Saved {single_filename}")

        page_width = page.rect.width
        x_off = i * (page_width + X_MARGIN)
        convert_pdf_page_to_dxf(page, master_dxf, x_offset=x_off, y_offset=0.0, scale=1.0)

    master_filename = "PF_260436_HiCTB16_Master_All_Pages.dxf"
    master_filepath = os.path.join(output_dir, master_filename)
    master_dxf.saveas(master_filepath)
    print(f"\n=======================================================")
    print(f"Master DXF created successfully:")
    print(f"  {master_filepath}")
    print(f"=======================================================")

if __name__ == '__main__':
    main()
