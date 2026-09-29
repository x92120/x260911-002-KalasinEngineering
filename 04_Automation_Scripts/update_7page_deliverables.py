#!/usr/bin/env python3
"""
Update Complete 7-Page Deliverables:
- Page 1: Executive Cover & System Overview
- Page 2: Visual Multi-Tier Automation Architecture Diagram
- Page 3: Industrial Network Topology (CPwE / DLR) with Software Overlay
- Page 4: Hardware & System Equipment Bill of Materials (BOM)
- Page 5: Workstation Functional Allocation Matrix (OWS 1, 2, 3 & EWS)
- Page 6: High Availability, Redundancy Strategy & Fault Tolerance Analysis
- Page 7: Comprehensive Software & License Bill of Materials (BOM) with Catalog Numbers

Outputs:
1. Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf (7 Pages)
2. Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx (7 Slides)
3. Automation_Architecture_Project_SPRINT_Kalasin.pptx (7 Slides)
4. Software_License_BOM_Slide_Light.png
"""

import os
import shutil
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import fitz
import io

OUT_DIR_PROJECT = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUT_DIR_WORKSPACE = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"

PDF_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf")
PDF_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pdf")
PPTX_LIGHT_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx")
PPTX_LIGHT_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin_Light.pptx")
PPTX_DARK_1 = os.path.join(OUT_DIR_PROJECT, "Automation_Architecture_Project_SPRINT_Kalasin.pptx")
PPTX_DARK_2 = os.path.join(OUT_DIR_WORKSPACE, "Automation_Architecture_Project_SPRINT_Kalasin.pptx")

import generate_light_architecture_slides as gls
import generate_all_architecture_deliverables as gad
import render_page7_bom as r7

def main():
    print("=== Generating 7-Page High-Resolution Slides Images ===")
    base_slides = gls.render_slide_images() # [Slide1_Cover, Slide2_Arch, Slide3_HW_BOM, Slide4_OWS_Alloc, Slide5_Redundancy]
    net_slide = gad.render_network_diagram_light() # Network Diagram
    bom7_slide = r7.render_software_bom_slide_light() # Page 7 Software BOM

    # Order of 7 Pages:
    # Page 1: Cover
    # Page 2: System Architecture Diagram
    # Page 3: Network Topology Diagram (with software annotations)
    # Page 4: Hardware BOM Table
    # Page 5: Workstation Allocation (OWS-01/02/03 & EWS)
    # Page 6: Redundancy & High Availability Strategy
    # Page 7: Complete Software & License Bill of Materials (BOM)
    seven_slides = [
        base_slides[0], # Page 1: Cover
        base_slides[1], # Page 2: Architecture
        net_slide,      # Page 3: Network Topology
        base_slides[2], # Page 4: Hardware BOM
        base_slides[3], # Page 5: Workstation Allocation
        base_slides[4], # Page 6: Redundancy & HA
        bom7_slide      # Page 7: Comprehensive Software & License BOM
    ]

    print(f"Total compiled pages: {len(seven_slides)}")

    # 1. Compile 7-Page PDF
    doc = fitz.open()
    page_w, page_h = 960, 540 # 16:9 widescreen
    for idx, im in enumerate(seven_slides):
        page = doc.new_page(width=page_w, height=page_h)
        buf = io.BytesIO()
        im.save(buf, format="PNG")
        page.insert_image(fitz.Rect(0, 0, page_w, page_h), stream=buf.getvalue())
        print(f"  -> Added Page {idx+1} to PDF")

    temp_pdf = "/tmp/Automation_Architecture_7Page.pdf"
    doc.save(temp_pdf, garbage=4, deflate=True)
    doc.close()

    shutil.copyfile(temp_pdf, PDF_LIGHT_1)
    shutil.copyfile(temp_pdf, PDF_LIGHT_2)
    size_mb = os.path.getsize(PDF_LIGHT_1) / (1024 * 1024)
    print(f"✅ Saved 7-Page PDF to:\n  - {PDF_LIGHT_1} ({size_mb:.2f} MB)")

    # 2. Build 7-Slide Light PPTX
    prs = pptx.Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for idx, im in enumerate(seven_slides):
        s = prs.slides.add_slide(blank_layout)
        tmp_slide_png = f"/tmp/slide_{idx+1}.png"
        im.save(tmp_slide_png)
        s.shapes.add_picture(tmp_slide_png, Inches(0.2), Inches(0.2), width=Inches(12.933))
        print(f"  -> Added Slide {idx+1} to PPTX")

    prs.save(PPTX_LIGHT_1)
    shutil.copyfile(PPTX_LIGHT_1, PPTX_LIGHT_2)
    shutil.copyfile(PPTX_LIGHT_1, PPTX_DARK_1)
    shutil.copyfile(PPTX_LIGHT_1, PPTX_DARK_2)
    print(f"✅ Saved 7-Slide PPTX to:\n  - {PPTX_LIGHT_1}")

if __name__ == "__main__":
    main()
