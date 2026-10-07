import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# Create Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Color Palette
C_NAVY_DARK = "1B365D"
C_NAVY_MED = "2C5E8A"
C_ICE_BLUE = "D9E1F2"
C_CARD_BG = "F2F5F9"
C_ZEBRA = "F9FAFC"
C_WHITE = "FFFFFF"
C_GREEN_BG = "E2EFDA"
C_GREEN_FG = "375623"
C_AMBER_BG = "FFF2CC"
C_AMBER_FG = "B25900"
C_BLUE_BG = "DDEBF7"
C_BLUE_FG = "1F4E79"
C_TOTAL_BG = "E9EEF4"
C_BORDER = "D9D9D9"

# Font Styles
f_title = Font(name="Segoe UI", size=15, bold=True, color="FFFFFF")
f_subtitle = Font(name="Segoe UI", size=10, italic=True, color="E0E6ED")
f_sec_hdr = Font(name="Segoe UI", size=11, bold=True, color=C_NAVY_DARK)
f_tbl_hdr = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
f_data = Font(name="Segoe UI", size=9, color="000000")
f_data_bold = Font(name="Segoe UI", size=9, bold=True, color="000000")
f_total = Font(name="Segoe UI", size=10, bold=True, color=C_NAVY_DARK)

f_status_pass = Font(name="Segoe UI", size=9, bold=True, color=C_GREEN_FG)
f_status_warn = Font(name="Segoe UI", size=9, bold=True, color=C_AMBER_FG)
f_status_info = Font(name="Segoe UI", size=9, bold=True, color=C_BLUE_FG)

# Fill Styles
fill_hdr = PatternFill("solid", fgColor=C_NAVY_DARK)
fill_subhdr = PatternFill("solid", fgColor=C_NAVY_MED)
fill_sec = PatternFill("solid", fgColor=C_ICE_BLUE)
fill_card = PatternFill("solid", fgColor=C_CARD_BG)
fill_zebra = PatternFill("solid", fgColor=C_ZEBRA)
fill_white = PatternFill("solid", fgColor=C_WHITE)
fill_total = PatternFill("solid", fgColor=C_TOTAL_BG)

fill_pass = PatternFill("solid", fgColor=C_GREEN_BG)
fill_warn = PatternFill("solid", fgColor=C_AMBER_BG)
fill_info = PatternFill("solid", fgColor=C_BLUE_BG)

# Border Styles
thin_side = Side(border_style="thin", color=C_BORDER)
double_bottom = Side(border_style="double", color=C_NAVY_DARK)
border_data = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
border_total = Border(left=thin_side, right=thin_side, top=thin_side, bottom=double_bottom)
border_hdr = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# Alignment Styles
align_center = Alignment(horizontal="center", vertical="center")
align_left = Alignment(horizontal="left", vertical="center")
align_right = Alignment(horizontal="right", vertical="center")
align_wrap_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

# ==============================================================================
# SHEET 1: 00_Executive_Summary
# ==============================================================================
ws1 = wb.create_sheet(title="00_Executive_Summary")
ws1.views.sheetView[0].showGridLines = True

# Header Banner
ws1.merge_cells("A1:K1")
ws1["A1"] = "KALASIN ENGINEERING  |  QUOTATION (BOQ) vs. JUNCTION BOX SCHEDULE COMPARISON REPORT"
ws1["A1"].font = f_title
ws1["A1"].fill = fill_hdr
ws1["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws1.row_dimensions[1].height = 32

ws1.merge_cells("A2:K2")
ws1["A2"] = "PROJECT: SPRINT 18K TPA SPRAY DRYER IN THAILAND  |  Quotation_R02.xlsx (BOQ D2510-775 R.2) vs. IO_List_By_Junction_Box.xlsx (Distributed Architecture)"
ws1["A2"].font = f_subtitle
ws1["A2"].fill = fill_subhdr
ws1["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws1.row_dimensions[2].height = 20

# KPI Metric Cards
kpis = [
    ("A", "B", "QUOTED CAPACITY (BOQ)", "720 Points", "39 Modules (Centralized Rack)", fill_card, Font(name="Segoe UI", size=18, bold=True, color=C_NAVY_DARK)),
    ("C", "E", "JB SCHEDULE CAPACITY", "1,274 Points", "54 Modules across 5 Chassis (+76.9%)", PatternFill("solid", fgColor="E7F1FF"), Font(name="Segoe UI", size=18, bold=True, color="004085")),
    ("F", "H", "SYSTEM ACTIVE SIGNALS", "713 Points", "Wired across 16 Field Enclosures", PatternFill("solid", fgColor="E2F0D9"), Font(name="Segoe UI", size=18, bold=True, color="385723")),
    ("I", "K", "RESERVE SPARE CAPACITY", "44.0% Spare", "561 Reserve Points (Exceeds >=20% Spec)", PatternFill("solid", fgColor="D4EDDA"), Font(name="Segoe UI", size=18, bold=True, color="155724"))
]

for start_c, end_c, title, val, sub, bg, f_val in kpis:
    start_idx = ord(start_c) - ord('A') + 1
    end_idx = ord(end_c) - ord('A') + 1
    
    # Set top-left values
    ws1[f"{start_c}4"] = title
    ws1[f"{start_c}5"] = val
    ws1[f"{start_c}6"] = sub
    
    # Merge cells
    ws1.merge_cells(f"{start_c}4:{end_c}4")
    ws1.merge_cells(f"{start_c}5:{end_c}5")
    ws1.merge_cells(f"{start_c}6:{end_c}6")
    
    for r in range(4, 7):
        for c in range(start_idx, end_idx + 1):
            cell = ws1.cell(r, c)
            cell.fill = bg
            cell.border = border_data
            if r == 4:
                cell.font = Font(name="Segoe UI", size=8, bold=True, color="555555")
                cell.alignment = align_center
            elif r == 5:
                cell.font = f_val
                cell.alignment = align_center
            elif r == 6:
                cell.font = Font(name="Segoe UI", size=8, italic=True, color="444444")
                cell.alignment = align_center
ws1.row_dimensions[4].height = 16
ws1.row_dimensions[5].height = 28
ws1.row_dimensions[6].height = 18

# Table 1: Master I/O Hardware & Capacity Comparison
ws1.merge_cells("A8:K8")
ws1["A8"] = "1. MASTER I/O POINT HARDWARE & CAPACITY RECONCILIATION"
ws1["A8"].font = f_sec_hdr
ws1["A8"].fill = fill_sec
ws1.row_dimensions[8].height = 22

t1_headers = [
    ("Signal Family", 18),
    ("Quotation Card Catalog", 22),
    ("Quoted Qty", 12),
    ("Quoted Cap", 12),
    ("JB Schedule Card Catalog", 22),
    ("JB Qty", 10),
    ("JB Total Cap", 14),
    ("Active Used", 12),
    ("Spare Pts", 12),
    ("Spare Margin", 14),
    ("Technical Variance & Scope Delta", 34)
]
for col_idx, (th, w) in enumerate(t1_headers, 1):
    c = ws1.cell(9, col_idx, th)
    c.font = f_tbl_hdr
    c.fill = fill_hdr
    c.alignment = align_center
    c.border = border_hdr
    ws1.column_dimensions[get_column_letter(col_idx)].width = w
ws1.row_dimensions[9].height = 26

t1_data = [
    ("Digital Input (DI)", "1756-IB32 (32-ch Sink)", 12, 384, "1756-IB32 (32-ch Sink/Source)", 17, 544, 334, 210, 0.3860, "+5 Modules (+160 pts / +41.7%). Fulfills >= 20% spare spec."),
    ("Digital Output (DO)", "1756-OB32 (32-ch Source)", 5, 160, "1756-OB32 (32-ch Transistor)", 10, 306, 122, 184, 0.6013, "+5 Modules (+146 pts / +91.3%). Massive expansion for valve interlocks."),
    ("Analog Input (AI)", "1756-IF8H (8-ch Isolated HART)", 18, 144, "1756-IF16 (16-ch High Density)", 26, 416, 252, 164, 0.3942, "+8 Modules (+272 pts / +188.9%). Value engineered to 16-channel card."),
    ("Analog Output (AO)", "1756-OF8H (8-ch Isolated HART)", 4, 32, "1756-OF8 (8-ch Voltage/Current)", 1, 8, 5, 3, 0.3750, "-3 Modules (-24 pts / -75.0%). Optimized for 5 active control valves."),
]

for row_idx, rdata in enumerate(t1_data, 10):
    ws1.row_dimensions[row_idx].height = 20
    ws1.cell(row_idx, 1, rdata[0]).alignment = align_left
    ws1.cell(row_idx, 2, rdata[1]).alignment = align_left
    ws1.cell(row_idx, 3, rdata[2]).alignment = align_center
    ws1.cell(row_idx, 4, rdata[3]).alignment = align_center
    ws1.cell(row_idx, 5, rdata[4]).alignment = align_left
    ws1.cell(row_idx, 6, rdata[5]).alignment = align_center
    ws1.cell(row_idx, 7, rdata[6]).alignment = align_center
    ws1.cell(row_idx, 8, rdata[7]).alignment = align_center
    ws1.cell(row_idx, 9, rdata[8]).alignment = align_center
    ws1.cell(row_idx, 10, rdata[9]).alignment = align_center
    ws1.cell(row_idx, 10).number_format = "0.0%"
    ws1.cell(row_idx, 11, rdata[10]).alignment = align_left
    
    for c in range(1, 12):
        ws1.cell(row_idx, c).font = f_data_bold if c in [1, 4, 7, 8, 10] else f_data
        ws1.cell(row_idx, c).border = border_data
        if row_idx % 2 == 1:
            ws1.cell(row_idx, c).fill = fill_zebra

# Process Total Row
r_tot = 14
ws1.row_dimensions[r_tot].height = 24
ws1.cell(r_tot, 1, "TOTAL PROCESS I/O").alignment = align_left
ws1.cell(r_tot, 2, "ControlLogix Baseline").alignment = align_left
ws1.cell(r_tot, 3, 39).alignment = align_center
ws1.cell(r_tot, 4, 720).alignment = align_center
ws1.cell(r_tot, 5, "ControlLogix Distributed").alignment = align_left
ws1.cell(r_tot, 6, 54).alignment = align_center
ws1.cell(r_tot, 7, 1274).alignment = align_center
ws1.cell(r_tot, 8, 713).alignment = align_center
ws1.cell(r_tot, 9, 561).alignment = align_center
ws1.cell(r_tot, 10, 0.4403).alignment = align_center
ws1.cell(r_tot, 10).number_format = "0.0%"
ws1.cell(r_tot, 11, "+15 Modules (+554 Pts / +76.9% Capacity Expansion)").alignment = align_left
for c in range(1, 12):
    ws1.cell(r_tot, c).font = f_total
    ws1.cell(r_tot, c).fill = fill_total
    ws1.cell(r_tot, c).border = border_total

# Power / Commons Row
r_pin = 15
ws1.row_dimensions[r_pin].height = 20
ws1.cell(r_pin, 1, "Common / Power Pins").alignment = align_left
ws1.cell(r_pin, 2, "Not Separately Listed").alignment = align_left
ws1.cell(r_pin, 3, "-").alignment = align_center
ws1.cell(r_pin, 4, "-").alignment = align_center
ws1.cell(r_pin, 5, "Dedicated 24VDC Bus Pins").alignment = align_left
ws1.cell(r_pin, 6, "-").alignment = align_center
ws1.cell(r_pin, 7, 150).alignment = align_center
ws1.cell(r_pin, 8, "-").alignment = align_center
ws1.cell(r_pin, 9, "-").alignment = align_center
ws1.cell(r_pin, 10, "-").alignment = align_center
ws1.cell(r_pin, 11, "68 DI Commons + 18 DO Commons + 52 AI Commons + 12 AO").alignment = align_left
for c in range(1, 12):
    ws1.cell(r_pin, c).font = f_data
    ws1.cell(r_pin, c).fill = fill_zebra
    ws1.cell(r_pin, c).border = border_data

# Grand Physical Terminals Row
r_grand = 16
ws1.row_dimensions[r_grand].height = 22
ws1.cell(r_grand, 1, "TOTAL PHYSICAL TERMINATIONS").alignment = align_left
ws1.cell(r_grand, 2, "720 Standard Terminals").alignment = align_left
ws1.cell(r_grand, 3, "-").alignment = align_center
ws1.cell(r_grand, 4, 720).alignment = align_center
ws1.cell(r_grand, 5, "1,424 Marshaled Pins").alignment = align_left
ws1.cell(r_grand, 6, "-").alignment = align_center
ws1.cell(r_grand, 7, 1424).alignment = align_center
ws1.cell(r_grand, 8, 713).alignment = align_center
ws1.cell(r_grand, 9, 711).alignment = align_center
ws1.cell(r_grand, 10, 0.4993).alignment = align_center
ws1.cell(r_grand, 10).number_format = "0.0%"
ws1.cell(r_grand, 11, "Complete physical wiring footprint across 16 enclosures").alignment = align_left
for c in range(1, 12):
    ws1.cell(r_grand, c).font = f_total
    ws1.cell(r_grand, c).fill = fill_sec
    ws1.cell(r_grand, c).border = border_total

# Table 2: Architecture & Controller Specification Upgrades
ws1.merge_cells("A18:K18")
ws1["A18"] = "2. SYSTEM ARCHITECTURE & CORE HARDWARE SPECIFICATION UPGRADE"
ws1["A18"].font = f_sec_hdr
ws1["A18"].fill = fill_sec
ws1.row_dimensions[18].height = 22

t2_headers = [
    ("Component / System Attribute", 26),
    ("Baseline Quotation Specification", 30),
    ("JB Schedule Implementation", 32),
    ("Engineering Rationale & Functional Advantage", 38)
]
for col_idx, (th, w) in enumerate(t2_headers, 1):
    c = ws1.cell(19, col_idx, th)
    c.font = f_tbl_hdr
    c.fill = fill_hdr
    c.alignment = align_center
    c.border = border_hdr
ws1.row_dimensions[19].height = 24

t2_data = [
    ("Process Controller (CPU)", "1x 1756-L81E (ControlLogix 5580, 3MB)", "1x 1756-L950TPSXT (Extreme Temp, 50MB, Conformal Coated)", "Handles larger distributed tag database and harsh environment (spray dryer heat/moisture)"),
    ("Network Scanner / Adapters", "4x 1756-EN2TR (100 Mbps Copper DLR)", "1756-EN4TR (1 Gbps High-Speed Fiber/Copper DLR Scanner)", "10x network bandwidth for real-time cyclic I/O transfer between 5 chassis"),
    ("System Architecture Topology", "Centralized Single Control Cabinet (CA1)", "Distributed 5-Chassis Suite (C1, C2 in CA1; C3, C4 in Tower; C5 in Slurry)", "Eliminates hundreds of long-distance multicores, prevents electrical noise coupling"),
    ("Field Enclosure Scope", "Line 143: 'JUNCTION BOX: 1 LOT' (Unspecified)", "16 Engineered Enclosures (8 SS304 Food-Grade + 4 ATEX Ex ia SS316 + 3 Panels)", "Resolves generic lump sum into certified sanitary food-grade and ATEX hazardous equipment"),
    ("Contractual Spare Compliance", "13.1% Overall Spare (Fails >=20% on DI: 6.5% & DO: 10.0%)", "44.0% Overall Spare (DI: 38.6%, DO: 60.1%, AI: 39.4%, AO: 37.5%)", "Guarantees complete compliance with plant automation expansion specifications")
]

for row_idx, rdata in enumerate(t2_data, 20):
    ws1.row_dimensions[row_idx].height = 24
    ws1.cell(row_idx, 1, rdata[0]).alignment = align_left
    ws1.cell(row_idx, 2, rdata[1]).alignment = align_left
    ws1.cell(row_idx, 3, rdata[2]).alignment = align_left
    ws1.cell(row_idx, 4, rdata[3]).alignment = align_wrap_left
    for c in range(1, 5):
        ws1.cell(row_idx, c).font = f_data_bold if c == 1 else f_data
        ws1.cell(row_idx, c).border = border_data
        if row_idx % 2 == 1:
            ws1.cell(row_idx, c).fill = fill_zebra

print("Built 00_Executive_Summary.")

# ==============================================================================
# SHEET 2: 01_Chassis_Architecture_C1_C5
# ==============================================================================
ws2 = wb.create_sheet(title="01_Chassis_Architecture_C1_C5")
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells("A1:P1")
ws2["A1"] = "DISTRIBUTED CONTROL CHASSIS (C1 TO C5) UTILIZATION & CAPACITY BREAKDOWN"
ws2["A1"].font = f_title
ws2["A1"].fill = fill_hdr
ws2["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[1].height = 28

ws2.merge_cells("A2:P2")
ws2["A2"] = "Detailed engineering breakdown of all 5 Rockwell ControlLogix racks, physical locations, signal loads, and reserve margins"
ws2["A2"].font = f_subtitle
ws2["A2"].fill = fill_subhdr
ws2["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws2.row_dimensions[2].height = 18

s2_headers = [
    ("Chassis Code", 14),
    ("Panel Tag", 12),
    ("Chassis Description & Function", 36),
    ("Physical Location / Enclosure", 32),
    ("DI Used", 10),
    ("DI Spare", 10),
    ("DI %", 10),
    ("DO Used", 10),
    ("DO Spare", 10),
    ("DO %", 10),
    ("AI Used", 10),
    ("AI Spare", 10),
    ("AI %", 10),
    ("AO Used", 10),
    ("AO Spare", 10),
    ("Total Active Pts", 16)
]
for col_idx, (th, w) in enumerate(s2_headers, 1):
    c = ws2.cell(4, col_idx, th)
    c.font = f_tbl_hdr
    c.fill = fill_hdr
    c.alignment = align_center
    c.border = border_hdr
    ws2.column_dimensions[get_column_letter(col_idx)].width = w
ws2.row_dimensions[4].height = 26

chassis_data = [
    ("C1", "Panel P1", "Main Controller Rack (1756-L950TPSXT)", "Main Control Room (Control Cabinet CA1)", 83, 45, 0.3516, 36, 32, 0.4706, 38, 26, 0.4063, 0, 0, 157),
    ("C2", "Panel P2", "Main Expansion High-Density Rack", "Main Control Room (Control Cabinet CA1)", 128, 32, 0.2000, 52, 50, 0.4902, 46, 82, 0.6406, 0, 0, 226),
    ("C3", "Panel P3", "Remote I/O Rack 1 (Spray Dryer / Packing)", "Field RIO Cabinet (Spray Dryer Tower 4F)", 33, 63, 0.6563, 25, 43, 0.6324, 102, 26, 0.2031, 5, 3, 165),
    ("C4", "Panel P4", "Remote I/O Rack 2 (Upper Tower & IS Hub)", "Upper Spray Dryer Tower (6th / 8th Floor)", 50, 46, 0.4792, 1, 33, 0.9706, 48, 16, 0.2500, 0, 0, 99),
    ("C5", "Panel P5", "Remote I/O Station 5 (Slurry Building RIO-200)", "Slurry Building 2nd Floor (RIO Room)", 40, 24, 0.3750, 8, 26, 0.7647, 18, 14, 0.4375, 0, 0, 66)
]

for row_idx, rdata in enumerate(chassis_data, 5):
    ws2.row_dimensions[row_idx].height = 20
    ws2.cell(row_idx, 1, rdata[0]).alignment = align_center
    ws2.cell(row_idx, 2, rdata[1]).alignment = align_center
    ws2.cell(row_idx, 3, rdata[2]).alignment = align_left
    ws2.cell(row_idx, 4, rdata[3]).alignment = align_left
    
    # DI
    ws2.cell(row_idx, 5, rdata[4]).alignment = align_center
    ws2.cell(row_idx, 6, rdata[5]).alignment = align_center
    ws2.cell(row_idx, 7, rdata[6]).alignment = align_center
    ws2.cell(row_idx, 7).number_format = "0.0%"
    
    # DO
    ws2.cell(row_idx, 8, rdata[7]).alignment = align_center
    ws2.cell(row_idx, 9, rdata[8]).alignment = align_center
    ws2.cell(row_idx, 10, rdata[9]).alignment = align_center
    ws2.cell(row_idx, 10).number_format = "0.0%"
    
    # AI
    ws2.cell(row_idx, 11, rdata[10]).alignment = align_center
    ws2.cell(row_idx, 12, rdata[11]).alignment = align_center
    ws2.cell(row_idx, 13, rdata[12]).alignment = align_center
    ws2.cell(row_idx, 13).number_format = "0.0%"
    
    # AO
    ws2.cell(row_idx, 14, rdata[13] if rdata[13] > 0 else "-").alignment = align_center
    ws2.cell(row_idx, 15, rdata[14] if rdata[14] > 0 else "-").alignment = align_center
    
    # Total
    ws2.cell(row_idx, 16, rdata[15]).alignment = align_center
    
    for c in range(1, 17):
        ws2.cell(row_idx, c).font = f_data_bold if c in [1, 2, 16] else f_data
        ws2.cell(row_idx, c).border = border_data
        if row_idx % 2 == 1:
            ws2.cell(row_idx, c).fill = fill_zebra

# Chassis Total Row
r_c_tot = 10
ws2.row_dimensions[r_c_tot].height = 22
ws2.cell(r_c_tot, 1, "TOTAL").alignment = align_center
ws2.cell(r_c_tot, 2, "5 Panels").alignment = align_center
ws2.cell(r_c_tot, 3, "Complete Rockwell Automation Distributed Platform").alignment = align_left
ws2.cell(r_c_tot, 4, "Central Control Room + Remote Field Stations").alignment = align_left
ws2.cell(r_c_tot, 5, 334).alignment = align_center
ws2.cell(r_c_tot, 6, 210).alignment = align_center
ws2.cell(r_c_tot, 7, 0.3860).alignment = align_center
ws2.cell(r_c_tot, 7).number_format = "0.0%"
ws2.cell(r_c_tot, 8, 122).alignment = align_center
ws2.cell(r_c_tot, 9, 184).alignment = align_center
ws2.cell(r_c_tot, 10, 0.6013).alignment = align_center
ws2.cell(r_c_tot, 10).number_format = "0.0%"
ws2.cell(r_c_tot, 11, 252).alignment = align_center
ws2.cell(r_c_tot, 12, 164).alignment = align_center
ws2.cell(r_c_tot, 13, 0.3942).alignment = align_center
ws2.cell(r_c_tot, 13).number_format = "0.0%"
ws2.cell(r_c_tot, 14, 5).alignment = align_center
ws2.cell(r_c_tot, 15, 3).alignment = align_center
ws2.cell(r_c_tot, 16, 713).alignment = align_center
for c in range(1, 17):
    ws2.cell(r_c_tot, c).font = f_total
    ws2.cell(r_c_tot, c).fill = fill_total
    ws2.cell(r_c_tot, c).border = border_total

print("Built 01_Chassis_Architecture_C1_C5.")

# ==============================================================================
# SHEET 3: 02_Field_Enclosure_Reconciliation
# ==============================================================================
ws3 = wb.create_sheet(title="02_Field_Enclosures")
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells("A1:M1")
ws3["A1"] = "FIELD ENCLOSURE & JUNCTION BOX RECONCILIATION (16 FIELD ENCLOSURES)"
ws3["A1"].font = f_title
ws3["A1"].fill = fill_hdr
ws3["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws3.row_dimensions[1].height = 28

ws3.merge_cells("A2:M2")
ws3["A2"] = "Detailed engineering specifications of the 16 Junction Boxes & Control Panels that replace Quotation Line 143 ('JUNCTION BOX: 1 LOT')"
ws3["A2"].font = f_subtitle
ws3["A2"].fill = fill_subhdr
ws3["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws3.row_dimensions[2].height = 18

s3_headers = [
    ("Enclosure Tag", 14),
    ("Classification", 16),
    ("Area Description & Process Location", 32),
    ("Dimensions (HWD mm)", 20),
    ("Material & Ingress Standard", 26),
    ("DI Used", 10),
    ("DO Used", 10),
    ("AI Used", 10),
    ("Active Pts", 12),
    ("Spare Pts", 12),
    ("Spare %", 12),
    ("Connected Control Panel / Rack", 26),
    ("Home-Run Trunk Cable Specification", 38)
]
for col_idx, (th, w) in enumerate(s3_headers, 1):
    c = ws3.cell(4, col_idx, th)
    c.font = f_tbl_hdr
    c.fill = fill_hdr
    c.alignment = align_center
    c.border = border_hdr
    ws3.column_dimensions[get_column_letter(col_idx)].width = w
ws3.row_dimensions[4].height = 26

enclosure_data = [
    ("JB-401", "Standard Process", "Jet Cooker Infeed Area 2nd Floor", "600 x 500 x 200 mm", "IP66 / SS304 (Sanitary Food Grade)", 50, 24, 22, 96, 24, 0.2000, "Panel P1 (Chassis C1)", "1x CVV 24Cx1.5 Control + 1x LiY-CY 16Px0.75 Analog to CA1"),
    ("JB-402", "Standard Process", "Jet Cooker Processing Skid 2F", "450 x 300 x 150 mm", "IP66 / SS304 (Industrial Hygienic)", 24, 9, 14, 48, 35, 0.4217, "Panel P2 (Chassis C2)", "1x CVV 18Cx1.5 Control + 1x LiY-CY 12Px0.75 Analog to CA1"),
    ("JB-601", "Standard Process", "Spray Dryer Ground Floor 1F (Base / Discharge)", "500 x 400 x 200 mm", "IP66 / SS304 (Sanitary Food Grade)", 33, 12, 16, 61, 35, 0.3646, "Panel P1 (Chassis C1)", "1x CVV 18Cx1.5 Control + 1x LiY-CY 12Px0.75 Analog to CA1"),
    ("JB-602", "Standard Process", "Spray Dryer Cyclone & Exhaust 3F", "800 x 600 x 250 mm", "IP66 / SS304 (Sanitary Food Grade)", 23, 3, 58, 86, 27, 0.2389, "Panel P2 (Chassis C2)", "1x CVV 24Cx1.5 Control + 1x LiY-CY 16Px0.75 Analog to CA1"),
    ("JB-606", "Standard Process", "Spray Dryer Filter Cleaning 6F (Burner Area)", "300 x 200 x 150 mm", "IP66 / SS304 (Sanitary Food Grade)", 4, 1, 4, 9, 45, 0.8333, "Panel P3 (Chassis C3)", "1x CVV 18Cx1.5 Control + 1x LiY-CY 4Px0.75 Analog to Panel P3"),
    ("JB-607", "Standard Process", "Spray Dryer Mid-Tower 4F (Main High-Density Hub)", "1000 x 800 x 300 mm", "IP66 / SS304 (Sanitary Food Grade)", 81, 40, 54, 177, 51, 0.2237, "Panel P2 & P3", "2x CVV 24Cx1.5 Control + 2x LiY-CY 16Px0.75 Analog Trunks"),
    ("JB-608", "Standard Process", "Spray Dryer Tower 4F Auxiliary Deck", "300 x 200 x 150 mm", "IP66 / SS304 (Sanitary Food Grade)", 0, 0, 2, 2, 0, 0.0000, "Panel P3 (Chassis C3)", "Local Field Multi-core Trunk to Panel P3"),
    ("JB-612", "Standard Process", "Spray Dryer Upper Tower (6th / 8th Floor)", "600 x 500 x 200 mm", "IP66 / SS304 (Sanitary Food Grade)", 17, 15, 8, 40, 32, 0.4444, "Panel P4 (Chassis C4)", "1x CVV 24Cx1.5 Control + 1x LiY-CY 12Px0.75 Analog to Panel P4"),
    ("JB-618", "Standard Process", "Packing Tower Building & Storage Silos", "600 x 400 x 200 mm", "IP66 / SS304 (Sanitary Food Grade)", 12, 9, 8, 29, 53, 0.6463, "Panel P3 (Chassis C3)", "1x CVV 18Cx1.5 Control + 1x LiY-CY 8Px0.75 Analog to Panel P3"),
    ("IS-JB-603", "ATEX Ex ia (IS)", "Hazardous Area Zone 1/21 (Lower Chamber)", "500 x 400 x 200 mm", "ATEX Ex ia (Blue Terminals / IP66 SS316)", 15, 0, 14, 29, 3, 0.0938, "Panel P4 (IS Hub)", "1x LiY-CY(EB) 16Px0.75 Blue IS Shielded Trunk to IS Barriers"),
    ("IS-JB-608", "ATEX Ex ia (IS)", "Hazardous Area Zone 1/21 (Mid-Tower Powder Zone)", "600 x 500 x 200 mm", "ATEX Ex ia (Blue Terminals / IP66 SS316)", 7, 1, 12, 20, 38, 0.6552, "Panel P4 (IS Hub)", "1x LiY-CY(EB) 16Px0.75 Blue IS Shielded Trunk to IS Barriers"),
    ("IS-JB-612", "ATEX Ex ia (IS)", "Hazardous Area Zone 1/21 (Upper Tower / Vent)", "500 x 400 x 200 mm", "ATEX Ex ia (Blue Terminals / IP66 SS316)", 15, 0, 4, 19, 17, 0.4722, "Panel P4 (IS Hub)", "1x LiY-CY(EB) 12Px0.75 Blue IS Shielded Trunk to IS Barriers"),
    ("IS-JB-618", "ATEX Ex ia (IS)", "Hazardous Area Zone 1/21 (Packing Tower & Silos)", "600 x 400 x 200 mm", "ATEX Ex ia (Blue Terminals / IP66 SS316)", 13, 0, 18, 31, 29, 0.4833, "Panel P4 (IS Hub)", "1x LiY-CY(EB) 16Px0.75 Blue IS Shielded Trunk to IS Barriers"),
    ("CA1", "Main Control Room", "Central Automation Suite (Chassis C1 & C2)", "2200 x 800 x 800 mm", "NEMA 12 / IP54 Rittal TS8 Bayed Suite", 0, 0, 0, 33, 98, 0.7481, "Main Controller Suite", "Pre-wired Marshaling Trunks, Internal Busbars & 1Gbps Fiber DLR"),
    ("RIO-200", "Remote Control Panel", "Slurry Building 2nd Floor (RIO Room)", "1200 x 800 x 300 mm", "IP66 / SS304 (Sanitary Food Grade)", 40, 8, 18, 66, 64, 0.4923, "Panel P5 (Chassis C5)", "1Gbps EtherNet/IP DLR Fiber Ring + Local Field Multi-core Trunks"),
    ("MCC", "Switchgear Interface", "Motor Control Center Switchgear Room", "Centerline 2500 Suite", "Form 4b Arc-Resistant Switchgear", 0, 0, 0, 0, 0, 0.0000, "Motor Feeders & VFDs", "Multi-Core Control Trunks (CVV 7Cx1.5) & EtherNet/IP Backbone")
]

for row_idx, rdata in enumerate(enclosure_data, 5):
    ws3.row_dimensions[row_idx].height = 20
    ws3.cell(row_idx, 1, rdata[0]).alignment = align_center
    ws3.cell(row_idx, 2, rdata[1]).alignment = align_center
    ws3.cell(row_idx, 3, rdata[2]).alignment = align_left
    ws3.cell(row_idx, 4, rdata[3]).alignment = align_center
    ws3.cell(row_idx, 5, rdata[4]).alignment = align_left
    ws3.cell(row_idx, 6, rdata[5]).alignment = align_center
    ws3.cell(row_idx, 7, rdata[6]).alignment = align_center
    ws3.cell(row_idx, 8, rdata[7]).alignment = align_center
    ws3.cell(row_idx, 9, rdata[8]).alignment = align_center
    ws3.cell(row_idx, 10, rdata[9]).alignment = align_center
    ws3.cell(row_idx, 11, rdata[10]).alignment = align_center
    ws3.cell(row_idx, 11).number_format = "0.0%"
    ws3.cell(row_idx, 12, rdata[11]).alignment = align_left
    ws3.cell(row_idx, 13, rdata[12]).alignment = align_left
    
    # Classification pill
    st_c = ws3.cell(row_idx, 2)
    if "Ex ia" in rdata[1]:
        st_c.fill = fill_info
        st_c.font = f_status_info
    elif "Standard" in rdata[1]:
        st_c.fill = fill_pass
        st_c.font = f_status_pass
    else:
        st_c.fill = fill_zebra
        st_c.font = f_data_bold

    for c in range(1, 14):
        if c != 2:
            ws3.cell(row_idx, c).font = f_data_bold if c in [1, 9] else f_data
            if row_idx % 2 == 1:
                ws3.cell(row_idx, c).fill = fill_zebra
        ws3.cell(row_idx, c).border = border_data

# Total Row
r_enc_tot = 21
ws3.row_dimensions[r_enc_tot].height = 22
ws3.cell(r_enc_tot, 1, "TOTAL").alignment = align_center
ws3.cell(r_enc_tot, 2, "16 Enclosures").alignment = align_center
ws3.cell(r_enc_tot, 3, "Complete Plant-Wide Enclosure Infrastructure").alignment = align_left
ws3.cell(r_enc_tot, 4, "-").alignment = align_center
ws3.cell(r_enc_tot, 5, "SS304 Food Grade & SS316 ATEX").alignment = align_left
ws3.cell(r_enc_tot, 6, 334).alignment = align_center
ws3.cell(r_enc_tot, 7, 122).alignment = align_center
ws3.cell(r_enc_tot, 8, 252).alignment = align_center
ws3.cell(r_enc_tot, 9, 713).alignment = align_center
ws3.cell(r_enc_tot, 10, 561).alignment = align_center
ws3.cell(r_enc_tot, 11, 0.4403).alignment = align_center
ws3.cell(r_enc_tot, 11).number_format = "0.0%"
ws3.cell(r_enc_tot, 12, "5 Interconnected Chassis").alignment = align_left
ws3.cell(r_enc_tot, 13, "Standardized Trunk Schedule & DLR Fiber Ring").alignment = align_left
for c in range(1, 14):
    ws3.cell(r_enc_tot, c).font = f_total
    ws3.cell(r_enc_tot, c).fill = fill_total
    ws3.cell(r_enc_tot, c).border = border_total

print("Built 02_Field_Enclosures.")

# ==============================================================================
# SHEET 4: 03_Cabling_and_Network_Delta
# ==============================================================================
ws4 = wb.create_sheet(title="03_Cabling_and_Network_Delta")
ws4.views.sheetView[0].showGridLines = True

ws4.merge_cells("A1:H1")
ws4["A1"] = "CABLING, NETWORK & MARSHALING INFRASTRUCTURE DELTA"
ws4["A1"].font = f_title
ws4["A1"].fill = fill_hdr
ws4["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws4.row_dimensions[1].height = 28

ws4.merge_cells("A2:H2")
ws4["A2"] = "Technical and commercial reconciliation of Cable Schedule (BOQ D3) & Cable Tray (BOQ D4) vs. Distributed Field Engineering"
ws4["A2"].font = f_subtitle
ws4["A2"].fill = fill_subhdr
ws4["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
ws4.row_dimensions[2].height = 18

s4_headers = [
    ("Infrastructure Item", 24),
    ("Quotation BOQ Specification", 28),
    ("Quoted Qty", 14),
    ("JB Schedule Implementation", 32),
    ("Engineered Qty / Spec", 22),
    ("Engineering Variance & Rationale", 34),
    ("Commercial Impact", 24),
    ("Variation Order (VO) Status", 22)
]
for col_idx, (th, w) in enumerate(s4_headers, 1):
    c = ws4.cell(4, col_idx, th)
    c.font = f_tbl_hdr
    c.fill = fill_hdr
    c.alignment = align_center
    c.border = border_hdr
    ws4.column_dimensions[get_column_letter(col_idx)].width = w
ws4.row_dimensions[4].height = 26

cabling_data = [
    ("Digital Control Multicore Trunks", "CVV 32C x 1 mm² (Central Home-Run)", "1,980 m", "Standardized CVV 18C & 24C x 1.5 mm² to local RIOs", "Localized Modular Trunks", "Eliminates heavy 32C copper pulling; improves signal voltage drop with 1.5 mm²", "Cost Offset (Shorter run)", "Scope Substitution"),
    ("Analog Multi-Pair Trunks", "LiY-CY TP 20P x 1 mm² (Overall Shielded)", "1,980 m", "Standardized LiY-CY 12P & 16P x 0.75 mm² to local RIOs", "Localized Shielded Pairs", "Reduces conduit fill factor; terminates directly into high-density IF16 cards", "Cost Offset (Shorter run)", "Scope Substitution"),
    ("Hazardous Area IS Trunks", "Generic LiY-CY instrument pairs", "Included in D3", "Specialized LiY-CY(EB) Blue-Sheathed IS Trunk Cables", "Dedicated 12P & 16P Trunks", "Mandatory compliance for ATEX Zone 1/21 explosion-proof intrinsically safe circuits", "Premium Material", "Compliant Upgrade"),
    ("Plant Network Backbone", "CAT6 Outdoor (465m) + Fiber (1 LOT)", "465 m / 1 LOT", "1 Gbps Multimode Fiber Optic Device Level Ring (DLR)", "Fiber DLR Ring + Copper Drops", "Provides fault-tolerant 1 Gbps ring redundancy between CA1, RIO-1, RIO-2 & RIO-200", "In Base Budget", "Approved Design"),
    ("Field Junction Boxes", "Line 143: 'JUNCTION BOX: 1 LOT'", "1 LOT", "16 Specific Enclosures (8 SS304 Food Grade + 4 SS316 ATEX)", "16 Physical Enclosures", "Upgrades generic lump sum into certified SS304/SS316 sanitary and explosion-proof boxes", "High Addition", "Primary VO Item"),
    ("Remote Control Enclosures", "2x TS Enclosure 2000x800x800mm in CA1", "2 Ea (CA1)", "CA1 Bayed Suite + RIO-200 (1200x800x300mm SS304) + Tower RIOs", "3 Cabinets + 2 Racks", "Provides local housing for Chassis C3, C4 and C5 in field areas", "Additional Hardware", "VO Claim: RIO Enclosures")
]

for row_idx, rdata in enumerate(cabling_data, 5):
    ws4.row_dimensions[row_idx].height = 24
    ws4.cell(row_idx, 1, rdata[0]).alignment = align_left
    ws4.cell(row_idx, 2, rdata[1]).alignment = align_left
    ws4.cell(row_idx, 3, rdata[2]).alignment = align_center
    ws4.cell(row_idx, 4, rdata[3]).alignment = align_left
    ws4.cell(row_idx, 5, rdata[4]).alignment = align_center
    ws4.cell(row_idx, 6, rdata[5]).alignment = align_wrap_left
    ws4.cell(row_idx, 7, rdata[6]).alignment = align_center
    
    st_cell = ws4.cell(row_idx, 8, rdata[7])
    st_cell.alignment = align_center
    if "VO" in rdata[7] or "Primary" in rdata[7]:
        st_cell.fill = fill_warn
        st_cell.font = f_status_warn
    elif "Approved" in rdata[7] or "Compliant" in rdata[7]:
        st_cell.fill = fill_pass
        st_cell.font = f_status_pass
    else:
        st_cell.fill = fill_info
        st_cell.font = f_status_info

    for c in range(1, 9):
        if c != 8:
            ws4.cell(row_idx, c).font = f_data_bold if c in [1, 3, 5] else f_data
            if row_idx % 2 == 1:
                ws4.cell(row_idx, c).fill = fill_zebra
        ws4.cell(row_idx, c).border = border_data

print("Built 03_Cabling_and_Network_Delta.")

# Save Workbook
output_path = '/Users/x92120/xApp-001/x260911-002-KalasinEngineering/IO_VO/IO_Comparison_Quotation_vs_Junction_Box.xlsx'
wb.save(output_path)
print(f"Successfully generated and saved: {output_path}")
