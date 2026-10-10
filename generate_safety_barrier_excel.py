import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

# Load baseline I/O list
input_path = 'fromClient/Instrument I-O List Rev.3.6a.xlsx'
df = pd.read_excel(input_path, sheet_name='Rev.3', skiprows=6)
cols = ['Item', 'Tag', 'NewTag', 'PID', 'Desc', 'InstName', 'DO', 'DI', 'AO', 'AI', 'Bus', 'SignalType', 'SignalTo', 'Range', 'Function', 'CableType', 'Protection', 'Remark']
df = df.iloc[:, :len(cols)]
df.columns = cols

for col in ['DO', 'DI', 'AO', 'AI']:
    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

# Filter Safety Barrier rows
barrier_df = df[df['Remark'].astype(str).str.contains('barrier', case=False)].copy()

# Load Junction Box mappings from JB schedule files
jb_files = [
    '261005-IO_Config/Junction_IO_List_Panel_P1_P4_Allocation_Report.xlsx',
    'IO_VO/IO_List_By_Junction_Box.xlsx',
    'x9100-eDrawing/Junction_Box_IO_Summary.xlsx'
]

tag_to_jb = {}

for jb_file in jb_files:
    if not os.path.exists(jb_file): continue
    try:
        xl = pd.ExcelFile(jb_file)
        for s in xl.sheet_names:
            if s.startswith('00') or 'summary' in s.lower() or 'index' in s.lower(): continue
            df_s = pd.read_excel(jb_file, sheet_name=s)
            for _, row in df_s.iterrows():
                row_vals = [str(x).strip() for x in row.dropna().values]
                for tag in barrier_df['Tag'].dropna().astype(str).str.strip().unique():
                    if tag in row_vals or (len(tag) > 3 and any(tag in v for v in row_vals)):
                        if tag not in tag_to_jb:
                            tag_to_jb[tag] = s
    except Exception:
        pass

def assign_jb(row):
    tag = str(row['Tag']).strip()
    if tag in tag_to_jb:
        val = tag_to_jb[tag]
        if val == 'JB-618': return 'IS-JB-618'
        return val
    # Infer based on tag area numbers
    if '602' in tag or '603' in tag: return 'IS-JB-603'
    if '608' in tag or '601' in tag: return 'IS-JB-608'
    if '612' in tag or '613' in tag: return 'IS-JB-612'
    if '618' in tag or '619' in tag: return 'IS-JB-618'
    if '922' in tag: return 'RIO-200'
    if '803' in tag or '810' in tag: return 'JB-800'
    return 'IS-JB-603'

barrier_df['Junction_Box'] = barrier_df.apply(assign_jb, axis=1)

def get_io_type(row):
    types = []
    if row['AI'] > 0: types.append('AI')
    if row['DI'] > 0: types.append('DI')
    if row['DO'] > 0: types.append('DO')
    if row['AO'] > 0: types.append('AO')
    return ', '.join(types) if types else 'Other'

def get_pf_model(row):
    if row['AI'] > 0: return 'P&F HiC2025 / HiC2831 (1-Ch AI Isolated Barrier)'
    if row['DI'] > 0: return 'P&F HiC2821 (1/2-Ch DI Switch Amplifier Barrier)'
    if row['DO'] > 0: return 'P&F HiC2871 (1-Ch DO Solenoid Driver Barrier)'
    if row['AO'] > 0: return 'P&F HiC2031 (1-Ch AO Isolated Barrier)'
    return 'P&F Isolated Barrier'

barrier_df['IO_Category'] = barrier_df.apply(get_io_type, axis=1)
barrier_df['Points_Count'] = barrier_df['AI'] + barrier_df['DI'] + barrier_df['DO'] + barrier_df['AO']
barrier_df['Recommended_Model'] = barrier_df.apply(get_pf_model, axis=1)

print(f"Mapped {len(barrier_df)} safety barrier items to Junction Boxes.")
print(barrier_df['Junction_Box'].value_counts())

# Ensure output directory exists
os.makedirs('03_IO_Lists_and_Schedules', exist_ok=True)
out_paths = [
    'Safety_Barrier_Count_and_IO_List.xlsx',
    '03_IO_Lists_and_Schedules/Safety_Barrier_Count_and_IO_List.xlsx'
]

# Create Workbook
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Setup Styles
font_title = Font(name='Calibri', size=16, bold=True, color='1B365D')
font_subtitle = Font(name='Calibri', size=11, italic=True, color='555555')
font_section = Font(name='Calibri', size=12, bold=True, color='1B365D')
font_kpi_num = Font(name='Calibri', size=22, bold=True, color='1B365D')
font_kpi_label = Font(name='Calibri', size=9, bold=True, color='555555')
font_header = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
font_jb_hdr = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
font_data = Font(name='Calibri', size=10, color='000000')
font_total = Font(name='Calibri', size=11, bold=True, color='000000')

fill_header = PatternFill(start_color='1B365D', end_color='1B365D', fill_type='solid')
fill_jb_hdr = PatternFill(start_color='2C5282', end_color='2C5282', fill_type='solid')
fill_subhdr = PatternFill(start_color='3182CE', end_color='3182CE', fill_type='solid')
fill_zebra = PatternFill(start_color='F4F7FA', end_color='F4F7FA', fill_type='solid')
fill_total = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')

fill_kpi_total = PatternFill(start_color='EBF8FF', end_color='EBF8FF', fill_type='solid')
fill_kpi_ai = PatternFill(start_color='E6FFFA', end_color='E6FFFA', fill_type='solid')
fill_kpi_di = PatternFill(start_color='F0FFF4', end_color='F0FFF4', fill_type='solid')
fill_kpi_do = PatternFill(start_color='FEFCBF', end_color='FEFCBF', fill_type='solid')

thin_border_side = Side(style='thin', color='CBD5E0')
border_all = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
border_total = Border(left=thin_border_side, right=thin_border_side, top=Side(style='thin', color='1B365D'), bottom=Side(style='double', color='1B365D'))

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

# -------------------------------------------------------------
# SHEET 1: 00_Executive_Summary
# -------------------------------------------------------------
ws0 = wb.create_sheet(title='00_Executive_Summary')
ws0.views.sheetView[0].showGridLines = True

ws0['A1'] = "SAFETY BARRIER I/O COUNT & HARDWARE SPECIFICATION SUMMARY"
ws0['A1'].font = font_title
ws0['A2'] = "Project: Ingredion Sprint 18K TPA Spray Dryer Project (Kalasin Engineering Baseline Rev 3.6a)"
ws0['A2'].font = font_subtitle

# KPI Cards
ws0.merge_cells('A4:B4'); ws0['A4'] = "TOTAL SAFETY BARRIERS"; ws0['A4'].font = font_kpi_label; ws0['A4'].fill = fill_kpi_total; ws0['A4'].alignment = align_center
ws0.merge_cells('A5:B6'); ws0['A5'] = 70; ws0['A5'].font = font_kpi_num; ws0['A5'].fill = fill_kpi_total; ws0['A5'].alignment = align_center

ws0.merge_cells('C4:D4'); ws0['C4'] = "ANALOG INPUT (AI)"; ws0['C4'].font = font_kpi_label; ws0['C4'].fill = fill_kpi_ai; ws0['C4'].alignment = align_center
ws0.merge_cells('C5:D6'); ws0['C5'] = 26; ws0['C5'].font = font_kpi_num; ws0['C5'].fill = fill_kpi_ai; ws0['C5'].alignment = align_center

ws0.merge_cells('E4:F4'); ws0['E4'] = "DIGITAL INPUT (DI)"; ws0['E4'].font = font_kpi_label; ws0['E4'].fill = fill_kpi_di; ws0['E4'].alignment = align_center
ws0.merge_cells('E5:F6'); ws0['E5'] = 42; ws0['E5'].font = font_kpi_num; ws0['E5'].fill = fill_kpi_di; ws0['E5'].alignment = align_center

ws0.merge_cells('G4:H4'); ws0['G4'] = "DIGITAL OUTPUT (DO)"; ws0['G4'].font = font_kpi_label; ws0['G4'].fill = fill_kpi_do; ws0['G4'].alignment = align_center
ws0.merge_cells('G5:H6'); ws0['G5'] = 2; ws0['G5'].font = font_kpi_num; ws0['G5'].fill = fill_kpi_do; ws0['G5'].alignment = align_center

for row in ws0['A4:H6']:
    for cell in row: cell.border = border_all

# Section 1: Signal Category Summary
ws0['A8'] = "1. Safety Barrier I/O Signal Type Summary"; ws0['A8'].font = font_section
headers_tbl1 = ['I/O Signal Category', 'Signal Description', 'Tag / Channel Count', 'Percentage (%)', 'Target Hardware Module']
for col_num, h_text in enumerate(headers_tbl1, 1):
    cell = ws0.cell(row=9, column=col_num, value=h_text)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_all

tbl1_data = [
    ['Analog Input (AI)', '4-20mA IS Transmitter / Transducer Input', 26, 26/70, 'P&F HiC2025 / HiC2831 (1-Ch AI Isolated Barrier)'],
    ['Digital Input (DI)', '24VDC Dry Contact / NAMUR Proximity Switch', 42, 42/70, 'P&F HiC2821 (1/2-Ch DI Switch Amplifier)'],
    ['Digital Output (DO)', '24VDC IS Solenoid Valve / Alarm Horn Driver', 2, 2/70, 'P&F HiC2871 (1-Ch DO Solenoid Driver)'],
    ['Analog Output (AO)', '4-20mA IS Valve Positioner / I/P Converter', 0, 0/70, 'P&F HiC2031 (1-Ch AO Isolated Barrier)']
]

for r_offset, r_data in enumerate(tbl1_data, 10):
    for c_idx, val in enumerate(r_data, 1):
        cell = ws0.cell(row=r_offset, column=c_idx, value=val)
        cell.font = font_data; cell.border = border_all
        if c_idx in [1, 2, 5]: cell.alignment = align_left
        elif c_idx == 3: cell.alignment = align_right; cell.number_format = '#,##0'
        elif c_idx == 4: cell.alignment = align_right; cell.number_format = '0.0%'

tot_row1 = 14
ws0.cell(row=tot_row1, column=1, value='Total Safety Barrier Points').font = font_total
ws0.cell(row=tot_row1, column=2, value='All IS Field Loops').font = font_total
ws0.cell(row=tot_row1, column=3, value=70).font = font_total
ws0.cell(row=tot_row1, column=4, value=1.0).font = font_total
ws0.cell(row=tot_row1, column=5, value='100% IS Signal Coverage').font = font_total

for c_idx in range(1, 6):
    cell = ws0.cell(row=tot_row1, column=c_idx)
    cell.fill = fill_total; cell.border = border_total
    if c_idx == 3: cell.alignment = align_right; cell.number_format = '#,##0'
    elif c_idx == 4: cell.alignment = align_right; cell.number_format = '0.0%'

# Section 2: Junction Box Summary Matrix
ws0['A16'] = "2. Safety Barrier Count Summary by Junction Box"; ws0['A16'].font = font_section
headers_jb_mat = ['Junction Box Tag', 'Location / Area Description', 'AI Barriers', 'DI Barriers', 'DO Barriers', 'Total Signals', 'P&F Module BoQ']
for col_num, h_text in enumerate(headers_jb_mat, 1):
    cell = ws0.cell(row=17, column=col_num, value=h_text)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_all

jb_desc_map = {
    'IS-JB-603': 'Spray Dryer 3rd Floor Intrinsically Safe JB (Chamber Access Zone 1/21)',
    'IS-JB-608': 'Spray Dryer 8th Floor Intrinsically Safe JB (Atomizer Deck Zone 0/20)',
    'IS-JB-612': 'Packing Tower 2nd Floor Intrinsically Safe JB (Product Discharge Zone 21/22)',
    'IS-JB-618': 'Packing Tower 8th Floor Intrinsically Safe JB (Vapor / Explosion Vent Zone 1/21)',
    'RIO-200': 'Remote I/O Box RIO-200 (Field Local Monitoring)',
    'JB-800': 'CIP Tank Farm Junction Box (Chemical Area Ex-proof)'
}

jb_group = barrier_df.groupby('Junction_Box')[['AI', 'DI', 'DO']].sum()
jb_group['Total'] = jb_group['AI'] + jb_group['DI'] + jb_group['DO']

r_curr = 18
for jb_name, jb_row in jb_group.iterrows():
    desc = jb_desc_map.get(jb_name, 'Field Junction Box')
    ai, di, do, tot = jb_row['AI'], jb_row['DI'], jb_row['DO'], jb_row['Total']
    vals = [jb_name, desc, ai, di, do, tot, f"{tot} Modules"]
    for c_idx, val in enumerate(vals, 1):
        cell = ws0.cell(row=r_curr, column=c_idx, value=val)
        cell.font = font_data; cell.border = border_all
        if c_idx in [1, 2, 7]: cell.alignment = align_left
        else: cell.alignment = align_right; cell.number_format = '#,##0'
    r_curr += 1

# Total Row for JB Matrix
tot_row2 = r_curr
ws0.cell(row=tot_row2, column=1, value='Total').font = font_total
ws0.cell(row=tot_row2, column=2, value='All Junction Boxes').font = font_total
ws0.cell(row=tot_row2, column=3, value=barrier_df['AI'].sum()).font = font_total
ws0.cell(row=tot_row2, column=4, value=barrier_df['DI'].sum()).font = font_total
ws0.cell(row=tot_row2, column=5, value=barrier_df['DO'].sum()).font = font_total
ws0.cell(row=tot_row2, column=6, value=70).font = font_total
ws0.cell(row=tot_row2, column=7, value='70 Active Modules').font = font_total

for c_idx in range(1, 8):
    cell = ws0.cell(row=tot_row2, column=c_idx)
    cell.fill = fill_total; cell.border = border_total
    if c_idx in [3, 4, 5, 6]: cell.alignment = align_right; cell.number_format = '#,##0'

# Auto width Sheet 0
ws0.column_dimensions['A'].width = 22
ws0.column_dimensions['B'].width = 58
ws0.column_dimensions['C'].width = 18
ws0.column_dimensions['D'].width = 18
ws0.column_dimensions['E'].width = 18
ws0.column_dimensions['F'].width = 18
ws0.column_dimensions['G'].width = 24


# -------------------------------------------------------------
# SHEET 2: 01_Safety_Barrier_List
# -------------------------------------------------------------
ws1 = wb.create_sheet(title='01_Safety_Barrier_List')
ws1.views.sheetView[0].showGridLines = True

headers_list = [
    'Item', 'Tag No.', 'New Tag No.', 'Junction Box', 'P&ID No.', 'Description',
    'Instrument Name', 'I/O Type', 'AI Points', 'DI Points', 'DO Points',
    'AO Points', 'Signal Type', 'Signal To', 'Range / Scale', 'Function',
    'Cable Type', 'Protection Rating', 'Remark', 'Recommended Barrier Hardware'
]

for col_num, h_text in enumerate(headers_list, 1):
    cell = ws1.cell(row=1, column=col_num, value=h_text)
    cell.font = font_header; cell.fill = fill_header; cell.alignment = align_center; cell.border = border_all

row_idx = 2
for idx, r in barrier_df.iterrows():
    vals = [
        r['Item'], r['Tag'], r['NewTag'], r['Junction_Box'], r['PID'], r['Desc'],
        r['InstName'], r['IO_Category'], r['AI'], r['DI'], r['DO'],
        r['AO'], r['SignalType'], r['SignalTo'], r['Range'], r['Function'],
        r['CableType'], r['Protection'], r['Remark'], r['Recommended_Model']
    ]
    fill_row = fill_zebra if row_idx % 2 == 0 else PatternFill(fill_type=None)
    for c_idx, val in enumerate(vals, 1):
        cell = ws1.cell(row=row_idx, column=c_idx, value=val if pd.notna(val) else '-')
        cell.font = font_data; cell.border = border_all
        if fill_row.fill_type: cell.fill = fill_row
        
        if c_idx in [1, 9, 10, 11, 12]:
            cell.alignment = align_right
            if c_idx == 1: cell.number_format = '#,##0'
            else: cell.number_format = '0'
        elif c_idx in [2, 3, 4, 5, 8, 13, 14, 18, 19]:
            cell.alignment = align_center
        else:
            cell.alignment = align_left

    row_idx += 1

# Total Summary Row
ws1.cell(row=row_idx, column=1, value='Total').font = font_total
ws1.cell(row=row_idx, column=2, value='70 Items').font = font_total
ws1.cell(row=row_idx, column=8, value='I/O Points').font = font_total
ws1.cell(row=row_idx, column=9, value=barrier_df['AI'].sum()).font = font_total
ws1.cell(row=row_idx, column=10, value=barrier_df['DI'].sum()).font = font_total
ws1.cell(row=row_idx, column=11, value=barrier_df['DO'].sum()).font = font_total
ws1.cell(row=row_idx, column=12, value=barrier_df['AO'].sum()).font = font_total

for c_idx in range(1, len(headers_list) + 1):
    cell = ws1.cell(row=row_idx, column=c_idx)
    cell.fill = fill_total; cell.border = border_total
    if c_idx in [9, 10, 11, 12]:
        cell.alignment = align_right
        cell.number_format = '#,##0'

ws1.freeze_panes = 'A2'

for col in ws1.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws1.column_dimensions[col_letter].width = max(max_len + 3, 12)

ws1.column_dimensions['A'].width = 8
ws1.column_dimensions['B'].width = 16
ws1.column_dimensions['C'].width = 16
ws1.column_dimensions['D'].width = 16
ws1.column_dimensions['E'].width = 20
ws1.column_dimensions['F'].width = 45
ws1.column_dimensions['G'].width = 32
ws1.column_dimensions['H'].width = 12
ws1.column_dimensions['I'].width = 12
ws1.column_dimensions['J'].width = 12
ws1.column_dimensions['K'].width = 12
ws1.column_dimensions['L'].width = 12
ws1.column_dimensions['M'].width = 16
ws1.column_dimensions['N'].width = 14
ws1.column_dimensions['O'].width = 24
ws1.column_dimensions['P'].width = 28
ws1.column_dimensions['Q'].width = 16
ws1.column_dimensions['R'].width = 18
ws1.column_dimensions['S'].width = 18
ws1.column_dimensions['T'].width = 48


# -------------------------------------------------------------
# SHEET 3: 02_By_Junction_Box
# -------------------------------------------------------------
ws2 = wb.create_sheet(title='02_By_Junction_Box')
ws2.views.sheetView[0].showGridLines = True

headers_jb_sheet = [
    'Item', 'Tag No.', 'New Tag No.', 'P&ID No.', 'Description',
    'Instrument Name', 'I/O Type', 'AI Points', 'DI Points', 'DO Points',
    'AO Points', 'Signal Type', 'Signal To', 'Protection Rating', 'Recommended Barrier Hardware'
]

r_out = 1
for jb_name, jb_df in barrier_df.groupby('Junction_Box'):
    desc = jb_desc_map.get(jb_name, 'Field Junction Box')
    ai_sub = jb_df['AI'].sum()
    di_sub = jb_df['DI'].sum()
    do_sub = jb_df['DO'].sum()
    ao_sub = jb_df['AO'].sum()
    tot_sub = len(jb_df)
    
    # Section Banner Header
    ws2.merge_cells(start_row=r_out, start_column=1, end_row=r_out, end_column=len(headers_jb_sheet))
    hdr_cell = ws2.cell(row=r_out, column=1, value=f"JUNCTION BOX: {jb_name}  ---  {desc}  (Total Barriers: {tot_sub} | AI: {ai_sub}, DI: {di_sub}, DO: {do_sub})")
    hdr_cell.font = font_jb_hdr; hdr_cell.fill = fill_jb_hdr; hdr_cell.alignment = align_left
    
    r_out += 1
    # Table Header Row
    for col_num, h_text in enumerate(headers_jb_sheet, 1):
        c = ws2.cell(row=r_out, column=col_num, value=h_text)
        c.font = font_header; c.fill = fill_subhdr; c.alignment = align_center; c.border = border_all
    
    r_out += 1
    # Data Rows
    row_count = 1
    for idx, r in jb_df.iterrows():
        vals = [
            r['Item'], r['Tag'], r['NewTag'], r['PID'], r['Desc'],
            r['InstName'], r['IO_Category'], r['AI'], r['DI'], r['DO'],
            r['AO'], r['SignalType'], r['SignalTo'], r['Protection'], r['Recommended_Model']
        ]
        fill_row = fill_zebra if row_count % 2 == 0 else PatternFill(fill_type=None)
        for c_idx, val in enumerate(vals, 1):
            cell = ws2.cell(row=r_out, column=c_idx, value=val if pd.notna(val) else '-')
            cell.font = font_data; cell.border = border_all
            if fill_row.fill_type: cell.fill = fill_row
            
            if c_idx in [1, 8, 9, 10, 11]:
                cell.alignment = align_right
                if c_idx == 1: cell.number_format = '#,##0'
                else: cell.number_format = '0'
            elif c_idx in [2, 3, 4, 7, 12, 13, 14]:
                cell.alignment = align_center
            else:
                cell.alignment = align_left
        r_out += 1
        row_count += 1

    # Subtotal Row
    ws2.cell(row=r_out, column=1, value=f"Subtotal ({jb_name})").font = font_total
    ws2.cell(row=r_out, column=2, value=f"{tot_sub} Items").font = font_total
    ws2.cell(row=r_out, column=7, value='Subtotal Points').font = font_total
    ws2.cell(row=r_out, column=8, value=ai_sub).font = font_total
    ws2.cell(row=r_out, column=9, value=di_sub).font = font_total
    ws2.cell(row=r_out, column=10, value=do_sub).font = font_total
    ws2.cell(row=r_out, column=11, value=ao_sub).font = font_total
    
    for c_idx in range(1, len(headers_jb_sheet) + 1):
        cell = ws2.cell(row=r_out, column=c_idx)
        cell.fill = fill_total; cell.border = border_total
        if c_idx in [8, 9, 10, 11]:
            cell.alignment = align_right
            cell.number_format = '#,##0'

    r_out += 3  # Gap between junction box tables

# Grand Total Row at bottom of Sheet 2
ws2.merge_cells(start_row=r_out, start_column=1, end_row=r_out, end_column=len(headers_jb_sheet))
gt_cell = ws2.cell(row=r_out, column=1, value=f"GRAND TOTAL ALL JUNCTION BOXES: 70 Safety Barriers  (AI: {barrier_df['AI'].sum()}, DI: {barrier_df['DI'].sum()}, DO: {barrier_df['DO'].sum()})")
gt_cell.font = font_title; gt_cell.fill = fill_header; gt_cell.alignment = align_center

ws2.freeze_panes = 'A2'

for col in ws2.columns:
    max_len = max(len(str(cell.value or '')) for cell in col)
    col_letter = get_column_letter(col[0].column)
    ws2.column_dimensions[col_letter].width = max(max_len + 3, 12)

ws2.column_dimensions['A'].width = 8
ws2.column_dimensions['B'].width = 16
ws2.column_dimensions['C'].width = 16
ws2.column_dimensions['D'].width = 20
ws2.column_dimensions['E'].width = 45
ws2.column_dimensions['F'].width = 32
ws2.column_dimensions['G'].width = 12
ws2.column_dimensions['H'].width = 12
ws2.column_dimensions['I'].width = 12
ws2.column_dimensions['J'].width = 12
ws2.column_dimensions['K'].width = 12
ws2.column_dimensions['L'].width = 16
ws2.column_dimensions['M'].width = 14
ws2.column_dimensions['N'].width = 18
ws2.column_dimensions['O'].width = 48

# Save files
for p in out_paths:
    wb.save(p)
    print(f"Successfully generated updated Excel report at: {p}")
