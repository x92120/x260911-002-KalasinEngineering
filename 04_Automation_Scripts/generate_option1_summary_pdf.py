#!/usr/bin/env python3
"""
generate_option1_summary_pdf.py
===============================
Generates an Executive Engineering PDF Summary for Option 1: Direct Slot Correlation (TB-CxSx)
using Typst.
Outputs:
  - 03_IO_Lists_and_Schedules/PLC_Terminal_Summary_Option1_C1_C2.pdf
  - 02_Electrical_and_eDrawing/02_ePlan_Exports/PLC_Terminal_Summary_Option1_C1_C2.pdf
"""

import os
import sys
import shutil
import subprocess
import openpyxl

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXCEL_PATH = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules", "Tags list PLC-Option1-DirectSlot-C1-C2.xlsx")
TYPST_SRC_PATH = os.path.join(BASE_DIR, "04_Automation_Scripts", "option1_summary.typ")
PDF_OUT_1 = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules", "PLC_Terminal_Summary_Option1_C1_C2.pdf")
PDF_OUT_2 = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing", "02_ePlan_Exports", "PLC_Terminal_Summary_Option1_C1_C2.pdf")

def load_summary_data():
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    ws = wb['Summary_By_Slot']
    rows = []
    for r in range(5, ws.max_row + 1):
        c_val = ws.cell(r, 2).value
        s_val = ws.cell(r, 3).value
        cx_sx = ws.cell(r, 4).value
        io_t = ws.cell(r, 5).value
        card = ws.cell(r, 6).value
        total_pts = ws.cell(r, 7).value
        active_pts = ws.cell(r, 8).value
        spare_pts = ws.cell(r, 9).value
        tb_strip = ws.cell(r, 10).value
        tb_range = ws.cell(r, 11).value
        if c_val:
            rows.append({
                'chassis': str(c_val),
                'slot_str': str(s_val),
                'cx_sx': str(cx_sx),
                'io_type': str(io_t),
                'card': str(card),
                'total': int(total_pts),
                'active': int(active_pts),
                'spare': int(spare_pts),
                'tb_strip': str(tb_strip),
                'tb_range': str(tb_range)
            })
    return rows

def generate_typst_content(rows):
    c1_rows = [r for r in rows if r['chassis'] == 'C1']
    c2_rows = [r for r in rows if r['chassis'] == 'C2']

    c1_total = sum(r['total'] for r in c1_rows)
    c1_active = sum(r['active'] for r in c1_rows)
    c1_spare = sum(r['spare'] for r in c1_rows)

    c2_total = sum(r['total'] for r in c2_rows)
    c2_active = sum(r['active'] for r in c2_rows)
    c2_spare = sum(r['spare'] for r in c2_rows)

    grand_total = c1_total + c2_total
    grand_active = c1_active + c2_active
    grand_spare = c1_spare + c2_spare

    # Build C1 Table Rows
    c1_table_lines = []
    for i, r in enumerate(c1_rows, 1):
        io_color = 'rgb("15803d")' if r['io_type'] == 'DI' else ('rgb("b45309")' if r['io_type'] == 'DO' else 'rgb("0369a1")')
        sample_tag = f"{r['cx_sx']}:1/{r['tb_strip']}:1A" if r['io_type'] in ['DI', 'DO'] else f"{r['cx_sx']}:1/{r['tb_strip']}:1(+)"
        c1_table_lines.append(f"""  text(size: 8pt)[{i}],
  text(size: 8pt, weight: "medium")[{r['slot_str']}],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[{r['cx_sx']}],
  text(size: 8pt, weight: "bold", fill: {io_color})[{r['io_type']}],
  text(size: 8pt)[{r['card']}],
  text(size: 8pt)[{r['total']}],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[{r['active']}],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[{r['spare']}],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[{r['tb_strip']}],
  text(size: 8pt)[{r['tb_range']}],
  text(size: 7.5pt, font: "Menlo")[{sample_tag}],""")

    # Build C2 Table Rows
    c2_table_lines = []
    for i, r in enumerate(c2_rows, 1):
        io_color = 'rgb("15803d")' if r['io_type'] == 'DI' else ('rgb("b45309")' if r['io_type'] == 'DO' else 'rgb("0369a1")')
        sample_tag = f"{r['cx_sx']}:1/{r['tb_strip']}:1A" if r['io_type'] in ['DI', 'DO'] else f"{r['cx_sx']}:1/{r['tb_strip']}:1(+)"
        c2_table_lines.append(f"""  text(size: 8pt)[{i}],
  text(size: 8pt, weight: "medium")[{r['slot_str']}],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[{r['cx_sx']}],
  text(size: 8pt, weight: "bold", fill: {io_color})[{r['io_type']}],
  text(size: 8pt)[{r['card']}],
  text(size: 8pt)[{r['total']}],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[{r['active']}],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[{r['spare']}],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[{r['tb_strip']}],
  text(size: 8pt)[{r['tb_range']}],
  text(size: 7.5pt, font: "Menlo")[{sample_tag}],""")

    c1_rows_str = "\n".join(c1_table_lines)
    c2_rows_str = "\n".join(c2_table_lines)

    typst_code = f"""
#set page(
  paper: "a4",
  flipped: true,
  margin: (top: 1.4cm, bottom: 1.4cm, left: 1.6cm, right: 1.6cm),
  header: context {{
    grid(
      columns: (1fr, 1fr),
      align: (left, right),
      text(size: 8pt, fill: rgb("64748b"), font: ("Arial", "Thonburi"))[
        *KALASIN ENGINEERING* | ControlLogix 1756 PLC System
      ],
      text(size: 8pt, fill: rgb("64748b"), font: ("Arial", "Thonburi"))[
        *Option 1 Summary (TB-CxSx)* | Page #counter(page).display()
      ]
    )
    line(length: 100%, stroke: 0.5pt + rgb("cbd5e1"))
  }},
  footer: context {{
    line(length: 100%, stroke: 0.5pt + rgb("cbd5e1"))
    v(2pt)
    grid(
      columns: (1fr, 1fr),
      align: (left, right),
      text(size: 8pt, fill: rgb("94a3b8"), font: ("Arial", "Thonburi"))[
        Document Ref: ENG-PLC-TB-OPT1-R01 | Direct Slot Correlation Standard
      ],
      text(size: 8pt, fill: rgb("94a3b8"), font: ("Arial", "Thonburi"))[
        Deliverable File: Tags list PLC-Option1-DirectSlot-C1-C2.xlsx
      ]
    )
  }}
)

#set text(font: ("Arial", "Thonburi"), size: 9pt, fill: rgb("0f172a"))
#set par(justify: true, leading: 0.55em)

// --- TITLE BLOCK ---
#block(
  fill: rgb("1e3a8a"),
  inset: (x: 16pt, y: 11pt),
  radius: 4pt,
  width: 100%
)[
  #grid(
    columns: (2.5fr, 1fr),
    align: (left + horizon, right + horizon),
    [
      #text(fill: white, size: 14pt, weight: "bold")[
        PLC TERMINAL BLOCK TAGGING & WIRE MARKING SUMMARY
      ]
      #v(2pt)
      #text(fill: rgb("93c5fd"), size: 9pt, weight: "medium")[
        OPTION 1: DIRECT SLOT CORRELATION (`TB-CxSx`) --- CHASSIS C1 & C2
      ]
    ],
    [
      #text(fill: white, size: 8pt)[
        *Project:* Kalasin Automation #linebreak()
        *PLC:* Rockwell 1756 ControlLogix #linebreak()
        *Status:* Engineering Approved #linebreak()
        *Date:* September 2026
      ]
    ]
  )
]

#v(6pt)

// --- KPI CARDS ROW ---
#grid(
  columns: (1fr, 1fr, 1fr, 1fr),
  gutter: 8pt,
  [
    #rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 7pt)[
      #text(size: 7.5pt, fill: rgb("64748b"))[*Total I/O Points*] #linebreak()
      #text(size: 15pt, fill: rgb("1e3a8a"), weight: "bold")[{grand_total}] #linebreak()
      #text(size: 7.5pt, fill: rgb("94a3b8"))[C1: {c1_total} | C2: {c2_total} Channels]
    ]
  ],
  [
    #rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 7pt)[
      #text(size: 7.5pt, fill: rgb("64748b"))[*Active Field Signals*] #linebreak()
      #text(size: 15pt, fill: rgb("15803d"), weight: "bold")[{grand_active}] #linebreak()
      #text(size: 7.5pt, fill: rgb("15803d"))[{grand_active / grand_total * 100:.1f}% System Utilization]
    ]
  ],
  [
    #rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 7pt)[
      #text(size: 7.5pt, fill: rgb("64748b"))[*Design Spare Channels*] #linebreak()
      #text(size: 15pt, fill: rgb("b45309"), weight: "bold")[{grand_spare}] #linebreak()
      #text(size: 7.5pt, fill: rgb("b45309"))[{grand_spare / grand_total * 100:.1f}% Spare for Expansion]
    ]
  ],
  [
    #rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 7pt)[
      #text(size: 7.5pt, fill: rgb("64748b"))[*Terminal Overlap Status*] #linebreak()
      #text(size: 13pt, fill: rgb("15803d"), weight: "bold")[0% Collision] #linebreak()
      #text(size: 7.5pt, fill: rgb("15803d"))[Fully Unique Slot-Tied IDs]
    ]
  ]
)

#v(6pt)

// --- ARCHITECTURE EXPLANATION ---
#rect(width: 100%, fill: rgb("eff6ff"), stroke: 1pt + rgb("bfdbfe"), radius: 4pt, inset: (x: 10pt, y: 7pt))[
  #grid(
    columns: (1.1fr, 1fr),
    gutter: 12pt,
    [
      #text(weight: "bold", fill: rgb("1e40af"), size: 8.5pt)[หลักการออกแบบ Option 1 (Direct Slot Correlation)] #linebreak()
      #text(size: 8pt)[
        • *ตั้งชื่อ Terminal Block ตาม Slot ตรงๆ (1:1):* Terminal Strip ทุกชุดจะใช้ชื่อ `TB-` ตามด้วยรหัส Slot เช่น `TB-C1S4`, `TB-C2S1`, `TB-C2S6` #linebreak()
        • *แก้ปัญหาทับซ้อนถาวร:* เดิม `C1S4` และ `C2S1` ชี้ไปที่ `TBDI1` ซ้ำกัน ใน Option 1 จะแยกเป็น `TB-C1S4` และ `TB-C2S1` ชัดเจน ไม่สับสน #linebreak()
        • *Self-Documenting:* ช่างหน้างานเห็นป้าย Terminal รู้ทันทีว่าเป็นการ์ดใบไหนในตู้ PLC โดยไม่ต้องเปิดตารางเทียบ
      ]
    ],
    [
      #text(weight: "bold", fill: rgb("1e40af"), size: 8.5pt)[โครงสร้างสูตร Wire Mark (Bidirectional Tagging)] #linebreak()
      #text(size: 8pt)[
        • *มาร์กปลายสายฝั่ง Field / Terminal:* `CxSx:Pin / TB-CxSx:Pin` #linebreak()
        #h(8pt)DI ตัวอย่าง: #text(fill: rgb("1e3a8a"), weight: "bold")[C1S4:1 / TB-C1S4:1A] หรือ #text(fill: rgb("1e3a8a"), weight: "bold")[C2S1:1 / TB-C2S1:1A] #linebreak()
        • *มาร์กปลายสายฝั่ง PLC Swing Arm:* `TB-CxSx:Pin / CxSx:Pin` #linebreak()
        #h(8pt)DI ตัวอย่าง: #text(fill: rgb("0f172a"), weight: "bold")[TB-C1S4:1A / C1S4:1] หรือ #text(fill: rgb("0f172a"), weight: "bold")[TB-C2S1:1A / C2S1:1] #linebreak()
        • *รูปแบบ Pin:* 32-pt DI/DO ใช้ `1A..16A` และ `1B..16B` | 16-pt AI ใช้ `1(+) .. 16(+)`
      ]
    ]
  )
]

#v(6pt)

// --- TABLE 1: CHASSIS C1 ---
#text(weight: "bold", fill: rgb("1e3a8a"), size: 9.5pt)[
  1. Chassis 1 (MCP-M1 - Main Control Cabinet) | 8 I/O Cards (224 Points)
]
#v(2pt)

#table(
  columns: (26pt, 48pt, 48pt, 36pt, 75pt, 40pt, 38pt, 38pt, 65pt, 75pt, 130pt),
  stroke: (x, y) => if y == 0 {{ 1pt + rgb("1e3a8a") }} else {{ 0.5pt + rgb("cbd5e1") }},
  fill: (col, row) => if row == 0 {{ rgb("1e3a8a") }} else if calc.even(row) {{ rgb("f8fafc") }} else {{ white }},
  align: (col, row) => if row == 0 {{ center + horizon }} else if col in (0, 1, 2, 3, 5, 6, 7, 8, 9) {{ center + horizon }} else {{ left + horizon }},
  inset: 4pt,

  text(fill: white, weight: "bold", size: 7.5pt)[No.],
  text(fill: white, weight: "bold", size: 7.5pt)[Physical Slot],
  text(fill: white, weight: "bold", size: 7.5pt)[Slot Code],
  text(fill: white, weight: "bold", size: 7.5pt)[Type],
  text(fill: white, weight: "bold", size: 7.5pt)[Module Model],
  text(fill: white, weight: "bold", size: 7.5pt)[Total],
  text(fill: white, weight: "bold", size: 7.5pt)[Active],
  text(fill: white, weight: "bold", size: 7.5pt)[Spare],
  text(fill: white, weight: "bold", size: 7.5pt)[Terminal Strip],
  text(fill: white, weight: "bold", size: 7.5pt)[Pin Range],
  text(fill: white, weight: "bold", size: 7.5pt)[Sample Field Wire Mark Tag],

{c1_rows_str}
)

#v(3pt)
#text(size: 7pt, fill: rgb("64748b"), style: "italic")[
  *หมายเหตุ C1:* Slot S0 คือ CPU (1756-L950TPSXT) และ Slot S1..S3 คือการ์ดสื่อสาร Ethernet (1756-EN4TR จำนวน 3 ใบ) การ์ด I/O เริ่มต้นที่ Slot S4 เป็นต้นไป
]

#pagebreak()

// --- PAGE 2: CHASSIS C2 ---
#text(weight: "bold", fill: rgb("1e3a8a"), size: 9.5pt)[
  2. Chassis 2 (MCP-M2 - Expansion I/O Cabinet) | 13 I/O Cards (327 Points)
]
#v(2pt)

#table(
  columns: (26pt, 48pt, 48pt, 36pt, 75pt, 40pt, 38pt, 38pt, 65pt, 75pt, 130pt),
  stroke: (x, y) => if y == 0 {{ 1pt + rgb("1e3a8a") }} else {{ 0.5pt + rgb("cbd5e1") }},
  fill: (col, row) => if row == 0 {{ rgb("1e3a8a") }} else if calc.even(row) {{ rgb("f8fafc") }} else {{ white }},
  align: (col, row) => if row == 0 {{ center + horizon }} else if col in (0, 1, 2, 3, 5, 6, 7, 8, 9) {{ center + horizon }} else {{ left + horizon }},
  inset: 4pt,

  text(fill: white, weight: "bold", size: 7.5pt)[No.],
  text(fill: white, weight: "bold", size: 7.5pt)[Physical Slot],
  text(fill: white, weight: "bold", size: 7.5pt)[Slot Code],
  text(fill: white, weight: "bold", size: 7.5pt)[Type],
  text(fill: white, weight: "bold", size: 7.5pt)[Module Model],
  text(fill: white, weight: "bold", size: 7.5pt)[Total],
  text(fill: white, weight: "bold", size: 7.5pt)[Active],
  text(fill: white, weight: "bold", size: 7.5pt)[Spare],
  text(fill: white, weight: "bold", size: 7.5pt)[Terminal Strip],
  text(fill: white, weight: "bold", size: 7.5pt)[Pin Range],
  text(fill: white, weight: "bold", size: 7.5pt)[Sample Field Wire Mark Tag],

{c2_rows_str}
)

#v(3pt)
#text(size: 7pt, fill: rgb("64748b"), style: "italic")[
  *หมายเหตุ C2:* Slot S0 คือการ์ดสื่อสาร Ethernet Adapter (1756-EN4TR) การ์ด I/O เริ่มต้นที่ Slot S1 ถึง S13 (DI: S1..S5, DO: S6..S8, AI: S9..S13)
]

#v(8pt)

// --- COMPARISON BLOCK ---
#rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 8pt)[
  #text(weight: "bold", fill: rgb("1e3a8a"), size: 9pt)[ตารางเปรียบเทียบก่อน-หลังการแก้ไขปัญหา Overlap (Before vs After)]
  #v(3pt)
  #table(
    columns: (1fr, 1.2fr, 1.2fr, 2.5fr),
    stroke: 0.5pt + rgb("cbd5e1"),
    fill: (col, row) => if row == 0 {{ rgb("f1f5f9") }} else {{ white }},
    inset: 4pt,
    text(weight: "bold", size: 7.5pt)[I/O Card],
    text(weight: "bold", size: 7.5pt)[Rev 01 เดิม (ชนกัน)],
    text(weight: "bold", size: 7.5pt)[Option 1 ใหม่ (แก้ไขแล้ว)],
    text(weight: "bold", size: 7.5pt)[ผลลัพธ์ทางวิศวกรรมและการทำงานจริง],

    [C1 Slot 4 (DI 1)],
    [TBDI1 (Pins 1A..16B)],
    [*TB-C1S4* (Pins 1A..16B)],
    [ระบุชัดเจนว่าเป็น Terminal Strip ของตู้ C1 Slot 4 ไม่สับสน],

    [C2 Slot 1 (DI 1)],
    [#text(fill: rgb("dc2626"), weight: "bold")[TBDI1 (ชนกับ C1S4!)]],
    [#text(fill: rgb("15803d"), weight: "bold")[*TB-C2S1* (Pins 1A..16B)]],
    [#text(fill: rgb("15803d"))[*หมดปัญหาทับซ้อน 100%* เชื่อมตรงไปยังการ์ด C2S1 ทันที]],

    [C2 Slot 6 (DO 1)],
    [#text(fill: rgb("dc2626"))[RL.1 (ชนกับ C1S8!)]],
    [#text(fill: rgb("15803d"), weight: "bold")[*TB-C2S6* (Pins 1A..16B)]],
    [แยกย่านอย่างเป็นอิสระ สามารถต่อตรงหรือผ่าน Fuse Terminal ได้อย่างเป็นระเบียบ]
  )
]

#v(8pt)
#align(center)[
  #text(size: 7.5pt, fill: rgb("94a3b8"))[
    *เอกสารไฟล์ Excel ฉบับเต็ม:* `Tags list PLC-Option1-DirectSlot-C1-C2.xlsx` (พร้อมชีต Summary, All_Marks_Sorted_CxSx และ Classic ePlan Tabs)
  ]
]
"""

    return typst_code

def main():
    print("Loading summary data from Excel...")
    rows = load_summary_data()
    print(f"Loaded {len(rows)} slot cards.")

    typst_code = generate_typst_content(rows)
    with open(TYPST_SRC_PATH, "w", encoding="utf-8") as f:
        f.write(typst_code)
    print(f"Wrote Typst source to {TYPST_SRC_PATH}")

    # Compile with Typst
    cmd = ["/opt/homebrew/bin/typst", "compile", TYPST_SRC_PATH, PDF_OUT_1]
    print(f"Compiling Typst to {PDF_OUT_1}...")
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        print("Typst compile error:")
        print(res.stderr)
        sys.exit(1)
    print("Typst compile successful!")

    # Copy to PDF_OUT_2
    os.makedirs(os.path.dirname(PDF_OUT_2), exist_ok=True)
    shutil.copyfile(PDF_OUT_1, PDF_OUT_2)
    print(f"Copied PDF to {PDF_OUT_2}")

if __name__ == "__main__":
    main()
