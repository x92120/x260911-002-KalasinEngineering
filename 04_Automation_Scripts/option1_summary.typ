
#set page(
  paper: "a4",
  flipped: true,
  margin: (top: 1.4cm, bottom: 1.4cm, left: 1.6cm, right: 1.6cm),
  header: context {
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
  },
  footer: context {
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
  }
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
      #text(size: 15pt, fill: rgb("1e3a8a"), weight: "bold")[557] #linebreak()
      #text(size: 7.5pt, fill: rgb("94a3b8"))[C1: 224 | C2: 333 Channels]
    ]
  ],
  [
    #rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 7pt)[
      #text(size: 7.5pt, fill: rgb("64748b"))[*Active Field Signals*] #linebreak()
      #text(size: 15pt, fill: rgb("15803d"), weight: "bold")[381] #linebreak()
      #text(size: 7.5pt, fill: rgb("15803d"))[68.4% System Utilization]
    ]
  ],
  [
    #rect(width: 100%, fill: rgb("f8fafc"), stroke: 1pt + rgb("e2e8f0"), radius: 4pt, inset: 7pt)[
      #text(size: 7.5pt, fill: rgb("64748b"))[*Design Spare Channels*] #linebreak()
      #text(size: 15pt, fill: rgb("b45309"), weight: "bold")[176] #linebreak()
      #text(size: 7.5pt, fill: rgb("b45309"))[31.6% Spare for Expansion]
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
  stroke: (x, y) => if y == 0 { 1pt + rgb("1e3a8a") } else { 0.5pt + rgb("cbd5e1") },
  fill: (col, row) => if row == 0 { rgb("1e3a8a") } else if calc.even(row) { rgb("f8fafc") } else { white },
  align: (col, row) => if row == 0 { center + horizon } else if col in (0, 1, 2, 3, 5, 6, 7, 8, 9) { center + horizon } else { left + horizon },
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

  text(size: 8pt)[1],
  text(size: 8pt, weight: "medium")[Slot 4],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S4],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[16],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[16],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S4],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C1S4:1/TB-C1S4:1A],
  text(size: 8pt)[2],
  text(size: 8pt, weight: "medium")[Slot 5],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S5],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[32],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[0],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S5],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C1S5:1/TB-C1S5:1A],
  text(size: 8pt)[3],
  text(size: 8pt, weight: "medium")[Slot 6],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S6],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[12],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[20],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S6],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C1S6:1/TB-C1S6:1A],
  text(size: 8pt)[4],
  text(size: 8pt, weight: "medium")[Slot 7],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S7],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[23],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[9],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S7],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C1S7:1/TB-C1S7:1A],
  text(size: 8pt)[5],
  text(size: 8pt, weight: "medium")[Slot 8],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S8],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[DO],
  text(size: 8pt)[1756-OB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[16],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[16],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S8],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C1S8:1/TB-C1S8:1A],
  text(size: 8pt)[6],
  text(size: 8pt, weight: "medium")[Slot 9],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S9],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[DO],
  text(size: 8pt)[1756-OB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[20],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[12],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S9],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C1S9:1/TB-C1S9:1A],
  text(size: 8pt)[7],
  text(size: 8pt, weight: "medium")[Slot 10],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S10],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[11],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[5],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S10],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C1S10:1/TB-C1S10:1(+)],
  text(size: 8pt)[8],
  text(size: 8pt, weight: "medium")[Slot 11],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C1S11],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[8],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[8],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C1S11],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C1S11:1/TB-C1S11:1(+)],
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
  stroke: (x, y) => if y == 0 { 1pt + rgb("1e3a8a") } else { 0.5pt + rgb("cbd5e1") },
  fill: (col, row) => if row == 0 { rgb("1e3a8a") } else if calc.even(row) { rgb("f8fafc") } else { white },
  align: (col, row) => if row == 0 { center + horizon } else if col in (0, 1, 2, 3, 5, 6, 7, 8, 9) { center + horizon } else { left + horizon },
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

  text(size: 8pt)[1],
  text(size: 8pt, weight: "medium")[Slot 1],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S1],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[24],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[8],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S1],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S1:1/TB-C2S1:1A],
  text(size: 8pt)[2],
  text(size: 8pt, weight: "medium")[Slot 2],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S2],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[23],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[9],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S2],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S2:1/TB-C2S2:1A],
  text(size: 8pt)[3],
  text(size: 8pt, weight: "medium")[Slot 3],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S3],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[29],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[29],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[0],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S3],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S3:1/TB-C2S3:1A],
  text(size: 8pt)[4],
  text(size: 8pt, weight: "medium")[Slot 4],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S4],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[32],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[0],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S4],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S4:1/TB-C2S4:1A],
  text(size: 8pt)[5],
  text(size: 8pt, weight: "medium")[Slot 5],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S5],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[DI],
  text(size: 8pt)[1756-IB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[20],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[12],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S5],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S5:1/TB-C2S5:1A],
  text(size: 8pt)[6],
  text(size: 8pt, weight: "medium")[Slot 6],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S6],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[DO],
  text(size: 8pt)[1756-OB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[12],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[20],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S6],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S6:1/TB-C2S6:1A],
  text(size: 8pt)[7],
  text(size: 8pt, weight: "medium")[Slot 7],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S7],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[DO],
  text(size: 8pt)[1756-OB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[31],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[1],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S7],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S7:1/TB-C2S7:1A],
  text(size: 8pt)[8],
  text(size: 8pt, weight: "medium")[Slot 8],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S8],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[DO],
  text(size: 8pt)[1756-OB32-0],
  text(size: 8pt)[32],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[9],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[23],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S8],
  text(size: 8pt)[1A .. 16B],
  text(size: 7.5pt, font: "Menlo")[C2S8:1/TB-C2S8:1A],
  text(size: 8pt)[9],
  text(size: 8pt, weight: "medium")[Slot 9],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S9],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[7],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[9],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S9],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C2S9:1/TB-C2S9:1(+)],
  text(size: 8pt)[10],
  text(size: 8pt, weight: "medium")[Slot 10],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S10],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[16],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[0],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S10],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C2S10:1/TB-C2S10:1(+)],
  text(size: 8pt)[11],
  text(size: 8pt, weight: "medium")[Slot 11],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S11],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[13],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[3],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S11],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C2S11:1/TB-C2S11:1(+)],
  text(size: 8pt)[12],
  text(size: 8pt, weight: "medium")[Slot 12],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S12],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[16],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[0],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S12],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C2S12:1/TB-C2S12:1(+)],
  text(size: 8pt)[13],
  text(size: 8pt, weight: "medium")[Slot 13],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[C2S13],
  text(size: 8pt, weight: "bold", fill: rgb("0369a1"))[AI],
  text(size: 8pt)[1756-IF16-0],
  text(size: 8pt)[16],
  text(size: 8pt, weight: "bold", fill: rgb("15803d"))[11],
  text(size: 8pt, weight: "bold", fill: rgb("b45309"))[5],
  text(size: 8pt, weight: "bold", fill: rgb("1e3a8a"))[TB-C2S13],
  text(size: 8pt)[1(+) .. 16(+)],
  text(size: 7.5pt, font: "Menlo")[C2S13:1/TB-C2S13:1(+)],
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
    fill: (col, row) => if row == 0 { rgb("f1f5f9") } else { white },
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
