
#set page(
  paper: "a4",
  flipped: true,
  margin: (top: 1.2cm, bottom: 1.2cm, left: 1.4cm, right: 1.4cm),
  header: context {
    grid(
      columns: (1fr, 1fr),
      align: (left, right),
      text(size: 8pt, fill: rgb("64748b"), font: "Arial")[
        *KALASIN ENGINEERING* | Turnkey xCIP-1545 Automation System (Ref: x2608003)
      ],
      text(size: 8pt, fill: rgb("64748b"), font: "Arial")[
        *Chassis C1 Destination & I/O Summary Report* | Page #counter(page).display()
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
      text(size: 8pt, fill: rgb("94a3b8"), font: "Arial")[
        Client: Ingredion (Thailand) Co., Ltd. / ISHITON | Document Ref: KAL-IO-REP-C1-001
      ],
      text(size: 8pt, fill: rgb("94a3b8"), font: "Arial")[
        Source File: `IO_List_xDev-R01-Tag35-7.xlsx` | Approved Rev: 3.6
      ]
    )
  }
)

#set text(font: "Arial", size: 8.5pt, fill: rgb("0f172a"))
#set par(justify: true, leading: 0.5em)

// --- TITLE BLOCK ---
#block(
  fill: rgb("1e293b"),
  inset: (x: 16pt, y: 11pt),
  radius: 4pt,
  width: 100%
)[
  #grid(
    columns: (2.5fr, 1fr),
    align: (left + horizon, right + horizon),
    [
      #text(fill: white, size: 14pt, weight: "bold")[
        CHASSIS C1: DESTINATION ROUTING & I/O POINT SUMMARY
      ]
      #v(2pt)
      #text(fill: rgb("93c5fd"), size: 9pt, weight: "medium")[
        ROCKWELL 1756 CONTROLLOGIX PROCESSOR RACK (CA1) --- POINT COUNT & FIELD ROUTING SCHEDULE
      ]
    ],
    [
      #text(fill: white, size: 8pt)[
        *Project:* Kalasin xCIP Turnkey #linebreak()
        *Controller:* 1756-L950TPSXT #linebreak()
        *Rack Enclosure:* Control Cabinet CA1
      ]
    ]
  )
]

#v(8pt)

// --- SECTION 1: EXECUTIVE KPI SUMMARY ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[1. Executive Summary & Routing Breakdown]
#v(4pt)

#grid(
  columns: (1fr, 1fr, 1fr, 1fr),
  gutter: 10pt,
  block(fill: rgb("f1f5f9"), stroke: 1pt + rgb("cbd5e1"), inset: 8pt, radius: 3pt)[
    #text(size: 8pt, fill: rgb("64748b"))[*TOTAL C1 CHANNELS*] #linebreak()
    #text(size: 16pt, weight: "bold", fill: rgb("0f172a"))[228 Points] #linebreak()
    #text(size: 7.5pt, fill: rgb("475569"))[Configured across Slots 00--11]
  ],
  block(fill: rgb("dcfce7"), stroke: 1pt + rgb("86efac"), inset: 8pt, radius: 3pt)[
    #text(size: 8pt, fill: rgb("166534"))[*ACTIVE SIGNALS*] #linebreak()
    #text(size: 16pt, weight: "bold", fill: rgb("14532d"))[138 Points] #linebreak()
    #text(size: 7.5pt, fill: rgb("15803d"))[60.5% Active Field Instruments]
  ],
  block(fill: rgb("fef9c3"), stroke: 1pt + rgb("fde047"), inset: 8pt, radius: 3pt)[
    #text(size: 8pt, fill: rgb("854d0e"))[*SPARE CAPACITY*] #linebreak()
    #text(size: 16pt, weight: "bold", fill: rgb("713f12"))[90 Points] #linebreak()
    #text(size: 7.5pt, fill: rgb("a16207"))[39.5% Expansion Reserve]
  ],
  block(fill: rgb("e0e7ff"), stroke: 1pt + rgb("a5b4fc"), inset: 8pt, radius: 3pt)[
    #text(size: 8pt, fill: rgb("3730a3"))[*DESTINATIONS*] #linebreak()
    #text(size: 16pt, weight: "bold", fill: rgb("1e1b4b"))[3 Locations] #linebreak()
    #text(size: 7.5pt, fill: rgb("4338ca"))[CA1, JB-401, JB-601]
  ]
)

#v(8pt)

// --- TABLE 1: DESTINATION SUMMARY ---
#table(
  columns: (1.5fr, 2.5fr, 1fr, 1fr, 1fr, 1fr, 1fr, 1fr, 1fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col < 2 { left + horizon } else { center + horizon },
  inset: 5pt,
  
  // Headers
  text(fill: white, weight: "bold", size: 8pt)[Destination],
  text(fill: white, weight: "bold", size: 8pt)[Plant Location / Role],
  text(fill: white, weight: "bold", size: 8pt)[Total Pts],
  text(fill: white, weight: "bold", size: 8pt)[Active],
  text(fill: white, weight: "bold", size: 8pt)[Spare],
  text(fill: white, weight: "bold", size: 8pt)[DI (24V)],
  text(fill: white, weight: "bold", size: 8pt)[DO (24V)],
  text(fill: white, weight: "bold", size: 8pt)[AI (4-20mA)],
  text(fill: white, weight: "bold", size: 8pt)[CPU / Comm],

  // Row 1: CA1
  [*CA1*], [Main Control Room (Local Cabinet Spares & Network)], [36], [0], [36], [16], [16], [0], [4],
  // Row 2: JB-401
  [*JB-401*], [Infeed 2nd Floor (Raw Material Prep & Cooking Infeed)], [112], [85], [27], [64], [32], [16], [0],
  // Row 3: JB-601
  [*JB-601*], [Spray Dryer 1st Floor (Chamber Bottom Discharge & Pumps)], [80], [53], [27], [48], [16], [16], [0],
  
  // Total Row
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[TOTAL]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[Chassis C1 Master Total]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[228]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[138]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[90]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[128]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[64]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[32]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[4]]
)

#v(10pt)

// --- SECTION 2: SLOT-BY-SLOT SCHEDULE ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[2. Chassis C1 Slot-to-Destination Routing Schedule]
#v(4pt)

#table(
  columns: (0.7fr, 1.4fr, 0.9fr, 1fr, 0.8fr, 0.8fr, 0.8fr, 3.5fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col in (0, 2, 3, 4, 5, 6) { center + horizon } else { left + horizon },
  inset: 4.5pt,

  text(fill: white, weight: "bold", size: 8pt)[Slot],
  text(fill: white, weight: "bold", size: 8pt)[Hardware Card],
  text(fill: white, weight: "bold", size: 8pt)[Type],
  text(fill: white, weight: "bold", size: 8pt)[Destination],
  text(fill: white, weight: "bold", size: 8pt)[Points],
  text(fill: white, weight: "bold", size: 8pt)[Active],
  text(fill: white, weight: "bold", size: 8pt)[Spare],
  text(fill: white, weight: "bold", size: 8pt)[Connected Equipment / Purpose],

  // Slot 00
  [00], [1756-L950TPSXT], [CPU], [CA1], [1], [0], [1], [ControlLogix Main Controller Head],
  // Slot 01
  [01], [1756-EN4TR], [ETH-1], [CA1], [1], [0], [1], [Plant Network Port 1 (Enterprise / SCADA)],
  // Slot 02
  [02], [1756-EN4TR], [ETH-2], [CA1], [1], [0], [1], [DLR Device Level Ring Supervisor Node],
  // Slot 03
  [03], [1756-EN4TR], [ETH-3], [CA1], [1], [0], [1], [Historian & Auxiliary Communications Link],
  // Slot 04
  [04], [1756-IB32], [DI], [CA1], [16], [0], [16], [Local Cabinet Spare DI Bank (Channel 00--15)],
  [04], [1756-IB32], [DI], [JB-401], [16], [16], [0], [Infeed Flow & Limit Switches (FS-40101, FS-40102, ZS-40101)],
  // Slot 05
  [05], [1756-IB32], [DI], [JB-401], [32], [32], [0], [Infeed Level Switches (LSH/LSL-40101--08) & Valve Feedback],
  // Slot 06
  [06], [1756-IB32], [DI], [JB-401], [16], [2], [14], [Infeed Interlocks & Permissives],
  [06], [1756-IB32], [DI], [JB-601], [16], [10], [6], [Spray Dryer Chamber Bottom Level & Safety Switches],
  // Slot 07
  [07], [1756-IB32], [DI], [JB-601], [32], [23], [9], [Spray Dryer 1st Fl Flow & Pressure (FS-60101, PSH-60101)],
  // Slot 08
  [08], [1756-OB32], [DO], [CA1], [16], [0], [16], [Local Cabinet Spare DO Bank (Channel 00--15)],
  [08], [1756-OB32], [DO], [JB-401], [16], [16], [0], [Infeed Solenoid Valves (XV-40101--XV-40116)],
  // Slot 09
  [09], [1756-OB32], [DO], [JB-401], [16], [8], [8], [Infeed Process Valves (XV-40117--XV-40124)],
  [09], [1756-OB32], [DO], [JB-601], [16], [12], [4], [Spray Dryer Discharge Solenoids (XV-60101--XV-60112)],
  // Slot 10
  [10], [1756-IF16], [AI], [JB-401], [16], [11], [5], [Infeed Transmitters (PT-40101, TT-40101--04, FT-40101)],
  // Slot 11
  [11], [1756-IF16], [AI], [JB-601], [16], [8], [8], [Spray Dryer 1st Fl Pressure & Temp (PT-60101, TT-60101--06)]
)

#pagebreak()

// --- SECTION 3: COMPLETE POINT-BY-POINT SCHEDULE (228 CHANNELS) ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[3. Chassis C1 Complete Point-by-Point Wiring & Loop Schedule (228 Channels)]
#v(4pt)

#table(
  columns: (0.6fr, 0.5fr, 1.2fr, 0.7fr, 1.4fr, 3.2fr, 1fr, 0.9fr, 0.8fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col in (0, 1, 3, 6, 7, 8) { center + horizon } else { left + horizon },
  inset: 3.5pt,

  text(fill: white, weight: "bold", size: 7.5pt)[Slot],
  text(fill: white, weight: "bold", size: 7.5pt)[Pt],
  text(fill: white, weight: "bold", size: 7.5pt)[Terminal],
  text(fill: white, weight: "bold", size: 7.5pt)[Type],
  text(fill: white, weight: "bold", size: 7.5pt)[Tag],
  text(fill: white, weight: "bold", size: 7.5pt)[Description],
  text(fill: white, weight: "bold", size: 7.5pt)[Dest JB],
  text(fill: white, weight: "bold", size: 7.5pt)[P&ID No.],
  text(fill: white, weight: "bold", size: 7.5pt)[Status],

  [00], [00], [C1S00-X1], [CPU], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [01], [00], [C1S01-X1], [ETH-1], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [02], [00], [C1S02-X1], [ETH-2], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [03], [00], [C1S03-X1], [ETH-3], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [00], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [01], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [02], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [03], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [04], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [05], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [06], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [07], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [08], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [09], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [10], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [11], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [12], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [13], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [14], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [15], [C1S04-X1], [DI], [SPARE], [Spare Reserve Channel], [CA1], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [04], [16], [C1S04-X2], [DI], [FSL-40201], [Flow Switch], [JB-401], [-], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [17], [C1S04-X2], [DI], [FSL-40202], [Flow Switch], [JB-401], [-], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [18], [C1S04-X2], [DI], [FT-40201], [Coriolis Flow Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [19], [C1S04-X2], [DI], [FT-40202], [Magnetic Flow Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [20], [C1S04-X2], [DI], [HV-40201], [Manual Valve With Position Switch], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [21], [C1S04-X2], [DI], [HV-40202], [Manual Valve With Position Switch], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [22], [C1S04-X2], [DI], [HV-40203], [Manual Valve With Position Switch], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [23], [C1S04-X2], [DI], [HV-40204], [Manual Valve With Position Switch], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [24], [C1S04-X2], [DI], [HV-40205], [Manual Valve With Position Switch], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [25], [C1S04-X2], [DI], [HV-40206], [Manual Valve With Position Switch], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [26], [C1S04-X2], [DI], [XV-40201A], [-], [JB-401], [-], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [27], [C1S04-X2], [DI], [XV-40201A], [-], [JB-401], [-], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [28], [C1S04-X2], [DI], [XV-40201B], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [29], [C1S04-X2], [DI], [XV-40201B], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [30], [C1S04-X2], [DI], [XV-40201C], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [04], [31], [C1S04-X2], [DI], [XV-40201C], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [00], [C1S05-X1], [DI], [XV-40201D], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [01], [C1S05-X1], [DI], [XV-40201D], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [02], [C1S05-X1], [DI], [XV-40201E], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [03], [C1S05-X1], [DI], [XV-40201E], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [04], [C1S05-X1], [DI], [XV-40201F], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [05], [C1S05-X1], [DI], [XV-40201F], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [06], [C1S05-X1], [DI], [XV-40201G], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [07], [C1S05-X1], [DI], [XV-40201G], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [08], [C1S05-X1], [DI], [XV-40201H], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [09], [C1S05-X1], [DI], [XV-40201H], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [11], [C1S05-X1], [DI], [XV-40202A], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [12], [C1S05-X1], [DI], [XV-40202B], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [13], [C1S05-X1], [DI], [XV-40202B], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [14], [C1S05-X1], [DI], [XV-40202C], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [15], [C1S05-X1], [DI], [XV-40202C], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [16], [C1S05-X2], [DI], [XV-40202D], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [17], [C1S05-X2], [DI], [XV-40202D], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [18], [C1S05-X2], [DI], [XV-40202E], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [19], [C1S05-X2], [DI], [XV-40202A], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [19], [C1S05-X2], [DI], [XV-40202E], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [20], [C1S05-X2], [DI], [XV-40202F], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [21], [C1S05-X2], [DI], [XV-40202F], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [22], [C1S05-X2], [DI], [XV-40202G], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [23], [C1S05-X2], [DI], [XV-40202G], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [24], [C1S05-X2], [DI], [XV-40202H], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [25], [C1S05-X2], [DI], [XV-40202H], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [26], [C1S05-X2], [DI], [XV-40203B], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [27], [C1S05-X2], [DI], [XV-40203B], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [28], [C1S05-X2], [DI], [XV-40203C], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [29], [C1S05-X2], [DI], [XV-40203C], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [30], [C1S05-X2], [DI], [XV-40203D], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [05], [31], [C1S05-X2], [DI], [XV-40203D], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [00], [C1S06-X1], [DI], [XV-40203E], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [01], [C1S06-X1], [DI], [XV-40203E], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [02], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [03], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [04], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [05], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [06], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [07], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [08], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [09], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [10], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [11], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [12], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [13], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [14], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [15], [C1S06-X1], [DI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [16], [C1S06-X2], [DI], [ES-60202], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [17], [C1S06-X2], [DI], [FS-60205], [Flow Switch], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [18], [C1S06-X2], [DI], [LSL-60201], [-], [JB-601], [325-01-600-PD-11], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [19], [C1S06-X2], [DI], [SV-60201], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [20], [C1S06-X2], [DI], [SV-60201], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [21], [C1S06-X2], [DI], [SV-60202], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [22], [C1S06-X2], [DI], [SV-60202], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [23], [C1S06-X2], [DI], [SV-60203], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [24], [C1S06-X2], [DI], [SV-60203], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [25], [C1S06-X2], [DI], [VS-60205], [Vibration Switch], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [06], [26], [C1S06-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [27], [C1S06-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [28], [C1S06-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [29], [C1S06-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [30], [C1S06-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [06], [31], [C1S06-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [00], [C1S07-X1], [DI], [XA-60202A], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [01], [C1S07-X1], [DI], [XA-60202B], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [02], [C1S07-X1], [DI], [XA-60202C], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [03], [C1S07-X1], [DI], [XA-60202D], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [04], [C1S07-X1], [DI], [XA-60202E], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [05], [C1S07-X1], [DI], [XA-60202F], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [06], [C1S07-X1], [DI], [XA-60202G], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [07], [C1S07-X1], [DI], [XA-60202H], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [08], [C1S07-X1], [DI], [XA-60202I], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [09], [C1S07-X1], [DI], [XA-60202J], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [10], [C1S07-X1], [DI], [XA-60202K], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [11], [C1S07-X1], [DI], [XA-60202L], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [12], [C1S07-X1], [DI], [XA-60202M], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [13], [C1S07-X1], [DI], [XV-60202], [-], [JB-601], [325-01-600-PD-11], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [14], [C1S07-X1], [DI], [XV-60202], [-], [JB-601], [325-01-600-PD-11], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [15], [C1S07-X1], [DI], [XV-60205A], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [16], [C1S07-X2], [DI], [XV-60205A], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [17], [C1S07-X2], [DI], [XV-60240D], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [18], [C1S07-X2], [DI], [XV-60240D], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [19], [C1S07-X2], [DI], [XV-60240D], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [20], [C1S07-X2], [DI], [XV-60240D], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [21], [C1S07-X2], [DI], [YA-60205], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [22], [C1S07-X2], [DI], [YI-60205], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [07], [23], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [24], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [25], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [26], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [27], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [28], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [29], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [30], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [07], [31], [C1S07-X2], [DI], [SPARE], [Spare Reserve Channel], [JB-601], [-], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [00], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [01], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [02], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [03], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [04], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [05], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [06], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [07], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [08], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [09], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [10], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [11], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [12], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [13], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [14], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [15], [C1S08-X1], [DO], [SPARE], [Spare Reserve Channel], [CA1], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [08], [16], [C1S08-X2], [DO], [XV-40202E], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [17], [C1S08-X2], [DO], [XV-40202F], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [18], [C1S08-X2], [DO], [XV-40202G], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [19], [C1S08-X2], [DO], [XV-40202H], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [20], [C1S08-X2], [DO], [XV-40203B], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [21], [C1S08-X2], [DO], [XV-40203C], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [22], [C1S08-X2], [DO], [XV-40203D], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [23], [C1S08-X2], [DO], [XV-40203E], [-], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [24], [C1S08-X2], [DO], [SV-40201], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [25], [C1S08-X2], [DO], [SV-40202], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [26], [C1S08-X2], [DO], [SV-40203], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [27], [C1S08-X2], [DO], [SV-40204], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [28], [C1S08-X2], [DO], [XV-40201A], [-], [JB-401], [-], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [29], [C1S08-X2], [DO], [XV-40201B], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [30], [C1S08-X2], [DO], [XV-40201C], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [08], [31], [C1S08-X2], [DO], [XV-40201D], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [00], [C1S09-X1], [DO], [XV-40201E], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [01], [C1S09-X1], [DO], [XV-40201F], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [02], [C1S09-X1], [DO], [XV-40201G], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [03], [C1S09-X1], [DO], [XV-40201H], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [04], [C1S09-X1], [DO], [XV-40202A], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [05], [C1S09-X1], [DO], [XV-40202B], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [06], [C1S09-X1], [DO], [XV-40202C], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [07], [C1S09-X1], [DO], [XV-40202D], [-], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [08], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [09], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [10], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [11], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [12], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [13], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [14], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [15], [C1S09-X1], [DO], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [16], [C1S09-X2], [DO], [SV-60201], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [17], [C1S09-X2], [DO], [SV-60202], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [18], [C1S09-X2], [DO], [SV-60230], [-], [JB-601], [325-01-600-PD-04], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [19], [C1S09-X2], [DO], [SV-60203], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [20], [C1S09-X2], [DO], [SV-60240A], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [21], [C1S09-X2], [DO], [SV-60231], [-], [JB-601], [325-01-600-PD-04], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [22], [C1S09-X2], [DO], [SV-60240B], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [23], [C1S09-X2], [DO], [SV-60240C], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [24], [C1S09-X2], [DO], [XV-60240D], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [25], [C1S09-X2], [DO], [XV-60251], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [26], [C1S09-X2], [DO], [XV-60202], [-], [JB-601], [325-01-600-PD-11], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [27], [C1S09-X2], [DO], [XV-60205A], [-], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [09], [28], [C1S09-X2], [DO], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [29], [C1S09-X2], [DO], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [30], [C1S09-X2], [DO], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [09], [31], [C1S09-X2], [DO], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [10], [00], [C1S10-X1], [AI], [FT-40201], [Coriolis Flow Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [01], [C1S10-X1], [AI], [FT-40201], [Coriolis Flow Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [02], [C1S10-X1], [AI], [FT-40202], [Magnetic Flow Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [03], [C1S10-X1], [AI], [LT-40203], [Level Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [04], [C1S10-X1], [AI], [LT-40201], [Level Transmitter], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [05], [C1S10-X1], [AI], [LT-40202], [Level Transmitter], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [06], [C1S10-X1], [AI], [pHT-40201], [PH Transmitter], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [07], [C1S10-X1], [AI], [pHT-40202], [PH Transmitter], [JB-401], [325-01-400-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [08], [C1S10-X1], [AI], [PT-40201], [Pressure Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [09], [C1S10-X1], [AI], [PT-40202], [Pressure Transmitter], [JB-401], [325-01-400-PD-02], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [10], [C1S10-X1], [AI], [PT-60202], [Pressure Transmitter], [JB-401], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [10], [11], [C1S10-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [10], [12], [C1S10-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [10], [13], [C1S10-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [10], [14], [C1S10-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [10], [15], [C1S10-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-401], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [00], [C1S11-X1], [AI], [DPI-60240], [-], [JB-601], [325-01-600-PD-05], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [01], [C1S11-X1], [AI], [DPT-60202], [Differential Pressure Transmitte], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [02], [C1S11-X1], [AI], [FT-60201], [-], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [03], [C1S11-X1], [AI], [FT-60204], [-], [JB-601], [325-01-600-PD-10], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [04], [C1S11-X1], [AI], [FT-61307], [-], [JB-601], [-], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [05], [C1S11-X1], [AI], [PT-60203], [Pressure Transmitter], [JB-601], [325-01-600-PD-08], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [06], [C1S11-X1], [AI], [PT-60204], [Pressure Transmitter], [JB-601], [325-01-600-PD-10], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [07], [C1S11-X1], [AI], [PT-60205], [Pressure Transmitter], [JB-601], [325-01-600-PD-01], table.cell(fill: rgb("dcfce7"))[#text(fill: rgb("14532d"), size: 7pt, weight: "bold")[ACTIVE]],
  [11], [08], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [09], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [10], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [11], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [12], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [13], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [14], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
  [11], [15], [C1S11-X1], [AI], [SPARE], [Spare Reserve Channel], [JB-601], [], table.cell(fill: rgb("fef9c3"))[#text(fill: rgb("854d0e"), size: 7pt, weight: "regular")[SPARE]],
)
