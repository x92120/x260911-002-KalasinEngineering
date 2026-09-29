
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
        *Chassis C1--C8 Slot to Destination Junction Box Master Report* | Page #counter(page).display()
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
        Client: Ingredion (Thailand) Co., Ltd. / ISHITON | Document Ref: KAL-IO-REP-C1C8-001
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
        CHASSIS C1 -- C8: MASTER SLOT TO DESTINATION JB REPORT
      ]
      #v(2pt)
      #text(fill: rgb("93c5fd"), size: 9pt, weight: "medium")[
        COMPLETE PLANT AUTOMATION I/O ROUTING DIRECTORY & FIELD DISTRIBUTION SCHEDULE
      ]
    ],
    [
      #text(fill: white, size: 8pt)[
        *Project Ref:* x2608003 #linebreak()
        *Scope:* Turnkey Automated CIP Skid System #linebreak()
        *Drawing Ref:* KAL-SYS-DIA-C1C8-001
      ]
    ]
  )
]

#v(8pt)

// --- SECTION 1: SYSTEM OVERVIEW ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[1. Master Chassis System Overview (C1 through C8)]
#v(4pt)

#table(
  columns: (1fr, 3.2fr, 2.5fr, 1fr, 1.1fr, 1fr, 1fr, 0.8fr, 0.8fr, 0.8fr, 1fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col < 3 { left + horizon } else { center + horizon },
  inset: 4.5pt,

  text(fill: white, weight: "bold", size: 8pt)[Rack],
  text(fill: white, weight: "bold", size: 8pt)[Chassis Description & Functional Role],
  text(fill: white, weight: "bold", size: 8pt)[Enclosure Location],
  text(fill: white, weight: "bold", size: 8pt)[Slots],
  text(fill: white, weight: "bold", size: 8pt)[Total Pts],
  text(fill: white, weight: "bold", size: 8pt)[Active],
  text(fill: white, weight: "bold", size: 8pt)[Spare],
  text(fill: white, weight: "bold", size: 8pt)[DI],
  text(fill: white, weight: "bold", size: 8pt)[DO],
  text(fill: white, weight: "bold", size: 8pt)[AI],
  text(fill: white, weight: "bold", size: 8pt)[AO / BUS],

  [*C1*], [Main Controller Rack (1756-L950TPSXT)], [Control Cabinet CA1 (MCR)], [12], [228], [138], [90], [128], [64], [32], [0],
  [*C2*], [Main Expansion Rack (High-Density I/O)], [Control Cabinet CA1 (MCR)], [11], [286], [203], [83], [157], [96], [32], [0],
  [*C3*], [Remote I/O Rack 1 (Spray Dryer / Tower)], [Field RIO Cabinet], [9], [109], [70], [39], [16], [32], [52], [8],
  [*C4*], [Remote I/O Rack 2 (Upper Spray Dryer)], [Field RIO Cabinet], [5], [89], [44], [45], [48], [24], [16], [0],
  [*C5*], [Remote I/O Station 5 (Slurry Building)], [Slurry Building 2nd Floor], [5], [113], [57], [56], [64], [32], [16], [0],
  [*C6*], [Motor Control Hardwired Interlocks], [MCC Room Ground Floor], [2], [18], [18], [0], [10], [8], [0], [0],
  [*C7*], [MCC & Fieldbus Supervisory System], [MCC Room Ground Floor], [8], [226], [158], [68], [118], [88], [0], [20],
  [*C8*], [Intrinsically Safe (IS) Field Barriers], [Field Ex Barrier Enclosures], [7], [149], [75], [74], [95], [8], [46], [0],

  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[TOTAL]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[All 8 Chassis Combined]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[All Plant Enclosures]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[59]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[1218]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[763]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[455]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[636]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[352]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[194]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[28]]
)


#v(10pt)

// --- SECTION 2: DESTINATION DIRECTORY ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[2. Destination Junction Box Directory (16 Destinations)]
#v(4pt)

#table(
  columns: (1.5fr, 2.5fr, 3.5fr, 1.5fr, 1.1fr, 1.1fr, 1.1fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col < 3 { left + horizon } else { center + horizon },
  inset: 4.5pt,

  text(fill: white, weight: "bold", size: 8pt)[Destination],
  text(fill: white, weight: "bold", size: 8pt)[Location],
  text(fill: white, weight: "bold", size: 8pt)[Process Functional Scope],
  text(fill: white, weight: "bold", size: 8pt)[Feeding Racks],
  text(fill: white, weight: "bold", size: 8pt)[Total Pts],
  text(fill: white, weight: "bold", size: 8pt)[Active],
  text(fill: white, weight: "bold", size: 8pt)[Spare],

  [*CA1*], [Main Control Room], [Central PLC Automation & Network Cabinets], [C1, C2, C3, C4, C5], [40], [0], [40],
  [*IS-JB-603*], [Spray Dryer 3rd Floor (IS)], [Intrinsically Safe Ex Zone Chamber Box], [C8], [32], [22], [10],
  [*IS-JB-608*], [Spray Dryer 8th Floor (IS)], [Intrinsically Safe Ex Zone Top Level Box], [C8], [32], [14], [18],
  [*IS-JB-612*], [Packing Tower 2nd Fl (IS)], [Intrinsically Safe Ex Zone Lower Tower Box], [C8], [41], [17], [24],
  [*IS-JB-618*], [Packing Tower 8th Fl (IS)], [Intrinsically Safe Ex Zone Upper Tower Box], [C8], [44], [22], [22],
  [*JB-401*], [Infeed 2nd Floor], [Infeed Raw Material Prep & Wet Mixing], [C1], [112], [85], [27],
  [*JB-402*], [Jet Cooker 2nd Floor], [Jet Cooker Thermal Cooking & Starch Conversion], [C2, C3], [65], [41], [24],
  [*JB-601*], [Spray Dryer 1st Floor], [Spray Dryer Bottom Discharge & Pumps], [C1, C7], [85], [58], [27],
  [*JB-602*], [Spray Dryer 3rd Floor], [Spray Dryer Middle Body & Chamber Instruments], [C2, C3], [82], [57], [25],
  [*JB-606*], [Spray Dryer 6th Floor], [Spray Dryer Air Filter & Observation Level], [C4], [26], [7], [19],
  [*JB-607*], [Spray Dryer 7th Floor], [Spray Dryer Cyclone & Bag Filter Exhaust], [C2, C3, C7], [195], [151], [44],
  [*JB-608*], [Spray Dryer 8th Floor], [Spray Dryer Air Dispenser & Scrubber Top], [C4], [6], [1], [5],
  [*JB-612*], [Packing Tower 2nd Floor], [Packing Tower Product Discharge & Conveyor], [C4, C7], [57], [37], [20],
  [*JB-618*], [Packing Tower 8th Floor], [Packing Tower Dehumidifier & Cyclone Air], [C3], [52], [25], [27],
  [*MCC*], [MCC Room Ground Floor], [Main Motor Control Center & VFD Inverters], [C6, C7], [237], [169], [68],
  [*RIO-200*], [Slurry Out-Building 2nd Fl], [Remote Slurry Processing Building Station], [C5], [112], [57], [55],

)

#pagebreak()

// --- SECTION 3: AGGREGATED SLOT TO JB SCHEDULE ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[3. Master Chassis C1--C8 Slot-to-Destination Schedule]
#v(4pt)

#table(
  columns: (0.7fr, 0.7fr, 1.5fr, 0.8fr, 1.2fr, 2fr, 0.8fr, 0.8fr, 0.8fr, 3.5fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col in (0, 1, 3, 4, 6, 7, 8) { center + horizon } else { left + horizon },
  inset: 4pt,

  text(fill: white, weight: "bold", size: 7.5pt)[Chassis],
  text(fill: white, weight: "bold", size: 7.5pt)[Slot],
  text(fill: white, weight: "bold", size: 7.5pt)[Module Card],
  text(fill: white, weight: "bold", size: 7.5pt)[Signal],
  text(fill: white, weight: "bold", size: 7.5pt)[Dest JB],
  text(fill: white, weight: "bold", size: 7.5pt)[Location],
  text(fill: white, weight: "bold", size: 7.5pt)[Total],
  text(fill: white, weight: "bold", size: 7.5pt)[Active],
  text(fill: white, weight: "bold", size: 7.5pt)[Spare],
  text(fill: white, weight: "bold", size: 7.5pt)[Sample Connected Loops / Tags],

  [*C1*], [00], [1756-L950TPSXT], [CPU], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C1*], [01], [1756-EN4TR], [ETH-1], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C1*], [02], [1756-EN4TR], [ETH-2], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C1*], [03], [1756-EN4TR], [ETH-3], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C1*], [04], [1756-IB32], [DI], [*CA1*], [Main Control Room], [16], [0], [16], [Reserve / Spare Bank],
  [*C1*], [04], [1756-IB32], [DI], [*JB-401*], [Infeed 2nd Floor], [16], [16], [0], [FSL-40201, FSL-40202, FT-40201 (+13 more)],
  [*C1*], [05], [1756-IB32], [DI], [*JB-401*], [Infeed 2nd Floor], [32], [32], [0], [XV-40201D, XV-40201D, XV-40201E (+29 more)],
  [*C1*], [06], [1756-IB32], [DI], [*JB-401*], [Infeed 2nd Floor], [16], [2], [14], [XV-40203E, XV-40203E],
  [*C1*], [06], [1756-IB32], [DI], [*JB-601*], [Spray Dryer 1st Floor], [16], [10], [6], [ES-60202, FS-60205, LSL-60201 (+7 more)],
  [*C1*], [07], [1756-IB32], [DI], [*JB-601*], [Spray Dryer 1st Floor], [32], [23], [9], [XA-60202A, XA-60202B, XA-60202C (+20 more)],
  [*C1*], [08], [1756-OB32], [DO], [*CA1*], [Main Control Room], [16], [0], [16], [Reserve / Spare Bank],
  [*C1*], [08], [1756-OB32], [DO], [*JB-401*], [Infeed 2nd Floor], [16], [16], [0], [XV-40202E, XV-40202F, XV-40202G (+13 more)],
  [*C1*], [09], [1756-OB32], [DO], [*JB-401*], [Infeed 2nd Floor], [16], [8], [8], [XV-40201E, XV-40201F, XV-40201G (+5 more)],
  [*C1*], [09], [1756-OB32], [DO], [*JB-601*], [Spray Dryer 1st Floor], [16], [12], [4], [SV-60201, SV-60202, SV-60230 (+9 more)],
  [*C1*], [10], [1756-IF16], [AI], [*JB-401*], [Infeed 2nd Floor], [16], [11], [5], [FT-40201, FT-40201, FT-40202 (+8 more)],
  [*C1*], [11], [1756-IF16], [AI], [*JB-601*], [Spray Dryer 1st Floor], [16], [8], [8], [DPI-60240, DPT-60202, FT-60201 (+5 more)],
  [*C2*], [00], [1756-EN4TR], [ETH-1], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C2*], [01], [1756-IB32], [DI], [*JB-402*], [Jet Cooker 2nd Floor], [32], [24], [8], [FS-40204, FS-40205, FT-40221 (+21 more)],
  [*C2*], [02], [1756-IB32], [DI], [*JB-602*], [Spray Dryer 3rd Floor], [32], [23], [9], [FT-60202, LSH-60251, PSL-60220 (+20 more)],
  [*C2*], [03], [1756-IB32], [DI], [*JB-607*], [Spray Dryer 7th Floor], [29], [29], [0], [XA-FH-60210A, XA-FH-60210B, XA-FH-60210C (+26 more)],
  [*C2*], [04], [1756-IB32], [DI], [*JB-607*], [Spray Dryer 7th Floor], [32], [32], [0], [XV-60261FC, XV-60261FC, XV-60264SQ (+29 more)],
  [*C2*], [05], [1756-IB32], [DI], [*JB-607*], [Spray Dryer 7th Floor], [32], [20], [12], [XV-60262SH, XV-60262SH, XV-60263 (+17 more)],
  [*C2*], [06], [1756-OB32], [DO], [*JB-402*], [Jet Cooker 2nd Floor], [16], [9], [7], [XV-40203F, XV-40204B, XV-40204C (+6 more)],
  [*C2*], [06], [1756-OB32], [DO], [*JB-602*], [Spray Dryer 3rd Floor], [16], [3], [13], [XS-60240B, XV-60214, XV-60241],
  [*C2*], [07], [1756-OB32], [DO], [*JB-607*], [Spray Dryer 7th Floor], [32], [31], [1], [SV-60207, XS-FH-60210C, XS-FH-60210D (+28 more)],
  [*C2*], [08], [1756-OB32], [DO], [*JB-607*], [Spray Dryer 7th Floor], [32], [9], [23], [XV-60264, XV-60264FN, XV-60264FP (+6 more)],
  [*C2*], [09], [1756-IF16], [AI], [*JB-402*], [Jet Cooker 2nd Floor], [16], [7], [9], [FT-40221, FT-40221, FT-40222 (+4 more)],
  [*C2*], [10], [1756-IF16], [AI], [*JB-602*], [Spray Dryer 3rd Floor], [16], [16], [0], [AT-60241, DPT-60203, FT-60202 (+13 more)],
  [*C3*], [00], [1756-EN4TR], [ETH-1], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C3*], [01], [1756-IB32], [DI], [*JB-618*], [Packing Tower 8th Floor], [11], [7], [4], [DV-61301, DV-61301, DV-61302 (+4 more)],
  [*C3*], [02], [1756-IB32], [DI], [*JB-618*], [Packing Tower 8th Floor], [5], [5], [0], [PSL-61311, SV-61304, SV-61304 (+2 more)],
  [*C3*], [03], [1756-OB32], [DO], [*JB-618*], [Packing Tower 8th Floor], [32], [9], [23], [DV-61301, DV-61302, SV-61301 (+6 more)],
  [*C3*], [05], [1756-IF16], [AI], [*JB-618*], [Packing Tower 8th Floor], [4], [4], [0], [PT-61303, PT-61306, TT-61303 (+1 more)],
  [*C3*], [06], [1756-N2], [AI], [*JB-602*], [Spray Dryer 3rd Floor], [16], [13], [3], [TT-60215B, TT-60216, TT-60240 (+10 more)],
  [*C3*], [07], [1756-N2], [AI], [*JB-607*], [Spray Dryer 7th Floor], [16], [16], [0], [FT-60207, FT-60221, FT-60261FA (+13 more)],
  [*C3*], [08], [1756-N2], [AI], [*JB-607*], [Spray Dryer 7th Floor], [16], [11], [5], [FT-60264FQ, FT-60264FR, PT-60201 (+8 more)],
  [*C3*], [11], [1756-N2], [AO], [*JB-402*], [Jet Cooker 2nd Floor], [1], [1], [0], [TCV-40222],
  [*C3*], [11], [1756-N2], [AO], [*JB-602*], [Spray Dryer 3rd Floor], [2], [2], [0], [TV-60202, TV-60214],
  [*C3*], [11], [1756-N2], [AO], [*JB-607*], [Spray Dryer 7th Floor], [5], [2], [3], [PV-60201, UY-FH-60210B],
  [*C4*], [00], [1756-EN4TR], [ETH-1], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C4*], [01], [1756-IB32], [DI], [*JB-606*], [Spray Dryer 6th Floor], [16], [4], [12], [FT-60203, LSH-60250, XA-60261 (+1 more)],
  [*C4*], [02], [1756-IB32], [DI], [*JB-612*], [Packing Tower 2nd Floor], [32], [17], [15], [PSL-61305, PSL-61306, PVM-61301 (+14 more)],
  [*C4*], [04], [1756-OB32], [DO], [*JB-606*], [Spray Dryer 6th Floor], [8], [1], [7], [SV-60220],
  [*C4*], [04], [1756-OB32], [DO], [*JB-612*], [Packing Tower 2nd Floor], [16], [15], [1], [PVM-61301, PVM-61301, SV-61307 (+12 more)],
  [*C4*], [05], [1756-IF16], [AI], [*JB-606*], [Spray Dryer 6th Floor], [2], [2], [0], [DPT-60201, FT-60203],
  [*C4*], [05], [1756-IF16], [AI], [*JB-608*], [Spray Dryer 8th Floor], [6], [1], [5], [TT-60245],
  [*C4*], [05], [1756-IF16], [AI], [*JB-612*], [Packing Tower 2nd Floor], [8], [4], [4], [PT-61302, PT-61309, PT-61302 (+1 more)],
  [*C5*], [00], [1756-EN4TR], [ETH-1], [*CA1*], [Main Control Room], [1], [0], [1], [Reserve / Spare Bank],
  [*C5*], [01], [1756-IB32], [DI], [*RIO-200*], [Slurry Out-Building 2nd Fl], [32], [32], [0], [FS-20201, FT-913201, FT-936201 (+29 more)],
  [*C5*], [02], [1756-IB32], [DI], [*RIO-200*], [Slurry Out-Building 2nd Fl], [32], [8], [24], [XV-20211, XV-20211, XV-20212 (+5 more)],
  [*C5*], [03], [1756-OB32], [DO], [*RIO-200*], [Slurry Out-Building 2nd Fl], [32], [8], [24], [XA-92201, XV-20201, XV-20202 (+5 more)],
  [*C5*], [04], [1756-IF16], [AI], [*RIO-200*], [Slurry Out-Building 2nd Fl], [16], [9], [7], [EXT-BIILER, FT-913201, FT-913201 (+6 more)],
  [*C6*], [05], [DO], [DO], [*MCC*], [MCC Room Ground Floor], [8], [8], [0], [PCM-2030, PCM-2030, PCM-2031 (+5 more)],
  [*C6*], [06], [DI], [DI], [*MCC*], [MCC Room Ground Floor], [10], [10], [0], [PCM-2030, PCM-2030, PCM-2031 (+7 more)],
  [*C7*], [00], [BUS], [BUS], [*JB-601*], [Spray Dryer 1st Floor], [1], [1], [0], [PCY-60205],
  [*C7*], [00], [BUS], [BUS], [*JB-607*], [Spray Dryer 7th Floor], [1], [1], [0], [COM-FH-60210B],
  [*C7*], [00], [BUS], [BUS], [*JB-612*], [Packing Tower 2nd Floor], [1], [1], [0], [Communication Signal],
  [*C7*], [00], [BUS], [BUS], [*MCC*], [MCC Room Ground Floor], [17], [17], [0], [ATY-40201, ATY-40202, ATY-40203 (+14 more)],
  [*C7*], [01], [DI], [DI], [*MCC*], [MCC Room Ground Floor], [32], [24], [8], [ATY-40201, ATY-40201, ATY-40201 (+21 more)],
  [*C7*], [02], [DI], [DI], [*JB-601*], [Spray Dryer 1st Floor], [2], [2], [0], [PCY-60205, PCY-60205],
  [*C7*], [02], [DI], [DI], [*MCC*], [MCC Room Ground Floor], [30], [30], [0], [PCM-60205-4, PCM-60205-5, PCM-60205-5 (+27 more)],
  [*C7*], [03], [DI], [DI], [*JB-601*], [Spray Dryer 1st Floor], [1], [1], [0], [PCY-60205],
  [*C7*], [03], [DI], [DI], [*MCC*], [MCC Room Ground Floor], [25], [25], [0], [RVM-61302, RVM-61302, RVM-61303 (+22 more)],
  [*C7*], [04], [DI], [DI], [*MCC*], [MCC Room Ground Floor], [28], [0], [28], [Reserve / Spare Bank],
  [*C7*], [05], [DO], [DO], [*MCC*], [MCC Room Ground Floor], [32], [24], [8], [ATY-40201, ATY-40202, ATY-40203 (+21 more)],
  [*C7*], [06], [DO], [DO], [*JB-601*], [Spray Dryer 1st Floor], [1], [1], [0], [PCY-60205],
  [*C7*], [06], [DO], [DO], [*MCC*], [MCC Room Ground Floor], [31], [31], [0], [BFY-60240, PCY-60215, XS-60240A (+28 more)],
  [*C7*], [07], [DO], [DO], [*MCC*], [MCC Room Ground Floor], [24], [0], [24], [Reserve / Spare Bank],
  [*C8*], [01], [DI], [DI], [*IS-JB-603*], [Spray Dryer 3rd Floor (IS)], [16], [15], [1], [AT-40201, AT-40202, AT-40203 (+12 more)],
  [*C8*], [01], [DI], [DI], [*IS-JB-608*], [Spray Dryer 8th Floor (IS)], [16], [7], [9], [ZS-60201A, ZS-60201B, ZS-60231C (+4 more)],
  [*C8*], [02], [DI], [DI], [*IS-JB-612*], [Packing Tower 2nd Fl (IS)], [31], [15], [16], [ISV-61303, ISV-61304, LSH-61304 (+12 more)],
  [*C8*], [03], [DI], [DI], [*IS-JB-618*], [Packing Tower 8th Fl (IS)], [32], [13], [19], [ISV-61301, LSH-61301, LSH-61303 (+10 more)],
  [*C8*], [04], [DO], [DO], [*IS-JB-608*], [Spray Dryer 8th Floor (IS)], [8], [1], [7], [XA-60201],
  [*C8*], [05], [AI], [AI], [*IS-JB-603*], [Spray Dryer 3rd Floor (IS)], [16], [7], [9], [PT-60230, PT-60231, TT-60224 (+4 more)],
  [*C8*], [06], [AI], [AI], [*IS-JB-608*], [Spray Dryer 8th Floor (IS)], [8], [6], [2], [PT-60220, TT-60220, TT-60221 (+3 more)],
  [*C8*], [06], [AI], [AI], [*IS-JB-612*], [Packing Tower 2nd Fl (IS)], [10], [2], [8], [DPT-61310, DPT-61312],
  [*C8*], [07], [AI], [AI], [*IS-JB-618*], [Packing Tower 8th Fl (IS)], [12], [9], [3], [DPT-61301, DPT-61302, LT-61301 (+6 more)],
)
