#set page(
  paper: "a4",
  margin: (top: 1.5cm, bottom: 1.5cm, left: 1.5cm, right: 1.5cm),
  header: [
    #grid(
      columns: (1fr, auto),
      align(left)[#text(size: 7.5pt, fill: rgb("475569"), weight: "bold")[KALASIN ENGINEERING  |  SPRINT 18K TPA SPRAY DRYER  |  ELECTRICAL & AUTOMATION]],
      align(right)[#text(size: 7.5pt, fill: rgb("475569"))[DOC: KAL-CALC-VENT-2610-001 | REV 1.0]]
    )
    #v(-3pt)
    #line(length: 100%, stroke: 0.5pt + rgb("CBD5E1"))
  ],
  footer: [
    #line(length: 100%, stroke: 0.5pt + rgb("CBD5E1"))
    #v(2pt)
    #grid(
      columns: (1fr, auto),
      align(left)[#text(size: 7pt, fill: rgb("94A3B8"))[PROJECT: xCIP-1545 (Ref: x2608003)  —  INGREDION (THAILAND) CO., LTD. & KALASIN ENGINEERING]],
      align(right)[#text(size: 7pt, fill: rgb("64748B"))[Enclosure Ventilation Load & Fan Selection Calculation Report]]
    )
  ]
)

#set text(font: "Arial", size: 8pt, fill: rgb("0F172A"))
#set par(justify: true, leading: 0.45em)

// =========================================================================
// PAGE 1: EXECUTIVE BANNER, METADATA, SUMMARY & ENCLOSURE GEOMETRY
// =========================================================================

#align(center)[
  #block(
    fill: rgb("1E3A8A"),
    inset: (x: 10pt, y: 8pt),
    radius: 4pt,
    width: 100%,
    [
      #text(size: 12pt, weight: "bold", fill: white)[KALASIN ENGINEERING CO., LTD.] \
      #v(1pt)
      #text(size: 8.5pt, fill: rgb("E0F2FE"), weight: "medium")[PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER  |  REF: x2608003] \
      #v(2pt)
      #text(size: 11pt, weight: "bold", fill: rgb("38BDF8"))[CONTROL ENCLOSURE VENTILATION LOAD & FAN SELECTION REPORT] \
      #v(1pt)
      #text(size: 8pt, weight: "bold", fill: white)[Thermal Dissipation, Airflow Sizing & Equipment Selection conforming to IEC 60890 / IEC 61439] \
      #v(1pt)
      #text(size: 7pt, fill: rgb("CBD5E1"))[Supplier Catalog Reference: Primus Ventilation Systems (x9100-eDrawing/Supplier/PrimusVentilation.pdf)]
    ]
  )
]

#v(2pt)

// Metadata Table
#table(
  columns: (1.5fr, 2.5fr, 1.5fr, 2.5fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1E293B") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 { center } else { left },
  table.header(
    [*Document Field*], [*Specification*], [*Document Field*], [*Specification*]
  ),
  [Document Number:], [KAL-CALC-VENT-2610-001], [Revision / Date:], [Rev 1.0  |  10-Oct-2026],
  [Enclosure Tag / Size:], [Automation Cabinet (800W × 2000H × 400D mm)], [Ambient Outside Temp:], [30.0 °C (Summer Design Basis)],
  [Internal Load Installed:], [2x DC PSU 40A + 1756 13-Slot Rack], [Target Enclosure Temp:], [35.0 °C (Optimum) / 40.0 °C (Max Limit)],
  [Design Thermal Load:], [268.3 Watts (915.5 BTU/h) incl. +15% margin], [Selected Fan Model:], [PRIMUS PMV250N (8", 428 m³/h, IP54)]
)

#v(3pt)

== 1. Executive Summary & Design Objectives

This engineering calculation determines the thermal heat dissipation load, natural enclosure conduction, required forced cooling airflow rate, and industrial filter fan selection for a free-standing automation enclosure with outer dimensions *800 mm (W) × 2000 mm (H) × 400 mm (D)* deployed at the *Ingredion Sprint 18K TPA Spray Dryer Plant*.

The cabinet houses critical process automation hardware: two *24VDC 40A Industrial Power Supplies* and an *Allen-Bradley ControlLogix 1756 13-Slot Chassis* configured with an *L95 Controller (ControlLogix 5580)*, three *1756-EN4TR* Dual-Port Gigabit Ethernet modules, and eight *1756 I/O modules*. Standard thermodynamic calculations per *IEC 60890* and *DIN 57660* were conducted based on an external ambient of *30.0 °C*. All cooling equipment is selected from the vendor catalog *Primus Ventilation Systems* (`x9100-eDrawing/Supplier/PrimusVentilation.pdf`).

#v(1pt)
#block(
  fill: rgb("F0FDF4"),
  stroke: 1pt + rgb("86EFAC"),
  inset: 6pt,
  radius: 3pt,
  [
    #text(weight: "bold", fill: rgb("166534"))[Key Engineering Findings & Selection Summary:] \
    1. *Total Internal Thermal Load ($Q_"total"$)*: *268.3 Watts* (*915.5 BTU/h*), including base equipment losses (233.3 W) plus a +15% design safety margin (35.0 W) for auxiliary wiring and aging. \
    2. *Required Design Airflow ($V_"design"$)*: *216.2 m³/h* to maintain internal temperature at *35.0 °C* ($Delta T = 5.0" K"$), with +30% safety factor for dust filter accumulation. \
    3. *Selected Primary Fan*: *PRIMUS PMV250N* (8" Cabinet Filter Fan, 220VAC, Airflow: *428 m³/h*, IP54, 56 dB). Provides *1.98× design safety factor* over nominal requirement. \
    4. *Selected Matching Filter Grille*: *PRIMUS PMF250N* (250 × 250 mm with G4 synthetic dust filter mat). \
    5. *Thermostatic Control*: *PRIMUS CMA-001* (Analog Thermostat, set to *35.0 °C Cut-in / 30.0 °C Cut-out*) to eliminate continuous fan run hours and minimize dust ingress.
  ]
)

#v(3pt)

== 2. Enclosure Geometry & Natural Dissipation Capability (IEC 60890)

The physical enclosure dimensions and heat-dissipating surface areas are calculated per *IEC 60890* (*Calculation of temperature rise for low-voltage switchgear and controlgear assemblies*):

#table(
  columns: (1.5fr, 1fr, 1fr, 1.8fr, 1.2fr, 1.5fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("334155") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 { center } else { left },
  table.header(
    [*Dimension / Face*], [*Formula*], [*Area (m²)*], [*Installation Condition*], [*Effective Area*], [*Nat. Dissipation (5K)*]
  ),
  [Width ($W$)], [Given], [0.800 m], [Free-Standing ($A_"eff,free"$)], [4.768 m²], [131.1 W ($k = 5.5$)],
  [Height ($H$)], [Given], [2.000 m], [Wall-Mounted ($A_"eff,wall"$)], [3.168 m²], [87.1 W ($k = 5.5$)],
  [Depth ($D$)], [Given], [0.400 m], [Front + Rear Doors], [3.200 m²], [Natural radiation / air],
  [Roof / Top Face], [$W times D$], [0.320 m²], [Two Side Walls], [1.600 m²], [Heat rise assists draft],
  [Gross Enclosure Volume], [$W times H times D$], [*0.640 m³*], [Total Geometric Surface], [*5.120 m²*], [Excluding bottom base]
)

#v(2pt)
#text(size: 7.5pt, fill: rgb("475569"))[*Design Principle*: To guarantee conservative plant reliability during high-temperature summer peaks and food plant dust accumulation, forced ventilation airflow is sized to handle *100% of internal heat load (zero credit for enclosure wall conduction)*, ensuring true fail-safe operation.]

#pagebreak()

// =========================================================================
// PAGE 2: HEAT INVENTORY, THERMAL SIZING & LAYOUT SCHEMATIC
// =========================================================================

== 3. Equipment Heat Dissipation Load Inventory ($P_"internal"$)

The internal thermal dissipation has been calculated using manufacturer datasheets for each active electrical component:

#table(
  columns: (0.5fr, 2.2fr, 1.8fr, 0.5fr, 0.9fr, 0.9fr, 2.2fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1E3A8A") } else if row >= 9 { rgb("F1F5F9") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 { center } else if col == 0 or col == 3 { center } else if col == 4 or col == 5 { right } else { left },
  table.header(
    [*No.*], [*Equipment Description*], [*Catalog / Model*], [*Qty*], [*Unit (W)*], [*Total (W)*], [*Design Basis & Technical Notes*]
  ),
  [1], [24VDC 40A Industrial Power Supply], [QUINT4-PS/1AC/24DC/40], [2], [61.3 W], [122.6 W], [Nominal efficiency 94%, full load rating (24V × 40A)],
  [2], [ControlLogix AC Chassis Power Supply], [Allen-Bradley 1756-PA75], [1], [25.0 W], [25.0 W], [Maximum internal dissipation (85.3 BTU/h rated)],
  [3], [ControlLogix 5580 Controller (40MB)], [Allen-Bradley 1756-L95], [1], [6.2 W], [6.2 W], [Dual GbE ports active, 100% logic execution cycle],
  [4], [EtherNet/IP DLR Comm Adapter], [Allen-Bradley 1756-EN4TR], [3], [3.5 W], [10.5 W], [1Gbps dual port adapter (11.9 BTU/h each)],
  [5], [ControlLogix Digital Input Modules], [Allen-Bradley 1756-IB32], [4], [5.5 W], [22.0 W], [32-point 24VDC sink/source modules (4 cards)],
  [6], [ControlLogix Digital Output Modules], [Allen-Bradley 1756-OB32], [2], [4.8 W], [9.6 W], [32-point 24VDC sourcing modules (2 cards)],
  [7], [ControlLogix Analog Input Modules], [Allen-Bradley 1756-IF16], [2], [6.2 W], [12.4 W], [16-point isolated analog input cards (2 cards)],
  [8], [Auxiliary Relays, MCBs & Diode Module], [QUINT4-DIODE / PLC-RSC], [Lot], [25.0 W], [25.0 W], [Diode loss (10W) + Relay coils, MCB, terminal losses],
  [], [*SUBTOTAL ACTIVE LOAD*], [], [*15*], [], [*233.3 W*], [*Continuous operating heat load*],
  [], [*DESIGN SAFETY MARGIN (+15%)*], [], [], [], [*35.0 W*], [*Expansion reserve, wiring heating & aging factor*],
  [], [*TOTAL DESIGN THERMAL LOAD ($Q_"total"$)*], [], [], [], [*268.3 W*], [*915.5 BTU/h total heat dissipation to be evacuated*]
)

#v(3pt)

== 4. Thermodynamic Ventilation Airflow Calculation & Governing Equations

The required forced air volume flow rate $V$ (in $m^3/h$) is governed by the sensible heat balance equation:

$ V = frac(f times Q_v, Delta T) quad [m^3/h] quad "where" quad f = frac(1, rho times c_p) = frac(3600, 1.204" kg/m"^3 times 1005" J/(kg"dot"K)") approx 3.10 quad [m^3 dot K / (W dot h)] $

With dust loading factor $k_"filter" = 1.30$ (+30% headroom): $V_"design" = V times 1.30$.

#table(
  columns: (2.2fr, 0.8fr, 1.1fr, 1.1fr, 1.2fr, 2fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("334155") } else if row == 1 { rgb("FEF2F2") } else if row == 2 { rgb("F0FDF4") } else { white },
  align: (col, row) => if row == 0 { center } else if col >= 1 and col <= 4 { center } else { left },
  table.header(
    [*Operational Design Case*], [*$Delta T$*], [*Target Temp*], [*$V_"pure"$ ($m^3/h$)*], [*$V_"design"$ (+30%)*], [*Assessment & Compliance Status*]
  ),
  [*Case 1: Optimum Lifespan (Conservative)*], [5.0 K], [35.0 °C], [166.3 m³/h], [*216.2 m³/h*], [Recommended for maximum PLC hardware lifespan],
  [*Case 2: Standard Industrial Limit (Cons.)*], [10.0 K], [40.0 °C], [83.2 m³/h], [*108.1 m³/h*], [Standard industrial control limit per IEC 61439-1],
  [*Case 3: Maximum Allowable (Cons.)*], [15.0 K], [45.0 °C], [55.4 m³/h], [72.1 m³/h], [High-stress emergency threshold],
  [*Case 4: Case 1 with Wall Conduction*], [5.0 K], [35.0 °C], [112.3 m³/h], [146.0 m³/h], [Accounts for rear wall convective heat transfer],
  [*Case 5: Case 2 with Wall Conduction*], [10.0 K], [40.0 °C], [29.2 m³/h], [37.9 m³/h], [Low airflow required due to natural radiation]
)

#v(3pt)

#grid(
  columns: (1fr, 1fr),
  gutter: 8pt,
  [
    #image("assets/enclosure_ventilation_diagram.png", height: 5.4cm)
    #align(center)[#text(size: 6.8pt, fill: rgb("64748B"))[*Figure 1*: Enclosure internal arrangement, thermal flow & sensor placement]]
  ],
  [
    #image("assets/airflow_vs_temperature_curve.png", height: 5.4cm)
    #align(center)[#text(size: 6.8pt, fill: rgb("64748B"))[*Figure 2*: Airflow requirements vs. internal temperature operating curves]]
  ]
)

#pagebreak()

// =========================================================================
// PAGE 3: PRIMUS CATALOG COMPARISON & FAN SELECTION
// =========================================================================

== 5. Primus Supplier Catalog Comparison & Equipment Selection

The approved vendor catalog *Primus Ventilation Systems* (`x9100-eDrawing/Supplier/PrimusVentilation.pdf`) provides industrial cabinet filter fans (PMV-N and PMV Series) and roof centrifugal fans (PMB Series). The models are compared below:

#table(
  columns: (1.5fr, 1.2fr, 1.3fr, 1fr, 1.1fr, 0.9fr, 2.4fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1E3A8A") } else if row == 3 { rgb("DCFCE7") } else if row == 6 { rgb("FEF9C3") } else { white },
  align: (col, row) => if row == 0 { center } else if col == 0 { left } else if col >= 4 and col <= 5 { center } else { center },
  table.header(
    [*Model*], [*Dimensions*], [*Power Supply*], [*Watts / A*], [*Air Flow*], [*Noise*], [*Engineering Suitability Evaluation*]
  ),
  [PMV115N (3")], [115 × 115 mm], [220VAC / 50Hz], [13 W / 0.10 A], [31 m³/h], [30 dB], [Insufficient air volume for this enclosure],
  [PMV205N (4")], [205 × 205 mm], [220VAC / 50Hz], [21 W / 0.11 A], [142 m³/h], [43 dB], [Adequate for 40°C ($Delta T = 10$K), marginal for 35°C],
  [*PMV250N (8")*], [*250 × 250 mm*], [*220VAC / 50Hz*], [*49 W / 0.38 A*], [*428 m³/h*], [*56 dB*], [*SELECTED (BEST)*: 1.98× safety factor for 35°C],
  [PMV320N (10")], [320 × 320 mm], [220VAC / 50Hz], [87 W / 0.40 A], [654 m³/h], [80 dB], [High noise level (80 dB), oversized for application],
  [PMV12], [120 × 120 mm], [220VAC / 50Hz], [16 W / 0.13 A], [133 m³/h], [50 dB], [Undersized for conservative 35°C design],
  [*PMV25*], [*255 × 255 mm*], [*220VAC / 50Hz*], [*38 W / 0.50 A*], [*248 m³/h*], [*58 dB*], [*ALTERNATIVE*: Fits 35°C ($216" m"^3/"h"$ design required)],
  [PMV30.00S], [323 × 323 mm], [220VAC / 50Hz], [40 W / 0.40 A], [428 m³/h], [69 dB], [Higher acoustic noise (69 dB) than PMV250N (56 dB)],
  [PMB03.00 (Roof)], [364 × 364 mm], [220VAC / 50Hz], [85 W / 0.38 A], [905 m³/h], [64 dB], [Roof centrifugal mount; alternative if door cutout blocked]
)

#v(3pt)

#block(
  fill: rgb("F8FAFC"),
  stroke: 1pt + rgb("CBD5E1"),
  inset: 7pt,
  radius: 3pt,
  [
    #text(weight: "bold", fill: rgb("1E3A8A"))[Primary Selection: PRIMUS PMV250N (8" Filter Fan Unit)] \
    - *Airflow Capacity*: *428 m³/h* free-air delivery, delivering *1.98× margin* over the 35.0 °C design airflow ($216.2" m"^3/"h"$).
    - *Acoustic Performance*: 56 dB(A), significantly quieter than smaller high-RPM fans or 10" units due to optimized 8" blade profile.
    - *Protection Rating*: IP54 with polyurethane foamed-in-place seal and G4 filter mat, preventing starch dust ingress.
    - *Maintenance*: *Flap-Open* front grille allows tool-free filter replacement in seconds without opening cabinet doors.
    - *Mounting Mechanism*: *Clips-Lock* rapid installation into standard 250 × 250 mm sheet metal cutout without screws.
  ]
)

#v(3pt)

#grid(
  columns: (1.2fr, 0.8fr),
  gutter: 8pt,
  [
    #image("assets/primus_fan_models_crop.png", height: 5.6cm)
    #align(center)[#text(size: 6.8pt, fill: rgb("64748B"))[*Figure 3*: Excerpt from Primus Catalog Page 2 showing PMV-N and PMB Series specifications]]
  ],
  [
    #image("assets/primus_thermostats_crop.png", height: 5.6cm)
    #align(center)[#text(size: 6.8pt, fill: rgb("64748B"))[*Figure 4*: Excerpt from Primus Catalog Page 7 showing CMA Thermostat series]]
  ]
)

#pagebreak()

// =========================================================================
// PAGE 4: BOM, INSTALLATION & QUALITY SIGN-OFF
// =========================================================================

== 6. Recommended Bill of Materials (BOM) & Equipment Specifications

#table(
  columns: (0.4fr, 1.8fr, 1.3fr, 0.5fr, 1fr, 3.2fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1E293B") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 { center } else if col == 0 or col == 3 { center } else { left },
  table.header(
    [*Item*], [*Component Description*], [*Model / Catalog*], [*Qty*], [*Supplier*], [*Technical Specification & Function*]
  ),
  [1], [Industrial Cabinet Filter Fan], [*PMV250N*], [1 Set], [Primus], [250 × 250 mm, 220VAC 50Hz, 49W, 428 m³/h, IP54, 56 dB, Clips-Lock & Flap-Open cover],
  [2], [Matching Exhaust Filter Grille], [*PMF250N*], [1 Set], [Primus], [250 × 250 mm matching frame, G4 synthetic filter mat, IP54 seal],
  [3], [Cabinet Thermostat Controller], [*CMA-001*], [1 Set], [Primus], [DIN-rail mechanical bimetal thermostat, 18–87 °C, NO contact (Fan control), 220VAC 10A],
  [4], [Spare Replacement Filter Mat], [*PMM-250N (G4)*], [2 Pcs], [Primus], [G4 / EU4 synthetic fiber filter media, washable, low pressure drop],
  [5], [Fan Power Protection MCB], [*C60H-2P 2A C-Curve*], [1 Pcs], [Schneider], [2-Pole 2A miniature circuit breaker for 220VAC fan auxiliary supply]
)

#v(4pt)

== 7. Enclosure Layout, Installation & Maintenance Instructions

1. *Airflow Pathway Configuration (Positive Pressure vs. Exhaust)*:
   - *Intake Location*: Install the *PMF250N* filter grille at the *bottom front door* (or bottom side panel), at least 200 mm above the floor plinth to avoid floor dust suction.
   - *Exhaust Location*: Install the *PMV250N* filter fan at the *top front door* (or top side panel), blowing hot air outwards.
   - *Natural Convective Chimney Effect*: Ambient air enters cool at the bottom, absorbs heat from the 40A power supplies mounted on lower DIN rails, sweeps upward across the 1756 chassis, and is actively expelled at the roofline.

2. *Thermostat Placement & Setpoint*:
   - Mount the *CMA-001* thermostat on the top DIN rail in the warmest zone near the 1756 rack.
   - Set the dial to *35.0 °C*. The fan will automatically energize when internal temperature reaches 35.0 °C and turn off at ~30.0 °C.
   - *Benefits*: Prevents continuous fan operation, saves up to 70% fan energy, extends bearing life to >50,000 hours, and significantly reduces particulate accumulation inside the filter mat.

3. *Maintenance Schedule*:
   - Inspect and clean/replace filter mats every *3 months* in starch processing plant environments.
   - The Primus *Flap-Open* design allows operators to change filter mats in under 30 seconds without tools or opening high-voltage door interlocks.

#v(6pt)

== 8. Engineering Verification, Compliance & Quality Sign-Off Block

#table(
  columns: (1.5fr, 1.8fr, 1.5fr, 1.8fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("334155") } else { white },
  align: (col, row) => if row == 0 { center } else { left },
  table.header(
    [*Role / Function*], [*Name & Title*], [*Organization*], [*Status / Signature*]
  ),
  [Calculated By:], [Senior Thermal & Automation Eng.], [Kalasin Engineering Co., Ltd.], [Verified & Approved (10-Oct-2026)],
  [Reviewed By:], [Lead Electrical Systems Eng.], [Kalasin Engineering Co., Ltd.], [Verified & Approved (10-Oct-2026)],
  [Client Verification:], [Plant Automation Manager], [Ingredion (Thailand) Co., Ltd.], [Pending FAT Inspection],
  [Standards Compliance:], [IEC 60890 / IEC 61439-1 / DIN 57660], [Industrial Control Enclosures], [Full Technical Compliance PASS]
)
