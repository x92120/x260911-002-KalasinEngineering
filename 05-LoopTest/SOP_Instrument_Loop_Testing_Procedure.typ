#set page(
  paper: "a4",
  margin: (top: 2.5cm, bottom: 2.5cm, left: 2.2cm, right: 2.2cm),
  header: [
    #grid(
      columns: (1fr, auto),
      align(left)[#text(size: 8pt, fill: rgb("64748B"), weight: "bold")[KALASIN ENGINEERING  |  PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER]],
      align(right)[#text(size: 8pt, fill: rgb("64748B"))[DOC: KAL-SOP-INS-0545 | REV 1.0]]
    )
    #v(-3pt)
    #line(length: 100%, stroke: 0.5pt + rgb("CBD5E1"))
  ],
  footer: [
    #line(length: 100%, stroke: 0.5pt + rgb("CBD5E1"))
    #v(2pt)
    #grid(
      columns: (1fr, auto),
      align(left)[#text(size: 8pt, fill: rgb("94A3B8"))[CONFIDENTIAL  —  INGREDION (THAILAND) CO., LTD. & KALASIN ENGINEERING]],
      align(right)[#text(size: 8pt, fill: rgb("64748B"))[IEC 62381 / ISA-RP60.8 Standards]]
    )
  ]
)

#set text(size: 9.5pt, fill: rgb("0F172A"))
#set par(justify: true, leading: 0.6em)

// Cover Header
#align(center)[
  #block(
    fill: rgb("1B365D"),
    inset: 15pt,
    radius: 4pt,
    width: 100%,
    [
      #text(size: 14pt, weight: "bold", fill: white)[KALASIN ENGINEERING CO., LTD.] \
      #v(3pt)
      #text(size: 9.5pt, fill: rgb("E0F2FE"), weight: "medium")[PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER  |  REF: x2608003] \
      #v(5pt)
      #text(size: 12pt, weight: "bold", fill: rgb("38BDF8"))[STANDARD OPERATING PROCEDURE (SOP)] \
      #v(2pt)
      #text(size: 11pt, weight: "bold", fill: white)[INSTRUMENT LOOP TESTING & FIELD COMMISSIONING METHODOLOGY] \
      #v(3pt)
      #text(size: 8pt, fill: rgb("CBD5E1"))[Conforming to IEC 62381, ISA-RP60.8, and NFPA 79 Standards]
    ]
  )
]

#v(6pt)

// Metadata Summary Table
#table(
  columns: (1.5fr, 2.5fr, 1.5fr, 2.5fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("2D4A77") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 { center } else { left },
  table.header(
    [*Document Field*], [*Specification*], [*Document Field*], [*Specification*]
  ),
  [Document Reference:], [KAL-SOP-INS-0545], [Issue Date:], [07-October-2026],
  [Revision:], [Rev. 1.0 (Approved for Execution)], [Target Facility:], [Ingredion Kalasin Facility],
  [Lead C&I Contractor:], [Kalasin Engineering Co., Ltd.], [Client Organization:], [Ingredion (Thailand) Co., Ltd.],
  [Applicable Enclosures:], [15 Junction Boxes (JB-401 to JB-618, IS-JBs)], [Total Scope:], [395 Equipment / 620 Channels]
)

#v(8pt)

== 1. PURPOSE & APPLICABLE STANDARDS

This Standard Operating Procedure (SOP) defines the mandatory technical requirements, safety precautions, test connection topologies, and acceptance criteria for performing instrument loop checks and pre-commissioning verification.
All tests must strictly comply with:
- *IEC 62381*: Process Industry — Automation Systems — Activities during factory acceptance testing, site acceptance testing, and site integration.
- *ISA-RP60.8*: Hardware Guide for Control Centers & Instrument Loop Verification.
- *IEC 60079-14*: Explosive Atmospheres — Electrical installations design, selection, and erection (for Intrinsically Safe loops in IS-JB-603, 608, 612, 618).
- *NFPA 79*: Electrical Standard for Industrial Machinery.

#v(8pt)

== 2. PRE-COMMISSIONING COLD LOOP TEST METHODS

#block(
  fill: rgb("FEF2F2"),
  stroke: 1pt + rgb("EF4444"),
  inset: 9pt,
  radius: 4pt,
  [
    #text(weight: "bold", fill: rgb("B91C1C"))[CRITICAL SAFETY PRECAUTION: SENSITIVE ELECTRONICS ISOLATION] \
    #text(size: 8.5pt, fill: rgb("7F1D1D"))[
      High-voltage Megger insulation testing (500 VDC) *WILL PERMANENTLY DESTROY* solid-state instrument electronics, SMART positioners, and Allen-Bradley ControlLogix 1756 / 1794 FLEX I/O modules. Both cores (+ and -) *MUST BE FULLY DISCONNECTED* at the transmitter terminal head and at the CA1 marshalling rack before pressing the Megger test button.
    ]
  ]
)

#v(6pt)

=== 2.1 Cable Insulation Resistance Testing (Megger 500 VDC)
*Objective*: Verify the dielectric integrity of cable insulation to prove zero leakage paths exist between conductors, shields, and ground.
- *Test Instrument*: Calibrated 500 VDC Insulation Tester (e.g., Megger MIT410/2 or Fluke 1507). Calibration certificate must be valid within 12 months.
- *Test Duration*: 60 seconds continuous test per measurement point.
- *Acceptance Criterion*: *#text(weight: "bold")[≥ 20 MΩ @ 500 VDC]* (Typical healthy new cabling reads > 100 MΩ).

*Test Execution Matrix (4 Mandatory Test Points)*:
1. *Core-to-Core (A <-> B)*: Apply 500 VDC between Conductor A (+) and Conductor B (-). Minimum 20 MΩ.
2. *Core-to-Shield (A <-> SH & B <-> SH)*: Apply 500 VDC between each signal conductor and the cable shield drain wire. Minimum 20 MΩ.
3. *Core-to-Earth (A <-> PE & B <-> PE)*: Apply 500 VDC between each signal conductor and the plant Protective Earth (PE). Minimum 20 MΩ.
4. *Shield-to-Earth (SH <-> PE)*: Apply 500 VDC between the cable shield drain wire and Protective Earth. Minimum 20 MΩ (proves shield is floating at field/JB).
5. *Post-Test Discharge*: Allow the tester to automatically bleed off cable capacitive charge for at least 15 seconds before touching any conductor.

#v(6pt)

=== 2.2 Loop Conductor Resistance Testing (Continuity Check)
*Objective*: Confirm complete end-to-end electrical path, verify that no high-resistance cold-solder or loose screw connections exist, and confirm correct polarity.
- *Test Instrument*: Digital Multimeter (Fluke 87V or Fluke 789 ProcessMeter) set to Low-Ohm resistance mode.
- *Acceptance Criterion*: *#text(weight: "bold")[≤ 2.0 Ω]* total round-trip loop resistance (typical for 0.75 mm² to 1.5 mm² cable runs up to 150 meters).

*Step-by-Step Test Procedure*:
1. *Lead Nulling*: Touch meter test leads together and engage the `REL / ZERO` function to null out test lead internal resistance (0.00 Ω).
2. *Field Shorting*: At the field device terminal strip, attach an insulated temporary shorting bridge securely across Terminal 1 (+) and Terminal 2 (-).
3. *Resistance Measurement*: At the CA1 marshalling rack, measure resistance across the corresponding incoming trunk terminal pair. Record measured ohms.
4. *Open-Circuit Check*: Remove the field shorting bridge. The meter at CA1 must immediately transition to *O.L (Open Loop / Infinite Ω)*. If any finite resistance is observed, an unwanted parallel path or short circuit exists and must be repaired.
5. *Polarity Verification*: Apply a known +9.0 VDC battery across Core A (+) and Core B (-) at CA1. Measure DC voltage at the field device. Verify positive polarity on Core A.

#v(6pt)

=== 2.3 Single-Point Earth Shielding ("CUTBACK & TAPE" Rule)
*Objective*: Prevent low-frequency 50 Hz AC ground hum, radio-frequency interference (RFI), and circulating eddy currents caused by differences in earth potential across the facility.
- *Acceptance Criteria*:
  - Shield-to-Ground resistance at Field Instrument: *#text(weight: "bold")[Infinite (O.L / > 20 MΩ)]* (Completely floating).
  - Shield-to-Ground resistance at Junction Box: *#text(weight: "bold")[Infinite (O.L / > 20 MΩ)]* (Isolated from metallic chassis).
  - Shield-to-IE Busbar at CA1 Marshalling: *#text(weight: "bold")[< 0.1 Ω]* (Bonded cleanly).
  - Instrument Clean Earth (IE) to Ground Pit: *#text(weight: "bold")[< 1.0 Ω]*.

*Execution Rules*:
- *At Field Device*: Strip the outer cable jacket. Cut the aluminum foil shield and bare tinned copper drain wire flush with the sheath base. Encase the cut end in a heat-shrink insulating sleeve and wrap securely with self-amalgamating tape. *The shield must never contact the metallic instrument enclosure.*
- *At Junction Box*: Connect the incoming branch shield drain and the outgoing trunk multi-pair shield drain through an *insulated feedthrough terminal (`SH`)*. Keep completely isolated from the mounting DIN rail and enclosure chassis.
- *At CA1 Marshalling Cabinet*: Land all trunk shield drain wires directly onto the dedicated, insulated copper *Instrument Earth (IE)* busbar. This represents the single, centralized earthing reference for the entire control loop.

#v(6pt)

=== 2.4 Enclosure, Cable Gland & Tag Physical Inspection
*Objective*: Guarantee mechanical robustness, environmental sealing, and long-term operating reliability.

#table(
  columns: (2fr, 2.5fr, 3.5fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("2D4A77") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  table.header([*Inspection Item*], [*Standard / Acceptance Spec*], [*Verification Method*]),
  [Ingress Protection (IP66/67)], [Neoprene gasket intact, continuous, unpinched], [Visual check; verify cover bolts torqued to 3.5 N·m],
  [Cable Glands], [Certified Ex-d / Ex-e nickel-plated brass or SS316], [Hand tightness check; verify elastomeric seal compressed],
  [Breather / Drain Plug], [SS316 certified drain plug at lowest enclosure point], [Verify downward orientation for condensate drainage],
  [Stainless Steel Tag Plate], [SS316 engraved plate (Tag, P&ID, Service)], [Cross-reference with P&ID; verify secured with SS wire],
  [Wire Terminations], [Insulated bootlace ferrules on all stranded wires], [Pull test (gentle tug); check torque (0.5–0.6 N·m)],
  [Spare Cable Cores], [Fitted with insulated end ferrules, labeled 'SPARE'], [Neatly coiled in lower duct; landed on spare terminals]
)

#v(10pt)

== 3. LIVE HOT LOOP & CALIBRATION METHODOLOGY

=== 3.1 Analog Input (AI) 2-Pin Differential Pair Calibration
In the Allen-Bradley 1756-IF16 current input architecture, each analog channel utilizes *two physical pins*:
- *Pin A*: IN-x (Signal current input positive, landed on terminal strip `(A)`)
- *Pin B*: i RTN-x (Signal current return negative, landed on terminal strip `(B)`)

*5-Point Simulated Signal Injection Table*:
Inject calibrated current using a Fluke 789 / Fluke 754 process calibrator directly across Pin A and Pin B at the field junction box:

#table(
  columns: (1.5fr, 1.5fr, 2fr, 2fr, 1.2fr, 1.5fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1B365D") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => center,
  table.header([*Test Point*], [*Simulated Input*], [*PLC Raw Counts*], [*SCADA Display Value*], [*Max Error*], [*Result*]),
  [0% (Zero)], [4.000 mA], [0 counts], [0.0 % (LRV)], [± 0.20%], [PASS],
  [25% (Span 1/4)], [8.000 mA], [8192 counts], [25.0 % EU], [± 0.20%], [PASS],
  [50% (Mid)], [12.000 mA], [16384 counts], [50.0 % EU], [± 0.20%], [PASS],
  [75% (Span 3/4)], [16.000 mA], [24576 counts], [75.0 % EU], [± 0.20%], [PASS],
  [100% (Full Span)], [20.000 mA], [32767 counts], [100.0 % (URV)], [± 0.20%], [PASS]
)

#v(6pt)

=== 3.2 Discrete Valve & Multi-Signal Equipment Testing
For actuated on-off valves (`XV-40201A`, `HV`, `SOV`) and flowmeters (`FT-40201`):
1. *Solenoid Actuation (DO)*: Force DO bit in ControlLogix. Verify 24 VDC energizes the solenoid coil, pilot valve shifts cleanly, and valve strokes within specification (< 1.5 sec).
2. *Open / Close Limit Switches (DI)*:
   - When valve reaches 100% full open: Proximity target hit, Dry Contact closes, PLC `DI_1` bit changes 0 -> 1, SCADA icon turns steady *GREEN*.
   - When valve de-energizes to 0% full closed: Spring returns valve to fail-safe seat, PLC `DI_2` bit changes 0 -> 1, SCADA icon turns steady *RED*.
3. *Flowmeter Aux Power & Pulse Counters*: Verify 24 VDC / 220 VAC auxiliary instrument power bus voltage. Inject 100 Hz pulse train into high-speed digital input (`DI_1`). Verify totalizer increments accurately on SCADA.

#v(10pt)

== 4. TRIPARTITE COMMISSIONING ACCEPTANCE SIGN-OFF

Upon completion of all cold and hot loop checks:
1. *Green Loop Test Tag*: Attach a weather-proof green inspection tag to the instrument cable gland marked with Date, Loop Pass Status, and Inspector Initials.
2. *Certificate Archiving*: Sign and date the individual equipment sheet inside the corresponding Junction Box Workbook (`Loop_Test_Report_JB-xxx.xlsx`).
3. *Tripartite Signatories*:
   - *Tested By*: Lead C&I Engineer, Kalasin Engineering Co., Ltd.
   - *Witnessed By*: QA/QC Electrical & Instrumentation Inspector, Ingredion (Thailand) Co., Ltd.
   - *Approved By*: Project Commissioning Manager, Joint Commissioning Team.

