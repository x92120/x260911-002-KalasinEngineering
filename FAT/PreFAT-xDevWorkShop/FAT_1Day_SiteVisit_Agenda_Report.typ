#set page(
  paper: "a4",
  margin: (top: 1.8cm, bottom: 1.8cm, left: 1.6cm, right: 1.6cm),
  header: [
    #grid(
      columns: (1fr, auto),
      align(left)[#text(size: 8pt, fill: rgb("64748B"), weight: "bold")[KALASIN ENGINEERING  |  SPRINT 18K TPA SPRAY DRYER  |  MCP-01 (C1–C4)]],
      align(right)[#text(size: 8pt, fill: rgb("64748B"))[DOC: KAL-FAT-AGN-2608-C1C4 | REV 1.2]]
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
      align(right)[#text(size: 8pt, fill: rgb("64748B"))[Focus: Chassis C1, C2, C3, C4 in MCP-01]]
    )
  ]
)

#set text(size: 8.5pt, fill: rgb("0F172A"))
#set par(justify: true, leading: 0.52em)

// =========================================================================
// PAGE 1: COVER, METADATA & EXECUTIVE CONCEPT
// =========================================================================

#align(center)[
  #block(
    fill: rgb("1B365D"),
    inset: 11pt,
    radius: 4pt,
    width: 100%,
    [
      #text(size: 13pt, weight: "bold", fill: white)[KALASIN ENGINEERING CO., LTD.] \
      #v(2pt)
      #text(size: 9pt, fill: rgb("E0F2FE"), weight: "medium")[PROJECT SPRINT 18K TPA SPRAY DRYER / JET COOKER  |  REF: x2608003] \
      #v(3pt)
      #text(size: 11.5pt, weight: "bold", fill: rgb("38BDF8"))[FACTORY ACCEPTANCE TEST (FAT) — 1-DAY AGENDA REPORT] \
      #v(2pt)
      #text(size: 9pt, weight: "bold", fill: white)["Done on FAT (1 Day)" — Focused on Main Control Panel (MCP-01: Chassis C1, C2, C3, C4)] \
      #v(2pt)
      #text(size: 7.5pt, fill: rgb("CBD5E1"))[Conforming to IEC 62381 / ISA-RP60.8 Standards | Venue: xDev Engineering Workshop]
    ]
  )
]

#v(2pt)

// Metadata Table
#table(
  columns: (1.5fr, 2.5fr, 1.5fr, 2.5fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("2D4A77") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 { center } else { left },
  table.header(
    [*Document Field*], [*Specification*], [*Document Field*], [*Specification*]
  ),
  [Document Reference:], [KAL-FAT-AGN-2608-C1C4], [Issue Date:], [08-October-2026],
  [Revision:], [Rev. 1.2 (C1–C4 + DI/DO Loopback)], [Execution Window:], [09:00 – 17:00 (1 Single Day)],
  [Project Client:], [Ingredion (Thailand) Co., Ltd.], [Venue / Location:], [xDev Workshop Integration Center],
  [Target Enclosure:], [Main Control Panel MCP-01 (TAMCO)], [Target Chassis:], [Chassis C1, C2, C3, C4 (52 Slots)],
  [Total I/O Tested:], [480 DI, 256 DO, 192 AI, 8 AO (936 Total)], [IS Barriers:], [Pepperl+Fuchs K-System (Chassis C4)]
)

#v(3pt)

== 1. C1–C4 Architecture & 1-Day Site Visit Concept

The plant control core resides within *MCP-01*, consisting of *four 13-slot ControlLogix chassis (C1, C2, C3, C4)* linked by a 100Mbps ODVA Device Level Ring (DLR) and dual Cisco Catalyst 9300 Core switches.

#table(
  columns: (0.9fr, 1.8fr, 2.2fr, 1.2fr, 2fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1F4E79") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 or col == 0 or col == 3 { center } else { left },
  table.header([*Chassis*], [*Role / Subsystem*], [*Key Hardware Components*], [*I/O Count*], [*Destination Field Boxes*]),
  [Chassis C1], [*Master CPU & Core I/O*], [1756-L950TPSXT, 3x EN4TR, 4x DI, 2x DO, 2x AI], [128 DI / 64 DO / 32 AI], [Central Area (CA1), JB-401, JB-601],
  [Chassis C2], [*Process Remote I/O*], [1756-EN4TR, 5x DI, 3x DO, 4x AI], [160 DI / 96 DO / 64 AI], [Feed Prep, JB-402, JB-602, JB-607],
  [Chassis C3], [*Extended Process RIO*], [1756-EN4TR, 3x DI, 2x DO, 4x AI, 1x AO (OF8)], [96 DI / 64 DO / 64 AI / 8 AO], [Filter Bags, JB-606, 612, 618],
  [Chassis C4], [*Hazardous Area / IS RIO*], [1756-EN4TR, 3x DI, 1x DO, 2x AI + P&F Barriers], [96 DI / 32 DO / 32 AI (All IS)], [IS-JB-603, 608, 612, 618]
)

#v(4pt)

== 2. Execution Principles for 1-Day Completion
1. *100% Pre-FAT Internal Readiness*: Cold loop check, point-to-point continuity, and loopback simulation are fully completed prior to 09:00.
2. *Representative Audited Verification*: Client witnesses core failover, safety trips, and representative loop samples across C1–C4.
3. *Automated Closed-Loop Cross-Wiring*: Direct dry contact loopback from DO relays to DI terminals enables 100% wiring verification with automated auto-generated test logs (Detailed on Page 3).

#pagebreak()

// =========================================================================
// PAGE 2: MASTER TIMETABLE & SIGN-OFF ENDORSEMENT
// =========================================================================

== 3. Master FAT Timetable (Focus on C1–C4: 09:00 – 17:00)

#table(
  columns: (0.9fr, 0.6fr, 1.9fr, 3.2fr, 1.2fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 {
    rgb("1F4E79")
  } else if row == 4 or row == 7 {
    rgb("FEF3C7")
  } else if calc.even(row) {
    rgb("F8FAFC")
  } else {
    white
  },
  align: (col, row) => if row == 0 or col == 0 or col == 1 { center } else { left },
  table.header(
    [*Time*], [*Duration*], [*Milestone / Agenda Item*], [*Key Activities (Focus on Chassis C1–C4)*], [*Responsible*]
  ),
  [09:00 - 09:15], [15 min], [*Client Arrival & Welcome*], [Client reception, visitor badge, safety induction, ESD protection check before entering MCP-01 test bay.], [Safety Lead / Host],
  [09:15 - 09:30], [*15 min*], [*Detail on FAT (Kick-off)*], [15-minute FAT Scope Briefing: C1–C4 DLR ring architecture, testing boundary, acceptance rules, Cat A vs B punch categorization.], [Project Manager / Lead Automation],
  [09:30 - 10:45], [75 min], [*Station 1: C1 & C2 Core Audit*], [
    • C1 Master CPU (Slot 0: 1756-L950TPSXT) status, Cisco 9300 HSRP failover, Stratix DLR cable break test (recovery \< 3ms). \
    • C1 & C2 DI loop toggle simulation (1756-IB32) via TB-24VDC rail; DO relay forcing (1756-OB32 -> TBRL -> TB-0VDC).
  ], [Lead PLC Eng / Client Inspector],
  [10:45 - 12:00], [75 min], [*Station 2: C1 & C2 Analog Loops*], [
    • C1 (C1S10-S11) & C2 (C2S9-S12) 4-20mA calibration audit (Fluke 789 5-point test: 4, 8, 12, 16, 20mA). \
    • Differential input circuit audit (IN-x & i RTN-x) on 1756-IF16 cards; zero crosstalk & shield grounding verification.
  ], [Instrumentation Lead / Client Rep],
  [12:00 - 13:00], [60 min], [*Lunch Break & Discussion*], [Executive hot luncheon served in private dining room. Mid-day review of morning C1/C2 test logs and workflow alignment.], [All Participants],
  [13:00 - 14:15], [75 min], [*Station 3: C3 (AO) & C4 (P&F IS)*], [
    • Chassis C3 Modulating AO (1756-OF8: 4-20mA) valve simulation; measure current with precision DMM (+/-0.5% FS). \
    • Chassis C4 dedicated Hazardous Area RIO with Pepperl+Fuchs K-System Zener barrier audit & Line Fault wire-break alarm.
  ], [Instrumentation & IS Lead / Client],
  [14:15 - 15:00], [45 min], [*Station 4: SCADA & Safety Trips*], [
    • FactoryTalk View SE SCADA integration for C1-C4 tags, faceplates, alarms, and historical trends. \
    • Critical safety interlocks: Chamber high-temp trip, burner gas lockout, and Master E-Stop de-energization (\< 50ms).
  ], [SCADA Lead / Process Eng],
  [15:00 - 15:30], [*30 min*], [*Coffee Break & Consolidation*], [Afternoon coffee break, pastries and refreshments. Engineering team consolidates C1-C4 test sheets into preliminary Punch List.], [Hospitality / Eng Team],
  [15:30 - 16:30], [60 min], [*Summary Meeting & Comments*], [Presentation of C1-C4 test results & 100% pass rate. Detailed review of Client Comments. Categorize Punch List (Cat A vs B).], [PM, Leads & Client Delegation],
  [16:30 - 17:00], [30 min], [*FAT Sign-off & Departure*], [Formal signing of FAT Certificate of Acceptance for MCP-01 (C1-C4). Review packing protection, dispatch plan & departure.], [Signatories & Directors]
)

#v(4pt)

== 4. Punch List Protocol & Acceptance Criteria (C1–C4 Focus)

- *Category A (Critical / Blocker)*: Safety hazard, hardware damage risk, or core architecture failure. Must be rectified and re-verified prior to dispatch.
- *Category B (Minor / Site Item)*: Cosmetic touch-up, drawing label revision, or minor spare part addition. Recorded with assigned owner and due date; does not impede shipping.
- *FAT Acceptance Rule*: System is accepted when 100% of tested loops pass, network redundancy operates seamlessly, and zero Category A items remain open.

#v(6pt)

// Sign-off Endorsement Box
#align(center)[
  #block(
    stroke: 1pt + rgb("1B365D"),
    fill: rgb("F8FAFC"),
    inset: 8pt,
    radius: 4pt,
    width: 100%,
    [
      #text(weight: "bold", size: 9pt, fill: rgb("1B365D"))[FACTORY ACCEPTANCE TEST (FAT) ENDORSEMENT — MCP-01 (CHASSIS C1, C2, C3, C4)] \
      #v(2pt)
      #text(size: 8pt)[Project: Ingredion Kalasin — SPRINT 18K TPA Spray Dryer Plant | Ref: KAL-FAT-AGN-2608-C1C4] \
      #v(3pt)
      [  ] *C1–C4 Accepted without Reservation* #h(15pt) [  ] *C1–C4 Accepted with Category B Punch Items* #h(15pt) [  ] *Rejected* \
      #v(10pt)
      #grid(
        columns: (1fr, 1fr),
        align(center)[
          #line(length: 70%, stroke: 0.5pt + rgb("64748B"))
          #text(size: 8pt, weight: "bold")[Representative of Client / Owner] \
          #text(size: 7pt, fill: rgb("64748B"))[Ingredion (Thailand) Co., Ltd.] \
          #text(size: 7pt)[Date: #box(width: 50pt, stroke: (bottom: 0.5pt))[]]
        ],
        align(center)[
          #line(length: 70%, stroke: 0.5pt + rgb("64748B"))
          #text(size: 8pt, weight: "bold")[Representative of Kalasin Engineering] \
          #text(size: 7pt, fill: rgb("64748B"))[Lead Automation / System Integrator] \
          #text(size: 7pt)[Date: #box(width: 50pt, stroke: (bottom: 0.5pt))[]]
        ]
      )
    ]
  )
]

#pagebreak()

// =========================================================================
// PAGE 3: SPECIAL METHODOLOGY — AUTOMATED DI/DO CLOSED-LOOP CROSS-WIRING
// =========================================================================

#align(center)[
  #block(
    fill: rgb("0F4C81"),
    inset: 10pt,
    radius: 4pt,
    width: 100%,
    [
      #text(size: 11pt, weight: "bold", fill: white)[SPECIAL TESTING METHODOLOGY: AUTOMATED DI / DO LOOPBACK SIMULATION] \
      #v(2pt)
      #text(size: 8.5pt, fill: rgb("E0F2FE"))[Temporary Dry Contact Cross-Wiring from Interposing Relay DO to DI Input Terminal] \
      #v(2pt)
      #text(size: 7.5pt, fill: rgb("CBD5E1"))[100% Comprehensive Wiring & Hardware Validation with Autogenerated Verification Report]
    ]
  )
]

#v(4pt)

== 1. Testing Objective & Closed-Loop Principle

To verify with 100% certainty that the wiring from the *PLC DO Module to the Relay Coil (A1/A2)*, the *Relay Dry Contact (11/14)*, the *Relay Return Line (TB-0VDC)*, and the *DI Field Terminal to the PLC DI Module* are completely sound, an *Automated Closed-Loop Cross-Wiring Routine* is deployed:

#align(center)[
  #block(
    stroke: 0.8pt + rgb("1F4E79"),
    inset: 8pt,
    fill: rgb("F1F5F9"),
    radius: 3pt,
    width: 98%,
    [
      #text(weight: "bold", size: 8.5pt)[CLOSED-LOOP CIRCUIT ARCHITECTURE:] \
      #v(3pt)
      #text(size: 8pt)[
        *[PLC 1756-OB32 Output]* $-->$ *[Wire to Relay Coil A1]* $-->$ *[Relay Coil A2 to TB-0VDC]* \
        $arrow.b$ *(Relay Coil Energizes & Closes Mechanical NO Contact 11 / 14)* \
        *[TB-24VDC Rail]* $-->$ *[Relay Contact 11/14]* $-->$ *[Temporary Loopback Jumper]* $-->$ *[DI Terminal TBDI]* $-->$ *[PLC 1756-IB32 Input]*
      ]
    ]
  )
]

#v(3pt)

== 2. Automated Step-by-Step Execution Sequence

1. *Automated Trigger*: The Python Automation Test Engine / Studio 5000 FAT Routine commands the target DO channel to Force ON (`DO = 1`).
2. *Hardware Response*: The 1756-OB32 pin fires 24VDC to Relay Coil terminal `A1`. The interposing relay coil energizes, pulling down the mechanical armature and closing the Normally Open (NO) dry contact terminals `11` and `14`.
3. *Loopback Feedback*: Positive +24VDC from the `TB-24VDC` distribution rail passes through closed contact `11/14`, traverses the temporary cross-wiring jumper, and enters the paired DI terminal block (`TBDI`).
4. *DI Sensing & Latency Check*: The corresponding 1756-IB32 digital input channel detects +24VDC and transitions from `0` to `1`. The test engine senses this feedback transition, validates that response latency is within limit ($Delta t < 200 "ms"$), and marks the loop as *PASS*.
5. *De-energization Check*: The test engine forces `DO = 0`. The relay de-energizes, the contact opens, and the DI bit drops back to `0` confirming no welded contacts or sticky relays.
6. *Auto-Report Generation*: The results (DO Tag, Relay Tag, DI Tag, Timestamp, Response Latency, and Verdict) are automatically logged into the master database and rendered into an official PDF/Excel report.

#v(3pt)

== 3. Core Technical Value & Proof of Quality

- *100% Wiring Verification*: Validates all 4 segments simultaneously: (1) PLC DO to Relay, (2) Relay Coil 0V return, (3) Relay contact switching, and (4) DI Terminal to PLC DI module.
- *Zero Human Forcing Error*: Completely eliminates manual wire-tapping or human misreading of terminal numbers. Tests up to 256 DO channels in under 5 minutes.
- *Deterministic Traceability*: Generates an autogenerated, timestamped test certificate signed off by the lead automation engineer.

#v(3pt)

== 4. Sample Autogenerated Loopback Test Log (Chassis C1–C4 Excerpt)

#table(
  columns: (0.7fr, 1fr, 1.2fr, 1.2fr, 1fr, 0.8fr, 0.7fr, 0.7fr),
  stroke: 0.5pt + rgb("CBD5E1"),
  fill: (col, row) => if row == 0 { rgb("1F4E79") } else if calc.even(row) { rgb("F8FAFC") } else { white },
  align: (col, row) => if row == 0 or col == 0 or col == 5 or col == 6 or col == 7 { center } else { left },
  table.header([*Loop*], [*DO Slot / Pin*], [*Relay Tag (A1/A2)*], [*Paired DI Terminal*], [*DI Slot / Pin*], [*Latency*], [*State*], [*Verdict*]),
  [LP-01], [C1S8 - Pin 1], [P1-TBRL1-1 (A1/A2)], [P1-TBDI1-1 (11/14)], [C1S4 - Pin 1], [28 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]],
  [LP-02], [C1S8 - Pin 2], [P1-TBRL1-2 (A1/A2)], [P1-TBDI1-2 (11/14)], [C1S4 - Pin 2], [31 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]],
  [LP-03], [C1S8 - Pin 3], [P1-TBRL1-3 (A1/A2)], [P1-TBDI1-3 (11/14)], [C1S4 - Pin 3], [26 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]],
  [LP-04], [C2S6 - Pin 1], [P2-TBRL1-1 (A1/A2)], [P2-TBDI1-1 (11/14)], [C2S1 - Pin 1], [34 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]],
  [LP-05], [C2S6 - Pin 2], [P2-TBRL1-2 (A1/A2)], [P2-TBDI1-2 (11/14)], [C2S1 - Pin 2], [29 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]],
  [LP-06], [C3S4 - Pin 1], [P3-TBRL1-1 (A1/A2)], [P3-TBDI1-1 (11/14)], [C3S1 - Pin 1], [33 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]],
  [LP-07], [C4S4 - Pin 1], [P4-TBRL1-1 (A1/A2)], [P4-TBDI1-1 (11/14)], [C4S1 - Pin 1], [37 ms], [1 -> 1], [#text(weight: "bold", fill: rgb("15803D"))[PASS]]
)

#v(4pt)

#align(right)[
  #text(size: 7.5pt, style: "italic", fill: rgb("64748B"))[Generated by Kalasin Engineering Automated Loopback Test Engine v2.4 | Master Report: KAL-FAT-TR-DIDO-2608]
]
