# Junction Box JB-401 Engineering Drawing & Testing Package

**Project**: Project Sprint 18K TPA Spray Dryer / Jet Cooker  
**Client**: Ingredion (Thailand) Co., Ltd.  
**Contractor**: Kalasin Engineering Co., Ltd. (Ref: x2608003)  
**Location**: Infeed 2nd Floor  
**Enclosure**: JB-401 (Standard Stainless Steel Field Junction Box, 800 x 600 x 250 mm)  
**Channel Capacity**: 112 Total Channels (85 Active, 27 Spares / 24.1% Margin)

---

## 1. Master Drawing Deliverables in `IO_VO`

| File Name | Format | Description & Engineering Content |
| :--- | :---: | :--- |
| **`JB-401_Master_Engineering_Drawings_Complete.pdf`** | **PDF (20 Pages)** | Complete compiled master engineering drawing book for JB-401 containing the System Configuration Diagram (SCD), full ControlLogix 1756 Chassis C1 I/O schematics (DI, DO, AI), and all 13 instrument typical loop hook-up diagrams for instruments landed in JB-401. |
| **`JB-401_Terminal_Wiring_Report.pdf`** | **PDF (4 Pages)** | Point-by-point field terminal wiring schedule, terminal numbers (P1-TBDI1, P1-TBDI2, P1-TBDI3, P1-TBRL1, P1-TBRL2, P1-TBAI1), cable wire numbers, PLC rack/slots, and tag assignments. |
| **`Loop_Test_Report_JB-401.xlsx`** | **Excel (43 Sheets)** | Formal Loop Test Report: Sheet `00_JB_Index` + 42 dedicated equipment testing certificates (1 sheet per equipment) with consolidated multi-signal channels (AI: 2-pin differential A/B). |
| **`Loop_Test_Report_JB-401.pdf`** | **PDF (43 Sheets)** | Printable vector PDF export of the complete JB-401 Loop Test Report workbook. |
| **`JB-401_IO_Schedule.xlsx`** | **Excel** | Master engineering I/O allocation schedule specifically for destination JB-401. |

---

## 2. Individual Schematic Drawings in `IO_VO/JB-401_Drawings/`

| Sheet No. | Drawing Number | Signal Type & Terminal Strip | Equipment & Instruments Landed in JB-401 |
| :---: | :--- | :--- | :--- |
| **03** | `KAL-JC-WIR-Sheet_03` | Main Controller Architecture | ControlLogix 1756-A13 Chassis C1 in Panel M1 (CA1 Marshalling) |
| **04** | `KAL-JC-WIR-Sheet_04` | 1756-IB32 DI (`P1-TBDI1`) | FSL-40201/02, FT-40201/02 Pulse, HV-40201..06, XV-40201A..C Feedback |
| **05** | `KAL-JC-WIR-Sheet_05` | 1756-IB32 DI (`P1-TBDI2`) | XV-40201D..H, XV-40202A..E Open & Close Feedback Limit Switches |
| **06** | `KAL-JC-WIR-Sheet_06` | 1756-IB32 DI (`P1-TBDI3`) | XV-40202F..H, XV-40203A..D Open & Close Feedback Limit Switches |
| **08** | `KAL-JC-WIR-Sheet_08` | 1756-OB32 DO (`P1-TBRL1`) | Solenoid Actuation Commands for XV-40201A..D, Relay Interposing Coils |
| **09** | `KAL-JC-WIR-Sheet_09` | 1756-OB32 DO (`P1-TBRL2`) | Solenoid Actuation Commands for XV-40201E..H, XV-40202A..H, XV-40203A..D |
| **10** | `KAL-JC-WIR-Sheet_10` | 1756-IF16 AI (`P1-TBAI1`) | 4-20mA Current Loops: FT-40201 (Flow/Density), FT-40202, LT-40201..03, pHT-40201/02 |

*Both high-resolution vector PDF and SVG versions of each sheet are stored in `IO_VO/JB-401_Drawings/`.*

---

## 3. Instrument Typical Hook-Up Diagrams in `JB-401_Master_Engineering_Drawings_Complete.pdf`

- **Page 8 (TYP-003)**: pH Analytical Measurement Loop (`pHT-40201`, `pHT-40202`)
- **Page 9 (TYP-006)**: Magnetic Flowmeter 4-Wire Active Loop (`FT-40202`)
- **Page 10 (TYP-007)**: Coriolis Mass Flowmeter 4-Wire Loop (`FT-40201`)
- **Page 11 (TYP-010)**: Differential Pressure Level / Flow (`PDT-40201`)
- **Page 12 (TYP-011)**: Guided Wave / Free Space Radar Level (`LT-40201`, `LT-40202`, `LT-40203`)
- **Page 13 (TYP-012)**: RTD Temperature Transmitter Pt100 (`TT-40201`..`03`)
- **Page 14 (TYP-014)**: Control Valve with SMART Positioner (`FCV-40201`, `TCV-40201`)
- **Page 15 (TYP-017)**: Pneumatic On-Off Solenoid Valve with Dual Proximity Limit Switches (`XV-40201A`..`H`, `XV-40202A`..`H`, `XV-40203A`..`D`)
- **Page 16 (TYP-019)**: Vibrating Fork Level Switch High/Low (`LSH-40201`, `LSL-40201`)
- **Page 17 (TYP-020)**: Pressure Switch High/Low (`PSH-40201`, `PSL-40201`)
