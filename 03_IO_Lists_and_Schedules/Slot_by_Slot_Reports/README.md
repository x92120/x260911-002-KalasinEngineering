# Slot_by_Slot_Reports — Per-Chassis Hardware Slot Reports

This directory contains individual Excel workbooks and print-ready PDF reports for ControlLogix Chassis C1 through C7 at the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📁 Chassis Hardware Reports Register

| Chassis Tag | Excel Workbook | Vector PDF Report | Hardware Allocation | Active Slots |
| :---: | :--- | :--- | :--- | :---: |
| **C1** | `Chassis_C1_Slots_IO_Report.xlsx` | `Chassis_C1_Slots_IO_Report.pdf` | Main Controller & Local I/O (1756-L950TPSXT, IB32, OB32, IF16) | Slots 03 to 10 |
| **C2** | `Chassis_C2_Slots_IO_Report.xlsx` | `Chassis_C2_Slots_IO_Report.pdf` | Process I/O Expansion (1756-IB32, OB32, IF16, OF8) | Slots 01 to 13 |
| **C3** | `Chassis_C3_Slots_IO_Report.xlsx` | `Chassis_C3_Slots_IO_Report.pdf` | Evaporator & Skid I/O (1756-IB32, OB32, IF16) | Slots 01 to 13 |
| **C4** | `Chassis_C4_Slots_IO_Report.xlsx` | `Chassis_C4_Slots_IO_Report.pdf` | Spray Dryer & Exhaust I/O (1756-IB32, OB32, IF16) | Slots 01 to 06 |
| **C5** | `Chassis_C5_Slots_IO_Report.xlsx` | `Chassis_C5_Slots_IO_Report.pdf` | Remote I/O Skid RIO-200 (1734-AENTR / 1756-IB32, OB32) | Slots 01 to 04 |
| **C6** | `Chassis_C6_Slots_IO_Report.xlsx` | `Chassis_C6_Slots_IO_Report.pdf` | MCC Direct Auxiliary Drop (1756-IB32, OB32) | Slots 05 to 06 |
| **C7** | `Chassis_C7_Slots_IO_Report.xlsx` | `Chassis_C7_Slots_IO_Report.pdf` | MCC Bus & Interlocks Drop (1756-IB32, OB32) | Slots 00 to 07 |

---

## ⚙️ Report Structure

Each per-chassis report contains dedicated sections per installed slot module, detailing channel index, PLC terminal pin, wire tag, instrument tag, description, destination JB, terminal block number, P&ID drawing reference, and signal type (DI, DO, AI, AO, RTD).
