# I/O Lists & Instrument Schedules

**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Engineering Discipline**: Instrumentation, Control System I/O Allocation & Cable Scheduling  

---

## 📌 Module Directory Overview

This module houses all master Excel workbooks, PDF I/O schedules, per-enclosure destination reports, per-chassis slot reports, and instrument cable mapping matrices for the plant control system.

```text
03_IO_Lists_and_Schedules/
├── Destination_Reports/             # 16 Field Junction Box & MCC Enclosure Workbooks
│   └── README.md
├── IO_List/                         # Slot-Configured I/O Master Workbooks & Tag Databases
│   └── README.md
├── IO_List-Bak/                     # Revision Archives, Raw Tag Imports & Historical Summaries
│   └── README.md
├── Slot_by_Slot_Reports/            # Per-Chassis (C1–C7) Slot-by-Slot Reports (Excel & PDF)
│   └── README.md
├── Chassis_and_Slot_IO_Source_Destination_Report.pdf  # ⭐ MASTER PDF: 99-Page I/O Schedule
├── Chassis_IO_List_Per_Slot_Report.xlsx               # ⭐ MASTER EXCEL: 48 Tabs (1 Sheet per Slot)
├── Destination_IO_List_Master_Report.pdf              # ⭐ MASTER PDF: 82-Page Destination Schedule
└── Destination_IO_List_Master_Report.xlsx             # ⭐ MASTER EXCEL: 17 Tabs (1 Sheet per Destination)
```

---

## 📑 Key Master Workbooks & PDF Reports

| Report / Deliverable Name | Format | Purpose | Link / Location |
| :--- | :---: | :--- | :--- |
| **Destination I/O Master Schedule** | PDF (82 Pages) | Master Field Enclosure Vector Schedule | [Destination_IO_List_Master_Report.pdf](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List-Bak/Destination_IO_List_Master_Report.pdf) |
| **Destination I/O Master Excel** | Excel (17 Tabs) | 1 Tab Per Field Destination (16 JBs + Index) | [Destination_IO_List_Master_Report.xlsx](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List-Bak/Destination_IO_List_Master_Report.xlsx) |
| **Per-Destination Workbooks** | Excel (16 Files) | Independent Enclosure Schedule Workbooks | [Destination_Reports/](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_Reports/README.md) |
| **Chassis & Slot Master PDF** | PDF (99 Pages) | Hardware Rack Wiring Vector Deliverable | [Chassis_and_Slot_IO_Source_Destination_Report.pdf](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List-Bak/Chassis_and_Slot_IO_Source_Destination_Report.pdf) |
| **Chassis I/O List Per Slot** | Excel (48 Tabs) | 1 Tab Per Hardware Slot (47 Slot Tabs + Index) | [Chassis_IO_List_Per_Slot_Report.xlsx](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List-Bak/Chassis_IO_List_Per_Slot_Report.xlsx) |
| **Per-Chassis Slot Packages** | Excel + PDF | Per-Chassis Hardware Packages (C1–C7) | [Slot_by_Slot_Reports/](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/README.md) |
| **Instrument & Cable Matrix** | Excel Workbook | Master Instrument-to-Cable Schedule | [Instrument_IO_Mapping_with_Models_and_Cables.xlsx](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List-Bak/Instrument_IO_Mapping_with_Models_and_Cables.xlsx) |
