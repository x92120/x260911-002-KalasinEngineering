# I/O Lists & Instrument Schedules
**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Engineering Discipline**: Instrumentation, PLC I/O Allocation & Cable Scheduling  

---

## 📌 Module Overview

This folder contains all master Excel workbooks, I/O allocation schedules, and instrument-to-cable mapping matrices for the plant control system.

```text
03_IO_Lists_and_Schedules/
├── Destination_IO_List_Master_Report.pdf              # ⭐ MASTER PDF: 82-Page Destination I/O Schedule (Field Enclosure View)
├── Destination_IO_List_Master_Report.xlsx             # ⭐ MASTER: 1 Sheet Per Destination (16 Field Enclosure Tabs + Index)
├── Destination_Reports/                               # 📁 Per-Destination Workbooks (16 Individual Enclosure Workbooks)
│   ├── MCC_IO_Schedule.xlsx                          # Motor Control Center (237 Points: DI, DO, BUS)
│   ├── JB-607_IO_Schedule.xlsx                       # Spray Dryer 7F Junction Box (195 Points)
│   ├── JB-401_IO_Schedule.xlsx                       # Infeed 2F Junction Box (112 Points)
│   ├── RIO-200_IO_Schedule.xlsx                      # Slurry Outbuilding 2F Remote I/O Skid (112 Points)
│   ├── JB-601_IO_Schedule.xlsx                       # Spray Dryer 1F Junction Box (85 Points)
│   ├── JB-602_IO_Schedule.xlsx                       # Spray Dryer 3F Junction Box (82 Points)
│   ├── JB-402_IO_Schedule.xlsx                       # Jet Cooker 2F Junction Box (65 Points)
│   ├── JB-612_IO_Schedule.xlsx                       # Packing Tower 2F Junction Box (57 Points)
│   ├── JB-618_IO_Schedule.xlsx                       # Packing Tower 8F Junction Box (52 Points)
│   ├── IS-JB-618_IO_Schedule.xlsx                    # Packing Tower 8F Intrinsically Safe Ex i JB (44 Points)
│   ├── IS-JB-612_IO_Schedule.xlsx                    # Packing Tower 2F Intrinsically Safe Ex i JB (41 Points)
│   ├── CA1_IO_Schedule.xlsx                          # Main Control Panel CA1 Local Terminals (32 Points)
│   ├── IS-JB-603_IO_Schedule.xlsx                    # Spray Dryer 3F Intrinsically Safe Ex i JB (32 Points)
│   ├── IS-JB-608_IO_Schedule.xlsx                    # Spray Dryer 8F Intrinsically Safe Ex i JB (32 Points)
│   ├── JB-606_IO_Schedule.xlsx                       # Spray Dryer 6F Junction Box (26 Points)
│   └── JB-608_IO_Schedule.xlsx                       # Spray Dryer 8F Temperature JB (6 Points)
├── Chassis_and_Slot_IO_Source_Destination_Report.pdf  # ⭐ MASTER PDF: 99-Page Print-Ready Vector I/O Schedule
├── Chassis_IO_List_Per_Slot_Report.xlsx               # ⭐ MASTER: 1 Sheet Per 1 Slot (47 Slot Tabs + Interactive Index)
├── Chassis_and_Slot_IO_Source_Destination_Report.xlsx # ⭐ MASTER: Chassis & Slot Source-to-Destination Schedule (ePlan Tagged)
├── Slot_by_Slot_Reports/                             # 📁 Per-Chassis Workbooks & PDFs (1 Sheet per Slot for each Rack)
│   ├── Chassis_C1_Slots_IO_Report.xlsx / .pdf        # RACK C1: Main Controller & Local I/O (Slots 03 to 10)
│   ├── Chassis_C2_Slots_IO_Report.xlsx / .pdf        # RACK C2: Process I/O Expansion (Slots 01 to 13)
│   ├── Chassis_C3_Slots_IO_Report.xlsx / .pdf        # RACK C3: Evaporator & Skid I/O (Slots 01 to 13)
│   ├── Chassis_C4_Slots_IO_Report.xlsx / .pdf        # RACK C4: Spray Dryer & Exhaust (Slots 01 to 06)
│   ├── Chassis_C5_Slots_IO_Report.xlsx / .pdf        # RACK C5: Remote I/O Skid RIO-200 (Slots 01 to 04)
│   ├── Chassis_C6_Slots_IO_Report.xlsx / .pdf        # RACK C6: MCC Direct Auxiliary (Slots 05 to 06)
│   └── Chassis_C7_Slots_IO_Report.xlsx / .pdf        # RACK C7: MCC Bus & Interlocks (Slots 00 to 07)
├── IO_List_Sorted_by_Chassis_Slot.xlsx               # Master Chassis & Slot Hardware Allocation
├── Instrument_IO_Mapping_with_Models_and_Cables.xlsx # Master Field Instrument to Cable & I/O Mapping
├── IO_Comparison_Report_Tag35_vs_Dev36.xlsx          # Audit report comparing Tag35 vs Dev36 releases
├── IO Check List.xlsx                                # QC validation check sheet
├── IO Check List-1.xlsx                              # QC validation check sheet (Revision 1)
├── Tags list PLC.xlsx                                # Exported PLC tag schedule
├── IO_List_xDev-R01-Tag35-1.xlsx                     # Revision snapshot 1
├── IO_List_xDev-R01-Tag35-2.xlsx                     # Revision snapshot 2
├── IO_List_xDev-R01-Tag35-3.xlsx                     # Revision snapshot 3
├── IO_List_xDev-R01-Tag35-4.xlsx                     # Revision snapshot 4
├── IO_List_xDev-R01-Tag35-5.xlsx                     # Revision snapshot 5
├── IO_List_xDev-R01-Tag35-6.xlsx                     # Revision snapshot 6 (Baseline source)
├── IO_List_xDev-R01-Tag35-7.xlsx                     # Revision snapshot 7
└── scratch_wir_strings.txt                           # Wire string extraction scratchpad
```

---

## 📑 Key Workbooks & PDF Reports

| Report / Deliverable Name | Format | Purpose | Primary Use Case |
|---|:---:|---|---|
| [Destination_IO_List_Master_Report.pdf](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.pdf) | **PDF** (82 Pages) | **Master Destination Vector I/O Schedule** | Official field enclosure wiring & termination schedule for loop testing and site commissioning |
| [Destination_IO_List_Master_Report.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.xlsx) | **Excel** (17 Tabs) | **Master 1 Sheet Per Destination Report** | Dedicated tab per field junction box / MCC enclosure (16 destination tabs + Interactive Index) |
| [Destination_Reports/](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_Reports/) | **Excel Workbooks** | **Per-Destination Workbooks** | 16 individual workbooks for each field junction box & MCC enclosure |
| [Chassis_and_Slot_IO_Source_Destination_Report.pdf](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.pdf) | **PDF** (99 Pages) | **Master Chassis & Slot Vector Schedule** | Official hardware rack wiring deliverable (Slots 0 to 12 across Racks C1 to C7) |
| [Chassis_IO_List_Per_Slot_Report.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Chassis_IO_List_Per_Slot_Report.xlsx) | **Excel** (48 Tabs) | **Master 1 Sheet Per 1 Slot Report** | Dedicated tab per hardware slot (47 slot sheets + Interactive Hyperlinked Index) |
| [Slot_by_Slot_Reports/](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/) | **Excel + PDF** | **Per-Chassis Slot Packages** | Independent workbooks & vector PDFs for C1 through C7 with 1 section per slot |
| [Chassis_and_Slot_IO_Source_Destination_Report.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.xlsx) | **Excel** | Master Chassis & Slot Schedule | Comprehensive I/O report with ePlan wire tags, PLC card pins, JB locations, field instruments, and P&ID references |
| [IO_List_Sorted_by_Chassis_Slot.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List_Sorted_by_Chassis_Slot.xlsx) | Master Hardware I/O Schedule | Used for wiring termination and Studio 5000 tag assignment |
| [Instrument_IO_Mapping_with_Models_and_Cables.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Instrument_IO_Mapping_with_Models_and_Cables.xlsx) | Instrument & Cable Matrix | Used for procurement, cable pulling, and junction box sizing |
| [IO_Comparison_Report_Tag35_vs_Dev36.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_Comparison_Report_Tag35_vs_Dev36.xlsx) | Tag Audit & Revision Tracking | Verification of changes across engineering baselines |
| [Tags list PLC.xlsx](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Tags%20list%20PLC.xlsx) | ControlLogix Tag Database | Direct tag import for PLC logic & SCADA development |
