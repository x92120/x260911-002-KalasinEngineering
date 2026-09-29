# Ingredion Kalasin Spray Dryer Plant — Engineering Workspace
**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Engineering Disciplines**: Automation, Electrical Engineering, Industrial OT Networks, Instrumentation & I/O Systems  
**Control Platform**: Rockwell Automation ControlLogix 5580 (1756-L950TPSXT) / Studio 5000 / FactoryTalk View SE / Cisco Catalyst 9300 / Stratix 5700  

---

## 📁 Repository Structure

```text
x260911-002-KalasinEngineering/
├── 01_Network_Design/                  # Industrial OT Network Design, Cisco Core & Redundancy
│   ├── 01_Topology_and_Architecture/  # CPwE Topology Specs & Purdue Level 1-3 Segmentation
│   ├── 02_IP_Address_Allocation/      # Plant Subnet & IP Address Matrix (VLAN 10,20,30,50,99)
│   ├── 03_Switch_and_Hardware_Config/ # Cisco Catalyst 9300 HSRP, Stratix 5700 & Server Teaming
│   ├── 04_DLR_and_Redundancy/         # Device Level Ring (DLR) Design & Supervisor Node Rules
│   ├── 05_Drawings_and_Schematics/    # Network Master Drawings (Light Style SVG, PDF, PNG, HTML)
│   ├── 06_Cable_Schedule/             # Copper Cat6A & Fiber Optic Cable Schedules
│   └── 07_Commissioning_and_FAT_SAT/  # SAT Test Sheets & Failover Validation Checklists
│
├── 02_Electrical_and_eDrawing/         # Electrical Schematics, Panel Layouts & CAD Deliverables
│   ├── 01_Source_Drawings_DWG/        # Original Vendor AutoCAD DWG Drawings (SCD, LAY, WIR)
│   ├── 02_ePlan_Exports/              # Exported ePlan Project Packages (Spray Dryer & Electrical)
│   ├── 03_CAD_Exports_DXF_SVG/        # 154 Vectorized DXF & SVG Engineering Drawing Sheets
│   ├── 04_PDF_Reports_and_Drawings/   # 145 Engineering Drawing PDFs, JB Reports & Mounting Plans
│   ├── 05_Instrument_Manuals/         # 22 Vendor Instrument Technical & User Manuals
│   ├── 06_SCADA_Demo/                 # FactoryTalk View SE SCADA Demo Videos & Frames
│   └── 07_Presentations_and_Overviews/# PID Control Presentations & Loop Architecture Overviews
│
├── 03_IO_Lists_and_Schedules/          # Master Excel Workbooks & Tag Databases
│   ├── Destination_IO_List_Master_Report.pdf              # ⭐ MASTER PDF: 82-Page Destination I/O Schedule (Field Enclosure View)
│   ├── Destination_IO_List_Master_Report.xlsx             # ⭐ MASTER: 1 Sheet Per Destination (16 Field Enclosure Tabs + Index)
│   ├── Destination_Reports/                               # 📁 Per-Destination Workbooks (16 Individual Enclosure Workbooks)
│   ├── Chassis_and_Slot_IO_Source_Destination_Report.pdf  # ⭐ MASTER PDF: 99-Page Print-Ready Vector I/O Schedule
│   ├── Chassis_IO_List_Per_Slot_Report.xlsx               # ⭐ MASTER: 1 Sheet Per 1 Slot (47 Slot Tabs + Interactive Index)
│   ├── Chassis_and_Slot_IO_Source_Destination_Report.xlsx # Master Chassis & Slot Source-to-Destination Schedule (ePlan Tagged)
│   ├── Slot_by_Slot_Reports/                             # Per-Chassis Workbooks (1 Sheet per Slot)
│   ├── IO_List_Sorted_by_Chassis_Slot.xlsx               # Master Chassis & Slot Hardware Allocation
│   ├── Instrument_IO_Mapping_with_Models_and_Cables.xlsx # Master Instrument-to-Cable Schedule
│   ├── IO_Comparison_Report_Tag35_vs_Dev36.xlsx          # Tag Audit & Change Verification
│   ├── Tags list PLC.xlsx                                # ControlLogix Master Tag Database
│   └── IO Check List & Revision History                  # Quality Assurance Check Sheets
│
└── 04_Automation_Scripts/              # Automated Drawing Generators & Interactive Viewers
    ├── cable_schedule_viewer.html      # Interactive Cable Schedule Search & Inspection Tool
    ├── pid_loop_control_viewer.html    # Interactive PID Loop Control Strategy Viewer
    └── python scripts                  # Batch CAD, PDF, SVG, and Excel Report Generators
```

---

## 🎨 Master Visual Deliverables (Light Style Print-Ready)

| Deliverable | Formats | Description |
|---|---|---|
| **Destination I/O Master Schedule (Field Enclosure View)** | [PDF](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Destination_IO_List_Master_Report.pdf) • [Excel (17 Tabs)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_IO_List_Master_Report.xlsx) • [16 Enclosure Workbooks](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_Reports/) | Complete 82-Page Vector Schedule organized by Field Destination (16 Junction Boxes & MCC Enclosure) with physical location, terminal pin, instrument tag, and PLC Source mapping |
| **Chassis & Slot I/O Source-to-Destination Schedule** | [PDF](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Chassis_and_Slot_IO_Source_Destination_Report.pdf) • [Excel (48 Tabs)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Chassis_IO_List_Per_Slot_Report.xlsx) • [Master Excel](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Chassis_and_Slot_IO_Source_Destination_Report.xlsx) | Complete 99-Page Vector Schedule: 7 Chassis, 78 Slots, 1,210 Channels with ePlan Wire Tags, PLC Pins, JB Destinations, and P&ID references |
| **Redundant Cisco Core & Server Infrastructure** | [Web Viewer](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Redundant_Network_Diagram.html) • [PDF](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.pdf) • [PNG](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.png) • [SVG](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.svg) | Dual Cisco Catalyst 9300 (HSRP v2), Dual-Homed Servers, Heartbeat, MCC Trunk, 10G Fiber Uplink |
| **ControlLogix 5-Chassis DLR Network & 13-Slot Rack Elevation** | [Web Viewer](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Network_and_Rack_Design.html) • [PDF](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf) • [PNG](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.png) • [SVG](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg) | ODVA DLR Ring (CS1–CS4 in MCP-01, CS5 RIO200 via Fiber) & Balanced 13-Slot Hardware Redesign |
| **SPRINT 18K Control Layout & System Architecture** | [Web Viewer](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_SPRINT_18K_Control_Layout.html) • [DXF](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.dxf) • [PDF](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.pdf) • [SVG](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.svg) | Full Plant Control Architecture matching hand-drawn engineering draft S__103251990 |

---

## 📑 Core Documentation Links

* **Network Design**:
  * [01_Network_Design Overview](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/README.md)
  * [Cisco Redundant Switch Production Configuration](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/CISCO_REDUNDANT_CORE_SWITCH_CONFIG.md)
  * [Server Redundancy & Dual-Homed NIC Teaming Guide](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/SERVER_REDUNDANCY_AND_NIC_TEAMING_CONFIG.md)
  * [MCC Room Trunk & Stratix Switch Configuration](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/MCC_TRUNK_AND_SWITCH_CONFIG.md)
  * [IP Allocation Matrix](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/02_IP_Address_Allocation/IP_ALLOCATION_MATRIX.md)
* **Electrical & eDrawing**:
  * [02_Electrical_and_eDrawing Overview](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/README.md)
  * [Master Jet Cooker Engineering Drawings (PDF)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/KALASIN_JET_COOKER_MASTER_ENGINEERING_DRAWINGS_COMPLETE.pdf)
  * [Complete PLC Wiring Drawings (PDF)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/Jet_Cooker_PLC_IO_Wiring_Drawings_Complete.pdf)
* **I/O Lists & Schedules**:
  * [03_IO_Lists_and_Schedules Overview](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/README.md)
  * [Master Hardware I/O Schedule (Excel)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List_Sorted_by_Chassis_Slot.xlsx)
  * [Field Instrument to Cable Mapping Matrix (Excel)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Instrument_IO_Mapping_with_Models_and_Cables.xlsx)
* **Automation Tools**:
  * [04_Automation_Scripts Overview](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/04_Automation_Scripts/README.md)
  * [Interactive Cable Schedule Viewer (HTML)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/04_Automation_Scripts/cable_schedule_viewer.html)
  * [Interactive PID Control Strategy Viewer (HTML)](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/04_Automation_Scripts/pid_loop_control_viewer.html)
