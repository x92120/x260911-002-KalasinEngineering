# Ingredion Kalasin Spray Dryer Plant — Engineering Workspace

**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Engineering Disciplines**: Industrial OT Networks, Electrical Engineering, Instrumentation, Control Systems & CAD Automation  
**Control Platform**: Rockwell Automation ControlLogix 5580 (1756-L950TPSXT) / Studio 5000 / FactoryTalk View SE / Cisco Catalyst 9300 / Stratix 5700  
**Location**: Kalasin, Thailand  
**Ref / Order**: SPRINT 18K TPA Project  

---

## 📁 Workspace Directory Map

```text
x260911-002-KalasinEngineering/
├── 01_Network_Design/                  # Industrial OT Network Architecture, Cisco Core & Redundancy
│   ├── 01_Topology_and_Architecture/  # CPwE Topology Specs & Purdue Level 1-3 Network Architecture
│   ├── 02_IP_Address_Allocation/      # Subnet & IP Allocation Matrix (VLAN 10, 20, 30, 40, 50, 99)
│   ├── 03_Switch_and_Hardware_Config/ # Cisco Catalyst 9300 HSRP, Stratix 5700 & Server Teaming Configurations
│   ├── 04_DLR_and_Redundancy/         # Device Level Ring (DLR) Design & Supervisor Ring Rules
│   ├── 05_Drawings_and_Schematics/    # Network Riser Drawings (SVG, PDF, PNG, DXF & HTML Viewers)
│   ├── 06_Cable_Schedule/             # Cat6A Copper & OM3/OS2 Fiber Optic Cable Schedule
│   └── 07_Commissioning_and_FAT_SAT/  # SAT Test Sheets & Failover Validation Checklists
│
├── 02_Electrical_and_eDrawing/         # Electrical Schematics, CAD Deliverables & Instrument Specifications
│   ├── 01_Source_Drawings_DWG/        # Original Vendor AutoCAD DWG Drawings (SCD, LAY, WIR)
│   ├── 02_ePlan_Exports/              # Exported ePlan Project Packages (Spray Dryer & Electrical)
│   ├── 03_CAD_Exports_DXF_SVG/        # 154 Vectorized DXF & SVG Engineering Drawing Sheets
│   ├── 04_PDF_Reports_and_Drawings/   # 145 Engineering Drawing PDFs, JB Reports & Mounting Plans
│   ├── 05_Instrument_Manuals/         # 22 Vendor Instrument Technical & User Manuals & IS Barriers
│   ├── 06_SCADA_Demo/                 # FactoryTalk View SE HMI Demo Screens, Trends & Slides
│   └── 07_Presentations_and_Overviews/# PID Control Strategy Presentations & Loop Architecture
│
├── 03_IO_Lists_and_Schedules/          # Master Excel Workbooks, I/O Tag Databases & Junction Schedules
│   ├── Destination_Reports/           # 16 Field Junction Box & MCC Destination Schedule Workbooks
│   ├── IO_List/                       # Slot-Configured I/O Master Workbooks & Tag Databases
│   ├── IO_List-Bak/                   # Engineering Revision Archives & Raw Tag Imports
│   └── Slot_by_Slot_Reports/          # Per-Chassis (C1–C7) Slot-by-Slot Hardware Reports (Excel & PDF)
│
├── 04_Automation_Scripts/              # CAD Automation, PDF Report Builders & Interactive HTML Viewers
│   ├── cable_schedule_viewer.html      # Interactive Cable Schedule Search & Filter Tool
│   ├── pid_loop_control_viewer.html    # Interactive PID Loop Control Strategy Web Viewer
│   └── *.py                            # Python Scripts for CAD, PDF, SVG, and Excel Report Generation
│
├── 261005-IO_Config/                   # Remote I/O & Junction Panel Configuration Reports
├── Datasheet/                          # Pepperl+Fuchs Barrier & Instrument Field Datasheets
├── eDrawingTemplate/                   # AutoCAD DXF Drawing Frames, Title Blocks & Module Templates
├── x900-eDrawing/                      # Legacy Vendor ePlan EDB Projects & AutoCAD DWG Drawings
└── x9100-eDrawing/                     # 47 Individual Slot-by-Slot I/O DXF Schematics (C1S1–C5S4)
```

---

## 🎨 Visual Master Deliverables & Interactive Viewers

| Deliverable Sheet | Web / Vector / CAD Links | Description |
| :--- | :--- | :--- |
| **Redundant Cisco Core & Server Infrastructure** | [Web Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Redundant_Network_Diagram.html) • [PDF](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.pdf) • [SVG](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.svg) | Dual Cisco Catalyst 9300 HSRP v2, FT View SE Server Teaming, 10G Fiber Uplinks, MCC Switch Trunk |
| **ControlLogix 5-Chassis DLR Network & 13-Slot Rack** | [Web Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Network_and_Rack_Design.html) • [PDF](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf) • [SVG](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg) | ODVA DLR Ring (CS1–CS4 in MCP-01, CS5 RIO200 via Fiber) & 13-Slot Chassis Elevation |
| **SPRINT 18K Control Layout & System Architecture** | [Web Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_SPRINT_18K_Control_Layout.html) • [DXF](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.dxf) • [PDF](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.pdf) | Complete Plant Architecture matching hand-drawn draft S__103251990 |
| **Interactive Plant Cable Schedule Viewer** | [HTML Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/04_Automation_Scripts/cable_schedule_viewer.html) | Interactive tool for searching instrument cables, terminal pins, and destination JB routing |
| **Interactive PID Loop Control Viewer** | [HTML Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/04_Automation_Scripts/pid_loop_control_viewer.html) | Interactive viewer detailing PID control loops for Jet Cooker, Dryer Tower, and CIP |

---

## 📑 Core Documentation Index

### 1. Network Engineering (`01_Network_Design`)
* 📄 [Network Architecture Specification](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/01_Topology_and_Architecture/NETWORK_ARCHITECTURE_SPEC.md)
* 📄 [Plant Subnet & IP Allocation Matrix](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/02_IP_Address_Allocation/IP_ALLOCATION_MATRIX.md)
* 📄 [Cisco Core Switch Configuration Guide](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/CISCO_REDUNDANT_CORE_SWITCH_CONFIG.md)
* 📄 [Server Redundancy & Dual NIC Teaming](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/SERVER_REDUNDANCY_AND_NIC_TEAMING_CONFIG.md)
* 📄 [MCC Trunk & Stratix Switch Setup](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/MCC_TRUNK_AND_SWITCH_CONFIG.md)
* 📄 [Device Level Ring (DLR) Topology Design](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/04_DLR_and_Redundancy/DLR_TOPOLOGY_DESIGN.md)

### 2. Electrical & Drawings (`02_Electrical_and_eDrawing`)
* 📄 [Source DWG Drawings Overview](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/01_Source_Drawings_DWG/README.md)
* 📄 [ePlan Project Exports Index](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/02_ePlan_Exports/README.md)
* 📄 [CAD Exports (DXF/SVG) Index](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/README.md)
* 📄 [PDF Reports & Wiring PDF Drawings](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/README.md)
* 📄 [Pepperl+Fuchs IS Barriers Documentation](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/05_Instrument_Manuals/Pepperl+Fuchs%20IS%20Barriers/README.md)

### 3. I/O Lists & Schedules (`03_IO_Lists_and_Schedules`)
* 📄 [Destination Junction Box Schedules](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Destination_Reports/README.md)
* 📄 [Master Slot-Configured I/O Databases](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/IO_List/README.md)
* 📄 [Per-Chassis Slot-by-Slot Reports](file:///e:/xApp-01/x260911-002-KalasinEngineering/03_IO_Lists_and_Schedules/Slot_by_Slot_Reports/README.md)

### 4. CAD & Automation Tools (`04_Automation_Scripts`)
* 📄 [Automation Scripts Overview](file:///e:/xApp-01/x260911-002-KalasinEngineering/04_Automation_Scripts/README.md)
