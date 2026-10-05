# Industrial OT Network Design Specification

**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Standard**: Converged Plantwide Ethernet (CPwE) / ANSI/ISA-95 & ISA-99 (IEC 62443)  
**Plant Systems**: Area 202 (Pre-Slurry), Area 402 (Adjustment), Area 602 (Spray Dryer), Area 814 (Plant CIP)  
**Main Platform**: Rockwell Automation ControlLogix 5580, Stratix 5700/5400 Managed Switches, EtherNet/IP (CIP), Cisco Catalyst 9300 Core  

---

## 📌 Overview & Architecture Principles

The Industrial OT Network for the Ingredion Kalasin Spray Dryer Plant is architected according to CPwE guidelines to ensure:
1. **Deterministic Performance**: Dedicated EtherNet/IP bandwidth for real-time PLC-to-PLC communication and I/O scan loops.
2. **High Availability & Fault Tolerance**: Device Level Ring (DLR) for field I/O racks with sub-3ms media recovery time, redundant Cisco Catalyst 9300 core switches (HSRP v2), and dual-homed server NIC teaming.
3. **Purdue Security Segmentation**: Strict segmentation across Purdue Levels 0–3 using managed VLANs (10, 20, 30, 40, 50, 99) and access control policies.
4. **Structured Cabling**: Standardized Cat6A industrial copper and OM3/OS2 fiber optic links with defined cabling schedules.

---

## 📁 Directory Structure & Module Index

```text
01_Network_Design/
├── README.md                              # Main Industrial OT Network Specification
├── 01_Topology_and_Architecture/          # CPwE Architecture, Purdue Levels & Topology Specifications
│   ├── README.md
│   └── NETWORK_ARCHITECTURE_SPEC.md       # Core Architecture & Purdue Model Definition
├── 02_IP_Address_Allocation/              # Subnetting Plan, VLAN Assignments & Device Schedule
│   ├── README.md
│   └── IP_ALLOCATION_MATRIX.md            # Detailed IP Address Assignment Table
├── 03_Switch_and_Hardware_Config/         # Cisco IOS Core Configs, Stratix Switch Port Maps & Server Setup
│   ├── README.md
│   ├── CISCO_REDUNDANT_CORE_SWITCH_CONFIG.md  # Cisco Catalyst 9300 HSRP v2 & LACP Production Config
│   ├── MCC_TRUNK_AND_SWITCH_CONFIG.md         # Stratix 5700 MCC 802.1Q Trunk Configuration
│   ├── SERVER_REDUNDANCY_AND_NIC_TEAMING_CONFIG.md # FT View SE Server Redundancy & Teaming Guide
│   └── SWITCH_HARDWARE_SPEC.md                # Managed Switch Hardware Specifications
├── 04_DLR_and_Redundancy/                 # Device Level Ring (DLR) Topology & Ring Supervisor Rules
│   ├── README.md
│   └── DLR_TOPOLOGY_DESIGN.md             # DLR Ring Supervisor Rules & Active/Backup Node Setup
├── 05_Drawings_and_Schematics/            # Master Vector Network Riser Diagrams, Elevation & Web Viewers
│   ├── README.md
│   ├── Redundant_Cisco_Core_and_Server_Network_Diagram.svg
│   ├── ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg
│   ├── SPRINT_18K_Control_Layout_System_Architecture.dxf
│   └── *.html / *.pdf / *.png
├── 06_Cable_Schedule/                     # Copper Cat6A & Fiber Optic Cable Schedule
│   ├── README.md
│   └── CABLE_SCHEDULE_TEMPLATE.md         # Industrial Cable Schedule Specification
└── 07_Commissioning_and_FAT_SAT/          # SAT Checklists, Ping Sweep Records & Failover Testing
    ├── README.md
    └── NETWORK_COMMISSIONING_CHECKLIST.md # Network Testing & Validation Protocol
```

---

## 🎨 Master Deliverable Drawings

1. **Redundant Cisco Core & Server Infrastructure**:
   * [Interactive Web Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Redundant_Network_Diagram.html)
   * [Printable PDF](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.pdf)
   * [Vector SVG](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.svg)
2. **ControlLogix 5-Chassis DLR Network & 13-Slot Rack Redesign**:
   * [Interactive Web Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Network_and_Rack_Design.html)
   * [Printable PDF](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf)
   * [Vector SVG](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg)
3. **SPRINT 18K Control Layout & System Architecture**:
   * [Interactive Web Viewer](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_SPRINT_18K_Control_Layout.html)
   * [AutoCAD DXF (R2010)](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.dxf)

---

## 🏷️ Network Layer Summary

| Layer | Purdue Level | Typical Devices | Network Media | Subnet Example |
|---|---|---|---|---|
| **Site Operations / SCADA** | Level 3 | FactoryTalk SE HMI Server, Historian, Engineering Workstation (EWS) | 1Gbps Fiber / Cat6A | `192.168.10.0/24` (VLAN 10) |
| **Supervisory / Operator HMI** | Level 2 | PanelView 5510, Area Thin Clients, Engineering Laptops | 1Gbps / 100Mbps Copper | `192.168.20.0/24` (VLAN 20) |
| **Control & Inter-PLC** | Level 1 | ControlLogix 5580 CPUs, 1756-EN4TR, Safety PACs | 1000BASE-T / Ring Trunk | `192.168.30.0/24` (VLAN 30) |
| **Field I/O & Drives (DLR Rings)** | Level 0/1 | POINT I/O 1734-AENTR, PowerFlex 525/755 VFDs, Flowmeters | 100BASE-TX M12 / RJ45 DLR | `192.168.40.0/24` (VLAN 40) |
