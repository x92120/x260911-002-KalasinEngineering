# Industrial OT Network Design Specification
**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Standard**: Converged Plantwide Ethernet (CPwE) / ANSI/ISA-95 & ISA-99 (IEC 62443)  
**Plant Systems**: Area 202 (Pre-Slurry), Area 402 (Adjustment), Area 602 (Spray Dryer), Area 814 (Plant CIP)  
**Main Platform**: Rockwell Automation ControlLogix 5580, Stratix 5700/5400 Managed Switches, EtherNet/IP (CIP)  

---

## 📌 Overview & Design Goals

The Industrial Automation Network for Ingredion Kalasin is designed to ensure:
1. **Deterministic Performance**: Dedicated EtherNet/IP bandwidth for time-critical I/O and motion control.
2. **High Availability & Fault Tolerance**: Device Level Ring (DLR) sub-rings for field devices with sub-3ms media recovery time, and redundant trunk uplinks to the plant distribution switches.
3. **Cybersecurity & Segmentation**: Strict Purdue Model Level 1–3 network segmentation using Managed VLANs, Access Control Lists (ACLs), and Industrial Firewalls/NAT.
4. **Maintainability**: Standardized IP allocation, managed switch port mapping, and labeled cable schedule for rapid troubleshooting.

---

## 📂 Directory Layout

```text
01_Network_Design/
├── README.md                              # This specification index
├── 01_Topology_and_Architecture/          # CPwE architecture, Purdue model levels & physical/logical topologies
│   └── NETWORK_ARCHITECTURE_SPEC.md
├── 02_IP_Address_Allocation/              # IP addressing plan, subnets, VLAN ID assignments & device schedule
│   └── IP_ALLOCATION_MATRIX.md
├── 03_Switch_and_Hardware_Config/         # Stratix switch port maps, Cisco IOS configs & server setup
│   ├── CISCO_REDUNDANT_CORE_SWITCH_CONFIG.md  # Cisco Catalyst 9300 HSRP v2 & LACP Production Config
│   ├── SERVER_REDUNDANCY_AND_NIC_TEAMING_CONFIG.md # FT View SE Server Redundancy & Teaming Guide
│   ├── MCC_TRUNK_AND_SWITCH_CONFIG.md         # Stratix 5700 MCC 802.1Q Trunk Configuration
│   └── SWITCH_HARDWARE_SPEC.md
├── 04_DLR_and_Redundancy/                 # Device Level Ring (DLR) design, Ring Supervisors & PRP integration
│   └── DLR_TOPOLOGY_DESIGN.md
├── 05_Drawings_and_Schematics/            # Network Riser diagrams, cabinet connection schematics & CAD layouts
│   ├── Redundant_Cisco_Core_and_Server_Network_Diagram.svg # 🎨 Redundant Cisco & Server Architecture Drawing
│   ├── Redundant_Cisco_Core_and_Server_Network_Diagram.pdf # 📄 Printable A1/A2 PDF
│   ├── Redundant_Cisco_Core_and_Server_Network_Diagram.png # 🖼️ Ultra High-Res PNG (5417x3646)
│   ├── Viewer_Redundant_Network_Diagram.html               # 🌐 Interactive Redundant Network Viewer
│   ├── ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg # 🎨 ControlLogix 5-Chassis DLR Master Drawing
│   ├── ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf # 📄 ControlLogix Printable PDF
│   ├── ControlLogix_5Chassis_DLR_Network_and_Rack_Design.png # 🖼️ ControlLogix High-Res PNG
│   ├── Viewer_Network_and_Rack_Design.html                   # 🌐 ControlLogix Interactive Web Viewer
│   ├── SPRINT_18K_Control_Layout_System_Architecture.dxf   # 📐 AutoCAD DXF R2010
│   └── README.md
├── 06_Cable_Schedule/                     # Copper Cat6A & Fiber Optic (OM3/OS2) schedule & labeling
│   └── CABLE_SCHEDULE_TEMPLATE.md
└── 07_Commissioning_and_FAT_SAT/          # Network SAT test sheets, ping sweep records & ring failover tests

---

## 🎨 Master Deliverable Drawings
1. **Redundant Cisco Core & Server Infrastructure (Cisco HSRP, Server Teaming, MCC Trunk, Server Fiber)**:
   * **Vector SVG**: [Redundant_Cisco_Core_and_Server_Network_Diagram.svg](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.svg)
   * **Printable PDF**: [Redundant_Cisco_Core_and_Server_Network_Diagram.pdf](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.pdf)
   * **High-Res PNG**: [Redundant_Cisco_Core_and_Server_Network_Diagram.png](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Redundant_Cisco_Core_and_Server_Network_Diagram.png)
   * **Interactive Web Viewer**: [Viewer_Redundant_Network_Diagram.html](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Redundant_Network_Diagram.html)
2. **ControlLogix 5-Chassis DLR Network & 13-Slot Rack Redesign**:
   * **Vector SVG**: [ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg)
   * **Printable PDF**: [ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf)
   * **High-Res PNG**: [ControlLogix_5Chassis_DLR_Network_and_Rack_Design.png](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/ControlLogix_5Chassis_DLR_Network_and_Rack_Design.png)
   * **Interactive Web Viewer**: [Viewer_Network_and_Rack_Design.html](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_Network_and_Rack_Design.html)
3. **SPRINT 18K Control Layout & System Architecture (Matching Hand-drawn S__103251990)**:
   * **AutoCAD DXF (R2010)**: [SPRINT_18K_Control_Layout_System_Architecture.dxf](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/SPRINT_18K_Control_Layout_System_Architecture.dxf)
   * **Interactive Web Viewer**: [Viewer_SPRINT_18K_Control_Layout.html](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/01_Network_Design/05_Drawings_and_Schematics/Viewer_SPRINT_18K_Control_Layout.html)
   * **Interactive Web Viewer**: [Viewer_SPRINT_18K_Control_Layout.html](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/Network%20Design/05_Drawings_and_Schematics/Viewer_SPRINT_18K_Control_Layout.html)

---

## 🏷️ Key System Architecture Summary

| Layer | Purdue Level | Typical Devices | Network Media | Subnet Example |
|---|---|---|---|---|
| **Site Operations / SCADA** | Level 3 | FactoryTalk SE HMI Server, Historian, Engineering Workstation (EWS) | 1Gbps Fiber / Cat6A | `192.168.10.0/24` (VLAN 10) |
| **Supervisory / Operator HMI** | Level 2 | PanelView 5510, Area Thin Clients, Engineering Laptops | 1Gbps / 100Mbps Copper | `192.168.20.0/24` (VLAN 20) |
| **Control & Inter-PLC** | Level 1 | ControlLogix 5580 CPUs, 1756-EN4TR, Safety PACs | 1000BASE-T / Ring Trunk | `192.168.30.0/24` (VLAN 30) |
| **Field I/O & Drives (DLR Rings)** | Level 0/1 | POINT I/O 1734-AENTR, PowerFlex 525/755 VFDs, Flowmeters | 100BASE-TX M12 / RJ45 DLR | `192.168.40.0/24` (VLAN 40) |
