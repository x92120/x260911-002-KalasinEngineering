# 02_IP_Address_Allocation — Subnet & IP Matrix

This directory defines the plant-wide IP address allocation matrix and VLAN segmentation schema for the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📄 Key Document

* **[IP_ALLOCATION_MATRIX.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/02_IP_Address_Allocation/IP_ALLOCATION_MATRIX.md)**: Full master IP address allocation table for PLCs, HMIs, Remote I/O adapters, VFDs, Managed Switches, and SCADA Servers.

---

## 🌐 Industrial VLAN Assignment Table

| VLAN ID | Subnet CIDR | Description / Target Equipment | Gateway / HSRP |
| :---: | :--- | :--- | :--- |
| **VLAN 10** | `192.168.10.0/24` | SCADA Servers, Historian, Engineering Workstations (Level 3) | `192.168.10.1` |
| **VLAN 20** | `192.168.20.0/24` | PanelView HMI Touchscreens & Thin Clients (Level 2) | `192.168.20.1` |
| **VLAN 30** | `192.168.30.0/24` | Primary PAC Controller Inter-PLC Network (Level 1) | `192.168.30.1` |
| **VLAN 40** | `192.168.40.0/24` | Field Remote I/O & DLR EtherNet/IP Sub-ring (Level 0/1) | N/A (Isolated DLR) |
| **VLAN 50** | `192.168.50.0/24` | PowerFlex VFDs & Motor Control Center (MCC) Devices | `192.168.50.1` |
| **VLAN 99** | `10.10.99.0/24` | Network Infrastructure Management (Stratix / Cisco Switches) | `10.10.99.1` |
