# 01_Topology_and_Architecture — Industrial OT Network Architecture

This directory contains the core topology and architecture specifications for the **Ingredion Kalasin Spray Dryer Plant** industrial control system network, designed according to Converged Plantwide Ethernet (CPwE) standards.

---

## 📄 Key Documents & Specifications

* **[NETWORK_ARCHITECTURE_SPEC.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/01_Topology_and_Architecture/NETWORK_ARCHITECTURE_SPEC.md)**: Detailed architecture specification covering Purdue Model Levels 1–3, industrial VLAN segmentation, firewall boundaries, and network redundancy standards.

---

## 📐 CPwE Network Architecture Overview

```text
+-------------------------------------------------------------------+
| Level 3: Industrial Zone / SCADA Servers (FactoryTalk SE, EWS)    |
+-------------------------------------------------------------------+
                                  |
                   [Cisco Catalyst 9300 Core HSRP]
                                  |
+-------------------------------------------------------------------+
| Level 2: Supervisory Control / HMI & PanelView Touchscreens      |
+-------------------------------------------------------------------+
                                  |
                   [Stratix 5700 Industrial Switches]
                                  |
+-------------------------------------------------------------------+
| Level 1: Control (ControlLogix 5580 PACs & 1756-EN4TR Cards)     |
+-------------------------------------------------------------------+
                                  |
                   [Device Level Ring (DLR) Topology]
                                  |
+-------------------------------------------------------------------+
| Level 0: Field I/O (POINT I/O 1734-AENTR, PowerFlex VFDs, Sensors)|
+-------------------------------------------------------------------+
```

---

## 🎯 Primary Architectural Goals

1. **Sub-3ms Ring Recovery**: ODVA Device Level Ring (DLR) topology for field chassis (CS1–CS5).
2. **Core Redundancy**: Dual Cisco Catalyst 9300 core switches operating with Hot Standby Router Protocol (HSRP v2).
3. **Dual-Homed Server NIC Teaming**: FactoryTalk View SE servers configured with active/standby NIC teaming for seamless failover.
