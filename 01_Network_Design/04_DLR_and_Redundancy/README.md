# 04_DLR_and_Redundancy — Device Level Ring Topology

This directory details the ODVA EtherNet/IP Device Level Ring (DLR) design, ring supervisor configuration, and redundancy protocols for the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📄 Key Document

* **[DLR_TOPOLOGY_DESIGN.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/04_DLR_and_Redundancy/DLR_TOPOLOGY_DESIGN.md)**: Device Level Ring topology specification, active ring supervisor setup, backup supervisor rules, and sub-3ms recovery benchmarks.

---

## ⭕ DLR Ring Allocation Summary

| Ring ID | Location / Cabinet | Included Nodes / Chassis | Active Supervisor | Backup Supervisor | Media / Speed |
| :---: | :--- | :--- | :--- | :--- | :--- |
| **DLR-01** | MCP-01 (Main Control Panel) | Chassis CS1, CS2, CS3, CS4 | 1756-EN4TR (CS1 Slot 1) | 1756-EN4TR (CS2 Slot 1) | Cat6A Shielded (100Mbps Full-Duplex) |
| **DLR-02** | RIO200 (Remote I/O Cabinet) | Chassis CS5 (RIO200 1734-AENTR / 1756-RIO) | 1756-EN4TR (CS5 Slot 1) | Stratix 5700 Managed Switch | Single-Mode Fiber Optic Link |

---

## ⚡ Performance Specifications

* **Media Recovery Time**: < 3 ms upon single link disruption.
* **Beacon Interval**: 400 microseconds.
* **Timeout Value**: 1,960 microseconds.
