# 06_Cable_Schedule — Industrial Network Cable Specifications

This directory contains the cable schedule standards, media specifications, and labeling conventions for the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📄 Key Document

* **[CABLE_SCHEDULE_TEMPLATE.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/06_Cable_Schedule/CABLE_SCHEDULE_TEMPLATE.md)**: Master network cabling specification, cable ID format, wire colors, and shielding rules.

---

## 🔌 Cable Media & Standards

| Cable Code | Cable Type | Conductor / Spec | Typical Usage | Max Distance |
| :---: | :--- | :--- | :--- | :---: |
| **CAT6A-SH** | Cat6A Shielded S/FTP | 24 AWG Twisted Pair, Teal PUR Jacket | EtherNet/IP Field DLR & Switch Links | 100 m |
| **FO-SM-OS2** | Single-Mode Fiber OS2 | 9/125 µm Duplex LC | MCP-01 to RIO200 & MCC Fiber Trunk | 2 km |
| **FO-MM-OM3** | Multi-Mode Fiber OM3 | 50/125 µm Duplex LC | Server Room to Core Switch Interconnect | 300 m |

---

## 🏷️ Cable Labeling Format

Format: `[Source Device] - [Media Type] - [Destination Device] - [Cable No.]`  
Example: `CS1-EN4TR1 - CAT6A - CS2-EN4TR1 - CABLE-0104`
