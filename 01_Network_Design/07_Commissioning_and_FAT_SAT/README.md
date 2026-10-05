# 07_Commissioning_and_FAT_SAT — Network SAT & Validation Protocol

This directory contains Site Acceptance Testing (SAT) protocols, network commissioning checklists, ping sweep verification templates, and ring failover validation sheets for the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📄 Key Document

* **[NETWORK_COMMISSIONING_CHECKLIST.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/07_Commissioning_and_FAT_SAT/NETWORK_COMMISSIONING_CHECKLIST.md)**: Master Site Acceptance Testing (SAT) checklist, failover timing test steps, and sign-off form.

---

## 📋 Commissioning Test Steps Summary

1. **Physical Cable & Shield Continuity Test**: Fluke DSX-8000 CableAnalyzer verification for all Cat6A and fiber optic links.
2. **Ping & ARP Sweep Test**: Verification of IP accessibility across all VLANs (VLAN 10, 20, 30, 40, 50, 99).
3. **DLR Break Failover Test**: Physical disconnect of Ring Port 1 to verify sub-3ms switchover to Ring Port 2 without PLC minor/major fault.
4. **Cisco HSRP Core Failover Test**: Active Catalyst 9300 power cycle to verify virtual gateway transfer to Standby unit under 1 second.
