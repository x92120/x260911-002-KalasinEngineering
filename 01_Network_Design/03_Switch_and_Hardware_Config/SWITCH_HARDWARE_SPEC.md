# Industrial Switch & Hardware Configuration Spec
**Project**: Ingredion Kalasin — Sprint 18K TPA Spray Dryer Plant  
**Switch Hardware Standard**: Rockwell Automation Allen-Bradley Stratix 5700 / 5400 Managed Industrial Switches  

---

## 1. Switch Bill of Materials (BOM)

| Switch Tag | Model Number | Total Ports | Fiber SFP Ports | Location | Role / Uplink |
|---|---|---|---|---|---|
| `SW-COR-01` | **Stratix 5400 (1783-HMS16S4CGN)** | 20 Ports | 4x 1Gbps SFP Fiber | Central Server Rack | Core / Distribution Layer |
| `SW-201`    | **Stratix 5700 (1783-BMS10CGN)**   | 10 Ports | 2x Combo SFP Fiber | RIO-202 (Area 202) | Area 202 Aggregation & DLR |
| `SW-401`    | **Stratix 5700 (1783-BMS10CGN)**   | 10 Ports | 2x Combo SFP Fiber | RIO-402 (Area 402) | Area 402 Distribution |
| `SW-601`    | **Stratix 5700 (1783-BMS20CGN)**   | 20 Ports | 2x Combo SFP Fiber | MCP-602 (Dryer)    | Dryer Process & Burner DLR |
| `SW-801`    | **Stratix 5700 (1783-BMS10CGN)**   | 10 Ports | 2x Combo SFP Fiber | CIP-01 (CIP Skid)  | Area 814 Distribution |

---

## 2. Standard Stratix Managed Switch Configuration Settings

* **Spanning Tree Protocol**: Rapid Spanning Tree Protocol (RSTP - IEEE 802.1w) enabled on all inter-switch trunks.
* **QoS (Quality of Service)**:
  - Trust DSCP tags from Rockwell 1756-EN2TR/EN4TR and PowerFlex drives.
  - Priority Queue 1 assigned to PTP (Precision Time Protocol / IEEE 1588) and CIP Motion.
  - Priority Queue 2 assigned to CIP Implicit I/O (UDP 2222).
* **Multicast Management**:
  - IGMP Snooping: Enabled globally.
  - IGMP Querier: Enabled on `SW-COR-01` (Core Switch).
* **Security & Port Hardening**:
  - Unused ports: Admin down (`shutdown`) and assigned to a dummy quarantine VLAN (VLAN 999).
  - Port Security: MAC address sticky or limit of 2 MAC addresses on access ports.
  - DHCP Snooping: Enabled on all access ports to block rogue DHCP servers.
