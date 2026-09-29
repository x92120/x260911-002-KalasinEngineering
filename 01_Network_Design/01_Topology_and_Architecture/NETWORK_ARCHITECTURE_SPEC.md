# Industrial Network Architecture Specification
**Facility**: Ingredion Kalasin Spray Dryer Plant  
**Architecture Framework**: Cisco / Rockwell Automation Converged Plantwide Ethernet (CPwE)  

---

## 1. Network Topology Overview

The plant automation network employs a **Hierarchical Star with Redundant Device-Level Rings (DLR)**:

```
                  ┌────────────────────────────────────────┐
                  │ Level 3: FactoryTalk SCADA & Historian │
                  └───────────────────┬────────────────────┘
                                      │ (VLAN Trunk / 1Gbps Fiber)
                  ┌───────────────────┴────────────────────┐
                  │ Main Plant Distribution Switch (Core) │
                  │     Stratix 5400 / Cisco Catalyst      │
                  └─────────┬────────────────────┬─────────┘
                            │                    │ (Fiber Trunk / REP or RSTP)
            ┌───────────────┴────────┐  ┌────────┴───────────────┐
            │ Area 202/402 Switch    │  │ Area 602/814 Switch    │
            │ Stratix 5700 (SW-201)  │  │ Stratix 5700 (SW-601)  │
            └───┬────────────────────┘  └───┬────────────────────┘
                │                           │
       ┌────────┴────────┐         ┌────────┴────────┐
       │ DLR Ring 1      │         │ DLR Ring 2      │
       │ (Area 202/402)  │         │ (Area 602 Core) │
       │ ControlLogix    │         │ ControlLogix    │
       │ 1756-EN2TR      │         │ 1756-EN4TR      │
       │ 1734-AENTR I/O  │         │ PowerFlex VFDs  │
       │ PowerFlex 525   │         │ Remote Racks    │
       └─────────────────┘         └─────────────────┘
```

---

## 2. Network Segmentation & VLAN Definition

| VLAN ID | VLAN Name | Purpose | Subnet CIDR | Default Gateway |
|:---:|---|---|---|---|
| **VLAN 10** | `VLAN_SCADA_SRV` | FactoryTalk SE Servers, Historian, Batch Server | `192.168.10.0/24` | `192.168.10.1` |
| **VLAN 20** | `VLAN_HMI_CLIENT` | PanelView 5510, Area ThinManager Thin Clients, EWS | `192.168.20.0/24` | `192.168.20.1` |
| **VLAN 30** | `VLAN_PLC_CONTROL` | ControlLogix 5580 Process & Safety Controller Racks | `192.168.30.0/24` | `192.168.30.1` |
| **VLAN 40** | `VLAN_IO_AREA202` | Area 202 Pre-Slurry Remote I/O & Pumps (DLR 1) | `192.168.40.0/24` | `192.168.40.1` |
| **VLAN 42** | `VLAN_IO_AREA402` | Area 402 Adjustment Remote I/O & Valves | `192.168.42.0/24` | `192.168.42.1` |
| **VLAN 60** | `VLAN_IO_AREA602` | Area 602 Spray Dryer Burner & Core RIO (DLR 2) | `192.168.60.0/24` | `192.168.60.1` |
| **VLAN 80** | `VLAN_IO_AREA814` | Area 814 CIP Skid RIO & Dosing Pumps | `192.168.80.0/24` | `192.168.80.1` |
| **VLAN 99** | `VLAN_MGMT_SWITCH` | Industrial Switch Management (SSH/SNMP/HTTPS) | `192.168.99.0/24` | `192.168.99.1` |

---

## 3. Industrial Protocols & QoS Traffic Prioritization

* **CIP (Common Industrial Protocol) / EtherNet/IP**:
  - Class 1 Implicit I/O: UDP Port 2222 (Highest priority, DSCP 47/59 - Expedited Forwarding).
  - Class 3 Explicit Messaging: TCP Port 44818 (Medium priority, DSCP 0 - Best Effort).
  - PTP (IEEE 1588 / CIP Sync): Synchronization across motion & sequence timestamps.
* **IGMP Snooping & Querying**: Enabled on all Stratix switches to prevent EtherNet/IP multicast flooding.
