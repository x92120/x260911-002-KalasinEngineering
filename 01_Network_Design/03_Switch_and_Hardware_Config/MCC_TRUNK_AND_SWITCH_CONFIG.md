# MCC Room Trunk & Switch Configuration Specification
**Facility**: Ingredion Kalasin Spray Dryer Plant  
**MCC Switch Tag**: `SW-MCC-01` (Allen-Bradley Stratix 5700 `1783-BMS10CGN` / Cisco IE-4000)  
**Location**: Motor Control Center (MCC) Room  
**Trunk Link**: Dual/Single 802.1Q VLAN Trunk from Central Cisco Core Switches  

---

## 1. MCC Room Network Architecture & Port Schedule

```text
               ┌─────────────────────────────────────────────────┐
               │ Central Cisco Core Switches (Operation Room)     │
               │ SW-CORE-01 (Gi1/0/24)    SW-CORE-02 (Gi1/0/24)  │
               └─────────┬───────────────────────────────┬───────┘
                         │                               │
                         ▼ (Trunk Cable 1)               ▼ (Trunk Cable 2 / Standby STP)
               ┌─────────────────────────────────────────────────┐
               │ SW-MCC-01: Industrial Managed Switch (MCC Room) │
               │ Stratix 5700 (1783-BMS10CGN) / IP: 192.168.50.10│
               └─────────┬───────────────────────────────┬───────┘
                         │                               │
        ┌────────────────┼───────────────────────────────┼────────────────┐
        ▼ (Port 3)       ▼ (Port 4)                      ▼ (Port 5)       ▼ (Port 7)
┌──────────────┐ ┌──────────────┐                 ┌──────────────┐ ┌──────────────┐
│ PowerFlex 755│ │ PowerFlex 755│                 │ PowerFlex 525│ │ Power Meter  │
│ 110kW Fan    │ │ 132kW Fan    │                 │ 18.5kW Pump  │ │ 1408-BC3A    │
│ IP: .50.21   │ │ IP: .50.22   │                 │ IP: .50.31   │ │ IP: .50.50   │
└──────────────┘ └──────────────┘                 └──────────────┘ └──────────────┘
```

---

## 2. Port Allocation Table (`SW-MCC-01`)

| Port No. | Connector | Type | Connected Device | VLAN | IP Address | Notes |
|:---:|:---:|:---:|---|:---:|:---:|---|
| **Port 1** | RJ45 | **TRUNK** | Cisco Core `SW-CORE-01` (Gi1/0/24) | 30, 50, 99 | - | **Main MCC Trunk Link** |
| **Port 2** | RJ45 | **TRUNK** | Cisco Core `SW-CORE-02` (Gi1/0/24) | 30, 50, 99 | - | **Redundant MCC Trunk (STP)** |
| **Port 3** | RJ45 | Access | PowerFlex 755 (Supply Fan FN-60201) | VLAN 50 | `192.168.50.21` | 110 kW Combustion Air |
| **Port 4** | RJ45 | Access | PowerFlex 755 (Exhaust Fan FN-60202)| VLAN 50 | `192.168.50.22` | 132 kW Main Exhaust |
| **Port 5** | RJ45 | Access | PowerFlex 525 (Slurry Pump PC-20201)| VLAN 50 | `192.168.50.31` | 18.5 kW Transfer Pump 1 |
| **Port 6** | RJ45 | Access | PowerFlex 525 (Slurry Pump PC-20211)| VLAN 50 | `192.168.50.32` | 18.5 kW Transfer Pump 2 |
| **Port 7** | RJ45 | Access | Power Meter `1408-BC3A` | VLAN 50 | `192.168.50.50` | Main Incoming Power Meter |
| **Port 8** | RJ45 | Access | Spare / Maintenance Laptop | VLAN 50 | DHCP | Local Diagnostics |

---

## 3. Stratix 5700 / Cisco IE Switch Configuration Script

```cisco
! ==============================================================================
! CONFIGURATION FOR SW-MCC-01 (MCC ROOM INDUSTRIAL SWITCH)
! ==============================================================================
hostname SW-MCC-01
no ip domain-lookup
ip domain-name ingredion.local

! --- 1. Spanning Tree & VLANs ---
spanning-tree mode rapid-pvst
spanning-tree portfast default
spanning-tree portfast bpduguard default

vlan 30
 name VLAN_PLC_CONTROL
!
vlan 50
 name VLAN_MCC_DRIVES
!
vlan 99
 name VLAN_MGMT
!

! --- 2. Main Trunk Ports to Cisco Core Switches ---
interface FastEthernet1/1
 description TRUNK-TO-SW-CORE-01
 switchport mode trunk
 switchport trunk allowed vlan 30,50,99
 switchport trunk native vlan 99
 no shutdown
!
interface FastEthernet1/2
 description REDUNDANT-TRUNK-TO-SW-CORE-02
 switchport mode trunk
 switchport trunk allowed vlan 30,50,99
 switchport trunk native vlan 99
 no shutdown
!

! --- 3. PowerFlex VFD Drive Ports (VLAN 50) ---
interface range FastEthernet1/3 - 7
 switchport mode access
 switchport access vlan 50
 spanning-tree portfast
 storm-control broadcast level 2.0
 storm-control multicast level 5.0
 no shutdown
!

! --- 4. Switch Management SVI & Default Gateway ---
interface Vlan50
 description SVI_MCC_MANAGEMENT
 ip address 192.168.50.10 255.255.255.0
 no shutdown
!
ip default-gateway 192.168.50.1

! --- 5. EtherNet/IP CIP Multicast Optimization ---
ip igmp snooping
ip igmp snooping vlan 50
!
end
write memory
```
