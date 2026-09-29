# Cisco Redundant Core Switch Configuration Specification
**Facility**: Ingredion Kalasin Spray Dryer Plant  
**Architecture**: Cisco / Rockwell Automation CPwE (Converged Plantwide Ethernet)  
**Hardware Platform**: Redundant Cisco Catalyst 9300-24T / 9300-48T (or Cisco Industrial Ethernet IE-5000 / Stratix 5400)  
**Redundancy Protocols**: HSRP (Hot Standby Router Protocol v2), LACP (IEEE 802.3ad Port-Channel), Rapid-PVST+ (IEEE 802.1w)  

---

## 1. High Availability & Redundancy Architecture

```text
               ┌────────────────────────────────────────────────────────┐
               │              REDUNDANT SERVER INFRASTRUCTURE           │
               │  [FT View SE Server 01 (Pri)]   [FT View SE Server 02 (Sec)]   │
               │   NIC 1 (Active)  NIC 2 (Standby) NIC 1 (Standby) NIC 2 (Active)│
               └─────────┬───────────────┬──────────────┬───────────────┬───────┘
                         │ (LACP Teaming)│              │               │
                         ▼               │              ▼               │
               ┌──────────────────┐      │     ┌──────────────────┐     │
               │  SW-CORE-01      │◄─────┼────►│  SW-CORE-02      │◄────┘
               │  Cisco Core SW 1 │      └────►│  Cisco Core SW 2 │
               │  (HSRP Priority) │            │  (HSRP Standby)  │
               └────────┬─────────┘            └────────┬─────────┘
                        │                               │
   ┌────────────────────┴───────────────────────────────┴────────────────────┐
   │                                                                         │
   ▼ (Fiber Optic Backbone Trunk)                                            ▼ (Copper/Fiber 802.1Q Trunk)
┌─────────────────────────────────┐                       ┌─────────────────────────────────┐
│ EXISTING SERVER ROOM            │                       │ MCC ROOM (MOTOR CONTROL CENTER) │
│ • FDF-SR01 Fiber Patch Panel    │                       │ • SW-MCC-01 (Stratix / Cisco IE)│
│ • Domain Controllers DC01/DC02  │                       │ • PowerFlex 755/525 VFDs (VLAN50)│
│ • Historian / Batch Archives    │                       │ • Power Meter 1408-BC3A         │
└─────────────────────────────────┘                       └─────────────────────────────────┘
```

---

## 2. IP Addressing & HSRP Default Gateway Plan

| VLAN ID | VLAN Name | Subnet / CIDR | Virtual IP (HSRP VIP) | SW-CORE-01 (Active) | SW-CORE-02 (Standby) |
|:---:|---|---|:---:|:---:|:---:|
| **VLAN 10** | `VLAN_SCADA_SRV` | `192.168.10.0/24` | `192.168.10.1` (VIP) | `192.168.10.2` (Pri: 110) | `192.168.10.3` (Pri: 100) |
| **VLAN 20** | `VLAN_HMI_CLIENT`| `192.168.20.0/24` | `192.168.20.1` (VIP) | `192.168.20.2` (Pri: 110) | `192.168.20.3` (Pri: 100) |
| **VLAN 30** | `VLAN_PLC_CONTROL`| `192.168.30.0/24`| `192.168.30.1` (VIP) | `192.168.30.2` (Pri: 110) | `192.168.30.3` (Pri: 100) |
| **VLAN 50** | `VLAN_MCC_DRIVES` | `192.168.50.0/24` | `192.168.50.1` (VIP) | `192.168.50.2` (Pri: 110) | `192.168.50.3` (Pri: 100) |
| **VLAN 99** | `VLAN_MGMT`       | `192.168.99.0/24` | -                   | `192.168.99.2`            | `192.168.99.3`            |

---

## 3. Cisco SW-CORE-01 Configuration (Primary Core Switch)

```cisco
! ==============================================================================
! CISCO IOS-XE CONFIGURATION: SW-CORE-01 (PRIMARY / ACTIVE HSRP GATEWAY)
! ==============================================================================
hostname SW-CORE-01
no ip domain-lookup
ip domain-name ingredion.local

! --- 1. Global Services & Spanning Tree ---
spanning-tree mode rapid-pvst
spanning-tree portfast default
spanning-tree portfast bpduguard default
! Set SW-CORE-01 as primary STP root for all VLANs
spanning-tree vlan 1-99 root primary

! --- 2. VLAN Definitions ---
vlan 10
 name VLAN_SCADA_SRV
!
vlan 20
 name VLAN_HMI_CLIENT
!
vlan 30
 name VLAN_PLC_CONTROL
!
vlan 50
 name VLAN_MCC_DRIVES
!
vlan 99
 name VLAN_MGMT
!
vlan 999
 name VLAN_QUARANTINE_UNUSED
!

! --- 3. Inter-Switch Link (ISL Cross-Connect: Port-Channel 1) ---
! Dual 10Gbps or 1Gbps bonded links between SW-CORE-01 and SW-CORE-02
interface range TenGigabitEthernet1/0/1 - 2
 description ISL-TRUNK-TO-SW-CORE-02
 switchport mode trunk
 switchport trunk encapsulation dot1q
 switchport trunk allowed vlan 10,20,30,50,99
 channel-group 1 mode active
 no shutdown
!
interface Port-channel 1
 description ISL-PORT-CHANNEL-TRUNK
 switchport mode trunk
 switchport trunk allowed vlan 10,20,30,50,99
 no shutdown
!

! --- 4. Redundant Server Dual-Homed Ports (NIC Teaming / LACP) ---
! Primary SCADA Server 01 (Port-Channel 10)
interface GigabitEthernet1/0/1
 description FT-VIEW-SERVER-01_NIC1
 switchport mode access
 switchport access vlan 10
 channel-group 10 mode active
 spanning-tree portfast
 no shutdown
!
interface Port-channel 10
 description FT-VIEW-SERVER-01_TEAM_A
 switchport mode access
 switchport access vlan 10
 spanning-tree portfast
 no shutdown
!
! Secondary SCADA Server 02 (Cross-connected port)
interface GigabitEthernet1/0/2
 description FT-VIEW-SERVER-02_NIC2
 switchport mode access
 switchport access vlan 10
 channel-group 11 mode active
 spanning-tree portfast
 no shutdown
!
interface Port-channel 11
 description FT-VIEW-SERVER-02_TEAM_A
 switchport mode access
 switchport access vlan 10
 spanning-tree portfast
 no shutdown
!

! --- 5. Fiber Uplink to Existing Server Room (SFP Module) ---
interface TenGigabitEthernet1/1/1
 description FIBER-UPLINK-TO-SERVER-ROOM-FDF01
 switchport mode trunk
 switchport trunk allowed vlan 10,20,99
 speed nonegotiate
 no shutdown
!

! --- 6. Trunk to MCC Room (PowerFlex VFDs & Motor Starters) ---
interface GigabitEthernet1/0/24
 description TRUNK-TO-MCC-ROOM-SW-MCC-01
 switchport mode trunk
 switchport trunk allowed vlan 30,50,99
 switchport trunk native vlan 99
 spanning-tree portfast trunk
 no shutdown
!

! --- 7. Uplink to Main Control Panel (MCP-01 ControlLogix CS1) ---
interface GigabitEthernet1/0/12
 description UPLINK-TO-MCP01-CS1-EN4TR-S02
 switchport mode access
 switchport access vlan 30
 spanning-tree portfast
 no shutdown
!

! --- 8. Layer 3 SVI & HSRP v2 Gateway Configuration ---
ip routing
!
interface Vlan10
 description SVI_SCADA_SERVERS
 ip address 192.168.10.2 255.255.255.0
 standby version 2
 standby 10 ip 192.168.10.1
 standby 10 priority 110
 standby 10 preempt
 standby 10 authentication md5 key-string IngredionKalasin#10
 no shutdown
!
interface Vlan20
 description SVI_HMI_CLIENTS
 ip address 192.168.20.2 255.255.255.0
 standby version 2
 standby 20 ip 192.168.20.1
 standby 20 priority 110
 standby 20 preempt
 standby 20 authentication md5 key-string IngredionKalasin#20
 no shutdown
!
interface Vlan30
 description SVI_PLC_CONTROL
 ip address 192.168.30.2 255.255.255.0
 standby version 2
 standby 30 ip 192.168.30.1
 standby 30 priority 110
 standby 30 preempt
 standby 30 authentication md5 key-string IngredionKalasin#30
 no shutdown
!
interface Vlan50
 description SVI_MCC_DRIVES
 ip address 192.168.50.2 255.255.255.0
 standby version 2
 standby 50 ip 192.168.50.1
 standby 50 priority 110
 standby 50 preempt
 standby 50 authentication md5 key-string IngredionKalasin#50
 no shutdown
!
interface Vlan99
 description SVI_IN_BAND_MGMT
 ip address 192.168.99.2 255.255.255.0
 no shutdown
!

! --- 9. Industrial EtherNet/IP CIP Multicast Management ---
ip igmp snooping
ip igmp snooping querier
ip igmp snooping vlan 30 querier address 192.168.30.2
ip igmp snooping vlan 50 querier address 192.168.50.2
!
end
write memory
```

---

## 4. Cisco SW-CORE-02 Configuration (Secondary / Standby Core Switch)

```cisco
! ==============================================================================
! CISCO IOS-XE CONFIGURATION: SW-CORE-02 (SECONDARY / STANDBY HSRP GATEWAY)
! ==============================================================================
hostname SW-CORE-02
no ip domain-lookup
ip domain-name ingredion.local

! --- 1. Global Services & Spanning Tree ---
spanning-tree mode rapid-pvst
spanning-tree portfast default
spanning-tree portfast bpduguard default
! Set SW-CORE-02 as secondary STP root for fast failover
spanning-tree vlan 1-99 root secondary

! --- 2. VLAN Definitions ---
vlan 10
 name VLAN_SCADA_SRV
!
vlan 20
 name VLAN_HMI_CLIENT
!
vlan 30
 name VLAN_PLC_CONTROL
!
vlan 50
 name VLAN_MCC_DRIVES
!
vlan 99
 name VLAN_MGMT
!
vlan 999
 name VLAN_QUARANTINE_UNUSED
!

! --- 3. Inter-Switch Link (ISL Cross-Connect: Port-Channel 1) ---
interface range TenGigabitEthernet1/0/1 - 2
 description ISL-TRUNK-TO-SW-CORE-01
 switchport mode trunk
 switchport trunk encapsulation dot1q
 switchport trunk allowed vlan 10,20,30,50,99
 channel-group 1 mode active
 no shutdown
!
interface Port-channel 1
 description ISL-PORT-CHANNEL-TRUNK
 switchport mode trunk
 switchport trunk allowed vlan 10,20,30,50,99
 no shutdown
!

! --- 4. Redundant Server Dual-Homed Ports (NIC Teaming / LACP) ---
! Secondary connection for Server 01 (Port-Channel 10)
interface GigabitEthernet1/0/1
 description FT-VIEW-SERVER-01_NIC2
 switchport mode access
 switchport access vlan 10
 channel-group 10 mode active
 spanning-tree portfast
 no shutdown
!
interface Port-channel 10
 description FT-VIEW-SERVER-01_TEAM_B
 switchport mode access
 switchport access vlan 10
 spanning-tree portfast
 no shutdown
!
! Primary connection for Server 02 (Port-Channel 11)
interface GigabitEthernet1/0/2
 description FT-VIEW-SERVER-02_NIC1
 switchport mode access
 switchport access vlan 10
 channel-group 11 mode active
 spanning-tree portfast
 no shutdown
!
interface Port-channel 11
 description FT-VIEW-SERVER-02_TEAM_B
 switchport mode access
 switchport access vlan 10
 spanning-tree portfast
 no shutdown
!

! --- 5. Redundant Fiber Uplink to Server Room ---
interface TenGigabitEthernet1/1/1
 description FIBER-UPLINK-TO-SERVER-ROOM-FDF02
 switchport mode trunk
 switchport trunk allowed vlan 10,20,99
 speed nonegotiate
 no shutdown
!

! --- 6. Redundant Trunk to MCC Room ---
interface GigabitEthernet1/0/24
 description REDUNDANT-TRUNK-TO-MCC-ROOM
 switchport mode trunk
 switchport trunk allowed vlan 30,50,99
 switchport trunk native vlan 99
 spanning-tree portfast trunk
 no shutdown
!

! --- 7. Layer 3 SVI & HSRP v2 Gateway Configuration ---
ip routing
!
interface Vlan10
 description SVI_SCADA_SERVERS
 ip address 192.168.10.3 255.255.255.0
 standby version 2
 standby 10 ip 192.168.10.1
 standby 10 priority 100
 standby 10 preempt
 standby 10 authentication md5 key-string IngredionKalasin#10
 no shutdown
!
interface Vlan20
 description SVI_HMI_CLIENTS
 ip address 192.168.20.3 255.255.255.0
 standby version 2
 standby 20 ip 192.168.20.1
 standby 20 priority 100
 standby 20 preempt
 standby 20 authentication md5 key-string IngredionKalasin#20
 no shutdown
!
interface Vlan30
 description SVI_PLC_CONTROL
 ip address 192.168.30.3 255.255.255.0
 standby version 2
 standby 30 ip 192.168.30.1
 standby 30 priority 100
 standby 30 preempt
 standby 30 authentication md5 key-string IngredionKalasin#30
 no shutdown
!
interface Vlan50
 description SVI_MCC_DRIVES
 ip address 192.168.50.3 255.255.255.0
 standby version 2
 standby 50 ip 192.168.50.1
 standby 50 priority 100
 standby 50 preempt
 standby 50 authentication md5 key-string IngredionKalasin#50
 no shutdown
!
interface Vlan99
 description SVI_IN_BAND_MGMT
 ip address 192.168.99.3 255.255.255.0
 no shutdown
!

! --- 8. Industrial EtherNet/IP CIP Multicast Management ---
ip igmp snooping
ip igmp snooping querier
ip igmp snooping vlan 30 querier address 192.168.30.3
ip igmp snooping vlan 50 querier address 192.168.50.3
!
end
write memory
```
