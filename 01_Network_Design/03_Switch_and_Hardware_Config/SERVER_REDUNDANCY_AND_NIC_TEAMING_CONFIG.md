# Server Redundancy & NIC Teaming Configuration Specification
**Facility**: Ingredion Kalasin Spray Dryer Plant  
**Applications**: Rockwell FactoryTalk View SE Distributed, FactoryTalk Historian SE, FactoryTalk Batch  
**OS Platform**: Microsoft Windows Server 2022 / 2025 Standard (Fault-Tolerant High-Availability Cluster)  

---

## 1. Physical Server Hardware & NIC Allocation

Each physical server (HPE ProLiant DL380 Gen10/Gen11 or Dell PowerEdge R750) is equipped with:
* **Dual Hot-Plug Redundant Power Supplies (1+1 PSUs)** connected to separate PDU feeds (UPS A & UPS B).
* **4x 10GbE/1GbE Network Interface Cards (NICs)**:
  - **NIC 1 (Plant LAN A)**: Connected to Cisco Core Switch `SW-CORE-01`.
  - **NIC 2 (Plant LAN B)**: Connected to Cisco Core Switch `SW-CORE-02`.
  - **NIC 3 (Dedicated Heartbeat / Inter-Server Sync)**: Direct high-speed cross-connect cable between Server 01 and Server 02.
  - **NIC 4 (iLO / iDRAC Out-of-Band Management)**: Connected to Management Switch (`VLAN 99`).

---

## 2. Windows Server NIC Teaming (LACP / Switch Independent)

### PowerShell Configuration for Switch-Independent Active/Standby Team:
```powershell
# Create NIC Team with Active / Standby failover across Cisco Switches
New-NetLbfoTeam -Name "ProductionTeam" `
                -TeamMembers "NIC1","NIC2" `
                -TeamingMode SwitchIndependent `
                -LoadBalancingAlgorithm Dynamic

# Configure Static IP Address on the Team Interface (VLAN 10)
# Example for Primary SCADA Server (SRV-SCADA-01):
New-NetIPAddress -InterfaceAlias "ProductionTeam" `
                 -IPAddress 192.168.10.10 `
                 -PrefixLength 24 `
                 -DefaultGateway 192.168.10.1

Set-DnsClientServerAddress -InterfaceAlias "ProductionTeam" `
                           -ServerAddresses ("192.168.10.30", "192.168.10.31")
```

---

## 3. Dedicated Inter-Server Heartbeat Link (Split-Brain Prevention)

To prevent "Split-Brain" (where both servers believe they are the active master during a network hiccup), a dedicated private link is established over NIC 3:

| Server Name | NIC 3 IP Address | Subnet Mask | Description |
|---|---|---|---|
| **SRV-SCADA-01 (Primary)**   | `10.10.10.1` | `255.255.255.252` (`/30`) | Direct Cat6A Patch to Server 02 NIC 3 |
| **SRV-SCADA-02 (Secondary)** | `10.10.10.2` | `255.255.255.252` (`/30`) | Direct Cat6A Patch to Server 01 NIC 3 |

---

## 4. Rockwell FactoryTalk View SE Redundancy Setup

1. **FactoryTalk Directory (FTD) Redundancy**:
   - `SRV-SCADA-01` configured as **Primary Directory Server**.
   - `SRV-SCADA-02` configured as **Secondary Directory Server**.
   - Replication interval: **Real-time synchronized** with automatic certificate renewal.

2. **HMI Tag & Data Server Redundancy**:
   - **Active/Standby Pair**: In Studio 5000 / FactoryTalk Administration Console:
     - Primary Server: `SRV-SCADA-01` (`192.168.10.10`)
     - Secondary Server: `SRV-SCADA-02` (`192.168.10.11`)
     - Heartbeat Interval: **500 ms**
     - Failover Switchover Time: **< 1.5 seconds**
   - **Auto-Synchronization**: Graphic display edits, tag database changes, and security accounts automatically synchronize from Primary to Secondary.

3. **FactoryTalk Alarms and Events (FTA&E)**:
   - Alarm status, acknowledgment, and suppression states are mirrored continuously across both servers.
   - Operators on any station (Operator 1, Operator 2, or Station 2) acknowledge alarms once, updating both servers simultaneously.

---

## 5. FactoryTalk Historian SE Collective Mirroring

* Primary Historian Node: `192.168.10.20`
* Secondary Historian (Mirror Node): `192.168.10.21`
* **Dual-Buffering Architecture**: ControlLogix PLC tags are sent concurrently to local buffers on both historian servers. If one server is taken down for Windows updates or hardware maintenance, zero trend data points are lost.
