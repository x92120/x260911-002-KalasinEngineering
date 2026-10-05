# 03_Switch_and_Hardware_Config — Switch Configurations & Hardware Setup

This directory contains production-ready Cisco IOS and Rockwell Stratix managed switch configurations, along with server NIC teaming guidelines for the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📄 Key Production Configurations & Specifications

* 🛠️ **[CISCO_REDUNDANT_CORE_SWITCH_CONFIG.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/CISCO_REDUNDANT_CORE_SWITCH_CONFIG.md)**: Production Cisco Catalyst 9300 HSRP v2, LACP EtherChannel, and VLAN Trunk configuration scripts.
* 🛠️ **[MCC_TRUNK_AND_SWITCH_CONFIG.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/MCC_TRUNK_AND_SWITCH_CONFIG.md)**: Stratix 5700 802.1Q trunking configuration for the Motor Control Center (MCC) switch cabinet.
* 💻 **[SERVER_REDUNDANCY_AND_NIC_TEAMING_CONFIG.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/SERVER_REDUNDANCY_AND_NIC_TEAMING_CONFIG.md)**: Configuration guide for FT View SE Primary/Secondary SCADA Server NIC teaming and network failover.
* 📋 **[SWITCH_HARDWARE_SPEC.md](file:///e:/xApp-01/x260911-002-KalasinEngineering/01_Network_Design/03_Switch_and_Hardware_Config/SWITCH_HARDWARE_SPEC.md)**: Hardware specification sheet for Cisco Catalyst 9300 and Stratix 5700/5400 switches.

---

## 🔒 Security & Redundancy Features

1. **HSRP v2 Core Failover**: Sub-second virtual router failover between Active and Standby Catalyst 9300 core switches.
2. **802.1Q Trunking**: Standardized VLAN tagging across fiber optic uplinks between MCP-01 and field cabinets.
3. **Port Security & CIP Sync**: Optimized QoS configuration prioritizing PTP (Precision Time Protocol / IEEE 1588) and CIP I/O packets.
