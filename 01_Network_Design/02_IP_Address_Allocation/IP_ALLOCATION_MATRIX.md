# Plant-Wide IP Address Allocation Matrix
**Project**: Ingredion Kalasin — Sprint 18K TPA Spray Dryer Plant  
**Subnet Mask (Class C Standard)**: `255.255.255.0` (`/24`)  

---

## 1. 5-Chassis ControlLogix DLR Ring & Controller Allocation (VLAN 30: `192.168.30.0/24`)

| IP Address | Device Tag | Chassis / Slot | Hardware Model | DLR / Network Role | Panel Location |
|---|---|---|---|---|---|
| `192.168.30.1`  | `GW-COR-01`    | Core Switch L3 SVI | Stratix 5400 | Plant Default Gateway | Server Rack |
| `192.168.30.10` | `PLC-01-CPU`   | CS1 Slot 00 | 1756-L950TPSXT | Master Process Safety Controller (XT) | Main Control Panel (MCP-01) |
| `192.168.30.11` | `CS1-EN4TR-DLR`| CS1 Slot 01 | 1756-EN4TR | **Active Ring Supervisor (Beacon Port 2)** | Main Control Panel (MCP-01) |
| `192.168.30.12` | `CS2-EN4TR`    | CS2 Slot 00 | 1756-EN4TR | Remote I/O Adapter (Ring Node 1) | Main Control Panel (MCP-01) |
| `192.168.30.13` | `CS3-EN4TR`    | CS3 Slot 00 | 1756-EN4TR | Remote I/O Adapter (Ring Node 2) | Main Control Panel (MCP-01) |
| `192.168.30.14` | `CS4-EN4TR`    | CS4 Slot 00 | 1756-EN4TR | Remote I/O Adapter (Ring Node 3) | Main Control Panel (MCP-01) |
| `192.168.30.15` | `CS5-EN4TR`    | CS5 Slot 00 | 1756-EN4TR | Remote I/O Adapter (Ring Node 4 via Fiber) | Remote Panel (RIO200) |
| `192.168.30.20` | `PLC-01-EN4T2` | CS1 Slot 02 | 1756-EN4TR | Plant SCADA / FactoryTalk Uplink | Main Control Panel (MCP-01) |
| `192.168.30.250`| `MC-01-MCP`    | Din-Rail MCP-01 | 1783-ETAP2F / Stratix | Fiber DLR Tap / Transceiver (MCP side) | Main Control Panel (MCP-01) |
| `192.168.30.251`| `MC-02-RIO`    | Din-Rail RIO200 | 1783-ETAP2F / Stratix | Fiber DLR Tap / Transceiver (RIO side) | Remote Panel (RIO200) |
| `192.168.30.50` | `HMI-MCP-01`   | Door Panel | PanelView 5510 15" | Operator Touchscreen HMI | Main Control Panel (MCP-01) |

---

## 2. Area 202 & 402 Pre-Slurry & Adjustment (VLAN 40: `192.168.40.0/24`)

| IP Address | Device Tag | Device Description | Connected Equipment | Hardware Model |
|---|---|---|---|---|
| `192.168.40.1` | `GW-SW201` | Gateway / Switch SVI | Area 202 RIO Panel (RIO-202) | Stratix 5700 |
| `192.168.40.10`| `RIO-202-AENT`| Remote I/O Adapter 1 (Pre-Slurry TS2020) | `XV-20202`, `XV-20203`, `LT-20201` | 1734-AENTR |
| `192.168.40.11`| `RIO-202-AENT2`| Remote I/O Adapter 2 (Pre-Slurry TS2021) | `XV-20212`, `XV-20213`, `LT-20211` | 1734-AENTR |
| `192.168.40.21`| `VFD-PC-20201` | Slurry Transfer Pump 1 VFD (18.5 kW) | Slurry Pump `PC-20201` | PowerFlex 525 |
| `192.168.40.22`| `VFD-PC-20211` | Slurry Transfer Pump 2 VFD (18.5 kW) | Slurry Pump `PC-20211` | PowerFlex 525 |
| `192.168.40.31`| `VFD-ME-20201` | Slurry Agitator VFD (15 kW) | Slurry Agitator `ME-20201` | PowerFlex 525 |
| `192.168.40.32`| `VFD-ME-20211` | Slurry Agitator VFD (15 kW) | Slurry Agitator `ME-20211` | PowerFlex 525 |
| `192.168.40.50`| `FM-FT-20201` | Coriolis Mass Flowmeter | Transfer Flow to Area 402 | E+H Promass 300 |

---

## 3. Area 602 Spray Dryer & Atomization (VLAN 60: `192.168.60.0/24`)

| IP Address | Device Tag | Device Description | Connected Equipment | Hardware Model |
|---|---|---|---|---|
| `192.168.60.1` | `GW-SW601` | Gateway / Switch SVI | Dryer Main Panel (MCP-602) | Stratix 5700 |
| `192.168.60.10`| `RIO-602-AENT1`| Dryer Chamber RIO Rack (DLR Node 1) | Burner permissives, Temp transmitters | 1734-AENTR |
| `192.168.60.11`| `RIO-602-AENT2`| Exhaust & Cyclone RIO Rack (DLR Node 2)| Cyclone dampers, Baghouse pulses | 1734-AENTR |
| `192.168.60.20`| `VFD-PUPD-11`  | High Pressure Feed Pump (75 kW) | HP Atomizer Feed `PUPD-11` | PowerFlex 755 |
| `192.168.60.21`| `VFD-FN-60201` | Main Supply Air Fan VFD (110 kW) | Combustion Air Fan `FN-60201` | PowerFlex 755 |
| `192.168.60.22`| `VFD-FN-60202` | Exhaust Air Fan VFD (132 kW) | Main Exhaust Fan `FN-60202` | PowerFlex 755 |

---

## 4. Area 814 Plant CIP Skid (VLAN 80: `192.168.80.0/24`)

| IP Address | Device Tag | Device Description | Connected Equipment | Hardware Model |
|---|---|---|---|---|
| `192.168.80.1` | `GW-SW801` | Gateway / Switch SVI | CIP Skid Panel (CIP-01) | Stratix 5700 |
| `192.168.80.10`| `RIO-814-AENT` | CIP Skid Remote I/O Adapter | Routing Valves `XV-814xx` | 1734-AENTR |
| `192.168.80.20`| `VFD-PC-81401` | CIP Forward Supply Pump (22 kW) | CIP Supply `PC-81401` | PowerFlex 525 |
| `192.168.80.21`| `VFD-PC-81402` | CIP Scavenge Return Pump (15 kW) | CIP Return `PC-81402` | PowerFlex 525 |
| `192.168.80.30`| `CT-81401`     | CIP Conductivity Transmitter | Caustic/Acid return rinse | E+H Smartec CLD18 |
