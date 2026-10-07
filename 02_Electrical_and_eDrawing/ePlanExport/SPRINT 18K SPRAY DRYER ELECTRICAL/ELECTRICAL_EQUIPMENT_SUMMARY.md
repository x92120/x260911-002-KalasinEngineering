# SPRINT 18K SPRAY DRYER - ELECTRICAL EQUIPMENT & SYSTEM SUMMARY

**Drawing Location:** `x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL`  
**System Controller:** Allen-Bradley ControlLogix PLC System (5 I/O Chassis Racks)  

---

## 1. Power Distribution & Wiring Standards

### Power Distribution Components
| Sheet | Device Tag | Description / Rating |
|:---:|:---|:---|
| Sheet 3 | `TBFEED` | Main Power Feed Terminal Block |
| Sheet 3 | `CB1` | 32A Main Circuit Breaker |
| Sheet 3 | `CB2` - `CB5` | 6A Circuit Breakers for PLC Chassis AC Power Supplies (PA75.1 - PA75.4) |
| Sheet 3 | `CB6` - `CB9` | 6A Circuit Breakers for 24VDC Power Supplies (PS1 - PS3) |
| Sheet 3 | `PA75.1` - `PA75.4` | ControlLogix Standard AC Power Supply Modules |
| Sheet 3 | `PS1` - `PS3` | 24VDC 20A Industrial Power Supply Units |

### Standard Wiring Specifications
| Circuit Type | Signal Voltage | Wire Color | Wire Cross-Section |
|:---|:---:|:---:|:---:|
| 220VAC 1-Phase Line (L) | 220 VAC | Brown (BN) | 2.5 mm² |
| 220VAC 1-Phase Neutral (N) | 220 VAC | Blue (BL) | 2.5 mm² |
| 24VDC Power Supply (+) | +24 VDC | Red (RD) | 1.5 - 2.5 mm² |
| 24VDC Power Supply (-) | 0 VDC | Black (BK) | 1.5 - 2.5 mm² |
| 24VDC Signal (+) | +24 VDC | White (WH) | 0.75 mm² |
| 24VDC Signal (-) | 0 VDC | Black (BK) | 0.75 mm² |

---

## 2. PLC I/O System Rack & Slot Architecture

### Chassis Rack 1
| Slot | Module Function / Signal Type | Drawing Sheet | DXF File |
|:---:|:---|:---:|:---:|
| **Slot 3** | Analog Input | Sheet 4 | [4.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/4.dxf) |
| **Slot 4** | Analog Input | Sheet 5 | [5.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/5.dxf) |
| **Slot 5** | Analog Input | Sheet 6 | [6.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/6.dxf) |
| **Slot 6** | Analog input | Sheet 7 | [7.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/7.dxf) |
| **Slot 7** | Analog input | Sheet 8 | [8.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/8.dxf) |
| **Slot 8** | Analog input | Sheet 9 | [9.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/9.dxf) |
| **Slot 9** | Analog input | Sheet 10 | [10.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/10.dxf) |

### Chassis Rack 2
| Slot | Module Function / Signal Type | Drawing Sheet | DXF File |
|:---:|:---|:---:|:---:|
| **Slot 1** | Analog input | Sheet 11 | [11.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/11.dxf) |
| **Slot 2** | Analog Input | Sheet 12 | [12.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/12.dxf) |
| **Slot 3** | Analog Output | Sheet 13 | [13.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/13.dxf) |
| **Slot 4** | Digital Output | Sheet 14 | [14.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/14.dxf) |
| **Slot 5** | Digital Input | Sheet 15 | [15.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/15.dxf) |
| **Slot 6** | Digital Input | Sheet 16 | [16.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/16.dxf) |
| **Slot 7** | Digital Input | Sheet 17 | [17.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/17.dxf) |
| **Slot 8** | Digital input | Sheet 18 | [18.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/18.dxf) |
| **Slot 9** | Digital input | Sheet 19 | [19.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/19.dxf) |

### Chassis Rack 3
| Slot | Module Function / Signal Type | Drawing Sheet | DXF File |
|:---:|:---|:---:|:---:|
| **Slot 1** | Digital Input | Sheet 20 | [20.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/20.dxf) |
| **Slot 2** | Digital Input | Sheet 21 | [21.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/21.dxf) |
| **Slot 3** | Digital Input | Sheet 22 | [22.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/22.dxf) |
| **Slot 4** | Digital Input | Sheet 23 | [23.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/23.dxf) |
| **Slot 5** | Digital Input | Sheet 24 | [24.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/24.dxf) |
| **Slot 6** | Digital Input | Sheet 25 | [25.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/25.dxf) |
| **Slot 7** | Digital Input | Sheet 26 | [26.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/26.dxf) |
| **Slot 8** | Digital Input | Sheet 27 | [27.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/27.dxf) |
| **Slot 9** | Digital Input | Sheet 28 | [28.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/28.dxf) |

### Chassis Rack 4
| Slot | Module Function / Signal Type | Drawing Sheet | DXF File |
|:---:|:---|:---:|:---:|
| **Slot 1** | Digital Output | Sheet 29 | [29.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/29.dxf) |
| **Slot 2** | Digital Output | Sheet 30 | [30.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/30.dxf) |
| **Slot 3** | Digital Output | Sheet 31 | [31.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/31.dxf) |
| **Slot 4** | Digital Output | Sheet 32 | [32.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/32.dxf) |
| **Slot 5** | Digital Output | Sheet 33 | [33.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/33.dxf) |
| **Slot 6** | Digital Output | Sheet 34 | [34.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/34.dxf) |

### Chassis Rack 5
| Slot | Module Function / Signal Type | Drawing Sheet | DXF File |
|:---:|:---|:---:|:---:|
| **Slot 1** | Digital Input | Sheet 35 | [35.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/35.dxf) |
| **Slot 2** | Digital Input | Sheet 36 | [36.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/36.dxf) |
| **Slot 3** | Digital Output | Sheet 37 | [37.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/37.dxf) |
| **Slot 4** | Analog input | Sheet 38 | [38.dxf](file:///e:/xApp-01/x260911-002-Kalasin-Process/x900-eDrawing/ePlanExport/SPRINT 18K SPRAY DRYER ELECTRICAL/38.dxf) |

---

## 3. System Summary Statistics
- **Total I/O Chassis Racks:** 5 Racks
- **Analog Input (AI) Modules:** 9 Slots (Chassis 1: Slots 3-9, Chassis 2: Slots 1-2, Chassis 5: Slot 4)
- **Analog Output (AO) Modules:** 1 Slot (Chassis 2: Slot 3)
- **Digital Input (DI) Modules:** 18 Slots (Chassis 2: Slots 5-9, Chassis 3: Slots 1-9, Chassis 5: Slots 1-2)
- **Digital Output (DO) Modules:** 8 Slots (Chassis 2: Slot 4, Chassis 4: Slots 1-6, Chassis 5: Slot 3)
