# Ingredion Kalasin Plant — Project SPRINT
## SCADA & PlantPAx DCS Simulation Video & Interactive HMI Viewers

### 📹 Video & Interactive Web Deliverables
* **Video File:** `Kalasin_Plant_SCADA_Demo_Project_SPRINT_Rev3.5.mp4` (Full HD 1080p, 60s H.264)
* 🌐 **Interactive SCADA Dashboard**: [scada_process_dashboard.html](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/06_SCADA_Demo/scada_process_dashboard.html)
* 🌐 **Automation Architecture Presentation**: [automation_architecture_slides.html](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/06_SCADA_Demo/automation_architecture_slides.html)
* 🌐 **Light Theme Presentation**: [automation_architecture_slides_light.html](file:///e:/xApp-01/x260911-002-KalasinEngineering/02_Electrical_and_eDrawing/06_SCADA_Demo/automation_architecture_slides_light.html)

---

### ⏱️ Video Scene Breakdown & Timeline

| Timecode | Scene Description | Highlighted SCADA / PlantPAx Functionality | Snapshot |
| :--- | :--- | :--- | :---: |
| **00:00 – 00:08** | **Plant Overview & KPI Dashboard** | Overall process flow (Area 202 ➔ 402 ➔ 602 ➔ 613 ➔ Utilities), animated pipe flow pulses, daily tonnage progress gauge (38.4 / 50 T). | `01_Plant_Overview_HMI.png` |
| **00:08 – 00:18** | **Area 202 & 402: Slurry Prep & Jet Cooker** | Tanks `TS-20201/02`, rotating agitators `AG-20201/02`, transfer pump `PC-20201` running, pH meter `AT-40201` (6.25 pH), and Jet Cooker steam injection (105.2 °C, 4.3 barg). | `02_Area202_402_Slurry_JetCooker.png` |
| **00:18 – 00:29** | **Area 602: Spray Dryer Tower & HP Skid** | Dedert Triplex HP Pump `PD-60202` running at 282.5 barg, drying chamber `CHAM-21` (16 atomizing nozzles, 195.4 °C inlet / 92.2 °C outlet), Air Heater BMS, and Baghouse `DTEX-21`. | `03_Area602_Spray_Dryer_Tower.png` |
| **00:29 – 00:39** | **Interactive PlantPAx Faceplate Demo** | Operator clicks on Slurry Infeed Valve `FCV-60201`. Shows authentic PlantPAx faceplate window, adjusts Setpoint (40 ➔ 44 m³/h), dynamic bar graph response, and interlock permissives. | `04_PlantPAx_Faceplate_Interaction.png` |
| **00:39 – 00:47** | **Live Alarm Warning & Operator Ack** | Filter cake loading triggers `DPT-60203` warning (>160 mmH2O). Flashing top alarm banner, blinking bell icon, operator cursor clicks `[ACK ALL]`, auto pulse-jet recovery sequence. | `05_Active_Alarm_Warning_Event.png` |
| **00:47 – 00:55** | **Real-Time Multi-Pen Process Trends** | Industrial high-contrast dark trend recorder with 4 live pens: Inlet Temp, Outlet Temp, Slurry Feed Flow, and HP Pump Discharge. Scanning time cursor with live tooltip callouts. | `06_RealTime_Process_Trends.png` |
| **00:55 – 01:00** | **Engineering Standards & Outro Card** | Compliance verification: Rockwell PlantPAx 5.x, ANSI/ISA-101 HMI, ANSI/ISA-18.2 Alarms, 472 Tag I/O infrastructure aligned with Rev 3.5 P&ID. | - |

---

### 🏭 Engineering Standards Compliance
1. **ANSI/ISA-101 High Performance HMI:** Ergonomic neutral gray background (`#D8DDE4`), contextual color coding (Green = running/healthy, Red = trip/alarm, Amber = warning).
2. **ANSI/ISA-18.2 Alarm Management:** Alarm state machine (Normal ➔ Active Unacknowledged ➔ Active Acknowledged ➔ Normal), time-stamped alarm banner.
3. **P&ID Rev 3.5 Alignment:** Covers 38 P&ID drawing sheets (`3-25-001-xxx-PD-xx`) and 472 physical I/O points across RIO-200, RIO-400, RIO-600, and RIO-MCC.
