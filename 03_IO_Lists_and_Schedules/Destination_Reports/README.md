# Destination_Reports — Per-Enclosure Field I/O Schedules

This folder contains 16 individual Excel workbooks corresponding to each physical field destination enclosure (Junction Boxes and MCC Panel) at the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📁 Enclosure Workbooks Register (16 Destination Workbooks)

| Excel File Name | Destination Enclosure Tag | Point Count | Location / Floor | Primary Signals / Equipment |
| :--- | :--- | :---: | :--- | :--- |
| **`MCC_IO_Schedule.xlsx`** | MCC Panel | 237 | MCC Room (Ground) | Motor Feeders, VFD Auxiliary & Bus |
| **`JB-607_IO_Schedule.xlsx`** | JB-607 | 195 | Spray Dryer Tower 7F | Chamber Temp/Pressure & Exhaust Fans |
| **`JB-401_IO_Schedule.xlsx`** | JB-401 | 112 | Infeed Area 2F | Slurry Feed Pumps & Tank Agitators |
| **`RIO-200_IO_Schedule.xlsx`** | RIO-200 | 112 | Slurry Skid 2F | Remote I/O Skid Drop |
| **`JB-601_IO_Schedule.xlsx`** | JB-601 | 85 | Spray Dryer 1F | Powder Discharge & Fluid Bed Valves |
| **`JB-602_IO_Schedule.xlsx`** | JB-602 | 82 | Spray Dryer 3F | Air Heater & Burner Interlocks |
| **`JB-402_IO_Schedule.xlsx`** | JB-402 | 65 | Jet Cooker 2F | Jet Cooker Steam Control & Temp |
| **`JB-612_IO_Schedule.xlsx`** | JB-612 | 57 | Packing Tower 2F | Bagging Machine & Conveyor I/O |
| **`JB-618_IO_Schedule.xlsx`** | JB-618 | 52 | Packing Tower 8F | Top Elevator & Dust Extraction |
| **`IS-JB-618_IO_Schedule.xlsx`**| IS-JB-618 | 44 | Packing Tower 8F (Ex i) | Hazardous Dust Area NAMUR Switches |
| **`IS-JB-612_IO_Schedule.xlsx`**| IS-JB-612 | 41 | Packing Tower 2F (Ex i) | Hazardous Dust Area NAMUR Switches |
| **`CA1_IO_Schedule.xlsx`** | CA1 | 32 | Main Control Room | PLC Rack C1 Local Interlocks |
| **`IS-JB-603_IO_Schedule.xlsx`**| IS-JB-603 | 32 | Spray Dryer 3F (Ex i) | Hazardous Area Pressure/Temp Transmitters |
| **`IS-JB-608_IO_Schedule.xlsx`**| IS-JB-608 | 32 | Spray Dryer 8F (Ex i) | Hazardous Area Pressure/Temp Transmitters |
| **`JB-606_IO_Schedule.xlsx`** | JB-606 | 26 | Spray Dryer 6F | Atomizer Lube Oil & Vibration |
| **`JB-608_IO_Schedule.xlsx`** | JB-608 | 6 | Spray Dryer 8F | High Temp Thermocouple Junction |

---

## ⚙️ Generation Script

These individual destination workbooks are generated automatically by `04_Automation_Scripts/build_destination_io_report.py`.
