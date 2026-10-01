# 03_CAD_Exports_DXF_SVG

This directory contains categorized, production-grade AutoCAD DXF and SVG engineering drawing deliverables for the **Kalasin Starch Plant - Jet Cooker Project** (xCIP System, Project Ref: `x2608003`).

---

## Directory Structure & Drawing Inventory

```
03_CAD_Exports_DXF_SVG/
│
├── DI_Slot_Drawings/                     # 15 Digital Input (1756-IB32) Wiring Diagrams (One Slot per DXF)
│   ├── C1S4_DI_Wiring_Diagram.dxf
│   ├── C1S5_DI_Wiring_Diagram.dxf
│   ├── C1S6_DI_Wiring_Diagram.dxf
│   ├── C1S7_DI_Wiring_Diagram.dxf
│   ├── C2S1_DI_Wiring_Diagram.dxf
│   ├── C2S2_DI_Wiring_Diagram.dxf
│   ├── C2S3_DI_Wiring_Diagram.dxf
│   ├── C2S4_DI_Wiring_Diagram.dxf
│   ├── C2S5_DI_Wiring_Diagram.dxf
│   ├── C3S1_DI_Wiring_Diagram.dxf
│   ├── C3S2_DI_Wiring_Diagram.dxf
│   ├── C3S3_DI_Wiring_Diagram.dxf
│   ├── C4S1_DI_Wiring_Diagram.dxf
│   ├── C4S2_DI_Wiring_Diagram.dxf
│   └── C4S3_DI_Wiring_Diagram.dxf
│
├── DO_Slot_Drawings/                     # 9 Digital Output (1756-OB32 with SPDT Relay) Wiring Diagrams
│   ├── C1S8_DO_Wiring_Diagram.dxf
│   ├── C1S9_DO_Wiring_Diagram.dxf
│   ├── C2S6_DO_Wiring_Diagram.dxf
│   ├── C2S7_DO_Wiring_Diagram.dxf
│   ├── C2S8_DO_Wiring_Diagram.dxf
│   ├── C3S4_DO_Wiring_Diagram.dxf
│   ├── C3S5_DO_Wiring_Diagram.dxf
│   ├── C4S4_DO_Wiring_Diagram.dxf
│   └── C5S3_DO_Wiring_Diagram.dxf
│
├── 01_PLC_IO_Wiring_DXF/                 # 40 PLC I/O Modular Wiring Schematics (DXF)
│   ├── KAL-JC-WIR-00001_Sheet01_Cover.dxf
│   ├── KAL-JC-WIR-00002_Sheet02_Index.dxf
│   ├── KAL-JC-WIR-20001..20037 (Chassis C1-C5 per-slot schematics)
│   └── Jet_Cooker_PLC_IO_Wiring_Master_1to1.dxf
│
├── 02_Instrument_Typical_DXF/            # 26 Instrument Hookup & Installation Typicals (DXF)
│   ├── KAL-JC-TYP-000_Sheet01_Cover.dxf
│   ├── KAL-JC-TYP-000-IDX_Sheet02_Index.dxf
│   ├── KAL-JC-TYP-001..023 (Typical Hookups: Flowmeters, Transmitters, Valves)
│   └── Instrument_Hookup_and_Wiring_Templates_Master_1to1.dxf
│
├── 03_MCC_Panel_Layout/                  # 14 MCC Motor Control Center Panel Layouts (DXF + SVG)
│   ├── MCC_Panel_Layout_Master_1to1.dxf
│   ├── MCC_Panel_Layout_Sheet1_Front_Elevation (DXF & SVG)
│   ├── MCC_Panel_Layout_Sheet2_Internal_GA (DXF & SVG)
│   ├── MCC_Panel_Layout_Sheet3..6_Panel_M1..M4 (DXF & SVG)
│   └── MCC_Panel_Layout_Sheet7_BOM.svg
│
├── 04_System_Architecture_CAD/           # 4 System & Network Architecture Diagrams
│   ├── C1_C8_Slot_to_Destination_JB_Architecture_Diagram.dxf
│   ├── C1_C8_Slot_to_Destination_JB_Architecture_Diagram.svg
│   ├── C1_C8_Slot_to_Destination_JB_Architecture_Diagram.pdf
│   └── KAL-JC-SCD-10001_System_Configuration.dxf
│
└── 05_SVG_Vector_Sheets/                 # 73 Vectorized SVG Presentation & Review Sheets
    ├── Cable_Schedule_Slides/            # 8 Cable Schedule Presentation Slides
    │   └── Cable_Schedule_Slide_1..8.svg
    ├── Instrument_Templates/             # 25 Instrument Installation Vector Sheets
    │   └── Instrument_Template_Sheet_01..25.svg
    └── Wiring_Sheets/                    # 40 Jet Cooker Wiring Review Sheets
        └── Jet_Cooker_Wiring_Sheet_01..40.svg
```

---

## Deliverable Subfolder Breakdown

| Subfolder | Count | Format | Primary Source / Purpose |
| :--- | :---: | :---: | :--- |
| **`DI_Slot_Drawings/`** | 15 | `.dxf` | Individual DI drawings generated from `template_DI.dxf` and `IO_List-By_SlotConfig.xlsx` |
| **`DO_Slot_Drawings/`** | 9 | `.dxf` | Individual DO drawings (with SPDT Relay) generated from `template_DO.dxf` and `IO_List-By_SlotConfig.xlsx` |
| **`01_PLC_IO_Wiring_DXF/`** | 40 | `.dxf` | Full ControlLogix 1756 PLC I/O wiring sheets (`KAL-JC-WIR-*`) |
| **`02_Instrument_Typical_DXF/`** | 26 | `.dxf` | Field instrument installation and hookup typical drawings (`KAL-JC-TYP-*`) |
| **`03_MCC_Panel_Layout/`** | 14 | `.dxf`, `.svg` | Front elevation, internal GA, and panel layout sheets M1-M4 |
| **`04_System_Architecture_CAD/`** | 4 | `.dxf`, `.svg`, `.pdf` | Chassis C1-C8 slot-to-JB destination architecture and system configuration |
| **`05_SVG_Vector_Sheets/`** | 73 | `.svg` | High-fidelity vector graphical sheets for slide decks, web viewing, and documentation |
