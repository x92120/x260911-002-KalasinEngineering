# x9100-eDrawing — Individual Slot-by-Slot CAD Wiring Drawings

This directory contains individual AutoCAD `.dxf` single-slot wiring diagrams for **all 13 slots (Slots 0 to 12)** across Chassis `C1`, `C2`, `C3`, `C4`, and `C5`, formatted with explicit hardware module model tags (`CxSy-<CARD>.dxf`, e.g. `C1S4-IB32.dxf`), master multi-slot CAD compilations, Python generation scripts, and module templates for the **Ingredion Kalasin Spray Dryer Plant**.

---

## 📁 Group Subdirectories Index (`CxSy-<CARD>.dxf`)

* 📁 **[C1/](file:///e:/xApp-01/x260911-002-KalasinEngineering/x9100-eDrawing/C1)** — Chassis C1 Main Controller Rack Drawings (30 files: `C1S0-L950TPSXT` to `C1S12-N2` DI, DO, AI, CPU & COMM DXFs)
* 📁 **[C2/](file:///e:/xApp-01/x260911-002-KalasinEngineering/x9100-eDrawing/C2)** — Chassis C2 Process Expansion Rack Drawings (26 files: `C2S0-EN4TR` to `C2S12-IF16` COMM, DI, DO & AI DXFs)
* 📁 **[C3/](file:///e:/xApp-01/x260911-002-KalasinEngineering/x9100-eDrawing/C3)** — Chassis C3 Evaporator & Skid I/O Rack Drawings (26 files: `C3S0-EN4TR` to `C3S12-OF8` COMM, DI, DO, AI, Reserve & AO DXFs)
* 📁 **[C4/](file:///e:/xApp-01/x260911-002-KalasinEngineering/x9100-eDrawing/C4)** — Chassis C4 Spray Dryer & Exhaust Rack Drawings (26 files: `C4S0-EN4TR` to `C4S12-N2` COMM, DI, DO, AI & Reserve DXFs)
* 📁 **[C5/](file:///e:/xApp-01/x260911-002-KalasinEngineering/x9100-eDrawing/C5)** — Chassis C5 Remote I/O Skid RIO-200 Drawings (26 files: `C5S0-EN4TR` to `C5S12-N2` COMM, DI, DO, AI & Reserve DXFs)

---

## 📑 Complete 13-Slot Hardware Model Filename Schedule

| Chassis | Slot 0 | Slots 1–3 | Slots 4–7 | Slots 8–9 | Slots 10–11 | Slot 12 |
| :---: | :--- | :--- | :--- | :--- | :--- | :--- |
| **C1** | `C1S0-L950TPSXT.dxf` | `C1S1-EN4TR.dxf` .. `C1S3-EN4TR.dxf` | `C1S4-IB32.dxf` .. `C1S7-IB32.dxf` | `C1S8-OB32.dxf` .. `C1S9-OB32.dxf` | `C1S10-IF16.dxf` .. `C1S11-IF16.dxf` | `C1S12-N2.dxf` |
| **C2** | `C2S0-EN4TR.dxf` | `C2S1-IB32.dxf` .. `C2S3-IB32.dxf` | `C2S4-IB32.dxf` .. `C2S5-IB32.dxf` | `C2S6-OB32.dxf` .. `C2S8-OB32.dxf` | `C2S9-IF16.dxf` .. `C2S11-IF16.dxf` | `C2S12-IF16.dxf` |
| **C3** | `C3S0-EN4TR.dxf` | `C3S1-IB32.dxf` .. `C3S3-IB32.dxf` | `C3S4-OB32.dxf` .. `C3S5-OB32.dxf` | `C3S6-IF16.dxf` .. `C3S9-IF16.dxf` | `C3S10-N2.dxf` .. `C3S11-N2.dxf` | `C3S12-OF8.dxf` |
| **C4** | `C4S0-EN4TR.dxf` | `C4S1-IB32.dxf` .. `C4S3-IB32.dxf` | `C4S4-OB32.dxf` | `C4S5-IF16.dxf` .. `C4S6-IF16.dxf` | `C4S7-N2.dxf` .. `C4S11-N2.dxf` | `C4S12-N2.dxf` |
| **C5** | `C5S0-EN4TR.dxf` | `C5S1-IB32.dxf` .. `C5S2-IB32.dxf` | `C5S3-OB32.dxf` | `C5S4-IF16.dxf` | `C5S5-N2.dxf` .. `C5S11-N2.dxf` | `C5S12-N2.dxf` |

---

## ⚙️ Running Automation Scripts

```powershell
# Batch update all per-slot DXF drawings with CxSy-<CARD> naming:
python x9100-eDrawing/update_x9100_drawings.py

# Recompile master 65-slot (5 Rows x 13 Columns) DXF grid compilation file:
python x9100-eDrawing/build_master_dxf.py
```
