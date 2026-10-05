# eDrawingTemplate — AutoCAD DXF & DWG Drawing Frames and Templates

This directory contains standardized AutoCAD drawing title block frames, border templates, and module wiring templates used by automated script generators for the **Ingredion Kalasin Spray Dryer Plant** CAD deliverables.

---

## 📐 Template Inventory

### 1. Title Block & Frame Templates
* **`frame.dwg`** / **`frame.dxf`**: Standard A3 ISO drawing frame with title block, engineering revision block, project metadata, and logo attributes.
* **`Frame_rev01.dxf`**: Revision 01 updated title block frame.
* **`Drawing2.dxf`**: Reference AutoCAD layout workspace template.

### 2. Module Electrical Wiring Templates
* **`template_DI-r02.dwg`** / **`template_DI-r02.dxf`**: Revision 02 32-Channel Digital Input (1756-IB32) module wiring drawing template.
* **`template_DO-r02.dwg`** / **`template_DO-r02.dxf`**: Revision 02 32-Channel Digital Output (1756-OB32 + Relay) module wiring drawing template.
* **`template_AI-r02.dwg`** / **`template_AI-r02.dxf`**: Revision 02 16-Channel Analog Input (1756-IF16) module wiring drawing template.
* **`template_DI.dxf`** / **`template_DO.dxf`**: Legacy baseline single-channel wiring templates.

---

## ⚙️ Usage in Automation Pipeline

Python scripts in `04_Automation_Scripts` (such as `generate_di_slot_dxf_drawings.py` and `generate_do_slot_dxf_drawings.py`) read these DXF templates, clone entity structures, update tag text and wire numbers, and output per-slot drawings into `02_Electrical_and_eDrawing/03_CAD_Exports_DXF_SVG/` and `x9100-eDrawing/`.
