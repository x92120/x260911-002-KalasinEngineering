# 04_Automation_Scripts — Engineering CAD & Report Generators

This directory contains the Python automation tools, drawing generation engines, PDF compilers, and interactive HTML web viewers for the **Ingredion Kalasin Spray Dryer Plant** project.

---

## 🌐 Interactive Web Viewers

* 🌐 **[cable_schedule_viewer.html](file:///e:/xApp-01/x260911-002-KalasinEngineering/04_Automation_Scripts/cable_schedule_viewer.html)**: Dynamic cable schedule inspector with live search, field destination filtering, and cable routing breakdown.
* 🌐 **[pid_loop_control_viewer.html](file:///e:/xApp-01/x260911-002-KalasinEngineering/04_Automation_Scripts/pid_loop_control_viewer.html)**: Interactive PID loop control architecture viewer covering Jet Cooker, Dryer Tower, and CIP systems.

---

## 📁 Automation Scripts Inventory & Categorization

### 1. I/O Report & Master Excel Workbooks Generators
* ⚙️ **`build_chassis_slot_pdf_reports.py`**: Compiles per-chassis (C1–C7) slot-by-slot PDF vector reports.
* ⚙️ **`build_chassis_slot_source_destination_report.py`**: Builds 99-page master chassis & slot source-to-destination vector report.
* ⚙️ **`build_destination_io_report.py`**: Builds 82-page master destination report and 16 individual junction box workbooks.
* ⚙️ **`build_io_list_by_slot_config.py`**: Maps raw ePlan tag data into slot-configured Excel workbooks.
* ⚙️ **`build_junction_io_panel_report.py`**: Generates panel P1–P4 junction-to-panel allocation matrix & remote I/O reports.
* ⚙️ **`build_per_slot_io_report.py`**: Builds 48-tab master Excel workbook with 1 tab per slot.
* ⚙️ **`build_tags_list_plc_r02.py`**: Processes and updates PLC tag databases.

### 2. CAD Drawing Generators & Converters
* ⚙️ **`export_cad_svg_dxf.py`**: Batch CAD export pipeline converting ePlan DWG/DXF drawings into vectorized SVG presentation sheets.
* ⚙️ **`generate_di_slot_dxf_drawings.py`**: Generates 15 individual DI slot DXF schematics (C1S4–C4S3).
* ⚙️ **`generate_do_slot_dxf_drawings.py`**: Generates 9 individual DO slot DXF schematics (C1S8–C5S3).
* ⚙️ **`generate_plc_io_wiring_drawings.py`**: Generates 40-sheet PLC I/O modular wiring schematics (`KAL-JC-WIR-*`).
* ⚙️ **`generate_mcc_panel_layout.py`**: Generates 14-sheet MCC panel layout drawings (front elevation, internal GA, BOM).
* ⚙️ **`generate_instrument_drawing_templates.py`**: Generates 26 instrument installation hookup drawings (`KAL-JC-TYP-*`).
* ⚙️ **`generate_pf_is_barrier_dxf.py`**: Generates Pepperl+Fuchs IS barrier CAD wiring diagrams.
* ⚙️ **`convert_pf_datasheet_to_dxf.py`** & **`convert_yamamori_pdf_to_dxf.py`**: PDF datasheet-to-DXF CAD converter scripts.

### 3. Presentation & SCADA Assets Generators
* ⚙️ **`generate_automation_architecture_slides.py`**: Generates dark-theme automation architecture presentation slides.
* ⚙️ **`generate_light_architecture_slides.py`**: Generates light-theme architecture presentation slides.
* ⚙️ **`generate_cable_schedule_slides.py`**: Generates cable schedule SVG vector presentation slides.
* ⚙️ **`generate_pid_control_presentation.py`**: Generates PID loop control specification presentation slides.
* ⚙️ **`generate_scada_demo_video.py`**: Renders 60-second 1080p SCADA demo MP4 video animation.

---

## ⚙️ Running Generators

```powershell
# Build master destination I/O report & 16 enclosure workbooks:
python 04_Automation_Scripts/build_destination_io_report.py

# Build master chassis & slot 99-page vector PDF report:
python 04_Automation_Scripts/build_chassis_slot_source_destination_report.py

# Generate 40-sheet PLC wiring diagrams:
python 04_Automation_Scripts/generate_plc_io_wiring_drawings.py

# Generate MCC panel layout drawings:
python 04_Automation_Scripts/generate_mcc_panel_layout.py
```
