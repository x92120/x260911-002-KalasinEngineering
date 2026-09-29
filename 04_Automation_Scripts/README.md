# Automation Scripts & Engineering Tools
**Project**: Ingredion Thailand — Sprint 18K TPA Spray Dryer Plant (Kalasin)  
**Engineering Discipline**: Automated Drawing Generation, Report Compilation, & Interactive Viewers  

---

## 📌 Module Overview

This folder contains Python automation generators, CAD export pipelines, and interactive HTML viewers used to compile the deliverables in `02_Electrical_and_eDrawing/` and `03_IO_Lists_and_Schedules/`.

```text
04_Automation_Scripts/
├── cable_schedule_viewer.html                  # 🌐 Interactive Cable Schedule Explorer
├── pid_loop_control_viewer.html                # 🌐 Interactive PID Loop Control Architecture Viewer
├── export_cad_svg_dxf.py                       # Batch CAD export utility (DWG to DXF/SVG)
├── generate_all_architecture_deliverables.py   # Master compilation runner
├── generate_automation_architecture_slides.py  # Presentation slide generator (Dark theme)
├── generate_light_architecture_slides.py       # Presentation slide generator (Light theme)
├── generate_cable_schedule_slides.py           # Cable schedule slide and vector generator
├── generate_instrument_drawing_templates.py   # Instrument hookup and wiring template generator
├── generate_instrument_io_mapping.py           # Instrument-to-I/O mapping processor
├── generate_jb_pdf_reports.py                  # Junction box PDF report generator
├── generate_mcc_panel_layout.py                # MCC panel mounting and general arrangement generator
├── generate_pid_control_presentation.py        # PID control presentation generator
├── generate_plc_io_wiring_drawings.py          # 40-sheet PLC wiring diagram generator
├── generate_scada_demo_video.py                # SCADA demo animation renderer
├── generate_sorted_io_list.py                  # Chassis and slot sorting script
├── render_page7_bom.py                         # MCC BOM table vector renderer
├── update_7page_deliverables.py                # 7-page MCC drawing package compiler
└── update_network_diagram_fonts.py             # SVG font normalization utility
```

---

## 🌐 Interactive HTML Viewers

* [cable_schedule_viewer.html](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/04_Automation_Scripts/cable_schedule_viewer.html) — Dynamic cable schedule search, filter, and routing inspector.
* [pid_loop_control_viewer.html](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/04_Automation_Scripts/pid_loop_control_viewer.html) — Interactive PID loop cascade and feedforward diagram viewer.

---

## ⚙️ Running Generators

```bash
# Generate sorted I/O list:
python3 04_Automation_Scripts/generate_sorted_io_list.py

# Generate complete 40-sheet PLC wiring diagrams:
python3 04_Automation_Scripts/generate_plc_io_wiring_drawings.py

# Generate MCC panel layouts:
python3 04_Automation_Scripts/generate_mcc_panel_layout.py
```
