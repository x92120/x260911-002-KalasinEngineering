# 06 Enclosure Ventilation & Thermal Calculation

**Project**: Sprint 18K TPA Spray Dryer / Jet Cooker  
**Ref Code**: `xCIP-1545` (Ref: `x2608003`)  
**Client**: Ingredion (Thailand) Co., Ltd.  
**Contractor**: Kalasin Engineering Co., Ltd.  

---

## Deliverables in this Folder

1. **Official PDF Calculation Report**:
   - [`Enclosure_Ventilation_Calculation_Report.pdf`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/Enclosure_Ventilation_Calculation_Report.pdf) (4 Pages, ready for client submission & FAT binder)

2. **Full Markdown Engineering Report**:
   - [`Enclosure_Ventilation_Calculation_Report.md`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/Enclosure_Ventilation_Calculation_Report.md)

3. **Source Typst File**:
   - [`Enclosure_Ventilation_Calculation_Report.typ`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/Enclosure_Ventilation_Calculation_Report.typ)

4. **Automated Calculation Script**:
   - [`calculate_ventilation.py`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/calculate_ventilation.py)

5. **Diagrams & Catalog Excerpts in `assets/`**:
   - Enclosure Thermal Layout Diagram: [`assets/enclosure_ventilation_diagram.png`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/enclosure_ventilation_diagram.png)
   - Airflow vs. Temperature Characteristic Curves: [`assets/airflow_vs_temperature_curve.png`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/airflow_vs_temperature_curve.png)
   - Supplier Catalog Excerpt (Fans): [`assets/primus_fan_models_crop.png`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/primus_fan_models_crop.png)
   - Supplier Catalog Excerpt (Thermostats): [`assets/primus_thermostats_crop.png`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/primus_thermostats_crop.png)

---

## Quick Summary of Calculation Results

- **Enclosure**: 800 (W) × 2000 (H) × 400 (D) mm (Volume = 0.640 m³)
- **External Ambient**: 30.0 °C
- **Target Enclosure Temperature**: 35.0 °C ($\Delta T = 5.0\text{ K}$)
- **Total Thermal Dissipation**: **268.3 Watts** (**915.5 BTU/h**), including +15% engineering margin
- **Required Airflow**:
  - Pure Airflow: **166.3 m³/h**
  - Design Airflow (+30% filter margin): **216.2 m³/h**
- **Selected Primus Equipment**:
  - **Filter Fan Unit**: **Primus PMV250N** (8", 220VAC, **428 m³/h**, 56 dB, IP54) — **1.98× design safety factor**
  - **Exhaust Grille**: **Primus PMF250N** (250 × 250 mm with G4 washable filter mat)
  - **Thermostat**: **Primus CMA-001** (Mechanical bimetal, set to Cut-in 35°C / Cut-out 30°C)
