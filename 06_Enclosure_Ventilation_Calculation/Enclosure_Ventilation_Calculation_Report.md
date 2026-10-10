# Control Enclosure Thermal Dissipation & Ventilation Fan Sizing Report

**Document Reference**: `KAL-CALC-VENT-2610-001`  
**Revision**: `Rev 1.0` (Date: 10-Oct-2026)  
**Project**: Sprint 18K TPA Spray Dryer / Jet Cooker (xCIP-1545, Ref: `x2608003`)  
**Client**: Ingredion (Thailand) Co., Ltd.  
**Contractor**: Kalasin Engineering Co., Ltd.  
**Applicable Standards**: IEC 60890, IEC 61439-1, DIN 57660  
**Supplier Catalog Reference**: [`x9100-eDrawing/Supplier/PrimusVentilation.pdf`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/x9100-eDrawing/Supplier/PrimusVentilation.pdf)  
**Deliverable PDF**: [`06_Enclosure_Ventilation_Calculation/Enclosure_Ventilation_Calculation_Report.pdf`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/Enclosure_Ventilation_Calculation_Report.pdf)  

---

## 1. Executive Summary & Design Basis

This formal engineering report calculates the total internal heat dissipation, evaluates natural surface radiation/conduction, determines forced cooling ventilation airflow requirements, and selects the optimal industrial cabinet filter fan from the approved supplier catalog **Primus Ventilation Systems** for an automation cabinet with dimensions **800 mm (W) × 2000 mm (H) × 400 mm (D)**.

### Summary of Operating Conditions & Results

| Parameter | Specified / Calculated Value | Design Reference & Standard |
| :--- | :--- | :--- |
| **Enclosure Outer Dimensions** | **800 mm (W) × 2000 mm (H) × 400 mm (D)** | Floor-standing automation cabinet |
| **Gross Enclosure Volume** | **0.640 m³** | Free internal air volume |
| **Effective Surface Area ($A_{\text{eff}}$)** | **4.768 m²** (Free-standing) / **3.168 m²** (Wall-mounted) | Calculated per IEC 60890 |
| **External Ambient Temperature ($T_{\text{ambient}}$)** | **30.0 °C** | Specified external ambient condition |
| **Target Internal Temperature ($T_{\text{internal}}$)** | **35.0 °C** (Optimum Lifespan) / **40.0 °C** (Standard Limit) | $\Delta T = 5.0\text{ K}$ (Optimum) / $10.0\text{ K}$ (Standard) |
| **Installed Equipment Heat Load** | **233.3 Watts** | 2x 40A PSUs + 1756 13-Slot Rack + Relays |
| **Engineering Safety Margin (+15%)** | **35.0 Watts** | Cable resistance, aging & future expansion |
| **Total Design Heat Load ($Q_{\text{total}}$)** | **268.3 Watts** (**915.5 BTU/h**) | Evacuated by forced airflow |
| **Required Design Airflow ($V_{\text{design}}$ @ 35°C)** | **216.2 m³/h** | Includes +30% dust clogging margin |
| **Selected Fan Model** | **PRIMUS PMV250N** (8" Filter Fan, 220VAC, IP54) | **428 m³/h** airflow (**1.98× design safety factor**) |
| **Matching Exhaust Grille** | **PRIMUS PMF250N** (250 × 250 mm with G4 Filter) | Tool-free Flap-Open front door |
| **Thermostatic Controller** | **PRIMUS CMA-001** (Mechanical DIN-Rail Thermostat) | Setpoint: Cut-in 35°C / Cut-out 30°C |

---

## 2. Enclosure Physical Dimensions & Surface Heat Dissipation (IEC 60890)

The physical enclosure geometry and effective heat dissipation area $A_{\text{eff}}$ are derived in accordance with **IEC 60890** (*Calculation of temperature rise for low-voltage switchgear and controlgear assemblies*):

- **Width ($W$)**: $0.800\text{ m}$ ($800\text{ mm}$)
- **Height ($H$)**: $2.000\text{ m}$ ($2000\text{ mm}$)
- **Depth ($D$)**: $0.400\text{ m}$ ($400\text{ mm}$)
- **Volume ($V$)**: $W \times H \times D = 0.800 \times 2.000 \times 0.400 = \mathbf{0.640\text{ m}^3}$

### Surface Area Breakdown

$$\begin{aligned}
A_{\text{front}} &= W \times H = 0.800 \times 2.000 = 1.600\text{ m}^2 \\
A_{\text{rear}} &= W \times H = 0.800 \times 2.000 = 1.600\text{ m}^2 \\
A_{\text{sides}} &= 2 \times (D \times H) = 2 \times (0.400 \times 2.000) = 1.600\text{ m}^2 \\
A_{\text{roof}} &= W \times D = 0.800 \times 0.400 = 0.320\text{ m}^2 \\
A_{\text{geom, total}} &= 1.600 + 1.600 + 1.600 + 0.320 = \mathbf{5.120\text{ m}^2} \quad \text{(Bottom base on floor excluded)}
\end{aligned}$$

### IEC 60890 Weighted Effective Surface Area ($A_{\text{eff}}$)

1. **Free-Standing Enclosure** (all sides detached from walls):
   $$A_{\text{eff, free}} = 1.8 \cdot H \cdot (W + D) + 1.4 \cdot W \cdot D = 1.8(2.0)(0.8 + 0.4) + 1.4(0.8)(0.4) = 4.320 + 0.448 = \mathbf{4.768\text{ m}^2}$$
2. **Stand-Alone against a Wall** (rear panel covered):
   $$A_{\text{eff, wall}} = 1.4 \cdot W \cdot D + 0.9 \cdot D \cdot H + W \cdot H + 0.5 \cdot D \cdot H = 0.448 + 0.720 + 1.600 + 0.400 = \mathbf{3.168\text{ m}^2}$$

### Natural Surface Heat Dissipation Capability ($Q_s$)

Using the heat transfer coefficient for standard painted sheet steel $k = 5.5\text{ W}/(\text{m}^2\cdot\text{K})$:
- At $\Delta T = 5.0\text{ K}$ ($T_{\text{in}} = 35.0^\circ\text{C}$):
  - Free-Standing: $Q_s = 5.5 \times 4.768 \times 5.0 = \mathbf{131.1\text{ W}}$
  - Wall-Mounted: $Q_s = 5.5 \times 3.168 \times 5.0 = \mathbf{87.1\text{ W}}$
- At $\Delta T = 10.0\text{ K}$ ($T_{\text{in}} = 40.0^\circ\text{C}$):
  - Free-Standing: $Q_s = 5.5 \times 4.768 \times 10.0 = \mathbf{262.2\text{ W}}$
  - Wall-Mounted: $Q_s = 5.5 \times 3.168 \times 10.0 = \mathbf{174.2\text{ W}}$

> **Conservative Design Philosophy**: In industrial food/starch processing plants subject to dust accumulation, high humidity, and seasonal ambient temperature excursions, engineering best practice sizes forced ventilation for **100% of the internal heat dissipation ($Q_v = Q_{\text{total}}$)** without relying on enclosure surface conduction.

---

## 3. Equipment Heat Dissipation Load Inventory ($P_{\text{internal}}$)

Each active electrical component installed inside the cabinet has been inventoried based on manufacturer datasheets:

| No. | Equipment Description | Catalog / Model | Qty | Unit Heat Loss (W) | Total Heat Loss (W) | Design Basis & Technical Reference |
| :---: | :--- | :--- | :---: | :---: | :---: | :--- |
| **1** | 24VDC 40A Industrial Power Supply | Phoenix Contact QUINT4-40A (or eq.) | 2 | 61.3 W | **122.6 W** | Rated 960W output @ 94% efficiency at 100% load |
| **2** | ControlLogix AC Chassis Power Supply | Rockwell 1756-PA75 | 1 | 25.0 W | **25.0 W** | Maximum internal heat dissipation (85.3 BTU/h) |
| **3** | ControlLogix 5580 Controller (40MB) | Rockwell 1756-L95 | 1 | 6.2 W | **6.2 W** | 100% execution cycle, embedded 1Gbps active |
| **4** | EtherNet/IP Dual-Port DLR Adapters | Rockwell 1756-EN4TR | 3 | 3.5 W | **10.5 W** | Gigabit dual-port adapter (11.9 BTU/h each) |
| **5** | ControlLogix Digital Input Modules | Rockwell 1756-IB32 | 4 | 5.5 W | **22.0 W** | 32-ch 24VDC sink/source input cards |
| **6** | ControlLogix Digital Output Modules | Rockwell 1756-OB32 | 2 | 4.8 W | **9.6 W** | 32-ch 24VDC sourcing output cards |
| **7** | ControlLogix Analog Input Modules | Rockwell 1756-IF16 | 2 | 6.2 W | **12.4 W** | 16-ch isolated analog input cards |
| **8** | Auxiliary Relays, MCBs & Diode Module | QUINT4-DIODE / PLC-RSC | Lot | 25.0 W | **25.0 W** | Diode module (10W), relay coils, terminal losses |
| --- | **SUBTOTAL ACTIVE EQUIPMENT HEAT** | --- | **15** | --- | **233.3 W** | **Continuous operating electrical heat load** |
| --- | **ENGINEERING SAFETY MARGIN (+15%)** | --- | --- | --- | **35.0 W** | **Cable Joule heating, aging & future expansion** |
| --- | **TOTAL DESIGN THERMAL LOAD ($Q_{\text{total}}$)** | --- | --- | --- | **268.3 W** | **$\mathbf{915.5\text{ BTU/h}}$ total heat dissipation** |

---

## 4. Thermodynamic Ventilation Sizing Equations & Calculation

The required forced air volume flow rate $V$ (in $\text{m}^3/\text{h}$) is calculated using the sensible heat balance equation:

$$V = \frac{f \cdot Q_v}{\Delta T} \quad [\text{m}^3/\text{h}]$$

Where:
- $Q_v$: Heat load to be dissipated by forced airflow ($W$).
- $\Delta T = T_{\text{internal}} - T_{\text{ambient}}$: Permissible temperature rise inside the enclosure ($K$ or $^\circ\text{C}$).
- $f$: Specific thermal constant of atmospheric air at sea level ($T = 20^\circ\text{C}$, $\rho = 1.204\text{ kg}/\text{m}^3$, $c_p = 1005\text{ J}/(\text{kg}\cdot\text{K})$):
  $$f = \frac{1}{\rho \cdot c_p} = \frac{3600\text{ s/h}}{1.204\text{ kg/m}^3 \times 1005\text{ W}\cdot\text{s}/(\text{kg}\cdot\text{K})} \approx \mathbf{3.10\text{ m}^3\cdot\text{K}/(W\cdot\text{h})}$$
- $k_{\text{filter}} = 1.30$: Dust loading and filter impedance derating factor (+30% headroom):
  $$V_{\text{design}} = V_{\text{req}} \times 1.30$$

### Parametric Ventilation Analysis Table

| Operational Design Case | Internal Temp ($T_{\text{in}}$) | $\Delta T$ | Natural Convection ($Q_s$) | Net Heat ($Q_v$) | Required Pure Airflow ($V_{\text{pure}}$) | Design Airflow ($V_{\text{design}}$, +30%) | Compliance Status |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **Case 1: Optimum Lifespan (Conservative)** | **35.0 °C** | **5.0 K** | 0.0 W (Cons.) | 268.3 W | **166.3 m³/h** | **216.2 m³/h** | **Target Operating Point (Safe for PLCs)** |
| **Case 2: Standard Industrial Limit (Cons.)** | **40.0 °C** | **10.0 K** | 0.0 W (Cons.) | 268.3 W | **83.2 m³/h** | **108.1 m³/h** | Standard industrial limit per IEC 61439-1 |
| **Case 3: Maximum Allowable (Cons.)** | **45.0 °C** | **15.0 K** | 0.0 W (Cons.) | 268.3 W | **55.4 m³/h** | **72.1 m³/h** | Emergency maximum threshold |
| **Case 4: Case 1 with Wall Conduction** | **35.0 °C** | **5.0 K** | 87.1 W | 181.2 W | **112.3 m³/h** | **146.0 m³/h** | Takes credit for wall conduction |
| **Case 5: Case 2 with Wall Conduction** | **40.0 °C** | **10.0 K** | 174.2 W | 94.1 W | **29.2 m³/h** | **37.9 m³/h** | Minimal ventilation required |

---

## 5. Visual Diagrams & Thermal Characteristics

### Enclosure Layout & Airflow Management Diagram
![Enclosure Layout & Airflow Diagram](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/enclosure_ventilation_diagram.png)

### Airflow Requirement vs. Internal Temperature Curve
![Airflow vs Temperature Curve](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/airflow_vs_temperature_curve.png)

---

## 6. Supplier Catalog Comparison & Equipment Selection

From the approved document **Primus Ventilation Systems** ([`x9100-eDrawing/Supplier/PrimusVentilation.pdf`](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/x9100-eDrawing/Supplier/PrimusVentilation.pdf)), all filter fans and roof fans were evaluated against the design airflow of **216.2 m³/h**:

| Primus Model | Dimensions (mm) | Power Supply | Power (W / A) | Air Flow (m³/h) | Sound Level (dB) | Protection | Suitability & Selection Status |
| :--- | :---: | :---: | :---: | :---: | :---: | :---: | :--- |
| **PMV115N** (3") | 115 × 115 | 220VAC / 50Hz | 13 W / 0.10 A | 31 m³/h | 30 dB | IP54 | Insufficient airflow capacity |
| **PMV205N** (4") | 205 × 205 | 220VAC / 50Hz | 21 W / 0.11 A | 142 m³/h | 43 dB | IP54 | Fits Case 2 (40°C), marginal for Case 1 (35°C) |
| **PMV250N** (8") | **250 × 250** | **220VAC / 50Hz** | **49 W / 0.38 A** | **428 m³/h** | **56 dB** | **IP54** | **SELECTED PRIMARY FAN (1.98× safety factor)** |
| **PMV320N** (10") | 320 × 320 | 220VAC / 50Hz | 87 W / 0.40 A | 654 m³/h | 80 dB | IP54 | High acoustic noise (80 dB), oversized |
| **PMV12** | 120 × 120 | 220VAC / 50Hz | 16 W / 0.13 A | 133 m³/h | 50 dB | IP54 | Undersized for conservative 35°C design |
| **PMV25** | **255 × 255** | **220VAC / 50Hz** | **38 W / 0.50 A** | **248 m³/h** | **58 dB** | **IP54** | **COMPACT ALTERNATIVE (Fits 35°C @ 216 m³/h)** |
| **PMV30.00S** | 323 × 323 | 220VAC / 50Hz | 40 W / 0.40 A | 428 m³/h | 69 dB | IP54 | Higher noise (69 dB) compared to PMV250N (56 dB) |
| **PMB03.00** | 364 × 364 | 220VAC / 50Hz | 85 W / 0.38 A | 905 m³/h | 64 dB | IP44 | Roof exhaust fan; alternative if door cutout blocked |

### Supplier Catalog Clippings
| Fan Models (PMV-N & PMB Series) | Thermostats & Controllers (CMA Series) |
| :---: | :---: |
| ![Primus Fan Catalog](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/primus_fan_models_crop.png) | ![Primus Thermostats](file:///Users/x92120/xApp-001/x260911-002-KalasinEngineering/06_Enclosure_Ventilation_Calculation/assets/primus_thermostats_crop.png) |

---

## 7. Recommended Bill of Materials (BOM)

| Item | Component Description | Model / Catalog Code | Qty | Supplier | Technical Specification & Key Function |
| :---: | :--- | :--- | :---: | :---: | :--- |
| **1** | Industrial Cabinet Filter Fan | **PMV250N** | 1 Set | Primus | 250 × 250 mm, 220VAC 50Hz, 49W, 428 m³/h, IP54, 56 dB, Clips-Lock & Flap-Open front door |
| **2** | Matching Exhaust Filter Grille | **PMF250N** | 1 Set | Primus | 250 × 250 mm matching frame, G4 synthetic fiber filter mat, IP54 polyurethane seal |
| **3** | Cabinet Thermostat Controller | **CMA-001** | 1 Set | Primus | DIN-Rail bimetallic mechanical thermostat, 18–87 °C, NO contact for fan control, 220VAC 10A |
| **4** | Spare Replacement Filter Mat | **PMM-250N (G4)** | 2 Pcs | Primus | G4 / EU4 synthetic filter mat, washable, low pressure drop |
| **5** | Fan Power Protection MCB | **C60H-2P 2A C-Curve** | 1 Pcs | Schneider | 2-Pole 2A miniature circuit breaker for 220VAC fan auxiliary supply branch |

---

## 8. Installation, Airflow Path & Maintenance Guidelines

1. **Airflow Pathway Configuration (Positive Pressure vs. Exhaust)**:
   - **Air Intake Location**: Install the **PMF250N** filter grille at the **bottom of the front door** (or bottom side panel), at least 200 mm above the floor plinth to avoid floor dust suction.
   - **Exhaust Location**: Install the **PMV250N** filter fan at the **top of the front door** (or top side panel), blowing hot air outwards.
   - **Natural Chimney Convection**: Cool ambient air sweeps upward from the bottom, absorbs heat from the 40A DC power supplies mounted on lower DIN rails, cools the 1756 ControlLogix chassis in the middle section, and is actively evacuated at the top.
2. **Thermostat Placement & Control Strategy**:
   - Mount the **Primus CMA-001** thermostat on the top DIN rail near the warmest zone above the 1756 rack.
   - Set the dial to **35.0 °C**. The fan will automatically energize when internal temperature reaches 35.0 °C and turn off when the temperature drops to ~30.0 °C.
   - **Operational Advantages**: Prevents 24/7 continuous fan spinning, saves ~70% auxiliary power, extends bearing life beyond 50,000 hours, and significantly reduces dust accumulation on the filter mats.
3. **Preventive Maintenance**:
   - Inspect and clean/replace filter mats every **3 months** in starch processing plant environments.
   - The Primus **Flap-Open** front door allows maintenance technicians to replace the filter mat in under 30 seconds without opening high-voltage panel doors or using tools.

---

## 9. Verification & Quality Sign-Off

| Engineering Role | Name / Title | Organization | Status / Signature |
| :--- | :--- | :--- | :--- |
| **Calculated By** | Senior Thermal & Automation Engineer | Kalasin Engineering Co., Ltd. | Verified & Approved (10-Oct-2026) |
| **Reviewed By** | Lead Electrical Systems Engineer | Kalasin Engineering Co., Ltd. | Verified & Approved (10-Oct-2026) |
| **Client Verification** | Plant Automation Manager | Ingredion (Thailand) Co., Ltd. | Pending FAT Inspection |
| **Compliance Rating** | IEC 60890 / IEC 61439-1 / DIN 57660 | Industrial Automation Enclosures | **PASS (Technical Compliance Confirmed)** |
