#!/usr/bin/env python3
"""
Enclosure Ventilation Load & Fan Sizing Calculation Script
Project: Kalasin Engineering - Sprint 18K TPA Spray Dryer / Jet Cooker
System Reference: xCIP-1545 (Project Ref: x2608003)
Standards: IEC 60890 / IEC 61439-1 / DIN 57660
Supplier Reference Document: x9100-eDrawing/Supplier/PrimusVentilation.pdf
"""

import math

def calculate_enclosure_thermal():
    print("=" * 80)
    print("KALASIN ENGINEERING - ENCLOSURE THERMAL & VENTILATION SIZING CALCULATION")
    print("=" * 80)
    
    # 1. Enclosure Geometry
    W = 0.800  # meters
    H = 2.000  # meters
    D = 0.400  # meters
    volume = W * H * D
    
    print("\n1. ENCLOSURE PHYSICAL SPECIFICATIONS:")
    print(f"  - Dimensions: {W*1000:.0f} mm (W) x {H*1000:.0f} mm (H) x {D*1000:.0f} mm (D)")
    print(f"  - Enclosure Gross Volume: {volume:.3f} m³")
    
    # Surface Area Calculations (IEC 60890)
    # A_total = 2*(W*H + D*H) + W*D (excluding bottom)
    A_front = W * H
    A_rear = W * H
    A_side = D * H
    A_top = W * D
    A_geom = A_front + A_rear + 2 * A_side + A_top
    
    # IEC 60890 effective surface area:
    # Free-standing on all sides: A = 1.8 * H * (W + D) + 1.4 * W * D
    A_eff_free = 1.8 * H * (W + D) + 1.4 * W * D
    # Stand-alone with rear against wall: A = 1.4*W*D + 0.9*D*H + W*H + 0.5*D*H
    A_eff_wall = 1.4 * W * D + 0.9 * D * H + W * H + 0.5 * D * H
    
    print(f"  - Exposed Geometric Surface Area (excluding bottom): {A_geom:.3f} m²")
    print(f"  - IEC 60890 Effective Surface Area (Free-Standing): {A_eff_free:.3f} m²")
    print(f"  - IEC 60890 Effective Surface Area (Wall-Mounted / Stand-Alone): {A_eff_wall:.3f} m²")
    
    # 2. Heat Load Breakdown
    print("\n2. HEAT DISSIPATION LOAD BREAKDOWN (P_internal):")
    
    # DC Power Supply 40A 2 PCS (Phoenix Contact QUINT4-PS/1AC/24DC/40 or equivalent)
    # Output: 24V x 40A = 960W each. Nominal efficiency: 94%.
    # Full load heat dissipation per unit: 61.3 W
    psu_qty = 2
    psu_loss_ea = 61.3  # Watts
    psu_total = psu_qty * psu_loss_ea
    print(f"  - DC Power Supply 24VDC 40A ({psu_qty} PCS): {psu_total:.1f} W ({psu_loss_ea:.1f} W each)")
    
    # 1756 13-Slot Chassis (ControlLogix 1756-A13)
    # PSU: 1756-PA75 (85-265VAC)
    p_pa75 = 25.0  # Watts
    # CPU: 1756-L95 (5580 Controller 40MB)
    p_cpu = 6.2    # Watts
    # EN4TR: 3 PCS (1Gbps Dual-Port EtherNet/IP)
    en4tr_qty = 3
    p_en4tr_ea = 3.5
    p_en4tr_tot = en4tr_qty * p_en4tr_ea
    # I/O Cards: 8 PCS (Digital/Analog 1756 modules: IB32, OB32, IF16)
    io_qty = 8
    p_io_ea = 5.5
    p_io_tot = io_qty * p_io_ea
    
    p_rack_total = p_pa75 + p_cpu + p_en4tr_tot + p_io_tot
    print(f"  - Allen-Bradley ControlLogix 1756 13-Slot Chassis: {p_rack_total:.1f} W")
    print(f"      • Chassis Power Supply (1756-PA75): {p_pa75:.1f} W")
    print(f"      • CPU Controller (1756-L95, 1 PCS): {p_cpu:.1f} W")
    print(f"      • Communication Adapters (1756-EN4TR, {en4tr_qty} PCS): {p_en4tr_tot:.1f} W ({p_en4tr_ea:.1f} W each)")
    print(f"      • I/O Modules ({io_qty} PCS): {p_io_tot:.1f} W ({p_io_ea:.1f} W each)")
    
    # Auxiliary Equipment & Cable Losses
    p_aux = 25.0  # Watts (Diode redundancy modules, MCBs, terminal blocks, wiring losses)
    print(f"  - Auxiliary & Infrastructure (Relays, MCBs, Diode module, wiring): {p_aux:.1f} W")
    
    p_subtotal = psu_total + p_rack_total + p_aux
    margin_factor = 0.15  # +15% safety / aging / future expansion margin
    p_margin = p_subtotal * margin_factor
    p_total = p_subtotal + p_margin
    btu_total = p_total * 3.412142
    
    print(f"  -------------------------------------------------------------")
    print(f"  BASE TOTAL HEAT LOAD: {p_subtotal:.1f} W")
    print(f"  DESIGN SAFETY MARGIN (+15%): {p_margin:.1f} W")
    print(f"  TOTAL DESIGN THERMAL LOAD (Q_total): {p_total:.1f} W ({btu_total:.1f} BTU/h)")
    
    # 3. Environmental Conditions & Ventilation Calculations
    T_amb = 30.0  # °C (Outside temperature specified by user)
    k_steel = 5.5  # W/(m²·K) for sheet steel enclosure
    f_air = 3.1    # m³·K/(W·h) standard air thermal constant at sea level
    filter_safety = 1.30  # +30% margin for dust loading & filter pressure drop
    
    print("\n3. THERMAL DISSIPATION & VENTILATION AIRFLOW REQUIREMENTS:")
    print(f"  - Ambient Outside Air Temperature: {T_amb:.1f} °C")
    print(f"  - Heat Transfer Coefficient of Steel (k): {k_steel:.1f} W/(m²·K)")
    print(f"  - Air Specific Heat Factor (f): {f_air:.1f} m³·K/(W·h)")
    print(f"  - Filter Clogging Safety Factor: {filter_safety:.2f} (+30% design headroom)")
    
    criteria = [
        ("Optimum Lifespan (Target 35°C)", 35.0),
        ("Standard Max Industrial (Target 40°C)", 40.0),
        ("Threshold Emergency (Target 45°C)", 45.0)
    ]
    
    print("\n  " + "-" * 88)
    print(f"  {'Design Condition':<35} | {'ΔT (K)':<8} | {'Q_s Wall':<10} | {'V_req pure':<12} | {'V_design (+30%)':<15}")
    print("  " + "-" * 88)
    
    for desc, T_in in criteria:
        dT = T_in - T_amb
        # Conservative: Q_v = P_total (zero credit for enclosure walls)
        V_req_pure = f_air * p_total / dT
        V_design_cons = V_req_pure * filter_safety
        
        # Wall-mounted natural convection: Q_s = k * A * dT
        Q_s_wall = k_steel * A_eff_wall * dT
        Q_v_net = max(0, p_total - Q_s_wall)
        V_req_wall = f_air * Q_v_net / dT
        V_des_wall = V_req_wall * filter_safety
        
        print(f"  {desc:<35} | {dT:<8.1f} | {Q_s_wall:<10.1f} | {V_req_pure:<12.1f} | {V_design_cons:<15.1f}")
    print("  " + "-" * 88)
    
    # 4. Primus Equipment Selection
    print("\n4. PRIMUS VENTILATION EQUIPMENT SELECTION (Supplier Document: PrimusVentilation.pdf):")
    print("""
  ====================================================================================
  RECOMMENDED PRIMARY SOLUTION:
  ====================================================================================
  • Filter Fan Unit: PRIMUS PMV250N (Size 8", 250 x 250 mm)
    - Power Supply: 220VAC, 50/60Hz, 1-Phase
    - Power Consumption: 49 W (0.38 A)
    - Nominal Free-Air Delivery: 428 m³/h (252 CFM)
    - Sound Pressure Level: 56 dB(A)
    - Protection Rating: IP54 (with polyurethane sealing gasket and G4 filter mat)
    - Mounting Feature: Clips-Lock quick snap-in mounting & Flap-Open front door for easy filter replacement
    - Justification: Delivers 428 m³/h, providing a 1.98x safety factor over the conservative
      35°C requirement (216 m³/h). Maintains enclosure temperature within ΔT ≤ 2.0–3.0°C.

  • Matching Exhaust Filter Grille: PRIMUS PMF250N
    - Frame Dimensions: 250 x 250 mm
    - Matched Cutout with PMV250N for symmetrical aesthetics and balanced pressure drop
    - Replaceable synthetic fiber filter mat (G4 / EU4 filtration class)

  • Temperature Controller: PRIMUS CMA-001 (or CMA-004-1 Digital)
    - Type: DIN-Rail Mechanical Thermostat with Bimetal Sensor (Normally Open contact)
    - Setting Range: +18°C to +87°C
    - Recommended Setting: Cut-in at +35°C, Cut-out at +30°C
    - Advantage: Turns fan ON only when internal heat exceeds 35°C, drastically extending
      fan bearing life, reducing dust accumulation on filter mats, and saving electrical energy.

  ====================================================================================
  ALTERNATIVE COMPACT SOLUTION:
  ====================================================================================
  • Filter Fan Unit: PRIMUS PMV25 (Size 255 x 255 mm)
    - Power Supply: 220VAC, 50/60Hz
    - Power Consumption: 38 W (0.5 A)
    - Nominal Airflow: 248 m³/h
    - Sound Pressure Level: 58 dB(A)
    - Matching Exhaust Filter: PRIMUS PMF25 (255 x 255 mm)
    - Justification: Exactly matches the 35°C design requirement (216 m³/h) with 15% margin.
    """)

if __name__ == "__main__":
    calculate_enclosure_thermal()
