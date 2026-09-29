#!/usr/bin/env python3
"""
Generate complete publication-grade Instrument Hook-Up, Process Connection,
and Electrical Loop Wiring Drawing Templates for:
Project: KALASIN STARCH PLANT - JET COOKER PROJECT
Client: INGREDION (THAILAND) CO., LTD.
System: Rockwell Automation ControlLogix 1756 & Field Instrumentation
Based on: 11-instrument Manual/ and Instrument_IO_Mapping_with_Models_and_Cables.xlsx

Outputs:
1. Multi-page A3 Landscape PDF drawing package (Cover, Index, + 23 Instrument Templates).
2. Master AutoCAD DXF model space file (1:1 Metric mm) + 23 individual DXF files.
3. Vector SVG drawing sheets for web / CAD viewers.
4. High-resolution PNG preview images for key templates.
"""

import os
import sys
import math
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import fitz  # PyMuPDF
import ezdxf
from ezdxf import colors
from ezdxf.enums import TextEntityAlignment

WORKSPACE_DIR = os.path.dirname(os.path.abspath(__file__))
PDF_DIR = os.path.join(WORKSPACE_DIR, "reports_pdf")
CAD_DIR = os.path.join(WORKSPACE_DIR, "cad_exports")
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(CAD_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# 23 INSTRUMENT TEMPLATES SPECIFICATION DATABASE
# -----------------------------------------------------------------------------
TEMPLATES_DATA = [
    {
        "code": "TYP-01",
        "title": "DUST / PARTICULATE ANALYZER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-001",
        "category": "Dust / Particulate Analyzer",
        "tag_prefix": "AT-60241 Series",
        "service": "Exhaust Duct Dust Emission Monitoring",
        "mfr": "ENVEA / PCME",
        "model": "PCME QAL 991 (Electrodynamic / Optical)",
        "proc_conn": 'DN100 Flange w/ Purge Air (1.5" BSP)',
        "elec_int": "4-20mA Output + Alarm Relay (24V DC)",
        "cable_spec": "2Px1.5mm² IS Shielded Overall (Belden 8777)",
        "power_supply": "24V DC / 15W + Clean Dry Instrument Air (4-6 bar)",
        "haz_class": "ATEX Zone 21/22 Dust Ex tb IIIC T120°C",
        "enclosure": "IP65 Weatherproof Cast Aluminum",
        "plc_module": "1756-IF16 (AI) & 1756-IB32 (DI Alarm)",
        "signal_type": "AI: 4-20mA Dust Conc / DI: Fault Contact",
        "dest_jb": "JB-602 / CA1 Marshalling",
        "barrier": "937THAITXPDC1 (IS Analog Barrier)",
        "bom": [
            ("1", "Dust Transmitter Sensor Unit", "PCME QAL 991", "1 EA"),
            ("2", "Mating Counter-Flange DN100 PN16", "Carbon Steel / SS316", "1 EA"),
            ("3", "Purge Air Filter Regulator & Gauge", "SMC 1/4\" BSP w/ Bracket", "1 SET"),
            ("4", "Cable Gland M20x1.5 Ex d/e", "Nickel-plated Brass", "2 EA"),
            ("5", "Instrument Stand 2\" Pipe & Base", "Hot-dip Galvanized Steel", "1 SET")
        ],
        "notes": [
            "Mount sensor on straight run min. 5D downstream, 2D upstream of duct bends.",
            "Instrument air purge must remain continuously ON to prevent window fouling.",
            "Shield drain wire must be isolated in field and grounded only at Marshalling Cabinet CA1 earth bar."
        ]
    },
    {
        "code": "TYP-02",
        "title": "COMBUSTIBLE LEL GAS DETECTOR HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-002",
        "category": "Combustible LEL Gas Analyzer",
        "tag_prefix": "AIT-60101 / AIT-60201",
        "service": "Explosion Hazard Lower Explosive Limit (LEL) Monitor",
        "mfr": "Sensidyne / SmartGas",
        "model": "925FGD Explosionproof Hydrocarbon Gas Detector",
        "proc_conn": 'Ambient Air Diffusion / Splash Guard Assembly',
        "elec_int": "4-20mA LEL Signal + Low/High Alarm Dry Contacts",
        "cable_spec": "3Cx1.5mm² Shielded Fire-Resistant (Belden 9533)",
        "power_supply": "24V DC (18-32V DC, 200mA max)",
        "haz_class": "Class I, Div 1, Grp BCD / Ex d IIC T5 Gb",
        "enclosure": "NEMA 4X / IP66 Heavy-duty Explosionproof",
        "plc_module": "1756-IF16 (AI LEL) & 1756-IB32 (DI Alarm)",
        "signal_type": "AI: 0-100% LEL (4-20mA) / DI: Gas Alarm",
        "dest_jb": "IS-JB-603 / CA1 Marshalling",
        "barrier": "937THAITXPDC1 (IS Repeater Barrier)",
        "bom": [
            ("1", "Gas Detector Transmitter 925FGD", "Sensidyne Infrared / Catalytic", "1 EA"),
            ("2", "Weather Guard / Dust Protection Cover", "Hydrophobic SS316", "1 EA"),
            ("3", "Calibration Adapter & Cup", "Molded PVC / Viton O-ring", "1 EA"),
            ("4", "Explosionproof Cable Gland 3/4\" NPT", "SS316 Flameproof Ex d", "1 EA"),
            ("5", "Wall / Stanchion Mounting Bracket", "SS304 Universal Bracket", "1 SET")
        ],
        "notes": [
            "Install 0.3m to 0.5m above floor for gases heavier than air, or near ceiling for lighter gases.",
            "Field sensor must undergo zero and span gas calibration prior to plant commissioning.",
            "Wiring must run in rigid galvanized steel conduit or approved armored flameproof cable."
        ]
    },
    {
        "code": "TYP-03",
        "title": "pH ANALYTICAL MEASUREMENT LOOP HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-003",
        "category": "pH Analytical Loop",
        "tag_prefix": "AIT-40101 / AIT-40201",
        "service": "Cooker Slurry & Flash Tank pH Neutralization",
        "mfr": "Mettler Toledo",
        "model": "M300G2 4-Wire Transmitter + InPro3250i + InTrac 777P",
        "proc_conn": "DN25 Weld-in Socket (Inclined 15°) / Retractable Housing",
        "elec_int": "4-20mA Process pH + 4-20mA Temp + 2x Relay Contacts",
        "cable_spec": "VP6 High-Z Cable to Tx + 2Px1.5mm² Shielded to PLC",
        "power_supply": "24V DC / 10W (Optional 100-240V AC)",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6/T4/T3 Ga/Gb",
        "enclosure": "NEMA 4X / IP65 Polycarbonate Housing",
        "plc_module": "1756-IF16 (AI pH) & 1756-IB32 (DI Alarms)",
        "signal_type": "AI: 0-14 pH (4-20mA HART) / AI: Temp",
        "dest_jb": "JB-401 / CA1 Marshalling",
        "barrier": "Internal Galvanic Isolation (4-Wire Tx)",
        "bom": [
            ("1", "pH Multi-Parameter Transmitter M300G2", "Mettler Toledo M300 4-Wire", "1 EA"),
            ("2", "Intelligent pH Sensor InPro3250i", "Pre-pressurized Liquid Gel", "1 EA"),
            ("3", "Retractable Housing InTrac 777P", "SS316L / Hastelloy C w/ Flange", "1 EA"),
            ("4", "Weld-in Socket 40mm / DN25 (15°)", "SS316L Sloped Socket", "1 EA"),
            ("5", "2\" Pipe Mounting Kit for M300G2", "SS304 Bracket & U-Bolts", "1 SET")
        ],
        "notes": [
            "Weld socket must be installed inclined at min. 15° above horizontal to prevent air bubble entrapment.",
            "Use retractable housing safety interlock valve before extracting probe during process operation.",
            "Ground transmitter chassis directly to plant PE bar with min. 4mm² copper grounding strap."
        ]
    },
    {
        "code": "TYP-04",
        "title": "DIFFERENTIAL PRESSURE GAUGE HOOK-UP & MECHANICAL DETAIL",
        "dwg_no": "KAL-JC-TYP-004",
        "category": "Diff. Pressure Gauge",
        "tag_prefix": "PDI-40101 / PDI-60201",
        "service": "Bag Filter / Steam Strainer Differential Pressure",
        "mfr": "Ashcroft",
        "model": "Model 1132 (Piston-Type Differential Gauge)",
        "proc_conn": '1/4" NPT Female w/ 3-Valve Direct Manifold (V03)',
        "elec_int": "Mechanical Local Indicator (No Electrical Wiring)",
        "cable_spec": "N/A (Local Mechanical Dial)",
        "power_supply": "None (Direct Process Pressure Driven)",
        "haz_class": "Non-Electrical Apparatus / Standard Zone 1/2",
        "enclosure": "IP65 Stainless Steel Case, Glycerin Filled",
        "plc_module": "None (Local Operator Readout)",
        "signal_type": "Visual Dial 0-100 kPa / 0-1 bar",
        "dest_jb": "Direct Process Piping Mount",
        "barrier": "N/A",
        "bom": [
            ("1", "Differential Pressure Gauge 1132", "Ashcroft 100mm Dial SS316", "1 EA"),
            ("2", "3-Valve Instrument Manifold V03", "SS316 1/4\" NPT Female", "1 EA"),
            ("3", "Impulse Tubing 1/2\" OD x 0.049\" W", "ASTM A269 Seamless SS316", "AR"),
            ("4", "Compression Tube Fittings 1/2\" x 1/4\"", "Double Ferrule SS316 (Swagelok)", "4 EA"),
            ("5", "Gauge Mounting Bracket 2\" Pipe Stand", "SS304 Angle Bracket & Clamps", "1 SET")
        ],
        "notes": [
            "High (HP) and Low (LP) pressure impulse lines must slope min. 1:12 downward toward drain pots.",
            "Operate 3-valve manifold equalizing valve before opening process isolation valves to protect sensor.",
            "Provide blow-down bleed valves on both impulse legs for routine cleaning and purging."
        ]
    },
    {
        "code": "TYP-05",
        "title": "DIFFERENTIAL PRESSURE TRANSMITTER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-005",
        "category": "Diff. Pressure Transmitter",
        "tag_prefix": "DPT-60201 / DPT-60901",
        "service": "Filter Differential Pressure / Orifice Flow Drop",
        "mfr": "Endress+Hauser",
        "model": "Deltabar PMD55B (Piezoresistive / Ceramic)",
        "proc_conn": '1/4" NPT Female w/ 5-Valve Manifold DA63M',
        "elec_int": "2-Wire 4-20mA HART Loop Powered (24V DC)",
        "cable_spec": "1Px1.5mm² Twisted Shielded (Belden 8761 / IEC 60332)",
        "power_supply": "24V DC via 1756-IF16 Analog Input Loop",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6 Ga/Gb",
        "enclosure": "NEMA 4X / IP66/68 Dual-Compartment Aluminum",
        "plc_module": "1756-IF16 (16-Point Diff Analog Input)",
        "signal_type": "4-20mA HART Process Delta-P (0-500 mbar)",
        "dest_jb": "JB-602 / IS-JB-603",
        "barrier": "937THAITXPDC1 (IS Repeater Barrier)",
        "bom": [
            ("1", "Deltabar PMD55B Transmitter", "E+H Ceramic / SS316 Diaphragm", "1 EA"),
            ("2", "5-Valve Manifold DA63M (Flange-Flange)", "SS316 Integrated Bleed Valves", "1 EA"),
            ("3", "2\" Pipe Mounting Bracket Kit", "SS304 Stanchion Bracket & U-Bolts", "1 SET"),
            ("4", "Cable Gland M20x1.5 Ex d/e IP68", "Nickel-Plated Brass Armored Gland", "1 EA"),
            ("5", "Impulse Tubing 1/2\" OD x 0.049\" W", "ASTM A269 Seamless SS316", "AR")
        ],
        "notes": [
            "Mount transmitter below process tapping points for liquid service to maintain filled impulse lines.",
            "5-valve manifold sequence: Open equalizing, open HP & LP blocks, close equalizing, check zero.",
            "Connect 2-wire loop to PLC 1756-IF16 terminals: (+) to 24V loop power, (-) to Channel Input."
        ]
    },
    {
        "code": "TYP-06",
        "title": "CORIOLIS MASS FLOWMETER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-006",
        "category": "Coriolis Flowmeter",
        "tag_prefix": "FT-40201 Series",
        "service": "Starch Slurry High-Accuracy Mass Flow & Density",
        "mfr": "Endress+Hauser",
        "model": "Promass E 300 (8E3B50 / Dual Tube SS316L)",
        "proc_conn": "DN50 (2\") ASME B16.5 Cl.150 RF Flange",
        "elec_int": "Channel 1: 4-20mA HART (Mass) / Channel 2: 4-20mA (Density)",
        "cable_spec": "2Px1.5mm² Shielded Overall (Belden 8777 / BS 5308)",
        "power_supply": "24V DC / 10W (Transmitter Electronics)",
        "haz_class": "ATEX II 2G Ex db eb ia IIC T6...T1 Gb",
        "enclosure": "IP67 Die-cast Aluminum Transmitter Housing",
        "plc_module": "1756-IF16 (Ch. 1 Mass Flow / Ch. 2 Density)",
        "signal_type": "4-20mA HART Mass (kg/h) & Density (g/cm³)",
        "dest_jb": "JB-401 (Cooker Section 1)",
        "barrier": "Galvanically Isolated Active 4-20mA Outputs",
        "bom": [
            ("1", "Promass E 300 Coriolis Flowmeter", "E+H DN50 Cl.150 SS316L Dual Tube", "1 EA"),
            ("2", "Companion Counter-Flanges DN50 Cl.150", "ASTM A182 F316 / RF Flanges", "2 EA"),
            ("3", "PTFE / Spiral Wound Gaskets", "316SS with PTFE Filler", "2 EA"),
            ("4", "Cable Glands M20x1.5 Ex d/e", "Nickel-plated Brass Armored Glands", "2 EA"),
            ("5", "Pipe Support Clamps & Structural Stand", "Heavy Duty Rigidity Anchors", "2 SET")
        ],
        "notes": [
            "Sensor tubes must be installed pointing downwards (flag-down) for liquid slurry to prevent gas collection.",
            "Support pipeline rigidly on both upstream and downstream sides within 2 pipe diameters of meter flanges.",
            "Do not clamp or attach pipe supports directly to the Coriolis sensor housing."
        ]
    },
    {
        "code": "TYP-07",
        "title": "ELECTROMAGNETIC FLOWMETER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-007",
        "category": "Electromagnetic Flowmeter",
        "tag_prefix": "FT-40202 / FT-60201",
        "service": "Slurry & Process Water Volumetric Flow Measurement",
        "mfr": "Endress+Hauser",
        "model": "Promag P 300 (5P3B50 / 80 / 1H PTFE Liner)",
        "proc_conn": "DN50/80/100 ASME Cl.150 RF Flanges (PTFE Lined)",
        "elec_int": "4-20mA HART Active/Passive + Pulse Output",
        "cable_spec": "1Px1.5mm² (Analog) + 1Px1.5mm² (Pulse/Totalizer)",
        "power_supply": "24V DC / 8W (4-Wire Electronics)",
        "haz_class": "ATEX II 2G Ex eb ia IIC T6...T1 Gb",
        "enclosure": "NEMA 4X / IP67 Aluminum Transmitter",
        "plc_module": "1756-IF16 (4-20mA Flow) & 1756-IB32 (Pulse Counter)",
        "signal_type": "4-20mA HART (0-150 m³/h) + Scaled Pulses",
        "dest_jb": "JB-401 / JB-601",
        "barrier": "Internal Galvanic Isolation",
        "bom": [
            ("1", "Promag P 300 Magmeter Sensor & Tx", "E+H PTFE Lined, Hastelloy C Electrodes", "1 EA"),
            ("2", "Grounding Rings / Discs SS316L", "E+H Grounding Electrodes Kit", "2 EA"),
            ("3", "PTFE Envelope Flange Gaskets", "Chemical Resistant Gaskets", "2 EA"),
            ("4", "Cable Glands M20x1.5 Ex d/e", "Nickel-plated Brass Glands", "2 EA"),
            ("5", "Grounding Cables 6mm² Green/Yellow", "Copper Earthing Strap with Lugs", "2 EA")
        ],
        "notes": [
            "Straight piping requirements: min. 5D upstream and 2D downstream of flowmeter.",
            "Install grounding rings between meter flanges and pipe counter-flanges to ensure reference ground.",
            "Pipe must always remain 100% full of conductive liquid for accurate flow measurement."
        ]
    },
    {
        "code": "TYP-08",
        "title": "VORTEX STEAM FLOWMETER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-008",
        "category": "Vortex Flowmeter",
        "tag_prefix": "FT-60101 / FT-60202",
        "service": "High-Pressure Saturated Steam Flow Measurement",
        "mfr": "Endress+Hauser",
        "model": "Prowirl F 200 (7F2C80 / 1H with Temp Sensor)",
        "proc_conn": "DN80/100 (3\"/4\") ASME Cl.300 RF Flange",
        "elec_int": "2-Wire 4-20mA HART Loop Powered (Compensated Steam)",
        "cable_spec": "1Px1.5mm² Twisted Shielded (High Temp Jacketed)",
        "power_supply": "24V DC via 1756-IF16 Analog Channel",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6 Ga/Gb",
        "enclosure": "Dual-Compartment IP67 Die-cast Aluminum",
        "plc_module": "1756-IF16 (Compensated Mass Flow Steam)",
        "signal_type": "4-20mA HART Steam Flow (0-10,000 kg/h)",
        "dest_jb": "JB-601 / JB-602",
        "barrier": "937THAITXPDC1 (IS Loop Powered Repeater)",
        "bom": [
            ("1", "Prowirl F 200 Flanged Vortex Meter", "E+H SS316 Cast Body w/ DSC Sensor", "1 EA"),
            ("2", "Counter-Flanges DN80 Cl.300 RF", "ASTM A105 / A182 F316 Flanges", "2 EA"),
            ("3", "Spiral Wound Gaskets Cl.300 Graphite", "Graphite Filled SS316 Rings", "2 EA"),
            ("4", "Cable Gland M20x1.5 Ex d/e High Temp", "Brass Nickel-Plated Gland", "1 EA"),
            ("5", "Pipeline Insulation Spacer & Shroud", "Thermal Extension Neck Jacket", "1 SET")
        ],
        "notes": [
            "Straight piping: min. 15D upstream and 5D downstream from flow disturbers (valves, reducers).",
            "Insulate pipe line up to meter neck, but leave transmitter electronics housing uninsulated.",
            "Verify that internal integrated Pt100 temperature compensation is activated in HART configuration."
        ]
    },
    {
        "code": "TYP-09",
        "title": "THERMAL MASS GAS FLOWMETER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-009",
        "category": "Thermal Mass Flowmeter",
        "tag_prefix": "FT-80301 Series",
        "service": "LPG / Natural Gas Boiler Fuel Consumption",
        "mfr": "Endress+Hauser",
        "model": "t-mass F 300 / I 300 (6F3B / 6I3B Insertion)",
        "proc_conn": 'DN100 Flange / 1" NPT Retractable Hot-Tap Assembly',
        "elec_int": "4-20mA HART Active (Gas Mass Flow) + Status Relay",
        "cable_spec": "2Px1.5mm² Shielded Armored (Belden 8777)",
        "power_supply": "24V DC / 12W (4-Wire Electronics)",
        "haz_class": "ATEX II 2G Ex db eb ia IIC T4 Gb (Gas Zone 1)",
        "enclosure": "NEMA 4X / IP67 Aluminum Enclosure",
        "plc_module": "1756-IF16 (AI Gas Flow) & 1756-IB32 (DI Alarm)",
        "signal_type": "4-20mA HART Standard Nm³/h Flow",
        "dest_jb": "JB-602 / CA1 Marshalling",
        "barrier": "Galvanic Isolation / Ex d Installation",
        "bom": [
            ("1", "t-mass 300 Thermal Flowmeter Tx", "E+H Insertion Probe w/ Hot-Tap", "1 EA"),
            ("2", "Weld-in Boss 1\" NPT Female Thread", "ASTM A105 / Carbon Steel Boss", "1 EA"),
            ("3", "Isolation Ball Valve 1\" Full Bore", "SS316 Cl.300 Gas Ball Valve", "1 EA"),
            ("4", "Cable Gland M20x1.5 Ex d Flameproof", "Brass Gland for Armored Cable", "2 EA"),
            ("5", "Grounding Cable 6mm² w/ Lug", "Copper Earthing Conductor", "1 EA")
        ],
        "notes": [
            "Install probe at centerline of pipe (0.5 x Inside Diameter) matching flow direction arrow exactly.",
            "Verify straight pipe run: min. 20D upstream and 5D downstream of insertion tap.",
            "Equipotential bonding must be connected between meter body and gas piping system."
        ]
    },
    {
        "code": "TYP-10",
        "title": "ULTRASONIC CLAMP-ON FLOWMETER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-010",
        "category": "Ultrasonic Clamp-on Flowmeter",
        "tag_prefix": "FIT-40105 Series",
        "service": "Cooling Water & Non-Invasive Slurry Verification",
        "mfr": "Keyence",
        "model": "FD-H Series (FD-H20 / FD-H32 Clamp-on)",
        "proc_conn": "External Pipe Surface Mount (No Pipe Penetration)",
        "elec_int": "4-20mA Analog Flow + Pulse / Frequency Output",
        "cable_spec": "4Cx0.75mm² Factory Pre-Wired M12 Connector Cable",
        "power_supply": "20-30V DC / 5W",
        "haz_class": "General Industrial Area / Non-Hazardous",
        "enclosure": "IP67 / IP69K Hygienic Stainless Steel Housing",
        "plc_module": "1756-IF16 (AI Flow Rate) & 1756-IB32 (Pulse)",
        "signal_type": "4-20mA (0-100 L/min) + NPN/PNP Pulse",
        "dest_jb": "JB-401 / CA1 Marshalling",
        "barrier": "Direct 24V DC Field Power",
        "bom": [
            ("1", "Clamp-on Ultrasonic Sensor Head", "Keyence FD-H20 / H32 Unit", "1 EA"),
            ("2", "Acoustic Coupling Gel / Pad", "High-Stability Silicone Couplant", "1 TUBE"),
            ("3", "Pipe Mounting Clamp Bands SS304", "Adjustable Worm Gear Straps", "2 EA"),
            ("4", "M12 4-Pin Shielded Cable (5m)", "PUR Oil-Resistant Flex Cable", "1 EA"),
            ("5", "Field Termination Box (Local)", "IP66 Junction Box w/ DIN Rail", "1 EA")
        ],
        "notes": [
            "Clean pipe surface thoroughly with emery cloth to remove rust, paint, and scale before mounting.",
            "Apply generous bead of acoustic coupling compound between sensor transducers and pipe wall.",
            "Clamp unit securely on horizontal pipe at 3 o'clock or 9 o'clock position (never top or bottom)."
        ]
    },
    {
        "code": "TYP-11",
        "title": "MANUAL HAND VALVE POSITION SWITCH HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-011",
        "category": "Hand Valve with Switch Box",
        "tag_prefix": "ZSO-40101 / ZSC-40101",
        "service": "Manual Safety Isolation Valve Open/Closed Status",
        "mfr": "TeroFox / APL-HKC",
        "model": "TF-20DFS Ball Valve + APL-210N Limit Switch Box",
        "proc_conn": "Flange ASME B16.5 Cl.150 RF, SS316 Body",
        "elec_int": "2x SPDT Mechanical Dry Contact Limit Switches",
        "cable_spec": "4Cx1.5mm² Shielded Instrumentation Cable",
        "power_supply": "24V DC Wetted from PLC 1756-IB32 DI Module",
        "haz_class": "Weatherproof IP67 (Option Ex d IIC T6)",
        "enclosure": "Die-cast Aluminum Powder Coated Beacon Box",
        "plc_module": "1756-IB32 (32-Point 24VDC Digital Input)",
        "signal_type": "DI 1: Valve OPEN (ZSO) / DI 2: Valve CLOSED (ZSC)",
        "dest_jb": "JB-401 / JB-601",
        "barrier": "1492-JD3FB Fuse Protection Terminal (1A)",
        "bom": [
            ("1", "Manual Ball Valve TF-20DFS (DN50)", "TeroFox 2-Piece SS316 Cl.150", "1 EA"),
            ("2", "Limit Switch Box APL-210N", "Yellow/Red Visual Indicator", "1 EA"),
            ("3", "ISO 5211 Mounting Bracket & Drive", "Namur SS304 Bracket", "1 SET"),
            ("4", "Cable Gland M20x1.5 Brass IP66", "Compression Weatherproof Gland", "1 EA"),
            ("5", "Flange Gaskets & Fasteners Kit", "PTFE Gaskets & SS304 Studs/Nuts", "2 SET")
        ],
        "notes": [
            "Adjust internal cam switches: Cam 1 trips at 90° (Fully Open), Cam 2 trips at 0° (Fully Closed).",
            "High-visibility 3D indicator beacon shows YELLOW for OPEN, RED for CLOSED.",
            "Seal unused M20 conduit entry port with certified IP67 threaded blanking plug."
        ]
    },
    {
        "code": "TYP-12",
        "title": "LIQUID LEVEL SWITCH (VIBRATING FORK) HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-012",
        "category": "Liquid Level Switch",
        "tag_prefix": "LSH-40101 / LSL-40101",
        "service": "High / Low Slurry Level Interlock & Pump Protection",
        "mfr": "Endress+Hauser",
        "model": "Liquiphant FTL51B (Tuning Fork SS316L)",
        "proc_conn": '3/4" NPT Male Threaded / Tri-Clamp 2" Hygienic',
        "elec_int": "3-Wire DC-PNP / 2-Wire Relay Contact Output",
        "cable_spec": "3Cx1.5mm² Shielded Instrument Cable",
        "power_supply": "24V DC (10-55V DC, PNP Mode)",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6...T1 Ga/Gb",
        "enclosure": "NEMA 4X / IP66/68 Stainless Steel Housing",
        "plc_module": "1756-IB32 (24V DC Sinking/Sourcing Digital Input)",
        "signal_type": "DI: High Level Alarm (Dry / Submerged State)",
        "dest_jb": "JB-401 / JB-402",
        "barrier": "937THDISTSDC1 (IS Digital Input Barrier)",
        "bom": [
            ("1", "Liquiphant FTL51B Level Switch", "E+H Tuning Fork SS316L (100mm)", "1 EA"),
            ("2", "Welding Boss 3/4\" NPT Half Coupling", "SS316 Tank Welding Socket", "1 EA"),
            ("3", "Cable Gland M20x1.5 Ex d/e", "Nickel-plated Brass Armored Gland", "1 EA"),
            ("4", "Viton / EPDM Thread Seal Ring", "Process Sealing Gasket", "1 EA"),
            ("5", "Grounding Strap 4mm² with Lugs", "Copper Flexible Earth Braid", "1 EA")
        ],
        "notes": [
            "Install fork with tines aligned vertically to prevent slurry buildup on top of vibrating blades.",
            "Set fail-safe mode switch: MAX (Overfill Protection) or MIN (Dry Run Pump Protection).",
            "Keep fork at least 50mm away from tank walls, internal baffles, and agitator blades."
        ]
    },
    {
        "code": "TYP-13",
        "title": "SOLID LEVEL SWITCH (POWDER FORK) HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-013",
        "category": "Solid Level Switch",
        "tag_prefix": "LSH-60101 Series",
        "service": "Starch Powder Silo High Level Overfill Alarm",
        "mfr": "Endress+Hauser",
        "model": "Soliphant FTM51 (Vibrating Fork for Fine Bulk Solids)",
        "proc_conn": '1-1/2" NPT Male / DN50 Cl.150 Flange',
        "elec_int": "24V DC SPDT Relay Dry Contacts (Fail-Safe Max)",
        "cable_spec": "4Cx1.5mm² Shielded Overall (Belden 9534)",
        "power_supply": "24V DC (19-55V DC, 1.5W)",
        "haz_class": "ATEX II 1/2D Ex ta/tb IIIC T120°C Da/Db (Dust)",
        "enclosure": "NEMA 4X / IP66 Cast Aluminum Powder Coated",
        "plc_module": "1756-IB32 (24V DC Digital Input)",
        "signal_type": "DI: Silo Full Alarm Interlock",
        "dest_jb": "JB-601 / CA1 Marshalling",
        "barrier": "1492-JD3FB Fuse Terminal + Relays",
        "bom": [
            ("1", "Soliphant FTM51 Powder Level Switch", "E+H Fork Extension SS316L (300mm)", "1 EA"),
            ("2", "Welding Flange DN50 PN16 / Cl.150", "Carbon Steel / SS316 Flange", "1 EA"),
            ("3", "Weather Protective Sun / Rain Hood", "SS304 Protective Roof", "1 EA"),
            ("4", "Dust-Ex Cable Gland M20x1.5", "Nickel-plated Brass ATEX Dust Gland", "1 EA"),
            ("5", "Grounding Wire 6mm² Green/Yellow", "Copper Earthing Strap", "1 EA")
        ],
        "notes": [
            "Install angled downward 20° to 30° on silo side wall to allow powder to slide off tines easily.",
            "Do not mount directly in the filling stream or impact zone of falling starch powder.",
            "Ensure housing earth terminal is connected to silo metal structure to dissipate electrostatic charge."
        ]
    },
    {
        "code": "TYP-14",
        "title": "RADAR LEVEL TRANSMITTER (80 GHz) HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-014",
        "category": "Radar Level Transmitter",
        "tag_prefix": "LT-40101 / LT-40201",
        "service": "Cooker Slurry & Flash Tank Non-Contact Continuous Level",
        "mfr": "Endress+Hauser",
        "model": "Micropilot FMR67B (80 GHz High-Frequency Radar)",
        "proc_conn": 'DN80 (3") ASME Cl.150 RF Flange w/ PTFE Drip Antenna',
        "elec_int": "2-Wire 4-20mA HART Loop Powered (24V DC)",
        "cable_spec": "1Px1.5mm² Twisted Pair Shielded (Belden 8761)",
        "power_supply": "24V DC via 1756-IF16 Analog Channel",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6 Ga/Gb",
        "enclosure": "NEMA 4X / IP66/68 Dual Compartment Aluminum",
        "plc_module": "1756-IF16 (16-Point Differential AI)",
        "signal_type": "4-20mA HART Level (0-100% / 0-6.0m)",
        "dest_jb": "JB-401 / JB-402",
        "barrier": "937THAITXPDC1 (IS Repeater Barrier)",
        "bom": [
            ("1", "Micropilot FMR67B Radar Transmitter", "E+H 80GHz w/ Drip-off PTFE Antenna", "1 EA"),
            ("2", "Tank Nozzle Flange DN80 Cl.150 RF", "SS316 Tank Flange Neck", "1 EA"),
            ("3", "PTFE Envelope Gasket DN80 Cl.150", "Chemical Resistant Gasket", "1 EA"),
            ("4", "Cable Gland M20x1.5 Ex d/e IP68", "Nickel-plated Brass Gland", "1 EA"),
            ("5", "Air Purge Connection Kit 1/4\" NPT", "Purge Connection for Condensate", "1 SET")
        ],
        "notes": [
            "Mount nozzle height must not exceed antenna length so horn protrudes clear into vessel interior.",
            "Distance to vessel wall should be min. 1/6 of vessel diameter to prevent false sidewall echoes.",
            "Activate Heartbeat Technology and mapping curves to mask agitator blade echoes during commissioning."
        ]
    },
    {
        "code": "TYP-15",
        "title": "PRESSURE GAUGE & SIPHON HOOK-UP & MECHANICAL DETAIL",
        "dwg_no": "KAL-JC-TYP-015",
        "category": "Pressure Gauge",
        "tag_prefix": "PI-40101 / PI-60201",
        "service": "Steam Line & Cooker Chamber Direct Pressure Readout",
        "mfr": "Ashcroft",
        "model": "Model T5500 (100mm Dial All-Stainless Steel)",
        "proc_conn": '1/2" NPT Male w/ Pigtail Siphon & Isolation Needle Valve',
        "elec_int": "Mechanical Local Indicator (No Electrical Wiring)",
        "cable_spec": "N/A (Direct Local Readout)",
        "power_supply": "None (Bourdon Tube Mechanical)",
        "haz_class": "Non-Electrical Apparatus (Standard Zone 1/2)",
        "enclosure": "IP65 Stainless Steel Case, Glycerin / Dry",
        "plc_module": "None (Local Operator Gauge)",
        "signal_type": "Visual Dial Scale 0-16 bar (0-230 psi)",
        "dest_jb": "Direct Piping Mount",
        "barrier": "N/A",
        "bom": [
            ("1", "Pressure Gauge T5500 (0-16 bar)", "Ashcroft 100mm Dial SS316 Bourdon", "1 EA"),
            ("2", "Pigtail Siphon Schedule 80 SS316", "1/2\" NPT Steam Condensate Loop", "1 EA"),
            ("3", "Needle Isolation Valve 1/2\" NPT", "SS316 Cl.6000 Needle Valve", "1 EA"),
            ("4", "Gauge Snubber / Pulsation Dampener", "Porous SS316 Disc 1/2\" NPT", "1 EA"),
            ("5", "Thread Sealant Tape (PTFE/Graphite)", "High Temperature Steam Rated", "1 ROLL")
        ],
        "notes": [
            "Pigtail siphon must be pre-filled with clean water prior to steam startup to protect Bourdon tube.",
            "Orient dial facing primary walkway / operating platform for ease of operator reading.",
            "Close needle valve slightly during excessive pulsating flow to protect internal linkage."
        ]
    },
    {
        "code": "TYP-16",
        "title": "ELECTRONIC PRESSURE SWITCH HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-016",
        "category": "Pressure Switch",
        "tag_prefix": "PSH-60201 / PSL-60201",
        "service": "High / Low Pressure Safety Trip Interlock",
        "mfr": "Endress+Hauser",
        "model": "Ceraphant PTC31B (Ceramic Sensor Diaphragm)",
        "proc_conn": '1/2" NPT Male Threaded / Tri-Clamp Sanitary',
        "elec_int": "24V DC 3-Wire PNP Switch Output (Transistor)",
        "cable_spec": "3Cx1.5mm² Shielded Instrumentation Cable",
        "power_supply": "24V DC (10-30V DC, 20mA)",
        "haz_class": "General Industrial Zone (Option Ex ec IIC T6)",
        "enclosure": "IP65 / IP67 Stainless Steel 316L Housing",
        "plc_module": "1756-IB32 (24V DC Digital Input)",
        "signal_type": "DI: Trip Contact (Open on Alarm)",
        "dest_jb": "JB-602 / CA1 Marshalling",
        "barrier": "1492-JD3FB Fuse Protection Terminal",
        "bom": [
            ("1", "Ceraphant PTC31B Pressure Switch", "E+H Ceramic Cell, LED Display", "1 EA"),
            ("2", "2-Valve Isolation Manifold 1/2\" NPT", "SS316 Block & Bleed Valve", "1 EA"),
            ("3", "M12 4-Pin Field Wireable Connector", "IP67 Female M12 Connector", "1 EA"),
            ("4", "Cable Gland M16x1.5 Brass", "Nickel-plated Cable Gland", "1 EA"),
            ("5", "Impulse Adapter 1/2\" NPT x 1/2\" Tube", "Swagelok SS316 Tube Fitting", "1 EA")
        ],
        "notes": [
            "Program switch point and reset hysteresis via front push-buttons and LED display.",
            "Configure output as Normally Closed (NC) fail-safe: circuit de-energizes on overpressure or power loss.",
            "Provide block-and-bleed valve for routine verification with field calibrator."
        ]
    },
    {
        "code": "TYP-17",
        "title": "PRESSURE TRANSMITTER HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-017",
        "category": "Pressure Transmitter",
        "tag_prefix": "PT-40101 / PT-60202",
        "service": "Continuous Pressure Monitoring for Cooker & Flash Tank",
        "mfr": "Endress+Hauser",
        "model": "Cerabar PMP51B / PMC51B (Metallic / Ceramic)",
        "proc_conn": '1/2" NPT Male Thread / 2" Tri-Clamp Diaphragm',
        "elec_int": "2-Wire 4-20mA HART Loop Powered (24V DC)",
        "cable_spec": "1Px1.5mm² Twisted Shielded (Belden 8761)",
        "power_supply": "24V DC via 1756-IF16 Analog Channel",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6 Ga/Gb",
        "enclosure": "NEMA 4X / IP66/68 Dual Compartment Aluminum",
        "plc_module": "1756-IF16 (16-Point Differential AI)",
        "signal_type": "4-20mA HART (0-10 bar / 0-16 bar gauge)",
        "dest_jb": "JB-401 / JB-602",
        "barrier": "937THAITXPDC1 (IS Repeater Barrier)",
        "bom": [
            ("1", "Cerabar PMP51B Pressure Transmitter", "E+H Piezoresistive SS316L Cell", "1 EA"),
            ("2", "2-Valve Block & Bleed Manifold", "SS316 1/2\" NPT Male x Female", "1 EA"),
            ("3", "2\" Pipe Mounting Bracket Kit", "SS304 Universal Bracket & U-Bolts", "1 SET"),
            ("4", "Cable Gland M20x1.5 Ex d/e IP68", "Nickel-plated Brass Armored Gland", "1 EA"),
            ("5", "Siphon / Cooling Neck (Steam Duty)", "SS316 Pigtail Loop 1/2\" NPT", "1 EA")
        ],
        "notes": [
            "Use flush diaphragm seal (PMC51B) for thick starch slurry to eliminate clogging in impulse ports.",
            "Ground transmitter body to local pipe ground; shield wire grounded only at CA1 cabinet earth bar.",
            "HART configuration: Set damping to 1.0s to filter process pump pulsations."
        ]
    },
    {
        "code": "TYP-18",
        "title": "BIMETAL TEMPERATURE GAUGE & THERMOWELL HOOK-UP DETAIL",
        "dwg_no": "KAL-JC-TYP-018",
        "category": "Temperature Gauge",
        "tag_prefix": "TI-40101 / TI-60201",
        "service": "Direct Pipe & Tank Slurry Temperature Indicator",
        "mfr": "Ashcroft",
        "model": "Model FI (100mm Dial Everyangle Bimetal) + Thermowell",
        "proc_conn": '1/2" NPT Male Thread into Barstock Thermowell',
        "elec_int": "Mechanical Local Indicator (No Electrical Wiring)",
        "cable_spec": "N/A (Local Dial Readout)",
        "power_supply": "None (Bimetal Coil Driven)",
        "haz_class": "Non-Electrical Apparatus (Standard Zone 1/2)",
        "enclosure": "IP65 Hermetically Sealed Stainless Steel",
        "plc_module": "None (Local Operator Readout)",
        "signal_type": "Visual Scale 0-150°C (32-300°F)",
        "dest_jb": "Direct Piping Well Mount",
        "barrier": "N/A",
        "bom": [
            ("1", "Bimetal Thermometer Model FI (0-150°C)", "Ashcroft 100mm Dial Everyangle Stem", "1 EA"),
            ("2", "Barstock Thermowell SS316 (U=150mm)", "1/2\" NPT x 1\" NPT Process Thread", "1 EA"),
            ("3", "Weldolet / Threadolet 1\" NPT", "ASTM A182 F316 Pipe Fitting", "1 EA"),
            ("4", "Thermal Conductive Paste", "High Temperature Aluminum Oxide Gel", "1 TUBE"),
            ("5", "Safety Wire & Tag Plate SS316", "Laser Engraved Tag Plate", "1 EA")
        ],
        "notes": [
            "Thermowell tip must be immersed into center third of process pipe for representative measurement.",
            "Apply thermal conductive paste inside thermowell bore before inserting gauge stem to improve heat transfer.",
            "Thermowell allows removal of thermometer for calibration without process shutdown."
        ]
    },
    {
        "code": "TYP-19",
        "title": "TEMPERATURE TRANSMITTER (RTD Pt100) HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-019",
        "category": "Temp Transmitter (RTD)",
        "tag_prefix": "TT-40101 / TT-60201",
        "service": "Precision Slurry Cooking Temperature Control",
        "mfr": "Endress+Hauser",
        "model": "iTEMP TMT71 + Omnigrad M TM411 (Pt100 Class A)",
        "proc_conn": '1/2" NPT Thread into Barstock Thermowell (SS316L)',
        "elec_int": "2-Wire 4-20mA HART Loop Powered (Head Mounted)",
        "cable_spec": "1Px1.5mm² Twisted Shielded (Belden 8761)",
        "power_supply": "24V DC via 1756-IF16 Analog Channel",
        "haz_class": "ATEX II 1/2G Ex ia IIC T6...T4 Ga/Gb",
        "enclosure": "IP66 / IP68 Form B Head Aluminum Housing",
        "plc_module": "1756-IF16 (16-Point Differential AI)",
        "signal_type": "4-20mA HART Temp (0-150°C / 0-200°C)",
        "dest_jb": "JB-401 / JB-602",
        "barrier": "937THAITXPDC1 (IS Temperature Repeater)",
        "bom": [
            ("1", "Head Transmitter iTEMP TMT71", "E+H 4-20mA HART Head Mounted Puck", "1 EA"),
            ("2", "RTD Sensor Assembly TM411 Pt100", "Class A 3-Wire Mineral Insulated", "1 EA"),
            ("3", "Flanged / Threaded Thermowell SS316L", "Drilled Barstock Tapered Well", "1 EA"),
            ("4", "Cable Gland M20x1.5 Ex d/e IP68", "Nickel-plated Brass Gland", "1 EA"),
            ("5", "Grounding Strap 4mm² with Lugs", "Flexible Copper Earthing Braid", "1 EA")
        ],
        "notes": [
            "Connect Pt100 RTD using 3-wire compensation circuit to eliminate field lead wire resistance error.",
            "Install thermowell at elbow pointing upstream or inclined 45° if pipe diameter is under 4 inches.",
            "Verify loop current with multimeter test points on TMT71 without interrupting signal loop."
        ]
    },
    {
        "code": "TYP-20",
        "title": "PNEUMATIC ON-OFF CONTROL VALVE HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-020",
        "category": "Pneumatic On-Off Valve",
        "tag_prefix": "XV-40101 / XV-60201",
        "service": "Automated Slurry Routing & Emergency Process Isolation",
        "mfr": "TeroFox / El-O-Matic",
        "model": "TF-20DFS Ball Valve + F-Series Actuator + Solenoid + APL-210N",
        "proc_conn": "Flange ASME B16.5 Cl.150 RF, SS316 Body",
        "elec_int": "DO: 24VDC Solenoid (5W) + DI: 2x Limit Switches (ZSO/ZSC)",
        "cable_spec": "3Cx1.5mm² (Solenoid) + 4Cx1.5mm² (Limit Switches)",
        "power_supply": "24V DC from 1756-OB32 + 6 bar Clean Instrument Air",
        "haz_class": "ATEX II 2G Ex d IIC T6 Gb / IP67",
        "enclosure": "Hard Anodized Aluminum Actuator & IP67 Beacon",
        "plc_module": "1756-OB32 (DO Solenoid) & 1756-IB32 (DI Open/Closed)",
        "signal_type": "DO: 24VDC Energize to Open / DI: ZSO, ZSC",
        "dest_jb": "JB-401 / JB-602",
        "barrier": "700-HL Interposing Relay (DO) + Fuses",
        "bom": [
            ("1", "Pneumatic Ball Valve Assembly", "TeroFox TF-20DFS DN50 Cl.150", "1 EA"),
            ("2", "Pneumatic Spring-Return Actuator", "El-O-Matic F-Series Rack & Pinion", "1 EA"),
            ("3", "Namur Solenoid Valve 5/2-Way (24VDC)", "Asco / SMC Low Power Solenoid", "1 EA"),
            ("4", "Limit Switch Box APL-210N (2x SPDT)", "Visual Yellow/Red Open/Close Beacon", "1 EA"),
            ("5", "Air Filter Regulator & Gauge SMC", "Modular AFR 1/4\" w/ Auto-Drain", "1 SET")
        ],
        "notes": [
            "Spring-return fail-safe action: Valve drives to CLOSED position on electrical power or air failure.",
            "Interposing relay 700-HL at CA1 cabinet isolates PLC 1756-OB32 output from field coil inrush.",
            "Air supply lines must use seamless SS316 1/4\" or 3/8\" tubing with Swagelok compression fittings."
        ]
    },
    {
        "code": "TYP-21",
        "title": "MODULATING GLOBE CONTROL VALVE HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-021",
        "category": "Globe Control Valve",
        "tag_prefix": "TCV-40222 / TV-60202",
        "service": "High-Precision Steam & Slurry Flow Modulation",
        "mfr": "Samson",
        "model": "Type 3241 Globe Valve + Trovis 3730-1 Smart Positioner",
        "proc_conn": "Flange ASME B16.5 Cl.150/300 RF (DN50/80/100)",
        "elec_int": "AO: 4-20mA Setpoint (Positioner) + AI: 4-20mA Position Feedback",
        "cable_spec": "2Px1.5mm² Twisted Shielded (Belden 8777)",
        "power_supply": "2-Wire 4-20mA Loop + 6.0 bar Instrument Air",
        "haz_class": "ATEX II 2G Ex ia IIC T6 / Ex d IIC T6",
        "enclosure": "IP66 Die-cast Aluminum Positioner Case",
        "plc_module": "1756-OF8 (AO Setpoint) & 1756-IF16 (AI Feedback)",
        "signal_type": "AO: 4-20mA Valve Cmd (0-100%) / AI: Position",
        "dest_jb": "JB-402 / JB-602",
        "barrier": "Isolated Analog Output / IS Repeater",
        "bom": [
            ("1", "Globe Control Valve Type 3241", "Samson Cast Steel / SS316 Body", "1 EA"),
            ("2", "Pneumatic Diaphragm Actuator Type 3271", "Samson Multi-Spring Actuator", "1 EA"),
            ("3", "Smart Positioner Trovis 3730-1", "HART Communication, Auto-Tuning", "1 EA"),
            ("4", "SMC Air Filter Regulator AW20", "Clean Air 1/4\" w/ Pressure Gauge", "1 SET"),
            ("5", "Cable Glands M20x1.5 Ex d/e IP68", "Nickel-plated Brass Glands", "2 EA")
        ],
        "notes": [
            "Perform Trovis 3730-1 automatic initialization routine after mechanical mounting is completed.",
            "Connect 4-20mA setpoint from PLC 1756-OF8 to Positioner Terminals 11(+) and 12(-).",
            "Ensure air supply pressure to positioner is regulated strictly between 4.0 and 6.0 bar."
        ]
    },
    {
        "code": "TYP-22",
        "title": "ELECTRONIC TEMPERATURE SWITCH HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-022",
        "category": "Temperature Switch",
        "tag_prefix": "TSH-40101 / TSH-60201",
        "service": "High Temperature Safety Trip & Alarm",
        "mfr": "Endress+Hauser",
        "model": "Thermophant TTR35 (Compact Pt100 Switch)",
        "proc_conn": '1/2" NPT Male Threaded / Hygienic Clamp',
        "elec_int": "24V DC 3-Wire PNP Switch Output + Local LED Display",
        "cable_spec": "3Cx1.5mm² Shielded Instrument Cable",
        "power_supply": "24V DC (10-30V DC, 35mA)",
        "haz_class": "General Industrial / Hygienic Zone",
        "enclosure": "IP65 / IP67 Stainless Steel 316L Housing",
        "plc_module": "1756-IB32 (24V DC Digital Input)",
        "signal_type": "DI: High Temperature Alarm Trip",
        "dest_jb": "JB-401 / JB-602",
        "barrier": "1492-JD3FB Fuse Protection Terminal",
        "bom": [
            ("1", "Thermophant TTR35 Temp Switch", "E+H Pt100 Class A, Digital Display", "1 EA"),
            ("2", "Barstock Thermowell SS316L (U=100mm)", "1/2\" NPT x 1/2\" NPT Process Well", "1 EA"),
            ("3", "M12 4-Pin Female Cable Assembly (5m)", "M12 Right Angle Molded Connector", "1 EA"),
            ("4", "Threadolet / Boss 1/2\" NPT SS316", "Pipe Welding Fitting", "1 EA"),
            ("5", "Grounding Clamp & Lug", "Copper Earth Wire", "1 SET")
        ],
        "notes": [
            "Configure switch point via optical touch keys on front cover without opening enclosure.",
            "Normally closed (NC) contact output configuration guarantees fail-safe alarm on wire break.",
            "Thermowell allows switch replacement without draining the process line."
        ]
    },
    {
        "code": "TYP-23",
        "title": "PNEUMATIC TURBINE VIBRATOR HOOK-UP & LOOP DIAGRAM",
        "dwg_no": "KAL-JC-TYP-023",
        "category": "Pneumatic Turbine Vibrator",
        "tag_prefix": "XV-60901 Series",
        "service": "Powder Silo & Hopper Wall Discharge Agitation",
        "mfr": "Netter Vibration",
        "model": "NCT 5 Pneumatic Turbine Vibrator",
        "proc_conn": 'G 1/8" BSP Air Port / Rigid Base Mounting Plate',
        "elec_int": "24V DC Solenoid Valve Control (4W)",
        "cable_spec": "3Cx1.5mm² Shielded Cable (Belden 9533)",
        "power_supply": "24V DC from 1756-OB32 + 2 to 6 bar Lubricated/Dry Air",
        "haz_class": "ATEX II 2GD c T6 85°C (Mechanical Ex)",
        "enclosure": "Extruded Aluminum Hard Anodized Body",
        "plc_module": "1756-OB32 (24V DC Digital Output)",
        "signal_type": "DO: Pulsed Vibration Burst (Timer Controlled)",
        "dest_jb": "JB-601 / CA1 Marshalling",
        "barrier": "700-HL Interposing Relay with Diode",
        "bom": [
            ("1", "Pneumatic Turbine Vibrator NCT 5", "Netter Vibration High-Speed Unit", "1 EA"),
            ("2", "Solenoid Valve 3/2-Way 24VDC (G 1/4\")", "High Flow Fast-Cycling Solenoid", "1 EA"),
            ("3", "Silencer / Exhaust Air Filter G 1/8\"", "Sintered Bronze Sinter Filter", "1 EA"),
            ("4", "SMC Air Filter Regulator Kit", "Modular AFR w/ Pressure Gauge", "1 SET"),
            ("5", "Weld-on Hopper Mounting Channel Base", "Reinforced Channel Profile Steel", "1 EA")
        ],
        "notes": [
            "Mount vibrator on stiffened channel beam welded to hopper cone to transmit vibration into bulk powder.",
            "Operate in intermittent pulse mode (e.g. 5 seconds ON, 30 seconds OFF) via PLC timer logic.",
            "Interposing relay 700-HL at CA1 panel protects PLC output module from inductive flyback."
        ]
    }
]

# -----------------------------------------------------------------------------
# STEP 2: PROFESSIONAL ENGINEERING PDF GENERATOR (A3 LANDSCAPE)
# -----------------------------------------------------------------------------
class InstrumentTemplatePDF(FPDF):
    def __init__(self):
        super().__init__(orientation="landscape", unit="mm", format="A3")
        self.set_margins(10, 10, 10)
        self.set_auto_page_break(auto=False)
        self.total_sheets = 25

    def draw_frame(self, sheet_no, title, dwg_no="KAL-JC-TYP-XXXXX", scale="N.T.S."):
        margin = 10
        w = 420 - 2 * margin  # 400 mm
        h = 297 - 2 * margin  # 277 mm

        # Outer heavy border (0.7mm)
        self.set_draw_color(25, 40, 65)
        self.set_line_width(0.7)
        self.rect(margin, margin, w, h)

        # Inner border (0.35mm)
        self.set_line_width(0.35)
        self.rect(margin + 2, margin + 2, w - 4, h - 4)

        # Coordinate Grid 1..8
        self.set_font("Helvetica", "B", 6.5)
        self.set_text_color(110, 120, 140)
        for i, lbl in enumerate(["1", "2", "3", "4", "5", "6", "7", "8"]):
            x = margin + (i + 0.5) * (w / 8)
            self.text(x, margin + 1.6, lbl)
            self.text(x, margin + h - 0.5, lbl)
            self.line(margin + i * (w / 8), margin, margin + i * (w / 8), margin + 2)
            self.line(margin + i * (w / 8), margin + h - 2, margin + i * (w / 8), margin + h)

        # Coordinate Grid A..F
        for i, lbl in enumerate(["A", "B", "C", "D", "E", "F"]):
            y = margin + (i + 0.5) * (h / 6)
            self.text(margin + 0.7, y, lbl)
            self.text(margin + w - 1.6, y, lbl)
            self.line(margin, margin + i * (h / 6), margin + 2, margin + i * (h / 6))
            self.line(margin + w - 2, margin + i * (h / 6), margin + w, margin + i * (h / 6))

        # Title Block (Bottom Right: 175 x 36 mm)
        tb_w = 175.0
        tb_h = 36.0
        tb_x = margin + w - tb_w
        tb_y = margin + h - tb_h

        self.set_fill_color(255, 255, 255)
        self.set_draw_color(25, 40, 65)
        self.set_line_width(0.5)
        self.rect(tb_x, tb_y, tb_w, tb_h, "DF")

        # Divisions
        self.line(tb_x, tb_y + 11, tb_x + tb_w, tb_y + 11)
        self.line(tb_x, tb_y + 23, tb_x + tb_w, tb_y + 23)
        self.line(tb_x + 110, tb_y, tb_x + 110, tb_y + 23)
        self.line(tb_x + 55, tb_y + 23, tb_x + 55, tb_y + tb_h)
        self.line(tb_x + 110, tb_y + 23, tb_x + 110, tb_y + tb_h)
        self.line(tb_x + 145, tb_y + 23, tb_x + 145, tb_y + tb_h)

        # Client Section
        self.set_text_color(20, 35, 60)
        self.set_font("Helvetica", "B", 7.5)
        self.text(tb_x + 3, tb_y + 4.5, "INGREDION (THAILAND) CO., LTD.")
        self.set_font("Helvetica", "", 6.5)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 3, tb_y + 8.5, "KALASIN STARCH PLANT - JET COOKER PROJECT")

        # Contractor Section
        self.set_text_color(20, 35, 60)
        self.set_font("Helvetica", "B", 7.5)
        self.text(tb_x + 113, tb_y + 4.5, "AEC INDUSTRIAL ENGINEERING")
        self.set_font("Helvetica", "", 6.0)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 113, tb_y + 8.5, "SYSTEMS INTEGRATION & AUTOMATION")

        # Drawing Title
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(15, 30, 60)
        lines = self.multi_cell(105, 3.5, title, dry_run=True, output="LINES")
        if len(lines) == 1:
            self.text(tb_x + 3, tb_y + 16, title)
            self.set_font("Helvetica", "", 6.0)
            self.set_text_color(90, 100, 115)
            self.text(tb_x + 3, tb_y + 20.5, "INSTRUMENT HOOK-UP & TYPICAL LOOP DIAGRAM")
        else:
            self.text(tb_x + 3, tb_y + 15, lines[0])
            self.text(tb_x + 3, tb_y + 19, lines[1] if len(lines) > 1 else "")

        # Drawing Number & Rev
        self.set_font("Helvetica", "B", 6.8)
        self.set_text_color(25, 35, 55)
        self.text(tb_x + 113, tb_y + 15.5, "DWG NO:")
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(10, 25, 80)
        self.text(tb_x + 128, tb_y + 15.5, dwg_no)
        self.set_font("Helvetica", "", 6.5)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 113, tb_y + 20, "REV: 3.6    DATE: 2026-09-08")

        # Drawn / Checked / Scale / Sheet
        self.set_font("Helvetica", "", 6.2)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 3, tb_y + 27, "DRAWN: AEC-ENG")
        self.text(tb_x + 3, tb_y + 32, "CHECKED: LEAD-PE")

        self.set_font("Helvetica", "B", 6.5)
        self.set_text_color(25, 90, 45)
        self.text(tb_x + 58, tb_y + 27, "STATUS: APPROVED")
        self.set_font("Helvetica", "", 6.2)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 58, tb_y + 32, f"SCALE: {scale}")

        self.set_font("Helvetica", "B", 6.2)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 113, tb_y + 27, "PROJECT NO:")
        self.set_font("Helvetica", "", 6.5)
        self.set_text_color(20, 35, 60)
        self.text(tb_x + 113, tb_y + 32, "R5THS00138-JC")

        self.set_font("Helvetica", "B", 6.2)
        self.set_text_color(70, 80, 95)
        self.text(tb_x + 148, tb_y + 27, "SHEET NO:")
        self.set_font("Helvetica", "B", 8.0)
        self.set_text_color(15, 30, 60)
        self.text(tb_x + 148, tb_y + 32.5, f"{sheet_no:02d} / {self.total_sheets:02d}")


# -----------------------------------------------------------------------------
# STEP 3: RENDER COVER, INDEX & 23 TEMPLATE SHEETS
# -----------------------------------------------------------------------------
def render_template_cover_sheet(pdf):
    pdf.add_page()
    pdf.draw_frame(1, "INSTRUMENT HOOK-UP & LOOP DRAWING DIRECTORY", "KAL-JC-TYP-000")

    margin = 12
    # Title Banner
    pdf.set_fill_color(24, 43, 73)
    pdf.rect(margin + 20, 30, 360, 42, "F")

    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 18)
    pdf.text(margin + 30, 46, "KALASIN STARCH PLANT - JET COOKER PROJECT")
    pdf.set_font("Helvetica", "", 12)
    pdf.set_text_color(200, 225, 250)
    pdf.text(margin + 30, 56, "STANDARD INSTRUMENT HOOK-UP & TYPICAL ELECTRICAL LOOP DRAWINGS")
    pdf.set_font("Helvetica", "I", 9.0)
    pdf.text(margin + 30, 65, "Based on Equipment Packages in 11-instrument Manual/ & Master I/O Mapping Rev 3.6")

    card_y = 82
    # Left Card: Executive Description
    pdf.set_fill_color(248, 250, 252)
    pdf.set_draw_color(210, 220, 230)
    pdf.set_line_width(0.4)
    pdf.rect(margin + 20, card_y, 175, 120, "DF")

    pdf.set_fill_color(30, 55, 90)
    pdf.rect(margin + 20, card_y, 175, 10, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    pdf.text(margin + 25, card_y + 7, "DRAWING PACKAGE OBJECTIVES & ENGINEERING STANDARDS")

    objs = [
        ("Project Scope:", "Jet Cooker Slurry & Flash Tank Instrumentation"),
        ("Client / Owner:", "INGREDION (THAILAND) CO., LTD."),
        ("Plant Location:", "Kalasin Starch Manufacturing Plant, Thailand"),
        ("System Integrator:", "AEC INDUSTRIAL ENGINEERING CO., LTD."),
        ("Covered Packages:", "23 Complete Equipment Types (CS-DS-01 to CS-DS-22)"),
        ("Drawing Contents:", "Piping Hook-Up, Loop Wiring Schematic & BOM"),
        ("Electrical Standard:", "Rockwell ControlLogix 1756 Architecture (24V DC / 4-20mA)"),
        ("Hazardous Areas:", "Ex ia / Ex d Classifications (937T Barriers & Glands)"),
        ("CAD Standard:", "1:1 Metric Millimeters (AutoCAD DXF R2010 Layered)"),
        ("Piping Standard:", "ASME B16.5 Cl.150/300 RF Flanges & NPT Threads"),
        ("Document Status:", "APPROVED ENGINEERING TEMPLATES (Rev 3.6)")
    ]
    py = card_y + 18
    for lbl, val in objs:
        pdf.set_font("Helvetica", "B", 7.2)
        pdf.set_text_color(40, 50, 70)
        pdf.text(margin + 25, py, lbl)
        pdf.set_font("Helvetica", "", 7.2)
        pdf.set_text_color(20, 30, 50)
        pdf.text(margin + 68, py, val)
        py += 9.2

    # Right Card: Standard Legend & Color Conventions
    pdf.set_fill_color(248, 250, 252)
    pdf.rect(margin + 205, card_y, 175, 120, "DF")
    pdf.set_fill_color(30, 55, 90)
    pdf.rect(margin + 205, card_y, 175, 10, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    pdf.text(margin + 210, card_y + 7, "DRAWING SYMBOLOGY & CAD LAYER CONVENTIONS")

    legend = [
        ("Process Piping:", "Heavy solid line (0.7mm), showing process flow direction"),
        ("Impulse Tubing:", "Medium dashed line, 1/2\" OD seamless SS316 tubing"),
        ("Instrument Air:", "Pneumatic air supply line, 6.0 bar dry filtered air"),
        ("Analog Signal (AI/AO):", "4-20mA HART loop, shielded twisted pair (Belden 8761)"),
        ("Discrete Signal (DI/DO):", "24V DC wetted contact or solenoid drive (Belden 9533)"),
        ("IS Barrier (Ex ia):", "Sky Blue block, Rockwell 937T Series Intrinsically Safe"),
        ("Interposing Relay:", "Orange block, Allen-Bradley 700-HL Series 24VDC Relay"),
        ("Terminal Strip:", "Green circle/box, Marshalling Panel CA1 (C#S##-X1)"),
        ("Earth Grounding:", "Protective Earth (PE) & Instrument Isolated Clean Earth (IE)")
    ]
    ly = card_y + 18
    for lbl, val in legend:
        pdf.set_font("Helvetica", "B", 7.0)
        pdf.set_text_color(25, 45, 75)
        pdf.text(margin + 210, ly, lbl)
        pdf.set_font("Helvetica", "", 6.8)
        pdf.set_text_color(60, 70, 85)
        pdf.text(margin + 210, ly + 4.5, val)
        ly += 11.2

    # Signatures
    sy = card_y + 130
    pdf.set_fill_color(248, 250, 252)
    pdf.rect(margin + 20, sy, 360, 32, "DF")
    cols = [
        ("PREPARED BY", "AEC INDUSTRIAL ENG.", "Lead Instrumentation Eng."),
        ("CHECKED BY", "AEC AUTOMATION LEAD", "System Architect & PE"),
        ("VERIFIED BY", "PROJECT ELECTRICAL PE", "Licensed Professional Eng."),
        ("APPROVED BY", "INGREDION PROJECT MGR.", "Client Automation Lead")
    ]
    for idx, (role, person, title) in enumerate(cols):
        cx = margin + 25 + idx * 89
        pdf.set_font("Helvetica", "B", 6.5)
        pdf.set_text_color(100, 110, 125)
        pdf.text(cx, sy + 6, role)
        pdf.set_font("Helvetica", "B", 7.5)
        pdf.set_text_color(20, 35, 60)
        pdf.text(cx, sy + 15, person)
        pdf.set_font("Helvetica", "", 6.2)
        pdf.set_text_color(80, 90, 105)
        pdf.text(cx, sy + 21, title)
        pdf.set_draw_color(180, 190, 205)
        pdf.line(cx, sy + 26, cx + 75, sy + 26)


def render_template_index_sheet(pdf):
    pdf.add_page()
    pdf.draw_frame(2, "INSTRUMENT TEMPLATE DRAWING DIRECTORY", "KAL-JC-TYP-000-IDX")

    margin = 14
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(20, 35, 60)
    pdf.text(margin + 5, 25, "MASTER DRAWING INDEX - 23 INSTRUMENT HOOK-UP & LOOP DIAGRAM TEMPLATES")

    col_w = 185
    rows_per_col = 13

    for col_idx in range(2):
        cx = margin + 5 + col_idx * 195
        cy = 32

        pdf.set_fill_color(28, 48, 80)
        pdf.rect(cx, cy, col_w, 7.5, "F")
        pdf.set_font("Helvetica", "B", 6.5)
        pdf.set_text_color(255, 255, 255)
        pdf.text(cx + 2, cy + 5.2, "SHT")
        pdf.text(cx + 12, cy + 5.2, "DRAWING NO")
        pdf.text(cx + 45, cy + 5.2, "EQUIPMENT CATEGORY / TEMPLATE TITLE")
        pdf.text(cx + 145, cy + 5.2, "REV")
        pdf.text(cx + 160, cy + 5.2, "STATUS")

        start_idx = col_idx * rows_per_col
        end_idx = min(len(TEMPLATES_DATA), (col_idx + 1) * rows_per_col)

        ry = cy + 7.5
        for t_idx in range(start_idx, end_idx):
            item = TEMPLATES_DATA[t_idx]
            s_no = 3 + t_idx
            bg = (255, 255, 255) if t_idx % 2 == 0 else (245, 248, 252)
            pdf.set_fill_color(*bg)
            pdf.rect(cx, ry, col_w, 14.5, "F")

            pdf.set_font("Helvetica", "B", 6.8)
            pdf.set_text_color(30, 45, 70)
            pdf.text(cx + 2, ry + 6.0, f"{s_no:02d}")

            pdf.set_font("Helvetica", "", 6.2)
            pdf.text(cx + 12, ry + 6.0, item["dwg_no"])

            pdf.set_font("Helvetica", "B", 6.5)
            pdf.set_text_color(15, 30, 60)
            title_text = item["title"]
            if len(title_text) > 42:
                title_text = title_text[:40] + ".."
            pdf.text(cx + 45, ry + 5.5, title_text)

            pdf.set_font("Helvetica", "", 5.5)
            pdf.set_text_color(90, 100, 115)
            pdf.text(cx + 45, ry + 10.5, f"Mfr: {item['mfr']} | Model: {item['model'][:28]}")

            pdf.set_font("Helvetica", "", 6.2)
            pdf.set_text_color(30, 45, 70)
            pdf.text(cx + 145, ry + 6.0, "3.6")

            pdf.set_font("Helvetica", "B", 5.8)
            pdf.set_text_color(20, 100, 40)
            pdf.text(cx + 160, ry + 6.0, "APPROVED")

            pdf.set_draw_color(225, 230, 240)
            pdf.line(cx, ry + 14.5, cx + col_w, ry + 14.5)
            ry += 14.5


def render_instrument_template_sheet(pdf, sheet_no, data):
    pdf.add_page()
    pdf.draw_frame(sheet_no, data["title"], data["dwg_no"])

    margin = 12
    ox = margin + 3
    oy = 22

    # Top Banner with Tag Prefix & Equipment Category
    pdf.set_fill_color(242, 246, 252)
    pdf.set_draw_color(180, 200, 225)
    pdf.set_line_width(0.4)
    pdf.rect(ox, oy, 394, 11, "DF")

    pdf.set_font("Helvetica", "B", 8.0)
    pdf.set_text_color(15, 35, 70)
    pdf.text(ox + 4, oy + 7.5, f"TEMPLATE {data['code']}: {data['category'].upper()}")

    pdf.set_font("Helvetica", "B", 6.8)
    pdf.set_text_color(30, 60, 100)
    pdf.text(ox + 130, oy + 7.5, f"TYPICAL TAGS: {data['tag_prefix']}")

    pdf.set_font("Helvetica", "", 6.5)
    pdf.set_text_color(80, 95, 115)
    pdf.text(ox + 235, oy + 7.5, f"SERVICE: {data['service']}  |  PLC I/O: {data['plc_module']}")

    # -------------------------------------------------------------------------
    # PANE 1: PROCESS PIPING & MECHANICAL HOOK-UP (Left: Width 122mm)
    # -------------------------------------------------------------------------
    p1_x = ox
    p1_y = oy + 14
    p1_w = 122.0
    p1_h = 210.0

    pdf.set_fill_color(252, 253, 255)
    pdf.set_draw_color(200, 210, 225)
    pdf.rect(p1_x, p1_y, p1_w, p1_h, "DF")

    # Header
    pdf.set_fill_color(30, 50, 80)
    pdf.rect(p1_x, p1_y, p1_w, 7, "F")
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.set_text_color(255, 255, 255)
    pdf.text(p1_x + 4, p1_y + 5, "SECTION A: PROCESS PIPING & MECHANICAL HOOK-UP")

    # Mechanical Schematic Drawings (Vector Graphics inside Pane 1)
    diag_y = p1_y + 12
    # Draw Process Pipe Line
    pdf.set_draw_color(50, 70, 95)
    pdf.set_line_width(0.8)
    pdf.line(p1_x + 10, diag_y + 40, p1_x + 112, diag_y + 40)
    pdf.line(p1_x + 10, diag_y + 52, p1_x + 112, diag_y + 52)
    pdf.set_font("Helvetica", "B", 6.0)
    pdf.set_text_color(70, 85, 105)
    pdf.text(p1_x + 12, diag_y + 36, "PROCESS SLURRY / STEAM PIPE")
    # Flow arrow
    pdf.line(p1_x + 55, diag_y + 46, p1_x + 75, diag_y + 46)
    pdf.line(p1_x + 70, diag_y + 43, p1_x + 75, diag_y + 46)
    pdf.line(p1_x + 70, diag_y + 49, p1_x + 75, diag_y + 46)
    pdf.text(p1_x + 56, diag_y + 44, "FLOW DIRECTION")

    # Instrument Mounting Assembly
    pdf.set_draw_color(160, 90, 20)
    pdf.set_line_width(0.5)
    # Tapping neck
    pdf.rect(p1_x + 56, diag_y + 24, 10, 16)
    # Flange or fitting
    pdf.set_fill_color(240, 240, 245)
    pdf.rect(p1_x + 51, diag_y + 20, 20, 4, "DF")
    pdf.rect(p1_x + 51, diag_y + 16, 20, 4, "DF")

    # Instrument Sensor Body Box
    pdf.set_fill_color(230, 242, 255)
    pdf.set_draw_color(30, 80, 140)
    pdf.rect(p1_x + 46, diag_y - 2, 30, 18, "DF")
    pdf.set_font("Helvetica", "B", 6.2)
    pdf.set_text_color(20, 45, 80)
    pdf.text(p1_x + 48, diag_y + 6, data["mfr"][:12])
    pdf.set_font("Helvetica", "", 5.5)
    pdf.text(p1_x + 48, diag_y + 12, data["category"][:14])

    # 2" Pipe Stand Stanchion
    pdf.set_draw_color(90, 100, 115)
    pdf.set_line_width(0.6)
    pdf.line(p1_x + 25, diag_y + 65, p1_x + 25, diag_y + 115)
    pdf.rect(p1_x + 18, diag_y + 115, 14, 5)
    pdf.set_font("Helvetica", "", 5.0)
    pdf.text(p1_x + 12, diag_y + 124, "2\" STANCHION PIPE STAND")

    # Mounting Clamp Link
    pdf.set_line_width(0.4)
    pdf.set_draw_color(120, 130, 145)
    pdf.line(p1_x + 25, diag_y + 75, p1_x + 50, diag_y + 15)

    # Process Hook-Up Parameters Table
    t_y = diag_y + 130
    pdf.set_fill_color(240, 245, 252)
    pdf.rect(p1_x + 4, t_y, 114, 6, "F")
    pdf.set_font("Helvetica", "B", 6.0)
    pdf.set_text_color(25, 45, 75)
    pdf.text(p1_x + 6, t_y + 4.2, "MECHANICAL HOOK-UP SPECIFICATION")

    params = [
        ("Process Connection:", data["proc_conn"]),
        ("Mounting Type:", "Direct Flanged / 2\" Pipe Stand"),
        ("Enclosure Rating:", data["enclosure"]),
        ("Hazardous Class:", data["haz_class"]),
        ("Process Medium:", "Starch Slurry / Saturated Steam / Air"),
        ("Ambient Temp Range:", "-20°C to +60°C Operating")
    ]
    py = t_y + 10
    for lbl, val in params:
        pdf.set_font("Helvetica", "B", 5.5)
        pdf.set_text_color(40, 50, 70)
        pdf.text(p1_x + 6, py, lbl)
        pdf.set_font("Helvetica", "", 5.5)
        pdf.set_text_color(20, 30, 50)
        val_str = val if len(val) <= 30 else val[:28] + ".."
        pdf.text(p1_x + 50, py, val_str)
        py += 7.2

    # -------------------------------------------------------------------------
    # PANE 2: ELECTRICAL & INSTRUMENT LOOP WIRING (Center: Width 142mm)
    # -------------------------------------------------------------------------
    p2_x = ox + 126.0
    p2_y = oy + 14
    p2_w = 142.0
    p2_h = 210.0

    pdf.set_fill_color(252, 253, 255)
    pdf.set_draw_color(200, 210, 225)
    pdf.rect(p2_x, p2_y, p2_w, p2_h, "DF")

    pdf.set_fill_color(30, 50, 80)
    pdf.rect(p2_x, p2_y, p2_w, 7, "F")
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.set_text_color(255, 255, 255)
    pdf.text(p2_x + 4, p2_y + 5, "SECTION B: ELECTRICAL LOOP WIRING SCHEMATIC (FIELD TO PLC)")

    # 4-Stage Loop Block:
    # 1. Field Instrument -> 2. Junction Box -> 3. Marshalling Cabinet CA1 -> 4. PLC 1756 Module
    stages = [
        ("FIELD INSTRUMENT", p2_x + 5, 28),
        ("JUNCTION BOX", p2_x + 38, 26),
        ("MARSHALLING CA1", p2_x + 69, 36),
        ("PLC 1756 I/O", p2_x + 110, 27)
    ]
    for s_title, sx, sw in stages:
        pdf.set_fill_color(235, 242, 250)
        pdf.set_draw_color(120, 140, 170)
        pdf.rect(sx, p2_y + 10, sw, 6, "DF")
        pdf.set_font("Helvetica", "B", 5.0)
        pdf.set_text_color(30, 50, 80)
        pdf.text(sx + 1.5, p2_y + 14.2, s_title)

    # 1. Field Instrument Box
    fx = p2_x + 5
    fy = p2_y + 20
    pdf.set_fill_color(245, 250, 255)
    pdf.set_draw_color(30, 70, 120)
    pdf.rect(fx, fy, 28, 80, "DF")
    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(20, 45, 80)
    pdf.text(fx + 2, fy + 7, "SENSOR TX")
    pdf.set_font("Helvetica", "", 4.8)
    pdf.text(fx + 2, fy + 12, data["tag_prefix"][:14])

    # Terminals on Instrument
    pins = [("1", "+ (24V)"), ("2", "- (Sig)"), ("S", "SHIELD"), ("E", "EARTH")]
    for p_idx, (p_no, p_lbl) in enumerate(pins):
        py = fy + 22 + p_idx * 14
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(60, 80, 110)
        pdf.rect(fx + 18, py, 7, 6, "DF")
        pdf.set_font("Helvetica", "B", 5.0)
        pdf.text(fx + 19.5, py + 4.5, p_no)
        pdf.set_font("Helvetica", "", 4.5)
        pdf.text(fx + 2, py + 4.5, p_lbl)

    # 2. Junction Box (JB-xxx)
    jbx = p2_x + 38
    jby = p2_y + 20
    pdf.set_fill_color(255, 252, 245)
    pdf.set_draw_color(180, 130, 40)
    pdf.rect(jbx, jby, 26, 80, "DF")
    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(80, 50, 10)
    pdf.text(jbx + 2, jby + 7, data["dest_jb"][:10])
    pdf.set_font("Helvetica", "", 4.8)
    pdf.text(jbx + 2, jby + 12, "FIELD JB")

    for p_idx in range(4):
        py = jby + 22 + p_idx * 14
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(160, 110, 30)
        pdf.rect(jbx + 5, py, 16, 6, "DF")
        pdf.set_font("Helvetica", "B", 5.0)
        pdf.text(jbx + 8, py + 4.5, f"TB:{p_idx+1}")

        # Field Cable Wire Lines (from Instrument to JB)
        pdf.set_draw_color(180, 40, 20) if p_idx == 0 else pdf.set_draw_color(30, 70, 160)
        pdf.set_line_width(0.4)
        pdf.line(fx + 25, fy + 25 + p_idx * 14, jbx + 5, py + 3)

    # Field Cable Label
    pdf.set_font("Helvetica", "B", 4.8)
    pdf.set_text_color(120, 30, 20)
    pdf.text(p2_x + 10, fy + 88, "FIELD CABLE:")
    pdf.set_font("Helvetica", "", 4.5)
    pdf.set_text_color(60, 70, 85)
    pdf.text(p2_x + 10, fy + 93, data["cable_spec"][:28])

    # 3. Marshalling Cabinet CA1 (Protection / Barrier)
    mx = p2_x + 69
    my = p2_y + 20
    pdf.set_fill_color(248, 250, 255)
    pdf.set_draw_color(40, 70, 120)
    pdf.rect(mx, my, 36, 80, "DF")
    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(20, 45, 80)
    pdf.text(mx + 2, my + 7, "CABINET CA1")
    pdf.set_font("Helvetica", "", 4.8)
    pdf.text(mx + 2, my + 12, "MARSHALLING")

    # Protection / Barrier Block inside CA1
    bar_is = "IS" in data["barrier"].upper()
    pdf.set_fill_color(220, 240, 255) if bar_is else pdf.set_fill_color(255, 248, 230)
    pdf.set_draw_color(0, 110, 200) if bar_is else pdf.set_draw_color(180, 120, 20)
    pdf.rect(mx + 4, my + 18, 28, 56, "DF")
    pdf.set_font("Helvetica", "B", 5.0)
    pdf.set_text_color(0, 80, 170) if bar_is else pdf.set_text_color(140, 80, 10)
    pdf.text(mx + 6, my + 24, "IS BARRIER" if bar_is else "ISOLATOR / RLY")
    pdf.set_font("Helvetica", "", 4.5)
    pdf.text(mx + 6, my + 29, data["barrier"][:16])

    # Trunk Cable lines (from JB to CA1)
    for p_idx in range(4):
        py = jby + 22 + p_idx * 14
        pdf.set_draw_color(30, 90, 150)
        pdf.line(jbx + 21, py + 3, mx + 4, my + 28 + p_idx * 10)

    # 4. PLC 1756 I/O Card
    px = p2_x + 110
    py = p2_y + 20
    pdf.set_fill_color(245, 248, 245)
    pdf.set_draw_color(30, 110, 50)
    pdf.rect(px, py, 27, 80, "DF")
    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(20, 75, 30)
    pdf.text(px + 2, py + 7, "PLC 1756")
    pdf.set_font("Helvetica", "", 4.8)
    pdf.text(px + 2, py + 12, data["plc_module"][:12])

    for p_idx in range(2):
        ch_y = py + 26 + p_idx * 18
        pdf.set_fill_color(255, 255, 255)
        pdf.set_draw_color(30, 110, 50)
        pdf.rect(px + 5, ch_y, 17, 8, "DF")
        pdf.set_font("Helvetica", "B", 5.0)
        pdf.text(px + 7, ch_y + 5.5, f"IN {p_idx+1} (+)")

        # Wire from Barrier to PLC
        pdf.set_draw_color(25, 100, 45)
        pdf.line(mx + 32, my + 30 + p_idx * 18, px + 5, ch_y + 4)

    # Grounding Earth Bar Detail at Bottom of Loop
    ey = p2_y + 110
    pdf.set_fill_color(245, 245, 245)
    pdf.set_draw_color(140, 140, 140)
    pdf.rect(p2_x + 5, ey, 132, 12, "DF")
    pdf.set_draw_color(40, 140, 60)
    pdf.set_line_width(1.0)
    pdf.line(p2_x + 8, ey + 6, p2_x + 134, ey + 6)
    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(25, 90, 40)
    pdf.text(p2_x + 10, ey + 4.5, "CLEAN INSTRUMENT GROUND BUS (IE) - CONNECTED TO CA1 MASTER GROUND")

    # Loop Characteristics Table
    lt_y = p2_y + 130
    pdf.set_fill_color(240, 245, 252)
    pdf.rect(p2_x + 4, lt_y, 134, 6, "F")
    pdf.set_font("Helvetica", "B", 6.0)
    pdf.set_text_color(25, 45, 75)
    pdf.text(p2_x + 6, lt_y + 4.2, "ELECTRICAL & SIGNAL LOOP SPECIFICATIONS")

    elec_specs = [
        ("Signal Protocol:", data["signal_type"]),
        ("Power Supply:", data["power_supply"]),
        ("Field Cable Type:", data["cable_spec"]),
        ("Destination JB:", data["dest_jb"]),
        ("Marshalling Cabinet:", "CA1 / Master Control Panel (Bay M1..M4)"),
        ("Barrier / Isolation:", data["barrier"]),
        ("PLC Module:", data["plc_module"])
    ]
    ey_row = lt_y + 10
    for lbl, val in elec_specs:
        pdf.set_font("Helvetica", "B", 5.5)
        pdf.set_text_color(40, 50, 70)
        pdf.text(p2_x + 6, ey_row, lbl)
        pdf.set_font("Helvetica", "", 5.5)
        pdf.set_text_color(20, 30, 50)
        v_clean = val if len(val) <= 40 else val[:38] + ".."
        pdf.text(p2_x + 45, ey_row, v_clean)
        ey_row += 7.2

    # -------------------------------------------------------------------------
    # PANE 3: BILL OF MATERIALS & ENGINEERING NOTES (Right: Width 122mm)
    # -------------------------------------------------------------------------
    p3_x = ox + 272.0
    p3_y = oy + 14
    p3_w = 122.0
    p3_h = 210.0

    pdf.set_fill_color(252, 253, 255)
    pdf.set_draw_color(200, 210, 225)
    pdf.rect(p3_x, p3_y, p3_w, p3_h, "DF")

    pdf.set_fill_color(30, 50, 80)
    pdf.rect(p3_x, p3_y, p3_w, 7, "F")
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.set_text_color(255, 255, 255)
    pdf.text(p3_x + 4, p3_y + 5, "SECTION C: BILL OF MATERIALS (BOM) & NOTES")

    # Equipment Header Tag
    pdf.set_fill_color(240, 245, 252)
    pdf.rect(p3_x + 4, p3_y + 10, 114, 18, "DF")
    pdf.set_font("Helvetica", "B", 6.5)
    pdf.set_text_color(20, 40, 70)
    pdf.text(p3_x + 6, p3_y + 15, f"MANUFACTURER: {data['mfr'].upper()}")
    pdf.set_font("Helvetica", "", 6.0)
    pdf.text(p3_x + 6, p3_y + 20, f"MODEL: {data['model']}")
    pdf.text(p3_x + 6, p3_y + 25, f"SERVICE: {data['service']}")

    # BOM Table Header
    bom_y = p3_y + 32
    pdf.set_fill_color(28, 48, 80)
    pdf.rect(p3_x + 4, bom_y, 114, 6, "F")
    pdf.set_font("Helvetica", "B", 5.5)
    pdf.set_text_color(255, 255, 255)
    pdf.text(p3_x + 6, bom_y + 4.2, "ITEM")
    pdf.text(p3_x + 16, bom_y + 4.2, "MATERIAL / COMPONENT DESCRIPTION")
    pdf.text(p3_x + 85, bom_y + 4.2, "SPEC / MODEL")
    pdf.text(p3_x + 108, bom_y + 4.2, "QTY")

    by = bom_y + 6
    for item_no, comp_desc, spec_model, qty in data["bom"]:
        pdf.set_fill_color(255, 255, 255)
        pdf.rect(p3_x + 4, by, 114, 11.5, "F")
        pdf.set_font("Helvetica", "B", 5.5)
        pdf.set_text_color(30, 45, 65)
        pdf.text(p3_x + 6, by + 5.0, item_no)
        pdf.set_font("Helvetica", "", 5.2)
        pdf.text(p3_x + 16, by + 5.0, comp_desc[:30])
        pdf.text(p3_x + 16, by + 9.0, spec_model[:32])
        pdf.set_font("Helvetica", "B", 5.2)
        pdf.text(p3_x + 108, by + 5.0, qty)
        pdf.set_draw_color(225, 230, 240)
        pdf.line(p3_x + 4, by + 11.5, p3_x + 118, by + 11.5)
        by += 11.5

    # Engineering Installation Notes
    notes_y = by + 6
    pdf.set_fill_color(240, 245, 252)
    pdf.rect(p3_x + 4, notes_y, 114, 6, "F")
    pdf.set_font("Helvetica", "B", 6.0)
    pdf.set_text_color(25, 45, 75)
    pdf.text(p3_x + 6, notes_y + 4.2, "ENGINEERING INSTALLATION NOTES")

    ny = notes_y + 10
    for idx, note in enumerate(data["notes"]):
        pdf.set_font("Helvetica", "B", 5.5)
        pdf.set_text_color(20, 35, 60)
        pdf.text(p3_x + 6, ny, f"{idx+1}.")
        pdf.set_font("Helvetica", "", 5.2)
        pdf.set_text_color(60, 70, 85)
        # Word wrap note text
        note_lines = pdf.multi_cell(102, 3.2, note, dry_run=True, output="LINES")
        pdf.text(p3_x + 12, ny, note_lines[0])
        if len(note_lines) > 1:
            pdf.text(p3_x + 12, ny + 3.4, note_lines[1])
            ny += 3.4
        if len(note_lines) > 2:
            pdf.text(p3_x + 12, ny + 3.4, note_lines[2])
            ny += 3.4
        ny += 7.5


# -----------------------------------------------------------------------------
# STEP 4: AUTOCAD DXF 1:1 METRIC MM EXPORTER FOR INSTRUMENT TEMPLATES
# -----------------------------------------------------------------------------
class InstrumentDXFExporter:
    def __init__(self):
        self.doc = ezdxf.new("R2010", setup=True)
        self.doc.header["$INSUNITS"] = 4  # Millimeters
        self.msp = self.doc.modelspace()
        self._setup_layers()

    def _setup_layers(self):
        layers = [
            ("0_BORDER", 7, 70),             # White, 0.70mm
            ("0_TITLE_BLOCK", 4, 35),        # Cyan, 0.35mm
            ("IE-EQUIP", 2, 35),             # Yellow, 0.35mm
            ("IE-WIRE", 1, 25),              # Red, 0.25mm
            ("IE-CABLE", 6, 25),             # Magenta, 0.25mm
            ("TER", 3, 35),                  # Green, 0.35mm
            ("IS_BARRIERS", 140, 35),        # Sky Blue, 0.35mm
            ("RELAYS", 30, 35),              # Orange, 0.35mm
            ("NOTATIONS", 8, 18),            # Gray, 0.18mm
            ("TEXTS", 7, 25),                # White, 0.25mm
            ("DIMENSIONS", 1, 18)            # Red, 0.18mm
        ]
        for name, color, lw in layers:
            if name not in self.doc.layers:
                self.doc.layers.add(name, color=color, lineweight=lw)

    def add_rect(self, x, y, w, h, layer="0_BORDER"):
        pts = [(x, y), (x + w, y), (x + w, y + h), (x, y + h)]
        return self.msp.add_lwpolyline(pts, close=True, dxfattribs={"layer": layer})

    def add_text(self, text, x, y, height=2.5, layer="TEXTS", align=TextEntityAlignment.LEFT):
        txt = self.msp.add_text(text, dxfattribs={"height": height, "layer": layer})
        txt.set_placement((x, y), align=align)
        return txt

    def draw_frame(self, ox, oy, sheet_no, total_sheets, title, dwg_no):
        self.add_rect(ox, oy, 400.0, 277.0, layer="0_BORDER")
        self.add_rect(ox + 2.0, oy + 2.0, 396.0, 273.0, layer="NOTATIONS")

        # Title block
        tb_w = 175.0
        tb_h = 36.0
        x = ox + 400.0 - tb_w
        y = oy

        self.add_rect(x, y, tb_w, tb_h, layer="0_BORDER")
        self.msp.add_line((x, y + 11), (x + tb_w, y + 11), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x, y + 23), (x + tb_w, y + 23), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 110, y), (x + 110, y + 23), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 55, y + 23), (x + 55, y + tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 110, y + 23), (x + 110, y + tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})
        self.msp.add_line((x + 145, y + 23), (x + 145, y + tb_h), dxfattribs={"layer": "0_TITLE_BLOCK"})

        self.add_text("INGREDION (THAILAND) CO., LTD.", x + 3, y + 4.5, height=2.8, layer="0_TITLE_BLOCK")
        self.add_text("KALASIN STARCH PLANT - JET COOKER PROJECT", x + 3, y + 8.5, height=2.0, layer="TEXTS")
        self.add_text("AEC INDUSTRIAL ENGINEERING", x + 113, y + 4.5, height=2.8, layer="0_TITLE_BLOCK")
        self.add_text("SYSTEMS INTEGRATION & AUTOMATION", x + 113, y + 8.5, height=2.0, layer="TEXTS")

        self.add_text(title[:45], x + 3, y + 16, height=2.8, layer="0_TITLE_BLOCK")
        self.add_text("INSTRUMENT HOOK-UP & TYPICAL LOOP DIAGRAM", x + 3, y + 20.5, height=2.0, layer="TEXTS")

        self.add_text(f"DWG NO: {dwg_no}", x + 113, y + 15.5, height=2.5, layer="TEXTS")
        self.add_text("REV: 3.6    DATE: 2026-09-08", x + 113, y + 20, height=2.2, layer="TEXTS")

        self.add_text("DRAWN: AEC-ENG", x + 3, y + 27, height=2.0, layer="TEXTS")
        self.add_text("CHECKED: LEAD-PE", x + 3, y + 32, height=2.0, layer="TEXTS")
        self.add_text("STATUS: APPROVED", x + 58, y + 27, height=2.2, layer="0_TITLE_BLOCK")
        self.add_text("SCALE: N.T.S.", x + 58, y + 32, height=2.0, layer="TEXTS")

        self.add_text(f"SHEET: {sheet_no:02d} / {total_sheets:02d}", x + 148, y + 32.5, height=2.8, layer="0_TITLE_BLOCK")

    def draw_template_panes(self, ox, oy, data):
        # Top banner
        self.add_rect(ox + 5, oy + 245, 390, 15, layer="IE-EQUIP")
        self.add_text(f"TEMPLATE {data['code']}: {data['category'].upper()} | {data['service']}", ox + 10, oy + 252, height=3.2, layer="0_TITLE_BLOCK")

        # Pane 1: Process Hook-Up (Left: 125mm)
        p1_x = ox + 5
        p1_y = oy + 38
        self.add_rect(p1_x, p1_y, 125, 200, layer="NOTATIONS")
        self.add_rect(p1_x, p1_y + 192, 125, 8, layer="0_TITLE_BLOCK")
        self.add_text("PROCESS PIPING HOOK-UP", p1_x + 4, p1_y + 194.5, height=2.2, layer="0_TITLE_BLOCK")

        # Piping graphics
        self.msp.add_line((p1_x + 10, p1_y + 150), (p1_x + 115, p1_y + 150), dxfattribs={"layer": "IE-EQUIP"})
        self.msp.add_line((p1_x + 10, p1_y + 138), (p1_x + 115, p1_y + 138), dxfattribs={"layer": "IE-EQUIP"})
        self.add_text("PROCESS LINE", p1_x + 12, p1_y + 153, height=2.0, layer="TEXTS")

        # Instrument sensor
        self.add_rect(p1_x + 48, p1_y + 160, 28, 20, layer="IE-EQUIP")
        self.add_text(data["mfr"][:12], p1_x + 50, p1_y + 172, height=2.0, layer="TEXTS")
        self.add_text(data["model"][:14], p1_x + 50, p1_y + 165, height=1.8, layer="NOTATIONS")

        # Stand
        self.msp.add_line((p1_x + 25, p1_y + 40), (p1_x + 25, p1_y + 130), dxfattribs={"layer": "NOTATIONS"})
        self.add_text("2\" STANCHION", p1_x + 15, p1_y + 35, height=1.8, layer="NOTATIONS")

        # Specs
        self.add_text(f"CONN: {data['proc_conn'][:25]}", p1_x + 6, p1_y + 80, height=1.8, layer="TEXTS")
        self.add_text(f"ENCL: {data['enclosure'][:25]}", p1_x + 6, p1_y + 70, height=1.8, layer="TEXTS")
        self.add_text(f"HAZ: {data['haz_class'][:25]}", p1_x + 6, p1_y + 60, height=1.8, layer="TEXTS")

        # Pane 2: Electrical Loop (Center: 140mm)
        p2_x = ox + 133
        p2_y = oy + 38
        self.add_rect(p2_x, p2_y, 138, 200, layer="NOTATIONS")
        self.add_rect(p2_x, p2_y + 192, 138, 8, layer="0_TITLE_BLOCK")
        self.add_text("ELECTRICAL LOOP SCHEMATIC", p2_x + 4, p2_y + 194.5, height=2.2, layer="0_TITLE_BLOCK")

        # 4 loop blocks: Sensor, JB, CA1, PLC
        self.add_rect(p2_x + 4, p2_y + 90, 26, 95, layer="IE-EQUIP")
        self.add_text("SENSOR TX", p2_x + 6, p2_y + 175, height=2.0, layer="TEXTS")

        self.add_rect(p2_x + 36, p2_y + 90, 26, 95, layer="TER")
        self.add_text("FIELD JB", p2_x + 38, p2_y + 175, height=2.0, layer="TER")

        self.add_rect(p2_x + 68, p2_y + 90, 34, 95, layer="IS_BARRIERS" if "IS" in data["barrier"] else "RELAYS")
        self.add_text("CA1 BARRIER", p2_x + 70, p2_y + 175, height=2.0, layer="TEXTS")

        self.add_rect(p2_x + 108, p2_y + 90, 26, 95, layer="IE-EQUIP")
        self.add_text("PLC 1756", p2_x + 110, p2_y + 175, height=2.0, layer="TEXTS")

        # Wires
        for wy in [155, 140, 125, 110]:
            self.msp.add_line((p2_x + 30, p2_y + wy), (p2_x + 36, p2_y + wy), dxfattribs={"layer": "IE-WIRE"})
            self.msp.add_line((p2_x + 62, p2_y + wy), (p2_x + 68, p2_y + wy), dxfattribs={"layer": "IE-CABLE"})
            self.msp.add_line((p2_x + 102, p2_y + wy), (p2_x + 108, p2_y + wy), dxfattribs={"layer": "IE-WIRE"})

        self.add_text(f"CABLE: {data['cable_spec'][:32]}", p2_x + 6, p2_y + 70, height=2.0, layer="IE-CABLE")
        self.add_text(f"POWER: {data['power_supply'][:32]}", p2_x + 6, p2_y + 60, height=2.0, layer="TEXTS")
        self.add_text(f"SIGNAL: {data['signal_type'][:32]}", p2_x + 6, p2_y + 50, height=2.0, layer="TEXTS")

        # Pane 3: BOM & Notes (Right: 118mm)
        p3_x = ox + 275
        p3_y = oy + 38
        self.add_rect(p3_x, p3_y, 120, 200, layer="NOTATIONS")
        self.add_rect(p3_x, p3_y + 192, 120, 8, layer="0_TITLE_BLOCK")
        self.add_text("BOM & ENGINEERING NOTES", p3_x + 4, p3_y + 194.5, height=2.2, layer="0_TITLE_BLOCK")

        # BOM Items
        by = p3_y + 180
        self.add_text("ITEM  DESCRIPTION                    QTY", p3_x + 4, by, height=1.8, layer="0_TITLE_BLOCK")
        by -= 8
        for b_no, b_desc, b_spec, b_qty in data["bom"]:
            self.add_text(f"{b_no}     {b_desc[:24]}  {b_qty}", p3_x + 4, by, height=1.8, layer="TEXTS")
            by -= 8

        # Notes
        ny = p3_y + 90
        self.add_text("INSTALLATION NOTES:", p3_x + 4, ny, height=2.0, layer="0_TITLE_BLOCK")
        ny -= 8
        for idx, n in enumerate(data["notes"]):
            self.add_text(f"{idx+1}. {n[:45]}", p3_x + 4, ny, height=1.6, layer="NOTATIONS")
            ny -= 6


# -----------------------------------------------------------------------------
# STEP 5: MASTER EXECUTION ORCHESTRATION
# -----------------------------------------------------------------------------
def main():
    print("==================================================================")
    print("  KALASIN JET COOKER - INSTRUMENT DRAWING TEMPLATE GENERATOR")
    print("==================================================================")

    total_sheets = 2 + len(TEMPLATES_DATA)
    print(f"Total Drawings in Template Suite: {total_sheets} Sheets")

    # -------------------------------------------------------------------------
    # PART A: GENERATE MULTI-PAGE PDF DRAWING PACKAGE
    # -------------------------------------------------------------------------
    print("\n--- Part A: Building High-Resolution Vector PDF Drawing Package ---")
    pdf = InstrumentTemplatePDF()
    pdf.total_sheets = total_sheets

    print("  [Rendering Sheet 01] Template Cover Sheet...")
    render_template_cover_sheet(pdf)

    print("  [Rendering Sheet 02] Template Directory & Index Sheet...")
    render_template_index_sheet(pdf)

    for idx, data in enumerate(TEMPLATES_DATA):
        s_no = 3 + idx
        print(f"  [Rendering Sheet {s_no:02d}] {data['code']} - {data['category']}...")
        render_instrument_template_sheet(pdf, s_no, data)

    pdf_output_path = os.path.join(PDF_DIR, "Instrument_Installation_Hookup_and_Wiring_Templates.pdf")
    pdf.output(pdf_output_path)
    print(f"\n[SUCCESS] PDF Generated: {pdf_output_path} ({os.path.getsize(pdf_output_path):,} bytes)")

    # -------------------------------------------------------------------------
    # PART B: EXPORT SVG SHEETS & PNG PREVIEWS
    # -------------------------------------------------------------------------
    print("\n--- Part B: Rendering Scalable Vector Graphics (SVG) & High-Res PNGs ---")
    doc = fitz.open(pdf_output_path)
    preview_sheets = [1, 2, 3, 5, 8, 16, 22, 23]  # Cover, Index, Dust, DP Tx, Coriolis, Press Tx, Control Valve, On-Off Valve

    for idx, page in enumerate(doc):
        s_no = idx + 1
        svg_name = f"Instrument_Template_Sheet_{s_no:02d}.svg"
        svg_content = page.get_svg_image()
        with open(os.path.join(CAD_DIR, svg_name), "w", encoding="utf-8") as f:
            f.write(svg_content)
        with open(os.path.join(PDF_DIR, svg_name), "w", encoding="utf-8") as f:
            f.write(svg_content)

        if s_no in preview_sheets:
            pix = page.get_pixmap(dpi=150)
            png_name = f"instrument_template_sheet_{s_no:02d}.png"
            pix.save(os.path.join(PDF_DIR, png_name))
            print(f"  [Preview Rendered] Sheet {s_no:02d}: {png_name}")

    print(f"  [OK] Exported {len(doc)} vector SVG sheets.")

    # -------------------------------------------------------------------------
    # PART C: AUTOCAD DXF SUITE (1:1 MODEL SPACE & INDIVIDUAL SHEETS)
    # -------------------------------------------------------------------------
    print("\n--- Part C: Exporting AutoCAD DXF Suites (1:1 Metric mm) ---")
    master_builder = InstrumentDXFExporter()

    grid_cols = 5
    spacing_x = 450.0
    spacing_y = 320.0

    # Draw Cover & Index in master DXF
    master_builder.draw_frame(0, 0, 1, total_sheets, "INSTRUMENT HOOK-UP & LOOP DRAWING DIRECTORY", "KAL-JC-TYP-000")
    master_builder.draw_frame(spacing_x, 0, 2, total_sheets, "INSTRUMENT TEMPLATE DRAWING DIRECTORY", "KAL-JC-TYP-000-IDX")

    # Draw 23 templates in master grid
    for idx, data in enumerate(TEMPLATES_DATA):
        grid_idx = 2 + idx
        gx = (grid_idx % grid_cols) * spacing_x
        gy = -(grid_idx // grid_cols) * spacing_y
        s_no = 3 + idx
        master_builder.draw_frame(gx, gy, s_no, total_sheets, data["title"], data["dwg_no"])
        master_builder.draw_template_panes(gx, gy, data)

    master_dxf_path = os.path.join(CAD_DIR, "Instrument_Hookup_and_Wiring_Templates_Master_1to1.dxf")
    master_builder.doc.saveas(master_dxf_path)
    print(f"  [OK] Master AutoCAD DXF Saved: {master_dxf_path} ({os.path.getsize(master_dxf_path):,} bytes)")

    # Individual standalone DXF files
    print("  [Exporting 25 Individual Standalone DXF Files]...")
    # Cover
    dxf_cov = InstrumentDXFExporter()
    dxf_cov.draw_frame(0, 0, 1, total_sheets, "INSTRUMENT HOOK-UP & LOOP DRAWING DIRECTORY", "KAL-JC-TYP-000")
    dxf_cov.doc.saveas(os.path.join(CAD_DIR, "KAL-JC-TYP-000_Sheet01_Cover.dxf"))

    # Index
    dxf_idx = InstrumentDXFExporter()
    dxf_idx.draw_frame(0, 0, 2, total_sheets, "INSTRUMENT TEMPLATE DRAWING DIRECTORY", "KAL-JC-TYP-000-IDX")
    dxf_idx.doc.saveas(os.path.join(CAD_DIR, "KAL-JC-TYP-000-IDX_Sheet02_Index.dxf"))

    # 23 Templates
    for idx, data in enumerate(TEMPLATES_DATA):
        s_no = 3 + idx
        fname = f"{data['dwg_no']}_Sheet{s_no:02d}_{data['code']}_{data['category'].replace(' ', '_').replace('/', '_')}.dxf"
        single_dxf = InstrumentDXFExporter()
        single_dxf.draw_frame(0, 0, s_no, total_sheets, data["title"], data["dwg_no"])
        single_dxf.draw_template_panes(0, 0, data)
        single_dxf.doc.saveas(os.path.join(CAD_DIR, fname))

    print(f"  [OK] All {total_sheets} individual DXF files successfully exported to {CAD_DIR}")
    print("\n==================================================================")
    print("  ALL INSTRUMENT TEMPLATES GENERATED AND VERIFIED!")
    print("==================================================================")


if __name__ == "__main__":
    main()
