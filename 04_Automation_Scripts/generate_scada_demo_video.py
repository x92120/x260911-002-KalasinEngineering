#!/usr/bin/env python3
"""
Generate High-Definition SCADA Demo Video for Ingredion Kalasin Plant - Project SPRINT
Conforming to:
- P&ID Rev 3.5 Process & Instrumentation Baseline
- Rockwell Automation PlantPAx 5.x DCS HMI Standard
- ANSI/ISA-101 High Performance HMI Color & Symbol Guidelines
- ANSI/ISA-18.2 Alarm Management Standard
"""

import os
import math
import time
import shutil
import numpy as np
from PIL import Image, ImageDraw, ImageFont
import imageio.v2 as imageio

# Resolution & Video Config
WIDTH = 1920
HEIGHT = 1080
FPS = 30
TOTAL_SECONDS = 60
TOTAL_FRAMES = FPS * TOTAL_SECONDS

OUTPUT_DIR_1 = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/SCADA_Demo_Video"
OUTPUT_DIR_2 = "/Users/x92120/Library/CloudStorage/GoogleDrive-x92120@gmail.com/My Drive/0x01-proj/xPrj-2603001-Kalasin/IO_List/SCADA_Demo_Video"
OUTPUT_FILE_1 = os.path.join(OUTPUT_DIR_1, "Kalasin_Plant_SCADA_Demo_Project_SPRINT_Rev3.5.mp4")
OUTPUT_FILE_2 = os.path.join(OUTPUT_DIR_2, "Kalasin_Plant_SCADA_Demo_Project_SPRINT_Rev3.5.mp4")

os.makedirs(OUTPUT_DIR_1, exist_ok=True)
os.makedirs(OUTPUT_DIR_2, exist_ok=True)

# Load TrueType fonts
def get_font(size, bold=False):
    font_paths = [
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "/System/Library/Fonts/SFNSMono.ttf"
    ]
    for p in font_paths:
        if os.path.exists(p):
            try:
                return ImageFont.truetype(p, size)
            except Exception:
                pass
    return ImageFont.load_default()

FONT_TITLE = get_font(22, bold=True)
FONT_SUBTITLE = get_font(15, bold=True)
FONT_TAB = get_font(14, bold=True)
FONT_BODY = get_font(13, bold=False)
FONT_BODY_BOLD = get_font(13, bold=True)
FONT_TAG = get_font(12, bold=True)
FONT_VALUE = get_font(14, bold=True)
FONT_STATUS = get_font(11, bold=True)
FONT_ALARM = get_font(14, bold=True)
FONT_SMALL = get_font(10, bold=False)

# Colors (PlantPAx / ISA-101 Standard)
C_TOP_BAR = (24, 32, 44)
C_TOP_BAR_BORDER = (45, 58, 76)
C_BANNER_BG = (35, 46, 60)
C_NAV_BG = (220, 224, 230)
C_NAV_ACTIVE = (13, 110, 253)
C_CANVAS_BG = (216, 221, 228)
C_FOOTER_BG = (20, 26, 36)

C_WHITE = (255, 255, 255)
C_BLACK = (20, 20, 20)
C_GRAY_TEXT = (100, 110, 125)
C_DARK_TEXT = (30, 35, 45)

C_RUN_GREEN = (0, 190, 80)
C_STOP_GRAY = (140, 145, 155)
C_ALARM_RED = (220, 38, 38)
C_WARN_AMBER = (235, 145, 10)
C_PIPE_GRAY = (150, 158, 170)
C_WATER_BLUE = (30, 130, 230)
C_SLURRY_ORANGE = (215, 120, 40)
C_STEAM_CYAN = (0, 180, 216)
C_PANEL_BG = (242, 245, 249)
C_PANEL_BORDER = (185, 195, 208)

# Navigation Tabs
TABS = [
    "1. Overview",
    "2. Slurry & pH (202/402)",
    "3. Jet Cooker & Infeed",
    "4. Spray Dryer (602)",
    "5. Burner & BMS",
    "6. Starch Silo (613)",
    "7. Steam Boiler (922)",
    "8. Process Trends",
    "9. Alarms"
]

def draw_top_banner(d, frame_idx, current_tab_idx, alarm_state="NORMAL", alarm_msg=""):
    # 1. Dark Topmost Header (y: 0 - 46)
    d.rectangle([0, 0, WIDTH, 46], fill=C_TOP_BAR)
    d.line([0, 46, WIDTH, 46], fill=C_TOP_BAR_BORDER, width=2)
    
    # Logo / Plant Name
    d.rectangle([15, 8, 45, 38], fill=(0, 120, 215), outline=C_WHITE, width=1)
    d.text((23, 14), "IGD", font=FONT_BODY_BOLD, fill=C_WHITE)
    d.text((58, 12), "INGREDION (THAILAND) CO., LTD.  |  KALASIN PLANT — PROJECT SPRINT", font=FONT_TITLE, fill=C_WHITE)
    
    # System Status Indicator
    blink = (frame_idx // 15) % 2 == 0
    sys_color = C_RUN_GREEN if blink else (0, 230, 100)
    d.ellipse([980, 18, 992, 30], fill=sys_color)
    d.text((1000, 15), "SYSTEM: ONLINE [AUTO]", font=FONT_SUBTITLE, fill=C_RUN_GREEN)
    
    # Controller Redundancy
    d.text((1200, 16), "CPU: 1756-L83ES (PRIMARY - SYNC)", font=FONT_BODY, fill=(180, 195, 210))
    
    # User / Clock
    d.text((1520, 16), "USER: ENG_CHONLAWIT [ADMIN]", font=FONT_BODY_BOLD, fill=(220, 230, 242))
    
    # Dynamic Clock
    sim_sec = 30 + (frame_idx // FPS)
    sim_min = 48 + (sim_sec // 60)
    sim_sec = sim_sec % 60
    sim_hour = 14 + (sim_min // 60)
    sim_min = sim_min % 60
    clock_str = f"2026-09-08 {sim_hour:02d}:{sim_min:02d}:{sim_sec:02d}"
    d.text((1770, 16), clock_str, font=FONT_BODY_BOLD, fill=C_WHITE)
    
    # 2. Alarm Banner & Quick Toolbar (y: 46 - 92)
    banner_bg = C_BANNER_BG
    if alarm_state == "ACTIVE":
        banner_bg = C_ALARM_RED if ((frame_idx // 10) % 2 == 0) else (180, 20, 20)
    elif alarm_state == "ACKNOWLEDGED":
        banner_bg = (160, 100, 0)
    
    d.rectangle([0, 46, WIDTH, 92], fill=banner_bg)
    d.line([0, 92, WIDTH, 92], fill=(60, 75, 95), width=1)
    
    # Alarm Counters
    d.rectangle([15, 52, 105, 86], fill=(45, 55, 70), outline=(70, 85, 105), width=1)
    d.text((22, 56), "CRIT: 0", font=FONT_STATUS, fill=(180, 190, 200))
    d.text((22, 70), "WARN: 1" if alarm_state != "NORMAL" else "WARN: 0", font=FONT_STATUS, fill=C_WARN_AMBER if alarm_state != "NORMAL" else (180, 190, 200))
    
    # Center Alarm Message
    if alarm_state == "NORMAL":
        d.text((130, 60), "ALL SAFETY INTERLOCKS HEALTHY  |  JET COOKER & SPRAY DRYER IN SERVICE  |  ZERO ACTIVE TRIPS", font=FONT_ALARM, fill=(130, 220, 160))
    elif alarm_state == "ACTIVE":
        d.text((130, 60), f"ALARM ACTIVE: {alarm_msg}", font=FONT_ALARM, fill=C_WHITE)
    elif alarm_state == "ACKNOWLEDGED":
        d.text((130, 60), f"ACKNOWLEDGED: {alarm_msg} (SEQUENCE STABILIZING)", font=FONT_ALARM, fill=C_WHITE)
        
    # Quick buttons
    btn_w = 95
    bx = 1480
    for btn_name in ["SILENCE", "ACK ALL", "LOGS", "HELP"]:
        d.rectangle([bx, 53, bx + btn_w, 85], fill=(55, 68, 86), outline=(90, 110, 135), width=1)
        d.text((bx + 18, 62), btn_name, font=FONT_BODY_BOLD, fill=C_WHITE)
        bx += btn_w + 10

    # 3. Navigation Ribbon (y: 92 - 138)
    d.rectangle([0, 92, WIDTH, 138], fill=C_NAV_BG)
    d.line([0, 138, WIDTH, 138], fill=(180, 190, 205), width=2)
    
    tab_x = 15
    for idx, tab_label in enumerate(TABS):
        is_active = (idx == current_tab_idx)
        tw = 190
        if is_active:
            d.rectangle([tab_x, 96, tab_x + tw, 137], fill=C_NAV_ACTIVE)
            d.text((tab_x + 15, 108), tab_label, font=FONT_TAB, fill=C_WHITE)
            d.line([tab_x, 136, tab_x + tw, 136], fill=(255, 215, 0), width=3)
        else:
            d.rectangle([tab_x, 98, tab_x + tw, 135], fill=(235, 239, 245), outline=(195, 205, 218), width=1)
            d.text((tab_x + 15, 108), tab_label, font=FONT_TAB, fill=C_DARK_TEXT)
        tab_x += tw + 8

def draw_bottom_footer(d):
    # Bottom Status Bar (y: 1032 - 1080)
    d.rectangle([0, 1032, WIDTH, 1080], fill=C_FOOTER_BG)
    d.line([0, 1032, WIDTH, 1032], fill=(50, 65, 85), width=2)
    
    d.text((20, 1048), "PLC-01: 1756-L83ES | SCAN: 16ms | MEMORY: 32% | DLR RING 1Gbps: HEALTHY (RIO-200, RIO-400, RIO-600, RIO-MCC)", font=FONT_BODY, fill=(180, 200, 220))
    d.text((1200, 1048), "BATTERY LIMITS: STEAM 8.4 barg | AIR 6.8 barg | LPG 2.2 barg | WTP 3.5 barg | DRYER FEED: 44.0 m3/h", font=FONT_BODY_BOLD, fill=(200, 225, 255))

def draw_mouse_cursor(d, x, y, click=False):
    pts = [(x, y), (x, y + 24), (x + 7, y + 18), (x + 14, y + 27), (x + 18, y + 25), (x + 11, y + 15), (x + 20, y + 15)]
    d.polygon(pts, fill=(255, 255, 255), outline=(0, 0, 0))
    if click:
        d.ellipse([x - 12, y - 12, x + 12, y + 12], outline=(255, 230, 0), width=3)

# ----------------- SCENE RENDERERS -----------------

def render_scene_1_overview(d, frame_idx):
    units = [
        ("AREA 202: SLURRY PREP", 60, 180, 360, 520, "Re-Slurry Tanks TS-20201/02\nAgitators AG-20201/02\nTransfer Pump PC-20201\nStatus: RUNNING (44.5 m3/h)", C_WATER_BLUE),
        ("AREA 402: pH & JET COOKER", 460, 180, 760, 520, "pH Adjust TS-40201/02\npH Sensor AT-40201: 6.25 pH\nBooster Pump PC-40203\nJet Cooker: 105.4 °C / 4.3 barg", C_SLURRY_ORANGE),
        ("AREA 602: SPRAY DRYER TOWER", 860, 180, 1260, 680, "Triplex HP Pump PD-60202 (285 barg)\nDrying Chamber CHAM-21\nAir Heater FH-60210A (195.4 °C)\nBaghouse DTEX-21 (DP: 125 mmH2O)", (40, 160, 120)),
        ("AREA 613: SILO & PACKING", 1360, 180, 1660, 520, "Starch Silo TS-61301 (65%)\nVibratory Sifters SC-61301\nBig-Bag Packer ME-61306\nRate: 2,250 kg/h", (130, 90, 180)),
        ("UTILITIES & INFRASTRUCTURE", 60, 580, 760, 980, "Area 922: SPRINT Steam Boiler (8.4 barg, 6.2 t/h)\nArea 913: LPG Bullet Tank & Vaporizer EV-91301\nArea 930: WTP Clarifier & Sand Filters TF-93322\nArea 950: Rotary Air Compressors CO-95200 (6.8 barg)", (70, 100, 130)),
        ("PRODUCTION METRICS & KPI DASHBOARD", 860, 740, 1660, 980, "Overall Plant Efficiency: 94.2%  |  Steam Economy: 1.82 kg steam/kg powder\nDaily Production Total: 38.4 Tons (Target: 50.0 Tons)\nTotal Power Demand: 842 kW  |  Zero Safety Interlock Bypass", (50, 70, 90))
    ]
    
    for title, x1, y1, x2, y2, desc, col in units:
        d.rectangle([x1, y1, x2, y2], fill=C_PANEL_BG, outline=C_PANEL_BORDER, width=2)
        d.rectangle([x1, y1, x2, y1 + 38], fill=col)
        d.text((x1 + 15, y1 + 10), title, font=FONT_SUBTITLE, fill=C_WHITE)
        
        lines = desc.split("\n")
        ly = y1 + 55
        for l in lines:
            d.text((x1 + 20, ly), l, font=FONT_BODY_BOLD if "Status" in l or "Daily" in l else FONT_BODY, fill=C_DARK_TEXT)
            ly += 26
            
    pipe_y = 350
    d.line([360, pipe_y, 460, pipe_y], fill=C_PIPE_GRAY, width=8)
    d.line([760, pipe_y, 860, pipe_y], fill=C_PIPE_GRAY, width=8)
    d.line([1260, pipe_y, 1360, pipe_y], fill=C_PIPE_GRAY, width=8)
    
    pulse_offset = (frame_idx * 6) % 100
    for px_start in [360, 760, 1260]:
        px = px_start + pulse_offset
        d.ellipse([px - 5, pipe_y - 5, px + 5, pipe_y + 5], fill=C_SLURRY_ORANGE)
        
    d.rectangle([880, 890, 1640, 920], fill=(220, 225, 235), outline=C_PANEL_BORDER, width=1)
    bar_w = int(760 * 0.768)
    d.rectangle([880, 890, 880 + bar_w, 920], fill=C_RUN_GREEN)
    d.text((1200, 897), "76.8% COMPLETE (38.4 / 50.0 Tons)", font=FONT_STATUS, fill=C_WHITE)
    
    progress = min(1.0, frame_idx / 220.0)
    mx = int(960 + (250 - 960) * progress)
    my = int(540 + (115 - 540) * progress)
    click = (frame_idx > 210)
    draw_mouse_cursor(d, mx, my, click)

def render_scene_2_slurry(d, frame_idx):
    d.text((40, 150), "AREA 202 & 402: RE-SLURRY, pH ADJUSTMENT & CONTINUOUS JET COOKER", font=FONT_TITLE, fill=C_DARK_TEXT)
    
    # Tank 1: Re-Slurry TS-20201
    d.rectangle([80, 240, 280, 560], fill=(230, 235, 242), outline=C_DARK_TEXT, width=3)
    d.text((110, 205), "TANK TS-20201\n(RE-SLURRY 1)", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    lvl = 0.74 + 0.02 * math.sin(frame_idx * 0.1)
    lvl_top = int(560 - 320 * lvl)
    d.rectangle([83, lvl_top, 277, 557], fill=(180, 215, 245))
    d.text((130, 400), f"LEVEL: {lvl*100:.1f} %", font=FONT_BODY_BOLD, fill=(20, 60, 120))
    d.rectangle([165, 210, 195, 240], fill=(70, 85, 100))
    d.line([180, 240, 180, 500], fill=(50, 50, 50), width=4)
    blade_angle = frame_idx * 0.35
    b_len = 35
    bx1 = 180 + b_len * math.cos(blade_angle)
    by1 = 500 + 10 * math.sin(blade_angle)
    bx2 = 180 - b_len * math.cos(blade_angle)
    by2 = 500 - 10 * math.sin(blade_angle)
    d.line([bx1, by1, bx2, by2], fill=C_DARK_TEXT, width=6)
    d.text((140, 185), "AG-20201 [RUN]", font=FONT_STATUS, fill=C_RUN_GREEN)

    # Tank 2: Re-Slurry TS-20202
    d.rectangle([340, 240, 540, 560], fill=(230, 235, 242), outline=C_DARK_TEXT, width=3)
    d.text((370, 205), "TANK TS-20202\n(RE-SLURRY 2)", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    lvl2 = 0.62 + 0.01 * math.cos(frame_idx * 0.08)
    lvl_top2 = int(560 - 320 * lvl2)
    d.rectangle([343, lvl_top2, 537, 557], fill=(180, 215, 245))
    d.text((390, 400), f"LEVEL: {lvl2*100:.1f} %", font=FONT_BODY_BOLD, fill=(20, 60, 120))
    d.rectangle([425, 210, 455, 240], fill=(70, 85, 100))
    d.line([440, 240, 440, 500], fill=(50, 50, 50), width=4)
    bx1 = 440 + b_len * math.cos(blade_angle + 1)
    by1 = 500 + 10 * math.sin(blade_angle + 1)
    bx2 = 440 - b_len * math.cos(blade_angle + 1)
    by2 = 500 - 10 * math.sin(blade_angle + 1)
    d.line([bx1, by1, bx2, by2], fill=C_DARK_TEXT, width=6)
    d.text((400, 185), "AG-20202 [RUN]", font=FONT_STATUS, fill=C_RUN_GREEN)

    # Transfer Pump PC-20201
    d.ellipse([270, 650, 350, 730], fill=C_RUN_GREEN, outline=C_DARK_TEXT, width=3)
    d.text((285, 680), "PUMP\nPC-20201", font=FONT_STATUS, fill=C_WHITE)
    d.line([180, 560, 180, 690, 270, 690], fill=C_PIPE_GRAY, width=6)
    d.line([440, 560, 440, 690, 350, 690], fill=C_PIPE_GRAY, width=6)
    
    # Discharge pipe to pH adjustment tank
    d.line([350, 690, 650, 690, 650, 380, 750, 380], fill=C_PIPE_GRAY, width=6)
    d.rectangle([460, 650, 570, 715], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((470, 658), "FT-20201", font=FONT_TAG, fill=C_DARK_TEXT)
    flow_val = 44.8 + 0.6 * math.sin(frame_idx * 0.15)
    d.text((470, 682), f"{flow_val:.1f} m3/h", font=FONT_VALUE, fill=C_WATER_BLUE)

    # Tank 3: pH Adjustment TS-40201
    d.rectangle([750, 240, 950, 560], fill=(230, 235, 242), outline=C_DARK_TEXT, width=3)
    d.text((770, 205), "pH ADJUST TANK\nTS-40201", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    lvl3 = 0.71
    lvl_top3 = int(560 - 320 * lvl3)
    d.rectangle([753, lvl_top3, 947, 557], fill=(240, 210, 170))
    d.rectangle([980, 320, 1090, 385], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((990, 328), "AT-40201 (pH)", font=FONT_TAG, fill=C_DARK_TEXT)
    d.text((990, 352), "6.25 pH", font=FONT_VALUE, fill=C_RUN_GREEN)

    # Infeed Pump PC-40203
    d.ellipse([1020, 500, 1100, 580], fill=C_RUN_GREEN, outline=C_DARK_TEXT, width=3)
    d.text((1035, 530), "PUMP\nPC-40203", font=FONT_STATUS, fill=C_WHITE)
    d.line([950, 540, 1020, 540], fill=C_PIPE_GRAY, width=6)

    # Jet Cooker Chamber Unit
    d.rectangle([1180, 480, 1420, 600], fill=(210, 220, 235), outline=C_DARK_TEXT, width=3)
    d.text((1205, 495), "CONTINUOUS JET COOKER", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    d.text((1210, 530), "Direct Steam Injection", font=FONT_BODY, fill=C_GRAY_TEXT)
    d.line([1100, 540, 1180, 540], fill=C_PIPE_GRAY, width=6)
    
    # Steam feed line into Jet Cooker
    d.line([1300, 300, 1300, 480], fill=C_STEAM_CYAN, width=6)
    d.text((1240, 275), "STEAM (8.4 barg)", font=FONT_TAG, fill=C_STEAM_CYAN)
    d.polygon([(1290, 360), (1310, 360), (1300, 380)], fill=C_WHITE, outline=C_DARK_TEXT)
    d.polygon([(1290, 400), (1310, 400), (1300, 380)], fill=C_WHITE, outline=C_DARK_TEXT)
    d.text((1320, 375), "TCV-40201 [64%]", font=FONT_TAG, fill=C_DARK_TEXT)
    
    d.rectangle([1460, 460, 1580, 530], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((1470, 468), "TT-40201", font=FONT_TAG, fill=C_DARK_TEXT)
    temp_val = 105.2 + 0.3 * math.cos(frame_idx * 0.2)
    d.text((1470, 495), f"{temp_val:.1f} °C", font=FONT_VALUE, fill=C_ALARM_RED)
    
    d.rectangle([1460, 545, 1580, 615], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((1470, 553), "PT-40201", font=FONT_TAG, fill=C_DARK_TEXT)
    press_val = 4.3 + 0.05 * math.sin(frame_idx * 0.2)
    d.text((1470, 580), f"{press_val:.2f} barg", font=FONT_VALUE, fill=C_WATER_BLUE)
    
    d.line([1420, 540, 1680, 540], fill=C_PIPE_GRAY, width=6)
    d.text((1600, 515), "TO SPRAY DRYER INFEED -->", font=FONT_TAG, fill=C_DARK_TEXT)
    
    if 300 <= frame_idx <= 450:
        d.rectangle([250, 740, 520, 820], fill=(255, 255, 220), outline=(180, 170, 80), width=2)
        d.text((260, 750), "TAG: PC-20201 (RE-SLURRY PUMP A)", font=FONT_TAG, fill=C_BLACK)
        d.text((260, 770), "STATUS: RUNNING | AUTO | VFD: 50.0 Hz", font=FONT_BODY, fill=C_RUN_GREEN)
        d.text((260, 790), "INTERLOCKS: HEALTHY | CURRENT: 42.1 A", font=FONT_BODY, fill=C_DARK_TEXT)

    progress = max(0.0, min(1.0, (frame_idx - 450) / 70.0))
    mx = int(320 + (650 - 320) * progress)
    my = int(700 + (115 - 700) * progress)
    click = (frame_idx > 500)
    draw_mouse_cursor(d, mx, my, click)

def render_scene_3_spray_dryer(d, frame_idx):
    d.text((40, 150), "AREA 602: SPRAY DRYING TOWER, HIGH PRESSURE PUMP SKID & AIR HEATER", font=FONT_TITLE, fill=C_DARK_TEXT)
    
    # 1. High Pressure Triplex Pump PD-60202 (Left)
    d.rectangle([80, 480, 320, 720], fill=C_PANEL_BG, outline=C_PANEL_BORDER, width=2)
    d.text((100, 495), "TRIPLEX HP PUMP PD-60202", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    d.text((100, 525), "Manufacturer: DEDERT", font=FONT_BODY, fill=C_GRAY_TEXT)
    d.ellipse([140, 560, 240, 660], fill=C_RUN_GREEN, outline=C_DARK_TEXT, width=3)
    d.text((160, 600), "300 BARG\nRUNNING", font=FONT_STATUS, fill=C_WHITE)
    
    # High Pressure Readout PT-60201
    d.rectangle([100, 740, 280, 810], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((110, 748), "PT-60201 (DISCHARGE)", font=FONT_TAG, fill=C_DARK_TEXT)
    hp_press = 282.5 + 1.2 * math.sin(frame_idx * 0.3)
    d.text((110, 775), f"{hp_press:.1f} barg", font=FONT_VALUE, fill=C_ALARM_RED)

    # High pressure pipe to top of chamber
    d.line([240, 610, 420, 610, 420, 260, 580, 260], fill=C_PIPE_GRAY, width=6)
    
    # 2. Drying Chamber CHAM-21 (Center)
    d.rectangle([580, 240, 880, 540], fill=(225, 230, 238), outline=C_DARK_TEXT, width=3)
    d.polygon([(580, 540), (880, 540), (730, 720)], fill=(225, 230, 238), outline=C_DARK_TEXT)
    d.text((660, 340), "DRYING CHAMBER\n    CHAM-21\n(18K TPA CAPACITY)", font=FONT_TITLE, fill=C_DARK_TEXT)
    
    d.polygon([(730, 260), (660, 320), (800, 320)], fill=(200, 230, 255), outline=C_WATER_BLUE)
    d.text((700, 275), "16 NOZZLES", font=FONT_TAG, fill=C_DARK_TEXT)
    
    # Air Heater FH-60210A & Fan
    d.rectangle([380, 340, 520, 440], fill=(255, 235, 220), outline=C_DARK_TEXT, width=2)
    d.text((395, 355), "AIR HEATER\nFH-60210A", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    d.text((395, 400), "GAS BURNER [ON]", font=FONT_STATUS, fill=C_ALARM_RED)
    d.line([520, 390, 550, 390, 550, 290, 580, 290], fill=(220, 100, 50), width=8)
    
    # Temperatures
    d.rectangle([560, 175, 700, 235], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((570, 182), "TT-60210 (INLET)", font=FONT_TAG, fill=C_DARK_TEXT)
    inlet_temp = 195.4 + 0.4 * math.sin(frame_idx * 0.1)
    d.text((570, 205), f"{inlet_temp:.1f} °C", font=FONT_VALUE, fill=C_ALARM_RED)
    
    d.rectangle([740, 680, 880, 740], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((750, 688), "TT-60202 (OUTLET)", font=FONT_TAG, fill=C_DARK_TEXT)
    outlet_temp = 92.2 + 0.3 * math.cos(frame_idx * 0.12)
    d.text((750, 712), f"{outlet_temp:.1f} °C", font=FONT_VALUE, fill=C_WATER_BLUE)

    # Cyclone & Baghouse
    d.line([730, 720, 940, 720, 940, 460, 1020, 460], fill=C_PIPE_GRAY, width=8)
    d.rectangle([1020, 400, 1140, 540], fill=(220, 225, 235), outline=C_DARK_TEXT, width=2)
    d.polygon([(1020, 540), (1140, 540), (1080, 640)], fill=(220, 225, 235), outline=C_DARK_TEXT)
    d.text((1035, 450), "CYCLONE\nCYCL-21", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    
    # Baghouse
    d.rectangle([1220, 320, 1440, 620], fill=(215, 222, 232), outline=C_DARK_TEXT, width=3)
    d.text((1250, 340), "BAGHOUSE DUST\nCOLLECTOR DTEX-21", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    d.text((1250, 400), "Pulse-Jet Cleaning: OK\nExhaust Air: 42,000 m3/h", font=FONT_BODY, fill=C_GRAY_TEXT)
    
    d.rectangle([1240, 470, 1420, 535], fill=C_WHITE, outline=C_PANEL_BORDER, width=2)
    d.text((1250, 478), "DPT-60203 (DIFF PRESS)", font=FONT_TAG, fill=C_DARK_TEXT)
    dp_val = 124.0 + 1.5 * math.sin(frame_idx * 0.2)
    d.text((1250, 502), f"{dp_val:.1f} mmH2O", font=FONT_VALUE, fill=C_RUN_GREEN)

    # Exhaust Fan BF-60240
    d.ellipse([1500, 420, 1600, 520], fill=C_RUN_GREEN, outline=C_DARK_TEXT, width=3)
    d.text((1520, 455), "EXHAUST\n  FAN\nBF-60240", font=FONT_STATUS, fill=C_WHITE)
    d.line([1440, 470, 1500, 470], fill=C_PIPE_GRAY, width=8)
    d.line([1600, 470, 1660, 470, 1660, 300], fill=C_PIPE_GRAY, width=8)
    d.text((1630, 270), "EXHAUST STACK", font=FONT_TAG, fill=C_DARK_TEXT)

    d.rectangle([1050, 640, 1110, 690], fill=C_WHITE, outline=C_DARK_TEXT, width=2)
    d.text((1055, 658), "RV-60201", font=FONT_TAG, fill=C_DARK_TEXT)
    d.line([1080, 690, 1080, 780, 1500, 780], fill=C_PIPE_GRAY, width=6)
    d.text((1250, 795), "STARCH POWDER TO PACKING SILO (AREA 613) -->", font=FONT_TAG, fill=C_DARK_TEXT)

    progress = max(0.0, min(1.0, (frame_idx - 780) / 70.0))
    mx = int(600 + (320 - 600) * progress)
    my = int(300 + (380 - 300) * progress)
    draw_mouse_cursor(d, mx, my, (frame_idx > 840))

def render_scene_4_faceplate(d, frame_idx):
    render_scene_3_spray_dryer(d, 800)
    
    fx1, fy1, fx2, fy2 = 620, 240, 1300, 840
    d.rectangle([fx1, fy1, fx2, fy2], fill=(238, 242, 248), outline=(50, 70, 95), width=3)
    
    d.rectangle([fx1, fy1, fx2, fy1 + 42], fill=(24, 45, 75))
    d.text((fx1 + 18, fy1 + 10), "FCV-60201 : SPRAY DRYER SLURRY INFEED CONTROL VALVE", font=FONT_SUBTITLE, fill=C_WHITE)
    d.rectangle([fx2 - 36, fy1 + 6, fx2 - 8, fy1 + 34], fill=(200, 50, 50))
    d.text((fx2 - 27, fy1 + 10), "X", font=FONT_BODY_BOLD, fill=C_WHITE)
    
    sub_x = fx1 + 15
    for st in ["[ MAIN ]", "[ LIMITS ]", "[ ALARMS ]", "[ TREND ]", "[ INTERLOCKS ]"]:
        d.rectangle([sub_x, fy1 + 48, sub_x + 115, fy1 + 78], fill=(215, 225, 238) if "MAIN" in st else C_WHITE, outline=(180, 195, 210), width=1)
        d.text((sub_x + 10, fy1 + 54), st, font=FONT_TAG, fill=C_DARK_TEXT)
        sub_x += 125
        
    d.text((fx1 + 25, fy1 + 95), "CONTROL MODE:", font=FONT_BODY_BOLD, fill=C_DARK_TEXT)
    modes = [("MAN", False), ("AUTO", True), ("CASCADE", False)]
    mx_btn = fx1 + 160
    for mname, mact in modes:
        d.rectangle([mx_btn, fy1 + 90, mx_btn + 90, fy1 + 120], fill=C_NAV_ACTIVE if mact else C_WHITE, outline=(160, 175, 195), width=1)
        d.text((mx_btn + 22, fy1 + 98), mname, font=FONT_BODY_BOLD, fill=C_WHITE if mact else C_DARK_TEXT)
        mx_btn += 105

    sp_val = 40.0
    if frame_idx >= 960:
        ramp = min(1.0, (frame_idx - 960) / 120.0)
        sp_val = 40.0 + 4.0 * ramp
    pv_val = 39.8 + (sp_val - 40.0) * 0.95 + 0.2 * math.sin(frame_idx * 0.2)
    cv_val = 62.5 + (sp_val - 40.0) * 2.2 + 0.3 * math.cos(frame_idx * 0.2)
    
    d.rectangle([fx1 + 25, fy1 + 140, fx1 + 340, fy1 + 420], fill=C_WHITE, outline=(190, 205, 220), width=2)
    d.text((fx1 + 40, fy1 + 160), "SETPOINT (SP):", font=FONT_BODY_BOLD, fill=C_DARK_TEXT)
    d.rectangle([fx1 + 180, fy1 + 152, fx1 + 320, fy1 + 192], fill=(240, 248, 255), outline=C_NAV_ACTIVE, width=2)
    d.text((fx1 + 200, fy1 + 162), f"{sp_val:.1f} m3/h", font=FONT_VALUE, fill=C_NAV_ACTIVE)

    d.text((fx1 + 40, fy1 + 220), "PROCESS VAR (PV):", font=FONT_BODY_BOLD, fill=C_DARK_TEXT)
    d.rectangle([fx1 + 180, fy1 + 212, fx1 + 320, fy1 + 252], fill=(245, 255, 245), outline=C_RUN_GREEN, width=2)
    d.text((fx1 + 200, fy1 + 222), f"{pv_val:.1f} m3/h", font=FONT_VALUE, fill=C_RUN_GREEN)

    d.text((fx1 + 40, fy1 + 280), "OUTPUT (CV):", font=FONT_BODY_BOLD, fill=C_DARK_TEXT)
    d.rectangle([fx1 + 180, fy1 + 272, fx1 + 320, fy1 + 312], fill=(255, 250, 240), outline=(220, 140, 20), width=2)
    d.text((fx1 + 200, fy1 + 282), f"{cv_val:.1f} %", font=FONT_VALUE, fill=(200, 120, 10))

    d.text((fx1 + 40, fy1 + 345), "VALVE POSITION: OPEN (71.2%)", font=FONT_BODY_BOLD, fill=C_DARK_TEXT)
    d.text((fx1 + 40, fy1 + 375), "PID STATUS: NORMAL / IN REGULATION", font=FONT_STATUS, fill=C_RUN_GREEN)

    d.rectangle([fx1 + 370, fy1 + 140, fx2 - 25, fy1 + 420], fill=C_WHITE, outline=(190, 205, 220), width=2)
    d.text((fx1 + 390, fy1 + 155), "DYNAMIC BAR GRAPH", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    
    bar_base_y = fy1 + 390
    bar_max_h = 200
    
    sp_h = int(bar_max_h * (sp_val / 60.0))
    d.rectangle([fx1 + 420, bar_base_y - sp_h, fx1 + 470, bar_base_y], fill=C_NAV_ACTIVE)
    d.text((fx1 + 428, bar_base_y + 10), "SP", font=FONT_BODY_BOLD, fill=C_NAV_ACTIVE)
    
    pv_h = int(bar_max_h * (pv_val / 60.0))
    d.rectangle([fx1 + 490, bar_base_y - pv_h, fx1 + 540, bar_base_y], fill=C_RUN_GREEN)
    d.text((fx1 + 498, bar_base_y + 10), "PV", font=FONT_BODY_BOLD, fill=C_RUN_GREEN)

    cv_h = int(bar_max_h * (cv_val / 100.0))
    d.rectangle([fx1 + 560, bar_base_y - cv_h, fx1 + 610, bar_base_y], fill=(230, 130, 20))
    d.text((fx1 + 568, bar_base_y + 10), "CV", font=FONT_BODY_BOLD, fill=(230, 130, 20))

    d.rectangle([fx1 + 25, fy1 + 440, fx2 - 25, fy2 - 25], fill=(245, 248, 252), outline=(200, 215, 230), width=1)
    d.text((fx1 + 40, fy1 + 455), "PERMISSIVES & INTERLOCK SAFETY CHECKS:", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    checks = [
        ("TS-40203 Infeed Tank Level > 25%", True),
        ("High Pressure Pump PD-60202 Running Status", True),
        ("Drying Chamber CHAM-21 Deluge Valve Closed", True),
        ("Burner BMS Flame Scanner Signal Normal", True)
    ]
    cy = fy1 + 485
    for ctext, cok in checks:
        d.ellipse([fx1 + 40, cy + 3, fx1 + 52, cy + 15], fill=C_RUN_GREEN)
        d.text((fx1 + 65, cy), ctext, font=FONT_BODY, fill=C_DARK_TEXT)
        d.text((fx2 - 120, cy), "[ HEALTHY ]", font=FONT_STATUS, fill=C_RUN_GREEN)
        cy += 24

    if frame_idx < 950:
        mx, my = fx1 + 250, fy1 + 172
        click = (frame_idx > 920)
    elif frame_idx < 1080:
        mx, my = fx1 + 270, fy1 + 172
        click = False
    else:
        progress = min(1.0, (frame_idx - 1080) / 40.0)
        mx = int((fx1 + 270) + (fx2 - 22 - (fx1 + 270)) * progress)
        my = int((fy1 + 172) + (fy1 + 20 - (fy1 + 172)) * progress)
        click = (frame_idx > 1130)
    draw_mouse_cursor(d, mx, my, click)

def render_scene_5_alarm(d, frame_idx):
    render_scene_3_spray_dryer(d, 800)
    
    is_spiked = (frame_idx >= 1190)
    is_acked = (frame_idx >= 1290)
    
    dpt_val = 124.0
    if is_spiked:
        if not is_acked:
            dpt_val = 168.5 + 1.0 * math.sin(frame_idx * 0.4)
            box_col = C_ALARM_RED if ((frame_idx // 8) % 2 == 0) else C_WARN_AMBER
            status_text = "ALARM: HIGH-HIGH (> 160)"
            status_col = C_ALARM_RED
        else:
            rec = min(1.0, (frame_idx - 1290) / 100.0)
            dpt_val = 168.5 - 42.0 * rec + 0.5 * math.sin(frame_idx * 0.2)
            box_col = (180, 120, 0)
            status_text = "ACKED - JET PULSE ACTIVE"
            status_col = (180, 120, 0)
    else:
        box_col = C_PANEL_BORDER
        status_text = "NORMAL (< 160)"
        status_col = C_RUN_GREEN
        
    d.rectangle([1220, 460, 1460, 560], fill=(255, 250, 240) if is_spiked else C_WHITE, outline=box_col, width=4)
    d.text((1235, 470), "DPT-60203 (DIFF PRESS)", font=FONT_SUBTITLE, fill=C_DARK_TEXT)
    d.text((1235, 498), f"{dpt_val:.1f} mmH2O", font=get_font(22, bold=True), fill=box_col if is_spiked else C_RUN_GREEN)
    d.text((1235, 532), status_text, font=FONT_STATUS, fill=status_col)
    
    if is_spiked:
        d.rectangle([600, 750, 1320, 870], fill=(255, 245, 245) if not is_acked else (255, 255, 235), outline=box_col, width=3)
        d.rectangle([600, 750, 1320, 785], fill=box_col)
        card_title = "CRITICAL PROCESS WARNING: DPT-60203" if not is_acked else "ALARM ACKNOWLEDGED — RECOVERY SEQUENCE INITIATED"
        d.text((620, 758), card_title, font=FONT_SUBTITLE, fill=C_WHITE)
        d.text((620, 798), "Location: DTEX-21 Dust Collector (Spray Dryer Area 602)", font=FONT_BODY_BOLD, fill=C_DARK_TEXT)
        d.text((620, 822), f"Condition: Differential Pressure {dpt_val:.1f} mmH2O exceeded Warning Limit (160.0 mmH2O)", font=FONT_BODY, fill=C_DARK_TEXT)
        d.text((620, 844), "Action: Auto Reverse-Pulse Purge Triggered  |  Eng Chonlawit Logged", font=FONT_BODY, fill=C_GRAY_TEXT)

    if frame_idx < 1290:
        progress = max(0.0, min(1.0, (frame_idx - 1200) / 70.0))
        mx = int(960 + (1620 - 960) * progress)
        my = int(600 + (70 - 600) * progress)
        click = (frame_idx > 1270)
    else:
        mx, my = 1620, 70
        click = False
    draw_mouse_cursor(d, mx, my, click)

def render_scene_6_trends(d, frame_idx):
    d.text((40, 150), "SCADA TREND RECORDER & REAL-TIME ANALYTICS (AREA 602 & UTILITIES)", font=FONT_TITLE, fill=C_DARK_TEXT)
    
    tx1, ty1, tx2, ty2 = 60, 190, 1860, 960
    d.rectangle([tx1, ty1, tx2, ty2], fill=(20, 24, 32), outline=(60, 75, 95), width=3)
    
    d.rectangle([tx1, ty1, tx2, ty1 + 45], fill=(30, 38, 50))
    d.text((tx1 + 20, ty1 + 12), "LIVE TREND DISPLAY [SAMPLE INTERVAL: 100 ms | BUFFER: 24 HOURS]", font=FONT_SUBTITLE, fill=C_WHITE)
    
    tb_x = tx2 - 500
    for tb_btn in ["15 MIN", "1 HOUR", "8 HOURS", "ZOOM", "EXPORT CSV"]:
        d.rectangle([tb_x, ty1 + 8, tb_x + 85, ty1 + 36], fill=(50, 62, 80) if "1 HOUR" not in tb_btn else C_NAV_ACTIVE)
        d.text((tb_x + 12, ty1 + 15), tb_btn, font=FONT_TAG, fill=C_WHITE)
        tb_x += 95
        
    cx1, cy1, cx2, cy2 = tx1 + 100, ty1 + 65, tx2 - 50, ty2 - 80
    d.rectangle([cx1, cy1, cx2, cy2], fill=(12, 16, 22), outline=(45, 55, 70), width=2)
    
    for gy in range(cy1 + 40, cy2, 60):
        d.line([cx1, gy, cx2, gy], fill=(30, 40, 52), width=1)
        norm_val = 100 - int(100 * (gy - cy1) / (cy2 - cy1))
        d.text((cx1 - 45, gy - 8), f"{norm_val}%", font=FONT_TAG, fill=(120, 135, 150))
        
    for gx in range(cx1 + 120, cx2, 140):
        d.line([gx, cy1, gx, cy2], fill=(30, 40, 52), width=1)
        
    pens = [
        ("TT-60210 Dryer Inlet Air Temp (°C)", (240, 70, 70), 195.4, 180, 220, 0.05, 0.0),
        ("TT-60202 Dryer Outlet Air Temp (°C)", (70, 180, 255), 92.2, 70, 110, 0.08, 1.2),
        ("FT-60201 HP Slurry Feed Flow (m3/h)", (50, 220, 110), 44.0, 20, 60, 0.04, 2.5),
        ("PT-60201 Triplex Pump Pressure (barg)", (255, 210, 40), 282.5, 240, 320, 0.12, 0.8)
    ]
    
    chart_w = cx2 - cx1
    t_offset = frame_idx * 2
    
    for pen_label, pen_col, cur_val, min_v, max_v, freq, phase in pens:
        points = []
        for x_step in range(0, chart_w, 4):
            x_pos = cx1 + x_step
            sim_t = (t_offset + x_step) * freq + phase
            val = cur_val + (max_v - min_v) * 0.08 * math.sin(sim_t) + (max_v - min_v) * 0.03 * math.cos(sim_t * 2.3)
            norm_v = (val - min_v) / (max_v - min_v)
            y_pos = int(cy2 - norm_v * (cy2 - cy1))
            y_pos = max(cy1 + 2, min(cy2 - 2, y_pos))
            points.append((x_pos, y_pos))
            
        if len(points) > 1:
            d.line(points, fill=pen_col, width=3)
            
    cursor_x = cx1 + (frame_idx * 5) % chart_w
    d.line([cursor_x, cy1, cursor_x, cy2], fill=(255, 255, 255), width=2)
    
    d.rectangle([cursor_x + 10, cy1 + 20, cursor_x + 280, cy1 + 130], fill=(25, 32, 44), outline=(70, 90, 115), width=2)
    d.text((cursor_x + 20, cy1 + 28), "SAMPLE AT CURSOR:", font=FONT_STATUS, fill=(180, 200, 220))
    d.text((cursor_x + 20, cy1 + 48), f"TT-60210 : 195.4 °C", font=FONT_TAG, fill=(240, 70, 70))
    d.text((cursor_x + 20, cy1 + 68), f"TT-60202 :  92.2 °C", font=FONT_TAG, fill=(70, 180, 255))
    d.text((cursor_x + 20, cy1 + 88), f"FT-60201 :  44.0 m3/h", font=FONT_TAG, fill=(50, 220, 110))
    d.text((cursor_x + 20, cy1 + 108), f"PT-60201 : 282.5 barg", font=FONT_TAG, fill=(255, 210, 40))

    lx = cx1 + 20
    ly = cy2 + 18
    for pen_label, pen_col, cur_val, min_v, max_v, freq, phase in pens:
        d.rectangle([lx, ly + 2, lx + 25, ly + 14], fill=pen_col)
        d.text((lx + 32, ly), pen_label, font=FONT_TAG, fill=C_WHITE)
        lx += 370

def render_scene_7_outro(d, frame_idx):
    d.rectangle([0, 138, WIDTH, 1032], fill=(16, 22, 32))
    
    for gx in range(0, WIDTH, 80):
        d.line([gx, 138, gx, 1032], fill=(24, 34, 48), width=1)
    for gy in range(138, 1032, 80):
        d.line([0, gy, WIDTH, gy], fill=(24, 34, 48), width=1)
        
    ox1, oy1, ox2, oy2 = 360, 220, 1560, 940
    d.rectangle([ox1, oy1, ox2, oy2], fill=(24, 32, 45), outline=(50, 110, 190), width=3)
    
    d.rectangle([ox1, oy1, ox2, oy1 + 90], fill=(13, 71, 161))
    d.text((ox1 + 40, oy1 + 20), "INGREDION (THAILAND) CO., LTD. — KALASIN PLANT", font=get_font(26, bold=True), fill=C_WHITE)
    d.text((ox1 + 42, oy1 + 55), "PROJECT SPRINT: 18K TPA SPRAY DRYER PLANT AUTOMATION", font=FONT_SUBTITLE, fill=(180, 215, 255))
    
    d.text((ox1 + 50, oy1 + 120), "SCADA & DISTRIBUTED CONTROL SYSTEM (DCS) OVERVIEW", font=FONT_TITLE, fill=(255, 215, 0))
    
    bullets = [
        ("P&ID Baseline Alignment", "Complete coverage of all 38 drawing sheets (Rev 3.5 baseline, CAD DWG source)."),
        ("Rockwell PlantPAx 5.x Integration", "Standardized global objects, faceplates, interlocks and motor/valve templates."),
        ("High Performance HMI (ANSI/ISA-101)", "Ergonomic neutral-gray scheme with contextual color-coded process alerts."),
        ("Advanced Alarm Management (ANSI/ISA-18.2)", "Prioritized alarm routing, latching states, and time-stamped sequence of events."),
        ("472 I/O Tag Infrastructure", "Mapped across RIO-200 (Slurry), RIO-400 (Infeed), RIO-600 (Dryer) & RIO-MCC."),
        ("Safety & Critical Protection", "Dedicated BMS burner safety, chamber deluge, and IEP explosion suppression monitoring.")
    ]
    
    by = oy1 + 180
    for b_title, b_desc in bullets:
        d.text((ox1 + 60, by), f"[✔] {b_title}", font=get_font(15, bold=True), fill=C_WHITE)
        d.text((ox1 + 400, by + 1), b_desc, font=FONT_BODY, fill=(195, 210, 230))
        by += 44
        
    d.line([ox1 + 40, oy1 + 470, ox2 - 40, oy1 + 470], fill=(45, 65, 90), width=2)
    d.text((ox1 + 60, oy1 + 490), "Project Reference: xPrj-2603001-Kalasin | Deliverable: SCADA Simulation & Demo Video", font=FONT_BODY_BOLD, fill=(220, 230, 245))
    d.text((ox1 + 60, oy1 + 520), "Consultant & Engineering Integration: AEC Industrial Engineering / Systems Team", font=FONT_BODY, fill=(160, 180, 205))
    d.text((ox1 + 60, oy1 + 550), f"Generated: {time.strftime('%d-%b-%2026')} | Status: VERIFIED & APPROVED", font=FONT_STATUS, fill=C_RUN_GREEN)

# ----------------- MAIN COMPILER -----------------

def main():
    print(f"=== Starting SCADA Demo Video Generation ===")
    print(f"Output File 1: {OUTPUT_FILE_1}")
    print(f"Output File 2: {OUTPUT_FILE_2}")
    print(f"Resolution: {WIDTH}x{HEIGHT} | FPS: {FPS} | Total Seconds: {TOTAL_SECONDS} | Total Frames: {TOTAL_FRAMES}")

    start_time = time.time()
    
    writer1 = imageio.get_writer(
        OUTPUT_FILE_1,
        fps=FPS,
        codec='libx264',
        quality=9,
        pixelformat='yuv420p',
        ffmpeg_params=['-preset', 'fast', '-crf', '18']
    )

    for frame_idx in range(TOTAL_FRAMES):
        im = Image.new("RGB", (WIDTH, HEIGHT), C_CANVAS_BG)
        d = ImageDraw.Draw(im)
        
        alarm_state = "NORMAL"
        alarm_msg = ""
        
        if frame_idx < 240:
            tab_idx = 0
            render_scene_1_overview(d, frame_idx)
        elif frame_idx < 540:
            tab_idx = 1
            render_scene_2_slurry(d, frame_idx)
        elif frame_idx < 870:
            tab_idx = 3
            render_scene_3_spray_dryer(d, frame_idx)
        elif frame_idx < 1170:
            tab_idx = 3
            render_scene_4_faceplate(d, frame_idx)
        elif frame_idx < 1410:
            tab_idx = 3
            if frame_idx < 1190:
                alarm_state = "NORMAL"
            elif frame_idx < 1290:
                alarm_state = "ACTIVE"
                alarm_msg = "ALM-60203 DTEX-21 BAGHOUSE DIFF PRESS HIGH (> 160 mmH2O)"
            else:
                alarm_state = "ACKNOWLEDGED"
                alarm_msg = "ALM-60203 DTEX-21 DIFF PRESS ACKED BY ENG_CHONLAWIT"
            render_scene_5_alarm(d, frame_idx)
        elif frame_idx < 1650:
            tab_idx = 7
            render_scene_6_trends(d, frame_idx)
        else:
            tab_idx = 0
            render_scene_7_outro(d, frame_idx)

        draw_top_banner(d, frame_idx, tab_idx, alarm_state, alarm_msg)
        draw_bottom_footer(d)
        
        frame_np = np.array(im)
        writer1.append_data(frame_np)
        
        if (frame_idx + 1) % 150 == 0 or frame_idx == TOTAL_FRAMES - 1:
            elapsed = time.time() - start_time
            fps_speed = (frame_idx + 1) / elapsed
            print(f"Rendered {frame_idx + 1}/{TOTAL_FRAMES} frames ({(frame_idx + 1)/TOTAL_FRAMES*100:.1f}%) at {fps_speed:.1f} fps")

    writer1.close()
    
    shutil.copyfile(OUTPUT_FILE_1, OUTPUT_FILE_2)
    
    size_mb1 = os.path.getsize(OUTPUT_FILE_1) / (1024 * 1024)
    size_mb2 = os.path.getsize(OUTPUT_FILE_2) / (1024 * 1024)
    total_time = time.time() - start_time
    
    print(f"\n=== Video Generation Completed Successfully! ===")
    print(f"File 1: {OUTPUT_FILE_1} ({size_mb1:.2f} MB)")
    print(f"File 2: {OUTPUT_FILE_2} ({size_mb2:.2f} MB)")
    print(f"Total Time: {total_time:.2f} s")

if __name__ == "__main__":
    main()
