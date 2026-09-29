#!/usr/bin/env python3
"""
================================================================================
Kalasin Engineering xCIP System (Project Ref: x2608003)
CHASSIS C1-C8 MASTER SLOT TO DESTINATION JUNCTION BOX PDF REPORT (TYPST)
================================================================================
Generates:
  1. 03_IO_Lists_and_Schedules/C1_C8_Chassis_Slot_to_Destination_JB_Master_Report.pdf
  2. 02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/C1_C8_Chassis_Slot_to_Destination_JB_Master_Report.pdf
================================================================================
"""

import os
import subprocess
from collections import defaultdict, Counter
import openpyxl

BASE_DIR = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering"
SOURCE_EXCEL = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/IO_List_xDev-R01-Tag35-7.xlsx")
TYP_FILE = os.path.join(BASE_DIR, "04_Automation_Scripts/chassis_c1_c8_master_report.typ")
PDF_OUT_1 = os.path.join(BASE_DIR, "03_IO_Lists_and_Schedules/C1_C8_Chassis_Slot_to_Destination_JB_Master_Report.pdf")
PDF_OUT_2 = os.path.join(BASE_DIR, "02_Electrical_and_eDrawing/04_PDF_Reports_and_Drawings/C1_C8_Chassis_Slot_to_Destination_JB_Master_Report.pdf")
TYPST_BIN = "/opt/homebrew/bin/typst"

def escape_typst(s):
    if not s: return ""
    return str(s).replace('\\', '\\\\').replace('"', '\\"').replace('$', '\\$').replace('#', '\\#').replace('_', '\\_')

def load_all_records():
    wb = openpyxl.load_workbook(SOURCE_EXCEL, data_only=True)
    ws = wb['IO List']
    
    records = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        col7 = str(r[2]).strip() if r[2] is not None else ''
        ch = str(r[30]).strip() if r[30] is not None else ''
        slot = str(r[31]).strip() if r[31] is not None else ''
        if not ch and col7.startswith('C') and 'S' in col7:
            ch = col7.split('S')[0]
            slot = col7.split('S')[1].split(',')[0]
            
        if ch not in [f"C{i}" for i in range(1, 9)]:
            continue
            
        dest = str(r[1]).strip() if r[1] is not None else ''
        card = str(r[3]).strip() if r[3] is not None else ''
        card_type = str(r[26]).strip() if r[26] is not None else ''
        term = str(r[5]).strip() if r[5] is not None else ''
        tag = str(r[11]).strip() if r[11] is not None else ''
        desc = str(r[12]).strip() if r[12] is not None else ''
        io_t = str(r[20]).strip() if r[20] is not None else ''
        pt = str(r[32]).strip() if r[32] is not None else ''
        pid = str(r[14]).strip() if r[14] is not None else ''
        
        c_name = card if card and card != '#N/A' else card_type
        if '-' in c_name:
            parts = c_name.split('-')
            if len(parts) >= 2 and not parts[1].isdigit():
                clean_c = f"{parts[0]}-{parts[1]}"
            else:
                clean_c = c_name
        else:
            clean_c = c_name if c_name and c_name != '#N/A' else io_t
            
        is_sp = 'SPARE' in tag.upper() or not tag or tag == '#N/A' or tag == '0'
        records.append({
            'chassis': ch,
            'slot': int(slot) if slot.isdigit() else slot,
            'pt': int(pt) if str(pt).isdigit() else pt,
            'dest': dest,
            'card': clean_c,
            'term': term,
            'tag': tag if not is_sp else 'SPARE',
            'desc': desc if (desc and desc != '#N/A' and desc != '0') else ('Spare Reserve Channel' if is_sp else '-'),
            'io_t': io_t,
            'pid': pid if pid != '#N/A' else '-',
            'is_sp': is_sp
        })
    return records

def generate_typst_master():
    records = load_all_records()
    print(f"Loaded {len(records)} total records across C1-C8.")
    
    # Chassis statistics
    ch_stats = defaultdict(lambda: {'total': 0, 'active': 0, 'spare': 0, 'DI': 0, 'DO': 0, 'AI': 0, 'AO': 0, 'BUS': 0, 'slots': set(), 'dests': set()})
    for r in records:
        c = r['chassis']
        ch_stats[c]['total'] += 1
        if r['is_sp']: ch_stats[c]['spare'] += 1
        else: ch_stats[c]['active'] += 1
        io = r['io_t']
        if io in ch_stats[c]: ch_stats[c][io] += 1
        ch_stats[c]['slots'].add(r['slot'])
        ch_stats[c]['dests'].add(r['dest'])
        
    # Destination statistics
    dest_stats = defaultdict(lambda: {'total': 0, 'active': 0, 'spare': 0, 'chassis': set()})
    for r in records:
        d = r['dest']
        dest_stats[d]['total'] += 1
        if r['is_sp']: dest_stats[d]['spare'] += 1
        else: dest_stats[d]['active'] += 1
        dest_stats[d]['chassis'].add(r['chassis'])
        
    typ = []
    typ.append("""
#set page(
  paper: "a4",
  flipped: true,
  margin: (top: 1.2cm, bottom: 1.2cm, left: 1.4cm, right: 1.4cm),
  header: context {
    grid(
      columns: (1fr, 1fr),
      align: (left, right),
      text(size: 8pt, fill: rgb("64748b"), font: "Arial")[
        *KALASIN ENGINEERING* | Turnkey xCIP-1545 Automation System (Ref: x2608003)
      ],
      text(size: 8pt, fill: rgb("64748b"), font: "Arial")[
        *Chassis C1--C8 Slot to Destination Junction Box Master Report* | Page #counter(page).display()
      ]
    )
    line(length: 100%, stroke: 0.5pt + rgb("cbd5e1"))
  },
  footer: context {
    line(length: 100%, stroke: 0.5pt + rgb("cbd5e1"))
    v(2pt)
    grid(
      columns: (1fr, 1fr),
      align: (left, right),
      text(size: 8pt, fill: rgb("94a3b8"), font: "Arial")[
        Client: Ingredion (Thailand) Co., Ltd. / ISHITON | Document Ref: KAL-IO-REP-C1C8-001
      ],
      text(size: 8pt, fill: rgb("94a3b8"), font: "Arial")[
        Source File: `IO_List_xDev-R01-Tag35-7.xlsx` | Approved Rev: 3.6
      ]
    )
  }
)

#set text(font: "Arial", size: 8.5pt, fill: rgb("0f172a"))
#set par(justify: true, leading: 0.5em)

// --- TITLE BLOCK ---
#block(
  fill: rgb("1e293b"),
  inset: (x: 16pt, y: 11pt),
  radius: 4pt,
  width: 100%
)[
  #grid(
    columns: (2.5fr, 1fr),
    align: (left + horizon, right + horizon),
    [
      #text(fill: white, size: 14pt, weight: "bold")[
        CHASSIS C1 -- C8: MASTER SLOT TO DESTINATION JB REPORT
      ]
      #v(2pt)
      #text(fill: rgb("93c5fd"), size: 9pt, weight: "medium")[
        COMPLETE PLANT AUTOMATION I/O ROUTING DIRECTORY & FIELD DISTRIBUTION SCHEDULE
      ]
    ],
    [
      #text(fill: white, size: 8pt)[
        *Project Ref:* x2608003 #linebreak()
        *Scope:* Turnkey Automated CIP Skid System #linebreak()
        *Drawing Ref:* KAL-SYS-DIA-C1C8-001
      ]
    ]
  )
]

#v(8pt)

// --- SECTION 1: SYSTEM OVERVIEW ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[1. Master Chassis System Overview (C1 through C8)]
#v(4pt)

#table(
  columns: (1fr, 3.2fr, 2.5fr, 1fr, 1.1fr, 1fr, 1fr, 0.8fr, 0.8fr, 0.8fr, 1fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col < 3 { left + horizon } else { center + horizon },
  inset: 4.5pt,

  text(fill: white, weight: "bold", size: 8pt)[Rack],
  text(fill: white, weight: "bold", size: 8pt)[Chassis Description & Functional Role],
  text(fill: white, weight: "bold", size: 8pt)[Enclosure Location],
  text(fill: white, weight: "bold", size: 8pt)[Slots],
  text(fill: white, weight: "bold", size: 8pt)[Total Pts],
  text(fill: white, weight: "bold", size: 8pt)[Active],
  text(fill: white, weight: "bold", size: 8pt)[Spare],
  text(fill: white, weight: "bold", size: 8pt)[DI],
  text(fill: white, weight: "bold", size: 8pt)[DO],
  text(fill: white, weight: "bold", size: 8pt)[AI],
  text(fill: white, weight: "bold", size: 8pt)[AO / BUS],
""")

    ch_desc_map = {
        'C1': ('Main Controller Rack (1756-L950TPSXT)', 'Control Cabinet CA1 (MCR)'),
        'C2': ('Main Expansion Rack (High-Density I/O)', 'Control Cabinet CA1 (MCR)'),
        'C3': ('Remote I/O Rack 1 (Spray Dryer / Tower)', 'Field RIO Cabinet'),
        'C4': ('Remote I/O Rack 2 (Upper Spray Dryer)', 'Field RIO Cabinet'),
        'C5': ('Remote I/O Station 5 (Slurry Building)', 'Slurry Building 2nd Floor'),
        'C6': ('Motor Control Hardwired Interlocks', 'MCC Room Ground Floor'),
        'C7': ('MCC & Fieldbus Supervisory System', 'MCC Room Ground Floor'),
        'C8': ('Intrinsically Safe (IS) Field Barriers', 'Field Ex Barrier Enclosures'),
    }

    tot_pts = 0
    tot_act = 0
    tot_spr = 0
    tot_di = 0
    tot_do = 0
    tot_ai = 0
    tot_ao_bus = 0

    for ch_idx in range(1, 9):
        cid = f"C{ch_idx}"
        st = ch_stats[cid]
        d_title, d_loc = ch_desc_map[cid]
        ao_bus = st['AO'] + st['BUS']
        tot_pts += st['total']
        tot_act += st['active']
        tot_spr += st['spare']
        tot_di += st['DI']
        tot_do += st['DO']
        tot_ai += st['AI']
        tot_ao_bus += ao_bus
        
        typ.append(f'  [*{cid}*], [{d_title}], [{d_loc}], [{len(st["slots"])}], [{st["total"]}], [{st["active"]}], [{st["spare"]}], [{st["DI"]}], [{st["DO"]}], [{st["AI"]}], [{ao_bus}],')

    typ.append(f"""
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[TOTAL]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[All 8 Chassis Combined]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[All Plant Enclosures]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[59]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_pts}]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_act}]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_spr}]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_di}]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_do}]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_ai}]],
  table.cell(fill: rgb("334155"))[#text(fill: white, weight: "bold")[{tot_ao_bus}]]
)
""")

    typ.append("""
#v(10pt)

// --- SECTION 2: DESTINATION DIRECTORY ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[2. Destination Junction Box Directory (16 Destinations)]
#v(4pt)

#table(
  columns: (1.5fr, 2.5fr, 3.5fr, 1.5fr, 1.1fr, 1.1fr, 1.1fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col < 3 { left + horizon } else { center + horizon },
  inset: 4.5pt,

  text(fill: white, weight: "bold", size: 8pt)[Destination],
  text(fill: white, weight: "bold", size: 8pt)[Location],
  text(fill: white, weight: "bold", size: 8pt)[Process Functional Scope],
  text(fill: white, weight: "bold", size: 8pt)[Feeding Racks],
  text(fill: white, weight: "bold", size: 8pt)[Total Pts],
  text(fill: white, weight: "bold", size: 8pt)[Active],
  text(fill: white, weight: "bold", size: 8pt)[Spare],
""")

    jb_desc_map = {
        'CA1': ('Main Control Room', 'Central PLC Automation & Network Cabinets'),
        'JB-401': ('Infeed 2nd Floor', 'Infeed Raw Material Prep & Wet Mixing'),
        'JB-402': ('Jet Cooker 2nd Floor', 'Jet Cooker Thermal Cooking & Starch Conversion'),
        'JB-601': ('Spray Dryer 1st Floor', 'Spray Dryer Bottom Discharge & Pumps'),
        'JB-602': ('Spray Dryer 3rd Floor', 'Spray Dryer Middle Body & Chamber Instruments'),
        'JB-606': ('Spray Dryer 6th Floor', 'Spray Dryer Air Filter & Observation Level'),
        'JB-607': ('Spray Dryer 7th Floor', 'Spray Dryer Cyclone & Bag Filter Exhaust'),
        'JB-608': ('Spray Dryer 8th Floor', 'Spray Dryer Air Dispenser & Scrubber Top'),
        'JB-612': ('Packing Tower 2nd Floor', 'Packing Tower Product Discharge & Conveyor'),
        'JB-618': ('Packing Tower 8th Floor', 'Packing Tower Dehumidifier & Cyclone Air'),
        'RIO-200': ('Slurry Out-Building 2nd Fl', 'Remote Slurry Processing Building Station'),
        'MCC': ('MCC Room Ground Floor', 'Main Motor Control Center & VFD Inverters'),
        'IS-JB-603': ('Spray Dryer 3rd Floor (IS)', 'Intrinsically Safe Ex Zone Chamber Box'),
        'IS-JB-608': ('Spray Dryer 8th Floor (IS)', 'Intrinsically Safe Ex Zone Top Level Box'),
        'IS-JB-612': ('Packing Tower 2nd Fl (IS)', 'Intrinsically Safe Ex Zone Lower Tower Box'),
        'IS-JB-618': ('Packing Tower 8th Fl (IS)', 'Intrinsically Safe Ex Zone Upper Tower Box'),
    }

    for d in sorted(dest_stats.keys()):
        st = dest_stats[d]
        loc_str, func_str = jb_desc_map.get(d, ('Field Area', 'Process Field Enclosure'))
        ch_str = ", ".join(sorted(st['chassis'], key=lambda x: int(x[1:])))
        typ.append(f'  [*{d}*], [{loc_str}], [{func_str}], [{ch_str}], [{st["total"]}], [{st["active"]}], [{st["spare"]}],')

    typ.append("""
)

#pagebreak()

// --- SECTION 3: AGGREGATED SLOT TO JB SCHEDULE ---
#text(size: 11pt, weight: "bold", fill: rgb("1e3a8a"))[3. Master Chassis C1--C8 Slot-to-Destination Schedule]
#v(4pt)

#table(
  columns: (0.7fr, 0.7fr, 1.5fr, 0.8fr, 1.2fr, 2fr, 0.8fr, 0.8fr, 0.8fr, 3.5fr),
  fill: (col, row) => if row == 0 { rgb("1e293b") } else if calc.even(row) { rgb("f8fafc") } else { rgb("ffffff") },
  stroke: (x, y) => if y == 0 { 0.5pt + rgb("475569") } else { 0.5pt + rgb("e2e8f0") },
  align: (col, row) => if col in (0, 1, 3, 4, 6, 7, 8) { center + horizon } else { left + horizon },
  inset: 4pt,

  text(fill: white, weight: "bold", size: 7.5pt)[Chassis],
  text(fill: white, weight: "bold", size: 7.5pt)[Slot],
  text(fill: white, weight: "bold", size: 7.5pt)[Module Card],
  text(fill: white, weight: "bold", size: 7.5pt)[Signal],
  text(fill: white, weight: "bold", size: 7.5pt)[Dest JB],
  text(fill: white, weight: "bold", size: 7.5pt)[Location],
  text(fill: white, weight: "bold", size: 7.5pt)[Total],
  text(fill: white, weight: "bold", size: 7.5pt)[Active],
  text(fill: white, weight: "bold", size: 7.5pt)[Spare],
  text(fill: white, weight: "bold", size: 7.5pt)[Sample Connected Loops / Tags],
""")

    # Group by (chassis, slot, dest)
    csd = defaultdict(list)
    for r in records:
        csd[(r['chassis'], r['slot'], r['dest'])].append(r)
        
    s_keys = sorted(csd.keys(), key=lambda x: (int(x[0][1:]), int(x[1]) if str(x[1]).isdigit() else 99, x[2]))
    for cid, slot_no, d in s_keys:
        items = csd[(cid, slot_no, d)]
        card = escape_typst(items[0]['card'])
        io_t = escape_typst(items[0]['io_t'])
        tot = len(items)
        spr = sum(1 for it in items if it['is_sp'])
        act = tot - spr
        loc = escape_typst(jb_desc_map.get(d, ('Field Area', ''))[0])
        
        act_tags = [it['tag'] for it in items if not it['is_sp']]
        if act_tags:
            tag_str = ", ".join(act_tags[:3])
            if len(act_tags) > 3: tag_str += f" (+{len(act_tags)-3} more)"
        else:
            tag_str = "Reserve / Spare Bank"
            
        typ.append(f'  [*{cid}*], [{slot_no:02d}], [{card}], [{io_t}], [*{d}*], [{loc}], [{tot}], [{act}], [{spr}], [{tag_str}],')

    typ.append(")\n")
    
    with open(TYP_FILE, "w", encoding="utf-8") as f:
        f.write("\n".join(typ))
        
    print(f"Master Typst document created at: {TYP_FILE}")
    print(f"Compiling Typst Master PDF: {PDF_OUT_1}")
    res1 = subprocess.run([TYPST_BIN, "compile", TYP_FILE, PDF_OUT_1], capture_output=True, text=True)
    if res1.returncode != 0:
        print(f"Error compiling: {res1.stderr}")
        return
    res2 = subprocess.run([TYPST_BIN, "compile", TYP_FILE, PDF_OUT_2], capture_output=True, text=True)
    print(f"Master C1-C8 PDF Successfully compiled to:\n - {PDF_OUT_1}\n - {PDF_OUT_2}")

if __name__ == "__main__":
    generate_typst_master()
