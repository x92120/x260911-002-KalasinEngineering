import re

src_path = "/Users/x92120/xApp-001/x260911-002-KalasinEngineering/04_Automation_Scripts/build_loop_test_plan_and_schedule.py"
with open(src_path, "r", encoding="utf-8") as f:
    code = f.read()

# Replace build_master_loop_list_sheet and add build_cad_reference_sheet
old_func_start = "def build_master_loop_list_sheet(ws, inst_list, tag_to_channel):"
old_func_end = "# -----------------------------------------------------------------------------\n# Sheet 5: 04_ITP_Inspection_Plan"

new_funcs = """def classify_cad_typical(item, matched_first):
    tag = item['active_tag'].upper()
    sig_to = item['sig_to'] or ''
    sig_type = item['sig_type'] or ''
    do_sig = item['do']
    di_sig = item['di']
    ao_sig = item['ao']
    ai_sig = item['ai']

    is_is = ("IS" in sig_to) or ("IS-JB" in sig_to)
    if is_is:
        return ("TYP-IS-01", "TB-IS", "LOOP DIAGRAM WRING.dwg", "Intrinsically Safe Ex Loop (Galvanic Barrier in CA1)")

    if ao_sig == 1 or any(k in tag for k in ['FCV', 'PCV', 'LCV', 'TCV', 'CV', 'PV']):
        return ("TYP-AO-01", "TBAO", "LOOP DIAGRAM WRING.dwg", "4-20mA Analog Control Valve SMART Positioner")
    elif ai_sig == 1:
        if any(k in tag for k in ['FT', 'DT', 'AT', 'FIT', 'AIT']):
            return ("TYP-AI-02", "TBAI", "LOOP DIAGRAM WRING.dwg", "4-Wire Active Powered Transmitter (Flow/Density)")
        else:
            return ("TYP-AI-01", "TBAI", "LOOP DIAGRAM WRING.dwg", "2-Wire Loop-Powered 4-20mA Transmitter (PT/LT/TT)")
    elif do_sig == 1 or any(k in tag for k in ['XV', 'SOV', 'SV', 'PMP', 'MTR', 'FAN']):
        return ("TYP-DO-01", "TBDO", "PANEL WIRING DIAGRAM.dwg", "24VDC Discrete Output (Solenoid / Relay Coil)")
    elif di_sig == 1 or any(k in tag for k in ['ZSO', 'ZSC', 'LSH', 'LSL', 'PSH', 'PSL', 'TSH', 'TSL', 'PB', 'SW']):
        return ("TYP-DI-01", "TBDI", "PANEL WIRING DIAGRAM.dwg", "24VDC Discrete Input (Limit / Level / Pressure Switch)")
    elif 'Local' in sig_to or not sig_to or sig_type in ['Local', '-']:
        return ("TYP-MECH", "N/A", "Standard Mechanical Typical", "Mechanical Field Device (Dial Gauge / Manual Valve)")
    else:
        return ("TYP-GEN-01", "TBCI", "LOOP DIAGRAM WRING.dwg", "Auxiliary Package / Serial Comm / Remote Interface")

# -----------------------------------------------------------------------------
# Sheet 4: 03_Loop_Test_Master_List (560 Loops Point-by-Point Test Sheet)
# -----------------------------------------------------------------------------
def build_master_loop_list_sheet(ws, inst_list, tag_to_channel):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:Z1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  MASTER INSTRUMENT LOOP TESTING & CALIBRATION VERIFICATION SCHEDULE",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:Z2")
    set_cell(ws["A2"], "Point-by-Point Field Loop Check Record (Rev. 3.6a Master) Cross-Referenced with CAD Loop Typical Diagrams (DWG/DXF)",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    headers = [
        "Item #", "Active Tag No.", "Legacy Tag", "P&ID No.", "Instrument Service Description",
        "Instrument Type / Name", "Signal Type", "Calibrated Range",
        "CAD Typical ID", "CAD Strip Type", "CAD DWG Reference",
        "Field Junction Box", "Terminal Block ID", "PLC Rack / Slot",
        "Wire Tag (PLC Side)", "Wire Tag (Terminal Side)", "PLC Tag Name",
        "Cold Continuity (<2Ω)", "Cold Megger (>20MΩ)", "Hot Power (24VDC)", "Calibration 0% (4mA)",
        "Calibration 50% (12mA)", "Calibration 100% (20mA)", "Stroke / Trip Response",
        "Testing Status", "QA/QC Sign-off & Date"
    ]
    ws.row_dimensions[4].height = 28
    for col_idx, h in enumerate(headers, 1):
        align = ALIGN_CENTER if col_idx in [1, 2, 3, 4, 7, 9, 10, 14, 18, 19, 20, 21, 22, 23, 24, 25, 26] else ALIGN_LEFT
        set_cell(ws.cell(4, col_idx), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=align)

    row_idx = 5
    for item in inst_list:
        fill = ZEBRA_ODD if (row_idx % 2 == 1) else ZEBRA_EVEN
        t_act = item['active_tag'].upper()
        t_leg = item['tag'].upper()

        matched = tag_to_channel.get(t_act) or tag_to_channel.get(t_leg)
        if not matched:
            norm_act = t_act.replace(' ', '').replace('-', '')
            norm_leg = t_leg.replace(' ', '').replace('-', '')
            matched = tag_to_channel.get(norm_act) or tag_to_channel.get(norm_leg)

        if not matched:
            for k in [t_act, t_leg]:
                if k and any(k.endswith(sfx) for sfx in ['FB', 'SB', 'SF', 'FC']):
                    base_k = re.sub(r'(FB|SB|SF|FC)$', '', k)
                    if base_k in tag_to_channel:
                        matched = tag_to_channel[base_k]
                        break

        first_m = matched[0] if matched else None
        typ_id, strip_type, dwg_ref, typ_desc = classify_cad_typical(item, first_m)

        if matched:
            jb_tag = first_m['dest']
            tb_id = first_m['term_num']
            slot_info = f"{first_m['chassis']}/{first_m['slot']}"
            w_plc = first_m['wire_plc']
            w_term = first_m['wire_term']
            ptag = first_m['plc_tag']
            status_test = "READY FOR TEST"
            status_fill = PENDING_FILL
            status_font = PENDING_FONT
        else:
            jb_tag = item['sig_to'] if item['sig_to'] else "Local Dial"
            tb_id = "-"
            slot_info = "-"
            w_plc = "-"
            w_term = "-"
            ptag = "-"
            if any(w in item['sig_to'] for w in ["AHTR", "HCP", "HTRE", "Deluge", "Suppression", "MCC"]):
                status_test = "PACKAGE LOCAL"
                status_fill = SPARE_FILL
                status_font = SPARE_FONT
            elif "Local" in jb_tag or not item['sig_to']:
                status_test = "MECHANICAL"
                status_fill = ZEBRA_EVEN
                status_font = FONT_DATA_MUTED
            else:
                status_test = "RESERVE ALLOC"
                status_fill = SPARE_FILL
                status_font = SPARE_FONT

        ws.row_dimensions[row_idx].height = 20
        set_cell(ws.cell(row_idx, 1), item['item_no'], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 2), item['active_tag'], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 3), item['tag'] if item['tag'] != item['active_tag'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 4), item['pid'] if item['pid'] else "-", font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 5), item['desc'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 6), item['inst_name'], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 7), item['sig_type'] if item['sig_type'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 8), item['range'] if item['range'] else "-", font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)

        # CAD Typical & Drawing reference columns
        set_cell(ws.cell(row_idx, 9), typ_id, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 10), strip_type, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 11), dwg_ref, font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)

        set_cell(ws.cell(row_idx, 12), jb_tag, font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 13), tb_id, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 14), slot_info, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 15), w_plc, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 16), w_term, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(row_idx, 17), ptag, font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)

        # Pre-configured Checklist Verification Fields
        set_cell(ws.cell(row_idx, 18), "[  ] Pass", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 19), "____ MΩ", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 20), "[  ] 24V", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 21), "____ mA", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 22), "____ mA", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 23), "____ mA", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 24), "[  ] Pass", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)

        set_cell(ws.cell(row_idx, 25), status_test, font=status_font, fill=status_fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(row_idx, 26), "Pending Sign-off", font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_CENTER)

        row_idx += 1

    col_widths = {
        1: 8, 2: 18, 3: 16, 4: 18, 5: 38, 6: 28, 7: 16, 8: 18,
        9: 14, 10: 14, 11: 24, 12: 16, 13: 18, 14: 14, 15: 22, 16: 22, 17: 26,
        18: 14, 19: 14, 20: 12, 21: 14, 22: 14, 23: 14, 24: 16, 25: 18, 26: 22
    }
    for c_i, w in col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w

# -----------------------------------------------------------------------------
# Sheet 5: 04_ITP_Inspection_Plan"""

code = code.replace(old_func_start + code.split(old_func_start)[1].split(old_func_end)[0], new_funcs + "\n")

# Now let's add build_cad_reference_sheet before main
cad_ref_sheet_code = """
# -----------------------------------------------------------------------------
# Sheet 7: 06_CAD_Loop_Wiring_Reference (AutoCAD DWG/DXF Typicals & Wiring Rules)
# -----------------------------------------------------------------------------
def build_cad_reference_sheet(ws):
    ws.views.sheetView[0].showGridLines = True

    # Title Block
    ws.merge_cells("A1:K1")
    set_cell(ws["A1"], "KALASIN ENGINEERING  |  CAD INSTRUMENT LOOP DIAGRAM & PANEL WIRING SPECIFICATION",
             font=FONT_TITLE, fill=NAVY_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[1].height = 25

    ws.merge_cells("A2:K2")
    set_cell(ws["A2"], "Application of LOOP DIAGRAM WRING.dwg & PANEL WIRING DIAGRAM.dwg (DWG & DXF) to Project Sprint 18K TPA",
             font=FONT_SUBTITLE, fill=SUB_FILL, alignment=ALIGN_CENTER)
    ws.row_dimensions[2].height = 18

    # Section 1: CAD Drawing Master Register & Conversion Audit
    ws.merge_cells("A4:K4")
    set_cell(ws["A4"], "1. CAD SOURCE DRAWINGS & DXF CONVERSION MASTER AUDIT", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[4].height = 22

    headers_audit = ["Drawing Filename", "File Format", "File Size", "Modelspace Entities", "Text Annotations", "Block Definitions", "Conversion Tool", "Engineering Application Scope"]
    ws.row_dimensions[5].height = 24
    for c_i, h in enumerate(headers_audit, 1):
        set_cell(ws.cell(5, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    audit_data = [
        ("LOOP DIAGRAM WRING.dwg", "AutoCAD DWG (v2018)", "2.05 MB", "48,670 entities", "18,638 texts", "1,138 blocks", "Native Client Source", "Field-to-Marshalling 3-Zone Loop Typicals (AI, AO, DI, DO, IS)"),
        ("LOOP DIAGRAM WRING.dxf", "AutoCAD DXF (ASCII)", "28.37 MB", "48,670 entities", "18,638 texts", "1,138 blocks", "dwg2dxf (GNU LibreDWG)", "Fully Parsable Open CAD Format / ezdxf Programmatic Automation"),
        ("PANEL WIRING DIAGRAM.dwg", "AutoCAD DWG (v2018)", "4.53 MB", "22,994 entities", "9,462 texts", "1,107 blocks", "Native Client Source", "Marshalling CA1 & RIO-200 Internal Wiring, Fuses, Diode Redundancy"),
        ("PANEL WIRING DIAGRAM.dxf", "AutoCAD DXF (ASCII)", "100.82 MB", "22,994 entities", "9,462 texts", "1,107 blocks", "dwg2dxf (GNU LibreDWG)", "Full High-Resolution Vector Export for Electrical CAD Integration")
    ]
    cur_r = 6
    for row in audit_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 20
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_RIGHT)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), row[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 7), row[6], font=ACTIVE_FONT, fill=ACTIVE_FILL, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 8), row[7], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 2: 3-Zone Architecture & Terminal Numbering Standard
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    set_cell(ws.cell(cur_r, 1), "2. 3-ZONE INSTRUMENT WIRING ARCHITECTURE & TERMINAL DESIGNATIONS", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_arch = ["Physical Zone", "CAD Drawing Block / Column", "Terminal Prefix", "Terminal Model / Spec", "Wire Color Standard", "Core / Pair Tagging", "Shield Treatment Philosophy", "Inspection Checkpoint"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_arch, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    arch_data = [
        ("Zone 1: FIELD", "FIELD column (Transmitter, Valve, Switch)", "Terminals 1(+), 2(-), PE", "M20 / 1/2\" NPT Ex-d/Ex-e Gland", "White (+), Black (-)", "Branch 1Px0.75mm² or 1Px1.5mm²", "CUTBACK & TAPE (Floating, isolated from instrument housing)", "Verify cable gland IP66/67, seal, O-ring, and floating shield"),
        ("Zone 2: JUNCTION BOX", "JUNCTION BOX column (JB-401..618, IS-JB)", "TBAI, TBAO, TBDI, TBDO, TBCI", "Weidmüller WDU 2.5 / Phoenix UK", "White (+), Black (-)", "PR 1..PR 24 (Pair numbers marked)", "SH Feedthrough (Isolated rail, NOT bonded to JB enclosure frame)", "Check terminal torque 0.6Nm, wire ferrules, trunk gland earthing"),
        ("Zone 3: MARSHALLING", "CONTROL ROOM column (Cabinet CA1 / RIO-200)", "TBDAI, TBDAO, TBDDI, TBDDO", "WSI 6-LD (0.5A Fuse) / Disconnect", "White (+), Black (-), Red/Blue", "Multi-pair 24Px0.75mm² Re-2Y(st)-Yv", "IE Instrument Earth (Single-Point Star bonded to Clean Earth Pit)", "Verify fuse rating (0.5A fast), disconnect knife closed, IE <1.0Ω"),
        ("Zone 3: PLC I/O RACK", "1756 ControlLogix / 1794 FLEX I/O", "RTB Terminal (Chassis C1..C5)", "1756-TBCH / 1794-TB3 Removable TB", "Grey (DC+), Blue (DC-), White (Sig)", "Pre-formed wiring harness to CA1", "Internal chassis backplane isolated ground", "Verify card LED active, channel OK, tag scaled in ControlLogix")
    ]
    for row in arch_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 24
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), row[5], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 7), row[6], font=SPARE_FONT, fill=SPARE_FILL, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 8), row[7], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 3: 6 Master CAD Typical Configurations Mapped to Project Loops
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    set_cell(ws.cell(cur_r, 1), "3. MASTER CAD TYPICAL LOOP WIRING CONFIGURATIONS & COMMISSIONING PROCEDURES", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_typ = ["CAD Typical ID", "Instrument Loop Category", "Sample Project Tags", "Terminal Strip", "Signal Standard", "Power Source", "Commissioning Cold Check", "Commissioning Hot / Calibration Check"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_typ, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    typ_data = [
        ("TYP-AI-01", "2-Wire Loop Powered Analog Transmitter", "PT-10080, LT-10081, TT-10085, PDT-60101", "TBAI.xx", "4-20mA DC (2-Wire)", "24VDC from CA1 Marshalling", "Megger core-to-earth >20MΩ @ 500VDC, continuity <2.0Ω", "Inject 4, 8, 12, 16, 20 mA via Fluke 789 at field terminals; verify ControlLogix tag scaling ±0.2%"),
        ("TYP-AI-02", "4-Wire Active Powered Analog Transmitter", "FT-102900, FT-102901, FT-106003, DT-106004", "TBAI.xx", "4-20mA DC (Active Output)", "220VAC / 24VDC Aux Power", "Megger power and signal cores separately >20MΩ", "Power up instrument; measure 24VDC/220VAC; simulate 0-100% flow from local transmitter menu; check SCADA"),
        ("TYP-AO-01", "4-20mA SMART Control Valve Positioner", "FCV-102900-1, FCV-102901-1, FCV-106003-1", "TBAO.xx", "4-20mA Command + Travel FB", "4-20mA loop powered positioner", "Loop resistance check (typ. 250-450Ω), air supply 5.5 bar", "Force 4, 8, 12, 16, 20 mA from ControlLogix; verify 0, 25, 50, 75, 100% physical stroke and limit switches"),
        ("TYP-DI-01", "Discrete Status / Limit / Level Switch", "LSH-60101, LSL-60102, ZSO-102900, ZSC-102900", "TBDI.xx", "24VDC Interrogation Contact", "24VDC Wetting Voltage from CA1", "Contact open = ∞ MΩ, contact closed = <2.0Ω", "Verify 24VDC at field contact; actuate switch manually; verify PLC card LED on/off and HMI alarm latching"),
        ("TYP-DO-01", "Solenoid Valve / Interposing Relay Drive", "XV-102900, SOV-102901, SV-60101, PMP-40101", "TBDO.xx", "24VDC Switched Output", "24VDC from 1756-OB32 / Omron Relay", "Coil resistance check (typ. 30-120Ω), diode polarity check", "Command DO bit from ControlLogix; verify interposing relay pull-in, 24VDC at solenoid, pneumatic actuation"),
        ("TYP-IS-01", "Intrinsically Safe Ex Loop (Hazardous Area)", "Loops landed in IS-JB-603, 608, 612, 618", "TB-IS.xx (Blue)", "Ex-ia / Ex-ib Certified Loop", "Galvanic Isolator (MTL / P+F) in CA1", "Verify blue trunk cable, isolated shield, clearance >50mm", "Measure barrier output Voc, Isc; verify entity parameter matching (Ca > C_cable + C_inst, La > L_cable + L_inst)")
    ]
    for row in typ_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 26
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA_CODE, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA, fill=fill, alignment=ALIGN_CENTER)
        set_cell(ws.cell(cur_r, 6), row[5], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 7), row[6], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 8), row[7], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    cur_r += 1
    # Section 4: Panel Wiring & 24VDC Power Distribution Standard (from PANEL WIRING DIAGRAM.dwg)
    ws.merge_cells(f"A{cur_r}:K{cur_r}")
    set_cell(ws.cell(cur_r, 1), "4. MARSHALLING CABINET (CA1) INTERNAL ARRANGEMENT & 24VDC POWER SYSTEM", font=FONT_SECTION, fill=SECTION_FILL, alignment=ALIGN_LEFT)
    ws.row_dimensions[cur_r].height = 22

    cur_r += 1
    hdrs_pnl = ["Subsystem Component", "Specification / Model in DWG", "Location in CA1", "Redundancy / Protection Mechanism", "Testing & Verification Requirement"]
    ws.row_dimensions[cur_r].height = 24
    for c_i, h in enumerate(hdrs_pnl, 1):
        set_cell(ws.cell(cur_r, c_i), h, font=FONT_HEADER, fill=NAVY_FILL, border=HEADER_BORDER, alignment=ALIGN_CENTER)

    cur_r += 1
    pnl_data = [
        ("Primary Power Supplies (PSU1 & PSU2)", "Phoenix Contact Quint Power 24VDC / 40A", "Cabinet Base DIN Rail (Left & Right)", "N+1 Active Redundancy via Diode Decoupling Module", "Simulate AC supply failure on PSU1; verify seamless 24VDC transfer without PLC dip"),
        ("Branch Circuit Fuses", "Weidmüller WSI 6-LD / WDU 2.5 with 0.5A fast fuse", "Marshalling Terminal Rails TBX204..TBX207", "Individual channel fuse isolation with blown-fuse LED", "Measure loop voltage drop across fuse holder (<0.1VDC); verify fuse LED indicator"),
        ("Dual Earthing Busbars", "PE (Protective Earth) & IE (Instrument Earth)", "Cabinet Bottom Isolated Copper Bars (PE=Direct, IE=Isolated)", "Single-Point Grounding architecture prevents ground loop hum", "Measure resistance between IE bar and clean earth pit (<1.0Ω); check IE-to-PE isolation (>10MΩ)"),
        ("Internal Wire Ducting", "Slotted PVC Trunking 40x100mm & 70x100mm", "Vertical & Horizontal Cabinet Wireways", "Physical segregation between 24VDC signal and 220VAC power", "Inspect 50mm separation between IS blue ducts and non-IS grey ducts; 50% fill factor"),
        ("PLC Remote I/O Adapters", "1794-AENTR / 1756-EN2TR Dual Port EtherNet/IP", "Chassis Mounting Plates", "Device Level Ring (DLR) fault-tolerant network topology", "Perform network ring break test; verify zero packet loss to central ControlLogix processor")
    ]
    for row in pnl_data:
        fill = ZEBRA_ODD if (cur_r % 2 == 1) else ZEBRA_EVEN
        ws.row_dimensions[cur_r].height = 24
        set_cell(ws.cell(cur_r, 1), row[0], font=FONT_DATA_BOLD, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 2), row[1], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 3), row[2], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 4), row[3], font=FONT_DATA, fill=fill, alignment=ALIGN_LEFT)
        set_cell(ws.cell(cur_r, 5), row[4], font=FONT_DATA_MUTED, fill=fill, alignment=ALIGN_LEFT)
        cur_r += 1

    col_w = {1: 24, 2: 30, 3: 28, 4: 28, 5: 32, 6: 24, 7: 34, 8: 42}
    for c_i, w in col_w.items():
        ws.column_dimensions[get_column_letter(c_i)].width = w
"""

# Insert before main
code = code.replace("# -----------------------------------------------------------------------------\n# Main Execution", cad_ref_sheet_code + "\n# -----------------------------------------------------------------------------\n# Main Execution")

# Update main to create sheet 06_CAD_Loop_Wiring_Reference and update auto_filter ref to Z
code = code.replace("ws_list.auto_filter.ref = f\"A4:W{ws_list.max_row}\"", "ws_list.auto_filter.ref = f\"A4:Z{ws_list.max_row}\"")

old_main_sheet6 = """    # 6. Field Loop Check Certificate Template Sheet
    print("Generating Sheet: 05_Loop_Test_Form_Template...")
    ws_form = wb.create_sheet(title="05_Loop_Test_Form_Template")
    build_form_template_sheet(ws_form)
    ws_form.freeze_panes = 'A4'"""

new_main_sheets = """    # 6. Field Loop Check Certificate Template Sheet
    print("Generating Sheet: 05_Loop_Test_Form_Template...")
    ws_form = wb.create_sheet(title="05_Loop_Test_Form_Template")
    build_form_template_sheet(ws_form)
    ws_form.freeze_panes = 'A4'

    # 7. CAD Loop Diagram & Panel Wiring Specification Sheet
    print("Generating Sheet: 06_CAD_Loop_Wiring_Reference...")
    ws_cad = wb.create_sheet(title="06_CAD_Loop_Wiring_Reference")
    build_cad_reference_sheet(ws_cad)
    ws_cad.freeze_panes = 'A5'"""

code = code.replace(old_main_sheet6, new_main_sheets)

with open(src_path, "w", encoding="utf-8") as f:
    f.write(code)

print("build_loop_test_plan_and_schedule.py updated successfully!")
