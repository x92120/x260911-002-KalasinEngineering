#!/usr/bin/env python3
"""
Generate Industrial Vector Drawing & Companion PDF/PNG for:
Rockwell Automation ControlLogix 5580 (1756-L950TPSXT)
5-Chassis 13-Slot DLR Network Ring Cable Design & Rack Slot Redesign (LIGHT STYLE)
Facility: Ingredion Kalasin Spray Dryer Plant
Theme: Clean Industrial Light Style (Suitable for high-contrast engineering prints)
"""

import os
import fitz

SVG_WIDTH = 2500
SVG_HEIGHT = 1650

def generate_svg():
    svg_elements = []

    # Defs: Gradients, Filters, Markers, Patterns (Light Style)
    svg_elements.append(f'''
    <defs>
        <!-- Gradients for Cards -->
        <linearGradient id="cpuGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#1e40af" />
            <stop offset="60%" stop-color="#1d4ed8" />
            <stop offset="100%" stop-color="#172554" />
        </linearGradient>

        <linearGradient id="commGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#0e7490" />
            <stop offset="60%" stop-color="#0891b2" />
            <stop offset="100%" stop-color="#164e63" />
        </linearGradient>

        <linearGradient id="diGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#065f46" />
            <stop offset="60%" stop-color="#059669" />
            <stop offset="100%" stop-color="#064e3b" />
        </linearGradient>

        <linearGradient id="doGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#9a3412" />
            <stop offset="60%" stop-color="#ea580c" />
            <stop offset="100%" stop-color="#7c2d12" />
        </linearGradient>

        <linearGradient id="aiGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#581c87" />
            <stop offset="60%" stop-color="#7e22ce" />
            <stop offset="100%" stop-color="#3b0764" />
        </linearGradient>

        <linearGradient id="aoGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#9f1239" />
            <stop offset="60%" stop-color="#e11d48" />
            <stop offset="100%" stop-color="#881337" />
        </linearGradient>

        <linearGradient id="emptyGrad" x1="0%" y1="0%" x2="0%" y2="100%">
            <stop offset="0%" stop-color="#f1f5f9" />
            <stop offset="100%" stop-color="#e2e8f0" />
        </linearGradient>

        <!-- Drop Shadows -->
        <filter id="shadow" x="-5%" y="-5%" width="110%" height="115%">
            <feDropShadow dx="0" dy="4" stdDeviation="5" flood-color="#64748b" flood-opacity="0.16"/>
        </filter>

        <!-- Arrow Markers -->
        <marker id="arrowCopper" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto">
            <path d="M 0 1 L 7 4 L 0 7 z" fill="#059669" />
        </marker>
        <marker id="arrowFiber" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto">
            <path d="M 0 1 L 7 4 L 0 7 z" fill="#d97706" />
        </marker>
        <marker id="arrowUplink" markerWidth="8" markerHeight="8" refX="6" refY="4" orient="auto">
            <path d="M 0 1 L 7 4 L 0 7 z" fill="#0284c7" />
        </marker>

        <!-- Patterns -->
        <pattern id="diagonalHatch" width="8" height="8" patternTransform="rotate(45 0 0)" patternUnits="userSpaceOnUse">
            <line x1="0" y1="0" x2="0" y2="8" stroke="#cbd5e1" stroke-width="1.5" />
        </pattern>
        <pattern id="lightGrid" width="40" height="40" patternUnits="userSpaceOnUse">
            <path d="M 40 0 L 0 0 0 40" fill="none" stroke="#f1f5f9" stroke-width="1"/>
        </pattern>
    </defs>
    ''')

    # Background canvas (Crisp White + Light Technical Grid)
    svg_elements.append(f'<rect width="{SVG_WIDTH}" height="{SVG_HEIGHT}" fill="#ffffff" />')
    svg_elements.append(f'<rect width="{SVG_WIDTH}" height="{SVG_HEIGHT}" fill="url(#lightGrid)" />')

    # ==========================================
    # 1. HEADER & TITLE BLOCK
    # ==========================================
    header_svg = f'''
    <g id="Header_TitleBlock">
        <rect x="30" y="25" width="{SVG_WIDTH - 60}" height="95" rx="8" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" filter="url(#shadow)" />
        
        <text x="60" y="60" font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" font-size="22" font-weight="800" fill="#0f172a" letter-spacing="1.0">INGREDION THAILAND — SPRINT 18K TPA SPRAY DRYER PLANT (KALASIN)</text>
        <text x="60" y="90" font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" font-size="14" font-weight="700" fill="#0284c7" letter-spacing="0.5">ROCKWELL CONTROLLOGIX 5580 (1756-L950TPSXT) — 5-CHASSIS DLR NETWORK CABLE &amp; 13-SLOT REDESIGN</text>

        <!-- Technical Badges -->
        <g transform="translate({SVG_WIDTH - 760}, 45)">
            <rect x="0" y="0" width="130" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="65" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">DRAWING NO.</text>
            <text x="65" y="36" text-anchor="middle" font-family="sans-serif" font-size="12" fill="#0284c7" font-weight="800">DWG-260911-NET-01</text>
        </g>
        <g transform="translate({SVG_WIDTH - 610}, 45)">
            <rect x="0" y="0" width="70" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="35" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">REVISION</text>
            <text x="35" y="36" text-anchor="middle" font-family="sans-serif" font-size="12" fill="#059669" font-weight="800">REV B02</text>
        </g>
        <g transform="translate({SVG_WIDTH - 520}, 45)">
            <rect x="0" y="0" width="100" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="50" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">TOPOLOGY</text>
            <text x="50" y="36" text-anchor="middle" font-family="sans-serif" font-size="12" fill="#d97706" font-weight="800">ODVA DLR RING</text>
        </g>
        <g transform="translate({SVG_WIDTH - 400}, 45)">
            <rect x="0" y="0" width="150" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="75" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">CONTROLLER MODEL</text>
            <text x="75" y="36" text-anchor="middle" font-family="sans-serif" font-size="11" fill="#db2777" font-weight="800">1756-L950TPSXT (XT)</text>
        </g>
        <g transform="translate({SVG_WIDTH - 230}, 45)">
            <rect x="0" y="0" width="170" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="85" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">CHASSIS SPEC</text>
            <text x="85" y="36" text-anchor="middle" font-family="sans-serif" font-size="11" fill="#0f172a" font-weight="800">5 CHASSIS x 13 SLOTS</text>
        </g>
    </g>
    '''
    svg_elements.append(header_svg)

    # ==========================================
    # 2. SECTION 1: TOPOLOGY & CABLE DESIGN (Upper)
    # ==========================================
    mcp_x, mcp_y, mcp_w, mcp_h = 40, 140, 1680, 390
    rio_x, rio_y, rio_w, rio_h = 1760, 140, 700, 390

    top_section_svg = f'''
    <g id="Section1_Title">
        <text x="45" y="132" font-family="sans-serif" font-size="14" font-weight="800" fill="#334155" letter-spacing="1">PART 1: INDUSTRIAL ETHERNET DLR NETWORK ARCHITECTURE &amp; CABLE ROUTING</text>
    </g>

    <!-- MAIN CONTROL PANEL (MCP-01) BOUNDARY -->
    <g id="Panel_MCP">
        <rect x="{mcp_x}" y="{mcp_y}" width="{mcp_w}" height="{mcp_h}" rx="8" fill="#f0f9ff" stroke="#0284c7" stroke-width="1.8" stroke-dasharray="6,4" filter="url(#shadow)" />
        <rect x="{mcp_x}" y="{mcp_y}" width="380" height="32" rx="6" fill="#0284c7" />
        <text x="{mcp_x + 15}" y="{mcp_y + 21}" font-family="sans-serif" font-size="13" font-weight="700" fill="#ffffff">MAIN CONTROL PANEL (MCP-01) — SAME ENCLOSURE (CS1 - CS4)</text>
        <text x="{mcp_x + mcp_w - 240}" y="{mcp_y + 21}" font-family="monospace" font-size="11" font-weight="700" fill="#0369a1">MEDIA: INDUSTRIAL CAT6A S/FTP</text>
    </g>

    <!-- RIO200 REMOTE PANEL BOUNDARY -->
    <g id="Panel_RIO200">
        <rect x="{rio_x}" y="{rio_y}" width="{rio_w}" height="{rio_h}" rx="8" fill="#f0fdfa" stroke="#0d9488" stroke-width="1.8" stroke-dasharray="6,4" filter="url(#shadow)" />
        <rect x="{rio_x}" y="{rio_y}" width="360" height="32" rx="6" fill="#0d9488" />
        <text x="{rio_x + 15}" y="{rio_y + 21}" font-family="sans-serif" font-size="13" font-weight="700" fill="#ffffff">REMOTE I/O PANEL (RIO200) — PRE-SLURRY AREA</text>
        <text x="{rio_x + rio_w - 200}" y="{rio_y + 21}" font-family="monospace" font-size="11" font-weight="700" fill="#0f766e">MEDIA: FIBER OPTIC OM3/OS2</text>
    </g>
    '''
    svg_elements.append(top_section_svg)

    # Function to draw a network device node in Part 1 (Light Style)
    def draw_top_node(x, y, w, h, tag, title, model, role, ip, ports, is_sup=False):
        border_col = "#eab308" if is_sup else "#cbd5e1"
        badge_svg = ""
        if is_sup:
            badge_svg = f'''
            <rect x="{x + w - 120}" y="{y + 6}" width="112" height="18" rx="9" fill="#ca8a04" />
            <text x="{x + w - 64}" y="{y + 19}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#ffffff">★ DLR SUPERVISOR</text>
            '''
        else:
            badge_svg = f'''
            <rect x="{x + w - 85}" y="{y + 6}" width="78" height="18" rx="9" fill="#0f766e" />
            <text x="{x + w - 46}" y="{y + 19}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#ffffff">DLR NODE</text>
            '''

        # Ports
        ports_svg = []
        for p in ports:
            px, py, pname, pcol, ptype = p['x'], p['y'], p['name'], p['color'], p['type']
            ports_svg.append(f'''
            <circle cx="{px}" cy="{py}" r="7" fill="{pcol}" stroke="#ffffff" stroke-width="1.5" />
            <text x="{px}" y="{py - 11}" text-anchor="middle" font-family="monospace" font-size="9" font-weight="700" fill="#0f172a">{pname}</text>
            <text x="{px}" y="{py + 19}" text-anchor="middle" font-family="monospace" font-size="8" font-weight="600" fill="#64748b">{ptype}</text>
            ''')

        safe_role = role.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_title = title.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
        safe_model = model.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        return f'''
        <g id="TopNode_{tag}">
            <rect x="{x}" y="{y}" width="{w}" height="{h}" rx="6" fill="#ffffff" stroke="{border_col}" stroke-width="{2 if is_sup else 1.5}" filter="url(#shadow)"/>
            <rect x="{x}" y="{y}" width="{w}" height="28" rx="6" fill="#334155" />
            <text x="{x + 12}" y="{y + 19}" font-family="sans-serif" font-size="12" font-weight="800" fill="#ffffff">{tag} — {safe_title}</text>
            {badge_svg}

            <text x="{x + 12}" y="{y + 48}" font-family="monospace" font-size="11" font-weight="700" fill="#0284c7">{safe_model}</text>
            <text x="{x + 12}" y="{y + 66}" font-family="sans-serif" font-size="10" fill="#475569">{safe_role}</text>
            <text x="{x + 12}" y="{y + 84}" font-family="monospace" font-size="10" font-weight="700" fill="#059669">IP: {ip}</text>

            {''.join(ports_svg)}
        </g>
        '''

    # Nodes positioning in Part 1
    cs1_ports = [
        {'x': 100, 'y': 330, 'name': 'P1 (Ring In/Out)', 'color': '#059669', 'type': 'RJ45'},
        {'x': 190, 'y': 330, 'name': 'P2 (Beacon)', 'color': '#059669', 'type': 'RJ45'},
        {'x': 280, 'y': 330, 'name': 'UPLINK (S2)', 'color': '#0284c7', 'type': 'RJ45/Fiber'}
    ]
    svg_elements.append(draw_top_node(65, 200, 270, 160, "CS1", "MASTER CONTROLLER", "1756-L950TPSXT + EN4TR", "Active Ring Supervisor & Master CPU", "192.168.30.11", cs1_ports, is_sup=True))

    cs2_ports = [
        {'x': 430, 'y': 330, 'name': 'P1 (From CS1)', 'color': '#059669', 'type': 'RJ45'},
        {'x': 520, 'y': 330, 'name': 'P2 (To CS3)', 'color': '#059669', 'type': 'RJ45'}
    ]
    svg_elements.append(draw_top_node(375, 200, 210, 160, "CS2", "EXPANSION RACK 1", "1756-EN4TR (Slot 0)", "Remote I/O Adapter Node 1", "192.168.30.12", cs2_ports))

    cs3_ports = [
        {'x': 680, 'y': 330, 'name': 'P1 (From CS2)', 'color': '#059669', 'type': 'RJ45'},
        {'x': 770, 'y': 330, 'name': 'P2 (To CS4)', 'color': '#059669', 'type': 'RJ45'}
    ]
    svg_elements.append(draw_top_node(625, 200, 210, 160, "CS3", "EXPANSION RACK 2", "1756-EN4TR (Slot 0)", "Remote I/O Adapter Node 2", "192.168.30.13", cs3_ports))

    cs4_ports = [
        {'x': 930, 'y': 330, 'name': 'P1 (From CS3)', 'color': '#059669', 'type': 'RJ45'},
        {'x': 1020, 'y': 330, 'name': 'P2 (To MC-01)', 'color': '#059669', 'type': 'RJ45'}
    ]
    svg_elements.append(draw_top_node(875, 200, 210, 160, "CS4", "EXPANSION RACK 3", "1756-EN4TR (Slot 0)", "Remote I/O Adapter Node 3", "192.168.30.14", cs4_ports))

    mc1_ports = [
        {'x': 1180, 'y': 330, 'name': 'RJ45-1 (From CS4)', 'color': '#059669', 'type': 'Cat6A'},
        {'x': 1260, 'y': 330, 'name': 'RJ45-2 (To CS1)', 'color': '#059669', 'type': 'Cat6A'},
        {'x': 1370, 'y': 330, 'name': 'FO-1 (TX/RX Out)', 'color': '#d97706', 'type': 'LC Duplex'},
        {'x': 1460, 'y': 330, 'name': 'FO-2 (TX/RX In)', 'color': '#d97706', 'type': 'LC Duplex'}
    ]
    svg_elements.append(draw_top_node(1135, 200, 365, 160, "MC-01", "DLR FIBER TAP / SWITCH", "1783-ETAP2F / STRATIX", "Copper-to-Fiber DLR Transceiver (MCP)", "192.168.30.250", mc1_ports))

    # Inter-panel Fiber Optic Cable Conduit
    svg_elements.append(f'''
    <g id="Fiber_Trunk_Zone">
        <rect x="1530" y="220" width="200" height="120" rx="6" fill="#fffdf5" stroke="#f59e0b" stroke-width="1.5" stroke-dasharray="4,3" />
        <text x="1630" y="245" text-anchor="middle" font-family="sans-serif" font-size="11" font-weight="800" fill="#b45309">FIBER TRUNK CABLEWAY</text>
        <text x="1630" y="265" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#d97706">4-CORE ARMORED OM3/OS2</text>
        <text x="1630" y="285" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b">Length: ~150 - 300 m</text>
        <text x="1630" y="305" text-anchor="middle" font-family="monospace" font-size="9" font-weight="700" fill="#059669">TAG: FOB-RIO-01 / 02</text>
    </g>
    ''')

    mc2_ports = [
        {'x': 1820, 'y': 330, 'name': 'FO-1 (From MC1)', 'color': '#d97706', 'type': 'LC Duplex'},
        {'x': 1910, 'y': 330, 'name': 'FO-2 (To MC1)', 'color': '#d97706', 'type': 'LC Duplex'},
        {'x': 2010, 'y': 330, 'name': 'RJ45-1 (To CS5)', 'color': '#059669', 'type': 'Cat6A'},
        {'x': 2090, 'y': 330, 'name': 'RJ45-2 (From CS5)', 'color': '#059669', 'type': 'Cat6A'}
    ]
    svg_elements.append(draw_top_node(1780, 200, 350, 160, "MC-02", "DLR FIBER TAP / SWITCH", "1783-ETAP2F / STRATIX", "Fiber-to-Copper DLR Transceiver (RIO)", "192.168.30.251", mc2_ports))

    cs5_ports = [
        {'x': 2220, 'y': 330, 'name': 'P1 (From MC-02)', 'color': '#059669', 'type': 'RJ45'},
        {'x': 2310, 'y': 330, 'name': 'P2 (To MC-02)', 'color': '#059669', 'type': 'RJ45'}
    ]
    svg_elements.append(draw_top_node(2170, 200, 260, 160, "CS5", "REMOTE EXPANSION RACK", "1756-EN4TR (Slot 0)", "Remote I/O Adapter Node 4", "192.168.30.15", cs5_ports))

    # Cable Connection Lines
    cables_svg = f'''
    <g id="Cables_Routing" stroke-linecap="round">
        <!-- CBL-DLR-01: CS1 P1 -> CS2 P1 -->
        <path d="M 100 337 L 100 420 L 430 420 L 430 337" fill="none" stroke="#059669" stroke-width="3" marker-end="url(#arrowCopper)" />
        <rect x="230" y="410" width="100" height="20" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="280" y="424" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#065f46">CBL-DLR-01</text>

        <!-- CBL-DLR-02: CS2 P2 -> CS3 P1 -->
        <path d="M 520 337 L 520 440 L 680 440 L 680 337" fill="none" stroke="#059669" stroke-width="3" marker-end="url(#arrowCopper)" />
        <rect x="560" y="430" width="100" height="20" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="610" y="444" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#065f46">CBL-DLR-02</text>

        <!-- CBL-DLR-03: CS3 P2 -> CS4 P1 -->
        <path d="M 770 337 L 770 420 L 930 420 L 930 337" fill="none" stroke="#059669" stroke-width="3" marker-end="url(#arrowCopper)" />
        <rect x="815" y="410" width="100" height="20" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="865" y="424" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#065f46">CBL-DLR-03</text>

        <!-- CBL-DLR-04: CS4 P2 -> MC-01 RJ45-1 -->
        <path d="M 1020 337 L 1020 440 L 1180 440 L 1180 337" fill="none" stroke="#059669" stroke-width="3" marker-end="url(#arrowCopper)" />
        <rect x="1060" y="430" width="100" height="20" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="1110" y="444" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#065f46">CBL-DLR-04</text>

        <!-- FOB-RIO-01 (FIBER OUTBOUND): MC-01 FO-1 -> MC-02 FO-1 -->
        <path d="M 1370 337 L 1370 400 L 1820 400 L 1820 337" fill="none" stroke="#d97706" stroke-width="3.5" stroke-dasharray="8,5" marker-end="url(#arrowFiber)" />
        <rect x="1550" y="388" width="120" height="22" rx="4" fill="#fef3c7" stroke="#d97706" />
        <text x="1610" y="403" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#92400e">FOB-RIO-01 (Core 1-2)</text>

        <!-- CBL-DLR-05: MC-02 RJ45-1 -> CS5 P1 -->
        <path d="M 2010 337 L 2010 420 L 2220 420 L 2220 337" fill="none" stroke="#059669" stroke-width="3" marker-end="url(#arrowCopper)" />
        <rect x="2075" y="410" width="100" height="20" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="2125" y="424" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#065f46">CBL-DLR-05</text>

        <!-- CBL-DLR-06: CS5 P2 -> MC-02 RJ45-2 -->
        <path d="M 2310 337 L 2310 450 L 2090 450 L 2090 337" fill="none" stroke="#059669" stroke-width="3" marker-end="url(#arrowCopper)" />
        <rect x="2160" y="440" width="100" height="20" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="2210" y="454" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#065f46">CBL-DLR-06</text>

        <!-- FOB-RIO-02 (FIBER RETURN): MC-02 FO-2 -> MC-01 FO-2 -->
        <path d="M 1910 337 L 1910 480 L 1460 480 L 1460 337" fill="none" stroke="#d97706" stroke-width="3.5" stroke-dasharray="8,5" marker-end="url(#arrowFiber)" />
        <rect x="1640" y="468" width="120" height="22" rx="4" fill="#fef3c7" stroke="#d97706" />
        <text x="1700" y="483" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#92400e">FOB-RIO-02 (Core 3-4)</text>

        <!-- CBL-DLR-07 (CLOSING RING TO SUPERVISOR BEACON PORT): MC-01 RJ45-2 -> CS1 P2 -->
        <path d="M 1260 337 L 1260 500 L 190 500 L 190 337" fill="none" stroke="#059669" stroke-width="3.5" marker-end="url(#arrowCopper)" />
        <rect x="670" y="490" width="220" height="22" rx="4" fill="#ecfdf5" stroke="#059669" />
        <text x="780" y="505" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#065f46">CBL-DLR-07 (CLOSES DLR RING TO P2)</text>

        <!-- CBL-UPL-01: CS1 S02 Uplink to Plant Core SCADA Network -->
        <path d="M 280 337 L 280 180 L 320 180" fill="none" stroke="#0284c7" stroke-width="2.5" marker-end="url(#arrowUplink)" />
        <rect x="330" y="168" width="220" height="24" rx="4" fill="#e0f2fe" stroke="#0284c7" />
        <text x="440" y="184" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#0369a1">CBL-UPL-01: TO PLANT SCADA (VLAN 30)</text>
    </g>
    '''
    svg_elements.append(cables_svg)

    # ==========================================
    # 3. SECTION 2: 13-SLOT CHASSIS REDESIGN (Lower)
    # ==========================================
    racks_header_svg = f'''
    <g id="Section2_Title">
        <text x="45" y="565" font-family="sans-serif" font-size="14" font-weight="800" fill="#334155" letter-spacing="1">PART 2: REDESIGNED 13-SLOT CHASSIS SLOT DEFINITIONS &amp; HARDWARE ELEVATION</text>
        <text x="960" y="565" font-family="sans-serif" font-size="12" font-weight="700" fill="#059669">STANDARDIZED BALANCED I/O ARCHITECTURE (4 DI / 2 DO / 2 AI / 2 AO / 2 SPARE PER RACK)</text>
    </g>
    '''
    svg_elements.append(racks_header_svg)

    rack_data = [
        {
            'id': 'CS1',
            'title': 'CHASSIS 1 — MASTER CONTROLLER (MCP-01)',
            'loc': 'Main Control Panel / Rack 1',
            'net_role': 'Active DLR Supervisor + Supervisory Master',
            'slots': [
                {'slot': 0, 'card': '1756-L950TPSXT', 'name': 'CPU', 'desc': 'GuardLogix Process Safety Controller (XT)', 'sig': 'CPU', 'grad': 'cpuGrad', 'txtCol': '#93c5fd'},
                {'slot': 1, 'card': '1756-EN4TR', 'name': 'Eth_Module', 'desc': 'DLR Ring Supervisor (1Gbps Dual Port)', 'sig': 'DLR COMM', 'grad': 'commGrad', 'txtCol': '#67e8f9'},
                {'slot': 2, 'card': '1756-EN4TR', 'name': 'Eth_Module', 'desc': 'Plant SCADA / FactoryTalk Uplink', 'sig': 'UPLINK', 'grad': 'commGrad', 'txtCol': '#67e8f9'},
                {'slot': 3, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Sink/Source Digital Input (DI-01)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 4, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Sink/Source Digital Input (DI-02)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 5, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Sink/Source Digital Input (DI-03)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 6, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Sink/Source Digital Input (DI-04)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 7, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Sourcing Digital Output (DO-01)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 8, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Sourcing Digital Output (DO-02)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 9, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch High-Speed Analog Input (AI-01)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 10, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch High-Speed Analog Input (AI-02)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 11, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Voltage/Current Analog Output (AO-01)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 12, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Voltage/Current Analog Output (AO-02)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'}
            ]
        },
        {
            'id': 'CS2',
            'title': 'CHASSIS 2 — EXPANSION RACK 1 (MCP-01)',
            'loc': 'Main Control Panel / Rack 2',
            'net_role': 'DLR Ring Node 1',
            'slots': [
                {'slot': 0, 'card': '1756-EN4TR', 'name': 'Eth_Module', 'desc': 'Remote I/O Adapter (Slot 0 Redesigned)', 'sig': 'DLR COMM', 'grad': 'commGrad', 'txtCol': '#67e8f9'},
                {'slot': 1, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-05)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 2, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-06)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 3, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-07)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 4, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-08)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 5, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-03)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 6, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-04)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 7, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-03)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 8, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-04)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 9, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-03)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 10, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-04)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 11, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'},
                {'slot': 12, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'}
            ]
        },
        {
            'id': 'CS3',
            'title': 'CHASSIS 3 — EXPANSION RACK 2 (MCP-01)',
            'loc': 'Main Control Panel / Rack 3',
            'net_role': 'DLR Ring Node 2',
            'slots': [
                {'slot': 0, 'card': '1756-EN4TR', 'name': 'Eth_Module', 'desc': 'Remote I/O Adapter (Slot 0 Redesigned)', 'sig': 'DLR COMM', 'grad': 'commGrad', 'txtCol': '#67e8f9'},
                {'slot': 1, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-09)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 2, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-10)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 3, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-11)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 4, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-12)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 5, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-05)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 6, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-06)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 7, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-05)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 8, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-06)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 9, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-05)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 10, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-06)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 11, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'},
                {'slot': 12, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'}
            ]
        },
        {
            'id': 'CS4',
            'title': 'CHASSIS 4 — EXPANSION RACK 3 (MCP-01)',
            'loc': 'Main Control Panel / Rack 4',
            'net_role': 'DLR Ring Node 3',
            'slots': [
                {'slot': 0, 'card': '1756-EN4TR', 'name': 'Eth_Module', 'desc': 'Remote I/O Adapter (Slot 0 Redesigned)', 'sig': 'DLR COMM', 'grad': 'commGrad', 'txtCol': '#67e8f9'},
                {'slot': 1, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-13)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 2, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-14)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 3, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-15)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 4, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-16)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 5, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-07)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 6, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-08)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 7, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-07)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 8, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-08)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 9, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-07)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 10, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-08)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 11, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'},
                {'slot': 12, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'}
            ]
        },
        {
            'id': 'CS5',
            'title': 'CHASSIS 5 — REMOTE I/O RACK (RIO200 PANEL)',
            'loc': 'Remote I/O Panel RIO200 (Area 200) via Fiber',
            'net_role': 'DLR Ring Node 4 (Fiber Link)',
            'slots': [
                {'slot': 0, 'card': '1756-EN4TR', 'name': 'Eth_Module', 'desc': 'Remote I/O Adapter (Slot 0 Redesigned)', 'sig': 'DLR COMM', 'grad': 'commGrad', 'txtCol': '#67e8f9'},
                {'slot': 1, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-17)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 2, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-18)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 3, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-19)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 4, 'card': '1756-IB32', 'name': '32DI', 'desc': '24VDC Digital Input (DI-20)', 'sig': 'IN', 'grad': 'diGrad', 'txtCol': '#86efac'},
                {'slot': 5, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-09)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 6, 'card': '1756-OB32', 'name': '32DO', 'desc': '24VDC Digital Output (DO-10)', 'sig': 'OUT', 'grad': 'doGrad', 'txtCol': '#fdba74'},
                {'slot': 7, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-09)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 8, 'card': '1756-IF16', 'name': '16AI', 'desc': '16-Ch Analog Input (AI-10)', 'sig': 'IN', 'grad': 'aiGrad', 'txtCol': '#d8b4fe'},
                {'slot': 9, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-09)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 10, 'card': '1756-OF8', 'name': '8AO', 'desc': '8-Ch Analog Output (AO-10)', 'sig': 'IOUT', 'grad': 'aoGrad', 'txtCol': '#fda4af'},
                {'slot': 11, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'},
                {'slot': 12, 'card': '1756-N2', 'name': 'Empty', 'desc': 'Slot Filler / Spare Expansion', 'sig': 'EMPTY', 'grad': 'emptyGrad', 'txtCol': '#64748b'}
            ]
        }
    ]

    rack_start_y = 580
    rack_w = 1780
    slot_w = 130
    slot_gap = 5
    rack_h = 165
    rack_step_y = 180

    for r_idx, r in enumerate(rack_data):
        curr_y = rack_start_y + r_idx * rack_step_y
        border_c = "#0284c7" if r['id'] == 'CS1' else ("#0d9488" if r['id'] == 'CS5' else "#cbd5e1")

        rack_svg = f'''
        <g id="RackElevation_{r['id']}">
            <!-- Rack Chassis Frame -->
            <rect x="40" y="{curr_y}" width="{rack_w}" height="{rack_h}" rx="6" fill="#f8fafc" stroke="{border_c}" stroke-width="1.8" filter="url(#shadow)" />
            
            <!-- Rack Info Banner -->
            <rect x="40" y="{curr_y}" width="{rack_w}" height="26" rx="6" fill="#334155" />
            <text x="55" y="{curr_y + 18}" font-family="sans-serif" font-size="12" font-weight="800" fill="#ffffff">{r['title']}</text>
            <text x="550" y="{curr_y + 18}" font-family="sans-serif" font-size="11" fill="#cbd5e1">Location: <tspan fill="#ffffff" font-weight="700">{r['loc']}</tspan></text>
            <text x="1000" y="{curr_y + 18}" font-family="sans-serif" font-size="11" fill="#cbd5e1">Network: <tspan fill="#93c5fd" font-weight="700">{r['net_role']}</tspan></text>
            <text x="{rack_w - 40}" y="{curr_y + 18}" font-family="monospace" font-size="11" font-weight="800" fill="#fef08a">1756-A13 (13 SLOTS)</text>
        '''

        # Power Supply module on left
        ps_x = 52
        ps_y = curr_y + 32
        ps_w = 48
        ps_h = 120
        rack_svg += f'''
            <rect x="{ps_x}" y="{ps_y}" width="{ps_w}" height="{ps_h}" rx="4" fill="#ffffff" stroke="#cbd5e1" stroke-width="1" />
            <text x="{ps_x + ps_w/2}" y="{ps_y + 20}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#475569">PWR</text>
            <text x="{ps_x + ps_w/2}" y="{ps_y + 36}" text-anchor="middle" font-family="sans-serif" font-size="8" fill="#64748b">1756</text>
            <text x="{ps_x + ps_w/2}" y="{ps_y + 48}" text-anchor="middle" font-family="sans-serif" font-size="8" fill="#64748b">PA72</text>
            <circle cx="{ps_x + ps_w/2}" cy="{ps_y + 80}" r="4" fill="#059669" />
            <text x="{ps_x + ps_w/2}" y="{ps_y + 98}" text-anchor="middle" font-family="monospace" font-size="7" font-weight="700" fill="#059669">DC OK</text>
        '''

        # 13 slots
        slot_origin_x = 110
        for s in r['slots']:
            sx = slot_origin_x + s['slot'] * (slot_w + slot_gap)
            sy = curr_y + 32
            sw = slot_w
            sh = 120

            is_empty = s['name'] == 'Empty'
            hatch_fill = 'url(#diagonalHatch)' if is_empty else 'none'

            led1_col = "#cbd5e1" if is_empty else "#34d399"
            led2_col = "#cbd5e1" if is_empty else "#38bdf8"

            rack_svg += f'''
            <g id="Slot_{r['id']}_S{s['slot']:02d}">
                <rect x="{sx}" y="{sy}" width="{sw}" height="{sh}" rx="4" fill="url(#{s['grad']})" stroke="#cbd5e1" stroke-width="1" />
                {f'<rect x="{sx}" y="{sy}" width="{sw}" height="{sh}" rx="4" fill="{hatch_fill}" />' if is_empty else ''}

                <!-- Slot Number Badge -->
                <rect x="{sx}" y="{sy}" width="{sw}" height="20" rx="3" fill="#1e293b" />
                <text x="{sx + sw/2}" y="{sy + 14}" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="#f8fafc">SLOT {s['slot']:02d}</text>

                <!-- LED status dots -->
                <circle cx="{sx + 14}" cy="{sy + 28}" r="2.5" fill="{led1_col}" />
                <circle cx="{sx + 24}" cy="{sy + 28}" r="2.5" fill="{led2_col}" />
                <text x="{sx + sw - 8}" y="{sy + 31}" text-anchor="end" font-family="monospace" font-size="8" font-weight="700" fill="{s['txtCol']}">{s['sig']}</text>

                <!-- Card Name & Code -->
                <text x="{sx + sw/2}" y="{sy + 52}" text-anchor="middle" font-family="sans-serif" font-size="12" font-weight="800" fill="{ '#64748b' if is_empty else '#ffffff' }">{s['name']}</text>
                <text x="{sx + sw/2}" y="{sy + 70}" text-anchor="middle" font-family="monospace" font-size="10" font-weight="700" fill="{s['txtCol']}">{s['card']}</text>

                <!-- Terminal / Description -->
                <rect x="{sx + 8}" y="{sy + 82}" width="{sw - 16}" height="30" rx="3" fill="#1e293b" stroke="#334155" stroke-width="0.75" />
                <text x="{sx + sw/2}" y="{sy + 96}" text-anchor="middle" font-family="sans-serif" font-size="8" font-weight="600" fill="#e2e8f0">{s['desc'][:18]}</text>
                <text x="{sx + sw/2}" y="{sy + 107}" text-anchor="middle" font-family="sans-serif" font-size="8" fill="#94a3b8">{s['desc'][18:36]}</text>
            </g>
            '''

        rack_svg += '</g>'
        svg_elements.append(rack_svg)

    # ==========================================
    # 4. SIDE PANEL / SUMMARY TABLES (Right Column)
    # ==========================================
    side_x = 1840
    side_y = 580
    side_w = 620

    tables_svg = f'''
    <g id="RightSideTables">
        <!-- TABLE 1: MODULE QUANTITY SUMMARY -->
        <rect x="{side_x}" y="{side_y}" width="{side_w}" height="380" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" filter="url(#shadow)" />
        <rect x="{side_x}" y="{side_y}" width="{side_w}" height="32" rx="6" fill="#334155" />
        <text x="{side_x + 15}" y="{side_y + 21}" font-family="sans-serif" font-size="12" font-weight="800" fill="#ffffff">1. REDESIGNED BILL OF MATERIALS &amp; I/O QUANTITY MATRIX</text>

        <!-- Headers -->
        <g transform="translate({side_x + 15}, {side_y + 45})">
            <rect x="0" y="0" width="{side_w - 30}" height="24" fill="#f1f5f9" />
            <text x="10" y="16" font-family="sans-serif" font-size="10" font-weight="800" fill="#475569">CARD CODE</text>
            <text x="95" y="16" font-family="sans-serif" font-size="10" font-weight="800" fill="#475569">MODEL NO.</text>
            <text x="210" y="16" font-family="sans-serif" font-size="10" font-weight="800" fill="#475569">DESCRIPTION</text>
            <text x="440" y="16" font-family="sans-serif" font-size="10" font-weight="800" fill="#475569">Q'TY</text>
            <text x="500" y="16" font-family="sans-serif" font-size="10" font-weight="800" fill="#475569">CHANNELS</text>
        </g>

        <!-- Rows -->
        <!-- Row 1: CPU -->
        <g transform="translate({side_x + 15}, {side_y + 75})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#ffffff" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#1d4ed8">CPU</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-L950TPSXT</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">ControlLogix 5580 Process Safety (XT)</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#0284c7">1</text>
            <text x="510" y="20" font-family="sans-serif" font-size="10" fill="#64748b">50MB Memory</text>
        </g>
        <!-- Row 2: Comm -->
        <g transform="translate({side_x + 15}, {side_y + 110})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#f8fafc" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#0891b2">Eth_Module</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-EN4TR</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">Dual-Port 1Gbps DLR EtherNet/IP Adapter</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#0284c7">6</text>
            <text x="510" y="20" font-family="sans-serif" font-size="10" fill="#64748b">DLR / Uplink</text>
        </g>
        <!-- Row 3: 32DI -->
        <g transform="translate({side_x + 15}, {side_y + 145})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#ffffff" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#059669">32DI</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-IB32</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">24VDC Digital Input (Sink/Source)</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#059669">20</text>
            <text x="510" y="20" font-family="monospace" font-size="11" font-weight="700" fill="#059669">640 DI pts</text>
        </g>
        <!-- Row 4: 32DO -->
        <g transform="translate({side_x + 15}, {side_y + 180})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#f8fafc" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#ea580c">32DO</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-OB32</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">24VDC Sourcing Digital Output</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#ea580c">10</text>
            <text x="510" y="20" font-family="monospace" font-size="11" font-weight="700" fill="#ea580c">320 DO pts</text>
        </g>
        <!-- Row 5: 16AI -->
        <g transform="translate({side_x + 15}, {side_y + 215})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#ffffff" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#7e22ce">16AI</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-IF16</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">16-Ch Voltage/Current Analog Input</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#7e22ce">10</text>
            <text x="510" y="20" font-family="monospace" font-size="11" font-weight="700" fill="#7e22ce">160 AI pts</text>
        </g>
        <!-- Row 6: 8AO -->
        <g transform="translate({side_x + 15}, {side_y + 250})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#f8fafc" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#e11d48">8AO</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-OF8</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">8-Ch Voltage/Current Analog Output</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#e11d48">10</text>
            <text x="510" y="20" font-family="monospace" font-size="11" font-weight="700" fill="#e11d48">80 AO pts</text>
        </g>
        <!-- Row 7: Empty/Spare -->
        <g transform="translate({side_x + 15}, {side_y + 285})">
            <rect x="0" y="0" width="{side_w - 30}" height="32" fill="#ffffff" />
            <text x="10" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#64748b">Empty</text>
            <text x="95" y="20" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">1756-N2</text>
            <text x="210" y="20" font-family="sans-serif" font-size="10" fill="#334155">Chassis Slot Filler Covers (Spares)</text>
            <text x="450" y="20" font-family="monospace" font-size="11" font-weight="800" fill="#64748b">8</text>
            <text x="510" y="20" font-family="sans-serif" font-size="10" fill="#94a3b8">Future Exp</text>
        </g>
        <!-- Total Row -->
        <g transform="translate({side_x + 15}, {side_y + 325})">
            <rect x="0" y="0" width="{side_w - 30}" height="36" rx="4" fill="#e0f2fe" stroke="#0284c7" stroke-width="1.5" />
            <text x="15" y="23" font-family="sans-serif" font-size="11" font-weight="800" fill="#0369a1">TOTAL SLOTS ALLOCATED</text>
            <text x="210" y="23" font-family="sans-serif" font-size="10" font-weight="600" fill="#0369a1">5 CHASSIS x 13 SLOTS = 65 SLOTS</text>
            <text x="440" y="23" font-family="monospace" font-size="13" font-weight="800" fill="#0369a1">65</text>
            <text x="500" y="23" font-family="monospace" font-size="11" font-weight="800" fill="#0284c7">1,200 I/O</text>
        </g>

        <!-- TABLE 2: NETWORK CABLE SCHEDULE & CONNECTIONS -->
        <rect x="{side_x}" y="{side_y + 400}" width="{side_w}" height="380" rx="8" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" filter="url(#shadow)" />
        <rect x="{side_x}" y="{side_y + 400}" width="{side_w}" height="32" rx="6" fill="#334155" />
        <text x="{side_x + 15}" y="{side_y + 421}" font-family="sans-serif" font-size="12" font-weight="800" fill="#ffffff">2. DLR NETWORK CABLE SCHEDULE &amp; RUN LIST</text>

        <!-- Schedule Headers -->
        <g transform="translate({side_x + 15}, {side_y + 440})">
            <rect x="0" y="0" width="{side_w - 30}" height="22" fill="#f1f5f9" />
            <text x="10" y="15" font-family="sans-serif" font-size="9" font-weight="800" fill="#475569">CABLE TAG</text>
            <text x="100" y="15" font-family="sans-serif" font-size="9" font-weight="800" fill="#475569">FROM (PORT)</text>
            <text x="230" y="15" font-family="sans-serif" font-size="9" font-weight="800" fill="#475569">TO (PORT)</text>
            <text x="360" y="15" font-family="sans-serif" font-size="9" font-weight="800" fill="#475569">MEDIA / SPEC</text>
            <text x="510" y="15" font-family="sans-serif" font-size="9" font-weight="800" fill="#475569">EST. LEN</text>
        </g>
    '''

    cable_rows = [
        ("CBL-DLR-01", "CS1-EN4TR (P1)", "CS2-EN4TR (P1)", "Cat6A Shielded S/FTP", "1.5 m", "#059669"),
        ("CBL-DLR-02", "CS2-EN4TR (P2)", "CS3-EN4TR (P1)", "Cat6A Shielded S/FTP", "1.5 m", "#059669"),
        ("CBL-DLR-03", "CS3-EN4TR (P2)", "CS4-EN4TR (P1)", "Cat6A Shielded S/FTP", "1.5 m", "#059669"),
        ("CBL-DLR-04", "CS4-EN4TR (P2)", "MC-01 (RJ45-1)", "Cat6A Shielded S/FTP", "2.0 m", "#059669"),
        ("FOB-RIO-01", "MC-01 (FO-1 TX/RX)", "MC-02 (FO-1 TX/RX)", "4C Armored OM3/OS2", "200 m", "#d97706"),
        ("CBL-DLR-05", "MC-02 (RJ45-1)", "CS5-EN4TR (P1)", "Cat6A Shielded S/FTP", "1.5 m", "#059669"),
        ("CBL-DLR-06", "CS5-EN4TR (P2)", "MC-02 (RJ45-2)", "Cat6A Shielded S/FTP", "1.5 m", "#059669"),
        ("FOB-RIO-02", "MC-02 (FO-2 TX/RX)", "MC-01 (FO-2 TX/RX)", "4C Armored OM3/OS2", "200 m", "#d97706"),
        ("CBL-DLR-07", "MC-01 (RJ45-2)", "CS1-EN4TR (P2 Beacon)", "Cat6A S/FTP (Ring Close)", "2.0 m", "#059669"),
        ("CBL-UPL-01", "CS1-EN4TR (S02)", "Core Switch (VLAN 30)", "Cat6A / 1Gbps Fiber", "15 m", "#0284c7")
    ]

    for c_idx, c in enumerate(cable_rows):
        row_y = side_y + 466 + c_idx * 31
        bg = "#ffffff" if c_idx % 2 == 0 else "#f8fafc"
        tables_svg += f'''
        <g transform="translate({side_x + 15}, {row_y})">
            <rect x="0" y="0" width="{side_w - 30}" height="29" fill="{bg}" />
            <text x="10" y="19" font-family="monospace" font-size="10" font-weight="800" fill="{c[5]}">{c[0]}</text>
            <text x="100" y="19" font-family="sans-serif" font-size="9" font-weight="600" fill="#0f172a">{c[1]}</text>
            <text x="230" y="19" font-family="sans-serif" font-size="9" font-weight="600" fill="#0f172a">{c[2]}</text>
            <text x="360" y="19" font-family="sans-serif" font-size="9" fill="#334155">{c[3]}</text>
            <text x="510" y="19" font-family="monospace" font-size="9" font-weight="700" fill="#64748b">{c[4]}</text>
        </g>
        '''

    # Engineering Notes at bottom right
    tables_svg += f'''
        <g transform="translate({side_x}, {side_y + 790})">
            <rect x="0" y="0" width="{side_w}" height="90" rx="6" fill="#fffdf5" stroke="#f59e0b" stroke-width="1.2" />
            <text x="15" y="22" font-family="sans-serif" font-size="11" font-weight="800" fill="#b45309">ENGINEERING NOTES &amp; DLR DESIGN RULES:</text>
            <text x="15" y="40" font-family="sans-serif" font-size="9" fill="#334155">1. CS1 Slot 1 is the ACTIVE RING SUPERVISOR. Port 2 is beacon-blocked during normal operation.</text>
            <text x="15" y="55" font-family="sans-serif" font-size="9" fill="#334155">2. DLR Beacon Interval configured at 400µs; single cable fault failover recovers in &lt; 3ms.</text>
            <text x="15" y="70" font-family="sans-serif" font-size="9" fill="#334155">3. Expansion chassis CS2–CS5 standardized with 1756-EN4TR in Slot 00 per Rockwell CPwE guideline.</text>
        </g>
    </g>
    '''
    svg_elements.append(tables_svg)

    svg_elements.append('</svg>')

    full_svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {SVG_WIDTH} {SVG_HEIGHT}" width="{SVG_WIDTH}" height="{SVG_HEIGHT}">''' + ''.join(svg_elements)
    return full_svg

if __name__ == '__main__':
    output_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(output_dir, exist_ok=True)
    
    svg_content = generate_svg()
    svg_path = os.path.join(output_dir, "ControlLogix_5Chassis_DLR_Network_and_Rack_Design.svg")
    with open(svg_path, "w", encoding="utf-8") as f:
        f.write(svg_content)
    print(f"Successfully generated SVG: {svg_path}")

    # Export PDF
    pdf_doc = fitz.open(svg_path)
    pdf_bytes = pdf_doc.convert_to_pdf()
    pdf_final = fitz.open('pdf', pdf_bytes)
    pdf_path = os.path.join(output_dir, "ControlLogix_5Chassis_DLR_Network_and_Rack_Design.pdf")
    pdf_final.save(pdf_path)
    print(f"Successfully generated PDF: {pdf_path}")

    # Export High-Res PNG (150 DPI)
    page = pdf_final[0]
    pix = page.get_pixmap(dpi=150)
    png_path = os.path.join(output_dir, "ControlLogix_5Chassis_DLR_Network_and_Rack_Design.png")
    pix.save(png_path)
    print(f"Successfully generated PNG: {png_path} ({pix.width}x{pix.height})")
