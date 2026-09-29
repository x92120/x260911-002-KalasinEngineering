#!/usr/bin/env python3
"""
Generate Vector Diagram & Deliverables for:
Redundant Cisco Core Switches & Redundant Server Infrastructure (LIGHT STYLE)
Features:
- Dual Cisco Catalyst Core Switches (SW-CORE-01 & SW-CORE-02) with HSRP & LACP Port-Channel ISL
- Redundant FactoryTalk View SE Servers (Primary/Secondary) with Dual-Homed NIC Teaming & Dedicated Heartbeat
- Fiber Optic Backbone Trunk to Existing Server Room (Domain Controllers & Historian Collective)
- 802.1Q VLAN Trunk to MCC Room (Stratix 5700 & PowerFlex 755/525 Drives)
- Control LAN Uplink to ControlLogix 5580 Master Chassis (MCP-01)
Facility: Ingredion Kalasin Spray Dryer Plant
Theme: Clean Industrial Light Style (Suitable for crisp paper printing and high-contrast review)
"""

import os
import fitz

SVG_W = 2600
SVG_H = 1750

def generate_svg():
    svg = []

    # Defs: Gradients, Filters, Patterns for Light Style
    svg.append(f'''
    <defs>
        <filter id="shadow" x="-5%" y="-5%" width="110%" height="115%">
            <feDropShadow dx="0" dy="4" stdDeviation="5" flood-color="#64748b" flood-opacity="0.16"/>
        </filter>

        <pattern id="lightGrid" width="40" height="40" patternUnits="userSpaceOnUse">
            <path d="M 40 0 L 0 0 0 40" fill="none" stroke="#f1f5f9" stroke-width="1"/>
        </pattern>
    </defs>
    ''')

    # Background Canvas (Crisp White with subtle technical grid)
    svg.append(f'<rect width="{SVG_W}" height="{SVG_H}" fill="#ffffff" />')
    svg.append(f'<rect width="{SVG_W}" height="{SVG_H}" fill="url(#lightGrid)" />')

    # ====================================================
    # 1. HEADER & TITLE BLOCK
    # ====================================================
    svg.append(f'''
    <g id="Header">
        <rect x="40" y="25" width="{SVG_W - 80}" height="95" rx="8" fill="#f8fafc" stroke="#cbd5e1" stroke-width="1.5" filter="url(#shadow)" />
        <text x="70" y="60" font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" font-size="22" font-weight="800" fill="#0f172a" letter-spacing="0.8">INGREDION THAILAND — SPRINT 18K SPRAY DRYER PLANT (KALASIN)</text>
        <text x="70" y="90" font-family="-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif" font-size="14" font-weight="700" fill="#0284c7" letter-spacing="0.5">REDUNDANT CISCO CORE SWITCHES &amp; REDUNDANT SERVER INFRASTRUCTURE CONFIGURATION</text>

        <!-- Technical Badges -->
        <g transform="translate({SVG_W - 840}, 45)">
            <rect x="0" y="0" width="130" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="65" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">DRAWING NO.</text>
            <text x="65" y="36" text-anchor="middle" font-family="sans-serif" font-size="12" fill="#0284c7" font-weight="800">DWG-260911-NET-02</text>
        </g>
        <g transform="translate({SVG_W - 690}, 45)">
            <rect x="0" y="0" width="120" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="60" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">CISCO PROTOCOL</text>
            <text x="60" y="36" text-anchor="middle" font-family="sans-serif" font-size="12" fill="#0284c7" font-weight="800">HSRP v2 + LACP</text>
        </g>
        <g transform="translate({SVG_W - 550}, 45)">
            <rect x="0" y="0" width="130" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="65" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">SERVER REDUNDANCY</text>
            <text x="65" y="36" text-anchor="middle" font-family="sans-serif" font-size="12" fill="#059669" font-weight="800">FT VIEW SE CLUSTER</text>
        </g>
        <g transform="translate({SVG_W - 400}, 45)">
            <rect x="0" y="0" width="150" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="75" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">MCC TRUNK</text>
            <text x="75" y="36" text-anchor="middle" font-family="sans-serif" font-size="11" fill="#d97706" font-weight="800">802.1Q (VLAN 30,50,99)</text>
        </g>
        <g transform="translate({SVG_W - 230}, 45)">
            <rect x="0" y="0" width="170" height="45" rx="5" fill="#ffffff" stroke="#cbd5e1" />
            <text x="85" y="18" text-anchor="middle" font-family="sans-serif" font-size="10" fill="#64748b" font-weight="700">SERVER ROOM LINK</text>
            <text x="85" y="36" text-anchor="middle" font-family="sans-serif" font-size="11" fill="#db2777" font-weight="800">DUAL 10G FIBER TRUNK</text>
        </g>
    </g>
    ''')

    # ====================================================
    # 2. ZONE 1: EXISTING SERVER ROOM (Top Left)
    # ====================================================
    sr_x, sr_y, sr_w, sr_h = 50, 150, 680, 480
    svg.append(f'''
    <g id="Zone_ServerRoom">
        <rect x="{sr_x}" y="{sr_y}" width="{sr_w}" height="{sr_h}" rx="8" fill="#fff5f7" stroke="#f43f5e" stroke-width="1.8" stroke-dasharray="6,4" filter="url(#shadow)" />
        <rect x="{sr_x}" y="{sr_y}" width="280" height="32" rx="6" fill="#e11d48" />
        <text x="{sr_x + 15}" y="{sr_y + 21}" font-family="sans-serif" font-size="13" font-weight="700" fill="#ffffff">EXISTING SERVER ROOM (ZONE 1)</text>
        <text x="{sr_x + sr_w - 210}" y="{sr_y + 21}" font-family="monospace" font-size="11" font-weight="700" fill="#be123c">FIBER BACKBONE ENDPOINT</text>

        <!-- FDF-SR01 Fiber Distribution Panel -->
        <rect x="{sr_x + 30}" y="{sr_y + 50}" width="{sr_w - 60}" height="45" rx="4" fill="#ffffff" stroke="#f43f5e" stroke-width="1.5" />
        <text x="{sr_x + 45}" y="{sr_y + 77}" font-family="monospace" font-size="11" font-weight="800" fill="#be123c">FDF-SR01</text>
        <text x="{sr_x + 130}" y="{sr_y + 77}" font-family="sans-serif" font-size="11" font-weight="600" fill="#334155">12-Core OM3/OS2 Fiber Optic Patch Panel (LC Duplex)</text>
    ''')

    # Servers in Existing Server Room
    sr_servers = [
        ('DC01: Primary Domain Controller', 'Active Directory / DNS / NTP Master', '192.168.10.30', 'HPE DL380 Gen10', sr_y + 115),
        ('DC02: Secondary Domain Controller', 'Active Directory Backup / Failover', '192.168.10.31', 'HPE DL380 Gen10', sr_y + 195),
        ('FT Historian SE 01 (Primary)', 'FactoryTalk Process Data Archive', '192.168.10.20', 'HPE DL380 Gen11', sr_y + 275),
        ('FT Historian SE 02 (Mirror)', 'OSIsoft PI Collective Mirror Node', '192.168.10.21', 'HPE DL380 Gen11', sr_y + 355),
        ('Server Room Online UPS Unit', '5kVA Redundant Dual-Power Source', 'SNMP: 192.168.99.10', 'APC Smart-UPS RT', sr_y + 425)
    ]

    for s_title, s_role, s_ip, s_hw, sy in sr_servers:
        is_ups = 'UPS' in s_title
        border = '#94a3b8' if is_ups else '#fda4af'
        svg.append(f'''
        <g transform="translate({sr_x + 30}, {sy})">
            <rect x="0" y="0" width="{sr_w - 60}" height="65" rx="5" fill="#ffffff" stroke="{border}" stroke-width="1.2" />
            <text x="15" y="24" font-family="sans-serif" font-size="11" font-weight="800" fill="#0f172a">{s_title}</text>
            <text x="15" y="42" font-family="sans-serif" font-size="10" fill="#64748b">{s_role}</text>
            <rect x="{sr_w - 230}" y="10" width="150" height="22" rx="3" fill="#fff1f2" stroke="#fecdd3" />
            <text x="{sr_w - 155}" y="25" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#be123c">IP: {s_ip}</text>
            <text x="{sr_w - 155}" y="50" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="600" fill="#94a3b8">{s_hw}</text>
        </g>
        ''')

    svg.append('</g>')

    # ====================================================
    # 3. ZONE 2: CENTRAL OPERATION 1 ROOM (Center & Right)
    # ====================================================
    op_x, op_y, op_w, op_h = 780, 150, 1770, 950
    svg.append(f'''
    <g id="Zone_OperationRoom">
        <rect x="{op_x}" y="{op_y}" width="{op_w}" height="{op_h}" rx="8" fill="#f8fafc" stroke="#0284c7" stroke-width="1.8" filter="url(#shadow)" />
        <rect x="{op_x}" y="{op_y}" width="420" height="32" rx="6" fill="#0284c7" />
        <text x="{op_x + 15}" y="{op_y + 21}" font-family="sans-serif" font-size="13" font-weight="700" fill="#ffffff">OPERATION 1 ROOM — CENTRAL CONTROL &amp; NETWORK CABINET</text>
        <text x="{op_x + op_w - 320}" y="{op_y + 21}" font-family="monospace" font-size="11" font-weight="700" fill="#0369a1">REDUNDANT CORE SWITCH &amp; SERVERS</text>
    ''')

    # 3A. REDUNDANT CISCO CORE SWITCHES
    c1_x, c1_y, c_w, c_h = op_x + 40, op_y + 60, 520, 220
    c2_x, c2_y = op_x + 640, op_y + 60

    # Cisco SW-CORE-01 (Active)
    svg.append(f'''
    <g id="Switch_Cisco_Core_01">
        <rect x="{c1_x}" y="{c1_y}" width="{c_w}" height="{c_h}" rx="6" fill="#ffffff" stroke="#0284c7" stroke-width="2" filter="url(#shadow)" />
        <rect x="{c1_x}" y="{c1_y}" width="{c_w}" height="32" rx="6" fill="#0284c7" />
        <text x="{c1_x + 15}" y="{c1_y + 21}" font-family="sans-serif" font-size="13" font-weight="800" fill="#ffffff">SW-CORE-01 (CISCO CATALYST 9300)</text>
        <rect x="{c1_x + c_w - 140}" y="{c1_y + 7}" width="125" height="18" rx="4" fill="#0369a1" />
        <text x="{c1_x + c_w - 78}" y="{c1_y + 20}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#ffffff">★ HSRP ACTIVE (110)</text>

        <text x="{c1_x + 15}" y="{c1_y + 55}" font-family="sans-serif" font-size="11" fill="#334155">Role: <tspan fill="#0284c7" font-weight="800">Primary Layer 3 Gateway &amp; STP Root</tspan></text>
        <text x="{c1_x + 15}" y="{c1_y + 75}" font-family="monospace" font-size="11" font-weight="700" fill="#059669">Mgmt IP: 192.168.99.2 (VLAN 99)</text>

        <!-- SVI Table inside SW1 -->
        <rect x="{c1_x + 15}" y="{c1_y + 90}" width="{c_w - 30}" height="115" rx="4" fill="#f8fafc" stroke="#cbd5e1" />
        <text x="{c1_x + 25}" y="{c1_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">VLAN SVI</text>
        <text x="{c1_x + 130}" y="{c1_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">SWITCH SVI IP</text>
        <text x="{c1_x + 280}" y="{c1_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">HSRP VIP (GATEWAY)</text>
        <text x="{c1_x + 440}" y="{c1_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">STATE</text>

        <text x="{c1_x + 25}" y="{c1_y + 130}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 10 (SCADA)</text>
        <text x="{c1_x + 130}" y="{c1_y + 130}" font-family="monospace" font-size="10" font-weight="700" fill="#0284c7">192.168.10.2</text>
        <text x="{c1_x + 280}" y="{c1_y + 130}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.10.1</text>
        <text x="{c1_x + 440}" y="{c1_y + 130}" font-family="monospace" font-size="10" font-weight="800" fill="#0284c7">Active</text>

        <text x="{c1_x + 25}" y="{c1_y + 150}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 20 (HMI)</text>
        <text x="{c1_x + 130}" y="{c1_y + 150}" font-family="monospace" font-size="10" font-weight="700" fill="#0284c7">192.168.20.2</text>
        <text x="{c1_x + 280}" y="{c1_y + 150}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.20.1</text>
        <text x="{c1_x + 440}" y="{c1_y + 150}" font-family="monospace" font-size="10" font-weight="800" fill="#0284c7">Active</text>

        <text x="{c1_x + 25}" y="{c1_y + 170}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 30 (PLC)</text>
        <text x="{c1_x + 130}" y="{c1_y + 170}" font-family="monospace" font-size="10" font-weight="700" fill="#0284c7">192.168.30.2</text>
        <text x="{c1_x + 280}" y="{c1_y + 170}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.30.1</text>
        <text x="{c1_x + 440}" y="{c1_y + 170}" font-family="monospace" font-size="10" font-weight="800" fill="#0284c7">Active</text>

        <text x="{c1_x + 25}" y="{c1_y + 190}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 50 (MCC)</text>
        <text x="{c1_x + 130}" y="{c1_y + 190}" font-family="monospace" font-size="10" font-weight="700" fill="#0284c7">192.168.50.2</text>
        <text x="{c1_x + 280}" y="{c1_y + 190}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.50.1</text>
        <text x="{c1_x + 440}" y="{c1_y + 190}" font-family="monospace" font-size="10" font-weight="800" fill="#0284c7">Active</text>
    </g>
    ''')

    # Cisco SW-CORE-02 (Standby)
    svg.append(f'''
    <g id="Switch_Cisco_Core_02">
        <rect x="{c2_x}" y="{c2_y}" width="{c_w}" height="{c_h}" rx="6" fill="#ffffff" stroke="#0d9488" stroke-width="2" filter="url(#shadow)" />
        <rect x="{c2_x}" y="{c2_y}" width="{c_w}" height="32" rx="6" fill="#0d9488" />
        <text x="{c2_x + 15}" y="{c2_y + 21}" font-family="sans-serif" font-size="13" font-weight="800" fill="#ffffff">SW-CORE-02 (CISCO CATALYST 9300)</text>
        <rect x="{c2_x + c_w - 140}" y="{c2_y + 7}" width="125" height="18" rx="4" fill="#115e59" />
        <text x="{c2_x + c_w - 78}" y="{c2_y + 20}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#ffffff">HSRP STANDBY (100)</text>

        <text x="{c2_x + 15}" y="{c2_y + 55}" font-family="sans-serif" font-size="11" fill="#334155">Role: <tspan fill="#0d9488" font-weight="800">Secondary Gateway &amp; Hot Standby</tspan></text>
        <text x="{c2_x + 15}" y="{c2_y + 75}" font-family="monospace" font-size="11" font-weight="700" fill="#059669">Mgmt IP: 192.168.99.3 (VLAN 99)</text>

        <!-- SVI Table inside SW2 -->
        <rect x="{c2_x + 15}" y="{c2_y + 90}" width="{c_w - 30}" height="115" rx="4" fill="#f8fafc" stroke="#cbd5e1" />
        <text x="{c2_x + 25}" y="{c2_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">VLAN SVI</text>
        <text x="{c2_x + 130}" y="{c2_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">SWITCH SVI IP</text>
        <text x="{c2_x + 280}" y="{c2_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">HSRP VIP (GATEWAY)</text>
        <text x="{c2_x + 440}" y="{c2_y + 110}" font-family="monospace" font-size="10" font-weight="800" fill="#64748b">STATE</text>

        <text x="{c2_x + 25}" y="{c2_y + 130}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 10 (SCADA)</text>
        <text x="{c2_x + 130}" y="{c2_y + 130}" font-family="monospace" font-size="10" font-weight="700" fill="#0d9488">192.168.10.3</text>
        <text x="{c2_x + 280}" y="{c2_y + 130}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.10.1</text>
        <text x="{c2_x + 440}" y="{c2_y + 130}" font-family="monospace" font-size="10" font-weight="800" fill="#d97706">Standby</text>

        <text x="{c2_x + 25}" y="{c2_y + 150}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 20 (HMI)</text>
        <text x="{c2_x + 130}" y="{c2_y + 150}" font-family="monospace" font-size="10" font-weight="700" fill="#0d9488">192.168.20.3</text>
        <text x="{c2_x + 280}" y="{c2_y + 150}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.20.1</text>
        <text x="{c2_x + 440}" y="{c2_y + 150}" font-family="monospace" font-size="10" font-weight="800" fill="#d97706">Standby</text>

        <text x="{c2_x + 25}" y="{c2_y + 170}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 30 (PLC)</text>
        <text x="{c2_x + 130}" y="{c2_y + 170}" font-family="monospace" font-size="10" font-weight="700" fill="#0d9488">192.168.30.3</text>
        <text x="{c2_x + 280}" y="{c2_y + 170}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.30.1</text>
        <text x="{c2_x + 440}" y="{c2_y + 170}" font-family="monospace" font-size="10" font-weight="800" fill="#d97706">Standby</text>

        <text x="{c2_x + 25}" y="{c2_y + 190}" font-family="monospace" font-size="10" font-weight="600" fill="#0f172a">VLAN 50 (MCC)</text>
        <text x="{c2_x + 130}" y="{c2_y + 190}" font-family="monospace" font-size="10" font-weight="700" fill="#0d9488">192.168.50.3</text>
        <text x="{c2_x + 280}" y="{c2_y + 190}" font-family="monospace" font-size="10" font-weight="800" fill="#059669">192.168.50.1</text>
        <text x="{c2_x + 440}" y="{c2_y + 190}" font-family="monospace" font-size="10" font-weight="800" fill="#d97706">Standby</text>
    </g>
    ''')

    # ISL Cross-Connect Cable between Cisco Core Switches (Port-Channel 1)
    svg.append(f'''
    <!-- Inter-Switch Link (LACP Port-Channel 1) -->
    <path d="M {c1_x + c_w} {c1_y + 70} L {c2_x} {c2_y + 70}" stroke="#0284c7" stroke-width="4" />
    <path d="M {c1_x + c_w} {c1_y + 85} L {c2_x} {c2_y + 85}" stroke="#0284c7" stroke-width="4" />
    <rect x="{(c1_x + c_w + c2_x)/2 - 85}" y="{c1_y + 62}" width="170" height="32" rx="4" fill="#e0f2fe" stroke="#0284c7" stroke-width="1.5" />
    <text x="{(c1_x + c_w + c2_x)/2}" y="{c1_y + 82}" text-anchor="middle" font-family="monospace" font-size="11" font-weight="800" fill="#0369a1">ISL: LACP Po1 (20G)</text>
    ''')

    # 3B. REDUNDANT SCADA SERVERS (FactoryTalk View SE Distributed)
    srv1_x, srv_y, srv_w, srv_h = op_x + 40, op_y + 350, 520, 260
    srv2_x = op_x + 640

    # Server 01 (Primary)
    svg.append(f'''
    <g id="Server_SCADA_01">
        <rect x="{srv1_x}" y="{srv_y}" width="{srv_w}" height="{srv_h}" rx="6" fill="#ffffff" stroke="#2563eb" stroke-width="2" filter="url(#shadow)" />
        <rect x="{srv1_x}" y="{srv_y}" width="{srv_w}" height="32" rx="6" fill="#2563eb" />
        <text x="{srv1_x + 15}" y="{srv_y + 21}" font-family="sans-serif" font-size="13" font-weight="800" fill="#ffffff">SRV-SCADA-01 (PRIMARY HMI SERVER)</text>
        <rect x="{srv1_x + srv_w - 120}" y="{srv_y + 7}" width="105" height="18" rx="4" fill="#1e40af" />
        <text x="{srv1_x + srv_w - 68}" y="{srv_y + 20}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#ffffff">● ACTIVE MASTER</text>

        <text x="{srv1_x + 15}" y="{srv_y + 55}" font-family="sans-serif" font-size="11" fill="#334155">Role: <tspan fill="#1d4ed8" font-weight="800">Primary FactoryTalk View SE Data &amp; Alarm Server</tspan></text>
        <text x="{srv1_x + 15}" y="{srv_y + 75}" font-family="monospace" font-size="12" font-weight="800" fill="#059669">Cluster IP: 192.168.10.10 (VLAN 10)</text>

        <!-- NIC details -->
        <rect x="{srv1_x + 15}" y="{srv_y + 90}" width="{srv_w - 30}" height="150" rx="4" fill="#f8fafc" stroke="#cbd5e1" />
        <text x="{srv1_x + 25}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">INTERFACE</text>
        <text x="{srv1_x + 120}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">TEAM / ROLE</text>
        <text x="{srv1_x + 240}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">CONNECTED TO</text>
        <text x="{srv1_x + 400}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">STATUS</text>

        <text x="{srv1_x + 25}" y="{srv_y + 135}" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">NIC 1 (10G)</text>
        <text x="{srv1_x + 120}" y="{srv_y + 135}" font-family="sans-serif" font-size="10" font-weight="600" fill="#0284c7">LACP Team A</text>
        <text x="{srv1_x + 240}" y="{srv_y + 135}" font-family="monospace" font-size="10" fill="#0369a1">SW-CORE-01 (Gi1/0/1)</text>
        <text x="{srv1_x + 400}" y="{srv_y + 135}" font-family="sans-serif" font-size="10" font-weight="800" fill="#059669">Active</text>

        <text x="{srv1_x + 25}" y="{srv_y + 160}" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">NIC 2 (10G)</text>
        <text x="{srv1_x + 120}" y="{srv_y + 160}" font-family="sans-serif" font-size="10" font-weight="600" fill="#0284c7">LACP Team B</text>
        <text x="{srv1_x + 240}" y="{srv_y + 160}" font-family="monospace" font-size="10" fill="#0d9488">SW-CORE-02 (Gi1/0/1)</text>
        <text x="{srv1_x + 400}" y="{srv_y + 160}" font-family="sans-serif" font-size="10" font-weight="800" fill="#059669">Standby Link</text>

        <text x="{srv1_x + 25}" y="{srv_y + 185}" font-family="monospace" font-size="10" font-weight="700" fill="#b45309">NIC 3 (Sync)</text>
        <text x="{srv1_x + 120}" y="{srv_y + 185}" font-family="sans-serif" font-size="10" font-weight="600" fill="#b45309">Heartbeat Private</text>
        <text x="{srv1_x + 240}" y="{srv_y + 185}" font-family="monospace" font-size="10" fill="#b45309">Server 02 (10.10.10.1)</text>
        <text x="{srv1_x + 400}" y="{srv_y + 185}" font-family="sans-serif" font-size="10" font-weight="800" fill="#b45309">&lt; 500ms</text>

        <text x="{srv1_x + 25}" y="{srv_y + 210}" font-family="monospace" font-size="10" font-weight="700" fill="#64748b">NIC 4 (iLO)</text>
        <text x="{srv1_x + 120}" y="{srv_y + 210}" font-family="sans-serif" font-size="10" fill="#64748b">Hardware Mgmt</text>
        <text x="{srv1_x + 240}" y="{srv_y + 210}" font-family="monospace" font-size="10" fill="#64748b">SW-MGMT-01 (VLAN99)</text>
        <text x="{srv1_x + 400}" y="{srv_y + 210}" font-family="sans-serif" font-size="10" font-weight="700" fill="#64748b">Ready</text>
    </g>
    ''')

    # Server 02 (Secondary / Standby)
    svg.append(f'''
    <g id="Server_SCADA_02">
        <rect x="{srv2_x}" y="{srv_y}" width="{srv_w}" height="{srv_h}" rx="6" fill="#ffffff" stroke="#7c3aed" stroke-width="2" filter="url(#shadow)" />
        <rect x="{srv2_x}" y="{srv_y}" width="{srv_w}" height="32" rx="6" fill="#7c3aed" />
        <text x="{srv2_x + 15}" y="{srv_y + 21}" font-family="sans-serif" font-size="13" font-weight="800" fill="#ffffff">SRV-SCADA-02 (SECONDARY HMI SERVER)</text>
        <rect x="{srv2_x + srv_w - 120}" y="{srv_y + 7}" width="105" height="18" rx="4" fill="#5b21b6" />
        <text x="{srv2_x + srv_w - 68}" y="{srv_y + 20}" text-anchor="middle" font-family="sans-serif" font-size="9" font-weight="800" fill="#ffffff">● HOT STANDBY</text>

        <text x="{srv2_x + 15}" y="{srv_y + 55}" font-family="sans-serif" font-size="11" fill="#334155">Role: <tspan fill="#6d28d9" font-weight="800">Secondary FactoryTalk View SE Data &amp; Alarm Server</tspan></text>
        <text x="{srv2_x + 15}" y="{srv_y + 75}" font-family="monospace" font-size="12" font-weight="800" fill="#059669">Cluster IP: 192.168.10.11 (VLAN 10)</text>

        <!-- NIC details -->
        <rect x="{srv2_x + 15}" y="{srv_y + 90}" width="{srv_w - 30}" height="150" rx="4" fill="#f8fafc" stroke="#cbd5e1" />
        <text x="{srv2_x + 25}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">INTERFACE</text>
        <text x="{srv2_x + 120}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">TEAM / ROLE</text>
        <text x="{srv2_x + 240}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">CONNECTED TO</text>
        <text x="{srv2_x + 400}" y="{srv_y + 112}" font-family="sans-serif" font-size="10" font-weight="800" fill="#64748b">STATUS</text>

        <text x="{srv2_x + 25}" y="{srv_y + 135}" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">NIC 1 (10G)</text>
        <text x="{srv2_x + 120}" y="{srv_y + 135}" font-family="sans-serif" font-size="10" font-weight="600" fill="#7c3aed">LACP Team A</text>
        <text x="{srv2_x + 240}" y="{srv_y + 135}" font-family="monospace" font-size="10" fill="#0d9488">SW-CORE-02 (Gi1/0/2)</text>
        <text x="{srv2_x + 400}" y="{srv_y + 135}" font-family="sans-serif" font-size="10" font-weight="800" fill="#059669">Active</text>

        <text x="{srv2_x + 25}" y="{srv_y + 160}" font-family="monospace" font-size="10" font-weight="700" fill="#0f172a">NIC 2 (10G)</text>
        <text x="{srv2_x + 120}" y="{srv_y + 160}" font-family="sans-serif" font-size="10" font-weight="600" fill="#7c3aed">LACP Team B</text>
        <text x="{srv2_x + 240}" y="{srv_y + 160}" font-family="monospace" font-size="10" fill="#0369a1">SW-CORE-01 (Gi1/0/2)</text>
        <text x="{srv2_x + 400}" y="{srv_y + 160}" font-family="sans-serif" font-size="10" font-weight="800" fill="#059669">Standby Link</text>

        <text x="{srv2_x + 25}" y="{srv_y + 185}" font-family="monospace" font-size="10" font-weight="700" fill="#b45309">NIC 3 (Sync)</text>
        <text x="{srv2_x + 120}" y="{srv_y + 185}" font-family="sans-serif" font-size="10" font-weight="600" fill="#b45309">Heartbeat Private</text>
        <text x="{srv2_x + 240}" y="{srv_y + 185}" font-family="monospace" font-size="10" fill="#b45309">Server 01 (10.10.10.2)</text>
        <text x="{srv2_x + 400}" y="{srv_y + 185}" font-family="sans-serif" font-size="10" font-weight="800" fill="#b45309">&lt; 500ms</text>

        <text x="{srv2_x + 25}" y="{srv_y + 210}" font-family="monospace" font-size="10" font-weight="700" fill="#64748b">NIC 4 (iLO)</text>
        <text x="{srv2_x + 120}" y="{srv_y + 210}" font-family="sans-serif" font-size="10" fill="#64748b">Hardware Mgmt</text>
        <text x="{srv2_x + 240}" y="{srv_y + 210}" font-family="monospace" font-size="10" fill="#64748b">SW-MGMT-01 (VLAN99)</text>
        <text x="{srv2_x + 400}" y="{srv_y + 210}" font-family="sans-serif" font-size="10" font-weight="700" fill="#64748b">Ready</text>
    </g>
    ''')

    # Dedicated Direct Inter-Server Heartbeat Cable
    svg.append(f'''
    <!-- Direct Heartbeat Link between Server 1 and Server 2 -->
    <path d="M {srv1_x + srv_w} {srv_y + 215} L {srv2_x} {srv_y + 215}" stroke="#d97706" stroke-width="3.5" stroke-dasharray="6,4" />
    <rect x="{(srv1_x + srv_w + srv2_x)/2 - 110}" y="{srv_y + 200}" width="220" height="30" rx="4" fill="#fef3c7" stroke="#d97706" stroke-width="1.5" />
    <text x="{(srv1_x + srv_w + srv2_x)/2}" y="{srv_y + 220}" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#92400e">DEDICATED HEARTBEAT (10.10.10.x/30)</text>
    ''')

    # Dual-Homed Server Uplinks to Cisco Switches
    svg.append(f'''
    <!-- Server 1 to SW1 -->
    <path d="M {srv1_x + 120} {srv_y} L {srv1_x + 120} {c1_y + c_h}" stroke="#0284c7" stroke-width="2.5" />
    <text x="{srv1_x + 128}" y="{(srv_y + c1_y + c_h)/2}" font-family="monospace" font-size="9" font-weight="700" fill="#0369a1">NIC1 &#8594; Gi1/0/1</text>

    <!-- Server 1 to SW2 (Cross-connect) -->
    <path d="M {srv1_x + 220} {srv_y} L {srv1_x + 220} {srv_y - 30} L {c2_x + 100} {srv_y - 30} L {c2_x + 100} {c2_y + c_h}" stroke="#0d9488" stroke-width="2.5" stroke-dasharray="5,3" />
    <text x="{srv1_x + 280}" y="{srv_y - 36}" font-family="monospace" font-size="9" font-weight="700" fill="#0f766e">NIC2 &#8594; SW2 Gi1/0/1 (Cross)</text>

    <!-- Server 2 to SW2 -->
    <path d="M {srv2_x + 350} {srv_y} L {srv2_x + 350} {c2_y + c_h}" stroke="#0d9488" stroke-width="2.5" />
    <text x="{srv2_x + 358}" y="{(srv_y + c2_y + c_h)/2}" font-family="monospace" font-size="9" font-weight="700" fill="#0f766e">NIC1 &#8594; Gi1/0/2</text>

    <!-- Server 2 to SW1 (Cross-connect) -->
    <path d="M {srv2_x + 250} {srv_y} L {srv2_x + 250} {srv_y - 50} L {c1_x + 380} {srv_y - 50} L {c1_x + 380} {c1_y + c_h}" stroke="#0284c7" stroke-width="2.5" stroke-dasharray="5,3" />
    <text x="{srv1_x + 380}" y="{srv_y - 56}" font-family="monospace" font-size="9" font-weight="700" fill="#0369a1">NIC2 &#8594; SW1 Gi1/0/2 (Cross)</text>
    ''')

    # 3C. OPERATOR CLIENT WORKSTATIONS (VLAN 20)
    cli_x = op_x + 1220
    cli_y = op_y + 60
    cli_w = 510
    cli_h = 550
    svg.append(f'''
    <g id="ControlRoom_Clients">
        <rect x="{cli_x}" y="{cli_y}" width="{cli_w}" height="{cli_h}" rx="6" fill="#ffffff" stroke="#64748b" stroke-width="1.5" filter="url(#shadow)" />
        <rect x="{cli_x}" y="{cli_y}" width="{cli_w}" height="32" rx="6" fill="#475569" />
        <text x="{cli_x + 15}" y="{cli_y + 21}" font-family="sans-serif" font-size="12" font-weight="800" fill="#ffffff">CONTROL ROOM CLIENTS &amp; OVERVIEW DISPLAYS (VLAN 20)</text>

        <!-- Workstation List -->
    ''')

    clients_data = [
        ('Engineering Workstation (EWS)', 'Studio 5000 / FT Admin / System Config', '192.168.20.150', cli_y + 55),
        ('Operator Station 1 (Main Console)', 'FT View SE Client / Batch Control', '192.168.20.101', cli_y + 135),
        ('Operator Station 2 (Main Console)', 'FT View SE Client / Dryer Control', '192.168.20.102', cli_y + 215),
        ('Operator Station 3 (Remote Station 2)', 'Remote Thin Client / Field Operation', '192.168.20.103', cli_y + 295),
        ('55" Overhead Display: Process Overview', 'Central High-Res Overview Monitor', 'ThinManager', cli_y + 375),
        ('55" Overhead Display: Trending &amp; Alarms', 'Process Curve &amp; Realtime Trends', 'ThinManager', cli_y + 445)
    ]

    for c_title, c_role, c_ip, cy in clients_data:
        svg.append(f'''
        <g transform="translate({cli_x + 20}, {cy})">
            <rect x="0" y="0" width="{cli_w - 40}" height="60" rx="4" fill="#f8fafc" stroke="#cbd5e1" />
            <text x="12" y="24" font-family="sans-serif" font-size="11" font-weight="800" fill="#0f172a">{c_title}</text>
            <text x="12" y="42" font-family="sans-serif" font-size="10" fill="#64748b">{c_role}</text>
            <rect x="{cli_w - 180}" y="12" width="130" height="22" rx="3" fill="#e0f2fe" stroke="#bae6fd" />
            <text x="{cli_w - 115}" y="27" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#0284c7">{c_ip}</text>
        </g>
        ''')

    svg.append('</g>')  # End ControlRoom_Clients
    svg.append('</g>')  # End Zone_OperationRoom

    # ====================================================
    # 4. ZONE 3: MCC ROOM (MOTOR CONTROL CENTER) (Bottom Right)
    # ====================================================
    mcc_x, mcc_y, mcc_w, mcc_h = 1350, 1140, 1200, 560
    svg.append(f'''
    <g id="Zone_MCCRoom">
        <rect x="{mcc_x}" y="{mcc_y}" width="{mcc_w}" height="{mcc_h}" rx="8" fill="#fffdf5" stroke="#f59e0b" stroke-width="1.8" stroke-dasharray="6,4" filter="url(#shadow)" />
        <rect x="{mcc_x}" y="{mcc_y}" width="340" height="32" rx="6" fill="#d97706" />
        <text x="{mcc_x + 15}" y="{mcc_y + 21}" font-family="sans-serif" font-size="13" font-weight="700" fill="#ffffff">MCC ROOM — MOTOR CONTROL CENTER (ZONE 3)</text>
        <text x="{mcc_x + mcc_w - 280}" y="{mcc_y + 21}" font-family="monospace" font-size="11" font-weight="700" fill="#b45309">802.1Q TRUNK: VLAN 30, 50, 99</text>

        <!-- SW-MCC-01 Stratix Managed Switch -->
        <g transform="translate({mcc_x + 30}, {mcc_y + 50})">
            <rect x="0" y="0" width="{mcc_w - 60}" height="80" rx="6" fill="#ffffff" stroke="#d97706" stroke-width="2" filter="url(#shadow)" />
            <rect x="0" y="0" width="{mcc_w - 60}" height="28" rx="6" fill="#d97706" />
            <text x="15" y="19" font-family="sans-serif" font-size="12" font-weight="800" fill="#ffffff">SW-MCC-01: ALLEN-BRADLEY STRATIX 5700 (1783-BMS10CGN)</text>
            <text x="{mcc_w - 220}" y="19" font-family="monospace" font-size="11" font-weight="800" fill="#ffffff">IP: 192.168.50.10</text>

            <text x="15" y="48" font-family="sans-serif" font-size="11" fill="#334155">Trunk Port 1: <tspan fill="#0284c7" font-weight="800">From SW-CORE-01 (Gi1/0/24)</tspan></text>
            <text x="350" y="48" font-family="sans-serif" font-size="11" fill="#334155">Trunk Port 2: <tspan fill="#0d9488" font-weight="800">From SW-CORE-02 (Gi1/0/24 - Standby)</tspan></text>
            <text x="750" y="48" font-family="monospace" font-size="11" font-weight="800" fill="#059669">Gateway VIP: 192.168.50.1</text>
        </g>
    ''')

    # Drives in MCC Room
    mcc_drives = [
        ('PowerFlex 755 (110kW)', 'Supply Air Fan FN-60201', '192.168.50.21', 'VLAN 50', mcc_x + 30, mcc_y + 150),
        ('PowerFlex 755 (132kW)', 'Exhaust Air Fan FN-60202', '192.168.50.22', 'VLAN 50', mcc_x + 420, mcc_y + 150),
        ('PowerFlex 525 (18.5kW)', 'Slurry Transfer Pump PC-20201', '192.168.50.31', 'VLAN 50', mcc_x + 810, mcc_y + 150),
        ('PowerFlex 525 (18.5kW)', 'Slurry Transfer Pump PC-20211', '192.168.50.32', 'VLAN 50', mcc_x + 30, mcc_y + 270),
        ('Power Meter 1408-BC3A', 'Main Incoming 400V Metering', '192.168.50.50', 'VLAN 50', mcc_x + 420, mcc_y + 270),
        ('Agitator Starters (ME-20201/11)', 'Slurry Tanks Agitator VFDs', '192.168.50.33', 'VLAN 50', mcc_x + 810, mcc_y + 270)
    ]

    for d_title, d_role, d_ip, d_vlan, dx, dy in mcc_drives:
        svg.append(f'''
        <g transform="translate({dx}, {dy})">
            <rect x="0" y="0" width="360" height="95" rx="5" fill="#ffffff" stroke="#fed7aa" stroke-width="1.2" />
            <text x="15" y="24" font-family="sans-serif" font-size="11" font-weight="800" fill="#0f172a">{d_title}</text>
            <text x="15" y="42" font-family="sans-serif" font-size="10" fill="#64748b">{d_role}</text>
            <rect x="15" y="58" width="160" height="22" rx="3" fill="#fef3c7" stroke="#fde68a" />
            <text x="95" y="73" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#b45309">IP: {d_ip}</text>
            <rect x="200" y="58" width="120" height="22" rx="3" fill="#ecfdf5" stroke="#a7f3d0" />
            <text x="260" y="73" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#059669">{d_vlan}</text>
        </g>
        ''')

    svg.append('</g>')

    # ====================================================
    # 5. ZONE 4: MAIN CONTROL PANEL (MCP-01) (Bottom Left)
    # ====================================================
    mcp_x, mcp_y, mcp_w, mcp_h = 50, 660, 680, 1040
    svg.append(f'''
    <g id="Zone_MCP">
        <rect x="{mcp_x}" y="{mcp_y}" width="{mcp_w}" height="{mcp_h}" rx="8" fill="#f0fdf4" stroke="#10b981" stroke-width="1.8" filter="url(#shadow)" />
        <rect x="{mcp_x}" y="{mcp_y}" width="420" height="32" rx="6" fill="#059669" />
        <text x="{mcp_x + 15}" y="{mcp_y + 21}" font-family="sans-serif" font-size="13" font-weight="700" fill="#ffffff">MAIN CONTROL PANEL (MCP-01) — CONTROLLOGIX CS1 - CS4</text>

        <!-- CS1 Master Controller -->
        <g transform="translate({mcp_x + 20}, {mcp_y + 50})">
            <rect x="0" y="0" width="{mcp_w - 40}" height="140" rx="6" fill="#ffffff" stroke="#10b981" stroke-width="1.5" />
            <text x="15" y="24" font-family="sans-serif" font-size="12" font-weight="800" fill="#0f172a">RACK 1 (CS1): MASTER CONTROLLER &amp; DLR SUPERVISOR</text>
            <text x="15" y="44" font-family="sans-serif" font-size="10" fill="#334155">Slot 00: 1756-L950TPSXT Master Controller (50MB Process Safety)</text>
            <text x="15" y="62" font-family="sans-serif" font-size="10" font-weight="700" fill="#0284c7">Slot 01: 1756-EN4TR (DLR Supervisor Node) — IP: 192.168.30.11</text>
            <text x="15" y="80" font-family="sans-serif" font-size="10" font-weight="700" fill="#be123c">Slot 02: 1756-EN4TR (Control LAN SCADA Uplink) — IP: 192.168.30.20</text>
            <text x="15" y="98" font-family="sans-serif" font-size="10" fill="#64748b">Slots 03-12: Redesigned Balanced I/O (4 DI / 2 DO / 2 AI / 2 AO)</text>
            <text x="15" y="118" font-family="monospace" font-size="10" font-weight="800" fill="#059669">Gateway VIP: 192.168.30.1</text>
        </g>

        <!-- CS2, CS3, CS4 Summaries -->
        <g transform="translate({mcp_x + 20}, {mcp_y + 210})">
            <rect x="0" y="0" width="{mcp_w - 40}" height="70" rx="4" fill="#ffffff" stroke="#cbd5e1" />
            <text x="15" y="24" font-family="sans-serif" font-size="11" font-weight="800" fill="#0f172a">RACK 2 (CS2): Expansion Rack 1 (DLR Node 1)</text>
            <text x="15" y="44" font-family="sans-serif" font-size="10" fill="#64748b">Slot 00: 1756-EN4TR (IP: 192.168.30.12) + I/O Modules &amp; Spares</text>
        </g>
        <g transform="translate({mcp_x + 20}, {mcp_y + 295})">
            <rect x="0" y="0" width="{mcp_w - 40}" height="70" rx="4" fill="#ffffff" stroke="#cbd5e1" />
            <text x="15" y="24" font-family="sans-serif" font-size="11" font-weight="800" fill="#0f172a">RACK 3 (CS3): Expansion Rack 2 (DLR Node 2)</text>
            <text x="15" y="44" font-family="sans-serif" font-size="10" fill="#64748b">Slot 00: 1756-EN4TR (IP: 192.168.30.13) + I/O Modules &amp; Spares</text>
        </g>
        <g transform="translate({mcp_x + 20}, {mcp_y + 380})">
            <rect x="0" y="0" width="{mcp_w - 40}" height="70" rx="4" fill="#ffffff" stroke="#cbd5e1" />
            <text x="15" y="24" font-family="sans-serif" font-size="11" font-weight="800" fill="#0f172a">RACK 4 (CS4): Expansion Rack 3 (DLR Node 3)</text>
            <text x="15" y="44" font-family="sans-serif" font-size="10" fill="#64748b">Slot 00: 1756-EN4TR (IP: 192.168.30.14) + I/O Modules &amp; Spares</text>
        </g>

        <!-- Technical Table: Architecture Summary & Key Rules -->
        <g transform="translate({mcp_x + 20}, {mcp_y + 470})">
            <rect x="0" y="0" width="{mcp_w - 40}" height="280" rx="6" fill="#ffffff" stroke="#cbd5e1" stroke-width="1.5" />
            <rect x="0" y="0" width="{mcp_w - 40}" height="30" rx="6" fill="#334155" />
            <text x="15" y="20" font-family="sans-serif" font-size="11" font-weight="800" fill="#ffffff">TECHNICAL ARCHITECTURE SPECIFICATIONS</text>

            <text x="15" y="55" font-family="sans-serif" font-size="10" font-weight="800" fill="#0284c7">1. CISCO CORE SWITCH HIGH AVAILABILITY (HSRP v2):</text>
            <text x="25" y="72" font-family="sans-serif" font-size="9" fill="#334155">• SW-CORE-01 is configured with Priority 110 (Active Gateway for all VLANs).</text>
            <text x="25" y="87" font-family="sans-serif" font-size="9" fill="#334155">• SW-CORE-02 is Standby (Priority 100). Automated failover time &lt; 900 ms.</text>

            <text x="15" y="112" font-family="sans-serif" font-size="10" font-weight="800" fill="#059669">2. SERVER REDUNDANCY &amp; DUAL-HOMING:</text>
            <text x="25" y="129" font-family="sans-serif" font-size="9" fill="#334155">• FT View SE Primary (SRV-01) &amp; Secondary (SRV-02) have dual cross-connected NICs.</text>
            <text x="25" y="144" font-family="sans-serif" font-size="9" fill="#334155">• Private direct Cat6A Heartbeat link on NIC 3 ensures zero split-brain.</text>

            <text x="15" y="169" font-family="sans-serif" font-size="10" font-weight="800" fill="#d97706">3. 802.1Q TRUNK TO MCC ROOM:</text>
            <text x="25" y="186" font-family="sans-serif" font-size="9" fill="#334155">• Carries VLAN 30 (Control), VLAN 50 (MCC Drives), and VLAN 99 (Management).</text>
            <text x="25" y="201" font-family="sans-serif" font-size="9" fill="#334155">• Connected to SW-MCC-01 Stratix 5700 switch feeding PowerFlex VFDs.</text>

            <text x="15" y="226" font-family="sans-serif" font-size="10" font-weight="800" fill="#be123c">4. FIBER OPTIC BACKBONE TO SERVER ROOM:</text>
            <text x="25" y="243" font-family="sans-serif" font-size="9" fill="#334155">• Redundant 10G/1G multi-strand fiber connecting Core Switches to FDF-SR01.</text>
            <text x="25" y="258" font-family="sans-serif" font-size="9" fill="#334155">• Ensures continuous access to Historian, Batch Archives, and Active Directory.</text>
        </g>
    </g>
    ''')

    # ====================================================
    # 6. MAIN CABLING RUNS (FIBER & COPPER TRUNKS)
    # ====================================================
    svg.append(f'''
    <!-- 6A. FIBER OPTIC RUN TO EXISTING SERVER ROOM -->
    <!-- From SW-CORE-01 SFP to FDF-SR01 -->
    <path d="M {c1_x + 80} {c1_y} L {c1_x + 80} 120 L {sr_x + sr_w - 60} 120 L {sr_x + sr_w - 60} {sr_y + 70}" fill="none" stroke="#db2777" stroke-width="3.5" stroke-dasharray="8,5" />
    <rect x="{sr_x + sr_w + 30}" y="107" width="180" height="24" rx="4" fill="#fdf2f8" stroke="#db2777" stroke-width="1.2" />
    <text x="{sr_x + sr_w + 120}" y="123" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#9d174d">FOB-SR-01 (10G Fiber A)</text>

    <!-- From SW-CORE-02 SFP to FDF-SR01 (Redundant Fiber Path) -->
    <path d="M {c2_x + 80} {c2_y} L {c2_x + 80} 85 L {sr_x + sr_w - 90} 85 L {sr_x + sr_w - 90} {sr_y + 70}" fill="none" stroke="#db2777" stroke-width="3.5" stroke-dasharray="8,5" />
    <rect x="{sr_x + sr_w + 240}" y="73" width="180" height="24" rx="4" fill="#fdf2f8" stroke="#db2777" stroke-width="1.2" />
    <text x="{sr_x + sr_w + 330}" y="89" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#9d174d">FOB-SR-02 (10G Fiber B)</text>

    <!-- 6B. 802.1Q TRUNK TO MCC ROOM (Routed via cable corridor to avoid servers) -->
    <!-- From SW-CORE-01 Gi1/0/24 to SW-MCC-01 Port 1 -->
    <path d="M {c1_x + 480} {c1_y + c_h} L {c1_x + 480} 460 L 1940 460 L 1940 1080 L {mcc_x + 200} 1080 L {mcc_x + 200} {mcc_y + 50}" fill="none" stroke="#d97706" stroke-width="4" />
    <rect x="1560" y="1067" width="220" height="26" rx="4" fill="#fef3c7" stroke="#d97706" stroke-width="1.5" />
    <text x="1670" y="1084" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#92400e">TRK-MCC-01 (Main 802.1Q Trunk)</text>

    <!-- From SW-CORE-02 Gi1/0/24 to SW-MCC-01 Port 2 (Standby Trunk) -->
    <path d="M {c2_x + 480} {c2_y + c_h} L {c2_x + 480} 480 L 1970 480 L 1970 1105 L {mcc_x + 500} 1105 L {mcc_x + 500} {mcc_y + 50}" fill="none" stroke="#d97706" stroke-width="3" stroke-dasharray="6,4" />
    <rect x="1560" y="1095" width="240" height="24" rx="4" fill="#fef3c7" stroke="#d97706" stroke-width="1.5" />
    <text x="1680" y="1111" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#92400e">TRK-MCC-02 (Redundant Trunk STP)</text>

    <!-- 6C. CONTROL LAN UPLINK TO MAIN CONTROL PANEL (MCP-01) -->
    <path d="M {c1_x + 30} {c1_y + c_h} L {c1_x + 30} {mcp_y + 115} L {mcp_x + mcp_w - 20} {mcp_y + 115}" fill="none" stroke="#059669" stroke-width="3.5" />
    <rect x="{mcp_x + mcp_w + 10}" y="{mcp_y + 100}" width="160" height="26" rx="4" fill="#ecfdf5" stroke="#059669" stroke-width="1.5" />
    <text x="{mcp_x + mcp_w + 90}" y="{mcp_y + 117}" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#065f46">CBL-UPL-01 (VLAN 30)</text>

    <!-- 6D. UPLINK TO CONTROL ROOM CLIENTS (VLAN 20) -->
    <path d="M {c2_x + c_w} {c2_y + 110} L {cli_x} {c2_y + 110}" fill="none" stroke="#0284c7" stroke-width="3" />
    <rect x="{c2_x + c_w + 5}" y="{c2_y + 98}" width="70" height="24" rx="4" fill="#e0f2fe" stroke="#0284c7" stroke-width="1.2" />
    <text x="{c2_x + c_w + 40}" y="{c2_y + 114}" text-anchor="middle" font-family="monospace" font-size="10" font-weight="800" fill="#0369a1">VLAN 20</text>
    ''')

    # Closing
    svg.append('</svg>')

    full_svg = f'''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {SVG_W} {SVG_H}" width="{SVG_W}" height="{SVG_H}">''' + ''.join(svg)
    return full_svg

if __name__ == '__main__':
    out_dir = os.path.dirname(os.path.abspath(__file__))
    os.makedirs(out_dir, exist_ok=True)
    
    svg_str = generate_svg()
    svg_path = os.path.join(out_dir, "Redundant_Cisco_Core_and_Server_Network_Diagram.svg")
    with open(svg_path, "w", encoding="utf-8") as f:
        f.write(svg_str)
    print(f"Saved SVG: {svg_path}")

    # Generate PDF
    pdf_doc = fitz.open(svg_path)
    pdf_bytes = pdf_doc.convert_to_pdf()
    pdf_final = fitz.open('pdf', pdf_bytes)
    pdf_path = os.path.join(out_dir, "Redundant_Cisco_Core_and_Server_Network_Diagram.pdf")
    pdf_final.save(pdf_path)
    print(f"Saved PDF: {pdf_path}")

    # Generate high-res PNG (150 DPI)
    page = pdf_final[0]
    pix = page.get_pixmap(dpi=150)
    png_path = os.path.join(out_dir, "Redundant_Cisco_Core_and_Server_Network_Diagram.png")
    pix.save(png_path)
    print(f"Saved PNG: {png_path} ({pix.width}x{pix.height})")
