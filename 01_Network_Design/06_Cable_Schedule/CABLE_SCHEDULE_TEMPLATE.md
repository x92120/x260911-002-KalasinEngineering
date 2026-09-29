# Industrial Network Cable Schedule Template
**Facility**: Ingredion Kalasin Spray Dryer Plant  
**Cable Types**:
* **Fiber Optic Backbone**: 12-Core Single-mode OS2 / Multimode OM3 Armored Loose-Tube with LC Duplex connectors.
* **Industrial Copper**: Cat6A Shielded S/FTP Industrial TPE/PUR Jacket (teal color), RJ45 / M12 D-Code / X-Code.

---

## Cable Run Schedule — 5-Chassis ControlLogix DLR Ring & Fiber Optic

| Cable Tag | Origin (From Panel / Port) | Destination (To Panel / Port) | Cable Type / Media | Est. Length | Route / Conduit | Signal & Function |
|---|---|---|---|---|---|---|
| `CBL-DLR-01` | MCP-01 / `CS1-EN4TR` (Port 1) | MCP-01 / `CS2-EN4TR` (Port 1) | Cat6A Shielded S/FTP | 1.5 m | Panel Internal | DLR Ring Segment 1 |
| `CBL-DLR-02` | MCP-01 / `CS2-EN4TR` (Port 2) | MCP-01 / `CS3-EN4TR` (Port 1) | Cat6A Shielded S/FTP | 1.5 m | Panel Internal | DLR Ring Segment 2 |
| `CBL-DLR-03` | MCP-01 / `CS3-EN4TR` (Port 2) | MCP-01 / `CS4-EN4TR` (Port 1) | Cat6A Shielded S/FTP | 1.5 m | Panel Internal | DLR Ring Segment 3 |
| `CBL-DLR-04` | MCP-01 / `CS4-EN4TR` (Port 2) | MCP-01 / `MC-01` (RJ45-1) | Cat6A Shielded S/FTP | 2.0 m | Panel Internal | DLR to Fiber Converter |
| `FOB-RIO-01` | MCP-01 / `MC-01` (FO-1 TX/RX) | RIO200 / `MC-02` (FO-1 TX/RX) | 4-Core Armored OM3/OS2 (Core 1-2) | 200 m | Tray / Conduit | DLR Outbound Fiber Trunk |
| `CBL-DLR-05` | RIO200 / `MC-02` (RJ45-1) | RIO200 / `CS5-EN4TR` (Port 1) | Cat6A Shielded S/FTP | 1.5 m | Panel Internal | DLR to Remote Chassis CS5 |
| `CBL-DLR-06` | RIO200 / `CS5-EN4TR` (Port 2) | RIO200 / `MC-02` (RJ45-2) | Cat6A Shielded S/FTP | 1.5 m | Panel Internal | DLR from Remote Chassis CS5 |
| `FOB-RIO-02` | RIO200 / `MC-02` (FO-2 TX/RX) | MCP-01 / `MC-01` (FO-2 TX/RX) | 4-Core Armored OM3/OS2 (Core 3-4) | 200 m | Tray / Conduit | DLR Return Fiber Trunk |
| `CBL-DLR-07` | MCP-01 / `MC-01` (RJ45-2) | MCP-01 / `CS1-EN4TR` (Port 2) | Cat6A Shielded S/FTP | 2.0 m | Panel Internal | Closes DLR to Supervisor Beacon |
| `CBL-UPL-01` | MCP-01 / `CS1-EN4TR` (Slot 2) | Server Rack / Core Switch | Cat6A / 1Gbps Fiber | 15 m | Cable Tray | Plant SCADA / FT View Uplink |
