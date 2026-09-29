# Industrial Network FAT / SAT Commissioning Checklist
**Facility**: Ingredion Kalasin Spray Dryer Plant  
**Testing Protocols**: Physical Layer Verification, Ping Sweep, DLR Failover Timing, Multicast IGMP Verification  

---

## 1. Physical Layer Testing & Inspection
- [ ] **Fiber Optic Tier 1 Testing**: Optical Loss Test Set (OLTS) certification for all strands (`FO-TRK-01..03`). Loss < 3.0 dB end-to-end.
- [ ] **Copper Testing**: Fluke DSX Cat6A Channel Test (Wiremap, NEXT, Return Loss, Length).
- [ ] **Shielding & Grounding**: Single-point termination of shield drains verified against ground loop currents.

---

## 2. Layer 2 / Layer 3 Protocol Validation
- [ ] **VLAN Routing**: Verified ping between VLAN 10 (SCADA) and VLAN 30 (PLC) across Core Gateway.
- [ ] **VLAN Isolation**: Verified access restriction between VLAN 40 (Area 202 Field I/O) and untrusted networks.
- [ ] **IGMP Snooping**: Verified with Wireshark that CIP implicit multicast traffic does not flood non-subscriber switch ports.

---

## 3. DLR Ring Redundancy Failover Testing
- [ ] **Ring Break Test (Cable Disconnect)**:
  - Disconnect Port 1 on `PLC-01-EN4T`.
  - Record failover time (target < 3ms).
  - Verify zero I/O module fault flags on ControlLogix chassis.
- [ ] **Ring Restore Test (Cable Reconnect)**:
  - Reconnect cable.
  - Verify Ring Supervisor automatically transitions back to "Ring Normal" without communication glitch.
