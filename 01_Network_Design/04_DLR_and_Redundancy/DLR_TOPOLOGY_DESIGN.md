# Device Level Ring (DLR) & Redundancy Design
**Standard**: ODVA EtherNet/IP Device Level Ring (DLR) Specification / IEC 62439-3  
**Project**: Ingredion Kalasin Spray Dryer Plant  

---

## 1. DLR Overview & Fault Recovery
DLR is a Layer 2 ring protocol designed for rapid single-fault recovery (< 3ms) in industrial automation systems. In the event of a severed cable or failed node, the Ring Supervisor unblocks its secondary port within milliseconds, preventing communication timeout (`Connection Timeout`) on the ControlLogix processor.

```
       [ 1756-EN2TR / EN4TR Bridge ] <--- Active Ring Supervisor
        Port 1                  Port 2
          │                       │ (Normally Blocked Beacon Port)
          ▼                       ▼
    [1734-AENTR] ◄────────► [PowerFlex 525] ◄────────► [1734-AENTR Node 2]
     (Ring Node)             (Dual Port VFD)            (Ring Node)
```

---

## 2. Kalasin Plant DLR Ring Assignments

### Ring 1: Area 202 Pre-Slurry DLR
* **Ring Supervisor**: 1756-EN4TR in Main Control Panel (MCP-01) Slot 2 (IP: `192.168.40.10`)
* **Backup Supervisor**: 1783-ETAP (or Stratix 5700 port pair in DLR supervisor mode)
* **Ring Nodes (Daisy-chained)**:
  1. `RIO-202-AENT1` (1734-AENTR Slot 1/2)
  2. `VFD-PC-20201` (PowerFlex 525 with 25-COMM-E2P dual Ethernet)
  3. `VFD-PC-20211` (PowerFlex 525 with 25-COMM-E2P dual Ethernet)
  4. `VFD-ME-20201` (PowerFlex 525 dual Ethernet)
  5. `RIO-202-AENT2` (1734-AENTR Slot 1/2)
* **Beacon Interval**: $400\,\mu\text{s}$
* **Beacon Timeout**: $1960\,\mu\text{s}$ (less than 2ms total failover time)

### Ring 2: Area 602 Spray Dryer Burner & Process DLR
* **Ring Supervisor**: 1756-EN4TR in MCP-602 Slot 3 (IP: `192.168.60.10`)
* **Ring Nodes**:
  1. `RIO-602-AENT1` (Chamber Instrumentation)
  2. `RIO-602-AENT2` (Cyclone & Baghouse)
  3. `VFD-PUPD-11` (PowerFlex 755 with 20-750-ENETR dual port)
  4. `VFD-FN-60201` (PowerFlex 755 dual port)
  5. `VFD-FN-60202` (PowerFlex 755 dual port)
