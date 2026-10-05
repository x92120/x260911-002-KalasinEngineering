# 07_Presentations_and_Overviews — Control Strategy & Architecture Presentations

This directory contains technical presentations, slides, and engineering overviews for the **Ingredion Kalasin Spray Dryer Plant** PID loop control strategy and automation system.

---

## 📑 File Inventory

* 📊 **`PID_Loop_Control_Specification_Kalasin.pptx`**: PowerPoint presentation specifying PID loop control strategies, cascade control loops, feedforward compensation, and tuning parameters for the Jet Cooker and Spray Dryer processes.
* 🌐 **`PID_Loop_Control_Architecture_Overview.html`**: Standalone HTML interactive presentation viewer detailing PID control loop topology, setpoint ranges, transmitter tags, and actuator feedback loops.

---

## 🎯 Control Strategy Scope

1. **Jet Cooker Steam Temperature Cascade Control**: Primary loop controlling slurry output temp (`TIC-40201`), secondary inner loop controlling steam valve position (`PIC-40202`).
2. **Spray Dryer Tower Feed Flow & Atomizer Control**: Slurry feed pump VFD control (`FIC-60201`) with safety interlocks linked to drying chamber inlet temp (`TIC-60202`).
3. **Baghouse Pulse-Jet Cleaning Control**: Differential pressure-based auto pulsing (`DPIC-60203`).
