// ==============================================================================
// KALASIN ENGINEERING CO., LTD. — AUTOMATION ENGINEERING DIVISION
// PROJECT: SPRINT 18K TPA SPRAY DRYER PLANT (KALASIN)
// SYSTEM: ROCKWELL CONTROLLOGIX 5580 (CHASSIS C1 - C4 IN MCP-01)
// MODULE: AOI_FAT_DIDO_Loopback (Add-On Instruction for DI/DO Loopback Testing)
// LANGUAGE: STRUCTURED TEXT (IEC 61131-3)
// ==============================================================================

/*
DESCRIPTION:
This Add-On Instruction automates the FAT closed-loop verification of Digital Output 
interposing relays and Digital Input field circuits via temporary cross-wiring.

CIRCUIT TOPOLOGY:
1756-OB32 DO Pin -> TBRL Relay Coil (A1/A2) -> Relay Contact (11/14) -> 
Temporary Loopback Jumper -> TBDI Terminal -> 1756-IB32 DI Pin.

EXECUTION LOGIC:
1. Verifies DI is OFF initially (Stuck-On Pre-Check).
2. Forces DO Output bit ON.
3. Measures time delta until DI Input turns ON (Switching Latency).
4. Holds ON for visual operator confirmation.
5. De-energizes DO Output bit.
6. Verifies DI Input drops to OFF within timeout (Stuck-Contact Post-Check).
7. Sets PASS/FAIL flags and records latency in milliseconds.
*/

// ------------------------------------------------------------------------------
// UDT DEFINITION: UDT_FAT_Loopback_Channel
// ------------------------------------------------------------------------------
/*
TYPE UDT_FAT_Loopback_Channel :
STRUCT
    Loop_ID          : STRING[12]; // e.g. 'LPB-001'
    DO_Slot_Pin      : STRING[16]; // e.g. 'C1S8.0'
    DI_Slot_Pin      : STRING[16]; // e.g. 'C1S4.0'
    Relay_Tag        : STRING[20]; // e.g. 'P1-TBRL1-1'
    Timeout_Limit_ms : DINT;       // Max allowable response time (e.g. 200 ms)
    Hold_Time_ms     : DINT;       // Dwell time for visual LED check (e.g. 300 ms)
    Latency_ms       : DINT;       // Measured contact closure latency
    State_Step       : DINT;       // Current internal sequencer state
    Pass             : BOOL;       // Test passed successfully
    Fail             : BOOL;       // Test failed
    Fault_Code       : DINT;       // 1=DI Stuck High, 2=Pickup Timeout, 3=Drop Timeout
    Completed        : BOOL;       // Channel test cycle completed
END_STRUCT
END_TYPE
*/

// ------------------------------------------------------------------------------
// ADD-ON INSTRUCTION LOGIC: AOI_FAT_DIDO_Loopback
// ------------------------------------------------------------------------------

// Input Parameters:
//   Enable_Test      : BOOL; // Master enable command
//   Trigger_Start    : BOOL; // Single-shot trigger to test this channel
//   Reset_Fault      : BOOL; // Clears latching fault flags
//   DI_Feedback_Bit  : BOOL; // Physical DI feedback bit from 1756-IB32
//   Timeout_Limit_ms : DINT; // e.g. 200 ms
//   Hold_Time_ms     : DINT; // e.g. 300 ms

// Output Parameters:
//   DO_Command_Bit   : BOOL; // Driving bit to 1756-OB32 channel
//   Test_Busy        : BOOL; // Execution currently active
//   Test_Done        : BOOL; // Cycle finished
//   Test_Pass        : BOOL; // Result: Pass
//   Test_Fail        : BOOL; // Result: Fail
//   Measured_Latency : DINT; // Response time in ms
//   Fault_Code       : DINT; // 0=None, 1=Stuck High, 2=Pickup Timeout, 3=Dropout Timeout

// State Definitions
#define STATE_IDLE           0
#define STATE_PRECHECK_DI   10
#define STATE_FIRE_DO       20
#define STATE_WAIT_DI_ON    30
#define STATE_HOLD_PULSE    40
#define STATE_RELEASE_DO    50
#define STATE_WAIT_DI_OFF   60
#define STATE_COMPLETE      70
#define STATE_FAULT         99

// Reset logic
IF Reset_Fault THEN
    Test_Fail := FALSE;
    Fault_Code := 0;
    Step := STATE_IDLE;
    DO_Command_Bit := FALSE;
END_IF;

IF NOT Enable_Test THEN
    DO_Command_Bit := FALSE;
    Test_Busy := FALSE;
    Step := STATE_IDLE;
    RETURN;
END_IF;

CASE Step OF
    STATE_IDLE:
        Test_Busy := FALSE;
        DO_Command_Bit := FALSE;
        IF Trigger_Start AND NOT Test_Done THEN
            Test_Busy := TRUE;
            Test_Done := FALSE;
            Test_Pass := FALSE;
            Test_Fail := FALSE;
            Fault_Code := 0;
            Measured_Latency := 0;
            Step := STATE_PRECHECK_DI;
        END_IF;

    STATE_PRECHECK_DI:
        // Safety Pre-Check: DI must be 0 prior to energizing relay
        IF DI_Feedback_Bit THEN
            // DI is already energized without DO signal -> Field short or stuck high!
            Fault_Code := 1; // DI Stuck High Error
            Test_Fail := TRUE;
            Step := STATE_FAULT;
        ELSE
            Timer_Latency.PRE := Timeout_Limit_ms;
            Timer_Latency.RESET := 1;
            Timer_Latency.RESET := 0;
            Step := STATE_FIRE_DO;
        END_IF;

    STATE_FIRE_DO:
        // Command DO Relay to fire
        DO_Command_Bit := TRUE;
        Timer_Latency.TimerEnable := 1;
        Step := STATE_WAIT_DI_ON;

    STATE_WAIT_DI_ON:
        // Wait for dry contact to close and energize DI
        TONR(Timer_Latency);
        
        IF DI_Feedback_Bit THEN
            // Success: DI feedback detected!
            Measured_Latency := Timer_Latency.ACC;
            Timer_Hold.PRE := Hold_Time_ms;
            Timer_Hold.RESET := 1;
            Timer_Hold.RESET := 0;
            Step := STATE_HOLD_PULSE;
        ELSIF Timer_Latency.DN THEN
            // Timeout reached without DI feedback
            DO_Command_Bit := FALSE;
            Fault_Code := 2; // Pickup Timeout (Relay coil failed or wiring broken)
            Test_Fail := TRUE;
            Step := STATE_FAULT;
        END_IF;

    STATE_HOLD_PULSE:
        // Hold relay ON for visual verification on panel LED
        TONR(Timer_Hold);
        IF Timer_Hold.DN THEN
            DO_Command_Bit := FALSE; // Turn off DO
            Timer_Dropout.PRE := Timeout_Limit_ms;
            Timer_Dropout.RESET := 1;
            Timer_Dropout.RESET := 0;
            Step := STATE_WAIT_DI_OFF;
        END_IF;

    STATE_WAIT_DI_OFF:
        // Verify relay contact drops out cleanly and DI turns OFF
        TONR(Timer_Dropout);
        IF NOT DI_Feedback_Bit THEN
            // Relay opened successfully (No contact weld)
            Test_Pass := TRUE;
            Step := STATE_COMPLETE;
        ELSIF Timer_Dropout.DN THEN
            // Contact stuck or welded closed!
            Fault_Code := 3; // Dropout Timeout (Contact Welded)
            Test_Fail := TRUE;
            Step := STATE_FAULT;
        END_IF;

    STATE_COMPLETE:
        DO_Command_Bit := FALSE;
        Test_Busy := FALSE;
        Test_Done := TRUE;
        IF NOT Trigger_Start THEN
            Step := STATE_IDLE;
        END_IF;

    STATE_FAULT:
        DO_Command_Bit := FALSE;
        Test_Busy := FALSE;
        Test_Done := TRUE;
        // Latch in fault state until Reset_Fault or Trigger_Start cleared
        IF NOT Trigger_Start THEN
            Step := STATE_IDLE;
        END_IF;
END_CASE;
