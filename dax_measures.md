// ============================================================
// FILE: dax_measures.dax
// PROJECT: Semi-Automated Treasury Reconciliation
// AUTHOR: Wanlaya Sirirojanan
// LAST UPDATED: May 2026
//
// PURPOSE:
// All DAX measures used in the Power BI dashboard.
// Organised by section for easy reference.
//
// HOW TO USE:
// In Power BI Desktop → Enter Data / New Measure
// Copy each measure individually into the formula bar.
// ============================================================


// ------------------------------------------------------------
// SECTION 1: BALANCE MEASURES
// Core cash position calculations used across all visuals.
// ------------------------------------------------------------

// Opening balance of the selected period
// Reads from HeaderMonthend table linked via Period Bridge
Beginning Balance =
CALCULATE(
    SUM(HeaderMonthend[Bal.Amt.Element:Text]),
    HeaderMonthend[Balance_Period] = "Beginning Balance"
)


// Closing balance of the selected period
Ending Balance =
CALCULATE(
    SUM(HeaderMonthend[Bal.Amt.Element:Text]),
    HeaderMonthend[Balance_Period] = "Ending Balance"
)


// ------------------------------------------------------------
// SECTION 2: INFLOW & OUTFLOW MEASURES
// Used in the cash flow matrix and bar charts.
// ------------------------------------------------------------

// Total inflow for the filtered period
// Only counts credit (positive) signed amounts
Total Inflow =
CALCULATE(
    SUM(Consolidation[Signed_Amount_File]),
    Consolidation[Signed_Amount_File] > 0
)


// Total outflow for the filtered period
// Returns as positive number for display purposes
Total Outflow =
CALCULATE(
    -SUM(Consolidation[Signed_Amount_File]),
    Consolidation[Signed_Amount_File] < 0
)


// Net cash = inflow minus outflow
// Excludes rows where bank amount is blank (GL-only rows)
Net cash =
CALCULATE(
    [Total Inflow] - [Total Outflow],
    NOT(ISBLANK(Consolidation[Signed_Amount_Bank]))
)


// ------------------------------------------------------------
// SECTION 3: DAILY CASH POSITION MEASURES
// Used in the daily running balance table and line chart.
// Logic relies on RowIndex from Power Query to identify
// the last transaction of each day correctly.
// ------------------------------------------------------------

// Ending cash balance at the close of each day
// Finds the highest RowIndex for the day, then returns
// the RunningBalance on that row
Daily Ending Cash =
VAR LastRow =
    CALCULATE(
        MAX(Consolidation[RowIndex]),
        ALLEXCEPT(Consolidation, Consolidation[BookingDate])
    )
RETURN
    CALCULATE(
        MAX(Consolidation[RunningBalance]),
        Consolidation[RowIndex] = LastRow,
        ALLEXCEPT(Consolidation, Consolidation[BookingDate])
    )


// Opening cash balance at the start of each day
// Returns the ending balance of the previous business day.
// Falls back to Beginning Balance if no prior day exists.
Daily Beginning Cash =
VAR CurrentDate = MIN(Consolidation[BookingDate])
VAR PrevDate =
    CALCULATE(
        MAX(Consolidation[BookingDate]),
        ALL(Consolidation[BookingDate]),
        Consolidation[BookingDate] < CurrentDate,
        NOT ISBLANK(Consolidation[BookingDate])
    )
VAR PrevBalance =
    CALCULATE(
        MAX(Consolidation[RunningBalance]),
        ALL(Consolidation[BookingDate]),
        Consolidation[BookingDate] = PrevDate,
        NOT ISBLANK(Consolidation[BookingDate])
    )
RETURN
    IF(ISBLANK(PrevDate), [Beginning Balance], PrevBalance)


// ------------------------------------------------------------
// SECTION 4: MAX & MIN CASH MEASURES
// Used in the liquidity awareness KPI cards.
// Highlights the peak and trough cash positions in the period.
// ------------------------------------------------------------

// Highest daily ending cash in the selected period
Max Ending Cash =
MAXX(
    VALUES(Consolidation[BookingDate]),
    [Daily Ending Cash]
)


// Lowest daily ending cash in the selected period
Min Ending Cash =
MINX(
    VALUES(Consolidation[BookingDate]),
    [Daily Ending Cash]
)


// Date on which the highest daily ending cash occurred
Date of Max Cash =
CALCULATE(
    MAX(Consolidation[BookingDate]),
    FILTER(
        Consolidation,
        Consolidation[RunningBalance] = [Max Ending Cash]
    )
)


// Date on which the lowest daily ending cash occurred
Date of Min Cash =
CALCULATE(
    MAX(Consolidation[BookingDate]),
    FILTER(
        Consolidation,
        Consolidation[RunningBalance] = [Min Ending Cash]
    )
)

