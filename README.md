**Semi-Automated Treasury Reconciliation \& Cash Flow Dashboard**

**Author:** Wanlaya Sirirojanan

**Last Updated:** May 2026

**Tools:** Excel · Power Query (M Code) · Power BI (DAX) · CAMT.053 (ISO 20022) · XML

\---

## The Problem

In most small and mid-size companies, the treasury analyst starts every morning
the same way — opening a bank statement, copying transactions manually into
Excel, matching them against accounting records and other internal information
by hand, then building a cash position report for the morning meeting.

It works. But it takes 2–3 hours that should not need to be spent that way. Large multinationals solve this with a Treasury Management System (TMS) , such as WallstreetSuite, Kyriba. A TMS automatically receives bank statement files (MT940 or CAMT.053), reconciles transactions, and updates the cash position in real time. But a TMS costs significantly more than most companies can justify to implement and maintain.

This project fills the gap between "manual copy-paste every morning" and "full TMS implementation.

\---

## What This Project Does

Using only Excel, Power Query, and Power BI — tools available on every corporate
laptop — this model automates the three tasks a treasury analyst does every day:

**1. Bank Reconciliation**

Matches every bank transaction against the general ledger automatically, using
transaction reference numbers. Flags unmatched items and identifies whether the
gap sits on the bank side or the accounting side, so the right team knows what
to action.

**2. Mini Direct Cash Flow Report**

Classifies every transaction into operating or financing, and into specific
categories (domestic revenue, supplier payments, VAT, salaries, loan drawdown,
etc.). Builds the daily and weekly cash flow statement automatically.

**3. Cash Position Dashboard**

Shows opening and closing balance, net cash for the period, daily inflow and
outflow, and highlights the highest and lowest daily cash positions for
liquidity awareness — all refreshed from the latest bank statement with a
single click.

\---

## Dashboard Preview

![Overview](https://github.com/wsiriroj/treasury-recon/blob/main/Screenshots/overview_dashboard.png)



![Daily Cash Flow](https://github.com/wsiriroj/treasury-recon/blob/main/Screenshots/cashflow_daily.png)



![Reconciliation Status](https://github.com/wsiriroj/treasury-recon/blob/main/Screenshots/reconciliation_status.png)

\---

## How It Works

The process has three steps that run automatically once set up:

**Step 1 — File drop**

The treasury analyst saves the new bank statement file into a designated folder.
No opening, no copying, no formatting.

**Step 2 — Refresh**

Opening the Excel workbook and clicking Refresh triggers Power Query to read
the latest files, parse the bank statement, match it against the GL, classify
every transaction, and calculate the running cash balance.

**Step 3 — Review**

The Power BI dashboard updates. The analyst reviews the reconciliation status,
checks the cash position, and focuses attention only on the unmatched items
that need follow-up — typically 3–5 transactions out of 20+.

**Total time from bank statement to reviewed cash position: under one hour.**

*Parsing and refresh takes under 5 minutes. Remaining time is analyst review
of flagged items only.*

\---

## Reconciliation Logic

Matching runs in two passes:

* **Full reference match** — the bank’s end-to-end transaction ID matches the
document reference in the GL exactly. Highest confidence.
* **Key reference match** — a shortened reference key extracted from the
transaction description matches the GL. Catches cases where the reference
format differs slightly between systems.

Unmatched transactions are automatically labelled by responsible party:

* *Pending process by GL* — the bank has the transaction; accounting has not
booked it yet
* *Pending process by Bank* — the GL has an accrual; the bank settlement has
not arrived yet

This tells each team exactly what they own, without the treasury analyst having
to chase anyone manually.

\---

## Business Context — Mock Data

The model was built and tested using a mock company: **ABC Co Ltd**, a Thai
import-export T-shirt trading business with one THB bank account, overseas
supplier payments, FX conversion activity, a short-term non-revolving loan,
and a time deposit placement. Three months of data (April–June 2026) were used.

The mock data was generated to reflect realistic treasury scenarios including:

* A cash shortfall requiring a short-term borrowing in May
* A time deposit placement in June when cash surplus emerged
* Cross-border payments requiring currency conversion
* Month-end accruals not yet settled by the bank

\---

## Project Structure

* README.md
* reconciliation.m
* dax\_measures.dax
* PowerBI/

  * Treasury\_Dashboard.pbix
* Screenshots/

  * overview\_dashboard.png
  * cashflow\_daily.png
  * reconciliation\_status.png
* SampleData/

\---

## Limitations \& What Is Next

**Current version**

* One company, one bank account, one currency
* Matching logic handles standard reference-based reconciliation — complex
cases (split payments, partial settlements) are flagged for manual review
* The direct cash flow report is a simplified version — it does not consolidate
all cash positions into a single summary table. Users need to review both
the summary and detail views together
* Refresh is triggered manually by the user — save new files and click Refresh
when an update is required

**Next version**

* Multi-entity and multi-bank-account support
* Multi-currency with live FX rate feed
* Summary and status of FX, loan, and investment positions
* Variance analysis month over month
* Variance analysis against cash flow forecast
* Higher-volume matching using SQL in the ETL layer

\---

## Why This Project

This project reflects my experience in corporate treasury, a frustration I encountered repeatedly across teams.

Most treasury and finance teams cannot solve technical problems on their own. When they need support, they join a queue. A simple automation request that should take a week ends up taking three months, waiting behind other projects from other departments.
The goal was to change that. If a treasury analyst can solve the problem themselves — using tools already on their laptop, without writing a single line of code — they do not need to wait. The work gets done in weeks, not quarters.
This project is that solution. Built by a treasury person, for treasury people.

It does not replace a TMS and was never designed to. But it reduces the daily burden — filtering transactions, flagging exceptions, and delivering a cash position in under an hour. For a team that previously spent their morning on manual work, that is enough to matter.

*All business logic, mapping process and reconciliation design, and analytical decisions were
defined by the author.*

*AI assistance (Claude) was used for creating the mock data and technical
implementation of complex syntax — the same way a treasury analyst would
engage an internal IT team.*

