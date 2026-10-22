---
name: ayoah-payment-system
description: Complete standard operating guide, business rules, mathematical models, and automation architecture for the AYOAH Driver Payroll & Payment System. Use whenever generating payslips, calculating driver balances, managing loan repayments, applying promo rates, parsing monthly daily balance spreadsheets, or setting up payroll dashboards for AYOAH Go.
---

# 🚗 AYOAH Driver Payment & Payroll System Guide

## 1. Executive Summary & Architecture

The **AYOAH Driver Payment System** is an automated payroll engine designed for ride-hailing and fleet operations (AYOAH Go). It processes weekly driver earnings, tolls, cash reconciliations, advances, credits/debits, multi-week loan amortizations, and promotional incentive tiers across Google Sheets and Google Drive.

```
                      ┌──────────────────────────────────────┐
                      │    AYOAH GO Drivers — payables       │
                      │  (Weekly aggregate payout summary)   │
                      └──────────────────┬───────────────────┘
                                         │
┌───────────────────────────┐            │            ┌───────────────────────────┐
│ Google Drive: 2026 Folder │            │            │  Central Master Database  │
│ [Month] - Drivers Daily   │◄───────────┼───────────►│  • Settings (Promos)      │
│ Balance (Daily Breakdown) │            │            │  • Loans (Amortization)   │
└─────────────┬─────────────┘            │            └─────────────┬─────────────┘
              │                          ▼                          │
              │         ┌───────────────────────────────────┐       │
              └────────►│      STEP 1: Local Settings       │◄──────┘
                        │    (Interactive Dashboard Tab)    │
                        └─────────────────┬─────────────────┘
                                          │
                                          ▼
                        ┌───────────────────────────────────┐
                        │      STEP 2: PaySlip Sheet        │
                        │ (Detailed Multi-Driver Payslips)  │
                        └─────────────────┬─────────────────┘
                                          │ Idempotent Loan Updates
                                          ▼
                        ┌───────────────────────────────────┐
                        │    Master DB Loans Stamped        │
                        └───────────────────────────────────┘
```

---

## 2. Core Mathematical Formulas & Accounting Rules

### A. Pure Trip Gross Fare vs. Credits/Debits (Crucial Rule)
* **Rule:** Credits and Debits have **nothing** to do with Gross Trip Fares. Company service fees are **never** taken from credits or debits.
* On the Payables summary sheet, `Final Pay` already has credits added and debits subtracted ($\text{Final Pay} = \text{Trip Payout} + \text{Credits} - \text{Debits}$).
* **Pure Trip Gross Fare Calculation:**
  $$\text{Gross Fare} = \text{Cash} + \text{Final Pay} + |\text{Fees}| - |\text{Credits}| + |\text{Debits}|$$
* **Price no Toll (Net Trip Fare):**
  $$\text{Price no Toll} = \text{Gross Fare} - \text{Tolls}$$

### B. Driver Payout Rates
* **Standard System Rate:** **90%** (0.90) of Price no Toll $+$ 100% of Tolls.
* **Promo System Rate:** **95%** (0.95) of Price no Toll $+$ 100% of Tolls for up to **4 consecutive weeks**.
* **Tiered Drivers (e.g. Angel Yoy):**
  * $80\%$ on Price no Toll up to $\$1,000$.
  * $90\%$ on Price no Toll exceeding $\$1,000$.
  * Effective Rate: $\frac{(1000 \times 0.80) + ((\text{Price no Toll} - 1000) \times 0.90)}{\text{Price no Toll}}$.
* **Manual Override %:** If an operator enters an override rate (e.g. `92%` or `0.92`) in Column O of `Settings`, the formula immediately prioritizes the override over the system rate.

### C. Driver Total & Company Fee
* **Driver Pay:** $\text{Price no Toll} \times \text{Active Rate}$
* **Driver Total:** $\text{Driver Pay} + \text{Tolls}$
* **Company Fee:** $\text{Gross Fare} - \text{Driver Total}$

### D. Final Driver Balance (Take-Home Payout)
Credits and Debits are added **exactly once at the very end** to the balance:
$$\text{Driver Balance} = \text{Driver Total} - \text{Cash Collected} + (\text{Credits} - \text{Debits}) + \text{Refunds/Additions} - \text{Advances} - \text{Loan Deductions}$$

---

## 3. Promotional Incentive Rules (Master DB Settings)

1. **Starting Date Requirement:** Every driver must have a starting date recorded in Column B of the Master Database `Settings` sheet (format: `MM/DD/YYYY`). If any driver lacks a starting date, the script halts and notifies the operator.
2. **2-Monday Expiration Rule:**
   * A driver is eligible for the 95% promo for their first 4 payment weeks.
   * **Expiration Gate:** If the driver's starting date is **more than 2 Mondays before the payment Saturday**, the driver is permanently categorized as `None/Expired` (Standard 90% rate).
3. **Idempotent 4-Week Stamping:**
   * The Master DB records promo payments in Columns C through F (Payment 1 to 4).
   * Stamped with the payment Saturday's date (`MM/DD/YYYY`).
   * Re-running payroll during the same week does not double-stamp or consume multiple promo slots.

---

## 4. Advances & Loan Ledger Management

### A. Advances Sheet
* Local tab: `Advances` (Columns: `Driver Name`, `Amount`, `Fees`).
* **3% Advance Fee:** Managed via native formula: `=IF(B2<>"", B2*0.03, "")`.
* Both Amount and Fee are deducted from payout prior to loan eligibility checks.

### B. Master Database Loans (Amortization Ledger)
* Master DB tab: `Loans` (Columns A–H: `Driver Name`, `Loan Date`, `Total Loan`, `Weekly Installment`, `Paid So Far`, `Last Payment Date`, `Last Payment Amount`, `Remaining Balance`).
* **Buffer Rule:** A loan is only eligible for deduction if the Payslip Saturday is strictly greater than the Loan Issue Saturday (`stampDate > loanSaturday`).
* **Funds Availability Priority:**
  1. Driver Total $-$ Cash Collected $+$ Credits $-$ Debits $-$ Advances $+$ Additions = **Available Funds**.
  2. If $\text{Available Funds} > 0$, deduction is $\min(\text{Weekly Installment}, \text{Remaining Balance}, \text{Available Funds})$.
  3. A driver's balance is never forced below zero by loan repayments.
* **Idempotency:** If `Last Payment Date` equals the current pay week Saturday, the system locks the deducted amount and does not double-charge on subsequent script runs.

---

## 5. Google Drive Daily Balance File Integration

### File & Folder Hierarchy
* **Drive Base Folder:** `DAILY_BALANCE_BASE_FOLDER_ID` (`1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg`)
* **Year Folder:** E.g., `2026`
* **Monthly File:** E.g., `October - Drivers Daily Balance`
* **Target Sheet Tab:** `Summary`

### Column Structure (Strict Columns A through K)
Data **must** be isolated to Columns 1 through 11:
* **Col A (0):** `Date` (e.g. `Thursday, 01`, `Monday, 05` — merged cell with blank cells below)
* **Col B (1):** `Driver`
* **Col C (2):** `Gross Fare`
* **Col D (3):** `Hours`
* **Col E (4):** `Driver Avg/ Hour`
* **Col F (5):** `CPH`
* **Col G (6):** `Trips`
* **Col H (7):** `Cash`
* **Col I (8):** `Tolls`
* **Col J (9):** `Credit/Debit`
* **Col K (10):** `Balance`

> ⚠️ **CRITICAL ISOLATION RULE:** Never scan past Column K (index 10). Columns M+ often contain pivot summaries and charts with headers like `"Driver"`. If scanned, secondary headers will overwrite Column B mapping and corrupt payroll totals.

### Trip Discrepancy Defense Guard
If the parsed daily balance trips do not match the verified `AYOAH GO Drivers — payables` trips:
* The system triggers a safety alert listing the exact discrepancies.
* **Safety Action:** Summary totals from the Payables sheet are preserved to prevent erroneous payouts.

---

## 6. Formula-First Spreadsheet Architecture

Whenever generating or updating sheets, write **dynamic formulas** rather than static numbers:

### In `Settings` Sheet:
* **Price no Toll (Col D):** `=B{r}-C{r}`
* **Driver Pay (Col F):** `=D{r}*IF(ISNUMBER(O{r}), O{r}, E{r})`
* **Driver Total (Col G):** `=F{r}+C{r}`
* **Driver Balance (Col L):** `=G{r}-H{r}+I{r}+J{r}-K{r}`
* **Company Fee (Col M):** `=B{r}-G{r}`
* **Total Row:** `=SUM(...)` and weighted effective rate `=IF(D{sum}>0, F{sum}/D{sum}, "")`.

### In `PaySlip` Sheet:
* **Daily Total:** `=D{r}+E{r}`
* **Grand Total Row:** `=SUM(...)` across daily rows for Trips, Price, and Tolls.
* **Driver Total Box:** `=D{grandRow}+E{grandRow}`
* **Driver Balance:** `=F{grandRow}+SUM(F{firstDeduct}:F{lastDeduct})`
* **TOTAL PAYOUT:** `=SUMIF(D1:D{lastRow}, "Driver Balance", F1:F{lastRow})`

---

## 7. Operational Runbook

```
Step 1: Open Google Sheet -> 'AYOAH Automation -> 1. Setup Advances & Local Settings'
        • Check Settings tab
        • Verify promo statuses
        • Enter any Manual Override % in Col O or Refunds in Col J
        • Review live Totals at bottom of Settings

Step 2: 'AYOAH Automation -> 🧪 Test Daily Balance Reader'
        • Confirms 100% trip match with October Daily Balance file
        • Inspect 'Daily_Balance_Test_Report' tab

Step 3: 'AYOAH Automation -> 2. Generate Payslip'
        • Generates individual payslips on 'PaySlip' tab
        • Auto-updates Master DB Loans idempotently
        • Ready for distribution and driver payouts!
```
