---
name: spreadsheet-formula-first
description: Architectural guidelines and best practices for developing spreadsheet automations (Google Sheets, Excel, Apps Script, Python) using a "Formula-First" methodology. Mandates offloading all presentation math, relationships, overrides, running balances, and aggregations to native spreadsheet formulas, restricting script code strictly to complex data discovery, external I/O, multi-sheet reconciliation, and state-machine transactions.
---

# 📐 Spreadsheet Formula-First Engineering Guide

## 1. Core Philosophy: The Formula-First Principle

> **"If a native spreadsheet formula can calculate, derive, or aggregate a value, the automation script MUST write a dynamic formula into the sheet rather than computing a dead, hardcoded number in code."**

### Why Formula-First?
1. **Live Interactivity (Zero Script Latency)**: When users adjust an input (e.g. rate override, bonus, deduction), dependent cells and summary rows update immediately without running a script.
2. **Full Auditability**: Accountants, managers, and stakeholders can click on any cell and inspect its exact dependencies (e.g., `=B2-C2`, `=D2*IF(ISNUMBER(O2), O2, E2)`) rather than guessing how an opaque value was calculated.
3. **Execution Speed & Quota Efficiency**: Batch-writing formula patterns reduces script run times, memory usage, and execution quota consumption.
4. **Resilient Data Separation**: Code handles external APIs, file discovery, and database state machines; the spreadsheet handles calculations and display math.

---

## 2. The 3-Tier Separation of Concerns

```
┌─────────────────────────────────────────────────────────────┐
│ TIER 1: Native Spreadsheet Formulas (The Sheet Engine)      │
│ • Mathematical derivations (Gross - Tolls = Price No Toll)  │
│ • Interdependent rates (=Price * ActiveRate)               │
│ • User overrides (=IF(ISNUMBER(Override), Override, System))│
│ • Dynamic balances & fee splits                             │
│ • Summary totals (=SUM, =SUMIF, weighted averages)          │
└──────────────────────────────┬──────────────────────────────┘
                               │ Reads evaluated values (.getValues())
┌──────────────────────────────▼──────────────────────────────┐
│ TIER 2: Automation Scripts (Apps Script / Python Engine)    │
│ • External I/O (Google Drive search, file opening, APIs)    │
│ • Cross-system parsing (date normalization, multi-sheet)    │
│ • Idempotent ledger updates (Loans, Promos, Audit Logs)     │
│ • UI generation, menu dispatching, alerts, styling          │
└──────────────────────────────┬──────────────────────────────┘
                               │ Preserves & respects
┌──────────────────────────────▼──────────────────────────────┐
│ TIER 3: User Input Layer (Editable Dashboard Columns)       │
│ • Manual Override %                                         │
│ • Refunds / Additions / Custom Deductions                   │
│ • Notes & Annotations                                       │
└─────────────────────────────────────────────────────────────┘
```

---

## 3. Decision Matrix: Formula vs. Script

| Scenario / Task | Engine | Justification / Implementation |
| :--- | :---: | :--- |
| **Row Math** (e.g. `Gross - Tolls`) | **Formula** | `=B{r}-C{r}` — updates instantly if values change. |
| **Conditional Override** (e.g. manual rate) | **Formula** | `=D{r}*IF(ISNUMBER(O{r}), O{r}, E{r})` — allows real-time interactive override without script re-runs. |
| **Running Balances** (e.g. `Total - Cash + Credits - Loans`) | **Formula** | `=G{r}-H{r}+I{r}+J{r}-K{r}` — live financial reactivity. |
| **Column Totals & Summary Rows** | **Formula** | `=SUM(B2:B{r-1})`, `=IF(D{r}>0, F{r}/D{r}, "")` for weighted averages. |
| **Categorical Aggregations** | **Formula** | `=SUMIF(D:D, "Driver Balance", F:F)` — robust across dynamic rows. |
| **Drive File Discovery & Folder Traversal** | **Script** | Google Drive API, file permissions, fuzzy folder matching. |
| **Cross-Period Date Reconciliation** | **Script** | Parsing irregular headers (e.g. "Thursday, 01") into ISO dates. |
| **Permanent Ledger Mutations** | **Script** | Stamping database dates, writing loan payment histories idempotently. |
| **Dynamic Table Styling** | **Script** | Alternating row colors, borders, font weights, column widths. |

---

## 4. Key Implementation Patterns

### Pattern A: Dynamic Row Calculation in Batch Writes
When populating a table via `Range.setValues()`, generate formula strings that reference the specific row index `r = i + 2`:

```javascript
let output = driverRows.map((d, i) => {
  let r = i + 2;
  return [
    d.driver,                                               // Col A: Name
    d.gross,                                                // Col B: Gross Fare
    d.tolls,                                                // Col C: Tolls
    `=B${r}-C${r}`,                                         // Col D: Price no Toll
    d.sysRate,                                              // Col E: System Rate
    `=D${r}*IF(ISNUMBER(O${r}), O${r}, E${r})`,            // Col F: Driver Pay (Respects Override)
    `=F${r}+C${r}`,                                         // Col G: Driver Total
    d.cash,                                                 // Col H: Cash Collected
    d.debits,                                               // Col I: Credits & Debits
    d.additionalPay,                                        // Col J: User Additions
    d.loanDeduction,                                        // Col K: Loan Deduction
    `=G${r}-H${r}+I${r}+J${r}-K${r}`,                       // Col L: Driver Balance
    `=B${r}-G${r}`,                                         // Col M: Company Fee
    d.promoStatus,                                          // Col N: Promo Notes
    d.userOverride                                          // Col O: Manual Override %
  ];
});
sheet.getRange(2, 1, output.length, 15).setValues(output);
```

### Pattern B: Dynamic Sum Row with Weighted Averages
Never write hardcoded sums at the bottom of a dashboard:

```javascript
let sumRow = output.length + 2;
sheet.getRange(sumRow, 1).setValue("Total").setHorizontalAlignment("left");

let sumFormulas = [
  [
    `=SUM(B2:B${sumRow - 1})`,                      // Gross
    `=SUM(C2:C${sumRow - 1})`,                      // Tolls
    `=SUM(D2:D${sumRow - 1})`,                      // Net Price
    `=IF(D${sumRow}>0, F${sumRow}/D${sumRow}, "")`, // Weighted System Rate
    `=SUM(F2:F${sumRow - 1})`,                      // Driver Pay
    `=SUM(G2:G${sumRow - 1})`,                      // Driver Total
    `=SUM(H2:H${sumRow - 1})`,                      // Cash
    `=SUM(I2:I${sumRow - 1})`,                      // Debits
    `=SUM(J2:J${sumRow - 1})`,                      // Additions
    `=SUM(K2:K${sumRow - 1})`,                      // Loans
    `=SUM(L2:L${sumRow - 1})`,                      // Balances
    `=SUM(M2:M${sumRow - 1})`,                      // Company Fees
    "",
    ""
  ]
];
sheet.getRange(sumRow, 2, 1, 14).setFormulas(sumFormulas);
```

### Pattern C: Defending Scripts from Formula Rows
When a script subsequently reads data back from a sheet that contains formula summary rows (e.g. `Total`), **always filter out summary rows**:

```javascript
sheet.getRange(2, 1, lastRow - 1, lastCol).getValues().forEach(row => {
  let name = String(row[0] || "").trim();
  // CRITICAL DEFENSE: Never process Total/Sum rows as entity records!
  if (!name || name.toLowerCase() === "total" || name.toLowerCase().startsWith("total")) return;

  // Process valid data row...
});
```

### Pattern D: Live Hierarchical Payslip Totals
In complex multi-block payslips, use `SUM` and `SUMIF`:

```javascript
// Grand Total for daily items
let grandTotalFormula = `=D${grandTotalRow}+E${grandTotalRow}`;
totalCell.setFormula(grandTotalFormula);

// Driver Balance incorporating deductions
let balanceFormula = (lastDeductRow >= startDeductRow)
  ? `=F${grandTotalRow}+SUM(F${startDeductRow}:F${lastDeductRow})`
  : `=F${grandTotalRow}`;
balanceCell.setFormula(balanceFormula);

// Overall Sheet Grand Total using SUMIF
totalPayoutCell.setFormula(`=SUMIF(D1:D${currentRow - 1}, "Driver Balance", F1:F${currentRow - 1})`);
```

---

## 5. Checklist Before Committing Any Spreadsheet Automation

- [ ] Are simple subtractions/additions (`Gross - Tolls`) written as formulas?
- [ ] Do rate calculations dynamically observe user override cells?
- [ ] Are bottom summary rows using `=SUM(...)` and `=SUMIF(...)`?
- [ ] Does the downstream script read `.getValues()` (evaluated numbers) properly?
- [ ] Does the script guard against `"Total"` / `"Sum"` rows when reading data back?
- [ ] Are formatting and column widths applied cleanly after writing formulas?
