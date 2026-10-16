// ==========================================
// ⚙️ SYSTEM CONFIGURATION
// Your Master Database Link is permanently saved here:
// ==========================================
const MASTER_DB_URL = "https://docs.google.com/spreadsheets/d/1RJ-4CCGSqYm7fAM5XkQ8rFDzKyZOiADf-H8_OX8DxJI/edit";

/**
 * Standard menu creation function.
 * Runs automatically when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('AYOAH Automation')
    .addItem('1. Setup Advances & Local Settings', 'setupInitialSheets')
    .addSeparator()
    .addItem('2. Generate Payslip', 'runCompletePayroll')
    .addToUi();
}

/**
 * Helper: Opens the Master Database using the hardcoded link above.
 */
function getMasterDatabase() {
  try {
    return SpreadsheetApp.openByUrl(MASTER_DB_URL);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error: Could not open Master Database. Please check if the link at the top of the script is correct.");
    return null;
  }
}

/**
 * Helper: Cleans text strings with $ or commas into pure math numbers
 */
function parseMoney(val) {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  return Number(String(val).replace(/[^0-9.-]+/g, "")) || 0;
}

/**
 * Helper: Parses any date input safely without timezone shifting issues.
 */
function parseTripDate(val, timeZone) {
  if (!val) return null;
  let tz = timeZone || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  if (val instanceof Date) {
    let str = Utilities.formatDate(val, tz, "yyyy-MM-dd");
    let parts = str.split("-");
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0, 0);
  }
  if (typeof val === 'string') {
    let s = val.trim();
    let match = s.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    if (match) {
      return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), 12, 0, 0, 0);
    }
  }
  let dObj = new Date(val);
  if (!isNaN(dObj.getTime())) {
    let str = Utilities.formatDate(dObj, tz, "yyyy-MM-dd");
    let parts = str.split("-");
    return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]), 12, 0, 0, 0);
  }
  return null;
}

/**
 * Helper: Finds the Saturday of the week a date falls in (for the Buffer Rule)
 */
function getSaturdayOfDate(d) {
  let dt = parseTripDate(d) || new Date(d);
  let day = dt.getDay(); // 0 is Sunday, 6 is Saturday
  let diff = day === 0 ? -1 : 6 - day; // Group Sunday with the previous Saturday
  dt.setDate(dt.getDate() + diff);
  dt.setHours(0, 0, 0, 0);
  return dt;
}

/**
 * Helper: Parses flexible date strings (e.g. 10/5/2026, 2026-10-05) into a normalized Date object at noon.
 */
function parseFlexibleDate(str) {
  if (!str) return null;
  str = String(str).trim();
  let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let month = Number(m[1]);
    let day = Number(m[2]);
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    return new Date(year, month - 1, day, 12, 0, 0, 0);
  }
  let mIso = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (mIso) {
    return new Date(Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]), 12, 0, 0, 0);
  }
  let d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Helper: Parses pay period text like "Period: 10/5/2026 – 10/11/2026"
 */
function parsePayPeriodText(cellText) {
  if (!cellText) return null;
  let match = cellText.match(/(\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4})\s*[\s\-\u2013\u2014to]+\s*(\d{1,4}[\/\-]\d{1,2}[\/\-]\d{1,4})/i);
  if (match) {
    let s = parseFlexibleDate(match[1]);
    let e = parseFlexibleDate(match[2]);
    return { start: s, end: e, raw: cellText };
  }
  return null;
}

/**
 * Helper: Computes the Saturday stamp date for promo/loan tracking from period end date
 */
function getSaturdayStamp(endDate, timeZone) {
  let tz = timeZone || SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  let dayOfWeek = endDate.getDay();
  let daysToSunday = (7 - dayOfWeek) % 7;
  let endOfWeekSunday = new Date(endDate);
  endOfWeekSunday.setDate(endDate.getDate() + daysToSunday);

  let stampDate = new Date(endOfWeekSunday);
  stampDate.setDate(endOfWeekSunday.getDate() + 6);
  stampDate.setHours(0, 0, 0, 0);

  let stampString = Utilities.formatDate(stampDate, tz, "MM/dd/yyyy");
  return { stampDate, stampString };
}

/**
 * Helper: Computes the date of 2 Mondays before a given Saturday payment date
 */
function getTwoMondaysBefore(date) {
  let d = new Date(date);
  d.setHours(0, 0, 0, 0);
  let mondaysFound = 0;
  while (mondaysFound < 2) {
    d.setDate(d.getDate() - 1);
    if (d.getDay() === 1) { // 1 is Monday
      mondaysFound++;
    }
  }
  return d;
}

/**
 * Helper: Locates the payables summary sheet
 */
function findPayablesSheet(ss) {
  let named = ss.getSheetByName("AYOAH GO Drivers — payables");
  if (named) return named;

  for (let s of ss.getSheets()) {
    let n = s.getName().toLowerCase();
    if (n.includes("payable") || n.includes("driver") || n.includes("ayoah")) {
      return s;
    }
  }

  let active = ss.getActiveSheet();
  if (active && !["Advances", "Settings", "PaySlip"].includes(active.getName())) {
    return active;
  }

  for (let s of ss.getSheets()) {
    if (!["Advances", "Settings", "PaySlip"].includes(s.getName())) {
      return s;
    }
  }
  return ss.getSheets()[0];
}

/**
 * Helper: Validates whether a sheet has the payables table structure
 */
function isPayablesSheet(sheet) {
  if (!sheet) return false;
  let lastRow = Math.min(sheet.getLastRow(), 10);
  let lastCol = Math.min(sheet.getLastColumn(), 12);
  if (lastRow < 3) return false;

  let sample = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  for (let r = 0; r < sample.length; r++) {
    let rowStr = sample[r].map(c => String(c || "").trim().toLowerCase()).join(" ");
    if (rowStr.includes("driver net") || (rowStr.includes("trips") && rowStr.includes("toll")) || rowStr.includes("period:")) {
      return true;
    }
  }
  return false;
}

/**
 * Helper: Parses the summary payables sheet
 */
function parsePayablesSheet(sheet, tz) {
  let lastRow = sheet.getLastRow();
  let lastCol = Math.max(sheet.getLastColumn(), 10);
  if (lastRow < 5) return null;

  let allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  // 1. Pay Period
  let period = null;
  for (let r = 0; r < Math.min(allValues.length, 5); r++) {
    for (let c = 0; c < (allValues[r] || []).length; c++) {
      let cellText = String(allValues[r][c] || "").trim();
      if (/period[:\s]/i.test(cellText)) {
        let p = parsePayPeriodText(cellText);
        if (p) { period = p; break; }
      }
    }
    if (period) break;
  }

  // 2. Saturday stamp
  let stampDate = null;
  let stampString = "";
  if (period && period.end) {
    let sat = getSaturdayStamp(period.end, tz);
    stampDate = sat.stampDate;
    stampString = sat.stampString;
  }

  // 3. Header row detection
  let headerRowIdx = -1;
  let colMap = {
    driverId: -1, name: -1, trips: -1, driverNet: -1,
    fees: -1, tollways: -1, credits: -1, debits: -1,
    finalPay: -1, cash: -1
  };

  for (let r = 0; r < Math.min(allValues.length, 10); r++) {
    let row = (allValues[r] || []).map(cell => String(cell || "").trim().toLowerCase());
    if (row.includes("driver id") || (row.includes("name") && (row.includes("driver net") || row.includes("trips")))) {
      headerRowIdx = r;
      row.forEach((h, idx) => {
        if (h === "driver id") colMap.driverId = idx;
        else if (h === "name" || h === "driver name") colMap.name = idx;
        else if (h === "trips" || h === "num of trips") colMap.trips = idx;
        else if (h === "driver net" || h === "net") colMap.driverNet = idx;
        else if (h === "fees" || h === "fee") colMap.fees = idx;
        else if (h.includes("toll")) colMap.tollways = idx;
        else if (h === "credits" || h === "credit") colMap.credits = idx;
        else if (h === "debits" || h === "debit") colMap.debits = idx;
        else if (h.includes("final pay")) colMap.finalPay = idx;
        else if (h === "cash" || h.includes("cash")) colMap.cash = idx;
      });
      break;
    }
  }

  // Fallbacks based on standard column layout (Cols A-J)
  if (colMap.driverId === -1) colMap.driverId = 0;
  if (colMap.name === -1) colMap.name = 1;
  if (colMap.trips === -1) colMap.trips = 2;
  if (colMap.driverNet === -1) colMap.driverNet = 3;
  if (colMap.fees === -1) colMap.fees = 4;
  if (colMap.tollways === -1) colMap.tollways = 5;
  if (colMap.credits === -1) colMap.credits = 6;
  if (colMap.debits === -1) colMap.debits = 7;
  if (colMap.finalPay === -1) colMap.finalPay = 8;
  if (colMap.cash === -1) colMap.cash = 9;

  let driverInfo = {};
  let driverDebits = {};

  let startRow = headerRowIdx !== -1 ? headerRowIdx + 1 : 4;
  for (let r = startRow; r < allValues.length; r++) {
    let row = allValues[r] || [];
    let driverName = String(row[colMap.name] || "").trim();
    if (!driverName) {
      let idCheck = String(row[colMap.driverId] || "").trim();
      if (idCheck.toLowerCase().startsWith("total")) break;
      continue;
    }
    if (driverName.toLowerCase() === "total" || driverName.toLowerCase().startsWith("total")) {
      break;
    }

    let driverId = String(row[colMap.driverId] || "").trim();
    let trips = Number(row[colMap.trips]) || 0;
    let netFromCol = parseMoney(row[colMap.driverNet]);
    let fees = parseMoney(row[colMap.fees]);
    let tolls = parseMoney(row[colMap.tollways]);
    let credits = parseMoney(row[colMap.credits]);
    let debits = parseMoney(row[colMap.debits]);
    let finalPay = parseMoney(row[colMap.finalPay]);
    let cash = parseMoney(row[colMap.cash]);

    // Gross fare calculation requested by user:
    // "Gross fare in the new sheet is actually not included, you have to calculate it by adding Cash, Final Pay and Fees amounts together."
    let gross = cash + finalPay + Math.abs(fees);
    if (gross === 0 && netFromCol > 0) {
      gross = netFromCol + tolls;
    }
    // Price no Toll = Gross Fare - Tolls
    let net = gross - tolls;
    let netCreditsDebits = Math.abs(credits) - Math.abs(debits);

    driverInfo[driverName] = {
      id: driverId,
      name: driverName,
      trips: trips,
      gross: gross,
      tolls: tolls,
      net: net,
      cash: cash,
      credits: credits,
      debits: debits,
      netCreditsDebits: netCreditsDebits,
      earliestDate: period ? period.start : new Date()
    };

    driverDebits[driverName] = netCreditsDebits;
  }

  return { period, stampDate, stampString, driverInfo, driverDebits };
}

/**
 * Step 1: Creates the Local Dashboard AND Stamps the Master Database (Promos)!
 */
function setupInitialSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. SETUP LOCAL ADVANCES SHEET ---
  const advName = "Advances";
  let advSheet = ss.getSheetByName(advName);
  if (!advSheet) {
    advSheet = ss.insertSheet(advName);
    const headers = [["Driver Name", "Amount", "Fees"]];
    advSheet.getRange("A1:C1").setValues(headers).setFontWeight("bold").setBackground("#EFEFEF");
    advSheet.getRange("B2:C").setNumberFormat("$#,##0.00");

    let formulas = [];
    for (let i = 2; i <= 501; i++) {
      formulas.push([`=IF(B${i}<>"", B${i}*0.03, "")`]);
    }
    advSheet.getRange("C2:C501").setFormulas(formulas);
    advSheet.setColumnWidth(1, 180); advSheet.setColumnWidth(2, 100); advSheet.setColumnWidth(3, 100);
  }

  // --- 2. DETECT DATA SOURCE & READ DRIVER INFO ---
  let driverInfo = {};
  let driverDebits = {};
  let stampDate = null;
  let stampString = "";
  let period = null;
  const tz = ss.getSpreadsheetTimeZone();

  const payablesSheet = findPayablesSheet(ss);
  const tripSheet = ss.getSheetByName("Trip Information");

  if (payablesSheet && isPayablesSheet(payablesSheet)) {
    let parsedData = parsePayablesSheet(payablesSheet, tz);
    if (!parsedData || Object.keys(parsedData.driverInfo).length === 0) {
      return SpreadsheetApp.getUi().alert("Error: Found payables sheet '" + payablesSheet.getName() + "', but could not parse driver rows. Please check headers and data.");
    }
    driverInfo = parsedData.driverInfo;
    driverDebits = parsedData.driverDebits;
    stampDate = parsedData.stampDate;
    stampString = parsedData.stampString;
    period = parsedData.period;
  } else if (tripSheet && tripSheet.getLastRow() >= 3) {
    // Legacy Trip Information parsing
    const tripData = tripSheet.getRange(3, 1, tripSheet.getLastRow() - 2, 14).getValues();
    let maxTripTime = 0;

    tripData.forEach(row => {
      let driver = row[3];
      if (!driver) return;

      let gross = Number(row[12]) || 0;
      let tolls = Number(row[10]) || 0;
      let cash = Number(row[11]) || 0;
      let net = gross - tolls;
      let tDate = parseTripDate(row[4], tz);
      if (!tDate) return;

      if (!driverInfo[driver]) {
        driverInfo[driver] = { gross: 0, tolls: 0, net: 0, cash: 0, earliestDate: tDate };
      }

      driverInfo[driver].gross += gross;
      driverInfo[driver].tolls += tolls;
      driverInfo[driver].net += net;
      driverInfo[driver].cash += cash;

      if (!isNaN(tDate.getTime())) {
        if (tDate < driverInfo[driver].earliestDate) driverInfo[driver].earliestDate = tDate;
        if (tDate.getTime() > maxTripTime) maxTripTime = tDate.getTime();
      }
    });

    if (maxTripTime > 0) {
      let maxTripDate = new Date(maxTripTime);
      let dayOfWeek = maxTripDate.getDay();
      let daysToSunday = (7 - dayOfWeek) % 7;
      let endOfWeekSunday = new Date(maxTripDate);
      endOfWeekSunday.setDate(maxTripDate.getDate() + daysToSunday);

      stampDate = new Date(endOfWeekSunday);
      stampDate.setDate(endOfWeekSunday.getDate() + 6);
      stampString = Utilities.formatDate(stampDate, tz, "MM/dd/yyyy");
    }
  } else {
    return SpreadsheetApp.getUi().alert("Advances created. Could not find 'AYOAH GO Drivers — payables' or 'Trip Information' with data.");
  }

  // --- 3. LOAD ADVANCES AND DEBITS (For Dashboard Prediction Accuracy) ---
  const creditSheet = ss.getSheetByName("Credits & Debits Breakdown") || ss.getSheetByName("Credits & Debits");
  if (creditSheet && creditSheet.getLastRow() >= 3) {
    driverDebits = {};
    creditSheet.getRange(3, 1, creditSheet.getLastRow() - 2, 6).getValues().forEach(row => {
      if (row[1] && row[3] != 0) {
        if (!driverDebits[row[1]]) driverDebits[row[1]] = 0;
        driverDebits[row[1]] += Number(row[3]);
      }
    });
  }

  let driverAdvances = {};
  if (advSheet && advSheet.getLastRow() > 1) {
    advSheet.getRange(2, 1, advSheet.getLastRow() - 1, 3).getValues().forEach(row => {
      let amt = Number(row[1]) || 0; let fee = Number(row[2]) || 0;
      if (row[0] && (amt > 0 || fee > 0)) {
        if (!driverAdvances[row[0]]) driverAdvances[row[0]] = 0;
        driverAdvances[row[0]] += (amt + fee);
      }
    });
  }

  // --- 4. CHECK LOCAL OVERRIDES & FLAGS ---
  let localSettingsSheet = ss.getSheetByName("Settings");
  if (!localSettingsSheet) localSettingsSheet = ss.insertSheet("Settings");

  let existingLocalOverrides = {};
  let existingLocalAdditions = {};
  if (localSettingsSheet.getLastRow() >= 2) {
    let headers = localSettingsSheet.getRange(1, 1, 1, localSettingsSheet.getLastColumn()).getValues()[0];
    let overrideColIndex = headers.indexOf("Manual Override %");
    let additionColIndex = headers.indexOf("Refunds / Additions");
    if (overrideColIndex > -1 || additionColIndex > -1) {
      let oldData = localSettingsSheet.getRange(2, 1, localSettingsSheet.getLastRow() - 1, localSettingsSheet.getLastColumn()).getValues();
      oldData.forEach(r => {
        if (r[0]) {
          if (overrideColIndex > -1 && r[overrideColIndex] !== "") existingLocalOverrides[r[0]] = r[overrideColIndex];
          if (additionColIndex > -1 && r[additionColIndex] !== "") existingLocalAdditions[r[0]] = r[additionColIndex];
        }
      });
    }
  }

  // --- 5. FETCH MASTER DB (PROMOS & LOANS) ---
  let masterDB = getMasterDatabase();
  if (!masterDB) return;

  let masterSetSheet = masterDB.getSheetByName("Settings");
  if (!masterSetSheet) return SpreadsheetApp.getUi().alert("Error: 'Settings' sheet missing in Master Database!");

  let masterData = [];
  if (masterSetSheet.getLastRow() >= 2) {
    masterData = masterSetSheet.getRange(2, 1, masterSetSheet.getLastRow() - 1, 6).getValues();
  }
  let masterIndexMap = {};
  masterData.forEach((r, i) => {
    let key = String(r[0] || "").trim().toLowerCase();
    if (key) masterIndexMap[key] = i;
  });

  // --- 5B. VALIDATE THAT ALL DRIVERS HAVE A STARTING DATE IN MASTER DB ---
  let driversMissingStartDate = [];
  let newDriversAdded = false;

  for (let driver in driverInfo) {
    let key = String(driver).trim().toLowerCase();
    if (masterIndexMap[key] === undefined) {
      masterData.push([String(driver).trim(), "", "", "", "", ""]);
      masterIndexMap[key] = masterData.length - 1;
      newDriversAdded = true;
    }

    let mRowIdx = masterIndexMap[key];
    let mRow = masterData[mRowIdx];
    let mStartDateVal = mRow[1];
    let parsedStartDate = parseFlexibleDate(mStartDateVal) || parseTripDate(mStartDateVal, tz);

    if (!parsedStartDate || isNaN(parsedStartDate.getTime())) {
      driversMissingStartDate.push(driver);
    }
  }

  // If new drivers were added to Master DB array, write them back so user sees their rows ready for start dates
  if (newDriversAdded) {
    masterSetSheet.getRange(2, 1, Math.max(masterSetSheet.getLastRow() - 1, 1), 6).clearContent();
    masterSetSheet.getRange(2, 1, masterData.length, 6).setValues(masterData);
  }

  // If ANY driver is missing a starting date, HALT and do NOT create the Settings dashboard!
  if (driversMissingStartDate.length > 0) {
    let missingListStr = driversMissingStartDate.map(d => "  • " + d).join("\n");
    let msg = "⚠️ SETUP PAUSED: Missing Starting Dates!\n\n" +
      "The Settings dashboard cannot be generated until ALL drivers have a valid starting date in the Master Database.\n\n" +
      "Missing starting dates for " + driversMissingStartDate.length + " driver(s):\n" +
      missingListStr + "\n\n" +
      "👉 Please open the Master Database 'Settings' sheet, enter their starting date in Column B (e.g. MM/DD/YYYY), and click 'Setup Advances & Local Settings' again.";
    SpreadsheetApp.getUi().alert(msg);
    return;
  }

  let masterLoanSheet = masterDB.getSheetByName("Loans");
  if (!masterLoanSheet) return SpreadsheetApp.getUi().alert("Error: 'Loans' sheet missing in Master Database!");

  let activeLoans = {};
  let stampDateObj = new Date(stampString);
  stampDateObj.setHours(0, 0, 0, 0);

  if (masterLoanSheet.getLastRow() >= 2) {
    let loanData = masterLoanSheet.getRange(2, 1, masterLoanSheet.getLastRow() - 1, 8).getValues();
    loanData.forEach(row => {
      let dName = row[0];
      let lDate = row[1];
      let totLoan = parseMoney(row[2]); // Col C
      let install = parseMoney(row[3]); // Col D
      let paidSoFar = parseMoney(row[4]); // Col E

      let remBal = totLoan - paidSoFar; // Auto calculate Remaining Balance

      let lastPayDateStr = "";
      if (row[5] && !isNaN(new Date(row[5]).getTime())) { // Col F
        lastPayDateStr = Utilities.formatDate(new Date(row[5]), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
      }
      let lastPayAmt = parseMoney(row[6]); // Col G

      if (dName && lDate) {
        // Buffer Rule: Only count if Payslip Saturday is strictly greater than Loan Issue Saturday
        let loanSat = getSaturdayOfDate(lDate);
        if (stampDateObj.getTime() > loanSat.getTime()) {
          if (!activeLoans[dName]) activeLoans[dName] = 0;

          if (lastPayDateStr === stampString) {
            activeLoans[dName] += lastPayAmt; // Already locked in for this week
          } else if (remBal > 0) {
            activeLoans[dName] += Math.min(install, remBal);
          }
        }
      }
    });
  }

  // --- 6. APPLY RULES & STAMP MASTER ARRAY ---
  const EIGHTY_NINETY_DRIVERS = ["Angel Yoy"];
  const DEFAULT_RATE = 0.90;
  const PROMO_RATE = 0.95;

  let output = [];
  let updatedMaster = false;

  let periodEndObj = (period && period.end) ? new Date(period.end) : (stampDate ? new Date(stampDate) : new Date());
  periodEndObj.setHours(23, 59, 59, 999);
  let twoMondaysBefore = stampDate ? getTwoMondaysBefore(stampDate) : null;

  for (let driver in driverInfo) {
    let key = String(driver).trim().toLowerCase();
    let mRowIdx = masterIndexMap[key];
    let mRow = masterData[mRowIdx];
    let mStartDateVal = mRow[1];
    let startObj = parseFlexibleDate(mStartDateVal) || parseTripDate(mStartDateVal, tz);
    if (startObj) startObj.setHours(0, 0, 0, 0);

    let priorPaymentsCount = 0;
    let alreadyStampedThisWeek = false;
    let thisWeekCol = -1;

    for (let col = 2; col <= 5; col++) {
      let cellVal = mRow[col];
      if (cellVal && String(cellVal).trim() !== "") {
        let d = parseFlexibleDate(cellVal) || parseTripDate(cellVal, tz);
        if (d && !isNaN(d.getTime())) {
          let existingStr = Utilities.formatDate(d, tz, "MM/dd/yyyy");
          if (existingStr === stampString) {
            alreadyStampedThisWeek = true;
            thisWeekCol = col;
          } else {
            priorPaymentsCount++;
          }
        }
      }
    }

    let isActivePromo = false;
    let currentWeekNumber = 0;

    // Rule:
    // If starting date is more than 2 Mondays before the payment date (e.g. before Oct 5 for Oct 17 payment),
    // the driver is considered (None/Expired) and cannot get into the promo.
    // Also, promos are capped at 4 payments and expire within 35 days (5 weeks) from starting date.
    let daysSinceStart = (startObj && stampDate) ? Math.round((stampDate.getTime() - startObj.getTime()) / (1000 * 60 * 60 * 24)) : 999;
    let isWithinPromoWindow = (daysSinceStart >= 0 && daysSinceStart <= 35);

    if (priorPaymentsCount === 0) {
      // Driver has not yet received a promo payment: to get INTO the promo, start date must be on or after 2 Mondays before payment date
      if (startObj && twoMondaysBefore && startObj.getTime() >= twoMondaysBefore.getTime() && startObj.getTime() <= periodEndObj.getTime()) {
        isActivePromo = true;
        currentWeekNumber = 1;

        if (!alreadyStampedThisWeek && stampDate) {
          for (let col = 2; col <= 5; col++) {
            if (!mRow[col] || String(mRow[col]).trim() === "") {
              mRow[col] = stampDate;
              updatedMaster = true;
              break;
            }
          }
        }
      } else {
        // Start date is before 2 Mondays before payment date: CANNOT GET INTO PROMO!
        isActivePromo = false;
        // If mistakenly stamped for this week earlier, erase it from Master DB!
        if (alreadyStampedThisWeek && thisWeekCol !== -1) {
          mRow[thisWeekCol] = "";
          updatedMaster = true;
        }
      }
    } else if (priorPaymentsCount < 4 && isWithinPromoWindow) {
      // Driver is continuing an active 4-week promo
      isActivePromo = true;
      currentWeekNumber = priorPaymentsCount + 1;

      if (!alreadyStampedThisWeek && stampDate) {
        for (let col = 2; col <= 5; col++) {
          if (!mRow[col] || String(mRow[col]).trim() === "") {
            mRow[col] = stampDate;
            updatedMaster = true;
            break;
          }
        }
      }
    } else {
      // Expired promo or all 4 payments already used
      isActivePromo = false;
      if (alreadyStampedThisWeek && !isWithinPromoWindow && thisWeekCol !== -1) {
        mRow[thisWeekCol] = "";
        updatedMaster = true;
      }
    }

    let sysRate = DEFAULT_RATE;
    let promoDisplay = "None/Expired";

    if (isActivePromo) {
      sysRate = PROMO_RATE;
      promoDisplay = currentWeekNumber + " of 4";
    } else {
      if (EIGHTY_NINETY_DRIVERS.includes(driver)) {
        sysRate = (driverInfo[driver].net > 1000) ? ((1000 * 0.80) + ((driverInfo[driver].net - 1000) * 0.90)) / driverInfo[driver].net : 0.80;
      }
    }

    let override = existingLocalOverrides[driver] !== undefined ? existingLocalOverrides[driver] : "";
    let activeRate = override !== "" ? Number(override) : sysRate;

    let additionalPay = existingLocalAdditions[driver] !== undefined ? parseMoney(existingLocalAdditions[driver]) : 0;
    let additionalPayDisplay = existingLocalAdditions[driver] !== undefined ? existingLocalAdditions[driver] : "";

    // Financial Math (Advances First!)
    let driverPay = driverInfo[driver].net * activeRate;
    let driverTotal = driverPay + driverInfo[driver].tolls;
    let basicBalance = driverTotal - driverInfo[driver].cash;

    let debits = driverDebits[driver] || 0;
    let advances = driverAdvances[driver] || 0;

    // Calculate exactly what is available to cover the loan after advances & additional payments
    let actualAvailableFunds = basicBalance + debits - advances + additionalPay;

    let maxLoanIntended = activeLoans[driver] || 0;
    let predictedLoanDeduction = Math.min(maxLoanIntended, Math.max(0, actualAvailableFunds));

    let dashboardDriverBalance = basicBalance + debits + additionalPay - predictedLoanDeduction; // Dashboard Math
    let companyFee = driverInfo[driver].gross - driverTotal;

    output.push([
      driver, driverInfo[driver].gross, driverInfo[driver].tolls, driverInfo[driver].net,
      sysRate, driverPay, driverTotal, driverInfo[driver].cash, debits, additionalPayDisplay, predictedLoanDeduction,
      dashboardDriverBalance, companyFee, promoDisplay, override
    ]);
  }

  // --- 7. WRITE BACK TO MASTER DB IF WE STAMPED PROMOS ---
  let updateMessage = "Financial Dashboard generated locally!";
  if (updatedMaster || masterData.length > (masterSetSheet.getLastRow() - 1)) {
    // FIX: Clear the old dimensions first so it can dynamically expand to fit the new driver Array length!
    masterSetSheet.getRange(2, 1, Math.max(masterSetSheet.getLastRow() - 1, 1), 6).clearContent();
    masterSetSheet.getRange(2, 1, masterData.length, 6).setValues(masterData);

    if (updatedMaster) {
      localSettingsSheet.getRange("P1").setValue("DB_UPDATED").setFontColor("white");
      updateMessage += `\n\n(Master Database has been securely stamped with Saturday's date: ${stampString}).`;
    }
  }

  // --- 8. FORMATTING THE PROFESSIONAL DASHBOARD ---
  output.sort((a, b) => a[0].localeCompare(b[0]));
  localSettingsSheet.getRange("A:O").clearContent().clearFormat();

  let newHeaders = ["Driver Name", "Gross Fare", "Tolls", "Price no Toll", "System Rate", "Driver Pay", "Driver Total", "Cash Collected", "Credits & Debits", "Refunds / Additions", "Loan Deduction", "Driver Balance", "Company Fee", "Promo Payments Taken", "Manual Override %"];

  let headerRange = localSettingsSheet.getRange("A1:O1");
  headerRange.setValues([newHeaders])
    .setFontWeight("bold").setBackground("#4a86e8").setFontColor("white")
    .setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);

  localSettingsSheet.setRowHeight(1, 40);

  if (output.length > 0) {
    let dataRange = localSettingsSheet.getRange(2, 1, output.length, 15);
    dataRange.setValues(output);
    dataRange.setVerticalAlignment("middle");

    let fullTableRange = localSettingsSheet.getRange(1, 1, output.length + 1, 15);
    fullTableRange.setBorder(true, true, true, true, true, true, "#b7b7b7", SpreadsheetApp.BorderStyle.SOLID);

    for (let i = 0; i < output.length; i++) {
      if (i % 2 === 0) localSettingsSheet.getRange(i + 2, 1, 1, 15).setBackground("#f3f3f3");
    }

    localSettingsSheet.getRange(2, 2, output.length, 3).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    localSettingsSheet.getRange(2, 6, output.length, 4).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    localSettingsSheet.getRange(2, 10, output.length, 1).setBackground("#FFF2CC").setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    localSettingsSheet.getRange(2, 11, output.length, 3).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    localSettingsSheet.getRange(2, 1, output.length, 1).setHorizontalAlignment("left");
    localSettingsSheet.getRange(2, 5, output.length, 1).setNumberFormat("0.00%").setHorizontalAlignment("center");
    localSettingsSheet.getRange(2, 14, output.length, 1).setHorizontalAlignment("center");
    localSettingsSheet.getRange(2, 15, output.length, 1).setBackground("#FFF2CC").setNumberFormat("0.00%").setHorizontalAlignment("center");
  }

  localSettingsSheet.autoResizeColumns(1, 15);
  for (let c = 1; c <= 15; c++) {
    localSettingsSheet.setColumnWidth(c, localSettingsSheet.getColumnWidth(c) + 20);
  }

  SpreadsheetApp.getUi().alert(updateMessage);
}

/**
 * Step 2: Master function. Generates Payslips AND Deducts Loans.
 */
function runCompletePayroll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const HEADER_BLUE = "#CFE2F3";
  const BORDER_BLUE = "#3d85c6";

  const localSetSheet = ss.getSheetByName("Settings");
  if (!localSetSheet) return SpreadsheetApp.getUi().alert("Error: Run Step 1 first to generate the Settings sheet.");

  let finalRates = {};
  let additionalPayments = {};

  if (localSetSheet.getLastRow() >= 2) {
    let headers = localSetSheet.getRange(1, 1, 1, localSetSheet.getLastColumn()).getValues()[0];
    let rateColIndex = headers.indexOf("System Rate");
    let overrideColIndex = headers.indexOf("Manual Override %");
    let additionalColIndex = headers.indexOf("Refunds / Additions");

    localSetSheet.getRange(2, 1, localSetSheet.getLastRow() - 1, localSetSheet.getLastColumn()).getValues().forEach(row => {
      let driver = row[0];
      let sysRate = rateColIndex > -1 ? Number(row[rateColIndex]) || 0 : 0;
      let override = overrideColIndex > -1 ? row[overrideColIndex] : "";
      let addPay = additionalColIndex > -1 ? parseMoney(row[additionalColIndex]) : 0;

      if (driver) {
        finalRates[driver] = (override !== "" && override != null) ? Number(override) : sysRate;
        additionalPayments[driver] = addPay;
      }
    });
  }

  // ==========================================
  // PART A: FILL TRIP COLUMNS & FIND SATURDAY
  // ==========================================
  const tripSheet = ss.getSheetByName("Trip Information");
  if (!tripSheet || tripSheet.getLastRow() < 3) {
    const payablesSheet = findPayablesSheet(ss);
    if (payablesSheet && isPayablesSheet(payablesSheet)) {
      return SpreadsheetApp.getUi().alert("Settings & Advances are ready!\n\nStep 2 (Payslip generation with Daily Balance) will be enabled in the next update as planned.");
    }
    return SpreadsheetApp.getUi().alert("Error: 'Trip Information' sheet not found.");
  }

  const tripData = tripSheet.getRange(3, 1, lastRow - 2, 14).getValues();

  let outputData = [];
  let sums = { priceNoToll: 0, driverPay: 0, tolls: 0, driverTotal: 0, cash: 0, balance: 0, fee: 0 };
  let maxTripTime = 0;

  tripData.forEach(row => {
    let driver = row[3];
    let gross = Number(row[12]) || 0;
    let tolls = Number(row[10]) || 0;
    let cash = Number(row[11]) || 0;
    let netFare = gross - tolls;
    let tDate = parseTripDate(row[4], ss.getSpreadsheetTimeZone());

    if (tDate && !isNaN(tDate.getTime()) && tDate.getTime() > maxTripTime) {
      maxTripTime = tDate.getTime();
    }

    let rate = finalRates[driver] || 0.90;

    let cols = {
      priceNoToll: netFare, driverPay: netFare * rate, rate: rate, tolls: tolls,
      driverTotal: (netFare * rate) + tolls, cash: cash, balance: ((netFare * rate) + tolls) - cash,
      fee: gross - ((netFare * rate) + tolls)
    };

    sums.priceNoToll += cols.priceNoToll; sums.driverPay += cols.driverPay; sums.tolls += cols.tolls;
    sums.driverTotal += cols.driverTotal; sums.cash += cols.cash; sums.balance += cols.balance; sums.fee += cols.fee;

    outputData.push([cols.priceNoToll, cols.driverPay, cols.rate, cols.tolls, cols.driverTotal, cols.cash, cols.balance, cols.fee]);
  });

  let stampDate = null;
  let stampString = "";
  if (maxTripTime > 0) {
    let maxTripDate = new Date(maxTripTime);
    let dayOfWeek = maxTripDate.getDay();
    let daysToSunday = (7 - dayOfWeek) % 7;
    let endOfWeekSunday = new Date(maxTripDate);
    endOfWeekSunday.setDate(maxTripDate.getDate() + daysToSunday);

    stampDate = new Date(endOfWeekSunday);
    stampDate.setDate(endOfWeekSunday.getDate() + 6);
    stampString = Utilities.formatDate(stampDate, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
  }

  tripSheet.getRange("O2:V2").setValues([["Price no Toll", "Driver Pay", "Rate", "Tolls.", "Driver Total.", "Cash Collected.", "Driver Balance.", "Company Fee"]])
    .setFontWeight("bold").setHorizontalAlignment("right").setBackground("#EFEFEF");
  tripSheet.getRange(3, 15, outputData.length, 8).setValues(outputData);

  let avgRate = sums.priceNoToll > 0 ? (sums.driverPay / sums.priceNoToll) : 0;
  tripSheet.getRange("O1:V1").setValues([[sums.priceNoToll, sums.driverPay, avgRate, sums.tolls, sums.driverTotal, sums.cash, sums.balance, sums.fee]])
    .setFontWeight("bold").setFontSize(14).setVerticalAlignment("middle").setHorizontalAlignment("right").setBackground("#EFEFEF");

  tripSheet.getRange(1, 15, lastRow, 2).setNumberFormat("$#,##0.00");
  tripSheet.getRange(1, 18, lastRow, 5).setNumberFormat("$#,##0.00");
  tripSheet.getRange("Q1").setNumberFormat('"Avg" 0.00%');
  tripSheet.getRange(3, 17, lastRow - 2, 1).setNumberFormat("0.00%");
  tripSheet.setColumnWidths(15, 8, 130);
  SpreadsheetApp.flush();

  // ==========================================
  // FETCH MASTER DB LOANS FOR PAYSLIP
  // ==========================================
  let masterDB = getMasterDatabase();
  let driverLoans = {};
  let masterLoanData = [];
  let masterLoanSheet = null;

  let stampDateObj = new Date(stampString);
  stampDateObj.setHours(0, 0, 0, 0);

  if (masterDB) {
    masterLoanSheet = masterDB.getSheetByName("Loans");
    if (masterLoanSheet && masterLoanSheet.getLastRow() >= 2) {
      masterLoanData = masterLoanSheet.getRange(2, 1, masterLoanSheet.getLastRow() - 1, 8).getValues();
      masterLoanData.forEach((row, idx) => {
        let dName = row[0];
        let lDate = row[1];
        let totLoan = parseMoney(row[2]); // Col C
        let install = parseMoney(row[3]); // Col D
        let paidSoFar = parseMoney(row[4]); // Col E

        let remBal = totLoan - paidSoFar; // Auto calculate Remaining Balance
        row[7] = remBal;

        let lastPayDateStr = "";
        if (row[5] && !isNaN(new Date(row[5]).getTime())) { // Col F
          lastPayDateStr = Utilities.formatDate(new Date(row[5]), ss.getSpreadsheetTimeZone(), "MM/dd/yyyy");
        }
        let lastPayAmt = parseMoney(row[6]); // Col G

        if (dName && lDate) {
          // Buffer Rule Check
          let loanSat = getSaturdayOfDate(lDate);
          if (stampDateObj.getTime() > loanSat.getTime()) {
            if (!driverLoans[dName]) driverLoans[dName] = [];
            driverLoans[dName].push({
              idx: idx,
              date: lDate,
              remBal: remBal,
              install: install,
              lastPayDateStr: lastPayDateStr,
              lastPayAmt: lastPayAmt
            });
          }
        }
      });
    }
  }

  // ==========================================
  // PART B & C: READ DATA AND BUILD PAYSLIP
  // ==========================================
  const creditSheet = ss.getSheetByName("Credits & Debits Breakdown") || ss.getSheetByName("Credits & Debits");
  let driverDebits = {};
  if (creditSheet && creditSheet.getLastRow() >= 3) {
    creditSheet.getRange(3, 1, creditSheet.getLastRow() - 2, 6).getValues().forEach(row => {
      if (row[1] && row[3] != 0) {
        if (!driverDebits[row[1]]) driverDebits[row[1]] = [];
        driverDebits[row[1]].push({ desc: row[4] || "Other", amount: row[3], date: row[5] });
      }
    });
  }

  const advanceSheet = ss.getSheetByName("Advances");
  let driverAdvances = {};
  if (advanceSheet && advanceSheet.getLastRow() > 1) {
    advanceSheet.getRange(2, 1, advanceSheet.getLastRow() - 1, 3).getValues().forEach(row => {
      let amt = Number(row[1]) || 0; let fee = Number(row[2]) || 0;
      if (row[0] && (amt > 0 || fee > 0)) {
        if (!driverAdvances[row[0]]) driverAdvances[row[0]] = [];
        driverAdvances[row[0]].push({ amount: amt, fee: fee });
      }
    });
  }

  let drivers = {};
  tripData.forEach(row => {
    let name = row[3];
    if (!name) return;
    let tripDate = parseTripDate(row[4], ss.getSpreadsheetTimeZone());
    if (!tripDate) return;
    let isoDateStr = Utilities.formatDate(tripDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    let displayDateStr = Utilities.formatDate(tripDate, ss.getSpreadsheetTimeZone(), "dd-MM-yy");
    let gross = Number(row[12]) || 0; let tolls = Number(row[10]) || 0; let cash = Number(row[11]) || 0;
    let pay = (gross - tolls) * (finalRates[name] || 0.90);

    if (!drivers[name]) drivers[name] = { dates: {}, totalTrips: 0, totalPay: 0, totalTolls: 0, totalTotal: 0, totalCash: 0 };
    if (!drivers[name].dates[isoDateStr]) drivers[name].dates[isoDateStr] = { displayDate: displayDateStr, count: 0, pay: 0, tolls: 0, total: 0 };

    let d = drivers[name].dates[isoDateStr];
    d.count += 1; d.pay += pay; d.tolls += tolls; d.total += (pay + tolls);
    drivers[name].totalTrips += 1; drivers[name].totalPay += pay; drivers[name].totalTolls += tolls;
    drivers[name].totalTotal += (pay + tolls); drivers[name].totalCash += cash;
  });

  let paySheet = ss.getSheetByName("PaySlip");
  if (paySheet) ss.deleteSheet(paySheet);
  paySheet = ss.insertSheet("PaySlip");
  paySheet.setHiddenGridlines(true);
  paySheet.getRange("A:G").setBackground("white");

  let currentRow = 2;
  let totalAllBalances = 0;
  let masterLoanUpdates = [];

  for (let name in drivers) {
    let d = drivers[name];

    paySheet.getRange(currentRow, 2, 1, 5).setValues([["Driver", "Num of Trips", "Price", "Tolls", "Driver Total"]])
      .setFontWeight("bold").setBackground(HEADER_BLUE).setBorder(true, false, true, false, false, false, BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    let summaryRange = paySheet.getRange(currentRow, 2, 1, 5);
    summaryRange.setValues([[name, d.totalTrips, d.totalPay, d.totalTolls, d.totalTotal]]).setFontWeight("bold");
    paySheet.getRange(currentRow, 2).setHorizontalAlignment("left");
    summaryRange.setBorder(false, false, true, false, false, false, BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    Object.keys(d.dates).sort().forEach(isoDate => {
      let day = d.dates[isoDate];
      paySheet.getRange(currentRow, 2, 1, 5).setValues([[day.displayDate, day.count, day.pay, day.tolls, day.total]]);
      paySheet.getRange(currentRow, 2).setHorizontalAlignment("right");
      currentRow++;
    });

    paySheet.getRange(currentRow, 2, 1, 4).setValues([["Grand Total", d.totalTrips, d.totalPay, d.totalTolls]])
      .setFontWeight("bold").setBackground(HEADER_BLUE).setBorder(true, false, true, false, false, false, BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID);

    let totalCell = paySheet.getRange(currentRow, 6, 2, 1);
    totalCell.merge().setValue(d.totalTotal).setFontWeight("bold").setFontSize(20).setBackground(HEADER_BLUE)
      .setVerticalAlignment("middle").setHorizontalAlignment("right")
      .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    currentRow += 2;

    if (d.totalCash > 0) {
      paySheet.getRange(currentRow, 4).setValue("Sub: Cash Collected");
      paySheet.getRange(currentRow, 6).setValue(-d.totalCash);
      currentRow++;
    }

    let debitSum = 0;
    if (driverDebits[name]) {
      driverDebits[name].forEach(item => {
        let cleanDesc = String(item.desc).toLowerCase().includes("verrazano") || String(item.desc).toLowerCase().includes("bridge") || String(item.desc).toLowerCase().includes("wrong route") || String(item.desc).toLowerCase().includes("tunnel") ? "Penalty: Wrong Bridge/Route" :
          String(item.desc).toLowerCase().includes("cancel") ? "Cancellation Fee" :
            String(item.desc).toLowerCase().includes("late") ? "Penalty: Late Arrival" :
              String(item.desc).toLowerCase().includes("no show") ? "Penalty: No Show" :
                String(item.desc).replace(/TRIP ID/gi, "").replace(/ for \d+/gi, "").replace(/\b\d{5,}\b/g, "").replace(/\s+/g, " ").trim();
        let dateSuffix = (item.date && !isNaN(new Date(item.date).getTime())) ? " " + Utilities.formatDate(parseTripDate(item.date, ss.getSpreadsheetTimeZone()) || new Date(item.date), ss.getSpreadsheetTimeZone(), "dd-MM") : "";
        paySheet.getRange(currentRow, 4).setValue((item.amount < 0 ? "Sub: " : "Add: ") + cleanDesc + dateSuffix);
        paySheet.getRange(currentRow, 6).setValue(item.amount);
        debitSum += Number(item.amount);
        currentRow++;
      });
    }

    let advanceSum = 0;
    if (driverAdvances[name]) {
      driverAdvances[name].forEach(adv => {
        if (adv.amount > 0) {
          paySheet.getRange(currentRow, 4).setValue("Sub: Advanced Payment");
          paySheet.getRange(currentRow, 6).setValue(-adv.amount); advanceSum += adv.amount; currentRow++;
        }
        if (adv.fee > 0) {
          paySheet.getRange(currentRow, 4).setValue("Sub: Fees");
          paySheet.getRange(currentRow, 6).setValue(-adv.fee); advanceSum += adv.fee; currentRow++;
        }
      });
    }

    let addPayAmt = additionalPayments[name] || 0;
    if (addPayAmt !== 0) {
      paySheet.getRange(currentRow, 4).setValue(addPayAmt > 0 ? "Add: Refund / Addition" : "Sub: Refund / Deduction");
      paySheet.getRange(currentRow, 6).setValue(addPayAmt);
      currentRow++;
    }

    // --- NEW: 100% IDEMPOTENT LOAN CALCULATION (Advances & Additions First) ---
    let preLoanBalance = d.totalTotal - d.totalCash + debitSum - advanceSum + addPayAmt;
    let fundsAvailable = Math.max(0, preLoanBalance);
    let totalLoanDeducted = 0;

    if (driverLoans[name] && fundsAvailable > 0) {
      driverLoans[name].sort((a, b) => new Date(a.date) - new Date(b.date));

      driverLoans[name].forEach(loan => {
        if (fundsAvailable > 0) {
          let actualDeduct = 0;
          let isAlreadyProcessedThisWeek = (loan.lastPayDateStr === stampString);

          if (isAlreadyProcessedThisWeek) {
            actualDeduct = Math.min(loan.lastPayAmt, fundsAvailable);
          } else if (loan.remBal > 0) {
            let maxDeduct = Math.min(loan.install, loan.remBal);
            actualDeduct = Math.min(maxDeduct, fundsAvailable);
          }

          if (actualDeduct > 0) {
            let dStr = "";
            if (loan.date && !isNaN(new Date(loan.date).getTime())) {
              let origDate = parseTripDate(loan.date, ss.getSpreadsheetTimeZone()) || new Date(loan.date);
              dStr = " (" + Utilities.formatDate(origDate, ss.getSpreadsheetTimeZone(), "MM/dd/yy") + ")";
            }
            paySheet.getRange(currentRow, 4).setValue("Sub: Loan Repayment" + dStr);
            paySheet.getRange(currentRow, 6).setValue(-actualDeduct);
            currentRow++;

            fundsAvailable -= actualDeduct;
            totalLoanDeducted += actualDeduct;

            if (!isAlreadyProcessedThisWeek) {
              masterLoanUpdates.push({
                idx: loan.idx,
                addAmount: actualDeduct,
                stampDate: stampDate
              });
            }
          }
        }
      });
    }

    let finalBal = preLoanBalance - totalLoanDeducted;
    totalAllBalances += finalBal;

    paySheet.getRange(currentRow, 4).setValue("Driver Balance");
    paySheet.getRange(currentRow, 6).setValue(finalBal);
    paySheet.getRange(currentRow, 4, 1, 3).setBackground(HEADER_BLUE).setFontWeight("bold").setFontColor("black").setFontStyle("normal");
    currentRow += 4;
  }

  paySheet.getRange(currentRow, 4).setValue("TOTAL PAYOUT");
  paySheet.getRange(currentRow, 6).setValue(totalAllBalances);
  paySheet.getRange(currentRow, 4, 1, 3).setBackground(HEADER_BLUE).setFontWeight("bold").setFontSize(12)
    .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  paySheet.getRange(1, 4, currentRow + 1, 3).setNumberFormat("#,##0.00");
  paySheet.getRange(1, 3, currentRow, 1).setHorizontalAlignment("center");
  paySheet.getRange(1, 4, currentRow, 3).setHorizontalAlignment("right");
  paySheet.setColumnWidth(1, 60); paySheet.setColumnWidth(2, 160); paySheet.setColumnWidth(3, 100);
  paySheet.setColumnWidth(4, 200); paySheet.setColumnWidth(5, 100); paySheet.setColumnWidth(6, 120);

  if (paySheet.getMaxRows() > currentRow) paySheet.deleteRows(currentRow + 1, paySheet.getMaxRows() - currentRow);
  if (paySheet.getMaxColumns() > 7) paySheet.deleteColumns(8, paySheet.getMaxColumns() - 7);
  ss.setActiveSheet(paySheet);

  // --- FINALIZE IDEMPOTENT 100% AUTOMATED DB STAMPING FOR LOANS ---
  let alertMessage = "Payslips successfully generated!";
  if (masterLoanUpdates.length > 0 && masterLoanSheet) {
    masterLoanUpdates.forEach(upd => {
      let r = masterLoanData[upd.idx];
      let totLoan = parseMoney(r[2]);   // Col C
      let paidSoFar = parseMoney(r[4]); // Col E
      let newPaidSoFar = paidSoFar + upd.addAmount;

      r[4] = newPaidSoFar;              // Col E: Amount Paid So Far
      r[5] = upd.stampDate;             // Col F: Last Payment Date
      r[6] = upd.addAmount;             // Col G: Last Payment Amount
      r[7] = totLoan - newPaidSoFar;    // Col H: Remaining Balance (Auto Calculated!)
    });

    // Write back the whole perfectly calculated array
    masterLoanSheet.getRange(2, 1, masterLoanData.length, 8).setValues(masterLoanData);
    alertMessage += "\n\n(Master Database 'Loans' sheet has been securely updated. All balances auto-calculated. The system locked the dates so you can safely re-run this file without double counting!).";
  }

  SpreadsheetApp.getUi().alert(alertMessage);
}