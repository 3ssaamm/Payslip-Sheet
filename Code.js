// ==========================================
// ⚙️ SYSTEM CONFIGURATION
// Your Master Database Link is permanently saved here:
// ==========================================
const MASTER_DB_URL = "https://docs.google.com/spreadsheets/d/1RJ-4CCGSqYm7fAM5XkQ8rFDzKyZOiADf-H8_OX8DxJI/edit";
const DAILY_BALANCE_BASE_FOLDER_ID = "1iDd2ME2b6coX8GO6Dkom5u9jldkRSwVg";
// Optional: Paste direct URL to the monthly Drivers Daily Balance sheet (e.g. October) to guarantee 100% direct access:
const MONTHLY_BALANCE_FILE_URL = "";

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
    .addSeparator()
    .addItem('🧪 Test Daily Balance Reader', 'testDailyBalanceReader')
    .addItem('🔍 Diagnose Daily Balance File', 'diagnoseDailyBalanceConnection')
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
  if (str instanceof Date && !isNaN(str.getTime())) return str;
  str = String(str).trim();
  let mIso = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (mIso) {
    return new Date(Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]), 12, 0, 0, 0);
  }
  let m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (m) {
    let n1 = Number(m[1]);
    let n2 = Number(m[2]);
    let year = Number(m[3]);
    if (year < 100) year += 2000;
    let month, day;
    if (n1 > 12) {
      day = n1; month = n2;
    } else if (n2 > 12) {
      month = n1; day = n2;
    } else {
      // Standard US format default: MM-DD-YYYY or MM/DD/YYYY
      month = n1; day = n2;
    }
    return new Date(year, month - 1, day, 12, 0, 0, 0);
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
 * Helper: Resiliently matches a driver's name in the daily balance map
 */
function getDriverDailyMap(dailyBreakdown, driverName) {
  if (!dailyBreakdown || !driverName) return {};
  let clean = driverName.toLowerCase().replace(/\s+/g, " ").trim();
  if (dailyBreakdown[clean]) return dailyBreakdown[clean];

  let stripped = clean.replace(/[^a-z0-9]/g, "");
  for (let k in dailyBreakdown) {
    if (k.replace(/[^a-z0-9]/g, "") === stripped) {
      return dailyBreakdown[k];
    }
  }

  for (let k in dailyBreakdown) {
    if (k.includes(clean) || clean.includes(k)) {
      return dailyBreakdown[k];
    }
  }

  // Token-based matching: first and last name match
  let tokens = clean.split(/\s+/).filter(t => t.length > 2);
  if (tokens.length >= 2) {
    let first = tokens[0];
    let last = tokens[tokens.length - 1];
    for (let k in dailyBreakdown) {
      if (k.includes(first) && k.includes(last)) {
        return dailyBreakdown[k];
      }
    }
  }

  return {};
}

/**
 * Helper: Parses and resolves daily dates from the Monthly sheet (handling "Thursday, 01", "05-10-26", native Dates, ISO, etc.)
 */
function resolveDailyDate(rawCell, displayCell, targetYear, targetMonthIndex, tz) {
  let dayNum = null;
  let str = String(displayCell || rawCell || '').trim();
  if (!str) return null;

  // Pattern 1: Day name + day number (e.g. 'Thursday, 01', 'Monday, 05', 'Sat, 10', 'Thursday, 1')
  let mName = str.match(/(?:mon|tue|wed|thu|fri|sat|sun)?[a-z]*[,\s]+0?([1-9]|[12]\d|3[01])\b/i);
  if (mName && mName[1]) {
    dayNum = parseInt(mName[1], 10);
  }

  // Pattern 2: DD-MM-YYYY or DD/MM/YYYY or MM/DD/YYYY
  if (!dayNum) {
    let mDate = str.match(/^0?([1-9]|[12]\d|3[01])[\/\-]0?([1-9]|[12]\d|3[01])[\/\-](\d{2,4})/);
    if (mDate) {
      let n1 = parseInt(mDate[1], 10);
      let n2 = parseInt(mDate[2], 10);
      let targetMonthNum = targetMonthIndex + 1;
      if (n2 === targetMonthNum) {
        dayNum = n1; // DD-MM-YYYY
      } else if (n1 === targetMonthNum) {
        dayNum = n2; // MM-DD-YYYY
      } else if (n1 > 12) {
        dayNum = n1;
      } else {
        dayNum = n2;
      }
    }
  }

  // Pattern 3: Date object fallback
  if (!dayNum && rawCell instanceof Date && !isNaN(rawCell.getTime())) {
    if (rawCell.getMonth() === targetMonthIndex) {
      dayNum = rawCell.getDate();
    } else if (rawCell.getDate() === targetMonthIndex + 1) {
      dayNum = rawCell.getMonth() + 1;
    } else {
      dayNum = rawCell.getDate();
    }
  }

  // Pattern 4: Bare number (e.g. 1 to 31)
  if (!dayNum) {
    let mNum = str.match(/^0?([1-9]|[12]\d|3[01])$/);
    if (mNum) dayNum = parseInt(mNum[1], 10);
  }

  if (dayNum != null && dayNum >= 1 && dayNum <= 31) {
    return new Date(targetYear, targetMonthIndex, dayNum, 12, 0, 0, 0);
  }
  return null;
}

/**
/**
 * Helper: Fetches the daily trip and price breakdown for each driver from Google Drive Daily Balance sheets.
 * Path structure: Base Folder (DAILY_BALANCE_BASE_FOLDER_ID) -> Year (e.g. 2026) -> Monthly Sheet (e.g. "October - Drivers Daily Balance") -> "Summary" tab
 */
function fetchDailyBalanceBreakdown(period, tz) {
  let dailyBreakdown = {};
  dailyBreakdown._meta = { filesFound: [], selectedFile: null, errors: [] };

  if (!period || !period.start || !period.end) {
    Logger.log("fetchDailyBalanceBreakdown: No valid period provided.");
    dailyBreakdown._meta.errors.push("No valid pay period provided.");
    return dailyBreakdown;
  }

  const MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  // Identify all year/month targets across the period
  let targets = [];
  let cur = new Date(period.start);
  cur.setHours(0, 0, 0, 0);
  let end = new Date(period.end);
  end.setHours(23, 59, 59, 999);

  let visited = new Set();
  while (cur.getTime() <= end.getTime()) {
    let y = cur.getFullYear();
    let m = cur.getMonth();
    let key = `${y}-${m}`;
    if (!visited.has(key)) {
      visited.add(key);
      targets.push({ year: y, monthIndex: m, monthName: MONTH_NAMES[m] });
    }
    cur.setDate(cur.getDate() + 1);
  }

  let pStartIso = Utilities.formatDate(period.start, tz, "yyyy-MM-dd");
  let pEndIso = Utilities.formatDate(period.end, tz, "yyyy-MM-dd");

  for (let target of targets) {
    let dailySS = null;
    let targetFileName = target.monthName + " - Drivers Daily Balance";

    // 1. Direct URL takes absolute priority if configured
    if (MONTHLY_BALANCE_FILE_URL && String(MONTHLY_BALANCE_FILE_URL).trim() !== "") {
      try {
        dailySS = SpreadsheetApp.openByUrl(String(MONTHLY_BALANCE_FILE_URL).trim());
        dailyBreakdown._meta.selectedFile = {
          name: dailySS.getName(),
          id: dailySS.getId(),
          source: "Direct URL (MONTHLY_BALANCE_FILE_URL)",
          lastUpdated: "Active Sheet"
        };
      } catch (e) {
        Logger.log("Error opening MONTHLY_BALANCE_FILE_URL: " + e.message);
        dailyBreakdown._meta.errors.push("Could not open MONTHLY_BALANCE_FILE_URL: " + e.message);
      }
    }

    // 2. Drive Folder Search if direct URL is not configured
    if (!dailySS) {
      let baseFolder;
      try {
        baseFolder = DriveApp.getFolderById(DAILY_BALANCE_BASE_FOLDER_ID);
      } catch (e) {
        let err = "Could not open Daily Balance base folder ID (" + DAILY_BALANCE_BASE_FOLDER_ID + "): " + e.message;
        Logger.log(err);
        dailyBreakdown._meta.errors.push(err);
        continue;
      }

      let yearStr = String(target.year);
      let yearFolder = null;
      let yearFolders = baseFolder.getFoldersByName(yearStr);
      if (yearFolders.hasNext()) {
        yearFolder = yearFolders.next();
      } else {
        let allSubfolders = baseFolder.getFolders();
        while (allSubfolders.hasNext()) {
          let f = allSubfolders.next();
          if (f.getName().includes(yearStr)) {
            yearFolder = f;
            break;
          }
        }
      }

      if (!yearFolder) {
        let err = "Year folder '" + yearStr + "' not found in Google Drive folder.";
        Logger.log(err);
        dailyBreakdown._meta.errors.push(err);
        continue;
      }

      // Collect ALL candidate files in year folder
      let candidateFiles = [];
      let allFiles = yearFolder.getFiles();
      while (allFiles.hasNext()) {
        let f = allFiles.next();
        if (f.isTrashed()) continue;
        let fname = f.getName().trim();
        let fnameLower = fname.toLowerCase();
        let mLower = target.monthName.toLowerCase();
        let mShort = mLower.slice(0, 3);

        let isExact = (fnameLower === targetFileName.toLowerCase());
        let isMonthlyBalance = (fnameLower.includes(mLower) || fnameLower.includes(mShort)) &&
                               (fnameLower.includes("balance") || fnameLower.includes("daily"));
        let isSingleDayFile = /\b(?:0?[1-9]|[12]\d|3[01])\b/.test(fnameLower) && !fnameLower.includes("drivers daily balance");

        if (isExact) {
          candidateFiles.push({ file: f, priority: 1, name: fname, updated: f.getLastUpdated() });
        } else if (isMonthlyBalance && !isSingleDayFile) {
          candidateFiles.push({ file: f, priority: 2, name: fname, updated: f.getLastUpdated() });
        } else if (isMonthlyBalance) {
          candidateFiles.push({ file: f, priority: 3, name: fname, updated: f.getLastUpdated() });
        }
      }

      // Priority 1 (exact name) first, then newest lastUpdated timestamp first
      candidateFiles.sort((a, b) => {
        if (a.priority !== b.priority) return a.priority - b.priority;
        return b.updated.getTime() - a.updated.getTime();
      });

      dailyBreakdown._meta.filesFound = candidateFiles.map(c => ({
        name: c.name,
        id: c.file.getId(),
        updated: c.updated,
        priority: c.priority
      }));

      for (let cand of candidateFiles) {
        try {
          let testSS = SpreadsheetApp.openById(cand.file.getId());
          let testSummary = testSS.getSheetByName("Summary");
          if (!testSummary) {
            for (let s of testSS.getSheets()) {
              let sn = s.getName().toLowerCase();
              if (sn === "summary" || (sn.includes("summary") && !sn.includes("weekly") && !sn.includes("angel"))) {
                testSummary = s;
                break;
              }
            }
          }
          if (testSummary && testSummary.getLastRow() >= 3) {
            dailySS = testSS;
            dailyBreakdown._meta.selectedFile = {
              name: cand.name,
              id: cand.file.getId(),
              source: "Google Drive Folder '" + yearFolder.getName() + "'",
              lastUpdated: cand.updated
            };
            break;
          }
        } catch (e) {
          Logger.log("Could not open candidate file " + cand.name + ": " + e.message);
        }
      }
    }

    if (!dailySS) {
      let err = "No valid '" + targetFileName + "' spreadsheet found in Drive folder or direct URL.";
      Logger.log(err);
      dailyBreakdown._meta.errors.push(err);
      continue;
    }

    // Access the Summary sheet
    let summarySheet = dailySS.getSheetByName("Summary");
    if (!summarySheet) {
      for (let s of dailySS.getSheets()) {
        let sn = s.getName().toLowerCase();
        if (sn === "summary" || (sn.includes("summary") && !sn.includes("weekly") && !sn.includes("angel"))) {
          summarySheet = s;
          break;
        }
      }
    }
    if (!summarySheet) summarySheet = dailySS.getSheets()[0];
    if (!summarySheet || summarySheet.getLastRow() < 2) continue;

    // Restrict range to columns A through K (1 to 11) to isolate the data table and avoid charts/pivots in cols L+
    let range = summarySheet.getRange(1, 1, summarySheet.getLastRow(), 11);
    let allValues = range.getValues();
    let allDisplayValues = range.getDisplayValues();

    // Map Header Columns strictly within columns A-K (0 to 10)
    let headerRowIdx = -1;
    let colMap = {
      date: -1, driver: -1, grossFare: -1, hours: -1, driverAvg: -1,
      cph: -1, trips: -1, cash: -1, tolls: -1, creditDebit: -1, balance: -1
    };

    for (let r = 0; r < Math.min(allValues.length, 6); r++) {
      let row = (allValues[r] || []).slice(0, 11).map(c => String(c || "").trim().toLowerCase());
      if (row.includes("date") && (row.includes("driver") || row.includes("trips") || row.includes("hours") || row.includes("gross fare"))) {
        headerRowIdx = r;
        row.slice(0, 11).forEach((h, idx) => {
          if (h === "date" && colMap.date === -1) colMap.date = idx;
          else if ((h === "driver" || h === "driver name" || h === "name") && colMap.driver === -1) colMap.driver = idx;
          else if (h.includes("gross") && colMap.grossFare === -1) colMap.grossFare = idx;
          else if ((h === "hours" || h === "hour") && colMap.hours === -1) colMap.hours = idx;
          else if (h.includes("avg") && (h.includes("hour") || h.includes("driver")) && colMap.driverAvg === -1) colMap.driverAvg = idx;
          else if (h === "cph" && colMap.cph === -1) colMap.cph = idx;
          else if ((h === "trips" || h === "num of trips") && colMap.trips === -1) colMap.trips = idx;
          else if (h === "cash" && colMap.cash === -1) colMap.cash = idx;
          else if (h.includes("toll") && colMap.tolls === -1) colMap.tolls = idx;
          else if ((h.includes("credit") || h.includes("debit")) && colMap.creditDebit === -1) colMap.creditDebit = idx;
          else if (h.includes("balance") && colMap.balance === -1) colMap.balance = idx;
        });
        break;
      }
    }

    if (colMap.date === -1) colMap.date = 0;       // Col A
    if (colMap.driver === -1) colMap.driver = 1;   // Col B
    if (colMap.grossFare === -1) colMap.grossFare = 2; // Col C
    if (colMap.hours === -1) colMap.hours = 3;     // Col D
    if (colMap.driverAvg === -1) colMap.driverAvg = 4; // Col E
    if (colMap.cph === -1) colMap.cph = 5;         // Col F
    if (colMap.trips === -1) colMap.trips = 6;     // Col G
    if (colMap.cash === -1) colMap.cash = 7;       // Col H
    if (colMap.tolls === -1) colMap.tolls = 8;     // Col I
    if (colMap.creditDebit === -1) colMap.creditDebit = 9; // Col J
    if (colMap.balance === -1) colMap.balance = 10; // Col K

    let currentDate = null;
    let currentIsoDate = null;
    let currentDisplayDate = null;

    let startR = headerRowIdx !== -1 ? headerRowIdx + 1 : 2;

    // 5. Parse Data Rows
    for (let r = startR; r < allValues.length; r++) {
      let row = allValues[r] || [];
      let displayRow = allDisplayValues[r] || [];

      let rawDateCell = row[colMap.date];
      let displayDateCell = displayRow[colMap.date];

      // Only update date when the cell actually has text (prevents empty merged cells from resetting date)
      if (String(displayDateCell || rawDateCell || '').trim() !== "") {
        let d = resolveDailyDate(rawDateCell, displayDateCell, target.year, target.monthIndex, tz);
        if (d && !isNaN(d.getTime())) {
          currentDate = d;
          currentIsoDate = Utilities.formatDate(d, tz, "yyyy-MM-dd");
          currentDisplayDate = Utilities.formatDate(d, tz, "dd-MM-yy");
        }
      }

      if (!currentDate || !currentIsoDate) continue;

      // Filter by period bounds using ISO date string comparison
      if (currentIsoDate < pStartIso || currentIsoDate > pEndIso) continue;

      let driverName = String(row[colMap.driver] || "").trim();
      if (!driverName || driverName.toLowerCase().startsWith("total")) continue;

      let grossFare = colMap.grossFare !== -1 ? parseMoney(row[colMap.grossFare]) : 0;
      let hours = colMap.hours !== -1 ? Number(row[colMap.hours]) || 0 : 0;
      let driverAvg = colMap.driverAvg !== -1 ? parseMoney(row[colMap.driverAvg]) : 0;
      let trips = colMap.trips !== -1 ? (Number(row[colMap.trips]) || 0) : 0;
      let cash = colMap.cash !== -1 ? parseMoney(row[colMap.cash]) : 0;
      let tolls = colMap.tolls !== -1 ? parseMoney(row[colMap.tolls]) : 0;
      let creditDebit = colMap.creditDebit !== -1 ? parseMoney(row[colMap.creditDebit]) : 0;

      // Skip rows where the driver did not work or had zero activity
      if (trips === 0 && grossFare === 0 && tolls === 0 && cash === 0) continue;

      // Net fare without tolls
      let netFare = 0;
      if (grossFare > 0) {
        netFare = Math.max(0, grossFare - tolls);
      } else if (hours > 0 && driverAvg > 0) {
        netFare = hours * driverAvg;
        grossFare = netFare + tolls;
      }

      let key = driverName.toLowerCase().replace(/\s+/g, " ");
      if (!dailyBreakdown[key]) dailyBreakdown[key] = {};
      if (!dailyBreakdown[key][currentIsoDate]) {
        dailyBreakdown[key][currentIsoDate] = {
          displayDate: currentDisplayDate,
          count: 0,
          gross: 0,
          net: 0,
          tolls: 0,
          cash: 0,
          creditDebit: 0,
          hours: 0,
          driverAvg: 0
        };
      }

      dailyBreakdown[key][currentIsoDate].count += trips;
      dailyBreakdown[key][currentIsoDate].gross += grossFare;
      dailyBreakdown[key][currentIsoDate].net += netFare;
      dailyBreakdown[key][currentIsoDate].tolls += tolls;
      dailyBreakdown[key][currentIsoDate].cash += cash;
      dailyBreakdown[key][currentIsoDate].creditDebit += creditDebit;
      dailyBreakdown[key][currentIsoDate].hours += hours;
      dailyBreakdown[key][currentIsoDate].driverAvg = driverAvg;
    }
  }

  return dailyBreakdown;
}

/**
 * Step 2: Generates Payslip for the new Payables sheet workflow using Daily Balance breakdown.
 */
function generatePayablesPayslip(ss, payablesSheet, localSetSheet, finalRates, additionalPayments) {
  const HEADER_BLUE = "#CFE2F3";
  const BORDER_BLUE = "#3d85c6";
  const tz = ss.getSpreadsheetTimeZone();

  let payablesData = parsePayablesSheet(payablesSheet, tz);
  if (!payablesData || !payablesData.period) {
    return SpreadsheetApp.getUi().alert("Error: Could not parse pay period or driver rows from 'AYOAH GO Drivers — payables'.");
  }

  let stampDate = payablesData.stampDate;
  let stampString = payablesData.stampString;
  let stampDateObj = new Date(stampString);
  stampDateObj.setHours(0, 0, 0, 0);

  // 1. Fetch Daily Balance Breakdown from Google Drive
  let dailyBreakdown = fetchDailyBalanceBreakdown(payablesData.period, tz);

  // 2. Fetch Advances
  const advanceSheet = ss.getSheetByName("Advances");
  let driverAdvances = {};
  if (advanceSheet && advanceSheet.getLastRow() > 1) {
    advanceSheet.getRange(2, 1, advanceSheet.getLastRow() - 1, 3).getValues().forEach(row => {
      let amt = Number(row[1]) || 0;
      let fee = Number(row[2]) || 0;
      if (row[0] && (amt > 0 || fee > 0)) {
        let name = String(row[0]).trim();
        let key = name.toLowerCase().replace(/\s+/g, " ");
        if (!driverAdvances[key]) driverAdvances[key] = [];
        driverAdvances[key].push({ amount: amt, fee: fee });
      }
    });
  }

  // 3. Fetch Master DB Loans
  let masterDB = getMasterDatabase();
  let driverLoans = {};
  let masterLoanData = [];
  let masterLoanSheet = null;

  if (masterDB) {
    masterLoanSheet = masterDB.getSheetByName("Loans");
    if (masterLoanSheet && masterLoanSheet.getLastRow() >= 2) {
      masterLoanData = masterLoanSheet.getRange(2, 1, masterLoanSheet.getLastRow() - 1, 8).getValues();
      masterLoanData.forEach((row, idx) => {
        let dName = String(row[0] || "").trim();
        let lDate = row[1];
        let totLoan = parseMoney(row[2]); // Col C
        let install = parseMoney(row[3]); // Col D
        let paidSoFar = parseMoney(row[4]); // Col E

        let remBal = totLoan - paidSoFar; // Auto calculate Remaining Balance
        row[7] = remBal;

        let lastPayDateStr = "";
        if (row[5] && !isNaN(new Date(row[5]).getTime())) { // Col F
          lastPayDateStr = Utilities.formatDate(new Date(row[5]), tz, "MM/dd/yyyy");
        }
        let lastPayAmt = parseMoney(row[6]); // Col G

        if (dName && lDate) {
          // Buffer Rule Check: Only count if Payslip Saturday is strictly greater than Loan Issue Saturday
          let loanSat = getSaturdayOfDate(lDate);
          if (stampDateObj.getTime() > loanSat.getTime()) {
            let key = dName.toLowerCase().replace(/\s+/g, " ");
            if (!driverLoans[key]) driverLoans[key] = [];
            driverLoans[key].push({
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

  // 4. Build PaySlip Sheet
  let paySheet = ss.getSheetByName("PaySlip");
  if (paySheet) ss.deleteSheet(paySheet);
  paySheet = ss.insertSheet("PaySlip");
  paySheet.setHiddenGridlines(true);
  paySheet.getRange("A:G").setBackground("white");

  let currentRow = 2;
  let totalAllBalances = 0;
  let masterLoanUpdates = [];

  let sortedDriverNames = Object.keys(payablesData.driverInfo).sort((a, b) => a.localeCompare(b));

  let discrepancyList = [];

  for (let name of sortedDriverNames) {
    let info = payablesData.driverInfo[name];
    let normKey = name.toLowerCase().replace(/\s+/g, " ");

    let activeRate = finalRates[name] !== undefined ? finalRates[name] : 0.90;

    let driverDaily = getDriverDailyMap(dailyBreakdown, name);
    let dailyDates = Object.keys(driverDaily).sort();

    let hasDaily = dailyDates.length > 0;
    let dailySumTrips = 0;
    let dailySumNet = 0;
    let dailySumTolls = 0;
    let dailySumCash = 0;

    if (hasDaily) {
      dailyDates.forEach(isoDate => {
        let day = driverDaily[isoDate];
        dailySumTrips += day.count;
        dailySumNet += day.net;
        dailySumTolls += day.tolls;
        dailySumCash += day.cash;
      });
    }

    // Safety Defense: If Payables sheet has trips (e.g. 84 for Koba) and Daily Balance doesn't match:
    // Flag discrepancy and DO NOT allow partial numbers to corrupt payroll totals!
    let tripDiscrepancy = (info.trips > 0 && hasDaily && dailySumTrips !== info.trips);
    if (tripDiscrepancy) {
      discrepancyList.push({
        driver: name,
        summaryTrips: info.trips,
        dailyTrips: dailySumTrips,
        daysCount: dailyDates.length
      });
    }

    let totalTrips = (!tripDiscrepancy && hasDaily) ? dailySumTrips : info.trips;
    let totalTolls = (!tripDiscrepancy && hasDaily) ? dailySumTolls : info.tolls;
    let totalNet = (!tripDiscrepancy && hasDaily) ? dailySumNet : info.net;
    let totalCash = (info.cash > 0) ? info.cash : (hasDaily ? dailySumCash : 0);

    let driverPay = totalNet * activeRate;
    let driverTotal = driverPay + totalTolls;

    // Header Row
    let headRange = paySheet.getRange(currentRow, 2, 1, 5);
    headRange.setValues([["Driver", "Num of Trips", "Price", "Tolls", "Driver Total"]])
      .setFontWeight("bold").setBackground(HEADER_BLUE).setBorder(true, false, true, false, false, false, BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID);
    paySheet.getRange(currentRow, 2).setHorizontalAlignment("left");
    paySheet.getRange(currentRow, 3).setHorizontalAlignment("center");
    paySheet.getRange(currentRow, 4, 1, 3).setHorizontalAlignment("right");
    currentRow++;

    // Summary Row
    let summaryRange = paySheet.getRange(currentRow, 2, 1, 5);
    summaryRange.setValues([[name, totalTrips, driverPay, totalTolls, driverTotal]]).setFontWeight("bold");
    paySheet.getRange(currentRow, 2).setHorizontalAlignment("left");
    paySheet.getRange(currentRow, 3).setHorizontalAlignment("center");
    paySheet.getRange(currentRow, 4, 1, 3).setHorizontalAlignment("right");
    summaryRange.setBorder(false, false, true, false, false, false, BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID);
    currentRow++;

    // Daily Rows from Daily Balance
    if (hasDaily) {
      dailyDates.forEach(isoDate => {
        let day = driverDaily[isoDate];
        let dayPrice = day.net * activeRate;
        let dayDriverTotal = dayPrice + day.tolls;
        paySheet.getRange(currentRow, 2, 1, 5).setValues([
          [day.displayDate, day.count, dayPrice, day.tolls, dayDriverTotal]
        ]);
        paySheet.getRange(currentRow, 2).setHorizontalAlignment("right");
        paySheet.getRange(currentRow, 3).setHorizontalAlignment("center");
        paySheet.getRange(currentRow, 4, 1, 3).setHorizontalAlignment("right");
        currentRow++;
      });
    }

    // Grand Total Row
    let grandRange = paySheet.getRange(currentRow, 2, 1, 4);
    grandRange.setValues([["Grand Total", totalTrips, driverPay, totalTolls]])
      .setFontWeight("bold").setBackground(HEADER_BLUE).setBorder(true, false, true, false, false, false, BORDER_BLUE, SpreadsheetApp.BorderStyle.SOLID);
    paySheet.getRange(currentRow, 2).setHorizontalAlignment("left");
    paySheet.getRange(currentRow, 3).setHorizontalAlignment("center");
    paySheet.getRange(currentRow, 4, 1, 2).setHorizontalAlignment("right");

    let totalCell = paySheet.getRange(currentRow, 6, 2, 1);
    totalCell.merge().setValue(driverTotal).setFontWeight("bold").setFontSize(20).setBackground(HEADER_BLUE)
      .setVerticalAlignment("middle").setHorizontalAlignment("center")
      .setBorder(true, true, true, true, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    currentRow += 2;

    // Sub: Cash Collected
    if (totalCash > 0) {
      paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Sub: Cash Collected").setHorizontalAlignment("center");
      paySheet.getRange(currentRow, 6).setValue(-totalCash).setHorizontalAlignment("right");
      currentRow++;
    }

    // Sub: Debits
    let debitAmt = Math.abs(info.debits || 0);
    if (debitAmt > 0) {
      paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Sub: Debits").setHorizontalAlignment("center");
      paySheet.getRange(currentRow, 6).setValue(-debitAmt).setHorizontalAlignment("right");
      currentRow++;
    }

    // Add: Credits
    let creditAmt = Math.abs(info.credits || 0);
    if (creditAmt > 0) {
      paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Add: Credits").setHorizontalAlignment("center");
      paySheet.getRange(currentRow, 6).setValue(creditAmt).setHorizontalAlignment("right");
      currentRow++;
    }

    // Advances
    let advanceSum = 0;
    let advList = driverAdvances[normKey] || driverAdvances[name];
    if (advList) {
      advList.forEach(adv => {
        if (adv.amount > 0) {
          paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Sub: Advanced Payment").setHorizontalAlignment("center");
          paySheet.getRange(currentRow, 6).setValue(-adv.amount).setHorizontalAlignment("right");
          advanceSum += adv.amount;
          currentRow++;
        }
        if (adv.fee > 0) {
          paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Sub: Fees").setHorizontalAlignment("center");
          paySheet.getRange(currentRow, 6).setValue(-adv.fee).setHorizontalAlignment("right");
          advanceSum += adv.fee;
          currentRow++;
        }
      });
    }

    // Refunds / Additions
    let addPayAmt = additionalPayments[name] || 0;
    if (addPayAmt !== 0) {
      paySheet.getRange(currentRow, 4, 1, 2).merge().setValue(addPayAmt > 0 ? "Add: Refund / Addition" : "Sub: Refund / Deduction").setHorizontalAlignment("center");
      paySheet.getRange(currentRow, 6).setValue(addPayAmt).setHorizontalAlignment("right");
      currentRow++;
    }

    // Loan Repayment (Advances & Additions First!)
    let preLoanBalance = driverTotal - totalCash + creditAmt - debitAmt - advanceSum + addPayAmt;
    let fundsAvailable = Math.max(0, preLoanBalance);
    let totalLoanDeducted = 0;

    let loanList = driverLoans[normKey] || driverLoans[name];
    if (loanList && fundsAvailable > 0) {
      loanList.sort((a, b) => new Date(a.date) - new Date(b.date));

      loanList.forEach(loan => {
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
              let origDate = parseTripDate(loan.date, tz) || new Date(loan.date);
              dStr = " (" + Utilities.formatDate(origDate, tz, "MM/dd/yy") + ")";
            }
            paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Sub: Loan Repayment" + dStr).setHorizontalAlignment("center");
            paySheet.getRange(currentRow, 6).setValue(-actualDeduct).setHorizontalAlignment("right");
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

    // Blank row before Driver Balance
    currentRow++;

    let finalBal = preLoanBalance - totalLoanDeducted;
    totalAllBalances += finalBal;

    paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("Driver Balance").setHorizontalAlignment("center");
    paySheet.getRange(currentRow, 6).setValue(finalBal).setHorizontalAlignment("right");
    paySheet.getRange(currentRow, 4, 1, 3).setBackground(HEADER_BLUE).setFontWeight("bold").setFontColor("black").setFontStyle("normal");
    currentRow += 3; // Leaves exactly 2 blank rows before the next driver
  }

  // TOTAL PAYOUT
  paySheet.getRange(currentRow, 4, 1, 2).merge().setValue("TOTAL PAYOUT").setHorizontalAlignment("center");
  paySheet.getRange(currentRow, 6).setValue(totalAllBalances).setHorizontalAlignment("right");
  paySheet.getRange(currentRow, 4, 1, 3).setBackground(HEADER_BLUE).setFontWeight("bold").setFontSize(12)
    .setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);

  paySheet.getRange(1, 3, currentRow, 1).setNumberFormat("#,##0").setHorizontalAlignment("center");
  paySheet.getRange(1, 4, currentRow + 1, 3).setNumberFormat("#,##0.00");
  paySheet.getRange(1, 6, currentRow + 1, 1).setHorizontalAlignment("right");
  paySheet.setColumnWidth(1, 60);
  paySheet.setColumnWidth(2, 160);
  paySheet.setColumnWidth(3, 100);
  paySheet.setColumnWidth(4, 200);
  paySheet.setColumnWidth(5, 100);
  paySheet.setColumnWidth(6, 120);

  if (paySheet.getMaxRows() > currentRow) paySheet.deleteRows(currentRow + 1, paySheet.getMaxRows() - currentRow);
  if (paySheet.getMaxColumns() > 7) paySheet.deleteColumns(8, paySheet.getMaxColumns() - 7);
  ss.setActiveSheet(paySheet);

  // Update Master DB Loans if needed
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

    masterLoanSheet.getRange(2, 1, masterLoanData.length, 8).setValues(masterLoanData);
    alertMessage += "\n\n(Master Database 'Loans' sheet has been securely updated. All balances auto-calculated. The system locked the dates so you can safely re-run this file without double counting!).";
  }

  if (discrepancyList.length > 0) {
    let discLines = discrepancyList.map(d =>
      "  • " + d.driver + ": Payables sheet has " + d.summaryTrips + " trips, but Daily Balance file only had " + d.dailyTrips + " trips (" + d.daysCount + " day row(s))."
    ).join("\n");

    let meta = dailyBreakdown._meta || {};
    let fileName = (meta.selectedFile && meta.selectedFile.name) ? meta.selectedFile.name : "Unknown File";
    let fileSource = (meta.selectedFile && meta.selectedFile.source) ? meta.selectedFile.source : "Drive Folder";

    let warnMsg = "⚠️ WARNING: TRIP COUNT DISCREPANCY DETECTED!\n\n" +
      "The script detected a mismatch between your Payables summary sheet and the parsed Daily Balance sheet:\n\n" +
      "📁 File Parsed: " + fileName + "\n" +
      "📍 Source: " + fileSource + "\n\n" +
      "Discrepancies found for " + discrepancyList.length + " driver(s):\n" +
      discLines + "\n\n" +
      "🛡️ SAFETY ACTION APPLIED:\n" +
      "To protect your payroll and avoid incorrect payouts, the verified summary totals (e.g. trips and balances) from your Payables sheet have been PRESERVED on the payslips.\n\n" +
      "👉 To diagnose which file Google Drive is opening, click 'AYOAH Automation -> 🔍 3. Diagnose Daily Balance File'.\n" +
      "👉 Or paste the exact sheet link directly into MONTHLY_BALANCE_FILE_URL at the top of Code.js.";

    SpreadsheetApp.getUi().alert(warnMsg);
  } else {
    SpreadsheetApp.getUi().alert(alertMessage);
  }
}

/**
 * Legacy Step 2: Generates Payslips for original 'Trip Information' sheet.
 */
function generateLegacyPayslip(ss, tripSheet, localSetSheet, finalRates, additionalPayments) {
  const HEADER_BLUE = "#CFE2F3";
  const BORDER_BLUE = "#3d85c6";
  const lastRow = tripSheet.getLastRow();

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
  // READ DATA AND BUILD PAYSLIP
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

    // Loan Repayment (Advances & Additions First!)
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

/**
 * Step 2: Master function. Generates Payslips AND Deducts Loans.
 * Dispatches to generatePayablesPayslip for new payables sheet or generateLegacyPayslip for legacy Trip Information sheet.
 */
function runCompletePayroll() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

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

  const payablesSheet = findPayablesSheet(ss);
  const tripSheet = ss.getSheetByName("Trip Information");

  if (payablesSheet && isPayablesSheet(payablesSheet)) {
    generatePayablesPayslip(ss, payablesSheet, localSetSheet, finalRates, additionalPayments);
    return;
  }

  if (tripSheet && tripSheet.getLastRow() >= 3) {
    generateLegacyPayslip(ss, tripSheet, localSetSheet, finalRates, additionalPayments);
    return;
  }

  return SpreadsheetApp.getUi().alert("Error: Neither 'AYOAH GO Drivers — payables' nor 'Trip Information' sheet was found.");
}

/**
 * Diagnostic Tool: Inspects Google Drive Daily Balance folders and files,
 * validates connectivity, checks pay period bounds, and tests parsing.
 * Can be run from the 'AYOAH Automation -> 🔍 3. Diagnose Daily Balance File' menu.
 */
function diagnoseDailyBalanceConnection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  let log = [];
  log.push("=== AYOAH DAILY BALANCE CONNECTION DIAGNOSTICS ===");
  log.push("Timestamp: " + Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss"));
  log.push("Configured Base Folder ID: " + DAILY_BALANCE_BASE_FOLDER_ID);
  log.push("Configured Direct File URL: " + (MONTHLY_BALANCE_FILE_URL || "(None set - using automatic Google Drive search)"));

  // 1. Payables Period
  let payablesSheet = findPayablesSheet(ss);
  let payablesData = payablesSheet ? parsePayablesSheet(payablesSheet, tz) : null;
  if (payablesData && payablesData.period) {
    log.push("\n--- PAY PERIOD DETECTED ---");
    log.push("Raw Period Text: " + payablesData.period.raw);
    log.push("Start Date: " + Utilities.formatDate(payablesData.period.start, tz, "yyyy-MM-dd"));
    log.push("End Date: " + Utilities.formatDate(payablesData.period.end, tz, "yyyy-MM-dd"));
    log.push("Drivers in Payables: " + Object.keys(payablesData.driverInfo).length);
    for (let d in payablesData.driverInfo) {
      let inf = payablesData.driverInfo[d];
      log.push(`  • ${d}: ${inf.trips} trips, Net $${inf.net.toFixed(2)}, Tolls $${inf.tolls.toFixed(2)}, Cash $${inf.cash.toFixed(2)}`);
    }
  } else {
    log.push("\n⚠️ Could not detect pay period in payables sheet!");
  }

  // 2. Drive Folder Inspection
  log.push("\n--- GOOGLE DRIVE FOLDER INSPECTION ---");
  try {
    let baseFolder = DriveApp.getFolderById(DAILY_BALANCE_BASE_FOLDER_ID);
    log.push("✅ Connected to Base Folder: '" + baseFolder.getName() + "'");

    let subfolders = baseFolder.getFolders();
    let foundSubfolders = false;
    while (subfolders.hasNext()) {
      foundSubfolders = true;
      let sf = subfolders.next();
      log.push(`  📁 [Folder] '${sf.getName()}' (ID: ${sf.getId()})`);

      if (sf.getName().includes("2026")) {
        let files = sf.getFiles();
        log.push(`    Files inside '${sf.getName()}':`);
        let fCount = 0;
        while (files.hasNext()) {
          let f = files.next();
          fCount++;
          let upd = Utilities.formatDate(f.getLastUpdated(), tz, "yyyy-MM-dd HH:mm");
          let trashed = f.isTrashed() ? " [TRASHED]" : "";
          log.push(`      📄 ${f.getName()} | Updated: ${upd}${trashed} | ID: ${f.getId()}`);
        }
        if (fCount === 0) log.push("      (No files found)");
      }
    }
    if (!foundSubfolders) log.push("  (No subfolders found)");
  } catch (e) {
    log.push("❌ Error connecting to Base Folder: " + e.message);
  }

  // 3. Test Daily Breakdown Fetch
  if (payablesData && payablesData.period) {
    log.push("\n--- TEST PARSING DAILY BREAKDOWN ---");
    let testBreakdown = fetchDailyBalanceBreakdown(payablesData.period, tz);
    let meta = testBreakdown._meta || {};
    if (meta.selectedFile) {
      log.push(`✅ Successfully selected file: '${meta.selectedFile.name}'`);
      log.push(`   Source: ${meta.selectedFile.source}`);
      log.push(`   ID: ${meta.selectedFile.id}`);
      log.push(`   Last Updated: ${meta.selectedFile.lastUpdated}`);
    } else {
      log.push(`⚠️ No file selected! Errors: ` + (meta.errors || []).join("; "));
    }

    log.push("\nParsed Driver Daily Breakdowns:");
    let driverKeys = Object.keys(testBreakdown).filter(k => k !== "_meta");
    if (driverKeys.length === 0) {
      log.push("  (No driver breakdowns parsed)");
    } else {
      driverKeys.forEach(k => {
        let days = testBreakdown[k];
        let dateKeys = Object.keys(days);
        let sumTrips = dateKeys.reduce((acc, dt) => acc + days[dt].count, 0);
        log.push(`  • ${k}: ${sumTrips} trips across ${dateKeys.length} date(s): [${dateKeys.join(", ")}]`);
      });
    }
  }

  // 4. Output to dedicated Diagnostics Sheet Tab
  let diagSheetName = "Daily_Balance_Diagnostics";
  let diagSheet = ss.getSheetByName(diagSheetName);
  if (diagSheet) ss.deleteSheet(diagSheet);
  diagSheet = ss.insertSheet(diagSheetName);
  diagSheet.getRange(1, 1).setValue(log.join("\n")).setFontFamily("Consolas").setFontSize(10);
  diagSheet.setColumnWidth(1, 800);
  ss.setActiveSheet(diagSheet);

  SpreadsheetApp.getUi().alert("🔍 Diagnostics Complete!\n\nA detailed report has been generated in the '" + diagSheetName + "' sheet tab.");
}

/**
 * 🧪 Test Runner: Tests reading the Daily Balance sheet without modifying any payslips.
 * Compares each driver's trip counts between Payables and Daily Balance,
 * writes an interactive verification table to a 'Daily_Balance_Test_Report' tab, and shows a summary alert.
 */
function testDailyBalanceReader() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();

  let payablesSheet = findPayablesSheet(ss);
  if (!payablesSheet) {
    SpreadsheetApp.getUi().alert("❌ Error: Could not find 'AYOAH GO Drivers — payables' sheet tab.");
    return;
  }

  let payablesData = parsePayablesSheet(payablesSheet, tz);
  if (!payablesData || !payablesData.period) {
    SpreadsheetApp.getUi().alert("❌ Error: Could not parse pay period or drivers from payables sheet.");
    return;
  }

  let period = payablesData.period;
  let pStartStr = Utilities.formatDate(period.start, tz, "yyyy-MM-dd");
  let pEndStr = Utilities.formatDate(period.end, tz, "yyyy-MM-dd");

  let dailyBreakdown = fetchDailyBalanceBreakdown(period, tz);
  let meta = dailyBreakdown._meta || {};
  let fileName = (meta.selectedFile && meta.selectedFile.name) ? meta.selectedFile.name : "Unknown File";
  let fileSource = (meta.selectedFile && meta.selectedFile.source) ? meta.selectedFile.source : "Drive Folder";

  let results = [];
  let matchCount = 0;
  let mismatchCount = 0;
  let summaryLines = [];

  let sortedDrivers = Object.keys(payablesData.driverInfo).sort((a, b) => a.localeCompare(b));

  for (let name of sortedDrivers) {
    let pInfo = payablesData.driverInfo[name];
    let pTrips = pInfo.trips || 0;

    let driverDaily = getDriverDailyMap(dailyBreakdown, name);
    let dayKeys = Object.keys(driverDaily).sort();

    let dTrips = 0;
    let dNet = 0;
    let dTolls = 0;
    let dCash = 0;
    let dayDetails = [];

    dayKeys.forEach(dt => {
      let d = driverDaily[dt];
      dTrips += d.count;
      dNet += d.net;
      dTolls += d.tolls;
      dCash += d.cash;
      dayDetails.push(`${d.displayDate || dt} (${d.count} trips)`);
    });

    let isMatch = (pTrips === dTrips);
    if (isMatch) matchCount++;
    else mismatchCount++;

    let statusText = isMatch ? "✅ MATCH" : "❌ MISMATCH";
    results.push([
      name,
      pTrips,
      dTrips,
      statusText,
      dayKeys.length,
      dayDetails.join(", ") || "(No days recorded)",
      dCash,
      dTolls,
      dNet
    ]);

    summaryLines.push(
      `• ${name}: Payables = ${pTrips} trips | Daily = ${dTrips} trips (${dayKeys.length} days) ${isMatch ? '✅' : '❌'}`
    );
  }

  // Create or refresh 'Daily_Balance_Test_Report' tab
  const TAB_NAME = "Daily_Balance_Test_Report";
  let testSheet = ss.getSheetByName(TAB_NAME);
  if (testSheet) ss.deleteSheet(testSheet);
  testSheet = ss.insertSheet(TAB_NAME);

  // Title Block
  testSheet.getRange(1, 1).setValue("AYOAH DAILY BALANCE READER - VERIFICATION REPORT")
    .setFontWeight("bold").setFontSize(14).setFontColor("#1a73e8");
  testSheet.getRange(2, 1).setValue(`File Read: ${fileName} | Source: ${fileSource} | Period: ${pStartStr} to ${pEndStr}`)
    .setFontStyle("italic").setFontColor("#5f6368");
  testSheet.getRange(3, 1).setValue(`Total Drivers: ${sortedDrivers.length} | Matches: ${matchCount} | Mismatches: ${mismatchCount}`)
    .setFontWeight("bold").setFontColor(mismatchCount === 0 ? "#137333" : "#d93025");

  // Table Headers
  const HEADERS = [
    "Driver Name",
    "Payables Trips",
    "Daily Balance Trips",
    "Match Status",
    "Days Count",
    "Daily Breakdown (Date & Trips)",
    "Total Cash",
    "Total Tolls",
    "Total Net Fare"
  ];
  let hRange = testSheet.getRange(5, 1, 1, HEADERS.length);
  hRange.setValues([HEADERS])
    .setFontWeight("bold")
    .setBackground("#1a73e8")
    .setFontColor("white")
    .setHorizontalAlignment("center");

  if (results.length > 0) {
    let dRange = testSheet.getRange(6, 1, results.length, HEADERS.length);
    dRange.setValues(results);

    // Color code match status & formatting
    for (let r = 0; r < results.length; r++) {
      let rowNum = 6 + r;
      let matched = results[r][3] === "✅ MATCH";
      testSheet.getRange(rowNum, 4).setBackground(matched ? "#e6f4ea" : "#fce8e6")
        .setFontColor(matched ? "#137333" : "#c5221f")
        .setFontWeight("bold")
        .setHorizontalAlignment("center");
      
      testSheet.getRange(rowNum, 2, 1, 2).setHorizontalAlignment("center");
      testSheet.getRange(rowNum, 5).setHorizontalAlignment("center");
      testSheet.getRange(rowNum, 7, 1, 3).setNumberFormat("$#,##0.00").setHorizontalAlignment("right");
    }
  }

  testSheet.autoResizeColumns(1, HEADERS.length);
  testSheet.setColumnWidth(6, 420); // Width for day breakdowns
  ss.setActiveSheet(testSheet);

  let alertTitle = mismatchCount === 0 ? "🎉 SUCCESS: All Trip Counts Match 100%!" : "⚠️ Discrepancies Found";
  let alertBody = `${alertTitle}\n\n` +
    `📁 File Parsed: ${fileName}\n` +
    `📅 Pay Period: ${pStartStr} to ${pEndStr}\n\n` +
    `Drivers Checked: ${sortedDrivers.length}\n` +
    `✅ Matches: ${matchCount}\n` +
    `❌ Mismatches: ${mismatchCount}\n\n` +
    summaryLines.join("\n") + "\n\n" +
    (mismatchCount === 0 
      ? "👉 Full day-by-day audit is open in the 'Daily_Balance_Test_Report' tab.\nYou can now safely run '2. Generate Payslip'!"
      : "👉 Check the 'Daily_Balance_Test_Report' tab for details.");

  SpreadsheetApp.getUi().alert(alertBody);
}