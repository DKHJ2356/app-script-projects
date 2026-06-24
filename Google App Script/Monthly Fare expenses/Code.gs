// ── Constants ──────────────────────────────────────────────────────────────
const FARE_LOG_TAB = "FareLog";
const SUMMARY_TAB  = "Summary";
const USERS_TAB    = "Users";
const SHEET_ID     = "15Rxvef6ULWpT1WP0ZaMcKjClfK77iP_rJBuVdhl0VBQ";
const PAUSE_KEY    = "PAUSED";

// Days to run: Sunday=0, Monday=1, Wednesday=3
const RUN_DAYS = [0, 1, 3];

// ── Sheet Setup ────────────────────────────────────────────────────────────

function setupSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // ── Summary tab ──
  let summary = ss.getSheetByName(SUMMARY_TAB);
  if (!summary) summary = ss.insertSheet(SUMMARY_TAB);
  summary.getRange("A1").setValue("");
  summary.getRange("B1").setValue("Year");
  summary.getRange("C1").setValue("Month");
  summary.getRange("D1").setValue("Name");
  summary.getRange("E1").setValue("Total Fare");
  summary.getRange("A1:E1").setFontWeight("bold");
  summary.getRange("A2").setValue("");
  summary.getRange("B2").setValue(new Date().getFullYear());
  summary.getRange("C2").setValue(new Date().toLocaleString("default", { month: "long" }));
  summary.getRange("D2").setValue("(All)");
  summary.getRange("E2").setFormula(
    `=IF(D2="(All)",` +
    `SUMIFS(FareLog!E:E,FareLog!B:B,B2,FareLog!C:C,C2),` +
    `SUMIFS(FareLog!E:E,FareLog!B:B,B2,FareLog!C:C,C2,FareLog!D:D,D2))`
  );

  // Row 3 — spacer, Row 4 — filtered log headers, Row 5 — FILTER formula
  summary.getRange("A3").setValue("");
  summary.getRange("A4:F4").setValues([["Date", "Year", "Month", "Name", "Fare", "Status"]]);
  summary.getRange("A4:F4").setFontWeight("bold");
  summary.getRange("A5").setFormula(
    `=IF(D2="(All)",,` +
    `IFERROR(FILTER(FareLog!A5:F,` +
    `FareLog!B5:B=B2,` +
    `FareLog!C5:C=C2,` +
    `FareLog!D5:D=D2),` +
    `"No entries found."))`
  );
  // Fix Bug 2 — format date column in Summary so FILTER result shows yyyy-mm-dd not serial numbers
  summary.getRange("A5:A1000").setNumberFormat("yyyy-mm-dd");

  summary.getRange("B2").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["2024","2025","2026","2027"], true).build()
  );
  summary.getRange("C2").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(
      ["January","February","March","April","May","June",
       "July","August","September","October","November","December"], true
    ).build()
  );

  // Name dropdown — built dynamically from the Users tab
  const usersSheet = ss.getSheetByName(USERS_TAB);
  const userNames  = usersSheet
    ? usersSheet.getDataRange().getValues().slice(1).filter(r => r[0]).map(r => String(r[0]).trim())
    : [];
  const nameList = ["(All)"].concat(userNames);
  summary.getRange("D2").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(nameList, true).build()
  );

  // ── FareLog tab ──
  let fareLog = ss.getSheetByName(FARE_LOG_TAB);
  if (!fareLog) fareLog = ss.insertSheet(FARE_LOG_TAB);

  // Row 1 — input zone headers
  fareLog.getRange("B1").setValue("Your Name").setFontWeight("bold");
  fareLog.getRange("C1").setValue("Fare Amount").setFontWeight("bold");

  // Row 2 — clear entire input row (fixes stale FILLED/PENDING in F2)
  fareLog.getRange("A2:F2").clearContent();
  fareLog.getRange("C2").setBackground("#fffbeb"); // yellow highlight = fare input cell

  // Name dropdown on B2 — populated from Users tab
  if (userNames.length) {
    fareLog.getRange("B2").setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(userNames, true).build()
    );
  }

  // Row 3 — blank divider
  fareLog.getRange("A3:F3").clearContent();

  // Row 4 — log headers
  fareLog.getRange("A4:F4").setValues([["Date", "Year", "Month", "Name", "Fare", "Status"]]);
  fareLog.getRange("A4:F4").setFontWeight("bold");

  // Format date column so dates display as yyyy-mm-dd instead of serial numbers
  fareLog.getRange("A5:A1000").setNumberFormat("yyyy-mm-dd");

  // ── Users tab ──
  let users = ss.getSheetByName(USERS_TAB);
  if (!users) users = ss.insertSheet(USERS_TAB);
  if (!users.getRange("A1").getValue()) {
    users.getRange("A1:C1").setValues([["Name", "Slack Member ID", "DM Webhook URL"]]);
    users.getRange("A1:C1").setFontWeight("bold");
  }

  SpreadsheetApp.flush();
  Logger.log("setupSheet() complete.");
}

// ── User Registry ──────────────────────────────────────────────────────────

function getUsers() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(USERS_TAB);
  const rows  = sheet.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[0] && r[1])
    .map(r => ({
      name:      String(r[0]).trim(),
      slackId:   String(r[1]).trim(),
      dmWebhook: String(r[2] || "").trim()
    }));
}

// ── Input Zone Submit (installable trigger) ────────────────────────────────

function onEdit(e) {
  const sheet = e.range.getSheet();
  if (sheet.getName() !== FARE_LOG_TAB) return;

  const cell = e.range.getA1Notation();

  // C2 edited — attempt fare submission
  if (cell === "C2") {
    const name = sheet.getRange("B2").getValue();
    const fare = sheet.getRange("C2").getValue();
    if (name && fare && Number(fare) > 0) submitFare();
  }

  // B2 edited — filter log rows by selected name
  if (cell === "B2") {
    applyNameFilter(sheet.getRange("B2").getValue());
  }
}

function submitFare() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(FARE_LOG_TAB);
  const name  = sheet.getRange("B2").getValue();
  const fare  = sheet.getRange("C2").getValue();
  if (!name || !fare) return;

  const now     = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const year    = now.getFullYear();
  const month   = now.toLocaleString("default", { month: "long" });

  // Fix Bug 1 — find existing PENDING row for same name+date and fill it instead of always appending
  const data = sheet.getDataRange().getValues();
  let pendingRowIndex = -1;
  for (let i = 4; i < data.length; i++) {
    if (String(data[i][0]) === dateStr && String(data[i][3]) === name && data[i][5] === "PENDING") {
      pendingRowIndex = i + 1; // convert to 1-based sheet row
      break;
    }
  }

  if (pendingRowIndex > -1) {
    // Fill the existing PENDING row — no new row added
    sheet.getRange(pendingRowIndex, 5).setValue(fare);
    sheet.getRange(pendingRowIndex, 6).setValue("FILLED");
  } else {
    // No PENDING row found — append a new FILLED row
    sheet.appendRow([dateStr, year, month, name, fare, "FILLED"]);
  }

  // Clear input zone — including F2 to prevent stale status display
  sheet.getRange("B2").clearContent();
  sheet.getRange("C2").clearContent();
  sheet.getRange("F2").clearContent();

  SpreadsheetApp.flush();
  Logger.log("submitFare() — logged: " + name + " | " + fare + " | " + dateStr);
}

// ── Name Filter (called by onEdit when B2 changes) ─────────────────────────

function applyNameFilter(selectedName) {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const sheet   = ss.getSheetByName(FARE_LOG_TAB);
  const lastRow = sheet.getLastRow();

  // Remove any existing filter first
  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  if (!selectedName) return; // no name — show all rows

  // Fix Bug 3 — guard against empty log (no data rows yet)
  if (lastRow < 5) return;

  // Apply filter on log range (row 4 header + rows 5+ data), column D = Name
  const logRange = sheet.getRange(4, 1, lastRow - 3, 6);
  const filter   = logRange.createFilter();
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextEqualTo(selectedName)
    .build();
  filter.setColumnFilterCriteria(4, criteria); // col D = position 4 in the range
}

// ── Daily Entry ────────────────────────────────────────────────────────────

function appendTodayRow() {
  if (!_isRunDay()) return;
  if (_isPaused()) return;

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(FARE_LOG_TAB);
  const now   = new Date();
  const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const year  = now.getFullYear();
  const month = now.toLocaleString("default", { month: "long" });

  const users = getUsers();

  // Fix Bug 4 — skip users who already have a row for today to prevent duplicates
  const existing = sheet.getDataRange().getValues();
  const alreadyLogged = new Set(
    existing.slice(4).filter(r => String(r[0]) === dateStr).map(r => String(r[3]))
  );
  const toAppend = users.filter(u => !alreadyLogged.has(u.name));
  toAppend.forEach(u => sheet.appendRow([dateStr, year, month, u.name, "", "PENDING"]));
  SpreadsheetApp.flush();

  if (toAppend.length > 0) sendDMs("new", dateStr, toAppend);
  Logger.log("appendTodayRow() — " + users.length + " rows appended for " + dateStr);
}

// ── Escalation Check ───────────────────────────────────────────────────────

function checkMissedFare() {
  if (!_isRunDay()) return;
  if (_isPaused()) return;

  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(FARE_LOG_TAB);
  const data  = sheet.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const users = getUsers();

  // Start at i=4 (row 5) — skips input zone rows 1-4
  for (let i = 4; i < data.length; i++) {
    const rowDate = String(data[i][0]);
    const name    = String(data[i][3]);
    const fare    = data[i][4];
    const status  = String(data[i][5]);

    if (status === "PENDING" && rowDate < today && !fare) {
      sheet.getRange(i + 1, 6).setValue("ESCALATED");
      const user = users.find(u => u.name === name);
      if (user) sendDMs("escalation", rowDate, [user]);
      Logger.log("checkMissedFare() — escalated row for " + name + " on " + rowDate);
    }

    if (status === "PENDING" && fare) {
      sheet.getRange(i + 1, 6).setValue("FILLED");
    }
  }

  SpreadsheetApp.flush();
}

// ── DM Sender ──────────────────────────────────────────────────────────────

function sendDMs(type, dateStr, users) {
  const sheetUrl = "https://docs.google.com/spreadsheets/d/" + SHEET_ID;
  const label    = _friendlyDate(dateStr);
  const userList = users || [];

  userList.forEach(u => {
    if (!u.dmWebhook) {
      Logger.log("sendDMs — no DM webhook for " + u.name + ", skipping.");
      return;
    }

    const mention = u.slackId ? `<@${u.slackId}>` : u.name;
    const text = type === "new"
      ? `📋 Hi ${mention}! Please log your fare for today (${label}). <${sheetUrl}|Open Sheet>`
      : `⚠️ Hi ${mention}, your fare for *${label}* was never filled in. Please update the sheet. <${sheetUrl}|Open Sheet>`;

    try {
      UrlFetchApp.fetch(u.dmWebhook, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ text }),
        muteHttpExceptions: true
      });
      Logger.log("sendDMs — DM sent to " + u.name);
    } catch (err) {
      Logger.log("sendDMs error for " + u.name + ": " + err);
    }
  });
}

// ── Status Column Formatting ───────────────────────────────────────────────

function applyStatusFormatting() {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const fareLog = ss.getSheetByName(FARE_LOG_TAB);

  // Log rows start at row 4 (rows 1-3 = input zone + divider)
  const statusRange = fareLog.getRange("F4:F1000");
  statusRange.setHorizontalAlignment("center");

  fareLog.clearConditionalFormatRules();

  const filledRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("FILLED")
    .setBackground("#dcfce7")
    .setFontColor("#166534")
    .setRanges([statusRange])
    .build();

  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("PENDING")
    .setBackground("#fef3c7")
    .setFontColor("#92400e")
    .setRanges([statusRange])
    .build();

  const escalatedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("ESCALATED")
    .setBackground("#fee2e2")
    .setFontColor("#991b1b")
    .setRanges([statusRange])
    .build();

  fareLog.setConditionalFormatRules([filledRule, pendingRule, escalatedRule]);
  SpreadsheetApp.flush();
  Logger.log("applyStatusFormatting() complete.");
}

// ── Fix Summary Formula ────────────────────────────────────────────────────

function fixSummaryFormula() {
  const ss      = SpreadsheetApp.openById(SHEET_ID);
  const summary = ss.getSheetByName(SUMMARY_TAB);
  summary.getRange("E2").setFormula(
    `=IF(D2="(All)",` +
    `SUMIFS(FareLog!E:E,FareLog!B:B,B2,FareLog!C:C,C2),` +
    `SUMIFS(FareLog!E:E,FareLog!B:B,B2,FareLog!C:C,C2,FareLog!D:D,D2))`
  );
  SpreadsheetApp.flush();
  Logger.log("fixSummaryFormula() complete.");
}

// ── Pause / Resume ─────────────────────────────────────────────────────────

function pauseReminders() {
  PropertiesService.getScriptProperties().setProperty(PAUSE_KEY, "true");
  Logger.log("Reminders paused.");
}

function resumeReminders() {
  PropertiesService.getScriptProperties().deleteProperty(PAUSE_KEY);
  Logger.log("Reminders resumed.");
}

// ── Slack Slash Command Handler (Web App) ──────────────────────────────────

function doPost(e) {
  const params  = e.parameter || {};
  const command = params.command || "";
  const text    = (params.text || "").trim().toLowerCase();
  let message;

  if (command === "/fare-pause" || text === "pause") {
    pauseReminders();
    message = "Reminders paused. No pings until you resume.";
  } else if (command === "/fare-resume" || text === "resume") {
    resumeReminders();
    message = "Reminders resumed. You'll be pinged on the next scheduled day.";
  } else {
    message = "Unknown command. Use `/fare-pause` or `/fare-resume`.";
  }

  return ContentService
    .createTextOutput(JSON.stringify({ response_type: "ephemeral", text: message }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Custom Menu ────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Fare Tools")
    .addItem("Setup Sheet",         "setupSheet")
    .addItem("Apply Status Colors", "applyStatusFormatting")
    .addItem("Fix Summary Formula", "fixSummaryFormula")
    .addSeparator()
    .addItem("Pause Reminders",     "pauseReminders")
    .addItem("Resume Reminders",    "resumeReminders")
    .addSeparator()
    .addItem("Append Today Row",    "appendTodayRow")
    .addItem("Check Missed Fare",   "checkMissedFare")
    .addToUi();
}

// ── Trigger Setup ──────────────────────────────────────────────────────────

function createTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("appendTodayRow").timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger("checkMissedFare").timeBased().everyDays(1).atHour(9).nearMinute(5).create();
  Logger.log("Triggers created.");
}

// ── Helpers ────────────────────────────────────────────────────────────────

function _isRunDay() {
  return RUN_DAYS.includes(new Date().getDay());
}

function _isPaused() {
  return PropertiesService.getScriptProperties().getProperty(PAUSE_KEY) === "true";
}

function _friendlyDate(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "EEE MMM d");
}
