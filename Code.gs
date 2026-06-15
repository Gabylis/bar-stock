// ═══════════════════════════════════════════════════════════════
//  DAILY LIQUOR TRANSFER LOG - Web App
//  Deploy as Web App: Execute as "User accessing the web app"
// ═══════════════════════════════════════════════════════════════

function doGet(e) {
  var page = e && e.parameter && e.parameter.page ? e.parameter.page : "setup";
  if (page === "order") {
    return HtmlService.createHtmlOutput(getOrderHTML())
      .setTitle("Quick Order")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "restock") {
    return HtmlService.createHtmlOutput(getRestockHTML())
      .setTitle("Restock")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (page === "weight") {
    return HtmlService.createHtmlOutput(getWeightHTML())
      .setTitle("Bottle Weight Log")
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  return HtmlService.createHtmlOutput(getSetupHTML())
    .setTitle("Bar Log Setup")
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function runSetup(config) {
  try {
    var ss   = _createSpreadsheet(config);
    var form = _createForm(ss, config);
    _installTriggers(ss);
    _syncBrands(ss, form, config.brands);
    PropertiesService.getUserProperties().setProperty("ALERT_EMAILS", config.alertEmails || "");
    SpreadsheetApp.flush();
    var ssUrl   = ss.getUrl();
    var formUrl = form.getPublishedUrl();
    var appUrl  = ScriptApp.getService().getUrl();
    // Save links for later retrieval
    PropertiesService.getUserProperties().setProperty("SS_URL",   ssUrl);
    PropertiesService.getUserProperties().setProperty("FORM_URL", formUrl);
    return { success: true, spreadsheetUrl: ssUrl, formUrl: formUrl, appUrl: appUrl };
  } catch(e) {
    Logger.log("runSetup: " + e.message);
    return { success: false, error: e.message };
  }
}

function saveAlertEmails(emails) {
  try {
    PropertiesService.getUserProperties().setProperty("ALERT_EMAILS", emails);
    return { success: true };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function getAppUrl() {
  return ScriptApp.getService().getUrl();
}

function getSavedLinks() {
  var props = PropertiesService.getUserProperties();
  return {
    ssUrl:   props.getProperty("SS_URL")        || "",
    formUrl: props.getProperty("FORM_URL")      || "",
    emails:  props.getProperty("ALERT_EMAILS")  || ""
  };
}

function getManagementScriptForCopy() {
  var formId = PropertiesService.getUserProperties().getProperty("FORM_ID") || "";
  var emails = PropertiesService.getUserProperties().getProperty("ALERT_EMAILS") || "";
  var arr    = JSON.stringify(emails.split(",").map(function(e){return e.trim();}).filter(Boolean));
  // Return the management script with FORM_ID and ALERT_EMAILS injected
  return getManagementScriptText().replace("__FORM_ID__", formId).replace("__ALERT_EMAILS__", arr);
}

// ── Quick Order: load inventory ──────────────────────────────
function loadInventory() {
  try {
    var ssId = PropertiesService.getUserProperties().getProperty("SS_ID");
    if (!ssId) return { success: false, error: "No spreadsheet linked yet." };
    var ss   = SpreadsheetApp.openById(ssId);
    var hInv = ss.getSheetByName("Inventory");
    if (!hInv) return { success: false, error: "Inventory sheet not found." };
    var last = hInv.getLastRow();
    if (last < 3) return { success: true, brands: [], bars: [] };
    var brands = hInv.getRange(3, 1, last-2, 8).getValues()
      .filter(function(r) { return r[0] !== ""; })
      .map(function(r) {
        return { name: r[0], category: r[1], stock: Number(r[5]), min: Number(r[6]) };
      });
    var shDash = ss.getSheetByName("Dashboard");
    var bars = [];
    if (shDash && shDash.getLastRow() >= 5) {
      bars = shDash.getRange(5, 1, Math.min(shDash.getLastRow()-4, 10), 1).getValues()
        .map(function(r) { return r[0]; }).filter(function(v) { return v !== ""; });
    }
    return { success: true, brands: brands, bars: bars };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Quick Order: submit ──────────────────────────────────────
function submitOrder(order) {
  try {
    var ssId = PropertiesService.getUserProperties().getProperty("SS_ID");
    if (!ssId) return { success: false, error: "No spreadsheet linked yet." };
    var ss   = SpreadsheetApp.openById(ssId);
    var hLog = ss.getSheetByName("Form_Responses");
    var hInv = ss.getSheetByName("Inventory");
    if (!hLog) return { success: false, error: "Form_Responses sheet not found." };
    var ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
    var count = 0;
    order.items.forEach(function(item) {
      if (!item.qty || item.qty <= 0) return;
      hLog.appendRow([ts, "Exit — Transfer to Bar", order.bar,
        item.name + " - " + item.category, item.qty,
        order.deliveredBy, order.receivedBy, "Quick Order"]);
      count++;
    });
    // Force spreadsheet to recalculate before checking stock
    SpreadsheetApp.flush();
    Utilities.sleep(2000);
    // Check low stock and send alerts if needed
    if (hInv && count > 0) {
      var lastItem = order.items.filter(function(i) { return i.qty > 0; }).pop();
      if (lastItem) {
        _checkLowStockWebApp(ss, hInv, lastItem.name, order.bar, order.deliveredBy, order.receivedBy);
      }
    }
    return { success: true, count: count };
  } catch(e) { return { success: false, error: e.message }; }
}

// ── Restock: submit entries ──────────────────────────────────
function submitRestock(order) {
  try {
    var ssId = PropertiesService.getUserProperties().getProperty("SS_ID");
    if (!ssId) return { success: false, error: "No spreadsheet linked yet." };
    var ss   = SpreadsheetApp.openById(ssId);
    var hLog = ss.getSheetByName("Form_Responses");
    if (!hLog) return { success: false, error: "Form_Responses sheet not found." };
    var ts    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy HH:mm:ss");
    var notes = "Restock" + (order.invoice ? " | Invoice: " + order.invoice : "");
    var count = 0;
    order.items.forEach(function(item) {
      if (!item.qty || item.qty <= 0) return;
      hLog.appendRow([ts, "Entry — Stock Received", "",
        item.name + " - " + item.category, item.qty,
        order.deliveredBy, order.receivedBy, notes]);
      count++;
    });
    SpreadsheetApp.flush();
    return { success: true, count: count };
  } catch(e) { return { success: false, error: e.message }; }
}


function submitWeight(data) {
  // data = { bar, recordedBy, date, items: [{name, category, oz}] }
  try {
    var ssId = PropertiesService.getUserProperties().getProperty("SS_ID");
    if (!ssId) return { success: false, error: "Not configured." };
    var ss = SpreadsheetApp.openById(ssId);

    // Get or create Weight Log sheet
    var shWt = ss.getSheetByName("Weight Log");
    if (!shWt) shWt = ss.insertSheet("Weight Log");
    // Always ensure correct headers
    if (shWt.getRange(1,1).getValue() !== "Date") {
      shWt.clearContents();
      shWt.setTabColor("#8E44AD");
      shWt.getRange(1,1,1,7)
        .setValues([["Date","Bar Section","Brand","Category","Full Oz","Oz Remaining","Recorded By"]])
        .setBackground("#1A1A2E").setFontColor("#FFFFFF").setFontWeight("bold");
      shWt.setFrozenRows(1);
      [120,160,180,120,100,120,140].forEach(function(w,i){shWt.setColumnWidth(i+1,w);});
    }

    var count = 0;
    data.items.forEach(function(item) {
      if (item.oz === "" || item.oz === null || item.oz === undefined) return;
      shWt.appendRow([data.date, data.bar, item.name, item.category, item.fullOz || "", Number(item.oz), data.recordedBy]);
      count++;
    });
    SpreadsheetApp.flush();
    return { success: true, count: count };
  } catch(e) { return { success: false, error: e.message }; }
}

function _checkLowStockWebApp(ss, hInv, brand, dest, delivered, received) {
  try {
    var emails = PropertiesService.getUserProperties().getProperty("ALERT_EMAILS") || "";
    var alertEmails = emails.split(",").map(function(e) { return e.trim(); }).filter(Boolean);
    if (!alertEmails.length) return;
    var last = hInv.getLastRow(); if (last < 3) return;
    var data = hInv.getRange(3, 1, last-2, 8).getValues();
    var alerts = [];
    data.forEach(function(row, i) {
      var c = Number(row[5]), m = Number(row[6]);
      var r = hInv.getRange(i+3, 1, 1, 8);
      if (c <= 0)        r.setBackground("#FADBD8");
      else if (c <= m) { r.setBackground("#FDEBD0"); alerts.push({brand:row[0],current:c,min:m}); }
      else               r.setBackground("#EAFAF1");
    });
    if (!alerts.length) return;
    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm");
    var rows = alerts.map(function(a) {
      return "<tr><td style='padding:8px;border:1px solid #ddd'>" + a.brand +
             "</td><td style='padding:8px;border:1px solid #ddd;color:#e74c3c;font-weight:bold'>" + a.current +
             "</td><td style='padding:8px;border:1px solid #ddd'>" + a.min + "</td></tr>";
    }).join("");
    MailApp.sendEmail({
      to: alertEmails.join(","),
      subject: "Port 27 - Low Stock Alert - Quick Order - " + date,
      htmlBody: "<h2>Low Stock Alert</h2>" +
        "<p>Triggered by Quick Order to <b>" + dest + "</b> by <b>" + delivered + "</b></p>" +
        "<table border=1 cellpadding=6 style='border-collapse:collapse'>" +
        "<tr style='background:#E94560;color:white'><th>Brand</th><th>Current Stock</th><th>Minimum</th></tr>" +
        rows + "</table>"
    });
  } catch(e) { Logger.log("_checkLowStockWebApp: " + e.message); }
}


// ─── Management script text (what gets pasted into the Spreadsheet) ────────
// Uses placeholders __FORM_ID__ and __ALERT_EMAILS__ replaced at copy time
function getManagementScriptText() {
  var s = "";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "//  DAILY LIQUOR TRANSFER LOG \u2014 Spreadsheet Manager" + "\n";
  s += "//  Paste into Extensions > Apps Script > Code.gs" + "\n";
  s += "//  Save \u2192 select installTriggers \u2192 click Run" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "" + "\n";
  s += "var FORM_ID      = \"__FORM_ID__\";" + "\n";
  s += "var ALERT_EMAILS = __ALERT_EMAILS__;" + "\n";
  s += "var SHEET_LOG    = \"Form_Responses\";" + "\n";
  s += "var SHEET_INV    = \"Inventory\";" + "\n";
  s += "var SHEET_DASH   = \"Dashboard\";" + "\n";
  s += "var SHEET_CHARTS = \"Charts\";" + "\n";
  s += "var SHEET_WEIGHT = \"Weight Log\";" + "\n";
  s += "" + "\n";
  s += "function onOpen() {" + "\n";
  s += "  SpreadsheetApp.getActiveSpreadsheet().addMenu(\"Liquor Log\", [" + "\n";
  s += "    { name: \"Sync brands to Form\",    functionName: \"syncBrandsToForm\"     }," + "\n";
  s += "    { name: \"Refresh Dashboard\",      functionName: \"refreshDashboard\"     }," + "\n";
  s += "    { name: \"Refresh Charts\",         functionName: \"refreshCharts\"        }," + "\n";
  s += "    { name: \"Send daily report now\",  functionName: \"sendDailySummary\"     }," + "\n";
  s += "    { name: \"---\",                    functionName: \"noop\"                 }," + "\n";
  s += "    { name: \"Fix Inventory formulas\", functionName: \"fixInventoryFormulas\" }," + "\n";
  s += "    { name: \"Install triggers\",       functionName: \"installTriggers\"      }," + "\n";
  s += "  ]);" + "\n";
  s += "}" + "\n";
  s += "function noop() {}" + "\n";
  s += "" + "\n";
  s += "function installTriggers() {" + "\n";
  s += "  ScriptApp.getProjectTriggers().forEach(function(t) { ScriptApp.deleteTrigger(t); });" + "\n";
  s += "  var ss = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "  ScriptApp.newTrigger(\"onFormSubmit\")" + "\n";
  s += "    .forSpreadsheet(ss).onFormSubmit().create();" + "\n";
  s += "  ScriptApp.newTrigger(\"onWeightLogChange\")" + "\n";
  s += "    .forSpreadsheet(ss).onChange().create();" + "\n";
  s += "  ScriptApp.newTrigger(\"sendDailySummary\")" + "\n";
  s += "    .timeBased().atHour(23).everyDays(1).create();" + "\n";
  s += "  SpreadsheetApp.getUi().alert(\"Triggers installed!\");" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function onWeightLogChange(e) {" + "\n";
  s += "  try {" + "\n";
  s += "    var ss    = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "    var sheet = ss.getActiveSheet();" + "\n";
  s += "    if (sheet.getName() !== \"Weight Log\") return;" + "\n";
  s += "    // Only refresh if a row was added" + "\n";
  s += "    if (e && e.changeType !== \"INSERT_ROW\" && e.changeType !== \"EDIT\") return;" + "\n";
  s += "    var hInv   = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "    var shDash = ss.getSheetByName(SHEET_DASH);" + "\n";
  s += "    if (!hInv || !shDash) return;" + "\n";
  s += "    var last   = hInv.getLastRow();" + "\n";
  s += "    var brands = last >= 3" + "\n";
  s += "      ? hInv.getRange(3,1,last-2,2).getValues()" + "\n";
  s += "          .filter(function(r){return r[0]!==\"\";})" + "\n";
  s += "          .map(function(r){return {name:r[0],category:r[1]};})" + "\n";
  s += "      : [];" + "\n";
  s += "    var bars = _getBarsFromDash(shDash);" + "\n";
  s += "    var shCharts = ss.getSheetByName(SHEET_CHARTS);" + "\n";
  s += "    if (!shCharts) shCharts = ss.insertSheet(SHEET_CHARTS);" + "\n";
  s += "    _buildCharts(ss, shCharts, shDash, brands, bars);" + "\n";
  s += "  } catch(e) { Logger.log(\"onWeightLogChange: \" + e.message); }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function onFormSubmit(e) {" + "\n";
  s += "  try {" + "\n";
  s += "    var ss   = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "    var hLog = ss.getSheetByName(SHEET_LOG);" + "\n";
  s += "    var hInv = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "    var last = hLog.getLastRow();" + "\n";
  s += "    if (!hLog.getRange(last, 1).getValue())" + "\n";
  s += "      hLog.getRange(last, 1).setValue(new Date()).setNumberFormat(\"MM/DD/YYYY HH:mm\");" + "\n";
  s += "    var row = hLog.getRange(last, 1, 1, 8).getValues()[0];" + "\n";
  s += "    _checkLowStock(hInv, row[3], Number(row[4]), row[1], row[2], row[5], row[6]);" + "\n";
  s += "  } catch(e) { Logger.log(\"onFormSubmit: \" + e.message); }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function syncBrandsToForm() {" + "\n";
  s += "  try {" + "\n";
  s += "    if (!FORM_ID) { SpreadsheetApp.getUi().alert(\"No Form ID set.\"); return; }" + "\n";
  s += "    var ss     = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "    var hInv   = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "    var last   = hInv.getLastRow();" + "\n";
  s += "    if (last < 3) return;" + "\n";
  s += "    var brands = hInv.getRange(3, 1, last-2, 2).getValues()" + "\n";
  s += "      .filter(function(r) { return r[0] !== \"\"; })" + "\n";
  s += "      .map(function(r) { return r[0] + (r[1] ? \" - \" + r[1] : \"\"); });" + "\n";
  s += "    var form = FormApp.openById(FORM_ID);" + "\n";
  s += "    var q    = form.getItems().filter(function(i) {" + "\n";
  s += "      return i.getTitle().trim() === \"Liquor / Brand\";" + "\n";
  s += "    })[0];" + "\n";
  s += "    if (q && q.getType() === FormApp.ItemType.LIST)" + "\n";
  s += "      q.asListItem().setChoiceValues(brands);" + "\n";
  s += "    SpreadsheetApp.getUi().alert(\"Form updated with \" + brands.length + \" brands.\");" + "\n";
  s += "  } catch(e) { SpreadsheetApp.getUi().alert(\"Error: \" + e.message); }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function refreshDashboard() {" + "\n";
  s += "  try {" + "\n";
  s += "    var ss     = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "    var hInv   = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "    var shDash = ss.getSheetByName(SHEET_DASH);" + "\n";
  s += "    if (!hInv || !shDash) { SpreadsheetApp.getUi().alert(\"Sheets not found.\"); return; }" + "\n";
  s += "    var last   = hInv.getLastRow();" + "\n";
  s += "    var brands = last >= 3" + "\n";
  s += "      ? hInv.getRange(3, 1, last-2, 2).getValues()" + "\n";
  s += "          .filter(function(r) { return r[0] !== \"\"; })" + "\n";
  s += "          .map(function(r) { return { name: r[0], category: r[1] }; })" + "\n";
  s += "      : [];" + "\n";
  s += "    var bars = _getBarsFromDash(shDash);" + "\n";
  s += "    _buildDashboard(shDash, brands, bars);" + "\n";
  s += "    SpreadsheetApp.getUi().alert(\"Dashboard refreshed.\");" + "\n";
  s += "  } catch(e) { SpreadsheetApp.getUi().alert(\"Error: \" + e.message); }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function refreshCharts() {" + "\n";
  s += "  try {" + "\n";
  s += "    var ss     = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "    var hInv   = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "    var shDash = ss.getSheetByName(SHEET_DASH);" + "\n";
  s += "    if (!hInv || !shDash) { SpreadsheetApp.getUi().alert(\"Sheets not found.\"); return; }" + "\n";
  s += "    var last   = hInv.getLastRow();" + "\n";
  s += "    var brands = last >= 3" + "\n";
  s += "      ? hInv.getRange(3, 1, last-2, 2).getValues()" + "\n";
  s += "          .filter(function(r) { return r[0] !== \"\"; })" + "\n";
  s += "          .map(function(r) { return { name: r[0], category: r[1] }; })" + "\n";
  s += "      : [];" + "\n";
  s += "    var bars = _getBarsFromDash(shDash);" + "\n";
  s += "    // Get or create Charts sheet" + "\n";
  s += "    var shCharts = ss.getSheetByName(SHEET_CHARTS);" + "\n";
  s += "    if (!shCharts) shCharts = ss.insertSheet(SHEET_CHARTS);" + "\n";
  s += "    shCharts.setTabColor(\"#8E44AD\");" + "\n";
  s += "    _buildCharts(ss, shCharts, shDash, brands, bars);" + "\n";
  s += "    SpreadsheetApp.getUi().alert(\"Charts refreshed.\");" + "\n";
  s += "  } catch(e) { SpreadsheetApp.getUi().alert(\"Error: \" + e.message); }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function fixInventoryFormulas() {" + "\n";
  s += "  var ss   = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "  var hInv = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "  var last = hInv.getLastRow();" + "\n";
  s += "  for (var row = 3; row <= last; row++) {" + "\n";
  s += "    if (!hInv.getRange(row, 1).getValue()) continue;" + "\n";
  s += "    hInv.getRange(row, 4).setFormula(" + "\n";
  s += "      '=SUMPRODUCT((ISNUMBER(SEARCH(A' + row + ',Form_Responses!D:D)))' +" + "\n";
  s += "      '*(ISNUMBER(SEARCH(\"Entry\",Form_Responses!B:B)))' +" + "\n";
  s += "      '*IFERROR(VALUE(Form_Responses!E:E),0))'" + "\n";
  s += "    );" + "\n";
  s += "    hInv.getRange(row, 5).setFormula(" + "\n";
  s += "      '=SUMPRODUCT((ISNUMBER(SEARCH(A' + row + ',Form_Responses!D:D)))' +" + "\n";
  s += "      '*(ISNUMBER(SEARCH(\"Exit\",Form_Responses!B:B)))' +" + "\n";
  s += "      '*IFERROR(VALUE(Form_Responses!E:E),0))'" + "\n";
  s += "    );" + "\n";
  s += "    hInv.getRange(row, 6).setFormula('=C' + row + '+D' + row + '-E' + row);" + "\n";
  s += "    hInv.getRange(row, 8).setFormula(" + "\n";
  s += "      '=IF(F' + row + '<=0,\"OUT OF STOCK\",IF(F' + row + '<=G' + row + ',\"LOW\",\"OK\"))'" + "\n";
  s += "    );" + "\n";
  s += "  }" + "\n";
  s += "  SpreadsheetApp.getUi().alert(\"Formulas fixed for \" + (last-2) + \" brands.\");" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function sendDailySummary() {" + "\n";
  s += "  if (!ALERT_EMAILS || ALERT_EMAILS.length === 0) return;" + "\n";
  s += "  var ss      = SpreadsheetApp.getActiveSpreadsheet();" + "\n";
  s += "  var hInv    = ss.getSheetByName(SHEET_INV);" + "\n";
  s += "  var hLog    = ss.getSheetByName(SHEET_LOG);" + "\n";
  s += "  var emails  = ALERT_EMAILS.join(\",\");" + "\n";
  s += "  var date    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), \"MM/dd/yyyy\");" + "\n";
  s += "  var today   = new Date(); today.setHours(0,0,0,0);" + "\n";
  s += "  var logData = hLog.getRange(3, 1, Math.max(hLog.getLastRow()-2, 1), 8).getValues();" + "\n";
  s += "  var mov     = logData.filter(function(r) {" + "\n";
  s += "    var d = new Date(r[0]); d.setHours(0,0,0,0);" + "\n";
  s += "    return d.getTime() === today.getTime();" + "\n";
  s += "  });" + "\n";
  s += "  var lr = mov.map(function(r) {" + "\n";
  s += "    return \"<tr><td>\" + r[1] + \"</td><td>\" + r[3] + \"</td><td>\" + r[4] +" + "\n";
  s += "           \"</td><td>\" + (r[2]||\"-\") + \"</td><td>\" + r[5] + \"</td><td>\" + r[6] + \"</td></tr>\";" + "\n";
  s += "  }).join(\"\");" + "\n";
  s += "  var invData = hInv.getRange(3, 1, Math.max(hInv.getLastRow()-2, 1), 8).getValues();" + "\n";
  s += "  var ir = invData.map(function(r) {" + "\n";
  s += "    var s = Number(r[5]), m = Number(r[6]);" + "\n";
  s += "    var bg = s<=0 ? \"#FADBD8\" : s<=m ? \"#FDEBD0\" : \"#EAFAF1\";" + "\n";
  s += "    return '<tr style=\"background:' + bg + '\"><td>' + r[0] + \"</td><td>\" + r[1] +" + "\n";
  s += "           \"</td><td>\" + s + \"</td><td>\" + m + \"</td><td>\" + r[7] + \"</td></tr>\";" + "\n";
  s += "  }).join(\"\");" + "\n";
  s += "  MailApp.sendEmail({" + "\n";
  s += "    to: emails," + "\n";
  s += "    subject: \"Port 27 - Daily Bar Report - \" + date," + "\n";
  s += "    htmlBody: \"<h2>Daily Bar Report - \" + date + \"</h2>\" +" + "\n";
  s += "      \"<h3>Movements (\" + mov.length + \")</h3>\" +" + "\n";
  s += "      \"<table border=1 cellpadding=6><tr><th>Type</th><th>Brand</th><th>Qty</th>\" +" + "\n";
  s += "      \"<th>Dest</th><th>Delivered</th><th>Received</th></tr>\" + lr + \"</table>\" +" + "\n";
  s += "      \"<h3>Inventory</h3><table border=1 cellpadding=6><tr><th>Brand</th><th>Category</th>\" +" + "\n";
  s += "      \"<th>Stock</th><th>Min</th><th>Status</th></tr>\" + ir + \"</table>\"" + "\n";
  s += "  });" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "//  BUILD DASHBOARD" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "function _buildDashboard(shDash, brands, bars) {" + "\n";
  s += "  shDash.clear();" + "\n";
  s += "" + "\n";
  s += "  // Title" + "\n";
  s += "  shDash.getRange(1,1,1,9).merge()" + "\n";
  s += "    .setValue(\"DASHBOARD - Summary by Bar Section & Brand\")" + "\n";
  s += "    .setBackground(\"#1A1A2E\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "    .setFontSize(13).setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "  shDash.setRowHeight(1, 44);" + "\n";
  s += "" + "\n";
  s += "  // Section 1: Consumption by Section" + "\n";
  s += "  shDash.getRange(\"A3:D3\").merge().setValue(\"CONSUMPTION BY SECTION\")" + "\n";
  s += "    .setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "    .setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "  [\"Section\",\"Total Bottles\",\"Distinct Brands\",\"Last Transfer\"].forEach(function(h,i) {" + "\n";
  s += "    shDash.getRange(4,i+1).setValue(h)" + "\n";
  s += "      .setBackground(\"#E94560\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "      .setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "    shDash.setColumnWidth(i+1, [180,140,150,180][i]);" + "\n";
  s += "  });" + "\n";
  s += "  bars.forEach(function(bar, i) {" + "\n";
  s += "    var row = i+5;" + "\n";
  s += "    shDash.getRange(row,1).setValue(bar).setFontWeight(\"bold\");" + "\n";
  s += "    shDash.getRange(row,2).setFormula(" + "\n";
  s += "      '=IFERROR(SUMIFS(Form_Responses!E:E,Form_Responses!C:C,A'+row+',Form_Responses!B:B,\"Exit*\"),0)'" + "\n";
  s += "    );" + "\n";
  s += "    shDash.getRange(row,3).setFormula(" + "\n";
  s += "      '=IFERROR(SUMPRODUCT(((Form_Responses!C$2:Form_Responses!C$1000)=A'+row+')' +" + "\n";
  s += "      '*(ISNUMBER(SEARCH(\"Exit\",Form_Responses!B$2:Form_Responses!B$1000)))' +" + "\n";
  s += "      '*((Form_Responses!D$2:Form_Responses!D$1000)<>\"\")/COUNTIF(' +" + "\n";
  s += "      'Form_Responses!D$2:Form_Responses!D$1000,Form_Responses!D$2:Form_Responses!D$1000&\"\")),0)'" + "\n";
  s += "    );" + "\n";
  s += "    shDash.getRange(row,4).setFormula(" + "\n";
  s += "      '=IFERROR(TEXT(MAXIFS(Form_Responses!A:A,Form_Responses!C:C,A'+row+',Form_Responses!B:B,\"Exit*\"),\"MM/DD/YYYY HH:MM\"),\"-\")'" + "\n";
  s += "    );" + "\n";
  s += "    shDash.getRange(row,1,1,4).setHorizontalAlignment(\"center\")" + "\n";
  s += "      .setBorder(true,true,true,true,true,true,\"#CCCCCC\",SpreadsheetApp.BorderStyle.SOLID);" + "\n";
  s += "    if (i%2===1) shDash.getRange(row,1,1,4).setBackground(\"#EBF5FB\");" + "\n";
  s += "  });" + "\n";
  s += "  shDash.setColumnWidth(5, 30);" + "\n";
  s += "" + "\n";
  s += "  // Section 2: Top Brands by Exits" + "\n";
  s += "  shDash.getRange(\"F3:I3\").merge().setValue(\"TOP BRANDS BY EXITS\")" + "\n";
  s += "    .setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "    .setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "  [\"Brand\",\"Total Exits\",\"Total Entries\",\"Current Stock\"].forEach(function(h,i) {" + "\n";
  s += "    shDash.getRange(4,i+6).setValue(h)" + "\n";
  s += "      .setBackground(\"#E94560\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "      .setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "    shDash.setColumnWidth(i+6, [200,120,120,120][i]);" + "\n";
  s += "  });" + "\n";
  s += "  brands.forEach(function(b, i) {" + "\n";
  s += "    var dr = i+5, ir = i+3;" + "\n";
  s += "    shDash.getRange(dr,6).setFormula('=Inventory!A'+ir+'&\" - \"&Inventory!B'+ir).setFontWeight(\"bold\");" + "\n";
  s += "    shDash.getRange(dr,7).setFormula('=Inventory!E'+ir);" + "\n";
  s += "    shDash.getRange(dr,8).setFormula('=Inventory!D'+ir);" + "\n";
  s += "    shDash.getRange(dr,9).setFormula('=Inventory!F'+ir);" + "\n";
  s += "    shDash.getRange(dr,6,1,4).setHorizontalAlignment(\"center\")" + "\n";
  s += "      .setBorder(true,true,true,true,true,true,\"#CCCCCC\",SpreadsheetApp.BorderStyle.SOLID);" + "\n";
  s += "    if (i%2===1) shDash.getRange(dr,6,1,4).setBackground(\"#EBF5FB\");" + "\n";
  s += "  });" + "\n";
  s += "" + "\n";
  s += "  // Section 3: Consumption by Category" + "\n";
  s += "  var s3 = brands.length + 7;" + "\n";
  s += "  shDash.getRange(s3,1,1,9).merge()" + "\n";
  s += "    .setValue(\"CONSUMPTION BY BAR SECTION & CATEGORY\")" + "\n";
  s += "    .setBackground(\"#1A1A2E\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "    .setFontSize(12).setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "  shDash.setRowHeight(s3, 36);" + "\n";
  s += "  var cHdr = s3+1;" + "\n";
  s += "  [\"Bar Section\",\"Category\",\"Total Bottles\",\"% of Section\"].forEach(function(h,i) {" + "\n";
  s += "    shDash.getRange(cHdr,i+1).setValue(h)" + "\n";
  s += "      .setBackground(\"#E94560\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "      .setFontWeight(\"bold\").setFontSize(10).setHorizontalAlignment(\"center\");" + "\n";
  s += "  });" + "\n";
  s += "  var catSet = {};" + "\n";
  s += "  brands.forEach(function(b) { if (b.category) catSet[b.category] = true; });" + "\n";
  s += "  var cats = Object.keys(catSet);" + "\n";
  s += "  var cRow = cHdr+1;" + "\n";
  s += "  bars.forEach(function(bar, bi) {" + "\n";
  s += "    cats.forEach(function(cat, ci) {" + "\n";
  s += "      var cb = brands.filter(function(b) { return b.category === cat; });" + "\n";
  s += "      if (!cb.length) return;" + "\n";
  s += "      shDash.getRange(cRow,1).setValue(bar).setHorizontalAlignment(\"center\");" + "\n";
  s += "      shDash.getRange(cRow,2).setValue(cat).setHorizontalAlignment(\"center\");" + "\n";
  s += "      shDash.getRange(cRow,3).setFormula(" + "\n";
  s += "        '=IFERROR(SUMPRODUCT(' +" + "\n";
  s += "        '(Form_Responses!C:C=\"' + bar + '\")' +" + "\n";
  s += "        '*(ISNUMBER(SEARCH(\"Exit\",Form_Responses!B:B)))' +" + "\n";
  s += "        '*(IFERROR(VLOOKUP(IFERROR(LEFT(Form_Responses!D:D,FIND(\" - \",Form_Responses!D:D)-1),Form_Responses!D:D),Inventory!A:B,2,0),\"\")=\"' + cat + '\")' +" + "\n";
  s += "        '*IFERROR(VALUE(Form_Responses!E:E),0)),0)'" + "\n";
  s += "      );" + "\n";
  s += "      shDash.getRange(cRow,4).setFormula(" + "\n";
  s += "        '=IFERROR(IF(SUMIFS(Form_Responses!E:E,Form_Responses!C:C,\"' + bar +" + "\n";
  s += "        '\",Form_Responses!B:B,\"Exit*\")=0,\"-\",TEXT(C' + cRow +" + "\n";
  s += "        '/SUMIFS(Form_Responses!E:E,Form_Responses!C:C,\"' + bar +" + "\n";
  s += "        '\",Form_Responses!B:B,\"Exit*\"),\"0.0%\")),\"-\")'" + "\n";
  s += "      );" + "\n";
  s += "      shDash.getRange(cRow,1,1,4)" + "\n";
  s += "        .setBorder(true,true,true,true,true,true,\"#CCCCCC\",SpreadsheetApp.BorderStyle.SOLID);" + "\n";
  s += "      if ((bi+ci)%2===1) shDash.getRange(cRow,1,1,4).setBackground(\"#EBF5FB\");" + "\n";
  s += "      cRow++;" + "\n";
  s += "    });" + "\n";
  s += "  });" + "\n";
  s += "  shDash.setFrozenRows(2);" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "//  BUILD CHARTS SHEET" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "function _buildCharts(ss, shCharts, shDash, brands, bars) {" + "\n";
  s += "  // Clear existing charts and data" + "\n";
  s += "  shCharts.clearContents();" + "\n";
  s += "  var existingCharts = shCharts.getCharts();" + "\n";
  s += "  existingCharts.forEach(function(c) { shCharts.removeChart(c); });" + "\n";
  s += "" + "\n";
  s += "  shCharts.getRange(1,1,1,6).merge()" + "\n";
  s += "    .setValue(\"CHARTS - Visual Summary\")" + "\n";
  s += "    .setBackground(\"#1A1A2E\").setFontColor(\"#FFFFFF\")" + "\n";
  s += "    .setFontSize(13).setFontWeight(\"bold\").setHorizontalAlignment(\"center\");" + "\n";
  s += "  shCharts.setRowHeight(1, 44);" + "\n";
  s += "" + "\n";
  s += "  // \u2500\u2500 Chart 1: Exits by Bar Section \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500" + "\n";
  s += "  // Data for chart 1: pull from Dashboard section 1" + "\n";
  s += "  shCharts.getRange(3,1).setValue(\"Exits by Bar Section\")" + "\n";
  s += "    .setFontWeight(\"bold\").setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\");" + "\n";
  s += "  shCharts.getRange(3,2).setValue(\"Total Bottles\")" + "\n";
  s += "    .setFontWeight(\"bold\").setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\");" + "\n";
  s += "  bars.forEach(function(bar, i) {" + "\n";
  s += "    shCharts.getRange(4+i, 1).setFormula('=Dashboard!A' + (i+5));" + "\n";
  s += "    shCharts.getRange(4+i, 2).setFormula('=Dashboard!B' + (i+5));" + "\n";
  s += "  });" + "\n";
  s += "  shCharts.setColumnWidth(1, 180);" + "\n";
  s += "  shCharts.setColumnWidth(2, 140);" + "\n";
  s += "" + "\n";
  s += "  var chart1DataRange = shCharts.getRange(3, 1, bars.length+1, 2);" + "\n";
  s += "  var chart1 = shCharts.newChart()" + "\n";
  s += "    .setChartType(Charts.ChartType.PIE)" + "\n";
  s += "    .addRange(chart1DataRange)" + "\n";
  s += "    .setPosition(3, 4, 0, 0)" + "\n";
  s += "    .setOption(\"title\", \"Exits by Bar Section\")" + "\n";
  s += "    .setOption(\"width\", 420)" + "\n";
  s += "    .setOption(\"height\", 280)" + "\n";
  s += "    .setOption(\"pieHole\", 0.4)" + "\n";
  s += "    .setOption(\"legend\", {position: \"right\"})" + "\n";
  s += "    .setOption(\"colors\", [\"#E94560\",\"#0F3460\",\"#27AE60\",\"#F39C12\",\"#8E44AD\"])" + "\n";
  s += "    .build();" + "\n";
  s += "  shCharts.insertChart(chart1);" + "\n";
  s += "" + "\n";
  s += "  // \u2500\u2500 Chart 2: Exits by Category \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500" + "\n";
  s += "  var cat2Row = bars.length + 6;" + "\n";
  s += "  shCharts.getRange(cat2Row, 1).setValue(\"Category\")" + "\n";
  s += "    .setFontWeight(\"bold\").setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\");" + "\n";
  s += "  shCharts.getRange(cat2Row, 2).setValue(\"Total Bottles\")" + "\n";
  s += "    .setFontWeight(\"bold\").setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\");" + "\n";
  s += "" + "\n";
  s += "  // Aggregate category totals across all bars from Dashboard section 3" + "\n";
  s += "  var catSet = {};" + "\n";
  s += "  brands.forEach(function(b) { if (b.category) catSet[b.category] = true; });" + "\n";
  s += "  var cats = Object.keys(catSet);" + "\n";
  s += "  var dashCatStartRow = brands.length + 9; // approximate start of category section in Dashboard" + "\n";
  s += "" + "\n";
  s += "  // Simpler: compute category totals directly from Form_Responses" + "\n";
  s += "  cats.forEach(function(cat, ci) {" + "\n";
  s += "    shCharts.getRange(cat2Row+1+ci, 1).setValue(cat);" + "\n";
  s += "    shCharts.getRange(cat2Row+1+ci, 2).setFormula(" + "\n";
  s += "      '=IFERROR(SUMPRODUCT(' +" + "\n";
  s += "      '(ISNUMBER(SEARCH(\"Exit\",Form_Responses!B:B)))' +" + "\n";
  s += "      '*(IFERROR(VLOOKUP(IFERROR(LEFT(Form_Responses!D:D,FIND(\" - \",Form_Responses!D:D)-1),Form_Responses!D:D),Inventory!A:B,2,0),\"\")=\"' + cat + '\")' +" + "\n";
  s += "      '*IFERROR(VALUE(Form_Responses!E:E),0)),0)'" + "\n";
  s += "    );" + "\n";
  s += "  });" + "\n";
  s += "" + "\n";
  s += "  var chart2DataRange = shCharts.getRange(cat2Row, 1, cats.length+1, 2);" + "\n";
  s += "  var chart2 = shCharts.newChart()" + "\n";
  s += "    .setChartType(Charts.ChartType.PIE)" + "\n";
  s += "    .addRange(chart2DataRange)" + "\n";
  s += "    .setPosition(cat2Row, 4, 0, 0)" + "\n";
  s += "    .setOption(\"title\", \"Exits by Category\")" + "\n";
  s += "    .setOption(\"width\", 420)" + "\n";
  s += "    .setOption(\"height\", 280)" + "\n";
  s += "    .setOption(\"pieHole\", 0.4)" + "\n";
  s += "    .setOption(\"legend\", {position: \"right\"})" + "\n";
  s += "    .setOption(\"colors\", [\"#E94560\",\"#0F3460\",\"#27AE60\",\"#F39C12\",\"#8E44AD\",\"#1ABC9C\",\"#E67E22\"])" + "\n";
  s += "    .build();" + "\n";
  s += "  shCharts.insertChart(chart2);" + "\n";
  s += "" + "\n";
  s += "  // \u2500\u2500 Chart 3: Top Brands \u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500\u2500" + "\n";
  s += "  var brand3Row = cat2Row + cats.length + 3;" + "\n";
  s += "  shCharts.getRange(brand3Row, 1).setValue(\"Brand\")" + "\n";
  s += "    .setFontWeight(\"bold\").setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\");" + "\n";
  s += "  shCharts.getRange(brand3Row, 2).setValue(\"Total Exits\")" + "\n";
  s += "    .setFontWeight(\"bold\").setBackground(\"#0F3460\").setFontColor(\"#FFFFFF\");" + "\n";
  s += "  brands.forEach(function(b, i) {" + "\n";
  s += "    shCharts.getRange(brand3Row+1+i, 1).setFormula('=Dashboard!F' + (i+5));" + "\n";
  s += "    shCharts.getRange(brand3Row+1+i, 2).setFormula('=Dashboard!G' + (i+5));" + "\n";
  s += "  });" + "\n";
  s += "" + "\n";
  s += "  var chart3DataRange = shCharts.getRange(brand3Row, 1, brands.length+1, 2);" + "\n";
  s += "  var chart3 = shCharts.newChart()" + "\n";
  s += "    .setChartType(Charts.ChartType.BAR)" + "\n";
  s += "    .addRange(chart3DataRange)" + "\n";
  s += "    .setPosition(brand3Row, 4, 0, 0)" + "\n";
  s += "    .setOption(\"title\", \"Top Brands by Exits\")" + "\n";
  s += "    .setOption(\"width\", 500)" + "\n";
  s += "    .setOption(\"height\", Math.max(280, brands.length * 30))" + "\n";
  s += "    .setOption(\"legend\", {position: \"none\"})" + "\n";
  s += "    .setOption(\"colors\", [\"#E94560\"])" + "\n";
  s += "    .setOption(\"hAxis\", {title: \"Bottles\"})" + "\n";
  s += "    .setOption(\"vAxis\", {title: \"Brand\"})" + "\n";
  s += "    .build();" + "\n";
  s += "  shCharts.insertChart(chart3);" + "\n";
  s += "" + "\n";
  s += "  shCharts.setFrozenRows(1);" + "\n";
  s += "" + "\n";
  s += "  // \u2500\u2500 Charts 4+: Oz Consumed per Brand per Bar (monthly trend) \u2500" + "\n";
  s += "  var shWt = ss.getSheetByName(SHEET_WEIGHT);" + "\n";
  s += "  if (shWt && shWt.getLastRow() > 2) {" + "\n";
  s += "    var wtData = shWt.getRange(2, 1, shWt.getLastRow()-1, 7).getValues();" + "\n";
  s += "" + "\n";
  s += "    // Build: { bar: { brand: { month: oz_consumed } } }" + "\n";
  s += "    // Group by month (MM/YYYY), sum oz consumed per bottle" + "\n";
  s += "    var barData = {};" + "\n";
  s += "    bars.forEach(function(bar) { barData[bar] = {}; });" + "\n";
  s += "" + "\n";
  s += "    wtData.forEach(function(row) {" + "\n";
  s += "      var rawDate = row[0]; if (!rawDate) return;" + "\n";
  s += "      var month   = Utilities.formatDate(new Date(rawDate), Session.getScriptTimeZone(), \"MM/yyyy\");" + "\n";
  s += "      var bar     = String(row[1]).trim();" + "\n";
  s += "      var brand   = String(row[2]).trim();" + "\n";
  s += "      var fullOz  = row[4] !== \"\" ? Number(row[4]) : 0;" + "\n";
  s += "      var remOz   = Number(row[5]);" + "\n";
  s += "      var consumed = fullOz > 0 ? Math.max(0, fullOz - remOz) : 0;" + "\n";
  s += "      if (!barData[bar]) return;" + "\n";
  s += "      if (!barData[bar][brand]) barData[bar][brand] = {};" + "\n";
  s += "      barData[bar][brand][month] = (barData[bar][brand][month] || 0) + consumed;" + "\n";
  s += "    });" + "\n";
  s += "" + "\n";
  s += "    // Get all unique months sorted" + "\n";
  s += "    var allMonths = [];" + "\n";
  s += "    wtData.forEach(function(row) {" + "\n";
  s += "      if (!row[0]) return;" + "\n";
  s += "      var m = Utilities.formatDate(new Date(row[0]), Session.getScriptTimeZone(), \"MM/yyyy\");" + "\n";
  s += "      if (allMonths.indexOf(m) === -1) allMonths.push(m);" + "\n";
  s += "    });" + "\n";
  s += "    allMonths.sort(function(a, b) {" + "\n";
  s += "      var pa = a.split(\"/\"), pb = b.split(\"/\");" + "\n";
  s += "      return (Number(pa[1])*100 + Number(pa[0])) - (Number(pb[1])*100 + Number(pb[0]));" + "\n";
  s += "    });" + "\n";
  s += "" + "\n";
  s += "    var wtStartRow = brand3Row + brands.length + 4;" + "\n";
  s += "" + "\n";
  s += "    bars.forEach(function(bar, bi) {" + "\n";
  s += "      var barBrands = Object.keys(barData[bar]).filter(function(b) {" + "\n";
  s += "        return Object.keys(barData[bar][b]).length > 0;" + "\n";
  s += "      });" + "\n";
  s += "      if (!barBrands.length) return;" + "\n";
  s += "" + "\n";
  s += "      var tRow = wtStartRow + bi * (barBrands.length + 5);" + "\n";
  s += "" + "\n";
  s += "      // Header: Month | Brand1 | Brand2 | ..." + "\n";
  s += "      shCharts.getRange(tRow, 1).setNumberFormat(\"@\").setValue(\"Month\");" + "\n";
  s += "      barBrands.forEach(function(brand, bri) {" + "\n";
  s += "        shCharts.getRange(tRow, bri + 2).setNumberFormat(\"@\").setValue(brand);" + "\n";
  s += "      });" + "\n";
  s += "      shCharts.getRange(tRow, 1, 1, barBrands.length + 1)" + "\n";
  s += "        .setBackground(\"#8E44AD\").setFontColor(\"#FFFFFF\").setFontWeight(\"bold\");" + "\n";
  s += "" + "\n";
  s += "      // Data: one row per month" + "\n";
  s += "      allMonths.forEach(function(month, mi) {" + "\n";
  s += "        shCharts.getRange(tRow + 1 + mi, 1).setNumberFormat(\"@\").setValue(month);" + "\n";
  s += "        barBrands.forEach(function(brand, bri) {" + "\n";
  s += "          var val = barData[bar][brand][month] || 0;" + "\n";
  s += "          shCharts.getRange(tRow + 1 + mi, bri + 2).setNumberFormat(\"0.0\").setValue(Math.round(val * 10) / 10);" + "\n";
  s += "        });" + "\n";
  s += "      });" + "\n";
  s += "" + "\n";
  s += "      // Line chart: x=month, one line per brand" + "\n";
  s += "      var dataRange = shCharts.getRange(tRow, 1, allMonths.length + 1, barBrands.length + 1);" + "\n";
  s += "      var wtChart = shCharts.newChart()" + "\n";
  s += "        .setChartType(Charts.ChartType.LINE)" + "\n";
  s += "        .addRange(dataRange)" + "\n";
  s += "        .setNumHeaders(1)" + "\n";
  s += "        .setPosition(tRow, barBrands.length + 3, 0, 0)" + "\n";
  s += "        .setOption(\"title\", \"Monthly Oz Consumed \u2014 \" + bar)" + "\n";
  s += "        .setOption(\"width\", 540)" + "\n";
  s += "        .setOption(\"height\", 320)" + "\n";
  s += "        .setOption(\"legend\", {position: \"right\"})" + "\n";
  s += "        .setOption(\"colors\", [\"#E94560\",\"#0F3460\",\"#27AE60\",\"#F39C12\",\"#8E44AD\",\"#1ABC9C\",\"#E67E22\"])" + "\n";
  s += "        .setOption(\"hAxis\", {title: \"Month\"})" + "\n";
  s += "        .setOption(\"vAxis\", {title: \"Oz Consumed\", minValue: 0})" + "\n";
  s += "        .setOption(\"curveType\", \"function\")" + "\n";
  s += "        .setOption(\"pointSize\", 5)" + "\n";
  s += "        .build();" + "\n";
  s += "      shCharts.insertChart(wtChart);" + "\n";
  s += "    });" + "\n";
  s += "  }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "//  HELPERS" + "\n";
  s += "// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550" + "\n";
  s += "function _getBarsFromDash(shDash) {" + "\n";
  s += "  var last = shDash.getLastRow(); if (last < 5) return [];" + "\n";
  s += "  return shDash.getRange(5, 1, Math.min(last-4, 10), 1).getValues()" + "\n";
  s += "    .map(function(r) { return r[0]; })" + "\n";
  s += "    .filter(function(v) { return v !== \"\"; });" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  s += "function _checkLowStock(hInv, brand, qty, type, dest, delivered, received) {" + "\n";
  s += "  var last = hInv.getLastRow(); if (last < 3) return;" + "\n";
  s += "  var data = hInv.getRange(3, 1, last-2, 8).getValues();" + "\n";
  s += "  var alerts = [];" + "\n";
  s += "  data.forEach(function(row, i) {" + "\n";
  s += "    var c = Number(row[5]), m = Number(row[6]);" + "\n";
  s += "    var r = hInv.getRange(i+3, 1, 1, 8);" + "\n";
  s += "    if (c <= 0)        r.setBackground(\"#FADBD8\");" + "\n";
  s += "    else if (c <= m) { r.setBackground(\"#FDEBD0\"); alerts.push({brand:row[0],current:c,min:m}); }" + "\n";
  s += "    else               r.setBackground(\"#EAFAF1\");" + "\n";
  s += "  });" + "\n";
  s += "  if (alerts.length > 0 && ALERT_EMAILS.length > 0) {" + "\n";
  s += "    var date = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), \"MM/dd/yyyy HH:mm\");" + "\n";
  s += "    var rows = alerts.map(function(a) {" + "\n";
  s += "      return \"<tr><td>\" + a.brand + \"</td><td>\" + a.current + \"</td><td>\" + a.min + \"</td></tr>\";" + "\n";
  s += "    }).join(\"\");" + "\n";
  s += "    MailApp.sendEmail({" + "\n";
  s += "      to: ALERT_EMAILS.join(\",\")," + "\n";
  s += "      subject: \"Port 27 - Low Stock Alert - \" + date," + "\n";
  s += "      htmlBody: \"<h2>Low Stock Alert</h2><table border=1 cellpadding=6>\" +" + "\n";
  s += "        \"<tr><th>Brand</th><th>Current</th><th>Min</th></tr>\" + rows + \"</table>\"" + "\n";
  s += "    });" + "\n";
  s += "  }" + "\n";
  s += "}" + "\n";
  s += "" + "\n";
  return s;
}
// ─── Create Spreadsheet ─────────────────────────────────────────────────────
function _createSpreadsheet(config) {
  var ss  = SpreadsheetApp.create(config.spreadsheetName || "Daily Liquor Transfer Log");
  var LOG = "Form_Responses";
  var INV = "Inventory";
  var EXIT  = "Exit \u2014 Transfer to Bar";
  var ENTRY = "Entry \u2014 Stock Received";

  PropertiesService.getUserProperties().setProperty("SS_ID", ss.getId());
  ss.getActiveSheet().setName("_temp");

  // Inventory sheet
  var shInv = ss.insertSheet(INV);
  shInv.setTabColor("#0F3460");
  _styleTitle(shInv, "INVENTORY - Stock Control", 8);
  _writeHeaders(shInv,
    ["Liquor / Brand","Category","Opening Stock","Total Entries","Total Exits","Current Stock","Minimum Stock","Alert"],
    [200,130,120,120,120,120,120,110], 2, "#0F3460");

  config.brands.forEach(function(b, i) {
    var row = i + 3;
    shInv.getRange(row,1).setValue(b.name).setFontWeight("bold");
    shInv.getRange(row,2).setValue(b.category);
    shInv.getRange(row,3).setValue(Number(b.opening)).setFontColor("#0000FF");
    shInv.getRange(row,4).setFormula(
      "=SUMPRODUCT((ISNUMBER(SEARCH(A"+row+",'"+LOG+"'!D:D)))*(ISNUMBER(SEARCH(\"Entry\",'"+LOG+"'!B:B)))*IFERROR(VALUE('"+LOG+"'!E:E),0))"
    );
    shInv.getRange(row,5).setFormula(
      "=SUMPRODUCT((ISNUMBER(SEARCH(A"+row+",'"+LOG+"'!D:D)))*(ISNUMBER(SEARCH(\"Exit\",'"+LOG+"'!B:B)))*IFERROR(VALUE('"+LOG+"'!E:E),0))"
    );
    shInv.getRange(row,6).setFormula("=C"+row+"+D"+row+"-E"+row);
    shInv.getRange(row,7).setValue(Number(b.min)).setFontColor("#0000FF");
    shInv.getRange(row,8).setFormula(
      "=IF(F"+row+"<=0,\"OUT OF STOCK\",IF(F"+row+"<=G"+row+",\"LOW\",\"OK\"))"
    );
    if (i%2===1) shInv.getRange(row,1,1,8).setBackground("#EBF5FB");
  });
  shInv.getRange(3,1,config.brands.length,8)
    .setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
  shInv.setFrozenRows(2);

  // Dashboard placeholder
  ss.insertSheet("Dashboard").setTabColor("#27AE60");

  return ss;
}

// ─── Create Form ────────────────────────────────────────────────────────────
function _createForm(ss, config) {
  var EXIT  = "Exit \u2014 Transfer to Bar";
  var ENTRY = "Entry \u2014 Stock Received";
  var brands = config.brands.map(function(b) { return b.name + " - " + b.category; });

  var form = FormApp.create(config.spreadsheetName || "Daily Liquor Transfer Log");
  form.setCollectEmail(false).setShowLinkToRespondAgain(true)
      .setConfirmationMessage("Movement recorded. Thank you!");

  form.addSectionHeaderItem().setTitle("Movement Type");
  form.addMultipleChoiceItem().setTitle("Movement Type").setChoiceValues([EXIT, ENTRY]).setRequired(true);
  form.addSectionHeaderItem().setTitle("Destination").setHelpText("Leave blank for stock entries.");
  form.addMultipleChoiceItem().setTitle("Destination Bar").setChoiceValues(config.barSections).setRequired(false);
  form.addSectionHeaderItem().setTitle("Product");
  form.addListItem().setTitle("Liquor / Brand").setChoiceValues(brands).setRequired(true);
  form.addTextItem().setTitle("Quantity (bottles)").setRequired(true);
  form.addSectionHeaderItem().setTitle("Responsible Parties");
  form.addMultipleChoiceItem().setTitle("Delivered By").setChoiceValues(config.staff).setRequired(true);
  form.addTextItem().setTitle("Received By").setRequired(true);
  form.addParagraphTextItem().setTitle("Additional Notes").setRequired(false);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  SpreadsheetApp.flush();
  Utilities.sleep(4000);

  // Rename auto-created sheet
  ss.getSheets().forEach(function(sh) {
    var n = sh.getName();
    if (n !== "Form_Responses" && n.toLowerCase().indexOf("form response") !== -1) {
      sh.setName("Form_Responses");
      sh.setTabColor("#E94560");
    }
  });
  SpreadsheetApp.flush();
  Utilities.sleep(500);
  try { var ph=ss.getSheetByName("_temp"); if(ph) ss.deleteSheet(ph); } catch(e){}

  // Style Form_Responses
  var shLog = ss.getSheetByName("Form_Responses");
  if (shLog) {
    shLog.setTabColor("#E94560");
    shLog.getRange(1,1,1,8).setBackground("#E94560").setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setWrap(true);
    shLog.setFrozenRows(1);
    [160,200,160,200,140,140,140,200].forEach(function(w,i){shLog.setColumnWidth(i+1,w);});
  }

  // Build Dashboard
  var shDash = ss.getSheetByName("Dashboard");
  if (shDash) {
    var b2 = config.brands.map(function(b){return {name:b.name,category:b.category};});
    _buildDashboardDirect(shDash, b2, config.barSections);
    // Create Charts sheet
    var shCharts = ss.getSheetByName("Charts");
    if (!shCharts) shCharts = ss.insertSheet("Charts");
    shCharts.setTabColor("#8E44AD");
  }

  PropertiesService.getUserProperties().setProperty("FORM_ID", form.getId());
  return form;
}

// ─── Build Dashboard (server-side, during setup) ────────────────────────────
function _buildDashboardDirect(shDash, brands, bars) {
  shDash.clear();

  // Title
  shDash.getRange(1,1,1,9).merge()
    .setValue("DASHBOARD - Summary by Bar Section & Brand")
    .setBackground("#1A1A2E").setFontColor("#FFFFFF")
    .setFontSize(13).setFontWeight("bold").setHorizontalAlignment("center");
  shDash.setRowHeight(1, 44);

  // Section 1: Consumption by Section
  shDash.getRange("A3:D3").merge().setValue("CONSUMPTION BY SECTION")
    .setBackground("#0F3460").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");
  ["Section","Total Bottles","Distinct Brands","Last Transfer"].forEach(function(h,i) {
    shDash.getRange(4,i+1).setValue(h)
      .setBackground("#E94560").setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center");
    shDash.setColumnWidth(i+1, [180,140,150,180][i]);
  });
  bars.forEach(function(bar, i) {
    var row = i+5;
    shDash.getRange(row,1).setValue(bar).setFontWeight("bold");
    shDash.getRange(row,2).setFormula(
      '=IFERROR(SUMIFS(Form_Responses!E:E,Form_Responses!C:C,A'+row+',Form_Responses!B:B,"Exit*"),0)'
    );
    shDash.getRange(row,3).setFormula(
      '=IFERROR(SUMPRODUCT(((Form_Responses!C$2:Form_Responses!C$1000)=A'+row+')' +
      '*(ISNUMBER(SEARCH("Exit",Form_Responses!B$2:Form_Responses!B$1000)))' +
      '*((Form_Responses!D$2:Form_Responses!D$1000)<>"")/COUNTIF(' +
      'Form_Responses!D$2:Form_Responses!D$1000,Form_Responses!D$2:Form_Responses!D$1000&"")),0)'
    );
    shDash.getRange(row,4).setFormula(
      '=IFERROR(TEXT(MAXIFS(Form_Responses!A:A,Form_Responses!C:C,A'+row+',Form_Responses!B:B,"Exit*"),"MM/DD/YYYY HH:MM"),"-")'
    );
    shDash.getRange(row,1,1,4).setHorizontalAlignment("center")
      .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    if (i%2===1) shDash.getRange(row,1,1,4).setBackground("#EBF5FB");
  });
  shDash.setColumnWidth(5, 30);

  // Section 2: Top Brands by Exits
  shDash.getRange("F3:I3").merge().setValue("TOP BRANDS BY EXITS")
    .setBackground("#0F3460").setFontColor("#FFFFFF")
    .setFontWeight("bold").setHorizontalAlignment("center");
  ["Brand","Total Exits","Total Entries","Current Stock"].forEach(function(h,i) {
    shDash.getRange(4,i+6).setValue(h)
      .setBackground("#E94560").setFontColor("#FFFFFF")
      .setFontWeight("bold").setHorizontalAlignment("center");
    shDash.setColumnWidth(i+6, [200,120,120,120][i]);
  });
  brands.forEach(function(b, i) {
    var dr = i+5, ir = i+3;
    shDash.getRange(dr,6).setFormula('=Inventory!A'+ir+'&" - "&Inventory!B'+ir).setFontWeight("bold");
    shDash.getRange(dr,7).setFormula('=Inventory!E'+ir);
    shDash.getRange(dr,8).setFormula('=Inventory!D'+ir);
    shDash.getRange(dr,9).setFormula('=Inventory!F'+ir);
    shDash.getRange(dr,6,1,4).setHorizontalAlignment("center")
      .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
    if (i%2===1) shDash.getRange(dr,6,1,4).setBackground("#EBF5FB");
  });

  // Section 3: Consumption by Category
  var s3 = brands.length + 7;
  shDash.getRange(s3,1,1,9).merge()
    .setValue("CONSUMPTION BY BAR SECTION & CATEGORY")
    .setBackground("#1A1A2E").setFontColor("#FFFFFF")
    .setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center");
  shDash.setRowHeight(s3, 36);
  var cHdr = s3+1;
  ["Bar Section","Category","Total Bottles","% of Section"].forEach(function(h,i) {
    shDash.getRange(cHdr,i+1).setValue(h)
      .setBackground("#E94560").setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center");
  });
  var catSet = {};
  brands.forEach(function(b) { if (b.category) catSet[b.category] = true; });
  var cats = Object.keys(catSet);
  var cRow = cHdr+1;
  bars.forEach(function(bar, bi) {
    cats.forEach(function(cat, ci) {
      var cb = brands.filter(function(b) { return b.category === cat; });
      if (!cb.length) return;
      shDash.getRange(cRow,1).setValue(bar).setHorizontalAlignment("center");
      shDash.getRange(cRow,2).setValue(cat).setHorizontalAlignment("center");
      shDash.getRange(cRow,3).setFormula(
        '=IFERROR(SUMPRODUCT(' +
        '(Form_Responses!C:C="' + bar + '")' +
        '*(ISNUMBER(SEARCH("Exit",Form_Responses!B:B)))' +
        '*(IFERROR(VLOOKUP(IFERROR(LEFT(Form_Responses!D:D,FIND(" - ",Form_Responses!D:D)-1),Form_Responses!D:D),Inventory!A:B,2,0),"")="' + cat + '")' +
        '*IFERROR(VALUE(Form_Responses!E:E),0)),0)'
      );
      shDash.getRange(cRow,4).setFormula(
        '=IFERROR(IF(SUMIFS(Form_Responses!E:E,Form_Responses!C:C,"' + bar +
        '",Form_Responses!B:B,"Exit*")=0,"-",TEXT(C' + cRow +
        '/SUMIFS(Form_Responses!E:E,Form_Responses!C:C,"' + bar +
        '",Form_Responses!B:B,"Exit*"),"0.0%")),"-")'
      );
      shDash.getRange(cRow,1,1,4)
        .setBorder(true,true,true,true,true,true,"#CCCCCC",SpreadsheetApp.BorderStyle.SOLID);
      if ((bi+ci)%2===1) shDash.getRange(cRow,1,1,4).setBackground("#EBF5FB");
      cRow++;
    });
  });
  shDash.setFrozenRows(2);
}


function _installTriggers(ss) {
  ScriptApp.getProjectTriggers().forEach(function(t){ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger("onFormSubmit").forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger("sendDailySummary").timeBased().atHour(23).everyDays(1).create();
}

function _syncBrands(ss, form, brands) {
  var list = brands.map(function(b){return b.name+" - "+b.category;});
  var q    = form.getItems().filter(function(i){return i.getTitle()==="Liquor / Brand";})[0];
  if(q&&q.getType()===FormApp.ItemType.LIST) q.asListItem().setChoiceValues(list);
}

function _styleTitle(sheet, title, cols) {
  sheet.getRange(1,1,1,cols).merge().setValue(title)
    .setBackground("#1A1A2E").setFontColor("#FFFFFF").setFontSize(13).setFontWeight("bold")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.setRowHeight(1,44); sheet.setRowHeight(2,30);
}

function _writeHeaders(sheet, headers, widths, row, bg, startCol) {
  startCol = startCol||1;
  headers.forEach(function(h,i){
    var col=startCol+i;
    sheet.getRange(row,col).setValue(h).setBackground(bg).setFontColor("#FFFFFF")
      .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
    sheet.setColumnWidth(col,widths[i]);
  });
}

function getSetupHTML() {
  return "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1.0\">\n<title>Bar Stock</title>\n<style>\n@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:'Inter',sans-serif;background:#0d0d1a;color:#e8e8f0;min-height:100vh;padding:32px 16px 60px}\n.c{max-width:640px;margin:0 auto}\n.hdr{text-align:center;margin-bottom:40px}\n.icon{font-size:48px;margin-bottom:12px}\nh1{font-size:28px;font-weight:700;color:#fff}\n.hdr p{color:#8888aa;margin-top:8px;font-size:15px;line-height:1.5}\n.card{background:#16162a;border:1px solid #2a2a45;border-radius:16px;padding:28px;margin-bottom:20px}\n.ct{font-size:13px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#e94560;margin-bottom:20px}\n.field{margin-bottom:18px}\n.field label{display:block;font-size:14px;font-weight:500;color:#c8c8e0;margin-bottom:6px}\n.field small{display:block;font-size:12px;color:#666688;margin-top:4px}\ninput,textarea{width:100%;padding:10px 14px;background:#0d0d1a;border:1px solid #2a2a45;border-radius:8px;color:#e8e8f0;font-size:14px;font-family:inherit;outline:none}\ninput:focus,textarea:focus{border-color:#e94560}\ntextarea{resize:vertical;min-height:80px}\ntable.bt{width:100%;border-collapse:collapse;margin-top:8px}\ntable.bt th{font-size:11px;font-weight:600;text-transform:uppercase;color:#8888aa;text-align:left;padding:6px 8px;border-bottom:1px solid #2a2a45}\ntable.bt td{padding:4px 3px}\ntable.bt input{padding:7px 8px;font-size:13px;border-radius:6px;width:100%}\n.rmbtn{background:none;border:none;color:#555577;cursor:pointer;font-size:18px;padding:4px 8px}\n.rmbtn:hover{color:#e94560}\n.addbtn{background:none;border:1px dashed #2a2a45;color:#8888aa;border-radius:8px;padding:9px 16px;font-size:13px;cursor:pointer;width:100%;margin-top:10px;font-family:inherit}\n.addbtn:hover{border-color:#e94560;color:#e94560}\n.btn-p{width:100%;padding:16px;background:#e94560;color:#fff;border:none;border-radius:12px;font-size:16px;font-weight:600;font-family:inherit;cursor:pointer;margin-top:8px}\n.btn-p:hover{background:#c73550}\n.btn-p:disabled{background:#333355;color:#555577;cursor:not-allowed}\n.btn-s{width:100%;padding:14px;background:none;color:#e94560;border:2px solid #e94560;border-radius:12px;font-size:15px;font-weight:600;font-family:inherit;cursor:pointer;margin-top:8px}\n.btn-s:hover{background:#1a0a10}\n.step{display:flex;gap:12px;align-items:flex-start;margin-bottom:12px}\n.snum{min-width:28px;height:28px;background:#e94560;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;color:#fff;flex-shrink:0}\n.stitle{font-size:13px;color:#e8e8f0;font-weight:500}\n.sdesc{font-size:12px;color:#8888aa;margin-top:2px}\n.spinner{width:32px;height:32px;border:3px solid #2a2a45;border-top-color:#e94560;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 12px}\n@keyframes spin{to{transform:rotate(360deg)}}\n.lbox{background:#0d0d1a;border:1px solid #2a2a45;border-radius:8px;padding:10px 14px;font-size:13px;color:#8888aa;word-break:break-all;margin:8px 0}\n.lbox a{color:#e94560;text-decoration:none}\n.cpybtn{width:100%;margin-top:16px;padding:12px;background:#e94560;color:#fff;border:none;border-radius:8px;font-size:14px;font-weight:600;font-family:inherit;cursor:pointer}\n.hidden{display:none}\n.sbox{border-radius:12px;padding:20px 24px;margin-top:20px}\n.divider{display:flex;align-items:center;gap:12px;margin:20px 0;color:#444466;font-size:13px}\n.divider::before,.divider::after{content:\"\";flex:1;height:1px;background:#2a2a45}\n.copy-link-btn{padding:8px 14px;background:#2a2a45;color:#e8e8f0;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap;font-family:inherit}\n.copy-link-btn:hover{background:#e94560}\n</style>\n</head>\n<body>\n<div class=\"c\">\n  <div class=\"hdr\">\n    <div class=\"icon\">&#127870;</div>\n    <h1>Bar Stock</h1>\n    <p>Set up a new bar log or update your existing one.</p>\n  </div>\n\n  <div class=\"card\" id=\"modeCard\">\n    <button class=\"btn-p\" id=\"btnNew\">&#10133; Create a new Bar Log</button>\n    <div class=\"divider\">or</div>\n    <button class=\"btn-s\" id=\"btnUpdate\">&#9881;&#65039; Manage my Bar Log</button>\n  </div>\n\n  <!-- New setup form -->\n  <div id=\"setupForm\" class=\"hidden\">\n    <div class=\"card\"><div class=\"ct\">General</div>\n      <div class=\"field\"><label>Spreadsheet &amp; Form name</label><input id=\"ssName\" value=\"Daily Liquor Transfer Log\"></div>\n      <div class=\"field\"><label>Alert emails</label><input id=\"emails\" placeholder=\"manager@bar.com, owner@bar.com\"><small>Separate multiple emails with commas</small></div>\n    </div>\n    <div class=\"card\"><div class=\"ct\">Bar Sections</div>\n      <div class=\"field\"><label>Sections (one per line)</label><textarea id=\"barSections\">Tap Bar\nObservatory Bar\nTiki Bar</textarea></div>\n    </div>\n    <div class=\"card\"><div class=\"ct\">Staff</div>\n      <div class=\"field\"><label>Team members who deliver bottles (one per line)</label><textarea id=\"staff\">Junior\nJules\nMaria A</textarea></div>\n    </div>\n    <div class=\"card\"><div class=\"ct\">Initial Inventory</div>\n      <table class=\"bt\"><thead><tr><th>Brand name</th><th>Category</th><th>Opening stock</th><th>Min. stock</th><th></th></tr></thead>\n      <tbody id=\"brandRows\"></tbody></table>\n      <button class=\"addbtn\" id=\"addBtn\">+ Add brand</button>\n    </div>\n    <button class=\"btn-p\" id=\"subBtn\">Create My Bar Log</button>\n  </div>\n\n  <!-- Update form -->\n  <div id=\"updateForm\" class=\"hidden\">\n    <div class=\"card\"><div class=\"ct\">Update your existing Bar Log</div>\n      <p style=\"font-size:14px;color:#c8c8e0;margin-bottom:8px\">Your data stays intact. Only the script gets updated.</p>\n      <div class=\"field\" style=\"margin-top:16px\">\n        <label>Alert emails</label>\n        <div style=\"display:flex;gap:8px;align-items:center\">\n          <input id=\"updateEmails\" placeholder=\"manager@bar.com, owner@bar.com\" style=\"flex:1\">\n          <button class=\"copy-link-btn\" id=\"saveEmailsBtn\" onclick=\"saveEmails()\">Save</button>\n        </div>\n        <span id=\"emailSaved\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Emails saved!</span>\n      </div>\n      <div style=\"margin-top:20px\">\n        <div class=\"step\"><div class=\"snum\">1</div><div><div class=\"stitle\">Open your Spreadsheet</div></div></div>\n        <div class=\"step\"><div class=\"snum\">2</div><div><div class=\"stitle\">Extensions &rarr; Apps Script</div></div></div>\n        <div class=\"step\"><div class=\"snum\">3</div><div><div class=\"stitle\">Delete everything and paste the new script</div><div class=\"sdesc\">Use the Copy Script button below</div></div></div>\n        <div class=\"step\"><div class=\"snum\">4</div><div><div class=\"stitle\">Save, select installTriggers, click &#9654; Run</div></div></div>\n        <div class=\"step\"><div class=\"snum\">5</div><div><div class=\"stitle\">Close and reload Spreadsheet</div><div class=\"sdesc\">Then go to Liquor Log &rarr; Refresh Dashboard</div></div></div>\n        <button class=\"cpybtn\" id=\"cpyBtn2\">&#128203; Copy New Script to Clipboard</button>\n        <p class=\"hidden\" id=\"cpyOk2\" style=\"text-align:center;font-size:12px;color:#27ae60;margin-top:8px\">&#10003; Copied!</p>\n      </div>\n\n      <div style=\"margin-top:20px;padding-top:20px;border-top:1px solid #2a2a45\">\n        <div style=\"font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#e94560;margin-bottom:14px\">&#128279; Your Links</div>\n        <div style=\"margin-bottom:12px\">\n          <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#128202; Spreadsheet</div>\n          <div style=\"display:flex;gap:8px;align-items:center\">\n            <input id=\"inSS\" placeholder=\"Paste your Spreadsheet URL\" style=\"flex:1;font-size:13px;padding:8px 12px\">\n            <button class=\"copy-link-btn\" onclick=\"copyVal('inSS','cpUS')\">Copy</button>\n          </div>\n          <span id=\"cpUS\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n        </div>\n        <div style=\"margin-bottom:12px\">\n          <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#128203; Form (share with staff)</div>\n          <div style=\"display:flex;gap:8px;align-items:center\">\n            <input id=\"inForm\" placeholder=\"Paste your Form URL\" style=\"flex:1;font-size:13px;padding:8px 12px\">\n            <button class=\"copy-link-btn\" onclick=\"copyVal('inForm','cpUF')\">Copy</button>\n          </div>\n          <span id=\"cpUF\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n        </div>\n        <div style=\"margin-bottom:12px\">\n          <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#127870; Quick Order</div>\n          <div style=\"display:flex;gap:8px;align-items:center\">\n            <div class=\"lbox\" id=\"uLinkOrder\" style=\"flex:1;margin:0;font-size:13px\"></div>\n            <button class=\"copy-link-btn\" onclick=\"copyLink('uLinkOrder','cpUO')\">Copy</button>\n          </div>\n          <span id=\"cpUO\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n        </div>\n        <div style=\"margin-bottom:12px\">\n          <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#128230; Restock</div>\n          <div style=\"display:flex;gap:8px;align-items:center\">\n            <div class=\"lbox\" id=\"uLinkRestock\" style=\"flex:1;margin:0;font-size:13px\"></div>\n            <button class=\"copy-link-btn\" onclick=\"copyLink('uLinkRestock','cpUR')\">Copy</button>\n          </div>\n          <span id=\"cpUR\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n        </div>\n        <div>\n          <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#9878; Bottle Weight Log</div>\n          <div style=\"display:flex;gap:8px;align-items:center\">\n            <div class=\"lbox\" id=\"uLinkWeight\" style=\"flex:1;margin:0;font-size:13px\"></div>\n            <button class=\"copy-link-btn\" onclick=\"copyLink('uLinkWeight','cpUW')\">Copy</button>\n          </div>\n          <span id=\"cpUW\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n        </div>\n      </div>\n    </div>\n  </div>\n\n  <!-- Status boxes -->\n  <div class=\"sbox hidden\" id=\"stLoad\" style=\"background:#16162a;border:1px solid #2a2a45;text-align:center;color:#8888aa\">\n    <div class=\"spinner\"></div>Setting everything up &mdash; about 45 seconds&hellip;\n  </div>\n  <div class=\"sbox hidden\" id=\"stOk\" style=\"background:#0d2a1a;border:1px solid #27ae60\">\n    <h3 style=\"color:#27ae60;margin-bottom:12px\">&#10003; All done!</h3>\n    <p style=\"color:#c8c8e0;margin-bottom:12px\">Created in your Google Drive.</p>\n    <p style=\"font-size:13px;color:#8888aa;margin-bottom:6px\">Your Spreadsheet:</p>\n    <div class=\"lbox\" id=\"ssLink\"></div>\n    <p style=\"font-size:13px;color:#8888aa;margin:12px 0 6px\">Share this Form with your team:</p>\n    <div class=\"lbox\" id=\"fmLink\"></div>\n    <div style=\"background:#0d0d1a;border:1px solid #2a2a45;border-radius:12px;padding:20px;margin-top:20px\">\n      <p style=\"font-size:14px;font-weight:600;color:#e94560;margin-bottom:16px\">&#9889; One-time activation</p>\n      <div class=\"step\"><div class=\"snum\">1</div><div><div class=\"stitle\">Open your Spreadsheet</div></div></div>\n      <div class=\"step\"><div class=\"snum\">2</div><div><div class=\"stitle\">Extensions &rarr; Apps Script</div></div></div>\n      <div class=\"step\"><div class=\"snum\">3</div><div><div class=\"stitle\">Delete everything and paste</div></div></div>\n      <div class=\"step\"><div class=\"snum\">4</div><div><div class=\"stitle\">Save, select installTriggers, click &#9654; Run</div></div></div>\n      <div class=\"step\"><div class=\"snum\">5</div><div><div class=\"stitle\">Close and reload Spreadsheet</div></div></div>\n      <button class=\"cpybtn\" id=\"cpyBtn\">&#128203; Copy Script to Clipboard</button>\n      <p class=\"hidden\" id=\"cpyOk\" style=\"text-align:center;font-size:12px;color:#27ae60;margin-top:8px\">&#10003; Copied!</p>\n    </div>\n  </div>\n  <div class=\"sbox hidden\" id=\"stErr\" style=\"background:#2a0d14;border:1px solid #e94560\">\n    <h3 style=\"color:#e94560;margin-bottom:8px\">Something went wrong</h3>\n    <p id=\"errMsg\" style=\"color:#c8c8e0;font-size:14px\"></p>\n  </div>\n\n  <!-- Links section - shown after setup or always if SS exists -->\n  <div id=\"linksSection\" class=\"hidden\">\n    <div class=\"card\">\n      <div class=\"ct\">&#128279; Your Links</div>\n      <p style=\"font-size:13px;color:#8888aa;margin-bottom:16px\">Share these links with your team.</p>\n\n      <div style=\"margin-bottom:14px\">\n        <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#128202; Spreadsheet</div>\n        <div style=\"display:flex;gap:8px;align-items:center\">\n          <div class=\"lbox\" id=\"linkSS\" style=\"flex:1;margin:0\"></div>\n          <button class=\"copy-link-btn\" onclick=\"copyLink('linkSS','cpSS')\">Copy</button>\n        </div>\n        <span id=\"cpSS\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n      </div>\n\n      <div style=\"margin-bottom:14px\">\n        <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#128203; Form (share with staff)</div>\n        <div style=\"display:flex;gap:8px;align-items:center\">\n          <div class=\"lbox\" id=\"linkForm\" style=\"flex:1;margin:0\"></div>\n          <button class=\"copy-link-btn\" onclick=\"copyLink('linkForm','cpForm')\">Copy</button>\n        </div>\n        <span id=\"cpForm\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n      </div>\n\n      <div style=\"margin-bottom:14px\">\n        <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#127870; Quick Order (for managers)</div>\n        <div style=\"display:flex;gap:8px;align-items:center\">\n          <div class=\"lbox\" id=\"linkOrder\" style=\"flex:1;margin:0\"></div>\n          <button class=\"copy-link-btn\" onclick=\"copyLink('linkOrder','cpOrder')\">Copy</button>\n        </div>\n        <span id=\"cpOrder\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n      </div>\n\n      <div style=\"margin-bottom:14px\">\n        <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#128230; Restock (for managers)</div>\n        <div style=\"display:flex;gap:8px;align-items:center\">\n          <div class=\"lbox\" id=\"linkRestock\" style=\"flex:1;margin:0\"></div>\n          <button class=\"copy-link-btn\" onclick=\"copyLink('linkRestock','cpRestock')\">Copy</button>\n        </div>\n        <span id=\"cpRestock\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n      </div>\n      <div style=\"margin-bottom:4px\">\n        <div style=\"font-size:12px;color:#8888aa;margin-bottom:6px\">&#9878; Bottle Weight Log (for managers)</div>\n        <div style=\"display:flex;gap:8px;align-items:center\">\n          <div class=\"lbox\" id=\"linkWeight\" style=\"flex:1;margin:0\"></div>\n          <button class=\"copy-link-btn\" onclick=\"copyLink('linkWeight','cpWeight')\">Copy</button>\n        </div>\n        <span id=\"cpWeight\" style=\"display:none;font-size:11px;color:#27ae60;margin-top:4px\">&#10003; Copied!</span>\n      </div>\n    </div>\n  </div>\n</div>\n\n<script>\nvar BRANDS = [\n  [\"Jack Daniels\",\"Whiskey\",12,3],[\"Havana Club 7\",\"Rum\",6,3],\n  [\"Grey Goose\",\"Vodka\",8,2],[\"Bombay Sapphire\",\"Gin\",6,2],\n  [\"Johnnie Walker Black\",\"Whiskey\",10,3],[\"Bacardi Blanco\",\"Rum\",8,3],\n  [\"Patron Silver\",\"Tequila\",6,2],[\"Aperol\",\"Liqueur\",4,2],\n  [\"Campari\",\"Liqueur\",4,2],[\"Baileys\",\"Cream\",4,2]\n];\nBRANDS.forEach(function(b) { addRow(b[0],b[1],b[2],b[3]); });\n\ndocument.getElementById(\"btnNew\").onclick = function() {\n  document.getElementById(\"setupForm\").classList.remove(\"hidden\");\n  document.getElementById(\"modeCard\").classList.add(\"hidden\");\n};\ndocument.getElementById(\"btnUpdate\").onclick = function() {\n  document.getElementById(\"updateForm\").classList.remove(\"hidden\");\n  document.getElementById(\"modeCard\").classList.add(\"hidden\");\n  // Get real app URL from server\n  google.script.run\n    .withSuccessHandler(function(url) {\n      document.getElementById(\"uLinkOrder\").innerHTML = \"<a href='\" + url + \"?page=order' target='_blank'>\" + url + \"?page=order</a>\";\n      document.getElementById(\"uLinkRestock\").innerHTML = \"<a href='\" + url + \"?page=restock' target='_blank'>\" + url + \"?page=restock</a>\";\n      document.getElementById(\"uLinkWeight\").innerHTML = \"<a href='\" + url + \"?page=weight' target='_blank'>\" + url + \"?page=weight</a>\";\n    })\n    .getAppUrl();\n  // Load saved SS and Form URLs\n  google.script.run\n    .withSuccessHandler(function(r) {\n      if (r.ssUrl)   document.getElementById(\"inSS\").value   = r.ssUrl;\n      if (r.formUrl) document.getElementById(\"inForm\").value = r.formUrl;\n      if (r.emails)  document.getElementById(\"updateEmails\").value = r.emails;\n    })\n    .getSavedLinks();\n};\ndocument.getElementById(\"addBtn\").onclick = function() { addRow(\"\",\"\",\"\",2); };\ndocument.getElementById(\"subBtn\").onclick = submit;\ndocument.getElementById(\"cpyBtn\").onclick = function() { copyScript(\"cpyBtn\",\"cpyOk\",\"\"); };\ndocument.getElementById(\"cpyBtn2\").onclick = function() {\n  copyScript(\"cpyBtn2\",\"cpyOk2\", document.getElementById(\"updateEmails\").value);\n};\n\nfunction addRow(n,c,o,m) {\n  var tr = document.createElement(\"tr\");\n  tr.appendChild(mkCell(n, \"Brand name\", \"\"));\n  tr.appendChild(mkCell(c, \"Category\", \"\"));\n  tr.appendChild(mkCell(o, \"12\", \"70px\"));\n  tr.appendChild(mkCell(m, \"3\", \"70px\"));\n  var btn = document.createElement(\"button\");\n  btn.type = \"button\"; btn.className = \"rmbtn\"; btn.textContent = \"x\";\n  btn.onclick = function() { tr.parentNode.removeChild(tr); };\n  var td = document.createElement(\"td\"); td.appendChild(btn); tr.appendChild(td);\n  document.getElementById(\"brandRows\").appendChild(tr);\n}\nfunction mkCell(v,ph,w) {\n  var td = document.createElement(\"td\");\n  var inp = document.createElement(\"input\");\n  inp.type = \"text\"; inp.value = String(v !== undefined ? v : \"\"); inp.placeholder = ph;\n  if (w) inp.style.width = w;\n  td.appendChild(inp); return td;\n}\nfunction submit() {\n  var brands = [];\n  document.querySelectorAll(\"#brandRows tr\").forEach(function(tr) {\n    var ins = tr.querySelectorAll(\"input\");\n    var nm = ins[0].value.trim(); if (!nm) return;\n    brands.push({name:nm, category:ins[1].value.trim()||\"Other\", opening:Number(ins[2].value)||0, min:Number(ins[3].value)||2});\n  });\n  if (!brands.length) { alert(\"Add at least one brand.\"); return; }\n  var cfg = {\n    spreadsheetName: document.getElementById(\"ssName\").value.trim() || \"Daily Liquor Transfer Log\",\n    alertEmails: document.getElementById(\"emails\").value.trim(),\n    barSections: document.getElementById(\"barSections\").value.trim().split(\"\\n\").map(function(s){return s.trim();}).filter(Boolean),\n    staff: document.getElementById(\"staff\").value.trim().split(\"\\n\").map(function(s){return s.trim();}).filter(Boolean),\n    brands: brands\n  };\n  document.getElementById(\"subBtn\").disabled = true;\n  show(\"stLoad\"); hide(\"stOk\"); hide(\"stErr\");\n  google.script.run\n    .withSuccessHandler(function(r) {\n      hide(\"stLoad\");\n      if (r.success) {\n        document.getElementById(\"ssLink\").innerHTML = \"<a href='\" + r.spreadsheetUrl + \"' target='_blank'>\" + r.spreadsheetUrl + \"</a>\";\n        document.getElementById(\"fmLink\").innerHTML = \"<a href='\" + r.formUrl + \"' target='_blank'>\" + r.formUrl + \"</a>\";\n        show(\"stOk\");\n        document.getElementById(\"linkSS\").innerHTML = \"<a href='\" + r.spreadsheetUrl + \"' target='_blank'>\" + r.spreadsheetUrl + \"</a>\";\n        document.getElementById(\"linkForm\").innerHTML = \"<a href='\" + r.formUrl + \"' target='_blank'>\" + r.formUrl + \"</a>\";\n        document.getElementById(\"linkOrder\").innerHTML = \"<a href='\" + r.appUrl + \"?page=order' target='_blank'>\" + r.appUrl + \"?page=order</a>\";\n        document.getElementById(\"linkRestock\").innerHTML = \"<a href='\" + r.appUrl + \"?page=restock' target='_blank'>\" + r.appUrl + \"?page=restock</a>\";\n        document.getElementById(\"linkWeight\").innerHTML = \"<a href='\" + r.appUrl + \"?page=weight' target='_blank'>\" + r.appUrl + \"?page=weight</a>\";\n        show(\"linksSection\");\n      } else { document.getElementById(\"errMsg\").textContent = r.error; show(\"stErr\"); document.getElementById(\"subBtn\").disabled = false; }\n    })\n    .withFailureHandler(function(err) {\n      hide(\"stLoad\"); document.getElementById(\"errMsg\").textContent = err.message; show(\"stErr\"); document.getElementById(\"subBtn\").disabled = false;\n    })\n    .runSetup(cfg);\n}\nfunction copyScript(btnId, okId, emails) {\n  google.script.run.withSuccessHandler(function(t) {\n    if (emails) {\n      var arr = JSON.stringify(emails.split(\",\").map(function(e){return e.trim();}).filter(Boolean));\n      t = t.replace(\"__ALERT_EMAILS__\", arr);\n    }\n    navigator.clipboard.writeText(t).then(function() {\n      document.getElementById(okId).classList.remove(\"hidden\");\n      document.getElementById(btnId).textContent = \"Copied!\";\n    });\n  }).getManagementScriptForCopy();\n}\nfunction show(id) { document.getElementById(id).classList.remove(\"hidden\"); }\nfunction hide(id) { document.getElementById(id).classList.add(\"hidden\"); }\nfunction copyLink(srcId,okId) {\n  var el=document.getElementById(srcId);\n  var a=el.querySelector(\"a\");\n  var link=a?a.href:el.textContent.trim();\n  navigator.clipboard.writeText(link).then(function() {\n    var ok=document.getElementById(okId); ok.style.display=\"block\";\n    setTimeout(function(){ok.style.display=\"none\";},2000);\n  });\n}\nfunction copyVal(inputId,okId) {\n  var val=document.getElementById(inputId).value.trim();\n  if(!val) return;\n  navigator.clipboard.writeText(val).then(function() {\n    var ok=document.getElementById(okId); ok.style.display=\"block\";\n    setTimeout(function(){ok.style.display=\"none\";},2000);\n  });\n}\nfunction saveEmails() {\n  var emails = document.getElementById(\"updateEmails\").value.trim();\n  var btn = document.getElementById(\"saveEmailsBtn\");\n  btn.disabled = true; btn.textContent = \"Saving...\";\n  google.script.run\n    .withSuccessHandler(function(r) {\n      btn.disabled = false; btn.textContent = \"Save\";\n      if (r.success) {\n        document.getElementById(\"emailSaved\").style.display = \"block\";\n        setTimeout(function() { document.getElementById(\"emailSaved\").style.display = \"none\"; }, 2000);\n      } else { alert(\"Error: \" + r.error); }\n    })\n    .withFailureHandler(function(e) { btn.disabled = false; btn.textContent = \"Save\"; alert(\"Failed: \" + e.message); })\n    .saveAlertEmails(emails);\n}\n</script>\n</body>\n</html>";
}

function getOrderHTML() {
  return "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1.0\">\n<title>Quick Order</title>\n<style>\n@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:Inter,sans-serif;background:#0d0d1a;color:#e8e8f0;min-height:100vh;padding:0 0 90px}\n.header{background:#1A1A2E;padding:20px 24px;display:flex;align-items:center;gap:12px;border-bottom:1px solid #2a2a45;position:sticky;top:0;z-index:10}\n.header .icon{font-size:28px}\n.header h1{font-size:20px;font-weight:700;color:#fff}\n.header p{font-size:13px;color:#8888aa;margin-top:2px}\n.content{max-width:600px;margin:0 auto;padding:20px 16px}\n.section{background:#16162a;border:1px solid #2a2a45;border-radius:12px;padding:20px;margin-bottom:16px}\n.section-title{font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#e94560;margin-bottom:14px}\nselect,input[type=text]{width:100%;padding:10px 14px;background:#0d0d1a;border:1px solid #2a2a45;border-radius:8px;color:#e8e8f0;font-size:14px;font-family:inherit;outline:none}\nselect:focus,input:focus{border-color:#e94560}\n.field{margin-bottom:12px}\n.field label{display:block;font-size:13px;color:#8888aa;margin-bottom:6px}\n.brand-row{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid #1a1a2e}\n.brand-row:last-child{border-bottom:none}\n.brand-info{flex:1}\n.brand-name{font-size:14px;font-weight:500}\n.brand-cat{font-size:12px;color:#8888aa;margin-top:2px}\n.brand-stock{font-size:12px;font-weight:600;margin-top:2px}\n.ok{color:#27AE60}.low{color:#F39C12}.out{color:#E74C3C}\n.qty-wrap{display:flex;align-items:center;gap:6px}\n.qbtn{width:32px;height:32px;background:#2a2a45;border:none;border-radius:6px;color:#e8e8f0;font-size:20px;cursor:pointer;line-height:1}\n.qbtn:hover:not([disabled]){background:#e94560}\n.qbtn[disabled]{opacity:.3;cursor:not-allowed}\n.qinp{width:48px;text-align:center;padding:6px 4px;font-size:15px;font-weight:600}\n.submit-bar{position:fixed;bottom:0;left:0;right:0;background:#1A1A2E;border-top:1px solid #2a2a45;padding:16px 24px}\n.sbtn{width:100%;max-width:600px;margin:0 auto;display:block;padding:16px;background:#e94560;color:#fff;border:none;border-radius:12px;font-size:16px;font-weight:700;font-family:inherit;cursor:pointer}\n.sbtn:hover:not([disabled]){background:#c73550}\n.sbtn[disabled]{background:#333355;color:#555577;cursor:not-allowed}\n.badge{display:inline-block;background:#e94560;color:#fff;border-radius:20px;padding:1px 8px;font-size:12px;font-weight:700;margin-left:8px}\n.hidden{display:none}\n.success{text-align:center;padding:60px 24px}\n.success .icon{font-size:64px;margin-bottom:16px}\n.success h2{color:#27AE60;font-size:24px;margin-bottom:8px}\n.success p{color:#8888aa;font-size:15px;margin-bottom:24px}\n.success button{padding:14px 32px;background:#e94560;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;font-family:inherit}\n.loading{text-align:center;padding:60px 24px;color:#8888aa}\n.spinner{width:40px;height:40px;border:3px solid #2a2a45;border-top-color:#e94560;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 16px}\n@keyframes spin{to{transform:rotate(360deg)}}\n.cat-hdr{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#555577;padding:12px 0 6px}\n</style>\n</head>\n<body>\n<div class=\"header\">\n  <div class=\"icon\">&#127870;</div>\n  <div><h1>Quick Order</h1><p>Select items and submit</p></div>\n</div>\n<div class=\"content\">\n  <div id=\"loading\" class=\"loading\"><div class=\"spinner\"></div>Loading inventory...</div>\n  <div id=\"orderView\" class=\"hidden\">\n    <div class=\"section\">\n      <div class=\"section-title\">Destination &amp; Staff</div>\n      <div class=\"field\"><label>Bar Section</label><select id=\"barSel\"></select></div>\n      <div class=\"field\"><label>Delivered By</label><input type=\"text\" id=\"delivBy\" placeholder=\"Your name\"></div>\n      <div class=\"field\"><label>Received By</label><input type=\"text\" id=\"recvBy\" placeholder=\"Who receives at the bar\"></div>\n    </div>\n    <div class=\"section\">\n      <div class=\"section-title\">Brands <span class=\"badge\" id=\"count\">0</span></div>\n      <div id=\"brandList\"></div>\n    </div>\n  </div>\n  <div id=\"successView\" class=\"hidden success\">\n    <div class=\"icon\">&#9989;</div>\n    <h2>Order submitted!</h2>\n    <p id=\"successMsg\"></p>\n    <button onclick=\"reset()\">New Order</button>\n  </div>\n</div>\n<div id=\"submitBar\" class=\"submit-bar\">\n  <button class=\"sbtn\" id=\"submitBtn\" onclick=\"doSubmit()\" disabled>Submit Order</button>\n</div>\n<script>\nvar INV = null;\ngoogle.script.run\n  .withSuccessHandler(function(r) {\n    document.getElementById(\"loading\").classList.add(\"hidden\");\n    if (!r.success) { document.getElementById(\"loading\").textContent = \"Error: \" + r.error; document.getElementById(\"loading\").classList.remove(\"hidden\"); return; }\n    INV = r;\n    buildUI(r);\n    document.getElementById(\"orderView\").classList.remove(\"hidden\");\n  })\n  .withFailureHandler(function(e) {\n    document.getElementById(\"loading\").textContent = \"Error: \" + e.message;\n    document.getElementById(\"loading\").classList.remove(\"hidden\");\n  })\n  .loadInventory();\n\nfunction buildUI(r) {\n  var sel = document.getElementById(\"barSel\");\n  r.bars.forEach(function(b) { var o = document.createElement(\"option\"); o.value = b; o.textContent = b; sel.appendChild(o); });\n  var byCat = {};\n  r.brands.forEach(function(b) { if (!byCat[b.category]) byCat[b.category] = []; byCat[b.category].push(b); });\n  var list = document.getElementById(\"brandList\");\n  Object.keys(byCat).forEach(function(cat) {\n    var hdr = document.createElement(\"div\"); hdr.className = \"cat-hdr\"; hdr.textContent = cat; list.appendChild(hdr);\n    byCat[cat].forEach(function(b) {\n      var row = document.createElement(\"div\"); row.className = \"brand-row\";\n      var info = document.createElement(\"div\"); info.className = \"brand-info\";\n      var nm = document.createElement(\"div\"); nm.className = \"brand-name\"; nm.textContent = b.name;\n      var ct = document.createElement(\"div\"); ct.className = \"brand-cat\"; ct.textContent = b.category;\n      var sc = b.stock <= 0 ? \"out\" : b.stock <= b.min ? \"low\" : \"ok\";\n      var st = document.createElement(\"div\"); st.className = \"brand-stock \" + sc;\n      st.textContent = b.stock <= 0 ? \"Out of stock\" : b.stock + \" available\";\n      info.appendChild(nm); info.appendChild(ct); info.appendChild(st);\n      var wrap = document.createElement(\"div\"); wrap.className = \"qty-wrap\";\n      var bm = document.createElement(\"button\"); bm.className = \"qbtn\"; bm.textContent = \"-\"; bm.disabled = true;\n      bm.onclick = (function(n) { return function() { chg(n,-1); }; })(b.name);\n      var inp = document.createElement(\"input\"); inp.type = \"text\"; inp.className = \"qinp\";\n      inp.value = \"0\"; inp.readOnly = true; inp.id = \"q_\" + b.name.replace(/ /g,\"_\");\n      var bp = document.createElement(\"button\"); bp.className = \"qbtn\"; bp.textContent = \"+\"; bp.disabled = b.stock <= 0;\n      bp.id = \"p_\" + b.name.replace(/ /g,\"_\");\n      bp.onclick = (function(n) { return function() { chg(n,1); }; })(b.name);\n      wrap.appendChild(bm); wrap.appendChild(inp); wrap.appendChild(bp);\n      row.appendChild(info); row.appendChild(wrap); list.appendChild(row);\n    });\n  });\n}\n\nfunction chg(name, delta) {\n  var b = INV.brands.filter(function(x){return x.name===name;})[0];\n  var inp = document.getElementById(\"q_\" + name.replace(/ /g,\"_\"));\n  var val = Math.max(0, Math.min(b.stock, parseInt(inp.value||0)+delta));\n  inp.value = val;\n  inp.previousSibling.disabled = (val <= 0);\n  upd();\n}\nfunction upd() {\n  var n = 0;\n  INV.brands.forEach(function(b) { var i = document.getElementById(\"q_\"+b.name.replace(/ /g,\"_\")); if(i&&parseInt(i.value)>0) n++; });\n  document.getElementById(\"count\").textContent = n;\n  document.getElementById(\"submitBtn\").disabled = (n === 0);\n}\nfunction doSubmit() {\n  var bar = document.getElementById(\"barSel\").value;\n  var del = document.getElementById(\"delivBy\").value.trim();\n  var rec = document.getElementById(\"recvBy\").value.trim();\n  if (!del||!rec) { alert(\"Fill in Delivered By and Received By.\"); return; }\n  var items = [];\n  INV.brands.forEach(function(b) { var i=document.getElementById(\"q_\"+b.name.replace(/ /g,\"_\")); var q=parseInt(i?i.value:0); if(q>0) items.push({name:b.name,category:b.category,qty:q}); });\n  if (!items.length) return;\n  var btn = document.getElementById(\"submitBtn\"); btn.disabled=true; btn.textContent=\"Submitting...\";\n  google.script.run\n    .withSuccessHandler(function(r) {\n      if (r.success) {\n        document.getElementById(\"orderView\").classList.add(\"hidden\");\n        document.getElementById(\"submitBar\").classList.add(\"hidden\");\n        document.getElementById(\"successMsg\").textContent = r.count + \" item(s) logged to \" + bar + \".\";\n        document.getElementById(\"successView\").classList.remove(\"hidden\");\n      } else { alert(\"Error: \"+r.error); btn.disabled=false; btn.textContent=\"Submit Order\"; }\n    })\n    .withFailureHandler(function(e) { alert(\"Failed: \"+e.message); btn.disabled=false; btn.textContent=\"Submit Order\"; })\n    .submitOrder({bar:bar, deliveredBy:del, receivedBy:rec, items:items});\n}\nfunction reset() {\n  document.getElementById(\"successView\").classList.add(\"hidden\");\n  document.getElementById(\"submitBar\").classList.remove(\"hidden\");\n  document.getElementById(\"loading\").textContent=\"Refreshing inventory...\";\n  document.getElementById(\"loading\").classList.remove(\"hidden\");\n  document.getElementById(\"orderView\").classList.add(\"hidden\");\n  document.getElementById(\"brandList\").innerHTML=\"\";\n  var sel=document.getElementById(\"barSel\"); while(sel.options.length>0) sel.remove(0);\n  google.script.run\n    .withSuccessHandler(function(r){\n      document.getElementById(\"loading\").classList.add(\"hidden\");\n      if(!r.success){document.getElementById(\"loading\").textContent=\"Error: \"+r.error;document.getElementById(\"loading\").classList.remove(\"hidden\");return;}\n      INV=r; buildUI(r);\n      document.getElementById(\"count\").textContent=\"0\";\n      document.getElementById(\"submitBtn\").disabled=true;\n      document.getElementById(\"submitBtn\").textContent=\"Submit Order\";\n      document.getElementById(\"delivBy\").value=\"\";\n      document.getElementById(\"recvBy\").value=\"\";\n      document.getElementById(\"orderView\").classList.remove(\"hidden\");\n    })\n    .withFailureHandler(function(e){document.getElementById(\"loading\").textContent=\"Error: \"+e.message;document.getElementById(\"loading\").classList.remove(\"hidden\");})\n    .loadInventory();\n}\n</script>\n</body>\n</html>";
}

function getRestockHTML() {
  return "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1.0\">\n<title>Restock</title>\n<style>\n@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:Inter,sans-serif;background:#0d0d1a;color:#e8e8f0;min-height:100vh;padding:0 0 90px}\n.header{background:#1A1A2E;padding:20px 24px;display:flex;align-items:center;gap:12px;border-bottom:1px solid #2a2a45;position:sticky;top:0;z-index:10}\n.header .icon{font-size:28px}\n.header h1{font-size:20px;font-weight:700;color:#fff}\n.header p{font-size:13px;color:#8888aa;margin-top:2px}\n.content{max-width:600px;margin:0 auto;padding:20px 16px}\n.section{background:#16162a;border:1px solid #2a2a45;border-radius:12px;padding:20px;margin-bottom:16px}\n.section-title{font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#27AE60;margin-bottom:14px}\nselect,input[type=text]{width:100%;padding:10px 14px;background:#0d0d1a;border:1px solid #2a2a45;border-radius:8px;color:#e8e8f0;font-size:14px;font-family:inherit;outline:none}\nselect:focus,input:focus{border-color:#27AE60}\n.field{margin-bottom:12px}\n.field label{display:block;font-size:13px;color:#8888aa;margin-bottom:6px}\n.brand-row{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid #1a1a2e}\n.brand-row:last-child{border-bottom:none}\n.brand-info{flex:1}\n.brand-name{font-size:14px;font-weight:500}\n.brand-cat{font-size:12px;color:#8888aa;margin-top:2px}\n.brand-stock{font-size:12px;font-weight:600;margin-top:2px}\n.ok{color:#27AE60}.low{color:#F39C12}.out{color:#E74C3C}\n.qty-wrap{display:flex;align-items:center;gap:6px}\n.qbtn{width:32px;height:32px;background:#2a2a45;border:none;border-radius:6px;color:#e8e8f0;font-size:20px;cursor:pointer;line-height:1}\n.qbtn:hover:not([disabled]){background:#27AE60}\n.qbtn[disabled]{opacity:.3;cursor:not-allowed}\n.qinp{width:48px;text-align:center;padding:6px 4px;font-size:15px;font-weight:600}\n.submit-bar{position:fixed;bottom:0;left:0;right:0;background:#1A1A2E;border-top:1px solid #2a2a45;padding:16px 24px}\n.sbtn{width:100%;max-width:600px;margin:0 auto;display:block;padding:16px;background:#27AE60;color:#fff;border:none;border-radius:12px;font-size:16px;font-weight:700;font-family:inherit;cursor:pointer}\n.sbtn:hover:not([disabled]){background:#219a52}\n.sbtn[disabled]{background:#333355;color:#555577;cursor:not-allowed}\n.badge{display:inline-block;background:#27AE60;color:#fff;border-radius:20px;padding:1px 8px;font-size:12px;font-weight:700;margin-left:8px}\n.hidden{display:none}\n.success{text-align:center;padding:60px 24px}\n.success .icon{font-size:64px;margin-bottom:16px}\n.success h2{color:#27AE60;font-size:24px;margin-bottom:8px}\n.success p{color:#8888aa;font-size:15px;margin-bottom:24px}\n.success button{padding:14px 32px;background:#27AE60;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;font-family:inherit}\n.loading{text-align:center;padding:60px 24px;color:#8888aa}\n.spinner{width:40px;height:40px;border:3px solid #2a2a45;border-top-color:#27AE60;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 16px}\n@keyframes spin{to{transform:rotate(360deg)}}\n.cat-hdr{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#555577;padding:12px 0 6px}\n</style>\n</head>\n<body>\n<div class=\"header\">\n  <div class=\"icon\">&#128230;</div>\n  <div><h1>Restock</h1><p>Log incoming stock</p></div>\n</div>\n<div class=\"content\">\n  <div id=\"loading\" class=\"loading\"><div class=\"spinner\"></div>Loading inventory...</div>\n  <div id=\"orderView\" class=\"hidden\">\n    <div class=\"section\">\n      <div class=\"section-title\">Staff &amp; Invoice</div>\n      <div class=\"field\"><label>Invoice #</label><input type=\"text\" id=\"invoiceNo\" placeholder=\"e.g. INV-2024-001\" required></div>\n      <div class=\"field\"><label>Delivered By</label><input type=\"text\" id=\"delivBy\" placeholder=\"Supplier or staff name\"></div>\n      <div class=\"field\"><label>Received By</label><input type=\"text\" id=\"recvBy\" placeholder=\"Who receives the stock\"></div>\n    </div>\n    <div class=\"section\">\n      <div class=\"section-title\">Brands <span class=\"badge\" id=\"count\">0</span></div>\n      <div id=\"brandList\"></div>\n    </div>\n  </div>\n  <div id=\"successView\" class=\"hidden success\">\n    <div class=\"icon\">&#9989;</div>\n    <h2>Restock submitted!</h2>\n    <p id=\"successMsg\"></p>\n    <button onclick=\"reset()\">New Restock</button>\n  </div>\n</div>\n<div id=\"submitBar\" class=\"submit-bar\">\n  <button class=\"sbtn\" id=\"submitBtn\" onclick=\"doSubmit()\" disabled>Confirm Restock</button>\n</div>\n<script>\nvar INV = null;\ngoogle.script.run\n  .withSuccessHandler(function(r) {\n    document.getElementById(\"loading\").classList.add(\"hidden\");\n    if (!r.success) { document.getElementById(\"loading\").textContent = \"Error: \" + r.error; document.getElementById(\"loading\").classList.remove(\"hidden\"); return; }\n    INV = r;\n    buildUI(r);\n    document.getElementById(\"orderView\").classList.remove(\"hidden\");\n  })\n  .withFailureHandler(function(e) {\n    document.getElementById(\"loading\").textContent = \"Error: \" + e.message;\n    document.getElementById(\"loading\").classList.remove(\"hidden\");\n  })\n  .loadInventory();\n\nfunction buildUI(r) {\n  var byCat = {};\n  r.brands.forEach(function(b) { if (!byCat[b.category]) byCat[b.category] = []; byCat[b.category].push(b); });\n  var list = document.getElementById(\"brandList\");\n  Object.keys(byCat).forEach(function(cat) {\n    var hdr = document.createElement(\"div\"); hdr.className = \"cat-hdr\"; hdr.textContent = cat; list.appendChild(hdr);\n    byCat[cat].forEach(function(b) {\n      var row = document.createElement(\"div\"); row.className = \"brand-row\";\n      var info = document.createElement(\"div\"); info.className = \"brand-info\";\n      var nm = document.createElement(\"div\"); nm.className = \"brand-name\"; nm.textContent = b.name;\n      var ct = document.createElement(\"div\"); ct.className = \"brand-cat\"; ct.textContent = b.category;\n      var sc = b.stock <= 0 ? \"out\" : b.stock <= b.min ? \"low\" : \"ok\";\n      var st = document.createElement(\"div\"); st.className = \"brand-stock \" + sc;\n      st.textContent = \"Current stock: \" + b.stock;\n      info.appendChild(nm); info.appendChild(ct); info.appendChild(st);\n      var wrap = document.createElement(\"div\"); wrap.className = \"qty-wrap\";\n      var bm = document.createElement(\"button\"); bm.className = \"qbtn\"; bm.textContent = \"-\"; bm.disabled = true;\n      bm.onclick = (function(n) { return function() { chg(n,-1); }; })(b.name);\n      var inp = document.createElement(\"input\"); inp.type = \"text\"; inp.className = \"qinp\";\n      inp.value = \"0\"; inp.readOnly = true; inp.id = \"q_\" + b.name.replace(/ /g,\"_\");\n      var bp = document.createElement(\"button\"); bp.className = \"qbtn\"; bp.textContent = \"+\";\n      bp.onclick = (function(n) { return function() { chg(n,1); }; })(b.name);\n      wrap.appendChild(bm); wrap.appendChild(inp); wrap.appendChild(bp);\n      row.appendChild(info); row.appendChild(wrap); list.appendChild(row);\n    });\n  });\n}\n\nfunction chg(name, delta) {\n  var inp = document.getElementById(\"q_\" + name.replace(/ /g,\"_\"));\n  var val = Math.max(0, parseInt(inp.value||0) + delta);\n  inp.value = val;\n  inp.previousSibling.disabled = (val <= 0);\n  upd();\n}\nfunction upd() {\n  var n = 0;\n  INV.brands.forEach(function(b) { var i=document.getElementById(\"q_\"+b.name.replace(/ /g,\"_\")); if(i&&parseInt(i.value)>0) n++; });\n  document.getElementById(\"count\").textContent = n;\n  document.getElementById(\"submitBtn\").disabled = (n === 0);\n}\nfunction doSubmit() {\n  var inv = document.getElementById(\"invoiceNo\").value.trim();\n  var del = document.getElementById(\"delivBy\").value.trim();\n  var rec = document.getElementById(\"recvBy\").value.trim();\n  if (!inv) { alert(\"Invoice # is required.\"); return; }\n  if (!del||!rec) { alert(\"Fill in Delivered By and Received By.\"); return; }\n  var items = [];\n  INV.brands.forEach(function(b) { var i=document.getElementById(\"q_\"+b.name.replace(/ /g,\"_\")); var q=parseInt(i?i.value:0); if(q>0) items.push({name:b.name,category:b.category,qty:q}); });\n  if (!items.length) return;\n  var btn = document.getElementById(\"submitBtn\"); btn.disabled=true; btn.textContent=\"Submitting...\";\n  google.script.run\n    .withSuccessHandler(function(r) {\n      if (r.success) {\n        document.getElementById(\"orderView\").classList.add(\"hidden\");\n        document.getElementById(\"submitBar\").classList.add(\"hidden\");\n        document.getElementById(\"successMsg\").textContent = r.count + \" item(s) restocked successfully.\";\n        document.getElementById(\"successView\").classList.remove(\"hidden\");\n      } else { alert(\"Error: \"+r.error); btn.disabled=false; btn.textContent=\"Confirm Restock\"; }\n    })\n    .withFailureHandler(function(e) { alert(\"Failed: \"+e.message); btn.disabled=false; btn.textContent=\"Confirm Restock\"; })\n    .submitRestock({deliveredBy:del, receivedBy:rec, invoice:document.getElementById(\"invoiceNo\").value.trim(), items:items});\n}\nfunction reset() {\n  document.getElementById(\"successView\").classList.add(\"hidden\");\n  document.getElementById(\"submitBar\").classList.remove(\"hidden\");\n  document.getElementById(\"loading\").textContent=\"Refreshing inventory...\";\n  document.getElementById(\"loading\").classList.remove(\"hidden\");\n  document.getElementById(\"orderView\").classList.add(\"hidden\");\n  document.getElementById(\"brandList\").innerHTML=\"\";\n  google.script.run\n    .withSuccessHandler(function(r){\n      document.getElementById(\"loading\").classList.add(\"hidden\");\n      if(!r.success){document.getElementById(\"loading\").textContent=\"Error: \"+r.error;document.getElementById(\"loading\").classList.remove(\"hidden\");return;}\n      INV=r; buildUI(r);\n      document.getElementById(\"count\").textContent=\"0\";\n      document.getElementById(\"submitBtn\").disabled=true;\n      document.getElementById(\"submitBtn\").textContent=\"Confirm Restock\";\n      document.getElementById(\"invoiceNo\").value=\"\";\n      document.getElementById(\"delivBy\").value=\"\";\n      document.getElementById(\"recvBy\").value=\"\";\n      document.getElementById(\"orderView\").classList.remove(\"hidden\");\n    })\n    .withFailureHandler(function(e){document.getElementById(\"loading\").textContent=\"Error: \"+e.message;document.getElementById(\"loading\").classList.remove(\"hidden\");})\n    .loadInventory();\n}\n</script>\n</body>\n</html>";
}


function getWeightHTML() {
  return "<!DOCTYPE html>\n<html>\n<head>\n<meta charset=\"UTF-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1.0\">\n<title>Bottle Weight Log</title>\n<style>\n@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');\n*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}\nbody{font-family:Inter,sans-serif;background:#0d0d1a;color:#e8e8f0;min-height:100vh;padding:0 0 90px}\n.header{background:#1A1A2E;padding:20px 24px;display:flex;align-items:center;gap:12px;border-bottom:1px solid #2a2a45;position:sticky;top:0;z-index:10}\n.header .icon{font-size:28px}\n.header h1{font-size:20px;font-weight:700;color:#fff}\n.header p{font-size:13px;color:#8888aa;margin-top:2px}\n.content{max-width:600px;margin:0 auto;padding:20px 16px}\n.section{background:#16162a;border:1px solid #2a2a45;border-radius:12px;padding:20px;margin-bottom:16px}\n.section-title{font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:#8E44AD;margin-bottom:14px}\nselect,input[type=text],input[type=date],input[type=number]{width:100%;padding:10px 14px;background:#0d0d1a;border:1px solid #2a2a45;border-radius:8px;color:#e8e8f0;font-size:14px;font-family:inherit;outline:none}\nselect:focus,input:focus{border-color:#8E44AD}\n.field{margin-bottom:12px}\n.field label{display:block;font-size:13px;color:#8888aa;margin-bottom:6px}\n.brand-row{display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid #1a1a2e}\n.brand-row:last-child{border-bottom:none}\n.brand-info{flex:1}\n.brand-name{font-size:14px;font-weight:500}\n.brand-cat{font-size:12px;color:#8888aa;margin-top:2px}\n.oz-input{width:90px;text-align:center;padding:8px;font-size:15px;font-weight:600;border-radius:8px}\n.oz-label{font-size:12px;color:#8888aa;margin-left:6px}\n.submit-bar{position:fixed;bottom:0;left:0;right:0;background:#1A1A2E;border-top:1px solid #2a2a45;padding:16px 24px}\n.sbtn{width:100%;max-width:600px;margin:0 auto;display:block;padding:16px;background:#8E44AD;color:#fff;border:none;border-radius:12px;font-size:16px;font-weight:700;font-family:inherit;cursor:pointer}\n.sbtn:hover:not([disabled]){background:#7d3c98}\n.sbtn[disabled]{background:#333355;color:#555577;cursor:not-allowed}\n.badge{display:inline-block;background:#8E44AD;color:#fff;border-radius:20px;padding:1px 8px;font-size:12px;font-weight:700;margin-left:8px}\n.hidden{display:none}\n.success{text-align:center;padding:60px 24px}\n.success .icon{font-size:64px;margin-bottom:16px}\n.success h2{color:#8E44AD;font-size:24px;margin-bottom:8px}\n.success p{color:#8888aa;font-size:15px;margin-bottom:24px}\n.success button{padding:14px 32px;background:#8E44AD;color:#fff;border:none;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;font-family:inherit}\n.loading{text-align:center;padding:60px 24px;color:#8888aa}\n.spinner{width:40px;height:40px;border:3px solid #2a2a45;border-top-color:#8E44AD;border-radius:50%;animation:spin .8s linear infinite;margin:0 auto 16px}\n@keyframes spin{to{transform:rotate(360deg)}}\n.cat-hdr{font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#555577;padding:12px 0 6px}\n.hint{font-size:11px;color:#555577;margin-top:4px}\n.brand-block{border-bottom:1px solid #1a1a2e;padding:10px 0}\n.brand-block:last-child{border-bottom:none}\n.brand-row-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px}\n.brand-total{font-size:13px;font-weight:700;color:#8E44AD;white-space:nowrap}\n.bottle-row{display:flex;align-items:center;gap:8px;margin-bottom:6px}\n.bottle-lbl{font-size:12px;color:#555577;min-width:54px}\n.add-bottle-btn{background:none;border:1px dashed #2a2a45;color:#8888aa;border-radius:6px;padding:5px 12px;font-size:12px;cursor:pointer;margin-top:4px;font-family:inherit}\n.add-bottle-btn:hover{border-color:#8E44AD;color:#8E44AD}\n.rm-bottle-btn{background:none;border:none;color:#555577;cursor:pointer;font-size:18px;padding:2px 6px;line-height:1}\n.rm-bottle-btn:hover{color:#e94560}\n</style>\n</head>\n<body>\n<div class=\"header\">\n  <div class=\"icon\">&#9878;</div>\n  <div><h1>Bottle Weight Log</h1><p>Record oz remaining per open bottle</p></div>\n</div>\n<div class=\"content\">\n  <div id=\"loading\" class=\"loading\"><div class=\"spinner\"></div>Loading...</div>\n  <div id=\"orderView\" class=\"hidden\">\n    <div class=\"section\">\n      <div class=\"section-title\">Details</div>\n      <div class=\"field\"><label>Date</label><input type=\"date\" id=\"logDate\"></div>\n      <div class=\"field\"><label>Bar Section</label><select id=\"barSel\"></select></div>\n      <div class=\"field\"><label>Recorded By</label><input type=\"text\" id=\"recordedBy\" placeholder=\"Your name\"></div>\n    </div>\n    <div class=\"section\">\n      <div class=\"section-title\">Oz Remaining <span class=\"badge\" id=\"count\">0</span></div>\n      <p class=\"hint\" style=\"margin-bottom:12px;font-size:12px;color:#555577\">For each open bottle: enter the full bottle size and oz remaining. Leave empty to skip a brand.</p>\n      <div id=\"brandList\"></div>\n    </div>\n  </div>\n  <div id=\"successView\" class=\"hidden success\">\n    <div class=\"icon\">&#9989;</div>\n    <h2>Logged!</h2>\n    <p id=\"successMsg\"></p>\n    <button onclick=\"reset()\">New Log</button>\n  </div>\n</div>\n<div id=\"submitBar\" class=\"submit-bar\">\n  <button class=\"sbtn\" id=\"submitBtn\" onclick=\"doSubmit()\" disabled>Save Weight Log</button>\n</div>\n<script>\nvar INV = null;\n\n// Set today's date as default\ndocument.getElementById(\"logDate\").value = new Date().toISOString().split(\"T\")[0];\n\ngoogle.script.run\n  .withSuccessHandler(function(r) {\n    document.getElementById(\"loading\").classList.add(\"hidden\");\n    if (!r.success) { document.getElementById(\"loading\").textContent = \"Error: \" + r.error; document.getElementById(\"loading\").classList.remove(\"hidden\"); return; }\n    INV = r;\n    buildUI(r);\n    document.getElementById(\"orderView\").classList.remove(\"hidden\");\n  })\n  .withFailureHandler(function(e) {\n    document.getElementById(\"loading\").textContent = \"Error: \" + e.message;\n    document.getElementById(\"loading\").classList.remove(\"hidden\");\n  })\n  .loadInventory();\n\nfunction buildUI(r) {\n  var sel = document.getElementById(\"barSel\");\n  r.bars.forEach(function(b) { var o = document.createElement(\"option\"); o.value = b; o.textContent = b; sel.appendChild(o); });\n\n  var byCat = {};\n  r.brands.forEach(function(b) { if (!byCat[b.category]) byCat[b.category] = []; byCat[b.category].push(b); });\n  var list = document.getElementById(\"brandList\");\n\n  Object.keys(byCat).forEach(function(cat) {\n    var hdr = document.createElement(\"div\"); hdr.className = \"cat-hdr\"; hdr.textContent = cat; list.appendChild(hdr);\n    byCat[cat].forEach(function(b) {\n      var brandId = b.name.replace(/ /g,\"_\");\n      var wrap = document.createElement(\"div\");\n      wrap.className = \"brand-block\";\n      wrap.id = \"block_\" + brandId;\n\n      // Header row\n      var hdr = document.createElement(\"div\"); hdr.className = \"brand-row-hdr\";\n      var info = document.createElement(\"div\"); info.className = \"brand-info\";\n      var nm = document.createElement(\"div\"); nm.className = \"brand-name\"; nm.textContent = b.name;\n      var ct = document.createElement(\"div\"); ct.className = \"brand-cat\"; ct.textContent = b.category;\n      info.appendChild(nm); info.appendChild(ct);\n\n      // Total display\n      var total = document.createElement(\"div\"); total.className = \"brand-total\";\n      total.id = \"total_\" + brandId;\n      total.textContent = \"0 oz total\";\n\n      hdr.appendChild(info); hdr.appendChild(total);\n      wrap.appendChild(hdr);\n\n      // Bottles container\n      var bottles = document.createElement(\"div\");\n      bottles.id = \"bottles_\" + brandId;\n      wrap.appendChild(bottles);\n\n      // Add bottle button\n      var addBtn = document.createElement(\"button\");\n      addBtn.className = \"add-bottle-btn\";\n      addBtn.textContent = \"+ Add bottle\";\n      addBtn.type = \"button\";\n      addBtn.onclick = (function(bid, bname, bcat) {\n        return function() { addBottle(bid, bname, bcat); };\n      })(brandId, b.name, b.category);\n      wrap.appendChild(addBtn);\n\n      list.appendChild(wrap);\n      // Add first bottle by default\n      addBottle(brandId, b.name, b.category);\n    });\n  });\n}\n\nvar bottleCount = {};\n\nfunction addBottle(brandId, brandName, brandCat) {\n  var container = document.getElementById(\"bottles_\" + brandId);\n  if (!bottleCount[brandId]) bottleCount[brandId] = 0;\n  bottleCount[brandId]++;\n  var idx = bottleCount[brandId];\n\n  var row = document.createElement(\"div\"); row.className = \"bottle-row\";\n  row.id = \"bottlerow_\" + brandId + \"_\" + idx;\n\n  var lbl = document.createElement(\"span\"); lbl.className = \"bottle-lbl\";\n  lbl.textContent = \"Bottle \" + idx;\n\n  // Full size input\n  var sizeInp = document.createElement(\"input\");\n  sizeInp.type = \"number\"; sizeInp.className = \"oz-input\";\n  sizeInp.min = \"0\"; sizeInp.step = \"0.1\"; sizeInp.placeholder = \"full size\";\n  sizeInp.id = \"size_\" + brandId + \"_\" + idx;\n  sizeInp.dataset.sizebid = brandId;\n  var sizeLbl = document.createElement(\"span\"); sizeLbl.className = \"oz-label\"; sizeLbl.textContent = \"full sz\";\n\n  // Remaining input\n  var inp = document.createElement(\"input\");\n  inp.type = \"number\"; inp.className = \"oz-input\";\n  inp.min = \"0\"; inp.step = \"0.1\"; inp.placeholder = \"oz left\";\n  inp.id = \"oz_\" + brandId + \"_\" + idx;\n  inp.dataset.brand = brandName;\n  inp.dataset.cat   = brandCat;\n  inp.dataset.bid   = brandId;\n  inp.oninput = function() { updateTotal(brandId); upd(); };\n\n  var rmBtn = document.createElement(\"button\");\n  rmBtn.type = \"button\"; rmBtn.className = \"rm-bottle-btn\"; rmBtn.textContent = \"\u00d7\";\n  rmBtn.onclick = (function(bid, i) {\n    return function() {\n      var el = document.getElementById(\"bottlerow_\" + bid + \"_\" + i);\n      if (el) el.parentNode.removeChild(el);\n      updateTotal(bid); upd();\n    };\n  })(brandId, idx);\n\n  if (idx === 1) rmBtn.style.visibility = \"hidden\";\n\n  var ozLbl = document.createElement(\"span\"); ozLbl.className = \"oz-label\"; ozLbl.textContent = \"rem\";\n\n  row.appendChild(lbl);\n  row.appendChild(inp); row.appendChild(ozLbl);\n  row.appendChild(sizeInp); row.appendChild(sizeLbl);\n  row.appendChild(rmBtn);\n  container.appendChild(row);\n}\n\nfunction updateTotal(brandId) {\n  var inputs = document.querySelectorAll(\"[data-bid='\" + brandId + \"']\");\n  var total = 0;\n  inputs.forEach(function(inp) { if (inp.value !== \"\") total += parseFloat(inp.value) || 0; });\n  var el = document.getElementById(\"total_\" + brandId);\n  if (el) el.textContent = total > 0 ? total.toFixed(1) + \" oz total\" : \"0 oz total\";\n}\n\nfunction upd() {\n  var n = 0;\n  INV.brands.forEach(function(b) {\n    var bid = b.name.replace(/ /g,\"_\");\n    var inputs = document.querySelectorAll(\"[data-bid='\" + bid + \"']\");\n    inputs.forEach(function(inp) { if (inp.value !== \"\") n++; });\n  });\n  document.getElementById(\"count\").textContent = n;\n  document.getElementById(\"submitBtn\").disabled = (n === 0);\n}\n\nfunction doSubmit() {\n  var date = document.getElementById(\"logDate\").value;\n  var bar  = document.getElementById(\"barSel\").value;\n  var rec  = document.getElementById(\"recordedBy\").value.trim();\n  if (!date) { alert(\"Please select a date.\"); return; }\n  if (!rec)  { alert(\"Please fill in Recorded By.\"); return; }\n  var items = [];\n  INV.brands.forEach(function(b) {\n    var bid = b.name.replace(/ /g,\"_\");\n    var inputs = document.querySelectorAll(\"[data-bid='\" + bid + \"']\");\n    inputs.forEach(function(inp) {\n      if (inp.value !== \"\") {\n        var sizeEl = document.getElementById(\"size_\" + bid + \"_\" + inp.id.split(\"_\").pop());\n        var fullOz = sizeEl && sizeEl.value !== \"\" ? parseFloat(sizeEl.value) : null;\n        items.push({ name: b.name, category: b.category, oz: parseFloat(inp.value), fullOz: fullOz });\n      }\n    });\n  });\n  if (!items.length) return;\n  var btn = document.getElementById(\"submitBtn\"); btn.disabled = true; btn.textContent = \"Saving...\";\n  google.script.run\n    .withSuccessHandler(function(r) {\n      if (r.success) {\n        document.getElementById(\"orderView\").classList.add(\"hidden\");\n        document.getElementById(\"submitBar\").classList.add(\"hidden\");\n        document.getElementById(\"successMsg\").textContent = r.count + \" bottle(s) logged for \" + bar + \".\";\n        document.getElementById(\"successView\").classList.remove(\"hidden\");\n      } else { alert(\"Error: \" + r.error); btn.disabled = false; btn.textContent = \"Save Weight Log\"; }\n    })\n    .withFailureHandler(function(e) { alert(\"Failed: \" + e.message); btn.disabled = false; btn.textContent = \"Save Weight Log\"; })\n    .submitWeight({ bar: bar, recordedBy: rec, date: date, items: items });\n}\n\nfunction reset() {\n  document.getElementById(\"successView\").classList.add(\"hidden\");\n  document.getElementById(\"submitBar\").classList.remove(\"hidden\");\n  document.getElementById(\"loading\").textContent = \"Refreshing...\";\n  document.getElementById(\"loading\").classList.remove(\"hidden\");\n  document.getElementById(\"orderView\").classList.add(\"hidden\");\n  document.getElementById(\"brandList\").innerHTML = \"\";\n  var sel = document.getElementById(\"barSel\"); while(sel.options.length > 0) sel.remove(0);\n  google.script.run\n    .withSuccessHandler(function(r) {\n      document.getElementById(\"loading\").classList.add(\"hidden\");\n      if (!r.success) { document.getElementById(\"loading\").textContent = \"Error: \" + r.error; document.getElementById(\"loading\").classList.remove(\"hidden\"); return; }\n      INV = r; buildUI(r);\n      document.getElementById(\"count\").textContent = \"0\";\n      document.getElementById(\"submitBtn\").disabled = true;\n      document.getElementById(\"submitBtn\").textContent = \"Save Weight Log\";\n      document.getElementById(\"recordedBy\").value = \"\";\n      document.getElementById(\"logDate\").value = new Date().toISOString().split(\"T\")[0];\n      document.getElementById(\"orderView\").classList.remove(\"hidden\");\n    })\n    .withFailureHandler(function(e) { document.getElementById(\"loading\").textContent = \"Error: \" + e.message; document.getElementById(\"loading\").classList.remove(\"hidden\"); })\n    .loadInventory();\n}\n</script>\n</body>\n</html>";
}