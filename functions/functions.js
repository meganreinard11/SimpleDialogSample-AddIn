/* eslint-disable no-console */
"use strict";

// Local copies of IDs/headers (kept in sync with error-handler.js)
var DIAG_SHEET = "_Diagnostics";
var DIAG_TABLE_NAME = "ErrorLog";
var DIAG_HEADERS = ["Timestamp","Action","Message","Code","Location","Statement","Stack"];

var clickEvent;

// Wrap the writeToDoc in showNotification because showNotification is called in dialog.js but must be defined differently when the dialog is called from a task pane instead of a custom menu command.
function showNotification(text) {
    writeToDoc(text);
    //Required, call event.completed to let the platform know you are done processing.
    clickEvent.completed();
}

//Notice function needs to be in global namespace
function doSomethingAndShowDialog(event) {
    clickEvent = event;
    writeToDoc("Ribbon button clicked.");
    openDialog();
}

function writeToDoc(text) {
  Office.context.document.setSelectedDataAsync(text, function (asyncResult) {
      var error = asyncResult.error;
      if (asyncResult.status === "failed") {
          console.log("Unable to write to the document: " + asyncResult.error.message);
      }
  });
}

// Helpers that gracefully use ErrorHandler if present
function _getEnvFlag() {
  if (typeof window !== "undefined" && window.ErrorHandler && window.ErrorHandler.getEnvFlag) {
    return window.ErrorHandler.getEnvFlag();
  }
  return Promise.resolve({ debug: false });
}

function _enqueue(work) {
  if (typeof window !== "undefined" && window.ErrorHandler && window.ErrorHandler.enqueueDiagWrite) {
    return window.ErrorHandler.enqueueDiagWrite(work);
  }
  // Lightweight local queue if ErrorHandler isn't loaded yet
  _enqueue._q = _enqueue._q || Promise.resolve();
  _enqueue._q = _enqueue._q.then(work).catch(function (e) { console.warn("[diagnostics] queued write failed", e); });
  return _enqueue._q;
}

function _handleError(err, context) {
  if (typeof window !== "undefined" && window.ErrorHandler && window.ErrorHandler.handleError) {
    return window.ErrorHandler.handleError(err, context);
  }
  console.error("[Clear Diagnostics] failed", err);
  try { alert((context && context.userMessage) || "Couldn't clear the diagnostics log."); } catch (_) {}
  return Promise.resolve();
}

// Minimal colLetter for header range
function colLetter(n) {
  var s = "";
  while (n > 0) {
    var m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Fast “clear”: drop and recreate the diagnostics sheet/table. */
function clearDiagnostics() {
  return Excel.run(async function (ctx) {
    var wb = ctx.workbook;
    var ws = wb.worksheets.getItemOrNullObject(DIAG_SHEET);
    ws.load("isNullObject");
    await ctx.sync();

    if (!ws.isNullObject) {
      ws.delete(); // removes even if hidden
      await ctx.sync();
    }

    // Recreate sheet, headers, table
    ws = wb.worksheets.add(DIAG_SHEET);
    var headerRange = ws.getRange("A1:" + colLetter(DIAG_HEADERS.length) + "1");
    headerRange.values = [DIAG_HEADERS.slice(0)];
    headerRange.format.font.bold = true;

    var table = wb.tables.add(ws.name + "!A1:" + colLetter(DIAG_HEADERS.length) + "1", true);
    table.name = DIAG_TABLE_NAME;

    try { ws.visibility = "VeryHidden"; } catch (_) {}
    await ctx.sync();
  });
}

/**
 * Ribbon command onAction. Must call event.completed().
 */
async function clearDiagnosticsCommand(event) {
  try {
    var env = await _getEnvFlag();
    var proceed = true;
    if (env && env.debug) {
      try { proceed = confirm("Clear the diagnostics log? This cannot be undone."); } catch (_) {}
    }
    if (proceed) {
      await _enqueue(function () { return clearDiagnostics(); });
      try { alert("Diagnostics log cleared."); } catch (_) {}
    }
  } catch (err) {
    await _handleError(err, {
      action: "Clear Diagnostics",
      userMessage: "Couldn't clear the diagnostics log.",
      forceLogToSheet: true
    });
  } finally {
    try { event.completed(); } catch (_) {}
  }
}

// Expose for Office to locate it
if (typeof window !== "undefined") {
  window.clearDiagnosticsCommand = clearDiagnosticsCommand;
} else if (typeof self !== "undefined") {
  self.clearDiagnosticsCommand = clearDiagnosticsCommand;
}

// Optional: CommonJS export (handy for tests)
if (typeof module !== "undefined" && module.exports) {
  module.exports = { clearDiagnosticsCommand: clearDiagnosticsCommand, clearDiagnostics: clearDiagnostics };
}
