/* ════════════════════════════════════════════════════════════════
   Kascade — Office.js Add-in
   Features: Sheet Cascade  ·  Row Group Manager
   Requires: ExcelApi 1.9+
   ════════════════════════════════════════════════════════════════ */

// ── State ────────────────────────────────────────────────────────
let cachedRangeValues = []; // values from the selected named range

// ── Initialise ───────────────────────────────────────────────────
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    initTabs();
    initCascade();
    initGroups();
  }
});

/* ═══════════════════════════════════════════════════════════════
   TAB NAVIGATION
   ═══════════════════════════════════════════════════════════════ */
function initTabs() {
  document.querySelectorAll(".tab").forEach(function (tab) {
    tab.addEventListener("click", function () {
      document.querySelectorAll(".tab").forEach(function (t) { t.classList.remove("active"); });
      document.querySelectorAll(".tab-content").forEach(function (c) { c.classList.remove("active"); });
      tab.classList.add("active");
      document.getElementById("tab-" + tab.dataset.tab).classList.add("active");

      // Refresh data when switching tabs
      if (tab.dataset.tab === "groups") scanAndRefresh();
      if (tab.dataset.tab === "cascade") loadNamedRanges();
    });
  });
}

/* ═══════════════════════════════════════════════════════════════
   FEATURE 1: CASCADE SHEET
   ═══════════════════════════════════════════════════════════════ */
function initCascade() {
  document.getElementById("btn-pick-cell").addEventListener("click", pickActiveCell);
  document.getElementById("btn-refresh-ranges").addEventListener("click", loadNamedRanges);
  document.getElementById("cascade-range").addEventListener("change", onRangeSelected);
  document.getElementById("btn-cascade").addEventListener("click", executeCascade);

  // Initial load
  pickActiveCell();
  loadNamedRanges();
}

// Pick the currently selected cell address
function pickActiveCell() {
  Excel.run(function (ctx) {
    var cell = ctx.workbook.getSelectedRange();
    cell.load("address");
    return ctx.sync().then(function () {
      // address comes back as "Sheet1!$B$2", strip sheet prefix
      // For multi-cell selections, take only the top-left cell
      var addr = cell.address.replace(/^.*!/, "").replace(/\$/g, "").split(":")[0];
      document.getElementById("cascade-cell").value = addr;
    });
  }).catch(function (err) { showStatus("cascade", err.message, "error"); });
}

// Populate the named-range dropdown
function loadNamedRanges() {
  Excel.run(function (ctx) {
    var names = ctx.workbook.names;
    names.load("items");
    return ctx.sync().then(function () {
      var sel = document.getElementById("cascade-range");
      sel.innerHTML = '<option value="">— Select a named range —</option>';

      // Built-in names to skip
      var builtIn = ["print_area", "print_titles", "_xlnm.print_area", "_xlnm.print_titles",
                     "sheet_title", "_xlnm.database", "_xlnm.criteria", "_xlnm.extract"];
      var added = 0;

      names.items.forEach(function (n) {
        var nameLower = n.name.toLowerCase();
        // Skip built-in Excel names
        if (builtIn.indexOf(nameLower) !== -1) return;
        if (nameLower.indexOf("_xlnm.") === 0) return;

        var opt = document.createElement("option");
        opt.value = n.name;
        opt.textContent = n.name;
        sel.appendChild(opt);
        added++;
      });

      if (added === 0) {
        sel.innerHTML = '<option value="">No named ranges found</option>';
      }
    });
  }).catch(function (err) { showStatus("cascade", err.message, "error"); });
}

// When a named range is selected, preview its values
function onRangeSelected() {
  var name = document.getElementById("cascade-range").value;
  var previewSection = document.getElementById("cascade-preview-section");
  var btn = document.getElementById("btn-cascade");

  if (!name) {
    previewSection.style.display = "none";
    btn.disabled = true;
    cachedRangeValues = [];
    return;
  }

  Excel.run(function (ctx) {
    var namedItem = ctx.workbook.names.getItem(name);
    var range = namedItem.getRange();
    range.load("values");
    return ctx.sync().then(function () {
      // Flatten to a 1-D list of non-empty values
      cachedRangeValues = [];
      range.values.forEach(function (row) {
        row.forEach(function (val) {
          if (val !== null && val !== "") cachedRangeValues.push(String(val));
        });
      });

      // Render preview
      var list = document.getElementById("cascade-preview");
      list.innerHTML = "";
      cachedRangeValues.forEach(function (v) {
        var li = document.createElement("li");
        li.textContent = v;
        list.appendChild(li);
      });

      document.getElementById("cascade-count").textContent = cachedRangeValues.length + " values";
      previewSection.style.display = "block";
      btn.disabled = cachedRangeValues.length === 0;
    });
  }).catch(function (err) { showStatus("cascade", err.message, "error"); });
}

// Execute the cascade operation
function executeCascade() {
  var cellAddr = document.getElementById("cascade-cell").value.trim();
  if (!cellAddr) { showStatus("cascade", "Pick a target cell first.", "error"); return; }
  if (cachedRangeValues.length === 0) { showStatus("cascade", "Select a named range with values.", "error"); return; }

  var btn = document.getElementById("btn-cascade");
  btn.disabled = true;
  btn.textContent = "Working...";
  showStatus("cascade", "Creating " + cachedRangeValues.length + " sheets...", "info");

  Excel.run(function (ctx) {
    var sourceSheet = ctx.workbook.worksheets.getActiveWorksheet();
    sourceSheet.load("name");
    return ctx.sync().then(function () {
      var baseName = sourceSheet.name;
      var created = 0;
      var total = cachedRangeValues.length;

      // Build a promise chain that copies one sheet at a time
      var chain = Promise.resolve();

      cachedRangeValues.forEach(function (val, idx) {
        chain = chain.then(function () {
          return Excel.run(function (innerCtx) {
            var src = innerCtx.workbook.worksheets.getItem(baseName);
            var copy = src.copy(Excel.WorksheetPositionType.end);
            // Load all existing sheet names to ensure uniqueness
            var sheets = innerCtx.workbook.worksheets;
            sheets.load("items/name");
            copy.load("name");
            return innerCtx.sync().then(function () {
              // Build set of existing names
              var existingNames = {};
              sheets.items.forEach(function (s) { existingNames[s.name.toLowerCase()] = true; });
              // Rename the new sheet with uniqueness check
              var safeName = getUniqueSheetName(cleanSheetName(val), existingNames);
              copy.name = safeName;
              // Set the target cell
              copy.getRange(cellAddr).values = [[val]];
              return innerCtx.sync();
            }).then(function () {
              created++;
              showStatus("cascade", "Created " + created + " of " + total + "...", "info");
            });
          });
        });
      });

      return chain.then(function () {
        showStatus("cascade", "Done — created " + total + " sheet(s).", "success");
        btn.disabled = false;
        btn.textContent = "Cascade";
      });
    });
  }).catch(function (err) {
    showStatus("cascade", "Error: " + err.message, "error");
    btn.disabled = false;
    btn.textContent = "Cascade";
  });
}

// Sanitise a string for use as a sheet name (max 31 chars, no illegal chars)
function cleanSheetName(raw) {
  var s = String(raw).replace(/[\/\\?*\[\]:]/g, "_").trim();
  if (s.length > 31) s = s.substring(0, 31);
  if (s === "" || s === "History") s = "Sheet";
  return s;
}

// Ensure sheet name is unique by appending _2, _3, etc.
function getUniqueSheetName(baseName, existingNames) {
  if (!existingNames[baseName.toLowerCase()]) return baseName;
  var counter = 2;
  while (true) {
    var suffix = "_" + counter;
    var candidate = baseName.length + suffix.length > 31
      ? baseName.substring(0, 31 - suffix.length) + suffix
      : baseName + suffix;
    if (!existingNames[candidate.toLowerCase()]) return candidate;
    counter++;
    if (counter > 999) return baseName + "_" + Date.now(); // safety valve
  }
}

/* ═══════════════════════════════════════════════════════════════
   FEATURE 2: ROW GROUP MANAGER (inline cell tags)
   ═══════════════════════════════════════════════════════════════
   Workflow:
   1. User types tags in a chosen column (e.g. #hide, #detail, #assumptions)
   2. Clicks "Refresh" in the panel
   3. Add-in scans the column, discovers all unique tags
   4. Shows each group with hide/show/toggle buttons
   Tags must start with "#". Everything after # is the group name.
   ═══════════════════════════════════════════════════════════════ */

// In-memory cache of last scan results: { "#hide": [3, 5, 7], "#detail": [10, 11] }
var scannedGroups = {};

function initGroups() {
  document.getElementById("btn-refresh-groups").addEventListener("click", scanAndRefresh);
  document.getElementById("btn-show-all").addEventListener("click", showAllRows);
  document.getElementById("btn-hide-tagged").addEventListener("click", hideAllTagged);
}

// Scan ALL columns in the used range for # tags and build the groups list
function scanAndRefresh() {
  Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    usedRange.load(["values", "rowCount", "columnCount", "columnIndex"]);
    return ctx.sync().then(function () {
      var values = usedRange.values;
      var rowCount = usedRange.rowCount;
      var colCount = usedRange.columnCount;
      var startCol = usedRange.columnIndex; // 0-based column offset

      if (rowCount === 0) {
        scannedGroups = {};
        renderGroupsList();
        showStatus("groups", "Sheet is empty.", "info");
        return;
      }

      scannedGroups = {};
      var totalTagged = 0;
      var autoHideCols = []; // columns to auto-hide via #hidecol:X
      var autoHideRows = []; // rows to auto-hide via #hide / #hiderow

      // Scan every cell in the used range
      for (var r = 0; r < rowCount; r++) {
        for (var c = 0; c < colCount; c++) {
          var cellVal = String(values[r][c]).trim();
          if (cellVal.indexOf("#") !== 0 || cellVal.length <= 1) continue;

          var tag = cellVal.toLowerCase();
          var rowNum = r + 1; // 1-based row number (relative to row 1, not usedRange start)
          // Adjust if usedRange doesn't start at row 1
          // Actually usedRange.values[0] corresponds to the first row of the usedRange
          // We need absolute row numbers for hiding
          // usedRange starts at rowIndex (0-based)

          // We'll fix row numbering after loading rowIndex

          // Check for #hidecol:X,Y,Z syntax
          if (tag.indexOf("#hidecol:") === 0) {
            var colsPart = cellVal.substring(9).trim().toUpperCase();
            colsPart.split(",").forEach(function (ch) {
              ch = ch.trim();
              if (ch && autoHideCols.indexOf(ch) === -1) autoHideCols.push(ch);
            });
            totalTagged++;
            continue;
          }

          // #hide and #hiderow auto-hide the row
          if (tag === "#hide" || tag === "#hiderow") {
            if (autoHideRows.indexOf(rowNum) === -1) autoHideRows.push(rowNum);
            totalTagged++;
            continue;
          }

          // All other # tags become toggleable groups
          if (!scannedGroups[tag]) scannedGroups[tag] = [];
          if (scannedGroups[tag].indexOf(rowNum) === -1) {
            scannedGroups[tag].push(rowNum);
          }
          totalTagged++;
        }
      }

      // We need absolute row numbers — load rowIndex to adjust
      // usedRange.values[r] corresponds to absolute row (usedRange.rowIndex + r + 1) in 1-based
      // But we already loaded usedRange — let's reload with rowIndex
      // Actually rowIndex isn't directly on usedRange.load — we need to get it differently
      // The used range address tells us the start row
      usedRange.load("rowIndex");
      return ctx.sync().then(function () {
        var rowOffset = usedRange.rowIndex; // 0-based

        // Adjust all row numbers to be absolute (1-based)
        function adjustRows(arr) {
          return arr.map(function (r) { return r + rowOffset; });
        }

        autoHideRows = adjustRows(autoHideRows);

        Object.keys(scannedGroups).forEach(function (tag) {
          scannedGroups[tag] = adjustRows(scannedGroups[tag]);
          scannedGroups[tag].sort(function (a, b) { return a - b; });
        });

        // Auto-hide rows
        if (autoHideRows.length > 0) {
          autoHideRows.forEach(function (r) {
            sheet.getRange(r + ":" + r).format.rowHidden = true;
          });
        }

        // Auto-hide columns
        if (autoHideCols.length > 0) {
          autoHideCols.forEach(function (ch) {
            sheet.getRange(ch + ":" + ch).format.columnHidden = true;
          });
        }

        return ctx.sync().then(function () {
          renderGroupsList();

          var groupCount = Object.keys(scannedGroups).length;
          if (groupCount === 0 && autoHideRows.length === 0 && autoHideCols.length === 0) {
            showStatus("groups", "No # tags found.", "info");
          } else {
            var msg = "Scanned " + colCount + " column(s). Found " + totalTagged + " tag(s).";
            if (autoHideRows.length > 0) {
              msg += " Auto-hidden " + autoHideRows.length + " row(s).";
            }
            if (autoHideCols.length > 0) {
              msg += " Auto-hidden column(s): " + autoHideCols.join(", ") + ".";
            }
            showStatus("groups", msg, "success");
          }
        });
      });
    });
  }).catch(function (err) { showStatus("groups", err.message, "error"); });
}

// Render the discovered groups in the panel
function renderGroupsList() {
  var container = document.getElementById("groups-list");
  var keys = Object.keys(scannedGroups);

  if (keys.length === 0) {
    container.innerHTML = '<p class="hint" style="padding:12px;">No groups found. Type # tags in your column and click Refresh.</p>';
    return;
  }

  container.innerHTML = "";
  keys.forEach(function (tag) {
    var rows = scannedGroups[tag];
    var div = document.createElement("div");
    div.className = "group-item";
    div.innerHTML =
      '<div class="group-info">' +
        '<span class="group-name">' + escapeHtml(tag) + '</span>' +
        '<span class="group-meta">' + rows.length + ' row(s): ' + summariseRows(rows) + '</span>' +
      '</div>' +
      '<div class="group-actions">' +
        '<button class="btn-icon" data-action="toggle" data-group="' + escapeAttr(tag) + '" title="Toggle visibility">&#128065;</button>' +
        '<button class="btn-icon" data-action="hide" data-group="' + escapeAttr(tag) + '" title="Hide rows">&#8863;</button>' +
        '<button class="btn-icon" data-action="show" data-group="' + escapeAttr(tag) + '" title="Show rows">&#8862;</button>' +
      '</div>';
    container.appendChild(div);
  });

  // Attach event listeners
  container.querySelectorAll(".btn-icon").forEach(function (btn) {
    btn.addEventListener("click", function () {
      var action = btn.dataset.action;
      var group = btn.dataset.group;
      if (action === "toggle") toggleGroupVisibility(group);
      if (action === "hide")   setGroupVisibility(group, true);
      if (action === "show")   setGroupVisibility(group, false);
    });
  });
}

// Summarise row numbers: "1-5, 8, 10-12"
function summariseRows(rows) {
  if (rows.length === 0) return "";
  rows = rows.slice().sort(function (a, b) { return a - b; });
  var parts = [];
  var start = rows[0];
  var end = rows[0];
  for (var i = 1; i < rows.length; i++) {
    if (rows[i] === end + 1) {
      end = rows[i];
    } else {
      parts.push(start === end ? String(start) : start + "-" + end);
      start = rows[i];
      end = rows[i];
    }
  }
  parts.push(start === end ? String(start) : start + "-" + end);
  return parts.join(", ");
}

// Toggle visibility of a row group
function toggleGroupVisibility(groupName) {
  var rows = scannedGroups[groupName] || [];
  if (rows.length === 0) return;

  Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    // Check current state of first row to decide toggle direction
    var firstRowRange = sheet.getRange(rows[0] + ":" + rows[0]);
    firstRowRange.format.load("rowHidden");
    return ctx.sync().then(function () {
      var shouldHide = !firstRowRange.format.rowHidden;
      rows.forEach(function (r) {
        sheet.getRange(r + ":" + r).format.rowHidden = shouldHide;
      });
      return ctx.sync().then(function () {
        showStatus("groups", (shouldHide ? "Hidden" : "Shown") + " \"" + groupName + "\" (" + rows.length + " rows).", "success");
      });
    });
  }).catch(function (err) { showStatus("groups", err.message, "error"); });
}

// Explicitly hide or show a group
function setGroupVisibility(groupName, hide) {
  var rows = scannedGroups[groupName] || [];
  if (rows.length === 0) return;

  Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    rows.forEach(function (r) {
      sheet.getRange(r + ":" + r).format.rowHidden = hide;
    });
    return ctx.sync().then(function () {
      showStatus("groups", (hide ? "Hidden" : "Shown") + " \"" + groupName + "\" (" + rows.length + " rows).", "success");
    });
  }).catch(function (err) { showStatus("groups", err.message, "error"); });
}

// Show all rows on the active sheet
function showAllRows() {
  Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    var used = sheet.getUsedRange();
    used.format.rowHidden = false;
    return ctx.sync().then(function () {
      showStatus("groups", "All rows are now visible.", "success");
    });
  }).catch(function (err) { showStatus("groups", err.message, "error"); });
}

// Hide all tagged rows (every group found in the last scan)
function hideAllTagged() {
  var allRows = [];
  Object.keys(scannedGroups).forEach(function (tag) {
    scannedGroups[tag].forEach(function (r) {
      if (allRows.indexOf(r) === -1) allRows.push(r);
    });
  });

  if (allRows.length === 0) {
    showStatus("groups", "No tagged rows to hide. Click Refresh first.", "info");
    return;
  }

  Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    allRows.forEach(function (r) {
      sheet.getRange(r + ":" + r).format.rowHidden = true;
    });
    return ctx.sync().then(function () {
      showStatus("groups", "Hidden " + allRows.length + " tagged row(s).", "success");
    });
  }).catch(function (err) { showStatus("groups", err.message, "error"); });
}

/* ═══════════════════════════════════════════════════════════════
   UTILITIES
   ═══════════════════════════════════════════════════════════════ */
function showStatus(feature, message, type) {
  var el = document.getElementById(feature + "-status");
  el.textContent = message;
  el.className = "status " + type;
  el.style.display = "block";
  if (type === "success") {
    setTimeout(function () { el.style.display = "none"; }, 4000);
  }
}

function escapeHtml(s) {
  var d = document.createElement("div");
  d.appendChild(document.createTextNode(s));
  return d.innerHTML;
}

function escapeAttr(s) {
  return s.replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}
