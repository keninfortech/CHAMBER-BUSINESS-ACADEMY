/* global PDFLib, JSZip, XLSX */

var state = {
  templateBytes: null,
  participantsRaw: [],
  headers: [],
  participants: [],
  usedIds: new Set(),
  fontRegularBytes: null,
  fontBoldBytes: null,
};

var el = function(id) { return document.getElementById(id); };
var setStatus = function(msg) { el("genStatus").textContent = msg; };
var setTemplateStatus = function(msg) { el("templateStatus").textContent = msg; };
var setParticipantsStatus = function(msg) { el("participantsStatus").textContent = msg; };
var setMappingStatus = function(msg) { el("mappingStatus").textContent = msg; };
var setFontStatus = function(msg) { el("fontStatus").textContent = msg; };

function normalizeHeader(h) { return (h || "").toString().trim().toLowerCase(); }

function sanitizeCounty(county) {
  var v = (county || "").toString().trim();
  if (!v) return "Unknown_County";
  return v.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_").trim() || "Unknown_County";
}

function sanitizeNameForFile(name) {
  var v = (name || "").toString().trim().toUpperCase();
  return v.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_").slice(0, 50) || "NAME_NOT_PROVIDED";
}

function formatName(name) {
  var v = (name || "").toString().trim();
  return v ? v.toUpperCase() : "NAME NOT PROVIDED";
}

function padSerial(num, digits) {
  var s = String(num);
  return s.length >= digits ? s : "0".repeat(digits - s.length) + s;
}

function clampInt(n, min, max) {
  if (Number.isNaN(n)) return min;
  return Math.max(min, Math.min(max, n));
}

function getSettings() {
  return {
    serialPrefix: el("serialPrefix").value.trim() || "CBA",
    startSerial: parseInt(el("startSerial").value, 10) || 0,
    serialDigits: parseInt(el("serialDigits").value, 10) || 4,
    zipNamePrefix: el("zipNamePrefix").value.trim() || "KNCCI_Certificates",

    // Serial position (right after "No:")
    serialX: parseFloat(el("serialX").value) || 678,
    serialY: parseFloat(el("serialY").value) || 558,
    serialSize: parseFloat(el("serialSize").value) || 14,

    // Name position (centered)
    nameX: parseFloat(el("nameX").value) || 421,
    nameY: parseFloat(el("nameY").value) || 340,
    nameSize: parseFloat(el("nameSize").value) || 20,

    // Course: wipe only "DIGITAL TRADE..." line, keep "successfully participated"
    courseWipeX: parseFloat(el("courseWipeX").value) || 120,
    courseWipeY: parseFloat(el("courseWipeY").value) || 290,
    courseWipeW: parseFloat(el("courseWipeW").value) || 640,
    courseWipeH: parseFloat(el("courseWipeH").value) || 22,
    courseTextY: parseFloat(el("courseTextY").value) || 295,
    courseSize: parseFloat(el("courseSize").value) || 16,

    // Date: wipe "Held on 9th December , 2025"
    dateWipeX: parseFloat(el("dateWipeX").value) || 250,
    dateWipeY: parseFloat(el("dateWipeY").value) || 245,
    dateWipeW: parseFloat(el("dateWipeW").value) || 300,
    dateWipeH: parseFloat(el("dateWipeH").value) || 25,
    dateTextY: parseFloat(el("dateTextY").value) || 252,
    dateSize: parseFloat(el("dateSize").value) || 16,

    // Colors (R,G,B 0-255)
    courseColorR: parseFloat(el("courseColorR").value),
    courseColorG: parseFloat(el("courseColorG").value),
    courseColorB: parseFloat(el("courseColorB").value),
    dateColorR: parseFloat(el("dateColorR").value),
    dateColorG: parseFloat(el("dateColorG").value),
    dateColorB: parseFloat(el("dateColorB").value),
  };
}

async function loadDefaultTemplate() {
  try {
    var res = await fetch("assets/template.pdf");
    if (!res.ok) throw new Error("HTTP " + res.status);
    state.templateBytes = await res.arrayBuffer();
    setTemplateStatus("Template loaded: assets/template.pdf");
  } catch (e) {
    setTemplateStatus("Failed to load default template: " + e.message);
  }
}

async function handleTemplateUpload(file) {
  if (!file) return;
  state.templateBytes = await file.arrayBuffer();
  setTemplateStatus("Template loaded: " + file.name);
}

async function handleFontUpload(which, file) {
  if (!file) return;
  var bytes = await file.arrayBuffer();
  if (which === "regular") state.fontRegularBytes = bytes;
  if (which === "bold") state.fontBoldBytes = bytes;
  var reg = state.fontRegularBytes ? "custom regular" : "default regular";
  var bold = state.fontBoldBytes ? "custom bold" : "default bold";
  setFontStatus("Fonts: " + reg + " + " + bold);
}

function setMappingOptions(headers) {
  var ids = ["mapName", "mapId", "mapCounty", "mapCourse", "mapDate", "mapIssueDate"];
  for (var i = 0; i < ids.length; i++) {
    var s = el(ids[i]);
    s.innerHTML = "";
    var optNone = document.createElement("option");
    optNone.value = "";
    optNone.textContent = "(none)";
    s.appendChild(optNone);
    for (var j = 0; j < headers.length; j++) {
      var opt = document.createElement("option");
      opt.value = headers[j];
      opt.textContent = headers[j];
      s.appendChild(opt);
    }
  }
  autoSelect("mapName", headers, ["participant name", "name", "full name"]);
  autoSelect("mapId", headers, ["national id", "id number", "id", "what is your national id?"]);
  autoSelect("mapCounty", headers, ["business location", "county", "location"]);
  autoSelect("mapCourse", headers, ["course(s)", "course", "courses", "training", "ta needs", "type of ta"]);
  autoSelect("mapDate", headers, ["training date(s)", "training date", "date(s)", "date"]);
  autoSelect("mapIssueDate", headers, ["issue date", "certificate date"]);
}

function autoSelect(selectId, headers, candidates) {
  var s = el(selectId);
  var lowerMap = new Map(headers.map(function(h) { return [normalizeHeader(h), h]; }));
  for (var i = 0; i < candidates.length; i++) {
    var hit = lowerMap.get(candidates[i]);
    if (hit) { s.value = hit; return; }
  }
  var lcHeaders = headers.map(function(h) { return { h: h, lc: normalizeHeader(h) }; });
  for (var i = 0; i < candidates.length; i++) {
    var found = lcHeaders.find(function(x) { return x.lc.includes(candidates[i]); });
    if (found) { s.value = found.h; return; }
  }
}

function splitCsvLine(line) {
  var res = [], cur = "", inQ = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === "," && !inQ) { res.push(cur); cur = ""; }
    else { cur += ch; }
  }
  res.push(cur);
  return res.map(function(x) { return x.trim(); });
}

function parseCsv(text) {
  var lines = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").filter(function(l) { return l.trim().length; });
  if (!lines.length) return [];
  var headers = splitCsvLine(lines[0]);
  var out = [];
  for (var i = 1; i < lines.length; i++) {
    var cols = splitCsvLine(lines[i]);
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = cols[idx] || ""; });
    out.push(obj);
  }
  return out;
}

async function handleParticipantsUpload(file) {
  if (!file) return;
  var lower = file.name.toLowerCase();
  var buf = await file.arrayBuffer();
  var rows = [];
  try {
    if (lower.endsWith(".csv")) {
      rows = parseCsv(new TextDecoder("utf-8").decode(buf));
    } else {
      var wb = XLSX.read(buf, { type: "array" });
      rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
    }
  } catch (e) { setParticipantsStatus("Failed to read file: " + e.message); return; }
  state.participantsRaw = rows;
  state.headers = rows.length ? Object.keys(rows[0]) : [];
  setParticipantsStatus("Loaded " + rows.length + " row(s) from: " + file.name);
  setMappingOptions(state.headers);
  if (!applyMapping(true)) {
    setMappingStatus('Auto-mapping failed. Open "Column mapping" and apply manually.');
  }
}

function applyMapping(isAuto) {
  if (!state.participantsRaw.length) { setStatus("Upload a participant file first."); return false; }
  var map = {
    name: el("mapName").value, id: el("mapId").value, county: el("mapCounty").value,
    course: el("mapCourse").value, date: el("mapDate").value, issueDate: el("mapIssueDate").value,
  };
  if (!map.name || !map.id || !map.county) {
    if (!isAuto) setStatus("Mapping required: Participant Name, National ID, Business Location.");
    return false;
  }
  state.participants = [];
  state.usedIds = new Set();
  var dupes = 0, skipped = 0;
  for (var i = 0; i < state.participantsRaw.length; i++) {
    var r = state.participantsRaw[i];
    var name = (r[map.name] || "").toString().trim();
    var natId = (r[map.id] || "").toString().trim().replace(/\s+/g, "");
    var county = (r[map.county] || "").toString().trim();
    var courses = map.course ? (r[map.course] || "").toString().trim() : "";
    var dates = map.date ? (r[map.date] || "").toString().trim() : "";
    var issueDate = map.issueDate ? (r[map.issueDate] || "").toString().trim() : "";
    if (!name || !natId || !county) { skipped++; continue; }
    if (state.usedIds.has(natId)) { dupes++; continue; }
    state.usedIds.add(natId);
    state.participants.push({ name: name, natId: natId, county: county, courses: courses, dates: dates, issueDate: issueDate });
  }
  persist(); renderTable();
  setMappingStatus("Mapping applied. Participants ready: " + state.participants.length + " (skipped " + skipped + ", duplicates removed " + dupes + ").");
  setStatus("Ready to generate ZIP.");
  return true;
}

function renderTable() {
  var tbody = el("tbl").querySelector("tbody");
  tbody.innerHTML = "";
  var settings = getSettings();
  state.participants.forEach(function(p, idx) {
    var serial = settings.serialPrefix + padSerial(settings.startSerial + idx, settings.serialDigits);
    var tr = document.createElement("tr");
    tr.innerHTML =
      "<td>" + (idx + 1) + "</td><td>" + escapeHtml(serial) + "</td><td>" + escapeHtml(formatName(p.name)) +
      "</td><td>" + escapeHtml(p.natId) + "</td><td>" + escapeHtml(sanitizeCounty(p.county)) +
      "</td><td>" + escapeHtml(p.courses || "") + "</td><td>" + escapeHtml(p.dates || "") +
      "</td><td>" + escapeHtml(p.issueDate || "") + "</td>";
    tbody.appendChild(tr);
  });
}

function escapeHtml(str) {
  return (str || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function persist() {
  try { localStorage.setItem("kncci_participants_fixed", JSON.stringify(state.participants)); } catch (e) {}
}
function restore() {
  try {
    var raw = localStorage.getItem("kncci_participants_fixed");
    if (!raw) return;
    var arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return;
    state.participants = arr;
    state.usedIds = new Set(arr.map(function(x) { return (x.natId || "").toString().replace(/\s+/g, ""); }));
    setMappingStatus("Restored participants from browser storage: " + state.participants.length);
  } catch (e) {}
}

function downloadBlob(blob, filename) {
  var url = URL.createObjectURL(blob);
  var a = document.createElement("a");
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
}

async function downloadCsvTemplate() {
  try {
    var res = await fetch("assets/participants_template.csv");
    if (!res.ok) throw new Error("HTTP " + res.status);
    downloadBlob(await res.blob(), "KNCCI_participants_template.csv");
  } catch (e) { setStatus("Failed to download template: " + e.message); }
}

function exportCsvProcessed() {
  if (!state.participants.length) { setStatus("Nothing to export."); return; }
  var settings = getSettings();
  var headers = ["Serial", "Participant Name", "National ID", "Business Location", "Course(s)", "Training Date(s)", "Issue Date"];
  var rows = state.participants.map(function(p, idx) {
    var serial = settings.serialPrefix + padSerial(settings.startSerial + idx, settings.serialDigits);
    return [serial, formatName(p.name), p.natId, sanitizeCounty(p.county), p.courses || "", p.dates || "", p.issueDate || ""];
  });
  var csv = [headers].concat(rows).map(function(r) {
    return r.map(function(v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(",");
  }).join("\n");
  downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), "participants_processed.csv");
  setStatus("Exported processed CSV.");
}

/* ── Format date: "Held on 30th January , 2026" ── */
function formatDateForCert(dateStr) {
  if (!dateStr) return "";
  var str = dateStr.toString().trim();
  var d = null;
  // Try multiple date formats
  var fmts = [
    function(s) { return new Date(s); },
    function(s) { var p = s.split(/[\/\-]/); return p.length === 3 && parseInt(p[0]) > 12 ? new Date(p[2] + "-" + p[1] + "-" + p[0]) : new Date(p[2] + "-" + p[0] + "-" + p[1]); },
  ];
  for (var i = 0; i < fmts.length; i++) {
    try { d = fmts[i](str); if (!isNaN(d.getTime())) break; else d = null; } catch(e) { d = null; }
  }
  if (!d) return "Held on " + dateStr;
  var day = d.getDate();
  var suffix = "th";
  if (day < 11 || day > 13) {
    var mod = day % 10;
    if (mod === 1) suffix = "st"; else if (mod === 2) suffix = "nd"; else if (mod === 3) suffix = "rd";
  }
  var months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return "Held on " + day + suffix + " " + months[d.getMonth()] + " , " + d.getFullYear();
}

/* ── Build one certificate PDF ── */
async function buildCertificatePdfBytes(templateBytes, participant, serialText, settings) {
  var PDFDocument = PDFLib.PDFDocument;
  var StandardFonts = PDFLib.StandardFonts;
  var rgb = PDFLib.rgb;

  var pdfDoc = await PDFDocument.load(templateBytes);
  var page = pdfDoc.getPage(0);
  var pageWidth = page.getWidth();

  // Fonts
  var fontRegular, fontBold;
  if (state.fontRegularBytes) fontRegular = await pdfDoc.embedFont(state.fontRegularBytes);
  else fontRegular = await pdfDoc.embedFont(StandardFonts.TimesRoman);
  if (state.fontBoldBytes) fontBold = await pdfDoc.embedFont(state.fontBoldBytes);
  else fontBold = await pdfDoc.embedFont(StandardFonts.TimesRomanBold);

  // Colors
  var colorBlack = rgb(0, 0, 0);
  var colorCourse = rgb(settings.courseColorR / 255, settings.courseColorG / 255, settings.courseColorB / 255);
  var colorDate = rgb(settings.dateColorR / 255, settings.dateColorG / 255, settings.dateColorB / 255);
  var colorWhite = rgb(1, 1, 1);

  // ── Serial (black) ──
  page.drawText(serialText, {
    x: settings.serialX, y: settings.serialY,
    size: settings.serialSize, font: fontRegular, color: colorBlack,
  });

  // ── Name (black, centered, bold) ──
  var nameText = formatName(participant.name);
  var nw = fontBold.widthOfTextAtSize(nameText, settings.nameSize);
  page.drawText(nameText, {
    x: settings.nameX - nw / 2, y: settings.nameY,
    size: settings.nameSize, font: fontBold, color: colorBlack,
  });

  // ── Course: wipe ONLY "DIGITAL TRADE..." line, keep "successfully participated" ──
  if (participant.courses) {
    // White rectangle over course line
    page.drawRectangle({
      x: settings.courseWipeX, y: settings.courseWipeY,
      width: settings.courseWipeW, height: settings.courseWipeH,
      color: colorWhite, borderWidth: 0,
    });
    // Draw course text centered (golden color)
    var courseText = participant.courses.toUpperCase();
    var cw = fontRegular.widthOfTextAtSize(courseText, settings.courseSize);
    var cx = (pageWidth / 2) - (cw / 2);
    page.drawText(courseText, {
      x: cx, y: settings.courseTextY,
      size: settings.courseSize, font: fontRegular, color: colorCourse,
    });
  }

  // ── Date: wipe "Held on 9th December , 2025", replace with Excel date ──
  if (participant.dates) {
    // White rectangle over date line
    page.drawRectangle({
      x: settings.dateWipeX, y: settings.dateWipeY,
      width: settings.dateWipeW, height: settings.dateWipeH,
      color: colorWhite, borderWidth: 0,
    });
    // Draw formatted date centered (near-black, matching template)
    var formattedDate = formatDateForCert(participant.dates);
    var dw = fontRegular.widthOfTextAtSize(formattedDate, settings.dateSize);
    var dx = (pageWidth / 2) - (dw / 2);
    page.drawText(formattedDate, {
      x: dx, y: settings.dateTextY,
      size: settings.dateSize, font: fontRegular, color: colorDate,
    });
  }

  return await pdfDoc.save();
}

/* ── Generate ZIP ── */
async function generateZip() {
  if (!state.templateBytes) { setStatus("Load a template PDF first (Step 1)."); return; }
  if (!state.participants.length) { setStatus("Upload participants (Step 2)."); return; }
  var settings = getSettings();
  setStatus("Generating " + state.participants.length + " certificates...");
  var zip = new JSZip();
  for (var i = 0; i < state.participants.length; i++) {
    var p = state.participants[i];
    var serial = settings.serialPrefix + padSerial(settings.startSerial + i, settings.serialDigits);
    var folder = sanitizeCounty(p.county);
    var filename = serial + "_" + sanitizeNameForFile(p.name) + ".pdf";
    try {
      var pdfBytes = await buildCertificatePdfBytes(state.templateBytes, p, serial, settings);
      zip.folder(folder).file(filename, pdfBytes);
      if ((i + 1) % 50 === 0) setStatus("Progress: " + (i + 1) + "/" + state.participants.length);
    } catch (e) {
      console.error(e);
      zip.file("ERROR_" + serial + ".txt", String(e));
    }
  }
  setStatus("Packaging ZIP...");
  var zipBlob = await zip.generateAsync({ type: "blob" });
  var ts = new Date();
  var stamp = ts.getFullYear() + String(ts.getMonth() + 1).padStart(2, "0") + String(ts.getDate()).padStart(2, "0") + "_" + String(ts.getHours()).padStart(2, "0") + String(ts.getMinutes()).padStart(2, "0");
  downloadBlob(zipBlob, settings.zipNamePrefix + "_" + stamp + ".zip");
  setStatus("Done! ZIP downloaded.");
}

function resetLoadedList() {
  state.participantsRaw = []; state.headers = []; state.participants = []; state.usedIds = new Set();
  persist(); renderTable();
  setParticipantsStatus("No participant file loaded.");
  setMappingStatus("Mapping: not applied yet.");
  setStatus("Reset complete.");
}

function wireEvents() {
  el("btnLoadDefault").addEventListener("click", loadDefaultTemplate);
  el("templateFile").addEventListener("change", function(e) { handleTemplateUpload(e.target.files[0]); });
  el("participantsFile").addEventListener("change", function(e) { handleParticipantsUpload(e.target.files[0]); });
  el("btnApplyMapping").addEventListener("click", function() { applyMapping(false); });
  el("btnDownloadCsvTemplate").addEventListener("click", downloadCsvTemplate);
  el("fontRegularFile").addEventListener("change", function(e) { handleFontUpload("regular", e.target.files[0]); });
  el("fontBoldFile").addEventListener("change", function(e) { handleFontUpload("bold", e.target.files[0]); });
  el("btnGenerateZip").addEventListener("click", generateZip);
  el("btnExportCsv").addEventListener("click", exportCsvProcessed);
  el("btnResetList").addEventListener("click", resetLoadedList);
  ["serialPrefix", "startSerial", "serialDigits"].forEach(function(id) { el(id).addEventListener("input", renderTable); });
}

document.addEventListener("DOMContentLoaded", function() { restore(); wireEvents(); renderTable(); });
