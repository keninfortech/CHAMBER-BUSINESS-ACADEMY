/* global PDFLib, JSZip, XLSX */

const state = {
  templateBytes: null,
  participantsRaw: [],
  headers: [],
  participants: [],
  usedIds: new Set(),
  fontRegularBytes: null,
  fontBoldBytes: null,
};

const el = (id) => document.getElementById(id);
const setStatus = (msg) => (el("genStatus").textContent = msg);
const setTemplateStatus = (msg) => (el("templateStatus").textContent = msg);
const setParticipantsStatus = (msg) => (el("participantsStatus").textContent = msg);
const setMappingStatus = (msg) => (el("mappingStatus").textContent = msg);
const setFontStatus = (msg) => (el("fontStatus").textContent = msg);

function normalizeHeader(h) {
  return (h || "").toString().trim().toLowerCase();
}

function sanitizeCounty(county) {
  var v = (county || "").toString().trim();
  if (!v) return "Unknown_County";
  return v.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_").trim() || "Unknown_County";
}

function sanitizeNameForFile(name) {
  var v = (name || "").toString().trim().toUpperCase();
  var safe = v.replace(/[^\w\s-]/g, "").replace(/\s+/g, "_").slice(0, 50);
  return safe || "NAME_NOT_PROVIDED";
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

function rgbFromInputs() {
  var r = clampInt(parseInt(el("colorR").value, 10), 0, 255);
  var g = clampInt(parseInt(el("colorG").value, 10), 0, 255);
  var b = clampInt(parseInt(el("colorB").value, 10), 0, 255);
  return { r: r / 255, g: g / 255, b: b / 255 };
}

function getSettings() {
  return {
    serialPrefix: el("serialPrefix").value.trim() || "CBA",
    startSerial: parseInt(el("startSerial").value, 10) || 0,
    serialDigits: parseInt(el("serialDigits").value, 10) || 4,
    zipNamePrefix: el("zipNamePrefix").value.trim() || "KNCCI_Certificates",

    serialX: parseFloat(el("serialX").value) || 708,
    serialY: parseFloat(el("serialY").value) || 558,
    nameX: parseFloat(el("nameX").value) || 421,
    nameY: parseFloat(el("nameY").value) || 340,
    nameSize: parseFloat(el("nameSize").value) || 20,
    serialSize: parseFloat(el("serialSize").value) || 14,

    // Course wipe: covers "successfully participated..." + "DIGITAL TRADE..." (2 lines)
    courseWipeX: parseFloat(el("courseX").value) || 100,
    courseWipeY: parseFloat(el("courseWipeY").value) || 290,
    courseWipeW: parseFloat(el("courseW").value) || 700,
    courseWipeH: parseFloat(el("courseWipeH").value) || 45,
    // Course text: centered on the original course line
    courseTextY: parseFloat(el("courseY").value) || 310,
    courseSize: parseFloat(el("courseSize").value) || 16,
    courseAlign: el("courseAlign").value || "center",

    // Date wipe: covers "Held on 9th December , 2025"
    dateWipeX: parseFloat(el("dateX").value) || 250,
    dateWipeY: parseFloat(el("dateWipeY").value) || 246,
    dateWipeW: parseFloat(el("dateW").value) || 350,
    dateWipeH: parseFloat(el("dateWipeH").value) || 22,
    // Date text
    dateTextY: parseFloat(el("dateY").value) || 264,
    dateSize: parseFloat(el("dateSize").value) || 16,
    dateAlign: el("dateAlign").value || "center",

    textColor: rgbFromInputs(),
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
  var selects = ["mapName", "mapId", "mapCounty", "mapCourse", "mapDate", "mapIssueDate"].map(el);
  for (var i = 0; i < selects.length; i++) {
    var s = selects[i];
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
  var res = [];
  var cur = "";
  var inQ = false;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (ch === '"') {
      if (inQ && line[i + 1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if (ch === "," && !inQ) {
      res.push(cur); cur = "";
    } else {
      cur += ch;
    }
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
      var text = new TextDecoder("utf-8").decode(buf);
      rows = parseCsv(text);
    } else {
      var wb = XLSX.read(buf, { type: "array" });
      var ws = wb.Sheets[wb.SheetNames[0]];
      rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    }
  } catch (e) {
    setParticipantsStatus("Failed to read file: " + e.message);
    return;
  }
  state.participantsRaw = rows;
  state.headers = rows.length ? Object.keys(rows[0]) : [];
  setParticipantsStatus("Loaded " + rows.length + " row(s) from: " + file.name);
  setMappingOptions(state.headers);
  var ok = applyMapping(true);
  if (!ok) {
    setMappingStatus('Auto-mapping failed. Open "Column mapping" and apply manually.');
  }
}

function applyMapping(isAuto) {
  if (!state.participantsRaw.length) {
    setStatus("Upload a participant file first.");
    return false;
  }
  var map = {
    name: el("mapName").value,
    id: el("mapId").value,
    county: el("mapCounty").value,
    course: el("mapCourse").value,
    date: el("mapDate").value,
    issueDate: el("mapIssueDate").value,
  };
  if (!map.name || !map.id || !map.county) {
    if (!isAuto) setStatus("Mapping required: Participant Name, National ID, Business Location.");
    return false;
  }
  state.participants = [];
  state.usedIds = new Set();
  var dupes = 0;
  var skipped = 0;
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
  persist();
  renderTable();
  setMappingStatus("Mapping applied. Participants ready: " + state.participants.length + " (skipped " + skipped + ", duplicates removed " + dupes + ").");
  setStatus("Ready to generate ZIP.");
  return true;
}

function renderTable() {
  var tbody = el("tbl").querySelector("tbody");
  tbody.innerHTML = "";
  var settings = getSettings();
  state.participants.forEach(function(p, idx) {
    var serialNum = settings.startSerial + idx;
    var serial = settings.serialPrefix + padSerial(serialNum, settings.serialDigits);
    var tr = document.createElement("tr");
    tr.innerHTML =
      "<td>" + (idx + 1) + "</td>" +
      "<td>" + escapeHtml(serial) + "</td>" +
      "<td>" + escapeHtml(formatName(p.name)) + "</td>" +
      "<td>" + escapeHtml(p.natId) + "</td>" +
      "<td>" + escapeHtml(sanitizeCounty(p.county)) + "</td>" +
      "<td>" + escapeHtml(p.courses || "") + "</td>" +
      "<td>" + escapeHtml(p.dates || "") + "</td>" +
      "<td>" + escapeHtml(p.issueDate || "") + "</td>";
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
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(function() { URL.revokeObjectURL(url); }, 2000);
}

async function downloadCsvTemplate() {
  try {
    var res = await fetch("assets/participants_template.csv");
    if (!res.ok) throw new Error("HTTP " + res.status);
    var blob = await res.blob();
    downloadBlob(blob, "KNCCI_participants_template.csv");
  } catch (e) {
    setStatus("Failed to download template: " + e.message);
  }
}

function exportCsvProcessed() {
  if (!state.participants.length) { setStatus("Nothing to export."); return; }
  var settings = getSettings();
  var headers = ["Serial", "Participant Name", "National ID", "Business Location", "Course(s)", "Training Date(s)", "Issue Date"];
  var rows = state.participants.map(function(p, idx) {
    var serialNum = settings.startSerial + idx;
    var serial = settings.serialPrefix + padSerial(serialNum, settings.serialDigits);
    return [serial, formatName(p.name), p.natId, sanitizeCounty(p.county), p.courses || "", p.dates || "", p.issueDate || ""];
  });
  var csv = [headers].concat(rows).map(function(r) {
    return r.map(function(v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(",");
  }).join("\n");
  downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8" }), "participants_processed.csv");
  setStatus("Exported processed CSV.");
}

/* ── date formatter: "Held on 30th January, 2026" ── */
function formatDateForCert(dateStr) {
  if (!dateStr) return "";
  var str = dateStr.toString().trim();
  var d;
  try {
    d = new Date(str);
    if (isNaN(d.getTime())) {
      var parts = str.split(/[\/\-]/);
      if (parts.length === 3) {
        if (parseInt(parts[0]) > 12) d = new Date(parts[2] + "-" + parts[1] + "-" + parts[0]);
        else d = new Date(parts[2] + "-" + parts[0] + "-" + parts[1]);
      }
    }
    if (isNaN(d.getTime())) return "Held on " + dateStr;
  } catch (e) {
    return "Held on " + dateStr;
  }
  var day = d.getDate();
  var suffix = "th";
  if (day < 11 || day > 13) {
    var mod = day % 10;
    if (mod === 1) suffix = "st";
    else if (mod === 2) suffix = "nd";
    else if (mod === 3) suffix = "rd";
  }
  var months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  return "Held on " + day + suffix + " " + months[d.getMonth()] + ", " + d.getFullYear();
}

/* ── wipe rect + draw text ── */
function wipeRect(page, x, y, w, h) {
  var rgb = PDFLib.rgb;
  page.drawRectangle({
    x: x, y: y, width: w, height: h,
    color: rgb(1, 1, 1), borderColor: rgb(1, 1, 1), borderWidth: 0,
  });
}

function drawCentered(page, text, areaX, areaW, y, font, size, color) {
  var tw = font.widthOfTextAtSize(text, size);
  var x = areaX + (areaW / 2) - (tw / 2);
  page.drawText(text, { x: x, y: y, size: size, font: font, color: color });
}

/* ── build one certificate ── */
async function buildCertificatePdfBytes(templateBytes, participant, serialText, settings) {
  var PDFDocument = PDFLib.PDFDocument;
  var StandardFonts = PDFLib.StandardFonts;
  var rgb = PDFLib.rgb;

  var pdfDoc = await PDFDocument.load(templateBytes);
  var page = pdfDoc.getPage(0);

  // Fonts: use custom if uploaded, else Helvetica fallback
  var fontRegular, fontBold;
  if (state.fontRegularBytes) fontRegular = await pdfDoc.embedFont(state.fontRegularBytes);
  else fontRegular = await pdfDoc.embedFont(StandardFonts.TimesRoman);
  if (state.fontBoldBytes) fontBold = await pdfDoc.embedFont(state.fontBoldBytes);
  else fontBold = await pdfDoc.embedFont(StandardFonts.TimesRomanBold);

  var color = rgb(settings.textColor.r, settings.textColor.g, settings.textColor.b);

  // ── Serial number ──
  page.drawText(serialText, {
    x: settings.serialX, y: settings.serialY,
    size: settings.serialSize, font: fontRegular, color: color,
  });

  // ── Name (centered, bold) ──
  var nameText = formatName(participant.name);
  var nw = fontBold.widthOfTextAtSize(nameText, settings.nameSize);
  var nameDrawX = settings.nameX - (nw / 2);
  page.drawText(nameText, {
    x: nameDrawX, y: settings.nameY,
    size: settings.nameSize, font: fontBold, color: color,
  });

  // ── Course: wipe 2 lines ("successfully participated..." + "DIGITAL TRADE...") then draw course from Excel ──
  if (participant.courses) {
    wipeRect(page, settings.courseWipeX, settings.courseWipeY, settings.courseWipeW, settings.courseWipeH);
    var courseText = participant.courses.toUpperCase();
    drawCentered(page, courseText, settings.courseWipeX, settings.courseWipeW, settings.courseTextY, fontRegular, settings.courseSize, color);
  }

  // ── Date: wipe original "Held on 9th December, 2025" then draw new date from Excel ──
  if (participant.dates) {
    wipeRect(page, settings.dateWipeX, settings.dateWipeY, settings.dateWipeW, settings.dateWipeH);
    var formattedDate = formatDateForCert(participant.dates);
    drawCentered(page, formattedDate, settings.dateWipeX, settings.dateWipeW, settings.dateTextY, fontRegular, settings.dateSize, color);
  }

  return await pdfDoc.save();
}

/* ── generate ZIP ── */
async function generateZip() {
  if (!state.templateBytes) { setStatus("Load a template PDF first (Step 1)."); return; }
  if (!state.participants.length) { setStatus("Upload participants (Step 2)."); return; }
  var settings = getSettings();
  setStatus("Generating " + state.participants.length + " certificates...");
  var zip = new JSZip();
  for (var i = 0; i < state.participants.length; i++) {
    var p = state.participants[i];
    var serialNum = settings.startSerial + i;
    var serial = settings.serialPrefix + padSerial(serialNum, settings.serialDigits);
    var countyFolder = sanitizeCounty(p.county);
    var safeName = sanitizeNameForFile(p.name);
    var filename = serial + "_" + safeName + ".pdf";
    try {
      var pdfBytes = await buildCertificatePdfBytes(state.templateBytes, p, serial, settings);
      zip.folder(countyFolder).file(filename, pdfBytes);
      if ((i + 1) % 50 === 0) setStatus("Progress: " + (i + 1) + "/" + state.participants.length);
    } catch (e) {
      console.error(e);
      zip.file("ERROR_" + serial + "_" + sanitizeNameForFile(p.name) + ".txt", String(e));
    }
  }
  setStatus("Packaging ZIP...");
  var zipBlob = await zip.generateAsync({ type: "blob" });
  var ts = new Date();
  var stamp = ts.getFullYear() + String(ts.getMonth() + 1).padStart(2, "0") + String(ts.getDate()).padStart(2, "0") + "_" + String(ts.getHours()).padStart(2, "0") + String(ts.getMinutes()).padStart(2, "0");
  var outName = settings.zipNamePrefix + "_" + stamp + ".zip";
  downloadBlob(zipBlob, outName);
  setStatus("Done. Downloaded: " + outName);
}

function resetLoadedList() {
  state.participantsRaw = [];
  state.headers = [];
  state.participants = [];
  state.usedIds = new Set();
  persist();
  renderTable();
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

document.addEventListener("DOMContentLoaded", function() {
  restore();
  wireEvents();
  renderTable();
});
