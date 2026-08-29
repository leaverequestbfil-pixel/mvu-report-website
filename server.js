const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = Number(process.env.PORT || 5173);

app.use(express.json({ limit: "2mb" }));
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, "public")));

const ROOT = __dirname;
const DATA_DIR = path.join(ROOT, "data");
const UPLOAD_DIR = path.join(ROOT, "uploads");
const STORE_FILE = path.join(DATA_DIR, "mvu-data.json");

fs.mkdirSync(DATA_DIR, { recursive: true });
fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const upload = multer({
  dest: UPLOAD_DIR,
  limits: { fileSize: 50 * 1024 * 1024 }
});

function emptyStore() {
  return {
    villageMapping: {},
    weekOff: {},
    generatedReports: [],
    uploadLog: []
  };
}

function loadStore() {
  if (!fs.existsSync(STORE_FILE)) return emptyStore();
  try {
    const data = JSON.parse(fs.readFileSync(STORE_FILE, "utf8"));
    return {
      ...emptyStore(),
      ...data,
      villageMapping: data.villageMapping || {},
      weekOff: data.weekOff || {},
      generatedReports: data.generatedReports || [],
      uploadLog: data.uploadLog || []
    };
  } catch {
    return emptyStore();
  }
}

let store = loadStore();

function saveStore() {
  const tmp = STORE_FILE + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(store, null, 2), "utf8");
  fs.renameSync(tmp, STORE_FILE);
}

function now() {
  return new Date().toISOString();
}

function clean(v) {
  if (v === undefined || v === null) return "";
  return String(v).trim();
}

function norm(v) {
  return clean(v).toUpperCase();
}

function excelDateToDate(v) {
  if (v instanceof Date && !isNaN(v)) return v;

  if (typeof v === "number") {
    const parsed = XLSX.SSF.parse_date_code(v);
    if (parsed) {
      return new Date(
        parsed.y,
        parsed.m - 1,
        parsed.d,
        parsed.H || 0,
        parsed.M || 0,
        parsed.S || 0
      );
    }
  }

  const s = clean(v);
  if (!s) return null;

  const d = new Date(s.replace(" ", "T"));
  if (!isNaN(d)) return d;

  const m = s.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})/);
  if (m) {
    return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
  }

  return null;
}

function dateKey(v) {
  const d = excelDateToDate(v);
  if (!d) return "";

  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");

  return `${y}-${m}-${day}`;
}

function displayDate(key) {
  if (!key) return "";
  const [y, m, d] = key.split("-");
  return `${d}-${m}-${y}`;
}

function weekdayName(key) {
  const [y, m, d] = key.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString("en-US", {
    weekday: "long"
  });
}

function readSheetRows(filePath, sheetName) {
  const wb = XLSX.readFile(filePath, { cellDates: true });
  const ws = wb.Sheets[sheetName || wb.SheetNames[0]];
  if (!ws) throw new Error(`Sheet not found: ${sheetName}`);
  return XLSX.utils.sheet_to_json(ws, { defval: "" });
}

function logUpload(kind, filename, rowCount, status, message = "") {
  store.uploadLog.unshift({
    kind,
    filename,
    row_count: rowCount,
    status,
    message,
    uploaded_at: now()
  });

  store.uploadLog = store.uploadLog.slice(0, 100);
  saveStore();
}

function saveMapping(filePath, originalName) {
  const rows = readSheetRows(filePath, "UpdateVillageMapping");

  if (!rows.length) {
    throw new Error("Village mapping sheet is empty.");
  }

  for (const c of ["PanchayatID", "MVUNumber"]) {
    if (!(c in rows[0])) {
      throw new Error(`Missing column: ${c}`);
    }
  }

  let count = 0;

  for (const r of rows) {
    const p = clean(r.PanchayatID);
    const v = clean(r.MVUNumber);

    if (!p || !v) continue;

    store.villageMapping[p] = {
      panchayat_id: p,
      mvu_number: v,
      district_name: clean(r.DistrictName),
      block_name: clean(r.BlockName),
      panchayat_name: clean(r.PanchayatName),
      updated_at: now()
    };

    count++;
  }

  saveStore();

  logUpload(
    "mapping",
    originalName,
    rows.length,
    "success",
    "Mapping saved/updated permanently."
  );

  return count;
}

function saveWeekOff(filePath, originalName) {
  const wb = XLSX.readFile(filePath, { cellDates: true });

  const sheetName = wb.SheetNames.includes("Daily Data Entry")
    ? "Daily Data Entry"
    : wb.SheetNames[0];

  const ws = wb.Sheets[sheetName];

  if (!ws) throw new Error("Week Off sheet not found.");

  const raw = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    defval: ""
  });

  if (!raw.length) {
    throw new Error("Week Off file is empty.");
  }

  let headerRow = -1;

  for (let i = 0; i < Math.min(raw.length, 15); i++) {
    const h = (raw[i] || []).map(x => norm(x));

    const hasVehicle = h.some(x =>
      [
        "VEHICLE NUMBER",
        "VEHICLE NO.",
        "VEHICLE NO",
        "MVU/GADI NUMBER",
        "MVUNUMBER"
      ].includes(x)
    );

    const hasWeekOff = h.some(x =>
      ["WEEK OFF", "WEEKOFF"].includes(x)
    );

    if (hasVehicle && hasWeekOff) {
      headerRow = i;
      break;
    }
  }

  if (headerRow < 0) {
    throw new Error(
      "Week Off file must contain Vehicle Number and Week Off columns."
    );
  }

  const header = raw[headerRow] || [];

  const findCol = names =>
    header.findIndex(x => names.includes(norm(x)));

  const idx = {
    block: findCol(["BLOCK", "BLOCK / VEHICLE"]),
    district: findCol(["DISTRICT", "DISTRICT NAME"]),
    vehicle: findCol([
      "VEHICLE NUMBER",
      "VEHICLE NO.",
      "VEHICLE NO",
      "MVU/GADI NUMBER",
      "MVUNUMBER"
    ]),
    weekoff: findCol(["WEEK OFF", "WEEKOFF"])
  };

  if (idx.vehicle < 0 || idx.weekoff < 0) {
    throw new Error(
      "Week Off file must contain Vehicle Number and Week Off columns."
    );
  }

  const mappingByVehicle = new Map(
    Object.values(store.villageMapping).map(r => [
      norm(r.mvu_number),
      r
    ])
  );

  let district = "";
  const items = [];

  for (let i = headerRow + 1; i < raw.length; i++) {
    const r = raw[i] || [];

    if (idx.district >= 0 && clean(r[idx.district])) {
      district = clean(r[idx.district]);
    }

    const vehicle = clean(r[idx.vehicle]);
    if (!vehicle) continue;

    let block = idx.block >= 0 ? clean(r[idx.block]) : "";
    let rowDistrict =
      idx.district >= 0
        ? clean(r[idx.district])
        : district;

    const weekoff = clean(r[idx.weekoff]);
    if (!weekoff) continue;

    const m = mappingByVehicle.get(norm(vehicle));

    if (m) {
      if (!rowDistrict) rowDistrict = clean(m.district_name);
      if (!block) block = clean(m.block_name);
    }

    items.push({
      vehicle,
      district: rowDistrict,
      block,
      weekoff
    });
  }

  if (!items.length) {
    throw new Error("No valid Week Off records found.");
  }

  for (const x of items) {
    store.weekOff[x.vehicle] = {
      vehicle_no: x.vehicle,
      district: x.district,
      block: x.block,
      week_off: x.weekoff,
      updated_at: now()
    };
  }

  saveStore();

  logUpload(
    "weekoff",
    originalName,
    items.length,
    "success",
    "Week Off master saved/updated permanently."
  );

  return items.length;
}

function ticketAudio(row) {
  const ticket = norm(row.TicketStatus);

  if (
    ticket === "OPEN" ||
    ticket === "APPOINTED" ||
    ticket === "REASSIGN"
  ) {
    return "Pending";
  }

  return norm(row.SubStatus) === "VISITED FARMER"
    ? "Attend"
    : "Not Attend";
}

function processDetailed(filePath, originalName) {
  const rows = readSheetRows(filePath, "DetailedReport");

  if (!rows.length) {
    throw new Error("DetailedReport sheet is empty.");
  }

  const required = [
    "Panchayat ID",
    "CreatedDateTime",
    "LevelType",
    "Type",
    "CloseRemarks",
    "TicketStatus",
    "SubStatus"
  ];

  for (const c of required) {
    if (!(c in rows[0])) {
      throw new Error(`Detailed Report missing column: ${c}`);
    }
  }

  const filtered = rows.filter(r => {
    if (norm(r.LevelType) === "TA") return false;
    if (norm(r.Type) === "ENQUIRY") return false;
    if (norm(r.CloseRemarks).includes("WT")) return false;
    return true;
  });

  const mapping = new Map(
    Object.values(store.villageMapping).map(r => [
      clean(r.panchayat_id),
      r
    ])
  );

  const records = [];
  const unmatched = [];
  const dates = new Set();

  for (const r of filtered) {
    const d = dateKey(r.CreatedDateTime);
    if (!d) continue;

    dates.add(d);

    const p = clean(r["Panchayat ID"]);
    const m = mapping.get(p);

    if (!m) {
      unmatched.push({
        panchayatId: p,
        ticketId: clean(r.TicketID),
        date: d
      });
      continue;
    }

    records.push({
      date: d,
      vehicle: clean(m.mvu_number),
      district: clean(m.district_name) || clean(r.DISTRICT),
      block: clean(m.block_name) || clean(r.BLOCK),
      audio: ticketAudio(r)
    });
  }

  const sortedDates = [...dates].sort((a, b) => a.localeCompare(b));

  if (!sortedDates.length) {
    throw new Error(
      "No valid CreatedDateTime dates found after filtering."
    );
  }

  if (sortedDates.length > 2) {
    throw new Error(
      `Uploaded Detailed Report contains ${sortedDates.length} dates. The report format supports exactly two dates.`
    );
  }

  const firstDate = sortedDates[0];
  const secondDate = sortedDates.length > 1
    ? sortedDates[1]
    : "";

  /*
   * IMPORTANT:
   * Dist. Head Quarter is intentionally NOT removed.
   */
  const roster = Object.values(store.weekOff)
    .map(r => ({
      vehicle: r.vehicle_no,
      district: r.district,
      block: r.block,
      week_off: r.week_off
    }))
    .sort((a, b) =>
      `${a.district}|${a.block}|${a.vehicle}`.localeCompare(
        `${b.district}|${b.block}|${b.vehicle}`
      )
    );

  if (!roster.length) {
    throw new Error(
      "Week Off/Template master is empty. Upload it before generating the report."
    );
  }

  const agg = new Map();

  for (const r of records) {
    const key = `${norm(r.vehicle)}|${r.date}`;

    if (!agg.has(key)) {
      agg.set(key, {
        Attend: 0,
        "Not Attend": 0,
        Pending: 0
      });
    }

    agg.get(key)[r.audio]++;
  }

  function counts(vehicle, date) {
    const c = date
      ? (
          agg.get(`${norm(vehicle)}|${date}`) || {
            Attend: 0,
            "Not Attend": 0,
            Pending: 0
          }
        )
      : {
          Attend: 0,
          "Not Attend": 0,
          Pending: 0
        };

    return {
      received: c.Attend + c["Not Attend"] + c.Pending,
      attend: c.Attend,
      notAttend: c["Not Attend"],
      pending: c.Pending
    };
  }

  const rowsOut = roster.map(x => {
    const y = counts(x.vehicle, firstDate);
    const t = counts(x.vehicle, secondDate);

    const yesterdayRemark =
      x.week_off &&
      norm(x.week_off) === norm(weekdayName(firstDate))
        ? "Week off"
        : "";

    const todayRemark =
      secondDate &&
      x.week_off &&
      norm(x.week_off) === norm(weekdayName(secondDate))
        ? "Week off"
        : "";

    return {
      district: x.district || "",
      block: x.block || "",
      vehicle: x.vehicle,
      weekOff: x.week_off || "",

      yesterday: {
        received: y.received,
        attend: y.attend,
        remark: yesterdayRemark
      },

      today: {
        received: t.received,
        attended: t.attend,
        notAttend: t.notAttend,
        pending: t.pending,
        remark: todayRemark
      }
    };
  });

  const byDistrict = new Map();

  for (const r of rowsOut) {
    if (!byDistrict.has(r.district)) {
      byDistrict.set(r.district, []);
    }

    byDistrict.get(r.district).push(r);
  }

  function total(rows) {
    const out = {
      vehicles: rows.length,
      yesterdayReceived: 0,
      yesterdayAttend: 0,
      todayReceived: 0,
      attended: 0,
      notAttend: 0,
      pending: 0
    };

    for (const r of rows) {
      out.yesterdayReceived += Number(r.yesterday.received || 0);
      out.yesterdayAttend += Number(r.yesterday.attend || 0);

      out.todayReceived += Number(r.today.received || 0);
      out.attended += Number(r.today.attended || 0);
      out.notAttend += Number(r.today.notAttend || 0);
      out.pending += Number(r.today.pending || 0);
    }

    out.attendPct = out.todayReceived
      ? +(out.attended / out.todayReceived * 100).toFixed(2)
      : 0;

    /*
     * New top calculations:
     * Cases Received = both dates' received total
     * Attended Cases = both dates' attended total
     * Pending Cases = TODAY pending
     */
    out.casesReceived =
      out.yesterdayReceived + out.todayReceived;

    out.attendedCases =
      out.yesterdayAttend + out.attended;

    out.todayPending = out.pending;

    return out;
  }

  const districts = [...byDistrict.entries()]
    .map(([district, items]) => ({
      district,
      rows: items.sort((a, b) => {
        const aa =
          Number(a.yesterday.attend || 0) +
          Number(a.today.attended || 0);

        const bb =
          Number(b.yesterday.attend || 0) +
          Number(b.today.attended || 0);

        if (bb !== aa) return bb - aa;

        return clean(a.block).localeCompare(
          clean(b.block)
        );
      }),
      total: total(items)
    }))
    .sort((a, b) =>
      clean(a.district).localeCompare(clean(b.district))
    );

  const overall = total(rowsOut);

  const report = {
    generatedAt: now(),
    sourceFile: originalName,
    firstDate,
    secondDate,
    firstDateDisplay: displayDate(firstDate),
    secondDateDisplay: displayDate(secondDate),
    districts,
    overall,

    validation: {
      sourceRows: rows.length,
      rowsAfterFilter: filtered.length,
      unmatchedPanchayatRows: unmatched.length,
      datesFound: sortedDates
    }
  };

  store.generatedReports.unshift({
    generated_at: report.generatedAt,
    first_date: firstDate,
    second_date: secondDate,
    source_file: originalName,
    report_json: report
  });

  store.generatedReports =
    store.generatedReports.slice(0, 30);

  saveStore();

  logUpload(
    "detailed",
    originalName,
    rows.length,
    "success",
    `Generated report for ${displayDate(firstDate)}${
      secondDate
        ? " and " + displayDate(secondDate)
        : ""
    }. Filtered ${rows.length - filtered.length} rows.`
  );

  return report;
}

/* =========================
   API ROUTES
========================= */

app.get("/api/status", (req, res) => {
  const latest = store.generatedReports[0];

  const mappingCount =
    Object.keys(store.villageMapping).length;

  const weekOffCount =
    Object.keys(store.weekOff).length;

  res.json({
    ok: true,
    mappingCount,
    weekOffCount,
    mastersLocked:
      mappingCount > 0 && weekOffCount > 0,
    latest: latest ? latest.report_json : null,
    logs: store.uploadLog.slice(0, 8)
  });
});

app.post("/api/upload/mapping", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      throw new Error("Mapping file is required.");
    }

    if (Object.keys(store.villageMapping).length > 0) {
      throw new Error(
        "Village Mapping is locked. Use Hard Reset first."
      );
    }

    const count =
      saveMapping(
        req.file.path,
        req.file.originalname
      );

    if (fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    res.json({
      ok: true,
      message:
        `Village Mapping saved. ${count} source rows processed.`
    });
  } catch (e) {
    if (req.file?.path && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    logUpload(
      "mapping",
      req.file?.originalname || "",
      0,
      "error",
      e.message
    );

    res.status(400).json({
      ok: false,
      error: e.message
    });
  }
});

app.post("/api/upload/weekoff", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      throw new Error("Template file is required.");
    }

    if (Object.keys(store.weekOff).length > 0) {
      throw new Error(
        "Week Off Master is locked. Use Hard Reset first."
      );
    }

    const count =
      saveWeekOff(
        req.file.path,
        req.file.originalname
      );

    if (fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    res.json({
      ok: true,
      message:
        `Week Off master saved. ${count} vehicles processed.`
    });
  } catch (e) {
    if (req.file?.path && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    logUpload(
      "weekoff",
      req.file?.originalname || "",
      0,
      "error",
      e.message
    );

    res.status(400).json({
      ok: false,
      error: e.message
    });
  }
});

app.post("/api/generate", upload.single("file"), (req, res) => {
  try {
    if (!req.file) {
      throw new Error(
        "Detailed Report file is required."
      );
    }

    if (
      Object.keys(store.villageMapping).length === 0 ||
      Object.keys(store.weekOff).length === 0
    ) {
      throw new Error(
        "Please upload Village Mapping and Week Off master first."
      );
    }

    const report =
      processDetailed(
        req.file.path,
        req.file.originalname
      );

    if (fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    res.json({
      ok: true,
      report
    });
  } catch (e) {
    if (req.file?.path && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }

    logUpload(
      "detailed",
      req.file?.originalname || "",
      0,
      "error",
      e.message
    );

    res.status(400).json({
      ok: false,
      error: e.message
    });
  }
});

app.get("/api/report", (req, res) => {
  const latest = store.generatedReports[0];

  if (!latest) {
    return res.status(404).json({
      ok: false,
      error: "No report generated yet."
    });
  }

  res.json({
    ok: true,
    report: latest.report_json
  });
});

app.get("/api/report/export", (req, res) => {
  const latest = store.generatedReports[0];

  if (!latest) {
    return res.status(404).send(
      "No report generated yet."
    );
  }

  const report = latest.report_json;
  const out = [];

  for (const d of report.districts) {
    d.rows.forEach((r, i) => {
      out.push({
        District: i === 0 ? d.district : "",
        Block: r.block,
        "Vehicle No.": r.vehicle,
        [`${report.firstDateDisplay} Total Received`]:
          r.yesterday.received,
        [`${report.firstDateDisplay} Total Attend`]:
          r.yesterday.attend,
        [`${report.firstDateDisplay} Remark`]:
          r.yesterday.remark,
        [`${report.secondDateDisplay || "Today"} Total Received`]:
          r.today.received,
        "Total Attend": r.today.attended,
        "Not Attend": r.today.notAttend,
        Pending: r.today.pending,
        [`${report.secondDateDisplay || "Today"} Remark`]:
          r.today.remark
      });
    });
  }

  out.push({
    District: "OVERALL GRAND TOTAL",
    Block: "",
    "Vehicle No.": "",
    [`${report.firstDateDisplay} Total Received`]:
      report.overall.yesterdayReceived,
    [`${report.firstDateDisplay} Total Attend`]:
      report.overall.yesterdayAttend,
    [`${report.firstDateDisplay} Remark`]: "",
    [`${report.secondDateDisplay || "Today"} Total Received`]:
      report.overall.todayReceived,
    "Total Attend": report.overall.attended,
    "Not Attend": report.overall.notAttend,
    Pending: report.overall.pending,
    [`${report.secondDateDisplay || "Today"} Remark`]: ""
  });

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(out);

  XLSX.utils.book_append_sheet(
    wb,
    ws,
    "MVU Daily Report"
  );

  const buf = XLSX.write(wb, {
    type: "buffer",
    bookType: "xlsx"
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  res.setHeader(
    "Content-Disposition",
    `attachment; filename="MVU_Daily_Report_${report.firstDateDisplay}_${report.secondDateDisplay || ""}.xlsx"`
  );

  res.send(buf);
});

app.get("/api/weekoff/template", (req, res) => {
  const rows = Object.values(store.weekOff).map(r => ({
    "Block": r.block || "",
    "Vehicle Number": r.vehicle_no || "",
    "Week Off": r.week_off || ""
  }));

  const data = rows.length
    ? rows
    : [{
        "Block": "",
        "Vehicle Number": "",
        "Week Off": ""
      }];

  const ws = XLSX.utils.json_to_sheet(
    data,
    {
      header: [
        "Block",
        "Vehicle Number",
        "Week Off"
      ]
    }
  );

  ws["!cols"] = [
    { wch: 28 },
    { wch: 20 },
    { wch: 16 }
  ];

  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(
    wb,
    ws,
    "Daily Data Entry"
  );

  const buf = XLSX.write(wb, {
    type: "buffer",
    bookType: "xlsx"
  });

  res.setHeader(
    "Content-Type",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );

  res.setHeader(
    "Content-Disposition",
    'attachment; filename="MVU_Week_Off_Template.xlsx"'
  );

  res.send(buf);
});

/* =========================
   NORMAL REPORT RESET
========================= */

app.post("/api/reset/report", (req, res) => {
  try {
    store.generatedReports = [];
    saveStore();

    res.json({
      ok: true,
      message:
        "Daily report reset. Please upload a new Detailed Report."
    });
  } catch (e) {
    res.status(500).json({
      ok: false,
      error: e.message
    });
  }
});

/* =========================
   HARD RESET
   PASSWORD: 1122
========================= */

app.post("/api/reset/hard", (req, res) => {
  try {
    const password = clean(
      req.body?.password
    );

    if (password !== "1122") {
      return res.status(403).json({
        ok: false,
        error: "Wrong password."
      });
    }

    store.villageMapping = {};
    store.weekOff = {};
    store.generatedReports = [];
    store.uploadLog = [];

    saveStore();

    return res.json({
      ok: true,
      message:
        "Hard Reset completed. Village Mapping, Week Off Master and current report have been removed."
    });
  } catch (e) {
    console.error("Hard Reset Error:", e);

    return res.status(500).json({
      ok: false,
      error:
        "Hard Reset failed: " + e.message
    });
  }
});

/* =========================
   EXPLICIT PAGE ROUTES
   No Express wildcard (*)
========================= */

app.get("/", (req, res) => {
  res.sendFile(
    path.join(ROOT, "public", "upload.html")
  );
});

app.get("/upload.html", (req, res) => {
  res.sendFile(
    path.join(ROOT, "public", "upload.html")
  );
});

app.get("/result.html", (req, res) => {
  res.sendFile(
    path.join(ROOT, "public", "result.html")
  );
});

/* API 404 MUST RETURN JSON */

app.use("/api", (req, res) => {
  res.status(404).json({
    ok: false,
    error: `API route not found: ${req.method} ${req.originalUrl}`
  });
});

/* General 404 */

app.use((req, res) => {
  res.status(404).send("Page not found.");
});

/* =========================
   START SERVER
   ONLY PORT 5173
========================= */

const server = app.listen(PORT, () => {
  console.log(
    `MVU Report website running at http://localhost:${PORT}`
  );
});

server.on("error", err => {
  if (err.code === "EADDRINUSE") {
    console.error(
      `ERROR: Port ${PORT} is already in use.`
    );
    console.error(
      "Close the old Node.js server and run npm start again."
    );
    process.exit(1);
  }

  console.error(err);
  process.exit(1);
});
