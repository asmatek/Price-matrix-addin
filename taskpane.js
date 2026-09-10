/* global Office, XLSX */

/* ==================================================================== *
 * Matrix Price Finder — energy matrix engine
 *
 * Real supplier matrices (NRG, ERCOT-style broker sheets, Chariot, etc.)
 * don't share one layout. This engine auto-detects, per sheet, one of two
 * shapes:
 *
 *   WIDE:  one row per (utility, zone, profile, usage-cap, start date),
 *          with term months as columns ("Term - 12", "12 Mth", or bare
 *          numeric headers like 1..60).
 *
 *   LONG:  one row per (utility, zone, profile, term, start date), with
 *          annual-usage kWh brackets as columns (e.g. "0-200,000").
 *
 * It also picks up any zip-code lookup table it finds (zip -> utility
 * and/or zone) so a zip code alone can resolve the right sheet rows.
 * ==================================================================== */

Office.onReady(() => {
  document.getElementById("searchBtn").addEventListener("click", runSearch);
});

/* ------------------------------------------------------------------ *
 * Vocabulary: header keywords and utility/zone synonyms.
 * Extend these as you see more supplier files.
 * ------------------------------------------------------------------ */

const HEADER_KEYWORDS = {
  utility: ["utility", "dc", "tdsp"],
  zone: ["zone", "congestionzone", "loadzone"],
  profile: ["loadprofile", "loadfactor", "profile"],
  usageCap: ["usagegroupkwh", "annualusagekwhs", "annualusagekwh", "annualusage", "usagemwh", "annualvolumekwh"],
  startDate: ["startdate", "startmonth"],
  product: ["productname", "product"],
  term: ["term"],
  zip: ["zipcode", "zip"],
};

// Canonical utility key -> known abbreviations / spellings seen across files.
const UTILITY_SYNONYMS = {
  CENTERPOINT: ["cnp", "cnpt", "cpt", "centerpoint", "centerpointenergy"],
  AEP_CENTRAL: ["cpl", "aeptcc", "aepcentral", "aeptxcentral", "aeptexascentral"],
  AEP_NORTH: ["wtu", "aeptnc", "aepnorth", "aeptxnorth", "aeptexasnorth"],
  ONCOR: ["onc", "oncor"],
  TNMP: ["tnp", "tnmp", "texasnewmexicopower"],
  LPL: ["lpl", "lubbockpowerlight"],
};

const ZONE_SYNONYMS = {
  HOUSTON: ["houston", "houstonlz"],
  NORTH: ["north", "northlz"],
  SOUTH: ["south", "southlz"],
  WEST: ["west", "westlz"],
};

function normKey(s) {
  return String(s ?? "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function matchSynonym(dict, raw) {
  const n = normKey(raw);
  if (!n) return null;
  for (const [canon, syns] of Object.entries(dict)) {
    if (syns.includes(n)) return canon;
  }
  for (const [canon, syns] of Object.entries(dict)) {
    if (syns.some((s) => n.includes(s) || s.includes(n))) return canon;
  }
  return null;
}

/* ------------------------------------------------------------------ *
 * Main flow
 * ------------------------------------------------------------------ */

async function runSearch() {
  clearOutputs();

  const query = readQueryForm();
  if (!query.terms.length) {
    setStatus("Enter at least one term (months) to search for, e.g. 12,24,36.", "error");
    return;
  }

  const searchBtn = document.getElementById("searchBtn");
  searchBtn.disabled = true;
  setStatus("Reading attachments…");

  try {
    const { priceSheets, zipMap, sourcesSummary } = await collectAllData();
    renderSourcesList(sourcesSummary);

    if (priceSheets.length === 0) {
      setStatus(
        "No recognizable matrix data found. This scans .xlsx/.xlsm/.xls/.csv attachments " +
          "(and linked spreadsheet files in the body) for sheets with utility/zone/term columns " +
          "— check the 'Files scanned' list below.",
        "error"
      );
      return;
    }

    let resolvedUtility = null;
    let resolvedZone = null;
    let zipNote = null;
    if (query.zip) {
      const hit = zipMap[query.zip];
      if (hit) {
        resolvedUtility = hit.utilityKey || null;
        resolvedZone = hit.zoneKey || null;
        zipNote =
          "Zip " + query.zip + " → " + (hit.utilityRaw || "utility unknown") +
          ", " + (hit.zoneRaw || "zone unknown") + " zone.";
      } else {
        zipNote = "Zip " + query.zip + " wasn't found in any zip map included in these files.";
      }
    }
    if (query.utilityText) {
      const typed = matchSynonym(UTILITY_SYNONYMS, query.utilityText);
      resolvedUtility = typed || resolvedUtility;
    }

    // Filters are applied WHILE extracting (not after), because several
    // real supplier sheets pad tens/hundreds of thousands of rows — wide
    // sheets multiply out by every term column, long sheets by every usage
    // bracket column. Extracting everything first and filtering afterward
    // reliably exhausts memory in a real browser tab (it did in testing).
    const rowFilters = {
      utilityKey: resolvedUtility,
      zoneKey: resolvedZone,
      customerClass: query.customerClass === "any" ? null : query.customerClass,
      usageTier: query.usageTier === "any" ? null : query.usageTier,
      startYear: query.startYear,
      startMonthNum: query.startMonthNum,
    };

    const EXTRACTION_CAP = 20000;
    const priceRows = [];
    let capped = false;
    for (const ps of priceSheets) {
      if (priceRows.length >= EXTRACTION_CAP) {
        capped = true;
        break;
      }
      const rows = extractPriceRows(ps.grid, ps.detection, ps.source, rowFilters);
      for (let i = 0; i < rows.length && priceRows.length < EXTRACTION_CAP; i++) {
        priceRows.push(rows[i]);
      }
    }

    const groups = buildGroups(priceRows, { annualUsage: query.annualUsage });

    if (groups.length === 0) {
      let msg = "No matrix rows matched those filters.";
      if (!resolvedUtility && !query.utilityText && !query.zip) {
        msg += " Try entering a zip code or a utility name.";
      } else {
        msg += " Double-check the utility/zone spelling, customer class, and usage tier.";
      }
      setStatus(msg, "error");
      return;
    }

    const cappedNote = capped
      ? " (stopped after " + EXTRACTION_CAP + " rows — narrow the search for a complete result)"
      : "";
    setStatus(
      (zipNote ? zipNote + " " : "") + "Found " + groups.length + " matching price grid(s)." + cappedNote,
      capped ? "" : "success"
    );
    renderResults(groups, query.terms);
  } catch (err) {
    console.error(err);
    setStatus("Something went wrong: " + (err && err.message ? err.message : err), "error");
  } finally {
    searchBtn.disabled = false;
  }
}

function readQueryForm() {
  const zip = document.getElementById("zip").value.trim();
  const utilityText = document.getElementById("utility").value.trim();
  const customerClass = document.getElementById("customerClass").value;
  const usageTier = document.getElementById("usageTier").value;
  const annualUsageRaw = document.getElementById("annualUsage").value.trim();
  const annualUsage = annualUsageRaw ? parseFloat(annualUsageRaw) : null;
  const startMonthRaw = document.getElementById("startMonth").value; // "YYYY-MM" or ""
  const termsRaw = document.getElementById("terms").value.trim();

  let startYear = null;
  let startMonthNum = null;
  if (startMonthRaw) {
    const parts = startMonthRaw.split("-").map((v) => parseInt(v, 10));
    startYear = parts[0];
    startMonthNum = parts[1]; // 1-12
  }

  const terms = termsRaw
    .split(",")
    .map((t) => parseInt(t.trim(), 10))
    .filter((t) => Number.isFinite(t) && t > 0);

  return { zip, utilityText, customerClass, usageTier, annualUsage, startYear, startMonthNum, terms };
}

function setStatus(message, kind) {
  const el = document.getElementById("status");
  el.textContent = message;
  el.className = kind || "";
  el.classList.remove("hidden");
}

function clearOutputs() {
  document.getElementById("results").classList.add("hidden");
  document.getElementById("resultsList").innerHTML = "";
  document.getElementById("sourcesFound").classList.add("hidden");
  document.getElementById("sourcesList").innerHTML = "";
  document.getElementById("status").classList.add("hidden");
}

/* ------------------------------------------------------------------ *
 * Gathering data from every attachment
 * ------------------------------------------------------------------ */

async function collectAllData() {
  const item = Office.context.mailbox.item;
  const priceSheets = []; // {source, grid, detection} — extraction deferred until filters are known
  const zipMap = {};
  const sourcesSummary = [];

  if (!item) {
    return { priceSheets, zipMap, sourcesSummary };
  }

  // Scans a parsed SheetJS workbook for zip-mapping sheets (extracted right
  // away — these are small) and price-matrix sheets (NOT extracted yet —
  // just the grid + detected layout is kept, so extraction can apply
  // filters row-by-row later instead of materializing every term/bracket
  // combination for every row in the sheet up front. Some supplier sheets
  // have 100,000+ rows and 60 term columns; exploding all of that before
  // filtering reliably exhausts memory in a real browser tab.
  function processWorkbook(workbook, label) {
    let sheetsUsed = 0;
    let zipRowsFound = 0;

    workbook.SheetNames.forEach((sheetName) => {
      const sheet = workbook.Sheets[sheetName];

      // Some workbooks carry a hidden sheet (often tied to a VBA userform,
      // e.g. "UserControl") whose declared used-range is Excel's absolute
      // max (1,048,576 rows × 16,384 columns) even though it's really
      // empty. Converting that to a grid tries to materialize a
      // multi-billion-cell array and reliably crashes the tab — so any
      // sheet with an implausible declared size for real matrix data is
      // skipped outright rather than parsed.
      const ref = sheet["!ref"];
      if (ref) {
        const range = XLSX.utils.decode_range(ref);
        const rows = range.e.r + 1;
        const cols = range.e.c + 1;
        if (rows > 200000 || cols > 500) {
          sourcesSummary.push({
            text: label + " / " + sheetName + " — skipped (implausible sheet size: " + rows + " rows × " + cols + " cols, likely a hidden control sheet)",
            warn: false,
          });
          return;
        }
      }

      const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
      if (!grid || grid.length < 2) return;

      const detection = detectSheetLayout(grid);
      if (!detection) return;

      if (detection.kind === "zipmap") {
        const found = extractZipMap(grid, detection);
        Object.assign(zipMap, found);
        zipRowsFound += Object.keys(found).length;
        sheetsUsed++;
      } else {
        const dataRowCount = countNonBlankDataRows(grid, detection.headerRow);
        priceSheets.push({ source: label + " — " + sheetName, grid, detection, dataRowCount });
        sheetsUsed++;
      }
    });

    sourcesSummary.push({
      text:
        label +
        " — " +
        sheetsUsed +
        " usable sheet(s)" +
        (zipRowsFound ? ", " + zipRowsFound + " zip mapping(s)" : ""),
    });
  }

  // 1) Real file attachments.
  if (item.attachments && item.attachments.length > 0) {
    const fileAttachments = item.attachments.filter(
      (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File
    );

    for (const att of fileAttachments) {
      const name = att.name || "attachment";
      if (!/\.(xlsx|xlsm|xls|csv)$/i.test(name)) continue;

      const content = await getAttachmentContent(item, att);
      if (!content) {
        sourcesSummary.push({ text: name + " — could not be read", warn: true });
        continue;
      }
      if (content.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
        sourcesSummary.push({ text: name + " — cloud attachment, can't read client-side", warn: true });
        continue;
      }

      let workbook;
      try {
        workbook = XLSX.read(content.content, { type: "base64", cellDates: true });
      } catch (e) {
        sourcesSummary.push({ text: name + " — failed to parse as a spreadsheet", warn: true });
        continue;
      }
      processWorkbook(workbook, name);
    }
  }

  // 2) Linked files in the email body — suppliers often send a link to
  // Office's online viewer (view.officeapps.live.com/op/view.aspx?src=...)
  // instead of an actual attachment. Pull the real file URL out of that
  // wrapper (or use direct .xlsx/.xlsm/.xls/.csv links as-is) and try to
  // fetch it. This only works if the hosting server allows cross-origin
  // requests from this add-in — see the README for what to do if it doesn't.
  const linkedUrls = await findLinkedFileUrls(item);
  for (const url of linkedUrls) {
    const label = filenameFromUrl(url);
    try {
      const resp = await fetch(url, { mode: "cors" });
      if (!resp.ok) {
        sourcesSummary.push({ text: label + " (linked) — server returned HTTP " + resp.status, warn: true });
        continue;
      }
      const buf = await resp.arrayBuffer();
      const workbook = XLSX.read(buf, { type: "array", cellDates: true });
      processWorkbook(workbook, label + " (linked)");
    } catch (e) {
      sourcesSummary.push({
        text:
          label +
          " (linked) — couldn't be fetched (" +
          (e && e.message ? e.message : "network/CORS error") +
          "). The hosting server likely doesn't allow cross-origin requests; " +
          "ask the sender to attach the file directly, or see the README for a proxy option.",
        warn: true,
      });
    }
  }

  return { priceSheets, zipMap, sourcesSummary };
}

function countNonBlankDataRows(grid, headerRow) {
  let count = 0;
  for (let r = headerRow + 1; r < grid.length; r++) {
    const row = grid[r];
    if (row && row.some((v) => v !== "" && v !== null && v !== undefined)) count++;
  }
  return count;
}

function getAttachmentContent(item, attachment) {
  return new Promise((resolve) => {
    item.getAttachmentContentAsync(attachment.id, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("Could not read attachment", attachment.name, result.error);
        resolve(null);
        return;
      }
      resolve(result.value);
    });
  });
}

/* ------------------------------------------------------------------ *
 * Linked-file discovery (Office viewer links / direct spreadsheet links
 * pasted into the email body).
 * ------------------------------------------------------------------ */

function findLinkedFileUrls(item) {
  return new Promise((resolve) => {
    if (!item.body) {
      resolve([]);
      return;
    }
    item.body.getAsync(Office.CoercionType.Html, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        resolve([]);
        return;
      }
      try {
        const html = result.value;
        const candidates = new Set();

        // Anchor tags (normal case for Outlook-rendered links).
        const doc = new DOMParser().parseFromString(html, "text/html");
        doc.querySelectorAll("a[href]").forEach((a) => candidates.add(a.getAttribute("href")));

        // Fallback: bare URLs anywhere in the HTML source, in case a link
        // wasn't auto-linked into an <a> tag.
        const urlRegex = /https?:\/\/[^\s"'<>]+/gi;
        const raw = html.match(urlRegex) || [];
        raw.forEach((u) => candidates.add(u));

        const resolved = new Set();
        candidates.forEach((raw) => {
          const fileUrl = resolveToFileUrl(raw);
          if (fileUrl) resolved.add(fileUrl);
        });

        resolve(Array.from(resolved));
      } catch (e) {
        console.error("Failed scanning body for linked files", e);
        resolve([]);
      }
    });
  });
}

function resolveToFileUrl(rawUrl) {
  let u;
  try {
    u = new URL(rawUrl);
  } catch (e) {
    return null;
  }

  // Office Online viewer wrapper: the real file is in the ?src= param.
  if (/(^|\.)officeapps\.live\.com$/i.test(u.hostname) || /(^|\.)office\.com$/i.test(u.hostname)) {
    const src = u.searchParams.get("src");
    if (src) {
      try {
        return decodeURIComponent(src);
      } catch (e) {
        return src;
      }
    }
  }

  // Direct link to a spreadsheet file (ignoring any query string).
  if (/\.(xlsx|xlsm|xls|csv)$/i.test(u.pathname)) {
    return rawUrl;
  }

  return null;
}

function filenameFromUrl(url) {
  try {
    const u = new URL(url);
    const last = u.pathname.split("/").filter(Boolean).pop();
    return decodeURIComponent(last || url);
  } catch (e) {
    return url;
  }
}

/* ------------------------------------------------------------------ *
 * Sheet layout detection
 * ------------------------------------------------------------------ */

function headerCategoryFor(cellText) {
  const n = normKey(cellText);
  if (!n) return null;
  for (const [category, keywords] of Object.entries(HEADER_KEYWORDS)) {
    if (keywords.includes(n)) return category;
  }
  return null;
}

function parseTermFromHeader(cellText) {
  if (typeof cellText === "number" && cellText > 0 && cellText <= 120) {
    return Math.round(cellText);
  }
  const s = String(cellText ?? "").trim();
  let m = s.match(/^term\s*-?\s*(\d+)$/i);
  if (m) return parseInt(m[1], 10);
  m = s.match(/^(\d+)\s*mth\.?$/i);
  if (m) return parseInt(m[1], 10);
  m = s.match(/^(\d+)\s*months?$/i);
  if (m) return parseInt(m[1], 10);
  if (/^\d+(\.0+)?$/.test(s)) {
    const v = parseInt(s, 10);
    if (v > 0 && v <= 120) return v;
  }
  return null;
}

function parseTermFromCellValue(cellValue) {
  const s = String(cellValue ?? "").trim();
  const m = s.match(/(\d+)\s*months?/i);
  if (m) return parseInt(m[1], 10);
  if (/^\d+$/.test(s)) return parseInt(s, 10);
  return null;
}

function parseUsageBracketHeader(cellText) {
  const s = String(cellText ?? "").trim();
  let m = s.match(/^([\d,]+)\s*-\s*([\d,]+)$/);
  if (!m) m = s.match(/^([\d,]+)\s*to\s*([\d,]+)$/i);
  if (!m) return null;
  const min = parseInt(m[1].replace(/,/g, ""), 10);
  const max = parseInt(m[2].replace(/,/g, ""), 10);
  if (!Number.isFinite(min) || !Number.isFinite(max)) return null;
  return { label: s, min, max };
}

// Scans the first ~40 rows for the best header row, then classifies the
// sheet as a zip-map, a WIDE price matrix, or a LONG price matrix.
function detectSheetLayout(grid) {
  const maxScanRows = Math.min(grid.length, 40);
  let best = null;

  for (let r = 0; r < maxScanRows; r++) {
    const row = grid[r];
    if (!row) continue;

    const colMap = {};
    row.forEach((cell, c) => {
      const cat = headerCategoryFor(cell);
      if (cat && !(cat in colMap)) colMap[cat] = c;
    });

    const zipLike = "zip" in colMap && (("zone" in colMap) || ("utility" in colMap));
    const score =
      ["utility", "zone", "profile", "startDate", "product", "usageCap"].filter((k) => k in colMap).length;

    if (zipLike) {
      const candidate = { kind: "zipmap", headerRow: r, colMap };
      if (!best || best.kind !== "zipmap") best = candidate;
      continue;
    }

    if (score >= 3 && (!best || best.kind === "zipmap" || score > best.score)) {
      best = { kind: "pricematrix", headerRow: r, colMap, score };
    }
  }

  if (!best) return null;
  if (best.kind === "zipmap") return best;

  const headerRow = grid[best.headerRow] || [];
  const usedCols = new Set(Object.values(best.colMap));
  const termCols = [];
  const bracketCols = [];

  headerRow.forEach((cell, c) => {
    if (usedCols.has(c)) return;
    const bracket = parseUsageBracketHeader(cell);
    if (bracket) {
      bracketCols.push({ col: c, ...bracket });
      return;
    }
    if (!("term" in best.colMap)) {
      const term = parseTermFromHeader(cell);
      if (term !== null) termCols.push({ col: c, term });
    }
  });

  if ("term" in best.colMap) {
    return { ...best, layout: "long", bracketCols };
  }
  if (termCols.length > 0) {
    return { ...best, layout: "wide", termCols };
  }
  return null;
}

/* ------------------------------------------------------------------ *
 * Extraction
 * ------------------------------------------------------------------ */

function classifyProfile(profileRaw) {
  const n = String(profileRaw ?? "").trim().toUpperCase();
  if (!n) return { customerClass: "unknown", tier: null };

  if (n.includes("RES")) {
    let tier = null;
    if (n.includes("HI")) tier = "high";
    else if (n.includes("LO")) tier = "low";
    else if (n.includes("MED")) tier = "medium";
    return { customerClass: "residential", tier };
  }

  if (n.startsWith("BUS") || n === "HI" || n === "MED" || n === "LO" || n.includes("LIGHT")) {
    let tier = null;
    if (n.includes("HI")) tier = "high";
    else if (n.includes("MED")) tier = "medium";
    else if (n.includes("LO")) tier = "low";
    return { customerClass: "commercial", tier };
  }

  return { customerClass: "unknown", tier: null };
}

function toDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === "string" && value.trim()) {
    const d = new Date(value);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

// `filters` (optional) lets the caller skip rows immediately — before any
// term/bracket columns are expanded into individual price entries. This
// matters: a sheet with 140,000 data rows × 5 usage-bracket columns is
// 700,000 entries if extracted blind, but often only a few hundred once a
// utility, zone, customer class, and tier are known.
function extractPriceRows(grid, detection, source, filters) {
  const { colMap, headerRow, layout } = detection;
  const out = [];

  for (let r = headerRow + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row || row.every((v) => v === "" || v === null || v === undefined)) continue;

    const utilityRaw = colMap.utility !== undefined ? row[colMap.utility] : "";
    const zoneRaw = colMap.zone !== undefined ? row[colMap.zone] : "";
    const profileRaw = colMap.profile !== undefined ? row[colMap.profile] : "";
    const usageCapRaw = colMap.usageCap !== undefined ? row[colMap.usageCap] : "";
    const startDate = colMap.startDate !== undefined ? toDate(row[colMap.startDate]) : null;

    if (!utilityRaw && !zoneRaw) continue;

    const { customerClass, tier } = classifyProfile(profileRaw);
    const utilityKey = matchSynonym(UTILITY_SYNONYMS, utilityRaw);
    const zoneKey = matchSynonym(ZONE_SYNONYMS, zoneRaw);
    // WIDE-format usage caps are just a label by default (e.g. "0-300,000"),
    // but they follow the same "<min>-<max>" shape as LONG-format bracket
    // headers, so reuse that parser to get numeric bounds when possible —
    // this lets the smallest-bracket default and annualUsage filter apply
    // consistently regardless of which layout a supplier used.
    const usageCapBracket = usageCapRaw ? parseUsageBracketHeader(usageCapRaw) : null;

    if (filters) {
      if (filters.utilityKey && (!utilityKey || utilityKey !== filters.utilityKey)) continue;
      if (filters.zoneKey && (!zoneKey || zoneKey !== filters.zoneKey)) continue;
      if (filters.customerClass && customerClass !== filters.customerClass) continue;
      if (filters.usageTier && tier !== filters.usageTier) continue;
      if (filters.startYear && startDate) {
        if (startDate.getFullYear() !== filters.startYear) continue;
        if (startDate.getMonth() + 1 !== filters.startMonthNum) continue;
      }
    }

    const base = {
      source,
      startDate,
      utilityRaw: String(utilityRaw ?? ""),
      utilityKey,
      zoneRaw: String(zoneRaw ?? ""),
      zoneKey,
      customerClass,
      tier,
      usageCapLabel: usageCapRaw ? String(usageCapRaw) : null,
    };

    if (layout === "wide") {
      detection.termCols.forEach(({ col, term }) => {
        const price = row[col];
        if (typeof price === "number" && price > 0) {
          out.push({
            ...base,
            term,
            price,
            bracketLabel: base.usageCapLabel,
            bracketMin: usageCapBracket ? usageCapBracket.min : null,
            bracketMax: usageCapBracket ? usageCapBracket.max : null,
          });
        }
      });
    } else {
      const term = colMap.term !== undefined ? parseTermFromCellValue(row[colMap.term]) : null;
      if (term === null) continue;
      if (detection.bracketCols.length > 0) {
        detection.bracketCols.forEach(({ col, label, min, max }) => {
          const price = row[col];
          if (typeof price === "number" && price > 0) {
            out.push({ ...base, term, price, bracketLabel: label, bracketMin: min, bracketMax: max });
          }
        });
      }
    }
  }

  return out;
}

function extractZipMap(grid, detection) {
  const { colMap, headerRow } = detection;
  const out = {};
  for (let r = headerRow + 1; r < grid.length; r++) {
    const row = grid[r];
    if (!row) continue;
    const zipRaw = row[colMap.zip];
    if (zipRaw === "" || zipRaw === null || zipRaw === undefined) continue;
    const zip = String(zipRaw).trim().replace(/\.0$/, "");
    if (!/^\d{4,5}$/.test(zip)) continue;

    const utilityRaw = colMap.utility !== undefined ? row[colMap.utility] : "";
    const zoneRaw = colMap.zone !== undefined ? row[colMap.zone] : "";
    const utilityKey = utilityRaw ? matchSynonym(UTILITY_SYNONYMS, utilityRaw) : null;
    const zoneKey = zoneRaw ? matchSynonym(ZONE_SYNONYMS, zoneRaw) : null;

    const zip5 = zip.padStart(5, "0");
    if (!out[zip5] || (!out[zip5].utilityKey && utilityKey)) {
      out[zip5] = { utilityRaw: String(utilityRaw || ""), utilityKey, zoneRaw: String(zoneRaw || ""), zoneKey };
    }
  }
  return out;
}

/* ------------------------------------------------------------------ *
 * Grouping + filtering
 * ------------------------------------------------------------------ */

// By this point `priceRows` has already been filtered by utility/zone/class/
// tier/start-date during extraction (see extractPriceRows). This function's
// job is just to group into per-term price curves and pick the right usage
// bracket — `filters.annualUsage` is the only filter still relevant here.
function buildGroups(priceRows, filters) {
  const byKey = new Map();
  priceRows.forEach((row) => {
    const dateKey = row.startDate ? row.startDate.toISOString().slice(0, 7) : "unknown-date";
    const key = [
      row.source,
      row.utilityRaw,
      row.zoneRaw,
      row.customerClass,
      row.tier,
      dateKey,
      row.bracketLabel || "",
    ].join("|");
    if (!byKey.has(key)) {
      byKey.set(key, {
        source: row.source,
        utilityRaw: row.utilityRaw,
        zoneRaw: row.zoneRaw,
        customerClass: row.customerClass,
        tier: row.tier,
        startDate: row.startDate,
        bracketLabel: row.bracketLabel,
        bracketMin: row.bracketMin,
        bracketMax: row.bracketMax,
        terms: new Map(),
      });
    }
    byKey.get(key).terms.set(row.term, row.price);
  });

  let groups = Array.from(byKey.values());

  // Collapse to the earliest start date per key. When a start month was
  // requested, extraction already narrowed every row to that one month, so
  // this is a no-op; when no start month was given, this picks the nearest
  // upcoming pricing instead of showing every future month at once.
  {
    const earliestByKey = new Map();
    groups.forEach((g) => {
      const k = [g.source, g.utilityRaw, g.zoneRaw, g.customerClass, g.tier, g.bracketLabel || ""].join("|");
      const existing = earliestByKey.get(k);
      if (!existing || (g.startDate && existing.startDate && g.startDate < existing.startDate)) {
        earliestByKey.set(k, g);
      }
    });
    groups = Array.from(earliestByKey.values());
  }

  const withBracketKey = new Map();
  groups.forEach((g) => {
    const baseKey = [g.source, g.utilityRaw, g.zoneRaw, g.customerClass, g.tier].join("|");
    if (g.bracketMin === null || g.bracketMin === undefined) {
      withBracketKey.set(baseKey + "|" + (g.bracketLabel || ""), g);
      return;
    }
    if (filters.annualUsage != null) {
      if (filters.annualUsage >= g.bracketMin && filters.annualUsage <= g.bracketMax) {
        withBracketKey.set(baseKey, g);
      }
    } else {
      const existing = withBracketKey.get(baseKey);
      if (!existing || g.bracketMin < existing.bracketMin) {
        withBracketKey.set(baseKey, g);
      }
    }
  });

  const finalByKey = new Map();
  Array.from(withBracketKey.values()).forEach((g) => {
    const workbook = g.source.split(" — ")[0];
    const k = [workbook, g.utilityRaw, g.zoneRaw, g.customerClass, g.tier, g.bracketLabel || ""].join("|");
    const existing = finalByKey.get(k);
    if (!existing || g.terms.size > existing.terms.size) {
      finalByKey.set(k, g);
    }
  });

  return Array.from(finalByKey.values()).sort((a, b) => a.utilityRaw.localeCompare(b.utilityRaw));
}

/* ------------------------------------------------------------------ *
 * Sweet-spot detection
 * ------------------------------------------------------------------ */

function findSweetSpots(sortedTermPrices) {
  const sweet = new Set();
  for (let i = 0; i < sortedTermPrices.length; i++) {
    const prev = i > 0 ? sortedTermPrices[i - 1].price : null;
    const next = i < sortedTermPrices.length - 1 ? sortedTermPrices[i + 1].price : null;
    const cur = sortedTermPrices[i].price;
    const lowerThanPrev = prev === null || cur < prev;
    const lowerThanNext = next === null || cur < next;
    if (lowerThanPrev && lowerThanNext) sweet.add(sortedTermPrices[i].term);
  }
  if (sortedTermPrices.length > 0) {
    const min = sortedTermPrices.reduce((a, b) => (b.price < a.price ? b : a));
    sweet.add(min.term);
  }
  return sweet;
}

/* ------------------------------------------------------------------ *
 * Rendering
 * ------------------------------------------------------------------ */

function renderSourcesList(sourcesSummary) {
  const list = document.getElementById("sourcesList");
  list.innerHTML = "";
  sourcesSummary.forEach((s) => {
    const li = document.createElement("li");
    li.textContent = s.text;
    if (s.warn) li.className = "warn";
    list.appendChild(li);
  });
  document.getElementById("sourcesFound").classList.remove("hidden");
}

function formatMoney(v) {
  return Number(v).toFixed(5).replace(/0+$/, "").replace(/\.$/, "");
}

function renderResults(groups, requestedTerms) {
  const container = document.getElementById("resultsList");
  container.innerHTML = "";

  groups.forEach((g) => {
    const sortedTermPrices = Array.from(g.terms.entries())
      .map(([term, price]) => ({ term, price }))
      .sort((a, b) => a.term - b.term);
    const sweetSpots = findSweetSpots(sortedTermPrices);
    const priceByTerm = new Map(sortedTermPrices.map((tp) => [tp.term, tp.price]));

    const card = document.createElement("div");
    card.className = "match-card";

    const title = document.createElement("h3");
    title.textContent =
      g.utilityRaw + " — " + g.zoneRaw + " — " + (g.customerClass || "unknown") +
      (g.tier ? " (" + g.tier + ")" : "") +
      (g.startDate ? " — starts " + g.startDate.toLocaleDateString("en-US", { month: "long", year: "numeric" }) : "");
    card.appendChild(title);

    if (g.bracketLabel) {
      const bracket = document.createElement("div");
      bracket.className = "match-meta";
      bracket.textContent = "Usage bracket: " + g.bracketLabel + " kWh/yr";
      card.appendChild(bracket);
    }

    const pillRow = document.createElement("div");
    pillRow.className = "term-grid";
    const shown = new Set();
    requestedTerms.forEach((t) => {
      pillRow.appendChild(makeTermPill(t, priceByTerm.get(t), sweetSpots.has(t)));
      shown.add(t);
    });
    Array.from(sweetSpots)
      .filter((t) => !shown.has(t))
      .sort((a, b) => a - b)
      .forEach((t) => {
        pillRow.appendChild(makeTermPill(t, priceByTerm.get(t), true));
      });
    card.appendChild(pillRow);

    if (sortedTermPrices.length > requestedTerms.length) {
      const details = document.createElement("details");
      details.className = "all-terms";
      const summary = document.createElement("summary");
      summary.textContent = "All " + sortedTermPrices.length + " available terms";
      details.appendChild(summary);
      const table = document.createElement("table");
      sortedTermPrices.forEach(({ term, price }) => {
        const tr = document.createElement("tr");
        if (sweetSpots.has(term)) tr.className = "sweet";
        const tdTerm = document.createElement("td");
        tdTerm.textContent = term + " mo";
        const tdPrice = document.createElement("td");
        tdPrice.textContent = formatMoney(price);
        tr.appendChild(tdTerm);
        tr.appendChild(tdPrice);
        table.appendChild(tr);
      });
      details.appendChild(table);
      card.appendChild(details);
    }

    const meta = document.createElement("div");
    meta.className = "match-meta";
    meta.textContent = "Source: " + g.source;
    card.appendChild(meta);

    container.appendChild(card);
  });

  document.getElementById("results").classList.remove("hidden");
}

function makeTermPill(term, price, isSweet) {
  const pill = document.createElement("div");
  pill.className = "term-pill" + (price === undefined ? " missing" : "") + (isSweet && price !== undefined ? " sweet" : "");
  const label = document.createElement("div");
  label.className = "term-label";
  label.textContent = term + " mo";
  const val = document.createElement("div");
  val.className = "term-price";
  val.textContent = price === undefined ? "not found" : formatMoney(price);
  pill.appendChild(label);
  pill.appendChild(val);
  return pill;
}
