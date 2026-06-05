/**
 * url-normalizer.runtime.gs
 * Enhanced final version: status-sheet self-healing, exact-host fast path,
 * selected/visible-row batch processing, cached rules, optional redirect expansion.
 * Fast runtime URL normalizer logic for Google Sheets.
 *
 * Keep this file installed permanently.
 * It reads normalization rules from URL_RULES sheet.
 * One-time seed/reset logic should stay in url-rules-seeder.gs.
 *
 * Main runners:
 *   RUN_NORMALIZE_NOW()
 *   RUN_NORMALIZE_WITH_REDIRECTS()
 *   RESUME_NORMALIZE()
 *   SHOW_NORMALIZE_STATUS()
 *   RESET_NORMALIZE_STATUS()
 */

const CONFIG = CFG.URL_NORMALIZER;

const AUTO_NORMALIZE_SHEET_SET = new Set(CFG.JOB_SHEETS_FALLBACK);

const TRACKING_PARAM_PREFIXES = [
  "utm_",
  "pk_",
  "mc_",
  "mkt_",
  "ref_",
  "vero_",
  "cs_utm_"
];

const TRACKING_PARAMS_EXACT = new Set([
  "gclid",
  "gbraid",
  "wbraid",
  "fbclid",
  "msclkid",
  "yclid",
  "ttclid",
  "_gl",

  "gh_src",
  "lever-source",
  "lever-source[]",

  // Generic source/referral params that should not remain on canonical job URLs.
  "source",
  "src",
  "ref",
  "referrer",
  "from",
  "origin",
  "jobsite",
  "st",
  "can",
  "trid",
  "gns",

  "li_fat_id",

  "trk",
  "tracking",
  "trackingid",
  "tracking_id",

  "_hsenc",
  "_hsmi",
  "hsctatracking",

  "irclickid",
  "irgwc",

  "affid",
  "affiliate",
  "affiliate_id",

  "partner",
  "partnerid",

  "clickid",
  "click_id",

  "lever-origin",
  "promotion",
  "state",
  "apply",
  "_ghcid",
  "it",
  "day_range",
  "workplace",
  "employment_type",
  "experience",
  "query",
  "location",

  // Common publisher/referral id. Specific job boards can also drop it via URL_RULES.
  "pid"
]);

const GENERIC_ASHBY_PATH_RE = /^\/(?:about\/)?(?:careers|jobs|openings|positions)\/?$/i;

// Host/path-specific query params that are the REAL job identity.
// These are protected even when URL_RULES is stale/missing or the global fallback rule matches first.
const IDENTITY_QUERY_KEEP_RULES = [
  { host: "linkedin.com", pathRe: /^\/jobs\/search$/i, keep: ["currentjobid"] },
  { host: "linkedin.com", pathRe: /^\/jobs\/search\/?$/i, keep: ["currentjobid"] },
  { host: "jobs.paymentology.com", pathRe: /^\/detail$/i, keep: ["uid", "coref"] },
  { host: "ashbyhq.com", pathRe: /^\/careers$/i, keep: ["ashby_jid"] },
  { host: "kuali.co", pathRe: /^\/positions$/i, keep: ["gnk", "gni"] },
  { host: "job-boards.greenhouse.io", pathRe: /^\/embed\/job_app$/i, keep: ["token"] }
];


let MEMORY_RULES_CACHE = null;
let MEMORY_HEADER_COL_CACHE = Object.create(null);
let REDIRECTS_USED_THIS_RUN = 0;



/**
 * Adds a menu so batch functions run with proper authorization.
 * IMPORTANT: Do not run RUN_NORMALIZE_NOW() as a sheet formula.
 */


function assertManualRunner_() {
  // Best-effort guard. Custom functions cannot create sheets/write properties reliably.
  // Running from menu/editor/button gives authorization.
  try {
    SpreadsheetApp.getActiveSpreadsheet();
    PropertiesService.getScriptProperties().getProperty(CONFIG.NORMALIZER_RUNTIME_STATE_KEY);
  } catch (e) {
    throw new Error(
      "Run URL Normalizer from the Apps Script editor, button, or the URL Normalizer menu. " +
      "Do not run RUN_NORMALIZE_NOW() as a sheet formula/custom function."
    );
  }
}

function WARM_URL_NORMALIZER_RUNTIME() {
  assertManualRunner_();

  const rawRules = readRulesFromProperties_();

  if (!rawRules || !rawRules.length) {
    throw new Error(
      "Rules snapshot missing. Run UPDATE_RULES_SNAPSHOT_FROM_SHEET() first."
    );
  }

  writeRulesToScriptCache_(rawRules);
  MEMORY_RULES_CACHE = buildRuntimeRules_(rawRules);

  SpreadsheetApp.getActive().toast(
    "Runtime cache warmed from published snapshot.",
    "URL Normalizer",
    5
  );
}

function UPDATE_RULES_SNAPSHOT_FROM_SHEET() {
  assertManualRunner_();

  const rawRules = readRulesFromSheet_();

  if (!rawRules || !rawRules.length) {
    throw new Error("No rules found in URL_RULES sheet.");
  }

  validateRuntimeRules_(rawRules);

  writeRulesToPropertiesStrict_(rawRules);
  writeRulesToScriptCache_(rawRules);
  writeRulesMeta_(rawRules);

  MEMORY_RULES_CACHE = buildRuntimeRules_(rawRules);

  SpreadsheetApp.getActive().toast(
    "Rules published successfully: Sheet → Properties → Cache → Memory.",
    "URL Normalizer",
    5
  );
}

function RESET_RULE_SNAPSHOT_ONLY() {

  const props =
    PropertiesService.getScriptProperties();

  const baseKey =
    CONFIG.RULES_SNAPSHOT_PROPERTY_KEY;

  const chunkCount =
    Number(
      props.getProperty(
        baseKey + "_chunk_count"
      )
    ) || 0;

  props.deleteProperty(baseKey);

  for (let i = 0; i < chunkCount; i++) {

    props.deleteProperty(
      baseKey + "_" + i
    );
  }

  props.deleteProperty(
    baseKey + "_chunk_count"
  );

  props.deleteProperty(
    "url_rules_snapshot_meta"
  );

  Logger.log("Rules snapshot removed.");
}

/** Normal run, no redirect expansion. */
function RUN_NORMALIZE_NOW() {
  assertManualRunner_();
  runOrResumeNormalize_(false);
}

function RUN_NORMALIZE_WITH_REDIRECTS() {
  assertManualRunner_();
  runOrResumeNormalize_(true);
}

function runOrResumeNormalize_(expandRedirects) {
  const saved = getSavedStatus_();
  const unfinished = saved &&
    saved.sheetName &&
    saved.state !== CFG.VALUES.RUN_STATE.DONE &&
    saved.state !== CFG.VALUES.RUN_STATE.READY;

  normalizeRangeInPlaceSafe_({
    expandRedirects: unfinished ? !!saved.expandRedirects : expandRedirects,
    chunkSize: unfinished ? saved.chunkSize : (expandRedirects ? CONFIG.REDIRECT_CHUNK_SIZE : CONFIG.DEFAULT_CHUNK_SIZE),
    resume: !!unfinished
  });
}

/** Continue from last stopped row. */
function RESUME_NORMALIZE() {
  assertManualRunner_();
  const status = getSavedStatus_();
  if (!status || !status.sheetName) {
    SpreadsheetApp.getActive().toast("No saved normalize status found.", "URL Normalizer", 5);
    return;
  }

  normalizeRangeInPlaceSafe_({
    expandRedirects: !!status.expandRedirects,
    chunkSize: status.chunkSize || CONFIG.DEFAULT_CHUNK_SIZE,
    resume: true
  });
}

/** Show current saved status. */
function SHOW_NORMALIZE_STATUS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  assertManualRunner_();

  const status = getSavedStatus_();

  if (!status) {
    ss.toast("No saved normalize status.", "URL Normalizer", 5);
    return;
  }

  writeStatusSheet_(status);

  const runtimeMin = status.startedAt
    ? ((Date.now() - new Date(status.startedAt).getTime()) / 60000).toFixed(1)
    : "-";

  ss.toast(
    [
      `Status: ${status.state || "-"}`,
      `Sheet: ${status.sheetName || "-"}`,
      `Processed: ${status.processedCount || 0}/${status.totalRowsPlanned || 0}`,
      `Replaced: ${status.replacedCount || 0}`,
      `Last Row: ${status.lastProcessedRow || "-"}`,
      `Last Replaced: ${status.lastReplacedRow || "-"}`,
      `Redirects: ${status.expandRedirects ? CFG.VALUES.BOOLEAN_TEXT.ON : CFG.VALUES.BOOLEAN_TEXT.OFF}`,
      `Runtime: ${runtimeMin} min`
    ].join(" | "),
    "URL Normalizer",
    10
  );
}

/** Clear saved progress. */
function RESET_NORMALIZE_STATUS() {
  assertManualRunner_();
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.NORMALIZER_RUNTIME_STATE_KEY);
  const status = createStatus_(CFG.VALUES.RUN_STATE.READY, "manual", "", "", false, CONFIG.DEFAULT_CHUNK_SIZE, 0, null);
  status.processedCount = 0;
  status.replacedCount = 0;
  status.lastProcessedRow = "";
  status.lastReplacedRow = "";
  saveStatus_(status);
  writeStatusSheet_(status);
  SpreadsheetApp.getActive().toast("Normalize status reset.", "URL Normalizer", 5);
}

/**
 * Handles one-cell edits and multi-row pasted links.
 */
function onEdit(e) {

  try {
    if (!e || !e.range) return;

    const sh = e.range.getSheet();
    const sheetName = sh.getName();

    // If headers are edited/moved/renamed, clear cached header columns.
    if (e.range.getRow() === CONFIG.HEADER_ROW) {
      MEMORY_HEADER_COL_CACHE = Object.create(null);
    }

    if (sheetName === CONFIG.RULES_SHEET_NAME) {
      MEMORY_RULES_CACHE = null;

      try {
        CacheService.getScriptCache().remove(CONFIG.RULES_CACHE_KEY);
      } catch (e) { }

      SpreadsheetApp.getActive().toast(
        "URL_RULES edited, but not applied yet. Run UPDATE_RULES_SNAPSHOT_FROM_SHEET() to apply changes.",
        "URL Normalizer",
        5
      );
      return;
    }

    if (!AUTO_NORMALIZE_SHEET_SET.has(sheetName)) return;
    if (e.range.getRow() <= CONFIG.HEADER_ROW) return;

    const jobUrlCol = getHeaderColumn_(sh, CONFIG.JOB_URL_HEADER);
    if (!jobUrlCol) return;

    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;
    if (jobUrlCol < startCol || jobUrlCol > endCol) return;

    normalizeSpecificRowsSafe_(sh, jobUrlCol, e.range.getRow(), e.range.getNumRows(), {
      source: "onEdit",
      expandRedirects: false,
      chunkSize: CONFIG.DEFAULT_CHUNK_SIZE
    });
  } catch (err) {
    console.error("onEdit normalize error:", err && err.stack ? err.stack : err);
  }
}

/**
 * Main safe batch runner.
 * Processes selected visible rows first. If no selection, processes visible rows in whole sheet.
 */
function normalizeRangeInPlaceSafe_(options) {
  options = options || {};

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    SpreadsheetApp.getActive().toast("Normalizer is already running.", "URL Normalizer", 5);
    return;
  }

  const startedAt = Date.now();
  const deadline = startedAt + CONFIG.SAFE_RUNTIME_MS - CONFIG.STOP_BUFFER_MS;
  REDIRECTS_USED_THIS_RUN = 0;

  let status = null;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const saved = options.resume ? getSavedStatus_() : null;
    const sh = saved && saved.sheetName ? ss.getSheetByName(saved.sheetName) : ss.getActiveSheet();

    if (!sh) return;

    const sheetName = sh.getName();
    if (sheetName === CONFIG.STATUS_SHEET_NAME || sheetName === CONFIG.RULES_SHEET_NAME) {
      ss.toast("Open a data sheet first. Status/rules sheets are not normalized.", "URL Normalizer", 5);
      return;
    }
    if (!AUTO_NORMALIZE_SHEET_SET.has(sheetName)) {
      ss.toast("Active sheet is not configured for auto-normalize: " + sheetName, "URL Normalizer", 5);
      return;
    }

    const lastRow = sh.getLastRow();
    if (lastRow <= CONFIG.HEADER_ROW) return;

    const jobUrlCol = getHeaderColumn_(sh, CONFIG.JOB_URL_HEADER);
    if (!jobUrlCol) {
      ss.toast("Cannot find header: " + CONFIG.JOB_URL_HEADER, "URL Normalizer", 5);
      return;
    }

    const expandRedirects = !!options.expandRedirects;
    const chunkSize = clampInt_(options.chunkSize, 1, CONFIG.MAX_CHUNK_SIZE, CONFIG.DEFAULT_CHUNK_SIZE);
    const rowsToProcess = buildRowsToProcess_(ss, sh, lastRow, options.resume, saved);

    if (!rowsToProcess.length) {
      ss.toast("No rows to process.", "URL Normalizer", 5);
      return;
    }

    status = createStatus_(CFG.VALUES.RUN_STATE.RUNNING, "manual", sheetName, jobUrlCol, expandRedirects, chunkSize, rowsToProcess.length, saved);
    saveStatus_(status);
    writeStatusSheet_(status);
    let lastStatusWriteAt = Date.now();

    for (let i = 0; i < rowsToProcess.length; i += chunkSize) {
      if (Date.now() >= deadline) {
        stopBeforeTimeout_(ss, status);
        return;
      }

      const chunkRows = rowsToProcess.slice(i, i + chunkSize);
      const blocks = groupContiguous_(chunkRows);

      for (let b = 0; b < blocks.length; b++) {
        processBlock_(sh, jobUrlCol, blocks[b].start, blocks[b].end, {
          expandRedirects: expandRedirects,
          deadline: deadline,
          status: status,
          dryRun: !!options.dryRun
        });

        const nowForStatus = Date.now();
        if (
          status.processedCount % CONFIG.STATUS_UPDATE_EVERY_ROWS === 0 ||
          nowForStatus - lastStatusWriteAt >= CONFIG.STATUS_UPDATE_EVERY_MS ||
          nowForStatus >= deadline
        ) {
          saveStatus_(status);
          writeStatusSheet_(status);
          lastStatusWriteAt = nowForStatus;
          maybeToastProgress_(ss, status);
        }

        if (Date.now() >= deadline) {
          stopBeforeTimeout_(ss, status);
          return;
        }
      }
    }

    status.state = CFG.VALUES.RUN_STATE.DONE;
    status.updatedAt = new Date().toISOString();
    saveStatus_(status);
    writeStatusSheet_(status);

    ss.toast(
      "Done. Processed: " + status.processedCount +
      " | Replaced: " + status.replacedCount +
      " | Last replaced row: " + (status.lastReplacedRow || "-"),
      "URL Normalizer",
      10
    );
  } catch (err) {
    if (status) {
      status.state = CFG.VALUES.RUN_STATE.ERROR;
      status.lastErrorMessage = err && err.message ? err.message : String(err);
      status.updatedAt = new Date().toISOString();
      saveStatus_(status);
      writeStatusSheet_(status);
    }
    throw err;
  } finally {
    lock.releaseLock();
  }
}

/** Used by onEdit pasted rows. */
function normalizeSpecificRowsSafe_(sh, jobUrlCol, startRow, numRows, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const startedAt = Date.now();
  const deadline = startedAt + CONFIG.ON_EDIT_RUNTIME_MS;

  const status = createStatus_(
    CFG.VALUES.RUN_STATE.RUNNING,
    options.source || "onEdit",
    sh.getName(),
    jobUrlCol,
    false,
    options.chunkSize || CONFIG.DEFAULT_CHUNK_SIZE,
    numRows,
    null
  );

  const endRow = startRow + numRows - 1;

  processBlock_(sh, jobUrlCol, startRow, endRow, {
    expandRedirects: false,
    deadline: deadline,
    status: status
  });

  status.state = Date.now() >= deadline
    ? CFG.VALUES.RUN_STATE.STOPPED_BEFORE_TIMEOUT
    : CFG.VALUES.RUN_STATE.DONE;

  status.updatedAt = new Date().toISOString();
  if (options.source !== "onEdit") {
    saveStatus_(status);
  }

  if (numRows > 5 || status.state !== CFG.VALUES.RUN_STATE.DONE) {
    writeStatusSheet_(status);

    if (status.state === CFG.VALUES.RUN_STATE.STOPPED_BEFORE_TIMEOUT) {
      ss.toast(
        "Large paste partially normalized. Some rows stayed unchanged. Last processed: " +
        status.lastProcessedRow +
        ". Paste smaller batches or run Resume normalize.",
        "URL Normalizer",
        8
      );
    } else {
      ss.toast(
        "Paste normalize done. Processed: " + status.processedCount +
        " | Replaced: " + status.replacedCount +
        " | Last replaced: " + (status.lastReplacedRow || "-"),
        "URL Normalizer",
        6
      );
    }
  }
}

/**
 * Process contiguous rows safely.
 * Reads/writes one rectangular range; normalizes rows in memory.
 */
function processBlock_(sh, jobUrlCol, startRow, endRow, run) {
  const len = endRow - startRow + 1;
  const range = sh.getRange(startRow, jobUrlCol, len, 1);
  const values = range.getValues();
  const output = new Array(values.length);
  let changed = false;
  let outputLen = values.length;

  for (let i = 0; i < values.length; i++) {
    const rowNum = startRow + i;
    const raw = values[i][0];

    if (Date.now() >= run.deadline) {
      for (let j = i; j < values.length; j++) output[j] = [values[j][0]];
      break;
    }

    try {
      let next = raw;

      if (!isBlank_(raw)) {
        const expanded = run.expandRedirects ? expandUrlOnceCached_(raw) : raw;
        next = normalizeJobUrlInternal_(expanded);

        if (String(next) !== String(raw)) {
          changed = true;
          run.status.replacedCount++;
          run.status.lastReplacedRow = rowNum;
        }
      }

      output[i] = [next];
      run.status.processedCount++;
      run.status.lastProcessedRow = rowNum;
      run.status.updatedAt = new Date().toISOString();
    } catch (err) {
      output[i] = [raw];
      run.status.lastErrorRow = rowNum;
      run.status.lastErrorMessage = err && err.message ? err.message : String(err);
      run.status.processedCount++;
      run.status.lastProcessedRow = rowNum;
      run.status.updatedAt = new Date().toISOString();
    }
  }

  if (changed) {
    range.setValues(output.slice(0, outputLen));
  }
}

function normalizeJobUrlInternal_(input) {
  try {
    const raw = normalizeInputUrlString_(input);
    if (!raw) return "";

    const u = parseUrl_(raw);
    const host = canonicalHost_(u.hostname);
    const path = normalizePath_(u.pathname);

    const ruleResult = applyRuleSheet_(host, path, getRules_());
    const finalPath = ruleResult.finalPath || path || "/";

    // IMPORTANT: some job boards keep the real job ID only in query params.
    // This hardcoded safety layer fixes stale/missing URL_RULES and prevents
    // the global tracking cleanup from deleting identity params.
    const identityKeep = getIdentityKeepParams_(host, finalPath, u.searchParams);
    const keepOnlyParams = identityKeep || ruleResult.keepOnlyParams;

    const keptParams = filterParams_(u.searchParams, keepOnlyParams, ruleResult.dropParams);
    const keptFragment = getIdentityFragment_(host, finalPath, u.hash);

    return buildUrl_(host, finalPath, keptParams, keptFragment);
  } catch (e) {
    console.warn("NORMALIZE_JOB_URL failed:", e && e.message ? e.message : e, input);
    return String(input == null ? "" : input).trim();
  }
}

function applyRuleSheet_(host, path, rules) {
  // Fast path: exact host rules are stored in a hash map, so common ATS/job-board
  // domains avoid scanning the full regex list for every URL.
  const exactRules = rules && rules.exactByHost ? rules.exactByHost[host] : null;
  const exactMatch = exactRules ? applyRuleList_(host, path, exactRules, false) : null;
  if (exactMatch) return exactMatch;

  // Fallback path: regex host rules, wildcard rules, and subdomain patterns.
  const regexMatch = applyRuleList_(host, path, rules && rules.regexRules ? rules.regexRules : [], true);
  if (regexMatch) return regexMatch;

  return { finalPath: path, keepOnlyParams: null, dropParams: null };
}

function applyRuleList_(host, path, ruleList, checkHostRegex) {
  for (let i = 0; i < ruleList.length; i++) {
    const rule = ruleList[i];
    if (!rule.enabled) continue;
    if (checkHostRegex && !rule.hostRe.test(host)) continue;
    if (rule.pathRe && !rule.pathRe.test(path)) continue;

    const finalPath = rule.pathReplacement && rule.pathRe
      ? path.replace(rule.pathRe, rule.pathReplacement)
      : path;

    return {
      finalPath: finalPath,
      keepOnlyParams: rule.keepParamSet || null,
      dropParams: rule.dropParamSet || null
    };
  }

  return null;
}

function getRules_() {
  if (MEMORY_RULES_CACHE) return MEMORY_RULES_CACHE;

  let rawRules = readRulesFromScriptCache_();

  if (!rawRules) {
    rawRules = readRulesFromProperties_();

    if (rawRules && rawRules.length) {
      writeRulesToScriptCache_(rawRules);
    }
  }

  if (!rawRules || !rawRules.length) {
    throw new Error(
      "Rules snapshot missing. Run UPDATE_RULES_SNAPSHOT_FROM_SHEET() first."
    );
  }

  MEMORY_RULES_CACHE = buildRuntimeRules_(rawRules);
  return MEMORY_RULES_CACHE;
}


function buildRuntimeRules_(rawRules) {
  const exactByHost = Object.create(null);
  const regexRules = [];

  for (let i = 0; i < rawRules.length; i++) {
    const raw = rawRules[i];
    const rule = deserializeRule_(raw);

    if (raw.exactHost) {
      if (!exactByHost[raw.exactHost]) exactByHost[raw.exactHost] = [];
      exactByHost[raw.exactHost].push(rule);
    } else {
      regexRules.push(rule);
    }
  }

  return {
    exactByHost: exactByHost,
    regexRules: regexRules
  };
}

function readRulesFromScriptCache_() {
  try {
    const cached = CacheService.getScriptCache().get(CONFIG.RULES_CACHE_KEY);
    return cached ? JSON.parse(cached) : null;
  } catch (e) {
    return null;
  }
}

function writeRulesToScriptCache_(rawRules) {
  try {
    const payload = JSON.stringify(rawRules);
    // Apps Script cache values have a size limit. Silently skip if too large.
    if (payload.length < 95000) {
      CacheService.getScriptCache().put(CONFIG.RULES_CACHE_KEY, payload, CONFIG.CACHE_TTL_SECONDS);
    }
  } catch (e) {
    // Ignore cache write failures. Sheet remains source of truth.
  }
}

function readRulesFromProperties_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const baseKey = CONFIG.RULES_SNAPSHOT_PROPERTY_KEY;

    const chunkCount = Number(props.getProperty(baseKey + "_chunk_count")) || 0;

    if (chunkCount > 0) {
      let payload = "";

      for (let i = 0; i < chunkCount; i++) {
        const part = props.getProperty(baseKey + "_" + i);

        if (part == null) {
          throw new Error("Missing rules snapshot chunk: " + i);
        }

        payload += part;
      }

      return JSON.parse(payload);
    }

    // backward compatibility with old single-property snapshot
    const raw = props.getProperty(baseKey);
    return raw ? JSON.parse(raw) : null;

  } catch (e) {
    return null;
  }
}

function readRulesFromSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.RULES_SHEET_NAME);
  if (!sh) return [];

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const data = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = data[0].map(v => String(v).toLowerCase().trim());
  const idx = name => header.indexOf(String(name).toLowerCase().trim());

  const iEnabled = idx(CFG.HEADERS.URL_RULES.ENABLED);
  const iHost = idx(CFG.HEADERS.URL_RULES.HOST_REGEX);
  const iPath = idx(CFG.HEADERS.URL_RULES.PATH_REGEX);
  const iRepl = idx(CFG.HEADERS.URL_RULES.PATH_REPLACEMENT);
  const iKeep = idx(CFG.HEADERS.URL_RULES.KEEP_PARAMS);
  const iDrop = idx(CFG.HEADERS.URL_RULES.DROP_PARAMS);

  const rawRules = [];

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const hostRegex = iHost >= 0 ? String(row[iHost] || "").trim() : "";
    if (!hostRegex) continue;

    rawRules.push({
      enabled: parseBoolean_(iEnabled >= 0 ? row[iEnabled] : true),
      hostRegex: hostRegex,
      exactHost: exactHostFromAnchoredRegex_(hostRegex),
      pathRegex: iPath >= 0 ? String(row[iPath] || "").trim() || null : null,
      pathReplacement: iRepl >= 0 ? String(row[iRepl] || "").trim() || null : null,
      keepParams: iKeep >= 0 ? splitCsv_(row[iKeep]) : [],
      dropParams: iDrop >= 0 ? splitCsv_(row[iDrop]) : []
    });
  }

  return rawRules;
}

function exactHostFromAnchoredRegex_(hostRegex) {
  const s = String(hostRegex || "").trim();
  if (!/^\^[A-Za-z0-9\\.\-]+\$$/.test(s)) return "";

  const body = s.slice(1, -1);
  if (/\\[dwsDWS]|\[|\]|\(|\)|\*|\+|\?|\{|\}|\||\//.test(body)) return "";

  return canonicalHost_(body.replace(/\\\./g, "."));
}

function deserializeRule_(r) {
  const keepParams = r.keepParams || [];
  const dropParams = r.dropParams || [];

  return {
    enabled: !!r.enabled,
    hostRe: safeRegExp_(r.hostRegex, "i"),
    pathRe: r.pathRegex ? safeRegExp_(r.pathRegex, "i") : null,
    pathReplacement: r.pathReplacement || null,
    keepParams: keepParams,
    dropParams: dropParams,
    keepParamSet: keepParams.length ? toLowerSet_(keepParams) : null,
    dropParamSet: dropParams.length ? toLowerSet_(dropParams) : null,
    notes: r.notes || ""
  };
}

function toLowerSet_(values) {
  const set = new Set();
  for (let i = 0; i < values.length; i++) {
    set.add(String(values[i]).toLowerCase());
  }
  return set;
}

function safeRegExp_(pattern, flags) {
  try {
    return new RegExp(pattern, flags || "");
  } catch (e) {
    console.warn("Invalid regex ignored:", pattern, e && e.message ? e.message : e);
    return /a^/;
  }
}

function getIdentityKeepParams_(host, path, searchParams) {
  const h = canonicalHost_(host);
  const p = normalizePath_(path);

  for (let i = 0; i < IDENTITY_QUERY_KEEP_RULES.length; i++) {
    const rule = IDENTITY_QUERY_KEEP_RULES[i];

    if (rule.host !== h) continue;
    if (!rule.pathRe.test(p)) continue;

    return new Set(rule.keep);
  }

  if (
    GENERIC_ASHBY_PATH_RE.test(p) &&
    searchParams &&
    searchParams.get &&
    searchParams.get("ashby_jid")
  ) {
    return new Set(["ashby_jid"]);
  }

  return null;
}

function filterParams_(searchParams, keepOnlyLowerSet, dropLowerSet) {
  const kept = makeParams_();
  if (!searchParams || !searchParams.items || !searchParams.items.length) return kept;

  const entries = searchParams.items;

  for (let i = 0; i < entries.length; i++) {
    const k = entries[i][0];
    const v = entries[i][1];
    const keyLower = String(k).toLowerCase();

    // IMPORTANT:
    // Some ATS pages store the real job identity in query params, e.g.
    // LinkedIn currentJobId, Paymentology uid/coref, Ashby ashby_jid.
    // If a rule has keep_params, those params must survive even if their names
    // look like tracking/source params. Explicit drop_params still wins.
    if (keepOnlyLowerSet) {
      if (!keepOnlyLowerSet.has(keyLower)) continue;
      if (dropLowerSet && dropLowerSet.has(keyLower)) continue;
      kept.append(k, v);
      continue;
    }

    if (TRACKING_PARAMS_EXACT.has(keyLower)) continue;
    if (startsWithAny_(keyLower, TRACKING_PARAM_PREFIXES)) continue;
    if (dropLowerSet && dropLowerSet.has(keyLower)) continue;

    kept.append(k, v);
  }

  return kept;
}

function startsWithAny_(s, prefixes) {
  for (let i = 0; i < prefixes.length; i++) {
    if (s.indexOf(prefixes[i]) === 0) return true;
  }
  return false;
}

function normalizeInputUrlString_(input) {
  if (input == null) return "";

  let s = String(input).trim();
  if (!s) return "";

  s = s.replace(/[\u200B-\u200D\uFEFF]/g, "");
  s = s.replace(/^<|>$/g, "");
  s = s.replace(/^['"]|['"]$/g, "");
  s = s.trim();

  if (/\s/.test(s)) s = s.replace(/\s+/g, "");

  if (!/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//.test(s)) {
    if (/^[\w.-]+\.[a-z]{2,}(\/|\?|#|$)/i.test(s)) s = "https://" + s;
  }

  return s;
}

function parseUrl_(raw) {
  const s = normalizeInputUrlString_(raw);
  const m = String(s).match(/^([a-zA-Z][a-zA-Z0-9+.-]*):\/\/([^\/?#]+)([^?#]*)(\?[^#]*)?(#.*)?$/);
  if (!m) throw new Error("Invalid URL");

  const authority = m[2] || "";
  const hostOnly = authority.split("@").pop().split(":")[0];
  const pathname = m[3] || "/";
  const search = m[4] ? m[4].slice(1) : "";

  return {
    href: s,
    protocol: m[1] + ":",
    hostname: hostOnly,
    pathname: pathname || "/",
    search: search ? "?" + search : "",
    hash: m[5] || "",
    searchParams: makeParams_(search)
  };
}

function canonicalHost_(hostname) {
  return String(hostname || "").toLowerCase().replace(/^www\./, "");
}

function normalizePath_(pathname) {
  let p = pathname || "/";

  // Remove duplicate ending slashes, but keep root as "/".
  // IMPORTANT: this function must return p. Without this return, every
  // normal URL path becomes undefined and buildUrl_() falls back to root.
  p = p.replace(/\/+$/, "");
  if (!p) p = "/";

  return p;
}

function buildUrl_(hostname, pathname, params, fragment) {
  let host = canonicalHost_(hostname);
  let path = pathname || "/";
  if (path.charAt(0) !== "/") path = "/" + path;

  const q = params && params.toString ? params.toString() : "";
  return "https://" + host + path + (q ? "?" + q : "") + (fragment || "");
}

function makeParams_(input) {
  return new SimpleSearchParams_(input);
}

function SimpleSearchParams_(input) {
  this.items = [];

  if (!input) return;

  if (input instanceof SimpleSearchParams_) {
    this.items = input.items.map(x => [x[0], x[1]]);
    return;
  }

  if (typeof input === "string") {
    let q = input.charAt(0) === "?" ? input.slice(1) : input;
    if (!q) return;

    const parts = q.split("&");
    for (let i = 0; i < parts.length; i++) {
      if (!parts[i]) continue;

      const eq = parts[i].indexOf("=");
      const k = eq >= 0 ? parts[i].slice(0, eq) : parts[i];
      const v = eq >= 0 ? parts[i].slice(eq + 1) : "";
      this.append(decodeParam_(k), decodeParam_(v));
    }
  }
}

SimpleSearchParams_.prototype.append = function (k, v) {
  this.items.push([String(k), String(v == null ? "" : v)]);
};

SimpleSearchParams_.prototype.set = function (k, v) {
  const key = String(k);
  this.items = this.items.filter(x => x[0] !== key);
  this.append(key, v);
};

SimpleSearchParams_.prototype.get = function (k) {
  const key = String(k);
  for (let i = 0; i < this.items.length; i++) {
    if (this.items[i][0] === key) return this.items[i][1];
  }
  return null;
};

SimpleSearchParams_.prototype.entries = function () {
  return this.items.map(x => [x[0], x[1]]);
};

SimpleSearchParams_.prototype.toString = function () {
  const out = [];
  for (let i = 0; i < this.items.length; i++) {
    out.push(encodeParam_(this.items[i][0]) + "=" + encodeParam_(this.items[i][1]));
  }
  return out.join("&");
};

function decodeParam_(s) {
  try {
    return decodeURIComponent(String(s).replace(/\+/g, " "));
  } catch (e) {
    return String(s);
  }
}

function encodeParam_(s) {
  return encodeURIComponent(String(s)).replace(/%20/g, "+");
}

function expandUrlOnceCached_(rawUrl) {
  try {
    const url = normalizeInputUrlString_(rawUrl);
    if (!url) return rawUrl;

    const cache = CacheService.getScriptCache();
    const key = CONFIG.REDIRECT_CACHE_PREFIX + digestKey_(url);
    const cached = cache.get(key);
    if (cached) return cached;

    if (REDIRECTS_USED_THIS_RUN >= CONFIG.MAX_REDIRECTS_PER_RUN) return rawUrl;
    REDIRECTS_USED_THIS_RUN++;

    const resp = UrlFetchApp.fetch(url, {
      followRedirects: true,
      muteHttpExceptions: true,
      headers: { "User-Agent": "Mozilla/5.0" }
    });

    const finalUrl = resp.getFinalUrl() || rawUrl;
    cache.put(key, finalUrl, CONFIG.REDIRECT_CACHE_TTL_SECONDS);
    return finalUrl;
  } catch (e) {
    console.warn("Redirect expansion failed:", e && e.message ? e.message : e, rawUrl);
    return rawUrl;
  }
}

function digestKey_(s) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(bytes).slice(0, 40);
}

function CLEAR_URL_NORMALIZER_CACHE() {
  MEMORY_RULES_CACHE = null;
  MEMORY_HEADER_COL_CACHE = Object.create(null);

  try {
    CacheService.getScriptCache().remove(CONFIG.RULES_CACHE_KEY);
  } catch (e) { }
}

function getHeaderColumn_(sh, headerName) {
  const cacheKey = sh.getSheetId() + ":" + headerName;
  if (MEMORY_HEADER_COL_CACHE[cacheKey]) return MEMORY_HEADER_COL_CACHE[cacheKey];

  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return 0;

  const headers = sh.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const target = String(headerName).trim().toLowerCase();

  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i]).trim().toLowerCase() === target) {
      MEMORY_HEADER_COL_CACHE[cacheKey] = i + 1;
      return i + 1;
    }
  }

  return 0;
}

function buildRowsToProcess_(ss, sh, lastRow, resume, saved) {
  const rows = [];

  if (resume) {
    const startRow = Math.max(CONFIG.HEADER_ROW + 1, Number((saved || {}).lastProcessedRow || CONFIG.HEADER_ROW) + 1);
    for (let r = startRow; r <= lastRow; r++) {
      if (isRowVisible_(sh, r)) rows.push(r);
    }
    return rows;
  }

  const selectedRows = getSelectedVisibleRowNumbers_(ss, sh);
  if (selectedRows.length) return selectedRows;

  for (let r = CONFIG.HEADER_ROW + 1; r <= lastRow; r++) {
    if (isRowVisible_(sh, r)) rows.push(r);
  }

  return rows;
}

function getSelectedVisibleRowNumbers_(ss, sh) {
  const rl = ss.getActiveRangeList();
  const ranges = rl ? rl.getRanges() : [];
  if (!ranges || !ranges.length) return [];

  const rows = new Set();
  const sheetId = sh.getSheetId();

  for (let i = 0; i < ranges.length; i++) {
    const r = ranges[i];
    if (r.getSheet().getSheetId() !== sheetId) continue;

    const start = Math.max(r.getRow(), CONFIG.HEADER_ROW + 1);
    const end = start + r.getNumRows() - 1;

    for (let row = start; row <= end; row++) {
      if (isRowVisible_(sh, row)) rows.add(row);
    }
  }

  return Array.from(rows).sort((a, b) => a - b);
}

function isRowVisible_(sh, row) {
  try {
    if (sh.isRowHiddenByUser(row)) return false;
    if (typeof sh.isRowHiddenByFilter === "function" && sh.isRowHiddenByFilter(row)) return false;
    return true;
  } catch (e) {
    return true;
  }
}

function groupContiguous_(sortedRowsAsc) {
  if (!sortedRowsAsc.length) return [];

  const out = [];
  let start = sortedRowsAsc[0];
  let prev = sortedRowsAsc[0];

  for (let i = 1; i < sortedRowsAsc.length; i++) {
    const cur = sortedRowsAsc[i];

    if (cur === prev + 1) {
      prev = cur;
    } else {
      out.push({ start: start, end: prev });
      start = cur;
      prev = cur;
    }
  }

  out.push({ start: start, end: prev });
  return out;
}

function createStatus_(state, source, sheetName, jobUrlCol, expandRedirects, chunkSize, totalRowsPlanned, saved) {
  const now = new Date().toISOString();

  return {
    state: state,
    source: source,
    sheetName: sheetName,
    jobUrlColumn: jobUrlCol,
    expandRedirects: !!expandRedirects,
    chunkSize: chunkSize,
    totalRowsPlanned: totalRowsPlanned || 0,
    processedCount: 0,
    replacedCount: 0,
    lastProcessedRow: saved ? Number(saved.lastProcessedRow || 0) : 0,
    lastReplacedRow: saved ? Number(saved.lastReplacedRow || 0) : 0,
    lastErrorRow: "",
    lastErrorMessage: "",
    startedAt: now,
    updatedAt: now
  };
}


function maybeToastProgress_(ss, status) {
  const total = Number(status.totalRowsPlanned || 0);
  const processed = Number(status.processedCount || 0);
  if (!total || !processed) return;

  if (processed % Math.max(CONFIG.STATUS_UPDATE_EVERY_ROWS, status.chunkSize || 1) !== 0) return;

  const pct = Math.min(100, Math.round((processed / total) * 100));
  ss.toast(
    "Progress " + pct + "% | Last row: " + (status.lastProcessedRow || "-") +
    " | Replaced: " + (status.replacedCount || 0) +
    " | Redirects: " + (REDIRECTS_USED_THIS_RUN || 0),
    "URL Normalizer",
    3
  );
}

function stopBeforeTimeout_(ss, status) {
  status.state = CFG.VALUES.RUN_STATE.STOPPED_BEFORE_TIMEOUT;
  status.updatedAt = new Date().toISOString();
  saveStatus_(status);
  writeStatusSheet_(status);

  ss.toast(
    "Stopped before timeout. Last processed row: " + status.lastProcessedRow +
    " | Last replaced row: " + (status.lastReplacedRow || "-") +
    ". Run RESUME_NORMALIZE().",
    "URL Normalizer",
    10
  );
}

function isBlank_(value) {
  return value == null || String(value).trim() === "";
}

function clampInt_(value, min, max, fallback) {
  const n = parseInt(value, 10);
  if (isNaN(n)) return fallback;
  return Math.max(min, Math.min(max, n));
}

function splitCsv_(cell) {
  const s = cell == null ? "" : String(cell).trim();
  if (!s) return [];
  return s.split(",").map(x => x.trim()).filter(Boolean);
}

function parseBoolean_(value) {
  if (value === true) return true;
  const s = String(value || "").toLowerCase().trim();
  return s === "true" || s === "yes" || s === "1" || s === "y";
}

function saveStatus_(status) {
  PropertiesService.getScriptProperties().setProperty(
    CONFIG.NORMALIZER_RUNTIME_STATE_KEY,
    JSON.stringify(status)
  );
}

function getSavedStatus_() {
  const raw = PropertiesService.getScriptProperties().getProperty(CONFIG.NORMALIZER_RUNTIME_STATE_KEY);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function writeStatusSheet_(status) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(CONFIG.STATUS_SHEET_NAME);

  if (!sh) {
    try {
      sh = ss.insertSheet(CONFIG.STATUS_SHEET_NAME);
    } catch (e) {
      // Keep property status even if sheet creation is blocked.
      console.error("Cannot create status sheet:", e && e.stack ? e.stack : e);
      ss.toast(
        "Status saved in Script Properties, but cannot create STATUS sheet. Run from menu/editor, not as formula.",
        "URL Normalizer",
        10
      );
      return;
    }
  }

  const total = Number(status.totalRowsPlanned || 0);
  const processed = Number(status.processedCount || 0);
  const percent = total ? Math.min(100, Math.round((processed / total) * 10000) / 100) + "%" : "";

  const k = CFG.STATUS_KEYS.NORMALIZE;
  const rows = [
    [CFG.HEADERS.STATUS.NORMALIZE_TITLE, CFG.HEADERS.STATUS.VALUE],
    [k.STATE, status.state || ""],
    [k.SOURCE, status.source || ""],
    [k.SHEET, status.sheetName || ""],
    [k.JOB_URL_COLUMN, status.jobUrlColumn || ""],
    [k.EXPAND_REDIRECTS, status.expandRedirects ? CFG.VALUES.BOOLEAN_TEXT.TRUE : CFG.VALUES.BOOLEAN_TEXT.FALSE],
    [k.REDIRECTS_USED_THIS_RUN, REDIRECTS_USED_THIS_RUN || 0],
    [k.CHUNK_SIZE, status.chunkSize || ""],
    [k.TOTAL_ROWS_PLANNED, total],
    [k.PROCESSED_COUNT, processed],
    [k.PROGRESS, percent],
    [k.REPLACED_COUNT, status.replacedCount || 0],
    [k.LAST_PROCESSED_ROW, status.lastProcessedRow || ""],
    [k.LAST_REPLACED_ROW, status.lastReplacedRow || ""],
    [k.LAST_ERROR_ROW, status.lastErrorRow || ""],
    [k.LAST_ERROR_MESSAGE, status.lastErrorMessage || ""],
    [k.STARTED_AT, status.startedAt || ""],
    [k.UPDATED_AT, status.updatedAt || ""]
  ];

  // Only clear the normalize status block (A:B). Keep column C as spacer and D:E for job mapping.
  const clearRows = Math.max(sh.getLastRow(), rows.length, 1);
  sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.NORMALIZE.START_COL, clearRows, CFG.STATUS_LAYOUT.NORMALIZE.WIDTH).clearContent();
  sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.NORMALIZE.START_COL, rows.length, CFG.STATUS_LAYOUT.NORMALIZE.WIDTH).setValues(rows);
  sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.NORMALIZE.SPACER_COL).setValue("");
  sh.setFrozenRows(1);
  sh.autoResizeColumns(CFG.STATUS_LAYOUT.NORMALIZE.START_COL, CFG.STATUS_LAYOUT.NORMALIZE.WIDTH);
}

function TEST_ONE_URL() {
  Logger.log(NORMALIZE_JOB_URL("https://join.com/companies/akveloncom/16244954-middle-devops-engineer-poland-serbia?pid=e65242534431eadcb0c9"));
}


const IDENTITY_FRAGMENT_KEEP_RULES = [
  {
    hostRe: /.*/,
    pathRe: /\/(career|careers|job|jobs|position|positions|opening|openings)/i,
    fragmentRe: /^(?:job|jobs|position|opening|role|vacancy)[-_](?:\d+|[A-Za-z0-9][A-Za-z0-9-]{2,})$/i
  }
];

function getIdentityFragment_(host, path, hash) {
  const h = canonicalHost_(host);
  const p = normalizePath_(path);
  const fragment = String(hash || "").replace(/^#/, "");

  if (!fragment) return "";

  for (let i = 0; i < IDENTITY_FRAGMENT_KEEP_RULES.length; i++) {
    const rule = IDENTITY_FRAGMENT_KEEP_RULES[i];

    if (!rule.hostRe.test(h)) continue;
    if (!rule.pathRe.test(p)) continue;
    if (!rule.fragmentRe.test(fragment)) continue;

    return "#" + fragment;
  }

  return "";
}

function validateRuntimeRules_(rules) {
  rules.forEach((r, i) => {
    const row = i + 2;

    if (!r.hostRegex) {
      throw new Error("URL_RULES row " + row + ": missing host_regex");
    }

    try {
      new RegExp(r.hostRegex);
    } catch (e) {
      throw new Error("URL_RULES row " + row + ": invalid host_regex");
    }

    if (r.pathRegex) {
      try {
        new RegExp(r.pathRegex);
      } catch (e) {
        throw new Error("URL_RULES row " + row + ": invalid path_regex");
      }
    }
  });
}

function writeRulesToPropertiesStrict_(rawRules) {
  const props = PropertiesService.getScriptProperties();
  const baseKey = CONFIG.RULES_SNAPSHOT_PROPERTY_KEY;
  const payload = JSON.stringify(rawRules);

  const CHUNK_SIZE = 80000;
  const chunkCount = Math.ceil(payload.length / CHUNK_SIZE);

  const oldChunkCount = Number(props.getProperty(baseKey + "_chunk_count")) || 0;

  // Delete old single-property snapshot only.
  props.deleteProperty(baseKey);

  // Delete old chunked snapshot only.
  for (let i = 0; i < oldChunkCount; i++) {
    props.deleteProperty(baseKey + "_" + i);
  }

  // Write new chunks.
  for (let i = 0; i < chunkCount; i++) {
    props.setProperty(
      baseKey + "_" + i,
      payload.slice(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE)
    );
  }

  props.setProperty(baseKey + "_chunk_count", String(chunkCount));
}

function writeRulesMeta_(rawRules) {
  const payload = JSON.stringify(rawRules);

  const meta = {
    version: CONFIG.RULES_CACHE_KEY,
    count: rawRules.length,
    payloadChars: payload.length,
    chunkSize: 80000,
    chunks: Math.ceil(payload.length / 80000),
    updatedAt: new Date().toISOString(),
    updatedBy: Session.getActiveUser().getEmail() || "unknown"
  };

  PropertiesService
    .getScriptProperties()
    .setProperty("url_rules_snapshot_meta", JSON.stringify(meta));
}
