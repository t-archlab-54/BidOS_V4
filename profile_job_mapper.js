function mapSelectedVisibleJobsToProfiles() {
  assertNoFreshRunningJobMap_();
  discardOldJobMapStateIfAny_();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const range = sheet.getActiveRange();

  if (!range) {
    throw new Error("Please select rows first.");
  }

  const allowedSheets = new Set(CFG.SHEETS.JOB_FALLBACK);

  if (!allowedSheets.has(sheet.getName())) {
    throw new Error(
      "Invalid job sheet. Allowed sheets: " +
      Array.from(allowedSheets).join(", ")
    );
  }

  const state = {
    jobSheetName: sheet.getName(),
    startRow: range.getRow(),
    numRows: range.getNumRows(),
    nextOffset: 0,
    chunkSize: CFG.MAPPER_CHUNK_SIZE,
    visibleJobs: 0,
    added: 0,
    dupes: 0,
    noMatch: 0,
    errors: 0,
    status: CFG.VALUES.RUN_STATE.RUNNING,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    leaseUntil: Date.now() + CFG.MAPPER_LEASE_MS
  };
  
  state.leaseUntil = Date.now() + CFG.MAPPER_LEASE_MS;
  saveJobMapStateFast_(state);
  ensureJobMapStatusSheet_(ss);
  writeJobMapStatus_(ss, state);

  deleteJobMapTriggers_();
  processJobMapChunks_();
}

function RESUME_JOB_MAPPING() {
  const state = loadJobMapStateFast_();

  if (!state) {
    SpreadsheetApp.getActive().toast("No saved job mapping state found.", CFG.APP_NAME, 5);
    return;
  }

  if (state.status === CFG.VALUES.RUN_STATE.DONE) {
    SpreadsheetApp.getActive().toast("Job mapping already completed.", CFG.APP_NAME, 5);
    return;
  }

  processJobMapChunks_();
}

function processJobMapChunks_() {
  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(10000)) {
    return;
  }

  const started = Date.now();
  let lastStatusWriteAt = 0;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let state = loadJobMapStateFast_();
    if (!state || state.status === CFG.VALUES.RUN_STATE.DONE) return;

    state.leaseUntil = Date.now() + CFG.MAPPER_LEASE_MS;
    saveJobMapStateFast_(state);

    const sheet = ss.getSheetByName(state.jobSheetName);

    if (!sheet) {
      throw new Error("Job sheet not found: " + state.jobSheetName);
    }

    const lastCol = sheet.getLastColumn();

    const headers = normalizeHeaders_(
      sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    );

    const jobIdCol = getHeaderIndex_(headers, CFG.HEADERS.JOB.JOB_ID) + 1;

    if (!jobIdCol) {
      throw new Error("Missing required header: job_id");
    }

    const profiles = loadActiveProfiles_(ss);
    const regions = loadRegionGroups_(ss);
    const stackMap = loadStackFilterMap_(ss);


    prepareProfilesForMatching_(profiles, regions, stackMap);

    const existingCache = {};
    const profileSheetCache = {};

    const sheetTag = normalizeTag_(sheet.getName());

    while (state.nextOffset < state.numRows) {
      if (Date.now() - started > CFG.MAPPER_MAX_RUNTIME_MS) {
        state.updatedAt = new Date().toISOString();
        state.leaseUntil = Date.now() + CFG.MAPPER_LEASE_MS;

        saveJobMapStateFast_(state);
        writeJobMapStatus_(ss, state);
        scheduleJobMapResume_();

        ss.toast(
          "Job mapping paused. Resume trigger created.",
          CFG.APP_NAME,
          5
        );

        return;
      }

      const chunk = readVisibleJobRowsChunkFast_(sheet, state, headers, jobIdCol, lastCol, sheetTag, regions, stackMap);

      const pendingWrites = {};

      chunk.jobs.forEach(job => {
        let matchedProfiles = [];

        try {
          matchedProfiles = profiles.filter(profile =>
            isProfileJobMatch_(profile, job, regions)
          );
        } catch (err) {
          state.errors = (state.errors || 0) + 1;
          state.noMatch++;
          return;
        }

        if (!matchedProfiles.length) {
          state.noMatch++;
          return;
        }

        matchedProfiles.forEach(profile => {
          try {
            const profileSheetName = profile.sheet_name;

            if (!profileSheetCache[profileSheetName]) {
              profileSheetCache[profileSheetName] =
                getOrCreateProfileTab_(ss, profileSheetName);
            }

            if (!existingCache[profileSheetName]) {
              existingCache[profileSheetName] =
                loadExistingAppKeys_(profileSheetCache[profileSheetName]);
            }

            const appKey = `${profile.profile_id}|${job.job_id}`;

            if (existingCache[profileSheetName].has(appKey)) {
              state.dupes++;
              return;
            }

            if (!pendingWrites[profileSheetName]) {
              pendingWrites[profileSheetName] = [];
            }

            pendingWrites[profileSheetName].push(buildProfileOutputRow_(appKey, job.date, job.job_url, CFG.VALUES.PROFILE_OUTPUT_STATUS.TODO, "", ""));

            existingCache[profileSheetName].add(appKey);
            state.added++;

          } catch (err) {
            state.errors = (state.errors || 0) + 1;

            const safeSheetName =
              profile.sheet_name ||
              buildProfileSheetName_(
                profile.profile_name,
                profile.region_tags && profile.region_tags[0],
                profile.stack_tags && profile.stack_tags[0]
              );

            if (!profileSheetCache[safeSheetName]) {
              profileSheetCache[safeSheetName] =
                getOrCreateProfileTab_(ss, safeSheetName);
            }

            if (!pendingWrites[safeSheetName]) {
              pendingWrites[safeSheetName] = [];
            }

            pendingWrites[safeSheetName].push(buildProfileOutputRow_(
              `${profile.profile_id || "UNKNOWN"}|${job.job_id || "UNKNOWN"}`,
              job.date || "",
              job.job_url || "",
              CFG.VALUES.PROFILE_OUTPUT_STATUS.ERROR,
              String(err && err.message ? err.message : err),
              ""
            ));
          }
        });
      });

      Object.keys(pendingWrites).forEach(profileSheetName => {
        appendProfileRowsByHeaderFast_(
          profileSheetCache[profileSheetName],
          pendingWrites[profileSheetName]
        );
      });

      state.visibleJobs += chunk.jobs.length;
      state.updatedAt = new Date().toISOString();

      const shouldWriteStatus =
        Date.now() - lastStatusWriteAt >= 5000 ||
        state.nextOffset >= state.numRows;

      if (shouldWriteStatus) {
        state.leaseUntil = Date.now() + CFG.MAPPER_LEASE_MS;
        saveJobMapStateFast_(state);
        writeJobMapStatus_(ss, state);
        lastStatusWriteAt = Date.now();

        ss.toast(
          `Mapped: ${state.added}, Dupes: ${state.dupes}, No match: ${state.noMatch}, Errors: ${state.errors || 0}`,
          CFG.APP_NAME,
          3
        );
      }
    }

    state.status = CFG.VALUES.RUN_STATE.DONE;
    state.updatedAt = new Date().toISOString();

    state.leaseUntil = Date.now() + CFG.MAPPER_LEASE_MS;
    saveJobMapStateFast_(state);
    writeJobMapStatus_(ss, state);
    deleteJobMapTriggers_();

    ss.toast(
      `Job mapping done. Added: ${state.added}, Dupes: ${state.dupes}, No match: ${state.noMatch}, Errors: ${state.errors || 0}`,
      CFG.APP_NAME,
      8
    );
  } finally {
    lock.releaseLock();
  }
}

function readVisibleJobRowsChunkFast_(sheet, state, headers, jobIdCol, lastCol, sheetTag, regions, stackMap) {
  const jobs = [];
  const jobIdWrites = [];

  const take = Math.min(state.chunkSize, state.numRows - state.nextOffset);

  const startRow = state.startRow + state.nextOffset;

  const values = sheet.getRange(startRow, 1, take, lastCol).getValues();

  for (let i = 0; i < take; i++) {
    const rowNumber = startRow + i;

    state.nextOffset++;

    if (sheet.isRowHiddenByFilter(rowNumber)) { continue; }

    if (sheet.isRowHiddenByUser(rowNumber)) { continue; }

    const obj = rowToObject_(headers, values[i]);

    if (!getObjValue_(obj, CFG.HEADERS.JOB.JOB_URL)) { continue; }

    let jobId = String(getObjValue_(obj, CFG.HEADERS.JOB.JOB_ID) || "").trim();

    if (!jobId) {
      jobId = sheetTag + String(rowNumber).padStart(6, "0");
      jobIdWrites.push({ row: rowNumber, value: jobId });
    }

    const regionTags = parseTags_(getObjValue_(obj, CFG.HEADERS.JOB.REGION_TAGS) || sheetTag);
    const stackTags = parseTags_(getObjValue_(obj, CFG.HEADERS.JOB.STACK_TAGS) || CFG.DEFAULT_STACK);

    jobs.push({
      job_id: jobId,
      date: getObjValue_(obj, CFG.HEADERS.JOB.DATE) || "",
      job_url: String(getObjValue_(obj, CFG.HEADERS.JOB.JOB_URL)),
      region_tags: regionTags,
      stack_tags: stackTags,
      _expandedRegions: expandTags_(regionTags, regions || {}),
      _stackSet: expandStackTags_(stackTags, stackMap)
    });
  }

  writeSingleColumnRowUpdatesFast_(sheet, jobIdCol, jobIdWrites);

  return { jobs };
}

/********************
 * PROFILE / JOB SHEET HELPERS
 ********************/

function loadActiveProfiles_(ss) {
  const sh = ss.getSheetByName(CFG.PROFILE_SHEET_NAME);

  if (!sh) {
    throw new Error("Missing Profiles sheet.");
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return [];
  }

  const headers = normalizeHeaders_(values[0]);

  return values
    .slice(1)
    .map(row => {
      const obj = rowToObject_(headers, row);

      const profileName = getFirstObjValue_(obj, CFG.HEADERS.PROFILE.PROFILE_NAME_ALIASES) || "";

      const regionTag = getFirstObjValue_(obj, CFG.HEADERS.PROFILE.REGION_TAG_ALIASES) || "";

      const stackTag = getFirstObjValue_(obj, CFG.HEADERS.PROFILE.STACK_TAG_ALIASES) || CFG.DEFAULT_STACK;

      const profileId =
        String(getObjValue_(obj, CFG.HEADERS.PROFILE.PROFILE_ID) || profileName).trim();

      const sheetName =
        getObjValue_(obj, CFG.HEADERS.PROFILE.SHEET_NAME)
          ? String(getObjValue_(obj, CFG.HEADERS.PROFILE.SHEET_NAME)).trim()
          : buildProfileSheetName_(profileName, regionTag, stackTag);

      return {
        profile_id: profileId,
        profile_name: profileName,
        region_tags: parseTags_(regionTag),
        stack_tags: parseTags_(stackTag),
        active: isProfileActive_(obj),
        sheet_name: sheetName
      };
    })
    .filter(profile => profile.active && profile.profile_id);
}

function buildProfileSheetName_(profileName, regionTag, stackTag) {
  const firstWord = String(profileName || CFG.VALUES.PROFILE_DEFAULT_NAME).trim().split(/\s+/)[0] || CFG.VALUES.PROFILE_DEFAULT_NAME;

  const region = normalizeTag_(regionTag || CFG.VALUES.DEFAULT_REGION);
  const stack = normalizeTag_(stackTag || CFG.DEFAULT_STACK);

  return [firstWord, "(", region, "_", stack, ")"].join("").replace(",", "");
}

function getOrCreateProfileTab_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);

  if (!sh) {
    sh = ss.insertSheet(sheetName);

    sh.getRange(1, 1, 1, CFG.PROFILE_TAB_HEADERS.length)
      .setValues([CFG.PROFILE_TAB_HEADERS]);

    sh.setFrozenRows(1);
  }

  return sh;
}

function loadExistingAppKeys_(sheet) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return new Set();
  }

  const lastCol = sheet.getLastColumn();
  const headers = normalizeHeaders_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  const appKeyCol = getHeaderIndex_(headers, CFG.HEADERS.PROFILE_TAB.APP_KEY) + 1;

  if (!appKeyCol) {
    throw new Error("Missing required profile-tab header: " + CFG.HEADERS.PROFILE_TAB.APP_KEY);
  }

  const values = sheet
    .getRange(2, appKeyCol, lastRow - 1, 1)
    .getValues()
    .flat()
    .map(v => String(v).trim())
    .filter(Boolean);

  return new Set(values);
}

function loadRegionGroups_(ss) {
  const sh = ss.getSheetByName(CFG.TAGS_SHEET_NAME);

  if (!sh) {
    return {};
  }

  const values = sh.getDataRange().getValues();

  if (values.length < 2) {
    return {};
  }

  const headers = normalizeHeaders_(values[0]);
  const groupCol = getHeaderIndex_(headers, CFG.HEADERS.TAGS.GROUP);
  const tagCol = getHeaderIndex_(headers, CFG.HEADERS.TAGS.TAG);

  if (groupCol === -1 || tagCol === -1) {
    return {};
  }

  const groups = {};

  values.slice(1).forEach(row => {
    const group = normalizeTag_(row[groupCol]);
    const tag = normalizeTag_(row[tagCol]);

    if (!group || !tag) {
      return;
    }

    if (!groups[group]) {
      groups[group] = new Set();
    }

    groups[group].add(tag);
  });

  return groups;
}

function prepareProfilesForMatching_(profiles, regions, stackMap) {
  profiles.forEach(profile => {
    profile._expandedRegions = expandTags_(profile.region_tags, regions);
    profile._stackSet = expandStackTags_(profile.stack_tags, stackMap);
  });
}

function isProfileJobMatch_(profile, job, regions) {
  const pRegions = profile._expandedRegions || expandTags_(profile.region_tags, regions);
  const jRegions = job._expandedRegions || expandTags_(job.region_tags, regions);

  const pStacks = profile._stackSet || new Set(profile.stack_tags);
  const jStacks = job._stackSet || new Set(job.stack_tags);

  const regionOk =
    pRegions.size > 0 &&
    jRegions.size > 0 &&
    hasIntersection_(pRegions, jRegions);

  const stackOk =
    pStacks.size === 0 ||
    jStacks.size === 0 ||
    hasIntersection_(pStacks, jStacks);

  return regionOk && stackOk;
}

function expandTags_(tags, groups) {
  const result = new Set();

  tags.forEach(tag => {
    const normalized = normalizeTag_(tag);

    if (!normalized) {
      return;
    }

    result.add(normalized);

    if (groups[normalized]) {
      groups[normalized].forEach(child => result.add(child));
    }
  });

  return result;
}

function hasIntersection_(a, b) {
  for (const value of a) {
    if (b.has(value)) {
      return true;
    }
  }

  return false;
}

function buildProfileOutputRow_(appKey, date, jobUrl, status, error, jd) {
  const h = CFG.HEADERS.PROFILE_TAB;
  const row = {};

  row[normalizeHeaderName_(h.APP_KEY)] = appKey || "";
  row[normalizeHeaderName_(h.DATE)] = date || "";
  row[normalizeHeaderName_(h.JOB_URL)] = jobUrl || "";
  row[normalizeHeaderName_(h.STATUS)] = status || "";
  row[normalizeHeaderName_(h.BIDDER)] = "";
  row[normalizeHeaderName_(h.RESUME_URL)] = "";
  row[normalizeHeaderName_(h.FILE_NAME)] = "";
  row[normalizeHeaderName_(h.CREATED_AT)] = "";
  row[normalizeHeaderName_(h.PRICE)] = "";
  row[normalizeHeaderName_(h.ERROR)] = error || "";
  row[normalizeHeaderName_(h.JD)] = jd || "";

  return row;
}

function appendProfileRowsByHeaderFast_(sheet, rows) {
  if (!rows.length) {
    return;
  }

  ensureProfileTabHeaders_(sheet);

  const lastCol = sheet.getLastColumn();
  const headers = normalizeHeaders_(sheet.getRange(1, 1, 1, lastCol).getValues()[0]);
  const values = rows.map(rowObj => headers.map(header => rowObj[header] || ""));
  const startRow = Math.max(2, sheet.getLastRow() + 1);

  sheet.getRange(startRow, 1, values.length, lastCol).setValues(values);
}

function ensureProfileTabHeaders_(sheet) {
  const configuredHeaders = CFG.PROFILE_TAB_HEADERS;

  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, configuredHeaders.length).setValues([configuredHeaders]);
    sheet.setFrozenRows(1);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), configuredHeaders.length);
  const existing = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const normalizedExisting = normalizeHeaders_(existing);
  const missing = configuredHeaders.filter(header => !normalizedExisting.includes(normalizeHeaderName_(header)));

  if (missing.length) {
    sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }

  sheet.setFrozenRows(1);
}

function writeSingleColumnRowUpdatesFast_(sheet, col, updates) {
  if (!updates || !updates.length) {
    return;
  }

  updates.sort((a, b) => a.row - b.row);

  let blockStart = updates[0].row;
  let blockValues = [[updates[0].value]];
  let previousRow = updates[0].row;

  for (let i = 1; i < updates.length; i++) {
    const item = updates[i];

    if (item.row === previousRow + 1) {
      blockValues.push([item.value]);
    } else {
      sheet.getRange(blockStart, col, blockValues.length, 1).setValues(blockValues);
      blockStart = item.row;
      blockValues = [[item.value]];
    }

    previousRow = item.row;
  }

  sheet.getRange(blockStart, col, blockValues.length, 1).setValues(blockValues);
}

function saveJobMapStateFast_(state) {
  PropertiesService
    .getDocumentProperties()
    .setProperty(CFG.MAPPER_STATE_KEY, JSON.stringify(state));
}

function loadJobMapStateFast_() {
  const raw = PropertiesService
    .getDocumentProperties()
    .getProperty(CFG.MAPPER_STATE_KEY);

  return raw ? JSON.parse(raw) : null;
}

function SHOW_JOB_MAPPING_STATUS() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const state = loadJobMapStateFast_();

  ensureJobMapStatusSheet_(ss);

  if (state) {
    writeJobMapStatus_(ss, state);

    const runtimeMin = state.startedAt
      ? ((Date.now() - new Date(state.startedAt).getTime()) / 60000).toFixed(1)
      : "-";
    ss.toast(
      [
        `Status: ${state.status || "-"}`,
        `Sheet: ${state.jobSheetName || "-"}`,
        `Progress: ${state.nextOffset || 0}/${state.numRows || 0}`,
        `Visible: ${state.visibleJobs || 0}`,
        `Added: ${state.added || 0}`,
        `Dupes: ${state.dupes || 0}`,
        `No Match: ${state.noMatch || 0}`,
        `Errors: ${state.errors || 0}`,
        `Runtime: ${runtimeMin} min`
      ].join(" | "),
      CFG.APP_NAME + " Job Mapper",
      10
    );

  } else {

    ss.toast(
      "No saved job mapping state found.",
      CFG.APP_NAME + " Job Mapper",
      5
    );
  }

  Logger.log(state);
}

function RESET_JOB_MAPPING_STATUS() {
  deleteJobMapTriggers_();

  PropertiesService
    .getDocumentProperties()
    .deleteProperty(CFG.MAPPER_STATE_KEY);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CFG.MAPPER_STATUS_SHEET_NAME);

  if (sh) {
    const clearRows = Math.max(sh.getLastRow(), 1);
    sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.JOB_MAP.START_COL, clearRows, CFG.STATUS_LAYOUT.JOB_MAP.WIDTH).clearContent();
  }

  ss.toast("Job mapping status reset.", CFG.APP_NAME, 5);
}

function ensureJobMapStatusSheet_(ss) {
  let sh = ss.getSheetByName(CFG.MAPPER_STATUS_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(CFG.MAPPER_STATUS_SHEET_NAME);
  }

  // Job map status lives in D:E. Column C intentionally stays blank as a visual spacer.
  if (String(sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.JOB_MAP.START_COL).getValue()).trim() !== CFG.HEADERS.STATUS.JOB_MAP_TITLE) {
    sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.JOB_MAP.START_COL, 1, CFG.STATUS_LAYOUT.JOB_MAP.WIDTH).setValues([
      [CFG.HEADERS.STATUS.JOB_MAP_TITLE, CFG.HEADERS.STATUS.VALUE]
    ]);
  }

  return sh;
}

function writeJobMapStatus_(ss, state) {
  const sh = ensureJobMapStatusSheet_(ss);

  const k = CFG.STATUS_KEYS.JOB_MAP;
  const rows = [
    [k.STATUS, state.status],
    [k.JOB_SHEET_NAME, state.jobSheetName],
    [k.START_ROW, state.startRow],
    [k.NUM_ROWS, state.numRows],
    [k.NEXT_OFFSET, state.nextOffset],
    [k.CHUNK_SIZE, state.chunkSize],
    [k.VISIBLE_JOBS, state.visibleJobs],
    [k.ADDED, state.added],
    [k.DUPES, state.dupes],
    [k.NO_MATCH, state.noMatch],
    [k.ERRORS, state.errors || 0],
    [k.STARTED_AT, state.startedAt],
    [k.UPDATED_AT, state.updatedAt]
  ];

  // Only clear the job-map status block (D:E). Keep A:B for URL normalize status.
  const clearRows = Math.max(sh.getLastRow(), rows.length + 1, 1);
  sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.JOB_MAP.START_COL, clearRows, CFG.STATUS_LAYOUT.JOB_MAP.WIDTH).clearContent();

  sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.JOB_MAP.START_COL, 1, CFG.STATUS_LAYOUT.JOB_MAP.WIDTH).setValues([
    [CFG.HEADERS.STATUS.JOB_MAP_TITLE, CFG.HEADERS.STATUS.VALUE]
  ]);

  sh.getRange(2, CFG.STATUS_LAYOUT.JOB_MAP.START_COL, rows.length, CFG.STATUS_LAYOUT.JOB_MAP.WIDTH).setValues(rows);
  sh.getRange(CFG.STATUS_LAYOUT.HEADER_ROW, CFG.STATUS_LAYOUT.JOB_MAP.SPACER_COL).setValue("");
  sh.setFrozenRows(1);
  // sh.autoResizeColumns(CFG.STATUS_LAYOUT.JOB_MAP.START_COL, CFG.STATUS_LAYOUT.JOB_MAP.WIDTH);
}

function assertNoFreshRunningJobMap_() {
  const state = loadJobMapStateFast_();

  if (!state || state.status !== CFG.VALUES.RUN_STATE.RUNNING) return;

  const leaseUntil = Number(state.leaseUntil || 0);

  if (leaseUntil && leaseUntil > Date.now()) {
    throw new Error("Job mapping is already running. Please wait until it finishes.");
  }
}

function discardOldJobMapStateIfAny_() {
  const state = loadJobMapStateFast_();

  if (!state) return;

  const freshRunning =
    state.status === CFG.VALUES.RUN_STATE.RUNNING &&
    Number(state.leaseUntil || 0) > Date.now();

  if (freshRunning) return;

  deleteJobMapTriggers_();

  PropertiesService
    .getDocumentProperties()
    .deleteProperty(CFG.MAPPER_STATE_KEY);
}

function scheduleJobMapResume_() {
  deleteJobMapTriggers_();

  ScriptApp
    .newTrigger("RESUME_JOB_MAPPING")
    .timeBased()
    .after(CFG.JOB_MAPPER.RESUME_AFTER_MS)
    .create();
}

function deleteJobMapTriggers_() {
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === "RESUME_JOB_MAPPING") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

function normalizeHeaderName_(header) {
  return String(header || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_");
}

function getHeaderIndex_(normalizedHeaders, headerName) {
  return normalizedHeaders.indexOf(normalizeHeaderName_(headerName));
}

function getObjValue_(obj, headerName) {
  return obj[normalizeHeaderName_(headerName)];
}

function getFirstObjValue_(obj, headerNames) {
  for (let i = 0; i < headerNames.length; i++) {
    const value = getObjValue_(obj, headerNames[i]);
    if (value !== "" && value != null) return value;
  }
  return "";
}

function normalizeHeaders_(headers) {
  return headers.map(header =>
    String(header || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, "_")
  );
}

function isProfileActive_(obj) {
  const status = String(getFirstObjValue_(obj, CFG.HEADERS.PROFILE.ACTIVE_ALIASES) || CFG.VALUES.PROFILE_ACTIVE_DEFAULT)
    .trim()
    .toLowerCase();

  return !CFG.VALUES.PROFILE_INACTIVE.includes(status);
}


function rowToObject_(headers, row) {
  const obj = {};

  headers.forEach((header, index) => {
    obj[header] = row[index];
  });

  return obj;
}

function parseTags_(value) {
  return String(value || "").split(/[,\|;/]+/).map(normalizeTag_).filter(Boolean);
}

function normalizeTag_(value) {
  return String(value || "").trim().toUpperCase().replace(/\s+/g, "_");
}

function expandStackTags_(tags, stackMap) {

  if (!tags || !tags.length) { return new Set(); }

  const canonicalSet = CFG.CANONICAL_STACK_TAG_SET;

  const result = new Set();

  tags.forEach(tag => {

    const normalized = normalizeTag_(tag);

    if (!normalized) { return; }

    // already SWE/AI/QA

    if (canonicalSet.has(normalized)) {

      result.add(normalized);

      return;
    }

    // mapped ATS keyword

    if (stackMap && stackMap[normalized]) {

      result.add(stackMap[normalized]);

      return;

    }

    // unknown keyword

    result.add(normalized);

  });

  return result;

}

function loadStackFilterMap_(ss) {

  const sh =
    ss.getSheetByName(CFG.TAGS_SHEET_NAME);

  if (!sh) { return {}; }

  const values =
    sh.getRange(CFG.STACK_FILTER_TABLE_RANGE).getValues();

  const map = {};

  values.slice(1).forEach(row => {

    const outputStack = normalizeTag_(row[0]);

    if (!outputStack) { return; }

    // direct mapping

    map[outputStack] = outputStack;

    // keyword mapping

    row.slice(2).forEach(keyword => {

      const key = normalizeTag_(keyword);

      if (!key) { return; }

      map[key] = outputStack;

    });

  });

  return map;

}