/********************
 * GLOBAL ON EDIT ENTRYPOINT
 * Keep only one onEdit(e) in the whole project.
 ********************/

function onEdit(e) {
  try {
    routingOnEdit_(e);
  } catch (err) {
    console.error(
      "onEdit error:",
      err && err.stack ? err.stack : err
    );
  }
}

/********************
 * ON EDIT ROUTER
 * Responsibility: decide only.
 * No business logic here.
 ********************/

function routingOnEdit_(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  const row = e.range.getRow();
  const a1 = e.range.getA1Notation();

  if (isJobSheetEdit_(sheetName, row)) {
    handleJobSheetEdit_(e);
    return;
  }
  
  if (isDashboardChartEdit_(sheetName, a1)) {
    handleChartEdit_(e);
    return;
  }
  
  if (isUrlRulesEdit_(sheetName)) {
    handleUrlRulesEdit_(e);
    return;
  }

  if (isCustomEdit_(sheetName, a1)) {
    handleCustomEdit_(e);
    return;
  }
}

/********************
 * ROUTE CONDITIONS
 ********************/

function isUrlRulesEdit_(sheetName) {
  return sheetName === CFG.URL_NORMALIZER.RULES_SHEET_NAME;
}

function isJobSheetEdit_(sheetName, row) {
  return AUTO_NORMALIZE_SHEET_SET.has(sheetName) &&
    row > CFG.URL_NORMALIZER.HEADER_ROW;
}

function isDashboardChartEdit_(sheetName, a1) {
  if (sheetName !== CFG.CHARTS.DASHBOARD_SHEET) return false;

  return Object.values(CFG.CHARTS.TITLE_RULES)
    .some(rule => (rule.watchCells || []).includes(a1));
}

function isCustomEdit_(sheetName, a1) {
  return false;
}