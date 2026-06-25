function handleChartEdit_(e) {
  const a1 = e.range.getA1Notation();

  const rules = Object.values(CFG.CHARTS.TITLE_RULES)
    .filter(rule => (rule.watchCells || []).includes(a1));

  if (!rules.length) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  rules.forEach(rule => {
    try {
      updateChartTitle_(rule);
    } catch (err) {
      console.error(
        `Chart title update failed for ${rule.chart}:`,
        err && err.stack ? err.stack : err
      );
    }
  });
}


function updateChartTitle_(config) {
  if (!config || !config.title || !config.sheet) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const title = ss.getRange(config.title).getDisplayValue();
  if (!title) return;

  const sheet = ss.getSheetByName(config.sheet);
  if (!sheet) {
    ss.toast(`Sheet not found: ${config.sheet}`, "Error", 5);
    return;
  }

  const charts = sheet.getCharts();
  const chart = charts[config.chartIndex];

  if (!chart) {
    ss.toast(`Chart index not found: ${config.chartIndex}`,"Error",3);
    return;
  }

  const updatedChart = chart.modify().setOption("title", title).build();

  sheet.updateChart(updatedChart);

  ss.toast(`Updated Chart ${config.chartIndex}: ${title}`,"Done",3);
}