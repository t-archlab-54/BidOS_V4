/**
 * url-rules-seeder.gs
 * Strict final version: seed/reset only; runtime reads URL_RULES sheet.
 * One-time URL_RULES installer/reset script.
 *
 * Keep this as a separate Apps Script file.
 * Run SETUP_URL_RULES_BASE() only when you want to create/reset URL_RULES.
 * Normalization does not read this file during normal operation.
 */

/**
 * Creates or refreshes URL_RULES from separated seed rules.
 * Run this only when you want to install/reset the base rule sheet.
 *
 * After this, maintain rules directly in URL_RULES.
 * URL_RULES is the single source of truth used by NORMALIZE_JOB_URL().
 */

// function SETUP_PROJECT() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   // ensureSlackSettingsSheet_();
//   SETUP_URL_RULES_BASE();
//   UPDATE_RULES_SNAPSHOT_FROM_SHEET();
//   // ensureScheduledReportsTrigger_();

//   ss.toast('Project setup complete. Only slack_bot_token must be set manually.', CFG.APP_NAME, 8);
// }

function CLEAR_URL_NORMALIZER_CACHE_SAFE_() {
  try {
    if (typeof CLEAR_URL_NORMALIZER_CACHE === "function") {
      CLEAR_URL_NORMALIZER_CACHE();
      return;
    }
  } catch (e) {
    console.warn("CLEAR_URL_NORMALIZER_CACHE failed:", e && e.message ? e.message : e);
  }

  // Fallback if runtime file is not loaded yet.
  try {
    if (typeof CacheService !== "undefined") {
      CacheService.getScriptCache().remove(CFG.URL_NORMALIZER.RULES_CACHE_KEY);
    }
  } catch (e) { }
}

function SETUP_URL_RULES_BASE() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Do not rely on CONFIG during install if runtime file failed to load.
  const rulesSheetName = getRulesSheetNameSafe_();
  let sh = ss.getSheetByName(rulesSheetName);
  if (!sh) sh = ss.insertSheet(rulesSheetName);

  const rows = URL_RULE_SEED_ROWS_();
  validateUrlRuleRows_(rows);

  try {
    // Safer than clear() because it preserves sheet permissions, filters, widths, formatting.
    sh.clearContents();

    // Remove old extra rows/columns only when safe. Avoid deleting all columns/rows.
    ensureSheetSize_(sh, rows.length, rows[0].length);

    // Chunked write avoids random Apps Script "unknown error" on larger rule sheets.
    writeRowsInChunks_(sh, rows, 50);

    try { sh.setFrozenRows(1); } catch (e) { }
    try { sh.autoResizeColumns(1, rows[0].length); } catch (e) { }

    CLEAR_URL_NORMALIZER_CACHE_SAFE_();
    ss.toast("URL_RULES base seed installed: " + (rows.length - 1) + " rules.", "URL Normalizer", 5);
  } catch (err) {
    logUrlRulesInstallError_("SETUP_URL_RULES_BASE", err);
    throw err;
  }
}

function REFRESH_URL_RULES_NOW() {
  CLEAR_URL_NORMALIZER_CACHE_SAFE_();
  SpreadsheetApp.getActive().toast("URL rules cache refreshed.", "URL Normalizer", 4);
}

function getRulesSheetNameSafe_() {
  try {
    if (typeof CONFIG !== "undefined" && CONFIG && CONFIG.RULES_SHEET_NAME) {
      return CONFIG.RULES_SHEET_NAME;
    }
  } catch (e) { }
  return CFG.SHEETS.URL_RULES;
}

function ensureSheetSize_(sh, requiredRows, requiredCols) {
  const maxRows = sh.getMaxRows();
  const maxCols = sh.getMaxColumns();

  if (maxRows < requiredRows) {
    sh.insertRowsAfter(maxRows, requiredRows - maxRows);
  }
  if (maxCols < requiredCols) {
    sh.insertColumnsAfter(maxCols, requiredCols - maxCols);
  }
}

function validateUrlRuleRows_(rows) {
  if (!Array.isArray(rows) || !rows.length || !Array.isArray(rows[0])) {
    throw new Error("URL_RULE_SEED_ROWS_() returned invalid rows.");
  }

  const width = rows[0].length;
  for (let i = 0; i < rows.length; i++) {
    if (!Array.isArray(rows[i])) {
      throw new Error("Invalid URL_RULES row at index " + i + ": row is not an array.");
    }
    if (rows[i].length !== width) {
      throw new Error("Invalid URL_RULES row width at sheet row " + (i + 1) + ": expected " + width + ", got " + rows[i].length);
    }
  }
}

function writeRowsInChunks_(sh, rows, chunkSize) {
  const width = rows[0].length;
  for (let start = 0; start < rows.length; start += chunkSize) {
    const chunk = rows.slice(start, start + chunkSize);
    sh.getRange(start + 1, 1, chunk.length, width).setValues(chunk);
    SpreadsheetApp.flush();
  }
}

function logUrlRulesInstallError_(where, err) {
  const msg = err && err.stack ? err.stack : String(err && err.message ? err.message : err);
  console.error(where + " failed:", msg);
  try {
    SpreadsheetApp.getActive().toast(where + " failed. Check Executions log.", "URL Normalizer", 10);
  } catch (e) { }
}

/**
 * Separated seed rules.
 * Keep this function separate from the normalizer logic.
 * The normalizer does NOT read this during normal operation.
 * It only reads the URL_RULES sheet.
 */
function URL_RULE_SEED_ROWS_() {
  return [
    [CFG.HEADERS.URL_RULES.ENABLED, CFG.HEADERS.URL_RULES.HOST_REGEX, CFG.HEADERS.URL_RULES.PATH_REGEX, CFG.HEADERS.URL_RULES.PATH_REPLACEMENT, CFG.HEADERS.URL_RULES.KEEP_PARAMS, CFG.HEADERS.URL_RULES.DROP_PARAMS, CFG.HEADERS.URL_RULES.NOTES],

    [true, "^boards\\.greenhouse\\.io$", "^/([^/]+)/jobs/(\\d+).*", "/$1/jobs/$2", "", "", "Greenhouse classic company/jobs/id"],
    [true, "^boards\\.greenhouse\\.io$", "^/embed/job_app.*", "/embed/job_app", "token", "", "Greenhouse embed keeps token"],
    [true, "^job-boards(\\.eu)?\\.greenhouse\\.io$", "^/([^/]+)/jobs/(\\d+).*", "/$1/jobs/$2", "", "", "Greenhouse job boards"],
    [true, "^job-boards\\.greenhouse\\.io$", "^/embed/job_app.*", "/embed/job_app", "token", "source,src,ref,gh_src", "Greenhouse job-boards embed keeps token"],
    [true, "^boards\\.greenhouse\\.io$", "^/.*$", "", "gh_jid", "", "Greenhouse gh_jid fallback only"],
    [true, "^job-boards(\\.eu)?\\.greenhouse\\.io$", "^/.*$", "", "gh_jid", "", "Greenhouse job-board gh_jid fallback only"],

    [true, "(^|\\.)lever\\.co$", "^/(.+?)(/apply)?$", "/$1", "", "jr_id,source,src,ref", "Lever remove apply/drop source"],

    [true, "^jobs\\.ashbyhq\\.com$", "^/([^/]+)/([^/?#]+)(/application)?$", "/$1/$2", "", "", "Ashby company/job"],
    [true, "^ashbyhq\\.com$", "^/careers/?$", "/careers", "ashby_jid", "", "Ashby main careers keeps ashby_jid"],

    [true, "^apply\\.workable\\.com$", "^/([^/]+)/j/([A-Z0-9]+).*", "/$1/j/$2", "", "", "Workable company/job code"],
    [true, "^jobs\\.workable\\.com$", "^/search.*", "/search", "selectedJobId", "location,query,day_range,workplace,employment_type,experience", "Workable search keeps selectedJobId"],

    [true, "^ats\\.rippling\\.com$", "^/([^/]+)/([^/]+)/jobs/([^/?#]+).*", "/$1/$2/jobs/$3", "", "can,st,jobSite,source,src,ref", "Rippling locale/company/jobs/id; drop source params"],
    [true, "^ats\\.rippling\\.com$", "^/([^/]+)/jobs/([^/?#]+).*", "/$1/jobs/$2", "", "can,st,jobSite,source,src,ref", "Rippling company/jobs/id without locale; drop source params"],

    [true, "^wd\\d+\\.myworkdayjobs\\.com$", "^/(.+?/job/.+?/[^/?#]+).*", "/$1", "", "jr_id,source,src,ref", "Workday clean job path"],
    [true, "^wd\\d+\\.myworkdayjobs\\.com$", "^/(.+)$", "/$1", "", "", "Workday wd fallback"],
    [true, "(^|\\.)myworkdayjobs\\.com$", "^/(.+)$", "/$1", "", "", "Workday keep path fallback"],

    [true, "^jobs\\.dayforcehcm\\.com$", "^/(.+?/jobs/[^/?#]+).*", "/$1", "", "pid,conversionId,fmt,source,src,ref", "Dayforce clean job path"],
    [true, "^jobs\\.dayforcehcm\\.com$", "^/(.+)$", "/$1", "", "pid,conversionId,fmt,source,src,ref", "Dayforce fallback"],

    [true, "^jobs\\.smartrecruiters\\.com$", "^/([^/]+)/([^/?#]+).*", "/$1/$2", "", "trid,source,src,ref", "SmartRecruiters company/job; drop trid tracking"],
    [true, "(^|\\.)teamtailor\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Teamtailor job"],
    [true, "(^|\\.)recruitee\\.com$", "^/o/([^/?#]+).*", "/o/$1", "", "", "Recruitee opening"],
    [true, "(^|\\.)jobs\\.personio\\.(de|com)$", "^/job/(\\d+).*", "/job/$1", "", "display,language,source,src,pid,it,_ghcid,apply,_pc", "Personio clean job id"],
    [true, "(^|\\.)bamboohr\\.com$", "^/careers/(\\d+).*", "/careers/$1", "", "", "BambooHR careers id"],
    [true, "(^|\\.)breezy\\.hr$", "^/p/([^/?#]+).*", "/p/$1", "", "", "Breezy posting"],
    [true, "(^|\\.)freshteam\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Freshteam job"],
    [true, "(^|\\.)zohorecruit\\.com$", "^/jobs/Careers/(\\d+).*", "/jobs/Careers/$1", "", "", "Zoho Recruit career id"],
    [true, "(^|\\.)careers\\.hibob\\.com$", "^/jobs/([^/?#]+)(/apply)?$", "/jobs/$1", "", "", "HiBob remove /apply"],
    [true, "(^|\\.)oraclecloud\\.com$", "^/hcmUI/CandidateExperience/([^/]+)/job/([^/?#]+).*", "/hcmUI/CandidateExperience/$1/job/$2", "", "", "Oracle Cloud HCM"],
    [true, "(^|\\.)applytojob\\.com$", "^/apply/([^/]+)/([^/?#]+).*", "/apply/$1/$2", "", "", "ApplyToJob"],
    [true, "(^|\\.)factorialhr\\.com$", "^/job_posting/([^/?#]+).*", "/job_posting/$1", "", "", "Factorial job posting"],
    [true, "(^|\\.)factorial\\.[a-z.]+$", "^/job_posting/([^/?#]+).*", "/job_posting/$1", "", "", "Factorial regional job posting"],
    [true, "(^|\\.)app\\.loxo\\.co$", "^/(.+?)(/form)?$", "/$1", "", "", "Loxo remove /form"],
    [true, "(^|\\.)peopleforce\\.io$", "^/careers/v/([^/?#]+).*", "/careers/v/$1", "", "", "PeopleForce career"],
    [true, "(^|\\.)careers-page\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Careers Page job"],
    [true, "^interviewroom\\.ai$", "^/job_post/([^/?#]+).*", "/job_post/$1", "", "", "InterviewRoom job"],
    [true, "(^|\\.)hiresome\\.ai$", "^/apply_form/([^/?#]+).*", "/apply_form/$1", "", "", "Hiresome apply form"],
    [true, "(^|\\.)icims\\.com$", "^/jobs/(\\d+)/([^/?#]+)/job.*", "/jobs/$1/$2/job", "", "mode,hub,iis,in_iframe,source,src,ref", "ICIMS canonical job path"],
    [true, "(^|\\.)jobvite\\.com$", "^/([^/]+)/job/([^/?#]+).*", "/$1/job/$2", "", "__jvst,__jvsd,source,src,ref", "Jobvite company/job"],

    [true, "^linkedin\\.com$", "^/jobs/view/(\\d+).*", "/jobs/view/$1", "currentJobId", "", "LinkedIn job view"],
    [true, "^linkedin\\.com$", "^/jobs/view.*", "/jobs/view", "currentJobId", "", "LinkedIn currentJobId fallback"],
    [true, "^linkedin\\.com$", "^/jobs/search/?$", "/jobs/search", "currentJobId", "", "LinkedIn search keeps currentJobId"],

    [true, "(^|\\.)indeed\\.[a-z.]+$", "^/viewjob.*", "/viewjob", "jk", "", "Indeed viewjob keeps jk"],
    [true, "(^|\\.)jobstreet\\.[a-z.]+$", "^/job/(\\d+).*", "/job/$1", "", "", "Jobstreet job id"],
    [true, "(^|\\.)welcometothejungle\\.com$", "^/en/companies/([^/]+)/jobs/([^/?#]+).*", "/en/companies/$1/jobs/$2", "", "", "WTTJ canonical job path"],
    [true, "(^|\\.)otta\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Otta job slug"],
    [true, "(^|\\.)cord\\.co$", "^/companies/([^/]+)/jobs/([^/?#]+).*", "/companies/$1/jobs/$2", "", "", "Cord company/job path"],
    [true, "(^|\\.)remoteok\\.com$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "RemoteOK canonical job slug"],
    [true, "(^|\\.)weworkremotely\\.com$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "WWR canonical job slug"],
    [true, "(^|\\.)wellfound\\.com$", "^/jobs/(\\d+).*", "/jobs/$1", "", "", "Wellfound numeric job ID"],
    [true, "(^|\\.)startupjobs\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "StartupJobs canonical job slug"],
    [true, "(^|\\.)euremotejobs\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "EU Remote Jobs canonical path"],
    [true, "(^|\\.)remotive\\.com$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "Remotive canonical job slug"],
    [true, "(^|\\.)remote\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Remote.com canonical job slug"],
    [true, "^join\\.com$", "^/companies/([^/]+)/([^/?#]+).*", "/companies/$1/$2", "", "pid", "JOIN exact company/job slug; drop pid tracking"],
    [true, "(^|\\.)join\\.com$", "^/companies/([^/]+)/([^/?#]+).*", "/companies/$1/$2", "", "pid", "JOIN company/job slug fallback; drop pid tracking"],
    [true, "(^|\\.)joinhandshake\\.com$", "^/jobs/(\\d+).*", "/jobs/$1", "", "", "Handshake numeric job ID"],
    [true, "(^|\\.)monster\\.com$", "^/job-openings/([^/?#]+).*", "/job-openings/$1", "", "", "Monster job opening slug"],
    [true, "(^|\\.)dice\\.com$", "^/job-detail/([^/?#]+).*", "/job-detail/$1", "", "jr_id,source,src,ref", "Dice job detail slug"],
    [true, "(^|\\.)reed\\.co\\.uk$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Reed job slug"],
    [true, "(^|\\.)totaljobs\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Totaljobs canonical path"],
    [true, "(^|\\.)cv-library\\.co\\.uk$", "^/job/(\\d+).*", "/job/$1", "", "", "CV Library numeric ID"],
    [true, "(^|\\.)seek\\.com\\.au$", "^/job/(\\d+).*", "/job/$1", "", "", "SEEK numeric ID"],
    [true, "(^|\\.)glassdoor\\.[a-z.]+$", "^/job-listing/([^/?#]+).*", "/job-listing/$1", "", "", "Glassdoor job listing slug"],
    [true, "(^|\\.)ziprecruiter\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "ZipRecruiter canonical path"],
    [true, "(^|\\.)simplyhired\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "SimplyHired canonical path"],
    [true, "(^|\\.)careerbuilder\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "CareerBuilder canonical path"],
    [true, "(^|\\.)builtin\\.[a-z.]+$", "^/job/([^/?#]+).*", "/job/$1", "", "", "BuiltIn canonical path"],

    [true, "(^|\\.)himalayas\\.app$", "^/companies/([^/]+)/jobs/([^/?#]+).*", "/companies/$1/jobs/$2", "", "", "Himalayas company/job path"],
    [true, "(^|\\.)nodesk\\.co$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "NoDesk job slug"],
    [true, "(^|\\.)workingnomads\\.com$", "^/jobs/?$", "/jobs", "job", "tag,positionType,postedDate,source,src,ref", "WorkingNomads jobs keeps job param"],
    [true, "(^|\\.)workingnomads\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Working Nomads job slug"],
    [true, "(^|\\.)justremote\\.co$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "JustRemote job slug"],
    [true, "(^|\\.)remoteleaf\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "RemoteLeaf job slug"],
    [true, "(^|\\.)dynamitejobs\\.com$", "^/company/([^/]+)/remote-job/([^/?#]+).*", "/company/$1/remote-job/$2", "", "", "Dynamite Jobs company/job path"],
    [true, "(^|\\.)powertofly\\.com$", "^/jobs/detail/(\\d+).*", "/jobs/detail/$1", "", "", "PowerToFly numeric detail ID"],
    [true, "(^|\\.)flexjobs\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "FlexJobs canonical slug"],
    [true, "(^|\\.)cryptocurrencyjobs\\.co$", "^/([^/?#]+).*", "/$1", "", "", "CryptoCurrencyJobs slug"],
    [true, "(^|\\.)cryptojobslist\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "CryptoJobsList slug"],
    [true, "(^|\\.)web3\\.career$", "^/([^/?#]+).*", "/$1", "", "", "Web3 Career slug"],
    [true, "(^|\\.)remote3\\.co$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "Remote3 slug"],
    [true, "(^|\\.)ai-jobs\\.net$", "^/job/(\\d+)/([^/?#]+).*", "/job/$1/$2", "", "", "AI Jobs numeric id and slug"],
    [true, "(^|\\.)aijobs\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "AIJobs canonical slug"],
    [true, "(^|\\.)kaggle\\.com$", "^/jobs/(\\d+).*", "/jobs/$1", "", "", "Kaggle job id"],
    [true, "(^|\\.)stackoverflowjobs\\.com$", "^/jobs/(\\d+).*", "/jobs/$1", "", "", "StackOverflow legacy job id"],

    [true, "(^|\\.)relocate\\.me$", "^/([^/?#]+).*", "/$1", "", "", "Relocate.me slug"],
    [true, "(^|\\.)eurojobs\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "EuroJobs canonical path"],
    [true, "(^|\\.)eures\\.europa\\.eu$", "^/jobseekers/jobdetails/(\\d+).*", "/jobseekers/jobdetails/$1", "", "", "EURES numeric details"],
    [true, "(^|\\.)jobs\\.europa\\.eu$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "EU Jobs slug"],
    [true, "(^|\\.)arbeitnow\\.com$", "^/view/([^/?#]+).*", "/view/$1", "", "", "Arbeitnow job view slug"],
    [true, "(^|\\.)stepstone\\.[a-z.]+$", "^/stellenangebote--([^/?#]+).*", "/stellenangebote--$1", "", "", "StepStone DE style"],
    [true, "(^|\\.)stepstone\\.[a-z.]+$", "^/job/([^/?#]+).*", "/job/$1", "", "", "StepStone job path"],
    [true, "(^|\\.)jobs\\.ch$", "^/en/vacancies/detail/([^/?#]+).*", "/en/vacancies/detail/$1", "", "", "Jobs.ch detail path"],
    [true, "(^|\\.)jobup\\.ch$", "^/en/jobs/detail/([^/?#]+).*", "/en/jobs/detail/$1", "", "", "JobUp detail"],
    [true, "(^|\\.)jobsdb\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "JobsDB job path"],
    [true, "(^|\\.)foundit\\.[a-z.]+$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Foundit job path"],
    [true, "(^|\\.)naukri\\.com$", "^/job-listings-([^/?#]+).*", "/job-listings-$1", "", "", "Naukri job listing slug"],
    [true, "(^|\\.)timesjobs\\.com$", "^/job-detail/([^/?#]+).*", "/job-detail/$1", "", "", "TimesJobs detail slug"],
    [true, "(^|\\.)shine\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Shine job slug"],
    [true, "(^|\\.)bdjobs\\.com$", "^/jobdetails\\.asp.*", "/jobdetails.asp", "id", "", "BDJobs id param"],
    [true, "(^|\\.)kalibrr\\.com$", "^/c/([^/]+)/jobs/(\\d+)/([^/?#]+).*", "/c/$1/jobs/$2/$3", "", "", "Kalibrr company/job path"],
    [true, "(^|\\.)glints\\.com$", "^/opportunities/jobs/([^/?#]+).*", "/opportunities/jobs/$1", "", "", "Glints opportunity slug"],
    [true, "(^|\\.)bossjob\\.[a-z.]+$", "^/en-us/job/([^/?#]+).*", "/en-us/job/$1", "", "", "Bossjob job slug"],
    [true, "(^|\\.)onlinejobs\\.ph$", "^/jobseekers/job/([^/?#]+).*", "/jobseekers/job/$1", "", "", "OnlineJobs PH slug"],
    [true, "(^|\\.)workana\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Workana job slug"],
    [true, "(^|\\.)getonbrd\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "GetOnBoard job slug"],
    [true, "(^|\\.)bumeran\\.[a-z.]+$", "^/empleos/([^/?#]+).*", "/empleos/$1", "", "", "Bumeran job slug"],
    [true, "(^|\\.)computrabajo\\.[a-z.]+$", "^/ofertas-de-trabajo/oferta-de-trabajo-de-([^/?#]+).*", "/ofertas-de-trabajo/oferta-de-trabajo-de-$1", "", "", "Computrabajo offer slug"],
    [true, "(^|\\.)elempleo\\.com$", "^/co/ofertas-trabajo/([^/?#]+).*", "/co/ofertas-trabajo/$1", "", "", "Elempleo offer"],
    [true, "(^|\\.)occ\\.com\\.mx$", "^/empleo/oferta/([^/?#]+).*", "/empleo/oferta/$1", "", "", "OCC Mexico offer"],
    [true, "(^|\\.)trabajando\\.[a-z.]+$", "^/trabajo/([^/?#]+).*", "/trabajo/$1", "", "", "Trabajando job"],
    [true, "(^|\\.)infojobs\\.net$", "^/([^/]+)/([^/?#]+).*", "/$1/$2", "", "", "InfoJobs two-part canonical"],
    [true, "(^|\\.)tecnoempleo\\.com$", "^/([^/?#]+).*", "/$1", "", "", "Tecnoempleo slug"],
    [true, "(^|\\.)landing\\.jobs$", "^/at/([^/]+)/([^/?#]+).*", "/at/$1/$2", "", "", "Landing Jobs company/job"],
    [true, "(^|\\.)talent\\.com$", "^/view.*", "/view", "id", "", "Talent.com ID param"],
    [true, "(^|\\.)adzuna\\.[a-z.]+$", "^/details/(\\d+).*", "/details/$1", "", "", "Adzuna details id"],
    [true, "(^|\\.)jobrapido\\.[a-z.]+$", "^/jobpreview/([^/?#]+).*", "/jobpreview/$1", "", "", "Jobrapido preview id"],
    [true, "(^|\\.)thelocal\\.[a-z.]+$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "The Local job slug"],

    [true, "(^|\\.)angel\\.co$", "^/company/([^/]+)/jobs/(\\d+).*", "/company/$1/jobs/$2", "", "", "AngelList legacy"],
    [true, "(^|\\.)ycombinator\\.com$", "^/companies/([^/]+)/jobs/([^/?#]+).*", "/companies/$1/jobs/$2", "", "", "YC company/job"],
    [true, "(^|\\.)levels\\.fyi$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Levels.fyi job slug"],
    [true, "(^|\\.)arc\\.dev$", "^/remote-jobs/([^/?#]+).*", "/remote-jobs/$1", "", "", "Arc remote job slug"],
    [true, "(^|\\.)turing\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Turing job slug"],
    [true, "(^|\\.)gun\\.io$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Gun.io job slug"],
    [true, "(^|\\.)upwork\\.com$", "^/jobs/([^_/?#]+)_~([^/?#]+).*", "/jobs/$1_~$2", "", "", "Upwork canonical"],
    [true, "(^|\\.)freelancer\\.com$", "^/projects/([^/?#]+).*", "/projects/$1", "", "", "Freelancer project slug"],
    [true, "(^|\\.)peopleperhour\\.com$", "^/freelance-jobs/([^/?#]+).*", "/freelance-jobs/$1", "", "", "PeoplePerHour job slug"],
    [true, "(^|\\.)contra\\.com$", "^/opportunity/([^/?#]+).*", "/opportunity/$1", "", "", "Contra opportunity"],
    [true, "(^|\\.)toptal\\.com$", "^/freelance-jobs/([^/?#]+).*", "/freelance-jobs/$1", "", "", "Toptal freelance job slug"],
    [true, "(^|\\.)hired\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Hired job slug"],
    [true, "(^|\\.)triplebyte\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Triplebyte job slug"],
    [true, "(^|\\.)cutshort\\.io$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Cutshort job"],
    [true, "(^|\\.)usebraintrust\\.com$", "^/talent/jobs/([^/?#]+).*", "/talent/jobs/$1", "", "", "Braintrust talent job"],
    [true, "(^|\\.)braintrust\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Braintrust job slug"],
    [true, "(^|\\.)clouddevs\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "CloudDevs job slug"],
    [true, "(^|\\.)remoteplatz\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Remoteplatz job slug"],
    [true, "(^|\\.)remotasks\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Remotasks job slug"],
    [true, "(^|\\.)micro1\\.ai$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Micro1 AI job slug"],
    [true, "(^|\\.)mercor\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Mercor job slug"],
    [true, "(^|\\.)outlier\\.ai$", "^/opportunities/([^/?#]+).*", "/opportunities/$1", "", "", "Outlier opportunity"],
    [true, "(^|\\.)g2i\\.co$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "G2i job slug"],
    [true, "(^|\\.)x-team\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "X-Team job slug"],
    [true, "(^|\\.)andela\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Andela job slug"],
    [true, "(^|\\.)torre\\.co$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Torre job slug"],
    [true, "(^|\\.)vanhack\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "VanHack job slug"],
    [true, "(^|\\.)lemon\\.io$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Lemon.io job slug"],
    [true, "(^|\\.)epicjobs\\.co$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "EpicJobs job slug"],
    [true, "(^|\\.)jobspresso\\.co$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Jobspresso job slug"],
    [true, "(^|\\.)virtualvocations\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Virtual Vocations job"],
    [true, "(^|\\.)remotehub\\.com$", "^/jobs/details/([^/?#]+).*", "/jobs/details/$1", "", "", "RemoteHub details"],
    [true, "(^|\\.)dribbble\\.com$", "^/jobs/(\\d+).*", "/jobs/$1", "", "", "Dribbble numeric job ID"],
    [true, "(^|\\.)behance\\.net$", "^/joblist/([^/?#]+).*", "/joblist/$1", "", "", "Behance job slug"],
    [true, "(^|\\.)authenticjobs\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Authentic Jobs slug"],
    [true, "(^|\\.)krop\\.com$", "^/creative-jobs/([^/?#]+).*", "/creative-jobs/$1", "", "", "Krop creative job"],
    [true, "(^|\\.)mediabistro\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Mediabistro job slug"],
    [true, "(^|\\.)snaphunt\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "Snaphunt slug"],
    [true, "(^|\\.)grabjobs\\.co$", "^/job/([^/?#]+).*", "/job/$1", "", "", "GrabJobs slug"],
    [true, "(^|\\.)jooble\\.[a-z.]+$", "^/desc/([^/?#]+).*", "/desc/$1", "", "", "Jooble desc id"],
    [true, "(^|\\.)lensa\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Lensa job slug"],
    [true, "(^|\\.)zippia\\.com$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Zippia job slug"],
    [true, "(^|\\.)jora\\.[a-z.]+$", "^/job/([^/?#]+).*", "/job/$1", "", "", "Jora job slug"],
    [true, "(^|\\.)workopolis\\.com$", "^/jobsearch/viewjob/([^/?#]+).*", "/jobsearch/viewjob/$1", "", "", "Workopolis viewjob"],
    [true, "(^|\\.)jobbank\\.gc\\.ca$", "^/jobsearch/jobposting/(\\d+).*", "/jobsearch/jobposting/$1", "", "", "Canada Job Bank posting id"],
    [true, "(^|\\.)seek\\.co\\.nz$", "^/job/(\\d+).*", "/job/$1", "", "", "SEEK NZ numeric ID"],
    [true, "(^|\\.)trademe\\.co\\.nz$", "^/a/jobs/listing/(\\d+).*", "/a/jobs/listing/$1", "", "", "TradeMe listing id"],
    [true, "(^|\\.)careerjet\\.[a-z.]+$", "^/jobad/([^/?#]+).*", "/jobad/$1", "", "", "CareerJet job ad id"],
    [true, "(^|\\.)neuvoo\\.[a-z.]+$", "^/view/([^/?#]+).*", "/view/$1", "", "", "Neuvoo/Talent legacy view id"],

    [true, "^jobs\\.paymentology\\.com$", "^/detail/?$", "/detail", "uid,coref", "", "Paymentology detail keeps uid/coref"],
    [true, "(^|\\.)callstack\\.com$", "^/careers/?$", "/careers", "", "job", "Callstack careers page keeps fragment job id"],
    [true, "(^|\\.)li\\.me$", "^/about/careers/?$", "/about/careers", "ashby_jid", "utm_source,source,src,ref", "Custom Ashby li.me careers"],
    [true, "(^|\\.)smallpdf\\.com$", "^/jobs/?$", "/jobs", "ashby_jid", "utm_source,source,src,ref", "Custom Ashby Smallpdf jobs"],
    [true, "(^|\\.)creatoriq\\.com$", "^/careers/?$", "/careers", "ashby_jid", "utm_source,source,src,ref", "Custom Ashby CreatorIQ careers"],
    [true, "(^|\\.)vcluster\\.com$", "^/careers/?$", "/careers", "ashby_jid", "utm_source,source,src,ref", "Custom Ashby careers"],
    [true, "^kuali\\.co$", "^/positions/?$", "/positions", "gnk,gni", "gns,source,src,ref", "Kuali positions keeps job identity gnk/gni"],
    [true, "^remoteyeah\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "", "RemoteYeah job"],
    [true, "(^|\\.)alignerr\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "referral-source,source,src,ref", "Alignerr job id"],
    [true, "(^|\\.)rubyplay\\.com$", "^/en/postings/([^/?#]+)(/applications/new)?$", "/en/postings/$1", "", "", "RubyPlay remove applications/new"],
    [true, "(^|\\.)epam\\.com$", "^/en/vacancy/([^/?#]+).*", "/en/vacancy/$1", "", "country,source,src,ref", "EPAM vacancy id"],
    [true, "^jobs\\.elastic\\.co$", "^/jobs/pipeline/(.+)$", "/jobs/pipeline/$1", "gh_jid", "source,src,ref", "Elastic job keeps gh_jid"],
    [true, "(^|\\.)iqvia\\.com$", "^/jobs/([^/?#]+).*", "/jobs/$1", "", "amp;utm_source,amp;utm_medium,utm_source,utm_medium,source,src,ref", "IQVIA job id"],
    [true, "^app\\.screenloop\\.com$", "^/careers/([^/]+)/job_posts/([^/?#]+).*", "/careers/$1/job_posts/$2", "", "tab,sl_source,igbTracker,source,src,ref", "Screenloop job post"],
    [true, "^recruit\\.zoho\\.com$", "^/recruit/ViewJob\\.na.*", "/recruit/ViewJob.na", "digest", "embedsource,source,src,ref", "Zoho Recruit digest job"],
    [true, "(^|\\.)biospace\\.com$", "^/job/(\\d+).*", "/job/$1", "", "TrackID,utm_source,utm_medium,utm_campaign,source,src,ref", "BioSpace job id"],

    [true, ".*", "", "", "", "share,shared,fb_action_ids,fb_action_types,fb_source,irclickid,irgwc,affid,affiliate,affiliate_id,partner,partnerid,clickid,click_id", "Global extra drop params fallback"]
  ];
}
