/********************
 * SHARED LAYOUT / CONSTANTS
 * Single source for sheet names, headers, status keys, UI labels, and spec values.
 ********************/

const CFG = {
  APP_NAME: "BOC-E",

  UI_MENU: {
    ITEMS: [
      { label: "Map selected jobs to profiles", fn: "mapSelectedVisibleJobsToProfiles" },
      { label: "Resume job mapping", fn: "RESUME_JOB_MAPPING" },
      { label: "Show job mapping status", fn: "SHOW_JOB_MAPPING_STATUS" },
      { separator: true },
      { label: "Get File Name", fn: "fillFileNamesFromSelection" },
      { separator: true },
      { label: "Resume normalize", fn: "RESUME_NORMALIZE" },
      { label: "Show normalize status", fn: "SHOW_NORMALIZE_STATUS" },
      { separator: true },
      { label: "Publish URL Rules", fn: "UPDATE_RULES_SNAPSHOT_FROM_SHEET" },
      { label: "Warm cache", fn: "WARM_URL_NORMALIZER_RUNTIME" },
      { separator: true },
      { label: "Send manual message", fn: "uiSendManualMessage" },
      { label: "Run reports now", fn: "uiRunReportNow" },
      { label: "Clean bot history", fn: "uiCleanBotHistory" },
    ]
  },

  SHEETS: {
    PROFILES: "Profiles",
    TAGS: "Tags",
    STATUS: "STATUS",
    URL_RULES: "URL_RULES",
    JOB_FALLBACK: ["EU", "LATAM", "US", "APAC"],
  },

  HEADERS: {
    JOB: {
      JOB_ID: "job_id",
      DATE: "date",
      JOB_URL: "job_url",
      REGION_TAGS: "region_tags",
      STACK_TAGS: "stack_tags"
    },
    PROFILE: {
      PROFILE_ID: "profile_id",
      PROFILE_NAME: "profile_name",
      PROFILE_NAME_ALIASES: ["profile_name", "profilename", "name", "profile"],
      REGION_TAG_ALIASES: ["region_tag", "region_tags", "region"],
      STACK_TAG_ALIASES: ["stack_tag", "stack_tags", "stack"],
      ACTIVE_ALIASES: ["status", "active"],
      SHEET_NAME: "sheet_name"
    },
    TAGS: {
      GROUP: "group",
      TAG: "tag"
    },

    PROFILE_TAB: {
      APP_KEY: "app_key",
      DATE: "date",
      JOB_URL: "job_url",
      STATUS: "status",
      BIDDER: "bidder",
      RESUME_URL: "resume_url",
      FILE_NAME: "file_name",
      CREATED_AT: "created_at",
      PRICE: "price",
      ERROR: "error",
      JD: "JD"
    },
    URL_RULES: {
      ENABLED: "enabled",
      HOST_REGEX: "host_regex",
      PATH_REGEX: "path_regex",
      PATH_REPLACEMENT: "path_replacement",
      KEEP_PARAMS: "keep_params",
      DROP_PARAMS: "drop_params",
      NOTES: "notes"
    },
  },

  STATUS: {
    HEADER_ROW: 1,
    BLOCK_WIDTH: 2,

    VALUE: "value",

    NORMALIZE_TITLE: "Normalize status key",
    NORMALIZE_START_COL: 1,

    JOB_MAP_TITLE: "Job map status key",
    JOB_MAP_START_COL: 4
  },

  STATUS_KEYS: {
    NORMALIZE: {
      STATE: "Status",
      SOURCE: "Run Source",
      SHEET: "Sheet",
      JOB_URL_COLUMN: "Job URL Column",
      EXPAND_REDIRECTS: "Follow Redirects",
      REDIRECTS_USED_THIS_RUN: "Redirects Used",
      CHUNK_SIZE: "Batch Size",
      TOTAL_ROWS_PLANNED: "Total Rows",
      PROCESSED_COUNT: "Rows Processed",
      PROGRESS: "Progress",
      REPLACED_COUNT: "URLs Updated",
      LAST_PROCESSED_ROW: "Last Processed Row",
      LAST_REPLACED_ROW: "Last Updated Row",
      LAST_ERROR_ROW: "Last Error Row",
      LAST_ERROR_MESSAGE: "Last Error",
      STARTED_AT: "Started",
      UPDATED_AT: "Updated"
    },

    JOB_MAP: {
      STATUS: "Status",
      JOB_SHEET_NAME: "Job Sheet",
      START_ROW: "Start Row",
      NUM_ROWS: "Selected Rows",
      NEXT_OFFSET: "Resume Position",
      CHUNK_SIZE: "Batch Size",
      VISIBLE_JOBS: "Visible jobs processed",
      ADDED: "Added",
      DUPES: "Duplicates",
      NO_MATCH: "No Matches",
      ERRORS: "Errors",
      LAST_ERROR_ROW: "Last Error Row",
      LAST_ERROR_MESSAGE: "Last Error",
      STARTED_AT: "Started",
      UPDATED_AT: "Updated"
    }
  },


  STACK_FILTER_TABLE_RANGE: "F1:L16",

  VALUES: {
    DEFAULT_STACK: "SWE",
    DEFAULT_REGION: "GEN",
    PROFILE_DEFAULT_NAME: "Profile",
    PROFILE_ACTIVE_DEFAULT: "active",
    PROFILE_INACTIVE: ["inactive", "false", "no", "n", "0", "disabled", "off", ""],
    PROFILE_OUTPUT_STATUS: {
      TODO: "TODO",
      ERROR: "ERROR"
    },
    BOOLEAN_TEXT: {
      TRUE: "TRUE",
      FALSE: "FALSE",
      ON: "ON",
      OFF: "OFF"
    },
    RUN_STATE: {
      READY: "READY",
      RUNNING: "RUNNING",
      PAUSED: "PAUSED",
      DONE: "DONE",
      ERROR: "ERROR",
      STOPPED_BEFORE_TIMEOUT: "STOPPED_BEFORE_TIMEOUT"
    },
    FILE_LOOKUP: {
      INVALID_URL: "Invalid URL",
      NO_ACCESS: "No access",
      DATE_FORMAT: "MM/dd"
    }

  },

  STACK_TAGS: {

    SWE: "SWE",

    AI: "AI",

    QA: "QA"

  },

  CHARTS: {
    DASHBOARD_SHEET: "Dashboard",
    TITLE_RULES: {
      CHART1: {
        title: "Dashboard!A2",
        sheet: "Chart",
        chartIndex: 0,
        watchCells: ["C1", "F1"]
      },

      CHART2: {
        title: "Dashboard!A10",
        sheet: "Chart",
        chartIndex: 1,
        watchCells: ["C1", "F1", "A9"]
      }
    }
  },

  PROFILE_TAB_HEADERS: [
    "app_key",
    "date",
    "job_url",
    "status",
    "bidder",
    "resume_url",
    "file_name",
    "created_at",
    "price",
    "error",
    "JD"
  ],

  JOB_MAPPER: {
    STATE_KEY: "JOB_MAP_STATE_JSON",
    CHUNK_SIZE: 50,
    MAX_RUNTIME_MS: 4 * 60 * 1000,
    RESUME_AFTER_MS: 60000
  },

  URL_NORMALIZER: {
    RULES_SHEET_NAME: "URL_RULES",
    STATUS_SHEET_NAME: "STATUS",
    RULES_CACHE_KEY: "url_rules_v16_ats_query_fix",
    REDIRECT_CACHE_PREFIX: "redir_v6:",
    NORMALIZER_RUNTIME_STATE_KEY: "url_normalizer_status_v3", //
    RULES_SNAPSHOT_PROPERTY_KEY: "url_rules_snapshot_v1",
    CACHE_TTL_SECONDS: 21600,
    REDIRECT_CACHE_TTL_SECONDS: 86400,
    JOB_URL_HEADER: "job_url",
    HEADER_ROW: 1,
    DEFAULT_CHUNK_SIZE: 100,
    REDIRECT_CHUNK_SIZE: 20,
    MAX_CHUNK_SIZE: 500,
    MAX_REDIRECTS_PER_RUN: 80,
    SAFE_RUNTIME_MS: 5 * 60 * 1000,
    STOP_BUFFER_MS: 25 * 1000,
    STATUS_UPDATE_EVERY_ROWS: 100,
    STATUS_UPDATE_EVERY_MS: 12000,
    ON_EDIT_RUNTIME_MS: 25 * 1000
  }
};

CFG.PROFILE_SHEET_NAME = CFG.SHEETS.PROFILES;
CFG.TAGS_SHEET_NAME = CFG.SHEETS.TAGS;
CFG.JOB_SHEETS_FALLBACK = CFG.SHEETS.JOB_FALLBACK;
CFG.DEFAULT_STACK = CFG.VALUES.DEFAULT_STACK;
CFG.MAPPER_STATUS_SHEET_NAME = CFG.SHEETS.STATUS;
CFG.MAPPER_STATE_KEY = CFG.JOB_MAPPER.STATE_KEY;
CFG.MAPPER_CHUNK_SIZE = CFG.JOB_MAPPER.CHUNK_SIZE;
CFG.MAPPER_MAX_RUNTIME_MS = CFG.JOB_MAPPER.MAX_RUNTIME_MS;
CFG.MAPPER_LEASE_MS = CFG.MAPPER_MAX_RUNTIME_MS;
CFG.CANONICAL_STACK_TAG_SET = new Set(
  Object.values(CFG.STACK_TAGS).map(v =>
    String(v)
      .trim()
      .toUpperCase()
      .replace(/\s+/g, "_")
  )
);

function SHOW_ALL_SCRIPT_PROPERTIES() {
  Logger.log(
    PropertiesService
      .getScriptProperties()
      .getProperties()
  );
}