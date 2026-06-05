const SETTINGS_SHEET = 'STATUS';

const MEMBER_RANGE = 'G2:I8'; // User table range (User | Mail | SlackID)
const CHANNEL_RANGE = 'K2:L9'; // Channel table range (channel_name | channel_id)
const RULE_RANGE = 'G12:M17'; // Rules table range (no header, only rule rows)
const VALID_RULE_TYPES = ['daily', 'weekly', 'once'];

//Rules table header ->  report_title	type	target_names	report_range	day/date	time	enabled

let SLACK_TARGET_CACHE = null;

function getSlackToken_() {
  const token = PropertiesService
    .getScriptProperties()
    .getProperty('slack_bot_token');

  if (!token) {
    throw new Error('Missing slack_bot_token in Script Properties.');
  }

  return token;
}

function validateReportRule_(rule, options) {
  options = options || {};
  const errors = [];

  if (!rule) {
    errors.push('Rule not found.');
    return errors;
  }

  const type = String(rule.type || '').trim().toLowerCase();
  const day = String(rule.day || '').trim();
  const time = String(rule.time || '').trim();

  if (!String(rule.target_names || '').trim()) {
    errors.push('Missing target_names.');
  }

  if (!String(rule.report_range || '').trim()) {
    errors.push('Missing report_range.');
  }

  validateReportRanges_(rule.report_range, errors);

  if (options.requireSchedule) {
    if (!VALID_RULE_TYPES.includes(type)) {
      errors.push('Invalid type. Allowed: daily, weekly, once.');
    }

    if (!isValidTime_(time)) {
      errors.push('Invalid time. Use HH:mm from 00:00 to 23:59. Example: 09:30.');
    }

    if (type === 'daily') {
      if (day.toLowerCase() !== 'each') {
        errors.push('For daily type, day/date must be "each".');
      }
    }

    if (type === 'weekly') {
      if (!isValidWeekdayList_(day)) {
        errors.push('For weekly type, day/date must be weekday(s). Example: Mon or Mon,Thu.');
      }
    }

    if (type === 'once') {
      if (!isValidMonthDay_(day)) {
        errors.push('For once type, day/date must be MM/dd. Example: 12/31.');
      }
    }
  }

  return errors;
}

function validateReportRanges_(reportRange, errors) {
  const text = String(reportRange || '').trim();
  if (!text) return;

  const ranges = text
    .split(',')
    .map(r => r.trim())
    .filter(Boolean);

  ranges.forEach(rangeA1 => {
    try {
      SpreadsheetApp.getActive().getRange(rangeA1);
    } catch (e) {
      errors.push('Invalid report_range: ' + rangeA1);
    }
  });
}

function isValidTime_(value) {
  const text = String(value || '').trim();

  if (!/^\d{1,2}:\d{2}$/.test(text)) return false;

  const [h, m] = text.split(':').map(Number);

  return h >= 0 && h <= 23 && m >= 0 && m <= 59;
}

function normalizeTime_(value) {
  if (!isValidTime_(value)) return '';

  const [h, m] = String(value).trim().split(':');

  return String(h).padStart(2, '0') + ':' + String(m).padStart(2, '0');
}

function normalizeMonthDay_(value) {
  const parts = String(value || '').trim().split('/');
  if (parts.length !== 2) return '';
  return parts[0].padStart(2, '0') + '/' + parts[1].padStart(2, '0');
}

function isValidMonthDay_(value) {
  const text = String(value || '').trim();

  if (!/^\d{1,2}\/\d{1,2}$/.test(text)) return false;

  const [month, day] = text.split('/').map(Number);

  const daysInMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  return month >= 1 &&
    month <= 12 &&
    day >= 1 &&
    day <= daysInMonth[month - 1];
}

function isValidWeekdayList_(value) {
  const allowed = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];

  const days = String(value || '')
    .split(/[,\s/]+/)
    .map(d => d.trim().toLowerCase())
    .filter(Boolean);

  if (!days.length) return false;

  return days.every(d => allowed.includes(d));
}

function isScheduledDayMatch_(ruleDay, today, weekday, type) {
  const day = String(ruleDay || '').trim();
  const ruleType = String(type || '').trim().toLowerCase();

  if (!day || day === '-') return false;

  if (ruleType === 'daily') {
    return day.toLowerCase() === 'each';
  }

  if (ruleType === 'once') {
    return normalizeMonthDay_(day) === today;
  }

  if (ruleType === 'weekly') {
    const allowedDays = day
      .split(/[,\s/]+/)
      .map(d => d.trim().toLowerCase())
      .filter(Boolean);

    return allowedDays.includes(String(weekday || '').toLowerCase());
  }

  return false;
}

function runScheduledReports() {
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const today = Utilities.formatDate(now, tz, 'MM/dd');
  const weekday = Utilities.formatDate(now, tz, 'EEE').toLowerCase();
  const currentTime = Utilities.formatDate(now, tz, 'HH:mm');

  const rules = getReportRules_();

  rules.forEach((rule, index) => {

    if (String(rule.enabled).toUpperCase() !== 'TRUE') return;

    const errors = validateReportRule_(rule, { requireSchedule: true });

    if (errors.length) {
      console.warn(
        'Scheduled report rule #' + (index + 1) + ' skipped: ' + errors.join(', ')
      );
      return;
    }

    if (
      isScheduledDayMatch_(rule.day, today, weekday, rule.type) &&
      normalizeTime_(rule.time) === currentTime
    ) {
      sendReport_(rule);
    }
  });
}

function runReportsNow() {
  const ui = SpreadsheetApp.getUi();

  const box = ui.prompt(
    'Run report by rule position',
    'Enter rule position number. Example: 1',
    ui.ButtonSet.OK_CANCEL
  );

  if (box.getSelectedButton() !== ui.Button.OK) return;

  const input = String(box.getResponseText() || '').trim();

  if (!/^\d+$/.test(input)) {
    ui.alert('Invalid input. Please enter only a number.');
    return;
  }

  const rules = getReportRules_();
  const ruleNumber = Number(input);

  if (!Array.isArray(rules) || !rules.length) {
    ui.alert('No report rules found.');
    return;
  }

  if (ruleNumber < 1 || ruleNumber > rules.length) {
    ui.alert('Rule number out of range. Available: 1 - ' + rules.length);
    return;
  }

  const rule = rules[ruleNumber - 1];
  const errors = validateReportRule_(rule, { requireSchedule: false });

  if (errors.length) {
    ui.alert('Rule #' + ruleNumber + ' is invalid:\n\n' + errors.join('\n'));
    return;
  }

  sendReport_(rule);

  ui.alert('Report rule #' + ruleNumber + ' sent.');
}

function sendReport_(rule) {
  if (!rule.target_names || !rule.report_range) return;

  const ranges = rule.report_range
    .split(',')
    .map(r => r.trim())
    .filter(Boolean);

  const messages = ranges.map(rangeA1 =>
    buildSlackTableMessage_(rangeA1, rule.rule_name)
  );

  rule.target_names
    .split(',')
    .map(t => t.trim())
    .filter(Boolean)
    .forEach(target => {
      const id = getSlackId_(target);

      if (!id) {
        console.warn(`Slack target not found: ${target}`);
        return;
      }

      messages.forEach(message => postSlackTable_(id, message));
    });
}

function uiSendManualMessage() {
  const ui = SpreadsheetApp.getUi();

  const targetBox = ui.prompt(
    'Targets',
    'Enter names separated by commas. Example: Tommy, OPS, NEW_JOBS',
    ui.ButtonSet.OK_CANCEL
  );

  if (targetBox.getSelectedButton() !== ui.Button.OK) return;

  const msgBox = ui.prompt(
    'Message',
    'Enter Slack message:',
    ui.ButtonSet.OK_CANCEL
  );

  if (msgBox.getSelectedButton() !== ui.Button.OK) return;

  const targets = targetBox.getResponseText()
    .split(',')
    .map(t => t.trim())
    .filter(Boolean);

  const message = msgBox.getResponseText();

  targets.forEach(target => {
    const id = getSlackId_(target);

    if (!id) {
      console.warn(`Slack target not found: ${target}`);
      return;
    }

    postSlackText_(id, message);
  });

  ui.alert('Slack message sent.');
}

function buildSlackTableMessage_(rangeA1, title) {
  const values = SpreadsheetApp.getActive()
    .getRange(rangeA1)
    .getDisplayValues();

  if (!values.length) {
    return {
      title,
      blocks: [
        {
          type: 'section',
          text: {
            type: 'mrkdwn',
            text: `*${title}*\nNo data found.`
          }
        }
      ]
    };
  }

  const rows = values.map(row =>
    row.map(cell => ({
      type: 'raw_text',
      text: String(cell || ' ')
    }))
  );

  return {
    title,
    blocks: [
      {
        type: 'section',
        text: {
          type: 'mrkdwn',
          text: `*${title}*`
        }
      },
      {
        type: 'table',
        rows
      }
    ]
  };
}

function postSlackTable_(channelOrUserId, message) {
  const token = getSlackToken_();
  const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      channel: channelOrUserId,
      text: message.title,
      blocks: message.blocks
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    throw new Error(`Slack API error: ${result.error}`);
  }

  return result;
}

function postSlackText_(channelOrUserId, text) {
  const token = getSlackToken_();

  const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      channel: channelOrUserId,
      text: text
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    throw new Error(`Slack API error: ${result.error}`);
  }

  return result;
}

function getSlackTargetMap_() {
  if (SLACK_TARGET_CACHE) return SLACK_TARGET_CACHE;

  const sh = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);

  if (!sh) {
    throw new Error(`Sheet not found: ${SETTINGS_SHEET}`);
  }

  const map = {};

  sh.getRange(MEMBER_RANGE).getDisplayValues().forEach(row => {
    const name = String(row[0] || '').trim().toLowerCase();
    const id = String(row[2] || '').trim();
    if (name && id) map[name] = id;
  });

  sh.getRange(CHANNEL_RANGE).getDisplayValues().forEach(row => {
    const name = String(row[0] || '').trim().toLowerCase();
    const id = String(row[1] || '').trim();
    if (name && id) map[name] = id;
  });

  SLACK_TARGET_CACHE = map;
  return map;
}

function getSlackId_(name) {
  const target = String(name || '').trim().toLowerCase();
  return getSlackTargetMap_()[target] || null;
}

function getReportRules_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);

  if (!sh) {
    throw new Error(`Sheet not found: ${SETTINGS_SHEET}`);
  }

  return sh.getRange(RULE_RANGE)
    .getDisplayValues()
    .filter(row => String(row[0] || '').trim() !== '')
    .map(row => ({
      rule_name: String(row[0] || '').trim(),
      type: String(row[1] || '').trim().toLowerCase(),
      target_names: String(row[2] || '').trim(),
      report_range: String(row[3] || '').trim(),
      day: String(row[4] || '').trim(),
      time: String(row[5] || '').trim(),
      enabled: String(row[6] || '').trim().toUpperCase()
    }));
}