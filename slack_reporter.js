const SETTINGS_SHEET = 'STATUS';
const CHART_SHEETS = ['Chart'];
const MEMBER_RANGE = 'G2:I8'; // User table range (User | Mail | SlackID)
const CHANNEL_RANGE = 'K2:L9'; // Channel table range (channel_name | channel_id)
const RULE_RANGE = 'G12:M17'; // Rules table range
const VALID_RULE_TYPES = ['daily', 'weekly', 'once'];


let SLACK_TARGET_CACHE = null;

function testSlackAuth() {
  Logger.log(getBotUserId_());
}

function getSlackToken_() {
  const token = PropertiesService.getScriptProperties().getProperty('slack_bot_token');
  if (!token) throw new Error('Missing slack_bot_token in Script Properties.');
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

  validateReportTargets_(rule.target_names, errors);

  if (!String(rule.report_source || '').trim()) {
    errors.push('Missing report_source.');
  }

  validateReportSources_(rule.report_source, errors);

  if (options.requireSchedule) {
    if (!VALID_RULE_TYPES.includes(type)) {
      errors.push('Invalid type. Allowed: daily, weekly, once.');
    }

    if (!isValidTime_(time)) {
      errors.push('Invalid time. Use HH:mm from 00:00 to 23:59.');
    }

    if (type === 'daily' && day.toLowerCase() !== 'each') {
      errors.push('For daily type, day/date must be "each".');
    }

    if (type === 'weekly' && !isValidWeekdayList_(day)) {
      errors.push('For weekly type, day/date must be weekday(s). Example: Mon or Mon,Thu.');
    }

    if (type === 'once' && !isValidMonthDay_(day)) {
      errors.push('For once type, day/date must be MM/dd. Example: 12/31.');
    }
  }

  return errors;
}

function validateReportSources_(reportSource, errors) {
  const text = String(reportSource || '').trim();
  if (!text) return;

  text.split(',').map(s => s.trim()).filter(Boolean).forEach(source => {
    if (isChartSource_(source)) {
      try {
        getChartFromSource_(source);
      } catch (e) {
        errors.push(e.message);
      }
      return;
    }

    try {
      SpreadsheetApp.getActive().getRange(source);
    } catch (e) {
      errors.push('Invalid report_source: ' + source);
    }
  });
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

function isValidTime_(value) {
  const text = String(value || '').trim();
  if (!/^\d{1,2}:\d{2}$/.test(text)) return false;
  const [h, m] = text.split(':').map(Number);
  return h >= 0 && h <= 23 && m >= 0 && m <= 59;
}

function isValidMonthDay_(value) {
  const text = String(value || '').trim();
  if (!/^\d{1,2}\/\d{1,2}$/.test(text)) return false;

  const [month, day] = text.split('/').map(Number);
  const daysInMonth = [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

  return month >= 1 && month <= 12 && day >= 1 && day <= daysInMonth[month - 1];
}

function isValidWeekdayList_(value) {
  const allowed = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];

  const days = String(value || '')
    .split(/[,\s/]+/)
    .map(d => d.trim().toLowerCase())
    .filter(Boolean);

  return days.length > 0 && days.every(d => allowed.includes(d));
}

function isScheduledDayMatch_(ruleDay, today, weekday, type) {
  const day = String(ruleDay || '').trim();
  const ruleType = String(type || '').trim().toLowerCase();

  if (!day || day === '-') return false;

  if (ruleType === 'daily') return day.toLowerCase() === 'each';
  if (ruleType === 'once') return normalizeMonthDay_(day) === today;

  if (ruleType === 'weekly') {
    return day
      .split(/[,\s/]+/)
      .map(d => d.trim().toLowerCase())
      .filter(Boolean)
      .includes(String(weekday || '').toLowerCase());
  }

  return false;
}

function isChartSource_(source) {
  const text = String(source || '').trim();
  const match = text.match(/^([^!,:]+):(\d+)$/);
  if (!match) return false;

  const sheetName = match[1].trim();
  const chartIndex = Number(match[2]);

  return CHART_SHEETS.includes(sheetName) &&
    Number.isInteger(chartIndex) &&
    chartIndex >= 0;
}


function runScheduledReports() {
  const now = new Date();
  const tz = Session.getScriptTimeZone();

  const today = Utilities.formatDate(now, tz, 'MM/dd');
  const weekday = Utilities.formatDate(now, tz, 'EEE').toLowerCase();
  const currentTime = Utilities.formatDate(now, tz, 'HH:mm');

  getReportRules_().forEach((rule, index) => {
    if (String(rule.enabled).toUpperCase() !== 'TRUE') return;

    const errors = validateReportRule_(rule, { requireSchedule: true });

    if (errors.length) {
      console.warn('Scheduled report rule #' + (index + 1) + ' skipped: ' + errors.join(', '));
      return;
    }

    const currentHour = now.getHours();

    const setHour = Number(normalizeTime_(rule.time).split(':')[0]);

    const hourDiff = currentHour - setHour;

    if (
      isScheduledDayMatch_(rule.day, today, weekday, rule.type) && hourDiff >= 0 && hourDiff <= 1) {
      sendReport_(rule);
    }
  });
}

function uiRunReportNow() {
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

  if (!rules.length) {
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
  if (!rule.target_names || !rule.report_source) return;

  const sources = rule.report_source.split(',').map(r => r.trim()).filter(Boolean);

  const reportItems = sources.map(source => {
    if (isChartSource_(source)) return { type: 'chart', source };
    return { type: 'table', message: buildSlackTableMessage_(source, rule.rule_name) };
  });

  rule.target_names.split(',').map(t => t.trim()).filter(Boolean).forEach(target => {
    const resolved = resolveSlackTarget_(target, { verify: true });

    if (!resolved) {
      console.warn('Slack target not found or invalid: ' + target);
      return;
    }

    reportItems.forEach(item => {
      if (item.type === 'chart') {
        postSlackChartPng_(resolved.postId, item.source, rule.rule_name);
      } else {
        postSlackTable_(resolved.postId, item.message);
      }
    });
  });
}

function uiSendManualMessage() {
  const ui = SpreadsheetApp.getUi();

  const targetBox = ui.prompt(
    'Targets',
    'Enter names, Slack usernames, or Slack IDs separated by commas. Example: Tommy, OPS, @timo, U123, C123',
    ui.ButtonSet.OK_CANCEL
  );

  if (targetBox.getSelectedButton() !== ui.Button.OK) return;

  const msgBox = ui.prompt('Message', 'Enter Slack message:', ui.ButtonSet.OK_CANCEL);
  if (msgBox.getSelectedButton() !== ui.Button.OK) return;

  const targets = targetBox.getResponseText().split(',').map(t => t.trim()).filter(Boolean);
  const message = msgBox.getResponseText();

  targets.forEach(target => {
    const resolved = resolveSlackTarget_(target, { verify: true });

    if (!resolved) {
      console.warn('Slack target not found or invalid: ' + target);
      return;
    }

    postSlackText_(resolved.postId, message);
  });

  ui.alert('Slack message sent.');
}

function uiCleanBotHistory() {
  const ui = SpreadsheetApp.getUi();

  const box = ui.prompt(
    'Clean bot history',
    'Enter Slack name, member name, channel name, user ID, channel ID, or DM ID.',
    ui.ButtonSet.OK_CANCEL
  );

  if (box.getSelectedButton() !== ui.Button.OK) return;

  const input = String(box.getResponseText() || '').trim();

  if (!input) {
    ui.alert('Please enter a Slack name or ID.');
    return;
  }

  const resolved = resolveSlackTarget_(input, { verify: true });
  const resolvedId = resolved ? resolved.postId : input;

  const confirm = ui.alert(
    'Confirm cleanup',
    'Delete all messages/files posted by this bot in: ' + input + '?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  postSlackText_(resolvedId, '🧹 Bot cleanup started...');
  Utilities.sleep(1000);

  const result = cleanBotHistory_(resolvedId);

  ui.alert(
    'Cleanup finished.\n\nDeleted messages: ' +
    result.deletedMessages +
    '\nDeleted files: ' +
    result.deletedFiles
  );
}

function buildSlackTableMessage_(rangeA1, title) {
  const values = SpreadsheetApp.getActive().getRange(rangeA1).getDisplayValues();

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
  const channelId = getSlackConversationId_(channelOrUserId);

  const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      channel: channelId,
      text: message.title,
      blocks: message.blocks
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.ok) throw new Error(`Slack API error: ${result.error}`);
  return result;
}

function postSlackText_(channelOrUserId, text) {
  const token = getSlackToken_();
  const channelId = getSlackConversationId_(channelOrUserId);

  const response = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      channel: channelId,
      text: text
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.ok) throw new Error(`Slack API error: ${result.error}`);
  return result;
}

function postSlackChartPng_(channelOrUserId, chartSource, title) {
  const channelId = getSlackConversationId_(channelOrUserId);
  const blob = getChartPngBlobFromSource_(chartSource, title);
  uploadSlackPng_(channelId, blob, title);
}

function getChartPngBlobFromSource_(chartSource, title) {
  const info = getChartFromSource_(chartSource);
  return info.chart.getAs('image/png').setName((title || 'chart') + '.png');
}

function getSlackConversationId_(id) {
  const text = String(id || '').trim();

  if (/^[CGD]/i.test(text)) return text;

  const token = getSlackToken_();

  const response = UrlFetchApp.fetch('https://slack.com/api/conversations.open', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      users: text
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    throw new Error(`Slack conversations.open error: ${result.error}`);
  }

  return result.channel.id;
}

function uploadSlackPng_(channelId, blob, title) {
  const token = getSlackToken_();
  const bytes = blob.getBytes();

  if (!bytes || bytes.length <= 0) {
    throw new Error('Chart PNG is empty.');
  }

  const safeTitle = String(title || 'chart').replace(/[^\w\- ]+/g, '').trim() || 'chart';
  const filename = safeTitle + '.png';

  const urlResponse = UrlFetchApp.fetch('https://slack.com/api/files.getUploadURLExternal', {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: {
      filename: filename,
      length: String(bytes.length)
    },
    muteHttpExceptions: true
  });

  const uploadInfo = JSON.parse(urlResponse.getContentText());

  if (!uploadInfo.ok) {
    throw new Error(
      'Slack getUploadURLExternal error: ' +
      uploadInfo.error +
      ' / response=' +
      urlResponse.getContentText()
    );
  }

  const uploadResponse = UrlFetchApp.fetch(uploadInfo.upload_url, {
    method: 'post',
    contentType: 'application/octet-stream',
    payload: bytes,
    muteHttpExceptions: true
  });

  if (uploadResponse.getResponseCode() < 200 || uploadResponse.getResponseCode() >= 300) {
    throw new Error('Slack raw upload failed: ' + uploadResponse.getContentText());
  }

  const completeResponse = UrlFetchApp.fetch('https://slack.com/api/files.completeUploadExternal', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + token
    },
    payload: JSON.stringify({
      channel_id: channelId,
      initial_comment: safeTitle,
      files: [
        {
          id: uploadInfo.file_id,
          title: safeTitle
        }
      ]
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(completeResponse.getContentText());

  if (!result.ok) {
    throw new Error(
      'Slack completeUploadExternal error: ' +
      result.error +
      ' / response=' +
      completeResponse.getContentText()
    );
  }

  return result;
}

function getSlackTargetMap_() {
  if (SLACK_TARGET_CACHE) return SLACK_TARGET_CACHE;

  const sh = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${SETTINGS_SHEET}`);

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
  const resolved = resolveSlackTarget_(name, { verify: false });
  return resolved ? resolved.id : null;
}

function getSlackIdFromSheet_(name) {
  const target = String(name || '').trim().replace(/^[@#]/, '').toLowerCase();
  return getSlackTargetMap_()[target] || null;
}

function resolveSlackTarget_(value, options) {
  options = options || {};

  const input = String(value || '').trim();
  if (!input) return null;

  const sheetId = getSlackIdFromSheet_(input);

  if (sheetId) {
    if (options.verify && !verifySlackId_(sheetId)) return null;

    return {
      input: input,
      id: sheetId,
      postId: getSlackConversationId_(sheetId),
      source: 'sheet'
    };
  }

  if (isSlackId_(input)) {
    if (options.verify && !verifySlackId_(input)) return null;

    return {
      input: input,
      id: input,
      postId: getSlackConversationId_(input),
      source: 'direct_id'
    };
  }

  const user = findSlackUserByUsername_(input);

  if (user && user.id) {
    return {
      input: input,
      id: user.id,
      postId: getSlackConversationId_(user.id),
      source: 'slack_profile'
    };
  }

  return null;
}

function verifySlackId_(id) {
  const text = String(id || '').trim();

  if (/^[UW]/i.test(text)) return verifySlackUserId_(text);
  if (/^[CDG]/i.test(text)) return verifySlackConversationId_(text);

  return false;
}

function verifySlackUserId_(userId) {
  const token = getSlackToken_();

  const response = UrlFetchApp.fetch(
    'https://slack.com/api/users.info?user=' + encodeURIComponent(userId),
    {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(response.getContentText());
  return !!(result.ok && result.user && !result.user.deleted);
}

function verifySlackConversationId_(channelId) {
  const token = getSlackToken_();

  const response = UrlFetchApp.fetch(
    'https://slack.com/api/conversations.info?channel=' + encodeURIComponent(channelId),
    {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    }
  );

  const result = JSON.parse(response.getContentText());
  return !!result.ok;
}

function findSlackUserByUsername_(name) {
  const target = String(name || '').trim().replace(/^@/, '').toLowerCase();
  if (!target) return null;

  const token = getSlackToken_();
  let cursor = null;

  do {
    const params = { limit: 200 };
    if (cursor) params.cursor = cursor;

    const url =
      'https://slack.com/api/users.list?' +
      Object.keys(params)
        .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
        .join('&');

    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: {
        Authorization: 'Bearer ' + token
      },
      muteHttpExceptions: true
    });

    const result = JSON.parse(response.getContentText());

    if (!result.ok) {
      throw new Error('Slack users.list error: ' + result.error);
    }

    const found = (result.members || []).find(user => {
      if (!user || user.deleted || user.is_bot) return false;

      const profile = user.profile || {};

      const candidates = [
        user.name,
        profile.display_name,
        profile.real_name,
        profile.display_name_normalized,
        profile.real_name_normalized
      ]
        .map(v => String(v || '').trim().toLowerCase())
        .filter(Boolean);

      return candidates.includes(target);
    });

    if (found) return found;

    cursor = result.response_metadata && result.response_metadata.next_cursor;
  } while (cursor);

  return null;
}

function getReportRules_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${SETTINGS_SHEET}`);

  return sh.getRange(RULE_RANGE)
    .getDisplayValues()
    .filter(row => String(row[0] || '').trim() !== '')
    .map(row => ({
      rule_name: String(row[0] || '').trim(),
      type: String(row[1] || '').trim().toLowerCase(),
      target_names: String(row[2] || '').trim(),
      report_source: String(row[3] || '').trim(),
      day: String(row[4] || '').trim(),
      time: String(row[5] || '').trim(),
      enabled: String(row[6] || '').trim().toUpperCase()
    }));
}

function cleanBotHistory_(channelOrUserId) {
  const channelId = getSlackConversationId_(channelOrUserId);
  const botUserId = getBotUserId_();

  let cursor = null;
  let deletedMessages = 0;
  let deletedFiles = 0;

  do {
    const history = getSlackHistory_(channelId, cursor);

    (history.messages || []).forEach(msg => {
      const isBotMessage =
        msg.user === botUserId ||
        msg.bot_id ||
        msg.subtype === 'bot_message';

      if (!isBotMessage || !msg.ts) return;

      if (msg.files && msg.files.length) {
        msg.files.forEach(file => {
          if (file.id && deleteSlackFile_(file.id)) deletedFiles++;
        });
      }

      if (deleteSlackMessage_(channelId, msg.ts)) deletedMessages++;
    });

    cursor = history.response_metadata && history.response_metadata.next_cursor;
  } while (cursor);

  return {
    channelId,
    deletedMessages,
    deletedFiles
  };
}

function getSlackHistory_(channelId, cursor) {
  const token = getSlackToken_();
  const params = { channel: channelId, limit: 100 };

  if (cursor) params.cursor = cursor;

  const url =
    'https://slack.com/api/conversations.history?' +
    Object.keys(params)
      .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(params[k]))
      .join('&');

  const response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: {
      Authorization: `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());
  if (!result.ok) throw new Error(`Slack history error: ${result.error}`);
  return result;
}

function deleteSlackMessage_(channelId, ts) {
  const token = getSlackToken_();

  const response = UrlFetchApp.fetch('https://slack.com/api/chat.delete', {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: JSON.stringify({
      channel: channelId,
      ts: ts
    }),
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    console.warn(`Delete message failed: ${result.error} / ${ts}`);
    return false;
  }

  return true;
}

function deleteSlackFile_(fileId) {
  const token = getSlackToken_();

  const response = UrlFetchApp.fetch('https://slack.com/api/files.delete', {
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`
    },
    payload: {
      file: fileId
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    console.warn(`Delete file failed: ${result.error} / ${fileId}`);
    return false;
  }

  return true;
}

function getBotUserId_() {
  const token = getSlackToken_();

  const response = UrlFetchApp.fetch('https://slack.com/api/auth.test', {
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`
    },
    muteHttpExceptions: true
  });

  const result = JSON.parse(response.getContentText());

  if (!result.ok) {
    throw new Error(`Slack auth.test error: ${result.error}`);
  }

  return result.user_id;
}

function parseChartSource_(source) {
  const text = String(source || '').trim();
  const match = text.match(/^([^!,:]+):(\d+)$/);

  if (!match) return null;

  const sheetName = match[1].trim();
  const chartIndex = Number(match[2]);

  if (!CHART_SHEETS.includes(sheetName)) return null;

  if (!Number.isInteger(chartIndex) || chartIndex < 0) return null;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return null;

  return {
    sheetName,
    chartIndex
  };
}

function getChartFromSource_(source) {
  const parsed = parseChartSource_(source);

  if (!parsed) {
    throw new Error('Invalid chart source: ' + source + '. Example: Chart:0 or Chart1:1');
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(parsed.sheetName);
  const chart = sheet.getCharts()[parsed.chartIndex];

  if (!chart) {
    throw new Error('Chart not found: ' + parsed.sheetName + ':' + parsed.chartIndex);
  }

  return {
    chart,
    sheetName: parsed.sheetName,
    chartIndex: parsed.chartIndex
  };
}

function isSlackId_(value) {
  return /^[UCDGW][A-Z0-9]+$/i.test(String(value || '').trim());
}

function validateReportTargets_(targetNames, errors) {
  const targets = String(targetNames || '')
    .split(',')
    .map(t => t.trim())
    .filter(Boolean);

  if (!targets.length) {
    errors.push('Missing target_names.');
    return;
  }

  targets.forEach(target => {
    const resolved = resolveSlackTarget_(target, { verify: true });

    if (!resolved) {
      errors.push(
        'Unknown or invalid Slack target: ' +
        target +
        '. Allowed: member name from MEMBER_RANGE, channel name from CHANNEL_RANGE, Slack profile username, @username, or real Slack ID.'
      );
    }
  });
}