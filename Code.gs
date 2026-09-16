/**
 * JOB LEAD ENGINE (Google Sheets + Apps Script)
 * ============================================================
 * Version:  1.0.0
 * Built:    2026-09-15
 * Author:   Alexander Rea <alexander@alexanderrea.com>
 * Built with Claude (Anthropic), in Cowork.
 *
 * ------------------------------------------------------------
 * WHAT IT DOES
 * ------------------------------------------------------------
 * 1. Searches job listings (LinkedIn, Indeed, Glassdoor, company sites) via the JSearch API
 * 2. Filters by title / location / pay, then has Claude score each job against your brief
 * 3. Pulls full job details, and a pay estimate when the listing shows none
 * 4. Finds likely hiring contacts and their emails with Hunter.io
 * 5. Cross-references your LinkedIn connections export and your Gmail history
 * 6. Has Claude pick the best contact and draft a short intro email
 * 7. Writes everything to the Leads tab; tick "Send" to email from Gmail, now or in a set window
 * 8. Logs outreach to Insightly (lead + note + tag), records it on the Sent Log tab, and watches for replies
 *
 * ------------------------------------------------------------
 * SERVICES USED
 * ------------------------------------------------------------
 * JSearch (RapidAPI or OpenWeb Ninja)   Job listings, job details, salary estimates
 *     https://rapidapi.com/letscrape-6bRBa3QguO5/api/jsearch
 *     Endpoints: /search-v2, /job-details, /estimated-salary, /company-job-salary
 * Hunter.io                             Hiring contacts and email addresses
 *     https://hunter.io/api-documentation  -  Endpoints: /v2/domain-search, /v2/account
 * Anthropic (Claude)                    Job fit scoring, contact choice, email drafting
 *     https://docs.claude.com  -  Endpoint: /v1/messages
 * Insightly CRM                         Leads, notes and tags for every send
 *     https://api.na1.insightly.com/v3.1  -  Endpoints: /Leads, /Leads/{id}/Notes, /Leads/{id}/Tags
 * Google Apps Script services           Sheets (the whole workbook), Gmail (drafts, sends, reply
 *     tracking, prior-contact history), Drive (resume attachment), Triggers (schedules)
 *
 * ------------------------------------------------------------
 * SETUP
 * ------------------------------------------------------------
 * API keys live on the "API Keys" tab of the sheet. Everything else is on "Settings".
 * Menu: Job Leads > Set up sheet > Set API keys > Test API keys > Check setup > Install schedule.
 */

// ============================================================
// CONSTANTS
// ============================================================

const SCRIPT_VERSION = '1.0.0';
const SCRIPT_BUILD = '2026-09-15';

const SHEETS = { SETTINGS: 'Settings', KEYS: 'API Keys', LEADS: 'Leads', SENT: 'Sent Log', FILTERED: 'Filtered',
  CONNECTIONS: 'Connections', LOG: 'Log', SEEN: 'Seen' };

// Sent Log: one line per email that went out, newest at the bottom
const SENT_COLUMNS = ['#', 'Date', 'Company', 'Job Title', 'Job Posting', 'Status', 'Lead Row', 'Job Key'];
const SL = {};
SENT_COLUMNS.forEach(function (h, i) { SL[h] = i + 1; });

// Filtered tab: every job a run skipped, with the reason. Tick "Add to Leads" to rescue one.
const FILTERED_COLUMNS = ['Add to Leads', 'Reason', 'Fit', 'Fit Reason', 'Job Title', 'Company', 'Location', 'Salary',
  'Posted', 'Job Link', 'Filtered At', 'Added At', 'Job Key', 'Job Data'];
const F = {};
FILTERED_COLUMNS.forEach(function (h, i) { F[h] = i + 1; });
const FILTERED_MAX_ROWS = 2000;

const LEAD_COLUMNS = [
  'Status', 'Send', 'Fit', 'Fit Reason', 'Job Title', 'Company', 'Location', 'Work Type',
  'Salary', 'Salary Estimate', 'Posted', 'Job Link', 'Apply Options', 'Role Details',
  'Contact Name', 'Contact Title', 'Contact Email',
  'Email Confidence', 'Warm Connections', 'Email History', 'Subject', 'Email Body',
  'Warm Intro Ask', 'Sent At', 'Insightly Lead ID', 'Notes', 'Found At',
  'Gmail Thread ID', 'Job Key', 'Job Data'
];
const C = {}; // column name -> 1-based index
LEAD_COLUMNS.forEach(function (h, i) { C[h] = i + 1; });
const HIDDEN_COLUMNS = ['Gmail Thread ID', 'Job Key', 'Job Data'];

const STATUS = {
  ENRICHING: 'Enriching',   // found + scored, waiting for contact lookup and draft
  READY: 'Ready',           // draft + email found, ready to send
  QUEUED: 'Queued',         // you ticked Send and it goes out at the next send window
  NO_EMAIL: 'No Email',     // draft written, but no email found (use LinkedIn or apply link)
  DRAFTED: 'Drafted',       // Gmail draft created (SEND_MODE = draft)
  SENT: 'Sent',
  REPLIED: 'Replied',
  MEETING: 'Meeting',
  PASSED: 'Passed',
  ERROR: 'Error'
};

const API_KEYS = [
  ['JSEARCH_API_KEY', 'JSearch key (RapidAPI or OpenWeb Ninja)'],
  ['HUNTER_API_KEY', 'Hunter.io API key'],
  ['ANTHROPIC_API_KEY', 'Anthropic (Claude) API key'],
  ['INSIGHTLY_API_KEY', 'Insightly API key']
];
const API_KEY_NAMES = API_KEYS.map(function (k) { return k[0]; });

const TIME_BUDGET_MS = 4.5 * 60 * 1000; // Apps Script hard limit is 6 minutes per run

const DEFAULT_SETTINGS = [
  ['SEARCH_QUERIES',
    'Head of Innovation\nCreative Director\nVP Creative Technology\nDirector of Emerging Technology',
    'One search per line. Each runs once per location below. Replace these with the roles you want.'],
  ['LOCATIONS', 'New York, NY\nRemote', 'One per line. A "Remote" line searches remote-only jobs (see WORK_FROM_HOME).'],
  ['COUNTRY', 'us', 'Two-letter country code.'],
  ['DATE_POSTED', 'week', 'all | today | 3days | week | month'],
  ['EMPLOYMENT_TYPES', 'FULLTIME,CONTRACTOR,PARTTIME', 'Comma list: FULLTIME, CONTRACTOR, PARTTIME, INTERN. Blank = all types.'],
  ['WORK_FROM_HOME', 'mixed', 'mixed = follow LOCATIONS (a "Remote" line searches remote-only) | only = remote jobs only | none = skip remote jobs'],
  ['RADIUS_MILES', '50', 'Distance around each location (not used for Remote). Blank = no radius. Google treats it as a guide, not a hard cutoff.'],
  ['ALLOWED_STATES', '', 'Non-remote jobs must be in one of these states, e.g. NY,CT,NJ. Blank = anywhere in COUNTRY.'],
  ['MIN_SALARY', '', 'Annual USD. Drops jobs whose LISTED pay tops out below this (hourly/monthly pay is annualized). Jobs with no listed pay are kept. Blank = off.'],
  ['SALARY_ESTIMATE', 'company', 'company = company-specific estimate, falls back to market | market = title + location estimate | off. Only runs when the listing shows no pay.'],
  ['SALARY_EXPERIENCE', 'ABOVE_FIFTEEN', 'Experience band for estimates: ALL, LESS_THAN_ONE, ONE_TO_THREE, FOUR_TO_SIX, SEVEN_TO_NINE, TEN_TO_FOURTEEN, ABOVE_FIFTEEN'],
  ['JOB_DETAILS', 'yes', 'yes = pull full details for each new lead (full description, direct apply links, seniority, industry, employer rating) | no. 1 JSearch request per lead.'],
  ['TITLE_INCLUDE',
    'innovation, creative technolog, creative director, chief creative, emerging, immersive, experiential, head of creative, head of content, generative, ai, content supply, dco, production, technology',
    'Title must contain at least one. Matches word starts ("technolog" = Technologist). Blank = allow all.'],
  ['TITLE_EXCLUDE',
    'intern, junior, associate, coordinator, assistant, specialist, software engineer, developer, nurse, teacher, sales representative',
    'Titles with any of these whole words are skipped ("intern" does not skip "International").'],
  ['MIN_FIT_SCORE', '7', 'Claude scores 1-10. Below this is dropped.'],
  ['MAX_NEW_LEADS_PER_RUN', '15', 'Caps Hunter credits and Claude spend per run.'],
  ['TARGET_ROLE_BRIEF',
    'Describe the roles you want: seniority, discipline, the kind of company, full-time or contract, and where. ' +
    'Add what is NOT a fit. Claude scores every job against this, so be specific.',
    'What a good lead looks like. Rewrite this before your first run.'],
  ['MY_NAME', 'Your Name', 'Used as the sender name on outgoing email.'],
  ['MY_PITCH',
    'Your background in a paragraph: years of experience, the companies and clients that carry weight, ' +
    'two or three pieces of work worth naming, and the skills this search is about. Claude only uses what you put here.',
    'Background Claude draws proof points from. Rewrite this before your first run.'],
  ['EMAIL_STYLE',
    '90-150 words. Confident, warm, specific, no hype. One concrete hook about the company or role, one or two proof points that map to it, and a clear ask for a 15-minute conversation. Plain text, no bullet points, no em dashes.',
    'Tone rules for drafts.'],
  ['EMAIL_SIGNATURE', 'Your Name\nyourwebsite.com | linkedin.com/in/you\nyou@example.com', 'Appended to every email when it sends.'],
  ['SEND_MODE', 'send', 'send = email when Send is ticked | draft = create a Gmail draft instead'],
  ['SEND_TIMING', 'now', 'now = go as soon as you tick Send | window = hold until the next SEND_WINDOW day and hour'],
  ['SEND_WINDOW_DAYS', 'TUESDAY,WEDNESDAY,THURSDAY', 'Days queued emails go out (weekday names).'],
  ['SEND_WINDOW_HOUR', '9', 'Hour queued emails go out (0-23, sheet time zone). Google fires the timer somewhere inside that hour, not on the dot.'],
  ['MAX_SENDS_PER_WINDOW', '25', 'Most emails to release in one window. Gmail allows 100/day on a personal account, 1,500 on Workspace.'],
  ['BCC_ADDRESS', '', 'Optional. e.g. your Insightly email dropbox address so the email attaches to the record.'],
  ['ATTACH_RESUME', 'no', 'yes = attach RESUME_FILE to every outreach email | no. Attachments to strangers are more likely to be filtered as spam, so a link in your signature is the safer default.'],
  ['RESUME_FILE', '', 'Google Drive file name, ID or share link for the resume (e.g. Resume.pdf). Needed when ATTACH_RESUME is yes.'],
  ['JOBS_API', 'rapidapi', 'rapidapi | openwebninja (where you bought your JSearch key)'],
  ['JOBS_API_URL', 'https://jsearch.p.rapidapi.com/search-v2', 'OpenWeb Ninja direct: https://api.openwebninja.com/jsearch/search-v2'],
  ['HUNTER_SENIORITY', 'senior,executive', 'junior, senior, executive (comma list)'],
  ['HUNTER_DEPARTMENTS', 'executive,management,hr,marketing,design', 'Hunter department filters (comma list).'],
  ['SCORE_MODEL', 'claude-haiku-4-5-20251001', 'Fast, cheap model for fit scoring.'],
  ['DRAFT_MODEL', 'claude-sonnet-5', 'Model for contact choice + email drafting.'],
  ['INSIGHTLY_POD', 'na1', 'Found in Insightly > User Settings > API (e.g. na1, eu1).'],
  ['INSIGHTLY_TAG', 'Job-Outreach', 'Tag added to every lead the script creates. Insightly does not allow spaces, so they become hyphens.'],
  ['INSIGHTLY_ENABLED', 'yes', 'yes | no'],
  ['SEARCH_DAYS', 'MONDAY,THURSDAY', 'Days the job search runs (weekday names).'],
  ['SEARCH_HOUR', '7', 'Hour of day (0-23, sheet time zone).']
];

// ============================================================
// MENU + SETUP
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Job Leads')
    .addItem('1. Set up sheet', 'setupSheet')
    .addItem('2. Set API keys', 'setApiKeys')
    .addItem('3. Test API keys', 'testApiKeys')
    .addItem('Check setup', 'checkSetup')
    .addItem('4. Install schedule', 'installTriggers')
    .addSeparator()
    .addItem('Run job search now', 'runJobSearch')
    .addItem('Find contacts + draft emails now', 'enrichLeads')
    .addItem('Rewrite drafts for selected rows', 'redraftSelected')
    .addItem('Retry stuck leads', 'retryStuckLeads')
    .addItem('Send all ticked leads', 'sendTickedLeads')
    .addItem('Send queued leads now', 'sendQueuedNow')
    .addItem('Check for replies', 'checkReplies')
    .addItem('Refresh Sent Log', 'refreshSentLog')
    .addSeparator()
    .addItem('Forget scored jobs (re-score next run)', 'clearSeenJobs')
    .addItem('Remove schedule', 'removeTriggers')
    .addToUi();
}

function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  const settings = addMissingSettings_();
  settings.setColumnWidth(1, 200);
  settings.setColumnWidth(2, 520);
  settings.setColumnWidth(3, 360);
  settings.getRange('B:C').setWrap(true);
  settings.getRange('A:C').setVerticalAlignment('top');

  // Leads (also migrates an older column layout)
  const leads = getOrCreateSheet_(SHEETS.LEADS);
  ensureLeadLayout_(leads, true);

  // API Keys (moved out of script storage so they travel with the sheet)
  ensureKeysSheet_();

  // Sent Log
  syncSentLog_();

  // Filtered
  ensureFilteredSheet_();

  // Connections
  const conn = getOrCreateSheet_(SHEETS.CONNECTIONS);
  if (conn.getLastRow() === 0) {
    conn.getRange('A1').setValue('Paste your LinkedIn Connections.csv here (File > Import > Replace current sheet). ' +
      'The script finds the "First Name" header row automatically.');
  }

  // Seen (every job already scored, so rejected jobs are not re-scored next run)
  const seen = getOrCreateSheet_(SHEETS.SEEN);
  if (seen.getLastRow() === 0) seen.getRange(1, 1, 1, 4).setValues([['Job Key', 'Fit', 'Reason', 'Scored At']]).setFontWeight('bold');
  seen.hideSheet();

  // Log
  const log = getOrCreateSheet_(SHEETS.LOG);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 3).setValues([['Time', 'Level', 'Message']]).setFontWeight('bold');
    log.setFrozenRows(1);
    log.setColumnWidth(3, 700);
  }

  ss.setActiveSheet(settings);
  toast_('Sheet ready. Next: Job Leads > Set API keys.');
}

/** Adds any Settings keys that are missing (e.g. after a script update). Never overwrites values. */
function addMissingSettings_() {
  const settings = getOrCreateSheet_(SHEETS.SETTINGS);
  if (settings.getLastRow() === 0) {
    settings.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'Notes']]).setFontWeight('bold');
    settings.setFrozenRows(1);
  }
  const existingKeys = settings.getLastRow() > 1
    ? settings.getRange(2, 1, settings.getLastRow() - 1, 1).getValues().map(function (r) { return String(r[0]); })
    : [];
  const missing = DEFAULT_SETTINGS.filter(function (d) { return existingKeys.indexOf(d[0]) === -1; });
  if (missing.length) {
    settings.getRange(settings.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
    if (existingKeys.length) log_('INFO', 'Added new settings: ' + missing.map(function (m) { return m[0]; }).join(', '));
  }
  return settings;
}

/**
 * Makes sure the Leads tab matches LEAD_COLUMNS. If an older layout is found, moves every
 * existing value into its new column by header name, then reapplies formatting.
 */
function ensureLeadLayout_(sheet, forceFormat) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const matches = LEAD_COLUMNS.every(function (h, i) { return header[i] === h; }) &&
    header.slice(LEAD_COLUMNS.length).every(function (h) { return !h; });
  if (matches) {
    if (forceFormat) formatLeadsSheet_(sheet);
    return false;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow > 1 && header[0] === 'Status') {
    const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const rows = data.map(function (r) {
      return LEAD_COLUMNS.map(function (h) { const i = header.indexOf(h); return i === -1 ? '' : r[i]; });
    });
    const width = Math.max(lastCol, LEAD_COLUMNS.length);
    sheet.getRange(2, 1, lastRow - 1, width).clearContent();
    sheet.getRange(2, 1, rows.length, LEAD_COLUMNS.length).setValues(rows);
    log_('INFO', 'Leads tab migrated to the new column layout (' + rows.length + ' rows).');
  }
  if (lastCol > LEAD_COLUMNS.length) {
    sheet.getRange(1, LEAD_COLUMNS.length + 1, 1, lastCol - LEAD_COLUMNS.length).clearContent();
  }
  formatLeadsSheet_(sheet);
  return true;
}

function formatLeadsSheet_(leads) {
  leads.getRange(1, 1, 1, LEAD_COLUMNS.length).setValues([LEAD_COLUMNS])
    .setFontWeight('bold').setBackground('#1f1f1f').setFontColor('#ffffff');
  leads.setFrozenRows(1);
  leads.setFrozenColumns(2);
  const widths = { 'Status': 90, 'Send': 55, 'Fit': 45, 'Fit Reason': 260, 'Job Title': 240, 'Company': 170,
    'Location': 140, 'Salary Estimate': 260, 'Job Link': 160, 'Apply Options': 260, 'Role Details': 240,
    'Warm Connections': 220, 'Email History': 160, 'Subject': 220, 'Email Body': 420, 'Warm Intro Ask': 280, 'Notes': 220 };
  Object.keys(widths).forEach(function (k) { leads.setColumnWidth(C[k], widths[k]); });
  ['Email Body', 'Apply Options', 'Role Details'].forEach(function (k) { leads.getRange(2, C[k], 999, 1).setWrap(true); });
  leads.getRange(2, 1, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(STATUS).map(function (k) { return STATUS[k]; }), true)
      .setAllowInvalid(false).build());
  leads.showColumns(1, Math.max(leads.getLastColumn(), LEAD_COLUMNS.length));
  HIDDEN_COLUMNS.forEach(function (h) { leads.hideColumns(C[h]); });
  const statusRange = leads.getRange(2, 1, 999, 1);
  const colors = [[STATUS.READY, '#d9ead3'], [STATUS.NO_EMAIL, '#fff2cc'], [STATUS.SENT, '#cfe2f3'],
    [STATUS.DRAFTED, '#cfe2f3'], [STATUS.REPLIED, '#b6d7a8'], [STATUS.MEETING, '#93c47d'],
    [STATUS.ERROR, '#f4cccc'], [STATUS.PASSED, '#eeeeee'], [STATUS.ENRICHING, '#f3f3f3']];
  leads.setConditionalFormatRules(colors.map(function (c) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(c[0]).setBackground(c[1])
      .setRanges([statusRange]).build();
  }));
}

function getLeadsSheet_() {
  const sheet = getSheet_(SHEETS.LEADS);
  ensureLeadLayout_(sheet);
  return sheet;
}

/**
 * Builds the API Keys tab and, the first time, copies across any keys that were
 * stored in the script itself, so nothing has to be retyped.
 */
function ensureKeysSheet_() {
  const sheet = getOrCreateSheet_(SHEETS.KEYS);
  if (String(sheet.getRange(1, 1).getValue()) !== 'Key') {
    sheet.getRange(1, 1, 1, 3).setValues([['Key', 'Value', 'What it is']])
      .setFontWeight('bold').setBackground('#1f1f1f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 420);
    sheet.setColumnWidth(3, 320);
  }
  const rows = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues()
    .map(function (r) { return String(r[0]); }) : [];
  const props = PropertiesService.getScriptProperties();
  const missing = API_KEYS.filter(function (k) { return rows.indexOf(k[0]) === -1; })
    .map(function (k) { return [k[0], props.getProperty(k[0]) || '', k[1]]; });
  if (missing.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
    const moved = missing.filter(function (m) { return m[1]; }).length;
    if (moved) log_('INFO', 'Moved ' + moved + ' API key(s) from script storage into the API Keys tab.');
  }
  const warnRow = Math.max(sheet.getLastRow(), API_KEYS.length + 1) + 2;
  sheet.getRange(warnRow, 1).setValue(
    'Anyone you share this spreadsheet with can read these keys. Keep sharing tight, and regenerate a key in ' +
    'the provider\'s dashboard if the sheet ever goes somewhere it should not.').setFontStyle('italic');
  return sheet;
}

function setApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const sheet = ensureKeysSheet_();
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 2).getValues();
  API_KEYS.forEach(function (k) {
    const idx = values.findIndex(function (r) { return String(r[0]) === k[0]; });
    const current = idx === -1 ? '' : String(values[idx][1] || '');
    const res = ui.prompt('Set API keys', k[1] + (current ? ' (saved; leave blank to keep)' : ''), ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return;
    const val = res.getResponseText().trim();
    if (val) sheet.getRange(idx === -1 ? sheet.getLastRow() + 1 : idx + 2, 1, 1, 2).setValues([[k[0], val]]);
  });
  API_KEY_CACHE = null; // pick up the new values straight away
  toast_('API keys saved to the API Keys tab. Next: Job Leads > Test API keys.');
}

function testApiKeys() {
  const s = getSettings_();
  const results = [];
  const check = function (label, fn) {
    try { results.push('OK  ' + label + ': ' + fn()); }
    catch (err) { results.push('FAIL ' + label + ': ' + err.message); }
  };
  check('Hunter', function () {
    const d = httpJson_('get', 'https://api.hunter.io/v2/account?api_key=' + encodeURIComponent(requireKey_('HUNTER_API_KEY')));
    const req = d.data && d.data.requests && d.data.requests.searches;
    return req ? (req.available - req.used) + ' searches left this period' : 'connected';
  });
  check('Claude', function () {
    callClaude_(s.SCORE_MODEL, 'Reply with OK.', 'ping', 20);
    return 'connected';
  });
  check('Insightly', function () {
    if (!isYes_(s.INSIGHTLY_ENABLED)) return 'disabled in Settings';
    const me = insightly_('get', '/Users/Me');
    return 'connected as ' + (me.EMAIL_ADDRESS || me.FIRST_NAME || 'user');
  });
  check('Resume', function () {
    if (!isYes_(s.ATTACH_RESUME)) return 'ATTACH_RESUME is off, nothing will be attached';
    const blob = resumeBlob_(s);
    return 'will attach ' + blob.getName() + ' (' + Math.round(blob.getBytes().length / 1024) + ' KB)';
  });
  check('JSearch', function () {
    requireKey_('JSEARCH_API_KEY');
    const before = JSEARCH_REQUESTS;
    const jobs = searchJobs_(s, 'Creative Director', 'New York, NY', 'today');
    const parts = [jobs.length + ' jobs returned'];
    if (isYes_(s.JOB_DETAILS) && jobs[0] && jobs[0].job_id) {
      parts.push(fetchJobDetails_(s, jobs[0].job_id) ? 'job details OK' : 'job details empty');
    }
    const est = String(s.SALARY_ESTIMATE || '').toLowerCase();
    if (est && est !== 'off' && est !== 'no') {
      parts.push('salary: ' + fetchSalaryEstimate_(s, { title: 'Creative Director', company: 'Google', city: 'New York', state: 'NY' }).text);
    }
    return parts.join('; ') + ' (' + (JSEARCH_REQUESTS - before) + ' requests)';
  });
  SpreadsheetApp.getUi().alert(results.join('\n'));
}

/** Plain-language report: what is installed, what is set, and what is ready to send. */
function checkSetup() {
  const s = getSettings_();
  const out = ['Job Lead Engine v' + SCRIPT_VERSION + ' (built ' + SCRIPT_BUILD + ')', ''];
  const triggers = {};
  ScriptApp.getProjectTriggers().forEach(function (t) {
    triggers[t.getHandlerFunction()] = (triggers[t.getHandlerFunction()] || 0) + 1;
  });
  out.push(triggers['onLeadEdit']
    ? 'OK   Ticking Send works (the edit trigger is installed).'
    : 'FAIL Ticking Send does NOTHING: the edit trigger is missing. Run Job Leads > Install schedule.');
  out.push((triggers['runJobSearch'] ? 'OK   Searches run ' + s.SEARCH_DAYS + ' at ' + s.SEARCH_HOUR + ':00.'
    : 'FAIL No scheduled search. Run Install schedule.'));
  out.push(triggers['checkReplies'] ? 'OK   Replies are checked every 4 hours.' : 'FAIL Replies are not being checked.');

  const savedFingerprint = PropertiesService.getScriptProperties().getProperty('SCHEDULE_FINGERPRINT');
  if (savedFingerprint && savedFingerprint !== scheduleFingerprint_(s)) {
    out.push('FAIL Your day/hour settings changed since Install schedule was last run, so the timers still use the old ones. ' +
      'Run Install schedule again.');
  }
  if (String(s.SEND_TIMING || 'now').trim().toLowerCase() === 'window') {
    const when = nextSendWindow_(s);
    out.push(triggers['sendQueued']
      ? 'OK   Queued emails go out ' + s.SEND_WINDOW_DAYS + ' at ' + s.SEND_WINDOW_HOUR + ':00. Next: ' + (when ? formatWhen_(when) : 'unknown') + '.'
      : 'FAIL SEND_TIMING is "window" but no send trigger exists. Run Install schedule or nothing will leave the queue.');
  } else {
    out.push('NOTE SEND_TIMING is "now": ticking Send acts immediately, nothing is held for a window. ' +
      'Set it to "window" if you want queuing.');
  }
  out.push(String(s.SEND_MODE).toLowerCase() === 'draft'
    ? 'NOTE SEND_MODE is "draft": ticking Send creates a Gmail DRAFT, it does not send. Look in Gmail > Drafts.'
    : 'NOTE SEND_MODE is "send": ticking Send emails immediately.');
  const bccSetting = String(s.BCC_ADDRESS || '').trim();
  if (bccSetting && !isEmail_(bccSetting)) {
    out.push('FAIL BCC_ADDRESS is "' + bccSetting + '", which is not an email address. Clear that cell or put a real address in it.');
  }
  if (isYes_(s.ATTACH_RESUME)) {
    try { out.push('OK   Resume attachment: ' + resumeBlob_(s).getName()); }
    catch (err) { out.push('FAIL Resume attachment is on but the file is missing: ' + err.message + '. Nothing will send.'); }
  } else {
    out.push('NOTE Resume attachment is off.');
  }

  const sheet = getLeadsSheet_();
  const values = sheet.getDataRange().getValues();
  const counts = {};
  let readyToSend = 0, lastSent = null, ticked = 0;
  for (let r = 1; r < values.length; r++) {
    const st = values[r][C['Status'] - 1];
    if (!st) continue;
    counts[st] = (counts[st] || 0) + 1;
    if (values[r][C['Send'] - 1] === true) ticked++;
    if (st === STATUS.READY && String(values[r][C['Contact Email'] - 1] || '').indexOf('@') !== -1) readyToSend++;
    const sent = values[r][C['Sent At'] - 1];
    if (sent instanceof Date && (!lastSent || sent > lastSent)) lastSent = sent;
  }
  out.push('');
  out.push('Leads: ' + Object.keys(counts).map(function (k) { return counts[k] + ' ' + k; }).join(', '));
  out.push(readyToSend + ' lead(s) are Ready with an email address and can be sent now.');
  if (counts[STATUS.QUEUED]) out.push(counts[STATUS.QUEUED] + ' lead(s) are queued and waiting for the send window.');
  if (counts[STATUS.ENRICHING] || counts[STATUS.ERROR]) {
    out.push((counts[STATUS.ENRICHING] || 0) + ' still being enriched, ' + (counts[STATUS.ERROR] || 0) +
      ' errored. Job Leads > Retry stuck leads starts them over.');
  }
  if (ticked) out.push(ticked + ' row(s) still have Send ticked, which means those sends did not go through.');
  out.push(lastSent ? 'Last email sent or drafted: ' + lastSent : 'No email has been sent or drafted yet.');
  out.push('');
  out.push('Every send attempt, including the ones that stop, is written to the Log tab.');
  SpreadsheetApp.getUi().alert(out.join('\n'));
}

function scheduleFingerprint_(s) {
  return [s.SEARCH_DAYS, s.SEARCH_HOUR, s.SEND_WINDOW_DAYS, s.SEND_WINDOW_HOUR].join('|');
}

function installTriggers() {
  removeTriggers(true);
  const s = getSettings_();
  const ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onLeadEdit').forSpreadsheet(ss).onEdit().create();
  const hour = Math.max(0, Math.min(23, parseInt(s.SEARCH_HOUR, 10) || 7));
  splitList_(s.SEARCH_DAYS).forEach(function (day) {
    const wd = ScriptApp.WeekDay[day.toUpperCase()];
    if (!wd) throw new Error('Unknown day in SEARCH_DAYS: ' + day);
    ScriptApp.newTrigger('runJobSearch').timeBased().onWeekDay(wd).atHour(hour).create();
  });
  ScriptApp.newTrigger('enrichLeads').timeBased().everyHours(1).create();
  const sendHour = Math.max(0, Math.min(23, parseInt(s.SEND_WINDOW_HOUR, 10) || 9));
  splitList_(s.SEND_WINDOW_DAYS).forEach(function (day) {
    const wd = ScriptApp.WeekDay[day.trim().toUpperCase()];
    if (wd) ScriptApp.newTrigger('sendQueued').timeBased().onWeekDay(wd).atHour(sendHour).create();
  });
  ScriptApp.newTrigger('checkReplies').timeBased().everyHours(4).create();
  PropertiesService.getScriptProperties().setProperty('SCHEDULE_FINGERPRINT', scheduleFingerprint_(s));
  log_('INFO', 'Schedule installed: search ' + s.SEARCH_DAYS + ' at ' + hour + ':00, enrich hourly, replies every 4h, ' +
    'queued sends ' + s.SEND_WINDOW_DAYS + ' at ' + sendHour + ':00.');
  toast_('Schedule installed.');
}

function removeTriggers(silent) {
  const handlers = ['onLeadEdit', 'runJobSearch', 'enrichLeads', 'checkReplies', 'continueEnrichment', 'sendQueued'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (handlers.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  if (silent !== true) toast_('Schedule removed.');
}

// ============================================================
// STEP 1: FIND + SCORE JOBS
// ============================================================

function runJobSearch() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) { log_('WARN', 'Job search skipped: another run in progress.'); return; }
  const started = Date.now();
  try {
    addMissingSettings_();
    ensureFilteredSheet_();
    const s = getSettings_();
    const typeWords = splitList_(s.EMPLOYMENT_TYPES);
    const validTypes = employmentTypes_(s);
    if (typeWords.length && validTypes.length < typeWords.length) {
      log_('WARN', 'EMPLOYMENT_TYPES "' + s.EMPLOYMENT_TYPES + '" has words JSearch doesn\'t accept. Using: ' +
        (validTypes.length ? validTypes.join(',') : 'all job types') + '. Valid values: FULLTIME, CONTRACTOR, PARTTIME, INTERN.');
    }
    const leads = getLeadsSheet_();
    const leadKeys = leadJobKeys_(leads);
    const seen = seenJobs_();
    const seenSheet = getOrCreateSheet_(SHEETS.SEEN);

    // 1. Pull listings
    const queries = splitLines_(s.SEARCH_QUERIES);
    const locations = splitLines_(s.LOCATIONS);
    const wfh = wfhMode_(s);
    const plan = [];
    queries.forEach(function (q) {
      if (wfh === 'only') { plan.push([q, 'Remote']); return; }
      locations.forEach(function (loc) {
        if (wfh === 'none' && /^remote$/i.test(loc)) return;
        plan.push([q, loc]);
      });
    });
    const jsBefore = JSEARCH_REQUESTS;
    const tSearch = Date.now();
    const errors = {};
    let raw = [];
    plan.forEach(function (p) {
      try {
        const found = searchJobs_(s, p[0], p[1], s.DATE_POSTED);
        if (/^remote$/i.test(p[1])) found.forEach(function (j) { j._fromRemoteSearch = true; });
        raw = raw.concat(found);
      }
      catch (err) {
        const msg = String(err.message).replace(/"request_id":"[^"]*",?/, '');
        if (!errors[msg]) { errors[msg] = 0; log_('ERROR', 'JSearch "' + p[0] + '" / ' + p[1] + ': ' + err.message); }
        errors[msg]++;
      }
      Utilities.sleep(300);
    });
    Object.keys(errors).forEach(function (msg) {
      if (errors[msg] > 1) log_('ERROR', 'The same JSearch error repeated on ' + errors[msg] + ' of ' + plan.length + ' searches.');
    });

    // 2. Normalize, dedupe, filter. Every skipped job is written to the Filtered tab with its reason.
    const include = splitList_(s.TITLE_INCLUDE);
    const exclude = splitList_(s.TITLE_EXCLUDE);
    const states = splitList_(s.ALLOWED_STATES).map(function (x) { return x.toUpperCase(); });
    const minSalary = minSalary_(s);
    const minScore = Number(s.MIN_FIT_SCORE) || 7;
    const batchKeys = {};
    const candidates = [];
    const readmitted = [];
    const filtered = [];
    const count = { duplicate: 0, inLeads: 0, seenLow: 0, title: 0, remote: 0, location: 0, salary: 0, lowFit: 0, scoreFailed: 0, overCap: 0 };
    raw.map(normalizeJob_).forEach(function (job) {
      if (!job.title || !job.company) return;
      if (batchKeys[job.key]) { count.duplicate++; return; }   // same job from another search this run
      batchKeys[job.key] = true;
      if (leadKeys[job.key]) { count.inLeads++; return; }
      const reason = filterReason_(job, { include: include, exclude: exclude, states: states, wfh: wfh, minSalary: minSalary });
      if (reason) {
        count[reason.type]++;
        if (!seen[job.key]) filtered.push([reason.text, '', job]);
        return;
      }
      const prior = seen[job.key];
      if (prior) {
        // Scored on an earlier run. Re-admit it if it clears the current MIN_FIT_SCORE (e.g. you lowered it).
        if (prior.fit >= minScore) { job.fit = prior.fit; job.fitReason = prior.reason; readmitted.push(job); }
        else count.seenLow++;
        return;
      }
      candidates.push(job);
    });
    candidates.sort(function (a, b) { return String(b.posted).localeCompare(String(a.posted)); });
    const toScore = candidates.slice(0, 60);
    count.overCap = candidates.length - toScore.length; // not recorded, so they are scored next run

    const searchSecs = Math.round((Date.now() - tSearch) / 1000);

    // 3. Claude fit scoring (batched)
    const tScore = Date.now();
    const allScored = scoreJobs_(toScore, s);
    const scoreSecs = Math.round((Date.now() - tScore) / 1000);
    count.scoreFailed = toScore.length - allScored.length; // not recorded, so they are retried next run
    if (allScored.length) {
      const stamp = new Date();
      seenSheet.getRange(seenSheet.getLastRow() + 1, 1, allScored.length, 4).setValues(
        allScored.map(function (j) { return [j.key, j.fit, j.fitReason, stamp]; }));
    }
    allScored.forEach(function (j) {
      if (j.fit < minScore) { count.lowFit++; filtered.push(['Fit ' + j.fit + ' is below MIN_FIT_SCORE ' + minScore, j.fitReason, j]); }
    });
    const scored = allScored.filter(function (j) { return j.fit >= minScore; }).concat(readmitted);
    scored.sort(function (a, b) { return b.fit - a.fit; });
    const cap = Number(s.MAX_NEW_LEADS_PER_RUN) || 15;
    const keep = scored.slice(0, cap);
    scored.slice(cap).forEach(function (j) {
      filtered.push(['Over MAX_NEW_LEADS_PER_RUN (' + cap + ')', j.fitReason, j]);
    });

    // 4. Write rows
    if (keep.length) appendLeads_(leads, keep);
    writeFiltered_(filtered);

    const parts = [];
    const add = function (n, label) { if (n) parts.push(n + ' ' + label); };
    add(count.duplicate, 'duplicates across searches');
    add(count.inLeads, 'already in Leads');
    add(count.seenLow, 'scored low on an earlier run');
    add(count.title, 'title');
    add(count.remote, 'remote setting');
    add(count.location, 'location');
    add(count.salary, 'below MIN_SALARY');
    add(count.lowFit, 'below MIN_FIT_SCORE');
    add(count.scoreFailed, 'scoring failed (retry next run)');
    add(count.overCap, 'over the 60-per-run scoring cap (next run)');
    add(scored.length - keep.length, 'over MAX_NEW_LEADS_PER_RUN');
    log_('INFO', 'Search: ' + (JSEARCH_REQUESTS - jsBefore) + ' JSearch requests, ' + raw.length + ' listings, ' +
      toScore.length + ' scored, ' + readmitted.length + ' re-admitted, ' + keep.length + ' new leads. Skipped: ' +
      (parts.length ? parts.join(', ') : 'none') + '. Details in the Filtered tab. Took ' +
      Math.round((Date.now() - started) / 1000) + 's (searching ' + searchSecs + 's, scoring ' + scoreSecs + 's).');
    lock.releaseLock();

    // 5. Contacts + drafts: use leftover time if there's enough, otherwise hand off to a fresh run in a minute
    const left = TIME_BUDGET_MS - (Date.now() - started);
    if (left > 120000) enrichLeads(left);
    else if (keep.length) scheduleEnrichmentSoon_('search used most of this run\'s time');
  } catch (err) {
    log_('ERROR', 'runJobSearch failed: ' + err.message);
    try { lock.releaseLock(); } catch (e) { /* already released */ }
    throw err;
  }
}

let JSEARCH_REQUESTS = 0; // counted per execution, reported in the Log

function searchJobs_(s, query, location, datePosted) {
  const remote = /^remote$/i.test(location);
  const params = {
    query: remote ? query : query + ' in ' + location,
    num_pages: 1,
    country: s.COUNTRY || 'us',
    date_posted: datePosted || 'week'
  };
  if (remote) params.work_from_home = 'true';
  const miles = parseFloat(s.RADIUS_MILES);
  if (!remote && miles > 0) params.radius = Math.round(miles * 1.60934); // API expects km
  const types = employmentTypes_(s);
  if (types.length && !EMPLOYMENT_FILTER_LOCAL) params.employment_types = types.join(',');
  let json;
  try {
    json = jsearchGet_(s, '/search-v2', params);
  } catch (err) {
    if (!params.employment_types || !/employment.?types/i.test(err.message)) throw err;
    // Never lose a whole run over this filter: drop it for the rest of the run and filter locally instead
    EMPLOYMENT_FILTER_LOCAL = true;
    log_('WARN', 'JSearch rejected employment_types=' + params.employment_types + '. Searching without it and filtering job types locally for this run.');
    delete params.employment_types;
    json = jsearchGet_(s, '/search-v2', params);
  }
  let jobs = [];
  if (json.data && Array.isArray(json.data.jobs)) jobs = json.data.jobs; // search-v2: { data: { jobs, cursor } }
  else if (Array.isArray(json.data)) jobs = json.data;                  // legacy shape
  if (EMPLOYMENT_FILTER_LOCAL && types.length) {
    jobs = jobs.filter(function (j) { return jobTypeAllowed_(j, types); });
  }
  return jobs;
}

let EMPLOYMENT_FILTER_LOCAL = false; // set when the API rejects employment_types during a run

/** Turns the EMPLOYMENT_TYPES setting into valid JSearch values (FULLTIME, CONTRACTOR, PARTTIME, INTERN). */
function employmentTypes_(s) {
  const alias = {
    FULLTIME: 'FULLTIME', FULL: 'FULLTIME', PERMANENT: 'FULLTIME',
    CONTRACTOR: 'CONTRACTOR', CONTRACT: 'CONTRACTOR', FREELANCE: 'CONTRACTOR', FRACTIONAL: 'CONTRACTOR', TEMPORARY: 'CONTRACTOR',
    PARTTIME: 'PARTTIME', PART: 'PARTTIME',
    INTERN: 'INTERN', INTERNSHIP: 'INTERN'
  };
  const out = [];
  String(s.EMPLOYMENT_TYPES || '').split(/[,;|\n]+/).forEach(function (raw) {
    const v = alias[raw.toUpperCase().replace(/[^A-Z]/g, '')];
    if (v && out.indexOf(v) === -1) out.push(v);
  });
  return out;
}

function jobTypeAllowed_(job, types) {
  const labels = [].concat(job.job_employment_types || [], job.job_employment_type || [])
    .map(function (x) { return String(x).toUpperCase().replace(/[^A-Z]/g, ''); }).filter(Boolean);
  if (!labels.length) return true; // unknown type: keep
  return labels.some(function (l) {
    return types.some(function (t) { return l.indexOf(t === 'CONTRACTOR' ? 'CONTRACT' : t) !== -1; });
  });
}

/** One JSearch GET. Works with RapidAPI or OpenWeb Ninja keys; base comes from JOBS_API_URL. */
function jsearchGet_(s, path, params) {
  const key = requireKey_('JSEARCH_API_KEY');
  // Accept any JSearch URL in Settings (older /search, /search-v2, or a bare base URL)
  const base = String(s.JOBS_API_URL || 'https://jsearch.p.rapidapi.com').trim().replace(/\/+$/, '')
    .replace(/\/(search|search-v2|job-details|estimated-salary|company-job-salary)$/, '');
  const headers = String(s.JOBS_API).toLowerCase() === 'openwebninja'
    ? { 'x-api-key': key }
    : { 'X-RapidAPI-Key': key, 'X-RapidAPI-Host': hostOf_(base) };
  JSEARCH_REQUESTS++;
  return httpJson_('get', base + path + '?' + toQuery_(params), null, headers);
}

function wfhMode_(s) {
  const v = String(s.WORK_FROM_HOME || 'mixed').trim().toLowerCase();
  if (/^(only|yes|true|remote)$/.test(v)) return 'only';
  if (/^(none|no|false|onsite|on-site|exclude)$/.test(v)) return 'none';
  return 'mixed';
}

/** Full job details: longer description, apply options, seniority, industry, employer reviews. */
function fetchJobDetails_(s, jobId) {
  const json = jsearchGet_(s, '/job-details', { job_id: jobId, country: s.COUNTRY || 'us' });
  return Array.isArray(json.data) ? (json.data[0] || null) : (json.data || null);
}

function applyJobDetails_(job, d) {
  if (!d) return;
  job.detailsFetched = true;
  if (d.job_description && String(d.job_description).length > String(job.description || '').length) {
    job.description = String(d.job_description).slice(0, 8000);
  }
  if (!job.website && d.employer_website) job.website = d.employer_website;
  if (!job.applyLink && d.job_apply_link) job.applyLink = d.job_apply_link;
  if (!job.salary && (d.job_min_salary || d.job_max_salary)) {
    job.minSalary = d.job_min_salary || null;
    job.maxSalary = d.job_max_salary || null;
    job.salaryPeriod = d.job_salary_period || '';
    job.salary = formatSalary_(job.minSalary, job.maxSalary, job.salaryPeriod);
  }
  job.applyOptions = (Array.isArray(d.apply_options) ? d.apply_options : []).slice(0, 6).map(function (o) {
    return (o.publisher || 'Link') + (o.is_direct ? ' (direct)' : '') + ': ' + o.apply_link;
  }).join('\n');

  const lines = [];
  const add = function (label, v) { const t = detailText_(v); if (t) lines.push(label + ': ' + t); };
  add('Seniority', d.seniority_level);
  add('Experience (yrs)', d.required_experience_years);
  add('Industry', d.industry);
  add('Function', d.job_function);
  add('Manages people', d.has_management_responsibilities);
  add('Arrangement', d.work_arrangement);
  add('Contract length', d.contract_duration);
  add('Start', d.start_date);
  add('AI/ML involved', d.ai_ml_involved);
  const review = (Array.isArray(d.employer_reviews) ? d.employer_reviews : [])[0];
  if (review && review.score) {
    lines.push('Employer rating: ' + review.score + '/' + (review.max_score || 5) + ' on ' + (review.publisher || 'reviews') +
      (review.review_count ? ' (' + review.review_count + ' reviews)' : ''));
  }
  job.roleDetails = lines.join('\n');
}

function detailText_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'boolean') return v ? 'yes' : 'no';
  if (Array.isArray(v)) return v.map(detailText_).filter(Boolean).join(', ');
  if (typeof v === 'object') {
    return Object.keys(v).map(function (k) { const t = detailText_(v[k]); return t ? k + ' ' + t : ''; }).filter(Boolean).join(', ');
  }
  return String(v);
}

/** Salary estimate for listings with no pay. company mode falls back to market data. */
function fetchSalaryEstimate_(s, job) {
  const mode = String(s.SALARY_ESTIMATE || '').trim().toLowerCase();
  const place = [job.city, job.state].filter(Boolean).join(', ');
  const common = {
    job_title: job.title,
    location: place || 'United States',
    location_type: place ? 'ANY' : 'COUNTRY',
    years_of_experience: String(s.SALARY_EXPERIENCE || 'ALL').trim().toUpperCase()
  };
  let best = null, source = '';
  if (mode === 'company') {
    try {
      best = pickSalary_(jsearchGet_(s, '/company-job-salary', Object.assign({ company: job.company }, common)));
      source = job.company + ' data';
    } catch (err) { best = null; }
  }
  if (!best) {
    best = pickSalary_(jsearchGet_(s, '/estimated-salary', common));
    source = 'market data, ' + (place || 'US');
  }
  if (!best) return { text: 'No estimate found', flag: '' };
  const period = best.salary_period || 'YEAR';
  const text = formatSalary_(best.min_salary, best.max_salary, period) +
    (best.median_salary ? ', median ' + money_(best.median_salary) : '') +
    ' (' + source + (best.salary_count ? ', ' + best.salary_count + ' salaries' : '') +
    (best.confidence ? ', ' + String(best.confidence).toLowerCase() + ' confidence' : '') + ')';
  const min = minSalary_(s);
  const top = annualize_(best.max_salary || best.median_salary, period);
  return { text: text, flag: min && top && top < min ? 'Estimated pay tops out below MIN_SALARY.' : '' };
}

function pickSalary_(json) {
  const rows = (json && Array.isArray(json.data)) ? json.data : [];
  return rows.filter(function (r) { return r && (r.median_salary || r.max_salary); })
    .sort(function (a, b) { return (b.salary_count || 0) - (a.salary_count || 0); })[0] || null;
}

function minSalary_(s) {
  const n = parseFloat(String(s.MIN_SALARY || '').replace(/[^0-9.]/g, ''));
  return n > 0 ? n : 0;
}

function annualize_(amount, period) {
  const n = Number(amount);
  if (!n) return 0;
  const p = String(period || '').toUpperCase();
  const mult = { YEAR: 1, MONTH: 12, WEEK: 52, DAY: 260, HOUR: 2080 }[p];
  return n * (mult || (n < 1000 ? 2080 : 1));
}

function salaryMeetsMin_(job, min) {
  if (!min) return true;
  const top = job.maxSalary || job.minSalary;
  if (!top) return true; // no listed pay: keep it
  return annualize_(top, job.salaryPeriod) >= min;
}

function normalizeJob_(r) {
  const city = r.job_city || '';
  const state = r.job_state || '';
  const job = {
    id: r.job_id || '',
    title: String(r.job_title || '').trim(),
    company: String(r.employer_name || '').trim(),
    website: r.employer_website || '',
    applyLink: r.job_apply_link || r.job_google_link || '',
    publisher: r.job_publisher || '',
    remote: !!(r.job_is_remote || /remote/i.test(r.work_arrangement || '')),
    fromRemoteSearch: !!r._fromRemoteSearch, // found by a remote-only search, so its listed city doesn't matter
    city: city,
    state: state,
    location: r.job_location || [city, state].filter(Boolean).join(', '),
    type: r.job_employment_type || (r.job_employment_types || []).join(', '),
    salary: formatSalary_(r.job_min_salary, r.job_max_salary, r.job_salary_period),
    minSalary: r.job_min_salary || null,
    maxSalary: r.job_max_salary || null,
    salaryPeriod: r.job_salary_period || '',
    posted: r.job_posted_at_datetime_utc || r.job_posted_at || '',
    description: String(r.job_description || '').slice(0, 3500)
  };
  job.key = normName_(job.company) + '|' + normName_(job.title);
  return job;
}

const STATE_NAMES = { NY: 'new york', CT: 'connecticut', NJ: 'new jersey', MA: 'massachusetts', PA: 'pennsylvania',
  RI: 'rhode island', CA: 'california', IL: 'illinois', TX: 'texas', FL: 'florida', GA: 'georgia', WA: 'washington' };

function locationMatches_(job, allowedStates) {
  if (job.remote || job.fromRemoteSearch) return true;
  if (!allowedStates.length) return true;
  const st = String(job.state || '').trim();
  const loc = String(job.location || '').toLowerCase().trim();
  // unknown or country-wide location: let Claude judge
  if (!st && (!loc || /^(united states|usa|us|u\.s\.|anywhere|nationwide|remote|multiple locations)$/.test(loc))) return true;
  return allowedStates.some(function (code) {
    const name = STATE_NAMES[code] || '';
    return st.toUpperCase() === code || (name && st.toLowerCase() === name) ||
      new RegExp('\\b' + code.toLowerCase() + '\\b').test(loc) || (name && loc.indexOf(name) !== -1);
  });
}

function appendLeads_(leads, jobs, fitReasonPrefix) {
  const now = new Date();
  const rows = jobs.map(function (j) {
    const row = new Array(LEAD_COLUMNS.length).fill('');
    row[C['Status'] - 1] = STATUS.ENRICHING;
    row[C['Send'] - 1] = false;
    row[C['Fit'] - 1] = j.fit === undefined ? '' : j.fit;
    row[C['Fit Reason'] - 1] = (fitReasonPrefix || '') + (j.fitReason || '');
    row[C['Job Title'] - 1] = j.title;
    row[C['Company'] - 1] = j.company;
    row[C['Location'] - 1] = j.remote ? 'Remote' + (j.location ? ' / ' + j.location : '') : j.location;
    row[C['Work Type'] - 1] = j.type;
    row[C['Salary'] - 1] = j.salary;
    row[C['Posted'] - 1] = j.posted ? String(j.posted).slice(0, 10) : '';
    row[C['Job Link'] - 1] = j.applyLink;
    row[C['Found At'] - 1] = now;
    row[C['Job Key'] - 1] = j.key;
    row[C['Job Data'] - 1] = JSON.stringify(j);
    return row;
  });
  const start = leads.getLastRow() + 1;
  leads.getRange(start, 1, rows.length, LEAD_COLUMNS.length).setValues(rows);
  leads.getRange(start, C['Send'], rows.length, 1).insertCheckboxes();
}

/** Returns null if the job passes, or { type, text } explaining why it was skipped. */
function filterReason_(job, o) {
  const t = titleCheck_(job.title, o.include, o.exclude);
  if (t) return { type: 'title', text: t };
  if (o.wfh === 'only' && !job.remote && !job.fromRemoteSearch) return { type: 'remote', text: 'Not remote (WORK_FROM_HOME is "only")' };
  if (o.wfh === 'none' && job.remote) return { type: 'remote', text: 'Remote (WORK_FROM_HOME is "none")' };
  if (!locationMatches_(job, o.states)) {
    return { type: 'location', text: 'Location "' + (job.location || job.state) + '" is outside ALLOWED_STATES (' + o.states.join(', ') + ')' };
  }
  if (!salaryMeetsMin_(job, o.minSalary)) {
    const top = annualize_(job.maxSalary || job.minSalary, job.salaryPeriod);
    return { type: 'salary', text: 'Listed pay ' + job.salary + ' (about ' + money_(top) + '/yr) is below MIN_SALARY ' + money_(o.minSalary) };
  }
  return null;
}

/** Title rules: TITLE_EXCLUDE words match whole words only ("intern" skips "Intern", not "International").
 *  TITLE_INCLUDE words match the start of a word ("technolog" matches "Technologist"); words of 3 letters or fewer
 *  must match exactly. Add * to either list to match a word start explicitly (e.g. "develop*"). */
function titleCheck_(title, include, exclude) {
  const t = String(title || '').toLowerCase();
  for (let i = 0; i < exclude.length; i++) {
    if (termMatches_(t, exclude[i], true)) return 'Title contains excluded word "' + exclude[i].trim() + '"';
  }
  if (!include.length) return '';
  const hit = include.some(function (x) { return termMatches_(t, x, x.replace(/\*$/, '').trim().length <= 3); });
  return hit ? '' : 'Title has none of the TITLE_INCLUDE words';
}

function termMatches_(text, term, wholeWord) {
  const raw = String(term || '').trim().toLowerCase();
  if (!raw) return false;
  const star = /\*$/.test(raw);
  const core = raw.replace(/\*+$/, '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&').replace(/\s+/g, '\\s+');
  if (!core) return false;
  const end = wholeWord && !star ? '(?![a-z0-9])' : '';
  return new RegExp('(^|[^a-z0-9])' + core + end).test(text);
}

function leadJobKeys_(leadsSheet) {
  const out = {};
  if (leadsSheet.getLastRow() > 1) {
    leadsSheet.getRange(2, C['Job Key'], leadsSheet.getLastRow() - 1, 1).getValues()
      .forEach(function (r) { if (r[0]) out[r[0]] = true; });
  }
  return out;
}

/** key -> { fit, reason } for every job scored on an earlier run. */
function seenJobs_() {
  const out = {};
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SEEN);
  if (!sheet || sheet.getLastRow() < 2) return out;
  sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues().forEach(function (r) {
    if (r[0]) out[r[0]] = { fit: Number(r[1]) || 0, reason: String(r[2] || '') };
  });
  return out;
}

function clearSeenJobs() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.SEEN);
  if (sheet && sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent();
  log_('INFO', 'Cleared scored-job memory. The next search re-scores everything it finds.');
  toast_('Scored jobs forgotten. The next search re-scores everything.');
}

function ensureFilteredSheet_() {
  const sheet = getOrCreateSheet_(SHEETS.FILTERED);
  const header = sheet.getRange(1, 1, 1, FILTERED_COLUMNS.length).getValues()[0];
  if (String(header[0]) !== FILTERED_COLUMNS[0]) {
    sheet.getRange(1, 1, 1, FILTERED_COLUMNS.length).setValues([FILTERED_COLUMNS])
      .setFontWeight('bold').setBackground('#1f1f1f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    const widths = { 'Add to Leads': 90, 'Reason': 320, 'Fit Reason': 300, 'Job Title': 260, 'Company': 170, 'Location': 150, 'Job Link': 160 };
    Object.keys(widths).forEach(function (k) { sheet.setColumnWidth(F[k], widths[k]); });
    sheet.hideColumns(F['Job Key']);
    sheet.hideColumns(F['Job Data']);
  }
  return sheet;
}

function writeFiltered_(items) {
  if (!items.length) return;
  const sheet = ensureFilteredSheet_();
  // Don't repeat a job that is already listed with the same reason (searches return the same listings run after run)
  const listed = {};
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, FILTERED_COLUMNS.length).getValues().forEach(function (r) {
      listed[r[F['Job Key'] - 1] + '||' + r[F['Reason'] - 1]] = true;
    });
  }
  items = items.filter(function (it) {
    const k = it[2].key + '||' + it[0];
    if (listed[k]) return false;
    listed[k] = true;
    return true;
  });
  if (!items.length) return;
  const now = new Date();
  const rows = items.map(function (it) {
    const j = it[2];
    const row = new Array(FILTERED_COLUMNS.length).fill('');
    row[F['Add to Leads'] - 1] = false;
    row[F['Reason'] - 1] = it[0];
    row[F['Fit'] - 1] = j.fit === undefined ? '' : j.fit;
    row[F['Fit Reason'] - 1] = it[1] || '';
    row[F['Job Title'] - 1] = j.title;
    row[F['Company'] - 1] = j.company;
    row[F['Location'] - 1] = j.remote ? 'Remote' + (j.location ? ' / ' + j.location : '') : j.location;
    row[F['Salary'] - 1] = j.salary;
    row[F['Posted'] - 1] = j.posted ? String(j.posted).slice(0, 10) : '';
    row[F['Job Link'] - 1] = j.applyLink;
    row[F['Filtered At'] - 1] = now;
    row[F['Job Key'] - 1] = j.key;
    row[F['Job Data'] - 1] = JSON.stringify(j);
    return row;
  });
  const start = sheet.getLastRow() + 1;
  sheet.getRange(start, 1, rows.length, FILTERED_COLUMNS.length).setValues(rows);
  sheet.getRange(start, F['Add to Leads'], rows.length, 1).insertCheckboxes();
  const excess = sheet.getLastRow() - 1 - FILTERED_MAX_ROWS;
  if (excess > 0) sheet.deleteRows(2, excess); // keep the newest rows
}

/** Tick "Add to Leads" on the Filtered tab: the job moves into Leads and gets contacts + a draft about a minute later. */
function addFilteredRow_(sheet, rowNum) {
  const row = sheet.getRange(rowNum, 1, 1, FILTERED_COLUMNS.length).getValues()[0];
  if (row[F['Added At'] - 1]) { sheet.getRange(rowNum, F['Add to Leads']).setValue(false); return; }
  const job = safeJson_(row[F['Job Data'] - 1]);
  if (!job.company) { toast_('Row ' + rowNum + ': no job data to add.'); return; }
  const leads = getLeadsSheet_();
  if (leadJobKeys_(leads)[job.key]) {
    toast_(job.title + ' is already in Leads.');
  } else {
    appendLeads_(leads, [job], 'Added from Filtered (' + row[F['Reason'] - 1] + '). ');
    scheduleEnrichmentSoon_('added from Filtered');
    toast_('Added "' + job.title + '" to Leads. Contacts and a draft arrive in a minute or two.');
  }
  sheet.getRange(rowNum, F['Add to Leads']).setValue(false);
  sheet.getRange(rowNum, F['Added At']).setValue(new Date());
}

function scoreJobs_(jobs, s) {
  const out = [];
  const system = 'You screen job listings for one specific executive candidate. Be strict: score 8-10 only when ' +
    'seniority, discipline and location all fit. Return JSON only.';
  for (let i = 0; i < jobs.length; i += 10) {
    const batch = jobs.slice(i, i + 10);
    const listing = batch.map(function (j, n) {
      return '### Job ' + n + '\nTitle: ' + j.title + '\nCompany: ' + j.company + '\nLocation: ' +
        (j.remote ? 'Remote ' : '') + j.location + '\nType: ' + j.type + '\nSalary: ' + (j.salary || 'n/a') +
        '\nDescription: ' + j.description.slice(0, 1200);
    }).join('\n\n');
    const user = 'CANDIDATE BRIEF:\n' + s.TARGET_ROLE_BRIEF + '\n\nCANDIDATE BACKGROUND:\n' + s.MY_PITCH +
      '\n\nJOBS:\n' + listing +
      '\n\nReturn a JSON array, one object per job: [{"job": <number>, "score": <1-10>, "reason": "<max 20 words>"}]';
    try {
      const arr = parseJson_(callClaude_(s.SCORE_MODEL, system, user, 1500));
      (Array.isArray(arr) ? arr : []).forEach(function (r) {
        const j = batch[Number(r.job)];
        if (!j) return;
        j.fit = Number(r.score) || 0;
        j.fitReason = String(r.reason || '');
        out.push(j);
      });
    } catch (err) {
      log_('ERROR', 'Scoring batch failed: ' + err.message);
    }
  }
  return out;
}

// ============================================================
// STEP 2: ENRICH (contacts, warm paths, draft email)
// ============================================================

function enrichLeads(budgetMs) {
  const budget = typeof budgetMs === 'number' ? budgetMs : TIME_BUDGET_MS;
  const started = Date.now();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    // Another run (usually a search) holds the lock: try again shortly instead of waiting for the hourly run
    scheduleEnrichmentSoon_('another run was in progress');
    return;
  }
  let remaining = 0;
  try {
    const s = getSettings_();
    const sheet = getLeadsSheet_();
    const values = sheet.getDataRange().getValues();
    const pending = [];
    for (let r = 1; r < values.length; r++) {
      if (values[r][C['Status'] - 1] === STATUS.ENRICHING) pending.push(r);
    }
    if (!pending.length) return;
    // Best fits first, so the leads you most want get contacts and drafts soonest
    pending.sort(function (a, b) { return (Number(values[b][C['Fit'] - 1]) || 0) - (Number(values[a][C['Fit'] - 1]) || 0); });
    const connections = loadConnections_();
    const jsStart = JSEARCH_REQUESTS;
    let done = 0, failed = 0;
    for (let i = 0; i < pending.length; i++) {
      const r = pending[i];
      if (Date.now() - started > budget - 75000) break; // one lead can take ~60s (details, salary, Hunter, Claude)
      // Record the attempt before doing the work: if Google stops the run mid-lead, the count survives
      const prior = String(values[r][C['Notes'] - 1] || '').match(/^Enrichment attempt (\d+)/);
      const attempt = (prior ? parseInt(prior[1], 10) : 0) + 1;
      if (attempt > 3) {
        setCells_(sheet, r + 1, { 'Status': STATUS.ERROR, 'Notes': 'Gave up after 3 attempts. Use Retry stuck leads to try again.' });
        log_('WARN', 'Row ' + (r + 1) + ' (' + values[r][C['Company'] - 1] + ') gave up after 3 attempts.');
        failed++;
        continue;
      }
      setCells_(sheet, r + 1, { 'Notes': 'Enrichment attempt ' + attempt + ' started ' + formatDate_(new Date()) + '...' });
      try {
        enrichRow_(sheet, r + 1, values[r], s, connections);
        done++;
      } catch (err) {
        failed++;
        setCells_(sheet, r + 1, { 'Status': STATUS.ERROR, 'Notes': 'Enrich failed: ' + err.message });
        log_('ERROR', 'Row ' + (r + 1) + ' (' + values[r][C['Company'] - 1] + ') enrich failed: ' + err.message);
      }
    }
    remaining = pending.length - done - failed;
    log_('INFO', 'Enrichment: ' + done + ' lead(s) finished' + (failed ? ', ' + failed + ' failed' : '') + ' in ' +
      Math.round((Date.now() - started) / 1000) + 's, ' + (JSEARCH_REQUESTS - jsStart) + ' JSearch requests. ' +
      (remaining ? remaining + ' still pending.' : 'None pending.'));
  } finally {
    lock.releaseLock();
  }
  if (remaining) scheduleEnrichmentSoon_(remaining + ' lead(s) still pending');
  else deleteTriggersFor_('continueEnrichment'); // nothing left, so drop any queued follow-up run
}

/** Puts anything stuck or failed back in the queue and starts a fresh attempt. */
function retryStuckLeads() {
  const sheet = getLeadsSheet_();
  const values = sheet.getDataRange().getValues();
  let n = 0;
  for (let r = 1; r < values.length; r++) {
    const status = values[r][C['Status'] - 1];
    if (status !== STATUS.ENRICHING && status !== STATUS.ERROR) continue;
    setCells_(sheet, r + 1, { 'Status': STATUS.ENRICHING, 'Notes': '' });
    n++;
  }
  if (!n) { toast_('Nothing is stuck.'); return; }
  log_('INFO', 'Retrying ' + n + ' stuck or failed lead(s).');
  toast_('Retrying ' + n + ' lead(s). This runs in the background.');
  enrichLeads();
}

/** One-off trigger target: picks up enrichment about a minute after the previous run stopped. */
function continueEnrichment() {
  deleteTriggersFor_('continueEnrichment');
  enrichLeads();
}

function scheduleEnrichmentSoon_(why) {
  try {
    deleteTriggersFor_('continueEnrichment');
    ScriptApp.newTrigger('continueEnrichment').timeBased().after(60 * 1000).create();
    log_('INFO', 'Enrichment continues in about a minute (' + why + ').');
  } catch (err) {
    log_('WARN', 'Could not schedule the next enrichment run (' + err.message + '). Use Job Leads > Find contacts + draft emails now.');
  }
}

function deleteTriggersFor_(handler) {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === handler) ScriptApp.deleteTrigger(t);
  });
}

function enrichRow_(sheet, rowNum, rowValues, s, connections) {
  const job = JSON.parse(rowValues[C['Job Data'] - 1] || '{}');
  if (!job.company) throw new Error('Missing job data');
  const notes = [];

  // Full job details (JOB_DETAILS)
  if (isYes_(s.JOB_DETAILS) && job.id && !job.detailsFetched) {
    try { applyJobDetails_(job, fetchJobDetails_(s, job.id)); }
    catch (err) { notes.push('Job details failed: ' + err.message); }
    if (job.salary && !salaryMeetsMin_(job, minSalary_(s))) notes.push('Listed pay tops out below MIN_SALARY.');
  }

  // Salary estimate when the listing shows no pay (SALARY_ESTIMATE)
  let salaryEstimate = '';
  const estMode = String(s.SALARY_ESTIMATE || '').trim().toLowerCase();
  if (!job.salary && estMode && estMode !== 'off' && estMode !== 'no') {
    try {
      const est = fetchSalaryEstimate_(s, job);
      salaryEstimate = est.text;
      if (est.flag) notes.push(est.flag);
    } catch (err) { notes.push('Salary estimate failed: ' + err.message); }
  }

  // Contacts from Hunter
  const hunter = findContacts_(job, s);
  const contacts = hunter.contacts;
  const domain = hunter.domain;

  // Warm paths
  const warm = findWarmConnections_(connections, job.company, contacts);
  const domainHistory = gmailHistory_(domain ? domain : '');

  // Claude chooses contact + drafts
  const draft = draftEmail_(job, contacts, warm, domainHistory, s);
  const idx = Number(draft.contact_index);
  const chosen = (idx >= 0 && idx < contacts.length) ? contacts[idx] : null;

  let history = domainHistory.threads ? domainHistory.threads + ' thread(s) with @' + domain : 'None found';
  if (chosen && chosen.email) {
    const personal = gmailHistory_(chosen.email);
    if (personal.threads) history = personal.threads + ' thread(s) with ' + chosen.email + '; ' + history;
  }

  setCells_(sheet, rowNum, {
    'Status': chosen && chosen.email ? STATUS.READY : STATUS.NO_EMAIL,
    'Salary': job.salary || '',
    'Salary Estimate': salaryEstimate,
    'Apply Options': job.applyOptions || '',
    'Role Details': job.roleDetails || '',
    'Job Data': JSON.stringify(job),
    'Contact Name': chosen ? chosen.name : '',
    'Contact Title': chosen ? chosen.position : '',
    'Contact Email': chosen ? chosen.email : '',
    'Email Confidence': chosen ? chosen.confidence : '',
    'Warm Connections': warm.map(function (w) { return w.label; }).join('\n'),
    'Email History': history,
    'Subject': draft.subject || '',
    'Email Body': draft.body || '',
    'Warm Intro Ask': draft.warm_intro_ask || '',
    'Notes': [draft.contact_reason, contacts.length ? '' : 'Hunter found no contacts; try LinkedIn "Meet the hiring team" or the apply link.']
      .concat(notes).filter(Boolean).join(' ')
  });
}

function findContacts_(job, s) {
  const key = requireKey_('HUNTER_API_KEY');
  const domain = domainOf_(job.website);
  const base = { api_key: key, limit: 10, type: 'personal' };
  if (domain) base.domain = domain; else base.company = job.company;

  const filtered = Object.assign({}, base, {
    seniority: String(s.HUNTER_SENIORITY || '').replace(/\s/g, ''),
    department: String(s.HUNTER_DEPARTMENTS || '').replace(/\s/g, '')
  });
  let res = hunterSearch_(filtered);
  // Hunter only charges when emails come back, so an unfiltered retry is cheap
  if (!res.emails.length) res = hunterSearch_(base);
  return {
    domain: res.domain || domain,
    contacts: res.emails.map(function (e) {
      return {
        name: [e.first_name, e.last_name].filter(Boolean).join(' '),
        first: e.first_name || '',
        last: e.last_name || '',
        position: e.position || '',
        seniority: e.seniority || '',
        department: e.department || '',
        email: e.value || '',
        confidence: e.confidence || '',
        linkedin: e.linkedin || ''
      };
    })
  };
}

function hunterSearch_(params) {
  Object.keys(params).forEach(function (k) { if (params[k] === '' || params[k] == null) delete params[k]; });
  const url = 'https://api.hunter.io/v2/domain-search?' + toQuery_(params);
  const json = httpJson_('get', url, null, null, [400, 404, 422]);
  const d = json.data || {};
  return { domain: d.domain || '', emails: Array.isArray(d.emails) ? d.emails : [] };
}

function loadConnections_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.CONNECTIONS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  let h = -1;
  for (let i = 0; i < Math.min(values.length, 15); i++) {
    if (values[i].map(String).indexOf('First Name') !== -1) { h = i; break; }
  }
  if (h === -1) return [];
  const head = values[h].map(String);
  const col = function (name) { return head.indexOf(name); };
  const fi = col('First Name'), la = col('Last Name'), url = col('URL'), co = col('Company'), po = col('Position'), em = col('Email Address');
  return values.slice(h + 1).filter(function (r) { return r[fi] || r[la]; }).map(function (r) {
    return {
      name: (String(r[fi]) + ' ' + String(r[la])).trim(),
      company: co >= 0 ? String(r[co]) : '',
      position: po >= 0 ? String(r[po]) : '',
      url: url >= 0 ? String(r[url]) : '',
      email: em >= 0 ? String(r[em]) : ''
    };
  });
}

function findWarmConnections_(connections, company, contacts) {
  const target = normCompany_(company);
  const contactNames = {};
  contacts.forEach(function (c) { if (c.name) contactNames[c.name.toLowerCase()] = c; });
  const out = [];
  connections.forEach(function (p) {
    const pc = normCompany_(p.company);
    const sameCompany = pc && target && (pc === target ||
      (Math.min(pc.length, target.length) >= 5 && (pc.indexOf(target) === 0 || target.indexOf(pc) === 0)));
    const isContact = !!contactNames[p.name.toLowerCase()];
    if (sameCompany || isContact) {
      out.push({
        name: p.name, position: p.position, url: p.url, isHiringContact: isContact,
        label: p.name + (p.position ? ' (' + p.position + ')' : '') + (isContact ? ' [Hunter contact, 1st-degree]' : '') +
          (p.url ? ' ' + p.url : '')
      });
    }
  });
  return out.slice(0, 8);
}

function gmailHistory_(addressOrDomain) {
  const generic = /^(gmail|yahoo|outlook|hotmail|icloud|aol|linkedin|indeed|glassdoor)\./i;
  if (!addressOrDomain || generic.test(addressOrDomain)) return { threads: 0 };
  try {
    const threads = GmailApp.search('from:' + addressOrDomain + ' OR to:' + addressOrDomain, 0, 20);
    return { threads: threads.length };
  } catch (err) {
    return { threads: 0 };
  }
}

function draftEmail_(job, contacts, warm, history, s, extra) {
  const system = 'You are a sharp executive-search writer drafting one-to-one outreach for a senior creative ' +
    'technology leader. Never invent facts about the candidate, the company or any relationship. Return JSON only.';
  const contactList = contacts.length
    ? contacts.map(function (c, i) {
      return i + ': ' + c.name + ' | ' + c.position + ' | seniority=' + c.seniority + ' | dept=' + c.department +
        ' | confidence=' + c.confidence;
    }).join('\n')
    : '(none found)';
  const warmText = warm.length
    ? warm.map(function (w) { return '- ' + w.name + (w.position ? ', ' + w.position : '') + (w.isHiringContact ? ' (also in contact list)' : ''); }).join('\n')
    : '(none)';
  const user =
    'CANDIDATE: ' + s.MY_NAME + '\nBACKGROUND:\n' + s.MY_PITCH +
    '\n\nJOB:\nTitle: ' + job.title + '\nCompany: ' + job.company + '\nLocation: ' + (job.remote ? 'Remote ' : '') + job.location +
    '\nType: ' + job.type + '\nLink: ' + job.applyLink +
    (job.roleDetails ? '\nRole details:\n' + job.roleDetails : '') + '\nDescription:\n' + job.description +
    '\n\nPOSSIBLE CONTACTS AT THE COMPANY:\n' + contactList +
    '\n\nCANDIDATE\'S LINKEDIN CONNECTIONS AT THIS COMPANY:\n' + warmText +
    '\n\nPRIOR EMAIL THREADS WITH THIS COMPANY\'S DOMAIN: ' + (history.threads || 0) +
    '\n\nRESUME: ' + (attachResume_(s)
      ? 'attached to this email. You may refer to it once, briefly, near the end.'
      : 'NOT attached. Never say a resume, CV or attachment is enclosed.') +
    '\n\nTASKS:\n' +
    '1. Pick the contact most likely to own this hire (the hiring manager or the executive the role reports to; ' +
    'for small companies the founder/CEO; head of talent only if no better option). Use -1 if none are plausible.\n' +
    '2. Write the email to that person (greet by first name; if -1, write it to the hiring team). STYLE: ' + s.EMAIL_STYLE +
    ' Do not include a signature. If a LinkedIn connection is also the chosen contact, you may note "we\'re connected on LinkedIn". ' +
    'Do not claim to know anyone personally.\n' +
    '3. If there are LinkedIn connections at the company who are NOT the chosen contact, write a 2-3 sentence LinkedIn ' +
    'message asking the most relevant one for a quick intro or a word about the role. Otherwise empty string.\n\n' +
    (extra ? 'EXTRA INSTRUCTIONS FROM THE CANDIDATE (these win over the style rules above): ' + extra + '\n\n' : '') +
    'Return JSON: {"contact_index": <number>, "contact_reason": "<max 15 words>", "subject": "<max 8 words>", ' +
    '"body": "<email body, plain text, \\n for line breaks>", "warm_intro_ask": "<string>"}';
  return parseJson_(callClaude_(s.DRAFT_MODEL, system, user, 3000));
}

// ============================================================
// STEP 3: SEND (tick the checkbox) + LOG TO INSIGHTLY
// ============================================================

/** Installable onEdit trigger (created by Install schedule). */
function onLeadEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  if (sheet.getName() === SHEETS.FILTERED) {
    if (F['Add to Leads'] < range.getColumn() || F['Add to Leads'] > range.getColumn() + range.getNumColumns() - 1) return;
    for (let r = Math.max(2, range.getRow()); r <= range.getRow() + range.getNumRows() - 1; r++) {
      if (sheet.getRange(r, F['Add to Leads']).getValue() === true) addFilteredRow_(sheet, r);
    }
    return;
  }
  if (sheet.getName() !== SHEETS.LEADS) return;
  ensureLeadLayout_(sheet);
  const firstCol = range.getColumn();
  const lastCol = firstCol + range.getNumColumns() - 1;
  if (C['Send'] < firstCol || C['Send'] > lastCol) return;
  const rowStart = Math.max(2, range.getRow());
  const rowEnd = range.getRow() + range.getNumRows() - 1;
  for (let r = rowStart; r <= rowEnd; r++) {
    if (sheet.getRange(r, C['Send']).getValue() === true) sendLeadRow_(sheet, r);
  }
}

function sendTickedLeads() {
  const sheet = getLeadsSheet_();
  const values = sheet.getDataRange().getValues();
  let n = 0;
  for (let r = 1; r < values.length; r++) {
    if (values[r][C['Send'] - 1] === true) { sendLeadRow_(sheet, r + 1); n++; }
  }
  toast_(n + ' ticked lead(s) processed.');
}

function sendLeadRow_(sheet, rowNum, force) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return;
  try {
    const s = getSettings_();
    const row = sheet.getRange(rowNum, 1, 1, LEAD_COLUMNS.length).getValues()[0];
    const get = function (k) { return row[C[k] - 1]; };
    const status = get('Status');
    const to = String(get('Contact Email') || '').trim();
    const subject = String(get('Subject') || '').trim();
    const body = String(get('Email Body') || '').trim();

    const fail = function (msg) {
      setCells_(sheet, rowNum, { 'Send': false, 'Notes': msg });
      log_('WARN', 'Nothing sent for row ' + rowNum + ' (' + get('Company') + '): ' + msg);
      toast_('Row ' + rowNum + ': ' + msg);
    };
    if ([STATUS.SENT, STATUS.DRAFTED, STATUS.REPLIED, STATUS.MEETING].indexOf(status) !== -1) {
      return fail('Row is already marked ' + status + ', so nothing was sent. If you set that status by hand, ' +
        'change it back to Ready and tick Send again.');
    }
    if (!isEmail_(to)) return fail('No valid Contact Email. Add one, then tick Send again.');
    if (!subject || !body) return fail('Subject or Email Body is empty.');

    if (!force && String(s.SEND_TIMING || 'now').trim().toLowerCase() === 'window') {
      const when = nextSendWindow_(s);
      setCells_(sheet, rowNum, {
        'Status': STATUS.QUEUED,
        'Send': false,
        'Notes': 'Queued. Goes out ' + (when ? formatWhen_(when) : 'at the next send window') + '.'
      });
      toast_('Queued for ' + (when ? formatWhen_(when) : 'the next send window') + '.');
      return;
    }

    const fullBody = body + '\n\n' + String(s.EMAIL_SIGNATURE || s.MY_NAME || '');
    const opts = { name: s.MY_NAME || undefined };
    const bcc = String(s.BCC_ADDRESS || '').trim();
    if (bcc && isEmail_(bcc)) opts.bcc = bcc;
    else if (bcc) log_('WARN', 'BCC_ADDRESS is "' + bcc + '", which is not an email address, so it was ignored. Clear that cell or put a real address in it.');
    if (attachResume_(s)) {
      // Don't send at all if the promised attachment is missing: the email usually refers to it
      try { opts.attachments = [resumeBlob_(s)]; }
      catch (err) { return fail('Resume not attached, nothing sent: ' + err.message); }
    }

    let threadId, newStatus;
    try {
      const draft = GmailApp.createDraft(to, subject, fullBody, opts);
      if (String(s.SEND_MODE).toLowerCase() === 'draft') {
        threadId = draft.getMessage().getThread().getId();
        newStatus = STATUS.DRAFTED;
      } else {
        threadId = draft.send().getThread().getId();
        newStatus = STATUS.SENT;
      }
    } catch (err) {
      // Never let one bad row stop a whole send window
      return fail('Gmail refused this email: ' + err.message);
    }
    const updates = { 'Status': newStatus, 'Send': false, 'Sent At': new Date(), 'Gmail Thread ID': threadId };
    setCells_(sheet, rowNum, updates);
    log_('INFO', (newStatus === STATUS.DRAFTED ? 'Gmail draft created (see Drafts)' : 'Email sent') + ': ' +
      get('Job Title') + ' @ ' + get('Company') + ' -> ' + to + (opts.attachments ? ' with resume attached' : ''));

    try { addToSentLog_(sheet, rowNum); }
    catch (err) { log_('WARN', 'Sent Log update failed: ' + err.message); }

    if (isYes_(s.INSIGHTLY_ENABLED)) {
      try {
        const leadId = logToInsightly_(s, {
          name: get('Contact Name'), title: get('Contact Title'), email: to, company: get('Company'),
          jobTitle: get('Job Title'), jobLink: get('Job Link'), subject: subject, body: fullBody,
          website: safeJson_(get('Job Data')).website || '', status: newStatus
        });
        setCells_(sheet, rowNum, { 'Insightly Lead ID': leadId });
      } catch (err) {
        setCells_(sheet, rowNum, { 'Notes': 'Email ' + newStatus.toLowerCase() + ', but Insightly failed: ' + err.message });
        log_('ERROR', 'Insightly: ' + err.message);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function attachResume_(s) {
  return isYes_(s.ATTACH_RESUME) && !!String(s.RESUME_FILE || '').trim();
}

let RESUME_BLOB = null; // cached per execution

/** Finds RESUME_FILE in Google Drive by ID, share link or file name. */
function resumeBlob_(s) {
  if (RESUME_BLOB) return RESUME_BLOB;
  const ref = String(s.RESUME_FILE || '').trim();
  if (!ref) throw new Error('RESUME_FILE is empty');
  let file = null;
  const id = (ref.match(/[-\w]{25,}/) || [])[0];
  if (id) { try { file = DriveApp.getFileById(id); } catch (err) { file = null; } }
  if (!file) {
    const found = DriveApp.getFilesByName(ref);
    if (found.hasNext()) file = found.next();
  }
  if (!file) throw new Error('no Google Drive file matches "' + ref + '"');
  const blob = file.getBlob();
  if (blob.getBytes().length > 20 * 1024 * 1024) throw new Error('"' + file.getName() + '" is larger than 20MB');
  RESUME_BLOB = blob;
  return blob;
}

/**
 * Rewrites the Subject and Email Body for the rows you have selected on the Leads tab,
 * optionally following an instruction like "shorter, open with the Cannes work".
 * Uses no Hunter or JSearch credits: it only asks Claude again.
 */
function redraftSelected() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActive().getActiveSheet();
  if (sheet.getName() !== SHEETS.LEADS) { ui.alert('Select the rows you want rewritten on the Leads tab first.'); return; }
  const range = sheet.getActiveRange();
  const first = Math.max(2, range.getRow());
  const last = range.getRow() + range.getNumRows() - 1;
  if (last < 2) { ui.alert('Select the rows you want rewritten.'); return; }

  const res = ui.prompt('Rewrite ' + (last - first + 1) + ' draft(s)',
    'Anything you want changed? For example: "shorter and warmer", "lead with the Cannes work", ' +
    '"mention I can start in January". Leave blank for a straight rewrite.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const extra = res.getResponseText().trim();

  const s = getSettings_();
  let done = 0, skipped = 0;
  for (let r = first; r <= last; r++) {
    const row = sheet.getRange(r, 1, 1, LEAD_COLUMNS.length).getValues()[0];
    const get = function (k) { return row[C[k] - 1]; };
    const status = get('Status');
    if ([STATUS.SENT, STATUS.DRAFTED, STATUS.REPLIED, STATUS.MEETING].indexOf(status) !== -1 || !get('Job Title')) {
      skipped++;
      continue;
    }
    try {
      const job = safeJson_(get('Job Data'));
      if (!job.title) { job.title = get('Job Title'); job.company = get('Company'); job.description = get('Fit Reason'); }
      const contacts = get('Contact Email') ? [{
        name: get('Contact Name'), position: get('Contact Title'), email: get('Contact Email'),
        seniority: '', department: '', confidence: get('Email Confidence'), linkedin: ''
      }] : [];
      const warm = String(get('Warm Connections') || '').split('\n').filter(Boolean).map(function (line) {
        return { name: line.replace(/\s*https?:\/\/\S+/, '').trim(), position: '', isHiringContact: false };
      });
      const history = { threads: parseInt(String(get('Email History') || '0'), 10) || 0 };
      const draft = draftEmail_(job, contacts, warm, history, s, extra);
      setCells_(sheet, r, {
        'Subject': draft.subject || get('Subject'),
        'Email Body': draft.body || get('Email Body'),
        'Warm Intro Ask': draft.warm_intro_ask || get('Warm Intro Ask'),
        'Notes': 'Rewritten ' + formatDate_(new Date()) + (extra ? ': ' + extra : '') + '.'
      });
      done++;
    } catch (err) {
      log_('ERROR', 'Rewrite failed on row ' + r + ': ' + err.message);
      setCells_(sheet, r, { 'Notes': 'Rewrite failed: ' + err.message });
    }
  }
  log_('INFO', 'Rewrote ' + done + ' draft(s)' + (extra ? ' with instruction: ' + extra : '') + '.');
  ui.alert('Rewrote ' + done + ' draft(s).' + (skipped ? ' Skipped ' + skipped + ' (already sent, or empty).' : ''));
}

// ============================================================
// SENT LOG
// ============================================================

function ensureSentLogSheet_() {
  const sheet = getOrCreateSheet_(SHEETS.SENT);
  if (String(sheet.getRange(1, 1).getValue()) !== SENT_COLUMNS[0]) {
    sheet.getRange(1, 1, 1, SENT_COLUMNS.length).setValues([SENT_COLUMNS])
      .setFontWeight('bold').setBackground('#1f1f1f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    const widths = { '#': 45, 'Date': 150, 'Company': 200, 'Job Title': 280, 'Job Posting': 220, 'Status': 90, 'Lead Row': 100 };
    Object.keys(widths).forEach(function (k) { sheet.setColumnWidth(SL[k], widths[k]); });
    sheet.hideColumns(SL['Job Key']);
  }
  return sheet;
}

/** Link that jumps to this lead's row on the Leads tab. */
function leadRowLink_(leadsSheet, rowNum) {
  return '=HYPERLINK("#gid=' + leadsSheet.getSheetId() + '&range=A' + rowNum + '", "Row ' + rowNum + '")';
}

function addToSentLog_(leadsSheet, rowNum) {
  const sheet = ensureSentLogSheet_();
  const row = leadsSheet.getRange(rowNum, 1, 1, LEAD_COLUMNS.length).getValues()[0];
  const get = function (k) { return row[C[k] - 1]; };
  const key = get('Job Key');
  const last = sheet.getLastRow();
  // If this lead is already listed, refresh it instead of adding a second line
  if (last > 1) {
    const keys = sheet.getRange(2, SL['Job Key'], last - 1, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (key && String(keys[i][0]) === String(key)) {
        sheet.getRange(i + 2, SL['Status']).setValue(get('Status'));
        sheet.getRange(i + 2, SL['Lead Row']).setFormula(leadRowLink_(leadsSheet, rowNum));
        return;
      }
    }
  }
  const values = new Array(SENT_COLUMNS.length).fill('');
  values[SL['#'] - 1] = last;                       // header row means last = count of entries
  values[SL['Date'] - 1] = get('Sent At') || new Date();
  values[SL['Company'] - 1] = get('Company');
  values[SL['Job Title'] - 1] = get('Job Title');
  values[SL['Job Posting'] - 1] = get('Job Link');
  values[SL['Status'] - 1] = get('Status');
  values[SL['Job Key'] - 1] = key;
  sheet.getRange(last + 1, 1, 1, SENT_COLUMNS.length).setValues([values]);
  sheet.getRange(last + 1, SL['Lead Row']).setFormula(leadRowLink_(leadsSheet, rowNum));
}

/** Adds anything missing, refreshes statuses, and repoints the row links (rows move when you sort or delete). */
function syncSentLog_() {
  const sheet = ensureSentLogSheet_();
  const leads = getLeadsSheet_();
  const sent = [STATUS.SENT, STATUS.DRAFTED, STATUS.REPLIED, STATUS.MEETING];
  const values = leads.getDataRange().getValues();
  const byKey = {};
  for (let r = 1; r < values.length; r++) {
    const status = values[r][C['Status'] - 1];
    if (sent.indexOf(status) === -1) continue;
    const key = String(values[r][C['Job Key'] - 1] || ('row-' + (r + 1)));
    byKey[key] = { row: r + 1, values: values[r], status: status };
  }
  let updated = 0;
  const listed = {};
  if (sheet.getLastRow() > 1) {
    const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, SENT_COLUMNS.length).getValues();
    rows.forEach(function (r, i) {
      const key = String(r[SL['Job Key'] - 1] || '');
      const lead = byKey[key];
      if (!lead) return;
      listed[key] = true;
      if (r[SL['Status'] - 1] !== lead.status) {
        sheet.getRange(i + 2, SL['Status']).setValue(lead.status);
        updated++;
      }
      sheet.getRange(i + 2, SL['Lead Row']).setFormula(leadRowLink_(leads, lead.row));
      sheet.getRange(i + 2, SL['#']).setValue(i + 1);
    });
  }
  let added = 0;
  Object.keys(byKey).forEach(function (key) {
    if (listed[key]) return;
    addToSentLog_(leads, byKey[key].row);
    added++;
  });
  return { added: added, updated: updated };
}

function refreshSentLog() {
  const res = syncSentLog_();
  const msg = 'Sent Log refreshed: ' + res.added + ' added, ' + res.updated + ' status change(s).';
  log_('INFO', msg);
  toast_(msg);
}

/** Time-trigger target: releases everything sitting in the queue. */
function sendQueued() {
  const s = getSettings_();
  const sheet = getLeadsSheet_();
  const values = sheet.getDataRange().getValues();
  const cap = Number(s.MAX_SENDS_PER_WINDOW) || 25;
  let sent = 0, left = 0;
  for (let r = 1; r < values.length; r++) {
    if (values[r][C['Status'] - 1] !== STATUS.QUEUED) continue;
    if (sent >= cap) { left++; continue; }
    try { sendLeadRow_(sheet, r + 1, true); }
    catch (err) { log_('ERROR', 'Send window: row ' + (r + 1) + ' failed (' + err.message + '). Carrying on with the rest.'); }
    sent++;
    Utilities.sleep(2000); // a small gap between sends
  }
  if (sent || left) {
    log_('INFO', 'Send window: ' + sent + ' queued lead(s) released' +
      (left ? ', ' + left + ' held back by MAX_SENDS_PER_WINDOW (they go in the next window)' : '') + '.');
  }
  return sent;
}

function sendQueuedNow() {
  const n = sendQueued();
  toast_(n ? n + ' queued lead(s) sent.' : 'Nothing is queued.');
}

const DAY_INDEX = { SUNDAY: 0, MONDAY: 1, TUESDAY: 2, WEDNESDAY: 3, THURSDAY: 4, FRIDAY: 5, SATURDAY: 6 };

/** The next date and time the queue will be released, in the sheet's time zone. */
function nextSendWindow_(s) {
  const days = splitList_(s.SEND_WINDOW_DAYS).map(function (d) { return DAY_INDEX[d.trim().toUpperCase()]; })
    .filter(function (n) { return n !== undefined; });
  if (!days.length) return null;
  const hour = Math.max(0, Math.min(23, parseInt(s.SEND_WINDOW_HOUR, 10) || 9));
  const now = new Date();
  for (let i = 0; i < 8; i++) {
    const d = new Date(now.getFullYear(), now.getMonth(), now.getDate() + i, hour, 0, 0);
    if (d > now && days.indexOf(d.getDay()) !== -1) return d;
  }
  return null;
}

function formatWhen_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEEE d MMM \'at\' h:mm a');
}

function logToInsightly_(s, d) {
  const found = insightly_('get', '/Leads/Search?field_name=EMAIL&field_value=' + encodeURIComponent(d.email) + '&top=1');
  let leadId = Array.isArray(found) && found.length ? found[0].LEAD_ID : null;
  if (!leadId) {
    const parts = String(d.name || '').trim().split(/\s+/);
    const first = parts.length > 1 ? parts.slice(0, -1).join(' ') : (parts[0] || '');
    const last = parts.length > 1 ? parts[parts.length - 1] : '(unknown)';
    const lead = insightly_('post', '/Leads', {
      FIRST_NAME: first,
      LAST_NAME: last,
      EMAIL: d.email,
      TITLE: d.title || '',
      ORGANISATION_NAME: d.company,
      WEBSITE: d.website || '',
      LEAD_DESCRIPTION: 'Job outreach: ' + d.jobTitle + '\n' + d.jobLink
    });
    leadId = lead.LEAD_ID;
    // Insightly rejects tags containing spaces
    const tag = String(s.INSIGHTLY_TAG || 'Job-Outreach').trim().replace(/\s+/g, '-');
    try { insightly_('post', '/Leads/' + leadId + '/Tags', { TAG_NAME: tag }); }
    catch (err) { log_('WARN', 'Insightly tag failed: ' + err.message); }
  }
  insightly_('post', '/Leads/' + leadId + '/Notes', {
    TITLE: 'Outreach ' + d.status.toLowerCase() + ': ' + d.jobTitle,
    BODY: 'Role: ' + d.jobTitle + ' at ' + d.company + '\nLink: ' + d.jobLink + '\n\nSubject: ' + d.subject + '\n\n' + d.body
  });
  return leadId;
}

// ============================================================
// STEP 4: REPLY TRACKING
// ============================================================

function checkReplies() {
  const s = getSettings_();
  const sheet = getLeadsSheet_();
  const values = sheet.getDataRange().getValues();
  let replies = 0;
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (row[C['Status'] - 1] !== STATUS.SENT) continue;
    const threadId = row[C['Gmail Thread ID'] - 1];
    if (!threadId) continue;
    let thread;
    try { thread = GmailApp.getThreadById(threadId); } catch (err) { continue; }
    if (!thread) continue;
    const messages = thread.getMessages();
    const me = emailOf_(messages[0].getFrom());
    const reply = messages.slice(1).filter(function (m) { return emailOf_(m.getFrom()) !== me; })[0];
    if (!reply) continue;
    replies++;
    const snippet = reply.getPlainBody().replace(/\s+/g, ' ').slice(0, 300);
    setCells_(sheet, r + 1, { 'Status': STATUS.REPLIED, 'Notes': 'Reply ' + formatDate_(reply.getDate()) + ': ' + snippet });
    const leadId = row[C['Insightly Lead ID'] - 1];
    if (leadId && isYes_(s.INSIGHTLY_ENABLED)) {
      try {
        insightly_('post', '/Leads/' + leadId + '/Notes', {
          TITLE: 'Reply received: ' + row[C['Job Title'] - 1],
          BODY: 'From: ' + reply.getFrom() + '\nDate: ' + reply.getDate() + '\n\n' + reply.getPlainBody().slice(0, 5000)
        });
      } catch (err) { log_('WARN', 'Insightly reply note failed: ' + err.message); }
    }
  }
  if (replies) {
    log_('INFO', replies + ' new repl' + (replies === 1 ? 'y' : 'ies') + ' detected.');
    syncSentLog_();
  }
}

// ============================================================
// API HELPERS
// ============================================================

function callClaude_(model, system, user, maxTokens) {
  const key = requireKey_('ANTHROPIC_API_KEY');
  const payload = {
    model: model,
    max_tokens: maxTokens || 2000,
    system: system,
    messages: [{ role: 'user', content: user }]
  };
  const json = httpJson_('post', 'https://api.anthropic.com/v1/messages', payload, {
    'x-api-key': key,
    'anthropic-version': '2023-06-01'
  });
  const text = (json.content || []).filter(function (b) { return b.type === 'text'; })
    .map(function (b) { return b.text; }).join('');
  if (!text) throw new Error('Claude returned no text (stop_reason: ' + json.stop_reason + ')');
  return text;
}

function insightly_(method, path, body) {
  const s = getSettings_();
  const key = requireKey_('INSIGHTLY_API_KEY');
  const url = 'https://api.' + (s.INSIGHTLY_POD || 'na1') + '.insightly.com/v3.1' + path;
  return httpJson_(method, url, body, { Authorization: 'Basic ' + Utilities.base64Encode(key + ':') });
}

/** JSON HTTP call with retries on rate limits / overload / 5xx. */
function httpJson_(method, url, body, headers, okCodes) {
  const opts = { method: method, muteHttpExceptions: true, headers: headers || {} };
  if (body) { opts.contentType = 'application/json'; opts.payload = JSON.stringify(body); }
  let lastErr = '';
  for (let attempt = 0; attempt < 4; attempt++) {
    const res = UrlFetchApp.fetch(url, opts);
    const code = res.getResponseCode();
    const text = res.getContentText();
    if ((code >= 200 && code < 300) || (okCodes && okCodes.indexOf(code) !== -1)) {
      if (!text) return {};
      try { return JSON.parse(text); } catch (e) { return {}; }
    }
    lastErr = 'HTTP ' + code + ' ' + hostOf_(url) + ': ' + String(text).slice(0, 300);
    if ([429, 500, 502, 503, 504, 529].indexOf(code) === -1) break;
    Utilities.sleep(Math.pow(2, attempt) * 2000);
  }
  throw new Error(lastErr);
}

// ============================================================
// SHEET + UTILITY HELPERS
// ============================================================

function getSettings_() {
  const sheet = getSheet_(SHEETS.SETTINGS);
  const out = {};
  DEFAULT_SETTINGS.forEach(function (d) { out[d[0]] = d[1]; });
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0]) out[String(r[0]).trim()] = String(r[1]);
    });
  }
  return out;
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Missing "' + name + '" tab. Run Job Leads > Set up sheet.');
  return sh;
}

function getOrCreateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function setCells_(sheet, rowNum, map) {
  Object.keys(map).forEach(function (k) {
    if (!C[k]) throw new Error('Unknown column ' + k);
    const v = map[k];
    sheet.getRange(rowNum, C[k]).setValue(typeof v === 'string' ? v.slice(0, 49000) : v);
  });
}

let API_KEY_CACHE = null;

/** Keys come from the API Keys tab; anything still in script storage is used as a fallback. */
function requireKey_(name) {
  if (!API_KEY_CACHE) {
    API_KEY_CACHE = {};
    const sheet = SpreadsheetApp.getActive().getSheetByName(SHEETS.KEYS);
    if (sheet && sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues().forEach(function (r) {
        if (r[0] && r[1]) API_KEY_CACHE[String(r[0]).trim()] = String(r[1]).trim();
      });
    }
  }
  const v = API_KEY_CACHE[name] || PropertiesService.getScriptProperties().getProperty(name);
  if (!v) throw new Error(name + ' is empty. Put it in the "API Keys" tab (or use Job Leads > Set API keys).');
  return v;
}

function log_(level, msg) {
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.LOG);
    if (sh) sh.appendRow([new Date(), level, String(msg).slice(0, 5000)]);
  } catch (e) { /* logging must never break a run */ }
  console.log(level + ': ' + msg);
}

function toast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, 'Job Leads', 6); } catch (e) { console.log(msg); }
}

function parseJson_(text) {
  const t = String(text).replace(/```(?:json)?/gi, '').trim();
  const starts = [t.indexOf('{'), t.indexOf('[')].filter(function (i) { return i !== -1; });
  if (!starts.length) throw new Error('No JSON in model output: ' + t.slice(0, 200));
  const start = Math.min.apply(null, starts);
  const end = Math.max(t.lastIndexOf('}'), t.lastIndexOf(']'));
  const slice = t.slice(start, end + 1);
  try { return JSON.parse(slice); }
  catch (e) { return JSON.parse(escapeControlCharsInStrings_(slice)); }
}

/** Escapes raw newlines/tabs that appear inside JSON string literals. */
function escapeControlCharsInStrings_(src) {
  let out = '', inStr = false, esc = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (inStr) {
      if (esc) { out += ch; esc = false; continue; }
      if (ch === '\\') { out += ch; esc = true; continue; }
      if (ch === '"') { inStr = false; out += ch; continue; }
      if (ch === '\n') { out += '\\n'; continue; }
      if (ch === '\r') { continue; }
      if (ch === '\t') { out += '\\t'; continue; }
      out += ch;
    } else {
      if (ch === '"') inStr = true;
      out += ch;
    }
  }
  return out;
}

function safeJson_(v) {
  try { return JSON.parse(v || '{}') || {}; } catch (e) { return {}; }
}

function isEmail_(v) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(v || '').trim());
}

function emailOf_(from) {
  const m = String(from || '').match(/<([^>]+)>/);
  return (m ? m[1] : String(from || '')).trim().toLowerCase();
}

function splitLines_(v) {
  return String(v || '').split(/\r?\n/).map(function (x) { return x.trim(); }).filter(Boolean);
}

function splitList_(v) {
  return String(v || '').split(/[,\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
}

function lower_(x) { return String(x).toLowerCase(); }

function isYes_(v) { return /^(yes|y|true|1|on)$/i.test(String(v).trim()); }

function toQuery_(params) {
  return Object.keys(params).map(function (k) {
    // Keep commas literal: list params like employment_types=FULLTIME,CONTRACTOR are rejected when sent as %2C
    return encodeURIComponent(k) + '=' + encodeURIComponent(params[k]).replace(/%2C/gi, ',');
  }).join('&');
}

function hostOf_(url) {
  const m = String(url).match(/^https?:\/\/([^\/?#]+)/i);
  return m ? m[1] : '';
}

function domainOf_(website) {
  const host = hostOf_(/^https?:\/\//i.test(website) ? website : 'https://' + website).toLowerCase();
  return host.replace(/^www\./, '');
}

function normName_(v) {
  return String(v || '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function normCompany_(v) {
  return normName_(v)
    .replace(/\b(inc|llc|ltd|limited|corp|corporation|co|company|group|holdings|the|plc|gmbh|agency|usa|us|na)\b/g, ' ')
    .replace(/\s+/g, ' ').trim();
}

function money_(n) {
  n = Number(n);
  return n >= 1000 ? '$' + Math.round(n / 1000) + 'K' : '$' + Math.round(n * 100) / 100;
}

function formatSalary_(min, max, period) {
  if (!min && !max) return '';
  const f = money_;
  const range = min && max ? f(min) + '-' + f(max) : f(min || max);
  return range + (period ? ' / ' + String(period).toLowerCase() : '');
}

function formatDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'MMM d');
}
