/**
 * One-time setup: creates the agent's tabs (never touches your existing ones),
 * stores API keys, and manages the daily trigger.
 */

function setupTabs() {
  const ss = SpreadsheetApp.getActive();

  // Agent Settings
  let s = ss.getSheetByName(TABS.SETTINGS);
  if (!s) {
    s = ss.insertSheet(TABS.SETTINGS);
    s.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Help']]).setFontWeight('bold');
    s.getRange(2, 1, SETTINGS_DEFAULTS.length, 3).setValues(SETTINGS_DEFAULTS);
    s.setColumnWidth(1, 220); s.setColumnWidth(2, 420); s.setColumnWidth(3, 420);
    s.getRange('B:B').setWrap(true);
  }

  // Targets
  let t = ss.getSheetByName(TABS.TARGETS);
  if (!t) {
    t = ss.insertSheet(TABS.TARGETS);
    t.getRange(1, 1, 1, TARGET_HEADERS.length).setValues([TARGET_HEADERS]).setFontWeight('bold');
    const consultingTitles = 'Associate Consultant, Consultant, Senior Associate Consultant, Business Analyst, Associate';
    const penn = 'University of Pennsylvania';
    const examples = [
      [true, 'McKinsey', 'mckinsey.com', 'Consulting', consultingTitles, 'New York', penn, 5, ''],
      [true, 'BCG', 'bcg.com', 'Consulting', consultingTitles, 'New York', penn, 5, ''],
      [true, 'Bain', 'bain.com', 'Consulting', consultingTitles, 'New York', penn, 5, ''],
      [false, 'Oliver Wyman', 'oliverwyman.com', 'Consulting', consultingTitles, 'New York', penn, 3, ''],
      [false, 'EY-Parthenon', 'parthenon.ey.com', 'Consulting', consultingTitles, 'New York', penn, 3, ''],
      [false, 'L.E.K. Consulting', 'lek.com', 'Consulting', consultingTitles, 'New York', penn, 3, ''],
      [false, 'Kearney', 'kearney.com', 'Consulting', consultingTitles, 'New York', penn, 3, ''],
      [false, 'Universal Music Group', 'umusic.com', 'Music', 'Analyst, Associate, Coordinator, Manager', 'New York', penn, 3, ''],
      [false, 'Spotify', 'spotify.com', 'Music', 'Analyst, Associate, Strategy, Marketing Manager', 'New York', penn, 3, ''],
    ];
    t.getRange(2, 1, examples.length, TARGET_HEADERS.length).setValues(examples);
    t.getRange(2, 1, 200, 1).insertCheckboxes();
    t.getRange(2, 4, 200, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(TRACK_OPTIONS, true).build());
    t.setFrozenRows(1);
    t.setColumnWidth(5, 360);
  }

  // Outreach Queue
  let q = ss.getSheetByName(TABS.QUEUE);
  if (!q) {
    q = ss.insertSheet(TABS.QUEUE);
    q.getRange(1, 1, 1, QUEUE_HEADERS.length).setValues([QUEUE_HEADERS]).setFontWeight('bold');
    q.setFrozenRows(1);
    q.setFrozenColumns(3);
    const col = name => QUEUE_HEADERS.indexOf(name) + 1;
    q.getRange(2, col(Q.STATUS), 1000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(STATUS_OPTIONS, true).build());
    q.getRange(2, col(Q.CHANNEL), 1000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(CHANNEL_OPTIONS, true).build());
    q.getRange(2, col(Q.PENN), 1000, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'Likely', 'No'], true).build());
    [Q.BODY, Q.NOTE, Q.LI_DM, Q.ABOUT].forEach(n => { q.setColumnWidth(col(n), 380); q.getRange(2, col(n), 1000, 1).setWrap(true); });
    q.setColumnWidth(col(Q.SUBJECT), 240);
    // Color the Status column so the review queue is easy to scan.
    const statusRange = q.getRange(2, col(Q.STATUS), 1000, 1);
    const colors = { 'Draft Ready': '#fff2cc', 'Approved': '#cfe2f3', 'Sent': '#d9ead3', 'Replied': '#b6d7a8', 'Skip': '#eeeeee' };
    q.setConditionalFormatRules(Object.keys(colors).map(v =>
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(v).setBackground(colors[v]).setRanges([statusRange]).build()));
  }

  // Referrals
  let r = ss.getSheetByName(TABS.REFERRALS);
  if (!r) {
    r = ss.insertSheet(TABS.REFERRALS);
    r.getRange(1, 1, 1, REFERRAL_HEADERS.length).setValues([REFERRAL_HEADERS]).setFontWeight('bold');
    r.getRange(2, 6, 500, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(REFERRAL_STATUS, true).build());
    r.setFrozenRows(1);
  }

  if (!ss.getSheetByName(TABS.LOG)) {
    SpreadsheetApp.getUi().alert('Heads up: no "' + TABS.LOG + '" tab found. Sent outreach will not be logged there until it exists.');
  }
  toast_('Tabs ready. Next: Setup > Set API keys, then fill in Agent Settings.');
}

function setApiKeys() {
  const ui = SpreadsheetApp.getUi();
  const ask = (label, key) => {
    const has = getSecret_(key) ? ' (already set - leave blank to keep)' : '';
    const res = ui.prompt(label + has, ui.ButtonSet.OK_CANCEL);
    if (res.getSelectedButton() !== ui.Button.OK) return false;
    const v = res.getResponseText().trim();
    if (v) setSecret_(key, v);
    return true;
  };
  if (!ask('Apollo.io API key (Apollo > Settings > Integrations > API)', 'APOLLO_API_KEY')) return;
  ask('Anthropic API key (console.anthropic.com). Optional - without it, drafts use a simple template.', 'ANTHROPIC_API_KEY');
  toast_('Keys saved to your private script storage (not the sheet).');
}

function enableDailyTrigger() {
  disableDailyTrigger();
  ScriptApp.newTrigger('syncInbox').timeBased().everyDays(1).atHour(8).create();
  toast_('Daily inbox check is on (runs around 8am).');
}

function disableDailyTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncInbox')
    .forEach(t => ScriptApp.deleteTrigger(t));
}
