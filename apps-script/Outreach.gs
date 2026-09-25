/**
 * Step 3 (send approved emails), step 4 (LinkedIn queue), and the inbox check
 * that marks things Sent / Replied, drafts follow-ups, and logs everything
 * into your "Networking 2027" tab.
 */

// ---------- Step 3: send approved emails ----------

function sendApproved() {
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings_();
  const queue = openTable_(TABS.QUEUE);
  const approved = queue.rows.filter(r => r[Q.STATUS] === STATUS.APPROVED);
  const toEmail = approved.filter(r => channelHasEmail_(r[Q.CHANNEL]) && r[Q.EMAIL] && !r[Q.EMAIL_SENT]);
  const linkedInWaiting = approved.filter(r => channelHasLinkedIn_(r[Q.CHANNEL]) && !r[Q.LI_SENT]).length;

  if (!toEmail.length) {
    ui.alert('No approved emails to send.' + (linkedInWaiting ? ' ' + linkedInWaiting + ' LinkedIn message(s) are waiting in step 4.' : ''));
    return;
  }
  const names = toEmail.map(r => '- ' + r[Q.NAME] + ' (' + r[Q.COMPANY] + ')').join('\n');
  if (ui.alert('Send ' + toEmail.length + ' email(s)?\n\n' + names, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

  let sent = 0;
  const problems = [];
  toEmail.forEach(row => {
    try {
      sendRowEmail_(settings, queue, row);
      sent++;
    } catch (e) {
      problems.push(row[Q.NAME] + ': ' + e.message);
    }
  });

  ui.alert('Sent ' + sent + ' email(s) and logged them in "' + TABS.LOG + '".' +
    (linkedInWaiting ? '\n\n' + linkedInWaiting + ' LinkedIn message(s) waiting - open step 4.' : '') +
    (problems.length ? '\n\nProblems:\n- ' + problems.join('\n- ') : ''));
}

function sendRowEmail_(settings, queue, row) {
  let draft = null;
  if (row[Q.DRAFT_ID]) {
    try { draft = GmailApp.getDraft(String(row[Q.DRAFT_ID])); } catch (e) { draft = null; }
  }
  if (!draft) {
    // Draft is gone: maybe you already sent it from Gmail.
    const already = findSentTo_(row[Q.EMAIL], row[Q.FOUND]);
    if (already) { markEmailSent_(queue, row, already.getDate()); return; }
    draft = GmailApp.getDraft(createGmailDraft_(settings, row));
  }
  draft.send();
  markEmailSent_(queue, row, new Date());
}

function markEmailSent_(queue, row, when) {
  queue.set(row, Q.EMAIL_SENT, when);
  if (row[Q.STATUS] !== STATUS.REPLIED) queue.set(row, Q.STATUS, STATUS.SENT);
  logOutreach_(row, 'Email');
}

// ---------- Step 4: LinkedIn queue (sidebar) ----------

function openLinkedInQueue() {
  const html = HtmlService.createHtmlOutputFromFile('LinkedInQueue').setTitle('LinkedIn queue');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Called by the sidebar. Everyone with a written LinkedIn note whose LinkedIn
 * step isn't done - no separate approval needed: clicking "I sent it" is the approval.
 */
function getLinkedInQueue() {
  const queue = openTable_(TABS.QUEUE);
  const ready = [STATUS.DRAFT_READY, STATUS.APPROVED, STATUS.SENT];
  return queue.rows
    .filter(r => ready.indexOf(r[Q.STATUS]) !== -1 && r[Q.NOTE] &&
                 channelHasLinkedIn_(r[Q.CHANNEL]) && !r[Q.LI_SENT] && r[Q.NAME])
    .map(r => ({
      row: r._row,
      name: String(r[Q.NAME]),
      title: String(r[Q.TITLE] || ''),
      company: String(r[Q.COMPANY] || ''),
      about: String(r[Q.ABOUT] || ''),
      why: String(r[Q.WHY] || ''),
      appliedRole: String(r[Q.APPLIED_ROLE] || ''),
      url: String(r[Q.LINKEDIN] || linkedInSearchUrl_(r[Q.NAME], r[Q.COMPANY])),
      note: String(r[Q.NOTE] || ''),
      message: String(r[Q.LI_DM] || ''),
    }));
}

/** Called by the sidebar when you click "I sent it" or "Skip". */
function markLinkedIn(rowNum, action, note) {
  const queue = openTable_(TABS.QUEUE);
  const row = queue.rows.find(r => r._row === rowNum);
  if (!row) throw new Error('Row ' + rowNum + ' not found');
  if (note && note !== row[Q.NOTE]) queue.set(row, Q.NOTE, note);
  if (action === 'skip') {
    queue.set(row, Q.LI_SENT, 'Skipped');
    return;
  }
  queue.set(row, Q.LI_SENT, new Date());
  const emailDone = !channelHasEmail_(row[Q.CHANNEL]) || !row[Q.EMAIL] || row[Q.EMAIL_SENT];
  if (emailDone && row[Q.STATUS] !== STATUS.REPLIED) queue.set(row, Q.STATUS, STATUS.SENT);
  logOutreach_(row, 'LinkedIn');
}

// ---------- Inbox check (manual or daily trigger) ----------

function syncInbox() {
  const settings = getSettings_();
  const queue = openTable_(TABS.QUEUE);
  const followUpDays = settingNumber_(settings, 'Follow-up after (days)', 7);
  const liveDrafts = {};
  GmailApp.getDrafts().forEach(d => { liveDrafts[d.getId()] = true; });
  const now = new Date();
  const counts = { sent: 0, replied: 0, followups: 0 };

  queue.rows.forEach(row => {
    if (!row[Q.NAME] || row[Q.STATUS] === STATUS.SKIP) return;
    const email = String(row[Q.EMAIL] || '').trim();

    // 1. You sent the draft straight from Gmail.
    if (email && row[Q.DRAFT_ID] && !row[Q.EMAIL_SENT] && !liveDrafts[row[Q.DRAFT_ID]]) {
      const msg = findSentTo_(email, row[Q.FOUND]);
      if (msg) { markEmailSent_(queue, row, msg.getDate()); counts.sent++; }
      else if (String(row[Q.NOTES] || '').indexOf('Gmail draft deleted') === -1) {
        queue.set(row, Q.NOTES, ((row[Q.NOTES] || '') + ' Gmail draft deleted.').trim());
      }
    }

    // 2. They replied (email). LinkedIn replies: set Status to "Replied" yourself.
    if (!row[Q.REPLIED]) {
      if (row[Q.STATUS] === STATUS.REPLIED) {
        queue.set(row, Q.REPLIED, now);
        markLogResponse_(row);
        counts.replied++;
      } else if (email && row[Q.EMAIL_SENT] instanceof Date) {
        const reply = GmailApp.search('from:' + email + ' after:' + gmailDate_(row[Q.EMAIL_SENT]), 0, 1);
        if (reply.length) {
          queue.set(row, Q.REPLIED, reply[0].getLastMessageDate());
          queue.set(row, Q.STATUS, STATUS.REPLIED);
          markLogResponse_(row);
          counts.replied++;
        }
      }
    }

    // 3. No reply after N days -> follow-up draft in the same thread (you send it).
    if (row[Q.STATUS] === STATUS.SENT && !row[Q.REPLIED] && !row[Q.FOLLOWUP] && email &&
        row[Q.EMAIL_SENT] instanceof Date && daysBetween_(row[Q.EMAIL_SENT], now) >= followUpDays) {
      if (createFollowUpDraft_(settings, row)) {
        queue.set(row, Q.FOLLOWUP, 'Drafted ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'M/d'));
        counts.followups++;
      }
    }
  });

  const msg = 'Inbox check: ' + counts.sent + ' newly sent, ' + counts.replied + ' replies, ' + counts.followups + ' follow-up draft(s) created.';
  try { toast_(msg); } catch (e) { /* running from the daily trigger - no UI */ }
  return msg;
}

function findSentTo_(email, since) {
  let q = 'in:sent to:' + String(email).trim();
  if (since instanceof Date) q += ' after:' + gmailDate_(new Date(since.getTime() - 86400000));
  const threads = GmailApp.search(q, 0, 5);
  for (const t of threads) {
    const msgs = t.getMessages();
    for (let i = msgs.length - 1; i >= 0; i--) {
      if (msgs[i].getTo().toLowerCase().indexOf(String(email).toLowerCase()) !== -1) return msgs[i];
    }
  }
  return null;
}

function createFollowUpDraft_(settings, row) {
  const email = String(row[Q.EMAIL]).trim();
  const original = findSentTo_(email, row[Q.FOUND]);
  const body = followUpText_(settings, row);
  if (original) {
    // Reply-all on your own message addresses the original recipient, keeping the thread.
    const draft = original.createDraftReplyAll(body);
    if (draft.getMessage().getTo().toLowerCase().indexOf(email.toLowerCase()) !== -1) return true;
    draft.deleteDraft();
  }
  GmailApp.createDraft(email, 'Re: ' + (row[Q.SUBJECT] || 'Following up'), body, { name: settings['Your name'] });
  return true;
}

function followUpText_(settings, row) {
  const first = firstName_(row[Q.NAME]);
  const formal = ['Consulting', 'Finance'].indexOf(row[Q.TRACK]) !== -1;
  const text = formal
    ? 'Hi ' + first + ',\n\nI wanted to follow up on my note below in case it got buried. I know your schedule is busy, and I would still greatly appreciate ' + settings['The ask'] + ' whenever is convenient.'
    : 'Hi ' + first + ',\n\nJust bumping this in case it got lost in your inbox. I\'d still love ' + settings['The ask'] + ' if you have time in the next few weeks.';
  return emailWithSignature_(settings, text);
}

// ---------- Logging into your existing "Networking 2027" tab ----------

/** Adds (or updates) the person in Networking 2027 so your hand-kept log stays complete. */
function logOutreach_(row, channel) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TABS.LOG);
  if (!sheet) return;
  const log = openTable_(TABS.LOG);
  const existing = findLogRow_(log, row);
  if (existing) {
    const form = String(existing['FORM OF CONTACT'] || '');
    if (form.indexOf(channel) === -1) log.set(existing, 'FORM OF CONTACT', form ? form + ' + ' + channel : channel);
    if (row[Q.EMAIL] && !existing['Column 1'] && log.headers.indexOf('Column 1') !== -1) log.set(existing, 'Column 1', row[Q.EMAIL]);
    return;
  }
  const entry = {
    'NAME': row[Q.NAME],
    'COMPANY': row[Q.COMPANY],
    'FORM OF CONTACT': channel,
    'Column 1': row[Q.EMAIL] || '',
    'DATE': new Date(),
    'LOCATION': row[Q.LOCATION] || '',
    'REACHED OUT?': true,
    'RESPONSE?': false,
    'NOTES': [row[Q.TITLE], 'via agent'].filter(Boolean).join(' - '),
  };
  const r = log.append(entry, 'NAME');
  // Copy checkbox/dropdown formatting from the first data row.
  if (r._row > 2) {
    const width = log.headers.length;
    sheet.getRange(2, 1, 1, width).copyTo(sheet.getRange(r._row, 1, 1, width),
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }
}

function markLogResponse_(row) {
  if (!SpreadsheetApp.getActive().getSheetByName(TABS.LOG)) return;
  const log = openTable_(TABS.LOG);
  const existing = findLogRow_(log, row);
  if (existing && log.headers.indexOf('RESPONSE?') !== -1) log.set(existing, 'RESPONSE?', true);
}

function findLogRow_(log, row) {
  const key = normalize_(row[Q.NAME]) + '|' + normalize_(row[Q.COMPANY]);
  return log.rows.find(r => normalize_(r['NAME']) + '|' + normalize_(r['COMPANY']) === key);
}

// ---------- Referrals ----------

function addSelectedToReferrals() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.QUEUE) { SpreadsheetApp.getUi().alert('Select rows in the "' + TABS.QUEUE + '" tab first.'); return; }
  const selected = selectedRowNumbers_();
  const queue = openTable_(TABS.QUEUE);
  const refs = openTable_(TABS.REFERRALS);
  let n = 0;
  queue.rows.filter(r => selected[r._row] && r[Q.NAME]).forEach(r => {
    refs.append({
      'Name': r[Q.NAME],
      'Company': r[Q.COMPANY],
      'Status': 'To Ask',
      'Notes': [r[Q.EMAIL], r[Q.LINKEDIN]].filter(Boolean).join(' | '),
    }, 'Name');
    n++;
  });
  toast_('Added ' + n + ' contact(s) to Referrals. Fill in the role and job link there.');
}
