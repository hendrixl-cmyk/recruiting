/**
 * Shared constants and small helpers used by every other file.
 */

const TABS = {
  TARGETS: 'Targets',
  QUEUE: 'Outreach Queue',
  REFERRALS: 'Referrals',
  SETTINGS: 'Agent Settings',
  LOG: 'Networking 2027', // your existing tab - the agent only appends/updates rows here
  // Your applications tab is set in Agent Settings ("Applications tab").
};

// Column names in the Outreach Queue. The code looks columns up by name,
// so you can reorder or add columns in the sheet without breaking anything.
const Q = {
  STATUS: 'Status',
  CHANNEL: 'Channel',
  NAME: 'Name',
  TITLE: 'Title',
  COMPANY: 'Company',
  APPLIED_ROLE: 'Applied Role',
  WHY: 'Why Them',
  LOCATION: 'Location',
  ABOUT: 'About Them',
  LINKEDIN: 'LinkedIn URL',
  EMAIL: 'Email',
  EMAIL_STATUS: 'Email Status',
  TRACK: 'Track',
  SUBJECT: 'Email Subject',
  BODY: 'Email Body',
  NOTE: 'LinkedIn Note',
  LI_DM: 'LinkedIn Message (after they accept)',
  DRAFT_ID: 'Gmail Draft ID',
  FOUND: 'Date Found',
  EMAIL_SENT: 'Email Sent',
  LI_SENT: 'LinkedIn Sent',
  FOLLOWUP: 'Follow-up Sent/Drafted',
  REPLIED: 'Replied',
  NOTES: 'Notes',
  APOLLO_ID: 'Apollo ID',
};

const QUEUE_HEADERS = [
  Q.STATUS, Q.CHANNEL, Q.NAME, Q.TITLE, Q.COMPANY, Q.APPLIED_ROLE, Q.WHY, Q.LINKEDIN, Q.NOTE, Q.LI_DM, Q.ABOUT, Q.LOCATION, Q.EMAIL,
  Q.EMAIL_STATUS, Q.TRACK, Q.SUBJECT, Q.BODY, Q.DRAFT_ID, Q.FOUND,
  Q.EMAIL_SENT, Q.LI_SENT, Q.FOLLOWUP, Q.REPLIED, Q.NOTES, Q.APOLLO_ID,
];

// Contact lifecycle:
//   New -> Draft Ready -> (you) Approved -> Sent -> Replied
//                      \-> (you) Skip
const STATUS = {
  NEW: 'New',
  DRAFT_READY: 'Draft Ready',
  APPROVED: 'Approved',
  SENT: 'Sent',
  REPLIED: 'Replied',
  SKIP: 'Skip',
};
const STATUS_OPTIONS = [STATUS.NEW, STATUS.DRAFT_READY, STATUS.APPROVED, STATUS.SENT, STATUS.REPLIED, STATUS.SKIP];
const CHANNEL_OPTIONS = ['Email', 'LinkedIn', 'Both'];
const TRACK_OPTIONS = ['Consulting', 'Music', 'Consumer', 'Finance', 'Tech', 'Other'];

const TARGET_HEADERS = ['Active', 'Company', 'Domain', 'Apollo Org ID', 'Track', 'Applied Roles', 'Titles', 'Locations', 'Keywords', '# Contacts', 'Source', 'Last Run'];
const REFERRAL_HEADERS = ['Name', 'Company', 'Role / Job', 'Job Link', 'Date Asked', 'Status', 'Notes'];
const REFERRAL_STATUS = ['To Ask', 'Asked', 'Agreed', 'Submitted', 'Declined'];

const SETTINGS_DEFAULTS = [
  ['Your name', 'Hendrix Lee', 'Used in signatures.'],
  ['School & year', 'University of Pennsylvania, Class of 2028', ''],
  ['Major / focus', 'Philosophy, Politics & Economics (Choice & Behavior), minor in Cinema & Media Studies', ''],
  ['About me', 'From Miami Beach. Business development intern at Pivot Tokyo / AI Beyond Borders (sponsorship decks, investor outreach in Tokyo). Associate Consultant with Consult for America and on the consulting committee of Wharton Undergraduate Media & Entertainment. Member of Delta Sigma Pi (professional business fraternity). Dances with Penn Hype and Strictly Funk, and is Artistic Director of Pan Asian Dance Troupe. Into the music industry and GarageBand production.', 'Facts the AI can pull from. Keep it true and specific.'],
  ['Similarity keywords', 'University of Pennsylvania, Delta Sigma Pi, Miami Beach, Consult for America, Wharton Undergraduate Media, Penn Hype, Strictly Funk, Pan Asian Dance Troupe', 'Things you share with people. Each one is a separate (free) Apollo search per company; people who match more of them rank higher. Put the most common ties first.'],
  ['The ask', 'a 15-minute call to hear about their path and experience', 'What you are asking for in each message.'],
  ['Email signature', 'Best,\nHendrix Lee\nUniversity of Pennsylvania', 'Appended to every email.'],
  ['Applications tab', '2027 Recruiting Tracker', 'Tab with your applications (needs COMPANY, ROLE / POSITION, LOCATION, APPLICATION STATUS columns).'],
  ['Skip application statuses', 'Rejected', 'Comma-separated. Companies where every application has one of these statuses are skipped.'],
  ['Contacts per company', '3', 'How many people to find per company you applied to.'],
  ['Also draft emails', 'Yes', 'Yes = also write an email (Gmail draft) when Apollo finds a verified address. No = LinkedIn only.'],
  ['Follow-up after (days)', '7', 'Days with no email reply before a follow-up draft is created.'],
  ['Max enrichments per run', '15', 'Each enrichment costs 1 Apollo credit (that is how you get the full name, LinkedIn and email).'],
  ['LinkedIn note max characters', '200', 'Free LinkedIn accounts: 200. Premium: 300.'],
  ['Claude model', 'claude-opus-5', 'Model used to write drafts.'],
  ['Default locations', 'New York', 'Used when an application has no location. Comma-separated.'],
];

// ---------- Secrets (stored per-user, never in the sheet) ----------

function getSecret_(key) {
  return PropertiesService.getUserProperties().getProperty(key) || '';
}

function setSecret_(key, value) {
  PropertiesService.getUserProperties().setProperty(key, value);
}

// ---------- Settings ----------

function getSettings_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TABS.SETTINGS);
  const settings = {};
  SETTINGS_DEFAULTS.forEach(([k, v]) => { settings[k] = v; });
  if (!sheet) return settings;
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0]).trim();
    if (key) settings[key] = String(values[i][1]);
  }
  return settings;
}

function settingNumber_(settings, key, fallback) {
  const n = parseInt(settings[key], 10);
  return isNaN(n) ? fallback : n;
}

// ---------- Table helper: read/write rows by header name ----------

/**
 * Wraps a sheet whose first row is headers. Rows are plain objects keyed by
 * header name, plus `_row` (1-based sheet row number).
 */
function openTable_(sheetName) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sheet) throw new Error('Missing tab "' + sheetName + '". Run Recruiting Agent > Set up tabs first.');
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const obj = { _row: i + 1 };
    headers.forEach((h, j) => { if (h) obj[h] = values[i][j]; });
    rows.push(obj);
  }
  return {
    sheet: sheet,
    headers: headers,
    rows: rows,
    col(name) {
      const idx = headers.indexOf(name);
      if (idx === -1) throw new Error('Column "' + name + '" not found in ' + sheetName);
      return idx + 1;
    },
    set(row, name, value) {
      sheet.getRange(row._row, this.col(name)).setValue(value);
      row[name] = value;
    },
    /**
     * Writes `obj` into the first row whose `keyName` cell is blank (sheets
     * with checkboxes pre-filled far down have "empty" rows that
     * appendRow would skip past), or after the last row.
     */
    append(obj, keyName) {
      const key = keyName || headers.find(h => h);
      const line = headers.map(h => (h in obj ? obj[h] : ''));
      let target = rows.find(r => r._blankClaimed !== true && String(r[key] || '').trim() === '');
      let rowNum;
      if (target) {
        rowNum = target._row;
        target._blankClaimed = true;
      } else {
        rowNum = (rows.length ? rows[rows.length - 1]._row : 1) + 1;
      }
      sheet.getRange(rowNum, 1, 1, line.length).setValues([line]);
      const r = Object.assign({ _row: rowNum }, obj);
      if (target) Object.assign(target, obj); else rows.push(r);
      return target || r;
    },
  };
}

// ---------- Pure helpers (unit-tested in tests/) ----------

function normalize_(s) {
  return String(s || '').toLowerCase().replace(/[^a-z0-9@.]/g, '');
}

function splitList_(s) {
  return String(s || '').split(/[,;\n]/).map(x => x.trim()).filter(Boolean);
}

function channelHasEmail_(channel) {
  return channel === 'Email' || channel === 'Both';
}

function channelHasLinkedIn_(channel) {
  return channel === 'LinkedIn' || channel === 'Both';
}

function pickChannel_(email, emailStatus, linkedinUrl, draftEmails) {
  const goodEmail = draftEmails !== false && email && emailStatus !== 'unavailable';
  if (goodEmail && linkedinUrl) return 'Both';
  if (goodEmail) return 'Email';
  return 'LinkedIn';
}

function firstName_(fullName) {
  return String(fullName || '').trim().split(/\s+/)[0] || '';
}

function daysBetween_(a, b) {
  return Math.floor((b.getTime() - a.getTime()) / 86400000);
}

function gmailDate_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy/MM/dd');
}

function toast_(msg) {
  SpreadsheetApp.getActive().toast(msg, 'Recruiting Agent', 8);
}
