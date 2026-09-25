/**
 * Step 2: write a personalized email + LinkedIn note for every "New" contact,
 * and put the email in your Gmail Drafts folder for review.
 */

const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';

function writeDrafts() {
  const started = Date.now();
  const settings = getSettings_();
  const queue = openTable_(TABS.QUEUE);
  const useClaude = !!getSecret_('ANTHROPIC_API_KEY');
  let written = 0, remaining = 0;
  const problems = [];

  for (const row of queue.rows) {
    if (row[Q.STATUS] !== STATUS.NEW || !row[Q.NAME]) continue;
    if (Date.now() - started > 4.5 * 60 * 1000) { remaining++; continue; }
    try {
      const draft = useClaude ? claudeDraft_(settings, row) : templateDraft_(settings, row);
      queue.set(row, Q.SUBJECT, draft.subject);
      queue.set(row, Q.BODY, draft.email_body);
      queue.set(row, Q.NOTE, fitNote_(draft.linkedin_note, settingNumber_(settings, 'LinkedIn note max characters', 200)));
      queue.set(row, Q.LI_DM, draft.linkedin_message || '');
      if (channelHasEmail_(row[Q.CHANNEL]) && row[Q.EMAIL]) {
        queue.set(row, Q.DRAFT_ID, createGmailDraft_(settings, row));
      }
      queue.set(row, Q.STATUS, STATUS.DRAFT_READY);
      written++;
    } catch (e) {
      problems.push(row[Q.NAME] + ': ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Wrote ' + written + ' draft(s)' + (useClaude ? '' : ' using the basic template (add an Anthropic key for personalized drafts)') + '.\n\n' +
    'LinkedIn: the queue opens next - open each profile, copy, paste, send.\n' +
    'Email: drafts are in Gmail. Send from there, or set Status to "Approved" and run step 3.' +
    (remaining ? '\n\n' + remaining + ' contact(s) left - run step 2 again.' : '') +
    (problems.length ? '\n\nProblems:\n- ' + problems.join('\n- ') : ''));
  if (written) openLinkedInQueue();
}

function claudeDraft_(settings, row) {
  const model = settings['Claude model'] || 'claude-opus-5';
  const noteLimit = settingNumber_(settings, 'LinkedIn note max characters', 200);
  const formal = ['Consulting', 'Finance'].indexOf(row[Q.TRACK]) !== -1;

  const system = [
    'You write cold networking outreach for a college student who is recruiting. Every message goes to a real professional and is reviewed by the student before sending.',
    '',
    'Tone: ' + (formal
      ? 'polished and concise, the way a strong consulting candidate writes. Respectful of their time, clear ask, no slang, no exclamation points.'
      : 'warm and genuine with a bit of personality, still professional. Sound like a curious student, not a salesperson.'),
    '',
    'Rules:',
    '- Email body: 70-120 words, 2-3 short paragraphs. Start with "Hi <first name>,". Do not include a sign-off or signature; one is added automatically.',
    '- Say who the student is in one line, give one specific reason for reaching out to THIS person, then make the ask.',
    '- The best reason is something they share with the student (listed in "shared"), e.g. Penn, the same high school or hometown, the same club. Lead with it when present. These come from a keyword match and the student checks the profile before sending, so state them plainly but do not embellish.',
    '- Mention that the student applied to the role in "applied_role" at their company (if given) as context for why they are reaching out. Do not ask for a referral in the first message.',
    '- Never claim a connection, mutual friend, or fact that is not in the input.',
    '- Use "background" (their real career history) to pick the specific reason, e.g. a past employer, a switch into consulting, or time at the firm. Do not invent anything beyond the input.',
    '- Avoid filler: no "I hope this email finds you well", "I am reaching out because", "I would love to pick your brain", "passionate", or flattery. No em dashes.',
    '- Subject: under 8 words, specific, no clickbait.',
    '- LinkedIn note (connection request): at most ' + noteLimit + ' characters including spaces, first name greeting, one-line reason, the ask. No links.',
    '- LinkedIn message (sent after they accept): 300-500 characters, thanks them for connecting, one specific line about their background, then the ask with flexibility on timing.',
  ].join('\n');

  const input = {
    student: {
      name: settings['Your name'],
      school_and_year: settings['School & year'],
      major: settings['Major / focus'],
      about: settings['About me'],
      ask: settings['The ask'],
    },
    recipient: {
      first_name: firstName_(row[Q.NAME]),
      full_name: row[Q.NAME],
      title: row[Q.TITLE],
      company: row[Q.COMPANY],
      location: row[Q.LOCATION],
      background: row[Q.ABOUT] || '',
      shared: sharedFromWhy_(row[Q.WHY]),
      applied_role: row[Q.APPLIED_ROLE] || '',
      industry_track: row[Q.TRACK],
      extra_notes: row[Q.NOTES] || '',
    },
  };

  const body = {
    model: model,
    max_tokens: 8000,
    system: system,
    messages: [{ role: 'user', content: 'Write the outreach for this contact:\n' + JSON.stringify(input, null, 2) }],
    output_config: {
      effort: 'medium',
      format: {
        type: 'json_schema',
        schema: {
          type: 'object',
          properties: {
            subject: { type: 'string' },
            email_body: { type: 'string' },
            linkedin_note: { type: 'string' },
            linkedin_message: { type: 'string' },
          },
          required: ['subject', 'email_body', 'linkedin_note', 'linkedin_message'],
          additionalProperties: false,
        },
      },
    },
  };
  const headers = { 'x-api-key': getSecret_('ANTHROPIC_API_KEY'), 'anthropic-version': '2023-06-01' };
  // If the model declines a request, let the API retry on a fallback model automatically.
  if (/^claude-(opus-5|fable)/.test(model)) {
    body.fallbacks = 'default';
    headers['anthropic-beta'] = 'server-side-fallback-2026-07-01';
  }

  const resp = UrlFetchApp.fetch(ANTHROPIC_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: headers,
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  if (code !== 200) throw new Error('Claude API returned ' + code + ': ' + resp.getContentText().slice(0, 300));
  return parseClaudeDraft_(JSON.parse(resp.getContentText()));
}

function parseClaudeDraft_(message) {
  if (message.stop_reason === 'refusal') throw new Error('Claude declined to write this one - write it by hand.');
  if (message.stop_reason === 'max_tokens') throw new Error('Claude ran out of room - run step 2 again.');
  const text = (message.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  const draft = JSON.parse(text);
  if (!draft.subject || !draft.email_body) throw new Error('Claude returned an incomplete draft.');
  return draft;
}

/** Pulls the shared keywords back out of the "Why Them" text. */
function sharedFromWhy_(why) {
  const m = String(why || '').match(/Shared: ([^(]+)/);
  return m ? m[1].split(',').map(x => x.trim()).filter(Boolean) : [];
}

/** Fallback when there's no Anthropic key: simple, editable template. */
function templateDraft_(settings, row) {
  const first = firstName_(row[Q.NAME]);
  const me = settings['Your name'];
  const myFirst = me.split(' ')[0];
  const ask = settings['The ask'];
  const shared = sharedFromWhy_(row[Q.WHY]);
  const role = row[Q.APPLIED_ROLE];
  const tie = shared.length ? 'I saw we both have ' + shared[0] + ' in common' : 'I came across your profile';
  const why = tie + (role ? ' and I recently applied to the ' + role + ' role at ' + row[Q.COMPANY] : ' while looking into ' + row[Q.COMPANY]) + '.';
  return {
    subject: (shared.length ? shared[0] + ' - ' : '') + 'student interested in ' + row[Q.COMPANY],
    email_body: 'Hi ' + first + ',\n\nMy name is ' + me + ', a student at ' + settings['School & year'] + '. ' + why +
      '\n\nWould you be open to ' + ask + '? I am happy to work around your schedule.',
    linkedin_note: 'Hi ' + first + ', I\'m ' + myFirst + ', a Penn student. ' + why + ' Would love to connect!',
    linkedin_message: 'Hi ' + first + ', thanks for connecting! ' + why + ' Would you be open to ' + ask +
      '? Happy to work around your schedule.',
  };
}

/** Trims a LinkedIn note to the character limit at a word boundary. */
function fitNote_(note, limit) {
  note = String(note || '').trim();
  if (note.length <= limit) return note;
  const cut = note.slice(0, limit - 1);
  const lastSpace = cut.lastIndexOf(' ');
  return (lastSpace > limit * 0.6 ? cut.slice(0, lastSpace) : cut).replace(/[,;:\s]+$/, '') + '…';
}

function emailWithSignature_(settings, body) {
  const sig = String(settings['Email signature'] || '').replace(/\\n/g, '\n');
  return String(body).trim() + (sig ? '\n\n' + sig : '');
}

function createGmailDraft_(settings, row) {
  const draft = GmailApp.createDraft(
    String(row[Q.EMAIL]).trim(),
    row[Q.SUBJECT],
    emailWithSignature_(settings, row[Q.BODY]),
    { name: settings['Your name'] });
  return draft.getId();
}

/** After editing Subject/Body in the sheet, push those edits to Gmail. */
function rebuildSelectedDrafts() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() !== TABS.QUEUE) { SpreadsheetApp.getUi().alert('Select rows in the "' + TABS.QUEUE + '" tab first.'); return; }
  const settings = getSettings_();
  const queue = openTable_(TABS.QUEUE);
  const selected = selectedRowNumbers_();
  let n = 0;
  queue.rows.filter(r => selected[r._row]).forEach(row => {
    if (!row[Q.EMAIL] || row[Q.EMAIL_SENT]) return;
    deleteDraftQuietly_(row[Q.DRAFT_ID]);
    queue.set(row, Q.DRAFT_ID, createGmailDraft_(settings, row));
    n++;
  });
  toast_('Rebuilt ' + n + ' Gmail draft(s) from the sheet text.');
}

function deleteDraftQuietly_(draftId) {
  if (!draftId) return;
  try { GmailApp.getDraft(String(draftId)).deleteDraft(); } catch (e) { /* already sent or deleted */ }
}

function selectedRowNumbers_() {
  const rows = {};
  const list = SpreadsheetApp.getActiveSheet().getActiveRangeList();
  if (!list) return rows;
  list.getRanges().forEach(r => {
    for (let i = r.getRow(); i < r.getRow() + r.getNumRows(); i++) if (i > 1) rows[i] = true;
  });
  return rows;
}
