// Runs the Apps Script code in Node against fake Sheets / Gmail / Apollo / Claude.
// Usage: node tests/run.js
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

// ---------- Fakes ----------
function makeSheet(name, grid) {
  const data = grid.map(r => r.slice());
  const sheet = {
    name, data,
    getName: () => name,
    getDataRange: () => ({ getValues: () => data.map(r => r.slice()) }),
    getLastRow: () => data.length,
    getRange(row, col, nr = 1, nc = 1) {
      return {
        setValue(v) { ensure(row, col); data[row - 1][col - 1] = v; },
        setValues(vals) { vals.forEach((line, i) => line.forEach((v, j) => { ensure(row + i, col + j); data[row + i - 1][col + j - 1] = v; })); },
        copyTo() {},
      };
    },
  };
  function ensure(r, c) {
    while (data.length < r) data.push([]);
    while (data[r - 1].length < c) data[r - 1].push('');
  }
  return sheet;
}

function makeContext(sheets, opts = {}) {
  const alerts = [];
  const fetches = [];
  const drafts = {};
  let draftSeq = 0;
  const ss = {
    getSheetByName: n => sheets[n] || null,
    toast: () => {},
  };
  const ctx = {
    console,
    SpreadsheetApp: {
      getActive: () => ss,
      getUi: () => ({ alert: m => { alerts.push(m); return 'YES'; }, showSidebar: () => {}, ButtonSet: {}, Button: { YES: 'YES' } }),
      CopyPasteType: { PASTE_DATA_VALIDATION: 1 },
    },
    PropertiesService: { getUserProperties: () => ({ getProperty: k => (opts.secrets || {})[k] || null }) },
    UrlFetchApp: {
      fetch(url, o) {
        const body = JSON.parse(o.payload);
        fetches.push({ url, body, headers: o.headers });
        const res = opts.fetch(url, body);
        return { getResponseCode: () => res.code || 200, getContentText: () => JSON.stringify(res.json) };
      },
    },
    GmailApp: {
      createDraft(to, subject, body) {
        const id = 'd' + (++draftSeq);
        drafts[id] = { to, subject, body };
        return { getId: () => id };
      },
      getDrafts: () => Object.keys(drafts).map(id => ({ getId: () => id })),
    },
    Session: { getScriptTimeZone: () => 'America/New_York' },
    HtmlService: { createHtmlOutputFromFile: () => ({ setTitle() { return this; } }) },
    Utilities: { formatDate: d => d.toISOString().slice(0, 10).replace(/-/g, '/') },
  };
  vm.createContext(ctx);
  const dir = path.join(__dirname, '..', 'apps-script');
  fs.readdirSync(dir).filter(f => f.endsWith('.gs')).forEach(f => {
    vm.runInContext(fs.readFileSync(path.join(dir, f), 'utf8'), ctx, { filename: f });
  });
  return { ctx, alerts, fetches, drafts };
}

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('ok   ' + name); }
  catch (e) { console.log('FAIL ' + name + '\n     ' + (e.stack || e)); process.exitCode = 1; }
}

// ---------- Fixtures ----------
function baseSheets() {
  const { ctx } = makeContext({}, { fetch: () => ({ json: {} }) });
  const QH = Array.from(vm.runInContext('QUEUE_HEADERS', ctx));
  const TH = Array.from(vm.runInContext('TARGET_HEADERS', ctx));
  return {
    QH,
    sheets: {
      'Agent Settings': settingsSheet(ctx, { 'Similarity keywords': 'University of Pennsylvania, Miami Beach', 'Contacts per company': '2' }),
      '2027 Recruiting Tracker': makeSheet('2027 Recruiting Tracker', [
        ['COMPANY', 'ROLE / POSITION', 'LOCATION', 'DATE APPLIED', 'APPLICATION STATUS'],
        ['Bain', 'Associate Consultant Intern', 'NYC, SF, LA', '3/29/2025', 'Rejected'],
        ['BCG', 'Summer Internship', 'Brooklyn', '6/2/2026', 'Rejected'],
        ['JP Morgan', 'Firmwide Strategy Analyst', '', '8/28/2026', 'Applied'],
        ['JP Morgan', '2027 Markets Summer Analyst Program', 'NYC', '6/15/2026', 'Applied'],
      ]),
      'Targets': makeSheet('Targets', [TH]),
      'Outreach Queue': makeSheet('Outreach Queue', [QH]),
      'Networking 2027': makeSheet('Networking 2027', [
        ['NAME', 'COMPANY', 'FORM OF CONTACT', 'Column 1', 'DATE', 'LOCATION', 'REACHED OUT?', 'RESPONSE?', 'NOTES'],
        ['Amanda Cui', 'JP Morgan', 'Email', '', '3/17/2026', 'New York', true, false, ''],
        ['', '', '', '', '', '', false, false, ''], // pre-filled checkbox row
      ]),
    },
  };
}

// Search results depend on the keyword: Jordan matches Penn AND Miami Beach, Priya only Penn.
function apolloFake(url, body) {
  if (url.endsWith('mixed_companies/search')) {
    return { json: { organizations: [{ id: 'org_jpm', name: 'JPMorganChase', primary_domain: 'jpmorganchase.com' }] } };
  }
  if (url.endsWith('mixed_people/api_search')) {
    const P = {
      amanda: { id: 'a1', first_name: 'Amanda', last_name_obfuscated: 'Cu***i', title: 'Analyst' }, // already contacted
      jordan: { id: 'a2', first_name: 'Jordan', last_name_obfuscated: 'Pa***k', title: 'Strategy Analyst' },
      priya: { id: 'a3', first_name: 'Priya', last_name_obfuscated: 'Ra***o', title: 'Markets Analyst' },
      extra: { id: 'a4', first_name: 'Extra', last_name_obfuscated: 'Pe***n', title: 'Analyst' },
    };
    const k = body.q_keywords || '';
    if (k === 'University of Pennsylvania') return { json: { people: [P.amanda, P.priya, P.jordan] } };
    if (k === 'Miami Beach') return { json: { people: [P.jordan] } };
    if (k) return { json: { people: [] } };
    return { json: { people: [P.extra, P.priya, P.jordan, P.amanda] } };
  }
  if (url.endsWith('people/bulk_match')) {
    const db = {
      a2: { id: 'a2', name: 'Jordan Park', first_name: 'Jordan', last_name: 'Park', title: 'Strategy Analyst', email: 'jordan.park@jpmorgan.com', email_status: 'verified',
            linkedin_url: 'http://www.linkedin.com/in/jordanpark', city: 'New York', state: 'New York', headline: 'Firmwide Strategy at JPMorganChase',
            employment_history: [{ current: true, title: 'Strategy Analyst', organization_name: 'JPMorganChase', start_date: '2024-08-01' },
                                 { current: false, title: 'Analyst', organization_name: 'Deloitte', start_date: '2021-06-01', end_date: '2024-07-01' }] },
      a3: { id: 'a3', name: 'Priya Rao', first_name: 'Priya', last_name: 'Rao', title: 'Markets Analyst', email: null, email_status: 'unavailable', linkedin_url: null },
      a4: { id: 'a4', name: 'Extra Person', title: 'Analyst', linkedin_url: 'http://www.linkedin.com/in/extra' },
    };
    return { json: { matches: body.details.map(d => db[d.id] || null) } };
  }
  throw new Error('unexpected url ' + url);
}

function settingsSheet(ctx, overrides) {
  const rows = [['Setting', 'Value', 'Help']].concat(Array.from(vm.runInContext('SETTINGS_DEFAULTS', ctx)).map(r => Array.from(r)));
  Object.keys(overrides || {}).forEach(k => { rows.find(r => r[0] === k)[1] = overrides[k]; });
  return makeSheet('Agent Settings', rows);
}

// ---------- Tests ----------
test('pure helpers', () => {
  const { ctx } = makeContext({}, { fetch: () => ({}) });
  assert.strictEqual(ctx.pickChannel_('a@b.com', 'verified', 'x'), 'Both');
  assert.strictEqual(ctx.pickChannel_('a@b.com', 'unavailable', 'x'), 'LinkedIn');
  assert.strictEqual(ctx.pickChannel_('', '', ''), 'LinkedIn');
  assert.deepStrictEqual(Array.from(ctx.splitList_('a, b;c\n')), ['a', 'b', 'c']);
  const long = 'Hi Jordan, I am a Penn student recruiting for consulting and would really value hearing about your move from Deloitte to BCG. Would you be open to a short call sometime in the next couple of weeks? Thanks so much!';
  const fit = ctx.fitNote_(long, 200);
  assert.ok(fit.length <= 200, 'note fits: ' + fit.length);
  assert.ok(fit.endsWith('…'));
  assert.strictEqual(ctx.fitNote_('short', 200), 'short');
});

test('Claude response parsing', () => {
  const { ctx } = makeContext({}, { fetch: () => ({}) });
  const d = ctx.parseClaudeDraft_({ stop_reason: 'end_turn', content: [{ type: 'thinking', thinking: '' },
    { type: 'text', text: '{"subject":"Hi","email_body":"Body","linkedin_note":"N","linkedin_message":"M"}' }] });
  assert.strictEqual(d.subject, 'Hi');
  assert.throws(() => ctx.parseClaudeDraft_({ stop_reason: 'refusal', content: [] }), /declined/);
});

test('applications -> targets: groups by company, skips rejected, infers titles', () => {
  const { ctx } = makeContext({}, { fetch: () => ({}) });
  const groups = ctx.groupApplications_([
    { 'COMPANY': 'Bain', 'ROLE / POSITION': 'ACI', 'APPLICATION STATUS': 'Rejected' },
    { 'COMPANY': 'American Express', 'ROLE / POSITION': 'Marketing, US Consumer Services', 'LOCATION': 'NYC', 'APPLICATION STATUS': 'Applied' },
    { 'COMPANY': 'American Express ', 'ROLE / POSITION': 'Pricing Strategy', 'LOCATION': 'NYC', 'APPLICATION STATUS': 'Applied' },
    { 'COMPANY': 'Capital One', 'ROLE / POSITION': 'Strategy Consulting Intern', 'APPLICATION STATUS': 'Closed - Appeal Pending' },
  ], ['Rejected']);
  assert.deepStrictEqual(Array.from(groups, g => g.company), ['American Express', 'Capital One']);
  assert.deepStrictEqual(Array.from(groups[0].roles), ['Marketing, US Consumer Services', 'Pricing Strategy']);
  assert.deepStrictEqual(Array.from(groups[0].locations), ['New York']);
  assert.strictEqual(ctx.inferTrack_('American Express', ['Marketing']), 'Finance');
  assert.strictEqual(ctx.inferTrack_('Bain', ['ACI']), 'Consulting');
  assert.strictEqual(ctx.inferTrack_('UMG', ['Brand']), 'Music');
  assert.ok(ctx.inferTitles_(['Strategy Consulting Intern']).includes('Associate Consultant'));
  assert.strictEqual(ctx.matchAppliedRole_('Markets Analyst', 'Firmwide Strategy Analyst; 2027 Markets Summer Analyst Program'), '2027 Markets Summer Analyst Program');
  assert.strictEqual(ctx.matchAppliedRole_('Strategy Associate', 'Firmwide Strategy Analyst; 2027 Markets Summer Analyst Program'), 'Firmwide Strategy Analyst');
});

test('findContacts: sources from applications, ranks by similarity, dedupes', () => {
  const { sheets } = baseSheets();
  const { ctx, alerts, fetches } = makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake });
  sheets['Agent Settings'] = settingsSheet(ctx, { 'Similarity keywords': 'University of Pennsylvania, Miami Beach', 'Contacts per company': '2' });
  ctx.findContacts();

  // Targets: only JP Morgan (Bain/BCG all rejected), resolved to an Apollo org.
  const t = sheets['Targets'].data, TH = t[0];
  assert.strictEqual(t.length, 2, JSON.stringify(t) + alerts);
  assert.strictEqual(t[1][TH.indexOf('Company')], 'JP Morgan');
  assert.strictEqual(t[1][TH.indexOf('Apollo Org ID')], 'org_jpm');
  assert.strictEqual(t[1][TH.indexOf('Track')], 'Finance');
  assert.strictEqual(t[1][TH.indexOf('Locations')], 'New York');
  assert.match(t[1][TH.indexOf('Titles')], /Strategy Analyst/);

  // People: Jordan (2 matches) ranked above Priya (1 match); Extra (0) and Amanda (already contacted) left out.
  const q = sheets['Outreach Queue'].data, H = q[0];
  const get = (r, h) => q[r][H.indexOf(h)];
  assert.strictEqual(q.length, 3, JSON.stringify(q.slice(1)) + alerts);
  assert.strictEqual(get(1, 'Name'), 'Jordan Park');
  assert.match(get(1, 'Why Them'), /Shared: University of Pennsylvania, Miami Beach/);
  assert.strictEqual(get(1, 'Applied Role'), 'Firmwide Strategy Analyst');
  assert.strictEqual(get(1, 'Channel'), 'Both');
  assert.match(get(1, 'About Them'), /Before: Analyst @ Deloitte \(2021-2024\)/);
  assert.strictEqual(get(2, 'Name'), 'Priya Rao');
  assert.strictEqual(get(2, 'Applied Role'), '2027 Markets Summer Analyst Program');
  assert.strictEqual(get(2, 'Channel'), 'LinkedIn');
  assert.match(get(2, 'LinkedIn URL'), /linkedin\.com\/search\/results\/people\/\?keywords=Priya%20Rao%20JP%20Morgan/);

  const searches = fetches.filter(f => f.url.endsWith('api_search'));
  assert.strictEqual(searches.length, 3, 'one per keyword + one plain');
  assert.deepStrictEqual(searches[0].body.organization_ids, ['org_jpm']);
  const enrich = fetches.filter(f => f.url.endsWith('bulk_match'));
  assert.deepStrictEqual(enrich[0].body.details.map(d => d.id), ['a2', 'a3'], 'only enrich the top 2 -> no wasted credits');

  // Running again doesn't duplicate targets or people.
  makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake }).ctx.findContacts();
  assert.strictEqual(sheets['Targets'].data.length, 2);
  assert.ok(sheets['Outreach Queue'].data.filter(r => r[H.indexOf('Name')] === 'Jordan Park').length === 1);
});

test('LinkedIn only mode skips emails', () => {
  const { sheets } = baseSheets();
  const { ctx } = makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake });
  sheets['Agent Settings'] = settingsSheet(ctx, { 'Also draft emails': 'No', 'Contacts per company': '1' });
  ctx.findContacts();
  const q = sheets['Outreach Queue'].data, H = q[0];
  assert.strictEqual(q[1][H.indexOf('Channel')], 'LinkedIn');
  assert.strictEqual(q[1][H.indexOf('Email')], '');
});

test('writeDrafts with Claude: fills columns + Gmail draft', () => {
  const { sheets } = baseSheets();
  const apollo = makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake });
  apollo.ctx.findContacts();
  const claude = (url, body) => {
    assert.ok(url.includes('api.anthropic.com'));
    assert.strictEqual(body.model, 'claude-opus-5');
    assert.strictEqual(body.fallbacks, 'default');
    assert.strictEqual(body.output_config.format.type, 'json_schema');
    const input = body.messages[0].content;
    return { json: { stop_reason: 'end_turn', content: [{ type: 'text', text: JSON.stringify({
      subject: 'Penn student - JPM question',
      email_body: 'Hi ' + (input.includes('Jordan') ? 'Jordan' : 'Priya') + ',\n\nBody.',
      linkedin_note: 'x'.repeat(250),
      linkedin_message: 'Thanks for connecting!',
    }) }] } };
  };
  const { ctx, drafts, fetches } = makeContext(sheets, { secrets: { ANTHROPIC_API_KEY: 'sk' }, fetch: claude });
  ctx.writeDrafts();
  const q = sheets['Outreach Queue'].data, H = q[0];
  const get = (r, h) => q[r][H.indexOf(h)];
  assert.strictEqual(get(1, 'Status'), 'Draft Ready');
  assert.strictEqual(get(1, 'Gmail Draft ID'), 'd1');
  assert.ok(get(1, 'LinkedIn Note').length <= 200);
  assert.strictEqual(get(1, 'LinkedIn Message (after they accept)'), 'Thanks for connecting!');
  assert.strictEqual(drafts.d1.to, 'jordan.park@jpmorgan.com');
  assert.match(drafts.d1.body, /Best,\nHendrix Lee/);
  // Priya has no email -> no Gmail draft, still Draft Ready for LinkedIn.
  assert.strictEqual(get(2, 'Gmail Draft ID'), '');
  assert.strictEqual(get(2, 'Status'), 'Draft Ready');
  const sent = fetches[0].body.messages[0].content;
  assert.ok(sent.includes('Deloitte'), 'background passed to Claude');
  assert.ok(sent.includes('"applied_role": "Firmwide Strategy Analyst"'), 'applied role passed');
  assert.ok(/"shared": \[\s*"University of Pennsylvania"/.test(sent), 'shared keywords passed');
});

test('writeDrafts without Claude key uses template', () => {
  const { sheets } = baseSheets();
  makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake }).ctx.findContacts();
  const { ctx } = makeContext(sheets, { secrets: {}, fetch: () => { throw new Error('should not call'); } });
  ctx.writeDrafts();
  const q = sheets['Outreach Queue'].data, H = q[0];
  assert.match(q[1][H.indexOf('Email Body')], /^Hi Jordan,/);
  assert.match(q[1][H.indexOf('LinkedIn Note')], /University of Pennsylvania/);
  assert.ok(q[1][H.indexOf('LinkedIn Note')].length <= 200);
});

test('LinkedIn queue + logging into Networking 2027 blank row', () => {
  const { sheets } = baseSheets();
  makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake }).ctx.findContacts();
  const { ctx } = makeContext(sheets, { fetch: () => ({}) });
  ctx.writeDrafts();
  const q = sheets['Outreach Queue'].data, H = q[0];
  // No approval step: both drafted people show up right away.
  const items = ctx.getLinkedInQueue();
  assert.deepStrictEqual(Array.from(items, i => i.name), ['Jordan Park', 'Priya Rao']);
  assert.match(items[1].why, /Shared: University of Pennsylvania/);
  ctx.markLinkedIn(items[1].row, 'sent', 'edited note');
  assert.strictEqual(q[2][H.indexOf('Status')], 'Sent');
  assert.strictEqual(q[2][H.indexOf('LinkedIn Note')], 'edited note');
  const log = sheets['Networking 2027'].data;
  assert.strictEqual(log.length, 3, 'reused the blank checkbox row');
  assert.strictEqual(log[2][0], 'Priya Rao');
  assert.strictEqual(log[2][2], 'LinkedIn');
  assert.strictEqual(log[2][6], true);
  // Jordan still has an unsent email, so LinkedIn sent doesn't flip him to Sent.
  ctx.markLinkedIn(items[0].row, 'sent', items[0].note);
  assert.strictEqual(q[1][H.indexOf('Status')], 'Draft Ready');
  assert.strictEqual(ctx.getLinkedInQueue().length, 0);
});

console.log('\n' + passed + ' test(s) passed' + (process.exitCode ? ', some FAILED' : ''));
