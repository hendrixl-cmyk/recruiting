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
      getUi: () => ({ alert: m => { alerts.push(m); return 'YES'; }, ButtonSet: {}, Button: { YES: 'YES' } }),
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
      'Targets': makeSheet('Targets', [
        TH,
        [true, 'BCG', 'bcg.com', 'Consulting', 'Consultant, Associate', 'New York', 'University of Pennsylvania', 2, ''],
        [false, 'Bain', 'bain.com', 'Consulting', 'Consultant', '', '', 2, ''],
        [false, '', '', '', '', '', '', '', ''], // checkbox-only row
      ]),
      'Outreach Queue': makeSheet('Outreach Queue', [QH]),
      'Networking 2027': makeSheet('Networking 2027', [
        ['NAME', 'COMPANY', 'FORM OF CONTACT', 'Column 1', 'DATE', 'LOCATION', 'REACHED OUT?', 'RESPONSE?', 'NOTES'],
        ['Amanda Cui', 'BCG', 'Email', '', '3/17/2026', 'New York', true, false, ''],
        ['', '', '', '', '', '', false, false, ''], // pre-filled checkbox row
      ]),
    },
  };
}

function apolloFake(url, body) {
  if (url.endsWith('mixed_people/api_search')) {
    return { json: { people: [
      { id: 'a1', first_name: 'Amanda', last_name_obfuscated: 'Cu***i', title: 'Consultant' }, // already contacted
      { id: 'a2', first_name: 'Jordan', last_name_obfuscated: 'Pa***k', title: 'Associate' },
      { id: 'a3', first_name: 'Priya', last_name_obfuscated: 'Ra***o', title: 'Consultant' },
      { id: 'a4', first_name: 'Extra', last_name_obfuscated: 'Pe***n', title: 'Consultant' },
    ] } };
  }
  if (url.endsWith('people/bulk_match')) {
    const db = {
      a2: { id: 'a2', name: 'Jordan Park', first_name: 'Jordan', last_name: 'Park', title: 'Associate', email: 'park.jordan@bcg.com', email_status: 'verified',
            linkedin_url: 'http://www.linkedin.com/in/jordanpark', city: 'New York', state: 'New York', headline: 'Associate at BCG',
            employment_history: [{ current: true, title: 'Associate', organization_name: 'BCG', start_date: '2024-08-01' },
                                 { current: false, title: 'Analyst', organization_name: 'Deloitte', start_date: '2021-06-01', end_date: '2024-07-01' }] },
      a3: { id: 'a3', name: 'Priya Rao', first_name: 'Priya', last_name: 'Rao', title: 'Consultant', email: null, email_status: 'unavailable', linkedin_url: null },
    };
    return { json: { matches: body.details.map(d => db[d.id] || null) } };
  }
  throw new Error('unexpected url ' + url);
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

test('findContacts: dedupes, enriches, builds About/LinkedIn', () => {
  const { sheets } = baseSheets();
  const { ctx, alerts, fetches } = makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake });
  ctx.findContacts();
  const q = sheets['Outreach Queue'].data;
  const H = q[0];
  const get = (r, h) => q[r][H.indexOf(h)];
  assert.strictEqual(q.length, 3, 'two new rows: ' + JSON.stringify(q.slice(1)) + alerts);
  assert.strictEqual(get(1, 'Name'), 'Jordan Park');
  assert.strictEqual(get(1, 'Channel'), 'Both');
  assert.strictEqual(get(1, 'Penn Alum?'), 'Likely');
  assert.match(get(1, 'About Them'), /Now: Associate @ BCG \(since 2024-08\)\. Before: Analyst @ Deloitte \(2021-2024\)/);
  assert.strictEqual(get(2, 'Name'), 'Priya Rao');
  assert.strictEqual(get(2, 'Channel'), 'LinkedIn');
  assert.match(get(2, 'LinkedIn URL'), /linkedin\.com\/search\/results\/people\/\?keywords=Priya%20Rao%20BCG/);
  // Amanda (already in Networking 2027) was never enriched -> no wasted credit.
  const enrichCalls = fetches.filter(f => f.url.endsWith('bulk_match'));
  assert.deepStrictEqual(enrichCalls[0].body.details.map(d => d.id), ['a2', 'a3']);
  const search = fetches[0].body;
  assert.deepStrictEqual(search.q_organization_domains_list, ['bcg.com']);
  assert.strictEqual(search.q_keywords, 'University of Pennsylvania');
  assert.strictEqual(fetches[0].headers['X-Api-Key'], 'k');
  // Only the active target ran.
  assert.ok(!fetches.some(f => JSON.stringify(f.body).includes('bain.com')));
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
      subject: 'Penn student - BCG question',
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
  assert.strictEqual(drafts.d1.to, 'park.jordan@bcg.com');
  assert.match(drafts.d1.body, /Best,\nHendrix Lee/);
  // Priya has no email -> no Gmail draft, still Draft Ready for LinkedIn.
  assert.strictEqual(get(2, 'Gmail Draft ID'), '');
  assert.strictEqual(get(2, 'Status'), 'Draft Ready');
  assert.ok(fetches[0].body.messages[0].content.includes('Deloitte'), 'background passed to Claude');
});

test('writeDrafts without Claude key uses template', () => {
  const { sheets } = baseSheets();
  makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake }).ctx.findContacts();
  const { ctx } = makeContext(sheets, { secrets: {}, fetch: () => { throw new Error('should not call'); } });
  ctx.writeDrafts();
  const q = sheets['Outreach Queue'].data, H = q[0];
  assert.match(q[1][H.indexOf('Email Body')], /^Hi Jordan,/);
  assert.ok(q[1][H.indexOf('LinkedIn Note')].length <= 200);
});

test('LinkedIn queue + logging into Networking 2027 blank row', () => {
  const { sheets } = baseSheets();
  makeContext(sheets, { secrets: { APOLLO_API_KEY: 'k' }, fetch: apolloFake }).ctx.findContacts();
  const { ctx } = makeContext(sheets, { fetch: () => ({}) });
  ctx.writeDrafts();
  const q = sheets['Outreach Queue'].data, H = q[0];
  q[2][H.indexOf('Status')] = 'Approved'; // approve Priya (LinkedIn only)
  const items = ctx.getLinkedInQueue();
  assert.strictEqual(items.length, 1);
  assert.strictEqual(items[0].name, 'Priya Rao');
  ctx.markLinkedIn(items[0].row, 'sent', 'edited note');
  assert.strictEqual(q[2][H.indexOf('Status')], 'Sent');
  assert.strictEqual(q[2][H.indexOf('LinkedIn Note')], 'edited note');
  const log = sheets['Networking 2027'].data;
  assert.strictEqual(log.length, 3, 'reused the blank checkbox row');
  assert.strictEqual(log[2][0], 'Priya Rao');
  assert.strictEqual(log[2][2], 'LinkedIn');
  assert.strictEqual(log[2][6], true);
  assert.strictEqual(ctx.getLinkedInQueue().length, 0);
});

console.log('\n' + passed + ' test(s) passed' + (process.exitCode ? ', some FAILED' : ''));
