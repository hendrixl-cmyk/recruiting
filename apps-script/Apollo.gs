/**
 * Step 1: find people at your target companies with Apollo.io.
 *
 * Finds people at the companies in Targets (filled from your applications by
 * Sources.gs), ranked by how much they have in common with you.
 *
 * How Apollo's API works:
 *   - People API Search (mixed_people/api_search) is free and costs no credits,
 *     but hides emails (and part of the last name on free plans).
 *   - People Enrichment (people/bulk_match) costs 1 credit per person and
 *     returns the full name, LinkedIn URL and email.
 * So we search broadly, skip anyone you've already contacted, and only
 * enrich the people we're actually going to add.
 */

const APOLLO_BASE = 'https://api.apollo.io/api/v1/';

function findContacts() {
  const apiKey = getSecret_('APOLLO_API_KEY');
  if (!apiKey) { SpreadsheetApp.getUi().alert('Add your Apollo API key first: Recruiting Agent > Setup > Set API keys.'); return; }

  const started = Date.now();
  const settings = getSettings_();
  let enrichBudget = settingNumber_(settings, 'Max enrichments per run', 15);
  const defaultLocations = splitList_(settings['Default locations']);
  const similarity = splitList_(settings['Similarity keywords']);
  const draftEmails = String(settings['Also draft emails']).toLowerCase() !== 'no';
  const problems = [];

  // 1. Pull in any companies you've applied to since the last run.
  let newTargets = 0;
  try { newTargets = syncTargetsFromApplications_(apiKey, settings); }
  catch (e) { problems.push(e.message); }

  const targets = openTable_(TABS.TARGETS);
  const queue = openTable_(TABS.QUEUE);
  const existing = buildExistingIndex_(queue);
  let added = 0;

  // 2. Companies you haven't sourced yet go first, then the least recently sourced.
  const active = targets.rows
    .filter(t => t['Active'] === true && t['Company'])
    .sort((a, b) => lastRunTime_(a) - lastRunTime_(b));

  for (const target of active) {
    if (enrichBudget <= 0) { problems.push('Hit "Max enrichments per run" - run again later for the remaining companies.'); break; }
    if (Date.now() - started > 4.5 * 60 * 1000) { problems.push('Stopped early to stay under the 6-minute limit - run again to continue.'); break; }

    const company = String(target['Company']).trim();
    const want = Math.min(parseInt(target['# Contacts'], 10) || 3, enrichBudget);
    try {
      const keywords = similarity.concat(splitList_(target['Keywords']));
      const candidates = searchTarget_(apiKey, target, defaultLocations, keywords, existing, want);
      if (!candidates.length) { problems.push(company + ': no new people found (try broader Titles/Locations in Targets).'); targets.set(target, 'Last Run', new Date()); continue; }

      const people = apolloEnrich_(apiKey, candidates.map(c => c.id));
      enrichBudget -= candidates.length;

      people.forEach((p, i) => {
        if (!p) return;
        const c = candidates[i];
        const name = p.name || [p.first_name, p.last_name].filter(Boolean).join(' ');
        if (isExisting_(existing, { id: p.id, name: name, company: company, email: p.email, linkedin: p.linkedin_url })) return;
        const title = p.title || c.title || '';
        const appliedRole = matchAppliedRole_(title, target['Applied Roles']);
        const row = {};
        row[Q.STATUS] = STATUS.NEW;
        row[Q.CHANNEL] = pickChannel_(p.email, p.email_status, p.linkedin_url, draftEmails);
        row[Q.NAME] = name;
        row[Q.TITLE] = title;
        row[Q.COMPANY] = company;
        row[Q.APPLIED_ROLE] = appliedRole;
        row[Q.WHY] = whyThem_(c.matches, title, appliedRole);
        row[Q.LOCATION] = [p.city, p.state].filter(Boolean).join(', ');
        row[Q.LINKEDIN] = p.linkedin_url || linkedInSearchUrl_(name, company);
        row[Q.ABOUT] = summarizePerson_(p);
        row[Q.EMAIL] = draftEmails ? (p.email || '') : '';
        row[Q.EMAIL_STATUS] = p.email_status || (p.email ? '' : 'none found');
        row[Q.TRACK] = target['Track'] || 'Other';
        row[Q.FOUND] = new Date();
        row[Q.APOLLO_ID] = p.id;
        queue.append(row, Q.NAME);
        rememberExisting_(existing, { id: p.id, name: name, company: company, email: p.email, linkedin: p.linkedin_url });
        added++;
      });
      targets.set(target, 'Last Run', new Date());
    } catch (e) {
      problems.push(company + ': ' + e.message);
    }
  }

  const msg = (newTargets ? 'Added ' + newTargets + ' compan' + (newTargets === 1 ? 'y' : 'ies') + ' from your applications to Targets.\n' : '') +
    'Found ' + added + ' new people to connect with in "' + TABS.QUEUE + '".' +
    (added ? ' Next: step 2 writes their messages.' : '') +
    (problems.length ? '\n\nNotes:\n- ' + problems.join('\n- ') : '');
  SpreadsheetApp.getUi().alert(msg);
}

function lastRunTime_(t) {
  return t['Last Run'] instanceof Date ? t['Last Run'].getTime() : 0;
}

/**
 * Searches one company. Runs one free Apollo search per similarity keyword
 * (e.g. "University of Pennsylvania") plus one plain search, then ranks people
 * by how many keywords they matched. Returns the top `want` new people.
 */
function searchTarget_(apiKey, target, defaultLocations, keywords, existing, want) {
  const company = String(target['Company']).trim();
  const locations = splitList_(target['Locations']);
  const base = {
    person_titles: splitList_(target['Titles']),
    person_locations: locations.length ? locations : defaultLocations,
    per_page: 50,
    page: 1,
  };
  const orgId = String(target['Apollo Org ID'] || '').trim();
  const domain = String(target['Domain'] || '').trim();
  let companyKeyword = '';
  if (orgId) base.organization_ids = [orgId];
  else if (domain) base.q_organization_domains_list = [domain];
  else companyKeyword = company;

  const passes = keywords.map(k => ({ keyword: k, body: Object.assign({}, base, { q_keywords: [companyKeyword, k].filter(Boolean).join(' ') }) }));
  passes.push({ keyword: '', body: companyKeyword ? Object.assign({}, base, { q_keywords: companyKeyword }) : base });

  const byId = {};
  const order = [];
  passes.forEach(pass => {
    const res = apolloPost_(apiKey, 'mixed_people/api_search', pass.body);
    (res.people || []).forEach(p => {
      if (!byId[p.id]) {
        const last = p.last_name || p.last_name_obfuscated || '';
        if (isExisting_(existing, { id: p.id, first: p.first_name, lastInitial: last.charAt(0), company: company })) return;
        byId[p.id] = { id: p.id, title: p.title, matches: [], rank: order.length };
        order.push(p.id);
      }
      if (pass.keyword && byId[p.id].matches.indexOf(pass.keyword) === -1) byId[p.id].matches.push(pass.keyword);
    });
  });
  return rankCandidates_(order.map(id => byId[id])).slice(0, want);
}

/** Most shared keywords first; ties keep Apollo's order. Pure function (unit-tested). */
function rankCandidates_(candidates) {
  return candidates.slice().sort((a, b) => (b.matches.length - a.matches.length) || (a.rank - b.rank));
}

function whyThem_(matches, title, appliedRole) {
  const parts = [];
  if (matches.length) parts.push('Shared: ' + matches.join(', ') + ' (check their profile)');
  if (appliedRole) parts.push('Works near the role you applied for: ' + appliedRole);
  else if (title) parts.push('Role: ' + title);
  return parts.join('. ');
}

/** Enriches Apollo person IDs 10 at a time. Returns people in the same order (null if not matched). */
function apolloEnrich_(apiKey, ids) {
  const out = [];
  for (let i = 0; i < ids.length; i += 10) {
    const chunk = ids.slice(i, i + 10);
    const res = apolloPost_(apiKey, 'people/bulk_match', {
      details: chunk.map(id => ({ id: id })),
      reveal_personal_emails: false,
    });
    const matches = res.matches || [];
    chunk.forEach((_, j) => out.push(matches[j] || null));
  }
  return out;
}

function apolloPost_(apiKey, path, body) {
  const resp = UrlFetchApp.fetch(APOLLO_BASE + path, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Api-Key': apiKey, 'Cache-Control': 'no-cache' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const code = resp.getResponseCode();
  const text = resp.getContentText();
  if (code === 429) throw new Error('Apollo rate limit hit - wait a minute and run again.');
  if (code < 200 || code >= 300) throw new Error('Apollo ' + path + ' returned ' + code + ': ' + text.slice(0, 300));
  return JSON.parse(text);
}

/**
 * One-paragraph background from Apollo data, e.g.
 * "Consultant at BCG | ex-Deloitte. Now: Consultant @ BCG (since 2023-07). Before: Analyst @ Deloitte (2020-2023)."
 * Shown in the sheet and given to Claude so messages can reference real career facts.
 */
function summarizePerson_(p) {
  const parts = [];
  if (p.headline) parts.push(String(p.headline).trim());
  const jobs = p.employment_history || [];
  const current = jobs.filter(j => j.current);
  const past = jobs.filter(j => !j.current).slice(0, 3);
  const yr = d => (d ? String(d).slice(0, 4) : '');
  if (current.length) {
    parts.push('Now: ' + current.map(j => (j.title || 'Role') + ' @ ' + (j.organization_name || '') + (j.start_date ? ' (since ' + String(j.start_date).slice(0, 7) + ')' : '')).join('; '));
  }
  if (past.length) {
    parts.push('Before: ' + past.map(j => (j.title || 'Role') + ' @ ' + (j.organization_name || '') +
      (j.start_date || j.end_date ? ' (' + yr(j.start_date) + '-' + yr(j.end_date) + ')' : '')).join('; '));
  }
  const loc = [p.city, p.state].filter(Boolean).join(', ');
  if (loc && !parts.length) parts.push('Based in ' + loc);
  return parts.join('. ').replace(/\.\./g, '.');
}

/** When Apollo has no LinkedIn URL, link to a LinkedIn people search instead. */
function linkedInSearchUrl_(name, company) {
  return 'https://www.linkedin.com/search/results/people/?keywords=' + encodeURIComponent(name + ' ' + company);
}

// ---------- De-duplication against everyone you've already contacted ----------

function buildExistingIndex_(queue) {
  const idx = { ids: {}, emails: {}, linkedins: {}, names: {}, initials: {} };
  queue.rows.forEach(r => rememberExisting_(idx, {
    id: r[Q.APOLLO_ID], name: r[Q.NAME], company: r[Q.COMPANY], email: r[Q.EMAIL], linkedin: r[Q.LINKEDIN],
  }));
  // Your hand-kept networking tabs: NAME, COMPANY, and the email lives in "Column 1".
  ['Networking 2027', 'Networking 2026'].forEach(tab => {
    const sheet = SpreadsheetApp.getActive().getSheetByName(tab);
    if (!sheet) return;
    const t = openTable_(tab);
    t.rows.forEach(r => rememberExisting_(idx, { name: r['NAME'], company: r['COMPANY'], email: r['Column 1'] }));
  });
  return idx;
}

function rememberExisting_(idx, p) {
  if (p.id) idx.ids[p.id] = true;
  if (p.email) idx.emails[normalize_(p.email)] = true;
  if (p.linkedin && String(p.linkedin).indexOf('/search/') === -1) idx.linkedins[normalize_(p.linkedin)] = true;
  const name = String(p.name || '').trim();
  if (name) {
    const parts = name.split(/\s+/);
    idx.names[normalize_(name) + '|' + normalize_(p.company)] = true;
    if (parts.length > 1) {
      idx.initials[normalize_(parts[0]) + normalize_(parts[parts.length - 1]).charAt(0) + '|' + normalize_(p.company)] = true;
    }
  }
}

function isExisting_(idx, p) {
  if (p.id && idx.ids[p.id]) return true;
  if (p.email && idx.emails[normalize_(p.email)]) return true;
  if (p.linkedin && idx.linkedins[normalize_(p.linkedin)]) return true;
  if (p.name && idx.names[normalize_(p.name) + '|' + normalize_(p.company)]) return true;
  if (p.first && p.lastInitial &&
      idx.initials[normalize_(p.first) + normalize_(p.lastInitial) + '|' + normalize_(p.company)]) return true;
  return false;
}
