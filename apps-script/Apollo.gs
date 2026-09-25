/**
 * Step 1: find people at your target companies with Apollo.io.
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

  const targets = openTable_(TABS.TARGETS);
  const queue = openTable_(TABS.QUEUE);
  const existing = buildExistingIndex_(queue);

  let added = 0;
  const problems = [];

  for (const target of targets.rows) {
    if (target['Active'] !== true || !target['Company']) continue;
    if (enrichBudget <= 0) break;
    if (Date.now() - started > 4.5 * 60 * 1000) { problems.push('Stopped early to stay under the 6-minute limit - run again to continue.'); break; }

    const company = String(target['Company']).trim();
    const want = Math.min(parseInt(target['# Contacts'], 10) || 3, enrichBudget);
    try {
      const candidates = searchTarget_(apiKey, target, defaultLocations, existing, want);
      if (!candidates.length) { problems.push(company + ': no new people found (try broader titles/locations).'); continue; }

      const people = apolloEnrich_(apiKey, candidates.map(c => c.id));
      enrichBudget -= candidates.length;

      people.forEach((p, i) => {
        if (!p) return;
        const name = p.name || [p.first_name, p.last_name].filter(Boolean).join(' ');
        if (isExisting_(existing, { id: p.id, name: name, company: company, email: p.email, linkedin: p.linkedin_url })) return;
        const row = {};
        row[Q.STATUS] = STATUS.NEW;
        row[Q.CHANNEL] = pickChannel_(p.email, p.email_status, p.linkedin_url);
        row[Q.NAME] = name;
        row[Q.TITLE] = p.title || candidates[i].title || '';
        row[Q.COMPANY] = company;
        row[Q.LOCATION] = [p.city, p.state].filter(Boolean).join(', ');
        row[Q.LINKEDIN] = p.linkedin_url || linkedInSearchUrl_(name, company);
        row[Q.ABOUT] = summarizePerson_(p);
        row[Q.EMAIL] = p.email || '';
        row[Q.EMAIL_STATUS] = p.email_status || (p.email ? '' : 'none found');
        row[Q.PENN] = candidates[i].pennMatch ? 'Likely' : '';
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

  const msg = 'Added ' + added + ' new contact(s) to "' + TABS.QUEUE + '".' +
    (added ? ' Review them, fix "Penn Alum?" (Yes/No) if you know, then run step 2.' : '') +
    (problems.length ? '\n\nNotes:\n- ' + problems.join('\n- ') : '');
  SpreadsheetApp.getUi().alert(msg);
}

/**
 * Searches one target company. Penn alumni first (if Keywords is set), then
 * everyone else matching the titles, until we have `want` new people.
 */
function searchTarget_(apiKey, target, defaultLocations, existing, want) {
  const company = String(target['Company']).trim();
  const base = {
    person_titles: splitList_(target['Titles']),
    person_locations: splitList_(target['Locations']).length ? splitList_(target['Locations']) : defaultLocations,
    per_page: 50,
    page: 1,
  };
  const domain = String(target['Domain'] || '').trim();
  if (domain) base.q_organization_domains_list = [domain];
  else base.q_keywords = company;

  const passes = [];
  const keywords = String(target['Keywords'] || '').trim();
  if (keywords) passes.push({ pennMatch: true, body: Object.assign({}, base, { q_keywords: [base.q_keywords, keywords].filter(Boolean).join(' ') }) });
  passes.push({ pennMatch: false, body: base });

  const picked = [];
  const seen = {};
  for (const pass of passes) {
    if (picked.length >= want) break;
    const res = apolloPost_(apiKey, 'mixed_people/api_search', pass.body);
    for (const p of (res.people || [])) {
      if (picked.length >= want) break;
      if (seen[p.id]) continue;
      seen[p.id] = true;
      const last = p.last_name || p.last_name_obfuscated || '';
      if (isExisting_(existing, { id: p.id, first: p.first_name, lastInitial: last.charAt(0), company: company })) continue;
      picked.push({ id: p.id, title: p.title, pennMatch: pass.pennMatch });
    }
  }
  return picked;
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
