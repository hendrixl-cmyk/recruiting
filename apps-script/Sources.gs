/**
 * Turns your applications tab into the list of companies to network at.
 *
 * Every company you applied to (unless every application there is Rejected)
 * becomes a row in Targets with the roles you applied for, job titles to
 * search for, and the company's Apollo ID so the people search is exact.
 * Rows you add to Targets by hand (Source = "Manual") are left alone.
 */

function syncTargetsFromApplications_(apiKey, settings) {
  const appsTab = settings['Applications tab'] || '2027 Recruiting Tracker';
  if (!SpreadsheetApp.getActive().getSheetByName(appsTab)) {
    throw new Error('Could not find your applications tab "' + appsTab + '". Fix the name in Agent Settings.');
  }
  const companies = groupApplications_(openTable_(appsTab).rows, splitList_(settings['Skip application statuses']));
  const targets = openTable_(TABS.TARGETS);
  const perCompany = settingNumber_(settings, 'Contacts per company', 3);
  const defaultLocations = settings['Default locations'] || '';
  let added = 0;

  companies.forEach(c => {
    const existing = targets.rows.find(r => normalize_(r['Company']) === normalize_(c.company));
    const roles = c.roles.join('; ');
    if (existing) {
      // Keep the roles list current; never override what you edited by hand.
      if (existing['Source'] === 'Applications' && roles !== existing['Applied Roles']) {
        targets.set(existing, 'Applied Roles', roles);
        targets.set(existing, 'Titles', inferTitles_(c.roles).join(', '));
      }
      return;
    }
    const org = apiKey ? resolveOrganization_(apiKey, c.company) : null;
    targets.append({
      'Active': true,
      'Company': c.company,
      'Domain': org ? org.domain : '',
      'Apollo Org ID': org ? org.id : '',
      'Track': inferTrack_(c.company, c.roles),
      'Applied Roles': roles,
      'Titles': inferTitles_(c.roles).join(', '),
      'Locations': c.locations.length ? c.locations.join(', ') : defaultLocations,
      'Keywords': '',
      '# Contacts': perCompany,
      'Source': 'Applications',
    }, 'Company');
    added++;
  });
  return added;
}

/** Groups application rows by company. Pure function (unit-tested). */
function groupApplications_(rows, skipStatuses) {
  const skip = skipStatuses.map(s => s.toLowerCase());
  const byCompany = {};
  const order = [];
  rows.forEach(r => {
    const company = String(r['COMPANY'] || '').trim();
    if (!company) return;
    const key = normalize_(company);
    if (!byCompany[key]) { byCompany[key] = { company: company, roles: [], locations: [], live: false }; order.push(key); }
    const c = byCompany[key];
    const status = String(r['APPLICATION STATUS'] || '').trim().toLowerCase();
    if (skip.indexOf(status) === -1) c.live = true;
    const role = String(r['ROLE / POSITION'] || '').trim();
    if (role && c.roles.indexOf(role) === -1 && skip.indexOf(status) === -1) c.roles.push(role);
    splitList_(r['LOCATION']).forEach(loc => {
      const full = expandLocation_(loc);
      if (c.locations.indexOf(full) === -1) c.locations.push(full);
    });
  });
  return order.map(k => byCompany[k]).filter(c => c.live);
}

const LOCATION_ALIASES = {
  nyc: 'New York', ny: 'New York', brooklyn: 'New York', manhattan: 'New York',
  sf: 'San Francisco', bay: 'San Francisco', la: 'Los Angeles', scal: 'Los Angeles',
  dc: 'Washington, DC', chi: 'Chicago', philly: 'Philadelphia',
};

function expandLocation_(loc) {
  return LOCATION_ALIASES[normalize_(loc)] || loc;
}

function inferTrack_(company, roles) {
  const text = (company + ' ' + roles.join(' ')).toLowerCase();
  if (/mckinsey|bcg|boston consulting|bain|oliver wyman|parthenon|kearney|l\.?e\.?k|deloitte|accenture|strategy&|consult/.test(text)) return 'Consulting';
  if (/umg|universal music|sony music|warner music|spotify|tidal|vevo|live nation|caa|wme|music|record/.test(text)) return 'Music';
  if (/jp ?morgan|goldman|citi|morgan stanley|blackrock|bank|capital one|american express|amex|markets|trading|asset management|finance|investment/.test(text)) return 'Finance';
  if (/marketing|brand|merch|beauty|fashion|retail|consumer|dior|fendi|alo|elf|shiseido/.test(text)) return 'Consumer';
  if (/software|product|data|engineer|tech|tiktok|adobe|roblox/.test(text)) return 'Tech';
  return 'Other';
}

/** Job titles to search for, based on the roles you applied to. Early-career titles first. */
function inferTitles_(roles) {
  const text = roles.join(' ').toLowerCase();
  const titles = [];
  const add = list => list.forEach(t => { if (titles.indexOf(t) === -1) titles.push(t); });
  if (/consult/.test(text)) add(['Associate Consultant', 'Consultant', 'Business Analyst', 'Associate']);
  if (/strategy/.test(text)) add(['Strategy Analyst', 'Strategy Associate', 'Analyst']);
  if (/markets|trading|sales/.test(text)) add(['Sales and Trading Analyst', 'Markets Analyst', 'Trader', 'Analyst']);
  if (/asset management|investment/.test(text)) add(['Investment Analyst', 'Analyst', 'Associate']);
  if (/product/.test(text)) add(['Associate Product Manager', 'Product Manager', 'Product Analyst']);
  if (/marketing|brand/.test(text)) add(['Marketing Associate', 'Brand Manager', 'Marketing Manager']);
  if (/finance|fp&a|corporate analyst/.test(text)) add(['Financial Analyst', 'Analyst', 'Associate']);
  if (/risk|fraud|credit/.test(text)) add(['Risk Analyst', 'Credit Analyst']);
  if (/project management|operations/.test(text)) add(['Project Manager', 'Operations Analyst']);
  add(['Analyst', 'Associate']);
  return titles.slice(0, 10);
}

/** Looks up a company in Apollo by name. Returns {id, domain} or null. */
function resolveOrganization_(apiKey, name) {
  try {
    const res = apolloPost_(apiKey, 'mixed_companies/search', { q_organization_name: name, page: 1, per_page: 1 });
    const org = (res.organizations || res.accounts || [])[0];
    if (!org) return null;
    return { id: org.organization_id || org.id || '', domain: org.primary_domain || org.domain || '' };
  } catch (e) {
    return null; // Falls back to searching by company name; you can fill Domain in Targets.
  }
}

/** Picks which of your applied roles is closest to this person's title (for the message). */
function matchAppliedRole_(title, appliedRoles) {
  const roles = String(appliedRoles || '').split(';').map(r => r.trim()).filter(Boolean);
  if (!roles.length) return '';
  const words = String(title || '').toLowerCase().split(/[^a-z]+/).filter(w => w.length > 3);
  let best = roles[0], bestScore = 0;
  roles.forEach(r => {
    const lr = r.toLowerCase();
    const score = words.filter(w => lr.indexOf(w) !== -1).length;
    if (score > bestScore) { best = r; bestScore = score; }
  });
  return best;
}
