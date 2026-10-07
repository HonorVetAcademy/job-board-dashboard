'use strict';
/* ═══════════════════════════════════════════════════════════
   Job Board Usage Dashboard — app.js
   HonorVet Technologies
   Parses exact schema from Job Postings.xlsx + Resume Searches.xlsx
   ═══════════════════════════════════════════════════════════ */

// ── BOARD CONFIG ──────────────────────────────────────────
const BOARD_CFG = {
  'Dice':           { color: '#F04E23', icon: '🎲' },
  'LinkedIn':       { color: '#0077B5', icon: '💼' },
  'Indeed':         { color: '#003A9B', icon: '🔍' },
  'Vivian':         { color: '#7B2FBE', icon: '🏥' },
  'Website':        { color: '#00897B', icon: '🌐' },
  'DocCafe':        { color: '#E91E63', icon: '☕' },
  'Monster':        { color: '#6B0D1E', icon: '👾' },
  'CareerBuilder':  { color: '#F57C00', icon: '🏗️' },
  'Resume Library': { color: '#0288D1', icon: '📚' },
  'Signal Hire':    { color: '#43A047', icon: '📡' },
  'Crintell':       { color: '#9C27B0', icon: '🧠' },
};
function boardColor(name) {
  for (const [k, v] of Object.entries(BOARD_CFG)) {
    if (name.toLowerCase().includes(k.toLowerCase())) return v.color;
  }
  const palette = ['#0077B5','#F04E23','#00897B','#7B2FBE','#003A9B','#F57C00','#E91E63','#0288D1'];
  let h = 0; for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) | 0;
  return palette[Math.abs(h) % palette.length];
}

// ── APP STATE ─────────────────────────────────────────────
const APP = {
  raw:      { postings: null, searches: null },
  postRecs: [],   // Normalized posting records
  srchRecs: [],   // Normalized search records
  // Board summaries
  boardPost: {},  // { board: { jobs, views, applications, spend, recruiters, verticals } }
  boardSrch: {},  // { board: { searches, views, contacts, responses, respRate, recruiters } }
  // Recruiter summaries
  recPost:   {},  // { recruiter: { jobs, views, applications, boards } }
  recSrch:   {},  // { recruiter: { views, boards } }
  // Recruiter × Board breakdowns (for matrix tables)
  recPostByBoard: {},  // { recruiter: { board: { jobs, views, applications } } }
  recSrchByBoard: {},  // { recruiter: { board: { views } } }
  // Time series (keyed by 'YYYY-MM-DD')
  postByDate: {},
  srchByDate: {},
  // Filter options
  allBoards:    [],
  allVerticals: [],
  allRecruiters:[],
  // Active filtered copies
  fPostRecs: [],
  fSrchRecs: [],
  charts: {},
  trendMode: { post: 'daily', srch: 'daily' },
};

// ── UTILITIES ─────────────────────────────────────────────
const fmtNum  = n => n == null ? '—' : Number(n).toLocaleString();
const fmtDec  = (n,d=1) => n == null ? '—' : Number(n).toFixed(d);
const fmtPct  = n => n == null ? '—' : (n * 100).toFixed(1) + '%';
const fmtDol  = n => n == null ? '—' : '$' + Number(n).toLocaleString(undefined,{maximumFractionDigits:2});

function toDate(v) {
  if (!v) return null;
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number') {
    // Excel serial date: 1 = Jan 1 1900, correcting for leap year bug
    const ms = (v - 25569) * 86400000;
    const d = new Date(ms);
    return isNaN(d) ? null : d;
  }
  if (typeof v === 'string') {
    const d = new Date(v.trim());
    return isNaN(d) ? null : d;
  }
  return null;
}
function dateFmt(d, fmt = 'YYYY-MM-DD') {
  if (!d) return null;
  const Y = d.getFullYear();
  const M = String(d.getMonth() + 1).padStart(2, '0');
  const D = String(d.getDate()).padStart(2, '0');
  if (fmt === 'YYYY-MM-DD') return `${Y}-${M}-${D}`;
  if (fmt === 'MM/DD') return `${M}/${D}`;
  if (fmt === 'Mon D') {
    const mon = d.toLocaleString('en',{month:'short'});
    return `${mon} ${D}`;
  }
  return `${Y}-${M}-${D}`;
}
function isoWeek(d) {
  const jan1 = new Date(d.getFullYear(), 0, 1);
  const w = Math.ceil(((d - jan1) / 86400000 + jan1.getDay() + 1) / 7);
  return `${d.getFullYear()}-W${String(w).padStart(2,'0')}`;
}
function numVal(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  const n = parseFloat(String(v).replace(/[^0-9.\-]/g,''));
  return isNaN(n) ? 0 : n;
}
function strVal(v) { return v == null ? '' : String(v).trim(); }
function cleanName(v) {
  return strVal(v).replace(/^[ \t]+|[ \t]+$/g,'').replace(/\s+/,' ');
}

// ── SHEET PARSER HELPERS ──────────────────────────────────
function sheetToArray(wb, name) {
  const ws = wb.Sheets[name];
  if (!ws) return null;
  return XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true });
}

// Find header row by looking for known column names
function findHeader(rows, mustHave) {
  for (let i = 0; i < Math.min(25, rows.length); i++) {
    const row = (rows[i] || []).map(c => strVal(c).toLowerCase().trim());
    const hits = mustHave.filter(col => row.some(cell => cell.includes(col.toLowerCase())));
    if (hits.length >= Math.min(mustHave.length, 2)) {
      return { idx: i, headers: (rows[i] || []).map(c => strVal(c).trim()) };
    }
  }
  return null;
}

function makeRecord(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { if (h) obj[h] = row[i] != null ? row[i] : null; });
  return obj;
}

function dataRows(rows, headerResult) {
  if (!headerResult) return [];
  const out = [];
  for (let i = headerResult.idx + 1; i < rows.length; i++) {
    const r = rows[i] || [];
    if (r.every(c => c == null || c === '')) continue;
    out.push(makeRecord(headerResult.headers, r));
  }
  return out;
}

// ── PARSERS — JOB POSTINGS FILE ──────────────────────────

function parseDicePosting(wb) {
  const rows = sheetToArray(wb, 'Dice_Posting');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Job Name', 'Job Views', 'Applications']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => ({
    board: 'Dice',
    date: toDate(r['Created Date']),
    recruiter: cleanName(r['Posted By'] || '').replace('@honorvettech.com','').replace('@honorvet.com','').trim(),
    vertical: strVal(r['Vertical']) || 'IT',
    jobTitle: strVal(r['Job Name']),
    positionId: strVal(r['Position ID']),
    views: numVal(r['Job Views']),
    applications: numVal(r['Applications']),
    alerts: numVal(r['Job Alerts']),
    active: strVal(r['IsActive']).toLowerCase() === 'yes',
    spend: 0,
  })).filter(r => r.jobTitle && !r.jobTitle.toLowerCase().includes('total'));
}

function parseLinkedInPosting(wb) {
  const rows = sheetToArray(wb, 'LinkedIn_Posting');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Job Title', 'Total Views', 'Applications', 'Job Poster']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => ({
    board: 'LinkedIn',
    date: toDate(r['Original List Date']),
    recruiter: cleanName(r['Job Poster'] || ''),
    vertical: strVal(r['Vertical']) || 'IT',
    jobTitle: strVal(r['Job Title'] || r['Standardized Title']),
    positionId: strVal(r['Employer Job ID']),
    views: numVal(r['Total Views']),
    applications: numVal(r['Applications']),
    alerts: 0,
    active: strVal(r['Current Status']).toUpperCase() === 'LISTED',
    spend: 0,
  })).filter(r => r.jobTitle);
}

function parseIndeedPostings(wb) {
  const rows = sheetToArray(wb, 'Indeed Analysis');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Job Title', 'Impressions', 'Clicks', 'Applies', 'Spend']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => ({
    board: 'Indeed',
    date: toDate(r['Job Created'] || r['Last updated']),
    recruiter: cleanName(r['Posted By'] || ''),
    vertical: strVal(r['Vertical']) || '',
    jobTitle: strVal(r['Job Title']),
    positionId: strVal(r['Reference #']),
    views: numVal(r['Impressions']),
    applications: numVal(r['Applies']),
    alerts: numVal(r['Apply starts']),
    active: strVal(r['Job status']).toLowerCase() === 'active',
    spend: numVal(r['Spend']),
    clicks: numVal(r['Clicks']),
    cpa: numVal(r['Cost per apply (CPA)']),
  })).filter(r => r.jobTitle);
}

function parseWebsitePosting(wb) {
  const rows = sheetToArray(wb, 'Website_Posting');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Job Title', 'Posted by', 'Total Applicants']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => ({
    board: 'Website',
    date: toDate(r['Date Posted'] || r['Date']),
    recruiter: cleanName(r['Posted by'] || ''),
    vertical: strVal(r['Vertical']) || '',
    jobTitle: strVal(r['Job Title']),
    positionId: strVal(r['JobDiva #']),
    views: 0,
    applications: numVal(r['Total Applicants']),
    alerts: 0,
    active: strVal(r['Job Status']).toUpperCase() === 'OPEN',
    spend: 0,
  })).filter(r => r.jobTitle);
}

function parseDocCafeData(wb) {
  const rows = sheetToArray(wb, 'DocCafe Data');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Job Title', 'Views Count', 'Applications Count']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => ({
    board: 'DocCafe',
    date: toDate(r['Date Initially Published']),
    recruiter: cleanName(strVal(r['Candidate Contact']).split('@')[0].replace(/\./g,' ')),
    vertical: 'Healthcare',
    jobTitle: strVal(r['Job Title']),
    positionId: strVal(r['Job Id']),
    views: numVal(r['Views Count']),
    applications: numVal(r['Applications Count']),
    alerts: 0,
    active: strVal(r['Job Status']).toLowerCase() === 'active',
    spend: 0,
  })).filter(r => r.jobTitle);
}

function parseVivianPosting(wb) {
  const rows = sheetToArray(wb, 'Vivian_Posting');
  if (!rows) return [];
  // 90-column raw dump; header at row 0
  const hdr = { idx: 0, headers: (rows[0] || []).map(c => strVal(c).trim()) };
  const seen = new Set();
  return dataRows(rows, hdr).map(r => {
    const id = strVal(r['objectID'] || r['JobID']);
    if (seen.has(id)) return null;
    seen.add(id);
    return {
      board: 'Vivian',
      date: toDate(r['Date '] || r['Date'] || r['createdAt']),
      recruiter: cleanName(r['recruiterFullName'] || r['Recruiter'] || ''),
      vertical: 'Healthcare',
      jobTitle: strVal(r['Job'] || r['searchTitle']),
      positionId: id,
      views: 0,
      applications: numVal(r['Inbound']) + numVal(r['Proposal']),
      alerts: 0,
      active: true,
      inbound: numVal(r['Inbound']),
      proposal: numVal(r['Proposal']),
      spend: 0,
    };
  }).filter(Boolean).filter(r => r.jobTitle);
}

function parseAllPostings(wb) {
  const all = [
    ...parseDicePosting(wb),
    ...parseLinkedInPosting(wb),
    ...parseIndeedPostings(wb),
    ...parseWebsitePosting(wb),
    ...parseDocCafeData(wb),
    ...parseVivianPosting(wb),
  ];
  return all;
}

// ── PARSERS — RESUME SEARCHES FILE ───────────────────────

function parseJobDivaResumes(wb) {
  // Full log: Date, Recruiter, Source, Candidate, Vertical
  const rows = sheetToArray(wb, 'JobDiva Resumes');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Date', 'Recruiter', 'Source', 'Candidate']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => ({
    board: mapJobDivaSource(strVal(r['Source'])),
    date: toDate(r['Date']),
    recruiter: cleanName(r['Recruiter'] || ''),
    vertical: strVal(r['Vertical']) || '',
    candidate: strVal(r['Candidate']),
    source: strVal(r['Source']),
    views: 1,
    searches: 0,
    contacts: 0,
    responses: 0,
  })).filter(r => r.candidate);
}

function mapJobDivaSource(src) {
  const s = src.toLowerCase().trim();
  if (!s) return 'Other';
  if (s.includes('indeed'))         return 'Indeed';
  if (s.includes('linkedin'))       return 'LinkedIn';
  if (s.includes('monster'))        return 'Monster';
  if (s.includes('careerbuilder') || s.includes('crintell')) return 'CareerBuilder';
  if (s.includes('dice'))           return 'Dice';
  if (s.includes('vivian'))         return 'Vivian';
  if (s.includes('resume library') || s.includes('resume-library') || s.includes('resumelibrary')) return 'Resume Library';
  if (s.includes('doccafe') || s.includes('docafe'))  return 'DocCafe';
  if (s.includes('facebook'))       return 'Facebook';
  if (s.includes('ziprecruiter'))   return 'ZipRecruiter';
  if (s.includes('glassdoor'))      return 'Glassdoor';
  // Catch-all for internal/misc sources that aren't job boards
  if (s.includes('internal') || s.includes('emailed') || s.includes('formatted') || s.includes('sent')) return 'Internal / Email';
  if (s.includes('ilabor') || s.includes('staff') || s.includes('sourced')) return 'Other Sourcing';
  return src.trim() || 'Other';
}

function parseJobDivaAnalysis(wb) {
  // Pivot: Source, Date, Recruiter, Count of Candidate
  const rows = sheetToArray(wb, 'JobDiva Analysis');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Source', 'Date', 'Recruiter', 'Count']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => {
    const count = numVal(r['Count of Candidate'] || r['Count']);
    if (!count) return null;
    return {
      board: mapJobDivaSource(strVal(r['Source'])),
      date: toDate(r['Date']),
      recruiter: cleanName(r['Recruiter'] || ''),
      views: count,
      searches: count,
      contacts: 0, responses: 0,
    };
  }).filter(Boolean);
}

function parseDiceResumes(wb) {
  // Recruiter pivot — individual counts
  const rows = sheetToArray(wb, 'Dice Resume Analysis');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Recruiters', 'Grand Total']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => {
    const recruiter = cleanName(r['Recruiters'] || '');
    if (!recruiter || recruiter.toLowerCase().includes('grand')) return null;
    const total = numVal(r['Grand Total'] || r['Total']);
    return { board: 'Dice', recruiter, views: total, searches: total, contacts: 0, responses: 0, date: null };
  }).filter(Boolean);
}

function parseLinkedInSearches(wb) {
  // LinkedIn InMails: Seat Holders, Sends, Responses, Accepts
  const rows = sheetToArray(wb, 'Linkedn Analysis');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Seat Holders', 'Sends', 'Responses', 'Response Rate']);
  if (!hdr) return [];
  // Also extract totals from summary rows
  let creditsUsed = 0, creditsLeft = 0;
  for (let i = 0; i < hdr.idx; i++) {
    const r = rows[i] || [];
    r.forEach((c, ci) => {
      const s = strVal(c).toLowerCase();
      if (s.includes('credits used')) creditsUsed = numVal(r[ci + 1] || r[ci + 2] || 0);
      if (s.includes('remaining')) creditsLeft = numVal(r[ci + 1] || r[ci + 2] || 0);
    });
  }
  const records = dataRows(rows, hdr).map(r => {
    const recruiter = cleanName(r['Seat Holders'] || r[Object.keys(r)[0]] || '');
    if (!recruiter) return null;
    const sends = numVal(r['Sends']);
    const resp  = numVal(r['Responses']);
    const acc   = numVal(r['Accepts']);
    const rRate = numVal(r['Response Rate %']) / 100;
    return {
      board: 'LinkedIn', recruiter, date: null,
      searches: sends, views: sends, contacts: acc, responses: resp,
      responseRate: rRate,
    };
  }).filter(Boolean);
  APP.meta = APP.meta || {};
  APP.meta.linkedinCreditsUsed = creditsUsed;
  APP.meta.linkedinCreditsLeft = creditsLeft;
  return records;
}

function parseIndeedSearches(wb) {
  // Sheet in Resume Searches: smart sourcing
  const rows = sheetToArray(wb, 'Indeed Analysis');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Recruiter', 'Searches', 'Resumes viewed', 'Contacts used']);
  if (!hdr) return [];
  let cGranted = 0, cUsed = 0, cLeft = 0;
  for (let i = 0; i < hdr.idx; i++) {
    const r = rows[i] || [];
    r.forEach((c, ci) => {
      const s = strVal(c).toLowerCase();
      if (s.includes('contacts granted')) cGranted = numVal(r[ci + 1] || 0);
      if (s.includes('contacts used'))    cUsed    = numVal(r[ci + 1] || r[ci + 2] || 0);
      if (s.includes('contacts left'))    cLeft    = numVal(r[ci + 1] || r[ci + 2] || 0);
    });
  }
  APP.meta = APP.meta || {};
  APP.meta.indeedContactsGranted = cGranted;
  APP.meta.indeedContactsUsed = cUsed;
  return dataRows(rows, hdr).map(r => {
    const recruiter = cleanName(r['Recruiter'] || '');
    if (!recruiter) return null;
    return {
      board: 'Indeed',
      recruiter, date: null,
      searches: numVal(r['Searches']),
      views: numVal(r['Resumes viewed']),
      contacts: numVal(r['Contacts used']),
      responses: numVal(r['Responses']),
    };
  }).filter(Boolean);
}

function parseMonsterAnalysis(wb) {
  // Credits summary + search type breakdown
  const rows = sheetToArray(wb, 'Monster Analysis');
  if (!rows) return [];
  let creditsRcvd = 0, creditsUsed = 0, creditsAvail = 0;
  const fetches = [];
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i] || [];
    const rowTxt = r.map(strVal).join('|').toLowerCase();
    if (rowTxt.includes('credits received') && rowTxt.includes('credits used')) {
      // next row has values
      const next = rows[i + 1] || [];
      next.forEach((c, ci) => {
        const n = numVal(c);
        if (n > 0 && ci === 1) creditsRcvd = n;
        if (n > 0 && ci === 2) creditsUsed = n;
        if (n > 0 && ci === 3) creditsAvail = n;
      });
    }
    r.forEach((c, ci) => {
      const s = strVal(c).toLowerCase();
      if (s.includes('integrated search') || s.includes('resume search')) {
        const fetchCount  = numVal(r[ci + 1]);
        const creditSpend = numVal(r[ci + 2]);
        fetches.push({ type: strVal(c), count: fetchCount, spend: creditSpend });
      }
    });
  }
  const totalFetch = fetches.reduce((s, f) => s + f.count, 0);
  APP.meta = APP.meta || {};
  APP.meta.monsterCreditsRcvd  = creditsRcvd;
  APP.meta.monsterCreditsUsed  = creditsUsed;
  APP.meta.monsterCreditsAvail = creditsAvail;
  APP.meta.monsterFetches = fetches;
  if (totalFetch === 0) return [];
  return fetches.map(f => ({
    board: 'Monster', recruiter: 'System', date: null,
    searches: f.count, views: f.count, contacts: 0, responses: 0, spend: f.spend,
  }));
}

function parseResumeLibrary(wb) {
  const rows = sheetToArray(wb, 'Resume-Library Analysis');
  if (!rows) return [];
  const hdr = findHeader(rows, ['User Name', 'Searches', 'Views', 'Unlocks']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => {
    const recruiter = cleanName(r['User Name'] || '');
    if (!recruiter) return null;
    const searches = numVal(r['Searches']);
    const views    = numVal(r['Views']);
    const unlocks  = numVal(r['Unlocks']);
    if (searches === 0 && views === 0 && unlocks === 0) return null;
    return { board: 'Resume Library', recruiter, date: null, searches, views, contacts: unlocks, responses: 0 };
  }).filter(Boolean);
}

function parseSignalHire(wb) {
  const rows = sheetToArray(wb, 'Signal Hire');
  if (!rows) return [];
  const hdr = findHeader(rows, ['Account', 'Unlocks']);
  if (!hdr) return [];
  return dataRows(rows, hdr).map(r => {
    const recruiter = cleanName(strVal(r['Account']).split('@')[0].replace(/\./g,' '));
    const unlocks = numVal(r['Unlocks']);
    if (!recruiter || unlocks === 0) return null;
    return { board: 'Signal Hire', recruiter, date: null, searches: unlocks, views: unlocks, contacts: unlocks, responses: 0 };
  }).filter(Boolean);
}

function parseVivianCandidates(wb) {
  const rows = sheetToArray(wb, 'Vivian Candidates');
  if (!rows) return [];
  const hdr = { idx: 0, headers: (rows[0] || []).map(c => strVal(c).trim()) };
  return dataRows(rows, hdr).map(r => {
    const recruiter = cleanName((r['recruiterFirstName'] || '') + ' ' + (r['recruiterLastName'] || ''));
    const credits   = numVal(r['Credit Used'] || r['Premium Credits'] || 0);
    const date      = toDate(r['Date'] || r['createdAtDate']);
    return {
      board: 'Vivian', recruiter, date,
      searches: 1, views: 1, contacts: numVal(r['creditGranted'] || 0),
      responses: 0,
      isPremium: strVal(r['isPremium']).toLowerCase() === 'true',
      discipline: strVal(r['discipline']),
      credits,
    };
  }).filter(r => r.recruiter && r.recruiter.trim());
}

function parseAllSearches(wb) {
  // Use JobDiva Resumes (individual) as primary source for per-date data
  // Supplement with board-specific summaries
  const jdResumes = parseJobDivaResumes(wb);
  const liSearches = parseLinkedInSearches(wb);
  const indeedSrch = parseIndeedSearches(wb);
  const monsterSrch = parseMonsterAnalysis(wb);
  const rlSrch = parseResumeLibrary(wb);
  const signalSrch = parseSignalHire(wb);
  const vivCands = parseVivianCandidates(wb);
  const diceResumes = parseDiceResumes(wb);

  // For boards that appear in BOTH jdResumes AND specific parsers, avoid double-counting
  // jdResumes has: Indeed, LinkedIn, Monster, CareerBuilder, Dice (through JobDiva)
  // Specific parsers cover direct usage; use jdResumes for those boards' date-level data
  // Add board summaries only for boards NOT in jdResumes
  return [...jdResumes, ...liSearches, ...indeedSrch, ...monsterSrch,
          ...rlSrch, ...signalSrch, ...vivCands, ...diceResumes];
}

// ── AGGREGATION ───────────────────────────────────────────
function aggregate() {
  APP.boardPost      = {};
  APP.boardSrch      = {};
  APP.recPost        = {};
  APP.recSrch        = {};
  APP.recPostByBoard = {};
  APP.recSrchByBoard = {};
  APP.postByDate     = {};
  APP.srchByDate     = {};

  // Posting records
  for (const r of APP.fPostRecs) {
    const b = r.board;
    if (!APP.boardPost[b]) APP.boardPost[b] = { jobs: 0, views: 0, applications: 0, spend: 0, recruiters: new Set(), verticals: new Set() };
    APP.boardPost[b].jobs++;
    APP.boardPost[b].views += r.views;
    APP.boardPost[b].applications += r.applications;
    APP.boardPost[b].spend += r.spend || 0;
    if (r.recruiter) APP.boardPost[b].recruiters.add(r.recruiter);
    if (r.vertical)  APP.boardPost[b].verticals.add(r.vertical);

    if (r.recruiter) {
      if (!APP.recPost[r.recruiter]) APP.recPost[r.recruiter] = { jobs: 0, views: 0, applications: 0, boards: new Set(), vertical: r.vertical || '' };
      APP.recPost[r.recruiter].jobs++;
      APP.recPost[r.recruiter].views += r.views;
      APP.recPost[r.recruiter].applications += r.applications;
      APP.recPost[r.recruiter].boards.add(b);
      // Recruiter × Board breakdown
      if (!APP.recPostByBoard[r.recruiter]) APP.recPostByBoard[r.recruiter] = {};
      if (!APP.recPostByBoard[r.recruiter][b]) APP.recPostByBoard[r.recruiter][b] = { jobs:0, views:0, applications:0 };
      APP.recPostByBoard[r.recruiter][b].jobs++;
      APP.recPostByBoard[r.recruiter][b].views += r.views;
      APP.recPostByBoard[r.recruiter][b].applications += r.applications;
    }

    if (r.date) {
      const dk = dateFmt(r.date);
      if (!APP.postByDate[dk]) APP.postByDate[dk] = { total: 0, byBoard: {} };
      APP.postByDate[dk].total++;
      APP.postByDate[dk].byBoard[b] = (APP.postByDate[dk].byBoard[b] || 0) + 1;
    }
  }

  // Search records
  for (const r of APP.fSrchRecs) {
    const b = r.board;
    if (!APP.boardSrch[b]) APP.boardSrch[b] = { searches: 0, views: 0, contacts: 0, responses: 0, recruiters: new Set() };
    APP.boardSrch[b].searches += r.searches || 0;
    APP.boardSrch[b].views    += r.views    || 0;
    APP.boardSrch[b].contacts += r.contacts || 0;
    APP.boardSrch[b].responses+= r.responses|| 0;
    if (r.recruiter) APP.boardSrch[b].recruiters.add(r.recruiter);

    if (r.recruiter) {
      if (!APP.recSrch[r.recruiter]) APP.recSrch[r.recruiter] = { views: 0, boards: new Set() };
      APP.recSrch[r.recruiter].views += r.views || 0;
      APP.recSrch[r.recruiter].boards.add(b);
      // Recruiter × Board breakdown
      if (!APP.recSrchByBoard[r.recruiter]) APP.recSrchByBoard[r.recruiter] = {};
      if (!APP.recSrchByBoard[r.recruiter][b]) APP.recSrchByBoard[r.recruiter][b] = { views:0 };
      APP.recSrchByBoard[r.recruiter][b].views += r.views || 0;
    }

    if (r.date) {
      const dk = dateFmt(r.date);
      if (!APP.srchByDate[dk]) APP.srchByDate[dk] = { total: 0, byBoard: {} };
      APP.srchByDate[dk].total += r.views || 1;
      APP.srchByDate[dk].byBoard[b] = (APP.srchByDate[dk].byBoard[b] || 0) + (r.views || 1);
    }
  }
}

// ── FILTERS ───────────────────────────────────────────────
function applyFilters() {
  const from = document.getElementById('flt-from').value;
  const to   = document.getElementById('flt-to').value;
  const board    = document.getElementById('flt-board').value;
  const vertical = document.getElementById('flt-vertical').value;
  const recruiter= document.getElementById('flt-recruiter').value;

  const dFrom = from ? new Date(from) : null;
  const dTo   = to   ? new Date(to + 'T23:59:59') : null;

  APP.fPostRecs = APP.postRecs.filter(r => {
    if (dFrom && r.date && r.date < dFrom) return false;
    if (dTo   && r.date && r.date > dTo)   return false;
    if (board    && r.board    !== board)    return false;
    if (vertical && r.vertical !== vertical) return false;
    if (recruiter&& r.recruiter!== recruiter)return false;
    return true;
  });

  APP.fSrchRecs = APP.srchRecs.filter(r => {
    if (dFrom && r.date && r.date < dFrom) return false;
    if (dTo   && r.date && r.date > dTo)   return false;
    if (board    && r.board    !== board)    return false;
    if (recruiter&& r.recruiter!== recruiter)return false;
    return true;
  });

  const activeFilters = [from, to, board, vertical, recruiter].filter(Boolean).length;
  const countEl = document.getElementById('flt-count');
  countEl.textContent = activeFilters ? `${activeFilters} filter${activeFilters > 1 ? 's' : ''} active` : '';

  aggregate();
  renderKPIs();
  renderCharts();
  renderTables();
  renderInsights();
}

function resetFilters() {
  ['flt-from','flt-to','flt-board','flt-vertical','flt-recruiter'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  applyFilters();
}

function populateFilters() {
  const boards     = [...new Set(APP.postRecs.map(r => r.board).concat(APP.srchRecs.map(r => r.board)))].filter(Boolean).sort();
  const verticals  = [...new Set(APP.postRecs.map(r => r.vertical))].filter(Boolean).sort();
  const recruiters = [...new Set(APP.postRecs.map(r => r.recruiter).concat(APP.srchRecs.map(r => r.recruiter)))].filter(r => r && r.length > 1).sort();

  const sel = (id, opts) => {
    const el = document.getElementById(id);
    if (!el) return;
    const cur = el.value;
    el.innerHTML = `<option value="">All ${id.includes('board') ? 'Boards' : id.includes('vertical') ? 'Verticals' : 'Recruiters'}</option>`;
    opts.forEach(o => { const opt = document.createElement('option'); opt.value = o; opt.textContent = o; el.appendChild(opt); });
    el.value = cur;
  };
  sel('flt-board', boards);
  sel('flt-vertical', verticals);
  sel('flt-recruiter', recruiters);
}

// ── KPI CARDS ─────────────────────────────────────────────
function renderKPIs() {
  const totalJobs  = APP.fPostRecs.length;
  const totalViews = APP.fPostRecs.reduce((s, r) => s + r.views, 0);
  const totalApps  = APP.fPostRecs.reduce((s, r) => s + r.applications, 0);
  const totalSrch  = APP.fSrchRecs.reduce((s, r) => s + (r.views || 0), 0);
  const totalSpend = APP.fPostRecs.reduce((s, r) => s + (r.spend || 0), 0);

  const boardsByPost = Object.entries(APP.boardPost).sort((a, b) => b[1].jobs - a[1].jobs);
  const topPost  = boardsByPost[0];
  const leastPost= boardsByPost[boardsByPost.length - 1];

  const allRecruiterNames = new Set([
    ...Object.keys(APP.recPost),
    ...Object.keys(APP.recSrch),
  ]);
  const activeRecs = [...allRecruiterNames].filter(r => r && r.length > 1).length;

  // WoW & MoM growth (compare posting counts by week/month)
  const dateSorted = Object.keys(APP.postByDate).sort();
  let wowGrowth = null, momGrowth = null;
  if (dateSorted.length > 0) {
    const lastDate = new Date(dateSorted[dateSorted.length - 1]);
    const thisWeekStart = new Date(lastDate); thisWeekStart.setDate(lastDate.getDate() - 6);
    const prevWeekStart = new Date(thisWeekStart); prevWeekStart.setDate(thisWeekStart.getDate() - 7);
    let thisW = 0, prevW = 0;
    for (const [dk, v] of Object.entries(APP.postByDate)) {
      const d = new Date(dk);
      if (d >= thisWeekStart && d <= lastDate) thisW += v.total;
      if (d >= prevWeekStart && d < thisWeekStart) prevW += v.total;
    }
    if (prevW > 0) wowGrowth = (thisW - prevW) / prevW;
    else if (thisW > 0) wowGrowth = 1;

    const thisMonth = lastDate.getMonth();
    const thisYear  = lastDate.getFullYear();
    let thisMo = 0, prevMo = 0;
    for (const [dk, v] of Object.entries(APP.postByDate)) {
      const d = new Date(dk);
      if (d.getFullYear() === thisYear && d.getMonth() === thisMonth) thisMo += v.total;
      if (d.getFullYear() === thisYear && d.getMonth() === thisMonth - 1) prevMo += v.total;
      if (thisMonth === 0 && d.getFullYear() === thisYear - 1 && d.getMonth() === 11) prevMo += v.total;
    }
    if (prevMo > 0) momGrowth = (thisMo - prevMo) / prevMo;
  }

  function growthBadge(pct) {
    if (pct === null) return '<span class="kpi-badge badge-flat">—</span>';
    const cls = pct > 0 ? 'badge-up' : pct < 0 ? 'badge-down' : 'badge-flat';
    const arrow = pct > 0 ? '▲' : pct < 0 ? '▼' : '=';
    const display = Math.abs(pct * 100) > 999 ? '>999%' : Math.abs(pct * 100).toFixed(1) + '%';
    return `<span class="kpi-badge ${cls}">${arrow} ${display}</span>`;
  }

  const cpa = totalApps > 0 ? totalSpend / totalApps : null;

  const cards = [
    { label: 'Total Jobs Posted',     value: fmtNum(totalJobs),  sub: 'Across all boards',        color: 'c-navy',    icon: '📋' },
    { label: 'Total Job Views',       value: fmtNum(totalViews), sub: 'Impressions / clicks',     color: 'c-sky',     icon: '👁️' },
    { label: 'Total Applications',    value: fmtNum(totalApps),  sub: 'Applied candidates',       color: 'c-teal',    icon: '✅' },
    { label: 'Resumes Viewed',        value: fmtNum(totalSrch),  sub: 'Sourced this period',      color: 'c-purple',  icon: '🔍' },
    { label: 'Most Utilized Board',   value: topPost   ? topPost[0]   : '—', sub: topPost   ? `${fmtNum(topPost[1].jobs)} jobs posted`   : '', color: 'c-success', icon: '🏆' },
    { label: 'Least Utilized Board',  value: leastPost ? leastPost[0] : '—', sub: leastPost ? `${fmtNum(leastPost[1].jobs)} jobs posted` : '', color: 'c-warn',    icon: '⚠️' },
    { label: 'Week-over-Week Growth', value: wowGrowth !== null ? (Math.abs(wowGrowth*100)>999 ? '>999%' : (wowGrowth*100).toFixed(1)+'%') : '—', sub: growthBadge(wowGrowth) + ' vs prior week', color: wowGrowth > 0 ? 'c-success' : 'c-danger', icon: '📈' },
    { label: 'Active Recruiters',     value: fmtNum(activeRecs), sub: 'With posting or search activity', color: 'c-orange', icon: '👤' },
  ];

  if (totalSpend > 0) {
    cards.push({ label: 'Total Ad Spend', value: fmtDol(totalSpend), sub: cpa ? `CPA: ${fmtDol(cpa)}` : '', color: 'c-navy', icon: '💰' });
  }

  document.getElementById('kpi-grid').innerHTML = cards.map(c => `
    <div class="kpi-card ${c.color}">
      <div class="kpi-label">${c.icon} ${c.label}</div>
      <div class="kpi-value">${c.value}</div>
      <div class="kpi-sub">${c.sub}</div>
    </div>
  `).join('');

  // Period label
  const dates = APP.fPostRecs.map(r => r.date).concat(APP.fSrchRecs.map(r => r.date)).filter(Boolean);
  if (dates.length > 0) {
    const minD = new Date(Math.min(...dates));
    const maxD = new Date(Math.max(...dates));
    document.getElementById('kpi-lbl').textContent = `${dateFmt(minD,'Mon D')} – ${dateFmt(maxD,'Mon D')}, ${maxD.getFullYear()}`;
  }
}

// ── CHARTS ────────────────────────────────────────────────
function mkChart(id, type, data, opts = {}) {
  const ctx = document.getElementById(id);
  if (!ctx) return;
  if (APP.charts[id]) { APP.charts[id].destroy(); }
  APP.charts[id] = new Chart(ctx.getContext('2d'), {
    type,
    data,
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: 'bottom', labels: { boxWidth: 12, font: { size: 11 } } },
        tooltip: { mode: 'index', intersect: false },
      },
      ...opts,
    },
  });
}

function renderCharts() {
  renderPostingBarChart();
  renderPostingPieChart();
  renderPostingTrendChart();
  renderPostingMetricsChart();
  renderSearchBarChart();
  renderSearchPieChart();
  renderSearchTrendChart();
  renderInmailChart();
  renderRecruiterPostChart();
  renderRecruiterSrchChart();
}

function renderPostingBarChart() {
  const boards = Object.keys(APP.boardPost);
  mkChart('chart-post-bar', 'bar', {
    labels: boards,
    datasets: [{
      label: 'Jobs Posted',
      data: boards.map(b => APP.boardPost[b].jobs),
      backgroundColor: boards.map(b => boardColor(b) + 'CC'),
      borderColor:     boards.map(b => boardColor(b)),
      borderWidth: 1, borderRadius: 4,
    }],
  }, {
    plugins: { legend: { display: false } },
    scales: {
      x: { grid: { display: false } },
      y: { beginAtZero: true, ticks: { stepSize: 1 } },
    },
  });
}

function renderPostingPieChart() {
  const boards = Object.keys(APP.boardPost);
  if (!boards.length) return;
  mkChart('chart-post-pie', 'doughnut', {
    labels: boards,
    datasets: [{
      data: boards.map(b => APP.boardPost[b].jobs),
      backgroundColor: boards.map(b => boardColor(b) + 'CC'),
      hoverOffset: 6,
    }],
  }, {
    plugins: { legend: { position: 'right', labels: { boxWidth: 10, font: { size: 11 } } } },
    cutout: '62%',
  });
}

function buildTrendData(byDate, mode, limit = 30) {
  const keys = Object.keys(byDate).sort();
  if (!keys.length) return { labels: [], datasets: [] };

  // Group by mode
  const grouped = {};
  const allBoards = new Set();
  keys.forEach(dk => {
    const group = mode === 'weekly' ? isoWeek(new Date(dk)) : dk;
    if (!grouped[group]) grouped[group] = { total: 0, byBoard: {} };
    grouped[group].total += byDate[dk].total;
    Object.entries(byDate[dk].byBoard || {}).forEach(([b, v]) => {
      grouped[group].byBoard[b] = (grouped[group].byBoard[b] || 0) + v;
      allBoards.add(b);
    });
  });

  const sortedGroups = Object.keys(grouped).sort().slice(-limit);
  const labels = sortedGroups.map(g => mode === 'weekly' ? g : dateFmt(new Date(g), 'Mon D'));

  // One dataset per board
  const boards = [...allBoards];
  const datasets = boards.map(b => ({
    label: b,
    data: sortedGroups.map(g => grouped[g].byBoard[b] || 0),
    borderColor: boardColor(b),
    backgroundColor: boardColor(b) + '22',
    tension: 0.3, fill: true, pointRadius: 3,
  }));

  // Total line
  datasets.push({
    label: 'Total',
    data: sortedGroups.map(g => grouped[g].total),
    borderColor: '#1B2A4A', borderWidth: 2,
    backgroundColor: 'transparent',
    tension: 0.3, pointRadius: 3, borderDash: [4, 3],
  });

  return { labels, datasets };
}

function renderPostingTrendChart() {
  const tdata = buildTrendData(APP.postByDate, APP.trendMode.post);
  mkChart('chart-post-trend', 'line', tdata, {
    scales: { x: { grid: { display: false } }, y: { beginAtZero: true } },
    plugins: { legend: { position: 'bottom' } },
  });
}

function renderPostingMetricsChart() {
  const boards = Object.keys(APP.boardPost);
  mkChart('chart-post-metrics', 'bar', {
    labels: boards,
    datasets: [
      { label: 'Views', data: boards.map(b => APP.boardPost[b].views), backgroundColor: '#0277BD88', borderRadius: 3 },
      { label: 'Applications', data: boards.map(b => APP.boardPost[b].applications), backgroundColor: '#00897B88', borderRadius: 3 },
    ],
  }, {
    scales: { x: { grid: { display: false } }, y: { beginAtZero: true } },
    plugins: { legend: { position: 'bottom' } },
  });
}

function renderSearchBarChart() {
  const boards = Object.keys(APP.boardSrch);
  if (!boards.length) return;
  mkChart('chart-srch-bar', 'bar', {
    labels: boards,
    datasets: [
      { label: 'Searches / Sends', data: boards.map(b => APP.boardSrch[b].searches), backgroundColor: '#0077B588', borderRadius: 3 },
      { label: 'Views / Fetches',  data: boards.map(b => APP.boardSrch[b].views),    backgroundColor: '#00897B88', borderRadius: 3 },
      { label: 'Contacts',         data: boards.map(b => APP.boardSrch[b].contacts), backgroundColor: '#7B2FBE88', borderRadius: 3 },
    ],
  }, {
    scales: { x: { grid: { display: false } }, y: { beginAtZero: true } },
    plugins: { legend: { position: 'bottom' } },
  });
}

function renderSearchPieChart() {
  const boards = Object.keys(APP.boardSrch);
  if (!boards.length) return;
  mkChart('chart-srch-pie', 'doughnut', {
    labels: boards,
    datasets: [{
      data: boards.map(b => APP.boardSrch[b].views),
      backgroundColor: boards.map(b => boardColor(b) + 'CC'),
      hoverOffset: 6,
    }],
  }, {
    plugins: { legend: { position: 'right', labels: { boxWidth: 10, font: { size: 11 } } } },
    cutout: '62%',
  });
}

function renderSearchTrendChart() {
  const tdata = buildTrendData(APP.srchByDate, APP.trendMode.srch);
  mkChart('chart-srch-trend', 'line', tdata, {
    scales: { x: { grid: { display: false } }, y: { beginAtZero: true } },
    plugins: { legend: { position: 'bottom' } },
  });
}

function renderInmailChart() {
  // LinkedIn InMail recruiters from srchRecs
  const liRecs = APP.fSrchRecs.filter(r => r.board === 'LinkedIn' && r.recruiter);
  if (!liRecs.length) { document.getElementById('chart-inmail').closest('.chart-card.wide').style.display = 'none'; return; }
  document.getElementById('chart-inmail').closest('.chart-card.wide').style.display = '';
  // Group by recruiter
  const byRec = {};
  liRecs.forEach(r => {
    if (!byRec[r.recruiter]) byRec[r.recruiter] = { sends: 0, responses: 0, contacts: 0 };
    byRec[r.recruiter].sends     += r.searches || 0;
    byRec[r.recruiter].responses += r.responses|| 0;
    byRec[r.recruiter].contacts  += r.contacts || 0;
  });
  const recs = Object.keys(byRec).sort((a,b) => byRec[b].sends - byRec[a].sends);
  mkChart('chart-inmail', 'bar', {
    labels: recs,
    datasets: [
      { label: 'InMail Sends',    data: recs.map(r => byRec[r].sends),     backgroundColor: '#0077B588', borderRadius: 3 },
      { label: 'Responses',       data: recs.map(r => byRec[r].responses), backgroundColor: '#00897B88', borderRadius: 3 },
      { label: 'Accepts',         data: recs.map(r => byRec[r].contacts),  backgroundColor: '#F59E0B88', borderRadius: 3 },
    ],
  }, {
    scales: { x: { grid: { display: false } }, y: { beginAtZero: true } },
    plugins: { legend: { position: 'bottom' } },
  });
}

function renderRecruiterPostChart() {
  const SKIP = new Set(['system','total','grand total']);
  const recs = Object.entries(APP.recPost)
    .filter(([n]) => n.length > 1 && !SKIP.has(n.toLowerCase().trim()))
    .sort((a,b) => b[1].jobs - a[1].jobs)
    .slice(0, 20);
  if (!recs.length) return;
  mkChart('chart-rec-post', 'bar', {
    labels: recs.map(([n]) => n),
    datasets: [{
      label: 'Jobs Posted',
      data: recs.map(([,v]) => v.jobs),
      backgroundColor: '#1B2A4A88', borderRadius: 4,
    }],
  }, {
    indexAxis: 'y',
    plugins: { legend: { display: false } },
    scales: { x: { beginAtZero: true }, y: { grid: { display: false } } },
  });
}

function renderRecruiterSrchChart() {
  const SKIP = new Set(['system','total','grand total']);
  const recs = Object.entries(APP.recSrch)
    .filter(([n]) => n.length > 1 && !SKIP.has(n.toLowerCase().trim()))
    .sort((a,b) => b[1].views - a[1].views)
    .slice(0, 15);
  if (!recs.length) return;
  mkChart('chart-rec-srch', 'bar', {
    labels: recs.map(([n]) => n),
    datasets: [{
      label: 'Resumes Viewed',
      data: recs.map(([,v]) => v.views),
      backgroundColor: '#00897B88', borderRadius: 4,
    }],
  }, {
    indexAxis: 'y',
    plugins: { legend: { display: false } },
    scales: { x: { beginAtZero: true }, y: { grid: { display: false } } },
  });
}

// ── TABLES ────────────────────────────────────────────────
function renderTables() {
  renderPostingsTable();
  renderSearchesTable();
  renderRecruitersTable();
  renderRecruiterPostMatrix();
  renderRecruiterSrchMatrix();
}

function renderPostingsTable() {
  const tbody = document.getElementById('dtbl-postings-body');
  if (!tbody) return;
  const boards = Object.entries(APP.boardPost).sort((a,b) => b[1].jobs - a[1].jobs);
  tbody.innerHTML = boards.map(([board, d], i) => {
    const avgViews = d.jobs > 0 ? (d.views / d.jobs).toFixed(1) : '—';
    const convRate = d.views > 0 ? (d.applications / d.views * 100).toFixed(1) + '%' : '—';
    const topVert  = [...d.verticals][0] || '—';
    return `<tr>
      <td><span class="board-dot" style="background:${boardColor(board)}"></span>${board}</td>
      <td class="r"><strong>${fmtNum(d.jobs)}</strong></td>
      <td class="r">${fmtNum(d.views)}</td>
      <td class="r">${fmtNum(d.applications)}</td>
      <td class="r">${avgViews}</td>
      <td class="r">${convRate}</td>
      <td>${topVert}</td>
      <td class="r">${d.spend > 0 ? fmtDol(d.spend) : '—'}</td>
    </tr>`;
  }).join('');
}

function renderSearchesTable() {
  const tbody = document.getElementById('dtbl-searches-body');
  if (!tbody) return;
  const boards = Object.entries(APP.boardSrch).sort((a,b) => b[1].views - a[1].views);
  tbody.innerHTML = boards.map(([board, d]) => {
    const respRate = d.searches > 0 ? (d.responses / d.searches * 100).toFixed(1) + '%' : '—';
    const recsList = [...d.recruiters].slice(0, 3).join(', ') + (d.recruiters.size > 3 ? ` +${d.recruiters.size-3}` : '');
    return `<tr>
      <td><span class="board-dot" style="background:${boardColor(board)}"></span>${board}</td>
      <td class="r">${fmtNum(d.searches)}</td>
      <td class="r"><strong>${fmtNum(d.views)}</strong></td>
      <td class="r">${fmtNum(d.contacts)}</td>
      <td class="r">${fmtNum(d.responses)}</td>
      <td class="r">${respRate}</td>
      <td>${recsList || '—'}</td>
    </tr>`;
  }).join('');
}

function renderRecruitersTable() {
  const tbody = document.getElementById('dtbl-recruiters-body');
  if (!tbody) return;
  const SKIP = new Set(['system','total','grand total','','—']);
  const allRecs = new Set([...Object.keys(APP.recPost), ...Object.keys(APP.recSrch)]);
  const rows = [...allRecs]
    .filter(r => r && r.length > 1 && !SKIP.has(r.toLowerCase().trim()))
    .map(r => ({
      name: r,
      vertical: APP.recPost[r]?.vertical || '',
      jobs: APP.recPost[r]?.jobs || 0,
      views: APP.recPost[r]?.views || 0,
      applications: APP.recPost[r]?.applications || 0,
      srchViews: APP.recSrch[r]?.views || 0,
      boards: new Set([...(APP.recPost[r]?.boards || []), ...(APP.recSrch[r]?.boards || [])]),
    }))
    .sort((a,b) => (b.jobs + b.srchViews) - (a.jobs + a.srchViews));

  tbody.innerHTML = rows.map(r => `<tr>
    <td><strong>${r.name}</strong></td>
    <td>${r.vertical || '—'}</td>
    <td class="r">${fmtNum(r.jobs)}</td>
    <td class="r">${fmtNum(r.views)}</td>
    <td class="r">${fmtNum(r.applications)}</td>
    <td class="r">${fmtNum(r.srchViews)}</td>
    <td>${[...r.boards].join(', ') || '—'}</td>
  </tr>`).join('');
}

// ── RECRUITER × BOARD MATRIX ──────────────────────────────
function hexToRgba(hex, alpha) {
  const r = parseInt(hex.slice(1,3),16);
  const g = parseInt(hex.slice(3,5),16);
  const b = parseInt(hex.slice(5,7),16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function switchRecTab(btn, paneId) {
  document.querySelectorAll('.rec-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('.rec-tab-pane').forEach(p => p.style.display = 'none');
  btn.classList.add('active');
  const pane = document.getElementById(paneId);
  if (pane) pane.style.display = '';
}

function renderRecruiterPostMatrix() {
  const wrap = document.getElementById('rec-post-matrix-wrap');
  if (!wrap) return;
  const SKIP = new Set(['system','total','grand total','','—']);

  const recs = Object.keys(APP.recPostByBoard)
    .filter(r => r && r.length > 1 && !SKIP.has(r.toLowerCase().trim()));
  if (!recs.length) { wrap.innerHTML = '<p style="padding:1rem;color:#888">No posting data by board available.</p>'; return; }

  // Get all boards, sorted by total jobs desc
  const boardSet = new Set();
  for (const r of recs) Object.keys(APP.recPostByBoard[r]).forEach(b => boardSet.add(b));
  const boards = [...boardSet].sort((a,b) => {
    const ta = recs.reduce((s,r) => s + (APP.recPostByBoard[r][a]?.jobs||0), 0);
    const tb = recs.reduce((s,r) => s + (APP.recPostByBoard[r][b]?.jobs||0), 0);
    return tb - ta;
  });

  // Sort recruiters by total jobs desc
  recs.sort((a,b) => {
    const ta = Object.values(APP.recPostByBoard[a]).reduce((s,v) => s+v.jobs, 0);
    const tb = Object.values(APP.recPostByBoard[b]).reduce((s,v) => s+v.jobs, 0);
    return tb - ta;
  });

  // Max per board (for heat-map scaling)
  const boardMax = {};
  for (const bd of boards) boardMax[bd] = Math.max(...recs.map(r => APP.recPostByBoard[r][bd]?.jobs||0));

  let html = '<table class="dtbl matrix-tbl"><thead><tr>';
  html += '<th>Recruiter</th><th class="r">Total</th>';
  for (const bd of boards) {
    const c = boardColor(bd);
    html += `<th class="r matrix-col-hdr" style="border-bottom:3px solid ${c}"><span class="board-dot" style="background:${c}"></span>${bd}</th>`;
  }
  html += '</tr></thead><tbody>';

  for (const r of recs) {
    const total = Object.values(APP.recPostByBoard[r]).reduce((s,v) => s+v.jobs, 0);
    html += `<tr><td class="rec-name"><strong>${r}</strong></td><td class="r total-cell"><strong>${total}</strong></td>`;
    for (const bd of boards) {
      const d = APP.recPostByBoard[r][bd];
      const val = d?.jobs || 0;
      const alpha = boardMax[bd] > 0 ? 0.12 + (val/boardMax[bd])*0.55 : 0;
      const bg = val > 0 ? hexToRgba(boardColor(bd), alpha) : '#f6f7f9';
      const tip = val > 0 ? `title="${val} job${val>1?'s':''} | ${fmtNum(d.views)} views | ${fmtNum(d.applications)} apps"` : '';
      html += `<td class="r matrix-cell" style="background:${bg}" ${tip}>${val > 0 ? `<strong>${val}</strong>` : '<span class="zero">—</span>'}</td>`;
    }
    html += '</tr>';
  }

  // Totals row
  const grandTotal = recs.reduce((s,r) => s + Object.values(APP.recPostByBoard[r]).reduce((ss,v) => ss+v.jobs, 0), 0);
  html += `<tr class="matrix-total"><td><strong>Total</strong></td><td class="r"><strong>${grandTotal}</strong></td>`;
  for (const bd of boards) {
    const ct = recs.reduce((s,r) => s+(APP.recPostByBoard[r][bd]?.jobs||0), 0);
    html += `<td class="r"><strong>${ct}</strong></td>`;
  }
  html += '</tr></tbody></table>';
  wrap.innerHTML = html;
}

function renderRecruiterSrchMatrix() {
  const wrap = document.getElementById('rec-srch-matrix-wrap');
  if (!wrap) return;
  const SKIP = new Set(['system','total','grand total','','—']);

  const recs = Object.keys(APP.recSrchByBoard)
    .filter(r => r && r.length > 1 && !SKIP.has(r.toLowerCase().trim()));
  if (!recs.length) { wrap.innerHTML = '<p style="padding:1rem;color:#888">No resume search data by board available.</p>'; return; }

  const boardSet = new Set();
  for (const r of recs) Object.keys(APP.recSrchByBoard[r]).forEach(b => boardSet.add(b));
  const boards = [...boardSet].sort((a,b) => {
    const ta = recs.reduce((s,r) => s+(APP.recSrchByBoard[r][a]?.views||0), 0);
    const tb = recs.reduce((s,r) => s+(APP.recSrchByBoard[r][b]?.views||0), 0);
    return tb - ta;
  });

  recs.sort((a,b) => {
    const ta = Object.values(APP.recSrchByBoard[a]).reduce((s,v) => s+v.views, 0);
    const tb = Object.values(APP.recSrchByBoard[b]).reduce((s,v) => s+v.views, 0);
    return tb - ta;
  });

  const boardMax = {};
  for (const bd of boards) boardMax[bd] = Math.max(...recs.map(r => APP.recSrchByBoard[r][bd]?.views||0));

  let html = '<table class="dtbl matrix-tbl"><thead><tr>';
  html += '<th>Recruiter</th><th class="r">Total Views</th>';
  for (const bd of boards) {
    const c = boardColor(bd);
    html += `<th class="r matrix-col-hdr" style="border-bottom:3px solid ${c}"><span class="board-dot" style="background:${c}"></span>${bd}</th>`;
  }
  html += '</tr></thead><tbody>';

  for (const r of recs) {
    const total = Object.values(APP.recSrchByBoard[r]).reduce((s,v) => s+v.views, 0);
    html += `<tr><td class="rec-name"><strong>${r}</strong></td><td class="r total-cell"><strong>${fmtNum(total)}</strong></td>`;
    for (const bd of boards) {
      const val = APP.recSrchByBoard[r][bd]?.views || 0;
      const alpha = boardMax[bd] > 0 ? 0.12 + (val/boardMax[bd])*0.55 : 0;
      const bg = val > 0 ? hexToRgba(boardColor(bd), alpha) : '#f6f7f9';
      const tip = val > 0 ? `title="${fmtNum(val)} resumes viewed"` : '';
      html += `<td class="r matrix-cell" style="background:${bg}" ${tip}>${val > 0 ? `<strong>${fmtNum(val)}</strong>` : '<span class="zero">—</span>'}</td>`;
    }
    html += '</tr>';
  }

  const grandTotal = recs.reduce((s,r) => s+Object.values(APP.recSrchByBoard[r]).reduce((ss,v) => ss+v.views, 0), 0);
  html += `<tr class="matrix-total"><td><strong>Total</strong></td><td class="r"><strong>${fmtNum(grandTotal)}</strong></td>`;
  for (const bd of boards) {
    const ct = recs.reduce((s,r) => s+(APP.recSrchByBoard[r][bd]?.views||0), 0);
    html += `<td class="r"><strong>${fmtNum(ct)}</strong></td>`;
  }
  html += '</tr></tbody></table>';
  wrap.innerHTML = html;
}

// ── INSIGHTS ──────────────────────────────────────────────
function renderInsights() {
  const grid = document.getElementById('insights-grid');
  if (!grid) return;
  const insights = [];

  const bPost = Object.entries(APP.boardPost).sort((a,b) => b[1].jobs - a[1].jobs);
  const bSrch = Object.entries(APP.boardSrch).sort((a,b) => b[1].views - a[1].views);

  // Top posting board
  if (bPost.length > 0) {
    const [top, d] = bPost[0];
    const conv = d.views > 0 ? (d.applications/d.views*100).toFixed(1) : 0;
    insights.push({ type: 'top', icon: '🏆', title: 'Top Posting Board', metric: top,
      body: `<strong>${top}</strong> leads job posting with <strong>${fmtNum(d.jobs)}</strong> jobs, <strong>${fmtNum(d.views)}</strong> views, and <strong>${fmtNum(d.applications)}</strong> applications (${conv}% conversion).` });
  }

  // Least posting board
  if (bPost.length > 1) {
    const [bot, d] = bPost[bPost.length - 1];
    insights.push({ type: 'low', icon: '⚠️', title: 'Underutilized for Postings', metric: bot,
      body: `<strong>${bot}</strong> has the fewest postings this period with only <strong>${fmtNum(d.jobs)}</strong> jobs. Consider increasing utilization or reviewing ROI.` });
  }

  // Top resume source
  if (bSrch.length > 0) {
    const [top, d] = bSrch[0];
    insights.push({ type: 'top', icon: '📚', title: 'Top Resume Source', metric: top,
      body: `<strong>${top}</strong> is the leading resume source with <strong>${fmtNum(d.views)}</strong> resumes viewed and <strong>${d.recruiters.size}</strong> active recruiters.` });
  }

  // LinkedIn InMail performance
  const liSrch = APP.boardSrch['LinkedIn'];
  if (liSrch && liSrch.searches > 0) {
    const rRate = (liSrch.responses / liSrch.searches * 100).toFixed(1);
    const accRate = (liSrch.contacts / liSrch.searches * 100).toFixed(1);
    const tier = liSrch.responses/liSrch.searches > 0.25 ? 'strong' : liSrch.responses/liSrch.searches > 0.1 ? 'moderate' : 'low';
    insights.push({ type: 'trend', icon: '💼', title: 'LinkedIn InMail Performance', metric: rRate + '% response rate',
      body: `<strong>${fmtNum(liSrch.searches)}</strong> InMails sent this week with a <strong>${rRate}%</strong> response rate and <strong>${accRate}%</strong> accept rate. Engagement is <strong>${tier}</strong> vs industry avg ~25%.` });
  }

  // Indeed ad spend efficiency
  const indeedPost = APP.boardPost['Indeed'];
  if (indeedPost && indeedPost.spend > 0) {
    const cpa = indeedPost.applications > 0 ? indeedPost.spend / indeedPost.applications : null;
    const cpc = indeedPost.views > 0 ? indeedPost.spend / indeedPost.views : null;
    insights.push({ type: 'rec', icon: '💰', title: 'Indeed Ad Spend', metric: fmtDol(indeedPost.spend),
      body: `Indeed spend: <strong>${fmtDol(indeedPost.spend)}</strong>. ${cpa ? `Cost per apply: <strong>${fmtDol(cpa)}</strong>.` : ''} ${cpc ? `Cost per click: <strong>${fmtDol(cpc)}</strong>.` : ''} Review job-level performance for budget optimization.` });
  }

  // Most active recruiter
  const topRec = Object.entries(APP.recPost).sort((a,b) => b[1].jobs - a[1].jobs)[0];
  if (topRec) {
    const [name, d] = topRec;
    const srch = APP.recSrch[name];
    insights.push({ type: 'growth', icon: '🌟', title: 'Most Active Recruiter', metric: name,
      body: `<strong>${name}</strong> posted <strong>${fmtNum(d.jobs)}</strong> jobs across ${d.boards.size} board(s) with <strong>${fmtNum(d.applications)}</strong> applications received. ${srch ? `Also viewed <strong>${fmtNum(srch.views)}</strong> resumes.` : ''}` });
  }

  // Vivian candidate pipeline
  const vivSrch = APP.boardSrch['Vivian'];
  const vivPost  = APP.boardPost['Vivian'];
  if (vivPost || vivSrch) {
    const jobs  = vivPost ? vivPost.jobs : 0;
    const cands = vivSrch ? vivSrch.views : 0;
    insights.push({ type: 'trend', icon: '🏥', title: 'Vivian Healthcare Pipeline', metric: `${fmtNum(jobs)} jobs`,
      body: `Vivian shows <strong>${fmtNum(jobs)}</strong> active healthcare job postings with <strong>${fmtNum(cands)}</strong> candidate leads (inbound + proposals) in the pipeline.` });
  }

  // Monster usage
  if (APP.meta?.monsterCreditsUsed > 0) {
    const pct = (APP.meta.monsterCreditsUsed / APP.meta.monsterCreditsRcvd * 100).toFixed(1);
    insights.push({ type: 'alert', icon: '👾', title: 'Monster Credits Utilization', metric: pct + '% used',
      body: `Monster credits: <strong>${fmtNum(APP.meta.monsterCreditsUsed)}</strong> of <strong>${fmtNum(APP.meta.monsterCreditsRcvd)}</strong> used (${pct}%). <strong>${fmtNum(APP.meta.monsterCreditsAvail)}</strong> credits remaining.` });
  }

  // Signal Hire
  const shSrch = APP.boardSrch['Signal Hire'];
  if (shSrch) {
    insights.push({ type: 'rec', icon: '📡', title: 'Signal Hire Unlocks', metric: fmtNum(shSrch.contacts) + ' unlocks',
      body: `Signal Hire team unlocked <strong>${fmtNum(shSrch.contacts)}</strong> contact records this week across ${shSrch.recruiters.size} account(s). Use for direct outreach.` });
  }

  // Posting distribution recommendation
  if (bPost.length > 2) {
    const topPct = bPost[0][1].jobs / APP.fPostRecs.length;
    if (topPct > 0.6) {
      insights.push({ type: 'rec', icon: '📊', title: 'Posting Concentration Risk', metric: (topPct*100).toFixed(0) + '% on one board',
        body: `<strong>${(topPct*100).toFixed(0)}%</strong> of postings are on a single board. Diversifying across ${bPost.slice(1,4).map(([b])=>b).join(', ')} can improve reach and reduce dependency risk.` });
    }
  }

  grid.innerHTML = insights.map(ins => `
    <div class="insight-card type-${ins.type}">
      <div class="ins-header">
        <span class="ins-icon">${ins.icon}</span>
        <span class="ins-title">${ins.title}</span>
      </div>
      <div class="ins-metric">${ins.metric}</div>
      <div class="ins-body">${ins.body}</div>
    </div>
  `).join('');
}

// ── TREND MODE TOGGLE ─────────────────────────────────────
function setMode(btn) {
  const series = btn.dataset.series;
  const mode   = btn.dataset.mode;
  btn.closest('.toggle-grp').querySelectorAll('.chip').forEach(c => c.classList.remove('active'));
  btn.classList.add('active');
  APP.trendMode[series] = mode;
  if (series === 'post') renderPostingTrendChart();
  else renderSearchTrendChart();
}

// ── VIEW SWITCHER ─────────────────────────────────────────
function switchView(view) {
  const isRec = view === 'recruiter';
  document.getElementById('vtog-exec').classList.toggle('active', !isRec);
  document.getElementById('vtog-rec').classList.toggle('active', isRec);
  // Toggle executive sections
  ['sec-kpis','sec-postings','sec-searches','sec-boards','sec-recruiters','sec-insights'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = isRec ? 'none' : '';
  });
  document.getElementById('dash-nav').style.display = isRec ? 'none' : '';
  document.getElementById('filter-bar').style.display = isRec ? 'none' : '';
  document.getElementById('recruiter-view').style.display = isRec ? 'block' : 'none';
  if (isRec) populateRvPicker();
}

function populateRvPicker() {
  const sel = document.getElementById('rv-recruiter');
  if (!sel) return;
  const SKIP = new Set(['system','total','grand total','','—']);
  const allRecs = [...new Set([
    ...Object.keys(APP.recPost),
    ...Object.keys(APP.recSrch)
  ])].filter(r => r && r.length > 1 && !SKIP.has(r.toLowerCase().trim())).sort();
  const cur = sel.value;
  sel.innerHTML = '<option value="">— Select a recruiter —</option>' +
    allRecs.map(r => `<option value="${r}"${r===cur?' selected':''}>${r}</option>`).join('');
  if (cur && allRecs.includes(cur)) renderRecruiterView();
}

function renderRecruiterView() {
  const name = document.getElementById('rv-recruiter')?.value || '';
  if (!name) {
    document.getElementById('rv-kpis').innerHTML = '<p class="rv-empty">Select a recruiter above to load their dashboard.</p>';
    ['rv-portal-body','rv-jobs-body','rv-srch-body'].forEach(id => {
      const el = document.getElementById(id);
      if (el) el.innerHTML = '';
    });
    return;
  }
  renderRvKpis(name);
  renderRvPortalTable(name);
  renderRvJobsList(name);
  renderRvSearchActivity(name);
}

function renderRvKpis(name) {
  const p = APP.recPost[name]  || { jobs:0, views:0, applications:0, boards: new Set() };
  const s = APP.recSrch[name] || { views:0, boards: new Set() };
  const pbb = APP.recPostByBoard[name] || {};
  const conv = p.views > 0 ? (p.applications/p.views*100).toFixed(1)+'%' : '0%';
  const cards = [
    { label:'Jobs Posted',     val: fmtNum(p.jobs),         icon:'📋', cls:'c-navy'    },
    { label:'Total Views',     val: fmtNum(p.views),        icon:'👁',  cls:'c-sky'     },
    { label:'Applications',    val: fmtNum(p.applications), icon:'📩', cls:'c-teal'    },
    { label:'Conv. Rate',      val: conv,                   icon:'📈', cls:'c-success'  },
    { label:'Resumes Viewed',  val: fmtNum(s.views),        icon:'🔍', cls:'c-purple'  },
    { label:'Portals Active',  val: Object.keys(pbb).length, icon:'🏢', cls:'c-orange' },
  ];
  document.getElementById('rv-kpis').innerHTML = cards.map(c => `
    <div class="kpi-card ${c.cls}">
      <div class="kpi-icon">${c.icon}</div>
      <div class="kpi-body">
        <div class="kpi-val">${c.val}</div>
        <div class="kpi-lbl">${c.label}</div>
      </div>
    </div>`).join('');
}

function renderRvPortalTable(name) {
  const tbody = document.getElementById('rv-portal-body');
  if (!tbody) return;
  const alloc  = loadAllocations();
  const pbb    = APP.recPostByBoard[name] || {};
  const myPosts = APP.fPostRecs.filter(r => r.recruiter === name);
  const boards  = Object.keys(pbb).sort((a,b) => pbb[b].jobs - pbb[a].jobs);

  if (!boards.length) {
    tbody.innerHTML = '<tr><td colspan="9" class="empty-cell">No posting data found for this recruiter.</td></tr>';
    return;
  }
  tbody.innerHTML = boards.map(bd => {
    const d  = pbb[bd];
    const al = alloc[bd] || null;
    const activePosts  = myPosts.filter(r => r.board === bd && r.active === true).length;
    // If active field not parsed, use total (treat all as posted)
    const hasActive    = myPosts.some(r => r.board === bd && r.active !== undefined);
    const activeCount  = hasActive ? activePosts : d.jobs;
    const closedCount  = hasActive ? (d.jobs - activePosts) : 0;
    const available    = al !== null ? Math.max(0, al - d.jobs) : null;
    const conv         = d.views > 0 ? (d.applications/d.views*100).toFixed(1)+'%' : '—';
    const color        = boardColor(bd);
    const availHtml    = available === null ? '<span class="zero">—</span>'
                        : available > 0
                          ? `<span class="badge badge-blue">${available}</span>`
                          : `<span class="badge badge-warn">0</span>`;
    return `<tr>
      <td><span class="board-dot" style="background:${color}"></span><strong>${bd}</strong></td>
      <td class="r">${al ? `<strong>${al}</strong>` : '<span class="zero">—</span>'}</td>
      <td class="r"><strong>${d.jobs}</strong></td>
      <td class="r"><span class="badge badge-green">${activeCount}</span></td>
      <td class="r">${closedCount > 0 ? `<span class="badge badge-grey">${closedCount}</span>` : '<span class="zero">—</span>'}</td>
      <td class="r">${availHtml}</td>
      <td class="r">${fmtNum(d.views)}</td>
      <td class="r">${fmtNum(d.applications)}</td>
      <td class="r">${conv}</td>
    </tr>`;
  }).join('');
}

function renderRvJobsList(name) {
  const tbody = document.getElementById('rv-jobs-body');
  if (!tbody) return;
  const myPosts = APP.fPostRecs
    .filter(r => r.recruiter === name)
    .sort((a,b) => (b.date||0) - (a.date||0));

  if (!myPosts.length) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-cell">No job postings found for this recruiter.</td></tr>';
    return;
  }
  tbody.innerHTML = myPosts.map(r => {
    const conv  = r.views > 0 ? (r.applications/r.views*100).toFixed(1)+'%' : '—';
    const status = r.active === true  ? '<span class="badge badge-green">Active</span>'
                 : r.active === false ? '<span class="badge badge-grey">Closed</span>'
                 : '<span class="badge badge-blue">Posted</span>';
    return `<tr>
      <td>${r.jobTitle || '—'}</td>
      <td><span class="board-dot" style="background:${boardColor(r.board)}"></span>${r.board}</td>
      <td>${r.vertical || '—'}</td>
      <td>${r.date ? dateFmt(r.date) : '—'}</td>
      <td class="r">${fmtNum(r.views)}</td>
      <td class="r">${fmtNum(r.applications)}</td>
      <td class="r">${conv}</td>
      <td>${status}</td>
    </tr>`;
  }).join('');
}

function renderRvSearchActivity(name) {
  const tbody = document.getElementById('rv-srch-body');
  if (!tbody) return;
  const sbb = APP.recSrchByBoard[name] || {};
  const rows = Object.entries(sbb).sort((a,b) => b[1].views - a[1].views);
  if (!rows.length) {
    tbody.innerHTML = '<tr><td colspan="2" class="empty-cell">No resume search data found for this recruiter.</td></tr>';
    return;
  }
  tbody.innerHTML = rows.map(([bd, d]) => `<tr>
    <td><span class="board-dot" style="background:${boardColor(bd)}"></span><strong>${bd}</strong></td>
    <td class="r">${fmtNum(d.views)}</td>
  </tr>`).join('');
}

// ── ALLOCATIONS ───────────────────────────────────────────
const ALLOC_KEY = 'hv_board_allocations';
function loadAllocations() {
  try { return JSON.parse(localStorage.getItem(ALLOC_KEY) || '{}'); } catch { return {}; }
}
function saveAllocations() {
  const alloc = {};
  document.querySelectorAll('.alloc-input').forEach(inp => {
    const v = parseInt(inp.value);
    if (!isNaN(v) && v > 0) alloc[inp.dataset.board] = v;
  });
  localStorage.setItem(ALLOC_KEY, JSON.stringify(alloc));
  closeAllocModal();
  renderRecruiterView();
}
function openAllocModal() {
  const alloc  = loadAllocations();
  const boards = [...new Set([...Object.keys(APP.boardPost), ...Object.keys(APP.boardSrch)])].filter(Boolean).sort();
  document.getElementById('alloc-inputs').innerHTML = boards.map(b => `
    <div class="alloc-row">
      <span class="board-dot" style="background:${boardColor(b)}"></span>
      <label class="alloc-label">${b}</label>
      <input class="alloc-input" type="number" min="1" placeholder="—"
             data-board="${b}" value="${alloc[b] || ''}">
      <span class="alloc-unit">slots</span>
    </div>`).join('');
  document.getElementById('alloc-modal').style.display = 'flex';
}
function closeAllocModal(e) {
  if (e && e.target.id !== 'alloc-modal') return;
  document.getElementById('alloc-modal').style.display = 'none';
}

// ── EXPORT ────────────────────────────────────────────────
function exportCSV(type) {
  closeMenu('export-menu');
  let rows = [], filename = 'export.csv';
  if (type === 'postings' || type === 'postings-tbl') {
    filename = 'job_postings.csv';
    rows = [['Board','Date','Recruiter','Vertical','Job Title','Position ID','Views','Applications','Spend']];
    APP.fPostRecs.forEach(r => rows.push([
      r.board, r.date ? dateFmt(r.date) : '', r.recruiter, r.vertical,
      r.jobTitle, r.positionId, r.views, r.applications, r.spend||0
    ]));
  } else if (type === 'searches' || type === 'searches-tbl') {
    filename = 'resume_searches.csv';
    rows = [['Board','Date','Recruiter','Candidate / Note','Source','Searches','Views','Contacts','Responses']];
    APP.fSrchRecs.forEach(r => rows.push([
      r.board, r.date ? dateFmt(r.date) : '', r.recruiter, r.candidate || '', r.source || '',
      r.searches||0, r.views||0, r.contacts||0, r.responses||0
    ]));
  } else if (type === 'recruiter' || type === 'recruiter-tbl') {
    filename = 'recruiter_summary.csv';
    rows = [['Recruiter','Vertical','Jobs Posted','Views','Applications','Resumes Viewed','Boards']];
    const allRecs = new Set([...Object.keys(APP.recPost), ...Object.keys(APP.recSrch)]);
    [...allRecs].filter(r => r && r.length > 1).forEach(r => {
      const p = APP.recPost[r] || {};
      const s = APP.recSrch[r] || {};
      const boards = new Set([...(p.boards||[]), ...(s.boards||[])]);
      rows.push([r, p.vertical||'', p.jobs||0, p.views||0, p.applications||0, s.views||0, [...boards].join('; ')]);
    });
  } else if (type === 'rv-portal') {
    filename = 'my_portal_summary.csv';
    const name = document.getElementById('rv-recruiter')?.value || '';
    const alloc = loadAllocations();
    const pbb = APP.recPostByBoard[name] || {};
    rows = [['Portal','Recruiter','Allocated Slots','Jobs Posted','Active','Closed','Slots Available','Views','Applications','Conv %']];
    const myPosts = APP.fPostRecs.filter(r => r.recruiter === name);
    Object.entries(pbb).sort((a,b)=>b[1].jobs-a[1].jobs).forEach(([bd,d]) => {
      const al = alloc[bd]||null;
      const hasActive = myPosts.some(r=>r.board===bd && r.active!==undefined);
      const activeCount = hasActive ? myPosts.filter(r=>r.board===bd&&r.active===true).length : d.jobs;
      const closedCount = hasActive ? d.jobs-activeCount : 0;
      const available = al!==null ? Math.max(0,al-d.jobs) : '';
      const conv = d.views>0?(d.applications/d.views*100).toFixed(1)+'%':'—';
      rows.push([bd,name,al||'',d.jobs,activeCount,closedCount,available,d.views,d.applications,conv]);
    });
  } else if (type === 'rv-jobs') {
    filename = 'my_job_postings.csv';
    const name = document.getElementById('rv-recruiter')?.value || '';
    rows = [['Job Title','Portal','Vertical','Date Posted','Views','Applications','Conv %','Status']];
    APP.fPostRecs.filter(r=>r.recruiter===name).sort((a,b)=>(b.date||0)-(a.date||0)).forEach(r => {
      const conv = r.views>0?(r.applications/r.views*100).toFixed(1)+'%':'—';
      const status = r.active===true?'Active':r.active===false?'Closed':'Posted';
      rows.push([r.jobTitle||'',r.board,r.vertical||'',r.date?dateFmt(r.date):'',r.views,r.applications,conv,status]);
    });
  } else if (type === 'rv-search') {
    filename = 'my_resume_activity.csv';
    const name = document.getElementById('rv-recruiter')?.value || '';
    rows = [['Source / Board','Resumes Viewed']];
    const sbb = APP.recSrchByBoard[name] || {};
    Object.entries(sbb).sort((a,b)=>b[1].views-a[1].views).forEach(([bd,d])=>rows.push([bd,d.views]));
  } else if (type === 'rec-post-matrix') {
    filename = 'recruiter_postings_by_board.csv';
    const SKIP = new Set(['system','total','grand total','','—']);
    const recs = Object.keys(APP.recPostByBoard).filter(r => r && r.length > 1 && !SKIP.has(r.toLowerCase().trim()));
    const boardSet = new Set();
    for (const r of recs) Object.keys(APP.recPostByBoard[r]).forEach(b => boardSet.add(b));
    const boards = [...boardSet].sort((a,b) => {
      return recs.reduce((s,r)=>s+(APP.recPostByBoard[r][b]?.jobs||0),0) - recs.reduce((s,r)=>s+(APP.recPostByBoard[r][a]?.jobs||0),0);
    });
    rows = [['Recruiter','Total Jobs',...boards.flatMap(b=>[b+' Jobs',b+' Views',b+' Apps'])]];
    recs.sort((a,b)=>Object.values(APP.recPostByBoard[b]).reduce((s,v)=>s+v.jobs,0)-Object.values(APP.recPostByBoard[a]).reduce((s,v)=>s+v.jobs,0));
    recs.forEach(r => {
      const total = Object.values(APP.recPostByBoard[r]).reduce((s,v)=>s+v.jobs,0);
      rows.push([r, total, ...boards.flatMap(b=>[APP.recPostByBoard[r][b]?.jobs||0, APP.recPostByBoard[r][b]?.views||0, APP.recPostByBoard[r][b]?.applications||0])]);
    });
  } else if (type === 'rec-srch-matrix') {
    filename = 'recruiter_searches_by_board.csv';
    const SKIP = new Set(['system','total','grand total','','—']);
    const recs = Object.keys(APP.recSrchByBoard).filter(r => r && r.length > 1 && !SKIP.has(r.toLowerCase().trim()));
    const boardSet = new Set();
    for (const r of recs) Object.keys(APP.recSrchByBoard[r]).forEach(b => boardSet.add(b));
    const boards = [...boardSet].sort((a,b) => {
      return recs.reduce((s,r)=>s+(APP.recSrchByBoard[r][b]?.views||0),0) - recs.reduce((s,r)=>s+(APP.recSrchByBoard[r][a]?.views||0),0);
    });
    rows = [['Recruiter','Total Views',...boards]];
    recs.sort((a,b)=>Object.values(APP.recSrchByBoard[b]).reduce((s,v)=>s+v.views,0)-Object.values(APP.recSrchByBoard[a]).reduce((s,v)=>s+v.views,0));
    recs.forEach(r => {
      const total = Object.values(APP.recSrchByBoard[r]).reduce((s,v)=>s+v.views,0);
      rows.push([r, total, ...boards.map(b=>APP.recSrchByBoard[r][b]?.views||0)]);
    });
  }
  if (!rows.length) return;
  const csv = rows.map(r => r.map(c => `"${String(c).replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8;' });
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob);
  a.download = filename; a.click();
}

function exportPDF() {
  closeMenu('export-menu');
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('landscape', 'mm', 'a4');
  const W = doc.internal.pageSize.getWidth();
  const margin = 15;

  doc.setFillColor(27, 42, 74);
  doc.rect(0, 0, W, 22, 'F');
  doc.setTextColor(255,255,255);
  doc.setFontSize(16); doc.setFont(undefined,'bold');
  doc.text('Job Board Usage Dashboard — HonorVet Technologies', margin, 14);
  doc.setFontSize(9); doc.setFont(undefined,'normal');
  doc.text('Generated: ' + new Date().toLocaleString(), W - margin, 14, { align: 'right' });

  let y = 34;
  doc.setTextColor(27,42,74);
  doc.setFontSize(12); doc.setFont(undefined,'bold');
  doc.text('Posting Summary by Board', margin, y); y += 6;

  const postHeaders = ['Board','Jobs','Views','Applications','Avg Views','Conv%'];
  const postData = Object.entries(APP.boardPost).sort((a,b)=>b[1].jobs-a[1].jobs).map(([b,d]) => [
    b, d.jobs, d.views, d.applications,
    d.jobs > 0 ? (d.views/d.jobs).toFixed(1) : '—',
    d.views > 0 ? (d.applications/d.views*100).toFixed(1)+'%' : '—',
  ]);
  autoTable(doc, postHeaders, postData, margin, y, W - margin*2);
  y = doc.lastAutoTable?.finalY ? doc.lastAutoTable.finalY + 10 : y + postData.length*8 + 10;

  if (y > 160) { doc.addPage(); y = 20; }
  doc.setFontSize(12); doc.setFont(undefined,'bold');
  doc.text('Resume Activity by Source', margin, y); y += 6;

  const srchHeaders = ['Source','Searches','Views','Contacts','Responses','Resp%'];
  const srchData = Object.entries(APP.boardSrch).sort((a,b)=>b[1].views-a[1].views).map(([b,d]) => [
    b, d.searches, d.views, d.contacts, d.responses,
    d.searches > 0 ? (d.responses/d.searches*100).toFixed(1)+'%' : '—',
  ]);
  autoTable(doc, srchHeaders, srchData, margin, y, W - margin*2);

  doc.save('job_board_dashboard.pdf');
}

function autoTable(doc, headers, rows, x, y, maxW) {
  const colW = maxW / headers.length;
  doc.setFontSize(9); doc.setFont(undefined,'bold');
  doc.setFillColor(240,244,248); doc.setTextColor(74,85,104);
  doc.rect(x, y-4, maxW, 8, 'F');
  headers.forEach((h, i) => doc.text(h, x + i*colW + 2, y));
  y += 6;
  doc.setFont(undefined,'normal'); doc.setTextColor(26,32,44);
  rows.forEach((row, ri) => {
    if (ri % 2 === 1) { doc.setFillColor(247,250,252); doc.rect(x, y-4, maxW, 7, 'F'); }
    row.forEach((c, i) => doc.text(String(c), x + i*colW + 2, y, { maxWidth: colW - 3 }));
    y += 7;
  });
  // Store for next y calculation
  doc.lastAutoTable = { finalY: y };
}

function dlChart(id) {
  const chart = APP.charts[id];
  if (!chart) return;
  const url = chart.toBase64Image();
  const a = document.createElement('a');
  a.href = url; a.download = id + '.png'; a.click();
}

// ── NAVIGATION ────────────────────────────────────────────
function setTab(el) {
  document.querySelectorAll('.nav-tab').forEach(t => t.classList.remove('active'));
  el.classList.add('active');
  const target = el.dataset.target;
  if (target) {
    const sec = document.getElementById(target);
    if (sec) sec.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}

// Highlight nav on scroll
function initScrollSpy() {
  const sections = [...document.querySelectorAll('.dash-sec')];
  const tabs = [...document.querySelectorAll('.nav-tab')];
  const offset = 170;
  window.addEventListener('scroll', () => {
    let active = sections[0];
    sections.forEach(s => { if (window.scrollY + offset >= s.offsetTop) active = s; });
    if (active) {
      tabs.forEach(t => t.classList.toggle('active', t.dataset.target === active.id));
    }
  }, { passive: true });
}

function toggleMenu(id) {
  document.getElementById(id).classList.toggle('open');
}
function closeMenu(id) {
  document.getElementById(id).classList.remove('open');
}
document.addEventListener('click', e => {
  if (!e.target.closest('.menu-wrap')) document.querySelectorAll('.menu-list').forEach(m => m.classList.remove('open'));
});

// ── FILE UPLOAD EVENTS ─────────────────────────────────────
const _files = { postings: null, searches: null };

function ev_DragOver(e, type) {
  e.preventDefault();
  document.getElementById('drop-' + type).classList.add('over');
}
function ev_DragLeave(e, type) {
  document.getElementById('drop-' + type).classList.remove('over');
}
function ev_Drop(e, type) {
  e.preventDefault();
  document.getElementById('drop-' + type).classList.remove('over');
  const f = e.dataTransfer.files[0];
  if (f) setFile(type, f);
}
function ev_FileInput(e, type) {
  const f = e.target.files[0];
  if (f) setFile(type, f);
}
function setFile(type, file) {
  _files[type] = file;
  const zone   = document.getElementById('drop-' + type);
  const status = document.getElementById('st-' + type);
  zone.classList.add('loaded');
  status.innerHTML = `<span class="dot dot-ok"></span> ${file.name} (${(file.size/1024).toFixed(0)} KB)`;
  checkReady();
}
function checkReady() {
  const gen = document.getElementById('btn-generate');
  gen.disabled = !(_files.postings && _files.searches);
}

function setProgress(pct, label) {
  document.getElementById('prog-fill').style.width = pct + '%';
  document.getElementById('prog-label').textContent = label;
}

async function generateDashboard() {
  const errEl = document.getElementById('upload-error');
  errEl.style.display = 'none';
  document.getElementById('upload-progress').style.display = 'block';
  setProgress(5, 'Reading Job Postings file…');

  try {
    const [wbPost, wbSrch] = await Promise.all([
      readWorkbook(_files.postings),
      readWorkbook(_files.searches),
    ]);
    setProgress(30, 'Parsing posting sheets…');
    APP.postRecs = parseAllPostings(wbPost);
    APP.raw.postings = wbPost;

    setProgress(55, 'Parsing resume search sheets…');
    APP.srchRecs = parseAllSearches(wbSrch);
    APP.raw.searches = wbSrch;

    if (APP.postRecs.length === 0 && APP.srchRecs.length === 0) {
      throw new Error('No recognizable data found. Verify the files are Job Postings.xlsx and Resume Searches.xlsx.');
    }

    setProgress(70, 'Aggregating data…');
    APP.fPostRecs = APP.postRecs;
    APP.fSrchRecs = APP.srchRecs;
    aggregate();
    populateFilters();

    setProgress(80, 'Building charts…');
    showDashboard();

    setProgress(90, 'Rendering KPIs & insights…');
    renderKPIs();
    renderCharts();
    renderTables();
    renderInsights();

    setProgress(100, 'Done!');

    // Header period
    const allDates = APP.postRecs.map(r => r.date).concat(APP.srchRecs.map(r => r.date)).filter(Boolean);
    if (allDates.length) {
      const minD = new Date(Math.min(...allDates));
      const maxD = new Date(Math.max(...allDates));
      document.getElementById('hdr-period').textContent = `${dateFmt(minD,'Mon D')} – ${dateFmt(maxD,'Mon D')}, ${maxD.getFullYear()}`;
    }
    document.getElementById('footer-ts').textContent = 'Generated ' + new Date().toLocaleString();

  } catch (err) {
    document.getElementById('upload-progress').style.display = 'none';
    errEl.textContent = '❌ ' + err.message;
    errEl.style.display = 'block';
    console.error(err);
  }
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false });
        resolve(wb);
      } catch (ex) { reject(ex); }
    };
    reader.onerror = () => reject(new Error('Failed to read file: ' + file.name));
    reader.readAsArrayBuffer(file);
  });
}

function showDashboard() {
  document.getElementById('upload-screen').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';
  initScrollSpy();
  initNavClicks();
}
function goUpload() {
  document.getElementById('dashboard').style.display = 'none';
  document.getElementById('upload-screen').style.display = '';
  document.getElementById('upload-progress').style.display = 'none';
}

function initNavClicks() {
  document.querySelectorAll('.nav-tab').forEach(tab => {
    tab.addEventListener('click', function(e) {
      e.preventDefault();
      setTab(this);
    });
  });
}

// ── INIT ──────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  // Chart.js global defaults
  Chart.defaults.font.family = "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif";
  Chart.defaults.font.size   = 12;
  Chart.defaults.color       = '#4A5568';
  Chart.defaults.plugins.tooltip.backgroundColor = 'rgba(27,42,74,0.92)';
  Chart.defaults.plugins.tooltip.titleFont = { weight: '600' };
  Chart.defaults.plugins.tooltip.padding = 10;
  Chart.defaults.plugins.tooltip.cornerRadius = 6;
  Chart.defaults.plugins.legend.labels.usePointStyle = true;
  Chart.defaults.plugins.legend.labels.pointStyleWidth = 10;
});
