/**
 * AMS 325 Final Project Peer Evaluation — Google Apps Script web app
 *
 * Part 1  Teammate evaluation (CONFIDENTIAL: only visible in this spreadsheet).
 * Part 2  Presentation evaluation of other groups.
 *
 * Day-to-day switches live in the "Settings" sheet (no redeploy needed):
 *   TEAMMATE_OPEN        TRUE / FALSE
 *   PRESENTATION_OPEN    TRUE / FALSE
 *   PRESENTATION_GROUPS  e.g. 1-7   (day 1)   or   8-13   (day 2)   or   1,2,5
 *
 * Results: Sheet menu "Peer Eval" → "Build results".
 */
const TITLE = 'AMS 325 Final Project Peer Evaluation';

// ----- Part 1: teammate questions (each teammate is rated on each question, 1–5) -----
const TEAM_RATING_QUESTIONS = [
  { q: 'Contribution: How much did this person contribute to the project?',
    low: 'Almost nothing', high: 'A major contributor' },
  { q: 'Reliability: Did this person finish their assigned tasks on time and with good quality?',
    low: 'Rarely', high: 'Always' },
  { q: 'Communication: How well did this person communicate and cooperate with the team?',
    low: 'Poorly', high: 'Excellently' },
];
const TEAM_TEXT_QUESTION = 'What did each teammate mainly work on? (one short sentence each)';
const SELF_QUESTION = 'What was YOUR main contribution to the project?';

// ----- Part 2: presentation questions (1–5) -----
const PRES_QUESTIONS = [
  { q: 'Research questions: The group clearly stated TWO distinct research questions, described the dataset, and explained why the questions matter.',
    low: 'Unclear / missing', high: 'Very clear and well motivated' },
  { q: 'Data & EDA: Data cleaning choices were justified, and the visualizations were meaningful and clearly interpreted (not just shown).',
    low: 'Weak / unexplained', high: 'Thorough and insightful' },
  { q: 'Methods: The group used at least 3 appropriate methods per question, explained WHY each was chosen and its assumptions, and compared them.',
    low: 'Unjustified / unclear', high: 'Well chosen and well explained' },
  { q: 'Results & conclusions: Appropriate metrics were used, results were interpreted in context, and the conclusions actually answer the research questions (with limitations noted).',
    low: 'Not convincing', high: 'Convincing and well supported' },
  { q: 'Delivery: Slides were clear and organized, the talk stayed within 15 minutes, every member presented, and questions were answered well.',
    low: 'Poor', high: 'Excellent' },
];

// ----- Class roster: EDIT THIS for your class -----
// One line per student: name, SBU ID (digits only, or "" if unknown), group number.
// Group 0 is a hidden TEST group: it only appears when the link ends with ?test=1,
// and its responses are excluded from all results.
const ROSTER = [
  {"name": "Student One",   "id": "111111111", "group": 1},
  {"name": "Student Two",   "id": "111111112", "group": 1},
  {"name": "Student Three", "id": "111111113", "group": 1},
  {"name": "Student Four",  "id": "111111114", "group": 2},
  {"name": "Student Five",  "id": "111111115", "group": 2},
  {"name": "Student Six",   "id": "111111116", "group": 2},
  {"name": "Student Seven", "id": "111111117", "group": 3},
  {"name": "Student Eight", "id": "111111118", "group": 3},
  {"name": "Student Nine",  "id": "",          "group": 3},
  // Group 99 = instructors and TAs (rate every group, extra comment fields, scored separately).
  {"name": "Instructor Name", "id": "222222221", "group": 99},
  {"name": "TA Name",         "id": "222222222", "group": 99},
  {"name": "Test Student A", "id": "000000001", "group": 0},
  {"name": "Test Student B", "id": "000000002", "group": 0},
  {"name": "Test Student C", "id": "000000003", "group": 0}
];

const TEST_GROUP = 0;  // hidden test group: only shown with ?test=1, excluded from results
const STAFF_GROUP = 99;  // instructors & TAs: rate every group, extra comments, scored separately
const INSTRUCTORS = ['Instructor Name'];  // can also enter the final grade (/100)
let PREVIEW_TEST_ = false;  // true only inside previewWithTestData()
const SH = { SETTINGS: 'Settings', TEAM: 'Teammate', SELF: 'Self', PRES: 'Presentation', STAFF: 'Staff' };
const COLORS = { dg: '#57d163', lg: '#c3efc8', wh: '#ffffff', ye: '#fff176', re: '#f4a3a0', na: '#eeeeee' };

// ===================== Setup & settings =====================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Peer Eval')
    .addItem('Set up sheets', 'setup')
    .addItem('Build results', 'buildResults')
    .addItem('Preview results incl. test data', 'previewWithTestData')
    .addItem('Clear test data', 'clearTestData')
    .addToUi();
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  const st = ensureSheet_(ss, SH.SETTINGS, ['Setting', 'Value', 'Notes']);
  if (st.getLastRow() < 2) {
    st.getRange(2, 2, 3, 1).setNumberFormat('@');  // keep "1-7" as text, not a date
    st.getRange(2, 1, 3, 3).setValues([
      ['TEAMMATE_OPEN', 'TRUE', 'TRUE / FALSE'],
      ['PRESENTATION_OPEN', 'TRUE', 'TRUE / FALSE'],
      ['PRESENTATION_GROUPS', '1-7', 'Groups that can be rated now, e.g. 1-7 or 8-13 or 1,2,5'],
    ]);
  }
  ensureSheet_(ss, SH.TEAM, teamHeader_());
  ensureSheet_(ss, SH.SELF, ['Timestamp', 'Rater', 'Group', 'ID Verified', 'Own contribution', 'Confidential comment']);
  ensureSheet_(ss, SH.PRES, presHeader_());
  ensureSheet_(ss, SH.STAFF, staffHeader_());
}

function teamHeader_() {
  return ['Timestamp', 'Rater', 'Group', 'ID Verified', 'Teammate']
    .concat(TEAM_RATING_QUESTIONS.map((_, i) => 'Q' + (i + 1)), ['What they worked on']);
}
function presHeader_() {
  return ['Timestamp', 'Rater', 'Rater Group', 'ID Verified', 'Rated Group']
    .concat(PRES_QUESTIONS.map((_, i) => 'Q' + (i + 1)), ['Comment']);
}

function ensureSheet_(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(header);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, header.length).setFontWeight('bold');
  }
  return sh;
}

function settings_() {
  const out = { TEAMMATE_OPEN: true, PRESENTATION_OPEN: true, PRESENTATION_GROUPS: '' };
  const st = SpreadsheetApp.getActive().getSheetByName(SH.SETTINGS);
  if (st && st.getLastRow() > 1) {
    st.getRange(2, 1, st.getLastRow() - 1, 2).getValues().forEach(r => {
      const k = String(r[0]).trim(), v = r[1];
      if (k === 'TEAMMATE_OPEN' || k === 'PRESENTATION_OPEN') out[k] = (v === true || String(v).toUpperCase() === 'TRUE');
      if (k === 'PRESENTATION_GROUPS') {
        // Sheets auto-converts "1-7" / "8-13" into a date (Jan 7 / Aug 13): turn it back into a range.
        out[k] = (v instanceof Date) ? (v.getMonth() + 1) + '-' + v.getDate() : String(v);
      }
    });
  }
  out.groups = parseGroups_(out.PRESENTATION_GROUPS);
  return out;
}

function parseGroups_(s) {
  const all = allGroups_();
  s = String(s || '').trim();
  if (!s) return all;
  const set = new Set();
  s.split(/[,\s]+/).forEach(part => {
    const m = part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (m) { for (let g = +m[1]; g <= +m[2]; g++) set.add(g); }
    else if (/^\d+$/.test(part)) set.add(+part);
  });
  return all.filter(g => set.has(g));
}

function allGroups_() {
  return Array.from(new Set(ROSTER.map(s => s.group))).filter(g => g !== TEST_GROUP && g !== STAFF_GROUP).sort((a, b) => a - b);
}

// ===================== Web app =====================
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Never sends SBU IDs to the browser.
function getInitData(includeTest) {
  const gs = allGroups_().concat([STAFF_GROUP], includeTest ? [TEST_GROUP] : []);
  return {
    title: TITLE,
    groups: gs.map(g => ({
      group: g, label: g === TEST_GROUP ? 'TEST (instructors)' : g === STAFF_GROUP ? 'Instructor / TA' : 'Group ' + g,
      members: ROSTER.map((s, i) => ({ idx: i, name: s.name, group: s.group }))
        .filter(s => s.group === g).map(s => ({ idx: s.idx, name: s.name })),
    })),
    teamQ: TEAM_RATING_QUESTIONS, teamText: TEAM_TEXT_QUESTION, selfQ: SELF_QUESTION,
    presQ: PRES_QUESTIONS,
  };
}

function verify_(idx, sid) {
  const s = ROSTER[Number(idx)];
  if (!s) throw new Error('Please select your name.');
  const clean = String(sid || '').replace(/\D/g, '');
  if (s.id) {
    if (clean !== s.id) throw new Error('The SBU ID does not match our records for ' + s.name + '. Please check and try again.');
    return { s: s, verified: 'yes' };
  }
  if (clean.length < 6) throw new Error('Please enter your SBU ID.');
  return { s: s, verified: 'no (entered ' + clean + ')' };
}

function login(idx, sid) {
  const v = verify_(idx, sid);
  return state_(Number(idx), v.s);
}

function state_(idx, s) {
  if (s.group === STAFF_GROUP) return staffState_(idx, s);
  const st = settings_();
  return {
    idx: idx, name: s.name, group: s.group,
    teamOpen: st.TEAMMATE_OPEN, presOpen: st.PRESENTATION_OPEN,
    presGroups: st.groups.filter(g => g !== s.group),
    teammates: ROSTER.map((x, i) => ({ idx: i, name: x.name, group: x.group }))
      .filter(x => x.group === s.group && x.idx !== idx).map(x => ({ idx: x.idx, name: x.name })),
    teamDone: hasRow_(SH.SELF, s.name),
    presDone: ratedGroups_(s.name),
  };
}

function hasRow_(sheetName, rater) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return false;
  return sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().some(r => r[0] === rater);
}

function ratedGroups_(rater) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.PRES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(r => r[1] === rater).map(r => Number(r[4]));
}

function score_(x) {
  const v = Number(x);
  if (!(v >= 1 && v <= 5 && Math.floor(v) === v)) throw new Error('Please answer every rating question (1–5).');
  return v;
}
function text_(x) { return String(x || '').trim().slice(0, 2000); }

// Insert or overwrite the row whose Rater (col 2) and key column (col 5) match.
function upsert_(sh, rater, key, row) {
  const n = sh.getLastRow();
  if (n > 1) {
    const vals = sh.getRange(2, 1, n - 1, 5).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][1] === rater && String(vals[i][4]) === String(key)) {
        sh.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sh.appendRow(row);
}

/**
 * payload = { ratings: { <mateIdx>: [q1, q2, ...] }, did: { <mateIdx>: text }, self: text, confidential: text }
 */
function submitTeam(idx, sid, payload) {
  const v = verify_(idx, sid);
  if (v.s.group === STAFF_GROUP) throw new Error('Instructors and TAs do not fill in Part 1.');
  if (!settings_().TEAMMATE_OPEN) throw new Error('Teammate evaluation is closed.');
  const me = Number(idx);
  const mates = ROSTER.map((x, i) => ({ i: i, x: x })).filter(o => o.x.group === v.s.group && o.i !== me);
  const nq = TEAM_RATING_QUESTIONS.length;
  const rows = mates.map(o => {
    const sc = (payload.ratings || {})[o.i];
    if (!Array.isArray(sc) || sc.length !== nq) throw new Error('Please rate ' + o.x.name + ' on every question.');
    return { name: o.x.name, sc: sc.map(score_), did: text_((payload.did || {})[o.i]) };
  });
  const self = text_(payload.self);
  if (!self) throw new Error('Please describe your own main contribution.');
  const now = new Date();
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActive();
    const t = ensureSheet_(ss, SH.TEAM, teamHeader_());
    rows.forEach(r => upsert_(t, v.s.name, r.name, [now, v.s.name, v.s.group, v.verified, r.name].concat(r.sc, [r.did])));
    const s = ensureSheet_(ss, SH.SELF, ['Timestamp', 'Rater', 'Group', 'ID Verified', 'Own contribution', 'Confidential comment']);
    upsertSelf_(s, v.s.name, [now, v.s.name, v.s.group, v.verified, self, text_(payload.confidential)]);
  } finally { lock.releaseLock(); }
  return state_(me, v.s);
}

function upsertSelf_(sh, rater, row) {
  const n = sh.getLastRow();
  if (n > 1) {
    const vals = sh.getRange(2, 2, n - 1, 1).getValues();
    for (let i = 0; i < vals.length; i++) {
      if (vals[i][0] === rater) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  sh.appendRow(row);
}

function submitPresentation(idx, sid, group, scores, comment) {
  const v = verify_(idx, sid);
  const st = settings_();
  if (v.s.group === STAFF_GROUP) throw new Error('Please reload the page to use the instructor/TA form.');
  if (!st.PRESENTATION_OPEN) throw new Error('Presentation evaluation is closed.');
  group = Number(group);
  if (group === v.s.group) throw new Error('You cannot rate your own group.');
  if (st.groups.indexOf(group) < 0) throw new Error('Group ' + group + ' is not open for rating right now.');
  if (!Array.isArray(scores) || scores.length !== PRES_QUESTIONS.length) throw new Error('Please answer every question (1–5).');
  const sc = scores.map(score_);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SpreadsheetApp.getActive(), SH.PRES, presHeader_());
    upsert_(sh, v.s.name, group, [new Date(), v.s.name, v.s.group, v.verified, group].concat(sc, [text_(comment)]));
  } finally { lock.releaseLock(); }
  return state_(Number(idx), v.s);
}

// ===================== Instructor / TA evaluation =====================
// Stored in the "Staff" sheet, never mixed into student peer averages.
function range_(n) { return Array.from({ length: n }, (_, i) => i); }

function staffHeader_() {
  return ['Timestamp', 'Rater', 'Role', 'ID Verified', 'Rated Group']
    .concat(PRES_QUESTIONS.map((_, i) => 'Q' + (i + 1)),
      PRES_QUESTIONS.map((_, i) => 'Q' + (i + 1) + ' comment'),
      ['Comment to group (shared)', 'Private note (staff only)', 'Comments on presenters', 'Final grade (/100)']);
}

function parsePresenters_(txt) {
  const out = {};
  String(txt || '').split('\n').forEach(line => {
    const k = line.indexOf(': ');
    if (k > 0) out[line.slice(0, k)] = line.slice(k + 2);
  });
  return out;
}

function staffState_(idx, s) {
  const groups = allGroups_();
  const nq = PRES_QUESTIONS.length;
  const prev = {};
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.STAFF);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, staffHeader_().length).getValues().forEach(r => {
      if (r[1] !== s.name) return;
      prev[Number(r[4])] = {
        scores: r.slice(5, 5 + nq).map(Number),
        qc: r.slice(5 + nq, 5 + 2 * nq).map(x => String(x || '')),
        pub: String(r[5 + 2 * nq] || ''),
        priv: String(r[6 + 2 * nq] || ''),
        pres: parsePresenters_(r[7 + 2 * nq]),
        grade: (r[8 + 2 * nq] === '' || r[8 + 2 * nq] === undefined) ? '' : String(r[8 + 2 * nq]),
      };
    });
  }
  const members = {};
  groups.forEach(g => { members[g] = ROSTER.filter(x => x.group === g).map(x => x.name); });
  return {
    idx: idx, name: s.name, group: s.group, isStaff: true,
    teamOpen: false, presOpen: true, presGroups: groups, teammates: [], teamDone: false,
    presDone: Object.keys(prev).map(Number), staffPrev: prev, members: members,
    canGrade: INSTRUCTORS.indexOf(s.name) >= 0,
  };
}

/** p = { scores: [..], qc: [..per-question comments..], pub: text, priv: text, pres: { name: text } } */
function submitStaff(idx, sid, group, p) {
  const v = verify_(idx, sid);
  if (v.s.group !== STAFF_GROUP) throw new Error('Only instructors and TAs can use this form.');
  group = Number(group);
  if (allGroups_().indexOf(group) < 0) throw new Error('Unknown group.');
  p = p || {};
  if (!Array.isArray(p.scores) || p.scores.length !== PRES_QUESTIONS.length) throw new Error('Please answer every question (1–5).');
  const sc = p.scores.map(score_);
  const qc = PRES_QUESTIONS.map((_, i) => text_((p.qc || [])[i]));
  let grade = '';
  if (INSTRUCTORS.indexOf(v.s.name) >= 0 && p.grade !== undefined && p.grade !== null && String(p.grade).trim() !== '') {
    grade = Number(p.grade);
    if (!(grade >= 0 && grade <= 100)) throw new Error('Final grade must be a number from 0 to 100.');
  }
  const pres = ROSTER.filter(x => x.group === group).map(x => [x.name, text_((p.pres || {})[x.name]).replace(/\s*\n\s*/g, ' ')])
    .filter(x => x[1]).map(x => x[0] + ': ' + x[1]).join('\n');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = ensureSheet_(SpreadsheetApp.getActive(), SH.STAFF, staffHeader_());
    sh.getRange(1, 1, 1, staffHeader_().length).setValues([staffHeader_()]).setFontWeight('bold');
    upsert_(sh, v.s.name, group, [new Date(), v.s.name, 'Instructor/TA', v.verified, group]
      .concat(sc, qc, [text_(p.pub), text_(p.priv), pres, grade]));
  } finally { lock.releaseLock(); }
  return state_(Number(idx), v.s);
}

// "Staff Results": instructor/TA scores and comments per group, next to the student average.
function buildStaffResults_(ss) {
  const nq = PRES_QUESTIONS.length;
  const rows = readRows_(ss, SH.STAFF, staffHeader_().length);
  const stud = readRows_(ss, SH.PRES, 6 + nq);
  const sh = freshSheet_(ss, 'Staff Results');
  let r = writeBlock_(sh, 1, [
    ['Instructor / TA evaluation (scored separately from students)'],
    ['"Comment to group" can be shared with the group. "Private notes" are for instructors/TAs only.'],
  ], null, [0]) + 1;
  const head = ['Group', '# staff'].concat(PRES_QUESTIONS.map((_, i) => 'Q' + (i + 1)),
    ['Staff overall', 'Student overall', 'Final grade (/100)', 'Comments to group (shared)', 'Per-question comments', 'Comments on presenters', 'Private notes (staff only)']);
  const vals = [head];
  allGroups_().forEach(g => {
    const got = rows.filter(x => Number(x[4]) === g);
    const qAvg = range_(nq).map(q => mean_(got.map(x => Number(x[5 + q]))));
    const sAll = stud.filter(x => Number(x[4]) === g).map(x => mean_(x.slice(5, 5 + nq).map(Number)));
    const tag = (x, t) => '[' + x[1] + '] ' + t;
    const pub = got.filter(x => String(x[5 + 2 * nq]).trim()).map(x => tag(x, String(x[5 + 2 * nq]).trim())).join('\n');
    const qcs = [];
    got.forEach(x => range_(nq).forEach(q => {
      const t = String(x[5 + nq + q]).trim();
      if (t) qcs.push('Q' + (q + 1) + ' ' + tag(x, t));
    }));
    const pres = got.filter(x => String(x[7 + 2 * nq]).trim()).map(x => tag(x, '\n' + String(x[7 + 2 * nq]).trim())).join('\n');
    const priv = got.filter(x => String(x[6 + 2 * nq]).trim()).map(x => tag(x, String(x[6 + 2 * nq]).trim())).join('\n');
    vals.push(['Group ' + g, got.length].concat(qAvg.map(r1_), [r1_(mean_(qAvg.filter(x => !isNaN(x)))), r1_(mean_(sAll)), got.map(x => x[8 + 2 * nq]).filter(x => x !== '' && x !== undefined && x !== null).join(', '), pub, qcs.join('\n'), pres, priv]));
  });
  writeBlock_(sh, r, vals, null, [0]);
  for (let c = head.length - 3; c <= head.length; c++) sh.setColumnWidth(c, 300);
  sh.setColumnWidth(head.length - 4, 300);
  sh.getRange(r, 1, vals.length, head.length).setWrap(true).setVerticalAlignment('top');
}

// Deletes every response submitted by the hidden TEST group.
function clearTestData() {
  const ss = SpreadsheetApp.getActive();
  [SH.TEAM, SH.SELF, SH.PRES].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh || sh.getLastRow() < 2) return;
    const vals = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
    for (let i = vals.length - 1; i >= 0; i--) {
      if (vals[i][0] !== '' && Number(vals[i][0]) === TEST_GROUP) sh.deleteRow(i + 2);
    }
  });
}

// Builds the result sheets INCLUDING the hidden TEST group, so the colors can be checked while testing.
// Run "Build results" again afterwards to get the real (test-free) results.
function previewWithTestData() {
  PREVIEW_TEST_ = true;
  try { buildResults(); } finally { PREVIEW_TEST_ = false; }
}

// ===================== Results (VIP-style colors) =====================
/**
 * Every cell shows "score / rater's average" where the rater's average is taken over all
 * scores that rater gave on that question. Color = z = (score - rater avg) / rater SD:
 *   dark green z > 1, light green 0.5..1, white -0.5..0.5, yellow -1..-0.5, red < -1.
 * Summary cells use the average z of all raters.
 */
function buildResults() {
  const ss = SpreadsheetApp.getActive();
  buildTeamResults_(ss);
  buildPresResults_(ss);
  buildPresDetail_(ss);
  buildStaffResults_(ss);
}

function color_(z) {
  if (z === null || z === undefined || isNaN(z)) return COLORS.na;
  if (z > 1) return COLORS.dg;
  if (z > 0.5) return COLORS.lg;
  if (z >= -0.5) return COLORS.wh;
  if (z >= -1) return COLORS.ye;
  return COLORS.re;
}
function mean_(a) { return a.length ? a.reduce((x, y) => x + y, 0) / a.length : NaN; }
function sd_(a) { const m = mean_(a); return Math.sqrt(mean_(a.map(x => (x - m) * (x - m)))); }
function r1_(x) { return isNaN(x) ? '' : Math.round(x * 10) / 10; }
function z_(x, stats) { return stats.sd > 1e-9 ? (x - stats.m) / stats.sd : 0; }

// stats[rater][q] = {m, sd}
function raterStats_(rows, nq, firstQCol) {
  const by = {};
  rows.forEach(r => { (by[r[1]] = by[r[1]] || []).push(r); });
  const out = {};
  Object.keys(by).forEach(k => {
    out[k] = [];
    for (let q = 0; q < nq; q++) {
      const vals = by[k].map(r => Number(r[firstQCol + q]));
      out[k][q] = { m: mean_(vals), sd: sd_(vals) };
    }
  });
  return out;
}

function readRows_(ss, name, ncol) {
  const sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, ncol).getValues()
    .filter(r => r[1] !== '' && (PREVIEW_TEST_ || !(r[2] !== '' && Number(r[2]) === TEST_GROUP)));
}

function freshSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  return sh;
}

// Writes a 2D block of values plus a parallel 2D block of background colors (null = leave default).
function writeBlock_(sh, row, values, bgs, boldRows) {
  if (!values.length) return row;
  const w = Math.max.apply(null, values.map(r => r.length));
  const vals = values.map(r => r.concat(Array(w - r.length).fill('')));
  const rng = sh.getRange(row, 1, vals.length, w);
  rng.setValues(vals);
  if (bgs) rng.setBackgrounds(bgs.map(r => r.concat(Array(w - r.length).fill(null)).map(c => c || '#ffffff')));
  (boldRows || []).forEach(i => sh.getRange(row + i, 1, 1, w).setFontWeight('bold'));
  return row + vals.length;
}

function buildTeamResults_(ss) {
  const nq = TEAM_RATING_QUESTIONS.length;
  const rows = readRows_(ss, SH.TEAM, 6 + nq);
  const selfRows = readRows_(ss, SH.SELF, 6);
  const stats = raterStats_(rows, nq, 5);
  const sh = freshSheet_(ss, 'Team Results');
  let r = 1;
  r = writeBlock_(sh, r, [
    ['Teammate evaluation — CONFIDENTIAL (instructor/TA only)'],
    ['Cells show score / that rater\'s average on the question. Colors: dark green > +1 SD, light green +0.5..1, white ±0.5, yellow -0.5..-1, red < -1 SD (relative to how that rater scored their other teammates).'],
  ], null, [0]) + 1;

  allGroups_().concat(PREVIEW_TEST_ ? [TEST_GROUP] : []).forEach(g => {
    const members = ROSTER.filter(s => s.group === g).map(s => s.name);
    r = writeBlock_(sh, r, [['Group ' + g]], [[COLORS.na]], [0]);

    // Summary per member
    const head = ['Student'].concat(TEAM_RATING_QUESTIONS.map((_, i) => 'Q' + (i + 1) + ' avg'),
      ['Overall', '# raters', 'Submitted?', 'Own contribution (self)', 'What teammates say they did', 'Confidential comment they wrote']);
    const vals = [head], bgs = [head.map(() => null)];
    members.forEach(m => {
      const got = rows.filter(x => x[4] === m);
      const qAvg = [], qZ = [];
      for (let q = 0; q < nq; q++) {
        qAvg.push(mean_(got.map(x => Number(x[5 + q]))));
        qZ.push(got.length ? mean_(got.map(x => z_(Number(x[5 + q]), stats[x[1]][q]))) : NaN);
      }
      const overall = mean_(qAvg.filter(x => !isNaN(x)));
      const overallZ = mean_(qZ.filter(x => !isNaN(x)));
      const self = selfRows.find(x => x[1] === m);
      const did = got.filter(x => String(x[5 + nq]).trim()).map(x => '• ' + String(x[5 + nq]).trim()).join('\n');
      vals.push([m].concat(qAvg.map(r1_), [r1_(overall), got.length, self ? 'yes' : 'NO',
        self ? self[4] : '', did, self ? self[5] : '']));
      bgs.push([null].concat(qZ.map(color_), [color_(overallZ), null, self ? null : COLORS.re, null, null, null]));
    });
    r = writeBlock_(sh, r, vals, bgs, [0]);

    // VIP-style detail matrix per question: rows = rater, cols = ratee
    TEAM_RATING_QUESTIONS.forEach((tq, q) => {
      const dv = [['Q' + (q + 1) + ': ' + tq.q, ''].concat(members.map(() => ''))];
      const db = [[null, null].concat(members.map(() => null))];
      dv.push(['Rater ↓  /  Rated →', ''].concat(members));
      db.push([null, null].concat(members.map(() => null)));
      members.forEach(rater => {
        const line = [rater, ''], lb = [null, null];
        members.forEach(ratee => {
          const hit = rows.find(x => x[1] === rater && x[4] === ratee);
          if (rater === ratee) { line.push('—'); lb.push(COLORS.na); }
          else if (!hit) { line.push(''); lb.push(null); }
          else {
            const s = Number(hit[5 + q]), st = stats[rater][q];
            line.push(s + ' / ' + r1_(st.m)); lb.push(color_(z_(s, st)));
          }
        });
        dv.push(line); db.push(lb);
      });
      r = writeBlock_(sh, r, dv, db, [1]);
    });
    r += 2;
  });
  sh.setColumnWidth(1, 170);
  for (let c = 2; c <= nq + 4; c++) sh.setColumnWidth(c, 90);
  sh.getRange(1, 1, Math.max(r, 2), 12).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 2, 1).setWrap(false);
}

function buildPresResults_(ss) {
  const nq = PRES_QUESTIONS.length;
  const rows = readRows_(ss, SH.PRES, 6 + nq);
  const stats = raterStats_(rows, nq, 5);
  const sh = freshSheet_(ss, 'Presentation Results');
  let r = writeBlock_(sh, 1, [
    ['Presentation evaluation — group summary'],
    ['Cells show the raw average score. Color = average of each rater\'s z-score (score vs. that rater\'s own average): dark green > +1 SD, light green +0.5..1, white ±0.5, yellow -0.5..-1, red < -1 SD.'],
  ], null, [0]) + 1;
  const head = ['Group', '# raters'].concat(PRES_QUESTIONS.map((x, i) => 'Q' + (i + 1)), ['Overall', 'Adjusted rank', 'Comments (anonymous)']);
  const items = allGroups_().map(g => {
    const got = rows.filter(x => Number(x[4]) === g);
    const qAvg = [], qZ = [];
    for (let q = 0; q < nq; q++) {
      qAvg.push(mean_(got.map(x => Number(x[5 + q]))));
      qZ.push(got.length ? mean_(got.map(x => z_(Number(x[5 + q]), stats[x[1]][q]))) : NaN);
    }
    const comments = got.map(x => String(x[5 + nq]).trim()).filter(c => c).map(c => '• ' + c).join('\n');
    return { g: g, n: got.length, qAvg: qAvg, qZ: qZ, ov: mean_(qAvg.filter(x => !isNaN(x))), ovZ: mean_(qZ.filter(x => !isNaN(x))), comments: comments };
  });
  const ranked = items.filter(i => i.n).slice().sort((a, b) => b.ovZ - a.ovZ);
  const vals = [head], bgs = [head.map(() => null)];
  items.forEach(i => {
    const rank = ranked.indexOf(i);
    vals.push(['Group ' + i.g, i.n].concat(i.qAvg.map(r1_), [r1_(i.ov), rank >= 0 ? rank + 1 : '', i.comments]));
    bgs.push([null, null].concat(i.qZ.map(color_), [color_(i.ovZ), null, null]));
  });
  r = writeBlock_(sh, r, vals, bgs, [0]) + 2;
  const legend = PRES_QUESTIONS.map((x, i) => ['Q' + (i + 1), x.q]);
  writeBlock_(sh, r, legend, null, []);
  sh.setColumnWidth(1, 90); sh.setColumnWidth(head.length, 420);
  sh.getRange(1, 1, r + legend.length, head.length).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 2, 1).setWrap(false);
}

function buildPresDetail_(ss) {
  const nq = PRES_QUESTIONS.length;
  const rows = readRows_(ss, SH.PRES, 6 + nq);
  const sh = freshSheet_(ss, 'Presentation Detail');
  const groups = allGroups_();
  // Overall per row = mean of the questions; rater stats computed on that overall.
  const ov = rows.map(x => ({ rater: x[1], rg: x[2], g: Number(x[4]), s: mean_(x.slice(5, 5 + nq).map(Number)) }));
  const by = {};
  ov.forEach(o => { (by[o.rater] = by[o.rater] || []).push(o); });
  const raters = Object.keys(by).sort((a, b) => by[a][0].rg - by[b][0].rg || a.localeCompare(b));
  const head = ['Rater', 'Rater group', 'Rater avg'].concat(groups.map(g => 'G' + g));
  const vals = [['Presentation evaluation — each rater × group (overall score / rater\'s average), colored by rater z-score'], head];
  const bgs = [[null], head.map(() => null)];
  raters.forEach(k => {
    const list = by[k], st = { m: mean_(list.map(o => o.s)), sd: sd_(list.map(o => o.s)) };
    const line = [k, list[0].rg, r1_(st.m)], lb = [null, null, null];
    groups.forEach(g => {
      const hit = list.find(o => o.g === g);
      if (g === list[0].rg) { line.push('own'); lb.push(COLORS.na); }
      else if (!hit) { line.push(''); lb.push(null); }
      else { line.push(r1_(hit.s) + ' / ' + r1_(st.m)); lb.push(color_(z_(hit.s, st))); }
    });
    vals.push(line); bgs.push(lb);
  });
  writeBlock_(sh, 1, vals, bgs, [0, 1]);
  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 160);
}
