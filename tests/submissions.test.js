/**
 * Regression tests for the Mythos Ascendant submission and point-calculation logic.
 *
 * Every case below corresponds to a real defect that reached the live sheet: student submissions
 * silently vanishing, and the Student Roster not totalling from Student Submissions. Run with
 * `npm test`.
 */
const { Sheet, loadCode } = require('./apps-script-mock');

let passed = 0;
let failed = 0;
let currentGroup = '';

function group(name) {
  currentGroup = name;
  console.log('\n' + name);
}

function check(description, condition, detail) {
  if (condition) {
    passed++;
    console.log('  ok    ' + description);
  } else {
    failed++;
    console.log('  FAIL  ' + description + (detail ? '\n          ' + detail : ''));
  }
}

const HEADERS = {
  roster: ['Student Name', 'Student Email', 'Class Period', 'Total Points Earned', 'Current Title Earned'],
  submissions: ['Timestamp', 'Student Email', 'Type of Media', 'Title of Media', 'Bonus Points (Yes/No)',
                'Reflection: Date/Time', 'Reflection: Mythological Connection', 'Reflection: Analysis',
                'Points', 'Teacher Verified?'],
};

/** Builds a spreadsheet in the shape setupMythosSheets() produces. */
function buildSheets(options = {}) {
  const settings = new Sheet('Journey Settings', [
    ['Points', 'Title', 'Congratulations Message', 'Image URL'],
    [0, 'Gnome', 'msg', ''],
    [20, 'Gremlin', 'msg', ''],
    [30, 'Kobold', 'msg', ''],
    [],
    ['System Setting', 'Value'],
    ['Enable Teacher Verification', options.verification === undefined ? 'TRUE' : options.verification],
    ['Main Logo', ''],
  ]);

  const roster = new Sheet('Student Roster', [HEADERS.roster]);
  (options.rosterRows || []).forEach(row => roster.rows.push(row));

  const submissions = new Sheet('Student Submissions', [HEADERS.submissions]);
  (options.submissionRows || []).forEach(row => submissions.rows.push(row));

  return {
    'Journey Settings': settings,
    'Student Roster': roster,
    'Student Submissions': submissions,
  };
}

const ANALYSIS = 'A long analysis about fate and free will.';

/** A complete, valid form payload. */
const form = (overrides = {}) => Object.assign({
  studentEmail: 'kid@orono.k12.mn.us',
  mediaType: 'Video Game',
  mediaTitle: 'Hades',
  bonusPoints: 'No',
  reflectionDate: 'Friday, Sept 12, 1:00 PM - 1:45 PM',
  reflectionConnection: 'It uses the Greek underworld.',
  reflectionAnalysis: ANALYSIS,
}, overrides);

/** A submission row as it appears in the sheet. */
const submissionRow = (o = {}) => [
  o.timestamp || new Date(),
  o.email || 'kid@orono.k12.mn.us',
  o.mediaType || 'Video Game',
  o.title || 'Hades',
  o.bonus || 'No',
  'd', 'c',
  o.analysis || ANALYSIS,
  o.points === undefined ? 10 : o.points,
  o.verified === undefined ? false : o.verified,
];

// ---------------------------------------------------------------------------------------------

group('A normal submission is recorded');
{
  const sheets = buildSheets();
  const code = loadCode(sheets);
  const result = code.processSubmission(form());
  const submissions = sheets['Student Submissions'];

  check('returns success', result.status === 'success', JSON.stringify(result));
  check('appends exactly one row', submissions.getLastRow() === 2, 'lastRow=' + submissions.getLastRow());
  check('a first Video Game scores 10', submissions._cell(2, 9) === 10, 'got ' + submissions._cell(2, 9));
  check('is left unverified while verification is on', submissions._cell(2, 10) === false);
  check('creates the roster row', sheets['Student Roster'].getLastRow() === 2);
  check('gives the new roster row a points formula', /SUMIFS/.test(sheets['Student Roster'].formulas['2,4'] || ''));
  check('gives the new roster row a title formula', /VLOOKUP/.test(sheets['Student Roster'].formulas['2,5'] || ''));
}

group('Roster totals: a hand-typed roster gets the formulas it never had');
{
  // setupMythosSheets() used to write the SUMIFS to D2 only, so students pasted in by hand
  // silently never accumulated points.
  const sheets = buildSheets({ rosterRows: [
    ['Ada', 'ada@orono.k12.mn.us', '3', '', ''],
    ['Ben', 'ben@orono.k12.mn.us', '3', '', ''],
    ['Cy',  'cy@orono.k12.mn.us',  '4', '', ''],
  ] });
  const code = loadCode(sheets);
  const repaired = code.ensureRosterFormulas();
  const roster = sheets['Student Roster'];

  check('repairs two columns for each of three students', repaired === 6, 'got ' + repaired);
  check('every student gets a points formula',
        [2, 3, 4].every(r => /SUMIFS/.test(roster.formulas[r + ',4'] || '')));
  check('each formula points at its own row', (roster.formulas['4,4'] || '').includes('$B4'), roster.formulas['4,4']);
  check('the title lookup is bounded to the numeric title rows',
        (roster.formulas['4,5'] || '').includes("'Journey Settings'!$A$2:$B$4"), roster.formulas['4,5']);
  check('running it again changes nothing', code.ensureRosterFormulas() === 0);
}

group('Roster totals: a deleted formula heals itself');
{
  const sheets = buildSheets({ rosterRows: [['Kid', 'kid@orono.k12.mn.us', '3', '', '']] });
  const code = loadCode(sheets);
  code.ensureRosterFormulas();

  // The teacher clears column D.
  delete sheets['Student Roster'].formulas['2,4'];
  sheets['Student Roster']._set(2, 4, '');

  code.processSubmission(form());
  check('the next submission restores it', /SUMIFS/.test(sheets['Student Roster'].formulas['2,4'] || ''));
  check('without adding a duplicate roster row', sheets['Student Roster'].getLastRow() === 2);
}

group('Roster totals: email capitalization does not split a student in two');
{
  const sheets = buildSheets({ rosterRows: [['Kid', 'Kid@Orono.K12.MN.US', '3', 0, 'Gnome']] });
  const code = loadCode(sheets);
  code.processSubmission(form());

  check('matches the existing roster row', sheets['Student Roster'].getLastRow() === 2,
        'roster students=' + (sheets['Student Roster'].getLastRow() - 1));
}

group('Duplicate guard: does not swallow real submissions');
{
  // A script/spreadsheet timezone mismatch makes stored timestamps read as being in the future.
  // The guard used to test `now - rowTime < 10000` with no lower bound, which matches forever.
  const future = new Date(Date.now() + 6 * 3600 * 1000);
  const sheets = buildSheets({ submissionRows: [submissionRow({ timestamp: future, verified: true })] });
  const code = loadCode(sheets);
  const result = code.processSubmission(form());

  check('a future-dated row does not block a new submission', sheets['Student Submissions'].getLastRow() === 3,
        'lastRow=' + sheets['Student Submissions'].getLastRow());
  check('and it is not reported as a duplicate', !/duplicate/i.test(result.message), result.message);
}
{
  const sheets = buildSheets({ submissionRows: [submissionRow({ timestamp: new Date(Date.now() - 86400000), verified: true })] });
  const code = loadCode(sheets);
  code.processSubmission(form());

  check('the same title a day later is recorded, not swallowed', sheets['Student Submissions'].getLastRow() === 3);
  check('and scores as a second submission (5, not 10)', sheets['Student Submissions']._cell(3, 9) === 5,
        'got ' + sheets['Student Submissions']._cell(3, 9));
}

group('Duplicate guard: still catches a genuine double-click');
{
  const sheets = buildSheets({ submissionRows: [submissionRow({ timestamp: new Date(Date.now() - 1000) })] });
  const code = loadCode(sheets);
  const result = code.processSubmission(form());

  check('writes no second row', sheets['Student Submissions'].getLastRow() === 2);
  check('and says so honestly', result.status === 'success' && /duplicate/i.test(result.message), result.message);
}
{
  // The guard used to scan lastRow-5..lastRow-1, never looking at the newest row - the one a
  // double-click would just have created.
  const rows = [];
  for (let i = 0; i < 9; i++) {
    rows.push(submissionRow({ email: 'other@orono.k12.mn.us', timestamp: new Date(Date.now() - 500000), title: 'T' + i, analysis: 'a' + i, verified: true }));
  }
  rows.push(submissionRow({ timestamp: new Date(Date.now() - 1000) }));

  const sheets = buildSheets({ submissionRows: rows });
  const code = loadCode(sheets);
  code.processSubmission(form());

  check('catches the duplicate even on a long sheet', sheets['Student Submissions'].getLastRow() === 11,
        'lastRow=' + sheets['Student Submissions'].getLastRow());
}

group('Input validation: incomplete submissions are reported, not thrown or filed blank');
{
  const sheets = buildSheets();
  const code = loadCode(sheets);
  const result = code.processSubmission(form({ bonusPoints: undefined, mediaTitle: '   ' }));

  check('names the missing field', result.status === 'error' && /Title of Media/.test(result.message), result.message);
  check('writes nothing', sheets['Student Submissions'].getLastRow() === 1);
}
{
  // The email input is made readOnly once auto-populated, which exempts it from the browser's
  // "required" check, so a blank value really can reach the server.
  const sheets = buildSheets();
  const code = loadCode(sheets, { activeUser: 'kid@orono.k12.mn.us' });
  const result = code.processSubmission(form({ studentEmail: '' }));

  check('a blank email falls back to the signed-in user',
        result.status === 'success' && sheets['Student Submissions']._cell(2, 2) === 'kid@orono.k12.mn.us',
        JSON.stringify(result));
}
{
  const sheets = buildSheets();
  const code = loadCode(sheets, { activeUser: '' });
  const result = code.processSubmission(form({ studentEmail: '' }));

  check('with no session either, it is rejected rather than filed unattributable',
        result.status === 'error' && sheets['Student Submissions'].getLastRow() === 1, JSON.stringify(result));
}

group('A mail failure never reports a saved submission as failed');
{
  const sheets = buildSheets({ verification: 'FALSE' });
  const code = loadCode(sheets);
  code.MailApp.sendEmail = () => { throw new Error('Service invoked too many times for one day: email.'); };
  const result = code.processSubmission(form());

  check('still reports success', result.status === 'success', JSON.stringify(result));
  check('because the row really was saved', sheets['Student Submissions'].getLastRow() === 2);
  check('and is marked verified with verification off', sheets['Student Submissions']._cell(2, 10) === true);
}

group('Verification: recalculating does not email the whole class');
{
  const rows = [];
  for (let i = 0; i < 5; i++) rows.push(submissionRow({ title: 'T' + i, analysis: 'a' + i, points: 0 }));

  const sheets = buildSheets({ rosterRows: [['Kid', 'kid@orono.k12.mn.us', '3', 25, 'Gremlin']], submissionRows: rows });
  const code = loadCode(sheets);
  for (let row = 2; row <= 6; row++) code.verifySubmission(row, true);

  const points = [2, 3, 4, 5, 6].map(r => sheets['Student Submissions']._cell(r, 9)).join(',');
  check('skipEmail is honoured', code.sentEmails.length === 0, code.sentEmails.length + ' emails sent');
  check('points follow the first/second/third+ ladder', points === '10,5,1,1,1', 'got ' + points);
  check('all rows end up verified', [2, 3, 4, 5, 6].every(r => sheets['Student Submissions']._cell(r, 10) === true));
}
{
  const sheets = buildSheets({
    rosterRows: [['Kid', 'kid@orono.k12.mn.us', '3', 25, 'Gremlin']],
    submissionRows: [submissionRow({ email: 'Kid@Orono.K12.MN.US', points: 0 })],
  });
  const code = loadCode(sheets);
  code.verifySubmission(2);

  check('verifying one submission does email the student', code.sentEmails.length === 1);
  check('and finds them despite differing capitalization',
        code.sentEmails.length === 1 && code.sentEmails[0].to === 'Kid@Orono.K12.MN.US');
}

group('The edit trigger tolerates being run by hand and on the wrong cell');
{
  const sheets = buildSheets({ submissionRows: [submissionRow({ points: 0 })] });
  const code = loadCode(sheets);
  const submissions = sheets['Student Submissions'];

  let threw = false;
  try { code.onEdit(undefined); } catch (e) { threw = true; }
  check('onEdit() with no event object does not throw', !threw);

  code.onEdit({ range: submissions.getRange(2, 3, 1, 1) });
  check('an edit outside column J is ignored', submissions._cell(2, 10) === false);

  code.onEdit({ range: submissions.getRange(2, 10, 1, 1).setValue(true) });
  check('checking the box awards the points', submissions._cell(2, 9) === 10, 'got ' + submissions._cell(2, 9));
  check('the simple trigger sends no mail, since it is not authorized to', code.sentEmails.length === 0);
}

group('A contended lock is reported honestly rather than losing the row');
{
  const sheets = buildSheets();
  const code = loadCode(sheets, { lockFails: true });
  const result = code.processSubmission(form());

  check('tells the student to retry', result.status === 'error' && /busy/i.test(result.message), result.message);
  check('and leaves nothing half-written', sheets['Student Submissions'].getLastRow() === 1);
}

// ---------------------------------------------------------------------------------------------

console.log(`\n${passed} passed, ${failed} failed\n`);
process.exit(failed > 0 ? 1 : 0);
