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

group('Reflection copy: the student is emailed what they just wrote');
{
  const sheets = buildSheets();
  const code = loadCode(sheets);
  const result = code.processSubmission(form({ reflectionConnection: 'Hades rules the <b>underworld</b>.' }));
  const copy = code.sentEmails.find(m => /reflection/i.test(m.subject));

  check('sends one email even while verification is pending', code.sentEmails.length === 1,
        code.sentEmails.length + ' emails sent');
  check('to the student', copy && copy.to === 'kid@orono.k12.mn.us');
  check('naming the media in the subject', copy && /Hades/.test(copy.subject), copy && copy.subject);
  check('with all three reflection answers in the plain-text body',
        copy && copy.body.includes(ANALYSIS) && copy.body.includes('underworld') &&
        copy.body.includes('Friday, Sept 12'), copy && copy.body);
  check('and tells the student it was sent', /copy of your reflection/i.test(result.message), result.message);

  const template = code.templates.find(t => t.file === 'Reflections');
  check('hands the HTML template the raw text, which its <?= ?> tags escape',
        template && template.entries[0].reflectionConnection === 'Hades rules the <b>underworld</b>.');
}
{
  const sheets = buildSheets({ verification: 'FALSE' });
  const code = loadCode(sheets);
  code.processSubmission(form());

  check('with verification off it arrives alongside the points email',
        code.sentEmails.length === 2 && code.sentEmails.some(m => /Journey Update/.test(m.subject)),
        code.sentEmails.map(m => m.subject).join(' | '));
}
{
  const sheets = buildSheets();
  sheets['Journey Settings'].rows.push(['Email Reflection Copy On Submit', 'FALSE']);
  const code = loadCode(sheets);
  const result = code.processSubmission(form());

  check('setting it to FALSE turns the copy off', code.sentEmails.length === 0);
  check('and the student is not told one is coming', !/copy of your reflection/i.test(result.message));
}
{
  const sheets = buildSheets();
  const code = loadCode(sheets);
  code.MailApp.sendEmail = () => { throw new Error('Service invoked too many times for one day: email.'); };
  const result = code.processSubmission(form());

  check('a failed copy still reports the saved submission as a success',
        result.status === 'success' && sheets['Student Submissions'].getLastRow() === 2, JSON.stringify(result));
  check('without claiming the copy was sent', !/copy of your reflection/i.test(result.message));
}
{
  const sheets = buildSheets({ submissionRows: [submissionRow({ timestamp: new Date(Date.now() - 1000) })] });
  const code = loadCode(sheets);
  code.processSubmission(form());

  check('a blocked double-click does not send a second copy', code.sentEmails.length === 0);
}

group('Reflection logs: the teacher can email each student everything they have written');
{
  const day = n => new Date(Date.UTC(2026, 9, n));
  const options = () => ({
    rosterRows: [
      ['Ada', 'ada@orono.k12.mn.us', '3', 30, 'Kobold'],
      ['Ben', 'ben@orono.k12.mn.us', '4', 10, 'Gnome'],
    ],
    // Deliberately out of order, with Ada's address typed in a different case.
    submissionRows: [
      submissionRow({ email: 'ada@orono.k12.mn.us', title: 'Percy Jackson', timestamp: day(20), verified: true }),
      submissionRow({ email: 'ben@orono.k12.mn.us', title: 'Thor', timestamp: day(18) }),
      submissionRow({ email: 'ADA@orono.k12.mn.us', title: 'Hades', timestamp: day(5), verified: true }),
      submissionRow({ email: 'drop-in@orono.k12.mn.us', title: 'Moana', timestamp: day(7) }),
    ],
  });

  {
    const code = loadCode(buildSheets(options()), { promptText: '' });
    code.emailReflectionLogs();

    const ada = code.sentEmails.find(m => m.to === 'ada@orono.k12.mn.us');
    check('a blank answer emails every student who has submitted, once each', code.sentEmails.length === 3,
          code.sentEmails.map(m => m.to).join(', '));
    check('a student\'s rows are gathered despite differing capitalization',
          ada && /Hades/.test(ada.body) && /Percy Jackson/.test(ada.body));
    check('oldest first, even when the sheet is not in order',
          ada && ada.body.indexOf('Hades') < ada.body.indexOf('Percy Jackson'));
    check('with their total and title', ada && /Glory Points: 30/.test(ada.body) && /Heroic Title: Kobold/.test(ada.body));
    check('a student missing from the roster still gets theirs',
          code.sentEmails.some(m => m.to === 'drop-in@orono.k12.mn.us'));
    check('and the teacher is told how many went', code.alerts.some(a => /Sent reflection logs to 3/.test(a)),
          code.alerts.join(' | '));
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: '4' });
    code.emailReflectionLogs();
    check('a class period sends to only that class',
          code.sentEmails.length === 1 && code.sentEmails[0].to === 'ben@orono.k12.mn.us',
          code.sentEmails.map(m => m.to).join(', '));
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: ' Ada@Orono.k12.mn.us ' });
    code.emailReflectionLogs();
    check('an email address sends to only that student',
          code.sentEmails.length === 1 && code.sentEmails[0].to === 'ada@orono.k12.mn.us');
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: '', mailQuota: 2 });
    code.emailReflectionLogs();
    check('it refuses up front rather than running out of quota halfway', code.sentEmails.length === 0);
    check('and explains why', code.alerts.some(a => /only send 2 more today/.test(a)), code.alerts.join(' | '));
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: null });
    code.emailReflectionLogs();
    check('cancelling the prompt sends nothing', code.sentEmails.length === 0);
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: '', confirm: false });
    code.emailReflectionLogs();
    check('answering No to the confirmation sends nothing', code.sentEmails.length === 0);
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: '7' });
    code.emailReflectionLogs();
    check('a period with no submissions says so', code.sentEmails.length === 0 &&
          code.alerts.some(a => /No submissions found/.test(a)), code.alerts.join(' | '));
  }
  {
    const code = loadCode(buildSheets(options()), { promptText: '' });
    const realSend = code.MailApp.sendEmail;
    code.MailApp.sendEmail = m => { if (m.to === 'ben@orono.k12.mn.us') throw new Error('bounced'); realSend(m); };
    code.emailReflectionLogs();
    check('one failed address does not stop the rest', code.sentEmails.length === 2);
    check('and is named to the teacher', code.alerts.some(a => /ben@orono\.k12\.mn\.us/.test(a)),
          code.alerts.join(' | '));
  }
}

// ---------------------------------------------------------------------------------------------

console.log(`\n${passed} passed, ${failed} failed\n`);
process.exit(failed > 0 ? 1 : 0);
