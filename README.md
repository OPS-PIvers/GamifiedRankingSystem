# GamifiedRankingSystem

**Mythos Ascendant** — a Google Apps Script web app that lets students log the mythology-related
media they consume, awards points, and ranks them on a leaderboard.

## Layout

| File | Purpose |
| --- | --- |
| `Code.js` | Apps Script backend: form handling, point calculation, verification, email |
| `index.html` | The student-facing web app (submission form + leaderboard) |
| `Email.html` | Template for the confirmation email |
| `appsscript.json` | Apps Script manifest (timezone, web app access) |
| `.clasp.json` | `clasp` project binding |

The spreadsheet has three tabs, created by **Mythos Admin → Setup Mythos Sheets**:
`Journey Settings` (titles, point thresholds, system settings), `Student Roster`, and
`Student Submissions`.

## Admin menu

Opening the spreadsheet adds a **Mythos Admin** menu:

- **Setup Mythos Sheets** — creates/resets the three tabs.
- **Verify All Pending** — verifies every unverified submission at once.
- **Recalculate All Submissions** — recomputes column I for every row.
- **Repair Roster Formulas** — rewrites the `Total Points Earned` and `Current Title Earned`
  formulas for every student. Useful after pasting in a new class list.
- **Install Verification Email Trigger** — required for students to be emailed when you check
  their "Teacher Verified?" box. The simple `onEdit` trigger runs unauthorized and cannot send
  mail; points are still awarded without it.

The roster's calculated columns (D and E) repair themselves on open, on every submission, and on
every verification, so a formula cleared by accident comes back on its own.

## Tests

The submission and point-calculation logic is covered by a regression suite that runs the real
`Code.js` against an in-memory mock of the Apps Script services (`tests/apps-script-mock.js`).

```sh
npm test
```

No dependencies — it runs on the Node standard library alone, and on every pull request via
GitHub Actions. The mock is deliberately small: it models cells, formulas, `appendRow` and
`getLastRow` closely enough to catch ordering and off-by-one mistakes, and stubs the rest. Anything
it does not model still needs checking in a live spreadsheet.

## Deploying

Push with [`clasp`](https://github.com/google/clasp), then redeploy the web app so students get the
new version:

```sh
clasp push
```
