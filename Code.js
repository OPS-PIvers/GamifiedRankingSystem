/**
 * @OnlyCurrentDoc
 *
 * A Google Apps Script to serve as the backend for the Mythos Ascendant web app.
 * It handles form submissions, calculates points, and serves the leaderboard data.
 */

const JOURNEY_SETTINGS_SHEET_NAME = "Journey Settings";
const STUDENT_ROSTER_SHEET_NAME = "Student Roster";
const STUDENT_SUBMISSIONS_SHEET_NAME = "Student Submissions";

const POINT_SYSTEM = {
  'Written Story (book, online, etc)': { first: 20, second: 10, thirdPlus: 5 },
  'Movie/TV Show/Play/Musical': { first: 10, second: 5, thirdPlus: 1 },
  'Video Game': { first: 10, second: 5, thirdPlus: 1 },
  'Podcast/Audio': { first: 10, second: 5, thirdPlus: 1 },
  'Graphic Novel/Comic Book': { first: 10, second: 5, thirdPlus: 1 },
  'Other': { first: 5, second: 1, thirdPlus: 0 }
};

const BONUS_POINTS_VALUE = 5; // Points for each myth read that connects to the modern story

// How close together two identical submissions have to be to count as an accidental double-click.
const DUPLICATE_WINDOW_MS = 10000;

// How long a submission will wait for the lock before giving up and asking the student to retry.
const SUBMISSION_LOCK_TIMEOUT_MS = 30000;

/**
 * Normalizes an email address for comparison.
 * Emails typed by a teacher into the roster and emails returned by Session can differ
 * in case and whitespace, which would otherwise create duplicate roster rows.
 * @param {*} value The raw value to normalize.
 * @returns {string} The trimmed, lower-cased email.
 */
function normalizeEmail(value) {
  return String(value === null || value === undefined ? "" : value).trim().toLowerCase();
}

/**
 * Adds a custom menu to the spreadsheet when it's opened.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Mythos Admin')
      .addItem('Setup Mythos Sheets', 'setupMythosSheets')
      .addItem('Verify All Pending', 'batchVerifyPending')
      .addItem('Recalculate All Submissions', 'recalculateAllSubmissions')
      .addItem('Repair Roster Formulas', 'repairRosterFormulas')
      .addItem('Install Verification Email Trigger', 'installVerificationTrigger')
      .addToUi();

  // Self-heal the roster totals every time the spreadsheet is opened, in case a formula in
  // column D or E was cleared or overwritten.
  try {
    ensureRosterFormulas();
  } catch (error) {
    console.log("Could not repair roster formulas on open:", error);
  }
}

/**
 * Iterates through all submissions and recalculates points based on media type and history.
 * Useful for retroactively fixing points if they show as 0.
 */
function recalculateAllSubmissions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
  const lastRow = submissionsSheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No submissions found to recalculate.");
    return;
  }

  const confirm = SpreadsheetApp.getUi().alert(
    "Recalculate All",
    "Are you sure you want to recalculate points for ALL submissions? This will update Column I for every row.",
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (confirm !== SpreadsheetApp.getUi().Button.YES) return;

  // Process rows one by one to ensure the "previous submission count" logic works correctly for each row.
  // skipEmail is essential here: without it a recalculation of a full class would send one email per
  // row and exhaust the account's daily MailApp quota.
  for (let i = 2; i <= lastRow; i++) {
    verifySubmission(i, true);
  }

  // Recalculating points is pointless if the roster has no formulas pointing at them.
  refreshRosterFormulas();

  SpreadsheetApp.getUi().alert("All submissions have been recalculated.");
}

/**
 * Rewrites the "Total Points Earned" and "Current Title Earned" formulas for every student
 * on the roster.
 *
 * Both setupMythosSheets() and processSubmission() used to only ever write these formulas to a
 * single row, so any student typed or pasted into the roster by hand had an empty column D and
 * never accumulated points from the Student Submissions sheet.
 */
function refreshRosterFormulas() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);
  const lastRow = rosterSheet.getLastRow();

  if (lastRow < 2) return 0;

  const pointsFormulas = [];
  const titleFormulas = [];
  for (let row = 2; row <= lastRow; row++) {
    pointsFormulas.push([buildRosterPointsFormula(row)]);
    titleFormulas.push([buildRosterTitleFormula(row)]);
  }

  rosterSheet.getRange(2, 4, pointsFormulas.length, 1).setFormulas(pointsFormulas);
  rosterSheet.getRange(2, 5, titleFormulas.length, 1).setFormulas(titleFormulas);
  SpreadsheetApp.flush();

  return lastRow - 1;
}

/**
 * Restores only the roster formulas that are missing or have been overwritten.
 *
 * Columns D and E hold formulas, and a teacher sorting, clearing or pasting over the roster can
 * wipe them - which silently stops points accumulating with no visible error. This is called on
 * open and on every submission and verification so the damage repairs itself.
 *
 * @returns {number} How many cells were repaired.
 */
function ensureRosterFormulas() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);
  const lastRow = rosterSheet ? rosterSheet.getLastRow() : 0;

  if (lastRow < 2) return 0;

  const numRows = lastRow - 1;
  const existingFormulas = rosterSheet.getRange(2, 4, numRows, 2).getFormulas();
  let repaired = 0;

  for (let i = 0; i < numRows; i++) {
    const row = i + 2;

    if (!isPointsFormula(existingFormulas[i][0])) {
      rosterSheet.getRange(row, 4).setFormula(buildRosterPointsFormula(row));
      repaired++;
    }

    if (!isTitleFormula(existingFormulas[i][1])) {
      rosterSheet.getRange(row, 5).setFormula(buildRosterTitleFormula(row));
      repaired++;
    }
  }

  if (repaired > 0) {
    console.log(`Repaired ${repaired} roster formula cell(s).`);
    SpreadsheetApp.flush();
  }

  return repaired;
}

/**
 * @param {string} formula A formula string read from the roster.
 * @returns {boolean} True if it still looks like the points total formula.
 */
function isPointsFormula(formula) {
  return typeof formula === 'string' && formula.toUpperCase().indexOf('SUMIFS(') !== -1;
}

/**
 * @param {string} formula A formula string read from the roster.
 * @returns {boolean} True if it still looks like the title lookup formula.
 */
function isTitleFormula(formula) {
  return typeof formula === 'string' && formula.toUpperCase().indexOf('VLOOKUP(') !== -1;
}

/**
 * Puts a warning-only protection on the roster's calculated columns so that clearing or typing
 * over them prompts the teacher first. Warning-only is deliberate: the owner can still make
 * deliberate changes, and the script can still write to the range.
 */
function protectRosterFormulaColumns() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);
  if (!rosterSheet) return;

  const description = "Mythos: calculated by formula - do not edit";

  // Drop any previous copy so repeated runs do not stack protections.
  rosterSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .filter(protection => protection.getDescription() === description)
      .forEach(protection => protection.remove());

  rosterSheet.getRange("D2:E")
      .protect()
      .setDescription(description)
      .setWarningOnly(true);
}

/**
 * Menu entry point for the roster repair so the teacher can fix the roster
 * after pasting in a new class list or accidentally clearing a column.
 */
function repairRosterFormulas() {
  const updated = refreshRosterFormulas();
  if (updated === 0) {
    SpreadsheetApp.getUi().alert("No students found on the roster.");
    return;
  }

  try {
    protectRosterFormulaColumns();
  } catch (error) {
    console.log("Could not protect roster formula columns:", error);
  }

  SpreadsheetApp.getUi().alert(
    `Recalculation formulas restored for ${updated} student(s).\n\n` +
    "Columns D and E are now marked as protected, so editing them will show a warning first. " +
    "They are also repaired automatically whenever the sheet is opened or a student submits."
  );
}

/**
 * Builds the SUMIFS formula that totals a student's verified points.
 * @param {number} rosterRow The roster row (1-indexed) the formula belongs to.
 */
function buildRosterPointsFormula(rosterRow) {
  const submissions = "'" + STUDENT_SUBMISSIONS_SHEET_NAME + "'!";
  return "=IFERROR(SUMIFS(" + submissions + "$I$2:$I," +
         submissions + "$B$2:$B,$B" + rosterRow + "," +
         submissions + "$J$2:$J,TRUE),0)";
}

/**
 * Builds the VLOOKUP formula that turns a point total into a title.
 * The lookup range is bounded to the numeric title rows: the "Journey Settings" sheet also holds
 * text-keyed system settings below them, and an approximate-match VLOOKUP over a column that mixes
 * text and numbers is not sorted ascending, so it can silently return the wrong title.
 * @param {number} rosterRow The roster row (1-indexed) the formula belongs to.
 */
function buildRosterTitleFormula(rosterRow) {
  const lastTitleRow = getLastTitleRow();
  return "=IFERROR(VLOOKUP($D" + rosterRow + ",'" + JOURNEY_SETTINGS_SHEET_NAME +
         "'!$A$2:$B$" + lastTitleRow + ",2,TRUE),\"Gnome\")";
}

/**
 * Finds the last row of the "Journey Settings" sheet that holds a title (a numeric point threshold
 * in column A). Everything below that is system settings.
 * @returns {number} The 1-indexed last title row, or 1 if there are no titles.
 */
function getLastTitleRow() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);

  if (!settingsSheet || settingsSheet.getLastRow() < 2) return 1;

  const columnA = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 1).getValues();
  let lastTitleRow = 1;
  for (let i = 0; i < columnA.length; i++) {
    if (typeof columnA[i][0] === 'number') {
      lastTitleRow = i + 1; // Convert the 0-indexed array position to a 1-indexed row.
    }
  }

  return lastTitleRow;
}

/**
 * Reads the title thresholds from the "Journey Settings" sheet, ignoring the system settings rows.
 * @returns {Array<Array>} Rows of [points, title, message, imageUrl].
 */
function getTitlesData() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
  const lastTitleRow = getLastTitleRow();

  if (!settingsSheet || lastTitleRow < 2) return [];

  return settingsSheet.getRange(2, 1, lastTitleRow - 1, 4).getValues();
}

/**
 * Installs an installable on-edit trigger.
 *
 * The simple onEdit(e) trigger runs unauthorized, so it can update points but can never call
 * MailApp - which is why checking the "Teacher Verified?" box never emailed the student. An
 * installable trigger runs with authorization and can.
 */
function installVerificationTrigger() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Remove any previously installed copy so verifying a submission cannot send two emails.
  ScriptApp.getProjectTriggers()
      .filter(trigger => trigger.getHandlerFunction() === 'onVerificationEdit')
      .forEach(trigger => ScriptApp.deleteTrigger(trigger));

  ScriptApp.newTrigger('onVerificationEdit')
      .forSpreadsheet(spreadsheet)
      .onEdit()
      .create();

  ui.alert("Verification emails are now enabled. Students will be emailed when you check their \"Teacher Verified?\" box.");
}

/**
 * Automatically triggers when a cell in the spreadsheet is edited.
 * Handles the teacher checking the "Verified" checkbox in the Student Submissions sheet.
 *
 * This is the simple trigger, which runs without authorization: it awards the points but cannot
 * send email. Run "Install Verification Email Trigger" from the Mythos Admin menu to also notify
 * the student.
 */
function onEdit(e) {
  handleVerificationEdit(e, true);
}

/**
 * Installed by installVerificationTrigger(). Same job as onEdit(), but authorized, so it also
 * sends the student their confirmation email.
 */
function onVerificationEdit(e) {
  handleVerificationEdit(e, false);
}

/**
 * Shared handler for the simple and installable edit triggers.
 * @param {Object} e The edit event.
 * @param {boolean} skipEmail Whether to suppress the confirmation email.
 */
function handleVerificationEdit(e, skipEmail) {
  // Guard against a missing event object, which happens if the function is run manually.
  if (!e || !e.range) return;

  const range = e.range;
  const sheet = range.getSheet();

  // Only process single-cell edits in the Student Submissions sheet, Column J (Verified status)
  if (sheet.getName() !== STUDENT_SUBMISSIONS_SHEET_NAME) return;
  if (range.getColumn() !== 10 || range.getNumColumns() !== 1) return;
  if (range.getRow() < 2) return;

  // If the checkbox was checked (TRUE)
  if (range.getValue() === true) {
    verifySubmission(range.getRow(), skipEmail);
  }
}

/**
 * Helper function to verify all pending submissions at once.
 */
function batchVerifyPending() {
  const pending = getPendingSubmissions();
  if (pending.length === 0) {
    SpreadsheetApp.getUi().alert("No pending submissions found.");
    return;
  }
  
  const confirm = SpreadsheetApp.getUi().alert(
    "Verify All",
    `Are you sure you want to verify all ${pending.length} pending submissions?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );
  
  if (confirm === SpreadsheetApp.getUi().Button.YES) {
    const rows = pending.map(p => p.rowNumber);
    verifyMultipleSubmissions(rows);
    SpreadsheetApp.getUi().alert(`Successfully verified ${pending.length} submissions.`);
  }
}

/**
 * Serves the HTML file for the web app.
 * This function is automatically called when a user visits the web app URL.
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Mythos Ascendant')
      .setFaviconUrl('https://img.icons8.com/color/48/000000/mythology.png');
}

/**
 * Gets the email of the user accessing the web app.
 * Since the app is deployed "as me" but "only for students in our domain",
 * this will return the student's email.
 */
function getUserEmail() {
  return Session.getActiveUser().getEmail();
}

/**
 * Gets the main logo URL for the web app.
 * This function is called by JavaScript in the web app.
 */
function getMainLogoUrl() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
  
  try {
    if (settingsSheet && settingsSheet.getLastRow() > 1) {
      // Find the row with "Main Logo" in column A and get the value from column B
      const allSettingsData = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
      const mainLogoRow = allSettingsData.find(row => row[0] === "Main Logo");
      if (mainLogoRow && mainLogoRow[1]) {
        const rawMainLogoUrl = mainLogoRow[1];
        console.log("Raw main logo URL from settings:", rawMainLogoUrl);
        
        // Process the URL using the existing helper function
        const processedMainLogoUrl = getPublicUrl(rawMainLogoUrl);
        if (validateImageUrl(processedMainLogoUrl)) {
          console.log("Using processed main logo URL:", processedMainLogoUrl);
          return processedMainLogoUrl;
        } else {
          console.log("Main logo URL failed validation");
        }
      }
    }
  } catch (error) {
    console.log("Error fetching main logo URL for web app:", error);
  }
  
  return ""; // Empty string means no logo
}

/**
 * Initializes the spreadsheet with the necessary tabs and headers.
 * This should be run manually once before deploying the web app.
 */
function setupMythosSheets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Data for the Journey Settings tab
  const titlesData = [
    { points: 0, title: "Gnome", message: "Congratulations! Your journey has begun! As a Gnome, you are a small, earth-dwelling spirit, and your adventure is just starting to take root.", imageUrl: "" },
    { points: 20, title: "Gremlin", message: "Congratulations! You've earned the title of Gremlin. Your mischievous nature and ability to cause minor disruptions are making an impact.", imageUrl: "" },
    { points: 30, title: "Kobold", message: "Congratulations! You have achieved the title of Kobold. Like this small, house-dwelling spirit, you are showing your presence and building your influence.", imageUrl: "" },
    { points: 35, title: "Dryad", message: "Congratulations! For reaching 35 points, you are now a Dryad. Your connection to your environment and ability to grow stronger are becoming apparent.", imageUrl: "" },
    { points: 38, title: "Satyr", message: "Congratulations! You've earned the title of Satyr. Your playful, half-goat nature is now recognized, a sign of your spirited approach to the game.", imageUrl: "" },
    { points: 40, title: "Gorgon", message: "Congratulations! You've reached 40 points and are now a Gorgon. While a monstrous being, you are showing your power and ability to freeze your opponents in their tracks.", imageUrl: "" },
    { points: 42, title: "The Answer to the Ultimate Question", message: "You have achieved the ultimate answer of 42 points and earned the title of The Answer to the Ultimate Question. Be sure to never occupy the same universe as the Ultimate Question.", imageUrl: "" },
    { points: 45, title: "Griffin", message: "Congratulations! For reaching 45 points, you are now a Griffin. Your powerful physical presence and dominance are becoming undeniable.", imageUrl: "" },
    { points: 48, title: "Minotaur", message: "Congratulations! You have achieved the title of Minotaur. Like this strong, formidable beast, you're a force to be reckoned with in the labyrinth of challenges.", imageUrl: "" },
    { points: 50, title: "The Sphinx", message: "Congratulations! You've earned the ultimate title of The Sphinx. Your intelligence and ability to outsmart your opponents are now your greatest weapons.", imageUrl: "" },
    { points: 52, title: "Hydra", message: "Congratulations! You've reached 52 points and are now a Hydra. Your ability to regenerate and bounce back from challenges is unmatched.", imageUrl: "" },
    { points: 55, title: "Fenrir", message: "Congratulations! You have achieved the title of Fenrir. A powerful, giant wolf, you are feared by your opponents and are poised to challenge even the strongest.", imageUrl: "" },
    { points: 58, title: "Valkyrie", message: "Congratulations! For reaching 58 points, you are now a Valkyrie. Your prowess in battle is a sight to behold, guiding the fallen and proving your dominance.", imageUrl: "" },
    { points: 60, title: "The Chimera", message: "Congratulations! You have earned the title of The Chimera. Your diverse skills and abilities are blending together into something truly monstrous and unique.", imageUrl: "" },
    { points: 62, title: "The Kraken", message: "Congratulations! You've reached 62 points and are now known as The Kraken. Your influence is growing, and your power can be felt across the entire game.", imageUrl: "" },
    { points: 65, title: "Dragon", message: "Congratulations! With 65 points, you have reached a new level of power and earned the legendary title of Dragon. You are an awe-inspiring force of nature, a creature of myth and legend, whose might is known throughout the land.", imageUrl: "" },
    { points: 68, title: "The Djinn", message: "Congratulations! You have earned the title of The Djinn. Your control over magic and your reality-bending skills are truly powerful.", imageUrl: "" },
    { points: 70, title: "Anubis", message: "Congratulations! With 70 points, you are now known as Anubis. Your mastery of the darkest parts of the game and your ability to guide others through the unknown is unmatched.", imageUrl: "" },
    { points: 75, title: "Hel", message: "Congratulations! You have earned the title of Hel. Like the ruler of the underworld, you hold absolute power over those who have been defeated.", imageUrl: "" },
    { points: 80, title: "Odin", message: "Congratulations! For reaching 80 points, you are now Odin. Your wisdom, command, and ability to see all make you a true leader and a god among men.", imageUrl: "" },
    { points: 85, title: "Shiva", message: "Congratulations! You have achieved the title of Shiva the Destroyer. You are a supreme force of destruction and transformation, changing the game with your every move.", imageUrl: "" },
    { points: 90, title: "Amaterasu", message: "Congratulations! For reaching 90 points, you have achieved the divine title of Amaterasu. Like the supreme sun goddess, your influence is a source of ultimate life and power, illuminating all who cross your path", imageUrl: "" },
    { points: 95, title: "Zeus", message: "Congratulations! You've earned the ultimate title of Zeus, King of Olympus. You command the sky, and your power over all aspects of the game is undeniable.", imageUrl: "" },
    { points: 100, title: "Chaos", message: "Congratulations! You've reached the pinnacle with 100 points and earned the ultimate title of Chaos. You are the primordial force, the beginning and the end of all things. Your dominance is complete.", imageUrl: "" }
  ];

  // Set up the Journey Settings tab
  let settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
  if (!settingsSheet) {
    settingsSheet = spreadsheet.insertSheet(JOURNEY_SETTINGS_SHEET_NAME, 0);
  }
  settingsSheet.clear();
  const settingsHeaders = ["Points", "Title", "Congratulations Message", "Image URL"];
  settingsSheet.getRange(1, 1, 1, settingsHeaders.length).setValues([settingsHeaders]).setFontWeight("bold");
  const settingsData = titlesData.map(row => [row.points, row.title, row.message, row.imageUrl]);
  settingsSheet.getRange(2, 1, settingsData.length, settingsData[0].length).setValues(settingsData);
  
  // Add a verification toggle
  settingsSheet.getRange(titlesData.length + 4, 1, 1, 2).setValues([["System Setting", "Value"]]).setFontWeight("bold");
  settingsSheet.getRange(titlesData.length + 5, 1, 1, 2).setValues([["Enable Teacher Verification", "TRUE"]]);
  settingsSheet.getRange(titlesData.length + 6, 1, 1, 2).setValues([["Main Logo", ""]]);

  // Set up the Student Roster tab
  let rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);
  if (!rosterSheet) {
    rosterSheet = spreadsheet.insertSheet(STUDENT_ROSTER_SHEET_NAME, 1);
  }
  rosterSheet.clear();
  const rosterHeaders = ["Student Name", "Student Email", "Class Period", "Total Points Earned", "Current Title Earned"];
  rosterSheet.getRange(1, 1, 1, rosterHeaders.length).setValues([rosterHeaders]).setFontWeight("bold");


  // Set up the Student Submissions tab
  let submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
  if (!submissionsSheet) {
    submissionsSheet = spreadsheet.insertSheet(STUDENT_SUBMISSIONS_SHEET_NAME, 2);
  }
  submissionsSheet.clear();
  const submissionsHeaders = ["Timestamp", "Student Email", "Type of Media", "Title of Media", "Bonus Points (Yes/No)", "Reflection: Date/Time", "Reflection: Mythological Connection", "Reflection: Analysis", "Points", "Teacher Verified?"];
  submissionsSheet.getRange(1, 1, 1, submissionsHeaders.length).setValues([submissionsHeaders]).setFontWeight("bold");

  // Apply the roster formulas to every student row that exists, and warn before they get edited.
  refreshRosterFormulas();
  try {
    protectRosterFormulaColumns();
  } catch (error) {
    console.log("Could not protect roster formula columns:", error);
  }

  SpreadsheetApp.getUi().alert("Mythos Ascendant sheets have been successfully set up!");
}

/**
 * Handles form submissions from the web app.
 * @param {Object} formData An object containing data from the form.
 */
function processSubmission(formData) {
  // The web app is deployed to run as the deploying user, so every student's submission executes
  // under the same account. Without a lock, a class submitting at once can interleave
  // getLastRow()/appendRow() and quietly overwrite each other's rows.
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(SUBMISSION_LOCK_TIMEOUT_MS);
  } catch (lockError) {
    console.log("Could not acquire submission lock:", lockError);
    return {
      status: "error",
      message: "The scrolls are busy with other students right now. Please wait a moment and submit again."
    };
  }

  try {
    return writeSubmission(formData);
  } catch (e) {
    console.log("Error processing submission:", e);
    return { status: "error", message: "An error occurred during submission: " + e.message };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Does the actual work of recording a submission. Always called while holding the submission lock.
 * @param {Object} formData An object containing data from the form.
 */
function writeSubmission(formData) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
  const submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
  const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);

  // Get system settings
  let enableVerification = "TRUE"; // Default
  try {
    const settingsData = settingsSheet.getDataRange().getValues();
    const verificationRow = settingsData.find(row => row[0] === "Enable Teacher Verification");
    if (verificationRow) {
      enableVerification = String(verificationRow[1]).toUpperCase();
    }
  } catch (e) {
    console.log("Error reading verification setting:", e);
  }
  
  // Get submission data. Everything is coerced and trimmed: the form can hand back undefined for
  // a radio group with nothing selected, and calling .trim() on that used to throw before a
  // single row was written.
  const asText = value => String(value === null || value === undefined ? "" : value).trim();

  // The email input is read-only once auto-populated, which exempts it from the browser's
  // "required" validation, so an empty one really can reach the server. Fall back to the signed-in
  // user rather than filing the submission under a blank name.
  let email = asText(formData.studentEmail);
  if (!email) {
    try {
      email = asText(Session.getActiveUser().getEmail());
    } catch (sessionError) {
      console.log("Could not read the active user's email:", sessionError);
    }
  }

  const mediaType = asText(formData.mediaType);
  const mediaTitle = asText(formData.mediaTitle);
  const bonusPoints = asText(formData.bonusPoints) === "Yes" ? "Yes" : "No";
  const reflectionDate = asText(formData.reflectionDate);
  const reflectionConnection = asText(formData.reflectionConnection);
  const reflectionAnalysis = asText(formData.reflectionAnalysis);

  // Reject incomplete submissions up front instead of recording an unusable row.
  const missingFields = [];
  if (!email) missingFields.push("Student Email");
  if (!mediaType) missingFields.push("Type of Media");
  if (!mediaTitle) missingFields.push("Title of Media");
  if (!reflectionDate) missingFields.push("Date and Time");
  if (!reflectionConnection) missingFields.push("Mythological Connection");
  if (!reflectionAnalysis) missingFields.push("Analysis");

  if (missingFields.length > 0) {
    return {
      status: "error",
      message: "Your offering is incomplete. Please fill in: " + missingFields.join(", ") + "."
    };
  }

  const normalizedEmail = normalizeEmail(email);

  // Guard against an accidental double submission (the same student sending the same title and
  // analysis twice within a few seconds).
  const lastRow = submissionsSheet.getLastRow();
  if (lastRow > 1) {
    // Scan the most recent rows INCLUDING the last one - the previous version stopped at
    // lastRow - 1, so it never looked at the row a double-click would just have created.
    const startRow = Math.max(2, lastRow - 4);
    const recentSubmissions = submissionsSheet.getRange(startRow, 1, lastRow - startRow + 1, 8).getValues();
    const now = Date.now();

    for (const row of recentSubmissions) {
      const rowTime = row[0] instanceof Date ? row[0].getTime() : NaN;
      if (isNaN(rowTime)) continue;

      // Only an age inside the window counts. The previous version tested `now - rowTime < 10000`
      // with no lower bound, so any row whose timestamp read as being in the future - which a
      // script/spreadsheet timezone mismatch will do - matched forever and silently swallowed
      // legitimate resubmissions.
      const age = now - rowTime;
      if (age < 0 || age > DUPLICATE_WINDOW_MS) continue;

      if (normalizeEmail(row[1]) === normalizedEmail &&
          asText(row[3]) === mediaTitle &&
          asText(row[7]) === reflectionAnalysis) {
        console.log("Duplicate submission detected and blocked for:", email);
        return {
          status: "success",
          message: "Submission received! (We ignored an accidental duplicate of the one you just sent.)"
        };
      }
    }
  }

  // Make sure the roster can actually total this student's points before we write anything.
  ensureRosterFormulas();

  // Fetch existing student info to check for level-up
  let rosterData = [];
  if (rosterSheet.getLastRow() > 1) {
    rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues();
  }
  // Compare normalized emails so a roster entry typed with different capitalization does not
  // silently create a second row for the same student.
  const studentRowIndex = rosterData.findIndex(row => normalizeEmail(row[1]) === normalizedEmail);

  let oldTitle = "Gnome";

  if (studentRowIndex === -1) {
      // New student, append a new row to the roster
      rosterSheet.appendRow(["", email, "", 0, "Gnome"]);

      const newRowNum = rosterSheet.getLastRow();
      rosterSheet.getRange(newRowNum, 4).setFormula(buildRosterPointsFormula(newRowNum));
      rosterSheet.getRange(newRowNum, 5).setFormula(buildRosterTitleFormula(newRowNum));

      // Force calculation of the new formulas
      SpreadsheetApp.flush();
  } else {
      // Existing student, get their current title so we can tell if this submission levels them up
      oldTitle = rosterData[studentRowIndex][4] || "Gnome";
  }

  // Count past submissions by this student for this media type
  let allSubmissions = [];
  if (submissionsSheet.getLastRow() > 1) {
      allSubmissions = submissionsSheet.getRange(2, 2, submissionsSheet.getLastRow() - 1, 2).getValues();
  }
  let submissionCount = 0;
  for (const row of allSubmissions) {
    if (normalizeEmail(row[0]) === normalizedEmail && asText(row[1]) === mediaType) {
      submissionCount++;
    }
  }

  // Calculate points based on submission count
  let points = 0;
  const mediaSettings = POINT_SYSTEM[mediaType] || POINT_SYSTEM['Other'];
  if (submissionCount === 0) {
    points = mediaSettings.first;
  } else if (submissionCount === 1) {
    points = mediaSettings.second;
  } else {
    points = mediaSettings.thirdPlus;
  }

  // Add bonus points if applicable
  if (bonusPoints === "Yes") {
    points += BONUS_POINTS_VALUE;
  }

  // Handle points based on verification setting
  let verificationStatus = false;

  if (enableVerification !== "TRUE") {
      verificationStatus = true; // Mark as verified so points are counted in Roster
  }

  // Write the submission back to the sheet - ALWAYS record the points for visibility
  const newRow = [new Date(), email, mediaType, mediaTitle, bonusPoints, reflectionDate, reflectionConnection, reflectionAnalysis, points, verificationStatus];
  submissionsSheet.appendRow(newRow);

  // Force the spreadsheet to recalculate all formulas
  SpreadsheetApp.flush();

  // Handle email sending based on verification setting.
  // The submission is already saved at this point, so the email is sent in its own try/catch:
  // a bounced address or an exhausted daily mail quota must never make a saved submission
  // report itself as failed and send the student off to submit all over again.
  if (enableVerification !== "TRUE") {
    try {
      const newRosterData = rosterSheet.getLastRow() > 1
          ? rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues()
          : [];
      const updatedRow = newRosterData.find(row => normalizeEmail(row[1]) === normalizedEmail);

      if (updatedRow) {
        sendConfirmationEmail(email, updatedRow[3], oldTitle, updatedRow[4]);
      } else {
        console.log("No roster row found for " + email + "; skipping confirmation email.");
      }
    } catch (emailError) {
      console.log("Error sending confirmation email:", emailError);
    }
  }
  // If verification is enabled, email will be sent when teacher verifies

  const successMessage = enableVerification !== "TRUE"
      ? "Submission received! An email has been sent to you with an update on your points."
      : "Submission received! Your submission is pending teacher verification.";
  return { status: "success", message: successMessage };
}

/**
 * Verifies a submission by moving pending points to calculated points.
 * @param {number} submissionRow The row number of the submission to verify (1-indexed).
 * @param {boolean} skipEmail Whether to skip sending the confirmation email (useful for bulk updates).
 */
function verifySubmission(submissionRow, skipEmail = false) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
    
    if (submissionRow < 2 || submissionRow > submissionsSheet.getLastRow()) {
      throw new Error("Invalid submission row number");
    }
    
    // Get the submission data (now 10 columns)
    const submissionData = submissionsSheet.getRange(submissionRow, 1, 1, 10).getValues()[0];
    const email = submissionData[1]; // Column B (Student Email)
    const mediaType = String(submissionData[2] || "").trim(); // Column C (Type of Media)
    const bonusPoints = submissionData[4]; // Column E (Bonus Points)
    const normalizedEmail = normalizeEmail(email);

    // Recalculate the points for this submission
    // Count past submissions by this student for this media type (only rows ABOVE this one)
    let submissionCount = 0;
    if (submissionRow > 2) {
        const priorSubmissions = submissionsSheet.getRange(2, 2, submissionRow - 2, 2).getValues();
        for (const row of priorSubmissions) {
          if (normalizeEmail(row[0]) === normalizedEmail && String(row[1] || "").trim() === mediaType) {
            submissionCount++;
          }
        }
    }

    // Calculate points based on submission count with safety fallback
    let points = 0;
    const mediaSettings = POINT_SYSTEM[mediaType] || POINT_SYSTEM['Other'];
    if (submissionCount === 0) {
      points = mediaSettings.first;
    } else if (submissionCount === 1) {
      points = mediaSettings.second;
    } else {
      points = mediaSettings.thirdPlus;
    }

    // Add bonus points if applicable
    if (bonusPoints === "Yes") {
      points += BONUS_POINTS_VALUE;
    }
    
    // Set the calculated points and mark as verified
    submissionsSheet.getRange(submissionRow, 9).setValue(points); // Set Points (Column I)
    submissionsSheet.getRange(submissionRow, 10).setValue(true); // Set checkbox to checked (Column J)

    // The roster only picks these points up through its formulas, so make sure they are intact.
    ensureRosterFormulas();

    // Force recalculation
    SpreadsheetApp.flush();

    // Send confirmation email to student about their points update.
    // skipEmail is honoured here: recalculateAllSubmissions() walks every row, and without this
    // a single recalculation would email the whole class once per submission and blow through the
    // account's daily mail quota.
    if (!skipEmail) {
      try {
        const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);

        // Get the student's current roster data
        const rosterData = rosterSheet.getLastRow() > 1
            ? rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues()
            : [];
        const studentRow = rosterData.find(row => normalizeEmail(row[1]) === normalizedEmail);

        if (studentRow) {
            // For verification emails, try to determine if this is a level up by checking point ranges
            // Get the previous point total (current minus the points just added)
            const newTotalPoints = studentRow[3];
            const newTitle = studentRow[4];
            const previousPoints = Math.max(0, newTotalPoints - points);

            // Find what title they would have had with previous points
            let oldTitle = "Gnome";
            for (const row of getTitlesData()) {
                if (previousPoints >= row[0]) {
                    oldTitle = row[1];
                }
            }

            console.log(`Verification email: Previous points: ${previousPoints}, Old title: ${oldTitle}, New points: ${newTotalPoints}, New title: ${newTitle}`);
            sendConfirmationEmail(email, newTotalPoints, oldTitle, newTitle);
        } else {
            console.log("No roster row found for " + email + "; skipping verification email.");
        }
      } catch (emailError) {
        console.log("Error sending verification email:", emailError);
        // Don't fail the verification if email fails
      }
    }

    return {
      status: "success",
      message: skipEmail
          ? "Submission verified successfully."
          : "Submission verified successfully and student has been notified."
    };
  } catch (e) {
    return { status: "error", message: "Error verifying submission: " + e.message };
  }
}

/**
 * Verifies multiple submissions at once.
 * @param {Array} submissionRows An array of row numbers to verify.
 */
function verifyMultipleSubmissions(submissionRows) {
  const results = [];
  
  for (const rowNum of submissionRows) {
    const result = verifySubmission(rowNum);
    results.push({ row: rowNum, result: result });
  }
  
  return results;
}

/**
 * Gets all pending submissions for teacher review.
 */
function getPendingSubmissions() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
  
  if (submissionsSheet.getLastRow() < 2) {
    return [];
  }
  
  const allData = submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, 10).getValues();
  const pendingSubmissions = [];
  
  allData.forEach((row, index) => {
    if (row[9] === false) { // Column J is checkbox - false means unchecked/pending
      pendingSubmissions.push({
        rowNumber: index + 2, // +2 because we started from row 2 and arrays are 0-indexed
        timestamp: row[0],
        studentEmail: row[1],
        mediaType: row[2],
        mediaTitle: row[3],
        bonusPoints: row[4],
        reflectionDate: row[5],
        reflectionConnection: row[6],
        reflectionAnalysis: row[7],
        currentPoints: row[8], // Provisional points; they only count once column J is TRUE
        status: row[9] // false = pending, true = verified
      });
    }
  });
  
  return pendingSubmissions;
}

/**
 * Gets unique class periods from the Student Roster.
 */
function getClassPeriods() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);
  
  if (rosterSheet.getLastRow() < 2) {
    return [];
  }
  
  const classPeriodData = rosterSheet.getRange(2, 3, rosterSheet.getLastRow() - 1, 1).getValues();
  const uniquePeriods = [...new Set(classPeriodData.flat().filter(period => period !== ""))];
  
  return uniquePeriods.sort();
}

/**
 * Retrieves student data for the leaderboard, sorted by points.
 * @param {string} classPeriodFilter Optional class period to filter by.
 */
function getLeaderboardData(classPeriodFilter) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);

  // Get all students and their points from the roster sheet
  let rosterData = [];
  if (rosterSheet.getLastRow() > 1) {
    rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues();
  }

  // Debug logging
  console.log("Roster data:", rosterData);
  console.log("Number of students found:", rosterData.length);
  console.log("Class period filter received:", classPeriodFilter);
  console.log("Filter type:", typeof classPeriodFilter);

  // Map roster data to a more usable format
  let students = rosterData
    .map(row => ({
      name: row[0],
      email: row[1], 
      classPeriod: row[2],
      points: Number(row[3]) || 0, // Convert to number, default to 0
      title: row[4]
    }));
    
  console.log("Students before filtering:", students.map(s => ({
    name: s.name || 'Empty',
    email: s.email || 'Empty', 
    classPeriod: s.classPeriod,
    classPeriodType: typeof s.classPeriod
  })));
    
  // Apply class period filter if specified
  let filteredStudents = students;
  if (classPeriodFilter && 
      classPeriodFilter !== "" && 
      classPeriodFilter !== "All Classes" && 
      classPeriodFilter !== null && 
      classPeriodFilter !== undefined) {
    console.log("Applying class period filter:", classPeriodFilter);
    const beforeFilter = students.length;
    
    // Filter students, but handle empty class periods gracefully
    filteredStudents = students.filter(student => {
      const studentPeriod = String(student.classPeriod).trim();
      const filterPeriod = String(classPeriodFilter).trim();
      const matches = studentPeriod === filterPeriod;
      
      if (!matches) {
        console.log(`Student ${student.name || student.email || 'Unknown'}: period '${studentPeriod}' != filter '${filterPeriod}'`);
      }
      
      return matches;
    });
    
    console.log(`Filter applied: ${beforeFilter} -> ${filteredStudents.length} students`);
  } else {
    console.log("No class period filter applied (showing all classes)");
    console.log("Filter reason:", !classPeriodFilter ? "No filter provided" : 
                classPeriodFilter === "All Classes" ? "All Classes selected" :
                "Filter is null/empty");
  }
  
  // Sort by points
  const sortedStudents = filteredStudents.sort((a, b) => b.points - a.points);
  
  console.log("Filtered students:", sortedStudents);
  console.log("Number of students after filtering:", sortedStudents.length);

  // Map to a cleaner format for the frontend
  const leaderboard = sortedStudents.map((student, index) => {
    return {
      rank: index + 1,
      name: student.name || (student.email ? student.email.split('@')[0] : 'Student ' + (index + 1)),
      email: student.email,
      classPeriod: student.classPeriod,
      points: student.points,
      title: student.title || 'Gnome' // Use existing title, fallback to Gnome
    };
  });

  console.log("=== DEBUG: Final leaderboard being returned ===");
  console.log("Leaderboard length:", leaderboard.length);
  console.log("Leaderboard data:", leaderboard);
  console.log("=== END DEBUG ===");

  return leaderboard;
}

/**
 * Helper function to convert a Google Drive URL to an embeddable image URL.
 * @param {string} url The original Google Drive URL.
 * @returns {string} The embeddable URL.
 */
function getPublicUrl(url) {
  console.log("Processing image URL:", url);
  
  if (!url || typeof url !== 'string') {
    console.log("Invalid URL provided, using fallback");
    return "https://placehold.co/150x150?text=Badge";
  }
  
  try {
    let fileId = null;
    
    // Handle different Google Drive URL formats
    if (url.includes("/file/d/")) {
      // Format: https://drive.google.com/file/d/FILE_ID/view?usp=sharing
      const match = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
      if (match) {
        fileId = match[1];
      }
    } else if (url.includes("/open?id=")) {
      // Format: https://drive.google.com/open?id=FILE_ID
      const match = url.match(/id=([a-zA-Z0-9_-]+)/);
      if (match) {
        fileId = match[1];
      }
    } else if (url.includes("drive.google.com") && url.includes("/d/")) {
      // Legacy format or other variations with /d/
      try {
        fileId = url.split("/d/")[1].split("/")[0];
      } catch (e) {
        console.log("Failed to parse legacy format:", e);
      }
    }
    
    if (fileId) {
      const convertedUrl = `https://drive.google.com/uc?export=view&id=${fileId}`;
      console.log("Converted URL:", convertedUrl);
      return convertedUrl;
    } else {
      console.log("No file ID found, checking if URL is already direct");
      // If it's already a direct image URL or uc format, return as-is
      if (url.includes("drive.google.com/uc") || url.match(/\.(jpg|jpeg|png|gif|webp)$/i)) {
        return url;
      } else {
        console.log("Unknown URL format, using fallback");
        return "https://placehold.co/150x150?text=Badge";
      }
    }
  } catch (error) {
    console.log("Error processing URL:", error);
    return "https://placehold.co/150x150?text=Badge";
  }
}

/**
 * Validates if an image URL is accessible.
 * @param {string} url The image URL to validate.
 * @returns {boolean} True if the URL appears to be valid.
 */
function validateImageUrl(url) {
  try {
    // Basic URL validation
    if (!url || typeof url !== 'string') {
      return false;
    }
    
    // Check if it looks like a valid URL
    if (!url.startsWith('http://') && !url.startsWith('https://')) {
      return false;
    }
    
    // For Google Drive URLs, we can't easily test accessibility without making a request,
    // but we can validate the format
    if (url.includes('drive.google.com')) {
      return url.includes('uc?export=view&id=') || 
             url.includes('/file/d/') || 
             url.includes('/open?id=');
    }
    
    // For other URLs, check if they look like image URLs
    return url.match(/\.(jpg|jpeg|png|gif|webp)$/i) || 
           url.includes('placehold.co') ||
           url.includes('placeholder');
           
  } catch (error) {
    console.log("Error validating image URL:", error);
    return false;
  }
}

/**
 * Debug function to test image URL processing for all titles.
 * Call this manually to check if image URLs are working.
 */
function debugImageUrls() {
  console.log("=== DEBUG: Testing all image URLs ===");
  
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
  
  if (settingsSheet.getLastRow() <= 1) {
    console.log("No title data found in Journey Settings");
    return;
  }
  
  const titlesData = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 1, 4).getValues();
  console.log("Found", titlesData.length, "titles to test");
  
  titlesData.forEach((row, index) => {
    const points = row[0];
    const title = row[1];
    const message = row[2];
    const rawImageUrl = row[3];
    
    console.log(`\n--- Testing title ${index + 1}: ${title} ---`);
    console.log("Raw URL:", rawImageUrl);
    
    if (!rawImageUrl) {
      console.log("❌ No image URL provided");
      return;
    }
    
    const convertedUrl = getPublicUrl(rawImageUrl);
    console.log("Converted URL:", convertedUrl);
    
    const isValid = validateImageUrl(convertedUrl);
    console.log("Validation result:", isValid ? "✅ VALID" : "❌ INVALID");
    
    if (convertedUrl.includes("placehold.co")) {
      console.log("⚠️ Using fallback placeholder");
    }
  });
  
  console.log("\n=== DEBUG: Complete ===");
}

/**
 * Sends a confirmation email to the student with their new total and title.
 * @param {string} recipientEmail The student's email address.
 * @param {number} newTotalPoints The student's updated point total.
 * @param {string} oldTitle The student's previous title.
 * @param {string} newTitle The student's current title.
 */
function sendConfirmationEmail(recipientEmail, newTotalPoints, oldTitle, newTitle) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
    
    // Get titles data from the spreadsheet
    let titlesData = [];
    let mainLogoUrl = "https://img.icons8.com/color/96/000000/mythology.png"; // Default fallback
    
    if (settingsSheet.getLastRow() > 1) {
      // Read the title rows by detecting where they end rather than assuming the system settings
      // occupy exactly the last 6 rows - adding or removing a setting used to silently drop or
      // include the wrong rows here.
      titlesData = getTitlesData();

      // Get the main logo URL from the settings
      try {
        // Find the row with "Main Logo" in column A and get the value from column B
        const allSettingsData = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
        const mainLogoRow = allSettingsData.find(row => row[0] === "Main Logo");
        if (mainLogoRow && mainLogoRow[1]) {
          const rawMainLogoUrl = mainLogoRow[1];
          console.log("Raw main logo URL from settings:", rawMainLogoUrl);
          
          // Process the URL using the existing helper function
          const processedMainLogoUrl = getPublicUrl(rawMainLogoUrl);
          if (validateImageUrl(processedMainLogoUrl)) {
            mainLogoUrl = processedMainLogoUrl;
            console.log("Using processed main logo URL:", mainLogoUrl);
          } else {
            console.log("Main logo URL failed validation, using default");
          }
        }
      } catch (error) {
        console.log("Error fetching main logo URL:", error);
      }
    }
    
    const titlesMap = new Map(titlesData.map(row => [row[1], { message: row[2], imageUrl: row[3] }]));

    let levelUpMessage = `You've earned new points! Your total is now ${newTotalPoints}. Keep going to reach the next title: ${newTitle}.`;
    let badgeImageUrl = "https://placehold.co/100x100?text=Points";
    let isLevelUp = oldTitle !== newTitle;
    
    // Always get badge for current title, regardless of level-up status
    console.log("=== BADGE IMAGE PROCESSING DEBUG ===");
    console.log("Processing badge image for title:", newTitle);
    console.log("Is level up:", isLevelUp);
    console.log("Old title:", oldTitle, "| New title:", newTitle);
    
    if (titlesMap.has(newTitle)) {
        const titleInfo = titlesMap.get(newTitle);
        
        // Only update message if it's a level up
        if (isLevelUp) {
            levelUpMessage = titleInfo.message;
        }
        
        console.log("Title info from map:", titleInfo);
        console.log("Raw image URL from settings:", titleInfo.imageUrl);
        console.log("Image URL type:", typeof titleInfo.imageUrl);
        console.log("Image URL length:", titleInfo.imageUrl ? titleInfo.imageUrl.length : 'null/undefined');
        
        try {
            // Use the helper function to convert the image URL
            let processedImageUrl = getPublicUrl(titleInfo.imageUrl);
            console.log("Processed image URL:", processedImageUrl);
            
            // Validate the processed URL
            if (validateImageUrl(processedImageUrl)) {
                badgeImageUrl = processedImageUrl;
                console.log("✓ Using processed image URL:", badgeImageUrl);
            } else {
                console.log("✗ Processed URL failed validation, using fallback");
                badgeImageUrl = "https://placehold.co/150x150?text=" + encodeURIComponent(newTitle);
            }
        } catch (error) {
            console.log("✗ Error processing image URL:", error);
            badgeImageUrl = "https://placehold.co/150x150?text=" + encodeURIComponent(newTitle);
        }
    } else {
        // No title info found in map
        console.log("No title info found for:", newTitle);
        badgeImageUrl = "https://placehold.co/150x150?text=" + encodeURIComponent(newTitle);
    }
    
    console.log("Final badge image URL:", badgeImageUrl);
    console.log("=== END BADGE IMAGE PROCESSING DEBUG ===");

    const template = HtmlService.createTemplateFromFile('Email');
    template.isLevelUp = isLevelUp;
    template.newPoints = newTotalPoints;
    template.newTitle = newTitle;
    template.levelUpMessage = levelUpMessage;
    template.badgeImageUrl = badgeImageUrl;
    template.mainLogoUrl = mainLogoUrl;

    const htmlBody = template.evaluate().getContent();

    MailApp.sendEmail({
        to: recipientEmail,
        subject: `Mythos Ascendant: Your Journey Update!`,
        htmlBody: htmlBody
    });
}
