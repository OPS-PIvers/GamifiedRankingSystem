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
  rosterSheet.getRange('D2').setFormula("=IFERROR(SUMIFS('" + STUDENT_SUBMISSIONS_SHEET_NAME + "'!G:G,'" + STUDENT_SUBMISSIONS_SHEET_NAME + "'!B:B,B2,'" + STUDENT_SUBMISSIONS_SHEET_NAME + "'!H:H,TRUE),0)");
  rosterSheet.getRange('E2').setFormula("=IFERROR(VLOOKUP(D2,'" + JOURNEY_SETTINGS_SHEET_NAME + "'!A:B,2,TRUE),\"Gnome\")");


  // Set up the Student Submissions tab
  let submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
  if (!submissionsSheet) {
    submissionsSheet = spreadsheet.insertSheet(STUDENT_SUBMISSIONS_SHEET_NAME, 2);
  }
  submissionsSheet.clear();
  const submissionsHeaders = ["Timestamp", "Student Email", "Type of Media", "Title of Media", "Bonus Points (Yes/No)", "Reflection", "Points", "Teacher Verified?"];
  submissionsSheet.getRange(1, 1, 1, submissionsHeaders.length).setValues([submissionsHeaders]).setFontWeight("bold");

  SpreadsheetApp.getUi().alert("Mythos Ascendant sheets have been successfully set up!");
}

/**
 * Handles form submissions from the web app.
 * @param {Object} formData An object containing data from the form.
 */
function processSubmission(formData) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const settingsSheet = spreadsheet.getSheetByName(JOURNEY_SETTINGS_SHEET_NAME);
    const submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
    const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);

    // Get system settings
    const verificationCell = settingsSheet.getRange(settingsSheet.getLastRow(), 2);
    const enableVerification = verificationCell.getValue();
    
    // Get submission data
    const email = formData.studentEmail.trim();
    const mediaType = formData.mediaType;
    const mediaTitle = formData.mediaTitle;
    const bonusPoints = formData.bonusPoints;
    const reflection = formData.reflection;

    // Fetch existing student info to check for level-up
    let rosterData = [];
    if (rosterSheet.getLastRow() > 1) {
      rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues();
    }
    const studentRowIndex = rosterData.findIndex(row => row[1] === email);
    
    let oldPoints = 0;
    let oldTitle = "Gnome";
    let rosterRow = [];

    if (studentRowIndex === -1) {
        // New student, append a new row to the roster
        rosterRow = ["", email, "", 0, "Gnome"];
        rosterSheet.appendRow(rosterRow);
        
        // Get the new row number and copy formulas
        const newRowNum = rosterSheet.getLastRow();
        
        // Copy the SUMIF formula for total points (Column D)
        const pointsFormula = "=IFERROR(SUMIF('" + STUDENT_SUBMISSIONS_SHEET_NAME + "'!B:B,B" + newRowNum + ",'" + STUDENT_SUBMISSIONS_SHEET_NAME + "'!G:G),0)";
        rosterSheet.getRange('D' + newRowNum).setFormula(pointsFormula);
        
        // Copy the VLOOKUP formula for title (Column E)  
        const titleFormula = "=IFERROR(VLOOKUP(D" + newRowNum + ",'" + JOURNEY_SETTINGS_SHEET_NAME + "'!A:B,2,TRUE),\"Gnome\")";
        rosterSheet.getRange('E' + newRowNum).setFormula(titleFormula);
        
        // Force calculation of the new formulas
        SpreadsheetApp.flush();
    } else {
        // Existing student, get their current stats
        rosterRow = rosterSheet.getRange(studentRowIndex + 2, 1, 1, 5).getValues()[0];
        oldPoints = rosterRow[3];
        oldTitle = rosterRow[4];
    }
    
    // Count past submissions by this student for this media type
    let allSubmissions = [];
    if (submissionsSheet.getLastRow() > 1) {
        allSubmissions = submissionsSheet.getRange(2, 2, submissionsSheet.getLastRow() - 1, 3).getValues();
    }
    let submissionCount = 0;
    for (const row of allSubmissions) {
      if (row[0] === email && row[1] === mediaType) {
        submissionCount++;
      }
    }

    // Calculate points based on submission count
    let points = 0;
    if (submissionCount === 0) {
      points = POINT_SYSTEM[mediaType].first;
    } else if (submissionCount === 1) {
      points = POINT_SYSTEM[mediaType].second;
    } else {
      points = POINT_SYSTEM[mediaType].thirdPlus;
    }

    // Add bonus points if applicable
    if (bonusPoints === "Yes") {
      points += BONUS_POINTS_VALUE;
    }
    
    // Handle points based on verification setting
    let earnedPoints = 0;
    let verificationStatus = false; // ALWAYS start false (unchecked checkbox)
    
    if (enableVerification !== "TRUE") {
        earnedPoints = points; // Award points immediately when verification is disabled
        // verificationStatus stays false - checkbox unchecked but points already awarded
    } else {
        earnedPoints = 0; // No points until teacher verifies by checking the box
        // verificationStatus stays false - teacher must check box to award points
    }

    // Write the submission back to the sheet
    const newRow = [new Date(), email, mediaType, mediaTitle, bonusPoints, reflection, earnedPoints, verificationStatus];
    submissionsSheet.appendRow(newRow);

    // Force the spreadsheet to recalculate all formulas
    SpreadsheetApp.flush();

    // Handle email sending based on verification setting
    if (enableVerification !== "TRUE") {
        // Verification disabled - send email immediately with updated points
        // Force the spreadsheet to recalculate all formulas first
        SpreadsheetApp.flush();
        
        // Get the fresh, updated data from the spreadsheet
        const newRosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues();
        const newStudentRowIndex = newRosterData.findIndex(row => row[1] === email);
        const newRosterRow = newRosterData[newStudentRowIndex];
        const newTotalPoints = newRosterRow[3];
        const newTitle = newRosterRow[4];

        // Send confirmation email
        sendConfirmationEmail(email, newTotalPoints, oldTitle, newTitle);
    }
    // If verification is enabled, email will be sent when teacher verifies

    const successMessage = enableVerification !== "TRUE" 
        ? "Submission received! An email has been sent to you with an update on your points."
        : "Submission received! Your submission is pending teacher verification.";
    return { status: "success", message: successMessage };
  } catch (e) {
    return { status: "error", message: "An error occurred during submission: " + e.message };
  }
}

/**
 * Verifies a submission by moving pending points to calculated points.
 * @param {number} submissionRow The row number of the submission to verify (1-indexed).
 */
function verifySubmission(submissionRow) {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const submissionsSheet = spreadsheet.getSheetByName(STUDENT_SUBMISSIONS_SHEET_NAME);
    
    if (submissionRow < 2 || submissionRow > submissionsSheet.getLastRow()) {
      throw new Error("Invalid submission row number");
    }
    
    // Get the submission data (now only 8 columns)
    const submissionData = submissionsSheet.getRange(submissionRow, 1, 1, 8).getValues()[0];
    const email = submissionData[1]; // Column B (Student Email)
    const mediaType = submissionData[2]; // Column C (Type of Media)
    const bonusPoints = submissionData[4]; // Column E (Bonus Points)
    const currentPoints = submissionData[6]; // Column G (Points)
    const currentStatus = submissionData[7]; // Column H (Teacher Verified?)
    
    if (currentStatus === true) {
      throw new Error("Submission is already verified");
    }
    
    // Recalculate the points for this submission
    // Count past submissions by this student for this media type
    let allSubmissions = [];
    if (submissionsSheet.getLastRow() > 1) {
        allSubmissions = submissionsSheet.getRange(2, 2, submissionsSheet.getLastRow() - 1, 3).getValues();
    }
    let submissionCount = 0;
    for (const row of allSubmissions) {
      if (row[0] === email && row[1] === mediaType) {
        submissionCount++;
      }
    }
    // Subtract 1 because we're counting the current submission too
    submissionCount = Math.max(0, submissionCount - 1);

    // Calculate points based on submission count
    let points = 0;
    if (submissionCount === 0) {
      points = POINT_SYSTEM[mediaType].first;
    } else if (submissionCount === 1) {
      points = POINT_SYSTEM[mediaType].second;
    } else {
      points = POINT_SYSTEM[mediaType].thirdPlus;
    }

    // Add bonus points if applicable
    if (bonusPoints === "Yes") {
      points += BONUS_POINTS_VALUE;
    }
    
    // Set the calculated points and mark as verified
    submissionsSheet.getRange(submissionRow, 7).setValue(points); // Set Points (Column G)
    submissionsSheet.getRange(submissionRow, 8).setValue(true); // Set checkbox to checked (Column H)
    
    // Force recalculation
    SpreadsheetApp.flush();
    
    // Send confirmation email to student about their points update
    try {
        const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
        const rosterSheet = spreadsheet.getSheetByName(STUDENT_ROSTER_SHEET_NAME);
        
        // Get the student's current roster data
        const rosterData = rosterSheet.getRange(2, 1, rosterSheet.getLastRow() - 1, 5).getValues();
        const studentRowIndex = rosterData.findIndex(row => row[1] === email);
        
        if (studentRowIndex !== -1) {
            const studentRow = rosterData[studentRowIndex];
            const newTotalPoints = studentRow[3];
            const newTitle = studentRow[4];
            
            // For verification emails, try to determine if this is a level up by checking point ranges
            // Get the previous point total (current minus the points just added)
            const previousPoints = Math.max(0, newTotalPoints - points);
            
            // Find what title they would have had with previous points
            const titlesData = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 6, 4).getValues(); // Subtract 6 for system settings
            let oldTitle = "Gnome";
            for (const row of titlesData) {
                if (previousPoints >= row[0]) {
                    oldTitle = row[1];
                }
            }
            
            console.log(`Verification email: Previous points: ${previousPoints}, Old title: ${oldTitle}, New points: ${newTotalPoints}, New title: ${newTitle}`);
            sendConfirmationEmail(email, newTotalPoints, oldTitle, newTitle);
        }
    } catch (emailError) {
        console.log("Error sending verification email:", emailError);
        // Don't fail the verification if email fails
    }
    
    return { status: "success", message: "Submission verified successfully and student has been notified." };
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
  
  const allData = submissionsSheet.getRange(2, 1, submissionsSheet.getLastRow() - 1, 8).getValues();
  const pendingSubmissions = [];
  
  allData.forEach((row, index) => {
    if (row[7] === false) { // Column H is checkbox - false means unchecked/pending
      pendingSubmissions.push({
        rowNumber: index + 2, // +2 because we started from row 2 and arrays are 0-indexed
        timestamp: row[0],
        studentEmail: row[1],
        mediaType: row[2],
        mediaTitle: row[3],
        bonusPoints: row[4],
        reflection: row[5],
        currentPoints: row[6], // This will be 0 for pending submissions
        status: row[7] // false = pending, true = verified
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
      titlesData = settingsSheet.getRange(2, 1, settingsSheet.getLastRow() - 6, 4).getValues();
      
      // DEBUG: Log the titles data structure
      console.log("=== TITLES DATA DEBUG ===");
      console.log("Titles data array:");
      titlesData.forEach((row, index) => {
        console.log(`Row ${index}: [${row[0]}, ${row[1]}, ${row[2]}, ${row[3]}]`);
        console.log(`  Column A (row[0]): ${row[0]}`);
        console.log(`  Column B (row[1]): ${row[1]}`);
        console.log(`  Column C (row[2]): ${row[2]}`);
        console.log(`  Column D (row[3]): ${row[3]}`);
      });
      console.log("=== END TITLES DATA DEBUG ===");
      
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
