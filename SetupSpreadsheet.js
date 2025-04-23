/**
 * Interactive Video Overlay Tool - Spreadsheet Setup
 * 
 * This script initializes the spreadsheet structure with proper columns,
 * dropdown menus, validations, and sample data for the Interactive Video Overlay Tool.
 */

/**
 * Creates a custom menu in the spreadsheet UI
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Video Overlay App')
    .addItem('Setup Spreadsheet', 'setupSpreadsheet')
    .addItem('Add Sample Data', 'addSampleData')
    .addItem('Reset Sheets to Default', 'resetSheets')
    .addItem('Configure Settings', 'showSettingsDialog')
    .addItem('Deploy Web App', 'showDeploymentInstructions')
    .addToUi();
}

/**
 * Main function to setup the entire spreadsheet structure
 */
function setupSpreadsheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Confirm before proceeding
  const response = ui.alert(
    'Setup Spreadsheet', 
    'This will set up the required sheets and structure for the Interactive Video Overlay Tool. Existing sheets with the same names will be modified. Continue?', 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Create/update all required sheets
    setupVideosSheet(ss);
    setupOverlaysSheet(ss);
    setupQuizOptionsSheet(ss);
    setupAnalyticsSheet(ss);
    setupUserDataSheet(ss);
    setupSettingsSheet(ss);
    
    // Format the spreadsheet
    formatSpreadsheet(ss);
    
    ui.alert('Setup Complete', 'The spreadsheet has been set up successfully!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'An error occurred during setup: ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('Setup error: ' + error.toString());
  }
}

/**
 * Sets up the Videos sheet with appropriate columns and validations
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function setupVideosSheet(ss) {
  // Get or create Videos sheet
  let sheet = ss.getSheetByName('Videos');
  if (!sheet) {
    sheet = ss.insertSheet('Videos');
  }

  // Clear existing content
  sheet.clear();

  // Set column headers
  const headers = [
    'Video Title', 'Video URL', 'Description', 'Active'
  ];

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');

  // Set column widths
  sheet.setColumnWidth(1, 200); // Video Title
  sheet.setColumnWidth(2, 300); // Video URL
  sheet.setColumnWidth(3, 300); // Description
  sheet.setColumnWidth(4, 100); // Active

  // Add validation for Active column - CORRECTED LINE
  const activeValidation = SpreadsheetApp.newDataValidation()
    .requireCheckbox() // Use requireCheckbox() for TRUE/FALSE validation via checkboxes
    .setAllowInvalid(false)
    // You can customize checked/unchecked values if needed, e.g.:
    // .requireCheckbox('YES', 'NO')
    .build();
  sheet.getRange('D2:D1000').setDataValidation(activeValidation);

  // Add helper text
  sheet.getRange('A2').setValue('Enter your video title here');
  sheet.getRange('B2').setValue('https://www.youtube.com/watch?v=...');
  sheet.getRange('C2').setValue('Brief description of the video');
  sheet.getRange('D2').setValue(true); // Set default value to TRUE (checked checkbox)
  sheet.getRange('A2:D2').setBackground('#f1f8ff');
  sheet.getRange('A2:D2').setFontStyle('italic');

  // Freeze header row
  sheet.setFrozenRows(1);

  // Add conditional formatting to highlight active row
  // Note: The condition needs to check the cell value directly for TRUE/FALSE
  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=($D2=TRUE)') // Check if the cell in column D is TRUE
    .setBackground('#e6f4ea')
    .setRanges([sheet.getRange('A2:D1000')]) // Apply to the whole row based on column D
    .build();

  const rules = sheet.getConditionalFormatRules();
  // Clear existing rules to avoid duplicates if run multiple times
  while(rules.length > 0) {
      rules.pop();
  }
  rules.push(activeRule);
  sheet.setConditionalFormatRules(rules);

  // Add instructions
  sheet.getRange('F1').setValue('INSTRUCTIONS:');
  sheet.getRange('F1').setFontWeight('bold');
  sheet.getRange('F2').setValue('1. Enter your YouTube video information in the columns to the left.');
  sheet.getRange('F3').setValue('2. The first video with the "Active" box checked (TRUE) will be used in the web app.'); // Updated instruction text
  sheet.getRange('F4').setValue('3. After entering a video, go to the "Overlays" tab to add interaction points.');
  sheet.getRange('F2:F10').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setColumnWidth(6, 400);
}
/**
 * Sets up the Overlays sheet with appropriate columns and validations
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function setupOverlaysSheet(ss) {
  // Get or create Overlays sheet
  let sheet = ss.getSheetByName('Overlays');
  if (!sheet) {
    sheet = ss.insertSheet('Overlays');
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set column headers
  const headers = [
    'Video Title', 'Timestamp (seconds)', 'Overlay Title', 'Content', 
    'Interaction Type', 'Next Action', 'Correct Answer', 
    'Incorrect Answer 1', 'Incorrect Answer 2', 'Incorrect Answer 3',
    'Group Name', 'Explanation', 'Correct Feedback', 'Incorrect Feedback',
    'Image URL', 'Image Width', 'Image Height'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Set column widths
  sheet.setColumnWidth(1, 200);  // Video Title
  sheet.setColumnWidth(2, 100);  // Timestamp
  sheet.setColumnWidth(3, 200);  // Overlay Title
  sheet.setColumnWidth(4, 300);  // Content
  sheet.setColumnWidth(5, 150);  // Interaction Type
  sheet.setColumnWidth(6, 150);  // Next Action
  sheet.setColumnWidth(7, 200);  // Correct Answer
  sheet.setColumnWidth(8, 200);  // Incorrect Answer 1
  sheet.setColumnWidth(9, 200);  // Incorrect Answer 2
  sheet.setColumnWidth(10, 200); // Incorrect Answer 3
  sheet.setColumnWidth(11, 150); // Group Name
  sheet.setColumnWidth(12, 300); // Explanation
  sheet.setColumnWidth(13, 200); // Correct Feedback
  sheet.setColumnWidth(14, 200); // Incorrect Feedback
  sheet.setColumnWidth(15, 250); // Image URL
  sheet.setColumnWidth(16, 100); // Image Width
  sheet.setColumnWidth(17, 100); // Image Height
  
  // Add data validations
  
  // Video Title dropdown (based on Videos sheet)
  const videosSheet = ss.getSheetByName('Videos');
  if (videosSheet) {
    const videoTitlesRange = 'Videos!$A$2:$A$1000';
    const videoTitleValidation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(ss.getRange(videoTitlesRange), true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange('A2:A1000').setDataValidation(videoTitleValidation);
  }
  
  // Interaction Type dropdown
  const interactionTypes = ['info', 'quiz', 'true_false', 'matching'];
  const interactionTypeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(interactionTypes, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('E2:E1000').setDataValidation(interactionTypeValidation);
  
  // Next Action dropdown
  const nextActions = ['continue', 'next_question', 'if_correct', 'if_incorrect', 'end'];
  const nextActionValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(nextActions, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('F2:F1000').setDataValidation(nextActionValidation);
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Add conditional formatting to highlight different interaction types
  const infoRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E1="info"')
    .setBackground('#e6f4ea')
    .setRanges([sheet.getRange('A2:Q1000')])
    .build();
  
  const quizRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E1="quiz"')
    .setBackground('#fce8e6')
    .setRanges([sheet.getRange('A2:Q1000')])
    .build();
  
  const trueFalseRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E1="true_false"')
    .setBackground('#fff7e6')
    .setRanges([sheet.getRange('A2:Q1000')])
    .build();
  
  const matchingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$E1="matching"')
    .setBackground('#e6f4ff')
    .setRanges([sheet.getRange('A2:Q1000')])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(infoRule, quizRule, trueFalseRule, matchingRule);
  sheet.setConditionalFormatRules(rules);
  
  // Add instructions
  sheet.getRange('S1').setValue('INSTRUCTIONS:');
  sheet.getRange('S1').setFontWeight('bold');
  sheet.getRange('S2').setValue('1. Video Title: Select from dropdown (must be added to Videos tab first)');
  sheet.getRange('S3').setValue('2. Timestamp: Enter seconds when overlay should appear (e.g., 30 for 0:30)');
  sheet.getRange('S4').setValue('3. Interaction Type: Choose the type of interaction');
  sheet.getRange('S5').setValue('   - info: Simple information display');
  sheet.getRange('S6').setValue('   - quiz: Multiple choice question');
  sheet.getRange('S7').setValue('   - true_false: True/False question');
  sheet.getRange('S8').setValue('   - matching: Matching options question');
  sheet.getRange('S9').setValue('4. Next Action: What happens after this overlay');
  sheet.getRange('S10').setValue('   - continue: Continue playing the video (default)');
  sheet.getRange('S11').setValue('   - next_question: Go to next question in sequence');
  sheet.getRange('S12').setValue('   - if_correct: Custom logic based on correct/incorrect answer');
  sheet.getRange('S13').setValue('   - if_incorrect: Custom logic based on correct/incorrect answer');
  sheet.getRange('S14').setValue('   - end: End the video and show the summary report');
  sheet.getRange('S15').setValue('5. For quizzes, enter correct answer(s) and incorrect options');
  sheet.getRange('S16').setValue('   - Multiple correct answers can be separated with pipe symbol (|)');
  sheet.getRange('S17').setValue('6. Group Name: Optionally group related overlays together');
  sheet.getRange('S18').setValue('7. Explanation: Additional context shown after answering');
  sheet.getRange('S19').setValue('8. Image URL: Optional image to display in the overlay');
  
  sheet.getRange('S1:S30').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setColumnWidth(19, 400);
}

/**
 * Sets up the Quiz Options sheet for backward compatibility
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function setupQuizOptionsSheet(ss) {
  // Get or create Quiz Options sheet (mostly for backward compatibility)
  let sheet = ss.getSheetByName('Quiz Options');
  if (!sheet) {
    sheet = ss.insertSheet('Quiz Options');
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set column headers
  const headers = [
    'Video Title', 'Overlay Title', 'Option Text', 'Is Correct', 'Feedback'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Add note about this sheet
  sheet.getRange('A2').setValue('NOTE: This sheet is maintained for backward compatibility. It is recommended to use the Overlays sheet for quiz options.');
  sheet.getRange('A2:E2').merge();
  sheet.getRange('A2').setFontStyle('italic');
  sheet.getRange('A2').setFontColor('red');
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Sets up the Analytics sheet for quiz performance data
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function setupAnalyticsSheet(ss) {
  // Get or create Analytics sheet
  let sheet = ss.getSheetByName('Quiz Analytics');
  if (!sheet) {
    sheet = ss.insertSheet('Quiz Analytics');
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set column headers
  const headers = [
    'Timestamp', 'User ID', 'Video Title', 'Overlay ID', 
    'Quiz Type', 'Was Correct', 'Selected Option', 
    'Time to Answer (sec)', 'Session ID'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Add data filtering
  sheet.getRange(1, 1, 1, headers.length).createFilter();
}

/**
 * Sets up the User Data sheet for tracking user activities
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function setupUserDataSheet(ss) {
  // Get or create User Data sheet
  let sheet = ss.getSheetByName('User Data');
  if (!sheet) {
    sheet = ss.insertSheet('User Data');
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set column headers
  const headers = [
    'Timestamp', 'Session ID', 'User ID', 'Video Title',
    'Event Type', 'Event Data', 'Browser', 'Device'
  ];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Add data filtering
  sheet.getRange(1, 1, 1, headers.length).createFilter();
}

/**
 * Sets up the Settings sheet for application configuration
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function setupSettingsSheet(ss) {
  // Get or create Settings sheet
  let sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    sheet = ss.insertSheet('Settings');
  }
  
  // Clear existing content
  sheet.clear();
  
  // Set column headers
  const headers = ['Setting', 'Value', 'Description'];
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('white');
  
  // Default settings
  const settings = [
    ['TeacherModeEnabled', 'TRUE', 'Enable teacher mode toggle in the web app'],
    ['AllowNotes', 'TRUE', 'Allow students to take notes during the video'],
    ['ShowProgressBar', 'TRUE', 'Show progress bar with overlay markers'],
    ['ShowStudentReport', 'TRUE', 'Show performance report at the end of the video'],
    ['PrimaryColor', '#4285f4', 'Primary theme color (hex code)'],
    ['SecondaryColor', '#34a853', 'Secondary theme color (hex code)'],
    ['AllowSkipping', 'FALSE', 'Allow students to skip ahead in the video'],
    ['RequireCorrectAnswers', 'FALSE', 'Require correct answers to continue'],
    ['ShowCorrectAnswers', 'TRUE', 'Show correct answers after quiz attempt']
  ];
  
  sheet.getRange(2, 1, settings.length, 3).setValues(settings);
  
  // Set column widths
  sheet.setColumnWidth(1, 200);  // Setting
  sheet.setColumnWidth(2, 150);  // Value
  sheet.setColumnWidth(3, 400);  // Description
  
  // Add data validation for boolean settings
  const booleanSettings = ['TeacherModeEnabled', 'AllowNotes', 'ShowProgressBar', 
                           'ShowStudentReport', 'AllowSkipping', 'RequireCorrectAnswers',
                           'ShowCorrectAnswers'];
  
  for (let i = 0; i < settings.length; i++) {
    if (booleanSettings.includes(settings[i][0])) {
      const validation = SpreadsheetApp.newDataValidation()
        .requireValueInList(['TRUE', 'FALSE'], true)
        .setAllowInvalid(false)
        .build();
      sheet.getRange(i + 2, 2).setDataValidation(validation);
    }
  }
  
  // Freeze header row
  sheet.setFrozenRows(1);
}

/**
 * Adds sample data to the spreadsheet
 */
function addSampleData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Confirm before proceeding
  const response = ui.alert(
    'Add Sample Data', 
    'This will add sample data to the spreadsheet. Continue?', 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Add sample video
    const videosSheet = ss.getSheetByName('Videos');
    if (videosSheet) {
      videosSheet.getRange('A2').setValue('Sample Tutorial');
      videosSheet.getRange('B2').setValue('https://www.youtube.com/watch?v=dQw4w9WgXcQ');
      videosSheet.getRange('C2').setValue('A sample video to test the interactive overlay functionality');
      videosSheet.getRange('D2').setValue('TRUE');
    }
    
    // Add sample overlays
    const overlaysSheet = ss.getSheetByName('Overlays');
    if (overlaysSheet) {
      // Info overlay
      overlaysSheet.getRange('A2').setValue('Sample Tutorial');
      overlaysSheet.getRange('B2').setValue(5);
      overlaysSheet.getRange('C2').setValue('Key Point #1');
      overlaysSheet.getRange('D2').setValue('This is an important concept explained at this timestamp.');
      overlaysSheet.getRange('E2').setValue('info');
      overlaysSheet.getRange('F2').setValue('continue');
      
      // Quiz overlay
      overlaysSheet.getRange('A3').setValue('Sample Tutorial');
      overlaysSheet.getRange('B3').setValue(10);
      overlaysSheet.getRange('C3').setValue('Interactive Quiz');
      overlaysSheet.getRange('D3').setValue('Based on what you\'ve seen so far, which option is correct?');
      overlaysSheet.getRange('E3').setValue('quiz');
      overlaysSheet.getRange('F3').setValue('continue');
      overlaysSheet.getRange('G3').setValue('This response is correct');
      overlaysSheet.getRange('H3').setValue('Incorrect option #1');
      overlaysSheet.getRange('I3').setValue('Incorrect option #2');
      overlaysSheet.getRange('J3').setValue('Incorrect option #3');
      overlaysSheet.getRange('L3').setValue('The correct answer demonstrates understanding of the concept.');
      overlaysSheet.getRange('M3').setValue('Well done! You\'ve understood the concept.');
      overlaysSheet.getRange('N3').setValue('Not quite. Review the video again for clarification.');
      
      // True/False question
      overlaysSheet.getRange('A4').setValue('Sample Tutorial');
      overlaysSheet.getRange('B4').setValue(15);
      overlaysSheet.getRange('C4').setValue('True or False');
      overlaysSheet.getRange('D4').setValue('The statement presented in the video at 0:12 is accurate.');
      overlaysSheet.getRange('E4').setValue('true_false');
      overlaysSheet.getRange('F4').setValue('next_question');
      overlaysSheet.getRange('G4').setValue('TRUE');
      overlaysSheet.getRange('H4').setValue('FALSE');
      
      // Branching question
      overlaysSheet.getRange('A5').setValue('Sample Tutorial');
      overlaysSheet.getRange('B5').setValue(20);
      overlaysSheet.getRange('C5').setValue('Branching Question');
      overlaysSheet.getRange('D5').setValue('Do you want to learn more about this topic?');
      overlaysSheet.getRange('E5').setValue('quiz');
      overlaysSheet.getRange('F5').setValue('if_correct:25');
      overlaysSheet.getRange('G5').setValue('Yes, tell me more');
      overlaysSheet.getRange('H5').setValue('No, continue with the video');
      
      // Final question with end action
      overlaysSheet.getRange('A6').setValue('Sample Tutorial');
      overlaysSheet.getRange('B6').setValue(30);
      overlaysSheet.getRange('C6').setValue('Final Question');
      overlaysSheet.getRange('D6').setValue('What was the main takeaway from this video?');
      overlaysSheet.getRange('E6').setValue('quiz');
      overlaysSheet.getRange('F6').setValue('end');
      overlaysSheet.getRange('G6').setValue('The correct main takeaway');
      overlaysSheet.getRange('H6').setValue('An incorrect interpretation');
      overlaysSheet.getRange('I6').setValue('Another incorrect option');
      overlaysSheet.getRange('L6').setValue('The main takeaway helps synthesize the key concepts presented throughout the video.');
    }
    
    ui.alert('Sample Data Added', 'Sample data has been added to the spreadsheet!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'An error occurred while adding sample data: ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('Sample data error: ' + error.toString());
  }
}

/**
 * Resets the spreadsheet to its default structure
 */
function resetSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Confirm before proceeding
  const response = ui.alert(
    'Reset Sheets', 
    'This will delete all sheets and recreate them with default settings. All data will be lost. Continue?', 
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    // Delete all sheets
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      if (sheets.length > 1) {  // Keep at least one sheet
        ss.deleteSheet(sheets[i]);
      }
    }
    
    // Recreate sheets
    setupSpreadsheet();
    
    ui.alert('Reset Complete', 'The spreadsheet has been reset to default settings!', ui.ButtonSet.OK);
  } catch (error) {
    ui.alert('Error', 'An error occurred during reset: ' + error.toString(), ui.ButtonSet.OK);
    Logger.log('Reset error: ' + error.toString());
  }
}

/**
 * Shows a dialog with settings configuration
 */
function showSettingsDialog() {
  const html = HtmlService.createHtmlOutput(`
    <h2>Settings Configuration</h2>
    <p>You can configure the application settings in the "Settings" sheet.</p>
    <p>Changes to settings will take effect the next time the web app is loaded.</p>
    <button onclick="google.script.host.close()">Close</button>
  `)
    .setWidth(400)
    .setHeight(200);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Settings Configuration');
}

/**
 * Shows instructions for deploying the web app
 */
function showDeploymentInstructions() {
  const html = HtmlService.createHtmlOutput(`
    <h2>Deploy Web App Instructions</h2>
    <ol>
      <li>In the Apps Script editor, click on "Deploy" &gt; "New deployment"</li>
      <li>Select "Web app" as the deployment type</li>
      <li>Fill in the following settings:
        <ul>
          <li>Description: "Interactive Video Overlay Tool"</li>
          <li>Execute as: "Me" (or your account)</li>
          <li>Who has access: Choose appropriate option (e.g., "Anyone" or "Anyone within [your domain]")</li>
        </ul>
      </li>
      <li>Click "Deploy" and authorize the app</li>
      <li>Copy the provided web app URL to share with your students</li>
    </ol>
    <button onclick="google.script.host.close()">Close</button>
  `)
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Deploy Web App Instructions');
}

/**
 * Formats the entire spreadsheet for better readability
 * @param {SpreadsheetApp.Spreadsheet} ss - The active spreadsheet
 */
function formatSpreadsheet(ss) {
  // Set spreadsheet properties
  ss.setSpreadsheetTheme(SpreadsheetApp.getDefaultSpreadsheetTheme()
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT1, '#4285f4')
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT2, '#34a853')
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT3, '#fbbc05')
    .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT4, '#ea4335')
  );
  
  // Set the order of sheets
  const sheetsOrder = ['Videos', 'Overlays', 'Quiz Options', 'Settings', 'Quiz Analytics', 'User Data'];
  const sheets = ss.getSheets();
  
  // Reorder sheets
  for (let i = 0; i < sheetsOrder.length; i++) {
    const sheet = ss.getSheetByName(sheetsOrder[i]);
    if (sheet) {
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(i + 1);
    }
  }
  
  // Set active sheet to Videos
  const videosSheet = ss.getSheetByName('Videos');
  if (videosSheet) {
    ss.setActiveSheet(videosSheet);
  }
}