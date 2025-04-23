/**
 * Interactive Video Overlay Tool - Enhanced Version
 * 
 * This application enhances YouTube videos with interactive overlays 
 * that appear at specific timestamps during playback.
 * Added features:
 * - Branching logic based on quiz responses
 * - Student performance reports
 * - Consecutive question support
 * - Enhanced analytics
 */

// Configuration
const CONFIG = {
  SHEETS: {
    VIDEOS: 'Videos',
    OVERLAYS: 'Overlays',
    QUIZ_ANALYTICS: 'Quiz Analytics',
    USER_DATA: 'User Data',
    SETTINGS: 'Settings',
    USER_NOTES: 'User Notes'
  },
  DEFAULTS: {
    ANIMATION_DURATION: 400,
    CACHE_DURATION: 1800 // 30 minutes in seconds
  },
  OVERLAY_TYPES: {
    INFO: 'info',
    QUIZ_MULTIPLE_CHOICE: 'quiz',
    QUIZ_TRUE_FALSE: 'true_false',
    QUIZ_MATCHING: 'matching'
  },
  NEXT_ACTIONS: {
    CONTINUE: 'continue',
    NEXT_QUESTION: 'next_question',
    IF_CORRECT: 'if_correct',
    IF_INCORRECT: 'if_incorrect',
    END: 'end'
  }
};

/**
 * Serves the web application HTML page
 * @param {Object} e - Event object
 * @returns {HtmlOutput} The HTML page
 */
function doGet(e) {
  // Initialize analytics and settings if they don't exist
  ensureRequiredSheets();
  
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Interactive Video Overlay')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Ensures all required sheets exist in the spreadsheet
 */
function ensureRequiredSheets() {
  const ss = SpreadsheetApp.getActive();
  
  // Create Quiz Analytics sheet if it doesn't exist
  if (!ss.getSheetByName(CONFIG.SHEETS.QUIZ_ANALYTICS)) {
    const analyticsSheet = ss.insertSheet(CONFIG.SHEETS.QUIZ_ANALYTICS);
    analyticsSheet.appendRow([
      'Timestamp', 'User ID', 'Video Title', 'Overlay ID', 
      'Quiz Type', 'Was Correct', 'Selected Option', 
      'Time to Answer (sec)', 'Session ID'
    ]);
  }
  
  // Create User Data sheet if it doesn't exist
  if (!ss.getSheetByName(CONFIG.SHEETS.USER_DATA)) {
    const userDataSheet = ss.insertSheet(CONFIG.SHEETS.USER_DATA);
    userDataSheet.appendRow([
      'Timestamp', 'Session ID', 'User ID', 'Video Title',
      'Event Type', 'Event Data', 'Browser', 'Device'
    ]);
  }
  
  // Create Settings sheet if it doesn't exist
  if (!ss.getSheetByName(CONFIG.SHEETS.SETTINGS)) {
    const settingsSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
    settingsSheet.appendRow(['Setting', 'Value', 'Description']);
    
    // Default settings
    settingsSheet.appendRow(['TeacherModeEnabled', 'TRUE', 'Enable teacher mode toggle in the web app']);
    settingsSheet.appendRow(['AllowNotes', 'TRUE', 'Allow students to take notes during the video']);
    settingsSheet.appendRow(['ShowProgressBar', 'TRUE', 'Show progress bar with overlay markers']);
    settingsSheet.appendRow(['ShowStudentReport', 'TRUE', 'Show performance report at the end of the video']);
    settingsSheet.appendRow(['PrimaryColor', '#4285f4', 'Primary theme color (hex code)']);
    settingsSheet.appendRow(['SecondaryColor', '#34a853', 'Secondary theme color (hex code)']);
    settingsSheet.appendRow(['AllowSkipping', 'FALSE', 'Allow students to skip ahead in the video']);
    settingsSheet.appendRow(['RequireCorrectAnswers', 'FALSE', 'Require correct answers to continue']);
  }
  
  // Create User Notes sheet if it doesn't exist
  if (!ss.getSheetByName(CONFIG.SHEETS.USER_NOTES)) {
    const notesSheet = ss.insertSheet(CONFIG.SHEETS.USER_NOTES);
    notesSheet.appendRow([
      'Timestamp', 'User ID', 'Video Title', 'Video Time (sec)',
      'Note Content', 'Session ID'
    ]);
  }
}

/**
 * Includes a file content within the main HTML file
 * @param {string} filename - Name of the file to include
 * @return {string} The content of the file
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Gets the default video from row 2 of the Videos tab
 * @returns {Object} Video data object or error
 */
function getDefaultVideo() {
  try {
    // Access the spreadsheet
    const ss = SpreadsheetApp.getActive();
    const videosSheet = ss.getSheetByName(CONFIG.SHEETS.VIDEOS);
    
    if (!videosSheet) {
      return { error: "Videos tab not found" };
    }
    
    // Find the first active video
    const videoData = videosSheet.getDataRange().getValues();
    let activeVideo = null;
    
    // Skip header row
    for (let i = 1; i < videoData.length; i++) {
      const row = videoData[i];
      if (row[3] === true || row[3] === 'TRUE') {  // Column D: Active status
        activeVideo = row;
        break;
      }
    }
    
    if (!activeVideo) {
      return { error: "No active video found. Please set at least one video to Active=TRUE in the Videos tab." };
    }
    
    const videoTitle = activeVideo[0]; // Column A: Video Title
    const videoUrl = activeVideo[1];   // Column B: Video URL
    
    // Extract YouTube video ID
    const videoId = extractYouTubeVideoId(videoUrl);
    if (!videoId) {
      return { error: "Invalid YouTube URL in the active video row." };
    }
    
    // Get app settings
    const settings = getAppSettings();
    
    return {
      videoId: videoId,
      videoTitle: videoTitle,
      settings: settings
    };
  } catch (error) {
    Logger.log("Error in getDefaultVideo: " + error.toString());
    return { error: "Error getting video: " + error.toString() };
  }
}

/**
 * Gets application settings from the Settings sheet
 * @returns {Object} Settings object
 */
function getAppSettings() {
  try {
    const ss = SpreadsheetApp.getActive();
    const settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
    
    if (!settingsSheet) {
      return getDefaultSettings();
    }
    
    const settingsData = settingsSheet.getDataRange().getValues();
    const settings = {};
    
    // Skip header row
    for (let i = 1; i < settingsData.length; i++) {
      const setting = settingsData[i][0];
      const value = settingsData[i][1];
      
      // Convert string booleans to actual booleans
      if (value === "TRUE" || value === "FALSE") {
        settings[setting] = (value === "TRUE");
      } else {
        settings[setting] = value;
      }
    }
    
    return settings;
  } catch (error) {
    Logger.log("Error in getAppSettings: " + error.toString());
    return getDefaultSettings();
  }
}

/**
 * Returns default settings if Settings sheet doesn't exist
 * @returns {Object} Default settings object
 */
function getDefaultSettings() {
  return {
    TeacherModeEnabled: true,
    AllowNotes: true,
    ShowProgressBar: true,
    ShowStudentReport: true,
    PrimaryColor: '#4285f4',
    SecondaryColor: '#34a853',
    AllowSkipping: false,
    RequireCorrectAnswers: false
  };
}

/**
 * Gets all overlays for the specified video title
 * @param {string} videoTitle - The title of the video to get overlays for
 * @returns {Object} Overlays data object or error
 */
function getOverlaysForVideo(videoTitle) {
  try {
    const ss = SpreadsheetApp.getActive();
    const overlaysSheet = ss.getSheetByName(CONFIG.SHEETS.OVERLAYS);
    
    if (!overlaysSheet) {
      return { error: "Overlays tab not found" };
    }
    
    // Get overlays data
    const overlaysData = overlaysSheet.getDataRange().getValues();
    const overlays = [];
    const groups = {};
    let overlaysByTitle = {};
    
    // Process overlays and build groups
    // Skip header row
    for (let i = 1; i < overlaysData.length; i++) {
      const row = overlaysData[i];
      
      // Check if this overlay belongs to our video
      if (row[0] && row[0] === videoTitle) {
        const timestamp = parseInt(row[1], 10);
        const title = row[2];
        const content = row[3];
        const type = row[4] ? row[4].toLowerCase() : 'info';
        const nextAction = row[5] || "continue";
        
        // Skip if missing essential data
        if (!timestamp || !title || !content) continue;
        
        // Parse next action logic
        let nextActionData = parseNextAction(nextAction);
        
        const overlay = {
          id: `overlay-${i}`,
          timestamp: timestamp,
          title: title,
          content: content,
          type: type,
          nextAction: nextActionData.action,
          actionParam: nextActionData.param,
          options: [],
          explanation: row[11] || "", // Column L: Explanation
          correctFeedback: row[12] || "", // Column M: Correct feedback
          incorrectFeedback: row[13] || "", // Column N: Incorrect feedback
          groupName: row[10] || "" // Column K: Group name
        };
        
        // If it's a quiz type, get the answer options
        if (type.includes('quiz') || type.includes('true_false')) {
          // Column G (index 6) is the correct answer
          // Columns H, I, J (indices 7, 8, 9) are incorrect answers
          const correctAnswer = row[6];
          
          if (correctAnswer) {
            const correctAnswers = correctAnswer.toString().split('|');
            
            // Add each correct answer as a separate option
            correctAnswers.forEach(answer => {
              if (answer.trim() !== '') {
                overlay.options.push({
                  text: answer.trim(),
                  isCorrect: true,
                  feedback: overlay.correctFeedback || "Correct!"
                });
              }
            });
            
            // Add incorrect answers (if they exist)
            for (let j = 7; j <= 9; j++) {
              if (row[j] && row[j].toString().trim() !== '') {
                overlay.options.push({
                  text: row[j],
                  isCorrect: false,
                  feedback: overlay.incorrectFeedback || "Incorrect."
                });
              }
            }
            
            // For true_false, ensure we only have TRUE and FALSE options
            if (type === 'true_false') {
              let hasTrueOption = false;
              let hasFalseOption = false;
              
              overlay.options.forEach(option => {
                if (option.text.toUpperCase() === 'TRUE') hasTrueOption = true;
                if (option.text.toUpperCase() === 'FALSE') hasFalseOption = true;
              });
              
              // Add missing options if needed
              if (!hasTrueOption) {
                overlay.options.push({
                  text: 'TRUE',
                  isCorrect: false,
                  feedback: overlay.incorrectFeedback || "Incorrect."
                });
              }
              
              if (!hasFalseOption) {
                overlay.options.push({
                  text: 'FALSE',
                  isCorrect: false,
                  feedback: overlay.incorrectFeedback || "Incorrect."
                });
              }
            }
            
            // Shuffle the options so correct answer isn't always first
            overlay.options = shuffleArray(overlay.options);
          }
        }
        
        // Process multimedia if any
        if (row[14]) { // Column O: Image URL
          overlay.image = {
            url: row[14],
            width: row[15] || "auto",  // Column P: Image width
            height: row[16] || "auto"  // Column Q: Image height
          };
        }
        
        overlays.push(overlay);
        
        // Store by title for lookup
        overlaysByTitle[title] = overlay;
        
        // Add to group if specified
        if (overlay.groupName) {
          if (!groups[overlay.groupName]) {
            groups[overlay.groupName] = [];
          }
          groups[overlay.groupName].push(overlay.id);
        }
      }
    }
    
    // Sort overlays by timestamp
    overlays.sort((a, b) => a.timestamp - b.timestamp);
    
    // Process consecutive questions
    // We need to do this after all overlays are loaded so we can find the next questions
    for (let i = 0; i < overlays.length; i++) {
      const overlay = overlays[i];
      
      if (overlay.nextAction === CONFIG.NEXT_ACTIONS.NEXT_QUESTION) {
        // Find the next question after this one
        let nextQuestion = null;
        for (let j = i + 1; j < overlays.length; j++) {
          if (overlays[j].type.includes('quiz') || overlays[j].type.includes('true_false')) {
            nextQuestion = overlays[j];
            break;
          }
        }
        
        if (nextQuestion) {
          overlay.actionParam = nextQuestion.timestamp;
        }
      }
    }
    
    return {
      overlays: overlays,
      groups: groups,
      overlaysByTitle: overlaysByTitle
    };
  } catch (error) {
    Logger.log("Error in getOverlaysForVideo: " + error.toString());
    return { error: "Error getting overlays: " + error.toString() };
  }
}

/**
 * Parse the next action value from the spreadsheet
 * @param {string} nextAction - The next action from the spreadsheet
 * @returns {Object} Parsed action and parameter
 */
function parseNextAction(nextAction) {
  if (!nextAction) {
    return { action: CONFIG.NEXT_ACTIONS.CONTINUE, param: null };
  }
  
  // Handle direct action values
  if ([CONFIG.NEXT_ACTIONS.CONTINUE, CONFIG.NEXT_ACTIONS.NEXT_QUESTION, CONFIG.NEXT_ACTIONS.END].includes(nextAction)) {
    return { action: nextAction, param: null };
  }
  
  // Handle if_correct:timestamp format
  if (nextAction.startsWith(CONFIG.NEXT_ACTIONS.IF_CORRECT + ':')) {
    const param = nextAction.split(':')[1].trim();
    return { action: CONFIG.NEXT_ACTIONS.IF_CORRECT, param: param };
  }
  
  // Handle if_incorrect:timestamp format
  if (nextAction.startsWith(CONFIG.NEXT_ACTIONS.IF_INCORRECT + ':')) {
    const param = nextAction.split(':')[1].trim();
    return { action: CONFIG.NEXT_ACTIONS.IF_INCORRECT, param: param };
  }
  
  // Default to continue if we don't recognize the format
  return { action: CONFIG.NEXT_ACTIONS.CONTINUE, param: null };
}

/**
 * Records quiz attempt data in the analytics sheet
 * @param {Object} quizData - Quiz attempt data
 * @returns {Object} Success message or error
 */
function recordQuizAttempt(quizData) {
  try {
    const ss = SpreadsheetApp.getActive();
    let analyticsSheet = ss.getSheetByName(CONFIG.SHEETS.QUIZ_ANALYTICS);
    
    if (!analyticsSheet) {
      analyticsSheet = ss.insertSheet(CONFIG.SHEETS.QUIZ_ANALYTICS);
      analyticsSheet.appendRow([
        'Timestamp', 'User ID', 'Video Title', 'Overlay ID', 
        'Quiz Type', 'Was Correct', 'Selected Option', 
        'Time to Answer (sec)', 'Session ID'
      ]);
    }
    
    // Add new row with quiz data
    analyticsSheet.appendRow([
      new Date(),
      quizData.userId || 'anonymous',
      quizData.videoTitle || '',
      quizData.overlayId || '',
      quizData.quizType || 'quiz',
      quizData.wasCorrect,
      quizData.selectedOption || '',
      quizData.timeToAnswer || 0,
      quizData.sessionId || ''
    ]);
    
    return { success: true, message: "Quiz data recorded successfully" };
  } catch (error) {
    Logger.log("Error in recordQuizAttempt: " + error.toString());
    return { error: "Failed to record quiz data: " + error.toString() };
  }
}

/**
 * Records user viewing events for analytics
 * @param {Object} eventData - User event data
 * @returns {Object} Success message or error
 */
function recordUserEvent(eventData) {
  try {
    const ss = SpreadsheetApp.getActive();
    let userDataSheet = ss.getSheetByName(CONFIG.SHEETS.USER_DATA);
    
    if (!userDataSheet) {
      userDataSheet = ss.insertSheet(CONFIG.SHEETS.USER_DATA);
      userDataSheet.appendRow([
        'Timestamp', 'Session ID', 'User ID', 'Video Title',
        'Event Type', 'Event Data', 'Browser', 'Device'
      ]);
    }
    
    // Add new row with event data
    userDataSheet.appendRow([
      new Date(),
      eventData.sessionId || '',
      eventData.userId || 'anonymous',
      eventData.videoTitle || '',
      eventData.eventType || '',
      eventData.eventData || '',
      eventData.browser || '',
      eventData.device || ''
    ]);
    
    return { success: true, message: "User event recorded successfully" };
  } catch (error) {
    Logger.log("Error in recordUserEvent: " + error.toString());
    return { error: "Failed to record user event: " + error.toString() };
  }
}

/**
 * Saves a user note for a specific video timestamp
 * @param {Object} noteData - Note data
 * @returns {Object} Success message or error
 */
function saveUserNote(noteData) {
  try {
    const ss = SpreadsheetApp.getActive();
    let notesSheet = ss.getSheetByName(CONFIG.SHEETS.USER_NOTES);
    
    if (!notesSheet) {
      notesSheet = ss.insertSheet(CONFIG.SHEETS.USER_NOTES);
      notesSheet.appendRow([
        'Timestamp', 'User ID', 'Video Title', 'Video Time (sec)',
        'Note Content', 'Session ID'
      ]);
    }
    
    // Add new row with note data
    notesSheet.appendRow([
      new Date(),
      noteData.userId || 'anonymous',
      noteData.videoTitle || '',
      noteData.videoTime || 0,
      noteData.noteContent || '',
      noteData.sessionId || ''
    ]);
    
    return { success: true, message: "Note saved successfully" };
  } catch (error) {
    Logger.log("Error in saveUserNote: " + error.toString());
    return { error: "Failed to save note: " + error.toString() };
  }
}

/**
 * Gets user notes for a specific video
 * @param {string} videoTitle - Video title
 * @param {string} userId - User identifier
 * @returns {Object} User notes or error
 */
function getUserNotes(videoTitle, userId = 'anonymous') {
  try {
    const ss = SpreadsheetApp.getActive();
    const notesSheet = ss.getSheetByName(CONFIG.SHEETS.USER_NOTES);
    
    if (!notesSheet) {
      return { notes: [] };
    }
    
    const notesData = notesSheet.getDataRange().getValues();
    const notes = [];
    
    // Skip header row
    for (let i = 1; i < notesData.length; i++) {
      const row = notesData[i];
      
      // Check if this note matches the video and user
      if (row[2] === videoTitle && row[1] === userId) {
        notes.push({
          timestamp: row[0],
          videoTime: row[3],
          content: row[4]
        });
      }
    }
    
    return { notes: notes };
  } catch (error) {
    Logger.log("Error in getUserNotes: " + error.toString());
    return { error: "Failed to retrieve notes: " + error.toString() };
  }
}

/**
 * Generates student performance report
 * @param {string} videoTitle - Video title
 * @param {string} sessionId - Session identifier
 * @param {string} userId - User identifier
 * @returns {Object} Performance report data
 */
function getStudentReport(videoTitle, sessionId, userId = 'anonymous') {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Get quiz analytics data
    const analyticsSheet = ss.getSheetByName(CONFIG.SHEETS.QUIZ_ANALYTICS);
    if (!analyticsSheet) {
      return { 
        error: "No quiz analytics data available",
        summary: createDefaultReportSummary()
      };
    }
    
    const analyticsData = analyticsSheet.getDataRange().getValues();
    
    // Get user event data
    const userDataSheet = ss.getSheetByName(CONFIG.SHEETS.USER_DATA);
    let viewingData = [];
    if (userDataSheet) {
      viewingData = userDataSheet.getDataRange().getValues();
    }
    
    // Get notes data
    const notesSheet = ss.getSheetByName(CONFIG.SHEETS.USER_NOTES);
    let notesData = [];
    if (notesSheet) {
      notesData = notesSheet.getDataRange().getValues();
    }
    
    // Initialize report data
    const report = {
      videoTitle: videoTitle,
      userId: userId,
      sessionId: sessionId,
      quizPerformance: {
        totalQuestions: 0,
        correctAnswers: 0,
        incorrectAnswers: 0,
        accuracyPercentage: 0,
        averageTimeToAnswer: 0,
        quizDetails: []
      },
      viewingStatistics: {
        startTime: null,
        endTime: null,
        totalTimeSpent: 0,
        completionPercentage: 0,
        pauseCount: 0
      },
      notesCount: 0,
      summary: createDefaultReportSummary()
    };
    
    // Process quiz analytics
    let totalTimeToAnswer = 0;
    
    // Skip header row
    for (let i = 1; i < analyticsData.length; i++) {
      const row = analyticsData[i];
      
      // Only include data for this session and video
      if (row[8] === sessionId && row[2] === videoTitle) {
        const overlayId = row[3];
        const quizType = row[4];
        const wasCorrect = row[5] === true || row[5] === "TRUE";
        const selectedOption = row[6];
        const timeToAnswer = parseFloat(row[7]) || 0;
        
        report.quizPerformance.totalQuestions++;
        if (wasCorrect) {
          report.quizPerformance.correctAnswers++;
        } else {
          report.quizPerformance.incorrectAnswers++;
        }
        
        totalTimeToAnswer += timeToAnswer;
        
        // Add to quiz details
        report.quizPerformance.quizDetails.push({
          overlayId: overlayId,
          quizType: quizType,
          wasCorrect: wasCorrect,
          selectedOption: selectedOption,
          timeToAnswer: timeToAnswer
        });
      }
    }
    
    // Calculate quiz performance metrics
    if (report.quizPerformance.totalQuestions > 0) {
      report.quizPerformance.accuracyPercentage = 
        (report.quizPerformance.correctAnswers / report.quizPerformance.totalQuestions) * 100;
      report.quizPerformance.averageTimeToAnswer = 
        totalTimeToAnswer / report.quizPerformance.totalQuestions;
    }
    
    // Process viewing data
    let startEvent = null;
    let endEvent = null;
    let pauseCount = 0;
    
    // Skip header row
    for (let i = 1; i < viewingData.length; i++) {
      const row = viewingData[i];
      
      // Only include data for this session and video
      if (row[1] === sessionId && row[3] === videoTitle) {
        const timestamp = row[0]; // Timestamp
        const eventType = row[4]; // Event Type
        
        // Track start/end times
        if (eventType === 'activity_started' && (!startEvent || timestamp < startEvent)) {
          startEvent = timestamp;
        }
        if (eventType === 'video_completed' && (!endEvent || timestamp > endEvent)) {
          endEvent = timestamp;
        }
        
        // Count pauses
        if (eventType === 'video_paused') {
          pauseCount++;
        }
      }
    }
    
    // Set viewing statistics
    if (startEvent && endEvent) {
      report.viewingStatistics.startTime = startEvent;
      report.viewingStatistics.endTime = endEvent;
      report.viewingStatistics.totalTimeSpent = 
        (endEvent.getTime() - startEvent.getTime()) / 1000; // in seconds
      report.viewingStatistics.completionPercentage = 100; // Assuming completed if we have end event
    }
    
    report.viewingStatistics.pauseCount = pauseCount;
    
    // Count notes
    for (let i = 1; i < notesData.length; i++) {
      const row = notesData[i];
      
      // Only include data for this session and video
      if (row[5] === sessionId && row[2] === videoTitle) {
        report.notesCount++;
      }
    }
    
    // Generate summary text
    report.summary = generateReportSummary(report);
    
    return report;
  } catch (error) {
    Logger.log("Error in getStudentReport: " + error.toString());
    return { 
      error: "Failed to generate student report: " + error.toString(),
      summary: createDefaultReportSummary()
    };
  }
}

/**
 * Creates default report summary if no data is available
 * @returns {Object} Default summary object
 */
function createDefaultReportSummary() {
  return {
    title: "Activity Completed",
    message: "You have completed this video activity.",
    grade: "N/A",
    feedback: "No quiz data available for this session."
  };
}

/**
 * Generates a human-readable summary from report data
 * @param {Object} report - The performance report data
 * @returns {Object} Summary object with title, message, grade, and feedback
 */
function generateReportSummary(report) {
  // Initialize summary
  const summary = {
    title: "Activity Completed",
    message: `You have completed "${report.videoTitle}".`,
    grade: "N/A",
    feedback: ""
  };
  
  // Add quiz performance data if available
  if (report.quizPerformance.totalQuestions > 0) {
    const accuracy = report.quizPerformance.accuracyPercentage.toFixed(1);
    
    summary.grade = accuracy + "%";
    
    // Generate feedback based on performance
    if (accuracy >= 90) {
      summary.feedback = "Excellent work! You demonstrated a strong understanding of the material.";
    } else if (accuracy >= 75) {
      summary.feedback = "Good job! You have a solid grasp of most concepts.";
    } else if (accuracy >= 60) {
      summary.feedback = "You're on the right track, but may want to review the material again to strengthen your understanding.";
    } else {
      summary.feedback = "It looks like you might need to revisit this content. Consider rewatching the video and paying close attention to the key points.";
    }
    
    // Add detailed statistics
    summary.message += ` You answered ${report.quizPerformance.correctAnswers} out of ${report.quizPerformance.totalQuestions} questions correctly.`;
  } else {
    summary.feedback = "No quiz data available for this session.";
  }
  
  // Add note-taking info if applicable
  if (report.notesCount > 0) {
    summary.message += ` You took ${report.notesCount} note${report.notesCount !== 1 ? 's' : ''} during the video.`;
  }
  
  return summary;
}

/**
 * Generates quiz performance report for teacher view
 * @param {string} videoTitle - Optional: filter by video title
 * @returns {Object} Quiz performance data or error
 */
function getQuizPerformanceReport(videoTitle = null) {
  try {
    const ss = SpreadsheetApp.getActive();
    const analyticsSheet = ss.getSheetByName(CONFIG.SHEETS.QUIZ_ANALYTICS);
    
    if (!analyticsSheet) {
      return { error: "No quiz analytics data available" };
    }
    
    const analyticsData = analyticsSheet.getDataRange().getValues();
    const report = {
      totalAttempts: 0,
      correctAttempts: 0,
      incorrectAttempts: 0,
      averageTimeToAnswer: 0,
      quizzesByOverlay: {},
      userPerformance: {}
    };
    
    let totalTimeToAnswer = 0;
    
    // Skip header row
    for (let i = 1; i < analyticsData.length; i++) {
      const row = analyticsData[i];
      const rowVideoTitle = row[2];
      
      // Skip if filtering by video title and this row doesn't match
      if (videoTitle && rowVideoTitle !== videoTitle) {
        continue;
      }
      
      const userId = row[1];
      const overlayId = row[3];
      const wasCorrect = row[5] === true || row[5] === "TRUE";
      const timeToAnswer = parseFloat(row[7]) || 0;
      
      // Update totals
      report.totalAttempts++;
      if (wasCorrect) {
        report.correctAttempts++;
      } else {
        report.incorrectAttempts++;
      }
      
      totalTimeToAnswer += timeToAnswer;
      
      // Update by overlay
      if (!report.quizzesByOverlay[overlayId]) {
        report.quizzesByOverlay[overlayId] = {
          totalAttempts: 0,
          correctAttempts: 0,
          incorrectAttempts: 0
        };
      }
      
      report.quizzesByOverlay[overlayId].totalAttempts++;
      if (wasCorrect) {
        report.quizzesByOverlay[overlayId].correctAttempts++;
      } else {
        report.quizzesByOverlay[overlayId].incorrectAttempts++;
      }
      
      // Update by user
      if (!report.userPerformance[userId]) {
        report.userPerformance[userId] = {
          totalAttempts: 0,
          correctAttempts: 0,
          incorrectAttempts: 0
        };
      }
      
      report.userPerformance[userId].totalAttempts++;
      if (wasCorrect) {
        report.userPerformance[userId].correctAttempts++;
      } else {
        report.userPerformance[userId].incorrectAttempts++;
      }
    }
    
    // Calculate averages
    if (report.totalAttempts > 0) {
      report.averageTimeToAnswer = totalTimeToAnswer / report.totalAttempts;
      report.correctPercentage = (report.correctAttempts / report.totalAttempts) * 100;
    }
    
    // Calculate percentages for each overlay and user
    for (const overlayId in report.quizzesByOverlay) {
      const overlay = report.quizzesByOverlay[overlayId];
      if (overlay.totalAttempts > 0) {
        overlay.correctPercentage = (overlay.correctAttempts / overlay.totalAttempts) * 100;
      }
    }
    
    for (const userId in report.userPerformance) {
      const user = report.userPerformance[userId];
      if (user.totalAttempts > 0) {
        user.correctPercentage = (user.correctAttempts / user.totalAttempts) * 100;
      }
    }
    
    return report;
  } catch (error) {
    Logger.log("Error in getQuizPerformanceReport: " + error.toString());
    return { error: "Failed to generate report: " + error.toString() };
  }
}

/**
 * Updates application settings
 * @param {Object} settings - New settings object
 * @returns {Object} Success message or error
 */
function updateAppSettings(settings) {
  try {
    const ss = SpreadsheetApp.getActive();
    let settingsSheet = ss.getSheetByName(CONFIG.SHEETS.SETTINGS);
    
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet(CONFIG.SHEETS.SETTINGS);
      settingsSheet.appendRow(['Setting', 'Value', 'Description']);
    }
    
    // Clear existing settings
    const lastRow = settingsSheet.getLastRow();
    if (lastRow > 1) {
      settingsSheet.deleteRows(2, lastRow - 1);
    }
    
    // Add new settings
    for (const setting in settings) {
      settingsSheet.appendRow([setting, settings[setting].toString()]);
    }
    
    return { success: true, message: "Settings updated successfully" };
  } catch (error) {
    Logger.log("Error in updateAppSettings: " + error.toString());
    return { error: "Failed to update settings: " + error.toString() };
  }
}

/**
 * Shuffles array elements randomly
 * @param {Array} array - The array to shuffle
 * @returns {Array} The shuffled array
 */
function shuffleArray(array) {
  const newArray = [...array];
  
  for (let i = newArray.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [newArray[i], newArray[j]] = [newArray[j], newArray[i]];
  }
  
  return newArray;
}

/**
 * Extracts YouTube video ID from URL
 * @param {string} url - The YouTube URL
 * @returns {string|null} The video ID or null if invalid URL
 */
function extractYouTubeVideoId(url) {
  if (!url) return null;
  
  // Handle various YouTube URL formats
  const match = url.match(/(?:youtube\.com\/(?:[^\/]+\/.+\/|(?:v|e(?:mbed)?)\/|.*[?&]v=)|youtu\.be\/)([^"&?\/\s]{11})/);
  
  return match ? match[1] : null;
}