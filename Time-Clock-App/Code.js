// Global variables 
const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
const ss = SpreadsheetApp.openById(SPREADSHEET_ID);





function doGet(e) {
  const page = e.parameter.page || '';
  Logger.log('Page parameter: ' + page); // Log to confirm the parameter
  if (page === 'tv') {
    return HtmlService
      .createHtmlOutputFromFile('TVDisplay')
      .setTitle('Employee Status Board')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3666/3666228.png');
  } else if (page === 'mobile') {
    return HtmlService
      .createHtmlOutputFromFile('MobileDisplay')
      .setTitle('Employee Time Tracking')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3666/3666228.png');
  } else if (page === 'manager') {
    return HtmlService
      .createHtmlOutputFromFile('ManagerDashboard')
      .setTitle('Employee Time Tracking')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3666/3666228.png');
  }else {
    return HtmlService
      .createHtmlOutputFromFile('Kiosk')
      .setTitle('Employee Time Tracking')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3666/3666228.png');
  }
}



// Global cache for storing break timers
// This will persist between function calls but not between script executions
const BREAK_TIMERS = {};

/**
 * Records the start time of an employee's break
 * @param {string} employeeId - The employee's ID
 * @param {string} breakType - The type of break ('regular' or 'lunch')
 * @param {Date} startTime - The break start time (optional, defaults to current time)
 */
function recordBreakStart(employeeId, breakType, startTime = null) {
  if (!employeeId || !breakType) return;
  
  // Convert employeeId to string to ensure consistent keys
  const empId = String(employeeId);
  
  // Use provided time or current time
  const breakStartTime = startTime || new Date();
  
  // Store break information
  BREAK_TIMERS[empId] = {
    employeeId: empId,
    breakType: breakType,
    startTime: breakStartTime,
    timeLimit: breakType === 'lunch' ? 30 : 15 // 30 min for lunch, 15 min for regular breaks
  };
  
  Logger.log(`Break timer started for employee ${empId}: ${breakType} break at ${breakStartTime.toISOString()}`);
}

/**
 * Clears the break timer for an employee
 * @param {string} employeeId - The employee's ID
 */
function clearBreakTimer(employeeId) {
  if (!employeeId) return;
  
  // Convert employeeId to string to ensure consistent keys
  const empId = String(employeeId);
  
  // Check if there's a timer for this employee
  if (BREAK_TIMERS[empId]) {
    Logger.log(`Break timer cleared for employee ${empId}`);
    delete BREAK_TIMERS[empId];
  }
}

/**
 * Gets the current break timer for an employee
 * @param {string} employeeId - The employee's ID
 * @return {Object|null} Break timer information or null if not on break
 */
function getBreakTimer(employeeId) {
  if (!employeeId) return null;
  
  // Convert employeeId to string to ensure consistent keys
  const empId = String(employeeId);
  
  return BREAK_TIMERS[empId] || null;
}






// Function to initialize spreadsheet (used by getLiveEmployeeStatus)
function initSpreadsheet() {
    try {
        if (typeof ss === 'undefined' || !ss) {
          ss = SpreadsheetApp.getActiveSpreadsheet();
          if (!ss) {
            Logger.log("No active spreadsheet found");
            return false;
          }
        }
        return true;
      } catch (e) {
        Logger.log("Error initializing spreadsheet: " + e.toString());
        return false;
      }
}


// Function to authenticate employee
function authenticateEmployee(employeeId, pin) {
  const employeeSheet = ss.getSheetByName('Employee Master Data');
  const employeeData = employeeSheet.getDataRange().getValues();
  
  // Skip header row
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][0] == employeeId && employeeData[i][5] == pin && employeeData[i][8] == "Active") {
      return {
        success: true,
        employeeId: employeeData[i][0],
        firstName: employeeData[i][1],
        lastName: employeeData[i][2],
        department: employeeData[i][3]
      };
    }
  }
  return { success: false, message: "Invalid PIN" };
}

// Modified clockIn function to include total missed minutes for the pay period
function clockIn(employeeId) {
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const today = new Date();
    const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    // Store full datetime for clock-in
    const fullDateTimeStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
    // Generate log ID
    const logId = "TL" + today.getTime().toString().slice(-8);
    
    // Check if employee has any incomplete time logs (regardless of date)
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    for (let i = 1; i < timeLogsData.length; i++) {
      if (timeLogsData[i][1] == employeeId && timeLogsData[i][15] == "Incomplete") {
        const incompleteDate = Utilities.formatDate(new Date(timeLogsData[i][2]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        return { 
          success: false, 
          message: `You have an incomplete time log from ${incompleteDate}. Please complete that clock-out first.` 
        };
      }
    }
    
    // Check if employee is late based on their assigned shift
    const lateMinutes = checkIfLate(employeeId, today);
    
    // Add new time log with full datetime format
    timeLogsSheet.appendRow([
      logId,                  // A: Log ID
      employeeId,             // B: Employee ID
      fullDateTimeStr,        // C: Date
      fullDateTimeStr,        // D: Clock-in time (FULL DATETIME)
      "",                     // E: Clock-out time
      "",                     // F: Regular Break Start 1
      "",                     // G: Regular Break End 1
      "",                     // H: Regular Break Start 2
      "",                     // I: Regular Break End 2
      "",                     // J: Lunch Break Start
      "",                     // K: Lunch Break End
      "",                     // L: Total Hours (formula added at clock-out)
      "",                     // M: Regular Break Total (formula added at clock-out)
      "",                     // N: Lunch Break Total (formula added at clock-out)
      "",                     // O: Net Hours (formula added at clock-out)
      "Incomplete",           // P: Status
      "",                     // Q: Notes
      "",                     // R: Regular Break 1 Total (NEW)
      "",                     // S: Regular Break 2 Total (NEW)
      "",                     // T: Lunch Break Total (NEW)
      "",                     // U: Regular Break 1 Missed Minutes (NEW)
      "",                     // V: Regular Break 2 Missed Minutes (NEW)
      "",                     // W: Lunch Break Missed Minutes (NEW)
      lateMinutes,            // X: Late Minutes (NEW)
      "",                     // Y: Early Departure Minutes (NEW)
      "",                     // Z: Total Missed Minutes (NEW)
    ]);
    
    // Get the row number for the newly added row
    const newRow = timeLogsSheet.getLastRow();
  
    // Set dynamic formulas for calculations based on current row
    let totalHoursFormula = `=IF(AND(ISDATE(D${newRow}),ISDATE(E${newRow})),DAYS(E${newRow},D${newRow})*24+HOUR(E${newRow})-HOUR(D${newRow})+(MINUTE(E${newRow})-MINUTE(D${newRow}))/60+(SECOND(E${newRow})-SECOND(D${newRow}))/3600,"")`;
    timeLogsSheet.getRange(newRow, 12).setFormula(totalHoursFormula);
    
    let regularBreakFormula = `=IF(AND(ISDATE(F${newRow}),ISDATE(G${newRow})),DAYS(G${newRow},F${newRow})*24+HOUR(G${newRow})-HOUR(F${newRow})+(MINUTE(G${newRow})-MINUTE(F${newRow}))/60+(SECOND(G${newRow})-SECOND(F${newRow}))/3600,0) + IF(AND(ISDATE(H${newRow}),ISDATE(I${newRow})),DAYS(I${newRow},H${newRow})*24+HOUR(I${newRow})-HOUR(H${newRow})+(MINUTE(I${newRow})-MINUTE(H${newRow}))/60+(SECOND(I${newRow})-SECOND(H${newRow}))/3600,0)`;
    timeLogsSheet.getRange(newRow, 13).setFormula(regularBreakFormula);
    
    let lunchBreakFormula = `=IF(AND(ISDATE(J${newRow}),ISDATE(K${newRow})),DAYS(K${newRow},J${newRow})*24+HOUR(K${newRow})-HOUR(J${newRow})+(MINUTE(K${newRow})-MINUTE(J${newRow}))/60+(SECOND(K${newRow})-SECOND(J${newRow}))/3600,0)`;
    timeLogsSheet.getRange(newRow, 14).setFormula(lunchBreakFormula);
    
    let netHoursFormula = `=IF(L${newRow}<>"",MAX(0,L${newRow}-M${newRow}-N${newRow}),"")`;
    timeLogsSheet.getRange(newRow, 15).setFormula(netHoursFormula);
    
    // Adding formulas for individual break calculations
    let regBreak1TotalFormula = `=IF(AND(ISDATE(F${newRow}),ISDATE(G${newRow})),(G${newRow}-F${newRow})*24*60,"")`;
    timeLogsSheet.getRange(newRow, 18).setFormula(regBreak1TotalFormula);
    
    let regBreak2TotalFormula = `=IF(AND(ISDATE(H${newRow}),ISDATE(I${newRow})),(I${newRow}-H${newRow})*24*60,"")`;
    timeLogsSheet.getRange(newRow, 19).setFormula(regBreak2TotalFormula);
    
    let lunchBreakTotalFormula = `=IF(AND(ISDATE(J${newRow}),ISDATE(K${newRow})),(K${newRow}-J${newRow})*24*60,"")`;
    timeLogsSheet.getRange(newRow, 20).setFormula(lunchBreakTotalFormula);
    
    // Column U - Missed Minutes for Regular Break 1
    let missedBreak1Formula = `=IF(AND(ISNUMBER(R${newRow}),R${newRow}>15),R${newRow}-15,"")`;
    timeLogsSheet.getRange(newRow, 21).setFormula(missedBreak1Formula);
    
    // Column V - Missed Minutes for Regular Break 2
    let missedBreak2Formula = `=IF(AND(ISNUMBER(S${newRow}),S${newRow}>15),S${newRow}-15,"")`;
    timeLogsSheet.getRange(newRow, 22).setFormula(missedBreak2Formula);
    
    // Column W - Missed Minutes for Lunch Break
    let missedLunchFormula = `=IF(AND(ISNUMBER(T${newRow}),T${newRow}>30),T${newRow}-30,"")`;
    timeLogsSheet.getRange(newRow, 23).setFormula(missedLunchFormula);
    
    // Column Z - Total Missed Minutes (including late and early departure)
    let totalMissedFormula = `=SUM(IF(ISBLANK(U${newRow}),0,U${newRow}),IF(ISBLANK(V${newRow}),0,V${newRow}),IF(ISBLANK(W${newRow}),0,W${newRow}),IF(ISBLANK(X${newRow}),0,X${newRow}),IF(ISBLANK(Y${newRow}),0,Y${newRow}))`;
    timeLogsSheet.getRange(newRow, 26).setFormula(totalMissedFormula);
    
    // Add a note if employee is late
    if (lateMinutes > 0) {
      timeLogsSheet.getRange(newRow, 17).setValue("Late clock-in");
    }
    
    // Get total missed minutes for the current pay period
    let payPeriodMissedMinutes = { total: 0 };
    if (lateMinutes > 0) {
      payPeriodMissedMinutes = getEmployeePayPeriodMissedMinutes(employeeId);
    }
    
    return { 
      success: true, 
      message: "Clock-in successful", 
      logId: logId, 
      lateMinutes: lateMinutes > 0 ? lateMinutes : 0,
      payPeriodMissedMinutes: payPeriodMissedMinutes.total
    };
}



/**
 * Creates a new time log entry
 * @param {Object} timeLogData - The time log data
 * @return {Object} Result of the operation
 */
function createTimeLog(timeLogData) {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const today = new Date();
    
    // Generate log ID
    const logId = "TL" + today.getTime().toString().slice(-8);
    
    // Parse the clock-in datetime
    const clockInDateTime = new Date(timeLogData.clockInDateTime);
    
    // Check if employee exists
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    const employeeData = employeeSheet.getDataRange().getValues();
    let employeeExists = false;
    
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] == timeLogData.employeeId) {
        employeeExists = true;
        break;
      }
    }
    
    if (!employeeExists) {
      return { success: false, message: "Employee not found" };
    }
    
    // Check if employee is late based on their assigned shift
    let lateMinutes = 0;
    if (clockInDateTime) {
      lateMinutes = checkIfLate(timeLogData.employeeId, clockInDateTime);
    }
    
    // Check if early departure if clock-out time is provided
    let earlyMinutes = 0;
    let clockOutDateTime = null;
    if (timeLogData.clockOutDateTime) {
      clockOutDateTime = new Date(timeLogData.clockOutDateTime);
      earlyMinutes = checkIfEarlyDeparture(timeLogData.employeeId, clockOutDateTime);
    }
    
    // Determine status
    const status = clockOutDateTime ? "Complete" : "Incomplete";
    
    // Prepare notes
    let notes = timeLogData.notes || `Manually created from Manager Dashboard at ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss")}`;
    if (lateMinutes > 0) {
      notes = notes ? notes + ", Late clock-in" : "Late clock-in";
    }
    if (earlyMinutes > 0) {
      notes = notes ? notes + ", Early departure" : "Early departure";
    }
    
    // Extract date from clock-in datetime for the date column
    const logDate = new Date(clockInDateTime);
    logDate.setHours(0, 0, 0, 0);
    
    // Add new time log with properly formatted dates
    timeLogsSheet.appendRow([
      logId,                  // A: Log ID
      timeLogData.employeeId, // B: Employee ID
      clockInDateTime,        // C: Date (Same as Clock-in date/time)
      clockInDateTime,        // D: Clock-in time (properly formatted date object)
      clockOutDateTime || "", // E: Clock-out time (properly formatted date object or empty)
      "",                     // F: Regular Break Start 1
      "",                     // G: Regular Break End 1
      "",                     // H: Regular Break Start 2
      "",                     // I: Regular Break End 2
      "",                     // J: Lunch Break Start
      "",                     // K: Lunch Break End
      "",                     // L: Total Hours (formula added below)
      "",                     // M: Regular Break Total (formula added below)
      "",                     // N: Lunch Break Total (formula added below)
      "",                     // O: Net Hours (formula added below)
      status,                 // P: Status
      notes,                  // Q: Notes
      "",                     // R: Regular Break 1 Total (NEW)
      "",                     // S: Regular Break 2 Total (NEW)
      "",                     // T: Lunch Break Total (NEW)
      "",                     // U: Regular Break 1 Missed Minutes (NEW)
      "",                     // V: Regular Break 2 Missed Minutes (NEW)
      "",                     // W: Lunch Break Missed Minutes (NEW)
      lateMinutes > 0 ? lateMinutes : "", // X: Late Minutes (NEW)
      earlyMinutes > 0 ? earlyMinutes : "", // Y: Early Departure Minutes (NEW)
      "",                     // Z: Total Missed Minutes (NEW)
    ]);
    
    // Get the row number for the newly added row
    const newRow = timeLogsSheet.getLastRow();
  
    // Set dynamic formulas for calculations based on current row
    let totalHoursFormula = `=IF(AND(ISDATE(D${newRow}),ISDATE(E${newRow})),DAYS(E${newRow},D${newRow})*24+HOUR(E${newRow})-HOUR(D${newRow})+(MINUTE(E${newRow})-MINUTE(D${newRow}))/60+(SECOND(E${newRow})-SECOND(D${newRow}))/3600,"")`;
    timeLogsSheet.getRange(newRow, 12).setFormula(totalHoursFormula);
    
    let regularBreakFormula = `=IF(AND(ISDATE(F${newRow}),ISDATE(G${newRow})),DAYS(G${newRow},F${newRow})*24+HOUR(G${newRow})-HOUR(F${newRow})+(MINUTE(G${newRow})-MINUTE(F${newRow}))/60+(SECOND(G${newRow})-SECOND(F${newRow}))/3600,0) + IF(AND(ISDATE(H${newRow}),ISDATE(I${newRow})),DAYS(I${newRow},H${newRow})*24+HOUR(I${newRow})-HOUR(H${newRow})+(MINUTE(I${newRow})-MINUTE(H${newRow}))/60+(SECOND(I${newRow})-SECOND(H${newRow}))/3600,0)`;
    timeLogsSheet.getRange(newRow, 13).setFormula(regularBreakFormula);
    
    let lunchBreakFormula = `=IF(AND(ISDATE(J${newRow}),ISDATE(K${newRow})),DAYS(K${newRow},J${newRow})*24+HOUR(K${newRow})-HOUR(J${newRow})+(MINUTE(K${newRow})-MINUTE(J${newRow}))/60+(SECOND(K${newRow})-SECOND(J${newRow}))/3600,0)`;
    timeLogsSheet.getRange(newRow, 14).setFormula(lunchBreakFormula);
    
    let netHoursFormula = `=IF(L${newRow}<>"",MAX(0,L${newRow}-M${newRow}-N${newRow}),"")`;
    timeLogsSheet.getRange(newRow, 15).setFormula(netHoursFormula);
    
    // Adding formulas for individual break calculations
    let regBreak1TotalFormula = `=IF(AND(ISDATE(F${newRow}),ISDATE(G${newRow})),(G${newRow}-F${newRow})*24*60,"")`;
    timeLogsSheet.getRange(newRow, 18).setFormula(regBreak1TotalFormula);
    
    let regBreak2TotalFormula = `=IF(AND(ISDATE(H${newRow}),ISDATE(I${newRow})),(I${newRow}-H${newRow})*24*60,"")`;
    timeLogsSheet.getRange(newRow, 19).setFormula(regBreak2TotalFormula);
    
    let lunchBreakTotalFormula = `=IF(AND(ISDATE(J${newRow}),ISDATE(K${newRow})),(K${newRow}-J${newRow})*24*60,"")`;
    timeLogsSheet.getRange(newRow, 20).setFormula(lunchBreakTotalFormula);
    
    // Column U - Missed Minutes for Regular Break 1
    let missedBreak1Formula = `=IF(AND(ISNUMBER(R${newRow}),R${newRow}>15),R${newRow}-15,"")`;
    timeLogsSheet.getRange(newRow, 21).setFormula(missedBreak1Formula);
    
    // Column V - Missed Minutes for Regular Break 2
    let missedBreak2Formula = `=IF(AND(ISNUMBER(S${newRow}),S${newRow}>15),S${newRow}-15,"")`;
    timeLogsSheet.getRange(newRow, 22).setFormula(missedBreak2Formula);
    
    // Column W - Missed Minutes for Lunch Break
    let missedLunchFormula = `=IF(AND(ISNUMBER(T${newRow}),T${newRow}>30),T${newRow}-30,"")`;
    timeLogsSheet.getRange(newRow, 23).setFormula(missedLunchFormula);
    
    // Column Z - Total Missed Minutes (including late and early departure)
    let totalMissedFormula = `=SUM(IF(ISBLANK(U${newRow}),0,U${newRow}),IF(ISBLANK(V${newRow}),0,V${newRow}),IF(ISBLANK(W${newRow}),0,W${newRow}),IF(ISBLANK(X${newRow}),0,X${newRow}),IF(ISBLANK(Y${newRow}),0,Y${newRow}))`;
    timeLogsSheet.getRange(newRow, 26).setFormula(totalMissedFormula);
    
    return { 
      success: true, 
      message: "Time log created successfully", 
      logId: logId
    };
  } catch (error) {
    Logger.log("Error in createTimeLog: " + error.toString());
    return { success: false, message: error.toString() };
  }
}



  

// Modified clockOut function to include total missed minutes for the pay period
// and prioritize the most recent incomplete log
function clockOut(employeeId) {
  const timeLogsSheet = ss.getSheetByName('Time Logs');
  const today = new Date();
  
  // Store full datetime for clock-out
  const fullDateTimeStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
  
  // Find the most recent log for this employee
  const timeLogsData = timeLogsSheet.getDataRange().getValues();
  Logger.log("Looking for most recent incomplete log for employee: " + employeeId);
  
  // Track incomplete logs
  let mostRecentIncompleteLog = null;
  let mostRecentIncompleteLogIndex = -1;
  let mostRecentIncompleteLogDate = null;
  
  // Process from newest to oldest to find incomplete logs
  for (let i = timeLogsData.length - 1; i >= 1; i--) {
    try {
      // Skip rows without proper data
      if (!timeLogsData[i] || timeLogsData[i].length < 16) { // Check for status column too
        continue;
      }
      
      // Check if employee ID matches
      if (timeLogsData[i][1] != employeeId) {
        continue;
      }
      
      // Check if this is an incomplete log
      const status = timeLogsData[i][15] ? String(timeLogsData[i][15]) : "";
      const clockInTime = timeLogsData[i][3] ? timeLogsData[i][3] : "";
      const clockOutTime = timeLogsData[i][4] ? String(timeLogsData[i][4]) : "";
      
      // Consider a log incomplete if it has "Incomplete" status or no clock out time
      const isIncomplete = status === "Incomplete" || (!clockOutTime || clockOutTime.trim() === "");
      
      if (isIncomplete && clockInTime) {
        // If this is the first incomplete log we've found, or it's more recent than our previous one
        if (!mostRecentIncompleteLog || (clockInTime > mostRecentIncompleteLogDate)) {
          mostRecentIncompleteLog = timeLogsData[i];
          mostRecentIncompleteLogIndex = i;
          mostRecentIncompleteLogDate = clockInTime;
          Logger.log("Found incomplete log for employee " + employeeId + " at row " + (i+1) + " with date " + clockInTime);
        }
      }
    } catch (rowError) {
      Logger.log("Error processing row " + i + " during incomplete log search: " + rowError.toString());
      continue;
    }
  }
  
  // If no incomplete log was found
  if (!mostRecentIncompleteLog) {
    return { success: false, message: "No active clock-in found" };
  }
  
  // Convert to 1-indexed for sheet row
  const rowIndex = mostRecentIncompleteLogIndex + 1;
  const logDate = new Date(mostRecentIncompleteLog[2]); // Column C: Date
  
  Logger.log("Found most recent incomplete log at row " + rowIndex + " with date " + logDate);
  
  // Check if employee is leaving early (only if clocking out on the same day they clocked in)
  const logDateStr = Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  let earlyMinutes = 0;
  if (logDateStr === todayStr) {
    earlyMinutes = checkIfEarlyDeparture(employeeId, today);
  }
  
  // Update the time log with clock-out time and early departure info
  timeLogsSheet.getRange(rowIndex, 5).setValue(fullDateTimeStr); // Column E: Clock-out time
  timeLogsSheet.getRange(rowIndex, 16).setValue("Complete"); // Column P: Status
  
  // Add early departure minutes (only if applicable)
  if (earlyMinutes > 0) {
    timeLogsSheet.getRange(rowIndex, 25).setValue(earlyMinutes); // Column Y: Early Departure Minutes
    
    // Add a note if employee is leaving early
    const currentNotes = timeLogsSheet.getRange(rowIndex, 17).getValue();
    const newNotes = currentNotes ? currentNotes + ", Early departure" : "Early departure";
    timeLogsSheet.getRange(rowIndex, 17).setValue(newNotes); // Column Q: Notes
  }
  
  // If clocking out on a different day, add a note about it
  if (logDateStr !== todayStr) {
    const currentNotes = timeLogsSheet.getRange(rowIndex, 17).getValue();
    const newNotes = currentNotes ? 
      currentNotes + `, Clocked out on ${todayStr} (different day)` : 
      `Clocked out on ${todayStr} (different day)`;
    timeLogsSheet.getRange(rowIndex, 17).setValue(newNotes);
  }
  
  // Get total missed minutes for the current pay period
  let payPeriodMissedMinutes = { total: 0 };
  if (earlyMinutes > 0) {
    payPeriodMissedMinutes = getEmployeePayPeriodMissedMinutes(employeeId);
  }
  
  return { 
    success: true, 
    message: "Clock-out successful", 
    earlyMinutes: earlyMinutes > 0 ? earlyMinutes : 0,
    differentDay: logDateStr !== todayStr,
    payPeriodMissedMinutes: payPeriodMissedMinutes.total
  };
}



  




/**
 * Modifies the startBreak function to record break start time
 */
function startBreak(employeeId, breakType) {
  try {
    // Convert employeeId to string and ensure breakType is valid
    employeeId = String(employeeId);
    breakType = breakType === 'lunch' ? 'lunch' : 'regular';
    
    // Get current employee status
    const status = getEmployeeStatus(employeeId);
    
    // Check if employee is clocked in
    if (status.status !== "Clocked In") {
      return { success: false, message: "You must be clocked in to take a break" };
    }
    
    // Check if employee already has too many breaks
    if (breakType === 'regular' && status.regularBreaksTaken >= 2) {
      return { success: false, message: "You have already taken your allowed regular breaks" };
    }
    
    if (breakType === 'lunch' && status.lunchBreakTaken) {
      return { success: false, message: "You have already taken your lunch break" };
    }
    
    // Get the time logs sheet
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Find the current incomplete time log for this employee
    let rowIndex = -1;
    
    // First approach: Look for rows marked as "Incomplete"
    for (let i = 1; i < timeLogsData.length; i++) {
      if (
        String(timeLogsData[i][1]) === employeeId && 
        timeLogsData[i][15] === "Incomplete"
      ) {
        rowIndex = i + 1; // +1 because array is 0-based but sheet is 1-based
        break;
      }
    }
    
    // Second approach: If we couldn't find by "Incomplete", try using the status.logId
    if (rowIndex === -1 && status.logId) {
      for (let i = 1; i < timeLogsData.length; i++) {
        if (String(timeLogsData[i][0]) === String(status.logId)) {
          rowIndex = i + 1;
          break;
        }
      }
    }
    
    // Third approach: Look for any log with clock in but no clock out
    if (rowIndex === -1) {
      for (let i = timeLogsData.length - 1; i >= 1; i--) { // Start from most recent
        if (
          String(timeLogsData[i][1]) === employeeId && 
          (timeLogsData[i][3] && !timeLogsData[i][4]) // Has clock in but no clock out
        ) {
          rowIndex = i + 1;
          break;
        }
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: "No active time log found" };
    }
    
    // Current time
    const now = new Date();
    const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
    // Update the appropriate column based on break type and count
    if (breakType === 'regular') {
      if (status.regularBreaksTaken === 0) {
        // First regular break
        timeLogsSheet.getRange(rowIndex, 6).setValue(formattedTime); // Column F: Regular Break 1 Start
      } else {
        // Second regular break
        timeLogsSheet.getRange(rowIndex, 8).setValue(formattedTime); // Column H: Regular Break 2 Start
      }
    } else {
      // Lunch break
      timeLogsSheet.getRange(rowIndex, 10).setValue(formattedTime); // Column J: Lunch Break Start
    }
    
    // Record the break start time in our server-side cache
    recordBreakStart(employeeId, breakType, now);
    
    return { 
      success: true, 
      message: `${breakType === 'lunch' ? 'Lunch' : 'Regular'} break started`,
      breakType: breakType,
      startTime: formattedTime
    };
  } catch (error) {
    Logger.log("Error in startBreak: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}

// Modify the existing endBreak function
function endBreak(employeeId, breakType) {
  try {
    // Convert employeeId to string and ensure breakType is valid
    employeeId = String(employeeId);
    breakType = breakType === 'lunch' ? 'lunch' : 'regular';
    
    // Get current employee status
    const status = getEmployeeStatus(employeeId);
    
    // Check if employee is on the correct break type
    if (
      (breakType === 'regular' && status.status !== "On Regular Break") ||
      (breakType === 'lunch' && status.status !== "On Lunch Break")
    ) {
      return { success: false, message: `You are not currently on a ${breakType} break` };
    }
    
    // Get the time logs sheet
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Find the current incomplete time log for this employee
    let rowIndex = -1;
    let breakStartTime = null;
    let regularBreaksTaken = 0;
    
    // Look for the active time log with the matching break
    for (let i = timeLogsData.length - 1; i >= 1; i--) {
      // Check if this is a log for our employee
      if (String(timeLogsData[i][1]) === employeeId) {
        // Check if this log has clock in but no clock out (active log)
        const hasClockIn = timeLogsData[i][3] && String(timeLogsData[i][3]).trim() !== "";
        const hasClockOut = timeLogsData[i][4] && String(timeLogsData[i][4]).trim() !== "";
        
        if (hasClockIn && !hasClockOut) {
          // This is an active log, now check if it has the right break started
          if (breakType === 'regular') {
            // Check for regular break 1
            const hasBreak1Start = timeLogsData[i][5] && String(timeLogsData[i][5]).trim() !== "";
            const hasBreak1End = timeLogsData[i][6] && String(timeLogsData[i][6]).trim() !== "";
            
            // Check for regular break 2
            const hasBreak2Start = timeLogsData[i][7] && String(timeLogsData[i][7]).trim() !== "";
            const hasBreak2End = timeLogsData[i][8] && String(timeLogsData[i][8]).trim() !== "";
            
            // If break 1 is started but not ended
            if (hasBreak1Start && !hasBreak1End) {
              rowIndex = i + 1; // +1 because array is 0-based but sheet is 1-based
              breakStartTime = new Date(timeLogsData[i][5]);
              regularBreaksTaken = 1;
              break;
            } 
            // If break 2 is started but not ended
            else if (hasBreak2Start && !hasBreak2End) {
              rowIndex = i + 1;
              breakStartTime = new Date(timeLogsData[i][7]);
              regularBreaksTaken = 2;
              break;
            }
          } else { // lunch break
            // Check for lunch break
            const hasLunchStart = timeLogsData[i][9] && String(timeLogsData[i][9]).trim() !== "";
            const hasLunchEnd = timeLogsData[i][10] && String(timeLogsData[i][10]).trim() !== "";
            
            // If lunch is started but not ended
            if (hasLunchStart && !hasLunchEnd) {
              rowIndex = i + 1;
              breakStartTime = new Date(timeLogsData[i][9]);
              break;
            }
          }
        }
      }
    }
    
    if (rowIndex === -1) {
      // If we couldn't find the active log in the sheet, try to use the status data
      if (status.logId) {
        // Find the row with this log ID
        for (let i = 1; i < timeLogsData.length; i++) {
          if (String(timeLogsData[i][0]) === String(status.logId)) {
            rowIndex = i + 1;
            
            // Determine which break is being ended based on status
            if (breakType === 'regular') {
              // Check which regular break is active
              const hasBreak1Start = timeLogsData[i][5] && String(timeLogsData[i][5]).trim() !== "";
              const hasBreak1End = timeLogsData[i][6] && String(timeLogsData[i][6]).trim() !== "";
              const hasBreak2Start = timeLogsData[i][7] && String(timeLogsData[i][7]).trim() !== "";
              const hasBreak2End = timeLogsData[i][8] && String(timeLogsData[i][8]).trim() !== "";
              
              if (hasBreak1Start && !hasBreak1End) {
                breakStartTime = new Date(timeLogsData[i][5]);
                regularBreaksTaken = 1;
              } else if (hasBreak2Start && !hasBreak2End) {
                breakStartTime = new Date(timeLogsData[i][7]);
                regularBreaksTaken = 2;
              }
            } else { // lunch break
              const hasLunchStart = timeLogsData[i][9] && String(timeLogsData[i][9]).trim() !== "";
              if (hasLunchStart) {
                breakStartTime = new Date(timeLogsData[i][9]);
              }
            }
            break;
          }
        }
      }
    }
    
    // If we still don't have a row, try one more approach - look for any incomplete log
    if (rowIndex === -1) {
      for (let i = 1; i < timeLogsData.length; i++) {
        if (
          String(timeLogsData[i][1]) === employeeId && 
          (timeLogsData[i][15] === "Incomplete" || 
           (timeLogsData[i][3] && !timeLogsData[i][4])) // Has clock in but no clock out
        ) {
          rowIndex = i + 1;
          
          // Determine which break is being ended
          if (breakType === 'regular') {
            // Check which regular break is active
            if (timeLogsData[i][5] && !timeLogsData[i][6]) {
              breakStartTime = new Date(timeLogsData[i][5]);
              regularBreaksTaken = 1;
            } else if (timeLogsData[i][7] && !timeLogsData[i][8]) {
              breakStartTime = new Date(timeLogsData[i][7]);
              regularBreaksTaken = 2;
            }
          } else { // lunch break
            if (timeLogsData[i][9] && !timeLogsData[i][10]) {
              breakStartTime = new Date(timeLogsData[i][9]);
            }
          }
          break;
        }
      }
    }
    
    // If we still don't have a row index, we can't find the active log
    if (rowIndex === -1) {
      Logger.log("Could not find active time log for employee " + employeeId);
      return { success: false, message: "No active time log found" };
    }
    
    // If we don't have a break start time from the sheet, try to get it from the break timer
    if (!breakStartTime) {
      const breakTimer = getBreakTimer(employeeId);
      if (breakTimer) {
        breakStartTime = breakTimer.startTime;
      } else {
        Logger.log("Could not determine break start time for employee " + employeeId);
        return { success: false, message: "Could not determine break start time" };
      }
    }
    
    // Current time
    const now = new Date();
    const formattedTime = Utilities.formatDate(now, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
    // Calculate break duration in minutes
    const breakDurationMs = now - breakStartTime;
    const breakDurationMinutes = breakDurationMs / (1000 * 60);
    
    // Check if break was extended
    const timeLimit = breakType === 'lunch' ? 30 : 15;
    const extendedMinutes = breakDurationMinutes > timeLimit ? Math.round((breakDurationMinutes - timeLimit) * 100) / 100 : 0;
    
    Logger.log(`Ending ${breakType} break for employee ${employeeId} at row ${rowIndex}, started at ${breakStartTime}, duration: ${breakDurationMinutes.toFixed(2)} min`);
    
    // Update the appropriate column based on break type and count
    if (breakType === 'regular') {
      if (regularBreaksTaken === 1) {
        // First regular break
        timeLogsSheet.getRange(rowIndex, 7).setValue(formattedTime); // Column G: Regular Break 1 End
        
        // Update missed minutes if extended
        if (extendedMinutes > 0) {
          timeLogsSheet.getRange(rowIndex, 21).setValue(extendedMinutes); // Column U: Regular Break 1 Missed Minutes
        }
      } else {
        // Second regular break
        timeLogsSheet.getRange(rowIndex, 9).setValue(formattedTime); // Column I: Regular Break 2 End
        
        // Update missed minutes if extended
        if (extendedMinutes > 0) {
          timeLogsSheet.getRange(rowIndex, 22).setValue(extendedMinutes); // Column V: Regular Break 2 Missed Minutes
        }
      }
    } else {
      // Lunch break
      timeLogsSheet.getRange(rowIndex, 11).setValue(formattedTime); // Column K: Lunch Break End
      
      // Update missed minutes if extended
      if (extendedMinutes > 0) {
        timeLogsSheet.getRange(rowIndex, 23).setValue(extendedMinutes); // Column W: Lunch Break Missed Minutes
      }
    }
    
    // Clear the break timer from our server-side cache
    clearBreakTimer(employeeId);
    
    // Calculate total missed minutes
    let totalMissedMinutes = 0;
    
    // Get existing missed minutes (late arrival, early departure)
    const lateMinutes = timeLogsData[rowIndex-1][23] || 0; // Column X
    const earlyMinutes = timeLogsData[rowIndex-1][24] || 0; // Column Y
    
    // Add all missed minutes
    totalMissedMinutes = lateMinutes + earlyMinutes + extendedMinutes;
    
    // Get pay period missed minutes for this employee
    const payPeriodMissedMinutes = getEmployeePayPeriodMissedMinutes(employeeId);
    
    return { 
      success: true, 
      message: `${breakType === 'lunch' ? 'Lunch' : 'Regular'} break ended`,
      extendedMinutes: extendedMinutes,
      payPeriodMissedMinutes: payPeriodMissedMinutes
    };
  } catch (error) {
    Logger.log("Error in endBreak: " + error.toString());
    return { success: false, message: "Error: " + error.toString() };
  }
}



// Function to get employee's current status with improved error handling and date handling
function getEmployeeStatus(employeeId) {
  try {
    // Make sure spreadsheet is initialized
    if (typeof ss === 'undefined' || !ss) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) {
        Logger.log("No active spreadsheet found");
        return { status: "Error", message: "Failed to initialize spreadsheet" };
      }
    }
    
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    if (!timeLogsSheet) {
      Logger.log("Time Logs sheet not found");
      return { status: "Error", message: "Time Logs sheet not found" };
    }
    
    const today = new Date();
    Logger.log("Looking for most recent incomplete log for employee: " + employeeId);
    
    // Find the most recent log for this employee
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    Logger.log("Total rows in Time Logs: " + timeLogsData.length);
    
    // Track incomplete logs
    let mostRecentIncompleteLog = null;
    let mostRecentIncompleteLogIndex = -1;
    let mostRecentIncompleteLogDate = null;
    
    // Process from newest to oldest to find incomplete logs
    for (let i = timeLogsData.length - 1; i >= 1; i--) {
      try {
        // Skip rows without proper data
        if (!timeLogsData[i] || timeLogsData[i].length < 16) { // Check for status column too
          continue;
        }
        
        // Check if employee ID matches
        if (timeLogsData[i][1] != employeeId) {
          continue;
        }
        
        // Check if this is an incomplete log
        const status = timeLogsData[i][15] ? String(timeLogsData[i][15]) : "";
        const clockInTime = timeLogsData[i][3] ? timeLogsData[i][3] : "";
        const clockOutTime = timeLogsData[i][4] ? String(timeLogsData[i][4]) : "";
        
        // Consider a log incomplete if it has "Incomplete" status or no clock out time
        const isIncomplete = status === "Incomplete" || (!clockOutTime || clockOutTime.trim() === "");
        
        if (isIncomplete && clockInTime) {
          // If this is the first incomplete log we've found, or it's more recent than our previous one
          if (!mostRecentIncompleteLog || (clockInTime > mostRecentIncompleteLogDate)) {
            mostRecentIncompleteLog = timeLogsData[i];
            mostRecentIncompleteLogIndex = i;
            mostRecentIncompleteLogDate = clockInTime;
            Logger.log("Found incomplete log for employee " + employeeId + " at row " + (i+1) + " with date " + clockInTime);
          }
        }
      } catch (rowError) {
        Logger.log("Error processing row " + i + " during incomplete log search: " + rowError.toString());
        continue;
      }
    }
    
    // If we found an incomplete log, process it
    if (mostRecentIncompleteLog) {
      const i = mostRecentIncompleteLogIndex;
      try {
        // Get the time values, ensuring they are strings
        const clockInTime = mostRecentIncompleteLog[3] ? String(mostRecentIncompleteLog[3]) : "";
        const clockOutTime = mostRecentIncompleteLog[4] ? String(mostRecentIncompleteLog[4]) : "";
        const break1Start = mostRecentIncompleteLog[5] ? String(mostRecentIncompleteLog[5]) : "";
        const break1End = mostRecentIncompleteLog[6] ? String(mostRecentIncompleteLog[6]) : "";
        const break2Start = mostRecentIncompleteLog[7] ? String(mostRecentIncompleteLog[7]) : "";
        const break2End = mostRecentIncompleteLog[8] ? String(mostRecentIncompleteLog[8]) : "";
        const lunchStart = mostRecentIncompleteLog[9] ? String(mostRecentIncompleteLog[9]) : "";
        const lunchEnd = mostRecentIncompleteLog[10] ? String(mostRecentIncompleteLog[10]) : "";
        
        Logger.log("Processing most recent incomplete log for employee " + employeeId + " at row " + (i+1));
        Logger.log("Time values: clockIn=" + clockInTime + ", clockOut=" + clockOutTime);
        
        // Count breaks taken
        let regularBreaksTaken = 0;
        let lunchBreakTaken = false;
        
        // First regular break is used if it has both start and end times
        if (break1Start && break1Start.trim() !== "" && break1End && break1End.trim() !== "") {
          regularBreaksTaken++;
        }
        
        // Second regular break is used if it has both start and end times
        if (break2Start && break2Start.trim() !== "" && break2End && break2End.trim() !== "") {
          regularBreaksTaken++;
        }
        
        // First regular break is in progress if it has start but no end time
        const onRegularBreak1 = break1Start && break1Start.trim() !== "" && (!break1End || break1End.trim() === "");
        
        // Second regular break is in progress if it has start but no end time
        const onRegularBreak2 = break2Start && break2Start.trim() !== "" && (!break2End || break2End.trim() === "");
        
        // Lunch break is used if it has both start and end times
        if (lunchStart && lunchStart.trim() !== "" && lunchEnd && lunchEnd.trim() !== "") {
          lunchBreakTaken = true;
        }
        
        // Lunch break is in progress if it has start but no end time
        const onLunchBreak = lunchStart && lunchStart.trim() !== "" && (!lunchEnd || lunchEnd.trim() === "");
        
        // Check status for the incomplete log
        if (onLunchBreak) {
          // On lunch break
          const result = {
            status: "On Lunch Break",
            time: lunchStart,
            logId: mostRecentIncompleteLog[0],
            regularBreaksTaken: regularBreaksTaken,
            lunchBreakTaken: false, // Not fully taken yet, still in progress
            onBreak: "lunch"
          };
          
          // Add break timer information
          const breakTimer = getBreakTimer(employeeId);
          if (breakTimer) {
            // Use existing timer
            result.breakStartTime = breakTimer.startTime.toISOString();
            result.breakTimeLimit = breakTimer.timeLimit;
          } else {
            // Create new timer based on lunch start time
            try {
              const breakStartTime = new Date(lunchStart);
              recordBreakStart(employeeId, "lunch", breakStartTime);
              
              // Now get the newly created timer
              const newTimer = getBreakTimer(employeeId);
              if (newTimer) {
                result.breakStartTime = newTimer.startTime.toISOString();
                result.breakTimeLimit = newTimer.timeLimit;
              }
            } catch (timerError) {
              Logger.log("Error creating break timer: " + timerError.toString());
            }
          }
          
          return result;
        } else if (onRegularBreak1 || onRegularBreak2) {
          // On regular break
          const breakTime = onRegularBreak1 ? break1Start : break2Start;
          const breakNumber = onRegularBreak1 ? 1 : 2;
          
          const result = {
            status: "On Regular Break",
            time: breakTime,
            logId: mostRecentIncompleteLog[0],
            regularBreaksTaken: regularBreaksTaken,
            lunchBreakTaken: lunchBreakTaken,
            onBreak: "regular",
            breakNumber: breakNumber
          };
          
          // Add break timer information
          const breakTimer = getBreakTimer(employeeId);
          if (breakTimer) {
            // Use existing timer
            result.breakStartTime = breakTimer.startTime.toISOString();
            result.breakTimeLimit = breakTimer.timeLimit;
          } else {
            // Create new timer based on break start time
            try {
              const breakStartTime = new Date(breakTime);
              recordBreakStart(employeeId, "regular", breakStartTime);
              
              // Now get the newly created timer
              const newTimer = getBreakTimer(employeeId);
              if (newTimer) {
                result.breakStartTime = newTimer.startTime.toISOString();
                result.breakTimeLimit = newTimer.timeLimit;
              }
            } catch (timerError) {
              Logger.log("Error creating break timer: " + timerError.toString());
            }
          }
          
          return result;
        } else if (clockInTime && clockInTime.trim() !== "") {
          // Clocked in
          // Clear any break timer since employee is not on break
          clearBreakTimer(employeeId);
          
          return {
            status: "Clocked In",
            time: clockInTime,
            logId: mostRecentIncompleteLog[0],
            regularBreaksTaken: regularBreaksTaken,
            lunchBreakTaken: lunchBreakTaken,
            onBreak: null
          };
        }
      } catch (processError) {
        Logger.log("Error processing incomplete log: " + processError.toString());
      }
    }
    
    // If we didn't find or couldn't process an incomplete log, fall back to the original logic
    // Process from newest to oldest to get the most recent log
    for (let i = timeLogsData.length - 1; i >= 1; i--) {
      try {
        // Skip rows without proper data
        if (!timeLogsData[i] || timeLogsData[i].length < 11) {
          continue;
        }
        
        // Check if employee ID matches
        if (timeLogsData[i][1] != employeeId) {
          continue;
        }
        
        // Get the time values, ensuring they are strings
        const clockInTime = timeLogsData[i][3] ? String(timeLogsData[i][3]) : "";
        const clockOutTime = timeLogsData[i][4] ? String(timeLogsData[i][4]) : "";
        const break1Start = timeLogsData[i][5] ? String(timeLogsData[i][5]) : "";
        const break1End = timeLogsData[i][6] ? String(timeLogsData[i][6]) : "";
        const break2Start = timeLogsData[i][7] ? String(timeLogsData[i][7]) : "";
        const break2End = timeLogsData[i][8] ? String(timeLogsData[i][8]) : "";
        const lunchStart = timeLogsData[i][9] ? String(timeLogsData[i][9]) : "";
        const lunchEnd = timeLogsData[i][10] ? String(timeLogsData[i][10]) : "";
        
        Logger.log("Found log for employee " + employeeId + " at row " + (i+1));
        Logger.log("Time values: clockIn=" + clockInTime + ", clockOut=" + clockOutTime);
        
        // Count breaks taken
        let regularBreaksTaken = 0;
        let lunchBreakTaken = false;
        
        // First regular break is used if it has both start and end times
        if (break1Start && break1Start.trim() !== "" && break1End && break1End.trim() !== "") {
          regularBreaksTaken++;
        }
        
        // Second regular break is used if it has both start and end times
        if (break2Start && break2Start.trim() !== "" && break2End && break2End.trim() !== "") {
          regularBreaksTaken++;
        }
        
        // First regular break is in progress if it has start but no end time
        const onRegularBreak1 = break1Start && break1Start.trim() !== "" && (!break1End || break1End.trim() === "");
        
        // Second regular break is in progress if it has start but no end time
        const onRegularBreak2 = break2Start && break2Start.trim() !== "" && (!break2End || break2End.trim() === "");
        
        // Lunch break is used if it has both start and end times
        if (lunchStart && lunchStart.trim() !== "" && lunchEnd && lunchEnd.trim() !== "") {
          lunchBreakTaken = true;
        }
        
        // Lunch break is in progress if it has start but no end time
        const onLunchBreak = lunchStart && lunchStart.trim() !== "" && (!lunchEnd || lunchEnd.trim() === "");
        
        // Check if this is the most recent log
        if (i === timeLogsData.length - 1 || timeLogsData[i][1] != timeLogsData[i+1][1]) {
          // Most recent log for this employee
          if (!clockOutTime || clockOutTime.trim() === "") {
            if (onLunchBreak) {
              // On lunch break
              const result = {
                status: "On Lunch Break",
                time: lunchStart,
                logId: timeLogsData[i][0],
                regularBreaksTaken: regularBreaksTaken,
                lunchBreakTaken: false, // Not fully taken yet, still in progress
                onBreak: "lunch"
              };
              
              // Add break timer information
              const breakTimer = getBreakTimer(employeeId);
              if (breakTimer) {
                // Use existing timer
                result.breakStartTime = breakTimer.startTime.toISOString();
                result.breakTimeLimit = breakTimer.timeLimit;
              } else {
                // Create new timer based on lunch start time
                try {
                  const breakStartTime = new Date(lunchStart);
                  recordBreakStart(employeeId, "lunch", breakStartTime);
                  
                  // Now get the newly created timer
                  const newTimer = getBreakTimer(employeeId);
                  if (newTimer) {
                    result.breakStartTime = newTimer.startTime.toISOString();
                    result.breakTimeLimit = newTimer.timeLimit;
                  }
                } catch (timerError) {
                  Logger.log("Error creating break timer: " + timerError.toString());
                }
              }
              
              return result;
            } else if (onRegularBreak1 || onRegularBreak2) {
              // On regular break
              const breakTime = onRegularBreak1 ? break1Start : break2Start;
              const breakNumber = onRegularBreak1 ? 1 : 2;
              
              const result = {
                status: "On Regular Break",
                time: breakTime,
                logId: timeLogsData[i][0],
                regularBreaksTaken: regularBreaksTaken,
                lunchBreakTaken: lunchBreakTaken,
                onBreak: "regular",
                breakNumber: breakNumber
              };
              
              // Add break timer information
              const breakTimer = getBreakTimer(employeeId);
              if (breakTimer) {
                // Use existing timer
                result.breakStartTime = breakTimer.startTime.toISOString();
                result.breakTimeLimit = breakTimer.timeLimit;
              } else {
                // Create new timer based on break start time
                try {
                  const breakStartTime = new Date(breakTime);
                  recordBreakStart(employeeId, "regular", breakStartTime);
                  
                  // Now get the newly created timer
                  const newTimer = getBreakTimer(employeeId);
                  if (newTimer) {
                    result.breakStartTime = newTimer.startTime.toISOString();
                    result.breakTimeLimit = newTimer.timeLimit;
                  }
                } catch (timerError) {
                  Logger.log("Error creating break timer: " + timerError.toString());
                }
              }
              
              return result;
            } else if (clockInTime && clockInTime.trim() !== "") {
              // Clocked in
              // Clear any break timer since employee is not on break
              clearBreakTimer(employeeId);
              
              return {
                status: "Clocked In",
                time: clockInTime,
                logId: timeLogsData[i][0],
                regularBreaksTaken: regularBreaksTaken,
                lunchBreakTaken: lunchBreakTaken,
                onBreak: null
              };
            }
          } else {
            // Clocked out
            // Clear any break timer since employee is not on break
            clearBreakTimer(employeeId);
            
            return {
              status: "Clocked Out",
              time: clockOutTime,
              logId: timeLogsData[i][0],
              regularBreaksTaken: regularBreaksTaken,
              lunchBreakTaken: lunchBreakTaken,
              onBreak: null
            };
          }
        }
      } catch (rowError) {
        Logger.log("Error processing row " + i + ": " + rowError.toString());
        continue;
      }
    }
    
    Logger.log("No incomplete log found for employee " + employeeId);
    
    // Clear any break timer since employee is not on break
    clearBreakTimer(employeeId);
    
    return {
      status: "Not Clocked In",
      time: "",
      logId: "",
      regularBreaksTaken: 0,
      lunchBreakTaken: false,
      onBreak: null
    };
  } catch (error) {
    Logger.log("Error in getEmployeeStatus: " + error.toString());
    return {
      status: "Error",
      message: error.toString(),
      regularBreaksTaken: 0,
      lunchBreakTaken: false,
      onBreak: null
    };
  }
}





// Helper function to get start of week (Monday)
function getStartOfWeek(date) {
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
  return new Date(date.setDate(diff));
}

// Function to check and enforce break rules
function enforceBreakRules(employeeId) {
  const timeLogsSheet = ss.getSheetByName('Time Logs');
  const today = new Date();
  const todayStr = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  
  // Find today's log
  const timeLogsData = timeLogsSheet.getDataRange().getValues();
  for (let i = 1; i < timeLogsData.length; i++) {
    const rowDate = Utilities.formatDate(new Date(timeLogsData[i][2]), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (rowDate === todayStr && timeLogsData[i][1] == employeeId) {
      const row = i + 1;
      
      // Check if employee has worked more than 5 hours without a lunch break
      if (timeLogsData[i][3] && !timeLogsData[i][9]) {
        const clockInTime = new Date(`${todayStr}T${timeLogsData[i][3]}`);
        const hoursWorked = (today - clockInTime) / (1000 * 60 * 60);
        
        if (hoursWorked >= 5) {
          // Add a warning note
          const currentNote = timeLogsSheet.getRange(row, 17).getValue() || "";
          timeLogsSheet.getRange(row, 17).setValue(currentNote + " WARNING: 5+ hours worked without lunch break.");
          
          return {
            enforced: true,
            message: "You have worked more than 5 hours without a lunch break. Please take a break now to comply with labor regulations."
          };
        }
      }
      
      // Check if regular breaks are being taken (at least one every 3 hours)
      if (timeLogsData[i][3] && !timeLogsData[i][5] && !timeLogsData[i][7]) {
        const clockInTime = new Date(`${todayStr}T${timeLogsData[i][3]}`);
        const hoursWorked = (today - clockInTime) / (1000 * 60 * 60);
        
        if (hoursWorked >= 3) {
          // Add a warning note
          const currentNote = timeLogsSheet.getRange(row, 17).getValue() || "";
          timeLogsSheet.getRange(row, 17).setValue(currentNote + " WARNING: 3+ hours worked without a regular break.");
          
          return {
            enforced: true,
            message: "You have worked more than 3 hours without a break. Please take a short break now."
          };
        }
      }
      
      break;
    }
  }
  
  return { enforced: false };
}


/**
 * Get live status of all employees who are currently clocked in or on break
 */
function getLiveEmployeeStatus() {
  try {
    // Get all active employees
    const employees = getActiveEmployees();
    
    // Create a map of employee IDs to names for quick lookup
    const employeeMap = {};
    employees.forEach(emp => {
      employeeMap[emp.employeeId] = {
        name: `${emp.firstName} ${emp.lastName}`,
        department: emp.department
      };
    });
    
    // Get the time logs sheet
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Map to store the latest status for each employee
    const employeeStatuses = {};
    
    // Process time logs to determine current status
    for (let i = 1; i < timeLogsData.length; i++) {
      const employeeId = String(timeLogsData[i][1]);
      const status = timeLogsData[i][15]; // Column P: Status
      
      // Only process active employees
      if (!employeeMap[employeeId]) continue;
      
      // Skip completed logs
      if (status === "Complete") continue;
      
      // Determine current status based on time log entries
      let currentStatus = "Not Clocked In";
      let statusTime = null;
      
      if (timeLogsData[i][3]) { // Has clock-in time
        if (timeLogsData[i][9] && !timeLogsData[i][10]) { // On lunch break
          currentStatus = "On Lunch Break";
          statusTime = new Date(timeLogsData[i][9]); // Lunch break start time
        } else if (
          (timeLogsData[i][5] && !timeLogsData[i][6]) || // On first regular break
          (timeLogsData[i][7] && !timeLogsData[i][8])    // On second regular break
        ) {
          currentStatus = "On Regular Break";
          statusTime = timeLogsData[i][5] && !timeLogsData[i][6] ? 
                       new Date(timeLogsData[i][5]) : 
                       new Date(timeLogsData[i][7]);
        } else if (!timeLogsData[i][4]) { // No clock-out time
          currentStatus = "Clocked In";
          statusTime = new Date(timeLogsData[i][3]); // Clock-in time
        }
      }
      
      // Store the status
      employeeStatuses[employeeId] = {
        employeeId: employeeId,
        name: employeeMap[employeeId].name,
        department: employeeMap[employeeId].department,
        status: currentStatus,
        time: statusTime ? Utilities.formatDate(statusTime, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") : null
      };
      
      // Add break timer information if on break
      if (currentStatus === "On Regular Break" || currentStatus === "On Lunch Break") {
        // Check if we have a break timer for this employee
        const breakTimer = getBreakTimer(employeeId);
        
        if (breakTimer) {
          // Use the cached break timer
          employeeStatuses[employeeId].breakStartTime = breakTimer.startTime.toISOString();
          employeeStatuses[employeeId].breakTimeLimit = breakTimer.timeLimit;
        } else {
          // If no cached timer, create one based on the status time
          recordBreakStart(
            employeeId, 
            currentStatus === "On Regular Break" ? 'regular' : 'lunch',
            statusTime
          );
          
          // Now get the newly created timer
          const newTimer = getBreakTimer(employeeId);
          if (newTimer) {
            employeeStatuses[employeeId].breakStartTime = newTimer.startTime.toISOString();
            employeeStatuses[employeeId].breakTimeLimit = newTimer.timeLimit;
          }
        }
      }
    }
    
    // Convert to array
    const result = Object.values(employeeStatuses);
    
    // Add employees who aren't in the status map (not clocked in)
    employees.forEach(emp => {
      if (!employeeStatuses[emp.employeeId]) {
        result.push({
          employeeId: emp.employeeId,
          name: `${emp.firstName} ${emp.lastName}`,
          department: emp.department,
          status: "Not Clocked In",
          time: null
        });
      }
    });
    
    return result;
  } catch (error) {
    Logger.log("Error in getLiveEmployeeStatus: " + error.toString());
    throw new Error("Failed to get employee status: " + error.toString());
  }
}

function getActiveEmployees() {
  // Initialize the spreadsheet if not already done
  if (!initSpreadsheet()) return [];
  
  const employeeSheet = ss.getSheetByName('Employee Master Data');
  if (!employeeSheet) {
    Logger.log("Employee Master Data sheet not found");
    return [];
  }
  
  const employeeData = employeeSheet.getDataRange().getValues();
  const employees = [];
  
  Logger.log("Processing " + employeeData.length + " rows of employee data");
  
  // Skip header row
  for (let i = 1; i < employeeData.length; i++) {
    // Add debug logging
    Logger.log("Processing row " + i + ": " + employeeData[i].join(", "));
    
    // Check if the employee is active (column I/index 8)
    if (employeeData[i][8] === "Active") {
      employees.push({
        employeeId: employeeData[i][0],
        firstName: employeeData[i][1],
        lastName: employeeData[i][2],
        department: employeeData[i][3]
      });
    }
  }
  
  // Sort by name
  employees.sort((a, b) => a.firstName.localeCompare(b.firstName));
  
  Logger.log("Found " + employees.length + " active employees");
  return employees;
}

/**
 * Properly calculates hours between two times, handling string formats and overnight shifts
 */
// Improvement to the calculateHours function
function calculateHours(startTime, endTime, startDate, endDate) {
  try {
    // Ensure inputs are strings to prevent errors
    startTime = String(startTime || "").trim();
    endTime = String(endTime || "").trim();
    
    // Early validation
    if (!startTime || !endTime) return 0;
    
    // Force 24-hour time format interpretation
    const startParts = startTime.split(":");
    const endParts = endTime.split(":");
    
    if (startParts.length < 2 || endParts.length < 2) return 0;
    
    // Create date objects for the same day
    const baseDate = new Date();
    baseDate.setHours(0, 0, 0, 0); // Reset to midnight
    
    const startDateTime = new Date(baseDate);
    startDateTime.setHours(
      parseInt(startParts[0], 10),
      parseInt(startParts[1], 10),
      startParts[2] ? parseInt(startParts[2], 10) : 0
    );
    
    const endDateTime = new Date(baseDate);
    endDateTime.setHours(
      parseInt(endParts[0], 10),
      parseInt(endParts[1], 10), 
      endParts[2] ? parseInt(endParts[2], 10) : 0
    );
    
    // Calculate hours difference
    let hoursDiff = (endDateTime - startDateTime) / (1000 * 60 * 60);
    
    // Handle overnight shift if needed
    if (hoursDiff < 0) {
      hoursDiff += 24;
    }
    
    return hoursDiff;
  } catch (e) {
    return 0;
  }
}


/**
 * Check if two dates are the same day
 */
function sameDay(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

/**
 * Helper function to format date as YYYY-MM-DD
 */
function formatDate(date) {
  if (!(date instanceof Date)) {
    date = new Date(date);
  }
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

/**
 * Gets an employee's shift details
 * @param {string|number} employeeId - The employee ID
 * @return {Object} Shift details or null if not found
 */
function getEmployeeShift(employeeId) {
    try {
      // Get employee's assigned shift
      const employeeSheet = ss.getSheetByName('Employee Master Data');
      const employeeData = employeeSheet.getDataRange().getValues();
      
      let shiftId = null;
      let employeeName = "";
      
      // Find the employee's shift ID
      for (let i = 1; i < employeeData.length; i++) {
        if (employeeData[i][0] == employeeId) {
          shiftId = employeeData[i][10]; // Column K has Shift ID
          employeeName = employeeData[i][1] + " " + employeeData[i][2]; // First and Last name
          break;
        }
      }
      
      if (!shiftId) {
        Logger.log("No shift assigned to employee ID: " + employeeId);
        return null; // No shift assigned
      }
      
      // Get shift details
      const shiftSheet = ss.getSheetByName('Shifts');
      const shiftData = shiftSheet.getDataRange().getValues();
      
      // Find the shift details
      for (let i = 1; i < shiftData.length; i++) {
        if (shiftData[i][0] == shiftId) {
          return {
            shiftId: shiftId,
            shiftName: shiftData[i][1], // Column B has Shift Name
            isOvernight: shiftData[i][2], // Column C has Is Overnight
            weekAStartTime: shiftData[i][3], // Column D has Week A Start Time
            weekAEndTime: shiftData[i][4], // Column E has Week A End Time
            weekBStartTime: shiftData[i][5], // Column F has Week B Start Time (if applicable)
            weekBEndTime: shiftData[i][6], // Column G has Week B End Time (if applicable)
            regBreakDuration: 15, // Default regular break duration (15 minutes)
            lunchDuration: 30, // Default lunch break duration (30 minutes)
            employeeName: employeeName
          };
        }
      }
      
      Logger.log("Shift ID found but details not available: " + shiftId);
      return null;
      
    } catch (e) {
      Logger.log("Error in getEmployeeShift: " + e.toString());
      return null;
    }
  }
  
/**
* Checks if an employee is late based on their shift
* @param {string|number} employeeId - The employee ID
* @param {Date} clockInTime - The clock-in time
* @return {number} Minutes late (0 if not late)
*/
function checkIfLate(employeeId, clockInTime) {
  try {
    // Get employee's shift details
    const shiftDetails = getEmployeeShift(employeeId);
    if (!shiftDetails || !shiftDetails.weekAStartTime) {
      return 0; // No shift or start time, so not late
    }

    // Create a date object for the shift start time on the same day as clock-in
    const shiftStartDateTime = new Date(clockInTime.getTime());
    
    // Handle different possible formats of weekAStartTime
    if (shiftDetails.weekAStartTime instanceof Date) {
      // If it's already a Date object, just get the hours and minutes
      shiftStartDateTime.setHours(shiftDetails.weekAStartTime.getHours());
      shiftStartDateTime.setMinutes(shiftDetails.weekAStartTime.getMinutes());
      shiftStartDateTime.setSeconds(shiftDetails.weekAStartTime.getSeconds());
    } else if (typeof shiftDetails.weekAStartTime === 'string') {
      // If it's a string, parse it
      const startTimeParts = shiftDetails.weekAStartTime.split(':');
      shiftStartDateTime.setHours(parseInt(startTimeParts[0], 10));
      shiftStartDateTime.setMinutes(parseInt(startTimeParts[1], 10));
      shiftStartDateTime.setSeconds(startTimeParts[2] ? parseInt(startTimeParts[2], 10) : 0);
    } else if (typeof shiftDetails.weekAStartTime === 'number') {
      // If it's a number (like decimal hours), convert to hours and minutes
      const hours = Math.floor(shiftDetails.weekAStartTime);
      const minutes = Math.round((shiftDetails.weekAStartTime - hours) * 60);
      shiftStartDateTime.setHours(hours);
      shiftStartDateTime.setMinutes(minutes);
      shiftStartDateTime.setSeconds(0);
    } else {
      // Log the unexpected format and return 0
      Logger.log("Unexpected format for weekAStartTime: " + typeof shiftDetails.weekAStartTime);
      Logger.log("Value: " + JSON.stringify(shiftDetails.weekAStartTime));
      return 0;
    }
    
    // Calculate minutes late (if negative, employee was early)
    const minutesLate = Math.round((clockInTime - shiftStartDateTime) / (1000 * 60));
    
    // Only return positive values (if employee is late)
    return minutesLate > 0 ? minutesLate : 0;
  } catch (e) {
    Logger.log("Error in checkIfLate: " + e.toString());
    Logger.log("shiftDetails: " + JSON.stringify(shiftDetails));
    return 0;
  }
}

/**
* Checks if an employee is leaving early based on their shift
* @param {string|number} employeeId - The employee ID
* @param {Date} clockOutTime - The clock-out time
* @return {number} Minutes early (0 if not early)
*/
function checkIfEarlyDeparture(employeeId, clockOutTime) {
  try {
    // Get employee's shift details
    const shiftDetails = getEmployeeShift(employeeId);
    if (!shiftDetails || !shiftDetails.weekAEndTime) {
      return 0; // No shift or end time, so not early
    }
    
    // Create a date object for the shift end time on the same day as clock-out
    const shiftEndDateTime = new Date(clockOutTime.getTime());
    
    // Handle different possible formats of weekAEndTime
    if (shiftDetails.weekAEndTime instanceof Date) {
      // If it's already a Date object, just get the hours and minutes
      shiftEndDateTime.setHours(shiftDetails.weekAEndTime.getHours());
      shiftEndDateTime.setMinutes(shiftDetails.weekAEndTime.getMinutes());
      shiftEndDateTime.setSeconds(shiftDetails.weekAEndTime.getSeconds());
    } else if (typeof shiftDetails.weekAEndTime === 'string') {
      // If it's a string, parse it
      const endTimeParts = shiftDetails.weekAEndTime.split(':');
      shiftEndDateTime.setHours(parseInt(endTimeParts[0], 10));
      shiftEndDateTime.setMinutes(parseInt(endTimeParts[1], 10));
      shiftEndDateTime.setSeconds(endTimeParts[2] ? parseInt(endTimeParts[2], 10) : 0);
    } else if (typeof shiftDetails.weekAEndTime === 'number') {
      // If it's a number (like decimal hours), convert to hours and minutes
      const hours = Math.floor(shiftDetails.weekAEndTime);
      const minutes = Math.round((shiftDetails.weekAEndTime - hours) * 60);
      shiftEndDateTime.setHours(hours);
      shiftEndDateTime.setMinutes(minutes);
      shiftEndDateTime.setSeconds(0);
    } else {
      // Log the unexpected format and return 0
      Logger.log("Unexpected format for weekAEndTime: " + typeof shiftDetails.weekAEndTime);
      Logger.log("Value: " + JSON.stringify(shiftDetails.weekAEndTime));
      return 0;
    }
    
    // Calculate minutes early (if negative, employee left after shift end)
    const minutesEarly = Math.round((shiftEndDateTime - clockOutTime) / (1000 * 60));
    
    // Only return positive values (if employee left early)
    return minutesEarly > 0 ? minutesEarly : 0;
  } catch (e) {
    Logger.log("Error in checkIfEarlyDeparture: " + e.toString());
    Logger.log("shiftDetails: " + JSON.stringify(shiftDetails));
    return 0;
  }
}

/**
 * Gets an employee's total missed minutes for the current pay period
 * @param {string|number} employeeId - The employee ID
 * @return {Object} Object containing missed minutes details
 */
function getEmployeePayPeriodMissedMinutes(employeeId) {
  try {
    // Find the current pay period
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    
    const today = new Date();
    let currentPayPeriod = null;
    
    // Skip header row
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][6] === 'Active') { // Column G has Status
        const startDate = new Date(payPeriodsData[i][2]); // Column C has Start Date
        const endDate = new Date(payPeriodsData[i][4]);   // Column E has End Date
        
        // Check if today falls within this pay period
        if (today >= startDate && today <= endDate) {
          currentPayPeriod = {
            id: payPeriodsData[i][0],
            name: payPeriodsData[i][1],
            startDate: startDate,
            endDate: endDate
          };
          break;
        }
      }
    }
    
    if (!currentPayPeriod) {
      Logger.log("No active pay period found");
      return { total: 0, details: [] };
    }
    
    // Get time logs for this employee within the pay period
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    let totalMissedMinutes = 0;
    const missedDetails = [];
    
    // Skip header row
    for (let i = 1; i < timeLogsData.length; i++) {
      // Check if this log belongs to the employee
      if (timeLogsData[i][1] == employeeId) {
        const logDate = new Date(timeLogsData[i][2]); // Column C has Date
        
        // Check if the log date falls within the current pay period
        if (logDate >= currentPayPeriod.startDate && logDate <= currentPayPeriod.endDate) {
          // Column Z (index 25) has Total Missed Minutes
          const missedMinutes = timeLogsData[i][25] || 0;
          
          if (missedMinutes > 0) {
            totalMissedMinutes += missedMinutes;
            
            // Add details for this entry
            missedDetails.push({
              date: Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
              missedMinutes: missedMinutes,
              logId: timeLogsData[i][0]
            });
          }
        }
      }
    }
    
    return {
      total: totalMissedMinutes,
      details: missedDetails,
      payPeriod: currentPayPeriod.name
    };
    
  } catch (e) {
    Logger.log("Error in getEmployeePayPeriodMissedMinutes: " + e.toString());
    return { total: 0, details: [] };
  }
}

  
//
//
//
//
//
//
//
//
//
//
//MANAGER DASHBOARD FUNCTIONS
//
//
//
//
//
//
//
//
//
//
//
//

// Function to get all employees
function getEmployees() {
  try {
      // Initialize spreadsheet
      if (!initSpreadsheet()) {
        return [];
      }
      
      const employeeSheet = ss.getSheetByName('Employee Master Data');
      const employeeData = employeeSheet.getDataRange().getValues();
      
      // Extract header row
      const headers = employeeData[0];
      
      // Find column indexes
      const idIndex = headers.indexOf('Employee ID');
      const firstNameIndex = headers.indexOf('First Name');
      const lastNameIndex = headers.indexOf('Last Name');
      const departmentIndex = headers.indexOf('Department');
      const emailIndex = headers.indexOf('Email');
      const pinIndex = headers.indexOf('PIN');
      const managerEmailIndex = headers.indexOf('Manager Email');
      const hireDateIndex = headers.indexOf('Hire Date');
      const statusIndex = headers.indexOf('Status');
      const shiftIndex = headers.indexOf('Shift');
      const shiftIdIndex = headers.indexOf('Shift ID');
      
      // Find indexes for the new fields
      const companyIndex = headers.indexOf('Company');
      const employmentTypeIndex = headers.indexOf('Employment Type');
      const scheduleTypeIndex = headers.indexOf('Schedule type');
      const paymentTypeIndex = headers.indexOf('Payment type');
      const hourlyRateIndex = headers.indexOf('Hourly rate');
      
      // Map data to objects
      const employees = [];
      for (let i = 1; i < employeeData.length; i++) {
        const row = employeeData[i];
        
        // Skip empty rows
        if (!row[idIndex] && !row[firstNameIndex] && !row[lastNameIndex]) {
          continue;
        }
        
        // Format hire date if it exists
        let hireDate = row[hireDateIndex];
        if (hireDate instanceof Date) {
          hireDate = Utilities.formatDate(hireDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        
        employees.push({
          id: row[idIndex],
          firstName: row[firstNameIndex],
          lastName: row[lastNameIndex],
          department: row[departmentIndex],
          email: row[emailIndex],
          pin: row[pinIndex],
          managerEmail: row[managerEmailIndex],
          hireDate: hireDate,
          status: row[statusIndex] || 'Active',
          shift: row[shiftIndex],
          shiftId: row[shiftIdIndex],
          // Add the new fields
          company: companyIndex >= 0 ? row[companyIndex] : '',
          employmentType: employmentTypeIndex >= 0 ? row[employmentTypeIndex] : '',
          scheduleType: scheduleTypeIndex >= 0 ? row[scheduleTypeIndex] : '',
          paymentType: paymentTypeIndex >= 0 ? row[paymentTypeIndex] : '',
          hourlyRate: hourlyRateIndex >= 0 ? row[hourlyRateIndex] : ''
        });
      }
      
      return employees;
    } catch (error) {
      Logger.log("Error in getEmployeesForManager: " + error.toString());
      return [];
    }
  }

  
  // Function to get all shifts
  function getShifts() {
    try {
      // Initialize spreadsheet
      if (!initSpreadsheet()) {
        return [];
      }
      
      const shiftSheet = ss.getSheetByName('Shifts');
      const shiftData = shiftSheet.getDataRange().getValues();
      
      // Skip header row
      const headers = shiftData[0];
      const shifts = [];
      
      for (let i = 1; i < shiftData.length; i++) {
        const row = shiftData[i];
        
        // Skip empty rows
        if (!row[0]) continue;
        
        shifts.push({
          id: row[0] || '',
          name: row[1] || '',
          isOvernight: row[2] || false
          // Add other shift properties as needed
        });
      }
      
      return shifts;
    } catch (error) {
      Logger.log('Error getting shifts: ' + error.toString());
      throw new Error('Failed to get shifts: ' + error.toString());
    }
  }
  

// Function to save employee (add new or update existing)
function saveEmployee(employeeData) {
  try {
    // Initialize spreadsheet
    if (!initSpreadsheet()) {
      return { success: false, message: 'Failed to initialize spreadsheet' };
    }
    
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    const data = employeeSheet.getDataRange().getValues();
    const headers = data[0];
    
    // Find column indexes for the new fields
    const companyIndex = headers.indexOf('Company');
    const employmentTypeIndex = headers.indexOf('Employment Type');
    const scheduleTypeIndex = headers.indexOf('Schedule type');
    const paymentTypeIndex = headers.indexOf('Payment type');
    const hourlyRateIndex = headers.indexOf('Hourly rate');
    
    // Check if this is a new employee
    if (employeeData.id === 'NEW') {
      // Generate new employee ID (next available number)
      let maxId = 0;
      for (let i = 1; i < data.length; i++) {
        const currentId = parseInt(data[i][0]) || 0;
        if (currentId > maxId) {
          maxId = currentId;
        }
      }
      employeeData.id = maxId + 1;
      
      // Create a row with all values
      const newRow = [];
      for (let i = 0; i < headers.length; i++) {
        if (i === 0) newRow.push(employeeData.id);
        else if (i === 1) newRow.push(employeeData.firstName);
        else if (i === 2) newRow.push(employeeData.lastName);
        else if (i === 3) newRow.push(employeeData.department);
        else if (i === 4) newRow.push(employeeData.email);
        else if (i === 5) newRow.push(employeeData.pin);
        else if (i === 6) newRow.push(employeeData.managerEmail);
        else if (i === 7) newRow.push(employeeData.hireDate);
        else if (i === 8) newRow.push(employeeData.status);
        else if (i === 9) newRow.push(employeeData.shift);
        else if (i === companyIndex) newRow.push(employeeData.company);
        else if (i === employmentTypeIndex) newRow.push(employeeData.employmentType);
        else if (i === scheduleTypeIndex) newRow.push(employeeData.scheduleType);
        else if (i === paymentTypeIndex) newRow.push(employeeData.paymentType);
        else if (i === hourlyRateIndex) newRow.push(employeeData.hourlyRate);
        else newRow.push(""); // Default empty for other columns
      }
      
      // Append new row
      employeeSheet.appendRow(newRow);
      
      // Get the new row index and set the formula for shiftId if needed
      const newRowIndex = employeeSheet.getLastRow();
      const shiftIdIndex = headers.indexOf('Shift ID');
      if (shiftIdIndex >= 0) {
        const shiftIdCell = employeeSheet.getRange(newRowIndex, shiftIdIndex + 1); // +1 because sheet columns are 1-indexed
        shiftIdCell.setFormula(`=IFERROR(INDEX(Shifts!A:A, MATCH(J${newRowIndex}, Shifts!B:B, 0)), "")`);
      }
      
      return { success: true, message: 'Employee added successfully', employeeId: employeeData.id };
    } else {
      // Update existing employee
      let rowIndex = -1;
      
      // Find the row with matching employee ID
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == employeeData.id) {
          rowIndex = i + 1; // +1 because sheet rows are 1-indexed
          break;
        }
      }
      
      if (rowIndex === -1) {
        return { success: false, message: 'Employee not found' };
      }
      
      // Update each field individually
      employeeSheet.getRange(rowIndex, 2).setValue(employeeData.firstName);
      employeeSheet.getRange(rowIndex, 3).setValue(employeeData.lastName);
      employeeSheet.getRange(rowIndex, 4).setValue(employeeData.department);
      employeeSheet.getRange(rowIndex, 5).setValue(employeeData.email);
      employeeSheet.getRange(rowIndex, 6).setValue(employeeData.pin);
      employeeSheet.getRange(rowIndex, 7).setValue(employeeData.managerEmail);
      employeeSheet.getRange(rowIndex, 8).setValue(employeeData.hireDate);
      employeeSheet.getRange(rowIndex, 9).setValue(employeeData.status);
      employeeSheet.getRange(rowIndex, 10).setValue(employeeData.shift);
      
      // Update the new fields
      if (companyIndex >= 0) {
        employeeSheet.getRange(rowIndex, companyIndex + 1).setValue(employeeData.company);
      }
      if (employmentTypeIndex >= 0) {
        employeeSheet.getRange(rowIndex, employmentTypeIndex + 1).setValue(employeeData.employmentType);
      }
      if (scheduleTypeIndex >= 0) {
        employeeSheet.getRange(rowIndex, scheduleTypeIndex + 1).setValue(employeeData.scheduleType);
      }
      if (paymentTypeIndex >= 0) {
        employeeSheet.getRange(rowIndex, paymentTypeIndex + 1).setValue(employeeData.paymentType);
      }
      if (hourlyRateIndex >= 0) {
        employeeSheet.getRange(rowIndex, hourlyRateIndex + 1).setValue(employeeData.hourlyRate);
      }
      
      // Set the formula for shiftId if needed
      const shiftIdIndex = headers.indexOf('Shift ID');
      if (shiftIdIndex >= 0) {
        const shiftIdCell = employeeSheet.getRange(rowIndex, shiftIdIndex + 1);
        shiftIdCell.setFormula(`=IFERROR(INDEX(Shifts!A:A, MATCH(J${rowIndex}, Shifts!B:B, 0)), "")`);
      }
      
      return { success: true, message: 'Employee updated successfully' };
    }
  } catch (error) {
    Logger.log('Error saving employee: ' + error.toString());
    return { success: false, message: 'Failed to save employee: ' + error.toString() };
  }
}



/**
 * Checks if an employee's email is missing and saves a provided email if needed
 * @param {string} employeeId - The ID of the employee
 * @param {string} email - The email address to save (optional)
 * @return {Object} Object with status and whether email is missing
 */
function checkAndSaveEmployeeEmail(employeeId, email) {
  try {
    Logger.log("Starting checkAndSaveEmployeeEmail function for employeeId: " + employeeId);
    
    // Initialize spreadsheet
    Logger.log("Initializing spreadsheet...");
    if (!initSpreadsheet()) {
      Logger.log("ERROR: Failed to initialize spreadsheet");
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    Logger.log("Spreadsheet initialized successfully");
    
    // Get employee sheet
    Logger.log("Getting 'Employee Master Data' sheet...");
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      Logger.log("ERROR: Employee Master Data sheet not found");
      return { success: false, message: "Employee Master Data sheet not found" };
    }
    Logger.log("Employee Master Data sheet found");
    
    // Get all employee data
    Logger.log("Retrieving employee data from sheet...");
    const employeeData = employeeSheet.getDataRange().getValues();
    Logger.log("Retrieved " + employeeData.length + " rows of employee data");
    
    let employeeRow = -1;
    let employeeEmail = "";
    
    // Find the employee's row and check if email exists
    Logger.log("Searching for employee with ID: " + employeeId);
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] == employeeId) {
        employeeRow = i + 1; // +1 because sheet rows are 1-indexed
        employeeEmail = employeeData[i][4] || ""; // Email is in column E (index 4)
        Logger.log("Employee found at row " + employeeRow + " with current email: '" + employeeEmail + "'");
        break;
      }
    }
    
    if (employeeRow === -1) {
      Logger.log("ERROR: Employee with ID " + employeeId + " not found");
      return { success: false, message: "Employee not found" };
    }
    
    // If we're just checking and not saving
    if (!email) {
      const emailMissing = !employeeEmail || employeeEmail.trim() === "";
      Logger.log("Check only mode. Email missing status: " + emailMissing);
      return { 
        success: true, 
        emailMissing: emailMissing
      };
    }
    
    // If we need to save the email
    Logger.log("Attempting to save email: " + email);
    
    // Validate the email format
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(email)) {
      Logger.log("ERROR: Invalid email format provided: " + email);
      return { success: false, message: "Invalid email format" };
    }
    Logger.log("Email format validation passed");
    
    // Update the email in column E (5th column)
    Logger.log("Updating email at row " + employeeRow + ", column 5 (E)");
    employeeSheet.getRange(employeeRow, 5).setValue(email);
    Logger.log("Email successfully updated to: " + email);
    
    return { 
      success: true, 
      message: "Email saved successfully" 
    };
    
  } catch (error) {
    const errorMsg = "Error in checkAndSaveEmployeeEmail: " + error.toString();
    Logger.log("CRITICAL ERROR: " + errorMsg);
    Logger.log("Stack trace: " + error.stack);
    return { success: false, message: "Error: " + error.toString() };
  }
}


  // Function to reset employee PIN
function resetEmployeePin(pinData) {
    try {
      // Initialize spreadsheet
      if (!initSpreadsheet()) {
        return { success: false, message: 'Failed to initialize spreadsheet' };
      }
      
      const employeeSheet = ss.getSheetByName('Employee Master Data');
      const data = employeeSheet.getDataRange().getValues();
      
      let rowIndex = -1;
      
      // Find the row with matching employee ID
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] == pinData.employeeId) {
          rowIndex = i + 1; // +1 because sheet rows are 1-indexed
          break;
        }
      }
      
      if (rowIndex === -1) {
        return { success: false, message: 'Employee not found' };
      }
      
      // Update the PIN in column F (index 5)
      employeeSheet.getRange(rowIndex, 6).setValue(pinData.pin);
      
      return { success: true, message: 'PIN reset successfully', pin: pinData.pin };
    } catch (error) {
      Logger.log('Error resetting PIN: ' + error.toString());
      return { success: false, message: 'Failed to reset PIN: ' + error.toString() };
    }
  }
  
  // Function to get the next available employee ID
  function getNextEmployeeId() {
    try {
      // Initialize spreadsheet
      if (!initSpreadsheet()) {
        return { success: false, message: 'Failed to initialize spreadsheet' };
      }
      
      const employeeSheet = ss.getSheetByName('Employee Master Data');
      const data = employeeSheet.getDataRange().getValues();
      
      let maxId = 0;
      for (let i = 1; i < data.length; i++) {
        const currentId = parseInt(data[i][0]) || 0;
        if (currentId > maxId) {
          maxId = currentId;
        }
      }
      
      return { success: true, nextId: maxId + 1 };
    } catch (error) {
      Logger.log('Error getting next employee ID: ' + error.toString());
      return { success: false, message: 'Failed to get next employee ID: ' + error.toString() };
    }
  }




//
//
//
///
//
//
//
//
//
//
//
//
//TIME TRACKING FUNCTIONS
//
//
//
//
//
//
//
//
//
//



/**
* Gets time logs with optional filtering
* @param {string} dateFilter - Optional date filter (YYYY-MM-DD)
* @param {string} employeeFilter - Optional employee ID filter
* @param {boolean} missedMinutesFilter - Optional filter for logs with missed minutes
* @return {Array} Filtered time logs
*/
// Function to get all time logs with optional filters
function getTimeLogs(dateFilter, employeeFilter, missedMinutesFilter) {
  const timeLogsSheet = ss.getSheetByName('Time Logs');
  const timeLogsData = timeLogsSheet.getDataRange().getValues();
  const employeeSheet = ss.getSheetByName('Employee Master Data');
  const employeeData = employeeSheet.getDataRange().getValues();
  
  // Create employee lookup map for quick reference
  const employeeMap = {};
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][0]) { // Only add if employee ID exists
      employeeMap[employeeData[i][0]] = {
        firstName: employeeData[i][1] || '',
        lastName: employeeData[i][2] || ''
      };
    }
  }
  
  // Process time logs
  const result = [];
  for (let i = 1; i < timeLogsData.length; i++) {
    const row = timeLogsData[i];
    if (!row[0]) continue; // Skip empty rows
    
    const logId = row[0];
    const employeeId = row[1];
    
    // Store the original date object for sorting later
    const dateObj = row[3] ? new Date(row[3]) : null;
    const date = dateObj ? Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "yyyy-MM-dd") : '';
    
    // Get the total missed minutes from column Z (index 25)
    const totalMissedMinutes = Number(row[25]) || 0;
    
    // Apply filters
    if (dateFilter && date !== dateFilter) continue;
    if (employeeFilter && employeeId != employeeFilter) continue;
    if (missedMinutesFilter === true && totalMissedMinutes <= 0) continue; // Skip logs without missed minutes if filter is active
    
    // Get employee name
    const employee = employeeMap[employeeId] || { firstName: '', lastName: '' };
    const employeeName = `${employee.firstName} ${employee.lastName}`.trim();
    
    // Format times for display
    const formatTime = (timeValue) => {
      if (!timeValue) return '';
      try {
        return Utilities.formatDate(new Date(timeValue), Session.getScriptTimeZone(), "HH:mm:ss");
      } catch (e) {
        return '';
      }
    };
    
    result.push({
      rowIndex: i + 1, // 1-based row index for updating later
      logId: logId,
      employeeId: employeeId,
      employeeName: employeeName,
      date: date,
      origDate: dateObj, // Store original date object for sorting
      clockInTime: formatTime(row[3]),
      clockOutTime: formatTime(row[4]),
      regularBreakStart1: formatTime(row[5]),
      regularBreakEnd1: formatTime(row[6]),
      regularBreakStart2: formatTime(row[7]),
      regularBreakEnd2: formatTime(row[8]),
      lunchBreakStart: formatTime(row[9]),
      lunchBreakEnd: formatTime(row[10]),
      status: row[15] || '',
      totalMissedMinutes: totalMissedMinutes // Include missed minutes in the result
    });
  }
  
  // Sort results by date in descending order (newest first)
  result.sort((a, b) => {
    // First try to sort by the original Date objects if available
    if (a.origDate && b.origDate) {
      return b.origDate - a.origDate; // Descending order
    }
    // Fall back to string comparison if needed
    return b.date.localeCompare(a.date);
  });
  
  // Remove the temporary origDate field we added for sorting
  result.forEach(item => delete item.origDate);
  
  return result;
}



/**
 * Updates an existing time log entry
 * @param {number} rowIndex - The row index in the sheet
 * @param {Object} timeLogData - The updated time log data
 * @return {Object} Result of the operation
 */
function updateTimeLog(rowIndex, timeLogData) {
  try {
    Logger.log("=== START updateTimeLog ===");
    Logger.log("Row Index: " + rowIndex);
    Logger.log("Time Log Data received: " + JSON.stringify(timeLogData));
    
    // Make sure spreadsheet is initialized
    Logger.log("Initializing spreadsheet...");
    if (!initSpreadsheet()) {
      Logger.log("ERROR: Failed to initialize spreadsheet");
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    Logger.log("Spreadsheet initialized successfully");
    
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    Logger.log("Time Logs sheet accessed");
    
    // Get employee ID from the row
    const employeeId = timeLogsSheet.getRange(rowIndex, 2).getValue();
    Logger.log("Employee ID from row " + rowIndex + ": " + employeeId);
    
    // Get the original date from the row (column C)
    const originalBaseDate = timeLogsSheet.getRange(rowIndex, 3).getValue();
    Logger.log("Original base date from sheet: " + originalBaseDate + " (type: " + typeof originalBaseDate + ")");
    
    // Use the provided clock in date if available, otherwise use the original date
    const baseDate = timeLogData.clockInDate || Utilities.formatDate(new Date(originalBaseDate), Session.getScriptTimeZone(), "yyyy-MM-dd");
    Logger.log("Using base date: " + baseDate);
    
    // Format the date string
    const dateStr = baseDate;
    Logger.log("Formatted date string: " + dateStr);
    Logger.log("Script timezone: " + Session.getScriptTimeZone());
    
    // Track changes for audit trail
    const changes = [];
    
    // Get current values from the sheet for comparison
    const currentValues = {
      date: Utilities.formatDate(new Date(originalBaseDate), Session.getScriptTimeZone(), "yyyy-MM-dd"),
      clockIn: timeLogsSheet.getRange(rowIndex, 4).getValue(),
      clockOut: timeLogsSheet.getRange(rowIndex, 5).getValue(),
      regBreak1Start: timeLogsSheet.getRange(rowIndex, 6).getValue(),
      regBreak1End: timeLogsSheet.getRange(rowIndex, 7).getValue(),
      regBreak2Start: timeLogsSheet.getRange(rowIndex, 8).getValue(),
      regBreak2End: timeLogsSheet.getRange(rowIndex, 9).getValue(),
      lunchStart: timeLogsSheet.getRange(rowIndex, 10).getValue(),
      lunchEnd: timeLogsSheet.getRange(rowIndex, 11).getValue()
    };
    
    // Format current values for comparison
    const formatTimeValue = (value) => {
      if (!value) return "";
      try {
        return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), "HH:mm:ss");
      } catch (e) {
        return "";
      }
    };
    
    const currentFormattedValues = {
      date: currentValues.date,
      clockIn: formatTimeValue(currentValues.clockIn),
      clockOut: formatTimeValue(currentValues.clockOut),
      regBreak1Start: formatTimeValue(currentValues.regBreak1Start),
      regBreak1End: formatTimeValue(currentValues.regBreak1End),
      regBreak2Start: formatTimeValue(currentValues.regBreak2Start),
      regBreak2End: formatTimeValue(currentValues.regBreak2End),
      lunchStart: formatTimeValue(currentValues.lunchStart),
      lunchEnd: formatTimeValue(currentValues.lunchEnd)
    };
    
    // Enhanced function to create a properly formatted datetime from date and time strings
    const createDateTime = (dateStr, timeStr) => {
      Logger.log(`Creating DateTime for date: ${dateStr}, time: ${timeStr}`);
      if (!timeStr || timeStr.trim() === "") {
        Logger.log("Empty time string, returning empty value");
        return "";
      }
      
      // If no date is provided, use the base date
      const useDate = dateStr || baseDate;
      if (!useDate) {
        Logger.log("No date available, returning empty value");
        return "";
      }
      
      try {
        // Parse the time string (HH:MM:SS)
        Logger.log("Parsing time string: " + timeStr);
        const timeParts = timeStr.split(':');
        Logger.log("Time parts: " + JSON.stringify(timeParts));
        
        const hours = Number(timeParts[0]);
        const minutes = Number(timeParts[1]);
        const seconds = timeParts.length > 2 ? Number(timeParts[2]) : 0;
        
        Logger.log("Parsed hours: " + hours + ", minutes: " + minutes + ", seconds: " + seconds);
        
        // Create a new date object using the provided date
        Logger.log("Creating new Date object with date: " + useDate);
        
        // Parse the date components explicitly
        const [year, month, day] = useDate.split('-').map(Number);
        Logger.log("Date components - year: " + year + ", month: " + (month-1) + " (0-based), day: " + day);
        
        // Create date with explicit year, month (0-based), day values to avoid timezone issues
        const dateTime = new Date(year, month-1, day, hours || 0, minutes || 0, seconds || 0);
        Logger.log("Created date object with explicit components: " + dateTime);
        
        return dateTime;
      } catch (e) {
        Logger.log("ERROR parsing time: " + e.toString());
        return "";
      }
    };
    
    // Create date objects for each time field using the appropriate date
    Logger.log("Creating date objects for each time field...");
    
    // For clock in, use the clock in date if provided
    Logger.log("Processing clockInTime with clockInDate: " + timeLogData.clockInDate);
    const clockInDateTime = createDateTime(timeLogData.clockInDate, timeLogData.clockInTime);
    Logger.log("Resulting clockInDateTime: " + clockInDateTime);
    
    // For clock out, use the clock out date if provided, otherwise use clock in date
    const clockOutDateToUse = timeLogData.clockOutDate || timeLogData.clockInDate;
    Logger.log("Processing clockOutTime with clockOutDate: " + clockOutDateToUse);
    const clockOutDateTime = createDateTime(clockOutDateToUse, timeLogData.clockOutTime);
    Logger.log("Resulting clockOutDateTime: " + clockOutDateTime);
    
    // For breaks, use the base date (which is now the clock in date if provided)
    Logger.log("Processing regularBreakStart1: " + timeLogData.regularBreakStart1);
    const regBreak1Start = createDateTime(baseDate, timeLogData.regularBreakStart1);
    Logger.log("Resulting regBreak1Start: " + regBreak1Start);
    
    Logger.log("Processing regularBreakEnd1: " + timeLogData.regularBreakEnd1);
    const regBreak1End = createDateTime(baseDate, timeLogData.regularBreakEnd1);
    Logger.log("Resulting regBreak1End: " + regBreak1End);
    
    Logger.log("Processing regularBreakStart2: " + timeLogData.regularBreakStart2);
    const regBreak2Start = createDateTime(baseDate, timeLogData.regularBreakStart2);
    Logger.log("Resulting regBreak2Start: " + regBreak2Start);
    
    Logger.log("Processing regularBreakEnd2: " + timeLogData.regularBreakEnd2);
    const regBreak2End = createDateTime(baseDate, timeLogData.regularBreakEnd2);
    Logger.log("Resulting regBreak2End: " + regBreak2End);
    
    Logger.log("Processing lunchBreakStart: " + timeLogData.lunchBreakStart);
    const lunchStart = createDateTime(baseDate, timeLogData.lunchBreakStart);
    Logger.log("Resulting lunchStart: " + lunchStart);
    
    Logger.log("Processing lunchBreakEnd: " + timeLogData.lunchBreakEnd);
    const lunchEnd = createDateTime(baseDate, timeLogData.lunchBreakEnd);
    Logger.log("Resulting lunchEnd: " + lunchEnd);
    
    // Calculate late minutes if clock-in time is provided
    let lateMinutes = 0;
    if (clockInDateTime) {
      Logger.log("Calculating late minutes...");
      lateMinutes = checkIfLate(employeeId, clockInDateTime);
      Logger.log("Late minutes calculated: " + lateMinutes);
    }
    
    // Calculate early departure minutes if clock-out time is provided
    let earlyMinutes = 0;
    if (clockOutDateTime) {
      Logger.log("Calculating early departure minutes...");
      earlyMinutes = checkIfEarlyDeparture(employeeId, clockOutDateTime);
      Logger.log("Early departure minutes calculated: " + earlyMinutes);
    }
    
    // Check if we have a clearedFields object to track explicitly cleared fields
    const clearedFields = timeLogData.clearedFields || {};
    
    // Update the time fields
    Logger.log("Updating time fields in the sheet...");
    
    // Update the date column (column C) if clock in date has changed
    if (timeLogData.clockInDate && timeLogData.clockInDate !== currentFormattedValues.date) {
      Logger.log("Updating date in column C to: " + timeLogData.clockInDate);
      // Create a date object for the new date at midnight
      const [year, month, day] = timeLogData.clockInDate.split('-').map(Number);
      const newBaseDate = new Date(year, month-1, day);
      timeLogsSheet.getRange(rowIndex, 3).setValue(newBaseDate);
      
      // Track change
      changes.push(`Date changed from ${currentFormattedValues.date} to ${timeLogData.clockInDate}`);
    }
    
    // Clock In Time
    if (clearedFields.clockInTime === true) {
      Logger.log("Clearing clockInTime in cell (row " + rowIndex + ", col 4)");
      timeLogsSheet.getRange(rowIndex, 4).clearContent();
      changes.push("Clock-in time cleared");
    } else if (clockInDateTime && timeLogData.clockInTime !== currentFormattedValues.clockIn) {
      Logger.log("Setting clockInDateTime in cell (row " + rowIndex + ", col 4): " + clockInDateTime);
      timeLogsSheet.getRange(rowIndex, 4).setValue(clockInDateTime);
      changes.push(`Clock-in time changed from ${currentFormattedValues.clockIn} to ${timeLogData.clockInTime}`);
    }
    
    // Clock Out Time
    if (clearedFields.clockOutTime === true) {
      Logger.log("Clearing clockOutTime in cell (row " + rowIndex + ", col 5)");
      timeLogsSheet.getRange(rowIndex, 5).clearContent();
      changes.push("Clock-out time cleared");
    } else if (clockOutDateTime && timeLogData.clockOutTime !== currentFormattedValues.clockOut) {
      Logger.log("Setting clockOutDateTime in cell (row " + rowIndex + ", col 5): " + clockOutDateTime);
      timeLogsSheet.getRange(rowIndex, 5).setValue(clockOutDateTime);
      changes.push(`Clock-out time changed from ${currentFormattedValues.clockOut} to ${timeLogData.clockOutTime}`);
    }
    
    // Regular Break 1 Start
    if (clearedFields.regularBreakStart1 === true) {
      Logger.log("Clearing regularBreakStart1 in cell (row " + rowIndex + ", col 6)");
      timeLogsSheet.getRange(rowIndex, 6).clearContent();
      changes.push("Regular break 1 start time cleared");
    } else if (regBreak1Start && timeLogData.regularBreakStart1 !== currentFormattedValues.regBreak1Start) {
      Logger.log("Setting regBreak1Start in cell (row " + rowIndex + ", col 6): " + regBreak1Start);
      timeLogsSheet.getRange(rowIndex, 6).setValue(regBreak1Start);
      changes.push(`Regular break 1 start changed from ${currentFormattedValues.regBreak1Start} to ${timeLogData.regularBreakStart1}`);
    }
    
    // Regular Break 1 End
    if (clearedFields.regularBreakEnd1 === true) {
      Logger.log("Clearing regularBreakEnd1 in cell (row " + rowIndex + ", col 7)");
      timeLogsSheet.getRange(rowIndex, 7).clearContent();
      changes.push("Regular break 1 end time cleared");
    } else if (regBreak1End && timeLogData.regularBreakEnd1 !== currentFormattedValues.regBreak1End) {
      Logger.log("Setting regBreak1End in cell (row " + rowIndex + ", col 7): " + regBreak1End);
      timeLogsSheet.getRange(rowIndex, 7).setValue(regBreak1End);
      changes.push(`Regular break 1 end changed from ${currentFormattedValues.regBreak1End} to ${timeLogData.regularBreakEnd1}`);
    }
    
    // Regular Break 2 Start
    if (clearedFields.regularBreakStart2 === true) {
      Logger.log("Clearing regularBreakStart2 in cell (row " + rowIndex + ", col 8)");
      timeLogsSheet.getRange(rowIndex, 8).clearContent();
      changes.push("Regular break 2 start time cleared");
    } else if (regBreak2Start && timeLogData.regularBreakStart2 !== currentFormattedValues.regBreak2Start) {
      Logger.log("Setting regBreak2Start in cell (row " + rowIndex + ", col 8): " + regBreak2Start);
      timeLogsSheet.getRange(rowIndex, 8).setValue(regBreak2Start);
      changes.push(`Regular break 2 start changed from ${currentFormattedValues.regBreak2Start} to ${timeLogData.regularBreakStart2}`);
    }
    
    // Regular Break 2 End
    if (clearedFields.regularBreakEnd2 === true) {
      Logger.log("Clearing regularBreakEnd2 in cell (row " + rowIndex + ", col 9)");
      timeLogsSheet.getRange(rowIndex, 9).clearContent();
      changes.push("Regular break 2 end time cleared");
    } else if (regBreak2End && timeLogData.regularBreakEnd2 !== currentFormattedValues.regBreak2End) {
      Logger.log("Setting regBreak2End in cell (row " + rowIndex + ", col 9): " + regBreak2End);
      timeLogsSheet.getRange(rowIndex, 9).setValue(regBreak2End);
      changes.push(`Regular break 2 end changed from ${currentFormattedValues.regBreak2End} to ${timeLogData.regularBreakEnd2}`);
    }
    
    // Lunch Break Start
    if (clearedFields.lunchBreakStart === true) {
      Logger.log("Clearing lunchBreakStart in cell (row " + rowIndex + ", col 10)");
      timeLogsSheet.getRange(rowIndex, 10).clearContent();
      changes.push("Lunch break start time cleared");
    } else if (lunchStart && timeLogData.lunchBreakStart !== currentFormattedValues.lunchStart) {
      Logger.log("Setting lunchStart in cell (row " + rowIndex + ", col 10): " + lunchStart);
      timeLogsSheet.getRange(rowIndex, 10).setValue(lunchStart);
      changes.push(`Lunch break start changed from ${currentFormattedValues.lunchStart} to ${timeLogData.lunchBreakStart}`);
    }
    
    // Lunch Break End
    if (clearedFields.lunchBreakEnd === true) {
      Logger.log("Clearing lunchBreakEnd in cell (row " + rowIndex + ", col 11)");
      timeLogsSheet.getRange(rowIndex, 11).clearContent();
      changes.push("Lunch break end time cleared");
    } else if (lunchEnd && timeLogData.lunchBreakEnd !== currentFormattedValues.lunchEnd) {
      Logger.log("Setting lunchEnd in cell (row " + rowIndex + ", col 11): " + lunchEnd);
      timeLogsSheet.getRange(rowIndex, 11).setValue(lunchEnd);
      changes.push(`Lunch break end changed from ${currentFormattedValues.lunchEnd} to ${timeLogData.lunchBreakEnd}`);
    }
    
    // Update late and early departure minutes
    Logger.log("Updating late and early departure minutes...");
    Logger.log("Setting late minutes in cell (row " + rowIndex + ", col 24): " + (lateMinutes > 0 ? lateMinutes : ""));
    timeLogsSheet.getRange(rowIndex, 24).setValue(lateMinutes > 0 ? lateMinutes : "");
    
    Logger.log("Setting early minutes in cell (row " + rowIndex + ", col 25): " + (earlyMinutes > 0 ? earlyMinutes : ""));
    timeLogsSheet.getRange(rowIndex, 25).setValue(earlyMinutes > 0 ? earlyMinutes : "");
    
    // Update status based on clock-in and clock-out times
    const status = (clockInDateTime && clockOutDateTime) ? "Complete" : "Incomplete";
    Logger.log("Setting status in cell (row " + rowIndex + ", col 16): " + status);
    timeLogsSheet.getRange(rowIndex, 16).setValue(status);
    
    // Update notes if needed to reflect late arrival or early departure
    Logger.log("Updating notes field...");
    const currentNotes = timeLogsSheet.getRange(rowIndex, 17).getValue() || "";
    Logger.log("Current notes: " + currentNotes);
    
    let newNotes = currentNotes;
    
    // Remove existing late/early notes
    newNotes = newNotes.replace(/Late clock-in,?\s*/g, "");
    newNotes = newNotes.replace(/Early departure,?\s*/g, "");
    newNotes = newNotes.replace(/,\s*$/, ""); // Remove trailing comma if any
    Logger.log("Notes after removing existing late/early tags: " + newNotes);
    
    // Add new notes if needed
    if (lateMinutes > 0) {
      newNotes = newNotes ? newNotes + ", Late clock-in" : "Late clock-in";
      Logger.log("Added late clock-in note. Notes now: " + newNotes);
    }
    
    if (earlyMinutes > 0) {
      newNotes = newNotes ? newNotes + ", Early departure" : "Early departure";
      Logger.log("Added early departure note. Notes now: " + newNotes);
    }
    
    // Add audit trail for changes made from manager dashboard
    if (changes.length > 0) {
      const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      const auditNote = `[${timestamp}] Manager edit: ${changes.join("; ")}`;
      
      // Append to existing notes or create new note
      newNotes = newNotes ? newNotes + "\n" + auditNote : auditNote;
      Logger.log("Added audit trail to notes: " + auditNote);
    }
    
    // Update notes field
    if (newNotes !== currentNotes) {
      Logger.log("Setting updated notes in cell (row " + rowIndex + ", col 17): " + newNotes);
      timeLogsSheet.getRange(rowIndex, 17).setValue(newNotes);
    } else {
      Logger.log("Notes unchanged, not updating");
    }
    
    // Recalculate the Total Missed Minutes formula in column Z (26)
    const totalMissedFormula = `=SUM(IF(ISBLANK(U${rowIndex}),0,U${rowIndex}),IF(ISBLANK(V${rowIndex}),0,V${rowIndex}),IF(ISBLANK(W${rowIndex}),0,W${rowIndex}),IF(ISBLANK(X${rowIndex}),0,X${rowIndex}),IF(ISBLANK(Y${rowIndex}),0,Y${rowIndex}))`;
    Logger.log("Setting total missed minutes formula in cell (row " + rowIndex + ", col 26): " + totalMissedFormula);
    timeLogsSheet.getRange(rowIndex, 26).setFormula(totalMissedFormula);
    
    Logger.log("Time log update completed successfully");
    Logger.log("=== END updateTimeLog ===");
    
    return { 
      success: true, 
      message: "Time log updated successfully",
      changes: changes // Return the changes made for potential client-side display
    };
  } catch (error) {
    Logger.log("ERROR in updateTimeLog: " + error.toString());
    Logger.log("Error stack: " + error.stack);
    return { success: false, message: error.toString() };
  }
}











function getReportData(payPeriodId, employeeId, startDate, endDate) {
  const timeLogsSheet = ss.getSheetByName('Time Logs');
  const payPeriodsSheet = ss.getSheetByName('Pay Periods');
  const employeeSheet = ss.getSheetByName('Employee Master Data');
  
  // Get all data
  const timeLogsData = timeLogsSheet.getDataRange().getValues();
  const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
  const employeeData = employeeSheet.getDataRange().getValues();
  
  // Create employee lookup map
  const employeeMap = {};
  employeeData.slice(1).forEach(row => {
    employeeMap[row[0]] = {
      firstName: row[1],
      lastName: row[2],
      department: row[3]
    };
  });
  
  // Filter time logs based on parameters
  let filteredLogs = timeLogsData.slice(1).filter(row => {
    const logDate = new Date(row[2]);
    
    // Apply filters
    if (startDate && endDate) {
      const start = new Date(startDate);
      const end = new Date(endDate);
      if (logDate < start || logDate > end) return false;
    }
    
    if (employeeId && row[1] != employeeId) return false;
    
    return true;
  });
  
  // Aggregate data by employee and pay period
  const reportData = {};
  
  filteredLogs.forEach(log => {
    const employeeId = log[1];
    const logDate = new Date(log[2]);
    
    // Find matching pay period
    const payPeriod = payPeriodsData.slice(1).find(pp => {
      const ppStart = new Date(pp[2]);
      const ppEnd = new Date(pp[4]);
      return logDate >= ppStart && logDate <= ppEnd;
    });
    
    if (!payPeriod || (payPeriodId && payPeriod[0] !== payPeriodId)) return;
    
    const key = `${employeeId}-${payPeriod[0]}`;
    if (!reportData[key]) {
      reportData[key] = {
        employeeId: employeeId,
        employeeName: `${employeeMap[employeeId]?.firstName || ''} ${employeeMap[employeeId]?.lastName || ''}`.trim(),
        payPeriodId: payPeriod[0],
        payPeriodName: payPeriod[1],
        totalHours: 0,
        regularHours: 0,
        breakTime: 0,
        lateMinutes: 0,
        earlyDeparture: 0,
        breakViolations: 0,
        totalViolations: 0
      };
    }
    
    // Aggregate metrics
    reportData[key].totalHours += Number(log[11]) || 0;
    reportData[key].regularHours += Number(log[14]) || 0;
    reportData[key].breakTime += (Number(log[12]) + Number(log[13])) || 0;
    reportData[key].lateMinutes += Number(log[23]) || 0;
    reportData[key].earlyDeparture += Number(log[24]) || 0;
    reportData[key].breakViolations += 
      (Number(log[20]) || 0) + 
      (Number(log[21]) || 0) + 
      (Number(log[22]) || 0);
    reportData[key].totalViolations += Number(log[25]) || 0;
  });
  
  return Object.values(reportData);
}

















function getReportData(date, payPeriod) {
  const timeLogsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Time Logs');
  const payPeriodsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pay Periods');
  
  const timeLogsData = timeLogsSheet.getDataRange().getValues();
  const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
  
  const report = [];
  const payPeriodDates = payPeriodsData.map(row => row[1]); 

  for (let i = 1; i < timeLogsData.length; i++) {
      const log = timeLogsData[i];
      const empId = log[1];
      const hoursWorked = log[11]; 
      const missedMinutes = log[25]; 

      if (payPeriod === 'all' || payPeriodDates.includes(payPeriod)) {
          report.push([empId, getEmployeeName(empId), hoursWorked, missedMinutes]);
      }
  }
  return report;
}

function getEmployeeName(empId) {
  const empSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Employee Master Data');
  const empData = empSheet.getDataRange().getValues();
  for (let i = 1; i < empData.length; i++) {
      if (empData[i][0] == empId) {
          return `${empData[i][1]} ${empData[i][2]}`; 
      }
  }
  return 'Unknown';
}


/**
 * Gets time logs data for reports based on filters
 * @param {Object} filters - Filter criteria (date, payPeriod)
 * @return {Array} Processed time logs data for reports
 */
function getTimeLogsReport(filters) {
  try {
    // Initialize spreadsheet
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    
    if (!timeLogsSheet || !employeeSheet) {
      return { success: false, message: "Required sheets not found" };
    }
    
    // Get all time logs data
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Get all employee data for lookup
    const employeeData = employeeSheet.getDataRange().getValues();
    const employeeMap = {};
    
    // Map employees by ID
    for (let i = 1; i < employeeData.length; i++) {
      const employeeId = employeeData[i][0];
      employeeMap[employeeId] = {
        firstName: employeeData[i][1],
        lastName: employeeData[i][2],
        department: employeeData[i][3]
      };
    }
    
    // Process filters
    let startDate = null;
    let endDate = null;
    
    if (filters.payPeriod && filters.payPeriod !== 'all') {
      // Get pay period dates
      const payPeriodsSheet = ss.getSheetByName('Pay Periods');
      const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
      
      for (let i = 1; i < payPeriodsData.length; i++) {
        if (payPeriodsData[i][0] === filters.payPeriod) {
          startDate = new Date(payPeriodsData[i][2]); // Start Date
          endDate = new Date(payPeriodsData[i][4]);   // End Date
          break;
        }
      }
    } else if (filters.date) {
      // Single date filter
      startDate = new Date(filters.date);
      startDate.setHours(0, 0, 0, 0);
      
      endDate = new Date(filters.date);
      endDate.setHours(23, 59, 59, 999);
    }
    
    // Aggregate data by employee
    const employeeStats = {};
    
    // Skip header row
    for (let i = 1; i < timeLogsData.length; i++) {
      const logDate = new Date(timeLogsData[i][2]); // Date column
      const employeeId = timeLogsData[i][1];        // Employee ID column
      
      // Apply date filter if specified
      if (startDate && endDate) {
        if (logDate < startDate || logDate > endDate) {
          continue; // Skip logs outside the date range
        }
      }
      
      // Skip logs without employee ID or incomplete logs
      if (!employeeId || timeLogsData[i][15] !== "Complete") {
        continue;
      }
      
      // Initialize employee stats if not already done
      if (!employeeStats[employeeId]) {
        const name = employeeMap[employeeId] ? 
          `${employeeMap[employeeId].firstName} ${employeeMap[employeeId].lastName}` : 
          `Unknown (ID: ${employeeId})`;
        
        employeeStats[employeeId] = {
          employeeId: employeeId,
          name: name,
          totalHoursWorked: 0,
          totalRegularBreakHours: 0,
          totalLunchHours: 0,
          totalMissedMinutes: 0
        };
      }
      
      // Add hours worked (Net Hours column)
      const netHours = timeLogsData[i][14] || 0;
      employeeStats[employeeId].totalHoursWorked += Number(netHours);
      
      // Add regular break hours (from minutes columns)
      const regularBreak1Minutes = timeLogsData[i][17] || 0;
      const regularBreak2Minutes = timeLogsData[i][18] || 0;
      employeeStats[employeeId].totalRegularBreakHours += 
        (Number(regularBreak1Minutes) + Number(regularBreak2Minutes)) / 60;
      
      // Add lunch break hours (from minutes column)
      const lunchBreakMinutes = timeLogsData[i][19] || 0;
      employeeStats[employeeId].totalLunchHours += Number(lunchBreakMinutes) / 60;
      
      // Add missed minutes
      const missedMinutes = timeLogsData[i][25] || 0;
      employeeStats[employeeId].totalMissedMinutes += Number(missedMinutes);
    }
    
    // Convert to array for easier client-side handling
    const reportData = Object.values(employeeStats);
    
    // Sort by name
    reportData.sort((a, b) => a.name.localeCompare(b.name));
    
    return { 
      success: true, 
      data: reportData 
    };
    
  } catch (error) {
    Logger.log("Error in getTimeLogsReport: " + error.toString());
    return { 
      success: false, 
      message: "Error generating report: " + error.toString() 
    };
  }
}
// Function to get pay periods for the dropdown
function getPayPeriods() {
  try {
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    
    const payPeriods = [];
    
    // Skip header row
    for (let i = 1; i < payPeriodsData.length; i++) {
      const row = payPeriodsData[i];
      
      // Check if we have valid data (Pay Period ID)
      if (!row[0]) continue;
      
      // Format dates and times for consistency
      let startDate, startTime, endDate, endTime;
      
      try {
        startDate = Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        startTime = Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "HH:mm:ss");
        endDate = Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "yyyy-MM-dd");
        endTime = Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "HH:mm:ss");
      } catch (e) {
        Logger.log("Error formatting dates for pay period " + row[0] + ": " + e.toString());
        continue;
      }
      
      payPeriods.push({
        id: row[0],
        name: row[1],
        startDate: startDate,
        startTime: startTime,
        endDate: endDate,
        endTime: endTime,
        status: row[6]
      });
    }
    
    return payPeriods;
  } catch (error) {
    Logger.log('Error in getPayPeriods: ' + error.toString());
    throw new Error('Failed to load pay periods: ' + error.toString());
  }
}

// Function to get pay period dates by ID
function getPayPeriodDates(payPeriodId) {
  const payPeriodsSheet = ss.getSheetByName('Pay Periods');
  const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
  
  // Find the pay period with the matching ID
  for (let i = 1; i < payPeriodsData.length; i++) {
    if (payPeriodsData[i][0] === payPeriodId) {
      return {
        startDate: Utilities.formatDate(new Date(payPeriodsData[i][2]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        endDate: Utilities.formatDate(new Date(payPeriodsData[i][4]), Session.getScriptTimeZone(), "yyyy-MM-dd")
      };
    }
  }
  
  return null;
}

// Function to get all departments
function getDepartments() {
  const employeeSheet = ss.getSheetByName('Employee Master Data');
  const employeeData = employeeSheet.getDataRange().getValues();
  
  // Skip header row and collect unique departments
  const departments = new Set();
  for (let i = 1; i < employeeData.length; i++) {
    if (employeeData[i][3]) { // Department is in column D (index 3)
      departments.add(employeeData[i][3]);
    }
  }
  
  return Array.from(departments).sort();
}













// Function to get pay periods for the dropdown
function getPayPeriods() {
  try {
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    
    const payPeriods = [];
    
    // Skip header row
    for (let i = 1; i < payPeriodsData.length; i++) {
      const row = payPeriodsData[i];
      
      // Check if we have valid data (Pay Period ID)
      if (!row[0]) continue;
      
      // Format dates and times for consistency
      const startDate = Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const startTime = Utilities.formatDate(new Date(row[3]), Session.getScriptTimeZone(), "HH:mm:ss");
      const endDate = Utilities.formatDate(new Date(row[4]), Session.getScriptTimeZone(), "yyyy-MM-dd");
      const endTime = Utilities.formatDate(new Date(row[5]), Session.getScriptTimeZone(), "HH:mm:ss");
      
      payPeriods.push({
        id: row[0],
        name: row[1],
        startDate: startDate,
        startTime: startTime,
        endDate: endDate,
        endTime: endTime,
        status: row[6]
      });
    }
    
    return payPeriods;
  } catch (error) {
    Logger.log('Error in getPayPeriods: ' + error.toString());
    throw new Error('Failed to load pay periods: ' + error.toString());
  }
}

// Function to generate employee time report based on date range
function generateEmployeeTimeReport(startDate, startTime, endDate, endTime) {
  try {
    Logger.log("Report parameters: " + startDate + " " + startTime + " to " + endDate + " " + endTime);
    
    // Format start and end datetime
    const startDateTime = new Date(`${startDate}T${startTime}`);
    const endDateTime = new Date(`${endDate}T${endTime}`);
    
    Logger.log("Start datetime: " + startDateTime);
    Logger.log("End datetime: " + endDateTime);
    Logger.log("Date range: " + startDateTime.toISOString() + " to " + endDateTime.toISOString());
    
    // Get time logs data
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Get employee data for names and additional information
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    const employeeData = employeeSheet.getDataRange().getValues();
    
    // Get headers to find column indices
    const employeeHeaders = employeeData[0];
    const employmentTypeIdx = employeeHeaders.indexOf('Employment Type');
    const hourlyRateIdx = employeeHeaders.indexOf('Hourly rate');
    const eligibleForPSIdx = employeeHeaders.indexOf('Eligible for Profit Sharing');
    
    // Create a map of employee IDs to employee details
    const employeeMap = {};
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0]) { // Check if employee ID exists
        const employeeId = employeeData[i][0].toString();
        const firstName = employeeData[i][1] || '';
        const lastName = employeeData[i][2] || '';
        
        employeeMap[employeeId] = {
          name: `${firstName} ${lastName}`.trim(),
          employmentType: employmentTypeIdx >= 0 ? employeeData[i][employmentTypeIdx] || '' : '',
          hourlyRate: hourlyRateIdx >= 0 ? employeeData[i][hourlyRateIdx] || '' : '',
          eligibleForProfitSharing: eligibleForPSIdx >= 0 ? employeeData[i][eligibleForPSIdx] || 'No' : 'No'
        };
      }
    }
    
    // Aggregate data by employee
    const employeeTotals = {};
    
    // Track employees with AUTO- logs
    const employeesWithAutoLogs = new Set();
    
    // Skip header row
    for (let i = 1; i < timeLogsData.length; i++) {
      const row = timeLogsData[i];
      // Check if we have valid data
      if (!row[0] || !row[1]) continue;
      
      const employeeId = row[1].toString();
      const logId = row[0].toString();
      
      try {
        // Get the log date (column C, index 2)
        let logDate;
        if (row[2]) {
          if (row[2] instanceof Date) {
            logDate = row[2];
          } else {
            try {
              logDate = new Date(row[2]);
            } catch (e) {
              Logger.log("Invalid log date format at row " + (i+1) + ": " + row[2]);
              continue;
            }
          }
        } else {
          continue; // Skip if no date
        }
        
        // Get clock in time if available
        let logDateTime = logDate;
        if (row[3] && row[3] instanceof Date) {
          // If we have a clock-in time, use it to create a full datetime
          const hours = row[3].getHours();
          const minutes = row[3].getMinutes();
          const seconds = row[3].getSeconds();
          
          logDateTime = new Date(logDate);
          logDateTime.setHours(hours, minutes, seconds);
        }
        
        // Check if the log date falls within the date range (using actual datetime)
        if (logDateTime >= startDateTime && logDateTime <= endDateTime) {
          // Check if this is an AUTO- log and add to tracking set if it is
          if (logId.includes("AUTO-")) {
            employeesWithAutoLogs.add(employeeId);
            Logger.log(`Found AUTO- log for employee ${employeeId}: ${logId}`);
          }
          
          // Initialize employee data if not exists
          if (!employeeTotals[employeeId]) {
            employeeTotals[employeeId] = {
              employeeId: employeeId,
              employeeName: employeeMap[employeeId] ? employeeMap[employeeId].name : `Unknown (ID: ${employeeId})`,
              employmentType: employeeMap[employeeId] ? employeeMap[employeeId].employmentType : '',
              hourlyRate: employeeMap[employeeId] ? employeeMap[employeeId].hourlyRate : '',
              eligibleForProfitSharing: employeeMap[employeeId] ? employeeMap[employeeId].eligibleForProfitSharing : 'No',
              totalHoursWorked: 0,
              regularBreakTime: 0,
              lunchBreakTime: 0,
              totalMissedMinutes: 0,
              recordCount: 0,
              dailyLogs: [],
              dailyHours: [],
              hasAutoLogs: false // Initialize flag for AUTO- logs
            };
          }
          
          // Add data from this row
          const totalHoursWorked = parseFloat(row[14] || 0); // Column O: Total Net Hours Worked
          const regularBreakTime = parseFloat(row[12] || 0); // Column M: Total Regular Break Time
          const lunchBreakTime = parseFloat(row[13] || 0); // Column N: Total Lunch Break Time
          const totalMissedMinutes = parseFloat(row[25] || 0); // Column Z: Total Missed Minutes
          
          // NEW: Get clock in/out times and status
          const clockInTime = row[3]; // Column D: Clock In Time 
          const clockOutTime = row[4]; // Column E: Clock Out Time
          const status = row[15] || ""; // Column P: Status (Complete/Incomplete)
          
          // Add cumulative data
          employeeTotals[employeeId].totalHoursWorked += totalHoursWorked;
          employeeTotals[employeeId].regularBreakTime += regularBreakTime;
          employeeTotals[employeeId].lunchBreakTime += lunchBreakTime;
          employeeTotals[employeeId].totalMissedMinutes += totalMissedMinutes;
          employeeTotals[employeeId].recordCount++;
          
          // NEW: Add detailed shift information
          const formattedDate = Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          
          // Add daily log with complete details
          employeeTotals[employeeId].dailyLogs.push({
            logId: logId,
            date: formattedDate,
            clockIn: clockInTime ? true : false,
            clockOut: clockOutTime ? true : false,
            hours: totalHoursWorked,
            status: status
          });
          
          // Add hours worked for this shift to the daily hours array
          employeeTotals[employeeId].dailyHours.push(totalHoursWorked);
          
          Logger.log(`Added data for employee ${employeeId} on ${logDate.toLocaleDateString()}: Hours=${totalHoursWorked}, RegBreak=${regularBreakTime}, LunchBreak=${lunchBreakTime}, MissedMin=${totalMissedMinutes}, Status=${status}`);
        }
      } catch (rowError) {
        Logger.log("Error processing row " + (i+1) + ": " + rowError.toString());
        continue;
      }
    }
    
    // Update hasAutoLogs flag for all employees with AUTO- logs
    employeesWithAutoLogs.forEach(employeeId => {
      if (employeeTotals[employeeId]) {
        employeeTotals[employeeId].hasAutoLogs = true;
        Logger.log(`Marked employee ${employeeId} as having AUTO- logs`);
      }
    });
    
    // Convert to array for return
    const reportData = Object.values(employeeTotals).map(employee => {
      // If employee has AUTO- logs, set paid break hours to 0
      if (employee.hasAutoLogs) {
        Logger.log(`Setting paid break hours to 0 for employee ${employee.employeeId} due to AUTO- logs`);
        employee.paidBreakHours = 0;
      } else {
        // Calculate paid break hours based on qualifying shifts
        const qualifyingShifts = Math.floor(employee.totalHoursWorked / 7);
        employee.paidBreakHours = qualifyingShifts * 0.5; // 30 minutes (0.5 hours) per qualifying shift
      }
      return employee;
    });
    
    // Sort by employee name
    reportData.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    
    Logger.log("Report generated with " + reportData.length + " employee records");
    return reportData;
  } catch (error) {
    Logger.log('Error in generateEmployeeTimeReport: ' + error.toString());
    throw new Error('Failed to generate report: ' + error.toString());
  }
}




// Generate PDF report
function generateReportPdf(reportData, startDate, endDate) {
  try {
    // Create a temporary HTML file for the PDF
    const htmlOutput = HtmlService.createTemplateFromFile('ReportPdfTemplate');
    
    // Pass data to the template
    htmlOutput.reportData = reportData;
    htmlOutput.startDate = startDate;
    htmlOutput.endDate = endDate;
    htmlOutput.generatedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
    
    // Evaluate the template to HTML
    const html = htmlOutput.evaluate().getContent();
    
    // Create a blob from the HTML
    const blob = Utilities.newBlob(html, 'text/html', 'report.html');
    
    // Create PDF from HTML
    const pdf = blob.getAs('application/pdf');
    
    // Set filename
    const filename = `Employee_Time_Report_${startDate.replace(/-/g, '')}_to_${endDate.replace(/-/g, '')}.pdf`;
    pdf.setName(filename);
    
    // Save to Drive temporarily
    const folder = DriveApp.getRootFolder();
    const file = folder.createFile(pdf);
    
    // Get the URL
    const url = file.getUrl();
    
    // Set expiration date to 5 minutes from now
    const expirationDate = new Date();
    expirationDate.setMinutes(expirationDate.getMinutes() + 5);
    
    // Make the file accessible via URL
    Drive.Files.update({
      'shared': true,
      'publishAuto': true,
      'publishedOutsideDomain': true
    }, file.getId());
    
    // Return the URL
    return url;
  } catch (error) {
    Logger.log('Error generating PDF: ' + error.toString());
    return null;
  }
}


// Get employee time report for the current pay period
function getEmployeeTimeReport(employeeId) {
  try {
    // Log for debugging
    Logger.log("Getting time report for employee ID: " + employeeId);
    
    // Make sure employeeId is a string for consistent comparison
    employeeId = String(employeeId).replace('.0', '');
    
    // Get the current pay period
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      return { success: false, message: "Pay Periods sheet not found" };
    }

    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    let currentPayPeriod = null;

    // First look for an active pay period
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][6] === "Active") {
        Logger.log("Found active pay period row: " + i);
        
        // Get start date from column C (index 2)
        let startDateValue = payPeriodsData[i][2];
        // Get start time from column D (index 3)
        let startTimeValue = payPeriodsData[i][3];
        
        // Get end date from column E (index 4)
        let endDateValue = payPeriodsData[i][4];
        // Get end time from column F (index 5)
        let endTimeValue = payPeriodsData[i][5];
        
        Logger.log("Raw start date: " + startDateValue);
        Logger.log("Raw start time: " + startTimeValue);
        Logger.log("Raw end date: " + endDateValue);
        Logger.log("Raw end time: " + endTimeValue);
        
        // Create start date with time
        let startDate = new Date(startDateValue);
        if (startTimeValue) {
          if (startTimeValue instanceof Date) {
            // If it's already a Date object, extract just the time components
            startDate.setHours(
              startTimeValue.getHours(),
              startTimeValue.getMinutes(),
              startTimeValue.getSeconds(),
              0
            );
          } else if (typeof startTimeValue === 'string') {
            // Parse time string (HH:MM:SS)
            const timeParts = startTimeValue.split(':');
            startDate.setHours(
              parseInt(timeParts[0] || 0),
              parseInt(timeParts[1] || 0),
              parseInt(timeParts[2] || 0),
              0
            );
          } else if (typeof startTimeValue === 'number') {
            // Handle Excel time format (decimal fraction of 24 hours)
            const hours = Math.floor(startTimeValue * 24);
            const minutes = Math.floor((startTimeValue * 24 * 60) % 60);
            const seconds = Math.floor((startTimeValue * 24 * 60 * 60) % 60);
            startDate.setHours(hours, minutes, seconds, 0);
          }
        }
        
        // Create end date with time
        let endDate = new Date(endDateValue);
        if (endTimeValue) {
          if (endTimeValue instanceof Date) {
            // If it's already a Date object, extract just the time components
            endDate.setHours(
              endTimeValue.getHours(),
              endTimeValue.getMinutes(),
              endTimeValue.getSeconds(),
              0
            );
          } else if (typeof endTimeValue === 'string') {
            // Parse time string (HH:MM:SS)
            const timeParts = endTimeValue.split(':');
            endDate.setHours(
              parseInt(timeParts[0] || 0),
              parseInt(timeParts[1] || 0),
              parseInt(timeParts[2] || 0),
              0
            );
          } else if (typeof endTimeValue === 'number') {
            // Handle Excel time format
            const hours = Math.floor(endTimeValue * 24);
            const minutes = Math.floor((endTimeValue * 24 * 60) % 60);
            const seconds = Math.floor((endTimeValue * 24 * 60 * 60) % 60);
            endDate.setHours(hours, minutes, seconds, 0);
          }
        }
        
        // IMPORTANT CHANGE: Store as Date objects, not strings
        currentPayPeriod = {
          id: payPeriodsData[i][0],
          name: payPeriodsData[i][1],
          startDate: startDate,
          endDate: endDate,
          // Add formatted versions for display purposes
          startDateFormatted: Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"),
          endDateFormatted: Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss")
        };
        
        Logger.log("Found active pay period: " + currentPayPeriod.id);
        Logger.log("Start date (parsed): " + currentPayPeriod.startDate.toLocaleString());
        Logger.log("End date (parsed): " + currentPayPeriod.endDate.toLocaleString());
        break;
      }
    }
    
    if (!currentPayPeriod) {
      return { success: false, message: "No active pay period found" };
    }
    
    // Get time logs for this employee within the pay period
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    if (!timeLogsSheet) {
      // Try alternate sheet name
      const altTimeLogsSheet = ss.getSheetByName('Time Logs Test');
      if (!altTimeLogsSheet) {
        return { success: false, message: "Time Logs sheet not found" };
      }
      timeLogsSheet = altTimeLogsSheet;
    }
    
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    const allTimeLogs = [];
    const incompleteLogs = [];
    let totalHoursWorked = 0;
    let totalRegularBreakTime = 0;
    let totalLunchBreakTime = 0;
    let totalMissedMinutes = 0;
    
    // For debugging - log the date range we're looking for
    Logger.log("Looking for logs between: " + currentPayPeriod.startDateFormatted + " and " + currentPayPeriod.endDateFormatted);
    
    // First pass: collect all logs and identify incomplete ones
    // Skip header row
    for (let i = 1; i < timeLogsData.length; i++) {
      // Skip empty rows or rows without employee ID
      if (!timeLogsData[i][1]) continue;
      
      // Convert employee ID to string for comparison and remove decimal if present
      const logEmployeeId = String(timeLogsData[i][1]).replace('.0', '');
      
      // Check if this log is for the requested employee
      if (logEmployeeId !== employeeId) continue;
      
      // Get the log date from column D(clockin time) (index 3)
      const logDate = new Date(timeLogsData[i][3]);
      
      // Check if the log date falls within the pay period
      if (logDate >= currentPayPeriod.startDate && logDate <= currentPayPeriod.endDate) {
        Logger.log(`Found time log for employee ${employeeId} on ${logDate.toLocaleDateString()}`);
        
        // Extract all the necessary data from the time log
        const timeLog = {
          id: timeLogsData[i][0],
          date: Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          clockIn: formatTimeForDisplay(timeLogsData[i][3]),
          clockOut: formatTimeForDisplay(timeLogsData[i][4]),
          regularBreakStart1: formatTimeForDisplay(timeLogsData[i][5]),
          regularBreakEnd1: formatTimeForDisplay(timeLogsData[i][6]),
          regularBreakStart2: formatTimeForDisplay(timeLogsData[i][7]),
          regularBreakEnd2: formatTimeForDisplay(timeLogsData[i][8]),
          lunchBreakStart: formatTimeForDisplay(timeLogsData[i][9]),
          lunchBreakEnd: formatTimeForDisplay(timeLogsData[i][10]),
          hoursWorked: Number(timeLogsData[i][11]) || 0,
          regularBreakTime: Number(timeLogsData[i][12]) || 0,
          lunchBreakTime: Number(timeLogsData[i][13]) || 0,
          netWorkingHours: Number(timeLogsData[i][14]) || 0,
          status: timeLogsData[i][15] || "",
          missedMinutes: Number(timeLogsData[i][25]) || 0,
          rowIndex: i, // Store the row index for reference
          logDate: logDate // Store the actual date object for sorting
        };
        
        allTimeLogs.push(timeLog);
        
        // Check if this log is incomplete
        if (timeLog.status === "Incomplete") {
          incompleteLogs.push(timeLog);
        }
        
        // Add to totals
        totalHoursWorked += timeLog.hoursWorked;
        totalRegularBreakTime += timeLog.regularBreakTime;
        totalLunchBreakTime += timeLog.lunchBreakTime;
        totalMissedMinutes += timeLog.missedMinutes;
      }
    }
    
    // Handle multiple incomplete logs if they exist
    let warningMessage = null;
    if (incompleteLogs.length > 1) {
      // Sort incomplete logs by date (most recent first)
      incompleteLogs.sort((a, b) => b.logDate - a.logDate);
      
      warningMessage = `Found ${incompleteLogs.length} incomplete time logs. Prioritizing the most recent one from ${incompleteLogs[0].date}.`;
      Logger.log(warningMessage);
      
      // Optionally, you could mark the older incomplete logs with a warning
      // This would require additional processing and updating the sheet
    }
    
    // Sort all time logs by date (most recent first)
    allTimeLogs.sort((a, b) => {
      // First prioritize incomplete logs
      if (a.status === "Incomplete" && b.status !== "Incomplete") return -1;
      if (a.status !== "Incomplete" && b.status === "Incomplete") return 1;
      
      // If both are incomplete or both are complete, sort by date (newest first)
      return b.logDate - a.logDate;
    });
    
    // Clean up the logs before returning (remove internal properties)
    const timeLogs = allTimeLogs.map(log => {
      const { rowIndex, logDate, ...cleanLog } = log;
      return cleanLog;
    });
    
    // Return the time logs and totals
    return {
      success: true,
      employeeId: employeeId,
      payPeriod: {
        id: currentPayPeriod.id,
        name: currentPayPeriod.name,
        startDate: currentPayPeriod.startDateFormatted,
        endDate: currentPayPeriod.endDateFormatted
      },
      timeLogs: timeLogs,
      totals: {
        hoursWorked: totalHoursWorked,
        regularBreakTime: totalRegularBreakTime,
        lunchBreakTime: totalLunchBreakTime,
        netWorkingHours: totalHoursWorked - totalRegularBreakTime - totalLunchBreakTime,
        missedMinutes: totalMissedMinutes
      },
      warning: warningMessage // Include warning about multiple incomplete logs if applicable
    };
    
  } catch (e) {
    Logger.log("Error in getEmployeeTimeReport: " + e.toString());
    return { success: false, message: "Error: " + e.toString() };
  }
}



function getEmployeeTimeReportCustom(employeeId, startDate, endDate, payPeriodId = null) {
  try {
    Logger.log("Getting custom time report for employee ID: " + employeeId);
    Logger.log("Start date input: " + startDate);
    Logger.log("End date input: " + endDate);
    Logger.log("Pay Period ID: " + payPeriodId);
    
    // Make sure employeeId is a string for consistent comparison
    employeeId = String(employeeId).replace('.0', '');
    
    let start, end;
    let payPeriod = null;
    
    // If payPeriodId is provided, get dates from the Pay Periods sheet
    if (payPeriodId) {
      const payPeriodsSheet = ss.getSheetByName('Pay Periods');
      if (!payPeriodsSheet) {
        return { success: false, message: "Pay Periods sheet not found" };
      }
      
      const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
      
      // Debug log to see what we're working with
      Logger.log("Pay Periods Data headers: " + JSON.stringify(payPeriodsData[0]));
      
      // Find the specified pay period
      for (let i = 1; i < payPeriodsData.length; i++) {
        // Convert both to strings for comparison and remove any decimal points
        const rowPayPeriodId = String(payPeriodsData[i][0]).replace('.0', '');
        const searchPayPeriodId = String(payPeriodId).replace('.0', '');
        
        Logger.log(`Comparing row ID: ${rowPayPeriodId} with search ID: ${searchPayPeriodId}`);
        
        if (rowPayPeriodId === searchPayPeriodId) {
          Logger.log("Found matching pay period at row " + i);
          
          // Check the column indices to make sure we're getting the right data
          const startDateValue = payPeriodsData[i][2]; // Column C: start date
          const endDateValue = payPeriodsData[i][4];   // Column E: end date
          
          Logger.log("Raw start date: " + startDateValue);
          Logger.log("Raw end date: " + endDateValue);
          
          // Check if we have valid date values
          if (!startDateValue || !endDateValue) {
            return { 
              success: false, 
              message: `Invalid dates found for pay period ${payPeriodId}. Start: ${startDateValue}, End: ${endDateValue}` 
            };
          }
          
          try {
            // Create proper Date objects
            if (startDateValue instanceof Date) {
              // Clone the date and set time to beginning of day in local timezone
              start = new Date(startDateValue.getFullYear(), startDateValue.getMonth(), startDateValue.getDate(), 0, 0, 0, 0);
            } else {
              // For string dates, parse with explicit components
              const startDateStr = String(startDateValue);
              if (startDateStr.includes('/')) {
                const [month, day, year] = startDateStr.split('/').map(Number);
                start = new Date(year, month - 1, day, 0, 0, 0, 0);
              } else if (startDateStr.includes('-')) {
                const [year, month, day] = startDateStr.split('-').map(Number);
                start = new Date(year, month - 1, day, 0, 0, 0, 0);
              } else {
                start = new Date(startDateValue);
                start.setHours(0, 0, 0, 0);
              }
            }
            
            if (endDateValue instanceof Date) {
              // Clone the date and set time to end of day in local timezone
              end = new Date(endDateValue.getFullYear(), endDateValue.getMonth(), endDateValue.getDate(), 23, 59, 59, 999);
            } else {
              // For string dates, parse with explicit components
              const endDateStr = String(endDateValue);
              if (endDateStr.includes('/')) {
                const [month, day, year] = endDateStr.split('/').map(Number);
                end = new Date(year, month - 1, day, 23, 59, 59, 999);
              } else if (endDateStr.includes('-')) {
                const [year, month, day] = endDateStr.split('-').map(Number);
                end = new Date(year, month - 1, day, 23, 59, 59, 999);
              } else {
                end = new Date(endDateValue);
                end.setHours(23, 59, 59, 999);
              }
            }
            
            // Verify the dates are valid
            if (isNaN(start.getTime()) || isNaN(end.getTime())) {
              throw new Error(`Invalid date conversion - Start: ${startDateValue}, End: ${endDateValue}`);
            }
            
            payPeriod = {
              id: payPeriodsData[i][0],
              name: payPeriodsData[i][1],
              startDate: start,
              endDate: end
            };
            
            Logger.log("Pay period dates set - Start: " + 
                      Utilities.formatDate(start, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") + 
                      ", End: " + 
                      Utilities.formatDate(end, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"));
            break;
          } catch (dateError) {
            Logger.log("Error processing dates: " + dateError);
            return { 
              success: false, 
              message: `Error processing dates for pay period ${payPeriodId}: ${dateError.message}` 
            };
          }
        }
      }
      
      if (!payPeriod) {
        return { success: false, message: `Pay period ${payPeriodId} not found` };
      }
    } else {
      // Handle custom date range
      if (!startDate || !endDate) {
        return { success: false, message: "Start date and end date are required for custom date range" };
      }
      
      try {
        // Parse start date properly based on format
        if (typeof startDate === 'string') {
          // Determine format and parse accordingly
          if (startDate.includes('-') && !startDate.includes('T') && !startDate.includes(' ')) {
            // Format: "YYYY-MM-DD"
            const [year, month, day] = startDate.split('-').map(Number);
            start = new Date(year, month - 1, day, 0, 0, 0, 0);
          } else if (startDate.includes('/') && !startDate.includes(' ')) {
            // Format: "MM/DD/YYYY"
            const [month, day, year] = startDate.split('/').map(Number);
            start = new Date(year, month - 1, day, 0, 0, 0, 0);
          } else {
            // Other formats, let JavaScript try to parse it
            start = new Date(startDate);
            start.setHours(0, 0, 0, 0);
          }
        } else if (startDate instanceof Date) {
          start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate(), 0, 0, 0, 0);
        } else {
          throw new Error("Invalid start date format");
        }
        
        // Parse end date properly based on format
        if (typeof endDate === 'string') {
          // Determine format and parse accordingly
          if (endDate.includes('-') && !endDate.includes('T') && !endDate.includes(' ')) {
            // Format: "YYYY-MM-DD"
            const [year, month, day] = endDate.split('-').map(Number);
            end = new Date(year, month - 1, day, 23, 59, 59, 999);
          } else if (endDate.includes('/') && !endDate.includes(' ')) {
            // Format: "MM/DD/YYYY"
            const [month, day, year] = endDate.split('/').map(Number);
            end = new Date(year, month - 1, day, 23, 59, 59, 999);
          } else {
            // Other formats, let JavaScript try to parse it
            end = new Date(endDate);
            end.setHours(23, 59, 59, 999);
          }
        } else if (endDate instanceof Date) {
          end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate(), 23, 59, 59, 999);
        } else {
          throw new Error("Invalid end date format");
        }
        
        // Verify custom dates are valid
        if (isNaN(start.getTime()) || isNaN(end.getTime())) {
          throw new Error(`Invalid date conversion`);
        }
      } catch (error) {
        return { 
          success: false, 
          message: `Error parsing custom date range: ${error.message}. Start: ${startDate}, End: ${endDate}` 
        };
      }
      
      // Log dates in local timezone format instead of ISO string (which is UTC)
      Logger.log("Using custom date range - Local Start: " + 
                Utilities.formatDate(start, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") + 
                ", Local End: " + 
                Utilities.formatDate(end, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"));
    }

    // Get time logs sheet
    let timeLogsSheet = ss.getSheetByName('Time Logs');
    if (!timeLogsSheet) {
      // Try alternate sheet name
      const altTimeLogsSheet = ss.getSheetByName('Time Logs Test');
      if (!altTimeLogsSheet) {
        return { success: false, message: "Time Logs sheet not found" };
      }
      timeLogsSheet = altTimeLogsSheet;
    }

    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    const allTimeLogs = [];
    const incompleteLogs = [];
    let totalHoursWorked = 0;
    let totalRegularBreakTime = 0;
    let totalLunchBreakTime = 0;
    let totalMissedMinutes = 0;

    // For debugging - log the date range in local timezone format
    Logger.log("Looking for logs between: " + 
               Utilities.formatDate(start, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss") + " and " + 
               Utilities.formatDate(end, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss"));

    // First pass: collect all logs and identify incomplete ones
    // Skip header row
    for (let i = 1; i < timeLogsData.length; i++) {
      // Skip empty rows or rows without employee ID
      if (!timeLogsData[i][1]) continue;
      
      // Convert employee ID to string for comparison and remove decimal if present
      const logEmployeeId = String(timeLogsData[i][1]).replace('.0', '');
      
      // Check if this log is for the requested employee
      if (logEmployeeId !== employeeId) continue;
      
      // Get the log date from column C (index 2)
      let logDate;
      const rawLogDate = timeLogsData[i][2];
      
      // Ensure logDate is properly parsed in local timezone
      if (rawLogDate instanceof Date) {
        logDate = new Date(rawLogDate.getFullYear(), rawLogDate.getMonth(), rawLogDate.getDate());
      } else {
        // Try to parse the date string
        try {
          logDate = new Date(rawLogDate);
          if (isNaN(logDate.getTime())) {
            Logger.log(`Skipping row ${i+1}: Invalid date format in column C: ${rawLogDate}`);
            continue;
          }
        } catch (e) {
          Logger.log(`Skipping row ${i+1}: Error parsing date in column C: ${rawLogDate}`);
          continue;
        }
      }
      
      // Normalized date comparison (ignoring time)
      const normalizedLogDate = new Date(logDate.getFullYear(), logDate.getMonth(), logDate.getDate());
      const normalizedStart = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      const normalizedEnd = new Date(end.getFullYear(), end.getMonth(), end.getDate());
      
      // Check if the log date falls within the date range
      if (normalizedLogDate >= normalizedStart && normalizedLogDate <= normalizedEnd) {
        Logger.log(`Found time log for employee ${employeeId} on ${logDate.toLocaleDateString()}`);
        
        const timeLog = {
          id: timeLogsData[i][0],
          date: Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          clockIn: formatTimeForDisplay(timeLogsData[i][3]),
          clockOut: formatTimeForDisplay(timeLogsData[i][4]),
          regularBreakStart1: formatTimeForDisplay(timeLogsData[i][5]),
          regularBreakEnd1: formatTimeForDisplay(timeLogsData[i][6]),
          regularBreakStart2: formatTimeForDisplay(timeLogsData[i][7]),
          regularBreakEnd2: formatTimeForDisplay(timeLogsData[i][8]),
          lunchBreakStart: formatTimeForDisplay(timeLogsData[i][9]),
          lunchBreakEnd: formatTimeForDisplay(timeLogsData[i][10]),
          hoursWorked: Number(timeLogsData[i][11]) || 0,
          regularBreakTime: Number(timeLogsData[i][12]) || 0,
          lunchBreakTime: Number(timeLogsData[i][13]) || 0,
          netWorkingHours: Number(timeLogsData[i][14]) || 0,
          status: timeLogsData[i][15] || "",
          missedMinutes: Number(timeLogsData[i][25]) || 0,
          rowIndex: i,
          logDate: logDate
        };
        
        allTimeLogs.push(timeLog);
        
        if (timeLog.status === "Incomplete") {
          incompleteLogs.push(timeLog);
        }
        
        totalHoursWorked += timeLog.hoursWorked;
        totalRegularBreakTime += timeLog.regularBreakTime;
        totalLunchBreakTime += timeLog.lunchBreakTime;
        totalMissedMinutes += timeLog.missedMinutes;
      }
    }
    
    // Handle multiple incomplete logs if found
    let warningMessage = null;
    if (incompleteLogs.length > 1) {
      incompleteLogs.sort((a, b) => b.logDate - a.logDate);
      warningMessage = `Found ${incompleteLogs.length} incomplete time logs. Prioritizing the most recent one from ${incompleteLogs[0].date}.`;
      Logger.log(warningMessage);
    }
    
    // Sort all time logs by status and date
    allTimeLogs.sort((a, b) => {
      if (a.status === "Incomplete" && b.status !== "Incomplete") return -1;
      if (a.status !== "Incomplete" && b.status === "Incomplete") return 1;
      return b.logDate - a.logDate;
    });
    
    // Remove internal properties before returning
    const timeLogs = allTimeLogs.map(log => {
      const { rowIndex, logDate, ...cleanLog } = log;
      return cleanLog;
    });
    
    // Calculate net working hours
    const netWorkingHours = totalHoursWorked - totalRegularBreakTime - totalLunchBreakTime;
    
    // Prepare and return the final result
    return {
      success: true,
      employeeId: employeeId,
      payPeriod: {
        id: payPeriodId || 'custom',
        name: payPeriod ? payPeriod.name : 'Custom Date Range',
        startDate: Utilities.formatDate(start, Session.getScriptTimeZone(), "MM/dd/yyyy"),
        endDate: Utilities.formatDate(end, Session.getScriptTimeZone(), "MM/dd/yyyy")
      },
      timeLogs: timeLogs,
      totals: {
        hoursWorked: totalHoursWorked,
        regularBreakTime: totalRegularBreakTime,
        lunchBreakTime: totalLunchBreakTime,
        netWorkingHours: netWorkingHours,
        // Round missed minutes to 2 decimal places
        missedMinutes: Math.round(totalMissedMinutes * 100) / 100
      },
      warning: warningMessage
    };
  } catch (error) {
    Logger.log("Error in getEmployeeTimeReportCustom: " + error.toString());
    return { success: false, message: error.toString() };
  }
}




function getEmployeeTimeReportByPayPeriod(employeeId, payPeriodId) {
  try {
    Logger.log("Getting time report for employee ID: " + employeeId + " and pay period ID: " + payPeriodId);
    
    // Make sure employeeId is a string for consistent comparison
    employeeId = String(employeeId).replace('.0', '');
    
    // Get the specified pay period
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      return { success: false, message: "Pay Periods sheet not found" };
    }
    
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    let payPeriod = null;
    
    // Find the specified pay period
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (String(payPeriodsData[i][0]) === String(payPeriodId)) {
        payPeriod = {
          id: payPeriodsData[i][0],
          name: payPeriodsData[i][1],
          startDate: new Date(payPeriodsData[i][2]),
          endDate: new Date(payPeriodsData[i][4])
        };
        break;
      }
    }
    
    if (!payPeriod) {
      return { success: false, message: "Pay period not found" };
    }
    
    // Get time logs for this employee within the pay period
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    if (!timeLogsSheet) {
      return { success: false, message: "Time Logs sheet not found" };
    }
    
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    const allTimeLogs = [];
    let totalHoursWorked = 0;
    let totalRegularBreakTime = 0;
    let totalLunchBreakTime = 0;
    let totalMissedMinutes = 0;
    
    // Process time logs
    for (let i = 1; i < timeLogsData.length; i++) {
      // Skip rows without employee ID
      if (!timeLogsData[i][1]) continue;
      
      // Convert employee ID to string for comparison
      const logEmployeeId = String(timeLogsData[i][1]).replace('.0', '');
      
      // Check if this log is for the requested employee
      if (logEmployeeId !== employeeId) continue;
      
      // Get the log date from column C
      const logDate = new Date(timeLogsData[i][2]);
      
      // Check if the log date falls within the pay period
      if (logDate >= payPeriod.startDate && logDate <= payPeriod.endDate) {
        // Round missed minutes to 2 decimal places for each log
        const missedMinutes = Number(timeLogsData[i][25]) || 0;
        const roundedMissedMinutes = Math.round(missedMinutes * 100) / 100;
        
        const timeLog = {
          id: timeLogsData[i][0],
          date: Utilities.formatDate(logDate, Session.getScriptTimeZone(), "yyyy-MM-dd"),
          clockIn: formatTimeForDisplay(timeLogsData[i][3]),
          clockOut: formatTimeForDisplay(timeLogsData[i][4]),
          hoursWorked: Number(timeLogsData[i][11]) || 0,
          regularBreakTime: Number(timeLogsData[i][12]) || 0,
          lunchBreakTime: Number(timeLogsData[i][13]) || 0,
          netWorkingHours: Number(timeLogsData[i][14]) || 0,
          status: timeLogsData[i][15] || "",
          missedMinutes: roundedMissedMinutes
        };
        
        allTimeLogs.push(timeLog);
        
        // Add to totals
        totalHoursWorked += timeLog.hoursWorked;
        totalRegularBreakTime += timeLog.regularBreakTime;
        totalLunchBreakTime += timeLog.lunchBreakTime;
        totalMissedMinutes += roundedMissedMinutes; // Add the rounded value
      }
    }
    
    // Sort logs by date (newest first)
    allTimeLogs.sort((a, b) => new Date(b.date) - new Date(a.date));
    
    // Round the total missed minutes to 2 decimal places
    const roundedTotalMissedMinutes = Math.round(totalMissedMinutes * 100) / 100;
    
    return {
      success: true,
      employeeId: employeeId,
      payPeriod: {
        id: payPeriod.id,
        name: payPeriod.name,
        startDate: Utilities.formatDate(payPeriod.startDate, Session.getScriptTimeZone(), "MM/dd/yyyy"),
        endDate: Utilities.formatDate(payPeriod.endDate, Session.getScriptTimeZone(), "MM/dd/yyyy")
      },
      timeLogs: allTimeLogs,
      totals: {
        hoursWorked: totalHoursWorked,
        regularBreakTime: totalRegularBreakTime,
        lunchBreakTime: totalLunchBreakTime,
        netWorkingHours: totalHoursWorked - totalRegularBreakTime - totalLunchBreakTime,
        missedMinutes: roundedTotalMissedMinutes
      }
    };
    
  } catch (error) {
    Logger.log("Error in getEmployeeTimeReportByPayPeriod: " + error.toString());
    return { success: false, message: error.toString() };
  }
}



// Helper function to format time values for display
function formatTimeForDisplay(timeValue) {
  if (!timeValue) return "";
  
  try {
    if (timeValue instanceof Date) {
      return Utilities.formatDate(timeValue, Session.getScriptTimeZone(), "HH:mm:ss");
    } else if (typeof timeValue === 'string') {
      // If it's a string that includes both date and time
      if (timeValue.includes(' ')) {
        const parts = timeValue.split(' ');
        if (parts.length > 1) {
          return parts[1]; // Return just the time part
        }
        return timeValue;
      }
      return timeValue;
    }
    return "";
  } catch (e) {
    Logger.log("Error formatting time: " + e.toString());
    return String(timeValue);
  }
}

// Helper function to format time values for display
function formatTimeForDisplay(timeValue) {
  if (!timeValue) return "";
  
  try {
    if (timeValue instanceof Date) {
      return Utilities.formatDate(timeValue, Session.getScriptTimeZone(), "HH:mm:ss");
    } else if (typeof timeValue === 'string') {
      // If it's a string that includes both date and time
      if (timeValue.includes(' ')) {
        const parts = timeValue.split(' ');
        if (parts.length > 1) {
          return parts[1]; // Return just the time part
        }
        return timeValue;
      }
      return timeValue;
    }
    return "";
  } catch (e) {
    Logger.log("Error formatting time: " + e.toString());
    return String(timeValue);
  }
}


//
//
//
//
//
//
//
//
//
//
//
//
// operator tabs  



/**
 * Analyzes operator attendance for benefits eligibility
 * @param {string} payPeriodId - The ID of the pay period to analyze
 * @return {Array} - Array of employee attendance results
 */
function analyzeOperatorAttendance(payPeriodId) {
  try {
    Logger.log("Starting analyzeOperatorAttendance for pay period: " + payPeriodId);
    
    // First, get the qualifying shifts data using the more accurate function
    const qualifyingShiftsData = calculateEmployeeQualifyingShifts(payPeriodId);
    Logger.log(`Retrieved qualifying shifts data for ${qualifyingShiftsData.length} employees`);
    
    // Create a lookup map by employee ID for quick access
    const qualifyingShiftsMap = {};
    qualifyingShiftsData.forEach(data => {
      qualifyingShiftsMap[data.employeeId] = data;
    });
    
    if (!initSpreadsheet()) {
      Logger.log("Failed to initialize spreadsheet");
      return [];
    }
    
    // Get pay period dates
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      Logger.log("Pay Periods sheet not found");
      return [];
    }
    
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    let payPeriod = null;
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][0] == payPeriodId) {
        payPeriod = {
          id: payPeriodsData[i][0],
          name: payPeriodsData[i][1],
          startDate: new Date(payPeriodsData[i][2]),
          endDate: new Date(payPeriodsData[i][4]) // Fix: Use index 4 for end date
        };
        break;
      }
    }
    
    if (!payPeriod) {
      Logger.log("Pay period not found with ID: " + payPeriodId);
      return [];
    }
    
    Logger.log("Analyzing pay period: " + payPeriod.name + 
              " (" + Utilities.formatDate(payPeriod.startDate, Session.getScriptTimeZone(), "yyyy-MM-dd") + 
              " to " + Utilities.formatDate(payPeriod.endDate, Session.getScriptTimeZone(), "yyyy-MM-dd") + ")");
    
    // Get employees with assigned shifts
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      Logger.log("Employee Master Data sheet not found");
      return [];
    }
    
    const employeeData = employeeSheet.getDataRange().getValues();
    const employees = [];
    
    // Find header row indices for easier reference
    const headers = employeeData[0];
    const idIdx = headers.indexOf('Employee ID');
    const firstNameIdx = headers.indexOf('First Name');
    const lastNameIdx = headers.indexOf('Last Name');
    const shiftIdx = headers.indexOf('Shift');
    const shiftIdIdx = headers.indexOf('Shift ID'); // Add this to get the Shift ID
    const statusIdx = headers.indexOf('Status');
    
    // Find all active employees with assigned shifts
    for (let i = 1; i < employeeData.length; i++) {
      if (
        employeeData[i][statusIdx] === 'Active' && 
        employeeData[i][shiftIdx] && 
        employeeData[i][shiftIdx].trim() !== ''
      ) {
        employees.push({
          id: employeeData[i][idIdx],
          firstName: employeeData[i][firstNameIdx],
          lastName: employeeData[i][lastNameIdx],
          name: employeeData[i][firstNameIdx] + ' ' + employeeData[i][lastNameIdx],
          shift: employeeData[i][shiftIdx],
          shiftId: employeeData[i][shiftIdIdx] // Store the shift ID as well
        });
      }
    }
    
    if (employees.length === 0) {
      Logger.log("No active employees with assigned shifts found");
      return [];
    }
    
    Logger.log("Found " + employees.length + " employees with assigned shifts");
    
    // Get shift details
    const shiftsSheet = ss.getSheetByName('Shifts');
    if (!shiftsSheet) {
      Logger.log("Shifts sheet not found");
      return [];
    }
    
    const shiftsData = shiftsSheet.getDataRange().getValues();
    const shifts = {};
    
    // Skip header row and process shifts
    for (let i = 1; i < shiftsData.length; i++) {
      // Use Shift ID as the key instead of name
      const shiftId = shiftsData[i][0]; // Assuming Shift ID is in column A
      const shiftName = shiftsData[i][1]; // Assuming Shift Name is in column B
      
      if (shiftId) {
        shifts[shiftId] = {
          id: shiftId,
          name: shiftName,
          weekAStartTime: shiftsData[i][3], // Adjust these indices based on your sheet
          weekAEndTime: shiftsData[i][4],
          weekBStartTime: shiftsData[i][5] || shiftsData[i][3], // Default to week A if not specified
          weekBEndTime: shiftsData[i][6] || shiftsData[i][4],   // Default to week A if not specified
          isOvernight: isOvernightShift(shiftsData[i][3], shiftsData[i][4])
        };
        
        // Also add an entry with the name as key for backward compatibility
        shifts[shiftName] = shifts[shiftId];
      }
    }
    
    // Get time logs for the pay period
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    if (!timeLogsSheet) {
      Logger.log("Time Logs sheet not found");
      return [];
    }
    
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Analyze attendance for each employee
    const results = [];
    
    for (const employee of employees) {
      Logger.log("Analyzing attendance for employee: " + employee.name);
      
      // Try to find the shift using shift ID first, then fall back to shift name
      let employeeShift = null;
      if (employee.shiftId && shifts[employee.shiftId]) {
        employeeShift = shifts[employee.shiftId];
      } else if (shifts[employee.shift]) {
        employeeShift = shifts[employee.shift];
      }
      
      if (!employeeShift) {
        Logger.log("Shift not found for employee: " + employee.name + ", shift: " + employee.shift + ", shiftId: " + employee.shiftId);
        
        // Create a default shift to avoid skipping the employee
        employeeShift = {
          name: employee.shift,
          weekAStartTime: "09:00:00",
          weekAEndTime: "17:00:00",
          weekBStartTime: "09:00:00",
          weekBEndTime: "17:00:00",
          isOvernight: false
        };
      }
      
      // Get all time logs for this employee within the pay period
      const employeeLogs = [];
      
      for (let i = 1; i < timeLogsData.length; i++) {
        const logDate = new Date(timeLogsData[i][2]);
        
        if (
          timeLogsData[i][1] == employee.id && 
          logDate >= payPeriod.startDate && 
          logDate <= payPeriod.endDate &&
          timeLogsData[i][15] === "Complete" // Only count completed logs
        ) {
          employeeLogs.push({
            logId: timeLogsData[i][0],
            date: logDate,
            clockIn: timeLogsData[i][3] ? new Date(timeLogsData[i][3]) : null,
            clockOut: timeLogsData[i][4] ? new Date(timeLogsData[i][4]) : null,
            regBreak1Start: timeLogsData[i][5] ? new Date(timeLogsData[i][5]) : null,
            regBreak1End: timeLogsData[i][6] ? new Date(timeLogsData[i][6]) : null,
            regBreak2Start: timeLogsData[i][7] ? new Date(timeLogsData[i][7]) : null,
            regBreak2End: timeLogsData[i][8] ? new Date(timeLogsData[i][8]) : null,
            lunchStart: timeLogsData[i][9] ? new Date(timeLogsData[i][9]) : null,
            lunchEnd: timeLogsData[i][10] ? new Date(timeLogsData[i][10]) : null,
            totalHours: parseFloat(timeLogsData[i][11]) || 0,
            regBreakTotal: parseFloat(timeLogsData[i][12]) || 0,
            lunchTotal: parseFloat(timeLogsData[i][13]) || 0,
            netHours: parseFloat(timeLogsData[i][14]) || 0,
            lateMinutes: parseFloat(timeLogsData[i][23]) || 0,     // Column X
            earlyMinutes: parseFloat(timeLogsData[i][24]) || 0,    // Column Y
            totalMissedMinutes: parseFloat(timeLogsData[i][25]) || 0, // Column Z
            notes: timeLogsData[i][16] || "",
            regBreak1Missed: parseFloat(timeLogsData[i][20]) || 0, // Column U
            regBreak2Missed: parseFloat(timeLogsData[i][21]) || 0, // Column V
            lunchBreakMissed: parseFloat(timeLogsData[i][22]) || 0, // Column W
          });
        }
      }
      
      if (employeeLogs.length === 0) {
        Logger.log("No time logs found for employee: " + employee.name);
        continue;
      }
      
      // Calculate statistics for eligibility
      let totalHours = 0;
      let totalMissedMinutes = 0;
      let lateMinutes = 0;
      let earlyMinutes = 0;
      let breakMissedMinutes = 0;
      const dailyLogs = [];
      
      // Sort logs by date
      employeeLogs.sort((a, b) => a.date - b.date);
      
      for (const log of employeeLogs) {
        totalHours += log.netHours;
        totalMissedMinutes += log.totalMissedMinutes;
        lateMinutes += log.lateMinutes;
        earlyMinutes += log.earlyMinutes;
        
        // Calculate missed break minutes (the difference between totalMissedMinutes and late/early minutes)
        const logBreakMissed = log.totalMissedMinutes - log.lateMinutes - log.earlyMinutes;
        breakMissedMinutes += logBreakMissed > 0 ? logBreakMissed : 0;
        
        // Format the date to YYYY-MM-DD for the set to count unique shifts
        const logDateStr = Utilities.formatDate(log.date, Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        // Add to daily logs for detailed view
        dailyLogs.push({
          logId: log.logId,
          date: logDateStr,
          clockIn: log.clockIn ? Utilities.formatDate(log.clockIn, Session.getScriptTimeZone(), "HH:mm:ss") : null,
          clockOut: log.clockOut ? Utilities.formatDate(log.clockOut, Session.getScriptTimeZone(), "HH:mm:ss") : null,
          hours: log.netHours,
          
          // Add these fields:
          regBreak1Missed: log.regBreak1Missed || 0,
          regBreak2Missed: log.regBreak2Missed || 0,
          lunchBreakMissed: log.lunchBreakMissed || 0,
          lateArrival: log.lateMinutes || 0,
          earlyDeparture: log.earlyMinutes || 0,
          missedMinutes: log.totalMissedMinutes || 0,
          notes: log.notes
        });
      }
      
      // Get the qualifying shifts count from the calculateEmployeeQualifyingShifts function
      const qualifyingData = qualifyingShiftsMap[employee.id] || { qualifyingShifts: 0, totalHours: 0 };
      const shiftsWorked = qualifyingData.qualifyingShifts;
      
      Logger.log(`Employee ${employee.name} has ${shiftsWorked} qualifying shifts according to calculateEmployeeQualifyingShifts`);
      
      // Determine eligibility using the qualifying shifts count
      const isEligible = (
        totalHours >= 66.5 && 
        shiftsWorked >= 7 && 
        totalMissedMinutes <= 20
      );
      
      results.push({
        employeeId: employee.id,
        name: employee.name,
        shift: employee.shift,
        totalHours: totalHours,
        shiftsWorked: shiftsWorked,  // Use the qualifying shifts count
        totalMissedMinutes: totalMissedMinutes,
        lateMinutes: lateMinutes,
        earlyMinutes: earlyMinutes,
        breakMissedMinutes: breakMissedMinutes,
        isEligible: isEligible,
        dailyLogs: dailyLogs
      });
      
      Logger.log(`Analysis for ${employee.name}: Hours=${totalHours.toFixed(2)}, Shifts=${shiftsWorked}, Missed=${totalMissedMinutes}, Eligible=${isEligible}`);
    }
    
    return results;
    
  } catch (error) {
    Logger.log("Error in analyzeOperatorAttendance: " + error.toString());
    return [];
  }
}



/**
 * Calculates qualifying shifts for all employees with proper overnight shift handling
 * @param {string} payPeriodId - The ID of the pay period to analyze
 * @return {Array} - Array of employee qualifying shift results
 */
function calculateEmployeeQualifyingShifts(payPeriodId) {
  try {
    Logger.log("Starting calculateEmployeeQualifyingShifts for pay period: " + payPeriodId);
    
    if (!initSpreadsheet()) {
      Logger.log("Failed to initialize spreadsheet");
      return [];
    }
    
    // Get pay period dates
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      Logger.log("Pay Periods sheet not found");
      return [];
    }
    
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    let payPeriod = null;
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][0] == payPeriodId) {
        payPeriod = {
          id: payPeriodsData[i][0],
          name: payPeriodsData[i][1],
          startDate: new Date(payPeriodsData[i][2]),
          endDate: new Date(payPeriodsData[i][4])
        };
        break;
      }
    }
    
    if (!payPeriod) {
      Logger.log("Pay period not found with ID: " + payPeriodId);
      return [];
    }
    
    Logger.log("Analyzing qualifying shifts for pay period: " + payPeriod.name);
    
    // Get all employees
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      Logger.log("Employee Master Data sheet not found");
      return [];
    }
    
    const employeeData = employeeSheet.getDataRange().getValues();
    const employees = [];
    
    // Find header row indices
    const headers = employeeData[0];
    const idIdx = headers.indexOf('Employee ID');
    const firstNameIdx = headers.indexOf('First Name');
    const lastNameIdx = headers.indexOf('Last Name');
    const statusIdx = headers.indexOf('Status');
    const shiftIdx = headers.indexOf('Shift');
    const shiftIdIdx = headers.indexOf('Shift ID');
    
    // Find all active employees
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][statusIdx] === 'Active') {
        employees.push({
          id: employeeData[i][idIdx],
          firstName: employeeData[i][firstNameIdx],
          lastName: employeeData[i][lastNameIdx],
          name: employeeData[i][firstNameIdx] + ' ' + employeeData[i][lastNameIdx],
          shift: employeeData[i][shiftIdx],
          shiftId: employeeData[i][shiftIdIdx]
        });
      }
    }
    
    if (employees.length === 0) {
      Logger.log("No active employees found");
      return [];
    }
    
    Logger.log("Found " + employees.length + " active employees");
    
    // Get shift details
    const shiftsSheet = ss.getSheetByName('Shifts');
    if (!shiftsSheet) {
      Logger.log("Shifts sheet not found");
      return [];
    }
    
    const shiftsData = shiftsSheet.getDataRange().getValues();
    const shifts = {};
    
    // Skip header row and process shifts
    for (let i = 1; i < shiftsData.length; i++) {
      const shiftId = shiftsData[i][0]; // Assuming Shift ID is in column A
      const shiftName = shiftsData[i][1]; // Assuming Shift Name is in column B
      
      if (shiftId) {
        shifts[shiftId] = {
          id: shiftId,
          name: shiftName,
          weekAStartTime: shiftsData[i][3],
          weekAEndTime: shiftsData[i][4],
          weekBStartTime: shiftsData[i][5] || shiftsData[i][3],
          weekBEndTime: shiftsData[i][6] || shiftsData[i][4],
          isOvernight: isOvernightShift(shiftsData[i][3], shiftsData[i][4])
        };
        
        // Also add an entry with the name as key for backward compatibility
        shifts[shiftName] = shifts[shiftId];
      }
    }
    
    // Get time logs for the pay period
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    if (!timeLogsSheet) {
      Logger.log("Time Logs sheet not found");
      return [];
    }
    
    const timeLogsData = timeLogsSheet.getDataRange().getValues();
    
    // Calculate qualifying shifts for each employee
    const results = [];
    
    for (const employee of employees) {
      Logger.log("Calculating qualifying shifts for employee: " + employee.name);
      
      // Try to find the shift using shift ID first, then fall back to shift name
      let employeeShift = null;
      if (employee.shiftId && shifts[employee.shiftId]) {
        employeeShift = shifts[employee.shiftId];
      } else if (shifts[employee.shift]) {
        employeeShift = shifts[employee.shift];
      }
      
      // Default to a regular day shift if no shift found
      if (!employeeShift) {
        Logger.log("Shift not found for employee: " + employee.name + ", using default day shift");
        employeeShift = {
          name: "Default Day Shift",
          isOvernight: false
        };
      }
      
      // Get all time logs for this employee within the pay period
      const employeeLogs = [];
      
      for (let i = 1; i < timeLogsData.length; i++) {
        const logDate = new Date(timeLogsData[i][2]);
        
        if (
          timeLogsData[i][1] == employee.id && 
          logDate >= payPeriod.startDate && 
          logDate <= payPeriod.endDate &&
          timeLogsData[i][15] === "Complete" // Only count completed logs
        ) {
          employeeLogs.push({
            logId: timeLogsData[i][0],
            date: logDate,
            clockIn: timeLogsData[i][3] ? new Date(timeLogsData[i][3]) : null,
            clockOut: timeLogsData[i][4] ? new Date(timeLogsData[i][4]) : null,
            totalHours: parseFloat(timeLogsData[i][11]) || 0,
            netHours: parseFloat(timeLogsData[i][14]) || 0,
            notes: timeLogsData[i][16] || ""
          });
        }
      }
      
      if (employeeLogs.length === 0) {
        Logger.log("No time logs found for employee: " + employee.name);
        results.push({
          employeeId: employee.id,
          name: employee.name,
          totalHours: 0,
          qualifyingShifts: 0,
          qualifyingShiftsWithPaidBreaks: 0,
          dailyLogs: []
        });
        continue;
      }
      
      // Sort logs by date and time
      employeeLogs.sort((a, b) => a.clockIn - b.clockIn);
      
      // Calculate statistics
      let totalHours = 0;
      const dailyShiftHours = {}; // Format: "YYYY-MM-DD" -> hours
      const dailyLogs = [];
      
      // Process each log with overnight shift handling
      for (const log of employeeLogs) {
        if (log.clockIn && log.clockOut) {
          const clockInDate = new Date(log.clockIn);
          let shiftDateStr;
          
          // Determine the shift date based on overnight status
          if (employeeShift.isOvernight) {
            // For overnight shifts, determine if this belongs to previous day's shift
            const clockInHour = clockInDate.getHours();
            const clockInMinutes = clockInDate.getMinutes();
            
            // If clock-in is between midnight and 5 AM, it belongs to previous day's shift
            if (clockInHour < 5 || (clockInHour === 5 && clockInMinutes === 0)) {
              const prevDate = new Date(clockInDate);
              prevDate.setDate(prevDate.getDate() - 1);
              shiftDateStr = Utilities.formatDate(prevDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            } else {
              shiftDateStr = Utilities.formatDate(clockInDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
            }
          } else {
            // For regular shifts, just use the calendar date
            shiftDateStr = Utilities.formatDate(clockInDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          
          // Add hours to the appropriate shift date
          if (!dailyShiftHours[shiftDateStr]) {
            dailyShiftHours[shiftDateStr] = 0;
          }
          dailyShiftHours[shiftDateStr] += log.netHours;
          
          // Add to total hours
          totalHours += log.netHours;
          
          // Add to daily logs
          dailyLogs.push({
            date: shiftDateStr,
            clockIn: Utilities.formatDate(log.clockIn, Session.getScriptTimeZone(), "HH:mm:ss"),
            clockOut: Utilities.formatDate(log.clockOut, Session.getScriptTimeZone(), "HH:mm:ss"),
            hours: log.netHours,
            notes: log.notes
          });
        }
      }
      
      // Count qualifying shifts (7+ hours)
      let qualifyingShifts = 0;
      let qualifyingShiftsWithPaidBreaks = 0;
      
      for (const shiftDate in dailyShiftHours) {
        const hours = dailyShiftHours[shiftDate];
        
        if (hours >= 7) {
          qualifyingShifts++;
          
          // For simplicity, assume all qualifying shifts have paid breaks
          // In a real implementation, you would check break compliance here
          qualifyingShiftsWithPaidBreaks++;
        }
      }
      
      results.push({
        employeeId: employee.id,
        name: employee.name,
        totalHours: totalHours,
        qualifyingShifts: qualifyingShifts,
        qualifyingShiftsWithPaidBreaks: qualifyingShiftsWithPaidBreaks,
        dailyLogs: dailyLogs,
        shiftDetails: Object.entries(dailyShiftHours).map(([key, hours]) => ({
          shiftDate: key,
          hours: hours,
          counted: hours >= 7
        })),
        shiftType: employeeShift.isOvernight ? "Overnight" : "Regular"
      });
      
      Logger.log(`Results for ${employee.name}: Hours=${totalHours.toFixed(2)}, Qualifying Shifts=${qualifyingShifts}, Qualifying Shifts with Paid Breaks=${qualifyingShiftsWithPaidBreaks}, ShiftType=${employeeShift.isOvernight ? "Overnight" : "Regular"}`);
    }
    
    return results;
    
  } catch (error) {
    Logger.log("Error in calculateEmployeeQualifyingShifts: " + error.toString());
    return [];
  }
}

/**
 * Helper function to determine if a shift is overnight based on start and end times
 * @param {string} startTime - Shift start time (HH:MM:SS)
 * @param {string} endTime - Shift end time (HH:MM:SS)
 * @return {boolean} - True if this is an overnight shift
 */
function isOvernightShift(startTime, endTime) {
  if (!startTime || !endTime) return false;
  
  // Parse times into hours
  const startParts = startTime.split(':');
  const endParts = endTime.split(':');
  
  if (startParts.length < 2 || endParts.length < 2) return false;
  
  const startHour = parseInt(startParts[0], 10);
  const endHour = parseInt(endParts[0], 10);
  
  // If end time is earlier than start time, it's an overnight shift
  // Also consider it overnight if start time is after 12:00 PM and end time is before 12:00 PM
  return endHour < startHour || (startHour >= 12 && endHour < 12);
}




/**
 * Determines if a shift is overnight by comparing start and end times
 * @param {string} startTime - The shift start time (e.g., "22:00:00")
 * @param {string} endTime - The shift end time (e.g., "06:00:00")
 * @return {boolean} - True if the shift is overnight
 */
function isOvernightShift(startTime, endTime) {
  try {
    if (!startTime || !endTime) return false;
    
    // Extract hours from the time strings
    let startHour = 0;
    let endHour = 0;
    
    if (typeof startTime === 'string') {
      const startParts = startTime.split(':');
      startHour = parseInt(startParts[0], 10);
    } else if (startTime instanceof Date) {
      startHour = startTime.getHours();
    }
    
    if (typeof endTime === 'string') {
      const endParts = endTime.split(':');
      endHour = parseInt(endParts[0], 10);
    } else if (endTime instanceof Date) {
      endHour = endTime.getHours();
    }
    
    // If end time is earlier than start time, it's overnight
    return endHour < startHour;
  } catch (error) {
    Logger.log("Error in isOvernightShift: " + error.toString());
    return false;
  }
}


/**
 * Gets detailed attendance data for a specific employee
 * @param {string} employeeId - The ID of the employee
 * @return {Object} - Employee attendance details
 */
function getEmployeeAttendanceDetails(employeeId) {
  try {
    Logger.log("Getting attendance details for employee: " + employeeId);
    
    // Get the currently active pay period
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    
    let activePeriodId = null;
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][6] === "Active") {
        activePeriodId = payPeriodsData[i][0];
        break;
      }
    }
    
    if (!activePeriodId) {
      Logger.log("No active pay period found");
      return { success: false, message: "No active pay period found" };
    }
    
    // Call the existing function to analyze attendance
    const allResults = analyzeOperatorAttendance(activePeriodId);
    
    // Find the specific employee
    const employeeResult = allResults.find(emp => String(emp.employeeId) === String(employeeId));
    
    if (!employeeResult) {
      Logger.log("Employee not found in attendance results");
      return { success: false, message: "Employee not found in attendance results" };
    }
    
    return { success: true, data: employeeResult };
    
  } catch (error) {
    Logger.log("Error in getEmployeeAttendanceDetails: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets a specific time log by ID and row index
 * @param {string} logId - Log ID to find
 * @param {number} rowIndex - The row index in the sheet
 * @return {Object} The time log data or null if not found
 */
function getTimeLogById(logId, rowIndex) {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return null;
    }
    
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    
    // Check if rowIndex is provided and valid
    if (rowIndex && rowIndex > 1 && rowIndex <= timeLogsSheet.getLastRow()) {
      const row = timeLogsSheet.getRange(rowIndex, 1, 1, timeLogsSheet.getLastColumn()).getValues()[0];
      
      // Get employee name
      const employeeId = row[1]; // Column B: Employee ID
      const employeeName = getEmployeeName(employeeId);
      
      return {
        rowIndex: rowIndex,
        logId: row[0], // Column A: Log ID
        employeeId: employeeId,
        employeeName: employeeName,
        date: Utilities.formatDate(new Date(row[2]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
        clockInTime: formatTimeForDisplay(row[3]), // Column D: Clock-in time
        clockOutTime: formatTimeForDisplay(row[4]), // Column E: Clock-out time
        regularBreakStart1: formatTimeForDisplay(row[5]), // Column F
        regularBreakEnd1: formatTimeForDisplay(row[6]), // Column G
        regularBreakStart2: formatTimeForDisplay(row[7]), // Column H
        regularBreakEnd2: formatTimeForDisplay(row[8]), // Column I
        lunchBreakStart: formatTimeForDisplay(row[9]), // Column J
        lunchBreakEnd: formatTimeForDisplay(row[10]), // Column K
        status: row[15] || '', // Column P: Status
        regBreak1Missed: row[20] || 0, // Column U: Regular Break 1 Missed Minutes
        regBreak2Missed: row[21] || 0, // Column V: Regular Break 2 Missed Minutes
        lunchBreakMissed: row[22] || 0, // Column W: Lunch Break Missed Minutes
        lateArrival: row[23] || 0, // Column X: Late Minutes
        earlyDeparture: row[24] || 0, // Column Y: Early Departure Minutes
        totalMissedMinutes: row[25] || 0 // Column Z: Total Missed Minutes
      };
    }
    
    // If rowIndex is not provided or invalid, search by logId
    const data = timeLogsSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === logId) {
        const employeeId = data[i][1];
        const employeeName = getEmployeeName(employeeId);
        
        return {
          rowIndex: i + 1,
          logId: data[i][0],
          employeeId: employeeId,
          employeeName: employeeName,
          date: Utilities.formatDate(new Date(data[i][2]), Session.getScriptTimeZone(), "yyyy-MM-dd"),
          clockInTime: formatTimeForDisplay(data[i][3]),
          clockOutTime: formatTimeForDisplay(data[i][4]),
          regularBreakStart1: formatTimeForDisplay(data[i][5]),
          regularBreakEnd1: formatTimeForDisplay(data[i][6]),
          regularBreakStart2: formatTimeForDisplay(data[i][7]),
          regularBreakEnd2: formatTimeForDisplay(data[i][8]),
          lunchBreakStart: formatTimeForDisplay(data[i][9]),
          lunchBreakEnd: formatTimeForDisplay(data[i][10]),
          status: data[i][15] || '',
          regBreak1Missed: data[i][20] || 0,
          regBreak2Missed: data[i][21] || 0,
          lunchBreakMissed: data[i][22] || 0,
          lateArrival: data[i][23] || 0,
          earlyDeparture: data[i][24] || 0,
          totalMissedMinutes: data[i][25] || 0
        };
      }
    }
    
    return null;
  } catch (error) {
    Logger.log("Error in getTimeLogById: " + error.toString());
    return null;
  }
}

/**
 * Helper function to format time values for display
 */
function formatTimeForDisplay(timeValue) {
  if (!timeValue) return '';
  
  try {
    return Utilities.formatDate(new Date(timeValue), Session.getScriptTimeZone(), "HH:mm:ss");
  } catch (e) {
    return '';
  }
}


/**
 * Grants eligible employees additional hours to reach 80 hours for the pay period
 * @param {string} payPeriodId - The ID of the pay period
 * @return {Object} Result of the operation
 */
function grantEligibleEmployees80Hours(payPeriodId) {
  try {
    Logger.log("Starting to grant 80 hours to eligible employees for pay period: " + payPeriodId);
    
    // Get the analysis results for the pay period
    const analysisResults = analyzeOperatorAttendance(payPeriodId);
    
    // Filter for eligible employees only
    const eligibleEmployees = analysisResults.filter(employee => employee.isEligible);
    
    Logger.log("Found " + eligibleEmployees.length + " eligible employees");
    if (eligibleEmployees.length === 0) {
      return { success: false, message: "No eligible employees found" };
    }
    
    // Get pay period details
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    let payPeriod = null;
    
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][0] == payPeriodId) {
        payPeriod = {
          id: payPeriodsData[i][0],
          name: payPeriodsData[i][1],
          startDate: new Date(payPeriodsData[i][2]),
          endDate: new Date(payPeriodsData[i][4])
        };
        break;
      }
    }
    
    if (!payPeriod) {
      return { success: false, message: "Pay period not found" };
    }
    
    // Get time logs sheet
    const timeLogsSheet = ss.getSheetByName('Time Logs');
    
    // Generate logs for each eligible employee
    const results = [];
    for (const employee of eligibleEmployees) {
      // Calculate additional hours needed to reach 80
      const currentHours = employee.totalHours;
      const additionalHoursNeeded = 80 - currentHours;
      
      if (additionalHoursNeeded <= 0) {
        Logger.log("Employee " + employee.name + " already has 80+ hours");
        results.push({
          employeeId: employee.employeeId,
          name: employee.name,
          success: false,
          message: "Employee already has 80+ hours"
        });
        continue;
      }
      
      Logger.log("Granting " + additionalHoursNeeded.toFixed(2) + " additional hours to " + employee.name);
      
      try {
        // Generate a unique log ID
        const logId = "AUTO-" + new Date().getTime() + "-" + employee.employeeId;
        
        // Set date to the last day of the pay period
        const logDate = new Date(payPeriod.endDate);

        // Subtract one day (24 hours) from the date
        logDate.setDate(logDate.getDate() - 1);
        
        // Calculate clock times
        const clockInTime = new Date(logDate);
        clockInTime.setHours(1, 0, 0, 0); // 1:00 AM
        
        const clockOutTime = new Date(clockInTime);
        // Add the necessary hours
        clockOutTime.setTime(clockInTime.getTime() + (additionalHoursNeeded * 60 * 60 * 1000));
        
        // Format datetime strings
        const dateStr = Utilities.formatDate(logDate, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
        const clockInTimeStr = Utilities.formatDate(clockInTime, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
        const clockOutTimeStr = Utilities.formatDate(clockOutTime, Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
        
        // Calculate duration in hours
        const hours = additionalHoursNeeded;
        const regularBreakHours = 0; // No breaks for these auto-generated entries
        const lunchBreakHours = 0;
        const netHours = hours;
        
        // Append new time log
        timeLogsSheet.appendRow([
          logId,                // Column A: Log ID
          employee.employeeId,  // Column B: Employee ID
          dateStr,              // Column C: Date
          clockInTimeStr,       // Column D: Clock-in time
          clockOutTimeStr,      // Column E: Clock-out time
          "",                   // Column F: Regular Break 1 Start
          "",                   // Column G: Regular Break 1 End
          "",                   // Column H: Regular Break 2 Start
          "",                   // Column I: Regular Break 2 End
          "",                   // Column J: Lunch Break Start
          "",                   // Column K: Lunch Break End
          hours,                // Column L: Total Hours Worked
          regularBreakHours,    // Column M: Total Regular Break Time
          lunchBreakHours,      // Column N: Total Lunch Break Time
          netHours,             // Column O: Net Working Hours
          "Complete",           // Column P: Status
          "80 hours Payperiod Incentive remaining hours ", // Column Q: Notes
        ]);
        
        results.push({
          employeeId: employee.employeeId,
          name: employee.name,
          success: true,
          additionalHours: additionalHoursNeeded,
          totalHours: 80
        });
        
      } catch (error) {
        Logger.log("Error creating time log for employee " + employee.name + ": " + error.toString());
        results.push({
          employeeId: employee.employeeId,
          name: employee.name,
          success: false,
          message: error.toString()
        });
      }
    }
    
    return {
      success: true,
      message: `Granted additional hours to ${results.filter(r => r.success).length} eligible employees`,
      results: results
    };
    
  } catch (error) {
    Logger.log("Error in grantEligibleEmployees80Hours: " + error.toString());
    return { success: false, message: error.toString() };
  }
}
















/**
 * Gets the current PTO balance for an employee
 * @param {string} employeeId - The employee's ID
 * @return {Object} PTO balance information
 */
function getEmployeePTOBalance(employeeId) {
  try {
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    const employeeData = employeeSheet.getDataRange().getValues();
    
    // Skip header row and find employee
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] == employeeId) {
        return {
          success: true,
          balance: employeeData[i][13], // Current PTO Balance column
          accrualRate: employeeData[i][12] // PTO Accrual Rate column
        };
      }
    }
    
    return { success: false, message: "Employee not found" };
  } catch (e) {
    Logger.log("Error getting PTO balance: " + e.toString());
    return { success: false, message: "Error retrieving PTO balance" };
  }
}


/**
 * Submits a PTO request
 * @param {Object} ptoRequest - The PTO request data
 * @return {Object} Result of the operation
 */
function submitPTORequest(ptoRequest) {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Try both possible sheet names
    let ptoRequestsSheet = ss.getSheetByName('PTO Requests');
    if (!ptoRequestsSheet) {
      // Try alternative names
      const possibleNames = ['PTO Requests', 'PTORequests', 'PTO_Requests', 'PTO'];
      for (const name of possibleNames) {
        ptoRequestsSheet = ss.getSheetByName(name);
        if (ptoRequestsSheet) break;
      }
      
      if (!ptoRequestsSheet) {
        Logger.log("PTO Requests sheet not found. Creating new sheet.");
        // Create the sheet if it doesn't exist
        ptoRequestsSheet = ss.insertSheet('PTO Requests');
        
        // Add headers based on your CSV structure
        ptoRequestsSheet.appendRow([
          "Request ID", 
          "Employee ID", 
          "Employee Name", 
          "Request Date", 
          "Start Date", 
          "End Date", 
          "Number of Hours Requested", 
          "Reason", 
          "Status", 
          "Manager Comments", 
          "Calendar Event ID (for Google Calendar sync)", 
          "Email Notification Status",
          "Employee Notes" // Added Employee Notes header
        ]);
      }
    }
    
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    
    // Generate request ID
    const today = new Date();
    const requestId = "PTO" + today.getTime().toString().slice(-8);
    
    // Get employee name, email, and manager email
    const employeeData = employeeSheet.getDataRange().getValues();
    let employeeName = "";
    let employeeEmail = ""; // Added employee email variable
    let managerEmail = "";
    let currentPTOBalance = 0;
    
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] == ptoRequest.employeeId) {
        employeeName = employeeData[i][1] + " " + employeeData[i][2]; // First + Last name
        employeeEmail = employeeData[i][4]; // Employee Email column E
        managerEmail = employeeData[i][6]; // Manager Email column
        currentPTOBalance = employeeData[i][13]; // Current PTO Balance column
        break;
      }
    }
    
    // Check if employee has enough PTO balance
    if (currentPTOBalance < ptoRequest.totalHours) {
      return { 
        success: false, 
        message: "Insufficient PTO balance. You have " + currentPTOBalance + 
                " hours available, but requested " + ptoRequest.totalHours + " hours." 
      };
    }
    
    // Format the request date
    const requestDate = Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd/yyyy");
    
    // Log what we're about to insert
    Logger.log("Inserting PTO request with ID: " + requestId);
    Logger.log("Employee ID: " + ptoRequest.employeeId);
    Logger.log("Employee Name: " + employeeName);
    Logger.log("Request Date: " + requestDate);
    Logger.log("Start Date: " + ptoRequest.startDate);
    Logger.log("End Date: " + ptoRequest.endDate);
    Logger.log("Hours: " + ptoRequest.totalHours);
    Logger.log("Type: " + ptoRequest.requestType);
    Logger.log("Notes: " + ptoRequest.notes); // Log notes
    
    // Create Google Calendar event
    let calendarEventId = "";
    try {
      const calendarId = "0c9bf77626ae7467b6e1fec960ed070161a164e4921905f3333282f003be419f@group.calendar.google.com";
      
      // Parse dates for calendar event - with timezone handling
      Logger.log("Parsing start date: " + ptoRequest.startDate);
      Logger.log("Parsing end date: " + ptoRequest.endDate);
      
      // Properly parse the date strings with timezone handling
      let startDate, endDate;
      
      // Function to parse date string to Date object with proper timezone handling
      function parseDate(dateString) {
        // For YYYY-MM-DD format
        if (dateString.match(/^\d{4}-\d{2}-\d{2}$/)) {
          const [year, month, day] = dateString.split('-').map(num => parseInt(num, 10));
          // Create date using year, month (0-indexed), day
          return new Date(year, month - 1, day);
        }
        // For MM/DD/YYYY format
        else if (dateString.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
          const [month, day, year] = dateString.split('/').map(num => parseInt(num, 10));
          return new Date(year, month - 1, day);
        }
        // If it's already a Date object
        else if (dateString instanceof Date) {
          return new Date(dateString);
        }
        // Default fallback
        else {
          // Try to parse as is, but this might have timezone issues
          const date = new Date(dateString);
          // Adjust for timezone if needed
          return new Date(date.getFullYear(), date.getMonth(), date.getDate());
        }
      }
      
      startDate = parseDate(ptoRequest.startDate);
      endDate = parseDate(ptoRequest.endDate);
      
      // Log the parsed dates for debugging
      Logger.log("Parsed start date: " + startDate);
      Logger.log("Parsed end date: " + endDate);
      
      // If end date is the same as start date, make it an all-day event for that day
      if (startDate.toDateString() === endDate.toDateString()) {
        // For all-day events, the end date should be the next day for Google Calendar
        endDate.setDate(endDate.getDate() + 1);
      } else {
        // For multi-day events, add one day to end date to make it inclusive
        endDate.setDate(endDate.getDate() + 1);
      }
      
      // Create event title and description
      const eventTitle = `PTO: ${employeeName} - ${ptoRequest.requestType}`;
      const eventDescription = `PTO Request Details:
Employee: ${employeeName}
Type: ${ptoRequest.requestType}
Hours: ${ptoRequest.totalHours}
Request ID: ${requestId}
Status: Pending
${ptoRequest.notes ? "Notes: " + ptoRequest.notes : ""}`;
      
      // Create calendar event with yellow color (color ID 5)
      const event = CalendarApp.getCalendarById(calendarId).createAllDayEvent(
        eventTitle,
        startDate,
        endDate,
        {
          description: eventDescription
        }
      );
      
      // Set the event color to yellow (color ID 5)
      event.setColor(CalendarApp.EventColor.YELLOW);
      
      // Get the event ID for reference
      calendarEventId = event.getId();
      Logger.log("Created calendar event with ID: " + calendarEventId);
      Logger.log("Event start date: " + event.getAllDayStartDate());
      Logger.log("Event end date: " + event.getAllDayEndDate());
      Logger.log("Event color set to yellow");
    } catch (calendarError) {
      Logger.log("Error creating calendar event: " + calendarError.toString());
      // Continue with request submission even if calendar event creation fails
    }
    
    // Add new PTO request - match the exact column structure from your CSV
    ptoRequestsSheet.appendRow([
      requestId,                  // Request ID
      ptoRequest.employeeId,      // Employee ID
      employeeName,               // Employee Name
      requestDate,                // Request Date
      ptoRequest.startDate,       // Start Date
      ptoRequest.endDate,         // End Date
      ptoRequest.totalHours,      // Number of Hours Requested
      ptoRequest.requestType,     // Reason
      "Pending",                  // Status
      "",                         // Manager Comments
      calendarEventId,            // Calendar Event ID
      "",                         // Email Notification Status
      ptoRequest.notes || ""      // Employee Notes - Column N
    ]);
    
    // Email notification section - updated to include additional recipient
    try {
      // Calendar link for employee email
      const calendarLink = "https://calendar.google.com/calendar/u/0?cid=MGM5YmY3NzYyNmFlNzQ2N2I2ZTFmZWM5NjBlZDA3MDE2MWExNjRlNDkyMTkwNWYzMzMzMjgyZjAwM2JlNDE5ZkBncm91cC5jYWxlbmRhci5nb29nbGUuY29t";
      
      // Prepare email content for manager
      const emailSubject = "New PTO Request: " + employeeName;
      const emailBody = "A new PTO request has been submitted:\n\n" +
                "Employee: " + employeeName + "\n" +
                "Request Type: " + ptoRequest.requestType + "\n" +
                "Dates: " + ptoRequest.startDate + " to " + ptoRequest.endDate + "\n" +
                "Total Hours: " + ptoRequest.totalHours + "\n" +
                (ptoRequest.notes ? "Notes: " + ptoRequest.notes + "\n\n" : "\n") +
                "Please log in to the manager dashboard to approve or deny this request.";
      
      // Additional email recipient
      const additionalEmail = "humboldtmachinemanager@gmail.com";
      
      // Send to manager if available
      if (managerEmail) {
        try {
          MailApp.sendEmail({
            to: managerEmail,
            subject: emailSubject,
            body: emailBody
          });
          Logger.log("Email sent to manager: " + managerEmail);
        } catch (emailError) {
          Logger.log("Error sending email to manager: " + emailError.toString());
          // Continue with other notifications even if this one fails
        }
      }
      
      // Always send to the additional email address
      try {
        MailApp.sendEmail({
          to: additionalEmail,
          subject: emailSubject,
          body: emailBody
        });
        Logger.log("Email sent to additional recipient: " + additionalEmail);
      } catch (emailError) {
        Logger.log("Error sending email to additional recipient: " + emailError.toString());
      }
      
// Send confirmation email to employee
if (employeeEmail) {
  try {
    // Prepare email content for employee with updated subject including dates
    const employeeEmailSubject = `Your PTO Request Confirmation: ${ptoRequest.startDate} to ${ptoRequest.endDate}`;
    const employeeEmailBody = "Your PTO request has been submitted successfully:\n\n" +
              "Request Type: " + ptoRequest.requestType + "\n" +
              "Dates: " + ptoRequest.startDate + " to " + ptoRequest.endDate + "\n" +
              "Total Hours: " + ptoRequest.totalHours + "\n" +
              (ptoRequest.notes ? "Notes: " + ptoRequest.notes + "\n\n" : "\n") +
              "Your request is now pending approval. You will be notified when your request is approved or denied.";
    
    MailApp.sendEmail({
      to: employeeEmail,
      subject: employeeEmailSubject,
      body: employeeEmailBody
    });
    Logger.log("Confirmation email sent to employee: " + employeeEmail);
  } catch (emailError) {
    Logger.log("Error sending confirmation email to employee: " + emailError.toString());
    // Continue even if employee email fails
  }
} else {
  Logger.log("No employee email found, skipping confirmation email");
}

      
    } catch (emailError) {
      Logger.log("Error in email notification process: " + emailError.toString());
      // Continue with request submission even if email fails
    }
    
    return { 
      success: true, 
      message: "PTO request submitted successfully", 
      requestId: requestId 
    };
  } catch (e) {
    Logger.log("Error submitting PTO request: " + e.toString());
    return { success: false, message: "Error submitting request: " + e.toString() };
  }
}



/**
 * Gets the employee's PTO balance including pending requests
 * @param {string} employeeId - The employee's ID
 * @return {Object} PTO balance information with pending hours
 */
function getEmployeePTOBalanceWithPending(employeeId) {
  try {
    // Get current PTO balance
    const balanceResult = getEmployeePTOBalance(employeeId);
    if (!balanceResult.success) {
      return balanceResult;
    }
    
    // Calculate pending PTO hours
    const ptoRequestsSheet = ss.getSheetByName('PTO Requests');
    if (!ptoRequestsSheet) {
      // If the sheet doesn't exist, return just the current balance
      return {
        success: true,
        balance: balanceResult.balance,
        pendingHours: 0,
        balanceAfterPending: balanceResult.balance
      };
    }
    
    const ptoRequestsData = ptoRequestsSheet.getDataRange().getValues();
    let pendingHours = 0;
    
    // Skip header row and find pending requests for this employee
    for (let i = 1; i < ptoRequestsData.length; i++) {
      // Check if this row is for the current employee and has "Pending" status
      if (ptoRequestsData[i][1] == employeeId && ptoRequestsData[i][8] === "Pending") {
        // Add the requested hours to pending total
        pendingHours += Number(ptoRequestsData[i][6]) || 0;
      }
    }
    
    // Calculate balance after pending requests
    const balanceAfterPending = balanceResult.balance - pendingHours;
    
    return {
      success: true,
      balance: balanceResult.balance,
      pendingHours: pendingHours,
      balanceAfterPending: balanceAfterPending
    };
  } catch (e) {
    Logger.log("Error getting PTO balance with pending: " + e.toString());
    return { success: false, message: "Error retrieving PTO balance" };
  }
}








// Profit Sharing Functions

/**
 * Calculates profit sharing eligibility for an employee based on their full time start date and schedule type
 * @param {string|number} employeeId - The employee's ID (optional - if not provided, will update all employees)
 * @return {Object} Result of the operation
 */
function calculateProfitSharingEligibility(employeeId = null) {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Get employee sheet
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      return { success: false, message: "Employee Master Data sheet not found" };
    }
    
    // Get pay periods sheet for determining eligibility dates
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      return { success: false, message: "Pay Periods sheet not found" };
    }
    
    // Get all employee data
    const employeeData = employeeSheet.getDataRange().getValues();
    
    // Find column indices
    const headers = employeeData[0];
    const idIdx = headers.indexOf('Employee ID');
    const hireDateIdx = headers.indexOf('Hire Date');
    const fullTimeStartDateIdx = headers.indexOf('Full Time Start Date'); // New column X
    const eligibilityIdx = headers.indexOf('Eligible for Profit Sharing');
    const eligibilityDateIdx = headers.indexOf('Date Eligible for Profit Sharing'); // Column W
    const statusIdx = headers.indexOf('Status');
    const scheduleTypeIdx = headers.indexOf('Schedule type');
    
    if (fullTimeStartDateIdx === -1) {
      return { success: false, message: "Full Time Start Date column not found in column X" };
    }
    
    if (eligibilityIdx === -1) {
      return { success: false, message: "Eligible for Profit Sharing column not found in column V" };
    }
    
    if (eligibilityDateIdx === -1) {
      return { success: false, message: "Date Eligible for Profit Sharing column not found in column W" };
    }
    
    if (scheduleTypeIdx === -1) {
      return { success: false, message: "Schedule type column not found" };
    }
    
    // Get pay periods data
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    
    // Sort pay periods by start date
    const payPeriods = [];
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][0] && payPeriodsData[i][2] && payPeriodsData[i][4]) { // Has ID, start date, and end date
        payPeriods.push({
          id: payPeriodsData[i][0],
          name: payPeriodsData[i][1],
          startDate: new Date(payPeriodsData[i][2]),
          endDate: new Date(payPeriodsData[i][4])
        });
      }
    }
    
    // Sort pay periods by end date
    payPeriods.sort((a, b) => a.endDate - b.endDate);
    
    // Group pay periods by month (using end date)
    const payPeriodsByMonth = {};
    payPeriods.forEach(period => {
      const year = period.endDate.getFullYear();
      const month = period.endDate.getMonth();
      const key = `${year}-${month}`;
      
      if (!payPeriodsByMonth[key]) {
        payPeriodsByMonth[key] = [];
      }
      
      payPeriodsByMonth[key].push(period);
    });
    
    // Process employees
    const results = [];
    const today = new Date();
    
    for (let i = 1; i < employeeData.length; i++) {
      // Skip if we're looking for a specific employee and this isn't it
      if (employeeId !== null && employeeData[i][idIdx] != employeeId) {
        continue;
      }
      
      // Skip inactive employees
      if (statusIdx !== -1 && employeeData[i][statusIdx] !== "Active") {
        continue;
      }
      
      // Check if employee is Full Time - if not, they're not eligible
      const scheduleType = employeeData[i][scheduleTypeIdx];
      const isFullTime = scheduleType === "Full Time";
      
      // Get full time start date (use hire date as fallback if not available)
      const fullTimeStartDate = employeeData[i][fullTimeStartDateIdx];
      const hireDate = employeeData[i][hireDateIdx];
      
      // Use Full Time Start Date if available, otherwise fall back to Hire Date
      const startDate = (fullTimeStartDate && fullTimeStartDate instanceof Date) ? fullTimeStartDate : hireDate;
      
      if (!startDate || !(startDate instanceof Date)) {
        results.push({
          employeeId: employeeData[i][idIdx],
          success: false,
          message: "Invalid start date (neither Full Time Start Date nor Hire Date is valid)"
        });
        continue;
      }
      
      // Calculate the date 4 months after start date
      const fourMonthsAfterStart = new Date(startDate);
      fourMonthsAfterStart.setMonth(fourMonthsAfterStart.getMonth() + 4);
      
      // Calculate the month after that (5th month)
      const fifthMonth = new Date(fourMonthsAfterStart);
      fifthMonth.setMonth(fifthMonth.getMonth() + 1);
      fifthMonth.setDate(1); // First day of the month
      
      // Find the second pay period in the fifth month
      const fifthMonthKey = `${fifthMonth.getFullYear()}-${fifthMonth.getMonth()}`;
      let eligibilityDate = null;
      
      if (payPeriodsByMonth[fifthMonthKey] && payPeriodsByMonth[fifthMonthKey].length >= 2) {
        // Get the second pay period of the month (using end date)
        const secondPayPeriod = payPeriodsByMonth[fifthMonthKey][1];
        eligibilityDate = secondPayPeriod.endDate;
      } else {
        // If we can't find the exact pay period, estimate it as the last day of the month
        eligibilityDate = new Date(fifthMonth.getFullYear(), fifthMonth.getMonth() + 1, 0);
      }
      
      // Determine if employee is eligible now - must be both past eligibility date AND full time
      const isEligible = (today >= eligibilityDate) && isFullTime;
      
      // Format the eligibility date for display
      const formattedEligibilityDate = Utilities.formatDate(eligibilityDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
      
      // Update the eligibility status in column V
      employeeSheet.getRange(i + 1, eligibilityIdx + 1).setValue(isEligible ? "Yes" : "No");
      
      // Add a note with the eligibility date for reference and schedule type info if not eligible
      let noteText = `Eligible from ${formattedEligibilityDate}`;
      if (!isFullTime) {
        noteText += ` (Not eligible: Schedule type is "${scheduleType}" - must be "Full Time")`;
      }
      employeeSheet.getRange(i + 1, eligibilityIdx + 1).setNote(noteText);
      
      // Update the eligibility date in column W, but only if they meet the schedule type requirement
      if (isFullTime) {
        employeeSheet.getRange(i + 1, eligibilityDateIdx + 1).setValue(eligibilityDate);
      } else {
        // Clear the eligibility date if they're not full time
        employeeSheet.getRange(i + 1, eligibilityDateIdx + 1).clearContent();
      }
      
      // Format the start date for display
      const formattedStartDate = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
      
      // Add to results
      results.push({
        employeeId: employeeData[i][idIdx],
        startDate: formattedStartDate,
        startDateType: fullTimeStartDate ? "Full Time Start Date" : "Hire Date (fallback)",
        fourMonthsAfterStart: Utilities.formatDate(fourMonthsAfterStart, Session.getScriptTimeZone(), "MM/dd/yyyy"),
        eligibilityDate: formattedEligibilityDate,
        scheduleType: scheduleType || "Not Set",
        isFullTime: isFullTime,
        isEligible: isEligible,
        success: true
      });
      
      // If we were looking for a specific employee and found them, we can stop
      if (employeeId !== null) {
        break;
      }
    }
    
    return {
      success: true,
      message: `Processed ${results.length} employees`,
      results: results
    };
    
  } catch (error) {
    Logger.log("Error in calculateProfitSharingEligibility: " + error.toString());
    return { success: false, message: error.toString() };
  }
}







/**
 * Gets profit sharing eligibility for a specific employee
 * @param {string|number} employeeId - The employee's ID
 * @return {Object} Eligibility information
 */
function getEmployeeProfitSharingEligibility(employeeId) {
  const result = calculateProfitSharingEligibility(employeeId);
  
  if (result.success && result.results && result.results.length > 0) {
    return {
      success: true,
      employeeId: result.results[0].employeeId,
      isEligible: result.results[0].isEligible,
      eligibilityDate: result.results[0].eligibilityDate
    };
  }
  
  return { 
    success: false, 
    message: result.message || "Failed to determine eligibility" 
  };
}

/**
 * Updates profit sharing eligibility for all employees with detailed logging
 * @return {Object} Result of the operation with logging information
 */
function updateAllEmployeeProfitSharingEligibility() {
  // Start logging
  Logger.log("Starting profit sharing eligibility update for all employees");
  const startTime = new Date();
  
  // Track changes
  const changes = {
    total: 0,
    becameEligible: [],
    becameIneligible: [],
    unchanged: 0,
    errors: []
  };
  
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      Logger.log("Failed to initialize spreadsheet");
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Get employee sheet to capture current state before changes
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      Logger.log("Employee Master Data sheet not found");
      return { success: false, message: "Employee Master Data sheet not found" };
    }
    
    // Get current eligibility status before making changes
    const currentData = employeeSheet.getDataRange().getValues();
    const headers = currentData[0];
    const idIdx = headers.indexOf('Employee ID');
    const eligibilityIdx = headers.indexOf('Eligible for Profit Sharing');
    
    // Create a map of current eligibility status
    const currentEligibility = {};
    for (let i = 1; i < currentData.length; i++) {
      if (currentData[i][idIdx]) {
        currentEligibility[currentData[i][idIdx]] = currentData[i][eligibilityIdx] === "Yes";
      }
    }
    
    // Run the calculation
    Logger.log("Calling calculateProfitSharingEligibility function");
    const result = calculateProfitSharingEligibility();
    
    // If the calculation was successful, analyze changes
    if (result.success && result.results) {
      Logger.log(`Successfully processed ${result.results.length} employees`);
      
      // Analyze each employee result to track changes
      result.results.forEach(employee => {
        const employeeId = employee.employeeId;
        const wasEligible = currentEligibility[employeeId] || false;
        const isNowEligible = employee.isEligible;
        
        if (wasEligible !== isNowEligible) {
          changes.total++;
          
          if (isNowEligible) {
            changes.becameEligible.push({
              employeeId: employeeId,
              eligibleFrom: employee.eligibilityDate,
              reason: "Passed eligibility date and is Full Time"
            });
            Logger.log(`Employee ${employeeId} became eligible for profit sharing`);
          } else {
            changes.becameIneligible.push({
              employeeId: employeeId,
              reason: employee.isFullTime ? 
                "No longer meets date criteria" : 
                `Schedule type is "${employee.scheduleType}" (not Full Time)`
            });
            Logger.log(`Employee ${employeeId} became ineligible for profit sharing: ${employee.isFullTime ? "No longer meets date criteria" : "Not Full Time"}`);
          }
        } else {
          changes.unchanged++;
        }
      });
      
      // Log any errors from the results
      result.results.forEach(employee => {
        if (!employee.success) {
          changes.errors.push({
            employeeId: employee.employeeId,
            error: employee.message
          });
          Logger.log(`Error processing employee ${employee.employeeId}: ${employee.message}`);
        }
      });
      
      // Calculate execution time
      const endTime = new Date();
      const executionTime = (endTime - startTime) / 1000; // in seconds
      
      // Create a detailed log entry in the spreadsheet
      logUpdateToSpreadsheet({
        timestamp: new Date(),
        totalEmployees: result.results.length,
        changesCount: changes.total,
        becameEligibleCount: changes.becameEligible.length,
        becameIneligibleCount: changes.becameIneligible.length,
        unchangedCount: changes.unchanged,
        errorsCount: changes.errors.length,
        executionTime: executionTime,
        details: JSON.stringify(changes)
      });
      
      // Return enhanced result with change information
      return {
        success: true,
        message: `Processed ${result.results.length} employees. ${changes.total} changes made (${changes.becameEligible.length} became eligible, ${changes.becameIneligible.length} became ineligible).`,
        changes: changes,
        executionTime: `${executionTime.toFixed(2)} seconds`,
        timestamp: new Date().toISOString()
      };
    } else {
      // Log the error
      Logger.log(`Error in calculation: ${result.message}`);
      return result;
    }
  } catch (error) {
    const errorMessage = `Error in updateAllEmployeeProfitSharingEligibility: ${error.toString()}`;
    Logger.log(errorMessage);
    return { 
      success: false, 
      message: errorMessage,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * Logs the update operation to a dedicated log sheet in the spreadsheet
 * @param {Object} logData - Data to log
 */
function logUpdateToSpreadsheet(logData) {
  try {
    // Get or create log sheet
    let logSheet = ss.getSheetByName('Profit Sharing Eligibility Logs');
    if (!logSheet) {
      logSheet = ss.insertSheet('Profit Sharing Eligibility Logs');
      // Add headers
      logSheet.appendRow([
        'Timestamp', 
        'Total Employees', 
        'Changes', 
        'Became Eligible', 
        'Became Ineligible', 
        'Unchanged',
        'Errors',
        'Execution Time (s)',
        'Details'
      ]);
      
      // Format the header row
      logSheet.getRange(1, 1, 1, 9).setFontWeight('bold');
    }
    
    // Format timestamp
    const formattedTimestamp = Utilities.formatDate(
      logData.timestamp, 
      Session.getScriptTimeZone(), 
      "MM/dd/yyyy HH:mm:ss"
    );
    
    // Add log entry
    logSheet.appendRow([
      formattedTimestamp,
      logData.totalEmployees,
      logData.changesCount,
      logData.becameEligibleCount,
      logData.becameIneligibleCount,
      logData.unchangedCount,
      logData.errorsCount,
      logData.executionTime.toFixed(2),
      logData.details
    ]);
    
    // Auto-resize columns for better readability
    logSheet.autoResizeColumns(1, 9);
    
    Logger.log("Update logged to spreadsheet successfully");
  } catch (error) {
    Logger.log(`Error logging to spreadsheet: ${error.toString()}`);
  }
}



function createDailyTrigger() {
  ScriptApp.newTrigger('updateAllEmployeeProfitSharingEligibility')
    .timeBased()
    .everyDays(1)
    .atHour(1) // Run at 1 AM
    .create();
}

/**
 * Updates profit sharing information for a pay period
 * @param {string} payPeriodId - The ID of the pay period
 * @param {number} profitSharingAmount - Total profit sharing amount to distribute
 * @return {Object} Result of the operation
 */
function updateProfitSharingInfo(payPeriodId, profitSharingAmount) {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Get pay periods sheet
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      return { success: false, message: "Pay Periods sheet not found" };
    }
    
    // Find the pay period row
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    let payPeriodRowIndex = -1;
    
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][0] == payPeriodId) {
        payPeriodRowIndex = i + 1; // +1 because sheet rows are 1-indexed
        break;
      }
    }
    
    if (payPeriodRowIndex === -1) {
      return { success: false, message: "Pay period not found" };
    }
    
    // Get eligible employees for profit sharing
    const eligibleEmployees = getEligibleEmployeesForProfitSharing();
    
    // Calculate profit sharing details
    const numRecipients = eligibleEmployees.length;
    let amountPerRecipient = 0;
    let recipientNames = "";
    
    if (numRecipients > 0 && profitSharingAmount > 0) {
      amountPerRecipient = profitSharingAmount / numRecipients;
      recipientNames = eligibleEmployees.map(emp => `${emp.firstName} ${emp.lastName}`).join(", ");
    }
    
    // Update the pay period with profit sharing information
    payPeriodsSheet.getRange(payPeriodRowIndex, 9).setValue(profitSharingAmount); // Column I: Total Profit Sharing Amount
    payPeriodsSheet.getRange(payPeriodRowIndex, 10).setValue(numRecipients); // Column J: # of Profit Sharing Recipients
    payPeriodsSheet.getRange(payPeriodRowIndex, 11).setValue(amountPerRecipient); // Column K: Amount per Recipient
    payPeriodsSheet.getRange(payPeriodRowIndex, 12).setValue(recipientNames); // Column L: Names of Recipients
    
    return {
      success: true,
      message: "Profit sharing information updated successfully",
      numRecipients: numRecipients,
      amountPerRecipient: amountPerRecipient,
      totalAmount: profitSharingAmount
    };
    
  } catch (error) {
    Logger.log("Error in updateProfitSharingInfo: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets employees eligible for profit sharing
 * @return {Array} Array of eligible employees
 */
function getEligibleEmployeesForProfitSharing() {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return [];
    }
    
    // Get employee sheet
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      return [];
    }
    
    const employeeData = employeeSheet.getDataRange().getValues();
    const eligibleEmployees = [];
    
    // Find the column index for "Eligible for Profit Sharing" (should be column V)
    const headers = employeeData[0];
    const eligibilityIdx = headers.indexOf('Eligible for Profit Sharing');
    const statusIdx = headers.indexOf('Status');
    
    if (eligibilityIdx === -1) {
      Logger.log("Eligible for Profit Sharing column not found");
      return [];
    }
    
    // Find eligible active employees
    for (let i = 1; i < employeeData.length; i++) {
      // Check if employee is active and eligible for profit sharing
      if (
        employeeData[i][statusIdx] === 'Active' && 
        employeeData[i][eligibilityIdx] === 'Yes'
      ) {
        eligibleEmployees.push({
          employeeId: employeeData[i][0],
          firstName: employeeData[i][1],
          lastName: employeeData[i][2]
        });
      }
    }
    
    return eligibleEmployees;
    
  } catch (error) {
    Logger.log("Error in getEligibleEmployeesForProfitSharing: " + error.toString());
    return [];
  }
}

/**
 * Gets profit sharing information for a pay period
 * @param {string} payPeriodId - The ID of the pay period
 * @return {Object} Profit sharing information
 */
function getProfitSharingInfo(payPeriodId) {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Get pay periods sheet
    const payPeriodsSheet = ss.getSheetByName('Pay Periods');
    if (!payPeriodsSheet) {
      return { success: false, message: "Pay Periods sheet not found" };
    }
    
    // Find the pay period row
    const payPeriodsData = payPeriodsSheet.getDataRange().getValues();
    
    for (let i = 1; i < payPeriodsData.length; i++) {
      if (payPeriodsData[i][0] == payPeriodId) {
        return {
          success: true,
          totalAmount: payPeriodsData[i][8] || 0, // Column I: Total Profit Sharing Amount
          numRecipients: payPeriodsData[i][9] || 0, // Column J: # of Profit Sharing Recipients
          amountPerRecipient: payPeriodsData[i][10] || 0, // Column K: Amount per Recipient
          recipientNames: payPeriodsData[i][11] || "" // Column L: Names of Recipients
        };
      }
    }
    
    return { success: false, message: "Pay period not found" };
    
  } catch (error) {
    Logger.log("Error in getProfitSharingInfo: " + error.toString());
    return { success: false, message: error.toString() };
  }
}
