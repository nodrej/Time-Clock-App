/******************************************
 * PTO MANAGEMENT FUNCTIONS
 ******************************************/

/**
 * Gets all employees for the PTO management dashboard
 * @return {Object} Object with success indicator and employee data
 */
function getEmployeesForPTOManagement() {
  try {
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      return { success: false, message: 'Employee data sheet not found' };
    }

    const data = employeeSheet.getDataRange().getValues();
    const headers = data[0];
    
    const idIndex = headers.indexOf('Employee ID');
    const firstNameIndex = headers.indexOf('First Name');
    const lastNameIndex = headers.indexOf('Last Name');
    const departmentIndex = headers.indexOf('Department');
    const hireDateIndex = headers.indexOf('Hire Date');
    const emailIndex = headers.indexOf('Email');
    
    if (idIndex < 0 || firstNameIndex < 0 || lastNameIndex < 0) {
      return { success: false, message: 'Required employee columns not found' };
    }

    const employees = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      employees.push({
        employeeId: row[idIndex].toString(),
        firstName: row[firstNameIndex],
        lastName: row[lastNameIndex],
        department: departmentIndex >= 0 ? row[departmentIndex] : 'Unassigned',
        hireDate: hireDateIndex >= 0 ? formatDate(row[hireDateIndex]) : null,
        email: emailIndex >= 0 ? row[emailIndex] : '',
        fullName: `${row[firstNameIndex]} ${row[lastNameIndex]}`
      });
    }

    return { success: true, data: employees };
    
  } catch (error) {
    console.error('Error in getEmployeesForPTOManagement:', error);
    return { success: false, message: error.toString() };
  }
}

/**
* Gets all PTO requests
* @return {Object} Object with success indicator and PTO request data
*/
function getPTORequests() {
  try {
    // Check if PTO requests sheet exists, create if it doesn't
    let ptoSheet = ss.getSheetByName('PTO Requests');
    if (!ptoSheet) {
      initializePTORequestsSheet();
      ptoSheet = ss.getSheetByName('PTO Requests');
      
      // Return empty data since sheet was just created
      return { 
        success: true, 
        data: [],
        message: "PTO Requests sheet was just created. No data available yet." 
      };
    }

    const data = ptoSheet.getDataRange().getValues();
    const headers = data[0];
    
    const requestIdIndex = headers.indexOf('Request ID');
    const employeeIdIndex = headers.indexOf('Employee ID');
    const employeeNameIndex = headers.indexOf('Employee Name');
    const requestDateIndex = headers.indexOf('Request Date');
    const startDateIndex = headers.indexOf('Start Date');
    const endDateIndex = headers.indexOf('End Date');
    const hoursIndex = headers.indexOf('Hours Requested');
    const reasonIndex = headers.indexOf('Reason');
    const statusIndex = headers.indexOf('Status');
    const managerCommentsIndex = headers.indexOf('Manager Comments');
    
    if (requestIdIndex < 0 || employeeIdIndex < 0 || startDateIndex < 0 || endDateIndex < 0 || hoursIndex < 0 || statusIndex < 0) {
      return { success: false, message: 'Required PTO request columns not found' };
    }

    const requests = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      requests.push({
        requestId: row[requestIdIndex].toString(),
        employeeId: row[employeeIdIndex].toString(),
        employeeName: employeeNameIndex >= 0 ? row[employeeNameIndex] : 'Unknown',
        requestDate: requestDateIndex >= 0 ? formatDate(row[requestDateIndex]) : null,
        startDate: formatDate(row[startDateIndex]),
        endDate: formatDate(row[endDateIndex]),
        hours: parseFloat(row[hoursIndex]),
        reason: reasonIndex >= 0 ? row[reasonIndex] : '',
        status: row[statusIndex],
        managerComments: managerCommentsIndex >= 0 ? row[managerCommentsIndex] : '',
        rowIndex: i + 1  // Store row index for updates
      });
    }

    return { success: true, data: requests };
    
  } catch (error) {
    console.error('Error in getPTORequests:', error);
    return { success: false, message: error.toString() };
  }
}

/**
* Gets all PTO requests
* @return {Object} Object with success indicator and PTO request data
*/
function getPTORequests() {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Check if sheet exists, create if needed
    let ptoSheet = ss.getSheetByName('PTO Requests');
    if (!ptoSheet) {
      const result = initializePTORequestsSheet();
      if (!result.success) {
        return result;
      }
      ptoSheet = ss.getSheetByName('PTO Requests');
      return { success: true, data: [], message: "PTO Requests sheet was just created" };
    }

    // Get the data - using direct column indexes matching your existing code style
    const data = ptoSheet.getDataRange().getValues();
    
    // If the sheet is empty or only has headers
    if (data.length <= 1) {
      return { success: true, data: [] };
    }
    
    const requests = [];
    
    // Process rows skipping header
    for (let i = 1; i < data.length; i++) {
      if (!data[i][0]) continue; // Skip empty rows
      
      requests.push({
        requestId: data[i][0].toString(),
        employeeId: data[i][1].toString(),
        employeeName: data[i][2] || 'Unknown',
        requestDate: formatDate(data[i][3]),
        startDate: formatDate(data[i][4]),
        endDate: formatDate(data[i][5]),
        hours: parseFloat(data[i][6]),
        reason: data[i][7] || '',
        status: data[i][8] || 'Pending',
        managerComments: data[i][9] || '',
        rowIndex: i + 1  // Store row index for updates
      });
    }

    return { success: true, data: requests };
    
  } catch (error) {
    Logger.log("Error in getPTORequests: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets PTO balances for all employees
 * @return {Object} Object with success indicator and PTO balance data
 */
function getEmployeePTOBalances() {
  try {
    // Make sure spreadsheet is initialized
    if (!initSpreadsheet()) {
      return { success: false, message: "Failed to initialize spreadsheet" };
    }
    
    // Get employees
    const employeeSheet = ss.getSheetByName('Employee Master Data');
    if (!employeeSheet) {
      return { success: false, message: 'Employee data sheet not found' };
    }
    
    const employeeData = employeeSheet.getDataRange().getValues();
    
    // Create transactions sheet if needed
    let transactionsSheet = ss.getSheetByName('PTO Transactions');
    if (!transactionsSheet) {
      const result = initializePTOTransactionsSheet();
      if (!result.success) {
        return result;
      }
      transactionsSheet = ss.getSheetByName('PTO Transactions');
      
      // Initialize employee balances
      initializeEmployeePTOBalances();
    }
    
    // Get transaction data
    const transData = transactionsSheet.getDataRange().getValues();
    
    // Get PTO requests for pending calculation
    const requestsResult = getPTORequests();
    if (!requestsResult.success) {
      return requestsResult;
    }
    const pendingRequests = requestsResult.data.filter(req => req.status === 'Pending');
    
    // Build employee lookup map with index 0 as ID, 1 as first name, 2 as last name, 3 as department
    const employees = [];
    for (let i = 1; i < employeeData.length; i++) {
      if (employeeData[i][0] && employeeData[i][8] === "Active") { // Check for ID and active status
        employees.push({
          employeeId: employeeData[i][0].toString(),
          firstName: employeeData[i][1],
          lastName: employeeData[i][2],
          department: employeeData[i][3] || 'Unassigned',
          hireDate: formatDate(employeeData[i][7]) // Hire date
        });
      }
    }
    
    // Calculate balances for each employee
    const balances = employees.map(employee => {
      // Find all transactions for this employee - column 1 is Employee ID
      const employeeTransactions = [];
      for (let i = 1; i < transData.length; i++) {
        if (String(transData[i][1]) === employee.employeeId) {
          employeeTransactions.push({
            date: transData[i][2], // Date column
            amount: parseFloat(transData[i][3]) || 0, // Amount column
            type: transData[i][4] || '' // Type column
          });
        }
      }
      
      // Calculate current balance
      let ptoBalance = 0;
      employeeTransactions.forEach(transaction => {
        ptoBalance += transaction.amount;
      });
      
      // Calculate YTD used (sum of negative transactions in current year)
      const currentYear = new Date().getFullYear();
      let ytdUsed = 0;
      employeeTransactions.forEach(transaction => {
        const transYear = transaction.date instanceof Date ? 
          transaction.date.getFullYear() : 
          new Date(transaction.date).getFullYear();
        
        if (transYear === currentYear && transaction.amount < 0) {
          ytdUsed -= transaction.amount; // Convert to positive value
        }
      });
      
      // Calculate pending hours
      let pendingHours = 0;
      pendingRequests.forEach(req => {
        if (req.employeeId === employee.employeeId) {
          pendingHours += req.hours;
        }
      });
      
      // Default annual accrual rate of 80 hours
      let accrualRate = 80;
      
      // Get accrual settings if they exist
      const accrualSettings = getPTOAccrualSettings();
      if (accrualSettings.success && accrualSettings.data) {
        accrualRate = accrualSettings.data.defaultAccrualRate;
        
        // Apply department-specific rules
        if (accrualSettings.data.departmentRules && 
            accrualSettings.data.departmentRules[employee.department] && 
            accrualSettings.data.departmentRules[employee.department].accrualRate) {
          accrualRate = accrualSettings.data.departmentRules[employee.department].accrualRate;
        }
        
        // Apply tenure-based rules if hire date is available
        if (employee.hireDate && accrualSettings.data.tenureRules && accrualSettings.data.tenureRules.length > 0) {
          const hireDate = new Date(employee.hireDate);
          const today = new Date();
          const yearsOfService = (today - hireDate) / (1000 * 60 * 60 * 24 * 365.25);
          
          // Sort rules by years descending
          const sortedRules = [...accrualSettings.data.tenureRules].sort((a, b) => b.years - a.years);
          
          // Find the first rule that applies
          for (const rule of sortedRules) {
            if (yearsOfService >= rule.years) {
              accrualRate = rule.accrualRate;
              break;
            }
          }
        }
      }
      
      // Return balance information
      return {
        employeeId: employee.employeeId,
        firstName: employee.firstName,
        lastName: employee.lastName,
        department: employee.department,
        hireDate: employee.hireDate,
        ptoBalance: ptoBalance.toFixed(1),
        ytdUsed: ytdUsed.toFixed(1),
        pendingHours: pendingHours.toFixed(1),
        accrualRate: accrualRate.toFixed(1)
      };
    });
    
    return { success: true, data: balances };
    
  } catch (error) {
    Logger.log("Error in getEmployeePTOBalances: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets PTO accrual settings
 * @return {Object} Object with success indicator and accrual settings
 */
function getPTOAccrualSettings() {
  try {
    // Make sure sheets exist
    let settingsSheet = ss.getSheetByName('PTO Settings');
    if (!settingsSheet) {
      const result = initializePTOSettingsSheet();
      if (!result.success) {
        return result;
      }
      settingsSheet = ss.getSheetByName('PTO Settings');
    }
    
    // Get settings data
    const settingsData = settingsSheet.getDataRange().getValues();
    const settings = {};
    
    // Parse basic settings from rows 1+ (skipping header)
    for (let i = 1; i < settingsData.length; i++) {
      if (settingsData[i].length >= 2) {
        const key = settingsData[i][0];
        const value = settingsData[i][1];
        
        switch(key) {
          case 'Accrual Frequency':
            settings.accrualFrequency = value;
            break;
          case 'Default Accrual Rate':
            settings.defaultAccrualRate = parseFloat(value);
            break;
          case 'Max Carryover':
            settings.maxCarryover = parseFloat(value);
            break;
          case 'Reset Date':
            settings.resetDate = value;
            break;
          case 'Probation Period':
            settings.probationPeriod = parseInt(value);
            break;
          case 'Auto Approve Under 8 Hours':
            settings.autoApproveUnder8Hours = (value === 'true' || value === true);
            break;
        }
      }
    }
    
    // Read tenure rules
    const tenureRulesSheet = ss.getSheetByName('PTO Tenure Rules');
    let tenureRules = [];
    
    if (tenureRulesSheet) {
      const rulesData = tenureRulesSheet.getDataRange().getValues();
      
      // Skip header row
      for (let i = 1; i < rulesData.length; i++) {
        if (rulesData[i][0] !== undefined) {
          tenureRules.push({
            years: parseInt(rulesData[i][0]),
            accrualRate: parseFloat(rulesData[i][1]),
            maxBalance: parseFloat(rulesData[i][2]) || 0
          });
        }
      }
    } else {
      // Create default tenure rules
      tenureRules = [
        { years: 0, accrualRate: 80, maxBalance: 120 },
        { years: 5, accrualRate: 120, maxBalance: 160 },
        { years: 10, accrualRate: 160, maxBalance: 200 }
      ];
    }
    
    // Read department rules
    const deptRulesSheet = ss.getSheetByName('PTO Department Rules');
    let departmentRules = {};
    
    if (deptRulesSheet) {
      const deptData = deptRulesSheet.getDataRange().getValues();
      
      // Skip header row
      for (let i = 1; i < deptData.length; i++) {
        const deptName = deptData[i][0];
        if (deptName) {
          departmentRules[deptName] = {
            accrualRate: parseFloat(deptData[i][1]) || null,
            maxBalance: parseFloat(deptData[i][2]) || null
          };
        }
      }
    }
    
    // Combine all settings
    settings.tenureRules = tenureRules;
    settings.departmentRules = departmentRules;
    
    return { success: true, data: settings };
    
  } catch (error) {
    Logger.log("Error in getPTOAccrualSettings: " + error.toString());
    return { success: false, message: error.toString() };
  }
}

/**
 * Updates a PTO request status
 * @param {string} requestId - The request ID to update
 * @param {string} status - The new status (Approved, Denied)
 * @param {string} comments - Manager comments
 * @return {Object} Object with success indicator
 */
function updatePTORequest(requestId, status, comments) {
  try {
    const ptoSheet = ss.getSheetByName('PTO Requests');
    if (!ptoSheet) {
      return { success: false, message: 'PTO Requests sheet not found' };
    }
    
    // Get all requests to find the right one
    const requestsResult = getPTORequests();
    if (!requestsResult.success) {
      return requestsResult;
    }
    
    // Find the request
    const request = requestsResult.data.find(req => req.requestId === requestId);
    if (!request) {
      return { success: false, message: 'Request not found' };
    }
    
    // Update the status and comments
    const data = ptoSheet.getDataRange().getValues();
    const headers = data[0];
    const statusIndex = headers.indexOf('Status');
    const commentsIndex = headers.indexOf('Manager Comments');
    
    if (statusIndex < 0) {
      return { success: false, message: 'Status column not found' };
    }
    
    // Update the status
    ptoSheet.getRange(request.rowIndex, statusIndex + 1).setValue(status);
    
    // Update comments if the column exists
    if (commentsIndex >= 0) {
      ptoSheet.getRange(request.rowIndex, commentsIndex + 1).setValue(comments);
    }
    
    // If approved, deduct hours from PTO balance
    if (status === 'Approved') {
      addPTOTransaction(
        request.employeeId, 
        new Date(), 
        -request.hours, 
        'Used', 
        `PTO Request ${requestId} approved`
      );
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error in updatePTORequest:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Updates multiple PTO requests at once
 * @param {Array} requestIds - Array of request IDs
 * @param {string} action - 'approve' or 'deny'
 * @param {string} comments - Manager comments for all requests
 * @return {Object} Object with success indicator
 */
function bulkUpdatePTORequests(requestIds, action, comments) {
  try {
    if (!requestIds || !Array.isArray(requestIds) || requestIds.length === 0) {
      return { success: false, message: 'No request IDs provided' };
    }
    
    let allSuccess = true;
    const errors = [];
    
    // Process each request
    requestIds.forEach(requestId => {
      const status = action === 'approve' ? 'Approved' : 'Denied';
      const result = updatePTORequest(requestId, status, comments);
      
      if (!result.success) {
        allSuccess = false;
        errors.push(`Request ${requestId}: ${result.message}`);
      }
    });
    
    if (allSuccess) {
      return { success: true };
    } else {
      return { 
        success: false, 
        message: `Some requests could not be updated: ${errors.join('; ')}` 
      };
    }
    
  } catch (error) {
    console.error('Error in bulkUpdatePTORequests:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Adjusts an employee's PTO balance
 * @param {string} employeeId - The employee ID
 * @param {string} adjustType - 'add', 'subtract', or 'set'
 * @param {number} hours - Hours to adjust by
 * @param {string} reason - Reason for adjustment
 * @return {Object} Object with success indicator
 */
function adjustPTOBalance(employeeId, adjustType, hours, reason) {
  try {
    const hoursNum = parseFloat(hours);
    if (isNaN(hoursNum)) {
      return { success: false, message: 'Invalid hours value' };
    }
    
    let transactionHours = 0;
    let transactionType = '';
    
    switch (adjustType) {
      case 'add':
        transactionHours = hoursNum;
        transactionType = 'Manual Adjustment (Add)';
        break;
      case 'subtract':
        transactionHours = -hoursNum;
        transactionType = 'Manual Adjustment (Subtract)';
        break;
      case 'set':
        // Get current balance to calculate difference
        const balanceResult = getEmployeePTOBalance(employeeId);
        if (!balanceResult.success) {
          return balanceResult;
        }
        
        const currentBalance = parseFloat(balanceResult.data.currentBalance);
        transactionHours = hoursNum - currentBalance;
        transactionType = 'Manual Adjustment (Set)';
        break;
      default:
        return { success: false, message: 'Invalid adjustment type' };
    }
    
    // Add the transaction
    return addPTOTransaction(
      employeeId,
      new Date(),
      transactionHours,
      transactionType,
      reason
    );
    
  } catch (error) {
    console.error('Error in adjustPTOBalance:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets a specific employee's PTO balance
 * @param {string} employeeId - The employee ID
 * @return {Object} Object with success indicator and balance data
 */
function getEmployeePTOBalance(employeeId) {
  try {
    // Get balances for all employees
    const balancesResult = getEmployeePTOBalances();
    if (!balancesResult.success) {
      return balancesResult;
    }
    
    // Find the employee's balance
    const balance = balancesResult.data.find(b => b.employeeId === employeeId);
    if (!balance) {
      return { success: false, message: 'Employee not found' };
    }
    
    return {
      success: true,
      data: {
        employeeId: balance.employeeId,
        employeeName: `${balance.firstName} ${balance.lastName}`,
        department: balance.department,
        currentBalance: balance.ptoBalance,
        ytdUsed: balance.ytdUsed,
        pendingHours: balance.pendingHours,
        accrualRate: balance.accrualRate
      }
    };
    
  } catch (error) {
    console.error('Error in getEmployeePTOBalance:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Gets an employee's PTO transaction history
 * @param {string} employeeId - The employee ID
 * @return {Object} Object with success indicator and transaction history
 */
function getEmployeePTOHistory(employeeId) {
  try {
    const transactionsSheet = ss.getSheetByName('PTO Transactions');
    if (!transactionsSheet) {
      return { success: false, message: 'PTO Transactions sheet not found' };
    }
    
    const data = transactionsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const idIndex = headers.indexOf('Transaction ID');
    const empIdIndex = headers.indexOf('Employee ID');
    const dateIndex = headers.indexOf('Date');
    const amountIndex = headers.indexOf('Amount (Hours)');
    const typeIndex = headers.indexOf('Type');
    const notesIndex = headers.indexOf('Notes');
    const balanceIndex = headers.indexOf('Balance After');
    
    if (empIdIndex < 0 || dateIndex < 0 || amountIndex < 0) {
      return { success: false, message: 'Required transaction columns not found' };
    }
    
    const transactions = [];
    let runningBalance = 0;
    
    // Collect all transactions for the employee
    for (let i = 1; i < data.length; i++) {
      if (data[i][empIdIndex].toString() === employeeId.toString()) {
        const amount = parseFloat(data[i][amountIndex]);
        runningBalance += amount;
        
        transactions.push({
          id: idIndex >= 0 ? data[i][idIndex] : `TRANS${i}`,
          date: formatDate(data[i][dateIndex]),
          hours: amount,
          type: typeIndex >= 0 ? data[i][typeIndex] : '',
          notes: notesIndex >= 0 ? data[i][notesIndex] : '',
          balanceAfter: balanceIndex >= 0 ? data[i][balanceIndex] : runningBalance.toFixed(1)
        });
      }
    }
    
    // Sort by date, newest first
    transactions.sort((a, b) => {
      const dateA = new Date(a.date);
      const dateB = new Date(b.date);
      return dateB - dateA;
    });
    
    return { success: true, data: transactions };
    
  } catch (error) {
    console.error('Error in getEmployeePTOHistory:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Saves PTO accrual settings
 * @param {Object} settings - PTO accrual settings object
 * @return {Object} Object with success indicator
 */
function savePTOAccrualSettings(settings) {
  try {
    let settingsSheet = ss.getSheetByName('PTO Settings');
    
    // Create the settings sheet if it doesn't exist
    if (!settingsSheet) {
      settingsSheet = ss.insertSheet('PTO Settings');
      settingsSheet.getRange('A1:B1').setValues([['Setting', 'Value']]);
    } else {
      // Clear existing settings
      const lastRow = Math.max(2, settingsSheet.getLastRow());
      settingsSheet.getRange(2, 1, lastRow - 1, 2).clear();
    }
    
    // Write basic settings
    settingsSheet.getRange('A2:B7').setValues([
      ['Accrual Frequency', settings.accrualFrequency],
      ['Default Accrual Rate', settings.defaultAccrualRate],
      ['Max Carryover', settings.maxCarryover],
      ['Reset Date', settings.resetDate],
      ['Probation Period', settings.probationPeriod],
      ['Auto Approve Under 8 Hours', settings.autoApproveUnder8Hours ? 'true' : 'false']
    ]);
    
    // Save tenure rules
    if (settings.tenureRules && Array.isArray(settings.tenureRules)) {
      let tenureSheet = ss.getSheetByName('PTO Tenure Rules');
      
      // Create the tenure rules sheet if it doesn't exist
      if (!tenureSheet) {
        tenureSheet = ss.insertSheet('PTO Tenure Rules');
      } else {
        tenureSheet.clear();
      }
      
      // Set headers
      tenureSheet.getRange('A1:C1').setValues([['Years of Service', 'Accrual Rate', 'Max Balance']]);
      
      // Write rules
      if (settings.tenureRules.length > 0) {
        const tenureData = settings.tenureRules.map(rule => 
          [rule.years, rule.accrualRate, rule.maxBalance || '']
        );
        tenureSheet.getRange(2, 1, tenureData.length, 3).setValues(tenureData);
      }
    }
    
    // Save department rules
    if (settings.departmentRules) {
      let deptSheet = ss.getSheetByName('PTO Department Rules');
      
      // Create the department rules sheet if it doesn't exist
      if (!deptSheet) {
        deptSheet = ss.insertSheet('PTO Department Rules');
      } else {
        deptSheet.clear();
      }
      
      // Set headers
      deptSheet.getRange('A1:C1').setValues([['Department', 'Accrual Rate', 'Max Balance']]);
      
      // Write rules
      const departments = Object.keys(settings.departmentRules);
      if (departments.length > 0) {
        const deptData = departments.map(dept => {
          const rule = settings.departmentRules[dept];
          return [dept, rule.accrualRate || '', rule.maxBalance || ''];
        });
        deptSheet.getRange(2, 1, deptData.length, 3).setValues(deptData);
      }
    }
    
    return { success: true };
    
  } catch (error) {
    console.error('Error in savePTOAccrualSettings:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Generates a PTO report based on specified parameters
 * @param {string} reportType - 'usage', 'balance', 'forecast', or 'trend'
 * @param {string} period - 'current_month', 'current_year', 'last_month', 'last_year', 'ytd', or 'custom'
 * @param {string} groupBy - 'employee', 'department', 'month', or 'reason'
 * @param {string} department - Department filter, or 'all'
 * @param {string} startDate - Start date for custom period (optional)
 * @param {string} endDate - End date for custom period (optional)
 * @return {Object} Object with success indicator and report data
 */
function generatePTOReport(reportType, period, groupBy, department, startDate, endDate) {
  try {
    // Get date range based on period
    const dateRange = getDateRangeForPeriod(period, startDate, endDate);
    
    // Get PTO requests
    const requestsResult = getPTORequests();
    if (!requestsResult.success) {
      return requestsResult;
    }
    const requests = requestsResult.data;
    
    // Get employee data
    const employeesResult = getEmployeesForPTOManagement();
    if (!employeesResult.success) {
      return employeesResult;
    }
    const employees = employeesResult.data;
    
    // Get employee balances for balance and forecast reports
    let balances = [];
    if (reportType === 'balance' || reportType === 'forecast') {
      const balancesResult = getEmployeePTOBalances();
      if (!balancesResult.success) {
        return balancesResult;
      }
      balances = balancesResult.data;
    }
    
    // Filter requests by date range and department
    let filteredRequests = requests.filter(req => {
      const startDate = new Date(req.startDate);
      return startDate >= dateRange.start && startDate <= dateRange.end;
    });
    
    if (department && department !== 'all') {
      filteredRequests = filteredRequests.filter(req => {
        const employee = employees.find(e => e.employeeId === req.employeeId);
        return employee && employee.department === department;
      });
    }
    
    // Generate report based on type and grouping
    let reportData = { summary: {}, details: [] };
    
    // Calculate summary stats
    reportData.summary = {
      totalRequests: filteredRequests.length,
      totalHours: filteredRequests.reduce((sum, req) => sum + req.hours, 0).toFixed(1),
      avgRequestLength: filteredRequests.length > 0 ? 
        (filteredRequests.reduce((sum, req) => sum + req.hours, 0) / filteredRequests.length).toFixed(1) : 0,
      uniqueEmployees: [...new Set(filteredRequests.map(req => req.employeeId))].length
    };
    
    // Generate detailed report based on group by
    switch(reportType) {
      case 'usage':
        reportData.details = generateUsageReport(filteredRequests, employees, groupBy);
        break;
        
      case 'balance':
        reportData.details = generateBalanceReport(balances, groupBy);
        break;
        
      case 'forecast':
        reportData.details = generateForecastReport(balances, requests, employees, groupBy);
        break;
        
      case 'trend':
        reportData.details = generateTrendReport(filteredRequests, employees, groupBy, dateRange);
        break;
    }
    
    return { success: true, data: reportData };
    
  } catch (error) {
    console.error('Error in generatePTOReport:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Exports PTO balances to CSV format
 * @return {Object} Object with success indicator and CSV data
 */
function exportPTOBalancesToCSV() {
  try {
    // Get employee balances
    const balancesResult = getEmployeePTOBalances();
    if (!balancesResult.success) {
      return balancesResult;
    }
    
    const balances = balancesResult.data;
    
    // Create CSV header row
    let csv = 'Employee ID,First Name,Last Name,Department,Hire Date,PTO Balance,YTD Used,Pending Hours,Accrual Rate\n';
    
    // Add data rows
    balances.forEach(balance => {
      csv += `${balance.employeeId},`;
      csv += `${balance.firstName},`;
      csv += `${balance.lastName},`;
      csv += `${balance.department},`;
      csv += `${balance.hireDate || ''},`;
      csv += `${balance.ptoBalance},`;
      csv += `${balance.ytdUsed},`;
      csv += `${balance.pendingHours},`;
      csv += `${balance.accrualRate}\n`;
    });
    
    return { success: true, data: csv };
    
  } catch (error) {
    console.error('Error in exportPTOBalancesToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Exports a generated report to CSV format
 * @param {string} reportType - 'usage', 'balance', 'forecast', or 'trend'
 * @param {string} period - 'current_month', 'current_year', 'last_month', 'last_year', 'ytd', or 'custom'
 * @param {string} groupBy - 'employee', 'department', 'month', or 'reason'
 * @param {string} department - Department filter, or 'all'
 * @param {string} startDate - Start date for custom period (optional)
 * @param {string} endDate - End date for custom period (optional)
 * @return {Object} Object with success indicator and CSV data
 */
function exportPTOReportToCSV(reportType, period, groupBy, department, startDate, endDate) {
  try {
    // Generate the report first
    const reportResult = generatePTOReport(reportType, period, groupBy, department, startDate, endDate);
    if (!reportResult.success) {
      return reportResult;
    }
    
    const reportData = reportResult.data;
    let csv = '';
    
    // Create CSV header row based on report type and grouping
    switch(reportType) {
      case 'usage':
      case 'trend':
        csv += `${groupByToColumnName(groupBy)},Requests,Hours,Percentage\n`;
        break;
        
      case 'balance':
        csv += `${groupByToColumnName(groupBy)},Current Balance,YTD Used,Pending\n`;
        break;
        
      case 'forecast':
        csv += `${groupByToColumnName(groupBy)},Pending Requests,Pending Hours,Expected End Balance\n`;
        break;
    }
    
    // Add data rows
    reportData.details.forEach(row => {
      // First column based on grouping
      switch(groupBy) {
        case 'employee':
          csv += `${row.employee || 'Unknown'},`;
          break;
        case 'department':
          csv += `${row.department || 'Unknown'},`;
          break;
        case 'month':
          csv += `${row.month || 'Unknown'},`;
          break;
        case 'reason':
          csv += `${row.reason || 'Unknown'},`;
          break;
      }
      
      // Data columns based on report type
      if (reportType === 'usage' || reportType === 'trend') {
        csv += `${row.requests || 0},`;
        csv += `${row.hours || 0},`;
        csv += `${row.percentage || '0'}%\n`;
      } else if (reportType === 'balance') {
        csv += `${row.currentBalance || 0},`;
        csv += `${row.ytdUsed || 0},`;
        csv += `${row.pending || 0}\n`;
      } else if (reportType === 'forecast') {
        csv += `${row.pendingRequests || 0},`;
        csv += `${row.pendingHours || 0},`;
        csv += `${row.expectedEndBalance || 0}\n`;
      }
    });
    
    // Add summary row
    csv += '\nSummary Statistics\n';
    csv += `Total Requests,${reportData.summary.totalRequests}\n`;
    csv += `Total Hours,${reportData.summary.totalHours}\n`;
    csv += `Average Hours Per Request,${reportData.summary.avgRequestLength}\n`;
    csv += `Unique Employees,${reportData.summary.uniqueEmployees}\n`;
    
    return { success: true, data: csv };
    
  } catch (error) {
    console.error('Error in exportPTOReportToCSV:', error);
    return { success: false, message: error.toString() };
  }
}

/**
 * Adds a PTO transaction for an employee
 * @param {string} employeeId - The employee ID
 * @param {Date} date - Transaction date
 * @param {number} hours - Hours (positive for accrual, negative for usage)
 * @param {string} type - Transaction type
 * @param {string} notes - Transaction notes
 * @return {Object} Object with success indicator
 */
function addPTOTransaction(employeeId, date, hours, type, notes) {
  try {
    const transactionsSheet = ss.getSheetByName('PTO Transactions');
    
    // Create the transactions sheet if it doesn't exist
    if (!transactionsSheet) {
      const sheet = ss.insertSheet('PTO Transactions');
      sheet.getRange('A1:F1').setValues([
        ['Transaction ID', 'Employee ID', 'Date', 'Amount (Hours)', 'Type', 'Notes', 'Balance After']
      ]);
    }
    
    // Calculate current balance before this transaction
    let currentBalance = 0;
    const data = transactionsSheet.getDataRange().getValues();
    const headers = data[0];
    
    const empIdIndex = headers.indexOf('Employee ID');
    const amountIndex = headers.indexOf('Amount (Hours)');
    
    if (empIdIndex >= 0 && amountIndex >= 0) {
      for (let i = 1; i < data.length; i++) {
        if (data[i][empIdIndex].toString() === employeeId.toString()) {
          currentBalance += parseFloat(data[i][amountIndex]);
        }
      }
    }
    
    // Calculate new balance
    const newBalance = currentBalance + hours;
    
    // Generate transaction ID
    const transactionId = 'PTO' + new Date().getTime().toString().slice(-8);
    
    // Add new transaction
    transactionsSheet.appendRow([
      transactionId,
      employeeId,
      date,
      hours,
      type,
      notes,
      newBalance.toFixed(1)
    ]);
    
    return { success: true };
    
  } catch (error) {
    console.error('Error in addPTOTransaction:', error);
    return { success: false, message: error.toString() };
  }
}

/******************************************
 * HELPER FUNCTIONS
 ******************************************/

/**
 * Gets date range for specified period
 * @param {string} period - Period code
 * @param {string} startDate - Custom start date (optional)
 * @param {string} endDate - Custom end date (optional)
 * @return {Object} Object with start and end dates
 */
function getDateRangeForPeriod(period, startDate, endDate) {
  const today = new Date();
  const currentYear = today.getFullYear();
  const currentMonth = today.getMonth();
  
  let start, end;
  
  switch(period) {
    case 'current_month':
      start = new Date(currentYear, currentMonth, 1);
      end = new Date(currentYear, currentMonth + 1, 0);
      break;
      
    case 'last_month':
      start = new Date(currentYear, currentMonth - 1, 1);
      end = new Date(currentYear, currentMonth, 0);
      break;
      
    case 'current_year':
      start = new Date(currentYear, 0, 1);
      end = new Date(currentYear, 11, 31);
      break;
      
    case 'last_year':
      start = new Date(currentYear - 1, 0, 1);
      end = new Date(currentYear - 1, 11, 31);
      break;
      
    case 'ytd':
      start = new Date(currentYear, 0, 1);
      end = today;
      break;
      
    case 'custom':
      if (startDate && endDate) {
        start = new Date(startDate);
        end = new Date(endDate);
      } else {
        // Default to current month if custom dates aren't provided
        start = new Date(currentYear, currentMonth, 1);
        end = new Date(currentYear, currentMonth + 1, 0);
      }
      break;
      
    default:
      // Default to current month
      start = new Date(currentYear, currentMonth, 1);
      end = new Date(currentYear, currentMonth + 1, 0);
  }
  
  return { start, end };
}

/**
 * Generates usage report data
 * @param {Array} requests - Filtered PTO requests
 * @param {Array} employees - Employee data
 * @param {string} groupBy - Grouping method
 * @return {Array} Report details
 */
function generateUsageReport(requests, employees, groupBy) {
  const groups = {};
  const totalHours = requests.reduce((sum, req) => sum + req.hours, 0);
  
  // Group data
  requests.forEach(req => {
    let groupKey;
    
    switch(groupBy) {
      case 'employee':
        groupKey = req.employeeName;
        break;
        
      case 'department':
        const employee = employees.find(e => e.employeeId === req.employeeId);
        groupKey = employee ? employee.department : 'Unknown';
        break;
        
      case 'month':
        const date = new Date(req.startDate);
        groupKey = date.toLocaleString('default', { month: 'long', year: 'numeric' });
        break;
        
      case 'reason':
        groupKey = req.reason || 'No Reason Provided';
        break;
        
      default:
        groupKey = 'All';
    }
    
    if (!groups[groupKey]) {
      groups[groupKey] = { requests: 0, hours: 0 };
    }
    
    groups[groupKey].requests++;
    groups[groupKey].hours += req.hours;
  });
  
  // Convert to array and calculate percentages
  return Object.keys(groups).map(key => {
    const group = groups[key];
    const percentage = totalHours > 0 ? ((group.hours / totalHours) * 100).toFixed(1) : 0;
    
    const result = { requests: group.requests, hours: group.hours.toFixed(1), percentage };
    
    // Add the appropriate field based on grouping
    switch(groupBy) {
      case 'employee':
        result.employee = key;
        break;
      case 'department':
        result.department = key;
        break;
      case 'month':
        result.month = key;
        break;
      case 'reason':
        result.reason = key;
        break;
    }
    
    return result;
  }).sort((a, b) => b.hours - a.hours); // Sort by hours descending
}

/**
 * Generates balance report data
 * @param {Array} balances - Employee PTO balances
 * @param {string} groupBy - Grouping method
 * @return {Array} Report details
 */
function generateBalanceReport(balances, groupBy) {
  if (groupBy === 'employee') {
    // For employee grouping, return individual balances
    return balances.map(balance => ({
      employee: `${balance.firstName} ${balance.lastName}`,
      currentBalance: balance.ptoBalance,
      ytdUsed: balance.ytdUsed,
      pending: balance.pendingHours
    })).sort((a, b) => b.currentBalance - a.currentBalance);
  } else if (groupBy === 'department') {
    // For department grouping, aggregate by department
    const departments = {};
    
    balances.forEach(balance => {
      const dept = balance.department || 'Unassigned';
      
      if (!departments[dept]) {
        departments[dept] = {
          totalBalance: 0,
          totalYtdUsed: 0,
          totalPending: 0,
          count: 0
        };
      }
      
      departments[dept].totalBalance += parseFloat(balance.ptoBalance);
      departments[dept].totalYtdUsed += parseFloat(balance.ytdUsed);
      departments[dept].totalPending += parseFloat(balance.pendingHours);
      departments[dept].count++;
    });
    
    return Object.keys(departments).map(dept => ({
      department: dept,
      currentBalance: departments[dept].totalBalance.toFixed(1),
      ytdUsed: departments[dept].totalYtdUsed.toFixed(1),
      pending: departments[dept].totalPending.toFixed(1),
      employeeCount: departments[dept].count
    })).sort((a, b) => b.currentBalance - a.currentBalance);
  }
  
  // Default to employee grouping
  return generateBalanceReport(balances, 'employee');
}

/**
 * Generates forecast report data
 * @param {Array} balances - Employee PTO balances
 * @param {Array} requests - PTO requests
 * @param {Array} employees - Employee data
 * @param {string} groupBy - Grouping method
 * @return {Array} Report details
 */
function generateForecastReport(balances, requests, employees, groupBy) {
  // Get pending requests
  const pendingRequests = requests.filter(req => req.status === 'Pending');
  
  if (groupBy === 'employee') {
    // For employee grouping, return individual forecasts
    return balances.map(balance => {
      // Calculate pending requests for this employee
      const employeePending = pendingRequests.filter(req => req.employeeId === balance.employeeId);
      const pendingCount = employeePending.length;
      const pendingHours = employeePending.reduce((sum, req) => sum + req.hours, 0);
      const expectedBalance = parseFloat(balance.ptoBalance) - pendingHours;
      
      return {
        employee: `${balance.firstName} ${balance.lastName}`,
        pendingRequests: pendingCount,
        pendingHours: pendingHours.toFixed(1),
        expectedEndBalance: expectedBalance.toFixed(1)
      };
    }).sort((a, b) => b.pendingHours - a.pendingHours);
  } else if (groupBy === 'department') {
    // For department grouping, aggregate by department
    const departments = {};
    
    // Initialize department totals
    employees.forEach(employee => {
      const dept = employee.department || 'Unassigned';
      if (!departments[dept]) {
        departments[dept] = {
          pendingRequests: 0,
          pendingHours: 0,
          currentBalance: 0,
          expectedBalance: 0,
          count: 0
        };
      }
    });
    
    // Add current balances
    balances.forEach(balance => {
      const dept = balance.department || 'Unassigned';
      if (departments[dept]) {
        departments[dept].currentBalance += parseFloat(balance.ptoBalance);
        departments[dept].count++;
      }
    });
    
    // Add pending requests
    pendingRequests.forEach(req => {
      const employee = employees.find(e => e.employeeId === req.employeeId);
      if (employee) {
        const dept = employee.department || 'Unassigned';
        if (departments[dept]) {
          departments[dept].pendingRequests++;
          departments[dept].pendingHours += req.hours;
        }
      }
    });
    
    // Calculate expected balances
    Object.keys(departments).forEach(dept => {
      departments[dept].expectedBalance = departments[dept].currentBalance - departments[dept].pendingHours;
    });
    
    return Object.keys(departments).map(dept => ({
      department: dept,
      pendingRequests: departments[dept].pendingRequests,
      pendingHours: departments[dept].pendingHours.toFixed(1),
      expectedEndBalance: departments[dept].expectedBalance.toFixed(1),
      employeeCount: departments[dept].count
    })).sort((a, b) => b.pendingHours - a.pendingHours);
  }
  
  // Default to employee grouping
  return generateForecastReport(balances, requests, employees, 'employee');
}

/**
 * Generates trend report data
 * @param {Array} requests - Filtered PTO requests
 * @param {Array} employees - Employee data
 * @param {string} groupBy - Grouping method
 * @param {Object} dateRange - Date range object
 * @return {Array} Report details
 */
function generateTrendReport(requests, employees, groupBy, dateRange) {
  // For trend reports, we group by month first, then by the specified grouping
  const startMonth = dateRange.start.getMonth();
  const startYear = dateRange.start.getFullYear();
  const endMonth = dateRange.end.getMonth();
  const endYear = dateRange.end.getFullYear();
  
  // Create array of months in range
  const months = [];
  for (let year = startYear; year <= endYear; year++) {
    const monthStart = (year === startYear) ? startMonth : 0;
    const monthEnd = (year === endYear) ? endMonth : 11;
    
    for (let month = monthStart; month <= monthEnd; month++) {
      months.push({ year, month });
    }
  }
  
  // Group requests by month and groupBy
  const trends = {};
  
  months.forEach(m => {
    const monthKey = new Date(m.year, m.month, 1).toLocaleString('default', { month: 'long', year: 'numeric' });
    trends[monthKey] = {};
    
    // Filter requests for this month
    const monthRequests = requests.filter(req => {
      const date = new Date(req.startDate);
      return date.getMonth() === m.month && date.getFullYear() === m.year;
    });
    
    // Group by specified grouping
    monthRequests.forEach(req => {
      let groupKey;
      
      switch(groupBy) {
        case 'employee':
          groupKey = req.employeeName;
          break;
          
        case 'department':
          const employee = employees.find(e => e.employeeId === req.employeeId);
          groupKey = employee ? employee.department : 'Unknown';
          break;
          
        case 'reason':
          groupKey = req.reason || 'No Reason Provided';
          break;
          
        default:
          groupKey = 'All';
      }
      
      if (!trends[monthKey][groupKey]) {
        trends[monthKey][groupKey] = { requests: 0, hours: 0 };
      }
      
      trends[monthKey][groupKey].requests++;
      trends[monthKey][groupKey].hours += req.hours;
    });
  });
  
  // Flatten the data for reporting
  const result = [];
  
  Object.keys(trends).forEach(month => {
    const monthData = trends[month];
    const totalHours = Object.values(monthData).reduce((sum, g) => sum + g.hours, 0);
    
    Object.keys(monthData).forEach(group => {
      const groupData = monthData[group];
      const percentage = totalHours > 0 ? ((groupData.hours / totalHours) * 100).toFixed(1) : 0;
      
      const entry = {
        month,
        requests: groupData.requests,
        hours: groupData.hours.toFixed(1),
        percentage
      };
      
      // Add the appropriate field based on grouping
      switch(groupBy) {
        case 'employee':
          entry.employee = group;
          break;
        case 'department':
          entry.department = group;
          break;
        case 'reason':
          entry.reason = group;
          break;
      }
      
      result.push(entry);
    });
  });
  
  return result.sort((a, b) => {
    // Sort by month then by hours
    if (a.month !== b.month) {
      return new Date(a.month) - new Date(b.month);
    }
    return b.hours - a.hours;
  });
}

/**
 * Converts groupBy value to human-readable column name
 * @param {string} groupBy - Grouping method
 * @return {string} Column name
 */
function groupByToColumnName(groupBy) {
  switch(groupBy) {
    case 'employee':
      return 'Employee';
    case 'department':
      return 'Department';
    case 'month':
      return 'Month';
    case 'reason':
      return 'Reason';
    default:
      return 'Group';
  }
}

/**
 * Formats a date to YYYY-MM-DD
 * @param {Date|string} date - Date to format
 * @return {string} Formatted date
 */
function formatDate(date) {
  if (!date) return '';
  
  let d;
  if (date instanceof Date) {
    d = date;
  } else {
    d = new Date(date);
    if (isNaN(d.getTime())) return date; // Return original if invalid
  }
  
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  
  return `${year}-${month}-${day}`;
}
