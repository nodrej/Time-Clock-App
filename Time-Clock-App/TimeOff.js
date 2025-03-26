// Copyright 2020 Google LLC
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     https://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.


const Header = {
  Timestamp: 'Timestamp',
  EmailAddress: 'Email Address',
  Name: 'Name',
  StartDate: 'Start date',
  EndDate: 'End date',
  Reason: 'Reason',
  AdditionalEmail: 'Additional email',
  Approval: 'Approval',
  NotifiedStatus: 'Notified status',
};

const Reason = {
  Vacation: 'Vacation',
  SickLeave: 'Sick leave',
  MaternityPaternity: 'Maternity/Paternity',
  Breavement: 'Bereavement',
  LeaveOfAbsence: 'Leave of absence',
  PersonalTime: 'Personal time',
};

const Approval = {
  InProgress: 'In progress',
  Approved: 'Approved',
  NotApproved: 'Not approved',
};

const NotifiedStatus = {
  NotNotified: 'Not notified',
  Notified: 'Notified',
};


/**
 * Add custom menu items when opening the sheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Time off')
      .addItem('Form setup', 'formSetup')
      .addItem('Column setup', 'columnSetup')
      .addItem('Notify employees', 'notify')
      .addToUi();
}



/**
 * Set up the "Request time off" form, and link the form's trigger to 
 * optionally send an email to an additional address (like a manager).
 */
function formSetup() {
  let sheet = SpreadsheetApp.getActiveSpreadsheet();
  if (sheet.getFormUrl()) {
    let ui = SpreadsheetApp.getUi();
    ui.alert(
      'ℹ️ A Form already exists',
      'Unlink the form and try again.\n\n' +
      'From the top menu:\n' +
      'Click "Form" > "Unlink form"',
      ui.ButtonSet.OK
    );
    return;
  }

  // Create the form.
  let form = FormApp.create('Request time off')
      .setCollectEmail(true)
      .setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId())
      .setLimitOneResponsePerUser(false);

  form.addTextItem().setTitle(Header.Name).setRequired(false);
  form.addDateItem().setTitle(Header.StartDate).setRequired(true);
  form.addDateItem().setTitle(Header.EndDate).setRequired(true);
  form.addListItem().setTitle(Header.Reason).setChoiceValues(Object.values(Reason)).setRequired(false);
  form.addTextItem().setTitle(Header.AdditionalEmail).setRequired(false);
}


/**
 * Creates an "Approval" and "Notified status" column
 */
function columnSetup() {
  let sheet = SpreadsheetApp.getActiveSheet();

  appendColumn(sheet, Header.Approval, Object.values(Approval));
  appendColumn(sheet, Header.NotifiedStatus, Object.values(NotifiedStatus));
}


/**
 * Appends a new column.
 * 
 *  @param {SpreadsheetApp.Sheet} sheet - tab in sheet.
 *  @param {string} headerName - name of column.
 *  @param {(string[] | null)} maybeChoices - optional drop down values for validation.
 */
function appendColumn(sheet, headerName, maybeChoices) {
  let range = sheet.getRange(1, sheet.getLastColumn() + 1);

  // Create the header header name.
  range.setValue(headerName);

  // If we pass choices to the function, create validation rules.
  if (maybeChoices) {
    let rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(maybeChoices)
        .build();

    range.offset(sheet.getFrozenRows(), 0, sheet.getMaxRows())
        .setDataValidation(rule);
  }
}


/**
 * Checks the notification status of each entry and, if not notified,
 * notifies them of their status accordingly.
 */
function notify() {
  let sheet = SpreadsheetApp.getActiveSheet();
  let dataRange = sheet.getDataRange().getValues();
  let headers = dataRange.shift();

  validateSheetHeaders(headers, Header);

  let rows = dataRange
      .map((row, i) => asObject(headers, row, i))
      .filter(row => row[Header.NotifiedStatus] != NotifiedStatus.Notified)
      .map(process)
      .map(row => writeRowToSheet(sheet, headers, row));
}


/**
 * Validate that the sheet headers match a schema.
 */
function validateSheetHeaders(headers, schema) {
  for (let header of Object.values(schema)) {
    if (!headers.includes(header)) {
      throw `🦕 ⚠️ Header "${header}" not found in sheet: ${JSON.stringify(headers)}`;
    }
  }
}


/**
 * Convert the row arrays into objects.
 * Start with an empty object, then create a new field
 * for each header name using the corresponding row value.
 * 
 * @param {string[]} headers - list of column names.
 * @param {any[]} rowArray - values of a row as an array.
 * @param {int} rowIndex - index of the row.
 */
function asObject(headers, rowArray, rowIndex) {
  return headers.reduce(
    (row, header, i) => {
      row[header] = rowArray[i];
      return row;
    }, {rowNumber: rowIndex + 1});
}


/**
 * Checks if a row is marked as "approved". If approved a calendar
 * invite is created for the user. If not approved, an email
 * notification is sent.
 * 
 * @param {Object} row - values in a row.
 * @returns {Object} the row with a "notified status" column populated.
 */
function process(row) {
  let email = row[Header.EmailAddress];
  let additionalEmail = row[Header.AdditionalEmail];
  let startDate = row[Header.StartDate];
  let endDate = row[Header.EndDate];
  let approval = row[Header.Approval];
  let message = `Your vacation time request from `
      + `${startDate.toDateString()} to `
      + `${endDate.toDateString()}: ${approval}`;

  if (approval == Approval.NotApproved) {
    // If not approved, send an email.
    let subject = 'Your vacation time request was NOT approved';
    MailApp.sendEmail(email, subject, message);
    row[Header.NotifiedStatus] = NotifiedStatus.Notified;

    Logger.log(`Not approved, email sent, row=${JSON.stringify(row)}`);
  }

  else if (approval == Approval.Approved) {
    // If approved, create a calendar event.
    CalendarApp.getCalendarById(email)
        .createAllDayEvent(
            'OOO - Out Of Office',
            startDate,
            endDate,
            {
              description: message,
              guests: additionalEmail,
              sendInvites: true,
            });

    // Send a confirmation email.
    let subject = 'Confirmed, your vacation time request has been approved!';
    MailApp.sendEmail(email, subject, message, {cc: additionalEmail});

    row[Header.NotifiedStatus] = NotifiedStatus.Notified;

    Logger.log(`Approved, calendar event created, row=${JSON.stringify(row)}`);
  }

  else {
    row[Header.NotifiedStatus] = NotifiedStatus.NotNotified;

    Logger.log(`No action taken, row=${JSON.stringify(row)}`);
  }

  return row;
}


/**
 * Rewrites a row into the sheet.
 * 
 * @param {SpreadsheetApp.Sheet} sheet - tab in sheet.
 * @param {string[]} headers - list of column names.
 * @param {Object} row - values in a row.
 */
function writeRowToSheet(sheet, headers, row) {
  let rowArray = headers.map(header => row[header]);
  let rowNumber = sheet.getFrozenRows() + row.rowNumber;
  sheet.getRange(rowNumber, 1, 1, rowArray.length).setValues([rowArray]);
}
