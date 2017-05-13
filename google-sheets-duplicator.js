/* global Logger SpreadsheetApp */

/**
 * Retrieves all the rows in the active spreadsheet that contain data and logs the
 * values for each row.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */

function buildMonth() {  // eslint-disable-line no-unused-vars
  Logger.clear();
  const months = ['Gennaio', 'Febbraio', 'Marzo', 'Aprile', 'Maggio', 'Giugno', 'Luglio', 'Agosto', 'Settembre', 'Ottobre', 'Novembre', 'Dicembre'];

  const sheet = SpreadsheetApp.getActiveSheet();

  const newSheet = SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet();

  const lastName = sheet.getName();
  const oldMonth = lastName.replace(/[0-9]*$/gi, '');
  const oldYear = parseInt(lastName.replace(oldMonth, ''), 10);

  const getNewMonthAndNewYear = ({ months, oldMonth, oldYear }) => { // eslint-disable-line no-shadow
    const lastPos = months.indexOf(oldMonth);
    return lastPos === 11 ? ({
      newMonthIndex: 0,
      newYear: oldYear + 1,
    }) : ({
      newMonthIndex: lastPos + 1,
      newYear: oldYear,
    });
  };

  const { newMonthIndex, newYear } = getNewMonthAndNewYear({ months, oldMonth, oldYear });

  // Set sheet's name
  newSheet.setName(months[newMonthIndex] + newYear);
  // Reset content
  newSheet.clearContents();
  // Copy headers
  const rangeToCopy = sheet.getRange('A1:C2');
  rangeToCopy.copyTo(newSheet.getRange(1, 1));
  // Copy legend
  // rangeToCopy = sheet.getRange('D12:D17');
  // rangeToCopy.copyTo(newSheet.getRange(12, 4));
  const legend = newSheet.getRange('D:D');
  legend.setBackground('white');
  // Set sum cell
  newSheet.getRange(8, 4).setFormula('=SUM(B:B)');

  const dateRowRange = newSheet.getRange(3, 1, 32, 1);


  const initialDate = new Date(newYear, newMonthIndex, 1, 0, 0, 0, 0);

  Date.prototype.addDays = function addDays(days) {
    const dat = new Date(this.valueOf());
    dat.setDate(dat.getDate() + days);
    return dat;
  };

  const getNewDateArray = ({ oldDateRange }) => {
    const dateRowsCount = oldDateRange.getNumRows();
    const dateArray = [];
    for (let i = 0, numRows = dateRowsCount - 1; i <= numRows; i += 1) {
      dateArray.push([initialDate.addDays(i)]);
    }
    return dateArray;
  };
  const dateArray = getNewDateArray({ oldDateRange: dateRowRange });

  dateRowRange.setValues(dateArray);
  const row2RowRange = newSheet.getRange(3, 2, 32, 2);
  row2RowRange.setBackground('#F9CB9C');

  SpreadsheetApp.setActiveSheet(newSheet);
  SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(1);
}

/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the readRows() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() { // eslint-disable-line no-unused-vars
  const sheet = SpreadsheetApp.getActiveSpreadsheet();
  const entries = [{
    name: 'New Month',
    functionName: 'buildMonth',
  }];
  sheet.addMenu('Refactored google sheets duplicator', entries);
}
