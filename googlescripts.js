/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Generate step-by-step', functionName: 'generateStepByStep_'},
    {name: 'Trace Dependents', functionName: 'traceDependents'},
    {name: 'Read Other File', functionName: 'readRows'},
    {name: 'Sort', functionName: 'sortrange'}
  ];
  spreadsheet.addMenu("Jared's Custom Scripts", menuItems);
}

// Performs a saved sort with a button
function sortrange() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheetCandidateList = spreadsheet.getSheetByName('CandidateList');
  var lastRow = sheetCandidateList.getLastRow();
  var lastColumn = sheetCandidateList.getLastColumn();
  var range = sheetCandidateList.getRange(2, 1, lastRow-1, lastColumn);
  range.sort(2);
}

/**
 * A custom function that converts meters to miles.
 *
 * @param {Number} meters The distance in meters.
 * @return {Number} The distance in miles.
 */
function metersToMiles(meters) {
  if (typeof meters != 'number') {
    return null;
  }
  return meters / 1000 * 0.621371;
}

/**
 * A custom function that gets the driving distance between two addresses.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Number} The distance in meters.
 */
function drivingDistance(origin, destination) {
  var directions = getDirections_(origin, destination);
  return directions.routes[0].legs[0].distance.value;
}

/**
 * Creates a new sheet containing step-by-step directions between the two
 * addresses on the "Settings" sheet that the user selected.
 */
function generateStepByStep_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('Settings');
  settingsSheet.activate();

  // Prompt the user for a row number.
  var selectedRow = Browser.inputBox('Generate step-by-step',
      'Please enter the row number of the addresses to use' +
      ' (for example, "2"):',
      Browser.Buttons.OK_CANCEL);
  if (selectedRow == 'cancel') {
    return;
  }
  var rowNumber = Number(selectedRow);
  if (isNaN(rowNumber) || rowNumber < 2 ||
      rowNumber > settingsSheet.getLastRow()) {
    Browser.msgBox('Error',
        Utilities.formatString('Row "%s" is not valid.', selectedRow),
        Browser.Buttons.OK);
    return;
  }

  // Retrieve the addresses in that row.
  var row = settingsSheet.getRange(rowNumber, 1, 1, 2);
  var rowValues = row.getValues();
  var origin = rowValues[0][0];
  var destination = rowValues[0][1];
  if (!origin || !destination) {
    Browser.msgBox('Error', 'Row does not contain two addresses.',
        Browser.Buttons.OK);
    return;
  }

  // Get the raw directions information.
  var directions = getDirections_(origin, destination);

  // Create a new sheet and append the steps in the directions.
  var sheetName = 'Driving Directions for Row ' + rowNumber;
  var directionsSheet = spreadsheet.getSheetByName(sheetName);
  if (directionsSheet) {
    directionsSheet.clear();
    directionsSheet.activate();
  } else {
    directionsSheet =
        spreadsheet.insertSheet(sheetName, spreadsheet.getNumSheets());
  }
  var sheetTitle = Utilities.formatString('Driving Directions from %s to %s',
      origin, destination);
  var newRows = [
    [sheetTitle, '', ''],
    ['Step', 'Distance (Meters)', 'Distance (Miles)']
  ];
  for (var i = 0; i < directions.routes[0].legs[0].steps.length; i++) {
    var step = directions.routes[0].legs[0].steps[i];
    // Remove HTML tags from the instructions.
    var instructions = step.html_instructions.replace(/<br>|<div.*?>/g, '\n')
        .replace(/<.*?>/g, '');
    newRows.push([
      instructions,
      step.distance.value,
      '=METERSTOMILES(R[0]C[-1])'
    ]);
  }
  directionsSheet.getRange(1, 1, newRows.length, 3).setValues(newRows);

  // Format the new sheet.
  directionsSheet.getRange('A1:C1').merge().setBackground('#ddddee');
  directionsSheet.getRange('A1:2').setFontWeight('bold');
  directionsSheet.setColumnWidth(1, 500);
  directionsSheet.getRange('B2:C').setVerticalAlignment('top');
  directionsSheet.getRange('C2:C').setNumberFormat('0.00');
  var stepsRange = directionsSheet.getDataRange()
      .offset(2, 0, directionsSheet.getLastRow() - 2);
  setAlternatingRowBackgroundColors_(stepsRange, '#ffffff', '#eeeeee');
  directionsSheet.setFrozenRows(2);
  SpreadsheetApp.flush();
}

/**
 * Sets the background colors for alternating rows within the range.
 * @param {Range} range The range to change the background colors of.
 * @param {string} oddColor The color to apply to odd rows (relative to the
 *     start of the range).
 * @param {string} evenColor The color to apply to even rows (relative to the
 *     start of the range).
 */
function setAlternatingRowBackgroundColors_(range, oddColor, evenColor) {
  var backgrounds = [];
  for (var row = 1; row <= range.getNumRows(); row++) {
    var rowBackgrounds = [];
    for (var column = 1; column <= range.getNumColumns(); column++) {
      if (row % 2 == 0) {
        rowBackgrounds.push(evenColor);
      } else {
        rowBackgrounds.push(oddColor);
      }
    }
    backgrounds.push(rowBackgrounds);
  }
  range.setBackgrounds(backgrounds);
}

/**
 * A shared helper function used to obtain the full set of directions
 * information between two addresses. Uses the Apps Script Maps Service.
 *
 * @param {String} origin The starting address.
 * @param {String} destination The ending address.
 * @return {Object} The directions response object.
 */
function getDirections_(origin, destination) {
  var directionFinder = Maps.newDirectionFinder();
  directionFinder.setOrigin(origin);
  directionFinder.setDestination(destination);
  var directions = directionFinder.getDirections();
  if (directions.routes.length == 0) {
    throw 'Unable to calculate directions between these addresses.';
  }
  return directions;
}

/*
  functions for tracing dependents
*/

/*
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Trace Dependents', functionName: 'traceDependents'}
  ];
  spreadsheet.addMenu("Jared's Custom Scripts", menuItems);
}
*/

function traceDependents(){
  var dependents = [];
  var spreadsheet = SpreadsheetApp.getActive();
  var currentCell = spreadsheet.getActiveCell();
  var currentCellRef = currentCell.getA1Notation();
  var range = spreadsheet.getDataRange();

  var regex = new RegExp("\\b" + currentCellRef + "\\b");
  var formulas = range.getFormulas();

  for (var i = 0; i < formulas.length; i++){
    var row = formulas[i];

    for (var j = 0; j < row.length; j++){
      var cellFormula = row[j];
      if (regex.test(cellFormula)){
        dependents.push([i,j]);
      }
    }
  }

  var output = "Dependents: ";
  var dependentRefs = [];
  for (var k = 0; k < dependents.length; k++){
    var rowNum = dependents[k][0] + 1;
    var colNum = dependents[k][1] + 1;
    var cell = range.getCell(rowNum, colNum);
    var cellRef = cell.getA1Notation();
    dependentRefs.push(cellRef);
  }

  if(dependentRefs.length > 0){
    output += dependentRefs.join(", ");
  } else {
    output += " None";
  }

  currentCell.setNote(output);
}

function readRows()
{
  // get data - https://docs.google.com/spreadsheet/ccc?key=0AlF-5fwPmkDodDNiYTU4QUFZZld6RFNNanQwU09MOXc#gid=0
  var sheet = SpreadsheetApp.openById("0AlF-5fwPmkDodDNiYTU4QUFZZld6RFNNanQwU09MOXc")
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  // write data
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName("Sheet");

  Logger.log(rows);
  Logger.log(values);
  Logger.log(numRows);
  Logger.log(values[0][1]);

  // https://developers.google.com/apps-script/reference/spreadsheet/range

  var cell = sheet.getRange(2,2,3,4); // (row, column, height, width)
  cell.setValue(values);

  // cell.setValue(100);


  // write data to log
  for (var i = 0; i <= numRows - 1; i++) { var row = values[i]; Logger.log(row); }
};

function onEdit(event)
{
  // var range_active = event.source.getActiveSheet(); // if wanted this sheet
  var sheet = SpreadsheetApp.getActive().getSheetByName("onEdit");
  var cell = sheet.getRange(2,2);
  var value = cell.getValue();
  cell.setValue(value + 1);

  // duplicates any change into cell E1, puts row and col in E2, E3
  sheet.getRange(1,5).setValue(event.source.getActiveRange().getValue());
  sheet.getRange(2,5).setValue(event.source.getActiveRange().getRow());
  sheet.getRange(3,5).setValue(event.source.getActiveRange().getColumn());

  var last_column_edited = event.source.getActiveRange().getColumn();
  var range_last_column_edited = sheet.getRange(1, last_column_edited, sheet.getMaxRows(), 1).getA1Notation();
  sheet.getRange(4,5).setValue(range_last_column_edited);
  // sheet.getRange(5,5).setValue(sheet.getMaxRows());

  // Browser.msgBox(sheet.getRange(1, last_column_edited, sheet.getMaxRows(), 1).getValues());

  sheet.getRange(1,8).setValue(sheet.getRange(1, last_column_edited, sheet.getMaxRows(), 1).getValue());

}

// http://webapps.stackexchange.com/questions/53555/google-sheets-how-can-i-use-row-values-to-index-into-another-sheet
function dataSubset(data,subset) {
  var output = [];
  for(var i=0, iLen=subset.length; i<iLen; i++) {
    for(var j=1, jLen=data.length+1; j<jLen; j++) {
      if(parseInt(subset[i]) == parseInt(j)) {
        output.push(data[j-1]);
      }
    }
  }
  return output;
}

// Google Script to Add 5 New Rows to a Sheet
function addRow() {
  // Get the sheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Copy down columns C-F starting with the last row
  var lastRow = sheet.getLastRow();
  var rangeColumns = [ ["C", "F"], ["J", "L"], ["P", "R"] ];

  for (var i = 0; i < rangeColumns.length; i++) {
    var range = sheet.getRange(rangeColumns[i][0] + lastRow + ":" + rangeColumns[i][1] + lastRow);
    range.copyTo(sheet.getRange(rangeColumns[i][0] + (lastRow + 1) + ":" + rangeColumns[i][1] + (lastRow + 5)));
  }
}