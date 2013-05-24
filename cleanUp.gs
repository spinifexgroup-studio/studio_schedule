/**
 * perform cleanup functions if things are out of order
 */
 
 
 
function cleanupViaMenu() {

  SpreadsheetApp.getActiveSpreadsheet().toast("Cleaning up. Ok to start working in about 30 secs", "Please Wait!", 5);
  
  copyJobsIntoLegend();
  colouriseJobSchedule();
  cleanupGrid();
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Ok to start working", "Hey there!", 5);
  
};


function cleanupOnOpen() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Cleaning up. Ok to start working in about 30 secs.", "Please Wait!", 5);
  
  var sheetSchedule     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  sheetSchedule.hideColumns(1, 2);
  sheetSchedule.hideRows(sheetSchedule.getLastRow());

  copyJobsIntoLegend();

  SpreadsheetApp.getActiveSpreadsheet().toast("Ok to start working", "Hey there!", 5);
  
};
