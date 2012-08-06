/**
 * perform cleanup functions if things are out of order
 */
function cleanUp() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Please don't do anything until it's ok to start working", "Warning", 5);

  colouriseJobSchedule();
  copyJobsIntoLegend();
  hideEmptyRows();
  checkForWeekStart();

  SpreadsheetApp.getActiveSpreadsheet().toast("Ok to start working", "Hey there!", 5);
};