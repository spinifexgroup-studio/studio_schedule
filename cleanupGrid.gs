/**
 * perform cleanup functions if things are out of order
 */
 


function cleanupGrid() {
  var sheetSchedule     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var fullRange = sheetSchedule.getDataRange();
  
  fullRange.setBorder(true, true, true, true, true, true);
};
