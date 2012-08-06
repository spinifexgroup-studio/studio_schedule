
/**
 * deleteJob
 */
function deleteItem() {

  // Get data about spreadsheet
  var sheetSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studo Schedule");
  var sheetJobs     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("jobs");
  var sheetCurrent  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Check we are in a sheet where we can delete a job
  if (sheetCurrent.getIndex() == sheetJobs.getIndex() || sheetCurrent.getIndex() == sheetSchedule.getIndex()) {
  
    var currentRange = sheetCurrent.getActiveRange();
    
    // Check we have only one row selected
    if (currentRange.getNumRows() == 1) {
    
      // Get uID for the row
      var startRow  = currentRange.getRowIndex();
      var colIndex  = 1;
      var numRows   = 1;
      var colRange  = 10;
      var dataRange = sheetCurrent.getRange(startRow, colIndex, numRows, colRange);
      
      var uID = dataRange.getValue();
      

      
      // Check if uID is in job range
      if ( uID >= 10000 ) {
        Browser.msgBox ("Success!");
      }
      else if ( uID < 10000 && uID >= 100 ) {
        Browser.msgBox ("You need to select a job to delete it. Use \"Delete Resource\" to delete resources.");
      } 
      else {
        Browser.msgBox ("You need to select a job to delete it.");
      }
    }
    else {
      Browser.msgBox ("One, and only one row can be selected");
    }
  }
  else {
    Browser.msgBox ("You must be in the spreadsheet \"studioSchedule\" or \"Jobs\"");
  }
};
