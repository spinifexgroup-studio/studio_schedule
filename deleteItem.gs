
/**
 * deleteItem
 */
function deleteItem() {

  // Get data about spreadsheet
  var sheetSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  var sheetCurrent  = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Check we are in a sheet where we can delete a job
  if (sheetCurrent.getIndex() == sheetSchedule.getIndex()) {
  
    var currentRange = sheetCurrent.getActiveRange();
    
    // Check we have only one row selected
    if (currentRange.getNumRows() == 1) {
    
      // Get uID for the row
      var row  = currentRange.getRowIndex();
      var dataRange = sheetSchedule.getRange(row, 1);
      
      var uID = dataRange.getValue();
      if (typeof uID == "string" ){
        SpreadsheetApp.getActiveSpreadsheet().toast("You can't delete that item.", "Warning!", 5);

      }
      else {
        if ( uID >= 10000 ) {
          sheetSchedule.deleteRow(row);
          copyJobsIntoLegend();
        }
        else {
          sheetSchedule.deleteRow(row);
        }
      }
    }
    else {
      SpreadsheetApp.getActiveSpreadsheet().toast("Only one item can be deleted at a time.", "Warning!", 5);
    }
  }
  else {
    SpreadsheetApp.getActiveSpreadsheet().toast("You must be in the spreadsheet \"Studio Schedule\"", "Warning!", 5);
  }
};
