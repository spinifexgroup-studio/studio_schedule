





/**
 * Hide empty lines in studio schedule
 */
function hideEmptyRows() {

  var startRow = 11;

  // Get data about spreadsheet
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  
  //sheet.showColumns(1, sheet.getLastColumn());   // Show all the rows
  sheet.showRows(1, sheet.getLastRow());
  
  // Loop through all the rows

  for (var i = startRow-1; i <= values.length - 1; i++) {
    /*
    Get data from column A
    Iterate through rows until encountering a blank row
    Keep iterating through blank rows until finding a row with content
    Hide the empty rows
    */
    
    var cell = values[i][0];
    var rowToHide = i+1;
    if (cell.length == 0){
      var hideRowCount = 1;
      var doExit = 0;
      while ( doExit < 1 ) {
        i++;
        var cell = values[i][0];
        if (cell.length == 0){ 
          hideRowCount++;
        }
        else{
          doExit = 1;
        }
      }
      sheet.hideRows(rowToHide,hideRowCount);
    }
    else {
    //whilst we're at it - make permanaent people bold and freelancers normal text
      if ( values[i][1] == "P" ){
        sheet.getRange(i+1, 3, 1, 1).setFontWeight("bold");
      }
      else {
        sheet.getRange(i+1, 3, 1, 1).setFontWeight("normal");
      }
    }
  }
  sheet.hideColumns(1);   // Hide the ID column
};
