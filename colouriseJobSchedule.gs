
/**
 * Colourise the job section
 */
function colouriseJobSchedule () {
  var startRowID = 11;
  var startColumnID = 6;

  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  var rows = scheduleSheet.getDataRange();
  var values = rows.getValues();
  var lastJobRow = 11;
  
  for (var i = 0; i <= values.length - 1; i++){
    var cell = values[i][0];
    if (typeof cell == "string" ){
      if (cell == "desID" ){
        lastJobRow = i+1;
      }
    }
  }

  var scheduleRange = scheduleSheet.getRange(startRowID, startColumnID, lastJobRow-11, scheduleSheet.getLastColumn() - startRowID);
  var scheduleValues = scheduleRange.getValues();

  var scheduleIDRange = scheduleSheet.getRange(startRowID, 1, lastJobRow-startRowID, 1);
  var scheduleIDValues = scheduleIDRange.getValues();

  var jobsRange = scheduleSheet.getRange(startRowID, 1, lastJobRow-startRowID, 5);
  var jobsValues = jobsRange.getValues();

  for (var s = 0; s < lastJobRow-startRowID; s++) { //Iterate through all jobs on schedule page
    var scheduleID = scheduleIDValues[s][0];
    var jobType = jobsValues[s][1].substring(0,1);
    
    for ( var c = 0 ; c < scheduleRange.getNumColumns() ; c++ ){
      var cell = scheduleValues[s][c];
      
      if (cell == "C" || cell == "c" || cell == "H" || cell == "h") {
        
        if (jobType == "I" ) {
          scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setBackgroundColor("#ffd966");
        }

        else if (jobType == "S" ) {
          scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setBackgroundColor("#6fa8dc");
        }

        else if (jobType == "E" ) {
          scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setBackgroundColor("#38761d");
        }
      }
      else if (cell == "D" || cell == "d") {
        scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setBackgroundColor("#ff0000");
        scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setFontColor("#ffff00");
      }

      else if (cell == "S" || cell == "s") {
        scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setBackgroundColor("#999999");
        scheduleSheet.getRange(s+startRowID, startColumnID+c, 1, 1).setFontColor("#ffff00");
      }
    }
  }
}
