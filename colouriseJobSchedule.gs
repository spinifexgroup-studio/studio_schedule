
/**
 * Colourise the job section
 */
function colouriseJobSchedule () {
  var startRowID = 11;
  var startColumnID = 5;

  var jobsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("jobs");
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var scheduleRange = scheduleSheet.getRange( "E11:CX60" )
  var scheduleValues = scheduleRange.getValues();

  var scheduleIDRange = scheduleSheet.getRange( "A11:A60" )
  var scheduleIDValues = scheduleIDRange.getValues();

  var jobsRange = jobsSheet.getRange( "A11:F50" )
  var jobsValues = jobsRange.getValues();

  for (var s = 0; s < 40 ; s++) { //Iterate through all jobs on schedule page
    var scheduleID = scheduleIDValues[s][0];
    
    for (var j = 0; j < 40 ; j++) { //Iterate through all jobs on jobs pages
      var jobID = jobsValues[j][0];
      var scheduleIdRow = startRowID+j;
          
      if (scheduleID.length != 0) {
        if (scheduleID == jobID){
          for ( var c = 0 ; c < scheduleRange.getNumColumns() ; c++ ){
            var cell = scheduleValues[j][c];
            
            if (cell == "C" || cell == "c" || cell == "H" || cell == "h") {
   
              var jobType = jobsValues[j][3];
              
              if (jobType == "I" ) {
                scheduleSheet.getRange(scheduleIdRow, startColumnID+c, 1, 1).setBackgroundColor("#ffd966");
              }

              else if (jobType == "S" ) {
                scheduleSheet.getRange(scheduleIdRow, startColumnID+c, 1, 1).setBackgroundColor("#6fa8dc");
              }

              else if (jobType == "SI" ) {
                scheduleSheet.getRange(scheduleIdRow, startColumnID+c, 1, 1).setBackgroundColor("#38761d");
              }
              
            }
          }
        }
      }
    }
  }
}
