
/**
 * delete last week if a week has passed
 */

function deletePreviousWeek() {
  var variablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  var weekStart = variablesSheet.getRange("startDate").getValues();    // startDate is in ranges list
  
  var oldStartDate = new Date(weekStart);
  var oldStartTime = oldStartDate.getTime();
  var todayTime = new Date().getTime();
  
  var oneDayMS = 86400000;
  var timeDiff = todayTime-oldStartTime;
  
  if ( timeDiff > (9*oneDayMS)) {   // See if a week and 2 days in milliseconds has passed (we do this on a wednesday)
  
    // Set the new starting monday
    var newStartDate = new Date(oldStartDate.getTime() + (7*oneDayMS) );
    var newStartDateString = newStartDate.getDate()+"/"+(newStartDate.getMonth()+1)+"/"+newStartDate.getFullYear()  // Javascript months are 0-11
    
    // Get data about spreadsheet

    var rows = scheduleSheet.getDataRange();
    var values = rows.getValues();
    
    // Loop through all the rows and find ranges to move
    var moveRanges = new Array();
    var startRanges = new Array();
    
    var rangeRowStart = 0;
    var numRowsInRange = 0;
    
    for (var i = rangeRowStart; i <= values.length - 1; i++){
      var cell = values[i][0];
      if (typeof cell == "string" ){
        var cellID = cell.substring( cell.length-2, cell.length);
        if (cellID == "ID" ){
          
          if (numRowsInRange > 0){
            var numberColumnsInRange = values[i].length - 12;
            var moveRangeInA1Notation = scheduleSheet.getRange(rangeRowStart, 13, numRowsInRange, numberColumnsInRange ).getA1Notation();
            var startRangeInA1Notation = scheduleSheet.getRange (rangeRowStart, 6, 1, 1).getA1Notation();
            moveRanges.push ( moveRangeInA1Notation ); // column 13 is one week in
            startRanges.push ( startRangeInA1Notation ); //column 5 is start day
          }
          
          rangeRowStart = i+2;
          numRowsInRange = 0;
        }
      }
      else {
        numRowsInRange++;
        
        /*
        // Setup comments for all jobs and rsources for first day of first week
        var newFirstDayCell = values[i,12];
        var newFirstDayRange = scheduleSheet.getRange(i+1, 13)
        
        if ( newFirstDayRange.getComment().length == 0 && i>9 ) {
          newFirstDayRange.setComment( "Notes for the week:" )
        }
        */
  
      }
      
      
    }
           
    // Move all renges one week to left in studio overview
    for ( i in moveRanges) {
      var moveRange = moveRanges[i];
      var startRange = startRanges[i];
      scheduleSheet.getRange(moveRange).copyTo(scheduleSheet.getRange(startRange));
      
    }
    
    // Set the new start date in the variables sheet
    variablesSheet.getRange("startDate").setValue(newStartDateString); 

    // The code below shows a popup that disappears in 5 seconds
    SpreadsheetApp.getActiveSpreadsheet().toast("Last week was just deleted for your convienience", "", 5);
  }
  else {
    SpreadsheetApp.getActiveSpreadsheet().toast("Nothing to delete", "", 5);
  }
}


