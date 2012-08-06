
/**
 * delete last week if a week has passed
 */
function checkForWeekStart() {
  // Get data about spreadsheet
  var startDateRange = "A2:A2";
 
  var variablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var weekStart = variablesSheet.getRange(startDateRange).getValues();    // startDate is in ranges list
  
  var oldStartDate = new Date(weekStart);
  var oldStartTime = oldStartDate.getTime();
  var todayTime = new Date().getTime();
  
  var oneDayMS = 86400000;
  
  if ( todayTime-oldStartTime > (9*oneDayMS)) {   // See if a week and 2 days in milliseconds has passed (we do this on a wednesday)
  
      // Set the new starting monday
      var newStartDate = new Date(oldStartDate.getTime() + (7*oneDayMS) );
      var newStartDateString = newStartDate.getDate()+"/"+(newStartDate.getMonth()+1)+"/"+newStartDate.getFullYear()  // Javascript months are 0-11
      variablesSheet.getRange(startDateRange).setValue(newStartDateString);  

      var jobRange         = "L11:CX60";
      var resDesignRange   = "L62:CX81";
      var res3dRange       = "L83:CX102";
      var res2dRange       = "L104:CX123";
      var resEditRange     = "L125:CX144";
      var resTechRange     = "L146:CX165";
      var resIntRange      = "L167:CX186";
      var resDevRange      = "L188:CX207";
      var resProducerRange = "L209:CX228";
      var resHeadRange     = "L230:CX249";

      var jobStart         = "E11:E11";
      var resDesignStart   = "E62:E57";
      var res3dStart       = "E83:E78";
      var res2dStart       = "E104:E99";
      var resEditStart     = "E125:E120";
      var resTechStart     = "E146:E141";
      var resIntStart      = "E167:E162";
      var resDevStart      = "E188:E183";
      var resProducerStart = "E209:E204";
      var resHeadStart     = "E230:E225";


      var moveRanges = [jobRange,resDesignRange,res3dRange,res2dRange,resEditRange,resTechRange,resIntRange,resDevRange,resProducerRange,resHeadRange];
      var startRanges = [jobStart,resDesignStart,res3dStart,res2dStart,resEditStart,resTechStart,resIntStart,resDevStart,resProducerStart,resHeadStart];
           
      // Move all renges one week to left in studio overview
      for ( i in moveRanges) {
        var moveRange = moveRanges[i];
        var startRange = startRanges[i];
        
        var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
        scheduleSheet.getRange(moveRange).copyTo(scheduleSheet.getRange(startRange));
      }

      // The code below shows a popup that disappears in 5 seconds
      SpreadsheetApp.getActiveSpreadsheet().toast("Last week was just deleted for your convienience", "", 5);
    }
};
