/**
 * create month names and dates
 */

function createDateHeaders() {
  var variablesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var scheduleSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  var weekStart = variablesSheet.getRange("startDate").getValues();    // startDate is in ranges list
  
  var startDate = new Date(weekStart);
  var startTime = startDate.getTime();
  var timeList = new Array();
  var dayList = new Array();
  var monthList = new Array();
  var oneDayMS = 86400000;

  for (var i = 0; i <= 250; i++) {
    time = startTime + i*oneDayMS;
    timeList.push ( time );
    var nextDate = new Date (time);
    dayList.push  ( nextDate.getDate() );
    
    var monthNumber = nextDate.getMonth();
    
    if (monthNumber == 0) { monthList.push ("JANUARY"); }
    
    else if (monthNumber == 1) { monthList.push ("FEBRUARY");}

    else if (monthNumber == 2) { monthList.push ("MARCH"); }

    else if (monthNumber == 3) { monthList.push ("APRIL"); }

    else if (monthNumber == 4) { monthList.push ("MAY"); }

    else if (monthNumber == 5) { monthList.push ("JUNE"); }

    else if (monthNumber == 6) { monthList.push ("JULY"); }

    else if (monthNumber == 7) { monthList.push ("AUGUST"); }

    else if (monthNumber == 8) { monthList.push ("SEPTEMBER"); }

    else if (monthNumber == 9) { monthList.push ("OCTOBER"); }

    else if (monthNumber == 10) { monthList.push ("NOVEMBER"); }

    else { monthList.push ("DECEMBER"); }
  }

  // Get the cells so we can write to them.
  
  var dateNumberRange = scheduleSheet.getRange(10, 6, 1, 251); // first date number F10
  dayList = [dayList];
  dateNumberRange.setValues ( dayList );
  
  var dateMonthRange = scheduleSheet.getRange(9, 6, 1, 251); // first month number F9
  monthList = [monthList];
  dateMonthRange.setValues (monthList);
  
}
