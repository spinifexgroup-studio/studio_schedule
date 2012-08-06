/**
Things to do:

Delete Item function
Implement colourID into addJob
upgrade form 40 job system to 50 jobs
reporting system
sorting system

*/


/**
 * Fix Illegal edits
 */
 
function onEdit(event)
{
  var activeSheet = event.source.getActiveSheet();
  var activeRange = event.source.getActiveRange();
  
  if ( activeSheet.getName() == "Studio Schedule" ) {
    var firstColumn = activeRange.getColumn();
    var firstRow = activeRange.getRow();
    
    
    // Check if noneditable areas are being edited and warn
    if (firstColumn <= 4 || firstRow <= 10) {
        Browser.msgBox("You should only be editing the bookings squares. Naughty, naughty!");
    }
    if (firstRow == 61 || firstRow == 82 || firstRow == 103 || firstRow == 124 || firstRow == 145 || firstRow == 166 || firstRow == 187 || firstRow == 208|| firstRow == 229 ) {
        Browser.msgBox("You should only be editing the bookings squares. Naughty, naughty!");
    }
    
    // Check to see if job section was edited and colourise Cs and Hs
    
    
    if ( firstRow > 10 && firstRow < 61 && firstColumn > 4) {
      var values = activeRange.getValues();
      var cell = values[0][0];
      
      if ( cell == "C" || cell == "C" || cell == "C" || cell == "C" ){
      
        if (jobType == "I" ) {
          activeRange.setBackgroundColor("#ffd966");
        }

        else if (jobType == "S" ) {
          activeRange.setBackgroundColor("#6fa8dc");
        }

        else if (jobType == "SI" ) {
          activeRange.setBackgroundColor("#38761d");
        }
      }
    }
   
    
  }
}


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


/**
 * Hide empty lines in studio schedule
 */
function copyJobsIntoResourceOverview() {
  var jobsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("jobs");
  var resourceOverviewSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var rows = jobsSheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var jobColour = new Array();
  var jobFontColour = new Array();
  var jobDept = new Array();
  var jobNumber = new Array();
  var jobDesc = new Array();
  
  //iterate through the jobs and add the job colour and parametres to some arrays.
  for (var i = 10; i <= values.length - 1; i++) {
    var cell = values[i][0];
    if (cell.length != 0) {
        var dataRange = jobsSheet.getRange(i+1, 2, 1, 1); // Get the range of the job to pick the job colour from
        
        jobColour.push( dataRange.getBackgroundColor() );
        jobFontColour.push( dataRange.getFontColor() );
        jobDept.push( values[i][3] );
        jobNumber.push( values[i][4] );
        jobDesc.push(values[i][5] );
    }
  }
  
  
  // Set remaining job slots to nothing in colour and value
  if (jobNumber.length < 40) {
    for ( i = 0; i < 40 - jobNumber.length; i++) {
          jobColour.push( "#ffffff" );
          jobDept.push( "" );
          jobNumber.push( "" );
          jobDesc.push( "" );
    }
  }
  
  // Now we have our jobs in an array we can write them to the resource OverView
  
  var resourceRows = resourceOverviewSheet.getDataRange();
  var resourceValues = resourceRows.getValues();

  var rowCount = 0;
  var firstRowID = 4
  var jobDeptColID = 5;
  var jobColourColID = 6;
  var jobNumberColID = 7;
  var jobDescColID = 12;

  for ( i in jobNumber ) {
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobDeptColID, 1, 1).setValue(jobDept[i]);
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setValue("X");
    
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setValue(jobNumber[i]);
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setValue(jobDesc[i]);
    
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setBackgroundColor(jobColour[i]);
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setFontColor (jobFontColour[i]);
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#d9ead3");
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#d9ead3");

    rowCount++;
    
    if (rowCount >= 5){
      jobNumberColID = jobNumberColID+14;
      jobColourColID = jobColourColID+14;
      jobDeptColID = jobDeptColID+14;
      jobDescColID = jobDescColID+14;
      
      rowCount = 0;
    }
  }
}


/**
 * Hide empty lines in studio schedule
 */
function hideEmptyRows() {

  var startRow = 10;

  // Get data about spreadsheet
 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();
  
  
  //sheet.showColumns(1, sheet.getLastColumn());   // Show all the rows
  sheet.showRows(1, sheet.getLastRow());
  
  // Loop through all the rows

  for (var i = startRow; i <= values.length - 1; i++) {
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
  }
  sheet.hideColumns(1);   // Hide the ID column
};


/**
 * addResource
 */
function addResource() {
};


/**
 * addJob
 */
function addJob() {
  var app = UiApp.createApplication();
  app.add(app.loadComponent("addJobGui"));
  app.setTitle("Add a job");
  SpreadsheetApp.getActiveSpreadsheet().show(app);
};


/**
 * addJob
 */
function addJobRespondToSubmit( e ) {
  var app = UiApp.getActiveApplication();
  
  var jobNumber = e.parameter.jobNumber;
  var jobName = e.parameter.jobName;
  var statusCode = e.parameter.statusCode;
  var deptCode = e.parameter.deptCode;
  
  
  // Error Checking
  
  var error = "";
  
  if (jobNumber.length < 8 || jobNumber.length > 9) {
    error = error + "Job Number length must be 8 or 9 characters in length. ";
  }

  if (statusCode.toUpperCase() == "L" || statusCode.toUpperCase() == "A" || statusCode.toUpperCase() == "P" ) {
  }
  else {
    error = error + "Illegal status code. ";
  }

  if (deptCode.toUpperCase() == "S" || deptCode.toUpperCase() == "I" || deptCode.toUpperCase() == "SI" ) {
  }
  else {
    error = error + "Illegal department code. ";
  }
  
  if (error.length > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast( error, "Error", 5);
      app.close();
      return app;
  }
  
  // Good to go - add job to job sheet
 
  var sheetVariables = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var sheetJobs     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("jobs");

  var lastJobIdRange = sheetVariables.getRange("B2:B2");
  var lastJobId = lastJobIdRange.getValue();
  lastJobId = lastJobId+1;
  lastJobIdRange.setValue(lastJobId);

  jobNumber = jobNumber.toUpperCase();
  statusCode = statusCode.toUpperCase();
  deptCode = deptCode.toUpperCase();
  var jobData = [[(lastJobId+""),"",statusCode,deptCode,jobNumber,jobName]];
  var lastJobRow = sheetJobs.getLastRow();
  var jobsRange = sheetJobs.getRange(lastJobRow+1, 1, 1, 6);
  jobsRange.setValues(jobData);
  
  // Do some cleanup
  
  SpreadsheetApp.getActiveSpreadsheet().toast( "Give me a few seconds to clean up please.", "Hey there!", 5);
    
  colouriseJobSchedule();
  copyJobsIntoResourceOverview();
  hideEmptyRows();
  
  SpreadsheetApp.getActiveSpreadsheet().toast( "I'm ready to go.", "Cool!", 5);

  app.close();
  return app;
};



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


/**
 * perform cleanup functions if things are out of order
 */
function cleanUp() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Please don't do anything until it's ok to start working", "", 5);

  colouriseJobSchedule();
  copyJobsIntoResourceOverview();
  hideEmptyRows();
  checkForWeekStart();

  SpreadsheetApp.getActiveSpreadsheet().toast("Ok to start working", "", 5);
};

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



/**
 * do this on opening spreadsheet
 */
function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Please don't do anything until it's ok to start working!", "Warning", 5);

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  
  menuEntries.push({ name : "Add Job", functionName : "addJob"});
  menuEntries.push({ name : "Add Resource", functionName : "addResource"});
  menuEntries.push({ name : "Delete Item", functionName : "deleteItem"});
  menuEntries.push({ name : "Clean Up", functionName : "cleanUp"});

  sheet.addMenu("Functions", menuEntries);
  
  colouriseJobSchedule();
  copyJobsIntoResourceOverview();
  hideEmptyRows();
  checkForWeekStart();
  SpreadsheetApp.getActiveSpreadsheet().toast("Ok to start working!", "Hey there!", 5);

};