
/**
 * addJob
 */
function addJob() {
  SpreadsheetApp.getActiveSpreadsheet().toast("Getting ready. Please be patient", "Hey there!", 10);

  var app = UiApp.createApplication();
  app.add(app.loadComponent("addJobGui"));
  app.setTitle("Add a job");
  SpreadsheetApp.getActiveSpreadsheet().show(app);
};


/**
 * addJobRespondToSubmit
 */
function addJobRespondToSubmit( e ) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Adding the job to the spreadsheet. Sometimes this can take a minute or so, escpecially with many active users. Please be patient.", "Please Wait!", 120);

  var app = UiApp.getActiveApplication();
  
  var jobNumber = e.parameter.jobNumber;
  var jobName = e.parameter.jobName;
  var statusCode = e.parameter.statusCode;
  var deptCode = e.parameter.deptCode;
  var jobProducer = e.parameter.jobProducer;
  
  
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

  if (deptCode.toUpperCase() == "S" || deptCode.toUpperCase() == "I" || deptCode.toUpperCase() == "E" ) {
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
  
  statusCode = deptCode+statusCode; // make the status code two characters for the data column
 
  var sheetVariables = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var sheetSchedule     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var lastJobIDRange = sheetVariables.getRange("lastJobID");
  var lastJobID = lastJobIDRange.getValue();
  lastJobID = lastJobID+1;
  lastJobIDRange.setValue(lastJobID);

  /*
  // Get colour via ID umber on variable page
  var lastColourIDRange = sheetVariables.getRange("lastColourID");
  var lastColourID = lastColourIDRange.getValue();
  lastColourID = lastColourID+1;
  if (lastColourID > 50) { lastColourID = 1; };
  lastColourIDRange.setValue(lastColourID);

  var jobColourBG = sheetVariables.getRange(lastColourID+2, 6).getBackgroundColor();
  var jobColourFG = sheetVariables.getRange(lastColourID+2, 6).getFontColor();
  var jobWeight = sheetVariables.getRange(lastColourID+2, 6).getFontWeight();  
  */
  
  // Get colour based on next unused colour on variables page
  var jobSwatch = ["#00ff00","#ffffff","normal"];
  jobSwatch = getNextJobColour();
  var jobColourBG = jobSwatch[0];
  var jobColourFG = jobSwatch[1];
  var jobWeight = jobSwatch[2];

  


  jobNumber = jobNumber.toUpperCase();
  statusCode = statusCode.toUpperCase();
  
  //
  // Add the job into the studio schedule
  //
  
  var rows = sheetSchedule.getDataRange();
  var values = rows.getValues();
  
  var rangeRowStart = 0;
  var numRowsInRange = 0;
  
  for (var i = rangeRowStart; i <= values.length - 1; i++){
    var cell = values[i][0];
    if (typeof cell == "string" ){
      if (cell == "desID" ){
      
        sheetSchedule.insertRowAfter(i);
        var jobDataForSchedule = [[(lastJobID+""),statusCode,"X",(jobNumber+" ("+jobProducer+")"),jobName]];
        var jobsRangeForSchedule = sheetSchedule.getRange(i+1, 1, 1, 5);
        jobsRangeForSchedule.setValues(jobDataForSchedule);

        var jobScheduleColourCell = sheetSchedule.getRange(i+1,3);
        
        jobScheduleColourCell.setBackground(jobColourBG);
        jobScheduleColourCell.setFontColor(jobColourFG);
        jobScheduleColourCell.setFontWeight(jobWeight);
        
        var jobScheduleBGRange = sheetSchedule.getRange(i+1, 4, 1, 2);

        if ( statusCode == "EA" || statusCode == "SA" || statusCode == "IA" ) {
          jobScheduleBGRange.setBackground("#d9d2e9");
          jobScheduleBGRange.setFontWeight("normal");
        }

        else if ( statusCode == "EP" || statusCode == "SP" || statusCode == "IP" ) {
          jobScheduleBGRange.setBackground("#f4cccc");          
          jobScheduleBGRange.setFontWeight("normal");

        }

        else {
          jobScheduleBGRange.setBackground("#efefef");
          jobScheduleBGRange.setFontWeight("bold");
        }
      }
    }
  }
  
  // Do some cleanup

  copyJobsIntoLegend();
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Ready to go!", "Woot!", 10);

  
  app.close();
  return app;
};


/**
 * getNextJobColour
 */
function getNextJobColour() {
  var sheetVariables = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var sheetSchedule     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var rows = sheetSchedule.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var jobColour = new Array ();
  var jobFontColour = new Array ();
  var jobFontWeight = new Array ();
    
  var numberOfJobs = 0;

  //iterate through the jobs and add the job colour and parametres to some arrays.
  for (var i = 1; i <= values.length - 1; i++) {
    var cell = values[i][0];
    if (typeof cell != "string" ) {
      if (cell >= 10000) {
          var dataRange = sheetSchedule.getRange(i+1, 3, 1, 1); // Get the range of the job to pick the job colour from
          
          jobColour.push( dataRange.getBackgroundColor() );
          jobFontColour.push( dataRange.getFontColor() );
          jobFontWeight.push( dataRange.getFontWeight() );
          
          numberOfJobs++;
      }
    }
  }


  var vRows = sheetVariables.getDataRange();
  var vNumRows = vRows.getNumRows();
  var vValues = vRows.getValues();

  var variableColour = new Array ();
  var variableFontColour = new Array ();
  var variableFontWeight = new Array ();
  
  //iterate through the variables and add the job colour and parametres to some arrays.
  for (var i = 1; i <= 50 - 1; i++) {
    var cell = values[i+1][5];
      var dataRange = sheetVariables.getRange(i+2, 6, 1, 1); // Get the range of the job to pick the job colour from
      
      variableColour.push( dataRange.getBackgroundColor() );
      variableFontColour.push( dataRange.getFontColor() );
      variableFontWeight.push( dataRange.getFontWeight() );

  }

  var jobSwatch = ["#00ff00","#ffffff","normal"];
  var colourMatch = 0;
  
  for (var whichColour = 0; whichColour <=50-1 ;whichColour++){
    for (var whichJob = 0; whichJob<=numberOfJobs-1;whichJob++){
      if ( variableColour [whichColour] == jobColour [whichJob] && variableFontColour [whichColour] == jobFontColour [whichJob] && variableFontWeight [whichColour] == jobFontWeight [whichJob]){
        colourMatch = 1;
      }    
    }
    if (colourMatch == 0){
        jobSwatch[0] = variableColour [whichColour];
        jobSwatch[1] = variableFontColour [whichColour];
        jobSwatch[2] = variableFontWeight [whichColour];
        return jobSwatch;
    }
    else{
      colourMatch = 0;
      jobSwatch = ["#00ff00","#ffffff","normal"];
    }
  }
  return jobSwatch;
}
