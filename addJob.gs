
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
  
  cleanUp();
  app.close();
  return app;
};
