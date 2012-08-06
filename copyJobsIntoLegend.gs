/**
 * Hide empty lines in studio schedule
 */
function copyJobsIntoLegend() {
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
  var jobColourColID = 5;
  var jobNumberColID = 6;
  var jobDescColID = 12;

  for ( i in jobNumber ) {
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setValue("X");
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setBackgroundColor(jobColour[i]);
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setFontColor (jobFontColour[i]);

    resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setValue(jobNumber[i]);
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setValue(jobDesc[i]);

    resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setFontColor("#000000");
    resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setFontColor("#000000");
    
    if (jobDept[i] == "S"){
      resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#cfe2f3");
      resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#cfe2f3");
    }
    else if (jobDept[i] == "I"){
      resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#fff2cc");
      resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#fff2cc");
    }
    else {
      resourceOverviewSheet.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#d9ead3");
      resourceOverviewSheet.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#d9ead3");
     }

    rowCount++;
    
    if (rowCount >= 5){
      jobNumberColID = jobNumberColID+14;
      jobColourColID = jobColourColID+14;
      jobDescColID = jobDescColID+14;
      
      rowCount = 0;
    }
  }
}