/**
 * Hide empty lines in studio schedule
 */
function copyJobsIntoLegend() {
  var sheetSchedule = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var rows = sheetSchedule.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

  var jobColour = new Array();
  var jobFontColour = new Array();
  var jobFontWeight = new Array();
  var jobDept = new Array();
  var jobStatus = new Array();
  var jobNumber = new Array();
  var jobDesc = new Array();
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
          jobDept.push( values[i][1].substring(0,1) );
          jobStatus.push( values[i][1].substring(1,2) );
          jobNumber.push( values[i][3] );
          jobDesc.push(values[i][4] );
          
          numberOfJobs++;
          
      }
    }
  }
  
  
  // Set remaining job slots to nothing in colour and value
  if (numberOfJobs < 50) {
    for ( i = 0; i < 50 - numberOfJobs; i++) {
          jobColour.push( "#ffffff" );
          jobDept.push( "" );
          jobStatus.push( "" );
          jobNumber.push( "" );
          jobDesc.push( "" );
    }
  }
  
  // Now we have our jobs in an array we can write them to the resource OverView
  
  var rowCount = 0;
  var firstRowID = 4
  var jobColourColID = 6;
  var jobNumberColID = 7;
  var jobDescColID = 13;

  for ( i in jobNumber ) {
    sheetSchedule.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setValue("X");
    sheetSchedule.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setBackgroundColor(jobColour[i]);
    sheetSchedule.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setFontColor (jobFontColour[i]);
    sheetSchedule.getRange((firstRowID+rowCount), jobColourColID, 1, 1).setFontWeight (jobFontWeight[i]);

    sheetSchedule.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setValue(jobNumber[i]);
    sheetSchedule.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setValue(jobDesc[i]);

    sheetSchedule.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setFontColor("#000000");
    sheetSchedule.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setFontColor("#000000");
    
    if (jobDept[i] == "S"){
      sheetSchedule.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#cfe2f3");
      sheetSchedule.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#cfe2f3");
    }
    else if (jobDept[i] == "I"){
      sheetSchedule.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#fff2cc");
      sheetSchedule.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#fff2cc");
    }
    else {
      sheetSchedule.getRange((firstRowID+rowCount), jobNumberColID, 1, 1).setBackgroundColor("#d9ead3");
      sheetSchedule.getRange((firstRowID+rowCount), jobDescColID, 1, 1).setBackgroundColor("#d9ead3");
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
