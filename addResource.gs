

/**
 * addResource
 */
function addResource() {
  var app = UiApp.createApplication();
  app.add(app.loadComponent("addResourceGui"));
  app.setTitle("Add a resource");
  SpreadsheetApp.getActiveSpreadsheet().show(app);
};


function addResourceRespondToSubmit( e ) {
  SpreadsheetApp.getActiveSpreadsheet().toast("Adding the resource to the spreadsheet. Sometimes this can take a minute or so. Please be patient.", "Please Wait!", 10);

  var app = UiApp.getActiveApplication();
  
  var resourceName = e.parameter.resName;
  var resourceType = e.parameter.resType;
  var resourceLevel = e.parameter.resLvl;
  var resourceDept = e.parameter.resDept;
  
  // Error Checking
  
  var error = "";
  
  if (resourceType.toUpperCase() == "P" || resourceType.toUpperCase() == "F" || resourceType.toUpperCase() == "N" ) {
  }
  else {
    error = error + "Illegal type code. ";
  }

  if (resourceType.toUpperCase() != "N" ){
    if (resourceLevel > 0 && resourceLevel < 5 ) {
    }
    else {
      error = error + "Illegal level code. ";
    }

    if (resourceDept.toUpperCase() == "D" || resourceDept.toUpperCase() == "2D" || resourceDept.toUpperCase() == "3D" || resourceDept.toUpperCase() == "LA" || resourceDept.toUpperCase() == "E" || resourceDept.toUpperCase() == "T" || resourceDept.toUpperCase() == "ID" || resourceDept.toUpperCase() == "SD" || resourceDept.toUpperCase() == "P" || resourceDept.toUpperCase() == "H" ){
    }
    else {
      error = error + "Illegal department code. ";
    }
  }
  else {
    resourceLevel = ""; // Physpic items heve no level
  }
  
  if (error.length > 0) {
      SpreadsheetApp.getActiveSpreadsheet().toast( error, "Error", 5);
      app.close();
      return app;
  }
  
  //
  // Add the resource into the studio schedule
  //
  
  var resourceHeaderID = "";
  var nextResourceHeaderID = "";
  
  if (resourceDept.toUpperCase() == "D"){
    resourceHeaderID = "desID";
    nextResourceHeaderID = "3dID";
  }
  else if (resourceDept.toUpperCase() == "3D"){
    resourceHeaderID = "3dID";
    nextResourceHeaderID = "2dID";
  }
  else if (resourceDept.toUpperCase() == "2D"){
    resourceHeaderID = "2dID";
    nextResourceHeaderID = "laID";
  }
  else if (resourceDept.toUpperCase() == "LA"){
    resourceHeaderID = "laID";
    nextResourceHeaderID = "editID";
  }
  else if (resourceDept.toUpperCase() == "E"){
    resourceHeaderID = "editID";
    nextResourceHeaderID = "techID";
  }
  else if (resourceDept.toUpperCase() == "T"){
    resourceHeaderID = "techID";
    nextResourceHeaderID = "intDevID";
  }
  else if (resourceDept.toUpperCase() == "ID"){
    resourceHeaderID = "intDevID";
    nextResourceHeaderID = "softDevID";
  }
  else if (resourceDept.toUpperCase() == "SD"){
    resourceHeaderID = "softDevID";
    nextResourceHeaderID = "proID";
  }
  else if (resourceDept.toUpperCase() == "P"){
    resourceHeaderID = "proID";
    nextResourceHeaderID = "headID";
  }
  else if (resourceDept.toUpperCase() == "H"){
    resourceHeaderID = "headID";
    nextResourceHeaderID = "stuffID";
  }
  else {
    resourceHeaderID = "stuffID";
    nextResourceHeaderID = "endID";
 }

  var sheetVariables = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("variables");
  var sheetSchedule     = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Studio Schedule");

  var lastResourceIDRange = sheetVariables.getRange("lastResourceID");
  var lastResourceID = lastResourceIDRange.getValue();
  lastResourceID = lastResourceID+1;
  lastResourceIDRange.setValue(lastResourceID);

  var resourceTypeLevel = resourceType+resourceLevel;
  resourceTypeLevel = resourceTypeLevel.toUpperCase();
  
  //
  // Add the job into the studio schedule
  //
  
  var rows = sheetSchedule.getDataRange();
  var values = rows.getValues();
  
  var rangeRowStart = 0;
  var numRowsInRange = 0;
  
  for (var i = rangeRowStart; i <= values.length - 1; i++){
    var cell = values[i][0];
    if (typeof cell == "string" && cell == nextResourceHeaderID){

      sheetSchedule.insertRowAfter(i);
      var resourceDataForSchedule = [[(lastResourceID+""),resourceTypeLevel,resourceName]];
      var resourceRangeForSchedule = sheetSchedule.getRange(i+1, 1, 1, 3);
      resourceRangeForSchedule.setValues(resourceDataForSchedule);
      
      var resourceScheduleBGRange = sheetSchedule.getRange(i+1, 3, 1, 1);
      
      if ( resourceType.toUpperCase() == "P") {
        resourceScheduleBGRange.setFontWeight("bold");
        if ( resourceLevel == 4 ){
          resourceScheduleBGRange.setBackground("#00ff00");
        }
        else if ( resourceLevel == 3 ){
          resourceScheduleBGRange.setBackground("#b6d7a8");
        }
        else if ( resourceLevel == 2 ){
          resourceScheduleBGRange.setBackground("#d9ead3");
        }
        else{
          resourceScheduleBGRange.setBackground("#ffff00");
        }
      }
      else if ( resourceType.toUpperCase() == "F" ){
        resourceScheduleBGRange.setFontWeight("normal");
          if ( resourceLevel == 4 ){
          resourceScheduleBGRange.setBackground("#00ffff");
        }
        else if ( resourceLevel == 3 ){
          resourceScheduleBGRange.setBackground("#a4c2f4");
        }
        else if ( resourceLevel == 2 ){
          resourceScheduleBGRange.setBackground("#c9daf8");
        }
        else{
          resourceScheduleBGRange.setBackground("#ffff00");
        }
      }
      else{
        resourceScheduleBGRange.setBackground("#efefef");
        resourceScheduleBGRange.setFontWeight("normal");
      }
      
      var nameRange = sheetSchedule.getRange(i+1, 3, 1, 3);
      nameRange.merge();
    }
  }

  
  // Do some cleanup
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Ready to go!", "Woot!", 10);

  
  app.close();
  return app;


}
