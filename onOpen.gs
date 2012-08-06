/**
Things to do:

Delete Item function
Implement  colourID into addJob
upgrade form 40 job system to 50 jobs
reporting system
sorting system

*/


/**
 * do this on opening spreadsheet
 */
function onOpen() {

  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  
  menuEntries.push({ name : "Add Job", functionName : "addJob"});
  menuEntries.push({ name : "Add Resource", functionName : "addResource"});
  menuEntries.push({ name : "Delete Item", functionName : "deleteItem"});
  menuEntries.push({ name : "Clean Up", functionName : "cleanUp"});

  sheet.addMenu("Functions", menuEntries);
  
  cleanUp();
};