/**
Things to do:

Delete Item function
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
  menuEntries.push({ name : "Clean Up", functionName : "cleanupViaMenu"});
  
  sheet.addMenu("Functions", menuEntries);
  
  cleanupOnOpen();
};
