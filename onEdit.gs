/**
 * Fix Illegal edits
 */
 /*
function onEdit(event)
{
  var activeSheet = event.source.getActiveSheet();
  var activeRange = event.source.getActiveRange();
  
  if ( activeSheet.getName() == "Studio Schedule" ) {
    var firstColumn = activeRange.getColumn();
    var firstRow = activeRange.getRow();
    
    
    // Check if noneditable areas are being edited and warn
    if (firstColumn <= 4 || firstRow <= 10) {
        Browser.msgBox("You should only be editing the bookings squares. Naughty, naughty! Unless you're a ninja, you should probably undo what ever it is you just did....");
    }
    if (firstRow == 61 || firstRow == 82 || firstRow == 103 || firstRow == 124 || firstRow == 145 || firstRow == 166 || firstRow == 187 || firstRow == 208|| firstRow == 229 ) {
        Browser.msgBox("You should only be editing the bookings squares. Naughty, naughty! Unless you're a ninja, you should probably undo what ever it is you just did...");
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
*/
