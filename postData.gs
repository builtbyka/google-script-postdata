var SCRIPT_PROP = PropertiesService.getScriptProperties(); // new property service

function doPost(e){
  return handleResponse(e, type);
}
 
function handleResponse(e, type) {

  var lock = LockService.getPublicLock();
  lock.waitLock(30000);  // wait 30 seconds before conceding defeat.
   
  try {
    var id = e.parameters.id; // get spreadsheet key from ajax call
    var sheetID = e.parameters.sheet; // get sheetname from ajax call
    var doc = SpreadsheetApp.openById(id); // open the google spreadsheet
    var sheet = doc.getSheetByName(sheetID); // open the google sheet
    
    //if it doesn't exist create from master-sheet
    if(sheet === null){
        var copySheet = doc.getSheetByName('master-sheet');
        var headerCopy = copySheet.getRange(1, 1, 1, copySheet.getLastColumn());
        doc.insertSheet(String(sheetID));
        var sheet = doc.getSheetByName(String(sheetID));
        var headerInsert =  sheet.getRange(1, 1, 1, 1);
        headerCopy.copyTo(headerInsert);
    }
    
    var gid = sheet.getSheetId(); // get unique ID of sheet
    var sheetData = sheet.getDataRange().getValues(); //get all the values in the sheet already
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; // get the header values - this is important as it knows from this where to inject the data
    var rows = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    var injectRow = '';
    injectRow = sheet.getLastRow() + 1; // last populated row + 1
      var row = [];
    
      for (i in headers){
          if(e.parameter[headers[i]] === undefined){
            row.push(''); 
          }else{
              row.push(e.parameter[headers[i]]); //creates an array of values that match headers
          }
      } 
      
      sheet.getRange(injectRow, 1, 1, row.length).setValues([row]) // set the values of the first available row with the array
      
      // return json success results
      return ContentService
      .createTextOutput(JSON.stringify({"result":"success", "row": injectRow}))
      .setMimeType(ContentService.MimeType.JSON);
    
    } catch(e){
      // if error return this
      return ContentService
      .createTextOutput(JSON.stringify({"result":"error", "error": e}))
      .setMimeType(ContentService.MimeType.JSON);
    } finally { //release lock
      lock.releaseLock(); 
    }

} 

function convert(value){
  if (value === true) return 'True';
  if (value === false) return 'False';
  return value;
}
 
function setup() {
   var doc = SpreadsheetApp.getActiveSpreadsheet();
   SCRIPT_PROP.setProperty("key", doc.getId());
}
