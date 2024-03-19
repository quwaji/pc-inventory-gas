/**
 * PC Information gathering into Spreadsheet
 * 
 */

function doPost(e) {
  Logger.log('post received: ' + e.postData.contents);
  recordRequest(e);
  return ContentService.createTextOutput(JSON.stringify({"success":true})).setMimeType(ContentService.MimeType.JSON);
}

function recordRequest(e) {

  // Parse JSON
  if (e == null || e.postData == null || e.postData.contents == null) {
    return;
  }
  var requestJSON = e.postData.contents;
  var requestObj = JSON.parse(requestJSON);
  var type = '';
  var userId = '';
  var groupId = '';
  var messageText = '';
  // if(requestObj.events.length > 0) {
  //   type = requestObj.events[0].type;
  //   if(type == 'message' || type == 'follow' || type == 'unfollow') {
  //     userId = requestObj.events[0].source.userId;
  //   }
  //   if(type == 'message') {
  //     messageText = requestObj.events[0].message.text;
  //   }
  //   if(requestObj.events[0].source.type == 'group') {
  //     groupId = requestObj.events[0].source.groupId;
  //   }
  // }

  date = new Date();

  //  
  // Record to sheet
  //

  var ss = SpreadsheetApp.getActive()
  var sheet = ss.getSheetByName('Post Log');

  // Get header definintion
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];

  // Get data for each headers
  var values = [];
  for (i in headers){
    var header = headers[i];
    var val = "";
    switch(header) {
      case "date":
        val = new Date();
        break;
      case "body":
        val = e.postData.contents;
        break;
      default:
        val = requestObj[header];
        if (val == undefined) {
          val = "";
        }
        break;
    }
    values.push(val);
  }

  // Add row
  sheet.appendRow(values);
}
