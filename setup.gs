global_setup = {};
function getProperty(name) {
  if (typeof name == 'undefined') {
    name = 'json';
  }
  if (Object.keys(global_setup).length === 0 && global_setup.constructor === Object) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETUP");
    var range = sheet.getRange("setup").getValues();
    for (var i=0; i<range.length; i++) {
      if (range[i][0] != "") {
        global_setup[range[i][0]] = range[i][1];       
      }
    }
  }
  return global_setup[name];
}
