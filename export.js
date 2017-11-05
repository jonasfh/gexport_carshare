function onOpen() {
  var menuEntries = 
  [ 
    {
      name: "Exporter Autopass filer",
      functionName: "autopass_export"
    }
  ];
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.addMenu("Faktura-export",menuEntries);
}

/**
 * Convert Excel file to Sheets
 * @param {Blob} excelFile The Excel file blob data; Required
 * @param {String} filename File name on uploading drive; Required
 * @return {Spreadsheet} File object pointing to converted Google Spreadsheet
 **/
function convertExcel2Sheets(excelFile, filename) {

  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  
  var uploadParams = {
    method:'post',
    contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
    contentLength: excelFile.getBytes().length,
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: excelFile.getBytes()
  };

  // Upload file to Drive root folder and convert to Sheets
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);

  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());

  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename,
  };
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };

  // Update metadata (filename and parent folder(s)) of converted sheet
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);

  return DriveApp.getFileById(fileDataResponse.id);
}


function autopass_export(){
  var sh = SpreadsheetApp.getActiveSheet();
  // Folder Bildeleringen/letsgo/autopass
  var folder = DriveApp.getFolderById(getProperty('autopass'));
  // Folder Bildeleringen/letsgo/autopass_v1/JSON
  var json_folder = DriveApp.getFolderById(getProperty('json'));
  var files = folder.getFiles();
  var num_files = 0;
  while (files.hasNext()){
    var file = files.next();
    if (file.getMimeType() == 'application/vnd.ms-excel') {
      var gsfile = null;
      if (!folder.getFilesByName(file.getName() + '.gsheet').hasNext()) {
        gsfile = convertExcel2Sheets(file.getBlob(), file.getName());
        folder.addFile(gsfile);
        DriveApp.removeFile(gsfile);
        gsfile.setName(gsfile.getName() + '.gsheet');
      }
      else {
        gsfile = folder.getFilesByName(file.getName() + '.gsheet').next()
      }
      if(!json_folder.getFilesByName(file.getName() + '.json.txt').hasNext()) {
        var data = autopass_JSON_convert(gsfile);
        var output = DriveApp.createFile(file.getName() + '.json.txt', JSON.stringify(data, null, '\t'), 'application/json');
        json_folder.addFile(output);
        DriveApp.removeFile(output);
        num_files ++;
      }
    }
  }
  var show_num = num_files;
  if (num_files < 10) {
    var a = ['Null','En','To','Tre','Fire','Fem','Seks','Syv','Åtte','Ni'];
    show_num = a[num_files];
  }
  SpreadsheetApp.getUi().alert(show_num + ' JSON-filer generert uten feil');
}

function autopass_JSON_convert(file) {
  // if (typeof file == 'undefined') file = DriveApp.getFileById('1SHCjnEDgfSKtlQZ3HBp-jf-Es0IiJnieD78'); // default value
  var spreadsheet = SpreadsheetApp.open(file);
  var sheet = spreadsheet.getSheetByName('Sheet1');
  data = [];
  var row = 12;
  var values = sheet.getRange(row, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if((v[0]) == 'Antall passeringer:') break;
    var d = v[1];
    var date = d.substr(6,4) + '-' + d.substr(3,2) + '-' +d.substr(0,2) + ' ' + d.substr(12,5);
    var reg_id = v[4];
    var amount = v[3];
    var comment = v[2];
    data.push({date: date, reg_id: reg_id, amount: amount, comment: comment})
  }
  return data;
}
