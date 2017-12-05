// Structure with the global setup
global_setup = {};

// Called when the functionality is installed
function onInstall(e) {
  onOpen(e);
}

// Called when the spreadsheet is opened
function onOpen(e) {
  var menuEntries = 
  [ 
    {
      name: "Åpne eksportverktøy",
      functionName: "openSidebar"
    }
  ];
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  sheet.addMenu("Faktura-eksport",menuEntries);
}

/*
* Opens the html sidebare, with a nice ui for exporting files
*/
function openSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('sidebar.html');
  SpreadsheetApp.getUi().showSidebar(html);  
}

/*
* Get the for a folder with a given, symbolic name, used in the setup.
* Takes strings as input, and returns an array of {url, name} - objects.
*/
function getFolderUrls() {
  var output = [];
  for (var i = 0; i < arguments.length; i++) {
    var folder = DriveApp.getFolderById(getProperty(arguments[i]));
    output.push({url:folder.getUrl(), name: folder.getName()});
  }
  return output;
}

/*
* List files that are not already exported.
*/
function listUhandledFiles() {
  // Folder Bildeleringen/letsgo/autopass
  var folder = DriveApp.getFolderById(getProperty('autopass'));
  
  // Folder Bildeleringen/letsgo/autopass_v1/JSON
  var json_folder = DriveApp.getFolderById(getProperty('json'));
  var files = folder.getFiles();
  var outfiles = [];
  while (files.hasNext()){
    var file = files.next();
    if (file.getName().substr(-4) == '.xls') {
      var gsfile = null;
      if (file.getMimeType() == 'application/vnd.google-apps.spreadsheet') {
          gsfile = file;
      }
      else if (file.getMimeType() == 'application/vnd.ms-excel') {
        if (!folder.getFilesByName(file.getName() + '.gsheet').hasNext()) {
          gsfile = convertExcel2Sheets(file.getBlob(), file.getName());
          folder.addFile(gsfile);
          DriveApp.removeFile(gsfile);
          gsfile.setName(gsfile.getName() + '.gsheet');
        }
        else {
          gsfile = folder.getFilesByName(file.getName() + '.gsheet').next();
        }
      }
      if(!json_folder.getFilesByName(file.getName() + '.json.txt').hasNext()) {
        outfiles.push({"name":gsfile.getName(), "id":gsfile.getId()});
      }
    }
  }
  return outfiles;
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

/*
* Get errors and statistics revealed during the export.
*/
function report_basic_stats(obj) {
  // UI Spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STATS");
  if (sheet == null) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("STATS");
    sheet.getRange(1,1).setValue("Statistikk for konverterte filer").setFontSize(36);
    var s = "Nedenfor finner du statistikk for filer som er konvertert med verktøyet. Nyeste filer øverst.";
    sheet.getRange(2,1).setValue(s);
    sheet.autoResizeColumn(1);
  }
  sheet.getParent().setActiveSheet(sheet);
  var startAt = 5;
  // Calculate number of rows
  var numrows = obj.length * 5 + 4;
  for(i = 0; i<obj.length; i++) {
    numrows += obj[i].errors.length + 1;
  }
  sheet.insertRowsBefore(startAt, numrows);
  startAt++;
  for (i = 0; i < obj.length; i++) {
    var range = sheet.getRange(startAt, 1, 3, 6);
    range.setValues(
      [
        ["Eksportert dato: " + new Date().toString().substr(4, 21), "Filnavn", obj[i].file_name, "", "", ""],
        ["Ant. Linjer", "Største beløp", "Ant. 0 - beløp", "Ant rader OK", "Første tidspunkt", "Siste tidspunkt"], 
        [obj[i].num_lines, obj[i].max_amount, obj[i].num_zeros, obj[i].num_lines_ok, 
            df(obj[i].start, 'j. M kl H:i'), df(obj[i].end, 'j. M kl H:i')],
      ]
    );
    sheet.getRange(startAt, 1, 1, 3).setFontWeight("bold").setFontSize(14).setBackgroundRGB(252, 251, 224);
    if (obj[i].errors.length > 0) {
      sheet.getRange(startAt + 3, 1, 1, 1).setValue("Feilmeldinger").setFontWeight("bold");
    }
    else {
      sheet.getRange(startAt + 3, 1, 1, 1).setValue("Ingen feilmeldinger fra fileksport").setFontWeight("bold");
    }
    for (j = 0; j < obj[i].errors.length; j++) {
      var r2 = sheet.getRange(startAt + 4 + j, 1, 1, 3);
      r2.setValues([[obj[i].errors[j].message, "Linje nr:", obj[i].errors[j].line_no]]);
    }
    startAt += 6 + obj[i].errors.length;
  }
}

/*
* Convert the spreadsheet to a json structure.
*/
function autopass_JSON_convert(fid, gObject) {
  var spreadsheet = SpreadsheetApp.openById(fid);
  var report = false;
  if (typeof gObject == 'undefined') {
    report = true;
    gObject = {
      "file_name": spreadsheet.getName(),
      "num_lines": 0, 
      "num_lines_ok": 0, 
      "max_amount": 0, 
      "num_zeros": 0, 
      "errors": [],
    }
  }
  var sheet = spreadsheet.getSheetByName('Sheet1');
  data = [];
  var row = 12;
  var values = sheet.getRange(row, 1, sheet.getLastRow(), sheet.getLastColumn()).getDisplayValues();
  var reg_id_arr = {};
  for (var i = 0; i < values.length; i++) {
    var v = values[i];
    if((v[0]) == 'Antall passeringer:') break;
    var d = v[1];
    if (d.trim() == "" ) {
      gObject.errors.push({"line_no": row + i, "type": "MISSING TIME", "message": "Tidspunkt mangler for raden"});      
    }
    var date = d.substr(6,4) + '-' + d.substr(3,2) + '-' +d.substr(0,2) + ' ' + d.substr(12,5);
    var reg_id = v[4].trim();
    var chip_id = v[0].trim();
    var amount = +v[3].replace(',','.');
    var comment = v[2];
    gObject.num_lines++;
    if (amount == 0) {
      gObject.num_zeros++;
      continue;
    }
    
    if (gObject.max_amount - amount < 0) gObject.max_amount = amount;
    var jsdate = new Date(d.substr(6,4), d.substr(3,2) - 1, d.substr(0,2), d.substr(12,2), d.substr(15,2), 0, 0);
    if (typeof gObject.start === 'undefined' || jsdate < gObject.start) gObject.start = jsdate;
    if (typeof gObject.end === 'undefined' || jsdate > gObject.end) gObject.end = jsdate;
    if (amount < 0) {
      gObject.errors.push({"line_no": row + i, "type": "NEGATIVE AMOUNT",
        "message": "Det er negativt beløp på denne raden."});   
      continue;
    }
    else if (chip_id == "" && reg_id == "") {
      gObject.errors.push({"line_no": row + i, "type": "MISSING BOTH ID",
        "message": "Raden mangler både registrerings-nummer og og autopass-chip id."});      
    }
    else if (chip_id == "") {
      // Ignore, this is OK
      gObject.errors.push({"line_no": row + i, "type": "MISSING CHIP ID",
        "message": "Raden mangler autopass-chip id."});
    }
    else if (chip_id in reg_id_arr) {
      if (reg_id == "") {
        gObject.errors.push({"line_no": row + i, "type": "REPLACED REGID",
          "message": "Erstattet registreringsnummer med registreringsnummer for denne autopass-id på en tidligere rad."});      
        reg_id = reg_id_arr[chip_id];
      }
      else if (reg_id_arr[chip_id] != reg_id) {
        gObject.errors.push({"line_no": row + i, "type": "MULTIPLE CHIP REGID",
          "message": "Det er flere ulike registreringsnummer for denne chip-id: " + chip_id});
      }
    }
    else {
      if (reg_id == "") {
        gObject.errors.push({"line_no": row + i, "type": "MISSING REGID",
          "message": "Denne raden mangler registreringsnummer. Sjekk om det kan etterfylles i filen."});
      }
      else {
        reg_id_arr[chip_id] = reg_id;
      }
    }
    // add data
    data.push({date: date, reg_id: reg_id, amount: amount, comment: comment});
    gObject.num_lines_ok++;    
  }
  if (report) {
    var file = DriveApp.getFileById(fid);
    var json_folder = DriveApp.getFolderById(getProperty('json'));
    var output = DriveApp.createFile(file.getName() + '.json.txt', JSON.stringify(data, null, '\t'), 'application/json');
    json_folder.addFile(output);
    DriveApp.removeFile(output);
    report_basic_stats([gObject]);
    output.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    return {'fid':fid, 'url':output.getUrl(), 'name':output.getName() };
  }
  return data;
}

/*
* Get setup property, currently stored in the spreadsheet.
*/
function getProperty(name) {
  if (typeof name == 'undefined') {
    name = 'json';
  }
  if (Object.keys(global_setup).length === 0 && global_setup.constructor === Object) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETUP");
    var range = sheet.getRange("A5:B20").getValues();
    for (var i=0; i<range.length; i++) {
      if (range[i][0] != "") {
        global_setup[range[i][0]] = range[i][1];       
      }
    }
  }
  return global_setup[name];
}

/*
* Check if setup is ok. If not: Set up!
*/

function isSetup() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SETUP");
  if (sheet == null) {
    setup();
    return false;
  }
  return true;
}

function setup() {
  sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("SETUP");
  // Create autopass folder
  var autopass = DriveApp.getRootFolder().createFolder("Mappe for autopass faktura-filer")
  // Create json folder
  var json = autopass.createFolder("Mappe for eksporterte json-filer");
  sheet.getRange(1, 1).setValue("Oppsett-side for eksport-verktøyet").setFontSize(26);
  sheet.getRange(2,1).setValue(
    "Dette regnearket har oppsettet for eksportverktøyet. Forklaringer: \n" +
    "autopass - ID for mappen med autopass-faktureringer \n" +
    "json - ID for mappen der de eksporterte json-filene lagres \n" +
    "\n" +
    "ID for en mappe finner du i slutten av url-en når du er inne på en mappe, som f.eks: \n" +
    "https://drive.google.com/drive/u/0/folders/1vqeWfmSOuhZICljldsjfsfKUzg7XE8PvXf"
  );
  sheet.setRowHeight(2, 120);
  var r = sheet.getRange("A4:B6");
  r.setValues(
    [
      ["Nøkkelord", "Verdi"],
      ["autopass", autopass.getId()],
      ["json", json.getId()]
    ]
  );
  sheet.getRange("4:4").setFontWeight("bold");
  
  global_setup = {autopass:autopass.getId(), json: json.getId()};
}


// Simple date formatter
function df(d, f) {
  if(d == null || !d instanceof Date || typeof f !== 'string') {
    return '';
  }
  var months = [
    'Januar',
    'Februar',
    'Mars',
    'April',
    'Mai',
    'Juni',
    'Juli',
    'August',
    'September',
    'Oktober',
    'November',
    'Desember',
  ];
    // Formats like php
    var formats = {
    Y: "d.getFullYear()",
    n: "d.getMonth()+1",
    m: "('0'+(d.getMonth()+1)).slice(-2)",
    M: "months[d.getMonth()].slice(0,3)",
    F: "months[d.getMonth()]",
    d: "('0'+d.getDate()).slice(-2)",
    j: "d.getDate()",
    H: "('0'+d.getHours()).slice(-2)",
    i: "('0'+d.getMinutes()).slice(-2)",
    s: "('0'+d.getSeconds()).slice(-2)",
    };
    var output = '';
    for (var i=0, len=f.length; i<len; i++) {
    if (f[i] in formats) {
      output += eval(formats[f[i]])
    }
    else {
      output += f[i];
    }
  }
  return output;
}
