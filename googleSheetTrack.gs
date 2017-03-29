/* global variables */
var form_id = 'put the form id in here';
var form_url = 'put the full url to your form here'
var sheet_name = 'the name of the sheet';
var edit_col = int; // the column where form edit urls live (A = 1)
var id_row = int; // the column where your identifier will live (A = 1)

// Display prompt on sheet open
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.prompt('Add/Find Something', ui.ButtonSet.OK_CANCEL);

  // Process the user's response. If we find the id, send user to edit form. If no id found, send user to new form
  if (prompt.getSelectedButton() == ui.Button.OK) {
    var foundRow = findRow(prompt.getResponseText());
   
    if (foundRow) {
      var editUrl = findEditUrl(foundRow);
      var message = 'Existing ID: ' + prompt.getResponseText() + ' found! Edit current entry?';
      goToForm(editUrl, message, 'edit');  
    } else {
      var message = 'ID: ' + prompt.getResponseText() + ' not found! Create a new entry?';
      goToForm(form_url, message, 'new entry');
    }
  }
}

// Load a given form into new tab with a message and type (edit or new entry)
function goToForm(url, message, type) {
  var ui = SpreadsheetApp.getUi();
  var prompt = ui.alert(message, ui.ButtonSet.YES_NO);
  
  if (prompt == ui.Button.YES) {
    var sendFormToTab = '<script>window.open("' +  url + '")</script>';
    
    var output = HtmlService
    .createHtmlOutput(sendFormToTab)
    .setHeight(20);
    
    ui.showModalDialog(output, 'Sending ' + type + ' form to new tab (check your pop up blocker!)');
  }
}


// Get all the ids and their corresponding rows [row, id]
function getIds() {
  var allData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name).getDataRange().getValues();
  var ids = [];
  
  for (i = 1; i < allData.length; i++) {
    ids.push([i, allData[i][id_col - 1]]);
  }
  return ids;
}

// Find the row of a given id value
function findRow(value) {
  var ids = getIds();
  
  // if id (index 1) is equal to the value given, return the row it lives in (index 0)
  for (i = 0; i < ids.length; i++) {
    if(value === ids[i][1]) {
      return ids[i][0];
      break;
    }
  }
}

// Find the EditUrl of a given row
function findEditUrl(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  return sheet.getRange(row + 1, edit_col).getValue();
}
  
// Set this script to trigger on form submit. Adds an appropriate url to edit each response. Make sure your form allows editing of responses.
function addEditUrl() {
  var form = FormApp.openById(form_id);
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  var data = sheet.getDataRange().getValues();
  var urlCol = edit_col;
  var responses = form.getResponses();
  var timestamps = [], urls = [], resultUrls = [];
  
  for (var i = 0; i < responses.length; i++) {
    timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
    urls.push(responses[i].getEditResponseUrl());
  }
  
  for (var j = 1; j < data.length; j++) {
    resultUrls.push([data[j][0]?urls[timestamps.indexOf(data[j][0].setMilliseconds(0))]:'']);
  }
  sheet.getRange(2, urlCol, resultUrls.length).setValues(resultUrls);
}
