function output(text) {
  Logger.log(text);
}

function error(text) {
  SpreadsheetApp.getUi().alert('ERROR: ' + text);
}
