function toastMessage(title, message) {
  SpreadsheetApp.getActive().toast(message, title);
}

function toastError(message) {
  SpreadsheetApp.getActive().toast(message, '⚠️ Error');
}

function toastSuccess(message) {
  SpreadsheetApp.getActive().toast(message, '👍 Success');
}
