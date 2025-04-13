function forgetMe() {
  GrootApi.forgetLogin();
}

function forgetAT() {
  GrootApi.debugForgetAccessToken();
}

function saveReqToFile(data, filename = 'msf-req', filetype = 'json') {
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const dateTime = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd-HH:mm:ss');
  const folder = getFolder(_zaratoolsFolder, true);
  const file = `${filename}_${dateTime}.${filetype}`;
  try {
    if (filetype === 'json') {
      folder.createFile(file, JSON.stringify(data, null, 2), 'application/json');
    } else {
      folder.createFile(file, data);
    }
    toastSuccess(`${filename} successfully exported in your Google Drive to: ${folder.getName()}/${file}`);
  } catch (error) {
    toastError(`Couldn't create the file \n ${error}`);
  }
}
