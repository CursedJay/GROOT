function forgetMe() {
  GrootApi.forgetLogin();
}

function forgetAT() {
  GrootApi.debugForgetAccessToken();
}

function saveReqToFile(data, filename) {
  const dateTime = Utilities.formatDate(new Date(), 'GMT+8', "yyyy-MM-dd'T'HH:mm:ss.SS");

  const folder = getFolder(_zaratoolsFolder, true);
  const file = filename ?? `msf-req-${dateTime}.json`;
  folder.createFile(file, JSON.stringify(data, null, 2), 'application/json');
}
