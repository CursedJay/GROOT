function updateBetaVersion() {
  const build = getNamedRangeValue('_Version_Build_Beta');
  const versionId = getNamedRangeValue('_Version_Id_Beta');
  const versionNum = getNamedRangeValue('_Version_Numbers_Beta');

  updateVersionTo(Number(build), versionId, versionNum);
}

function updateVersion() {
  const build = getNamedRangeValue('_Version_Build_Latest');
  const versionId = getNamedRangeValue('_Version_Id_Latest');
  const versionNum = getNamedRangeValue('_Version_Numbers_Latest');

  updateVersionTo(Number(build), versionId, versionNum);
}

function updateVersionTo(buildNum, sheetId, version) {
  const ui = SpreadsheetApp.getUi();
  const current = Number(getNamedRangeValue('_Version_Build_Current'));

  if (current == 0 || buildNum == 0) {
    ui.alert("There's a temporary issue, please try again later.");
    return;
  }

  if (buildNum <= current) {
    ui.alert("You're already using the latest build!");
    return;
  }

  const sourceId = SpreadsheetApp.getActive().getId();
  const file = DriveApp.getFileById(sourceId);
  const fileParents = file.getParents();

  let folder = null;
  if (fileParents.hasNext()) {
    folder = fileParents.next();
  }

  const latestFile = DriveApp.getFileById(sheetId);
  const playerName = getNamedRangeValue('Profile_Name');
  const fileName = `${
    playerName.length == 0 || playerName == '<Your Name Here>' ? 'GROOT' : playerName
  } v${version}${_finishUpdateTag}`;

  let newFile = null;
  if (folder != null) newFile = latestFile.makeCopy(fileName, folder);
  else newFile = latestFile.makeCopy(fileName);

  // ---------------------------------------------------
  // Anyone with the link can view, then copy sharing settings
  newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  const thisFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const editors = thisFile.getEditors();
  for (let i = 0; i < editors.length; i++) {
    newFile.addEditor(editors[i]);
  }
  const viewers = thisFile.getViewers();
  for (let i = 0; i < viewers.length; i++) {
    newFile.addViewer(viewers[i]);
  }
  // ---------------------------------------------------

  setNamedRangeValue('_Version_NewDocId', newFile.getId());

  /*var cache = CacheService.getUserCache();
    cache.put("MSFMigrationTarget", newFile.getId());
    cache.put("MSFMigrationSource", sourceId);
    */
  saveJSON_Base(newFile.getId());
  ui.alert(
    "Process is almost done!\nYour new GROOT will be opened in an instant.\nPlease select 'Update version / Finish update' on this new document to finish the process"
  );

  openSheet(newFile.getId());
}

function openSheet(spreadsheetId) {
  const url = `https://docs.google.com/spreadsheets/d/${spreadsheetId}`;
  const html = `<script>window.open('${url}');google.script.host.close();</script>`;
  const userInterface = HtmlService.createHtmlOutput(html);
  SpreadsheetApp.getUi().showModalDialog(userInterface, 'Opening your new document');
}
