var _updateStart;
var updateVersion;

function SetUpdateProgress(newValue) {
  var oldValue = PropertiesService.getScriptProperties().getProperty('OldUpdateProgress');

  if (oldValue != null) {
    SpreadsheetApp.getActiveSpreadsheet().removeMenu(oldValue);
    PropertiesService.getScriptProperties().deleteProperty('OldUpdateProgress');
  }

  if (newValue != null) {
    var subMenu = [{ name: 'Continue', functionName: 'DataSourceUpdate_Continue' }];
    PropertiesService.getScriptProperties().setProperty('OldUpdateProgress', newValue);
    SpreadsheetApp.getActiveSpreadsheet().addMenu(newValue, subMenu);
  }
}

function DataSourceUpdate_Start() {
  _updateStart = new Date().getTime();
  updateVersion = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', updateVersion);

  DataSource_SetStep(0);
  DataSource_Loop(0);
}

function DataSourceUpdate_Continue() {
  _updateStart = new Date().getTime();
  if (updateVersion == null) {
    SetUpdateProgress(null);
    return;
  }

  var step = DataSource_GetStep();
  DataSource_Loop(step);
}

function DataSource_SetStep(step) {
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateStep', step);
  SetUpdateProgress('Updating... ' + (step + 1) + '/15');
}

function DataSource_GetStep() {
  return Number(PropertiesService.getScriptProperties().getProperty('DataSourceUpdateStep'));
}

function DataSource_Loop(step) {
  do {
    output('Loop = ' + step);
    if (step == -1 || DataSource_Timeout(60)) {
      // Less that a minute left, abord
      DataSource_TimeoutPopup();
      return;
    }
    step = DataSource_UpdateStep(step);
  } while (step < 666);

  SetUpdateProgress(null);
  SpreadsheetApp.getActiveSpreadsheet().removeMenu('New DataSource Available!');
  SpreadsheetApp.getUi().alert('Process completed!');
}

function DataSource_TimeoutPopup() {
  SpreadsheetApp.getUi().alert(
    'This script needs more time to complete. To avoid timeout please close this popup and run the update data source again to continue.'
  );
}

function DataSource_Timeout(secLeft) {
  var elapsed = new Date().getTime() - _updateStart;
  return Math.round(elapsed / 1000) + secLeft >= 6 * 60;
}

function DataSource_UpdateStep(step) {
  output('Update Step ' + step);

  updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');

  switch (step) {
    case 0:
      if (!DS_Update_HeroData()) return -1;
      break;

    case 1:
      if (!DS_Update_HeroDataFormula()) return -1;
      break;

    case 2:
      if (!DS_Update_GearTiers()) return -1;
      break;

    case 3:
      if (!DS_Update_GearTiersFormula()) return -1;
      break;

    case 4:
      if (!DS_Update_GearLibrary()) return -1;
      break;

    case 5:
      if (!DS_Update_GearLibraryCraftFormula()) return -1;
      break;

    case 6:
      if (!DS_Update_GearLibraryHeroFormula()) return -1;
      break;

    case 7:
      if (!DS_Update_Localization_Heroes()) return -1;
      break;

    case 8:
      if (!DS_Update_Localization_HeroesFormula()) return -1;
      break;

    case 9:
      if (!DS_Update_Localization_Gear()) return -1;
      break;

    case 10:
      if (!DS_Update_Localization_Roster()) return -1;
      break;

    case 11:
      if (!DS_Update_Localization_Challenges()) return -1;
      break;

    case 12:
      if (!DS_Update_MissionsData()) return -1;
      break;

    case 13:
      if (!DS_Update_MissionsDataFormula()) return -1;
      break;

    case 14:
      AddMissingToons();
      DS_Update_Links();
      setNamedRangeValue('_Version_DataSource_Current', updateVersion);
      return 666;
      break;
  }

  DataSource_SetStep(++step);
  return step;
}
