/**
 * @NotOnlyCurrentDoc
 */

function onOpen(event) {
  const docname = SpreadsheetApp.getActiveSpreadsheet().getName();

  if (docname.endsWith(_finishUpdateTag)) {
    SpreadsheetApp.getUi().createMenu('Finish Update').addItem('Go!', 'FinishUpdate').addToUi();
  } else {
    CreateMenu();

    const buildCurrent = Number(getNamedRangeValue('_Version_Build_Current'));
    const buildLatest = Number(getNamedRangeValue('_Version_Build_Latest'));

    if (buildCurrent < buildLatest) {
      SpreadsheetApp.getUi().createMenu('New Version Available!').addItem('Begin Update', 'updateVersion').addToUi();
    }

    const flagBeta = getNamedRangeValue('Preferences_Flag_Beta');
    if (flagBeta) {
      const buildBeta = getNamedRangeValue('_Version_Build_Beta');
      if (buildCurrent < buildBeta && buildLatest != buildBeta) {
        SpreadsheetApp.getUi()
          .createMenu('New Beta Version Available!')
          .addItem('Begin Update', 'updateBetaVersion')
          .addToUi();
      }
    }

    const dataSourceCurrent = getNamedRangeValue('_Version_DataSource_Current');
    const dataSourceLatest = getNamedRangeValue('_Version_DataSource_Latest');

    if (dataSourceCurrent != dataSourceLatest && dataSourceLatest != '#REF!') {
      SpreadsheetApp.getUi()
        .createMenu('New DataSource Available!')
        .addItem('Update', 'DataSourceUpdate_Start')
        .addToUi();
    }

    const infoSourceCurrent = getNamedRangeValue('_Version_InfoSource_Current');
    const infoSourceLatest = getNamedRangeValue('_Version_InfoSource_Latest');

    const t1 = infoSourceCurrent.getTime();
    const t2 = infoSourceLatest.getTime();

    // 10mn delta
    if (Math.abs(t1 - t2) > 10 * 60 * 1000) {
      SpreadsheetApp.getUi().createMenu('New InfoSource Available!').addItem('Update', 'InfoSourceUpdate').addToUi();
    }
  }
}

function FinishUpdate() {
  const docname = SpreadsheetApp.getActiveSpreadsheet().getName();
  loadJSON_Update();
  CreateMenu();

  try {
    SpreadsheetApp.getActiveSpreadsheet().setName(docname.substring(0, docname.lastIndexOf(_finishUpdateTag)));
    SpreadsheetApp.getActiveSpreadsheet().removeMenu('Finish Update');
  } catch (e) {
    error(e);
  }
}

function CreateMenu() {
  const ui = SpreadsheetApp.getUi();

  const msfapi = ui.createMenu('MSF API');
  msfapi.addItem('Update inventory', 'api_importInventory');
  msfapi.addItem('Import FULL inventory', 'api_importFullInventory');
  msfapi.addSeparator();
  msfapi.addItem('Import Profile', 'api_importProfile');
  msfapi.addSeparator();
  msfapi.addItem('Forget me', 'forgetMe');

  const dataSource = ui.createMenu('Data source');
  dataSource.addItem('Update data source', 'DataSourceUpdate_Start');
  dataSource.addItem('Add missing characters', 'AddMissingToons');

  //var menuUpdate = SpreadsheetApp.getUi().createMenu("Update");
  //menuUpdate.addItem("Begin update to latest version", "updateVersion");
  //menuUpdate.addItem("Begin update to latest beta version", "updateBetaVersion");

  const menuImport = ui.createMenu('Import');
  menuImport.addItem('Import roster from MSF.gg', 'msfgg_import');
  menuImport.addSeparator();
  menuImport.addItem('Import from Google Drive', 'loadJSON');
  menuImport.addItem('Import from MANTIS', 'mantis_import');
  menuImport.addItem('Import from MSF Toolbot', 'msftoolbot_import');
  menuImport.addSeparator();
  menuImport.addItem('Finish Update', 'FinishUpdate');

  const menuExport = ui.createMenu('Export');
  menuExport.addItem('Export MSF.gg War Defense (csv)', 'msfgg_exportWarDefense');
  menuExport.addSeparator();
  menuExport.addItem('Export to Google Drive', 'saveJSON');

  ui.createMenu('Zaratools')
    .addSubMenu(msfapi)
    .addSubMenu(dataSource)
    //.addSubMenu(menuUpdate)
    .addSubMenu(menuImport)
    .addSubMenu(menuExport)
    // Ingame Sort
    .addSeparator()
    .addItem('Ingame Sort', 'IngameSort')
    .addItem('Notes as shards required for the next level', 'ShardsToNextLevelNotes')
    .addToUi();
}

function onEdit(event) {
  UpdateModified();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();

  const docname = ss.getName();
  if (docname.endsWith(' [Master]')) return; // Prevent formula deletion for the Master doc

  if (activeSheet.getName() == 'Preferences') {
    if (ss.getActiveCell().getA1Notation() == 'B4') {
      output(`ChangeLanguage: ${event.value}`);
      ChangeLanguage(event.value);
    }
  } else if (activeSheet.getName() == 'Roster') {
    if (ss.getActiveCell().getA1Notation() == 'C1') {
      if (ss.getActiveCell().getValue() == 'IN-GAME SORT') {
        IngameSort();
        ss.getActiveCell().setValue(event.oldValue);
      } else FilterTeam();
    }
  } else if (activeSheet.getName() == 'Raid') {
    if (ss.getActiveCell().getA1Notation() == 'B3') {
      activeSheet.getRange(3, 3).clearContent();
    }
  } else if (activeSheet.getName() == 'Missions') {
    const column = ss.getActiveCell().getColumn();
    const row = ss.getActiveCell().getRow();

    if (ss.getActiveCell().getA1Notation() == 'B2') {
      activeSheet.getRange(2, 3).setValue('');
    }

    const firstCol = 6;
    const deltaCol = 7;

    const rowCategory = 3;
    const rowEvent = 4;
    const rowDelta = 11;

    if ((column - firstCol) % deltaCol == 0) {
      if ((row - rowCategory) % rowDelta == 0) {
        activeSheet.getRange(row + 1, column).clearContent();
        activeSheet.getRange(row + 4, column, 5, 1).clearContent();
      }
      if ((row - rowEvent) % rowDelta == 0) {
        activeSheet.getRange(row + 3, column, 5, 1).clearContent();
      }
    }
  } else if (activeSheet.getName() == 'Teams') {
    if (ss.getActiveCell().getA1Notation() == 'O3') {
      const sort = getNamedRangeValue('Teams_Sorting');
      if (sort != '') {
        SortTeamsByPower(sort);
        ss.getActiveCell().clearContent();
      }
    }
  }
}

function UpdateModified() {
  const date = new Date();
  const today = Utilities.formatDate(date, 'GMT-04:00', 'yyyy/MM/dd');
  const cache = CacheService.getUserCache();
  const docId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const lastUpdateRoster = cache.get(`${docId}_lastUpdateRoster`);
  const lastUpdate = cache.get(`${docId}_lastUpdate`);

  if (today != lastUpdateRoster) {
    if (today != lastUpdate) {
      cache.put(`${docId}_lastUpdate`, today);
      setNamedRangeValue('_GAMORA_UpdateDoc', today);
    }

    // Get the cell that was just modified.
    const activeSheet = SpreadsheetApp.getActiveSheet();
    const editedCell = activeSheet.getActiveCell();

    const powerRange = getNamedRange('Roster_Power');

    if (
      powerRange.getSheet().getSheetId() == activeSheet.getSheetId() &&
      powerRange.getColumn() == editedCell.getColumn()
    ) {
      cache.put(`${docId}_lastUpdateRoster`, today);
      setNamedRangeValue('_GAMORA_UpdateRoster', today);
    }
  }
}
