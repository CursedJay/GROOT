function Test_DS_Update_M3Localization_Challenges() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_Localization_Challenges();
}

// Will update sheet _M3Localization_Challenges
function DS_Update_Localization_Challenges() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const locFolder = versionFolder.getFoldersByName('m3localization').next();

  var listIndex = {};
  var data = [];

  // Parse folders and add loc
  var langFolders = locFolder.getFolders();
  var langIndex = {};

  for (var l = 0; l < _langIds.length; l++) {
    langIndex[_langIds[l].toUpperCase()] = 1 + l;
  }

  var row = 0;
  while (langFolders.hasNext()) {
    const folder = langFolders.next();
    const langindex = langIndex[folder.getName().toUpperCase()];

    ///////////////////////////////////////////////////////////////////////////
    var file = folder.getFilesByName('challenges.csv').next();
    var content = file.getBlob().getDataAsString();
    var input = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (var r = 1; r < input.length; r++) {
      var id = input[r][0];

      if (id != 'ID_EVENT_LEGENDARY' && id != 'ID_EVENT_FLASH_EVENT' && id != 'ID_EVENT_CHARACTER_EVENT') {
        if (id.indexOf('_NAME') < 0) continue;
        if (id.indexOf('_STAGE') >= 0) continue;
        id = id.substring(0, id.indexOf('_NAME'));
      }

      if (
        id.startsWith('ID_ECM_') ||
        id.startsWith('ID_CHAL_') ||
        id.startsWith('ID_EVENT_') ||
        id.startsWith('ID_FL_') ||
        id.startsWith('ID_LE_') ||
        id.startsWith('ID_TU_')
      )
        id = id.substring(3);
      else continue;

      if (!listIndex.hasOwnProperty(id)) {
        data[row] = [id];
        listIndex[id] = row++;
      }
      data[listIndex[id]][langindex] = input[r][1];
    }

    ///////////////////////////////////////////////////////////////////////////
    var file = folder.getFilesByName('missions.csv').next();
    var content = file.getBlob().getDataAsString();
    var input = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (var r = 1; r < input.length; r++) {
      var id = input[r][0];

      if (id.indexOf('_NAME') < 0) continue;
      id = id.substring(0, id.indexOf('_NAME'));

      if (id.startsWith('ID_TU_')) id = id.substring(3);
      else continue;

      if (!listIndex.hasOwnProperty(id)) {
        data[row] = [id];
        listIndex[id] = row++;
      }
      data[listIndex[id]][langindex] = input[r][1];
    }

    ////////////////////////////////////////////////////////////////////////////
    file = folder.getFilesByName('ui.csv').next();
    content = file.getBlob().getDataAsString();
    input = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (var r = 1; r < input.length; r++) {
      var id = input[r][0];

      if (
        id != 'ID_GAMEMODE_CAMPAIGNS' &&
        id != 'ID_GAMEMODE_CHALLENGES' &&
        id != 'ID_UI_TEAM_UP' &&
        id != 'ID_UI_SELECT_AVAILABLE'
      ) {
        if (id.indexOf('_NAME') < 0) continue;
        if (!id.startsWith('ID_GAMEMODE_CAMPAIGN_')) continue;
        id = id.substring(0, id.indexOf('_NAME'));
      }

      id = id.substring(3);

      if (!listIndex.hasOwnProperty(id)) {
        data[row] = [id];
        listIndex[id] = row++;
      }
      data[listIndex[id]][langindex] = input[r][1];
    }
  }

  for (var r = 0; r < data.length; r++) {
    // If there's empty slots anywhere, use the ID as name
    for (var l = 1; l < data[0].length; l++) {
      if (data[r][l] == null) data[r][l] = data[r][0];
    }
  }

  const sheet = GetSheet('_M3Localization_Challenges');
  if (!UpdateSortedRangeData(sheet, 1, 2, 1 + data[0].length, 1, data)) return false;

  // Do the formula part here considering there's just one cell
  // ------------------------------------------------------------------------------------------------------------------------------
  // Make sure column A is still fine (can only be wrong if there's an addition or removal at the top row)
  const currentformula = sheet.getRange(1, 1).getFormula();
  if (currentformula == '' || currentformula == null) {
    if (_setValuesCount >= _maxSetValue) return false;
    sheet.getRange(1, 1).setFormula('=OFFSET(B:B, 0, Preferences_LanguageIndex)');
    _setValuesCount++;
  }
  sheet.getRange(2, 1, sheet.getMaxRows() - 1).clearContent();
  // ------------------------------------------------------------------------------------------------------------------------------

  return true;
}
