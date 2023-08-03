function Test_DS_Update_M3Localization_Roster() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_Localization_Roster();
}

function DS_Update_Localization_Roster_Old() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const locFolder = versionFolder.getFoldersByName('m3localization').next();

  const listIndex = {};
  const data = [];

  // Parse folders and add loc
  const langFolders = locFolder.getFolders();
  const langIndex = {};

  for (let l = 0; l < _langIds.length; l++) {
    langIndex[_langIds[l].toUpperCase()] = 1 + l;
  }

  let row = 0;
  while (langFolders.hasNext()) {
    const folder = langFolders.next();
    const langindex = langIndex[folder.getName().toUpperCase()];

    const files = folder.getFilesByName('roster.csv');
    const file = files.next();

    const content = file.getBlob().getDataAsString();
    const input = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 1; r < input.length; r++) {
      let id = input[r][0];
      if (id.startsWith('ID_ENUM_HERO_TRAIT_')) id = id.substring(8);
      else if (
        id == 'ID_ROSTER_FILTER_ALL' ||
        id == 'ID_ROSTER_FILTER_AVAILABLE' ||
        id == 'ID_ROSTER_SORT_REDSTARS' ||
        id == 'ID_ROSTER_SORT_LOYALTY'
      )
        id = id.substring(3);
      else continue;

      if (!listIndex.hasOwnProperty(id)) {
        data[row] = [id];
        listIndex[id] = row++;
      }
      data[listIndex[id]][langindex] = input[r][1];
    }
  }

  for (let r = 0; r < data.length; r++) {
    // If there's empty slots anywhere, use the ID as name
    for (let l = 1; l < data[0].length; l++) {
      if (data[r][l] == null) data[r][l] = data[r][0];
    }
  }

  const sheet = GetSheet('_M3Localization_Roster');
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
