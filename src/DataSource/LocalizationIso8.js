function Test_DS_Update_M3Localization_Iso8() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_Localization_Iso8();
}

// Will update sheet _M3Localization_Iso8
// No need to let users use it, this is not supposed to change
function DS_Update_Localization_Iso8() {
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

  const isoIds = [
    'ID_UI_ISO8_TITLE',
    'ID_UI_ISO8_CLASS',
    'ID_UI_ISO8_ROLE_TRAIT',
    'ID_ISO8-TIER-0-CURRENCY_NAME',
    'ID_CAMPAIGN_PLAYER_ENERGY_ISO8_NAME',
    'ID_GAMEMODE_CAMPAIGN_ISO8_NAME'
  ];

  const abilitiesIds = [];

  const sourceFiles = ['iso.csv', 'extraAbility_iso8.csv'];
  const sourceIds = [isoIds, abilitiesIds];

  var row = 0;

  while (langFolders.hasNext()) {
    const folder = langFolders.next();
    const langindex = langIndex[folder.getName().toUpperCase()];

    for (var source = 0; source < sourceFiles.length; source++) {
      ////////////////////////////////////////////////////////////////////////////
      const file = folder.getFilesByName(sourceFiles[source]).next();
      const content = file.getBlob().getDataAsString();
      const input = Utilities.parseCsv(content, ','.charCodeAt(0));

      // Parse lines of the CSV file
      for (var r = 1; r < input.length; r++) {
        var id = input[r][0];

        if (!id.startsWith('ID_EXTRA_ABILITY_ISO8_') || !id.endsWith('_NAME')) {
          if (
            !sourceIds[source].includes(id) &&
            !id.startsWith('ID_ISO8_STAT_') &&
            !id.startsWith('ID_UI_ISO8_TRAIT_NAME_')
          )
            continue;
        }

        id = id.substring(3);
        if (id.startsWith('UI_')) id = id.substring(3);
        if (id.endsWith('_NAME')) id = id.substring(0, id.lastIndexOf('_NAME'));
        if (id.startsWith('EXTRA_ABILITY_ISO8_')) id = id.substring(14);
        id = id.replace('-', '_');

        if (!listIndex.hasOwnProperty(id)) {
          data[row] = [id];
          listIndex[id] = row++;
        }

        data[listIndex[id]][langindex] = input[r][1];
      }
    }
  }

  for (var r = 0; r < data.length; r++) {
    // If there's empty slots anywhere, use the ID as name
    for (var l = 1; l < data[0].length; l++) {
      if (data[r][l] == null) data[r][l] = data[r][0];
    }
  }

  const sheet = GetSheet('_M3Localization_Iso8');
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
