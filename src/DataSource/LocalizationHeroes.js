function Test_DS_Update_Localization_Heroes() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_Localization_Heroes();
}

// TODO: Make sure if a language is added the system will still work
function DS_Update_Localization_Heroes_Old() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const combat_dataFolder = versionFolder.getFoldersByName('combat_data').next();
  const characters = combat_dataFolder.getFilesByName('characters.json').next();
  const locFolder = versionFolder.getFoldersByName('m3localization').next();

  const heroIndex = {};
  const data = [[]];

  // characters.json ====================================================================
  let content = characters.getBlob().getDataAsString();
  const json = JSON.parse(content);

  const ids = Object.keys(json.Data);
  const count = ids.length;
  let row = 0;

  // Parse all the characters, and skip the irrelevant ones
  for (let r = 0; r < count; r++) {
    const id = ids[r];
    // Keep VIP_A considering he's got a specific portrait
    if (id != 'VIPMaleAOperator' && IsInvalidHero(id, false)) continue;

    const d = json.Data[id];
    const traits = d.traits;
    if (traits == null) continue;

    heroIndex[id.toUpperCase()] = row;

    // Hack VIP_A's name to fit his profile image name
    if (id == 'VIPMaleAOperator') data[row++] = ['VIP_A'];
    else data[row++] = [id];
  }

  // Parse folders and add loc
  const langFolders = locFolder.getFolders();
  const langIndex = [];

  for (let l = 0; l < _langIds.length; l++) {
    langIndex[_langIds[l].toUpperCase()] = l + 1;
  }

  while (langFolders.hasNext()) {
    const folder = langFolders.next();
    const langId = folder.getName().toUpperCase();

    if (!langIndex.hasOwnProperty(langId)) continue;

    const langindex = langIndex[langId];
    const file = folder.getFilesByName('heroes.csv').next();

    content = file.getBlob().getDataAsString();
    const input = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 0; r < input.length; r++) {
      const fullid = input[r][0].toString();
      if (fullid.indexOf('ID_SHARD_') != 0 || fullid.indexOf('_NAME') < 0) continue;
      const cid = fullid.substring(9, fullid.indexOf('_NAME'));

      // fill the appropriate column with the localized value
      if (!heroIndex.hasOwnProperty(cid)) continue;

      data[heroIndex[cid]][langindex] = input[r][1];
    }
  }

  return UpdateSortedRangeData(GetSheet('_M3Localization_Heroes'), 1, 3, 2 + data[0].length, 1, data);
}

function DS_Update_Localization_HeroesFormula_Old() {
  const sheet = GetSheet('_M3Localization_Heroes');

  // ------------------------------------------------------------------------------------------------------------------------------
  // Make sure column B is still fine (can only be wrong if there's an addition or removal at the top row)
  const currentformula = sheet.getRange(1, 2).getFormula();
  if (currentformula == '' || currentformula == null) {
    if (_setValuesCount >= _maxSetValue) return false;
    sheet.getRange(1, 2).setFormula('=OFFSET(C:C, 0, Preferences_LanguageIndex)');
    _setValuesCount++;
  }
  sheet.getRange(2, 2, sheet.getMaxRows() - 1).clearContent();
  // ------------------------------------------------------------------------------------------------------------------------------

  const rangeFormulas = sheet.getRange(1, 1, sheet.getMaxRows(), 1).getFormulas();
  const ranges = GetEmptyRows(rangeFormulas);

  for (let g = 0; g < ranges.length; g++) {
    const newData = [];
    for (let row = ranges[g][0]; row < ranges[g][1]; row++) {
      // Set formula table for the portrait
      newData.push([`=IF(Preferences_Images_Portraits,IMAGE(_Values_Url_Portrait&$C${row}&_Values_Image_Ext),)`]);
    }

    if (!copyToRangeFormula(sheet, ranges[g][0], 1, newData)) return false;
  }

  return true;
}
