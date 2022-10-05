function Test_DS_Update_Localization_Gear() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_Localization_Gear();
}

// Will update sheet _M3Localization_Gear
function DS_Update_Localization_Gear() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const locFolder = versionFolder.getFoldersByName('m3localization').next();

  const heroesFolder = versionFolder.getFoldersByName('heroes').next();
  const fileGlobal = heroesFolder.getFilesByName('M3GlobalData.json').next();
  const globalContent = fileGlobal.getBlob().getDataAsString();
  const json = JSON.parse(globalContent);

  let GearTierCap = Number(json.GearTierCap);
  const overrideGTC = getNamedRangeValue('Preferences_OverrideMaxGearTier');
  if (overrideGTC != '' && Number(overrideGTC) > GearTierCap) GearTierCap = Number(overrideGTC);

  const gearIndex = {};
  const data = [];

  // Parse folders and add loc
  const langFolders = locFolder.getFolders();
  const langIndex = [];

  for (let l = 0; l < _langIds.length; l++) {
    langIndex[_langIds[l].toUpperCase()] = 1 + l;
  }

  while (langFolders.hasNext()) {
    const folder = langFolders.next();
    const langindex = langIndex[folder.getName().toUpperCase()];

    const file = folder.getFilesByName('gear.csv').next();

    content = file.getBlob().getDataAsString();
    const input = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    let row = 0;
    for (let r = 1; r < input.length; r++) {
      let id = input[r][0];
      if (id.indexOf('ID_GEAR_') != 0 || id.indexOf('_NAME') < 0) continue;
      id = id.substring(8, id.indexOf('_NAME'));

      if (id == 'T1_PLACEHOLDERGEAR') continue;

      if (id.startsWith('RED_') && GearTierCap < 20) continue;
      if (id.startsWith('WHITE_') && GearTierCap < 21) continue;

      if (id.startsWith('T') && !id.startsWith('Teal')) {
        const parts = id.split('_');
        const tier = Number(parts[0].substring(1));
        if (tier > GearTierCap) continue;
      }

      const value = input[r][1];
      if (value == 'No one uses THIS gear piece') continue;
      if (id == 'T1_TECH_RESIST') continue; //Scopely removed 'No one uses THIS gear piece on that piece' in ZHS folder... lol

      if (!gearIndex.hasOwnProperty(id)) {
        data[row] = [];
        data[row] = [id];
        gearIndex[id] = row++;
      }
      data[gearIndex[id]][langindex] = value;
    }
  }

  const sheet = GetSheet('_M3Localization_Gear');
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
