function Test_DS_Update_M3Localization_UI() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_Localization_UI();
}

// Will update sheet _M3Localization_UI
// No need to let users use it, this is not supposed to change
function DS_Update_Localization_UI() {
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

  const uiIds = [
    'ID_PLAYER_PROFILE_PID',
    'ID_PLAYERSTAT_ARENA_WINS',
    'ID_PLAYERSTAT_BEST_ARENA_RANK',
    'ID_PLAYERSTAT_HEROES_AT_MAX_LOYALTY',
    'ID_PLAYERSTAT_LAST_ARENA_RANK',
    'ID_PLAYERSTAT_STRONGEST_TEAM_POWER',
    'ID_PLAYERSTAT_TOTAL_COLLECTION_POWER',
    'ID_PLAYERSTAT_TOTAL_HEROES_COLLECTED',
    'ID_PLAYERSTAT_LATEST_BATTLEGROUNDS_RANK',
    'ID_PLAYERSTAT_BEST_BATTLEGROUNDS_RANK',
    'ID_PLAYERSTAT_WAR_MVP',
    'ID_POWER',
    'ID_SELECT_BUTTON_2',
    'ID_SLOT',
    'ID_TIER',
    'ID_UI_ABILITIES',
    'ID_UI_FINDER_ARENA_ORB',
    'ID_UI_FINDER_BLITZ_ORB',
    'ID_UI_FINDER_RAID_ORB',
    'ID_UI_FINDER_ULTRA_ORB',
    'ID_UI_GEAR',
    'ID_LEVEL',
    'ID_UI_SELECT_AVAILABLE',
    'ID_UI_TEAM_POWER',
    'ID_UI_TRAITS',
    'ID_STORE_CATEGORYNAME_POWERCORES',
    'ID_STORE_NAME_BLITZ',
    'ID_STORE_NAME_RAID',
    'ID_STORE_NAME_ARENA',
    'ID_STORE_NAME_WAR',
    'ID_STORE_CATEGORYNAME_ORBS',
    'ID_UI_FINDER_WAR_ORB',
    'ID_LOYALTY_HEADER',
    'ID_UI_STORE'
  ];

  const storeIds = [
    'ID_GACHA_CRATE_MEGA_NAME',
    'ID_GACHA_CRATE_ARENA_SMALL_NAME',
    'ID_GACHA_CRATE_BASIC_NAME',
    'ID_GACHA_CRATE_DOWNCONVERSION_NAME',
    'ID_GACHA_CRATE_PREMIUM_NAME',
    'ID_GACHA_CRATE_MILESTONE_NAME',
    'ID_GACHA_CRATE_RAID_CHASE_SHARD_NAME',
    'ID_GACHA_CRATE_EVENT_RAID_ALPHA_MYSTIC_VILLAINS_202005_NAME',
    'ID_GACHA_CRATE_EVENT_RAID_BETA_201907_NAME',
    'ID_GACHA_CRATE_EVENT_RAID_BRAWLER_201907_NAME'
  ];

  const sourceFiles = ['ui.csv', 'store.csv'];
  const sourceIds = [uiIds, storeIds];

  let row = 0;

  while (langFolders.hasNext()) {
    const folder = langFolders.next();
    const langindex = langIndex[folder.getName().toUpperCase()];

    for (let source = 0; source < sourceFiles.length; source++) {
      ////////////////////////////////////////////////////////////////////////////
      const file = folder.getFilesByName(sourceFiles[source]).next();
      const content = file.getBlob().getDataAsString();
      const input = Utilities.parseCsv(content, ','.charCodeAt(0));

      // Parse lines of the CSV file
      for (let r = 1; r < input.length; r++) {
        let id = input[r][0];

        if (!sourceIds[source].includes(id)) continue;

        id = id.substring(3);
        if (id.endsWith('_NAME')) id = id.substring(0, id.lastIndexOf('_NAME'));

        if (!listIndex.hasOwnProperty(id)) {
          data[row] = [id];
          listIndex[id] = row++;
        }

        let value = input[r][1];

        while (value.includes('{') && value.includes('}')) {
          const val = value;
          value = '';
          if (val.indexOf('{') > 0) value = val.substring(0, val.indexOf('{'));
          value += val.substring(val.indexOf('}') + 1);
          value.replace('  ', ' ');
          if (value.startsWith(' ')) value = value.substring(1);
          if (value.endsWith(' ')) value = value.substring(0, value.length - 1);
        }

        data[listIndex[id]][langindex] = value;
      }
    }
  }

  for (let r = 0; r < data.length; r++) {
    // If there's empty slots anywhere, use the ID as name
    for (let l = 1; l < data[0].length; l++) {
      if (data[r][l] == null) data[r][l] = data[r][0];
    }
  }

  const sheet = GetSheet('_M3Localization_UI');
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
