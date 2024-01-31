// Will update sheet _M3Localization_Iso8
// This is not supposed to change. No need to run on every updates
function DS_Update_Localization_Iso8() {
  const listIndex = {};
  const data = [];

  const isoIds = [
    'ID_UI_ISO8_TITLE',
    'ID_UI_ISO8_CLASS',
    'ID_UI_ISO8_ROLE_TRAIT',
    'ID_ISO8-TIER-0-CURRENCY_NAME',
    'ID_CAMPAIGN_PLAYER_ENERGY_ISO8_NAME',
    'ID_GAMEMODE_CAMPAIGN_ISO8_NAME'
  ];

  const abilitiesIds = [];

  const isoLocalFiles = GrootApi.getLocalizationFile('iso');
  if (isoLocalFiles === false) return;
  const eaIsoLocalFiles = GrootApi.getLocalizationFile('extraAbility_iso8');
  if (eaIsoLocalFiles === false) return;

  const sourceFiles = [isoLocalFiles, eaIsoLocalFiles];
  const sourceIds = [isoIds, abilitiesIds];

  let row = 0;

  for (let index = 0; index < _langIds.length; index++) {
    const langId = _langIds[index].toLowerCase();

    for (let source = 0; source < sourceFiles.length; source++) {
      // Fetch the CSVs from the URLs returned by the API
      const res = UrlFetchApp.fetch(sourceFiles[source][langId]);
      const content = res.getContentText();
      const csv = Utilities.parseCsv(content, ','.charCodeAt(0));

      // Parse lines of the CSV file
      for (let r = 1; r < csv.length; r++) {
        let id = csv[r][0];

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

        data[listIndex[id]][index + 1] = csv[r][1];
      }
    }
  }

  for (let r = 0; r < data.length; r++) {
    // If there's empty slots anywhere, use the ID as name
    for (let l = 1; l < data[0].length; l++) {
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
