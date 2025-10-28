// Updates sheet _M3Localization_Roster
function DS_Update_Localization_Roster() {
  const localizationFiles = GrootApi.getLocalizationFile('roster');
  if (localizationFiles === false) return;

  const listIndex = {};
  const data = [];

  let row = 0;
  for (let index = 0; index < _langIds.length; index++) {
    const langId = _langIds[index].toLowerCase();

    if (!localizationFiles[langId]) continue; //language url is missing
    // Fetch the CSVs from the URLs returned by the API
    const res = UrlFetchApp.fetch(localizationFiles[langId]);
    const content = res.getContentText();
    const csv = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 1; r < csv.length; r++) {
      let id = csv[r][0];
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
      data[listIndex[id]][index + 1] = csv[r][1];
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
