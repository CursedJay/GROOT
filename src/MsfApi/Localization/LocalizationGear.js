// Updates sheet _M3Localization_Gear
function DS_Update_Localization_Gear() {
  const gearIndex = {};
  const data = [];

  const gearLocalFiles = GrootApi.getLocalizationFile('gear');
  if (gearLocalFiles === false) return;

  for (let index = 0; index < _langIds.length; index++) {
    const langId = _langIds[index].toLowerCase();

    if (!gearLocalFiles[langId]) continue; //language url is missing
    // Fetch the CSVs from the URLs returned by the API
    const res = UrlFetchApp.fetch(gearLocalFiles[langId]);
    const content = res.getContentText();
    const csv = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    let row = 0;
    for (let r = 1; r < csv.length; r++) {
      let id = csv[r][0];
      if (id.indexOf('ID_GEAR_') != 0 || id.indexOf('_NAME') < 0) continue;
      id = id.substring(8, id.indexOf('_NAME'));

      if (id == 'T1_PLACEHOLDERGEAR') continue;
      //if (id.startsWith('RED_') && GearTierCap < 19) continue;
      //if (id.startsWith('WHITE_') && GearTierCap < 21) continue;

      // if (id.startsWith('T') && !id.startsWith('Teal')) {
      //   const parts = id.split('_');
      //   const tier = Number(parts[0].substring(1));
      //   if (tier > GearTierCap) continue;
      // }

      const value = csv[r][1];
      if (value == 'No one uses THIS gear piece') continue;
      if (id == 'T1_TECH_RESIST') continue; //Scopely removed 'No one uses THIS gear piece on that piece' in ZHS folder... lol

      if (!gearIndex.hasOwnProperty(id)) {
        data[row] = [id];
        gearIndex[id] = row++;
      }
      data[gearIndex[id]][index + 1] = value;
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
