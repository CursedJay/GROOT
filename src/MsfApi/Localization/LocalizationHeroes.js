// Will update sheet _M3Localization_Heroes
function DS_Update_Localization_Heroes() {
  const heroIndex = {};
  const data = [[]];

  //Get playable characters list
  const characters = GrootApi.getCharacterList();
  if (characters === false) return;

  const heroesLocalFiles = GrootApi.getLocalizationFile('heroes');
  if (heroesLocalFiles === false) return;

  let row = 0;
  for (let index = 0; index < _langIds.length; index++) {
    const langId = _langIds[index].toLowerCase();

    // Fetch the CSVs from the URLs returned by the API
    const res = UrlFetchApp.fetch(heroesLocalFiles[langId]);
    const content = res.getContentText();
    const csv = Utilities.parseCsv(content, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 0; r < csv.length; r++) {
      const fullid = csv[r][0].toString();
      if (fullid.indexOf('ID_SHARD_') != 0 || fullid.indexOf('_NAME') < 0) continue;
      const cid = fullid.substring(9, fullid.indexOf('_NAME'));

      // fill the appropriate column with the localized value

      const charId = characters.findIndex((char) => cid === char.toUpperCase());
      if (charId === -1) continue;

      if (!heroIndex.hasOwnProperty(cid)) {
        data[row] = [characters[charId]];
        heroIndex[cid] = row++;
      }
      data[heroIndex[cid]][index + 1] = csv[r][1];
    }
  }
  if (!UpdateSortedRangeData(GetSheet('_M3Localization_Heroes'), 1, 3, 2 + data[0].length, 1, data)) return false;

  return DS_Update_Localization_HeroesFormula();
}

function DS_Update_Localization_HeroesFormula() {
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
