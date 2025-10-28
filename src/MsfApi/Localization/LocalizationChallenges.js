// Updates sheet _M3Localization_Challenges
function DS_Update_Localization_Challenges() {
  const listIndex = {};
  const data = [];

  const challengesLocalFiles = GrootApi.getLocalizationFile('challenges');
  if (challengesLocalFiles === false) return;

  const missionsLocalFiles = GrootApi.getLocalizationFile('missions');
  if (missionsLocalFiles === false) return;

  const uiLocalFiles = GrootApi.getLocalizationFile('ui');
  if (uiLocalFiles === false) return;

  let row = 0;
  for (let index = 0; index < _langIds.length; index++) {
    const langId = _langIds[index].toLowerCase();

    if (!challengesLocalFiles[langId]) continue; //language url is missing
    // Fetch the CSVs from the URLs returned by the API
    const challengesRes = UrlFetchApp.fetch(challengesLocalFiles[langId]);
    const challengesContent = challengesRes.getContentText();
    const challengesCsv = Utilities.parseCsv(challengesContent, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 1; r < challengesCsv.length; r++) {
      let id = challengesCsv[r][0];

      if (id.includes('NUETEST')) continue;
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
        id.startsWith('ID_TU_') ||
        id.startsWith('ID_MFE_')
      )
        id = id.substring(3);
      else continue;

      if (!listIndex.hasOwnProperty(id)) {
        data[row] = [id];
        listIndex[id] = row++;
      }
      data[listIndex[id]][index + 1] = challengesCsv[r][1];
    }

    if (!missionsLocalFiles[langId]) continue; //language url is missing
    // Fetch the CSVs from the URLs returned by the API
    const missionsRes = UrlFetchApp.fetch(missionsLocalFiles[langId]);
    const missionsContent = missionsRes.getContentText();
    const missionsCsv = Utilities.parseCsv(missionsContent, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 1; r < missionsCsv.length; r++) {
      let id = missionsCsv[r][0];

      if (id.indexOf('_NAME') < 0) continue;
      id = id.substring(0, id.indexOf('_NAME'));

      if (id.startsWith('ID_TU_')) id = id.substring(3);
      else continue;

      if (!listIndex.hasOwnProperty(id)) {
        data[row] = [id];
        listIndex[id] = row++;
      }
      data[listIndex[id]][index + 1] = missionsCsv[r][1];
    }

    if (!uiLocalFiles[langId]) continue; //language url is missing
    // Fetch the UI CSVs from the URLs returned by the API
    const uiRes = UrlFetchApp.fetch(uiLocalFiles[langId]);
    const uiContent = uiRes.getContentText();
    const uiCsv = Utilities.parseCsv(uiContent, ','.charCodeAt(0));

    // Parse lines of the CSV file
    for (let r = 1; r < uiCsv.length; r++) {
      let id = uiCsv[r][0];

      if (
        id != 'ID_GAMEMODE_CAMPAIGNS' &&
        id != 'ID_GAMEMODE_EPIC_CAMPAIGNS' &&
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
      data[listIndex[id]][index + 1] = uiCsv[r][1];
    }
  }

  for (let r = 0; r < data.length; r++) {
    // If there's empty slots anywhere, use the ID as name
    for (let l = 1; l < data[0].length; l++) {
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
