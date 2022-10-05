// -------------------------------------------------
// MSF ToolBot Import
// Author: Zarathos based on chronolinq's work
// -------------------------------------------------

// GLOBAL
// columnOf
// sourceRow
// classOf

function tb_value(columnId) {
  const val = '' + sourceRow[columnOf[columnId]];

  switch (val.toLowerCase()) {
    case 'none':
      return '';
    case 'max':
      return '';
    case 'y':
      return true;
    case 'n':
      return false;
    case 'stri':
      return classOf['ASSASSIN'];
    case 'fort':
      return classOf['FORTIFY'];
    case 'heal':
      return classOf['RESTORATION'];
    case 'skir':
      return classOf['SKIRMISH'];
    case 'raid':
      return classOf['GAMBLER'];
  }

  return val;
}

function msftoolbot_import() {
  const toolbotSheetId = getNamedRangeValue('Preferences_ToolBot_SheetId');

  try {
    // Get ToolBot Range and Values
    const toolBotSheet = SpreadsheetApp.openById(toolbotSheetId).getSheetByName('Roster');
    sourceData = toolBotSheet.getRange(1, 1, toolBotSheet.getMaxRows(), toolBotSheet.getMaxColumns()).getValues();
  } catch (err) {
    Browser.msgBox(
      'Cannot get MSF ToolBot data',
      'Please ensure the MSF ToolBot Sheet ID in the Preferences sheet is correct',
      Browser.Buttons.OK
    );
    return;
  }

  const importIds = getNamedRangeValues('Roster_Import_Ids');
  const importData = getNamedRangeValues('Roster_Import_Data');
  const rosterNotesRange = getNamedRange('Roster_Notes');
  const formulas = rosterNotesRange.getFormulas();
  const classData = getNamedRangeValues('_Option_Class');
  classOf = {};
  classOf[''] = '';
  for (let i = 0; i < classData.length; i++) {
    classOf[classData[i][0]] = classData[i][1];
  }

  // Rely on the headers to make sure it's futureproof
  columnOf = {};
  sourceRow = [];

  for (let i = 0; i < sourceData[0].length; i++) {
    columnOf[sourceData[0][i]] = i;
  }

  for (let i = 1; i < sourceData.length; i++) {
    sourceRow = sourceData[i];
    const id = tb_value('ID');
    const gearPieces = [
      tb_value('TopLeft'),
      tb_value('MidLeft'),
      tb_value('BottomLeft'),
      tb_value('TopRight'),
      tb_value('MidRight'),
      tb_value('BottomRight')
    ];
    const abilities = [tb_value('Basic'), tb_value('Special'), tb_value('Ultimate'), tb_value('Passive')];
    const redStars = Number(tb_value('UnclaimedRed')) + Number(tb_value('Red'));
    const iso8class = tb_value('IsoClass');

    // Find current position of this character
    let s = -1;
    let added = true;
    for (let r = 0; r < importIds.length; r++) {
      if (importIds[r][0] == id) {
        s = r;
        added = false;
        break;
      } else if (s == -1 && importIds[r] == '') s = r;
    }

    // Didn't find and no room to add more characters
    if (s == -1) continue;

    if (added) importIds[s] = [id];

    importData[s][0] = tb_value('Yellow');
    importData[s][1] = redStars == 0 ? '' : redStars;
    importData[s][2] = tb_value('Level');
    importData[s][3] = tb_value('Tier');

    for (let g = 0; g < gearPieces.length; g++) {
      if (gearPieces[g] != '') importData[s][4 + g] = gearPieces[g];
    }

    if (iso8class != '') {
      const isoMatrix = Number(tb_value('IsoPips'));
      importData[s][10] = iso8class;

      if (isoMatrix <= 5) {
        importData[s][11] = 'GREEN';
        importData[s][12] = isoMatrix;
      } else if (isoMatrix <= 10) {
        importData[s][11] = 'BLUE';
        importData[s][12] = isoMatrix - 5;
      } else importData[s][12] = isoMatrix;
    }

    for (let a = 0; a < gearPieces.length; a++) {
      if (abilities[a] != '') importData[s][18 + a] = abilities[a];
    }

    importData[s][22] = tb_value('Power');
    importData[s][23] = tb_value('Fragments');
  }

  setNamedRangeValues('Roster_Import_Ids', importIds);
  setNamedRangeValues('Roster_Import_Data', importData);

  for (let row = 0; row < formulas.length; row++) {
    const formula = formulas[row][0];
    if (formula) {
      rosterNotesRange.getCell(row + 1, 1).setFormula(formula);
    }
  }
}
