function FilterTeam() {
  const activeDoc = SpreadsheetApp.getActive();
  const filter = activeDoc.getSheetByName('_Roster_Filter').getDataRange().getValues();
  const rosterSheet = activeDoc.getSheetByName('Roster');

  // Clear filter
  if (filter[0][0] == '') {
    rosterSheet.showRows(3, rosterSheet.getLastRow() - 2);
    return;
  }

  rosterSheet.hideRows(3, rosterSheet.getLastRow() - 2);
  const toons = getNamedRangeValues('Roster_Id');

  for (let i = 0; i < filter.length; i++) {
    if (filter[i][0] == '') return;

    for (let k = 0; k < toons.length; k++) {
      if (toons[k][0] == filter[i][0]) {
        rosterSheet.showRows(3 + k);
      }
    }
  }
}
