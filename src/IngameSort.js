function IngameSort() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Roster');
  const shards = sheet.getRange('Roster_Shards').getColumn();
  const power = sheet.getRange('Roster_Power').getColumn();
  const fav = sheet.getRange('Roster_Fav').getColumn();
  const range = getNamedRange('Roster_Data');
  range.sort([
    { column: fav, ascending: false },
    { column: power, ascending: false },
    { column: shards, ascending: false }
  ]);
}
