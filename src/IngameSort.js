function IngameSort() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Roster');
  var shards = sheet.getRange('Roster_Shards').getColumn();
  var power = sheet.getRange('Roster_Power').getColumn();
  var fav = sheet.getRange('Roster_Fav').getColumn();
  var range = getNamedRange('Roster_Data');
  range.sort([
    { column: fav, ascending: false },
    { column: power, ascending: false },
    { column: shards, ascending: false }
  ]);
}
