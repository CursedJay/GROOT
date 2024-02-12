function rosterSortAsInGame() {
  //const sheet = GetSheet('Roster');
  const shards = getNamedRange('Roster_Shards').getColumn();
  const power = getNamedRange('Roster_Power').getColumn();
  const fav = getNamedRange('Roster_Fav').getColumn();
  const range = getNamedRange('Roster_Data');
  range.sort([
    { column: fav, ascending: false },
    { column: power, ascending: false },
    { column: shards, ascending: false }
  ]);
}

function ShardsToNextLevelNotes() {
  const rosterNotesRange = getNamedRange('Roster_Notes');
  const rosterNotesData = rosterNotesRange.getValues();
  for (let row = 0; row < rosterNotesRange.getNumRows(); row++) {
    rosterNotesData[row][0] =
      '=IF(AND(INDEX(R[0]:R[0],0,COLUMN(Roster_Stars))>0,INDEX(R[0]:R[0],' +
      '0,COLUMN(Roster_Stars))<7),MAX(INDEX({30;55;80;130;200;300;0},INDEX' +
      '(R[0]:R[0],0,COLUMN(Roster_Stars)))-INDEX(R[0]:R[0],0,COLUMN(Roster_Shards)),0),"")';
  }
  rosterNotesRange.setFormulas(rosterNotesData);
  // var rosterNoteNewTitle = rosterNotesRange.offset(-1, 0, 1);
  // rosterNoteNewTitle.setValue('SHARDS TO NEXT LEVEL'); // Needs localised text
}
