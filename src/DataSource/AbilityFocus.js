function ImportAbilityFocus_0() {
  ImportAbilityFocus(0);
}
function ImportAbilityFocus_1() {
  ImportAbilityFocus(1);
}
function ImportAbilityFocus_2() {
  ImportAbilityFocus(2);
}
function ImportAbilityFocus_3() {
  ImportAbilityFocus(3);
}
function ImportAbilityFocus_4() {
  ImportAbilityFocus(4);
}
function ImportAbilityFocus_5() {
  ImportAbilityFocus(5);
}
function ImportAbilityFocus_6() {
  ImportAbilityFocus(6);
}
function ImportAbilityFocus_7() {
  ImportAbilityFocus(7);
}
function ImportAbilityFocus_8() {
  ImportAbilityFocus(8);
}
function ImportAbilityFocus_9() {
  ImportAbilityFocus(9);
}
function ImportAbilityFocus_10() {
  ImportAbilityFocus(10);
}
function ImportAbilityFocus_11() {
  ImportAbilityFocus(11);
}
function ImportAbilityFocus_12() {
  ImportAbilityFocus(12);
}
function ImportAbilityFocus_13() {
  ImportAbilityFocus(13);
}
function ImportAbilityFocus_14() {
  ImportAbilityFocus(14);
}
function ImportAbilityFocus_15() {
  ImportAbilityFocus(15);
}
function ImportAbilityFocus_16() {
  ImportAbilityFocus(16);
}
function ImportAbilityFocus_17() {
  ImportAbilityFocus(17);
}
function ImportAbilityFocus_18() {
  ImportAbilityFocus(18);
}
function ImportAbilityFocus_19() {
  ImportAbilityFocus(19);
}

function ImportAbilityFocus(index) {
  const folderId = '1hIY1LFnY0wVrBfJw-nCUYoCHzXH-TRHk';
  const abilityFocusFolder = DriveApp.getFolderById(folderId);

  /*var files = abilityFocusFolder.getFilesByName("*.json");
    var file = null;

    while (files.hasNext())
    {
        var f = files.next();
        if (index == 0)
        {
            file = f;
            break;
        }
        index--;
    }

    if (file == null) return;
    */

  const afSources = getNamedRangeValues('_AbilityFocus_Source');
  const author = afSources[index][0];

  const file = abilityFocusFolder.getFilesByName(`${author}.json`).next();
  const content = file.getBlob().getDataAsString();
  const json = JSON.parse(content);

  const teams = json.Teams;

  const data = [];

  data[0] = [json.Author, 'HERO', 'BASIC', 'SPECIAL', 'ULTIMATE', 'PASSIVE', 'SECONDARY'];

  let _index = 1;
  for (let t = 0; t < teams.length; t++) {
    const teamId = teams[t].Id;

    for (let h = 0; h < teams[t].Hero.length; h++) {
      const hero = teams[t].Hero[h];

      output(`Hero = ${hero}`);

      data[_index++] = [
        teamId,
        hero.id,
        hero.Abilities[0],
        hero.Abilities[1],
        hero.Abilities[2],
        hero.Abilities[3],
        hero.Secondary
      ];
    }
  }

  const sheet = GetSheet('_AbilityFocusTable');

  // Clear current data
  ClearSheet(sheet, true);

  // Resize to fit new data
  ResizeSheet(sheet, data.length, data[0].length);

  // Paste new data
  const rangeData = sheet.getRange(1, 1, data.length, data[0].length);
  rangeData.setValues(data);
}
