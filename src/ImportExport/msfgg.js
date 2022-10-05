function msfgg_saveData() {
  const params = { method: 'get', headers: { Authorization: `Bearer ${msfggtoken}` } };
  const url = `https://api.msf.gg/services/getRoster?userGuid=${getNamedRangeValue('Preferences_MSFgg_Id')}`;
  const response = UrlFetchApp.fetch(url, params);
  const json = JSON.parse(response.getContentText());
  const folder = getFolder(_zaratoolsFolder, true);
  folder.createFile('msfggtest.json', JSON.stringify(json));
}

function msfgg_exportWarDefense() {
  let csvtext = 'position,character1,character2,character3,character4,character5\r\n';

  const folder = getFolder(_zaratoolsFolder, true);
  const playerName = getNamedRangeValue('Profile_Name');
  const warDefenseIds = getNamedRangeValues('_War_Heroes_Defense');

  for (let team = 0; team < 8; team++) {
    for (let hero = 0; hero < 5; hero++) {
      csvtext += warDefenseIds[team * 5 + hero];
      if (hero < 4) csvtext += ',';
    }
    csvtext += '\r\n';
  }

  const filename = `${playerName}-war-defense.csv`;
  folder.createFile(filename, csvtext);
  Browser.msgBox(`File "${filename}" is waiting in your Drive.\nCheck https://drive.google.com/drive/my-drive`);
}

function msfgg_import() {
  const params = { method: 'get', headers: { Authorization: `Bearer ${msfggtoken}` } };
  const url = `https://api.msf.gg/services/getRoster?userGuid=${getNamedRangeValue('Preferences_MSFgg_Id')}`;
  const response = UrlFetchApp.fetch(url, params);

  const roster = JSON.parse(response.getContentText());
  //const roster = getFileContent(); //to test

  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Import_Data');
  setIdConverter('_Option_Class_Id', '_Option_Class_Name');
  const rosterNotesRange = getNamedRange('Roster_Notes');
  const formulas = rosterNotesRange.getFormulas();

  const characters = Object.keys(roster);

  for (let i = 0; i < characters.length; i++) {
    const characterId = characters[i];
    const hero = roster[characterId];

    if (characterId === 'undefined') continue;
    let s = -1;
    for (let r = 0; r < rosterIds.length; r++) {
      if (rosterIds[r][0] == characterId) {
        s = r;
        break;
      } else if (s == -1 && rosterIds[r][0] == '') s = r;
    }
    rosterIds[s][0] = characterId;
    if (hero.favorite == 1) rosterIds[s][1] = true;
    else rosterIds[s][1] = false;

    rosterData[s][0] = hero.stars == 0 ? '' : hero.stars;
    rosterData[s][1] = hero.redStars == 0 ? '' : hero.redStars;
    rosterData[s][2] = hero.level == 0 ? '' : hero.level;
    rosterData[s][3] = hero.tier == 0 ? '' : hero.tier;

    for (let gearPart = 1; gearPart <= 6; gearPart++) {
      const slot = `slot${gearPart}`;
      const equipped = hero[slot];
      rosterData[s][4 + (gearPart - 1)] = Number(equipped) == 1;
    }

    rosterData[s][10] = !hero.isoSkillId ? '' : valueOf(hero.isoSkillId);
    rosterData[s][11] = !hero.isoSkillId ? '' : hero.isoMatrix == '' ? 'GREEN' : hero.isoMatrix.toUpperCase(); //todo: localisation
    rosterData[s][12] = hero.isoMatrixQuality == 0 ? '' : hero.isoMatrixQuality;
    rosterData[s][13] = hero.isoMatrixQuality_Armor == 0 ? '' : hero.isoMatrixQuality_Armor;
    rosterData[s][14] = hero.isoMatrixQuality_Resist == 0 ? '' : hero.isoMatrixQuality_Resist;
    rosterData[s][15] = hero.isoMatrixQuality_Health == 0 ? '' : hero.isoMatrixQuality_Health;
    rosterData[s][16] = hero.isoMatrixQuality_Focus == 0 ? '' : hero.isoMatrixQuality_Focus;
    rosterData[s][17] = hero.isoMatrixQuality_Damage == 0 ? '' : hero.isoMatrixQuality_Damage;

    rosterData[s][18] = hero.basic == 0 ? '' : hero.basic;
    rosterData[s][19] = hero.special == 0 ? '' : hero.special;
    rosterData[s][20] = hero.ultimate == 0 ? '' : hero.ultimate;
    rosterData[s][21] = hero.passive == 0 ? '' : hero.passive;

    rosterData[s][22] = hero.power;
    rosterData[s][23] = hero.shards;
  }

  setNamedRangeValues('Roster_Import_Ids', rosterIds);
  setNamedRangeValues('Roster_Import_Data', rosterData);

  for (let row = 0; row < formulas.length; row++) {
    const formula = formulas[row][0];
    if (formula) {
      rosterNotesRange.getCell(row + 1, 1).setFormula(formula);
    }
  }
}

// function msfgg_importInventory()
// {
//     var params = { method: "get", headers: { Authorization: "Bearer " + msfggtoken } };
//     var url = "https://api.msf.gg/services/api/getInventory?ID=" + getNamedRangeValue("Preferences_MSFgg_Id") + "&Key=" + getNamedRangeValue("Preferences_MSFgg_Key");
//     var response = UrlFetchApp.fetch(url, params);

//     var json = JSON.parse(response.getContentText());

//     var itemCount = new Object();
//     itemCount = [];

//     for (var i = 0; i < json.length; i++)
//     {
//         itemCount[json[i].ItemId.toUpperCase()] = json[i].Qty;
//     }

//     for (var column = 1; column <= 5; column++)
//     {
//         var ids = getNamedRangeValues("Inventory_Ids" + column);
//         var values = getNamedRangeValues("Inventory_Values" + column);

//         for (var row = 0; row < ids.length; row++)
//         {
//             var id = ids[row][0];
//             if (itemCount.hasOwnProperty(id)) values[row] = [itemCount[id]];
//         }

//         setNamedRangeValues("Inventory_Values" + column, values);
//     }
// }

// function msfgg_importWarDefense()
// Will need to manage this with teams system
/*
function msfgg_importCSV()
{
    const folder = getFolder(_zaratoolsFolder, false);
    const filename = "msf_gg.csv";
    var file = null;
    if (folder != null)
    {
        const files = folder.getFilesByName(filename);
        if (files != null)
        {
            file = files.next();
        }
    }

    if (file == null)
    {
        Browser.msgBox('Please export your roster on https://msf.gg and copy the file in your Google Drive at this location: "' + _dataFolder + '/' + filename + '" then run this script');
        return;
    }

    var rosterIds = getNamedRangeValues("Roster_Import_Ids");
    var rosterData = getNamedRangeValues("Roster_Import_Data");

    const content = file.getBlob().getDataAsString();
    const input = Utilities.parseCsv(content, ','.charCodeAt(0));

    for (var row = 1; row < input.length; row++)
    {
        const heroData = input[row];
        const id = heroData[0];
        if (id == "") continue;

        var s = -1;
        for (var r = 0; r < rosterIds.length; r++)
        {
            if (rosterIds[r][0] == id)
            {
                s = r;
                break;
            }
            else if (s == -1 && rosterIds[r] == "") s = r;
        }

        if (s < 0) continue;

        rosterIds[s][0] = id;

        rosterData[s][0] = heroData[5]; // STARS
        rosterData[s][1] = heroData[6]; // RED STARS
        rosterData[s][2] = heroData[3]; // LEVEL
        rosterData[s][3] = heroData[4]; // GEAR TIER

        for (var i = 0; i < 6; i++)
            rosterData[s][4 + i] = heroData[7 + i]; // GEAR PARTS

        for (var i = 0; i < 4; i++)
            rosterData[s][17 + i] = heroData[14 + i]; // ABILITIES

        rosterData[s][21] = heroData[2]; // POWER
        rosterData[s][22] = heroData[13]; // SHARDS
    }

    setNamedRangeValues("Roster_Import_Ids", rosterIds);
    setNamedRangeValues("Roster_Import_Data", rosterData);
}

// -------------------------------------------------
// MSF.GG EXPORT (Thanks to Maliciargh)
// -------------------------------------------------
function msfgg_exportCSV()
{
    var csvtext = "id,name,power,level,gearTier,stars,redStars,slot1,slot2,slot3,slot4,slot5,slot6,shards,basic,special,ultimate,passive\r\n";

    const folder = getFolder(_zaratoolsFolder, true);

    const rosterIds = getNamedRangeValues("Roster_Import_Ids");
    const rosterData = getNamedRangeValues("Roster_Import_Data");

    setIdConverter("_M3Localization_Roster_Id", "_M3Localization_Roster_Name");

    for (var row = 0; row < rosterIds.length; row++)
    {
        var id = rosterIds[row][0];
        if (id == "") continue;

        csvtext +=
            id + ',' +
            valueOf(id) + ',' + // NAME
            rosterData[row][21] + ',' + // POWER
            rosterData[row][2] + ',' + // LEVEL
            rosterData[row][3] + ',' + // GEAR
            rosterData[row][0] + ',' + // STAR
            rosterData[row][1]; // RED STAR

        for (var i = 0; i < 6; i++)
            csvtext += "," + rosterData[row][4 + i]; // GEAR PARTS

        csvtext += "," + rosterData[row][22]; // SHARDS

        for (var i = 0; i < 4; i++)
            csvtext += "," + rosterData[row][17 + i]; // ABILITIES

        csvtext += "\r\n";
    }

    var filename = "msf_gg.csv";
    folder.createFile(filename, csvtext);
    Browser.msgBox('File "' + folder + '/' + filename + '" is waiting in your Drive.\nCheck https://drive.google.com/drive/my-drive');
}
*/

function msfgg_export() {
  const roster = [];
  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Import_Data');

  setIdConverter('_Option_Class_Name', '_Option_Class_Id');

  for (let row = 0; row < rosterIds.length; row++) {
    const id = rosterIds[row][0];
    if (id == '') continue;

    let skillId = valueOf(rosterData[row][10]).toLowerCase();
    if (skillId == '') skillId = null;

    const hero = {
      HeroId: id,
      Power: rosterData[row][22] == '' ? 0 : Number(rosterData[row][22]),
      Level: rosterData[row][2] == '' ? 0 : Number(rosterData[row][2]),
      GearLevel: rosterData[row][3] == '' ? 0 : Number(rosterData[row][3]),
      Loyalty: rosterData[row][0] == '' ? 0 : Number(rosterData[row][0]),
      RedStar: rosterData[row][1] == '' ? 0 : Number(rosterData[row][1]),
      Shards: rosterData[row][23] == '' ? 0 : Number(rosterData[row][23]),
      Equipment: [
        rosterData[row][4],
        rosterData[row][5],
        rosterData[row][6],
        rosterData[row][7],
        rosterData[row][8],
        rosterData[row][9]
      ],
      Basic: rosterData[row][18] == '' ? 0 : Number(rosterData[row][18]),
      Special: rosterData[row][19] == '' ? 0 : Number(rosterData[row][19]),
      Ultimate: rosterData[row][20] == '' ? 0 : Number(rosterData[row][20]),
      Passive: rosterData[row][21] == '' ? 0 : Number(rosterData[row][21]),
      IsoSkillId: skillId,
      IsoMatrix: rosterData[row][11] == '' ? 0 : String(rosterData[row][11]).toLowerCase(),
      IsoMatrixQuality: rosterData[row][12] == '' ? 0 : Number(rosterData[row][12]),
      IsoMatrixQuality_Armor: rosterData[row][13] == '' ? 0 : Number(rosterData[row][13]),
      IsoMatrixQuality_Resist: rosterData[row][14] == '' ? 0 : Number(rosterData[row][14]),
      IsoMatrixQuality_Health: rosterData[row][15] == '' ? 0 : Number(rosterData[row][15]),
      IsoMatrixQuality_Focus: rosterData[row][16] == '' ? 0 : Number(rosterData[row][16]),
      IsoMatrixQuality_Damage: rosterData[row][17] == '' ? 0 : Number(rosterData[row][17])
    };

    roster.push(hero);

    //break;
  }

  const data = {
    ID: getNamedRangeValue('Preferences_MSFgg_Id'),
    Key: getNamedRangeValue('Preferences_MSFgg_Key'),
    Roster: roster
  };

  const fetchPost = {
    method: 'post',
    headers: { Authorization: `Bearer ${msfggtoken}` },
    contentType: 'application/json',
    payload: JSON.stringify(data)
  };

  const url = 'https://api.msf.gg/services/api/saveRoster';
  const res = UrlFetchApp.fetch(url, fetchPost);

  output(res);
}
