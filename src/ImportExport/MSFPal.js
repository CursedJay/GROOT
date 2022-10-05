function msfpal_import() {
  const allianceId = getNamedRangeValue('Preferences_MSFPal_AllianceId');
  const apiKey = getNamedRangeValue('Preferences_MSFPal_Key');
  const playerName = getNamedRangeValue('Profile_Name');

  const url = `https://msf.pal.gg/rest/v1/alliance/${allianceId}/${playerName}/characters?api-key=${apiKey}`;
  const response = UrlFetchApp.fetch(url);
  const heroes = JSON.parse(response.getContentText());
  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Import_Data');

  setIdConverter('_MSFPal_PalId', '_MSFPal_GameId');

  const rosterNotesRange = getNamedRange('Roster_Notes');
  const formulas = rosterNotesRange.getFormulas();

  const classOf = {};
  const classId = getNamedRangeValues('_Option_Class_Id');
  const className = getNamedRangeValues('_Option_Class_Name');
  classOf[''] = '';
  for (let i = 0; i < classId.length; i++) {
    classOf[classId[i][0].toLowerCase()] = className[i][0];
  }

  for (let i = 0; i < heroes.length; i++) {
    const hero = heroes[i];
    if (hero.id == '') continue;

    let id = valueOf(hero.id);
    if (id == '') id = hero.id;

    if (hero.unlocked) {
      let s = -1;
      for (let r = 0; r < rosterIds.length; r++) {
        if (rosterIds[r][0] == id) {
          s = r;
          break;
        } else if (s == -1 && rosterIds[r][0] == '') s = r;
      }
      rosterIds[s][0] = id;
      if (hero.hasOwnProperty('favorite')) rosterIds[s][1] = hero.favorite;

      rosterData[s][0] = hero.yellowStars;
      rosterData[s][1] = hero.hasOwnProperty('redStars') ? hero.redStars : '';
      rosterData[s][2] = hero.level;
      rosterData[s][3] = hero.gearLevel;

      rosterData[s][10] = hero.iso8Class == null ? '' : classOf[hero.iso8Class];
      rosterData[s][11] = ''; //todo: look how iso color is handled
      rosterData[s][12] = hero.iso8Level == null ? '' : hero.iso8Level;
      rosterData[s][13] = '';
      rosterData[s][14] = '';
      rosterData[s][15] = '';
      rosterData[s][16] = '';
      rosterData[s][17] = '';

      rosterData[s][18] = hero.hasOwnProperty('basic') ? hero.basic : '';
      rosterData[s][19] = hero.hasOwnProperty('special') ? hero.special : '';
      rosterData[s][20] = hero.hasOwnProperty('ultimate') ? hero.ultimate : '';
      rosterData[s][21] = hero.hasOwnProperty('passive') ? hero.passive : '';

      rosterData[s][22] = hero.power;
    }
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

function msfpal_export() {
  const allianceId = getNamedRangeValue('Preferences_MSFPal_AllianceId');
  const apiKey = getNamedRangeValue('Preferences_MSFPal_Key');
  const playerName = getNamedRangeValue('Profile_Name');

  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Import_Data');

  const heroes = [];
  setIdConverter('_MSFPal_GameId', '_MSFPal_PalId');

  const classOf = {};
  const classId = getNamedRangeValues('_Option_Class_Id');
  const className = getNamedRangeValues('_Option_Class_Name');
  classOf[''] = '';
  for (let i = 0; i < classId.length; i++) {
    classOf[className[i][0]] = classId[i][0].toLowerCase();
  }

  for (let r = 0; r < rosterIds.length; r++) {
    if (rosterIds[r][0] == '') continue;

    const hero = {};

    hero.id = valueOf(rosterIds[r][0]);
    if (hero.id == '') hero.id = rosterIds[r][0];

    hero.yellowStars = +rosterData[r][0];
    hero.unlocked = hero.yellowStars > 0;
    if (rosterIds[r][1] == true) hero.favorite = true;

    if (hero.unlocked) {
      const redStars = Number(0 + rosterData[r][1]);
      let level = Number(0 + rosterData[r][2]);
      let gear = Number(0 + rosterData[r][3]);
      const basic = Number(0 + rosterData[r][18]);
      const special = Number(0 + rosterData[r][19]);
      const ultimate = Number(0 + rosterData[r][20]);
      const passive = Number(0 + rosterData[r][21]);
      const power = Number(0 + rosterData[r][22]);
      const iso8class = rosterData[r][10];
      const iso8level = Number(0 + rosterData[r][12]);

      // Prevents fail when level is missing
      if (level == 0) level = 1;
      if (gear == 0) gear = 1;

      hero.redStars = redStars;
      hero.level = level;
      hero.gearLevel = gear;
      hero.power = power;

      if (iso8class != '') hero.iso8Class = classOf[iso8class];
      if (iso8level != 0) hero.iso8Level = iso8level;

      if (basic > 0) hero.basic = basic;
      if (special > 0) hero.special = special;
      if (ultimate > 0) hero.ultimate = ultimate;
      if (passive > 0) hero.passive = passive;
    }

    heroes[heroes.length] = hero;
  }

  const fetchPost = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(heroes)
  };

  output(JSON.stringify(heroes));

  const url = `https://msf.pal.gg/rest/v1/alliance/${allianceId}/${playerName}/characters?api-key=${apiKey}`;
  UrlFetchApp.fetch(url, fetchPost);
}

// -------------------------------------------------
// MSFPAL EXPORT (Thanks to Maliciargh)
// -------------------------------------------------
function msfpal_exportCSV() {
  let csvtext =
    'id,label,unlocked,power,level,gearLevel,yellowStars,redStars,basic,special,ultimate,passive,iso8Class,iso8Level';

  const folder = getFolder(_zaratoolsFolder, true);

  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Import_Data');

  setIdConverter('_MSFPal_GameId', '_MSFPal_PalId');

  const classOf = {};
  const classId = getNamedRangeValues('_Option_Class_Id');
  const className = getNamedRangeValues('_Option_Class_Name');
  classOf[''] = '';
  for (let i = 0; i < classId.length; i++) {
    classOf[className[i][0]] = classId[i][0].toLowerCase();
  }

  for (let row = 0; row < rosterIds.length; row++) {
    if (rosterIds[row][0] == '') continue;

    let id = valueOf(rosterIds[row][0]);
    if (id == '') id = rosterIds[row][0];

    const power = Number(rosterData[row][21]);

    csvtext += `\r\n${id},${
      // ID
      id
    },${
      //LABEL (Replace by name)
      +rosterData[row][0] > 0
    },${
      // UNLOCKED
      power
    },${
      // POWER
      rosterData[row][2]
    },${
      // LEVEL
      rosterData[row][3]
    },${
      // GEAR TIER
      rosterData[row][0]
    },${
      // STAR
      rosterData[row][1]
    },${
      // RED STAR
      rosterData[row][18]
    },${
      // BASIC
      rosterData[row][19]
    },${
      // SPECIAL
      rosterData[row][20]
    },${
      // ULTIMATE
      rosterData[row][21]
    },${
      // PASSIVE
      classOf[rosterData[row][10]]
    },${
      // iso8 class
      rosterData[row][12]
    }`; // iso8 level
  }

  folder.createFile('msfpal.csv', csvtext);
  Browser.msgBox(
    `File "${_zaratoolsFolder}/msfpal.csv" is waiting in your Drive.\nCheck https://drive.google.com/drive/my-drive`
  );
}
