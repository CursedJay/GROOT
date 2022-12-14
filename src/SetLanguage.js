let _nameOf;
let _locOf;
let _optionOf;

function GetLocalization(langId) {
  const activeDoc = SpreadsheetApp.getActive();
  _locOf = [];
  _nameOf = [];
  _optionOf = [];

  // Text localization
  let loc = activeDoc.getSheetByName('_Localization').getDataRange().getValues();
  let index = -1;
  for (let x = 1; x < loc[0].length; x++) {
    if (loc[0][x] == langId) {
      index = x;
      break;
    }
  }
  if (index == -1) return false;

  _locOf[''] = '';
  for (let y = 1; y < loc.length; y++) {
    _locOf[loc[y][0]] = loc[y][index];
  }

  // Options localization
  loc = activeDoc.getSheetByName('_Option').getDataRange().getValues();
  for (let y = 0; y < loc.length; y++) {
    _optionOf[loc[y][0]] = loc[y][1];
  }

  // Hero name localization
  loc = activeDoc.getSheetByName('_HeroesLoc').getDataRange().getValues();
  index = -1;
  for (let x = 1; x < loc[0].length; x++) {
    if (loc[0][x] == langId) {
      index = x;
      break;
    }
  }
  if (index == -1) return false;

  _nameOf[''] = '';
  for (let y = 1; y < loc.length; y++) {
    _nameOf[loc[y][0]] = loc[y][index];
  }

  return true;
}

function TestChangeLanguage() {
  ChangeLanguage('Français');
}

function ChangeLanguage(newValue) {
  output(`Change language to ${newValue}`);
  const langIds = getNamedRangeValues('_Option_Language');
  let newLangId = '';

  for (let i = 0; i < langIds.length; i++) {
    if (langIds[i][1] == newValue) {
      newLangId = langIds[i][0];
      break;
    }
  }

  if (newLangId == '') {
    output(`Unknown lang id: ${newValue}`);
    return;
  }

  // Backup data while language hasn't been changed
  json = getJSON();
  // Set the new language
  setNamedRangeValue('Preferences_LanguageId', newLangId);
  output(`Lang id = ${newLangId}`);
  GetLocalization(newLangId);

  // =================== PREFERENCES =========================
  const prefs = json.RosterOrganizer.Preferences;
  // UPDATE COLOR SETTINGS -----------------------------------
  const colorSettings = [];

  output(`${prefs.Colors.Heroes} > ${_locOf[prefs.Colors.Heroes]}`);
  colorSettings[0] = [_locOf[prefs.Colors.Heroes]];
  output(`${prefs.Colors.Gear} > ${_locOf[prefs.Colors.Gear]}`);
  colorSettings[1] = [_locOf[prefs.Colors.Gear]];
  output(`${prefs.Colors.Level} > ${_locOf[prefs.Colors.Level]}`);
  colorSettings[2] = [_locOf[prefs.Colors.Level]];
  output(`${prefs.Colors.Abilities} > ${_locOf[prefs.Colors.Abilities]}`);
  colorSettings[3] = [_locOf[prefs.Colors.Abilities]];
  output(`${prefs.Colors.Shards} > ${_locOf[prefs.Colors.Shards]}`);
  colorSettings[4] = [_locOf[prefs.Colors.Shards]];
  output('4');
  setNamedRangeValues('Preferences_Color_Settings', colorSettings);
  output('5');

  // UPDATE TEAM NAMES ---------------------------------------
  const colorTeam = [];
  for (let r = 0; r < 10; r++) {
    colorTeam[r] = [];
    for (let c = 0; c < 3; c++) {
      colorTeam[r][c * 2 + 0] = prefs.Colors.Teams.Index[c * 10 + r];
      colorTeam[r][c * 2 + 1] = _locOf[prefs.Colors.Teams.Id[c * 10 + r]];
    }
  }

  colorTeam[9][5] = _locOf['ID_DEFAULT'];
  setNamedRangeValues('Preferences_Color_Team', colorTeam);
  // ---------------------------------------------------------

  // =========================== ROSTER ======================
  const rosterFilters = [
    [
      _locOf[json.Roster.TeamFilter],
      '',
      _locOf[json.Roster.Filters[0]],
      _locOf[json.Roster.Filters[1]],
      _locOf[json.Roster.Filters[2]],
      _locOf[json.Roster.Filters[3]],
      _locOf[json.Roster.Filters[4]],
      _locOf[json.Roster.Filters[5]]
    ]
  ];
  setNamedRangeValues('Roster_Filters', rosterFilters);
  // ---------------------------------------------------------

  // =========================== TEAMS =======================
  setNamedRangeValue('Teams_Filter_Available', _locOf[json.Teams.Filters.Available]);
  setNamedRangeValue('Teams_Filter_Tier', _locOf[json.Teams.Filters.Tier]);

  for (let column = 0; column < 4; column++) {
    let y = 0;
    const teamsColumn = [];
    for (let t = 0; t < _teams_count / 4; t++) {
      const team = json.Teams.Teams[t * 4 + column];
      if (y > 0) {
        teamsColumn[y++] = [''];
        teamsColumn[y++] = [''];
      }
      teamsColumn[y++] = [team.Name];
      for (let h = 0; h < _team_size; h++) {
        teamsColumn[y++] = [team.Hero == null ? '' : _nameOf[team.Hero[h]]];
      }
    }

    setNamedRangeValues(`Teams_NamesColumn${column + 1}`, teamsColumn);
  }

  /*
    // =========================== WAR =========================
    var warSheet = _activeDoc.getSheetByName("War");
    
    // War
    SetTeamsLang(warSheet, 4,8, 2,4, 8,5, json.War.Defenders);
    SetTeamsLang(warSheet, 23,8, 3,4, 8,5, json.War.Attackers);  
    */
  // ---------------------------------------------------------

  // =========================== RAID ========================
  const raidFilter = [[_optionOf[json.Raid.Filter.Raid], _optionOf[json.Raid.Filter.Lane]]];
  setNamedRangeValues('Raid_Filter', raidFilter);

  const raidSheet = SpreadsheetApp.getActive().getSheetByName('Raid');
  SetTeamsLang(raidSheet, 4, 7, 1, 5, 8, 4, json.Raid.Ultimus, ['1', '2', '3', '4', '5']);
  SetTeamsLang(raidSheet, 15, 7, 2, 5, 8, 4, json.Raid.Alpha, _raid_alpha_teams);
  SetTeamsLang(raidSheet, 34, 7, 2, 5, 8, 4, json.Raid.Beta, _raid_beta_teams);
  SetTeamsLang(raidSheet, 15, 7, 2, 5, 8, 4, json.Raid.Gamma, _raid_gamma_teams);
}

function SetLanguage(newLangId) {
  let newLangName = '';
  const langIds = getNamedRangeValues('_Option_Language');
  let json;

  for (let i = 0; i < langIds.length; i++) {
    if (langIds[i][0] == newLangId) {
      newLangName = langIds[i][1];
      break;
    }
  }

  if (newLangName == '') {
    output(`Unknown lang id: ${newLangId}`);
    return;
  }

  setNamedRangeValue('Preferences_Language', newLangName);
}

function SetTeamsLang(sheet, row, col, county, countx, offsety, offsetx, teams, ids) {
  for (let x = 0; x < countx; x++) {
    for (let y = 0; y < county; y++) {
      const index = y + x * county;
      if (index >= ids.length) return;

      const team = teams[ids[index]];
      if (team == null) continue;

      const range = sheet.getRange(row + y * offsety, col + x * offsetx, _team_size, 1);
      const values = [];

      for (let r = 0; r < _team_size; r++) {
        output(`*${team.Hero[r]}* => ${_nameOf[team.Hero[r]]}`);
        values[r] = [_nameOf[team.Hero[r]]];
      }
      range.setValues(values);
    }
  }
}
