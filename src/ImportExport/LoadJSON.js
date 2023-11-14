let _loadVersion = 0;
let _loadToolName;
let _isGroot;

// Load saved roster
function loadJSON() {
  loadJSON_Base(_zaratoolsFolder, _grootFilename);
}

// When using update to latest version.
function loadJSON_Update() {
  const docId = SpreadsheetApp.getActive().getId();
  if (loadJSON_Base(_zaratoolsFolder, docId)) return;

  loadJSON_Base('MSF_Groots_Tools', docId);
}

let _loadErrorCount;
let _loadErrorMsg;

function loadJSON_Base(foldername, filename) {
  const folder = getFolder(foldername, false);
  if (folder == null) return false;

  const file = getFile(folder, filename);
  if (file == null) return false;

  const sheet = SpreadsheetApp.getActive().getSheetByName('_Import');
  const sheetRange = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns());
  _data = sheetRange.getValues();

  const content = file.getBlob().getDataAsString();
  const json = JSON.parse(content);

  _loadErrorCount = 0;
  _loadErrorMsg = '';

  setJSON(json);

  sheetRange.setValues(_data);
  AddMissingCharacters();

  if (_loadErrorCount > 0)
    SpreadsheetApp.getUi().alert(`Warning: Some errors occured in the process\n\n${_loadErrorMsg}`);
}

function setJSON(json) {
  _loadErrorCount = 0;

  _loadVersion = json.Header.Version;
  _loadToolName = json.Header.Source;

  _isGroot = _loadToolName == 'GROOT';

  try {
    setJSON_Profile(json.Profile);
  } catch (e) {
    _loadErrorCount++;
    _loadErrorMsg += `${e}\n`;
  }

  try {
    if (_isGroot) {
      setJSON_GROOT(json.GROOT);

      if (json.hasOwnProperty('MSFgg')) {
        _data[30][4] = json.MSFgg.UserID;
        _data[31][4] = json.MSFgg.Key;
      }

      _data[40][3] = json.MSFToolBot.SheetId;
    } else {
      setJSON_GROOT(json.Profile);

      if (json.Profile.hasOwnProperty('MSFPalAllianceId')) _data[_pref_Y + 9][_pref_X] = json.Profile.MSFPalAllianceId;

      if (json.Profile.hasOwnProperty('MSFPalApiKey')) _data[_pref_Y + 10][_pref_X] = json.Profile.MSFPalApiKey;
    }
  } catch (e) {
    _loadErrorCount++;
    _loadErrorMsg += `${e}\n`;
  }

  try {
    setJSON_Roster(json.Roster);
  } catch (e) {
    _loadErrorCount++;
    _loadErrorMsg += `${e}\n`;
  }

  try {
    setJSON_StarkTech(json.StarkTech);
  } catch (e) {
    _loadErrorCount++;
    _loadErrorMsg += `${e}\n`;
  }

  if (_isGroot) {
    try {
      setJSON_Teams(json.Teams);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }

    try {
      setJSON_War(json.War);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }

    try {
      setJSON_Raid(json.Raid);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }

    try {
      if (json.hasOwnProperty('Blitz')) setJSON_Blitz(json.Blitz);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }

    try {
      if (json.hasOwnProperty('Mission')) setJSON_Mission(json.Mission);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }

    try {
      if (json.hasOwnProperty('Links')) setJSON_Links(json.Links);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }
  } else {
    try {
      setJSON_Raid_RO(json.Raid);
    } catch (e) {
      _loadErrorCount++;
      _loadErrorMsg += `${e}\n`;
    }
  }

  try {
    setJSON_Inventory(json.Inventory);
  } catch (e) {
    _loadErrorCount++;
    _loadErrorMsg += `${e}\n`;
  }

  try {
    setJSON_Farming(json.Farming);
  } catch (e) {
    _loadErrorCount++;
    _loadErrorMsg += `${e}\n`;
  }

  try {
    setJSON_CustomSheets(SpreadsheetApp.openById(json.GROOT.SheetId));
  } catch (e) {
    error(e);
  }
}

// Preferences
function setJSON_Profile(profile) {
  _data[_profile_Y + 0][_profile_X] = profile.Name;
  _data[_profile_Y + 1][_profile_X] = _isGroot ? profile.Portrait : profile.Icon;
  _data[_profile_Y + 2][_profile_X] = profile.Id;

  if (_isGroot || profile.hasOwnProperty('ResetTime')) _data[_profile_Y + 3][_profile_X] = profile.ResetTime;

  _data[_profile_Y + 4][_profile_X] = profile.MVPCount;
  _data[_profile_Y + 5][_profile_X] = profile.ArenaRank;
  _data[_profile_Y + 6][_profile_X] = profile.ArenaTopRank;
  _data[_profile_Y + 7][_profile_X] = profile.BlitzWins;

  if (_isGroot) {
    _data[_pref_Y][_pref_X] = profile.Language;
  } else if (profile.Preferences.hasOwnProperty('Language')) {
    setJSON_Lang(profile.Preferences.Language);
  }
}

// Only for the Roster Organizer versions
function setJSON_Lang(langName) {
  switch (langName) {
    case 'Deutsch':
      _data[_pref_Y][_pref_X] = 'DE';
      return;
    case 'English':
      _data[_pref_Y][_pref_X] = 'EN';
      return;
    case 'Español':
      _data[_pref_Y][_pref_X] = 'ES';
      return;
    case 'Français':
      _data[_pref_Y][_pref_X] = 'FR';
      return;
    case 'Português':
      _data[_pref_Y][_pref_X] = 'PT';
      return;
    case 'Svenska':
      _data[_pref_Y][_pref_X] = 'SV';
      return;
    case 'Türkçe':
      _data[_pref_Y][_pref_X] = 'TR';
      return;
    case 'Русский':
      _data[_pref_Y][_pref_X] = 'RU';
      return;
    case '한국어':
      _data[_pref_Y][_pref_X] = 'KO';
      return;
    case '中國語文':
      _data[_pref_Y][_pref_X] = 'ZHT';
      return;
    case '日本語':
      _data[_pref_Y][_pref_X] = 'ZHS';
      return;
    case 'German':
      _data[_pref_Y][_pref_X] = 'DE';
      return;
    default:
      _data[_pref_Y][_pref_X] = 'EN';
  }
}

// Roster Organizer Preferences
function setJSON_GROOT(data) {
  const roColorSettingConvert = {};
  roColorSettingConvert['Current tier'] = 'ID_TIER_CURRENT';
  roColorSettingConvert['Available tier'] = 'ID_TIER_AVAILABLE';
  roColorSettingConvert['None'] = 'ID_NONE';
  roColorSettingConvert['Gear tier'] = 'ID_TIER_CURRENT';
  roColorSettingConvert['Rank up'] = 'ID_RANKUP';
  roColorSettingConvert['Level warning'] = 'ID_LEVEL_WARNING';

  _data[_pref_Y + 1][_pref_X] = data.Preferences.NotifyBeta;
  _data[_pref_Y + 2][_pref_X] = data.Preferences.WarDuplicate;
  _data[_pref_Y + 3][_pref_X] = data.Preferences.BlitzDuplicate;

  _data[_pref_Y + 12][_pref_X] = _isGroot
    ? data.Preferences.Colors.Gear
    : roColorSettingConvert[data.Preferences.ColorGear];
  _data[_pref_Y + 14][_pref_X] = _isGroot
    ? data.Preferences.Colors.Abilities
    : roColorSettingConvert[data.Preferences.ColorAbility];
  _data[81][_pref_X] = _isGroot ? data.Preferences.GearTierOverride : '';

  if (_isGroot) {
    _data[_pref_Y + 11][_pref_X] = data.Preferences.Colors.Heroes;
    _data[_pref_Y + 13][_pref_X] = data.Preferences.Colors.Level;
    _data[_pref_Y + 15][_pref_X] = data.Preferences.Colors.Shards;

    for (let r = 0; r < _pref_color_teams_H; r++) {
      _data[_pref_color_teams_Y + r][_pref_color_teams_X + 0] = data.Preferences.Colors.Teams.Index[r];
      _data[_pref_color_teams_Y + r][_pref_color_teams_X + 1] = data.Preferences.Colors.Teams.Id[r];
    }

    _data[_pref_Y + 4][_pref_X] = data.Preferences.Images.Portraits;
    _data[_pref_Y + 5][_pref_X] = data.Preferences.Images.Gear;
    _data[_pref_Y + 6][_pref_X] = data.Preferences.Images.Teams;
    _data[_pref_Y + 7][_pref_X] = data.Preferences.Images.Stars;
    _data[_pref_Y + 8][_pref_X] = data.Preferences.Images.Resource;
  }
}

function setJSON_Roster(roster) {
  const ids = Object.keys(roster.Hero);
  const count = ids.length;

  const classData = getNamedRangeValues('_Option_Class');
  const classOf = {};
  classOf[''] = '';
  for (let i = 0; i < classData.length; i++) {
    classOf[classData[i][0]] = classData[i][1];
  }

  if (!_isGroot && roster.hasOwnProperty('Language')) {
    setJSON_Lang(roster.Language);
  }

  // Set Flag names
  for (let f = 0; f < 4; f++) {
    _data[_roster_Y + f][_roster_X] = roster.FlagNames[f];
  }

  if (_isGroot) {
    // Set current team filter
    //_data[_roster_Y + 10][_roster_X] = roster.TeamFilter;
    // Disable for now, would need to activate the script and not sure it's something to save anyway

    // Set custom columns values
    for (let f = 0; f < 6; f++) {
      _data[_roster_Y + 9 - f][_roster_X] = roster.Filters[f];
    }
  }

  // Unlike most data, set directly all characters data directly in the roster page to save space on the import / export and manage checkboxes properly
  const rosterIds = [];
  const rosterData = [];

  for (let y = 0; y < count; y++) {
    const hero = roster.Hero[ids[y]];

    if (_isGroot) {
      rosterIds[y] = [hero.Id, hero.Favorite];

      const dataLine = [];

      dataLine.push(hero.StarLevel);
      dataLine.push(hero.RedStarLevel);
      dataLine.push(hero?.DiamondLevel || '');
      dataLine.push(hero.Level);

      dataLine.push(hero.GearTier);
      for (let g = 0; g < 6; g++) dataLine.push(hero.GearParts[g]);

      // Pre-iso8 version
      if (!hero.hasOwnProperty('Iso8')) {
        dataLine.push('');
        dataLine.push('');
        for (let p = 0; p < 5; p++) dataLine.push('');
      }
      // Old versions
      else if (hero.Iso8.hasOwnProperty('Class')) {
        dataLine.push(classOf[hero.Iso8.Class]);
        dataLine.push(hero.Iso8.Level);
        for (let p = 0; p < 5; p++) {
          if (hero.Iso8.Parts[p] == true || hero.Iso8.Parts[p] == false) dataLine.push('');
          else dataLine.push(hero.Iso8.Parts[p]);
        }
      } else {
        dataLine.push(classOf[hero.Iso8.SkillId]);
        dataLine.push(hero.Iso8.IsoMatrix);
        dataLine.push(hero.Iso8.MatrixQuality);
        for (let p = 0; p < 5; p++) {
          dataLine.push(hero.Iso8.Stats[p]);
        }
      }

      for (let a = 0; a < 4; a++) dataLine.push(hero.Abilities[a]);
      dataLine.push(hero.Power);
      dataLine.push(hero.Shards);
      dataLine.push(hero.Notes);
      for (let f = 0; f < 4; f++) dataLine.push(hero.Flags[f]);

      rosterData[y] = dataLine;
    } else {
      rosterIds[y] = [hero.Id, false];

      rosterData[y] = [
        hero.StarLevel,
        hero.RedStarLevel,
        hero.Level,
        hero.GearTier,
        false,
        false,
        false,
        false,
        false,
        false,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        undefined,
        hero.Abilities[0],
        hero.Abilities[1],
        hero.Abilities[2],
        hero.Abilities[3],
        hero.Power,
        hero.Shards,
        hero.Notes,
        hero.Flags[0] == true,
        hero.Flags[1] == true,
        hero.Flags[2] == true,
        hero.Flags[3] == true
      ];
    }
  }

  setNamedRangeValuesResize('Roster_Import_Ids', rosterIds);
  setNamedRangeValuesResize('Roster_Import_Data', rosterData);
}

function setJSON_StarkTech(starkTech) {
  let pos = _starktech_Y;
  for (let x = 0; x < _originType.length; x++) {
    for (let y = 0; y < _statsType.length; y++) {
      _data[pos++][_starktech_X] = starkTech.Boosts[_originType[x]][_statsType[y]];
    }
  }

  _data[_starktech_Y + 25][_starktech_X] = starkTech.XP;
  _data[_starktech_Y + 26][_starktech_X] = starkTech.Credits;

  const date = new Date();
  try {
    date.setTime(parseInt(starkTech.Date));
  } catch (e) {
    error(e);
  }
  _data[_starktech_Y + 27][_starktech_X] = date;

  if (_isGroot) {
    for (let i = 0; i < 3; i++) {
      _data[_starktech_Y + 28 + i][_starktech_X] = starkTech.Donations.Alliance.Usual[i];
      _data[_starktech_Y + 31 + i][_starktech_X] = starkTech.Donations.Alliance.Special[i];
    }

    _data[_starktech_Y + 34][_starktech_X] = starkTech.Donations.Days.Normal;
    _data[_starktech_Y + 35][_starktech_X] = starkTech.Donations.Days.Special;

    _data[_starktech_Y + 36][_starktech_X] = starkTech.Donations.Player[0];
    _data[_starktech_Y + 37][_starktech_X] = starkTech.Donations.Player[1];
  } else {
    for (let i = 0; i < 3; i++) {
      _data[_starktech_Y + 28 + i][_starktech_X] = starkTech.Donations['AllianceUsual'][i];
      _data[_starktech_Y + 31 + i][_starktech_X] = starkTech.Donations['AllianceSpecial'][i];
    }

    _data[_starktech_Y + 34][_starktech_X] = starkTech.SpecialCycle - starkTech.SpecialDays;
    _data[_starktech_Y + 35][_starktech_X] = starkTech.SpecialDays;
  }

  _data[_starktech_Y + 38][_starktech_X] = starkTech.MinPower;
}

function setJSON_Teams(jsonTeams) {
  // Manage checkboxes here considering they can't be set as dynamic references
  // Optim: ignore if all are checked
  const checkcol = [[], [], [], []];
  const colAllTrue = [true, true, true, true];
  const teamPerColumn = 16;
  const teamInterline = 8;

  // By default, all checkboxes are true
  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < (teamPerColumn - 1) * teamInterline + 1; r++) {
      if (r % teamInterline == 0) checkcol[c][r] = [true, true];
      else checkcol[c][r] = ['', ''];
    }
  }

  for (let t = 0; t < jsonTeams.Teams.length && t < teamPerColumn * 4; t++) {
    const team = jsonTeams.Teams[t];
    _data[_teams_names_Y + t][_teams_names_X] = team.Name;
    if (team.Hero != null) {
      for (let h = 0; h < team.Hero.length; h++) {
        _data[_teams_Y + (t % _teams_H) * 5 + h][_teams_X + Math.floor(t / _teams_H)] = team.Hero[h];
      }
    }

    checkcol[t % 4][Math.floor(t / 4) * teamInterline] = [team.Blitz, team.War];

    if (!team.Blitz || !team.War) {
      colAllTrue[t % 4] = false;
    }
  }

  _data[_teams_filters_Y][_teams_filters_X + 0] = jsonTeams.Filters.Available;

  if (jsonTeams.Filters.hasOwnProperty('AvailableStats')) {
    for (let i = 0; i < jsonTeams.Filters.AvailableStats.length; i++) {
      _data[_teams_filterstats_Y + i][_teams_filterstats_X] = jsonTeams.Filters.AvailableStats[i];
    }
  }

  if (jsonTeams.Filters.hasOwnProperty('TeamsStats')) {
    for (let i = 0; i < jsonTeams.Filters.TeamsStats.length; i++) {
      _data[_teams_filterstats_Y + i + 6][_teams_filterstats_X] = jsonTeams.Filters.TeamsStats[i];
    }
  }

  for (let c = 0; c < 4; c++) {
    if (!colAllTrue[c]) setNamedRangeValues(`Teams_Flags${c + 1}`, checkcol[c]);
  }
}

function setJSON_War(war) {
  for (let t = 0; t < _war_defense_teams; t++) {
    _data[_war_defense_team_Y + t][_war_defense_team_X] = war.Defenders[t].Name;
  }

  // Attackers
  for (let t = 0; t < _war_offense_teams; t++) {
    _data[_war_offense_team_Y + t][_war_offense_team_X] = war.Attackers[t].Name;
  }
}

function setJSON_Raid(raid) {
  const R = ['Ultimus', 'Alpha', 'Beta', 'Gamma'];
  const X = [_raid_ultimus_X, _raid_alpha_X, _raid_beta_X, _raid_gamma_X];
  const Y = [_raid_ultimus_Y, _raid_alpha_Y, _raid_beta_Y, _raid_gamma_Y];
  const L = [_raid_ultimus_teams, _raid_alpha_teams, _raid_beta_teams, _raid_gamma_teams];

  for (let r = 0; r < R.length; r++) {
    const x = X[r];
    let y = Y[r];
    const lanes = L[r];

    const raidTeams = raid[R[r]];

    for (let t = 0; t < lanes.length; t++) {
      const team = raidTeams[lanes[t]];
      if (team == null) {
        y += _team_size;
      } else {
        for (let h = 0; h < _team_size; h++) {
          _data[y++][x] = team.Hero[h];
        }
      }
    }
  }
}

function setJSON_Raid_RO(raid) {
  const R = ['Ultimus', 'Alpha', 'Beta', 'Gamma'];
  const X = [_raid_ultimus_X, _raid_alpha_X, _raid_beta_X, _raid_gamma_X];
  const Y = [_raid_ultimus_Y, _raid_alpha_Y, _raid_beta_Y, _raid_gamma_Y];

  const ultimusIndexConvert = [0, 1, 2, 3, 4, 5];
  const alphaIndexConvert = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
  const betaIndexConvert = [0, 1, 2, -1, 5, 6, 7, 8];
  const gammaIndexConvert = [0, 1, 8, 9, 12, 2, 3, 11, 10, 14, 13];

  const IndexOf = [ultimusIndexConvert, alphaIndexConvert, betaIndexConvert, gammaIndexConvert];

  for (let r = 0; r < R.length; r++) {
    const x = X[r];
    const y = Y[r];
    const lanes = IndexOf[r];
    const raidname = R[r];
    const raidTeams = raid[raidname];

    for (let t = 0; t < raidTeams.length; t++) {
      const index = lanes[t];
      if (index < 0) continue;

      const team = raidTeams[t];

      for (let h = 0; h < _team_size; h++) {
        _data[y + h + 5 * index][x] = team[h];
      }
    }
  }
}

function setJSON_Blitz(blitz) {
  for (let r = 0; r < blitz.Team.length; r++) {
    _data[_blitz_team_Y + r][_blitz_team_names_X] = blitz.Team[r].Name;
    _data[_blitz_team_Y + r][_blitz_team_tier_X] = blitz.Team[r].Tier;
  }

  _data[_blitz_filter_Y][_blitz_filter_X] = blitz.Filter;

  const rampup = [];
  for (let r = 0; r < blitz.Rampup.Team.length; r++) {
    rampup[r] = [blitz.Rampup.Team[r].Win, blitz.Rampup.Team[r].Name];
  }
  setNamedRangeValues('Blitz_RampUp', rampup);

  _data[_blitz_plan_Y][_blitz_plan_X] = blitz.FreeCycle.Count;
  _data[_blitz_plan_Y + 1][_blitz_plan_X] = blitz.DailyExtra.Count;
  _data[_blitz_plan_Y + 2][_blitz_plan_X] = blitz.FreeCycle.StartTier;
  _data[_blitz_plan_Y + 3][_blitz_plan_X] = blitz.FreeCycle.StartFight;

  const cycle = [];
  for (let r = 0; r < blitz.FreeCycle.Team.length; r++) {
    cycle[r] = [blitz.FreeCycle.Team[r].Win, blitz.FreeCycle.Team[r].Name];
  }
  setNamedRangeValues('Blitz_Cycle', cycle);

  for (let r = 0; r < blitz.DailyExtra.Team.length; r++) {
    _data[_blitz_plan_extra_Y + r][_blitz_plan_extra_X] = blitz.DailyExtra.Team[r].Tier;
    _data[_blitz_plan_extra_Y + _blitz_plan_extra_H + r][_blitz_plan_extra_X] = blitz.DailyExtra.Team[r].Fight;
    _data[_blitz_plan_extra_Y + 2 * _blitz_plan_extra_H + r][_blitz_plan_extra_X] = blitz.DailyExtra.Team[r].Name;
  }
}

function setJSON_Mission(mission) {
  for (let r = 0; r < mission.Event.length; r++) {
    _data[_mission_eventType_Y + r][_mission_eventType_X] = mission.Event[r].Type;
    _data[_mission_eventName_Y + r][_mission_eventName_X] = mission.Event[r].Name;
    _data[_mission_eventTier_Y + r][_mission_eventTier_X] = mission.Event[r].Tier;

    if (mission.Event[r].hasOwnProperty('Hero')) {
      for (let h = 0; h < mission.Event[r].Hero.length; h++) {
        _data[_mission_heroes_Y + h + r * 5][_mission_heroes_X] = mission.Event[r].Hero[h];
      }
    }
  }
}

function setJSON_Inventory(inventory) {
  let qt1 = 0;
  let qt2 = 0;
  let qt3 = 0;

  // ISOITEM_GREEN_PROTECTOR_ARMOR_1
  const roles = ['PROTECTOR', 'BLASTER', 'SUPPORT', 'BRAWLER', 'CONTROLLER'];
  const stats = ['ARMOR', 'RESIST', 'HEALTH', 'FOCUS', 'DAMAGE'];

  _data[_inventory_countflag_Y][_inventory_countflag_X] = inventory.CountFlag;

  if (Object.keys(inventory.iso).length > 0)
    for (const [key, value] of Object.entries(inventory.iso)) {
      const isoitemsplit = key.split('_');
      const role = roles.indexOf(isoitemsplit[2]);
      let statStr = isoitemsplit[3];
      let level;

      if (isoitemsplit.length === 5) {
        level = Number(isoitemsplit[4]) - 1;
      } // Fix for v0.4.5 and .6 when I was missing a _ character before level
      else {
        level = Number(statStr.substring(statStr.length - 1)) - 1;
        statStr = statStr.substring(0, statStr.length - 1);
      }
      const stat = stats.indexOf(statStr);

      _data[roles.length * stats.length * level + stats.length * role + stat][_inventory_iso8_X] = value;
    }

  if (Object.keys(inventory.mats).length > 0)
    for (const [key, value] of Object.entries(inventory.mats)) {
      if (key.toUpperCase() === 'T1_NOARMOR' || key.split('_').length > 2) {
        _data[qt1][_inventory_main_X] = key;
        _data[qt1][_inventory_main_X + 1] = value;
        qt1++;
      } else {
        _data[qt2][_inventory_unique_X] = key;
        _data[qt2][_inventory_unique_X + 1] = value;
        qt2++;
      }
    }

  if (Object.keys(inventory.crafted).length > 0)
    for (const [key, value] of Object.entries(inventory.crafted)) {
      _data[qt3][_inventory_crafted_X] = key;
      _data[qt3][_inventory_crafted_X + 1] = value;
      qt3++;
    }
}

function setJSON_Farming(farming) {
  const ids = Object.keys(farming.Hero);
  const count = ids.length;

  const names = [];
  const target = [];

  setIdConverter('_M3Localization_Heroes_Id', '_M3Localization_Heroes_Name');

  for (let i = 0; i < count; i++) {
    const id = ids[i];
    if (id == '') {
      names[i] = ['', '', ''];
      target[i] = ['', '', '', '', '', '', '', ''];
      continue;
    }
    const hero = farming.Hero[id];

    if (_isGroot) {
      names[i] = [hero.Priority, hero.Index, valueOf(id)];
      target[i] = [
        hero.Target.StarLevel,
        hero.Target.RedStarLevel,
        hero.Target.DiamondLevel,
        hero.Target.Level,
        hero.Target.GearTier,
        hero.Target.Iso8
      ];
    } else {
      names[i] = [hero.Priority, '', valueOf(id)];
      target[i] = [hero.Target.StarLevel, '', hero.Target.Level, hero.Target.GearTier];
    }

    for (let a = 0; a < 4; a++) {
      target[i][6 + a] = hero.Target.Abilities[a];
    }
  }

  setNamedRangeValuesResize('Farming_Import_List', names);
  setNamedRangeValuesResize('Farming_Import_Target', target);
}

function setJSON_Links(links) {
  for (let r = 0; r < links.length; r++) {
    _data[_links_Title_Y + r][_links_X] = links[r].Title;
    _data[_links_URL_Y + r][_links_X] = links[r].URL;
  }
}

function setJSON_CustomSheets(source) {
  const sheets = source.getSheets();
  const activeDoc = SpreadsheetApp.getActive();

  // Check the name of all the old sheets
  for (let i = 0; i < sheets.length; i++) {
    const name = sheets[i].getName();
    if (name[0] != '*') continue;

    const copy = sheets[i].copyTo(activeDoc).setName(name);
    activeDoc.setActiveSheet(copy);
    activeDoc.moveActiveSheet(sheets[i].getIndex());
  }
}
