const _saveVersion = 1;
let _saveErrorCount;
let _saveErrorMsg;

function saveJSON() {
  saveJSON_Base(_grootFilename);
}

function saveJSON_Base(filename) {
  const folder = getFolder(_zaratoolsFolder, true);
  const file = getFile(folder, filename);
  if (file != null) file.setTrashed(true);

  _saveErrorCount = 0;
  _saveErrorMsg = '';

  const json = getJSON();

  const jsonString = JSON.stringify(json);
  folder.createFile(filename, jsonString);

  if (_saveErrorCount > 0)
    SpreadsheetApp.getUi().alert(`Warning: Some errors occured in the process\n\n${_saveErrorMsg}`);
}

function getJSON() {
  _saveErrorCount = 0;
  const sheet = SpreadsheetApp.getActive().getSheetByName('_Export');
  _data = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).getValues();

  const json = {};

  try {
    json.Header = getJSON_Header();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Profile = getJSON_Profile();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.GROOT = getJSON_GROOT();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.MSFgg = { UserID: _data[30][4], Key: _data[31][4] };
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.MSFToolBot = { SheetId: _data[40][3] };
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Roster = getJSON_Roster();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.StarkTech = getJSON_StarkTech();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Teams = getJSON_Teams();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.War = getJSON_War();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Raid = getJSON_Raid();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Blitz = getJSON_Blitz();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Mission = getJSON_Mission();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Inventory = getJSON_Inventory();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Farming = getJSON_Farming();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  try {
    json.Links = getJSON_Links();
  } catch (e) {
    _saveErrorCount++;
    _saveErrorMsg += `${e}\n`;
  }

  return json;
}

// Header
function getJSON_Header() {
  let header = {};

  header = {};
  header.Signature = 'MSFPlayerData';
  header.Version = _saveVersion;
  header.Date = Utilities.formatDate(new Date(), 'GMT-04:00', 'yyyy/MM/dd');
  header.Source = 'GROOT';
  header.Source.Version = _data[_version_Y][_version_X];

  return header;
}

// Preferences
function getJSON_Profile() {
  const profile = {};

  profile.Name = _data[_profile_Y][_profile_X];
  profile.Portrait = _data[_profile_Y + 1][_profile_X];
  profile.Id = _data[_profile_Y + 2][_profile_X];
  profile.ResetTime = _data[_profile_Y + 3][_profile_X];
  profile.MVPCount = _data[_profile_Y + 4][_profile_X];
  profile.ArenaRank = _data[_profile_Y + 5][_profile_X];
  profile.ArenaTopRank = _data[_profile_Y + 6][_profile_X];
  profile.BlitzWins = _data[_profile_Y + 7][_profile_X];
  profile.Language = _data[_pref_Y][_pref_X];

  return profile;
}

// Roster Organizer Preferences
function getJSON_GROOT() {
  const ro = { Preferences: { Colors: { Duplicates: {}, Settings: {}, Teams: { Index: [], Id: [] } }, Images: {} } };

  ro.Build = _data[_version_Y + 1][_version_X];
  ro.SheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

  ro.Preferences.NotifyBeta = _data[_pref_Y + 1][_pref_X];

  ro.Preferences.WarDuplicate = _data[_pref_Y + 2][_pref_X];
  ro.Preferences.BlitzDuplicate = _data[_pref_Y + 3][_pref_X];

  ro.Preferences.Colors.Heroes = _data[_pref_Y + 11][_pref_X];
  ro.Preferences.Colors.Gear = _data[_pref_Y + 12][_pref_X];
  ro.Preferences.Colors.Level = _data[_pref_Y + 13][_pref_X];
  ro.Preferences.Colors.Abilities = _data[_pref_Y + 14][_pref_X];
  ro.Preferences.Colors.Shards = _data[_pref_Y + 15][_pref_X];
  ro.Preferences.GearTierOverride = _data[81][_pref_X];

  for (let r = 0; r < _pref_color_teams_H; r++) {
    ro.Preferences.Colors.Teams.Index[r] = _data[_pref_color_teams_Y + r][_pref_color_teams_X];
    ro.Preferences.Colors.Teams.Id[r] = _data[_pref_color_teams_Y + r][_pref_color_teams_X + 1];
  }
  ro.Preferences.Images.Portraits = _data[_pref_Y + 4][_pref_X];
  ro.Preferences.Images.Gear = _data[_pref_Y + 5][_pref_X];
  ro.Preferences.Images.Teams = _data[_pref_Y + 6][_pref_X];
  ro.Preferences.Images.Stars = _data[_pref_Y + 7][_pref_X];
  ro.Preferences.Images.Resource = _data[_pref_Y + 8][_pref_X];

  return ro;
}

function getJSON_Roster() {
  const roster = { Hero: {} };
  const _rosterData = getNamedRangeValues('Roster_Data');

  const classData = getNamedRangeValues('_Option_Class');
  const classOf = {};
  classOf[''] = '';
  for (let i = 0; i < classData.length; i++) {
    classOf[classData[i][1]] = classData[i][0];
  }

  roster.FlagNames = [
    _data[_roster_Y + 0][_roster_X],
    _data[_roster_Y + 1][_roster_X],
    _data[_roster_Y + 2][_roster_X],
    _data[_roster_Y + 3][_roster_X]
  ];

  roster.TeamFilter = _data[_roster_Y + 10][_roster_X];

  roster.Filters = [
    _data[_roster_Y + 9][_roster_X],
    _data[_roster_Y + 8][_roster_X],
    _data[_roster_Y + 7][_roster_X],
    _data[_roster_Y + 6][_roster_X],
    _data[_roster_Y + 5][_roster_X],
    _data[_roster_Y + 4][_roster_X]
  ];

  for (let y = 0; y < _rosterData.length; y++) {
    const line = _rosterData[y];
    const id = line[0];
    if (id == '' || id == null) continue;

    const hero = {};
    hero.Id = id;
    hero.Favorite = line[1];
    hero.StarLevel = line[17];
    hero.RedStarLevel = line[18];
    hero.DiamondLevel = line[19];
    hero.Level = line[20];
    hero.GearTier = line[21];
    hero.GearParts = [line[22], line[23], line[24], line[25], line[26], line[27]];

    hero.Iso8 = {};
    hero.Iso8.SkillId = classOf[line[28]];
    hero.Iso8.IsoMatrix = line[29]; //todo: localisation
    hero.Iso8.MatrixQuality = line[30];
    hero.Iso8.Stats = [line[31], line[32], line[33], line[34], line[35]];

    hero.Abilities = [line[36], line[37], line[38], line[39]];
    hero.Power = line[40];
    hero.Shards = line[41];
    hero.Notes = line[42];
    hero.Flags = [line[43], line[44], line[45], line[46]];

    roster.Hero[id] = hero;
  }

  return roster;
}

function getJSON_StarkTech() {
  const starkTech = { Boosts: {}, Donations: { Alliance: { Usual: [], Special: [] }, Player: [], Days: {} } };

  let pos = _starktech_Y;
  for (let x = 0; x < _originType.length; x++) {
    starkTech.Boosts[_originType[x]] = {};
    for (let y = 0; y < _statsType.length; y++) {
      starkTech.Boosts[_originType[x]][_statsType[y]] = _data[pos++][_starktech_X];
    }
  }

  starkTech.XP = _data[_starktech_Y + 25][_starktech_X];
  starkTech.Credits = _data[_starktech_Y + 26][_starktech_X];

  const date = _data[_starktech_Y + 27][_starktech_X];
  try {
    starkTech.Date = date.getTime();
  } catch (e) {
    error(e);
  }

  starkTech.Donations.Alliance.Usual = [
    _data[_starktech_Y + 28][_starktech_X],
    _data[_starktech_Y + 29][_starktech_X],
    _data[_starktech_Y + 30][_starktech_X]
  ];
  starkTech.Donations.Alliance.Special = [
    _data[_starktech_Y + 31][_starktech_X],
    _data[_starktech_Y + 32][_starktech_X],
    _data[_starktech_Y + 33][_starktech_X]
  ];

  starkTech.Donations.Days.Normal = _data[_starktech_Y + 34][_starktech_X];
  starkTech.Donations.Days.Special = _data[_starktech_Y + 35][_starktech_X];

  starkTech.Donations.Player = [_data[_starktech_Y + 36][_starktech_X], _data[_starktech_Y + 37][_starktech_X]];

  starkTech.MinPower = _data[_starktech_Y + 38][_starktech_X];

  return starkTech;
}

function getJSON_Teams() {
  const teams = [];

  let x = 0;
  let y = 0;
  let ty = 0;
  for (let t = 0; t < _teams_count; t++) {
    if (ty == _teams_H) {
      ty = 0;
      y = 0;
      x++;
    }
    ty++;

    const team = { Hero: [] };
    let empty = true;
    for (let h = 0; h < _team_size; h++) {
      team.Hero[h] = _data[_teams_Y + y][_teams_X + x];
      if (team.Hero[h] != '') empty = false;
      y++;
    }

    if (empty) team.Hero = null;

    team.Name = _data[_teams_names_Y + t][_teams_names_X];
    team.Blitz = _data[_teams_blitz_Y + t][_teams_blitz_X];
    team.War = _data[_teams_war_Y + t][_teams_war_X];
    teams[t] = team;
  }

  const availableStats = [];
  const teamStats = [];

  for (let i = 0; i < 6; i++) {
    availableStats.push(_data[_teams_filterstats_Y + i][_teams_filterstats_X]);
  }

  for (let i = 0; i < 3; i++) {
    teamStats.push(_data[_teams_filterstats_Y + 6 + i][_teams_filterstats_X]);
  }

  const jsonTeams = {
    Teams: teams,
    Filters: {
      Available: _data[_teams_filters_Y][_teams_filters_X],
      AvailableStats: availableStats,
      TeamsStats: teamStats
    }
  };

  return jsonTeams;
}

function getJSON_War() {
  const war = { Defenders: [], Attackers: [] };

  // Defenders
  for (let t = 0; t < _war_defense_teams; t++) {
    war.Defenders[t] = { Name: _data[_war_defense_team_Y + t][_war_defense_team_X] };
  }

  // Attackers
  for (let t = 0; t < _war_offense_teams; t++) {
    war.Attackers[t] = { Name: _data[_war_offense_team_Y + t][_war_offense_team_X] };
  }

  return war;
}

function getJSON_Raid() {
  const raid = { Filter: {} };

  const R = ['Ultimus', 'Alpha', 'Beta', 'Gamma'];
  const X = [_raid_ultimus_X, _raid_alpha_X, _raid_beta_X, _raid_gamma_X];
  const Y = [_raid_ultimus_Y, _raid_alpha_Y, _raid_beta_Y, _raid_gamma_Y];
  //var H = [_raid_ultimus_H, _raid_alpha_H, _raid_beta_H, _raid_gamma_H];

  const L = [_raid_ultimus_teams, _raid_alpha_teams, _raid_beta_teams, _raid_gamma_teams];

  for (let r = 0; r < R.length; r++) {
    const raidname = R[r];
    const x = X[r];
    let y = Y[r];
    const teams = L[r];
    raid[raidname] = {};

    for (let t = 0; t < teams.length; t++) {
      const team = { Hero: [] };
      let empty = true;

      for (let h = 0; h < _team_size; h++) {
        team.Hero[h] = _data[y++][x];
        if (team.Hero[h] != '') empty = false;
      }

      if (!empty) {
        raid[raidname][teams[t]] = team;
      }
    }
  }

  raid.Filter.Raid = _data[_raid_filters_Y + 0][_raid_filters_X];
  raid.Filter.Lane = _data[_raid_filters_Y + 1][_raid_filters_X];

  return raid;
}

function getJSON_Blitz() {
  const blitz = {
    Filter: _data[_blitz_filter_Y][_blitz_filter_X],
    Rampup: { Team: [] },
    FreeCycle: {
      Count: _data[_blitz_plan_Y][_blitz_plan_X],
      StartTier: _data[_blitz_plan_Y + 2][_blitz_plan_X],
      StartFight: _data[_blitz_plan_Y + 3][_blitz_plan_X],
      Team: []
    },
    DailyExtra: {
      Count: _data[_blitz_plan_Y + 1][_blitz_plan_X],
      Team: []
    },
    Team: []
  };

  for (let r = 0; r < _blitz_team_H; r++) {
    const name = _data[_blitz_team_Y + r][_blitz_team_names_X];
    const tier = _data[_blitz_team_Y + r][_blitz_team_tier_X];

    blitz.Team[r] = { Name: name, Tier: tier };
  }

  for (let r = 0; r < _blitz_plan_rampup_H; r++) {
    blitz.Rampup.Team[r] = {
      Win: _data[_blitz_plan_rampup_Y + r][_blitz_plan_rampup_X],
      Name: _data[_blitz_plan_rampup_Y + r][_blitz_plan_rampup_X + 1]
    };
  }

  for (let r = 0; r < _blitz_plan_cycle_H; r++) {
    blitz.FreeCycle.Team[r] = {
      Win: _data[_blitz_plan_cycle_Y + r][_blitz_plan_cycle_X],
      Name: _data[_blitz_plan_cycle_Y + r][_blitz_plan_cycle_X + 1]
    };
  }

  for (let r = 0; r < _blitz_plan_extra_H; r++) {
    blitz.DailyExtra.Team[r] = {
      Tier: _data[_blitz_plan_extra_Y + r][_blitz_plan_extra_X],
      Fight: _data[_blitz_plan_extra_Y + _blitz_plan_extra_H + r][_blitz_plan_extra_X],
      Name: _data[_blitz_plan_extra_Y + 2 * _blitz_plan_extra_H + r][_blitz_plan_extra_X]
    };
  }

  return blitz;
}

function getJSON_Mission() {
  const mission = { Event: [] };

  for (let r = 0; r < _mission_eventCount; r++) {
    const heroes = [];
    for (let h = 0; h < 5; h++) {
      heroes[h] = _data[_mission_heroes_Y + h + r * 5][_mission_heroes_X];
    }
    mission.Event[r] = {
      Type: _data[_mission_eventType_Y + r][_mission_eventType_X],
      Name: _data[_mission_eventName_Y + r][_mission_eventName_X],
      Tier: _data[_mission_eventTier_Y + r][_mission_eventTier_X],
      Hero: heroes
    };
  }

  return mission;
}

function getJSON_Inventory() {
  const inventory = { CountFlag: _data[_inventory_countflag_Y][_inventory_countflag_X] };
  inventory.mats = {};
  inventory.crafted = {};
  inventory.iso = {};

  // ISOITEM_GREEN_PROTECTOR_ARMOR_1
  const roles = ['PROTECTOR', 'BLASTER', 'SUPPORT', 'BRAWLER', 'CONTROLLER'];
  const stats = ['ARMOR', 'RESIST', 'HEALTH', 'FOCUS', 'DAMAGE'];

  let index = 0;
  //ISO 8
  for (let level = 0; level < 5; level++) {
    for (let role = 0; role < roles.length; role++) {
      for (let stat = 0; stat < stats.length; stat++) {
        inventory.iso[`ISOITEM_GREEN_${roles[role]}_${stats[stat]}_${level + 1}`] = _data[index][_inventory_iso8_X];

        index++;
      }
    }
  }

  //Inventory materials
  for (let col = 0; col < 2; col++) {
    for (let i = 0; i < _data.length; i++) {
      const qt = _data[i][_inventory_main_X + 1 + col * 2];
      const id = _data[i][_inventory_main_X + col * 2];

      if (qt === '' || qt <= 0 || id === '') continue;

      inventory.mats[id] = qt;
    }
  }

  //Crafted gear
  for (let i = 0; i < _data.length; i++) {
    const qt = _data[i][_inventory_crafted_X + 1];
    const id = _data[i][_inventory_crafted_X];

    if (qt === '' || qt <= 0 || id === '') continue;

    inventory.crafted[id] = qt;
  }

  return inventory;
}

function getJSON_Farming() {
  const names = getNamedRangeValues('Farming_Import_List');
  const target = getNamedRangeValues('Farming_Import_Target');
  setIdConverter('_M3Localization_Heroes_Name', '_M3Localization_Heroes_Id');

  const farming = { Hero: {} };

  for (let r = 0; r < names.length; r++) {
    const name = names[r][2];
    if (name == '') continue;

    const hero = {};
    hero.Priority = names[r][0];
    hero.Index = names[r][1];

    hero.Target = {
      StarLevel: target[r][0],
      RedStarLevel: target[r][1],
      DiamondLevel: target[r][2],
      Level: target[r][3],
      GearTier: target[r][4],
      Iso8: target[r][5],
      Abilities: [target[r][6], target[r][7], target[r][8], target[r][9]]
    };
    farming.Hero[valueOf(name)] = hero;
  }

  return farming;
}

function getJSON_Links() {
  const links = [];

  for (let r = 0; r < _links_H; r++) {
    const link = { Title: _data[_links_Title_Y + r][_links_X], URL: _data[_links_URL_Y + r][_links_X] };
    links.push(link);
  }

  return links;
}
