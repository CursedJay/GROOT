let _data;
const _zaratoolsFolder = 'MSF_Zaratools';
const _grootFilename = 'MSF_PlayerData.json';
const _dataSourceFolderId = '1Geyeo5qv0RO5zHor-RXtBiTPeqhpjJQx';
const msfggtoken = '362478fb-cdf7-11ea-87db-0a2363cf333b';
const _finishUpdateTag = ' [Finish Update]';

// Data source
function IsInvalidHero(id, invalidlist) {
  // TODO: USe player_Character from M3HeroSheet.json instead
  const invalidHero = ['DarkPhoenix', 'average', 'Greg', 'SpiderSupreme'];
  if (
    id.indexOf('S_') == 0 ||
    id.indexOf('NUE') >= 0 ||
    id.indexOf('NPC') >= 0 ||
    id.indexOf('Minion') >= 0 ||
    id.indexOf('Operator') >= 0 ||
    id.indexOf('War_') == 0 ||
    id.indexOf('Test') == 0 ||
    id.indexOf('_BBMinn') >= 0 ||
    id.indexOf('PvE_') == 0 ||
    id.indexOf('PVE_') == 0 ||
    id.indexOf('GT_') == 0
  )
    return true;

  if (invalidlist) {
    for (let i = 0; i < invalidHero.length; i++) {
      if (id == invalidHero[i]) {
        return true;
      }
    }
  }
  return false;
}

function GetTraitType(trait) {
  switch (trait) {
    case 'Villain':
    case 'Hero':
      return 'Alignment';

    case 'Bio':
    case 'Tech':
    case 'Skill':
    case 'Mystic':
    case 'Mutant':
      return 'Origin';

    case 'City':
    case 'Global':
    case 'Cosmic':
      return 'Location';

    case 'Protector':
    case 'Controller':
    case 'Brawler':
    case 'Blaster':
    case 'Support':
      return 'CombatRole';

    case 'Minion':
      return 'Minion';

    case '':
    case null:
    case 'Ultron':
      return 'None';

    default:
      return 'Affiliation';
  }
}

// JSON Import / Export
const _roster_X = 2;
const _roster_Y = 83;

const _version_X = 3;
const _version_Y = 30;

const _team_size = 5;

const _teams_X = 0;
const _teams_Y = 0;
const _teams_W = 2;
const _teams_H = 32;
const _teams_count = _teams_W * _teams_H;
const _teams_names_X = 2;
const _teams_names_Y = 96;
const _teams_blitz_X = 3;
const _teams_blitz_Y = 96;
const _teams_war_X = 4;
const _teams_war_Y = 96;
const _teams_filters_X = 2;
const _teams_filters_Y = 95;
const _teams_filterstats_X = 3;
const _teams_filterstats_Y = 41;

const _blitz_filter_X = 2;
const _blitz_filter_Y = 82;

const _blitz_team_names_X = 5;
const _blitz_team_tier_X = 6;
const _blitz_team_Y = 120;
const _blitz_team_H = 40;

const _blitz_plan_X = 6;
const _blitz_plan_Y = 116;

const _blitz_plan_rampup_X = 7;
const _blitz_plan_rampup_Y = 38;
const _blitz_plan_rampup_H = 82;

const _blitz_plan_cycle_X = 7;
const _blitz_plan_cycle_Y = 120;
const _blitz_plan_cycle_H = 40;

const _blitz_plan_extra_X = 9;
const _blitz_plan_extra_Y = 40;
const _blitz_plan_extra_H = 40;

const _mission_eventType_X = 2;
const _mission_eventType_Y = 61;
const _mission_eventCount = 20;
const _mission_eventName_X = 3;
const _mission_eventName_Y = 76;
const _mission_eventTier_X = 4;
const _mission_eventTier_Y = 76;
const _mission_heroes_X = 14;
const _mission_heroes_Y = 0;

const _inventory_main_X = 10;
const _inventory_unique_X = 12;
const _inventory_countflag_X = 9;
const _inventory_countflag_Y = 39;
const _inventory_iso8_X = 15;
const _inventory_crafted_X = 17;

const _links_X = 6;
const _links_Title_Y = 80;
const _links_URL_Y = 98;
const _links_H = 18;

const _raid_filters_X = 6;
const _raid_filters_Y = 80;

const _raid_ultimus_teams = ['1', '2', '3', '4', '5'];
const _raid_ultimus_X = 5;
const _raid_ultimus_Y = 0;

const _raid_alpha_teams = [
  'CITY',
  'GLOBAL_CITY',
  'GLOBAL',
  'GLOBAL_COSMIC',
  'COSMIC',
  'CITY2',
  'GLOBAL_CITY2',
  'GLOBAL2',
  'GLOBAL_COSMIC2',
  'COSMIC2'
];
const _raid_alpha_X = 5;
const _raid_alpha_Y = _raid_ultimus_Y + _raid_ultimus_teams.length * _team_size;

const _raid_beta_teams = [
  'MYSTIC_SKILL',
  'SKILL_TECH',
  'BIO_TECH',
  'MYSTIC_SKILL_TECH',
  'MYSTIC_TECH_SINISTER6',
  'BIO_MUTANT',
  'BIO_MYSTIC',
  'MYSTIC_TECH',
  'MUTANT_MYSTIC'
];
const _raid_beta_X = 5;
const _raid_beta_Y = _raid_alpha_Y + _raid_alpha_teams.length * _team_size;

const _raid_gamma_teams = [
  'AV_GOTG',
  'AV_SV',
  'BH_SHLD',
  'AIM_SHLD',
  'AV_GOTG_SV',
  'HAND_KREE_RAV',
  'AIM_BH_SHLD',
  'DEF_HYD_MERC',
  'HAND_KREE',
  'KREE_RAV',
  'DEF_MERC',
  'DEF_HYD',
  'BH_MERC_XM',
  'GOTG_MERC_XM',
  'DEF_SHLD_WAK',
  'KREE_SV_GOTG'
];
const _raid_gamma_X = 6;
const _raid_gamma_Y = 0;

const _war_defense_teams = 8;
const _war_offense_teams = 20;

const _war_defense_team_X = 4;
const _war_defense_team_Y = 0;
const _war_defense_team_H = _war_defense_teams;

const _war_offense_team_X = 4;
const _war_offense_team_Y = _war_defense_team_Y + _war_defense_team_H;
const _war_offense_team_H = _war_offense_teams;

const _pref_color_teams_X = 2;
const _pref_color_teams_Y = 0;
const _pref_color_teams_H = 30;

const _pref_X = 2;
const _pref_Y = 30;

const _profile_X = 3;
const _profile_Y = 32;

const _starktech_X = 4;
const _starktech_Y = 32;

const _statsType = ['Health', 'Damage', 'Armor', 'Focus', 'Resistance'];
const _statsTypeShort = ['Health', 'Damage', 'Armor', 'Focus', 'Resist'];

const _originType = ['Tech', 'Bio', 'Mystic', 'Skill', 'Mutant'];

const _langIds = ['DE', 'EN', 'ES', 'FR', 'ID', 'IT', 'JA', 'KO', 'NO', 'PT', 'RU', 'SV', 'TH', 'TR', 'ZHS', 'ZHT'];
