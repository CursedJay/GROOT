var powermul_table = [];
var statmod_table = [];
var statmod_levels = [];

/// Will update sheet _HeroData and _HeroData_PowerMul
function DS_Update_HeroData() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const combat_dataFolder = versionFolder.getFoldersByName('combat_data').next();
  const heroesFolder = versionFolder.getFoldersByName('heroes').next();

  const characters = combat_dataFolder.getFilesByName('characters.json').next();
  const M3HeroStats = heroesFolder.getFilesByName('M3HeroStats.json').next();
  const M3HeroSheet = heroesFolder.getFilesByName('M3HeroSheet.json').next();
  const M3GearTiers = heroesFolder.getFilesByName('M3GearTiers.csv').next();

  var heroDataSheet = GetSheet('_HeroData');
  var column = {};
  var newData = [];
  var heroIndex = {};

  const columns = [
    'HERO_ID',
    'unlock_StarLevel',
    'player_Character',
    'trait_alignment',
    'trait_origin',
    'trait_area',
    'trait_role',
    'affiliation1',
    'affiliation2',
    'affiliation3',
    'affiliation4',
    'speed',
    'base health',
    'base damage',
    'base armor',
    'base focus',
    'base resist',
    'power_mul_index',
    'Health_mod_index',
    'Damage_mod_index',
    'Armor_mod_index',
    'Focus_mod_index',
    'Resist_mod_index'
  ];

  for (var c = 0; c < columns.length; c++) {
    column[columns[c]] = c;
  }

  // characters.json ====================================================================
  var content = characters.getBlob().getDataAsString();
  var json = JSON.parse(content);

  var ids = Object.keys(json.Data);
  var count = ids.length;

  var row = 0;

  powermul_table = [];

  for (var r = 0; r < count; r++) {
    var id = ids[r];

    // MANAGE SKIPPED LINES -------------------------------------------------
    if (IsInvalidHero(id, true)) continue;

    var d = json.Data[id];

    if (!d.hasOwnProperty('special')) continue;

    var traits = d.traits;
    var minion = false;

    if (traits == null) continue;

    // Safety net: skip some NPC traits (should be skipped by IsInvalidHero already)
    var skip = false;
    for (var t = 0; t < traits.length; t++) {
      if (traits[t] == 'Summon' || traits[t] == 'Operator' || traits[t] == 'DoomBot') {
        skip = true;
        break;
      }
      if (traits[t] == 'Minion') {
        minion = true;
      }
    }
    if (skip) continue;
    // -----------------------------------------------------------------------

    // Create the data line
    var herodata = [];

    var n = column['affiliation1'];
    heroIndex[id] = row;

    herodata[column['HERO_ID']] = id;
    herodata[column['player_Character']] = false;

    if (minion) herodata[n++] = 'HERO_TRAIT_MINION';

    for (var t = 0; t < traits.length; t++) {
      var trait = typeof traits[t] == 'string' ? 'HERO_TRAIT_' + traits[t].toUpperCase() : null;

      switch (GetTraitType(traits[t])) {
        case 'Alignment':
          herodata[column['trait_alignment']] = trait;
          break;
        case 'Origin':
          herodata[column['trait_origin']] = trait;
          break;
        case 'Location':
          herodata[column['trait_area']] = trait;
          break;
        case 'CombatRole':
          herodata[column['trait_role']] = trait;
          break;
        case 'Affiliation':
          herodata[n++] = trait;
      }
    }

    herodata[column['power_mul_index']] = GetPowerMulIndex(d, minion);

    newData[row] = herodata;
    row++;
  }

  // M3HeroStats.json ====================================================================
  content = M3HeroStats.getBlob().getDataAsString();
  json = JSON.parse(content);

  ids = Object.keys(json);
  count = ids.length;

  for (var i = 0; i < count; i++) {
    var s = 100;
    var id = ids[i];

    // Skip the heroes not in the previous list
    if (!heroIndex.hasOwnProperty(id)) continue;
    var d = json[id];
    var index = heroIndex[id];

    if (d.hasOwnProperty('speed_Override')) s = d.speed_Override[1];

    newData[index][column['speed']] = s;
    newData[index][column['Health_mod_index']] = GetStatModIndex(d.health_Mod) + 1;
    newData[index][column['Damage_mod_index']] = GetStatModIndex(d.damage_Mod) + 1;
    newData[index][column['Armor_mod_index']] = GetStatModIndex(d.armor_Mod) + 1;
    newData[index][column['Focus_mod_index']] = GetStatModIndex(d.focus_Mod) + 1;
    newData[index][column['Resist_mod_index']] = GetStatModIndex(d.resist_Mod) + 1;
  }

  // M3HeroSheet.json ====================================================================
  content = M3HeroSheet.getBlob().getDataAsString();
  json = JSON.parse(content);

  ids = Object.keys(json);
  count = ids.length;

  for (var i = 0; i < count; i++) {
    var id = ids[i];

    // Skip the heroes not in the previous list
    if (!heroIndex.hasOwnProperty(id)) continue;
    var d = json[id];
    var index = heroIndex[id];

    newData[index][column['player_Character']] = d.player_Character;
    newData[index][column['unlock_StarLevel']] = d.unlock_StarLevel;
  }

  // M3GearTier ==========================================================================
  content = M3GearTiers.getBlob().getDataAsString();
  var input = Utilities.parseCsv(content, ','.charCodeAt(0));

  count = input.length;

  for (var i = 1; i < count; i++) {
    var id = input[i][0];

    // Skip the heroes not in the previous list
    if (!heroIndex.hasOwnProperty(id)) continue;
    var index = heroIndex[id];

    newData[index][column['base health']] = input[i][14];
    newData[index][column['base damage']] = input[i][15];
    newData[index][column['base armor']] = input[i][16];
    newData[index][column['base focus']] = input[i][17];
    newData[index][column['base resist']] = input[i][18];
  }

  // Power mul table ========================
  if (!UpdateRangeData(GetSheet('_HeroData_PowerMul'), 1, 1, powermul_table[0].length, powermul_table)) return false;

  // Compare to the current data and update accordingle
  if (!UpdateSortedRangeData(heroDataSheet, 1, 1, newData[0].length, 1, newData)) return false;

  // Stats mod table =========================
  var heroStats = [];
  heroStats[0] = [];
  var len = 0;

  // Header: levels
  statmod_levels[1] = true;
  for (var c = 0; c < statmod_levels.length; c++) {
    if (statmod_levels[c] == true) heroStats[0][len++] = c;
  }

  for (var r = 0; r < statmod_table.length; r++) {
    heroStats[r + 1] = [];
    var value = 0.3;
    var index = 0;
    var stats = statmod_table[r];
    var statsk = Object.keys(stats);
    var len = statsk.length;

    for (var c = 0; c < heroStats[0].length; c++) {
      if (index < statsk.length && statsk[index] == heroStats[0][c]) value = stats[statsk[index++]];

      heroStats[r + 1][c] = value;
    }
  }

  if (!UpdateRangeData(GetSheet('_M3HeroStats'), 1, 1, heroStats[0].length, heroStats)) return false;

  return true;
}

function DS_Update_HeroDataFormula() {
  const heroDataSheet = GetSheet('_HeroData');
  const height = heroDataSheet.getMaxRows();

  const rangeNameFormula = heroDataSheet.getRange(1, heroDataSheet.getMaxColumns(), height, 1).getFormulas();
  const ranges = GetEmptyRows(rangeNameFormula);

  for (var i = 0; i < ranges.length; i++) {
    // Formulas =============================
    var formulas = [];

    for (var index = ranges[i][0]; index < ranges[i][1]; index++) {
      const rindex = '$Y' + index; // Roster index column name
      const lindex = '$Z' + index; // Level column name
      const gindex = '$AC' + index; // Gear column name
      const hsindex = '$AD' + index; // HeroStats2 LevelIndex
      const gtindex = '$AE' + index; // GearTier Hero Index
      var forms = [];

      // Hero Name - Only used for color...
      forms.push(
        '=IFERROR(INDEX(_M3Localization_Heroes_Name,MATCH($A' +
          index +
          ',_M3Localization_Heroes_Id,0)),$A' +
          index +
          ')'
      );

      // Roster Index
      forms.push('=IFERROR(MATCH($A' + index + ',Roster_Id,0))');

      // Level
      forms.push('=IF(OR(ISBLANK(' + rindex + '),ISBLANK(' + gindex + ')),,INDEX(Roster_Level,' + rindex + '))');

      // Stars Factor
      forms.push(
        '=IF(ISBLANK(' +
          lindex +
          '),,INDEX(_M3GlobalData_StarFactor,IF(INDEX(Roster_Stars,' +
          rindex +
          ')+0=0,$B' +
          index +
          ',INDEX(Roster_Stars,' +
          rindex +
          ')+1)))'
      );

      // Red Stars Factor
      forms.push(
        '=IF(OR(ISBLANK(' +
          lindex +
          '),INDEX(Roster_Stars,' +
          rindex +
          ')=0,INDEX(Roster_RedStars,' +
          rindex +
          ')=0),1,INDEX(_M3GlobalData_RedStarFactor,1+MIN(INDEX(Roster_Stars,' +
          rindex +
          '),INDEX(Roster_RedStars,' +
          rindex +
          '))))'
      );

      // Gear
      forms.push('=IF(ISBLANK(' + rindex + '),,INDEX(Roster_Gear,' + rindex + '))');

      // HeroStats2 LevelIndex
      forms.push('=IF(ISBLANK(' + lindex + '),,MATCH(' + lindex + ',_M3HeroStats!$1:$1)-1)');

      // GearTiers Hero index
      forms.push('=IF(ISBLANK(' + lindex + '),,MATCH($A' + index + ',_GearTiers_HeroId,0))');

      // GearTiers Hero gear index
      forms.push('=IF(ISBLANK(' + lindex + '),,' + gtindex + '+' + gindex + '-INDEX(_GearTiers_Tier,' + gtindex + '))');

      // GearTiers - Gear Index
      const statsTiersIndex = [1, , 2, 3, 4, 5];
      for (var part = 0; part < 6; part++) {
        if (part == 1)
          forms.push('=IF(OR(ISBLANK(' + lindex + '),INDEX(Roster_GearParts,' + rindex + ',2)<>TRUE),,' + gindex + ')');
        else
          forms.push(
            '=IF(OR(ISBLANK(' +
              lindex +
              '),INDEX(Roster_GearParts,' +
              rindex +
              ',' +
              (part + 1) +
              ')<>TRUE),,INDEX(_GearTiers_StatsTiers,$AF' +
              index +
              ',' +
              statsTiersIndex[part] +
              ')+1)'
          );
      }

      // Gear stats
      var statsMods = [];
      const statscols = ['S', 'T', 'U', 'V', 'W']; // Mod index
      for (var c = 0; c < statscols.length; c++) {
        const statCell = statscols[c] + index;
        statsMods.push('OFFSET(_M3HeroStats!$A$1,' + statCell + ',' + hsindex + ')');
      }

      const baseStat = ['M', 'N', 'O', 'P', 'Q'];
      const gearCol = ['$AG' + index, '$AH' + index, '$AI' + index, '$AJ' + index, '$AK' + index, '$AL' + index];
      var gearStats = [];
      const gearLibStat = [
        '_GearLibrary_Stats_Focus',
        '_GearLibrary_Hero_Stats',
        '_GearLibrary_Stats_Damage',
        '_GearLibrary_Stats_Resist',
        '_GearLibrary_Stats_Armor',
        '_GearLibrary_Stats_Health'
      ];
      for (var c = 0; c < baseStat.length; c++) {
        var f = '(INDEX(_GearTiers_BaseStats,$AF' + index + ',' + (c + 1) + ')';

        for (var g = 0; g < gearCol.length; g++) {
          f += ' + IF(ISBLANK(' + gearCol[g] + '),,INDEX(' + gearLibStat[g] + ',' + gearCol[g] + ',' + (c + 1) + '))';
        }

        f += ')';

        gearStats.push(f);
      }

      // Stark Tech %
      var starkTech = [];
      for (var st = 1; st <= 5; st++)
        starkTech.push(
          'IF(' +
            lindex +
            "<20,1,1+INDEX('Stark Tech'!$C$1:$V$15," +
            st +
            '*2+4,MATCH($E' +
            index +
            ",'Stark Tech'!$C$1:$V$1,0)))"
        );

      const curveRange = ['$A:$A', '$A:$A', '$A:$A', '$B:$B', '$B:$B'];
      const statFactor = ['', '(0.06+1/4.25)*', '0.06*', '', ''];

      // Total stats
      for (var s = 0; s < statFactor.length; s++) {
        forms.push(
          '=IF(ISBLANK(' +
            lindex +
            '),,(INDEX(_M3GlobalData_StatsCurve!' +
            curveRange[s] +
            ',' +
            lindex +
            ')*' +
            statFactor[s] +
            statsMods[s] +
            '*$AA' +
            index +
            '+' +
            gearStats[s] +
            ')*$AB' +
            index +
            '*' +
            starkTech[s] +
            ')'
        );
      }

      // Color
      forms.push(
        '=IF(Preferences_Color_Hero_Id="OPTION_TEAM",INDEX(Preferences_Color_Team_ColorId,INDEX(_Export_ColorTeams_Value,MIN(' +
          'IF(ISBLANK(D' +
          index +
          '),30,IFERROR(MATCH(D' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(E' +
          index +
          '),30,IFERROR(MATCH(E' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(F' +
          index +
          '),30,IFERROR(MATCH(F' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(G' +
          index +
          '),30,IFERROR(MATCH(G' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(H' +
          index +
          '),30,IFERROR(MATCH(H' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(I' +
          index +
          '),30,IFERROR(MATCH(I' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(J' +
          index +
          '),30,IFERROR(MATCH(J' +
          index +
          ',_Export_ColorTeams_Id,0),30)),' +
          'IF(ISBLANK(K' +
          index +
          '),30,IFERROR(MATCH(K' +
          index +
          ',_Export_ColorTeams_Id,0),30))))),' +
          'VLOOKUP(INDEX(A' +
          index +
          ':G' +
          index +
          ',0,Preferences_Color_Hero_Index),_ColorCodes,2))'
      );

      formulas.push(forms);
    }

    if (!copyToRangeFormula(heroDataSheet, ranges[i][0], 24, formulas)) return false;
  }

  return true;
}

function GetPowerMulIndex(d, minion) {
  var basic = d.basic.power_mul;
  var special = d.special.power_mul;
  var ultimate = minion ? null : d.ultimate.power_mul;
  var passive = d.passive_power_mul || [1, 1, 1, 1, 1];

  var line = [];

  if (minion)
    line = [
      basic[1],
      basic[2],
      basic[3],
      basic[4],
      basic[5],
      basic[6],
      special[1],
      special[2],
      special[3],
      special[4],
      special[5],
      special[6],
      ,
      ,
      ,
      ,
      ,
      ,
      ,
      passive[0],
      passive[1],
      passive[2],
      passive[3],
      passive[4]
    ];
  else
    line = [
      basic[1],
      basic[2],
      basic[3],
      basic[4],
      basic[5],
      basic[6],
      special[1],
      special[2],
      special[3],
      special[4],
      special[5],
      special[6],
      ultimate[0],
      ultimate[1],
      ultimate[2],
      ultimate[3],
      ultimate[4],
      ultimate[5],
      ultimate[6],
      passive[0],
      passive[1],
      passive[2],
      passive[3],
      passive[4]
    ];

  var i = 0;
  for (i = 0; i < powermul_table.length; i++) {
    var power = powermul_table[i];
    var equals = true;
    for (var j = 0; j < power.length; j++) {
      if (line[j] != power[j]) {
        equals = false;
        break;
      }
    }
    if (equals) return i;
  }

  powermul_table[i] = line;

  return i;
}

function GetStatModIndex(mods) {
  var modsk = Object.keys(mods);
  var i = 0;

  for (i = 0; i < statmod_table.length; i++) {
    var stats = statmod_table[i];
    var statsk = Object.keys(stats);

    if (statsk.length != modsk.length) continue;

    var equals = true;

    for (var j = 0; j < statsk.length; j++) {
      if (statsk[j] != modsk[j] || stats[statsk[j]] != mods[statsk[j]]) {
        equals = false;
        break;
      }
    }
    if (equals) return i;
  }

  for (var j = 0; j < modsk.length; j++) {
    statmod_levels[modsk[j]] = true;
  }
  statmod_table[i] = mods;

  return i;
}
