function Test_DS_Update_GearLibrary() {
  const version = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', version);
  DS_Update_GearLibrary();
}

let _gearLibraryTable;
let _gearLibraryIndex;

function GetCostOf(material) {
  const index = _gearLibraryIndex[material];

  if (index == null) {
    output(`error with ${material}`);
    return -1;
  }

  let cost = Number(_gearLibraryTable[index][21]);

  const parts = [_gearLibraryTable[index][15], _gearLibraryTable[index][17], _gearLibraryTable[index][19]];
  const isEmpty = (str) => str === '';

  if (parts.every(isEmpty)) return 0;

  parts.forEach((part) => {
    if (!isEmpty(part)) cost += GetCostOf(part);
  });

  return cost;
}

function GetAllMaterial(material, count, onlyBasicCatalyst) {
  if (material === '') return null;
  const index = _gearLibraryIndex[material];
  if (index == null) {
    output(`No index for ${material}`);
    return;
  }
  const matParts = [_gearLibraryTable[index][15], _gearLibraryTable[index][17], _gearLibraryTable[index][19]];
  const qteParts = [
    Number(_gearLibraryTable[index][16]),
    Number(_gearLibraryTable[index][18]),
    Number(_gearLibraryTable[index][20])
  ];

  const parts = [];

  matParts.forEach((part, index) => {
    const mats = GetAllMaterial(part, qteParts[index] * count);
    if (mats != null) parts.push(mats);
  });

  const res = {};
  res.BasicCatalyst = 0;
  res.StatCatalyst = 0;
  res.Mats = [];

  if (parts.length === 0) {
    if (material.split('_')[1] === 'Basic') {
      res.BasicCatalyst = count;
    } else if (!onlyBasicCatalyst) {
      if (_statsTypeShort.includes(material.split('_')[1])) {
        res.StatCatalyst = count;
      } else {
        res.Mats[0] = {};
        res.Mats[0].Name = material;
        res.Mats[0].Count = count;
      }
    }
    return res;
  }

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    res.BasicCatalyst += part.BasicCatalyst;

    if (!onlyBasicCatalyst) {
      res.StatCatalyst += part.StatCatalyst;

      for (let m = 0; m < part.Mats.length; m++) {
        let found = false;
        for (let n = 0; n < res.Mats.length; n++) {
          if (part.Mats[m].Name === res.Mats[n].Name) {
            found = true;
            res.Mats[n].Count += part.Mats[m].Count;
            break;
          }
        }
        if (!found) {
          res.Mats.push(part.Mats[m]);
        }
      }
    }
  }

  return res;
}

function GetAllMaterialHero(material, count) {
  if (material === '') return null;
  const index = _gearLibraryIndex[material];

  const matParts = [_gearLibraryTable[index][15], _gearLibraryTable[index][17], _gearLibraryTable[index][19]];
  const qteParts = [
    Number(_gearLibraryTable[index][16]),
    Number(_gearLibraryTable[index][18]),
    Number(_gearLibraryTable[index][20])
  ];

  const parts = [];

  matParts.forEach((part, index) => {
    const mats = GetAllMaterialHero(part, qteParts[index] * count);
    if (mats != null) parts.push(mats);
  });

  const res = {};
  res.BasicCatalyst = 0;
  res.UniqueMat = null;
  res.Mats = [];
  res.CatalystMats = [];

  if (parts.length === 0) {
    const matSplit = material.split('_');

    if (matSplit[1] === 'Basic') {
      res.BasicCatalyst = count;
    } else if (_statsTypeShort.includes(matSplit[1])) {
      res.CatalystMats[0] = {};
      res.CatalystMats[0].Name = material;
      res.CatalystMats[0].Count = count;
    } else if (matSplit.length == 2) {
      res.UniqueMat = {};
      res.UniqueMat.Name = material;
      res.UniqueMat.Count = count;
    } else {
      res.Mats[0] = {};
      res.Mats[0].Name = material;
      res.Mats[0].Count = count;
    }
    return res;
  }

  for (let i = 0; i < parts.length; i++) {
    const part = parts[i];
    res.BasicCatalyst += part.BasicCatalyst;

    if (part.UniqueMat != null) {
      if (res.UniqueMat != null) res.UniqueMat.Name = 'Error: Different unique is not managed';
      else {
        res.UniqueMat = {};
        res.UniqueMat.Name = part.UniqueMat.Name;
        res.UniqueMat.Count = part.UniqueMat.Count;
      }
    }

    for (let m = 0; m < part.Mats.length; m++) {
      let found = false;
      for (let n = 0; n < res.Mats.length; n++) {
        if (part.Mats[m].Name === res.Mats[n].Name) {
          found = true;
          res.Mats[n].Count += part.Mats[m].Count;
          break;
        }
      }
      if (!found) {
        res.Mats.push(part.Mats[m]);
      }
    }

    for (let m = 0; m < part.CatalystMats.length; m++) {
      let found = false;
      for (let n = 0; n < res.CatalystMats.length; n++) {
        if (part.CatalystMats[m].Name === res.CatalystMats[n].Name) {
          found = true;
          res.CatalystMats[n].Count += part.CatalystMats[m].Count;
          break;
        }
      }
      if (!found) {
        res.CatalystMats.push(part.CatalystMats[m]);
      }
    }
  }

  return res;
}

// Has to be done after HeroData
// Will update _M3GearLibrary
function DS_Update_GearLibrary() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const heroesFolder = versionFolder.getFoldersByName('heroes').next();
  const file = heroesFolder.getFilesByName('M3GearLibrary.csv').next();
  const content = file.getBlob().getDataAsString();
  const extraMaterial = 6;
  const fileGlobal = heroesFolder.getFilesByName('M3GlobalData.json').next();
  const globalContent = fileGlobal.getBlob().getDataAsString();
  const json = JSON.parse(globalContent);

  let GearTierCap = Number(json.GearTierCap);
  const overrideGTC = getNamedRangeValue('Preferences_OverrideMaxGearTier');
  if (overrideGTC != '' && Number(overrideGTC) > GearTierCap) GearTierCap = Number(overrideGTC);

  _gearLibraryTable = Utilities.parseCsv(content, ','.charCodeAt(0));

  // DATA:
  // STAT, TIER, HEALTH, DAMAGE, ARMOR, FOCUS, RESIST, BASIC CATALYST COUNT, TYPE CATALYST, MAT1 NAME, MAT1 COUNT, MAT2 NAME, MAT2 COUNT, MAT3 NAME, MAT3 COUNT, [MAT4 NAME, MAT4 COUNT], BIO COST, SKILL COST, TECH COST, MYSTIC COST, MUTANT COST

  // FORMULA
  // BIO COUNT, SKILL COUNT, TECH COUNT, MYSTIC COUNT, MUTANT COUNT

  const dataStats = [];
  const dataCraft = [];
  const dataHero = [];
  const dataPicture = [];

  const statsPos = {
    Health: 0,
    Damage: 5,
    Armor: 10,
    Focus: 15,
    Resist: 20
  };

  for (let t = 0; t <= GearTierCap; t++) {
    dataStats[t] = [];
    for (let n = 0; n < 25; n++) {
      dataStats[t][n] = 0;
    }
  }

  // Easy access to each source row
  _gearLibraryIndex = new Object();
  for (let row = 1; row < _gearLibraryTable.length; row++) {
    _gearLibraryIndex[_gearLibraryTable[row][0]] = row;
  }

  for (let o = 0; o < _originType.length; o++) {
    const cr = [_originType[o], '1']; // Each origin at tier 1...
    cr.push(1); // Requires one catalyst of the gear part's stat type
    for (let s = 0; s < _statsTypeShort.length; s++) cr.push(0); // No basic catalyst no matter what
    for (let m = 0; m < extraMaterial; m++) cr.push('', ''); // No other material
    for (let s = 0; s < _statsTypeShort.length; s++) cr.push('0'); // And has no cost
    dataCraft.push(cr);
  }

  // Parse source doc
  const gearT1Ids = ['Green_Health_Mat', 'Green_Damage_Mat', 'Green_Armor_Mat', 'Green_Focus_Mat', 'Green_Resist_Mat'];

  for (let row = 1; row < _gearLibraryTable.length; row++) {
    const gearRow = _gearLibraryTable[row];
    const id = gearRow[0];

    // These are the real T1
    if (gearT1Ids.includes(id)) {
      const idPart = id.split('_');

      const pos = statsPos[idPart[1]];
      for (let g = 0; g < _statsTypeShort.length; g++) {
        dataStats[1][pos + g] = gearRow[g + 2];
      }

      //var st = [idPart[1], '1'];
      //for (var g = 0; g < _statsTypeShort.length; g++) st.push(gearRow[g + 2]);
      //dataStats.push(st);

      dataPicture.push([gearRow[0], `${gearRow[14]}.png`]);

      continue;
    }

    // Ignore any other receipe or material
    if (!id.startsWith('T') || id.startsWith('Teal')) {
      // Only keep end material, we don't need anything else
      if (gearRow[15] != '') continue;

      // As long as the red tiers are locked, no need to take this
      if (id.startsWith('Red_') && GearTierCap < 19) continue;
      if (id.startsWith('White_') && GearTierCap < 23) continue;

      dataPicture.push([gearRow[0], `${gearRow[14]}.png`]);
      continue;
    }

    // This is a special case, need to manage it differently (tier 0 in GROOT's sheets)
    if (id == 'T1_NoArmor') {
      // Poly Fiber
      // Tier 0 in this table to simplify the formulas
      //var st = ['Armor', '0']; // Armor only
      //for (var g = 0; g < _statsTypeShort.length; g++) st.push(gearRow[g + 2]);
      //dataStats.push(st);
      const pos = statsPos['Armor'];
      for (let g = 0; g < _statsTypeShort.length; g++) {
        dataStats[0][pos + g] = gearRow[g + 2];
      }

      for (let o = 0; o < _originType.length; o++) {
        const cr = [_originType[o], '0']; // Each origin at tier 0...
        cr.push(0); // Requires no catalyst
        for (let s = 0; s < _statsTypeShort.length; s++) cr.push(0); // No basic catalyst no matter what
        cr.push(id, '1'); // The only material is T1_NoArmor
        for (let m = 1; m < extraMaterial; m++) cr.push('', ''); // No other material
        for (let s = 0; s < _statsTypeShort.length; s++) cr.push('0'); // And has no cost
        dataCraft.push(cr);
      }
      dataPicture.push([gearRow[0], `${gearRow[14]}.png`]);
      continue;
    }

    // T1 are special case. They are all placeholders in this file
    //const stats = ['_Focus', '_Damage', '_Resist', '_Armor', '_Health'];
    if (
      id.startsWith('T1_') &&
      _originType.some((elem) => id.includes(`_${elem}_`)) &&
      _statsTypeShort.some((elem) => id.endsWith(`_${elem}`))
    )
      continue;

    const parts = id.split('_');
    const tier = Number(parts[0].substring(1));
    if (tier > GearTierCap) continue;

    const origin = parts[1];
    const isHeroGear = !_originType.includes(origin);

    // Hero gear is managed in another sheet
    if (isHeroGear) {
      const heroId = id.substring(parts[0].length + 1);
      if (IsInvalidHero(heroId, true)) continue;

      const allMat = GetAllMaterialHero(id, 1);

      // Hero Gear
      const rowHero = [heroId, tier];
      for (let g = 0; g < _statsTypeShort.length; g++) rowHero.push(gearRow[g + 2]);

      if (allMat.CatalystMats.length + (allMat.UniqueMat == null ? 0 : 1) > 2 || allMat.CatalystMats.length > 2) {
        error(`Too many material for ${id}`);
      }

      rowHero.push(allMat.BasicCatalyst);
      if (allMat.CatalystMats.length > 0) rowHero.push(allMat.CatalystMats[0].Name, allMat.CatalystMats[0].Count);
      else rowHero.push('', '');

      if (allMat.CatalystMats.length > 1) rowHero.push(allMat.CatalystMats[1].Name, allMat.CatalystMats[1].Count);
      else if (allMat.UniqueMat != null) rowHero.push(allMat.UniqueMat.Name, allMat.UniqueMat.Count);
      else rowHero.push('', '');

      for (let i = 0; i < 3; i++) {
        if (i < allMat.Mats.length) rowHero.push(allMat.Mats[i].Name, allMat.Mats[i].Count);
        else rowHero.push('', '');
      }

      rowHero.push(GetCostOf(id));
      dataHero.push(rowHero);
      continue;
    } else {
      // Crafted non hero gear image library T12+
      if (id.startsWith('T')) {
        if (tier >= 12) {
          dataPicture.push([gearRow[0], `${gearRow[14]}.png`]);
        }
      }

      const stat = parts[2];

      // Stats are rather straight forward
      if (origin === _originType[0]) {
        //var rowStats = [stat, tier];
        //for (var g = 0; g < _statsTypeShort.length; g++) rowStats.push(gearRow[g + 2]);
        //dataStats.push(rowStats);

        const pos = statsPos[stat];
        for (let g = 0; g < _statsTypeShort.length; g++) {
          dataStats[Number(tier)][pos + g] = gearRow[g + 2];
        }
      }
      // We just need one, getting the data we need for all stat type at once
      if (stat === _statsTypeShort[0]) {
        const allMat = GetAllMaterial(id, 1, false);
        const rowCraft = [origin, tier, allMat.StatCatalyst, allMat.BasicCatalyst];
        for (let s = 1; s < _statsTypeShort.length; s++) {
          const statMat = GetAllMaterial(`T${tier}_${origin}_${_statsTypeShort[s]}`, 1, true);
          rowCraft.push(statMat.BasicCatalyst);
        }

        if (allMat.Mats.length > extraMaterial) {
          error(`Too many material for ${id}`);
        }

        for (let i = 0; i < extraMaterial; i++) {
          if (i < allMat.Mats.length) {
            rowCraft.push(allMat.Mats[i].Name, allMat.Mats[i].Count);
          } else {
            rowCraft.push('', '');
          }
        }

        for (let s = 0; s < _statsTypeShort.length; s++)
          rowCraft.push(GetCostOf(`T${tier}_${origin}_${_statsTypeShort[s]}`));

        dataCraft.push(rowCraft);
      }
    }
  }

  if (!UpdateSortedRangeData(GetSheet('_GearLibrary_Craft'), 1, 1, dataCraft[0].length, 3, dataCraft)) return false;
  //if (!UpdateSortedRangeData(GetSheet("_GearLibrary_Stats"), 1, 1, dataStats[0].length, 3, dataStats)) return false;
  if (!UpdateRangeData(GetSheet('_GearLibrary_Stats'), 1, 1, dataStats[0].length, dataStats)) return false;
  if (!UpdateSortedRangeData(GetSheet('_GearLibrary_Picture'), 1, 1, dataPicture[0].length, 1, dataPicture))
    return false;
  return UpdateSortedRangeData(GetSheet('_GearLibrary_Hero'), 1, 1, dataHero[0].length, 2, dataHero);
}

function DS_Update_GearLibraryCraftFormula() {
  const sheet = GetSheet('_GearLibrary_Craft');
  const rangeFormulas = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1).getFormulas();

  const ranges = GetEmptyRows(rangeFormulas);

  for (let g = 0; g < ranges.length; g++) {
    const newData = [];
    for (let row = ranges[g][0]; row < ranges[g][1]; row++) {
      const newRow = [];

      // Count of gear we need of each type
      const statCraftedGearCol = ['Y:Y', 'Z:Z', 'AA:AA', 'AB:AB', 'AC:AC'];
      for (let s = 0; s < _statsTypeShort.length; s++)
        newRow.push(
          `=MAX(COUNTIFS(_GearTiers_Slot_${_statsTypeShort[s]}, "=" & TRUE, _GearTiers_Origin, "=" & A${row}, _GearTiers_Tier_${_statsTypeShort[s]}, "=" & B${row})-SUMIFS(CraftedGear!${statCraftedGearCol[s]}, CraftedGear!X:X, A${row}, CraftedGear!W:W, B${row}),0)`
        );
      // =max(COUNTIFS(_GearTiers_Slot_Health, "=" & TRUE, _GearTiers_Origin, "=" & A1, _GearTiers_Tier_Health, "=" & B1)
      // -sumifs('*AdvGear'!X:X,'*AdvGear'!$W:$W,A1,'*AdvGear'!$V:$V,B1),0)
      const statCountCol = ['Z', 'AA', 'AB', 'AC', 'AD'];

      // Total of gear we need for this row
      newRow.push(`=SUM(${statCountCol[0]}${row}: ${statCountCol[statCountCol.length - 1]}${row})`);

      // Stat catalysts
      for (let s = 0; s < statCountCol.length; s++) newRow.push(`=${statCountCol[s]}${row} * $C${row}`);

      // Basic catalyst
      newRow.push(
        `= ${statCountCol[0]}${row} * D${row} + ${statCountCol[1]}${row} * E${row} + ${statCountCol[2]}${row} * F${row} + ${statCountCol[3]}${row} * G${row} + ${statCountCol[4]}${row} * H${row}`
      );

      // Total material
      const matCountCol = ['J', 'L', 'N', 'P', 'R', 'T'];
      for (let s = 0; s < matCountCol.length; s++) newRow.push(`=${matCountCol[s]}${row} * $AE${row}`);

      newData.push(newRow);
    }

    if (!copyToRangeFormula(sheet, ranges[g][0], sheet.getMaxColumns() + 1 - newData[0].length, newData)) return false;
  }

  return true;
}

function DS_Update_GearLibraryHeroFormula() {
  const sheet = GetSheet('_GearLibrary_Hero');
  const rangeFormulas = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1).getFormulas();

  const ranges = GetEmptyRows(rangeFormulas);

  for (let g = 0; g < ranges.length; g++) {
    const newData = [];
    for (let row = ranges[g][0]; row < ranges[g][1]; row++) {
      const newRow = [];

      newRow.push(`=INDEX(_GearTiers_Slot_Hero, MATCH(A${row} & B${row}, _GearTiers_HeroId & _GearTiers_Tier, 0))`);

      newData.push(newRow);
    }

    if (!copyToRangeFormula(sheet, ranges[g][0], sheet.getMaxColumns() + 1 - newData[0].length, newData)) return false;
  }

  return true;
}
