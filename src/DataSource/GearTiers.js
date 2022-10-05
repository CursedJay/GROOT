function Test_Update_GearTiers() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_GearTiers();
}

function GetTierOf(gearId) {
  if (gearId == 'T1_NoArmor') return 0;
  if (gearId.startsWith('Green_')) return 1;
  return Number(gearId.substring(1, gearId.indexOf('_')));
}

// Has to be done after HeroData
function DS_Update_GearTiers() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const heroesFolder = versionFolder.getFoldersByName('heroes').next();
  const file = heroesFolder.getFilesByName('M3GearTiers.csv').next();

  const fileGlobal = heroesFolder.getFilesByName('M3GlobalData.json').next();
  const globalContent = fileGlobal.getBlob().getDataAsString();
  const json = JSON.parse(globalContent);

  let GearTierCap = Number(json.GearTierCap);
  const overrideGTC = getNamedRangeValue('Preferences_OverrideMaxGearTier');
  if (overrideGTC != '' && Number(overrideGTC) > GearTierCap) GearTierCap = Number(overrideGTC);

  const content = file.getBlob().getDataAsString();
  const input = Utilities.parseCsv(content, ','.charCodeAt(0));

  const newData = [];

  let id = '';
  let origin = '';
  for (let row = 1; row < input.length; row++) {
    const sourceStats = input[row - 1]; // Use the base of the tier rather than the max
    const source = input[row];
    if (IsInvalidHero(source[0], true)) continue;

    if (source[1] == 0) continue; // Skip tier 0

    const tier = Number(source[1]);
    if (tier > GearTierCap) continue;

    if (id != source[0]) {
      id = source[0];
      origin = input[row + 12][2];
      origin = origin.substring(origin.indexOf('_') + 1, origin.lastIndexOf('_'));
    }

    const focus = GetTierOf(source[2]);
    const damage = GetTierOf(source[6]);
    const resist = GetTierOf(source[8]);
    const armor = GetTierOf(source[10]);
    const health = GetTierOf(source[12]);
    line = [
      id,
      tier,
      origin,
      focus,
      damage,
      resist,
      armor,
      health,
      sourceStats[14],
      sourceStats[15],
      sourceStats[16],
      sourceStats[17],
      sourceStats[18]
    ];
    newData.push(line);
  }

  return UpdateSortedRangeData(GetSheet('_GearTiers'), 1, 1, newData[0].length, 2, newData);
}

function DS_Update_GearTiersFormula() {
  const sheet = GetSheet('_GearTiers');
  const rangeFormulas = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1).getFormulas();

  const ranges = GetEmptyRows(rangeFormulas);

  for (let g = 0; g < ranges.length; g++) {
    const newData = [];
    for (let row = ranges[g][0]; row < ranges[g][1]; row++) {
      const newRow = [];

      // COST
      newRow[0] =
        `= IF(Q${row}, INDEX(_GearLibrary_Craft_CostFocus,  MATCH(C${row} & D${row}, _GearLibrary_Craft_Origin & _GearLibrary_Craft_Tier, 0)), 0)\n` +
        `+ IF(R${row}, INDEX(_GearLibrary_Hero_Cost,        MATCH(A${row} & B${row}, _GearLibrary_Hero_Id      & _GearLibrary_Hero_Tier,  0)), 0)\n` +
        `+ IF(S${row}, INDEX(_GearLibrary_Craft_CostDamage, MATCH(C${row} & E${row}, _GearLibrary_Craft_Origin & _GearLibrary_Craft_Tier, 0)), 0)\n` +
        `+ IF(T${row}, INDEX(_GearLibrary_Craft_CostResist, MATCH(C${row} & F${row}, _GearLibrary_Craft_Origin & _GearLibrary_Craft_Tier, 0)), 0)\n` +
        `+ IF(U${row}, INDEX(_GearLibrary_Craft_CostArmor,  MATCH(C${row} & G${row}, _GearLibrary_Craft_Origin & _GearLibrary_Craft_Tier, 0)), 0)\n` +
        `+ IF(V${row}, INDEX(_GearLibrary_Craft_CostHealth, MATCH(C${row} & H${row}, _GearLibrary_Craft_Origin & _GearLibrary_Craft_Tier, 0)), 0)`;

      // Farming tab hero index
      newRow[1] = `=IFERROR(MATCH(A${row},Farming_Id,0))`;

      // Current gear tier for that hero
      newRow[2] = `=IF(ISBLANK(O${row}),,IF(OR(INDEX(Farming_Enable,O${row},1)<>TRUE, B${row}<INDEX(Farming_Gear,O${row},1),B${row}>=INDEX(Farming_Gear_Target,O${row},1)),,INDEX(Farming_Gear,O${row},1)+0))`;

      for (let c = 0; c < 6; c++) {
        newRow[3 + c] = `=IF(ISBLANK(P${row}),,IF(P${row}<>B${row},TRUE,NOT(INDEX(Farming_Gear,O${row},${c + 2}))))`;
      }
      newData.push(newRow);
    }

    if (!copyToRangeFormula(sheet, ranges[g][0], sheet.getMaxColumns() + 1 - newData[0].length, newData)) return false;
  }

  return true;
}
