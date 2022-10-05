function mantis_import() {
  const folder = getFolder('MSF_Zaratools', false);
  if (folder == null) return false;

  const file = getFile(folder, 'RosterExport.json');
  if (file == null) return false;

  const content = file.getBlob().getDataAsString();
  const heroes = JSON.parse(content);

  const importIds = getNamedRangeValues('Roster_Import_Ids');
  const importData = getNamedRangeValues('Roster_Import_Data');
  const rosterNotesRange = getNamedRange('Roster_Notes');
  const formulas = rosterNotesRange.getFormulas();
  const ids = Object.keys(heroes);

  const classData = getNamedRangeValues('_Option_Class');
  const classOf = {};
  classOf[''] = '';
  for (let i = 0; i < classData.length; i++) {
    classOf[classData[i][0]] = classData[i][1];
  }

  for (let i = 0; i < ids.length; i++) {
    const id = ids[i];
    const hero = heroes[id];

    let s = -1;
    let added = true;
    for (let r = 0; r < importIds.length; r++) {
      if (importIds[r][0] == id) {
        s = r;
        added = false;
        break;
      } else if (s == -1 && importIds[r] == '') s = r;
    }

    if (s == -1) continue;
    if (added) importIds[s] = [id];
    importIds[s][1] = hero.Favorite;

    importData[s][0] = hero.Stars;
    importData[s][1] = hero.RedStars;
    importData[s][2] = hero.Level == 0 ? '' : hero.Level;
    importData[s][3] = hero.GearTier == 0 ? '' : hero.GearTier;
    for (let g = 0; g < 6; g++) importData[s][4 + g] = hero.GearParts[g];

    if (hero.hasOwnProperty('IsoSkillId')) {
      // LATEST
      if (hero.IsoSkillId == '') {
        importData[s][10] = '';
        importData[s][11] = '';
        importData[s][12] = '';
      } else {
        importData[s][10] = classOf[hero.IsoSkillId.toUpperCase()];

        if (hero.IsoMatrixQuality <= 5) {
          importData[s][11] = 'GREEN';
          importData[s][12] = hero.IsoMatrixQuality;
        } else if (hero.IsoMatrixQuality <= 10) {
          importData[s][11] = 'BLUE';
          importData[s][12] = hero.IsoMatrixQuality - 5;
        } else importData[s][12] = hero.IsoMatrixQuality;
      }
      for (let p = 0; p < 5; p++) {
        importData[s][13 + p] = hero.IsoStat[p] == 0 ? '' : hero.IsoStat[p];
      }
    } else if (hero.hasOwnProperty('Iso')) {
      // Legacy
      if (hero.Iso.SkillId == '') {
        importData[s][10] = '';
        importData[s][12] = '';
      } else {
        importData[s][10] = classOf[hero.Iso.SkillId.toUpperCase()];
        importData[s][12] = hero.Iso.MatrixQuality;
      }
      for (let p = 0; p < 5; p++) {
        importData[s][13 + p] = hero.Iso.Stat[p] == 0 ? '' : hero.Iso.Stat[p];
      }
    }

    for (let a = 0; a < 4; a++) importData[s][18 + a] = hero.Abilities[a] == 0 ? '' : hero.Abilities[a];
    importData[s][22] = hero.Power;
    importData[s][23] = hero.Stars == 7 ? '' : hero.Shards;
  }

  setNamedRangeValues('Roster_Import_Ids', importIds);
  setNamedRangeValues('Roster_Import_Data', importData);

  for (let row = 0; row < formulas.length; row++) {
    const formula = formulas[row][0];
    if (formula) {
      rosterNotesRange.getCell(row + 1, 1).setFormula(formula);
    }
  }
}
