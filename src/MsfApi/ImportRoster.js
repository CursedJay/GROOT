function api_importRoster(since) {
  const rosterVersionCell = getNamedRange('_Version_Roster');
  const rosterVersionValue = rosterVersionCell.getValue();
  const _since = since ?? (rosterVersionValue ? rosterVersionValue : 'fresh');
  const roster = GrootApi.getRoster(_since);

  if (roster === false) return;

  const fullInv = GrootApi.getFullInventory();
  const rosterInv = {};
  rosterInv.shards = {};
  rosterInv.redstars = {};
  rosterInv.diamonds = {};

  for (const piece of fullInv.data) {
    const { item, quantity } = piece;
    if (item.id.startsWith('SHARD_')) {
      rosterInv.shards[item.characterId] = quantity;
    } else if (item.id.startsWith('RS_')) {
      const redStars = item.id.split('_').pop();
      rosterInv.redstars[item.characterId] = redStars;
    } else if (item.id.startsWith('PD_')) {
      const diamonds = item.id.split('_').pop();
      rosterInv.diamonds[item.characterId] = diamonds;
    }
  }

  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Update_Data');
  setIdConverter('_Option_Class_Id2', '_Option_Class_Name2');
  const rosterNotesRange = getNamedRange('Roster_Notes');
  const formulas = rosterNotesRange.getFormulas();

  const characters = roster.data;

  for (const character of characters) {
    const characterId = character?.id;

    if (characterId === 'undefined') continue;

    let s = -1;
    for (let r = 0; r < rosterIds.length; r++) {
      if (rosterIds[r][0] === characterId) {
        s = r;
        break;
      } else if (s === -1 && rosterIds[r][0] === '') s = r;
    }
    rosterIds[s][0] = characterId;
    rosterIds[s][1] = character?.favorite ?? false;

    rosterData[s][0] = character.activeYellow;
    //rosterData[s][1] = character.activeRed;
    rosterData[s][1] = rosterInv.redstars?.[characterId] ?? 0;
    rosterData[s][2] = rosterInv.diamonds?.[characterId] ?? '';
    rosterData[s][3] = character.level;
    rosterData[s][4] = character.gearTier;

    for (let gearPart = 1; gearPart <= 6; gearPart++) {
      const equipped = character.gearSlots[gearPart - 1]; //true or false
      rosterData[s][5 + (gearPart - 1)] = equipped;
    }

    //const noIso = Object.keys(character.iso8).length === 0; //no iso equipped returns an empty iso object
    const activeIso8 = character.iso8?.active;
    const matrixLevel = character.iso8?.[activeIso8];
    const isoColors = ['GREEN', 'BLUE', 'PURPLE'];
    const isoMatrixColor = () => {
      const level = matrixLevel ?? 0;
      if (level === 0) return '';
      return isoColors[Math.ceil(level / 5) - 1];
    };

    rosterData[s][11] = activeIso8 === undefined ? '' : valueOf(activeIso8);
    rosterData[s][12] = isoMatrixColor();
    rosterData[s][13] =
      matrixLevel === undefined
        ? ''
        : matrixLevel <= 5
          ? matrixLevel
          : matrixLevel <= 10
            ? matrixLevel - 5
            : matrixLevel - 10; //level 1-5 Green, 6-10 Blue, 11-15 purple
    rosterData[s][14] = character.iso8?.armor ?? '';
    rosterData[s][15] = character.iso8?.resist ?? '';
    rosterData[s][16] = character.iso8?.health ?? '';
    rosterData[s][17] = character.iso8?.focus ?? '';
    rosterData[s][18] = character.iso8?.damage ?? '';

    rosterData[s][19] = character.basic === 0 ? '' : character.basic;
    rosterData[s][20] = character.special === 0 ? '' : character.special;
    rosterData[s][21] = character.ultimate === 0 ? '' : character.ultimate;
    rosterData[s][22] = character.passive === 0 ? '' : character.passive;

    rosterData[s][23] = character.power;
    rosterData[s][24] = rosterInv.shards?.[characterId] ?? '';
  }

  setNamedRangeValues('Roster_Import_Ids', rosterIds);
  setNamedRangeValues('Roster_Update_Data', rosterData);

  for (let row = 0; row < formulas.length; row++) {
    const formula = formulas[row][0];
    if (formula) {
      rosterNotesRange.getCell(row + 1, 1).setFormula(formula);
    }
  }
  rosterVersionCell.setValue(roster.meta.asOf);
  recordProgress();
  // After updating the roster, re-calculate gear usage to make sure it's accurate based on current inventory and updated roster
  api_CalculateGearUsage();
}

function api_importFullRoster() {
  api_importRoster('fresh');
}

function saveRosterJSON() {
  const roster = GrootApi.getRoster('fresh');
  const dateTime = Utilities.formatDate(new Date(), 'GMT+8', "yyyy-MM-dd'T'HH:mm:ss.SS");
  const filename = `msf_roster_${dateTime}.json`;

  saveReqToFile(roster, filename);
}
