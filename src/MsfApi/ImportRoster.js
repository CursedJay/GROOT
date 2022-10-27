function api_importRoster(since) {
  //const rosterVersionCell = getNamedRangeValue('Roster_Since');
  //const _since = since ?? (rosterVersionCell ? rosterVersionCell : 'fresh');
  const _since = since;
  const roster = GrootApi.getRoster(_since);

  if (roster === false) return;

  const rosterIds = getNamedRangeValues('Roster_Import_Ids');
  const rosterData = getNamedRangeValues('Roster_Import_Data');
  setIdConverter('_Option_Class_Id2', '_Option_Class_Name2');
  const rosterNotesRange = getNamedRange('Roster_Notes');
  const formulas = rosterNotesRange.getFormulas();

  const characters = roster.data;

  for (const character of characters) {
    const characterId = character.id;

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
    rosterData[s][1] = character.activeRed;
    rosterData[s][2] = character.level;
    rosterData[s][3] = character.gearTier;

    for (let gearPart = 1; gearPart <= 6; gearPart++) {
      const equipped = character.gearSlots[gearPart - 1]; //true or false
      rosterData[s][4 + (gearPart - 1)] = equipped;
    }

    //const noIso = Object.keys(character.iso8).length === 0; //no iso equipped returns an empty iso object
    const activeIso8 = character.iso8?.active;
    const matrixLevel = character.iso8?.[activeIso8];

    rosterData[s][10] = activeIso8 === undefined ? '' : valueOf(activeIso8);
    rosterData[s][11] = character.iso8?.matrix?.toUpperCase() ?? (matrixLevel >= 1 ? 'GREEN' : ''); //todo: localisation
    rosterData[s][12] = matrixLevel === undefined ? '' : matrixLevel <= 5 ? matrixLevel : matrixLevel - 6; //level 1-5 Green, 6-10 Blue
    rosterData[s][13] = character.iso8?.armor ?? '';
    rosterData[s][14] = character.iso8?.resist ?? '';
    rosterData[s][15] = character.iso8?.health ?? '';
    rosterData[s][16] = character.iso8?.focus ?? '';
    rosterData[s][17] = character.iso8?.damage ?? '';

    rosterData[s][18] = character.basic === 0 ? '' : character.basic;
    rosterData[s][19] = character.special === 0 ? '' : character.special;
    rosterData[s][20] = character.ultimate === 0 ? '' : character.ultimate;
    rosterData[s][21] = character.passive === 0 ? '' : character.passive;

    rosterData[s][22] = character.power;
    //rosterData[s][23] = character.shards; //in inventory
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

function api_importFullRoster() {
  api_importRoster('fresh');
}

function saveRosterJSON() {
  const roster = GrootApi.getRoster('fresh');
  const dateTime = Utilities.formatDate(new Date(), 'GMT+8', "yyyy-MM-dd'T'HH:mm:ss.SS");
  const filename = `msf_roster_${dateTime}.json`;

  saveReqToFile(roster, filename);
}
