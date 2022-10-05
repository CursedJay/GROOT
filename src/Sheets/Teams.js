function TEST_SortNameAsc() {
  SortTeamsByPower('SORT_NAME_ASC');
}

function SortTeamsByPower(sorting) {
  let teamsCondition;
  let isNumber = false;

  const power = getNamedRangeValues('_Teams_Power');

  if (sorting.startsWith('SORT_POWER')) {
    isNumber = true;
    teamsCondition = power;
  } else if (sorting.startsWith('SORT_NAME')) {
    teamsCondition = getNamedRangeValues('_Teams_Names');
  } else return;

  const inColumns = [];
  const outColumns = [];
  const teamsBlitz = getNamedRangeValues('_Export_Teams_Blitz');
  const teamsWar = getNamedRangeValues('_Export_Teams_War');
  const teamInterline = 8;

  for (let c = 0; c < 4; c++) {
    inColumns[c] = getNamedRangeValues(`Teams_Column${c + 1}`);
    outColumns[c] = new Array(inColumns[c].length);
  }

  const valuesTable = [];
  const emptyTable = [];
  for (let i = 0; i < teamsCondition.length; i++) {
    let row;
    if (isNumber) row = [i, Number(teamsCondition[i])];
    else row = [i, teamsCondition[i]];

    if (Number(power[i]) == 0) emptyTable.push(row);
    else valuesTable.push(row);
  }

  for (let i = 0; i < valuesTable.length; i++) {
    output(`Was: ${valuesTable[i][1]}`);
  }

  // ASC Power
  if (sorting.endsWith('_ASC')) {
    valuesTable.sort((row1, row2) => {
      return row1[1] > row2[1] ? 1 : -1;
    });
  } else if (sorting.endsWith('_DESC')) {
    valuesTable.sort((row1, row2) => {
      return row1[1] < row2[1] ? 1 : -1;
    });
  } else return;

  for (let i = 0; i < valuesTable.length; i++) {
    output(`Now: ${valuesTable[i][1]}`);
  }

  for (let i = 0; i < emptyTable.length; i++) {
    valuesTable.push(emptyTable[i]);
  }

  for (let t = 0; t < valuesTable.length; t++) {
    const y = Math.floor(t / 4) * teamInterline;
    const x = 0;
    const c = t % 4;
    const ptt = valuesTable[t][0];

    for (let r = 0; r < 6; r++) {
      outColumns[c][y + r] = [inColumns[ptt % 4][Math.floor(ptt / 4) * 8 + r][0]];
    }
    outColumns[c][y + 6] = [''];
    outColumns[c][y + 7] = [''];
  }

  for (let c = 0; c < 4; c++) {
    setNamedRangeValues(`Teams_Column${c + 1}`, outColumns[c]);
  }

  // Blitz and War Checkboxes
  const checkcol = [[], [], [], []];
  const teamPerColumn = 16;

  // By default, all checkboxes are true
  for (let c = 0; c < 4; c++) {
    for (let r = 0; r < (teamPerColumn - 1) * teamInterline + 1; r++) {
      if (r % teamInterline == 0) checkcol[c][r] = [true, true];
      else checkcol[c][r] = ['', ''];
    }
  }

  for (let t = 0; t < valuesTable.length; t++) {
    const y = Math.floor(t / 4) * teamInterline;
    const x = 0;
    const c = t % 4;
    const ptt = valuesTable[t][0];

    checkcol[c][y] = [teamsBlitz[ptt], teamsWar[ptt]];
  }

  for (let c = 0; c < 4; c++) {
    setNamedRangeValues(`Teams_Flags${c + 1}`, checkcol[c]);
  }
}
