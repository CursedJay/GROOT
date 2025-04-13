function InfoSourceUpdate() {
  SpreadsheetApp.getActiveSpreadsheet().removeMenu('New InfoSource Available!');

  const json = readJSON('1Jx-l1hm8gzzO6MXI-cYTDZG7xDH-VL2C');

  const ids = Object.keys(json.Hero);
  const count = ids.length;

  const flags = [
    'ORB_BASIC',
    'ORB_PREMIUM',
    'ORB_MEGA1',
    'ORB_MEGA2',
    'ORB_MILLESTONE',
    'ORB_ULTIMUS',
    'ORB_BLITZ',
    'STORE_BLITZ',
    'STORE_RAID',
    'ORB_RAID',
    'ORB_SPOTLIGHT',
    'ORB_LEGACY1',
    'ORB_LEGACY2',
    'ORB_LEGACY3',
    'STORE_ARENA',
    'STORE_WAR',
    'STORE_CRUCIBLE',
    'STORE_BATTLEWORLD',
    'STORE_ULTIMUS',
    'ORB_REDSTAR',
    'STORE_REDSTAR'
  ];

  const data = [];
  for (let r = 0; r < count; r++) {
    const id = ids[r];
    const hero = json.Hero[id];
    const line = [];
    line.push(id);

    if (hero.hasOwnProperty('RELEASE')) {
      const date = new Date();
      try {
        date.setTime(parseInt(hero.RELEASE));
        line.push(date);
      } catch (e) {
        line.push('');
      }
    } else line.push('');

    for (let f = 0; f < flags.length; f++) line.push(hero.Flags.includes(flags[f]));

    data.push(line);
  }

  const sheet = GetSheet('_HeroInfo');
  ResizeSheet(sheet, data.length, data[0].length);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);

  const infoSourceLatest = getNamedRangeValue('_Version_InfoSource_Latest');
  setNamedRangeValue('_Version_InfoSource_Current', infoSourceLatest);
}
