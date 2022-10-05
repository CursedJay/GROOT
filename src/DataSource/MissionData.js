function Test_DS_Update_MissionsData() {
  const latest = getNamedRangeValue('_Version_DataSource_Latest');
  PropertiesService.getScriptProperties().setProperty('DataSourceUpdateVersion', latest);
  DS_Update_MissionsData();
}

// Will update sheet _Missions_Data
function DS_Update_MissionsData() {
  const updateVersion = PropertiesService.getScriptProperties().getProperty('DataSourceUpdateVersion');
  const dataSourceFolder = DriveApp.getFolderById(_dataSourceFolderId);
  const versionFolder = dataSourceFolder.getFoldersByName(updateVersion).next();
  const missionsFolder = versionFolder.getFoldersByName('m3missions').next();

  const fileNames = ['ChallengesDetails.json', 'EventCampaignDetails.json'];
  // ChallengesDetails covers Flash events, Legendary and Challenges
  // EventCampaignDetails covers character events
  // No campaign

  const data = [];
  for (let f = 0; f < fileNames.length; f++) {
    const challengesFile = missionsFolder.getFilesByName(fileNames[f]).next();
    const content = challengesFile.getBlob().getDataAsString();
    const json = JSON.parse(content);

    const ids = Object.keys(json.GameData);
    const count = ids.length;

    for (let e = 0; e < count; e++) {
      const eventId = ids[e];
      if (fileNames[f] == 'EventCampaignDetails.json' && !eventId.endsWith('_A')) continue;

      const event = json.GameData[eventId];

      if (!event.hasOwnProperty('m_name')) continue;
      if (!event.hasOwnProperty('m_chapterList')) continue;
      if (!event.m_chapterList[0].hasOwnProperty('m_missionList')) continue;

      const missions = event.m_chapterList[0].m_missionList;
      const mission = missions[missions.length - 1];
      if (!mission.hasOwnProperty('m_heroFilters')) continue;
      if (!mission.m_heroFilters.hasOwnProperty('filters')) continue;

      let allTraits = [];
      let anyTrait = [];
      let filterIds = [];

      if (mission.m_heroFilters.hasOwnProperty('anyFilters')) {
        //Mythic Legendary
        if (mission.m_heroFilters.anyFilters.length == 0) continue;
        for (let c = 0; c < mission.m_heroFilters.anyFilters.length; c++) {
          if (mission.m_heroFilters.anyFilters[c].hasOwnProperty('allTraits'))
            allTraits.push(mission.m_heroFilters.anyFilters[c].allTraits);
          if (mission.m_heroFilters.anyFilters[c].hasOwnProperty('anyTrait'))
            anyTrait.push(mission.m_heroFilters.anyFilters[c].anyTrait);
          if (mission.m_heroFilters.anyFilters[c].hasOwnProperty('id'))
            filterIds.push(mission.m_heroFilters.anyFilters[c].id);
        }
        if (allTraits.length == 0 && anyTrait == 0 && filterIds.length == 0) continue;
      } else {
        if (mission.m_heroFilters.filters.hasOwnProperty('allTraits'))
          allTraits = mission.m_heroFilters.filters.allTraits;
        if (mission.m_heroFilters.filters.hasOwnProperty('anyTrait')) anyTrait = mission.m_heroFilters.filters.anyTrait;
        if (mission.m_heroFilters.filters.hasOwnProperty('id')) filterIds = mission.m_heroFilters.filters.id;

        // Only take the event where one or more teams is required
        if (allTraits.length == 0 && anyTrait == 0 && filterIds == 0) continue;

        if (filterIds.length == 0 && mission.hasOwnProperty('m_mustHaveHeroes')) {
          filterIds = mission.m_mustHaveHeroes;
        }
        if (filterIds.length == 0 && event.hasOwnProperty('m_mustHaveHeroes')) {
          filterIds = event.m_mustHaveHeroes;
        }
      }

      let eventName = event.m_name;

      let rewards = [];
      if (mission.hasOwnProperty('m_rewardsIDList')) rewards = mission.m_rewardsIDList;

      let reward = '';
      if (rewards.length > 0) {
        if (rewards[0].hasOwnProperty('m_rewardID')) reward = rewards[0].m_rewardID;
      }

      let eventType;

      if (eventName.startsWith('ID_LE_')) eventType = 'EVENT_LEGENDARY';
      else if (eventName.startsWith('ID_FL_')) eventType = 'EVENT_FLASH_EVENT';
      else if (eventName.startsWith('ID_CHAL_')) eventType = 'GAMEMODE_CHALLENGES';
      else if (eventName.startsWith('ID_TU_')) eventType = 'UI_TEAM_UP';
      else if (eventName.startsWith('ID_ECM_')) {
        eventName = eventId.substring(0, eventId.lastIndexOf('_'));
        eventType = 'EVENT_CHARACTER_EVENT';
      } else continue; // Ignore all other type of event, deprecated

      if (eventType != 'EVENT_CHARACTER_EVENT') {
        eventName = eventName.substring(eventName.indexOf('_') + 1);
        eventName = eventName.substring(0, eventName.lastIndexOf('_'));
      }
      const row = [];

      row.push(eventType);
      row.push(eventName.toUpperCase());
      row.push(''); // This is for campaign missions

      if (reward.startsWith('legendary_')) {
        reward = reward.substring(reward.indexOf('_') + 1);
        reward = reward.substring(0, reward.lastIndexOf('_'));
        row.push(reward);
      } else if (reward.startsWith('m_event_') && reward.indexOf('_hard_') > 0) {
        reward = reward.substring(8);
        reward = reward.substring(0, reward.lastIndexOf('_hard_'));
        row.push(reward.replace('_', ''));
      } else row.push('');

      for (let c = 0; c < 5; c++) {
        if (allTraits.length > c) row.push('HERO_TRAIT_' + allTraits[c].toString().toUpperCase());
        else row.push('');
      }

      for (let c = 0; c < 5; c++) {
        if (anyTrait.length > c) row.push('HERO_TRAIT_' + anyTrait[c].toString().toUpperCase());
        else row.push('');
      }

      for (let c = 0; c < 5; c++) {
        if (filterIds.length > c) row.push(filterIds[c]);
        else row.push('');
      }

      data.push(row);
    }
  }

  // Campaign
  const missionFile = missionsFolder.getFilesByName('CampaignDetails.json').next();
  const content = missionFile.getBlob().getDataAsString();
  const json = JSON.parse(content);

  const ids = Object.keys(json.GameData);
  const count = ids.length;

  for (let e = 0; e < count; e++) {
    const eventId = ids[e];
    const event = json.GameData[eventId];

    if (!event.hasOwnProperty('m_name')) continue;
    if (!event.hasOwnProperty('m_chapterList')) continue;
    if (!event.m_chapterList[0].hasOwnProperty('m_missionList')) continue;

    let eventName = event.m_name;
    if (!eventName.startsWith('ID_GAMEMODE_') || !eventName.endsWith('_NAME')) continue;
    eventName = eventName.substring(3, eventName.lastIndexOf('_'));

    let globalFilter = null;
    if (event.hasOwnProperty('m_heroFilters')) globalFilter = event.m_heroFilters.filters;

    const chaptersList = event.m_chapterList;
    const chaptersFilters = [];
    for (let chapterNum = 0; chapterNum < chaptersList.length; chapterNum++) {
      const chapter = chaptersList[chapterNum];

      const missionList = chapter.m_missionList;

      // Only check the boss nodes, so far missions are always grouped by 3
      const missionsFilters = [];
      for (let missionNum = 2; missionNum < missionList.length; missionNum += 3) {
        const mission = missionList[missionNum];
        if (mission.hasOwnProperty('m_heroFilters'))
          missionsFilters.push(CombineFilters(globalFilter, mission.m_heroFilters.filters));
        else missionsFilters.push(globalFilter);
      }

      chaptersFilters.push(missionsFilters);
    }

    const lines = [];
    for (let chapterNum = 0; chapterNum < chaptersFilters.length; chapterNum++) {
      // All the missions of this chapter share the same filter, let's check if the following ones do as well
      if (SameFilters(chaptersFilters[chapterNum])) {
        let ch = chapterNum + 1;
        while (
          ch < chaptersFilters.length &&
          SameFilters(chaptersFilters[ch]) &&
          EqualFilters(chaptersFilters[chapterNum][0], chaptersFilters[ch][0])
        )
          ch++;

        if (chapterNum == 0 && ch == chaptersFilters.length)
          lines.push({ FirstChapter: -1, LastChapter: -1, Mission: -1, Filter: chaptersFilters[chapterNum][0] });
        else
          lines.push({
            FirstChapter: chapterNum + 1,
            LastChapter: ch,
            Mission: -1,
            Filter: chaptersFilters[chapterNum][0]
          });
        chapterNum = ch - 1;
        continue;
      }

      for (let mission = 0; mission < chaptersFilters[chapterNum].length; mission++) {
        // Ignore if a section has the global campaign filter unless it's on the first chapter, just in case
        if (chapterNum > 0 && !EqualFilters(globalFilter, chaptersFilters[chapterNum][mission]))
          lines.push({
            FirstChapter: chapterNum + 1,
            LastChapter: chapterNum + 1,
            Mission: mission,
            Filter: chaptersFilters[chapterNum][mission]
          });
      }
    }

    for (let l = 0; l < lines.length; l++) {
      const line = lines[l];
      const row = [];
      row.push('GAMEMODE_CAMPAIGNS');
      row.push(eventName);
      if (line.FirstChapter < 0 || (line.FirstChapter == 1 && line.Mission == -1)) {
        row.push('');
      } else if (line.Mission == -1 && line.FirstChapter == line.LastChapter) {
        row.push(line.FirstChapter);
      } else if (line.Mission == -1) {
        row.push(line.FirstChapter + '-' + line.LastChapter);
      } else {
        row.push(line.FirstChapter + String.fromCharCode(65 + line.Mission));
      }
      row.push(''); // Not yet shard rewards considering they aren't in these data. Will see later how we can handle this

      let allTraits = [];
      let anyTrait = [];
      let filterIds = [];
      if (line.Filter != null) {
        if (line.Filter.hasOwnProperty('allTraits')) allTraits = line.Filter.allTraits;
        if (line.Filter.hasOwnProperty('anyTrait')) anyTrait = line.Filter.anyTrait;
        if (line.Filter.hasOwnProperty('id')) filterIds = line.Filter.id;
      }
      for (let c = 0; c < 5; c++) {
        if (allTraits.length > c) row.push('HERO_TRAIT_' + allTraits[c].toUpperCase());
        else row.push('');
      }

      for (let c = 0; c < 5; c++) {
        if (anyTrait.length > c) row.push('HERO_TRAIT_' + anyTrait[c].toUpperCase());
        else row.push('');
      }

      for (let c = 0; c < 5; c++) {
        if (filterIds.length > c) row.push(filterIds[c]);
        else row.push('');
      }
      data.push(row);
    }
  }

  return UpdateSortedRangeData(GetSheet('_MissionsData'), 1, 1, data[0].length, 3, data);
}

function SameFilters(filters) {
  if (filters.length < 2) return true;

  for (let f = 1; f < filters.length; f++) if (!EqualFilters(filters[0], filters[f])) return false;

  return true;
}

function CombineFilters(filter1, filter2) {
  if (filter1 == null) return filter2;
  if (filter2 == null) return filter1;
  if (filter1.length == 0) return filter2;
  if (filter2.length == 0) return filter1;

  const res = {};
  if (filter1.hasOwnProperty('allTraits')) {
    res.allTraits = [];
    for (let t = 0; t < filter1.allTraits.length; t++) res.allTraits.push(filter1.allTraits[t]);
  }
  if (filter1.hasOwnProperty('anyTrait')) {
    res.anyTrait = [];
    for (let t = 0; t < filter1.anyTrait.length; t++) res.anyTrait.push(filter1.anyTrait[t]);
  }
  if (filter1.hasOwnProperty('id')) {
    res.id = [];
    for (let t = 0; t < filter1.id.length; t++) res.id.push(filter1.id[t]);
  }

  if (filter2.hasOwnProperty('allTraits')) {
    if (!res.hasOwnProperty('allTraits')) res.allTraits = [];
    for (let t = 0; t < filter2.allTraits.length; t++) {
      if (!res.allTraits.includes(filter2.allTraits[t])) res.allTraits.push(filter2.allTraits[t]);
    }
  }
  if (filter2.hasOwnProperty('anyTrait')) {
    if (!res.hasOwnProperty('anyTrait')) res.anyTrait = [];
    for (let t = 0; t < filter2.anyTrait.length; t++) {
      if (!res.anyTrait.includes(filter2.anyTrait[t])) res.anyTrait.push(filter2.anyTrait[t]);
    }
  }
  if (filter2.hasOwnProperty('id')) {
    if (!res.hasOwnProperty('id')) res.id = [];
    for (let t = 0; t < filter2.id.length; t++) {
      if (!res.id.includes(filter2.id[t])) res.id.push(filter2.id[t]);
    }
  }

  return res;
}

function EqualFilters(filter1, filter2) {
  if (filter1 == null && filter2 == null) return true;
  if (filter1 == null || filter2 == null) return false;

  if (filter1.hasOwnProperty('allTraits') != filter2.hasOwnProperty('allTraits')) return false;
  if (filter1.hasOwnProperty('anyTrait') != filter2.hasOwnProperty('anyTrait')) return false;
  if (filter1.hasOwnProperty('id') != filter2.hasOwnProperty('id')) return false;

  if (filter1.hasOwnProperty('allTraits')) {
    if (filter1.allTraits.length != filter2.allTraits.length) return false;

    for (let t = 0; t < filter1.allTraits.length; t++) {
      if (!filter2.allTraits.includes(filter1.allTraits[t])) return false;
    }
  }

  if (filter1.hasOwnProperty('anyTrait')) {
    if (filter1.anyTrait.length != filter2.anyTrait.length) return false;

    for (let t = 0; t < filter1.anyTrait.length; t++) {
      if (!filter2.anyTrait.includes(filter1.anyTrait[t])) return false;
    }
  }

  if (filter1.hasOwnProperty('id')) {
    if (filter1.id.length != filter2.id.length) return false;

    for (let t = 0; t < filter1.id.length; t++) {
      if (!filter2.id.includes(filter1.id[t])) return false;
    }
  }

  return true;
}

function DS_Update_MissionsDataFormula() {
  const sheet = GetSheet('_MissionsData');
  const rangeFormulas = sheet.getRange(1, sheet.getMaxColumns(), sheet.getMaxRows(), 1).getFormulas();

  const ranges = GetEmptyRows(rangeFormulas);

  for (let g = 0; g < ranges.length; g++) {
    const newData = [];
    for (let row = ranges[g][0]; row < ranges[g][1]; row++) {
      const newRow = [];

      newRow.push(
        '=IFERROR(INDEX(_M3Localization_Challenges_Name,MATCH(B' +
          row +
          ',_M3Localization_Challenges_Id,0)),B' +
          row +
          ') & " " & C' +
          row
      );
      newRow.push(
        '=IFERROR(IF(ISBLANK(D' +
          row +
          '),,INDEX(_M3Localization_Heroes_Name, MATCH(D' +
          row +
          ',_M3Localization_Heroes_Id,0))))'
      );
      newRow.push('=B' + row + ' & IF(NOT(ISBLANK(C' + row + ')),C' + row + ',)');

      newData.push(newRow);
    }

    if (!copyToRangeFormula(sheet, ranges[g][0], sheet.getMaxColumns() + 1 - newData[0].length, newData)) return false;
  }

  return true;
}
