function recordProgress() {
  const profileData = api_importProfile();
  const allianceData = getAlliancePower();

  if (!profileData) {
    toastError("Couldn't get player profile data!");
    return;
  }
  if (!allianceData) toastError("Couldn't get alliance data!");

  const { tcp, stp } = profileData;
  const sheet = GetSheet('_ProgressData');
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const date = new Date();
  const currentDate = Utilities.formatDate(date, tz, 'MM/dd/yyyy');

  const lastRow = sheet.getLastRow();
  const colTCP = getNamedRange('_Progress_TCP')
    .getA1Notation()
    .match(/([A-Z]+)/)[0];
  const colSTP = getNamedRange('_Progress_STP')
    .getA1Notation()
    .match(/([A-Z]+)/)[0];
  const lastTCPEntry = sheet.getRange(`${colTCP}${lastRow}`).getValue();

  //Checks if TCP changed
  //if (tcp <= lastTCPEntry) return;

  const colDate = getNamedRange('_Progress_Date')
    .getA1Notation()
    .match(/([A-Z]+)/)[0];
  const lastDateEntry =
    lastRow > 1 ? Utilities.formatDate(sheet.getRange(`${colDate}${lastRow}`).getValue(), tz, 'MM/dd/yyyy') : '';

  // Check if date is the same
  if (currentDate === lastDateEntry) {
    //Edit last row if date didn't change to limit data entries
    const lastRowRange = sheet.getRange(lastRow, getNamedRange('_Progress_TCP').getColumn(), 1, 2);
    lastRowRange.setValues([[tcp, stp]]);

    // Alliance TCP
    if (allianceData) {
      const lastRowRangeAlliance = sheet.getRange(lastRow, getNamedRange('_Progress_AllianceAvgTCP').getColumn(), 1, 3);
      lastRowRangeAlliance.setValues([
        [allianceData?.tcp || '', allianceData?.stp || '', allianceData?.totalTcp || '']
      ]);
    }
  } else {
    //Add a new row
    const deltaTCPform = `=${colTCP}${lastRow + 1}-${colTCP}${lastRow}`;
    const deltaSTPform = `=${colSTP}${lastRow + 1}-${colSTP}${lastRow}`;
    sheet.appendRow([
      currentDate,
      tcp,
      stp,
      deltaTCPform,
      deltaSTPform,
      allianceData?.tcp || '',
      allianceData?.stp || '',
      allianceData?.totalTcp || ''
    ]);
  }
  setNamedRangeValue('_Version_Last_RecordProgress', date);
  //Fix date format on cells
  //const column = sheet.getRange('A2:A');
  //column.setNumberFormat('MM/dd/yyyy');
}

//Function for the trigger, limited to once per day
function autoRecordProgress() {
  const prevRecord = getNamedRangeValue('_Version_Last_RecordProgress');
  const currentDate = new Date();
  const diffTime = Math.floor(currentDate - prevRecord);
  const diffMinutes = Math.ceil(diffTime / (1000 * 60));
  if (diffMinutes >= 600) {
    recordProgress();
    //toastSuccess('Progress Report updated!');
  }
}

// Add a Trigger for "onOpen" event
function addRecordProgressTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const trigger = allTriggers.find((trigger) => trigger.getHandlerFunction() === 'autoRecordProgress');
  if (!trigger) {
    const trigger = ScriptApp.newTrigger('autoRecordProgress')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
    console.log(`Trigger created: ${trigger.getHandlerFunction()}`);
  }
}

// Removes the trigger
function removeRecordProgressTrigger() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const trigger = allTriggers.find((trigger) => trigger.getHandlerFunction() === 'autoRecordProgress');
  if (!trigger) return;
  ScriptApp.deleteTrigger(trigger);
  console.log(`Trigger removed: ${trigger.getHandlerFunction()}`);
}
