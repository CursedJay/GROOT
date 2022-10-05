function Test_UpdateData() {
  const sheet = GetSheet('_GTTest');

  //var data = [[1, 2, 3, 4, 5], [6, 7, 8, 9, 10], [11, 12, 13, 14, 15]];
  const data = [];
  //var characters = ["AimControl_Infect", "AimDmg_Offense", "AimSupport_Heal", "Hela", "HumanTorch", "Hulk", "HydraDmg_AoE", "HydraDmg_Buff"];
  const characters = ['HumanTorch', 'Hulk', 'YoYo'];

  for (let h = 0; h < characters.length; h++) {
    for (let g = 1; g <= 3; g++) {
      data.push([characters[h], g, 'Tech']);
    }
  }

  let setValues = false;

  setValues = true;
  //----------------------------
  // SET VALUES
  if (setValues) {
    ResizeSheet(sheet, data.length, data[0].length);
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    return;
  }
  //------------------------------

  SortTable(data, 2);

  //UpdateRangeData(sheet, 1, 1, 5, data);
  UpdateSortedRangeData(sheet, 1, 1, data[0].length, 2, data);
}
// Change sheet size and update any data that has changed (if any)
function UpdateRangeData(sheet, firstRow, firstColumn, lastColumn, newData) {
  const currentData = sheet
    .getRange(firstRow, firstColumn, sheet.getMaxRows() - firstRow + 1, lastColumn - firstColumn + 1)
    .getValues();
  let rowCopy = -1;
  let cMin;
  let cMax;

  const copyArea = [];

  // Check all lines of the current sheet
  for (let r = 0; r < currentData.length; r++) {
    // Less data on the new table, let's skip
    if (r >= newData.length) break;

    let first = -1;
    for (let c = 0; c < currentData[r].length; c++) {
      if (currentData[r][c] == '' && newData[r][c] == null) continue;
      if (currentData[r][c] != newData[r][c]) {
        first = c;
        break;
      }
    }

    // Nothing on this line
    if (first < 0) {
      if (rowCopy >= 0) {
        // Backup the area
        copyArea.push([rowCopy, cMin, r, cMax]);
        rowCopy = -1;
      }
      continue;
    }

    // There's changes on this line, check the right limit
    let last;
    for (last = currentData[r].length - 1; last > first; last--) {
      if (currentData[r][last] == '' && newData[r][last] == null) continue;
      if (currentData[r][last] != newData[r][last]) break;
    }

    last++;

    if (rowCopy < 0) {
      rowCopy = r;
      cMin = first;
      cMax = last;
    } else {
      cMin = Math.min(cMin, first);
      cMax = Math.max(cMax, last);
    }
  }

  // Push last area
  if (rowCopy >= 0) {
    copyArea.push([rowCopy, cMin, currentData.length, cMax]);
  }

  // Add or remove rows if necessary
  if (newData.length > currentData.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), newData.length - currentData.length);
  } else if (newData.length < currentData.length) {
    const count = currentData.length - newData.length;
    sheet.deleteRows(sheet.getMaxRows() - count, count);
    copyArea[1][2] = newData.length;
  }

  // Patch data
  for (let i = 0; i < copyArea.length; i++) {
    const area = copyArea[i];
    const patch = [];
    for (let r = area[0]; r < area[2]; r++) {
      const row = [];
      for (let c = area[1]; c < area[3]; c++) {
        row.push(newData[r][c]);
      }
      patch.push(row);
    }

    if (!copyToRange(sheet, area[0] + firstRow, area[1] + firstColumn, patch)) return false;
  }

  return true;
}

// Change sheet size and update any formula that has changed (if any)
function UpdateRangeFormula(sheet, firstRow, firstColumn, lastColumn, newData) {
  const currentData = sheet
    .getRange(firstRow, firstColumn, sheet.getMaxRows() - firstRow + 1, lastColumn - firstColumn + 1)
    .getFormulas();
  let rowCopy = -1;
  let cMin;
  let cMax;

  const copyArea = [];

  // Check all lines of the current sheet
  for (let r = 0; r < currentData.length; r++) {
    // Less data on the new table, let's skip
    if (r >= newData.length) break;

    let first = -1;
    for (let c = 0; c < currentData[r].length; c++) {
      if (currentData[r][c] == '' && newData[r][c] == null) continue;
      if (currentData[r][c] != newData[r][c]) {
        first = c;
        break;
      }
    }

    // Nothing on this line
    if (first < 0) {
      if (rowCopy >= 0) {
        // Backup the area
        copyArea.push([rowCopy, cMin, r, cMax]);
        rowCopy = -1;
      }
      continue;
    }

    // There's changes on this line, check the right limit
    let last;
    for (last = currentData[r].length - 1; last > first; last--) {
      if (currentData[r][last] == '' && newData[r][last] == null) continue;
      if (currentData[r][last] != newData[r][last]) break;
    }

    last++;

    if (rowCopy < 0) {
      rowCopy = r;
      cMin = first;
      cMax = last;
    } else {
      cMin = Math.min(cMin, first);
      cMax = Math.max(cMax, last);
    }
  }

  // Push last area
  if (rowCopy >= 0) {
    copyArea.push([rowCopy, cMin, currentData.length, cMax]);
  }

  // Add or remove rows if necessary
  if (newData.length > currentData.length) {
    sheet.insertRowsAfter(sheet.getMaxRows(), newData.length - currentData.length);
  } else if (newData.length < currentData.length) {
    const count = currentData.length - newData.length;
    sheet.deleteRows(sheet.getMaxRows() - count, count);
  }

  // Patch data
  for (let i = 0; i < copyArea.length; i++) {
    const area = copyArea[i];
    const patch = [];
    for (let r = area[0]; r < area[2]; r++) {
      const row = [];
      for (let c = area[1]; c < area[3]; c++) {
        row.push(newData[r][c]);
      }
      patch.push(row);
    }

    if (!copyToRangeFormula(sheet, area[0] + firstRow, area[1] + firstColumn, patch)) return false;
  }

  return true;
}

// Update data when the table is sorted (insert, update or remove rows accordingly)
function UpdateSortedRangeData(sheet, firstRow, firstColumn, lastColumn, keyColumns, newData) {
  const currentData = sheet
    .getRange(firstRow, firstColumn, sheet.getMaxRows() - firstRow + 1, lastColumn - firstColumn + 1)
    .getValues();
  let rowCopy = -1;
  let cMin;
  let cMax;
  let bufferSize = 0;
  let incomplete = false;

  const copyArea = [];

  // Sort the data, just to be completely sure it's fine (ex: GearTier data is sorted... except HumanTorch is before Hulk)
  SortTable(newData, 2);

  // Check all lines of the current sheet
  for (let r = 0; r < currentData.length && r < newData.length; r++) {
    //output("row... " + r);
    //output(" > currentData[r][0] = " + currentData[r][0]);
    //output(" > newData[r][0] = " + newData[r][0]);

    const keyDiff = CompareData(currentData[r], newData[r], 0, keyColumns);
    //output(" - keyDiff = " + keyDiff);
    //  0: it's a match
    // -1: insert this one, check next
    //  1: delete current

    // Delete one or multiple rows
    if (keyDiff > 0) {
      let r2 = r + 1;
      for (; r2 < currentData.length; r2++) {
        const keyDiff2 = CompareData(currentData[r2], newData[r], 0, keyColumns);
        if (keyDiff2 <= 0) break;
      }

      const count = r2 - r;
      currentData.splice(r, count);
      sheet.deleteRows(firstRow + r, count);
      //output(" - Deleted " + count + " rows starting row " + (firstRow + r));
      r--; // Next loop, check the same line now that the unwanted lines have been removed
      continue;
    }

    // Insert one or multiple rows
    if (keyDiff < 0) {
      let r2 = r + 1;
      let thenExit = false;
      if (!canSetValues(newData[0].length)) return false;

      for (; r2 < newData.length; r2++) {
        const keyDiff2 = CompareData(currentData[r], newData[r2], 0, keyColumns);
        if (keyDiff2 >= 0) break;
        if (!canSetValues(newData[0].length * (r2 - r))) {
          r2--;
          thenExit = true;
          break;
        }
      }

      const count = r2 - r;
      if (count > 0) {
        sheet.insertRowsBefore(firstRow + r, count);
        const patch = [];
        for (let row = r; row < r2; row++) {
          patch.push(newData[row]);
          currentData.splice(r, 0, newData[row]);
        }
        //output(" - Inserted " + count + " rows starting row " + (firstRow + r));
        if (!copyToRange(sheet, firstRow + r, firstColumn, patch)) return false;
      }

      if (thenExit) return false;

      r += count - 1; // Skip all inserted lines, we know they are correct
      continue;
    }

    // Same key, different value: add to the copy area
    let first;
    //output('First...');
    for (first = keyColumns; first < currentData[r].length; first++) {
      if (currentData[r][first] == '' && newData[r][first] == null) continue;
      if (currentData[r][first] != newData[r][first]) break;
    }

    // No difference in the value either!
    if (first == currentData[r].length) {
      // End of the recorded difference, let's keep track of it
      if (rowCopy >= 0) {
        copyArea.push([rowCopy, cMin, r, cMax]);
        bufferSize += (r - rowCopy) * (cMax - cMin);
        rowCopy = -1;
      }
      continue;
    }

    let last;
    for (last = currentData[r].length - 1; last > first; last--) {
      if (currentData[r][last] == '' && newData[r][last] == null) continue;
      if (currentData[r][last] != newData[r][last]) {
        break;
      }
    }
    last++;

    if (rowCopy >= 0) {
      // Extend the area if necessary
      if (cMin < first) first = cMin;
      if (cMax > last) last = cMax;
    } else rowCopy = r;

    if (!canSetValues(bufferSize + (last - first) * (r - rowCopy))) {
      // Too much data to copy, remove this line and stop here for now
      r--;
      if (r == rowCopy && bufferSize == 0) return false;

      if (r > rowCopy) {
        // Copy the current buffer
        copyArea.push([rowCopy, cMin, r, cMax]);
      }
      rowCopy = -1;
      incomplete = true;
      break;
    }

    cMin = first;
    cMax = last;
    //output(" - Start area = " + rowCopy + "/" + cMin + "/" + cMax);
  }

  // Push last area
  if (rowCopy >= 0) {
    copyArea.push([rowCopy, cMin, newData.length, cMax]);
    //output(" - copyArea += " + copyArea[copyArea.length - 1]);
  } else if (newData.length > currentData.length) {
    copyArea.push([currentData.length, 0, newData.length, newData[0].length]);
    //output(" - copyArea += " + copyArea[copyArea.length - 1]);
  }

  // Add or remove rows if necessary
  if (newData.length > currentData.length) {
    //output("Added " + (newData.length - currentData.length) + " rows");
    sheet.insertRowsAfter(sheet.getMaxRows(), newData.length - currentData.length);
  } else if (newData.length < currentData.length) {
    const count = currentData.length - newData.length;
    //output("Removed " + count + " rows");
    sheet.deleteRows(sheet.getMaxRows() - count + 1, count);
  }

  // Patch data
  for (let i = 0; i < copyArea.length; i++) {
    const area = copyArea[i];
    const patch = [];
    for (let r = area[0]; r < area[2]; r++) {
      const row = [];
      for (let c = area[1]; c < area[3]; c++) {
        row.push(newData[r][c]);
      }
      patch.push(row);
    }

    //output("- Copy to range " + area);
    if (!copyToRange(sheet, area[0] + firstRow, area[1] + firstColumn, patch)) return false;
  }

  return !incomplete;
}
