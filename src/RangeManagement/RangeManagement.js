function GetSheet(sheetName) {
  return SpreadsheetApp.getActive().getSheetByName(sheetName);
}

function ResizeSheet(sheet, rows, columns) {
  const currentrows = sheet.getMaxRows();
  const currentcolumns = sheet.getMaxColumns();

  if (currentrows < rows) {
    // Insert before the last row to make sure we keep named ranges
    sheet.insertRowsBefore(currentrows, rows - currentrows);
  } else if (currentrows > rows) {
    // Delete last rows, we want to keep the header and potential special formulas
    sheet.deleteRows(rows + 1, currentrows - rows);
  }

  if (columns > 0) {
    if (currentcolumns < columns) {
      sheet.insertColumnsBefore(currentcolumns, columns - currentcolumns);
    } else if (currentcolumns > columns) {
      sheet.deleteColumns(columns + 1, currentcolumns - columns);
    }
  }
}

function ClearSheet(sheet, clearHeader) {
  if (clearHeader) {
    sheet.clear();
    return;
  }

  sheet.getRange(2, 1, sheet.getMaxRows() - 1, sheet.getMaxColumns()).clear();
}

function CompareData(currentData, newData, start, count) {
  for (let c = 0; c < count; c++) {
    if (currentData[start + c] == '' && newData[start + c] == null) continue;
    if (newData[start + c] < currentData[start + c]) return -1;
    if (newData[start + c] > currentData[start + c]) return 1;
  }
  return 0;
}

function GetEmptyRows(range) {
  let first = -1;
  const ranges = [];
  for (let r = 0; r < range.length; r++) {
    const isEmpty = range[r][0] == '';

    if (first < 0 && isEmpty) first = r + 1;
    else if (first >= 0 && !isEmpty) {
      ranges.push([first, r + 1]);
      first = -1;
    }
  }
  if (first >= 0) ranges.push([first, range.length + 1]);

  return ranges;
}

function SortTable(data, keyColumns) {
  data.sort((x, y) => {
    return CompareData(y, x, 0, keyColumns);
  });
}

function GetLastRowWithData(range) {
  for (let i = 0; range.length > i; i++) {
    if (!range[i][0]) {
      return i;
    }
  }
  return range.length;
}
