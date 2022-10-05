/*function GetHeaderIndexes(sheet)
{
    var res = {};
    var rangeData = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues();

    for (var i = 0; i < rangeData[0].length; i++)
    {
        res[rangeData[0][i]] = i;
    }

    return res;
}*/

function GetColumnValues(sheet, column, skipHeader) {
  const rangeData = sheet.getRange(skipHeader ? 2 : 1, column, sheet.getMaxRows() - skipHeader ? 1 : 0).getValues();
  let res = [];

  for (let i = 0; i < rangeData.length; i++) {
    res = rangeData[i][0];
  }

  return res;
}

function GetValuesIndexes(sheet, column, skipHeader) {
  const rangeData = sheet.getRange(skipHeader ? 2 : 1, column, sheet.getMaxRows() - skipHeader ? 1 : 0).getValues();
  const res = {};

  for (let i = 0; i < rangeData.length; i++) {
    res[rangeData[i][0]] = i;
  }

  return res;
}
