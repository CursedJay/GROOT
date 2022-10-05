function copyToRange(sheet, row, column, data) {
  const total = data.length * data[0].length;

  if (_setValuesCount + total > _maxSetValue) return false;
  _setValuesCount += total;

  sheet.getRange(row, column, data.length, data[0].length).setValues(data);
  return true;
}

function copyToRangeFormula(sheet, row, column, data) {
  const total = data.length * data[0].length;

  if (_setValuesCount + total > _maxSetValue) return false;
  _setValuesCount += total;

  sheet.getRange(row, column, data.length, data[0].length).setFormulas(data);
  return true;
}
