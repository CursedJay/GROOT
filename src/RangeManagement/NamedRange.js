function getNamedRangeValueOn(doc, name) {
  return doc.getRangeByName(name).getValue();
}

function getNamedRangeValuesOn(doc, name) {
  return doc.getRangeByName(name).getValues();
}

function getNamedRangeIntegerOn(doc, name) {
  try {
    return parseInt(getNamedRangeValueOn(doc, name));
  } catch (e) {
    return null;
  }
}

let _setValuesCount = 0;
const _maxSetValue = 20000;
function setNamedRangeValueOn(doc, name, value) {
  if (_setValuesCount + 1 > _maxSetValue) return false;

  _setValuesCount++;
  doc.getRangeByName(name).setValue(value);
  return true;
}

function setNamedRangeValuesOn(doc, name, values) {
  const total = values.length * values[0].length;
  if (_setValuesCount + total > _maxSetValue) return false;

  _setValuesCount += total;
  doc.getRangeByName(name).setValues(values);
  return true;
}

function getNamedRange(name) {
  return SpreadsheetApp.getActive().getRangeByName(name);
}

function getNamedRangeValue(name) {
  return getNamedRangeValueOn(SpreadsheetApp.getActive(), name);
}

function getNamedRangeValues(name) {
  return getNamedRangeValuesOn(SpreadsheetApp.getActive(), name);
}

function getNamedRangeInteger(name) {
  return getNamedRangeIntegerOn(SpreadsheetApp.getActive(), name);
}

function setNamedRangeValue(name, value) {
  return setNamedRangeValueOn(SpreadsheetApp.getActive(), name, value);
}

function setNamedRangeValues(name, values) {
  return setNamedRangeValuesOn(SpreadsheetApp.getActive(), name, values);
}

function getSetValuesCount() {
  return _setValuesCount;
}

function getSetValuesMax() {
  return _maxSetValue;
}

function canSetValues(count) {
  return getSetValuesCount() + count <= getSetValuesMax();
}

function setNamedRangeValuesResize(name, values) {
  let range = getNamedRange(name);
  range.clearContent();

  if (values.length == 0) return true;

  const total = values.length * values[0].length;
  if (_setValuesCount + total > _maxSetValue) return false;

  range = range.getSheet().getRange(range.getRow(), range.getColumn(), values.length, values[0].length);

  range.setValues(values);

  _setValuesCount += total;
  return true;
}
