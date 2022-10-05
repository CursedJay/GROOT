let _valueOf;

function valueOf(key) {
  const k = key.toString().toLowerCase();

  if (!_valueOf.hasOwnProperty(k)) return '';
  return _valueOf[k];
}

function setIdConverter(keyColumnName, valueColumnName) {
  _valueOf = new Object();
  _valueOf[''] = '';

  const keys = getNamedRangeValues(keyColumnName);
  const values = getNamedRangeValues(valueColumnName);

  for (let row = 0; row < keys.length; row++) {
    if (keys[row][0] == '' || values[row][0] == '') continue;
    _valueOf[keys[row][0].toLowerCase()] = values[row][0];
  }
}
