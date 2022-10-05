var _valueOf;

function valueOf(key) {
  var k = key.toString().toLowerCase();

  if (!_valueOf.hasOwnProperty(k)) return '';
  return _valueOf[k];
}

function setIdConverter(keyColumnName, valueColumnName) {
  _valueOf = new Object();
  _valueOf[''] = '';

  var keys = getNamedRangeValues(keyColumnName);
  var values = getNamedRangeValues(valueColumnName);

  for (var row = 0; row < keys.length; row++) {
    if (keys[row][0] == '' || values[row][0] == '') continue;
    _valueOf[keys[row][0].toLowerCase()] = values[row][0];
  }
}
