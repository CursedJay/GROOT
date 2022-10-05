function DS_Update_Links() {
  const file = DriveApp.getFileById('1MSKK15zDqC2J8AsX7T4GQ3EkKkluXcHP');
  const content = file.getBlob().getDataAsString();
  const json = JSON.parse(content);

  var range = getNamedRange('Links_Sources');
  range.clearContent();

  var data = [];

  for (var i = 0; i < range.getHeight(); i++) {
    data[i] = ['', '', ''];
    if (i < json.ContentCreators.length)
      data[i][0] = '=HYPERLINK("' + json.ContentCreators[i].URL + '", "' + json.ContentCreators[i].Title + '")';

    if (i < json.Tools.length) data[i][1] = '=HYPERLINK("' + json.Tools[i].URL + '", "' + json.Tools[i].Title + '")';

    if (i < json.Servers.length)
      data[i][2] = '=HYPERLINK("' + json.Servers[i].URL + '", "' + json.Servers[i].Title + '")';
  }

  range.setFormulas(data);
}
