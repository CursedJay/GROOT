function getFolder(folderName, createIfDoesntExist) {
  var folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) return folders.next();
  if (!createIfDoesntExist) return null;

  folder = DriveApp.createFolder(folderName);
  DriveApp.addFolder(folder);

  return folder;
}

function getFile(folder, filename) {
  var files = folder.getFilesByName(filename);

  while (files.hasNext()) {
    var file = files.next();
    return file;
  }
  return null;
}

function readJSON(fileId) {
  const file = DriveApp.getFileById(fileId);
  const content = file.getBlob().getDataAsString();
  return JSON.parse(content);
}
