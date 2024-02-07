function getFolder(folderName, createIfDoesntExist) {
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) return folders.next();
  if (!createIfDoesntExist) return null;

  const folder = DriveApp.createFolder(folderName);
  DriveApp.addFolder(folder);

  return folder;
}

function getFile(folder, filename) {
  const files = folder.getFilesByName(filename);

  while (files.hasNext()) {
    const file = files.next();
    return file;
  }
  return null;
}

function readJSON(fileId) {
  const file = DriveApp.getFileById(fileId);
  const content = file.getBlob().getDataAsString();
  return JSON.parse(content);
}
