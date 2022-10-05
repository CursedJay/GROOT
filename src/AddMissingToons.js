function AddMissingToons() {
  // Complete list of characters available ingame
  const fullRange = getNamedRange('_HeroData_Id');
  const fullList = fullRange.getValues();

  // Current characters listin the user's roster tab
  const currentRange = getNamedRange('Roster_Id');
  const currentList = currentRange.getValues();
  const currentCount = currentRange.getHeight();

  const fullCount = fullList.length;
  let availablePos = -1;

  output('Full toons list:' + fullCount);
  output('Old sheet toons:' + currentCount);

  // First, clear duplicates and invalid toons
  for (let i = 0; i < currentCount; i++) {
    const id = currentList[i][0];

    // Empty line: ignore
    if (id == '' || id == null) {
      if (availablePos < 0) availablePos = i;
      continue;
    }

    // Remove if not in the new list
    let found = false;
    for (let j = 0; j < fullCount; j++) {
      if (id == fullList[j][0]) {
        found = true;
        fullList[j][0] = null; // No need to check if this toon is missing, and make sure duplicates are removed
        break;
      }
    }

    if (!found) {
      currentList[i][0] = null;
      if (availablePos < 0) availablePos = i;
      continue;
    }
  }

  if (availablePos >= 0) {
    // Add missing toons (anything in fullList is missing from current list)
    const pos = 0;
    for (let i = 0; i < fullCount; i++) {
      const id = fullList[i][0];
      if (id == null) continue;

      currentList[availablePos][0] = id;
      while (currentList[availablePos][0] != '' && currentList[availablePos][0] != null) {
        availablePos++;
        if (availablePos >= currentCount) break;
      }
      if (availablePos >= currentCount) break;
    }
  }

  currentRange.setValues(currentList);
}
