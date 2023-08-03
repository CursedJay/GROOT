/** Calculate gear usage starting with a full wipe of cached data */
function api_CalculateFullGearUsage() {
  api_CalculateGearUsage(true);
}

/** Calculate gear usage for all characters included in farming */
function api_CalculateGearUsage(forceClean) {
  const craftedGearMap = updateCraftedGearUsage();

  if (!forceClean) {
    // Clear gear usage
    // 0               1
    // CraftedGearHash Needed
    const gearUsageUsageRange = getNamedRange('GearUsage_Usage');
    gearUsageUsageRange.clearContent();
  }

  const cachedCharIds = new Set();
  const gearUsageCacheRangeName = 'GearUsage_Cache';
  // 0           1        2             3          4           5               6
  // CharacterId GearTier CraftedGearId MaterialId MaterialQty CraftedGearHash Needed
  const gearUsageCacheRange = getNamedRange(gearUsageCacheRangeName);
  const gearUsageCache = gearUsageCacheRange.getValues();

  if (forceClean) {
    // If this is a force, then clear out the cached gear data
    gearUsageCacheRange.clearContent();
  } else {
    // Get the cached characters so we can skip retrieving any character that is cached already
    for (const gearUsageCacheRow of gearUsageCache) {
      const currentCharId = gearUsageCacheRow[0];

      if (!cachedCharIds.has(currentCharId)) {
        cachedCharIds.add(currentCharId);
      }
    }
  }

  // 0           1
  // CharacterId Include
  const idsAndInclusion = getNamedRangeValues('Farming_Ids_Inclusion');
  const idsNotInCache = new Set();
  const characterFarmingRowMap = {};

  for (let i = 0; i < idsAndInclusion.length; i++) {
    const idsAndInclusionRow = idsAndInclusion[i];
    const include = idsAndInclusionRow[1];

    if (include) {
      const charId = idsAndInclusionRow[0];

      // Always add to this map so we can calculate usage later
      characterFarmingRowMap[charId] = i;

      // If we're including someone in farming AND they're not in the cache, add them
      if (charId && !cachedCharIds.has(charId)) {
        idsNotInCache.add(charId);
      }
    }
  }

  // 0    1     2      3      4      5     6
  // Tier Focus Unique Damage Resist Armor Health
  const currentEquippedGears = getNamedRangeValues('Farming_Gear');
  const targetGearLevels = getNamedRangeValues('Farming_Gear_Target');
  const gearUsageCacheUpdatedValues = [];

  // Get gear data from API for anyone that isn't in the cache already
  if (idsNotInCache.size > 0) {
    for (const includedId of idsNotInCache) {
      const charId = includedId;
      const characterGear = GrootApi.getGearByCharacterId(charId);

      const currentGearLevel = currentEquippedGears[characterFarmingRowMap[charId]][0];

      // Get all the gear information for gears that are at the same tier or higher than current
      Object.entries(characterGear.gearTiers).forEach(([tierString, gearTier]) => {
        const tier = Number(tierString);
        if (tier < currentGearLevel) return;

        const slots = gearTier?.slots;
        if (!slots) return;

        for (const slot of slots) {
          if (slot.piece.flatCost) {
            Object.values(slot.piece.flatCost).forEach((flatItem) => {
              gearUsageCacheUpdatedValues.push([
                charId,
                tier,
                slot.piece.id,
                flatItem.item.id,
                flatItem.quantity,
                '',
                false
              ]);
            });
          } else {
            gearUsageCacheUpdatedValues.push([charId, tier, slot.piece.id, slot.piece.id, 1, '', false]);
          }
        }
      });
    }
  }

  let updateCount = 0;

  // Add the new gear rows to the gear cache values
  for (let i = 0; i < gearUsageCache.length; i++) {
    let currentRow = gearUsageCache[i];

    if (!currentRow[0] && updateCount < gearUsageCacheUpdatedValues.length) {
      currentRow = gearUsageCacheUpdatedValues[updateCount];
      gearUsageCache[i] = currentRow;
      updateCount++;
    }

    // Skip if there's not data in row
    const currentCharId = currentRow[0];
    if (!currentCharId) continue;

    // Skip if this character is not included in usage
    const farmingRow = characterFarmingRowMap[currentCharId];
    if (farmingRow === undefined) continue;

    // Skip if this tier is higher than the target tier
    const targetTier = targetGearLevels[farmingRow][0];
    const currentGearTier = currentRow[1];
    if (currentGearTier >= targetTier) continue;

    const currentCraftedGearId = currentRow[2];
    const currentEquippedGearRow = currentEquippedGears[farmingRow];

    // Check to see if it's already equipped
    if (
      isPieceEquippedAlready(currentEquippedGearRow, currentCraftedGearId, currentEquippedGearRow[0], currentGearTier)
    ) {
      currentRow[5] = 'AlreadyEquipped';
      continue;
    }

    // Check if there's a crafted piece we can use
    const foundForCurrent = craftedGearMap.find((craftedGearItem) => {
      return (
        // Crafted piece is already assigned to this slot
        (craftedGearItem.id === currentCraftedGearId &&
          craftedGearItem.characterId === currentCharId &&
          craftedGearItem.gearTier === currentGearTier) ||
        // Crafted piece is available for this slot
        (craftedGearItem.id === currentCraftedGearId &&
          craftedGearItem.characterId === '' &&
          craftedGearItem.gearTier === 0)
      );
    });

    if (foundForCurrent) {
      // Current gear is covered by crafted piece
      currentRow[5] = `${currentCharId}${currentGearTier}`;

      if (foundForCurrent.characterId === '') {
        foundForCurrent.characterId = currentCharId;
        foundForCurrent.gearTier = currentGearTier;
      }
    } else {
      // Need materials for this piece
      currentRow[6] = true;
    }
  }

  // Write the gear cache to the page
  // NOTE: Bypassing helper method to write a LOT of rows (10,000) to the sheet
  SpreadsheetApp.getActive().getRangeByName(gearUsageCacheRangeName).setValues(gearUsageCache);
}

/** Determine if the piece of equipment is equipped already */
function isPieceEquippedAlready(currentEquippedGear, craftedGearId, craftedGearTier, currentGearTier) {
  if (craftedGearTier !== currentGearTier) return false;

  const slot = craftedGearId.substring(craftedGearId.lastIndexOf('_') + 1).toLowerCase();

  switch (slot) {
    case 'focus':
      return currentEquippedGear[1];
    case 'damage':
      return currentEquippedGear[3];
    case 'resist':
      return currentEquippedGear[4];
    case 'armor':
      return currentEquippedGear[5];
    case 'health':
      return currentEquippedGear[6];
    default:
      return currentEquippedGear[2];
  }
}

/** Get crafted gear usage */
function updateCraftedGearUsage() {
  // 0             1
  // CraftedGearId CharacterAndTier
  const gearUsageCraftedGearRangeName = 'GearUsage_CraftedGear';
  const gearUsageCraftedGearRange = getNamedRange(gearUsageCraftedGearRangeName);
  gearUsageCraftedGearRange.clearContent();
  const gearUsageCraftedGear = gearUsageCraftedGearRange.getValues();

  let currentCraftedGearIds = [];

  // Get quantities of crafted gear
  for (let i = 1; i <= 5; i++) {
    const craftedGearIds = getNamedRangeValues(`CraftedGear_Ids${i}`);
    const craftedGearQuantities = getNamedRangeValues(`CraftedGear_Values${i}`);
    currentCraftedGearIds = currentCraftedGearIds.concat(
      resolveCraftedGearQuantities(craftedGearIds, craftedGearQuantities)
    );
  }

  const craftedGearMap = [];

  for (let i = 0; i < currentCraftedGearIds.length; i++) {
    const craftedGearId = currentCraftedGearIds[i];
    gearUsageCraftedGear[i][0] = craftedGearId;
    craftedGearMap.push({
      id: `GEAR_${craftedGearId}`, // TODO: What's the real name of these items?
      characterId: '',
      gearTier: 0
    });
  }

  // Update crafted gear range
  setNamedRangeValues(gearUsageCraftedGearRangeName, gearUsageCraftedGear);

  return craftedGearMap;
}

function resolveCraftedGearQuantities(craftedGearIds, craftedGearQuantities) {
  const currentCraftedGearIds = [];

  for (let i = 0; i < craftedGearIds.length; i++) {
    const quantity = craftedGearQuantities[i][0];

    if (quantity > 0) {
      const currentCraftedGearId = craftedGearIds[i][0];

      for (let j = 0; j < quantity; j++) {
        currentCraftedGearIds.push(currentCraftedGearId);
      }
    }
  }

  return currentCraftedGearIds;
}
