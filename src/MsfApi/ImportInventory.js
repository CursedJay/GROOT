// TODO add param itemType
function api_importInventory(since) {
  const inventoryVersionCell = getNamedRange('_Version_Inventory');
  const inventoryVersionValue = inventoryVersionCell.getValue();
  const _since = since ?? (inventoryVersionValue ? inventoryVersionValue : 'fresh');
  const inventory = GrootApi.getGearInventory(_since);

  if (inventory === false) return;

  for (let column = 1; column <= 5; column++) {
    const matIdsRange = `Inventory_Ids${column}`;
    const matValuesRange = `Inventory_Values${column}`;
    const craftedIdsRange = `CraftedGear_Ids${column}`;
    const craftedValuesRange = `CraftedGear_Values${column}`;

    updateInventoryValues(inventory.gear, matIdsRange, matValuesRange, 'GEAR_');
    updateInventoryValues(inventory.gear, craftedIdsRange, craftedValuesRange, 'GEAR_');
  }
  inventoryVersionCell.setValue(inventory.since);
}

function api_importFullInventory() {
  api_importInventory('fresh');
}

function updateInventoryValues(data, idsRangeName, valuesRangeName, idPrefix = '') {
  const ids = getNamedRangeValues(idsRangeName);
  const values = getNamedRangeValues(valuesRangeName);

  for (let row = 0; row < ids.length; row++) {
    const id = `${idPrefix}${ids[row][0]}`;

    if (data.hasOwnProperty(id)) {
      values[row] = [data[id]];
    }
  }

  setNamedRangeValues(valuesRangeName, values);
}

function api_importIso(since) {
  const idRanges = [];
  const valueRanges = [];

  for (let column = 1; column <= 10; column++) {
    idRanges.push(`Inventory_Ids_Iso_${column}`);
    valueRanges.push(`Inventory_Values_Iso_${column}`);
  }
  api_ImportInventoryGeneric(since, 'ISOITEM', '_Version_Iso', idRanges, valueRanges);
}

function api_importFullIso() {
  api_importIso('fresh');
}

function api_importTraining(since) {
  api_ImportInventoryGeneric(
    since,
    'CONSUMABLE',
    '_Version_Training',
    ['Inventory_Ids_Training'],
    ['Inventory_Values_Training']
  );
}

function api_importFullTraining() {
  api_importTraining('fresh');
}

function api_importAbility(since) {
  api_ImportInventoryGeneric(
    since,
    'ABILITY_MATERIAL',
    '_Version_Ability',
    ['Inventory_Ids_Ability'],
    ['Inventory_Values_Ability']
  );
}

function api_importFullAbility() {
  api_importAbility('fresh');
}

function api_ImportInventoryGeneric(since, itemType, versionRangeName, idRangeNames, valueRangeNames) {
  const inventoryVersionCell = getNamedRange(versionRangeName);
  const inventoryVersionValue = inventoryVersionCell.getValue();
  const _since = since ?? (inventoryVersionValue ? inventoryVersionValue : 'fresh');

  const inventory = GrootApi.getInventoryByType(itemType, _since);
  if (inventory === false) return;

  for (let rangeCount = 0; rangeCount < idRangeNames.length; rangeCount++) {
    updateInventoryValues(inventory.items, idRangeNames[rangeCount], valueRangeNames[rangeCount]);
  }

  inventoryVersionCell.setValue(inventory.since);
}

function saveInventoryJSON() {
  const inventory = GrootApi.getGearInventory('fresh');
  const dateTime = Utilities.formatDate(new Date(), 'GMT+8', "yyyy-MM-dd'T'HH:mm:ss.SS");
  const filename = `msf_inventory_${dateTime}.json`;

  saveReqToFile(inventory, filename);
}
