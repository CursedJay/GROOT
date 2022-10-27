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

    updateInventoryValues(inventory.gear, matIdsRange, matValuesRange);
    updateInventoryValues(inventory.gear, craftedIdsRange, craftedValuesRange);
  }
  inventoryVersionCell.setValue(inventory.since);
}

function api_importFullInventory() {
  api_importInventory('fresh');
}

function updateInventoryValues(data, idsRange, valuesRange) {
  const ids = getNamedRangeValues(idsRange);
  const values = getNamedRangeValues(valuesRange);

  for (let row = 0; row < ids.length; row++) {
    const id = `GEAR_${ids[row][0]}`;
    if (data.hasOwnProperty(id)) values[row] = [data[id]];
  }

  setNamedRangeValues(valuesRange, values);
}

function saveInventoryJSON() {
  const inventory = GrootApi.getGearInventory('fresh');
  const dateTime = Utilities.formatDate(new Date(), 'GMT+8', "yyyy-MM-dd'T'HH:mm:ss.SS");
  const filename = `msf_inventory_${dateTime}.json`;

  saveReqToFile(inventory, filename);
}
