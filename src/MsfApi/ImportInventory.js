// TODO add param itemType
function api_importInventory(since) {
  const inventoryVersionCell = getNamedRangeValue('Inventory_Since');
  const _since = since ?? (inventoryVersionCell ? inventoryVersionCell : 'fresh');
  const inventory = GrootApi.getInventory(_since);
  saveReqToFile(inventory, 'MEGATEST YEAH.json');

  if (inventory === false) return;

  for (let column = 1; column <= 5; column++) {
    const matIdsRange = `Inventory_Ids${column}`;
    const matValuesRange = `Inventory_Values${column}`;
    const craftedIdsRange = `CraftedGear_Ids${column}`;
    const craftedValuesRange = `CraftedGear_Values${column}`;

    updateInventoryValues(inventory, matIdsRange, matValuesRange);
    updateInventoryValues(inventory, craftedIdsRange, craftedValuesRange);
  }
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
