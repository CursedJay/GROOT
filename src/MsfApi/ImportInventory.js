function api_importInventory(since) {
  const inventoryVersionCell = getNamedRangeValue('Inventory_Since');
  const _since = since ?? (inventoryVersionCell ? inventoryVersionCell : 'fresh');
  const inventory = GrootApi.getInventory(_since);

  if (inventory === false) return;

  for (let column = 1; column <= 5; column++) {
    const ids = getNamedRangeValues(`Inventory_Ids${column}`);
    const values = getNamedRangeValues(`Inventory_Values${column}`);

    for (let row = 0; row < ids.length; row++) {
      const id = `GEAR_${ids[row][0]}`;
      if (inventory.hasOwnProperty(id)) values[row] = [inventory[id]];
    }

    setNamedRangeValues(`Inventory_Values${column}`, values);
  }
}

function api_importFullInventory() {
  api_importInventory('fresh');
}
