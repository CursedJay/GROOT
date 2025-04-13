function exportRosterToJSON() {
  const roster = GrootApi.getRoster('fresh');
  if (!roster) return;
  saveReqToFile(roster.data, 'msf_roster', 'json');
}

function exportInventoryToJSON() {
  exportInventoryToFile('msf_inventory', 'json', false);
}

function exportInventoryLocalizedToJSON() {
  exportInventoryToFile('msf_inventory_local', 'json', true);
}

function exportInventoryToCSV() {
  exportInventoryToFile('msf_inventory', 'csv', false);
}

function exportInventoryLocalizedToCSV() {
  exportInventoryToFile('msf_inventory_local', 'csv', true);
}

function exportInventoryToFile(filename = 'Inventory', filetype = 'json', localized = false) {
  const inventory = GrootApi.getInventoryByType('GEAR');
  if (!inventory) return;

  let data;

  if (localized) {
    const sheet = GetSheet('_M3Localization_Gear');
    const localizedGear = sheet.getRange(1, 1, sheet.getLastRow(), 2).getValues();

    if (filetype === 'csv') {
      data = '';

      for (const [id, qty] of Object.entries(inventory.gear)) {
        const _id = id.replace('GEAR_', '');
        const locName = localizedGear.filter((elem) => {
          return !!~elem.indexOf(_id);
        })[0][0];
        data += `${_id},${locName},${qty}\n`;
      }
    } else {
      data = {};

      for (const [id, qty] of Object.entries(inventory.gear)) {
        const _id = id.replace('GEAR_', '');
        const locName = localizedGear.filter((elem) => {
          return !!~elem.indexOf(_id);
        })[0][0];

        data[_id] = { name: locName, quantity: qty };
      }
    }
  } else {
    if (filetype === 'csv') {
      data = '';
      for (const [id, qty] of Object.entries(inventory.gear)) {
        data += `${id},${qty}\n`;
      }
    } else {
      data = inventory.gear;
    }
  }
  saveReqToFile(data, filename, filetype);
}
