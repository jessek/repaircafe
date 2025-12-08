function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate().setTitle('Repair Intake').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processNewIntake(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const values = sheet.getDataRange().getValues();
  
  // Validation: Check if Item ID already exists
  // Item ID is in index 1 (Column B)
  for (let i = 1; i < values.length; i++) {
    if (values[i][1].toString() === data.itemId.toString()) {
      throw new Error("Duplicate ID: An item with ID " + data.itemId + " is already checked in.");
    }
  }

  // If no duplicate, append the row
  sheet.appendRow([
    new Date(), 
    data.itemId, 
    data.clientName, 
    data.category, 
    data.itemName, 
    data.issue, 
    '', // Time Updated
    '', // Fixer Name
    '', // Resolution
    '', // Fixer Notes
    data.email, 
    data.mailingList, 
    data.clientIsFixer
  ]);
  
  return "Success! Item " + data.itemId + " has been registered.";
}

function lookupItem(itemId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][1].toString() === itemId.toString()) {
      return {
        clientName: values[i][2],
        category: values[i][3],
        itemName: values[i][4],
        issue: values[i][5],
        fixerName: values[i][7],
        resolution: values[i][8],
        notes: values[i][9],
        rowIndex: i + 1
      };
    }
  }
  throw new Error("Item ID not found.");
}

function processRepairUpdate(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  // Update columns G-J (Index 6-9): Time Updated, Fixer Name, Resolution, Fixer Notes
  sheet.getRange(data.rowIndex, 7, 1, 4).setValues([[new Date(), data.fixerName, data.resolution, data.notes]]);
  return "Repair record for ID " + data.itemId + " updated successfully.";
}

function getStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  const values = sheet.getDataRange().getValues();
  let stats = { total: values.length - 1, fixed: 0, diagnosed: 0, advised: 0, notfound: 0 };
  for (let i = 1; i < values.length; i++) {
    let res = values[i][8]; // Column I
    if (res === "Fixed") stats.fixed++;
    else if (res === "Diagnosed") stats.diagnosed++;
    else if (res === "Advised") stats.advised++;
    else if (res === "Client Not Found") stats.notfound++;
  }
  return stats;
}