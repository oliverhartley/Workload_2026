function syncWorkloadsFromScratch() {
  var sourceSpreadsheetId = "1snf2ryBk7Lizdu5FwTf-LpRc70KN4I-W2-LECgEOJZU";
  var targetSpreadsheetId = "1qsB7bD_26sUie6OyW-uyty-r9W3LYTdFjXgIYel3zck";
  
  var sourceSheetName = "Oliver - Workloads Partners";
  var targetSheetName = "Synced Workloads";
  
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error("Source sheet not found.");
  }
  
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  if (!targetSheet) {
    targetSheet = targetSpreadsheet.insertSheet(targetSheetName);
    Logger.log("Created target sheet: " + targetSheetName);
  }
  
  // 1. Define the expected headers based on user request
  var expectedHeaders = [
    "Account: Micro Region", "Account: Billing Country", "Workload: ID", "Workload: Workload Name",
    "Link Workload", "Primary CE Technical Owner", "Workload: Owner Name", "Workload Progress",
    "Primary Workload Pillar", "Partner", "Partner Name", "Workload Gross Annual Recurring Revenue (converted)",
    "Workload Gross Month Recurring Revenue (converted)", "Technical Win Date", "Workload End Date",
    "Account: Account Name", "Tier", "DSR", "DCE", "PSF Investment", "Account: Account Owner"
  ];
  
  // 2. Read source data
  var sourceData = sourceSheet.getDataRange().getValues();
  var actualSourceHeaders = sourceData[0];
  
  // 3. Find indices of expected headers in source sheet
  var headerIndices = {};
  expectedHeaders.forEach(function(header) {
    var index = actualSourceHeaders.indexOf(header);
    if (index === -1) {
      Logger.log("Warning: Header '" + header + "' not found in source sheet.");
    }
    headerIndices[header] = index;
  });
  
  // Find Workload: ID index specifically as it's the key
  var idColIndex = headerIndices["Workload: ID"];
  if (idColIndex === -1) {
    throw new Error("Critical: 'Workload: ID' column not found in source sheet.");
  }
  
  // 4. Clear target sheet and set new headers
  targetSheet.clear();
  targetSheet.appendRow(expectedHeaders);
  
  // 5. Process data and write to target
  var rowsToWrite = [];
  
  for (var i = 1; i < sourceData.length; i++) {
    var sourceRow = sourceData[i];
    var id = sourceRow[idColIndex];
    
    if (id) {
      var newRow = [];
      expectedHeaders.forEach(function(header) {
        var index = headerIndices[header];
        if (index !== -1) {
          newRow.push(sourceRow[index]);
        } else {
          newRow.push(""); // Fill with empty string if header not found in source
        }
      });
      rowsToWrite.push(newRow);
    }
  }
  
  if (rowsToWrite.length > 0) {
    targetSheet.getRange(2, 1, rowsToWrite.length, expectedHeaders.length).setValues(rowsToWrite);
    Logger.log("Synced " + rowsToWrite.length + " rows to target sheet.");
  } else {
    Logger.log("No rows with valid IDs found to sync.");
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync Menu')
      .addItem('Sync Workloads', 'syncWorkloadsFromScratch')
      .addItem('Sync Expert Requests', 'syncExpertRequests')
      .addItem('Send Workload Emails', 'sendWorkloadEmails')
      .addItem('Create Owner Spreadsheets', 'createOwnerSheets')
      .addToUi();
}
