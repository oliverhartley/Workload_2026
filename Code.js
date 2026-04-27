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
  
  // 1. Define the expected headers in the SPECIFIC ORDER requested by the user
  var expectedHeaders = [
    "Partner Name", "Account: Account Name", "Primary Workload Pillar", "Workload: Workload Name",
    "Workload Progress", "Workload Gross Annual Recurring Revenue (converted)", "Workload: Owner Name",
    "Technical Win Date", "Account: Micro Region", "Account: Billing Country", "Workload: ID",
    "Link Workload", "Primary CE Technical Owner", "Partner", "Workload Gross Month Recurring Revenue (converted)",
    "Workload End Date", "Tier", "DSR", "DCE", "PSF Investment", "Account: Account Owner"
  ];
  
  // 2. Read target data to map by ID
  var targetData = targetSheet.getDataRange().getValues();
  var targetHeaders = targetData[0];
  
  // If sheet was empty or headers don't match, write headers
  if (targetData.length === 1 && targetData[0][0] === "") {
    targetSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    targetData = [expectedHeaders];
    targetHeaders = expectedHeaders;
  }
  
  var idColIndex = expectedHeaders.indexOf("Workload: ID");
  if (idColIndex === -1) {
    throw new Error("Critical: 'Workload: ID' column not defined in expected headers.");
  }
  
  var targetMap = {};
  for (var i = 1; i < targetData.length; i++) {
    var row = targetData[i];
    var id = row[idColIndex];
    if (id) {
      targetMap[id] = { row: i + 1, values: row };
    }
  }
  
  // 3. Read source data
  var sourceData = sourceSheet.getDataRange().getValues();
  var actualSourceHeaders = sourceData[0];
  
  // 4. Find indices of expected headers in source sheet
  var headerIndices = {};
  expectedHeaders.forEach(function(header) {
    var index = actualSourceHeaders.indexOf(header);
    headerIndices[header] = index;
  });
  
  var sourceIdColIndex = headerIndices["Workload: ID"];
  if (sourceIdColIndex === -1) {
    throw new Error("Critical: 'Workload: ID' column not found in source sheet.");
  }
  
  // 5. Process data and sync
  for (var i = 1; i < sourceData.length; i++) {
    var sourceRow = sourceData[i];
    var id = sourceRow[sourceIdColIndex];
    
    if (id) {
      var newRowValues = [];
      expectedHeaders.forEach(function(header) {
        var index = headerIndices[header];
        if (index !== -1) {
          newRowValues.push(sourceRow[index]);
        } else {
          newRowValues.push("");
        }
      });
      
      var targetRecord = targetMap[id];
      
      if (!targetRecord) {
        // New Row (Light Green)
        targetSheet.appendRow(newRowValues);
        var lastRow = targetSheet.getLastRow();
        targetSheet.getRange(lastRow, 1, 1, newRowValues.length).setBackground("#E2EFDA");
        Logger.log("Added new row for ID: " + id);
      } else {
        // Existing Row, compare values
        var targetValues = targetRecord.values;
        var isChanged = false;
        
        for (var j = 0; j < newRowValues.length; j++) {
          if (String(newRowValues[j]) !== String(targetValues[j])) {
            isChanged = true;
            break;
          }
        }
        
        if (isChanged) {
          // Update Row (Light Yellow)
          var targetRowNumber = targetRecord.row;
          targetSheet.getRange(targetRowNumber, 1, 1, newRowValues.length).setValues([newRowValues]);
          targetSheet.getRange(targetRowNumber, 1, 1, newRowValues.length).setBackground("#FFF2CC");
          Logger.log("Updated row for ID: " + id);
        } else {
          // No change, clear background
          var targetRowNumber = targetRecord.row;
          targetSheet.getRange(targetRowNumber, 1, 1, newRowValues.length).setBackground(null);
        }
      }
    }
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
