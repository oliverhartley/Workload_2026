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
  
  // If sheet was empty, use expectedHeaders as default
  if (targetData.length === 1 && targetData[0][0] === "") {
    targetSheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    targetData = [expectedHeaders];
    targetHeaders = expectedHeaders;
  }
  
  var idColIndex = targetHeaders.indexOf("Workload: ID");
  if (idColIndex === -1) {
    throw new Error("Critical: 'Workload: ID' column not found in target sheet.");
  }
  
  var progressColIndex = targetHeaders.indexOf("Workload Progress");
  
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
  targetHeaders.forEach(function(header) {
    var index = actualSourceHeaders.indexOf(header);
    headerIndices[header] = index;
  });
  
  var sourceIdColIndex = actualSourceHeaders.indexOf("Workload: ID");
  if (sourceIdColIndex === -1) {
    throw new Error("Critical: 'Workload: ID' column not found in source sheet.");
  }
  
  // 5. Process data and sync
  for (var i = 1; i < sourceData.length; i++) {
    var sourceRow = sourceData[i];
    var id = sourceRow[sourceIdColIndex];
    
    if (id) {
      var targetRecord = targetMap[id];
      var newRowValues = [];
      
      if (targetRecord) {
        // Start with existing target values
        newRowValues = targetRecord.values.slice();
        
        // Overwrite with source values where available
        targetHeaders.forEach(function(header, j) {
          var sourceIndex = headerIndices[header];
          if (sourceIndex !== -1) {
            newRowValues[j] = sourceRow[sourceIndex];
          }
        });
      } else {
        // New Row, build from scratch
        targetHeaders.forEach(function(header) {
          var sourceIndex = headerIndices[header];
          if (sourceIndex !== -1) {
            newRowValues.push(sourceRow[sourceIndex]);
          } else {
            newRowValues.push("");
          }
        });
      }
      
      if (!targetRecord) {
        // New Row (Light Green)
        targetSheet.appendRow(newRowValues);
        var lastRow = targetSheet.getLastRow();
        targetSheet.getRange(lastRow, 1, 1, newRowValues.length).setBackground("#E2EFDA");
        Logger.log("Added new row for ID: " + id);
        
        // Log to Change Log sheet
        logChange(targetSpreadsheet, id, lastRow, "Insert", "New row added");
      } else {
        // Existing Row, compare values
        var targetValues = targetRecord.values;
        var isChanged = false;
        
        if (progressColIndex !== -1) {
          if (normalizeValue(newRowValues[progressColIndex]) !== normalizeValue(targetValues[progressColIndex])) {
            isChanged = true;
          }
        } else {
          // Fallback if header not found, check all columns
          for (var j = 0; j < newRowValues.length; j++) {
            if (normalizeValue(newRowValues[j]) !== normalizeValue(targetValues[j])) {
              isChanged = true;
              break;
            }
          }
        }
        
        var targetRowNumber = targetRecord.row;
        if (isChanged) {
          // Update Row (Light Yellow)
          targetSheet.getRange(targetRowNumber, 1, 1, newRowValues.length).setValues([newRowValues]);
          targetSheet.getRange(targetRowNumber, 1, 1, newRowValues.length).setBackground("#FFF2CC");
          Logger.log("Updated row for ID: " + id);
          
          // Log details of changes (only for Workload Progress or all if fallback used)
          if (progressColIndex !== -1) {
            var detail = "Workload Progress (Column E): '" + targetValues[progressColIndex] + "' -> '" + newRowValues[progressColIndex] + "'";
            logChange(targetSpreadsheet, id, targetRowNumber, "Update", detail);
          } else {
            var details = [];
            for (var j = 0; j < newRowValues.length; j++) {
              if (normalizeValue(newRowValues[j]) !== normalizeValue(targetValues[j])) {
                details.push(targetHeaders[j] + ": '" + targetValues[j] + "' -> '" + newRowValues[j] + "'");
              }
            }
            logChange(targetSpreadsheet, id, targetRowNumber, "Update", details.join(", "));
          }
        } else {
          // No change, clear background
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

function normalizeValue(val) {
  if (val === null || val === undefined) return "";
  var str = String(val).trim();
  
  if (/^(USD|\$)?\s*[\d,]+(\.\d+)?$/.test(str)) {
    str = str.replace(/^(USD|\$)/, "").replace(/,/g, "").trim();
    var num = parseFloat(str);
    if (!isNaN(num)) {
      return String(Math.round(num));
    }
  }
  
  return str;
}

function logChange(spreadsheet, id, rowNumber, action, details) {
  var logSheetName = "Change Log";
  var logSheet = spreadsheet.getSheetByName(logSheetName);
  if (!logSheet) {
    logSheet = spreadsheet.insertSheet(logSheetName);
    logSheet.appendRow(["Timestamp", "Workload: ID", "Row Number", "Action", "Details"]);
    logSheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  logSheet.appendRow([new Date(), id, rowNumber, action, details]);
}
