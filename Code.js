function copyWorkloadPartnersSheet() {
  var sourceSpreadsheetId = "1snf2ryBk7Lizdu5FwTf-LpRc70KN4I-W2-LECgEOJZU";
  var sourceSheetName = "Oliver - Workloads Partners";
  
  // Open the source spreadsheet and get the sheet
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error("Source sheet not found: " + sourceSheetName);
  }
  
  // Open the target spreadsheet. 
  // Assuming this script is bound to the target spreadsheet.
  // If not, replace with SpreadsheetApp.openById("YOUR_TARGET_SPREADSHEET_ID")
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!targetSpreadsheet) {
    throw new Error("Target spreadsheet not found. Is this script bound to a spreadsheet?");
  }
  
  // Check if a sheet with the same name already exists in target
  var existingSheet = targetSpreadsheet.getSheetByName(sourceSheetName);
  if (existingSheet) {
    // Option: Delete the existing sheet to replace it
    targetSpreadsheet.deleteSheet(existingSheet);
    Logger.log("Deleted existing sheet in target: " + sourceSheetName);
  }
  
  // Copy the sheet to the target spreadsheet
  var copiedSheet = sourceSheet.copyTo(targetSpreadsheet);
  
  // Rename the copied sheet to the original name
  copiedSheet.setName(sourceSheetName);
  
  Logger.log("Successfully copied sheet to target: " + sourceSheetName);
}

function syncWorkloadPartnersSheet() {
  var sourceSpreadsheetId = "1snf2ryBk7Lizdu5FwTf-LpRc70KN4I-W2-LECgEOJZU";
  var sourceSheetName = "Oliver - Workloads Partners";
  
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error("Source sheet not found: " + sourceSheetName);
  }
  
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = targetSpreadsheet.getSheetByName(sourceSheetName);
  
  // If target sheet doesn't exist, create it by copying the source
  if (!targetSheet) {
    targetSheet = sourceSheet.copyTo(targetSpreadsheet);
    targetSheet.setName(sourceSheetName);
    Logger.log("Created target sheet by copying source.");
    // We should still add the Workload_ID column if it's not there
  }
  
  var lastRow = sourceSheet.getLastRow();
  if (lastRow <= 1) {
    Logger.log("Source sheet has no data rows. Skipping sync.");
    return;
  }
  
  // Extract IDs from Column E (column 5) formulas of source sheet
  var formulaRange = sourceSheet.getRange(2, 5, lastRow - 1, 1);
  var formulas = formulaRange.getFormulas();
  var workloadIds = [];
  
  for (var i = 0; i < formulas.length; i++) {
    var formula = formulas[i][0];
    var workloadId = "";
    
    if (formula) {
      // Find ID between Workload__c/ (or Workload_c/) and /view
      var match = formula.match(/(?:Workload__c|Workload_c)\/([^\/]+)\/view/);
      if (match && match[1]) {
        workloadId = match[1];
      }
    }
    workloadIds.push(workloadId);
  }
  
  // Now we have the IDs. Let's manage the target sheet.
  // Ensure target sheet has columns AB and AC if not present
  // Column AB is 28, AC is 29
  var targetLastCol = targetSheet.getLastColumn();
  if (targetLastCol < 29) {
    // We need to add columns up to AC
    // Let's just set headers for AB and AC
    targetSheet.getRange(1, 28).setValue("Status");
    targetSheet.getRange(1, 29).setValue("New");
  }
  
  // Also add Workload_ID column at the end if requested
  // Let's place it at Column AD (column 30) or just after the last data column
  // The user said "add a Column at the end with the Workload_ID"
  // Let's assume Column AD is the end if AB and AC are used.
  targetSheet.getRange(1, 30).setValue("Workload_ID");
  
  // Map target data by Workload_ID (which is now in Column AD = 30)
  var targetData = targetSheet.getDataRange().getValues();
  var targetMap = {};
  for (var i = 1; i < targetData.length; i++) {
    var id = targetData[i][29]; // Column AD is index 29 (0-based)
    if (id) {
      targetMap[id] = { row: i + 1, values: targetData[i] };
    }
  }
  
  var sourceData = sourceSheet.getDataRange().getValues();
  
  // Iterate source data and compare
  for (var i = 1; i < sourceData.length; i++) {
    var sourceRow = sourceData[i];
    var id = workloadIds[i - 1]; // Corresponding ID for this row
    
    if (!id) continue;
    
    var targetRowInfo = targetMap[id];
    
    if (!targetRowInfo) {
      // New row
      // We need to construct the full row for target including the ID
      // Source row might have fewer columns than target if target has AB, AC, AD
      var fullRow = sourceRow.slice(); // Copy source values
      // Pad with empty values up to column AC if needed
      while (fullRow.length < 29) {
        fullRow.push("");
      }
      // Set status and new flags in Columns AB and AC
      fullRow[27] = ""; // Status
      fullRow[28] = "Yes"; // New
      fullRow[29] = id; // Workload_ID in Column AD
      
      targetSheet.appendRow(fullRow);
      Logger.log("Added new row with ID: " + id);
    } else {
      // Existing row, compare cells (up to source row length or all?)
      // The user wants to track changes of cell values.
      var targetRowNumber = targetRowInfo.row;
      var targetValues = targetRowInfo.values;
      
      for (var j = 0; j < sourceRow.length; j++) {
        if (sourceRow[j] !== targetValues[j]) {
          // Change detected
          var cell = targetSheet.getRange(targetRowNumber, j + 1);
          cell.setValue(sourceRow[j]);
          cell.setBackground("#FFFFE0"); // Light yellow
          Logger.log("Updated cell (" + targetRowNumber + "," + (j + 1) + ") for ID: " + id);
        }
      }
      
      // Update the Workload_ID in Column AD just in case
      targetSheet.getRange(targetRowNumber, 30).setValue(id);
    }
  }
}


function copyLinkSheet() {
  var sourceSpreadsheetId = "1snf2ryBk7Lizdu5FwTf-LpRc70KN4I-W2-LECgEOJZU";
  var sourceSheetName = "Link";
  
  // Open the source spreadsheet and get the sheet
  var sourceSpreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sourceSheet = sourceSpreadsheet.getSheetByName(sourceSheetName);
  
  if (!sourceSheet) {
    throw new Error("Source sheet not found: " + sourceSheetName);
  }
  
  // Open the target spreadsheet. 
  // Assuming this script is bound to the target spreadsheet.
  var targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  if (!targetSpreadsheet) {
    throw new Error("Target spreadsheet not found. Is this script bound to a spreadsheet?");
  }
  
  // Check if a sheet with the same name already exists in target
  var existingSheet = targetSpreadsheet.getSheetByName(sourceSheetName);
  if (existingSheet) {
    // Delete the existing sheet to replace it
    targetSpreadsheet.deleteSheet(existingSheet);
    Logger.log("Deleted existing sheet in target: " + sourceSheetName);
  }
  
  // Copy the sheet to the target spreadsheet
  var copiedSheet = sourceSheet.copyTo(targetSpreadsheet);
  
  // Rename the copied sheet to the original name
  copiedSheet.setName(sourceSheetName);
  
  Logger.log("Successfully copied sheet to target: " + sourceSheetName);
  
  // Now extract Workload ID and add column
  var lastRow = copiedSheet.getLastRow();
  if (lastRow > 1) { // Ensure there are rows besides the header
    // Column B is column 2
    var formulaRange = copiedSheet.getRange(2, 2, lastRow - 1, 1);
    var formulas = formulaRange.getFormulas();
    var workloadIds = [];
    
    for (var i = 0; i < formulas.length; i++) {
      var formula = formulas[i][0];
      var workloadId = "";
      
      if (formula) {
        // Find ID between Workload__c/ (or Workload_c/) and /view
        var match = formula.match(/(?:Workload__c|Workload_c)\/([^\/]+)\/view/);
        if (match && match[1]) {
          workloadId = match[1];
        }
      }
      workloadIds.push([workloadId]);
    }
    
    // Add new column at column D (column 4)
    copiedSheet.getRange(1, 4).setValue("Workload_ID");
    copiedSheet.getRange(2, 4, workloadIds.length, 1).setValues(workloadIds);
    Logger.log("Extracted and populated Workload_IDs.");
  }
}

