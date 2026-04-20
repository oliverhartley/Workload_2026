function syncWorkloadsFromScratch() {
  var sourceSpreadsheetId = "1snf2ryBk7Lizdu5FwTf-LpRc70KN4I-W2-LECgEOJZU";
  var targetSpreadsheetId = "1qsB7bD_26sUie6OyW-uyty-r9W3LYTdFjXgIYel3zck";
  
  var sourceSheetName = "Oliver - Worloads Partners";
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
  
  // 1. Read source data and extract IDs from Column E (Link)
  var sourceData = sourceSheet.getDataRange().getValues();
  var sourceHeaders = sourceData[0];
  
  // Find Workload Name column in source (Column D = index 3)
  var nameColIndex = 3; 
  // Find Progress column in source (Column H = index 7)
  var progressColIndex = 7;
  
  var sourceMap = {};
  var sourceRows = [];
  
  var lastRow = sourceSheet.getLastRow();
  
  if (lastRow > 1) {
    for (var i = 1; i < sourceData.length; i++) {
      var sourceRow = sourceData[i];
      var workloadName = sourceRow[nameColIndex];
      var id = "";
      
      // Try to get link from rich text of Column E (column 5)
      var cell = sourceSheet.getRange(i + 1, 5);
      var richText = cell.getRichTextValue();
      var url = "";
      
      if (richText) {
        url = richText.getLinkUrl();
      }
      
      if (url) {
        var match = url.match(/(?:Workload__c|Workload_c)\/([^\/]+)\/view/);
        if (match && match[1]) {
          id = match[1];
        }
      }
      
      if (id) {
        var fullRow = sourceRow.slice();
        fullRow.push(id);
        
        sourceMap[id] = { values: fullRow, progress: sourceRow[progressColIndex] };
        sourceRows.push(fullRow);
        Logger.log("Extracted ID: '" + id + "' for workload: '" + workloadName + "'");
      } else {
        Logger.log("Failed to extract ID for workload: '" + workloadName + "'");
      }
    }
  }
  
  // 3. Prepare target sheet headers if empty
  var targetLastRow = targetSheet.getLastRow();
  
  if (targetLastRow === 0) {
    var targetHeaders = sourceHeaders.slice();
    targetHeaders.push("Workload_ID");
    targetSheet.appendRow(targetHeaders);
    targetLastRow = 1;
  }
  
  // 4. Read target data and map by ID
  var targetData = targetSheet.getDataRange().getValues();
  var targetMap = {};
  
  // Find Workload_ID column index in target. It should be the last one.
  var targetIdColIndex = targetData[0].length - 1;
  
  for (var i = 1; i < targetData.length; i++) {
    var targetRow = targetData[i];
    var id = targetRow[targetIdColIndex];
    if (id) {
      targetMap[id] = { row: i + 1, values: targetRow, progress: targetRow[progressColIndex] };
    }
  }
  
  // 5. Sync and apply rules
  
  // Track processed source IDs to find removed workloads
  var processedSourceIds = {};
  
  // New and Updated rows
  for (var id in sourceMap) {
    processedSourceIds[id] = true;
    var sourceRecord = sourceMap[id];
    var sourceValues = sourceRecord.values;
    var targetRecord = targetMap[id];
    
    if (!targetRecord) {
      // New Workload
      targetSheet.appendRow(sourceValues);
      var lastRow = targetSheet.getLastRow();
      targetSheet.getRange(lastRow, 1, 1, sourceValues.length).setBackground("#E2EFDA"); // Light Green
      Logger.log("Added new row with ID: " + id);
    } else {
      // Existing Workload
      var targetRowNumber = targetRecord.row;
      var targetValues = targetRecord.values;
      
      // Update values
      targetSheet.getRange(targetRowNumber, 1, 1, sourceValues.length).setValues([sourceValues]);
      
      // Check change tracking on progress column
      if (sourceRecord.progress !== targetRecord.progress) {
        // Tracked Change (Yellow)
        targetSheet.getRange(targetRowNumber, 1, 1, sourceValues.length).setBackground("#FFF2CC"); // Yellow
        Logger.log("Updated cell and highlighted yellow for ID: " + id);
      } else {
        // Regular update, clear background
        targetSheet.getRange(targetRowNumber, 1, 1, sourceValues.length).setBackground(null);
      }
    }
  }
  
  // Removed Workloads
  for (var id in targetMap) {
    if (!processedSourceIds[id]) {
      // Removed Workload
      var targetRowNumber = targetMap[id].row;
      var targetValues = targetMap[id].values;
      
      // Highlight entire row with light red background
      targetSheet.getRange(targetRowNumber, 1, 1, targetValues.length).setBackground("#FCE4D6"); // Light Red
      Logger.log("Highlighted removed workload with ID: " + id);
    }
  }
}
