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
  
  // 1. Read source data and extract IDs from Column E (Link)
  var sourceData = sourceSheet.getDataRange().getValues();
  var sourceHeaders = sourceData[0];
  
  // Find Workload Name column in source (Column D = index 3)
  var nameColIndex = 3; 
  // Find Progress column in source (Column H = index 7)
  var progressColIndex = 7;
  // Find Production Date column in source (Column S = index 18)
  var prodDateColIndex = 18;
  
  var sourceMap = {};
  
  var lastRow = sourceSheet.getLastRow();
  
  if (lastRow > 1) {
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Reset time for accurate day calculation
    
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
        fullRow.push(url); // Append raw URL
        
        // Calculate Days to Production
        var prodDateVal = sourceRow[prodDateColIndex];
        var daysToProd = "";
        
        if (prodDateVal) {
          var prodDate = new Date(prodDateVal);
          if (!isNaN(prodDate.getTime())) {
            prodDate.setHours(0, 0, 0, 0);
            var timeDiff = prodDate.getTime() - today.getTime();
            daysToProd = Math.ceil(timeDiff / (1000 * 3600 * 24));
          }
        }
        fullRow.push(daysToProd);
        
        sourceMap[id] = { values: fullRow, progress: sourceRow[progressColIndex] };
        Logger.log("Source Map Add: '" + id + "' for workload: '" + workloadName + "'");
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
    targetHeaders.push("Workload_Link");
    targetHeaders.push("Days to Production");
    targetHeaders.push("Pendiente");
    targetHeaders.push("Comentario");
    targetHeaders.push("ER-Gemini");
    targetHeaders.push("ER");
    targetSheet.appendRow(targetHeaders);
    targetLastRow = 1;
  }
  
  // 4. Read target data and map by ID
  var targetData = targetSheet.getDataRange().getValues();
  var targetHeaders = targetData[0];
  
  var percepcionColIndex = targetHeaders.indexOf("Percepcion del Partner");
  if (percepcionColIndex === -1) {
    targetSheet.getRange(1, 38).setValue("Percepcion del Partner");
    Logger.log("Added column 'Percepcion del Partner' at column AL");
    // Re-read headers
    targetHeaders = targetSheet.getDataRange().getValues()[0];
  }
  
  var targetMap = {};
  
  // Find column indices by header name
  var targetIdColIndex = targetHeaders.indexOf("Workload_ID");
  var statusColIndex = targetHeaders.indexOf("Pendiente");
  var daysToProdColIndex = targetHeaders.indexOf("Days to Production");
  
  Logger.log("Target ID Column Index: " + targetIdColIndex);
  Logger.log("Target Status Column Index: " + statusColIndex);
  
  if (targetIdColIndex === -1) {
    throw new Error("Workload_ID column not found in target sheet headers.");
  }
  
  for (var i = 1; i < targetData.length; i++) {
    var targetRow = targetData[i];
    var id = targetRow[targetIdColIndex];
    if (id) {
      targetMap[id] = { row: i + 1, values: targetRow, progress: targetRow[progressColIndex] };
      Logger.log("Target Map Add: '" + id + "' at row " + (i + 1));
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
    
    // Construct the full row for target including status and comments placeholders
    var targetRowToWrite = sourceValues.slice();
    while (targetRowToWrite.length < targetHeaders.length) {
      targetRowToWrite.push(""); // Pad with empty strings for new columns
    }
    
    if (!targetRecord) {
      // New Workload
      if (statusColIndex !== -1) targetRowToWrite[statusColIndex] = "New";
      targetSheet.appendRow(targetRowToWrite);
      var lastRow = targetSheet.getLastRow();
      targetSheet.getRange(lastRow, 1, 1, targetRowToWrite.length).setBackground("#E2EFDA"); // Light Green
      Logger.log("Decision: NEW for ID: " + id + ". Appended at row " + lastRow);
    } else {
      // Existing Workload
      var targetRowNumber = targetRecord.row;
      var targetValues = targetRecord.values;
      
      // Preserve existing comments if they exist in target
      var commentIndex = targetHeaders.indexOf("Comentario");
      var commentErIndex = targetHeaders.indexOf("ER-Gemini");
      var erIndex = targetHeaders.indexOf("ER");
      var percepcionIndex = targetHeaders.indexOf("Percepcion del Partner");
      
      if (commentIndex !== -1) targetRowToWrite[commentIndex] = targetValues[commentIndex];
      if (commentErIndex !== -1) targetRowToWrite[commentErIndex] = targetValues[commentErIndex];
      if (erIndex !== -1) targetRowToWrite[erIndex] = targetValues[erIndex];
      if (percepcionIndex !== -1) targetRowToWrite[percepcionIndex] = targetValues[percepcionIndex];
      
      // Check change tracking on progress column
      if (sourceRecord.progress !== targetRecord.progress) {
        // Tracked Change (Yellow)
        if (statusColIndex !== -1) targetRowToWrite[statusColIndex] = "Updated";
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setValues([targetRowToWrite]);
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setBackground("#FFF2CC"); // Yellow
        Logger.log("Decision: UPDATE (Yellow) for ID: " + id + " at row " + targetRowNumber);
      } else {
        // Regular update, clear background
        if (statusColIndex !== -1) targetRowToWrite[statusColIndex] = targetValues[statusColIndex]; // Preserve status
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setValues([targetRowToWrite]);
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setBackground(null);
        Logger.log("Decision: UPDATE (Clear) for ID: " + id + " at row " + targetRowNumber);
      }
    }
  }
  
  // Removed Workloads
  for (var id in targetMap) {
    if (!processedSourceIds[id]) {
      // Removed Workload
      var targetRowNumber = targetMap[id].row;
      var targetValues = targetMap[id].values;
      
      // Update status to Removed
      if (statusColIndex !== -1) {
        targetSheet.getRange(targetRowNumber, statusColIndex + 1).setValue("Removed");
      }
      
      // Highlight entire row with light red background
      targetSheet.getRange(targetRowNumber, 1, 1, targetValues.length).setBackground("#FCE4D6"); // Light Red
      Logger.log("Decision: REMOVED (Red) for ID: " + id + " at row " + targetRowNumber);
    }
  }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sync Menu')
      .addItem('Sync Workloads', 'syncWorkloadsFromScratch')
      .addItem('Sync Expert Requests', 'syncExpertRequests')
      .addItem('Send Workload Emails', 'sendWorkloadEmails')
      .addToUi();
}
