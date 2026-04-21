function syncExpertRequests() {
  var sourceFolderId = "1RJIoDUxiz3mNPXLO6XDaP8DWXrNIXP_B";
  var targetSpreadsheetId = "1qsB7bD_26sUie6OyW-uyty-r9W3LYTdFjXgIYel3zck";
  var targetSheetName = "Expert Request";
  
  var folder = DriveApp.getFolderById(sourceFolderId);
  var files = folder.getFiles();
  var newestFile = null;
  var latestTime = 0;
  
  while (files.hasNext()) {
    var file = files.next();
    var time = file.getDateCreated().getTime();
    if (time > latestTime) {
      latestTime = time;
      newestFile = file;
    }
  }
  
  if (!newestFile) {
    Logger.log("No files found in folder.");
    return;
  }
  
  Logger.log("Newest file: " + newestFile.getName() + " (" + newestFile.getDateCreated() + ")");
  
  var sourceData = [];
  var mimeType = newestFile.getMimeType();
  Logger.log("File MimeType: " + mimeType);
  
  if (mimeType === "application/vnd.google-apps.spreadsheet") {
    var ss = SpreadsheetApp.open(newestFile);
    var sheet = ss.getSheets()[0]; // Assume first sheet
    sourceData = sheet.getDataRange().getValues();
  } else {
    try {
      var csvContent = newestFile.getBlob().getDataAsString();
      sourceData = Utilities.parseCsv(csvContent);
    } catch (e) {
      Logger.log("Error parsing CSV: " + e.message);
      if (csvContent) {
        Logger.log("File content preview: " + csvContent.substring(0, 100));
      }
      return;
    }
  }
  
  if (sourceData.length === 0) {
    Logger.log("Source data is empty.");
    return;
  }
  
  var sourceHeaders = sourceData[0];
  var idColIndex = -1;
  var statusColIndex = -1;
  
  for (var i = 0; i < sourceHeaders.length; i++) {
    var header = sourceHeaders[i].trim();
    if (header === "Expert Request: ID") {
      idColIndex = i;
    } else if (header === "Status") {
      statusColIndex = i;
    }
  }
  
  if (idColIndex === -1) {
    Logger.log("Error: 'Expert Request: ID' column not found in CSV.");
    Logger.log("Found headers: " + JSON.stringify(sourceHeaders));
    return;
  }
  
  var targetSpreadsheet = SpreadsheetApp.openById(targetSpreadsheetId);
  var targetSheet = targetSpreadsheet.getSheetByName(targetSheetName);
  
  if (!targetSheet) {
    targetSheet = targetSpreadsheet.insertSheet(targetSheetName);
    Logger.log("Created target sheet: " + targetSheetName);
  }
  
  // Map source data by ID
  var sourceMap = {};
  for (var i = 1; i < sourceData.length; i++) {
    var row = sourceData[i];
    var id = row[idColIndex];
    if (id) {
      sourceMap[id] = { values: row, status: statusColIndex !== -1 ? row[statusColIndex] : "" };
    }
  }
  
  // Setup headers if table is new (starting at row 3)
  var targetLastRow = targetSheet.getLastRow();
  
  if (targetLastRow < 3) {
    targetSheet.getRange(3, 1, 1, sourceHeaders.length).setValues([sourceHeaders]);
    targetLastRow = 3;
  }
  
  // Read target data starting from row 3 (headers)
  var targetDataRange = targetSheet.getRange(3, 1, targetSheet.getLastRow() - 2, targetSheet.getLastColumn());
  var targetData = targetDataRange.getValues();
  var targetHeaders = targetData[0];
  var targetMap = {};
  
  var targetIdColIndex = targetHeaders.indexOf("Expert Request: ID");
  var sourceHeaderCount = sourceHeaders.length;
  
  if (targetIdColIndex === -1) {
    Logger.log("Error: 'Expert Request: ID' column not found in target sheet.");
    return;
  }
  
  for (var i = 1; i < targetData.length; i++) {
    var row = targetData[i];
    var id = row[targetIdColIndex];
    if (id) {
      targetMap[id] = { row: i + 3, values: row, status: statusColIndex !== -1 ? row[statusColIndex] : "" };
    }
  }
  
  var processedSourceIds = {};
  
  // Sync
  for (var id in sourceMap) {
    processedSourceIds[id] = true;
    var sourceRecord = sourceMap[id];
    var sourceValues = sourceRecord.values;
    var targetRecord = targetMap[id];
    
    var targetRowToWrite = sourceValues.slice();
    
    if (targetRecord) {
      var targetValues = targetRecord.values;
      for (var j = sourceHeaderCount; j < targetValues.length; j++) {
        targetRowToWrite.push(targetValues[j]);
      }
    }
    
    if (!targetRecord) {
      // New Row
      while (targetRowToWrite.length < targetHeaders.length) {
        targetRowToWrite.push("");
      }
      
      targetSheet.getRange(targetLastRow + 1, 1, 1, targetRowToWrite.length).setValues([targetRowToWrite]);
      targetSheet.getRange(targetLastRow + 1, 1, 1, targetRowToWrite.length).setBackground("#E2EFDA"); // Green
      targetLastRow++;
      Logger.log("NEW row added for ID: " + id);
    } else {
      // Existing Row
      var targetRowNumber = targetRecord.row;
      
      while (targetRowToWrite.length < targetHeaders.length) {
        targetRowToWrite.push("");
      }
      
      if (sourceRecord.status !== targetRecord.status) {
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setValues([targetRowToWrite]);
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setBackground("#FFF2CC"); // Yellow
        Logger.log("UPDATE (Yellow) for ID: " + id);
      } else {
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setValues([targetRowToWrite]);
        targetSheet.getRange(targetRowNumber, 1, 1, targetRowToWrite.length).setBackground(null);
        Logger.log("UPDATE (Clear) for ID: " + id);
      }
    }
  }
  
  // Removed Rows
  for (var id in targetMap) {
    if (!processedSourceIds[id]) {
      var targetRecord = targetMap[id];
      var targetRowNumber = targetRecord.row;
      targetSheet.getRange(targetRowNumber, 1, 1, targetHeaders.length).setBackground("#FCE4D6"); // Red
      Logger.log("REMOVED for ID: " + id);
    }
  }
  
  // Update Cell A1
  var a1Range = targetSheet.getRange("A1");
  var now = new Date();
  a1Range.setValue("Last updated: " + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));
  a1Range.setFontWeight("bold");
  a1Range.setFontColor("red");
}
