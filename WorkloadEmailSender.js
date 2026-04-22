function sendWorkloadEmails() {
  var sourceSpreadsheetId = "1snf2ryBk7Lizdu5FwTf-LpRc70KN4I-W2-LECgEOJZU";
  var sheetName = "Oliver - Workloads Partners";
  
  var spreadsheet = SpreadsheetApp.openById(sourceSpreadsheetId);
  var sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found.");
  }
  
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Column indices (0-based)
  var partnerIndex = 0; // Column A
  var nameIndex = 3;    // Column D
  var progressIndex = 7; // Column H
  var ownerIndex = 14;   // Column O
  var prodDateIndex = 18; // Column S
  var linkIndex = 29;   // Column AD
  var daysToProdIndex = 30; // Column AE
  
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  
  var workloadsByOwner = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var progress = row[progressIndex];
    
    // Filter for test
    if (progress === "4.1: Pre Ramp / Mobilize") {
      var owner = row[ownerIndex];
      var partner = row[partnerIndex];
      var name = row[nameIndex];
      var prodDateVal = row[prodDateIndex];
      var link = row[linkIndex];
      
      if (!owner) {
        Logger.log("Skipping row " + (i + 1) + " because owner is empty.");
        continue;
      }
      
      // Calculate Days to Production if not provided or recalculate to be sure
      var daysToProd = "";
      if (prodDateVal) {
        var prodDate = new Date(prodDateVal);
        if (!isNaN(prodDate.getTime())) {
          prodDate.setHours(0, 0, 0, 0);
          var timeDiff = prodDate.getTime() - today.getTime();
          daysToProd = Math.ceil(timeDiff / (1000 * 3600 * 24));
        }
      }
      
      // Fallback to spreadsheet value if calculation fails or is empty
      if (daysToProd === "" && row[daysToProdIndex] !== "") {
        daysToProd = row[daysToProdIndex];
      }
      
      // Fallback to link extraction if raw value is not a link (as in Code.js)
      if (!link && sheet.getRange(i + 1, linkIndex + 1).getRichTextValue()) {
         var richText = sheet.getRange(i + 1, linkIndex + 1).getRichTextValue();
         link = richText.getLinkUrl();
      }
      
      var workloadInfo = {
        name: name,
        partner: partner,
        prodDate: prodDateVal,
        daysToProd: daysToProd,
        progress: progress,
        link: link || "No link provided"
      };
      
      if (!workloadsByOwner[owner]) {
        workloadsByOwner[owner] = [];
      }
      workloadsByOwner[owner].push(workloadInfo);
    }
  }
  
  // Send emails
  for (var owner in workloadsByOwner) {
    var workloads = workloadsByOwner[owner];
    var recipient = "oliverhartley@google.com"; // FOR TESTING: Override recipient
    
    // In production, it would be:
    // var recipient = owner + "@google.com";
    
    var bodyHtml = "<p>Hola,</p>" +
               "<p>Te escribo para saber si los workloads de GCP van por buen camino o si hay algún retraso.</p>" +
               "<p>En caso de haber algún bloqueo, ¿me podrías confirmar si es por parte del partner o del cliente?</p>" +
               "<p>Por último, avísame si hay alguna acción técnica que pueda realizar desde mi lado para ayudar al partner a que las cosas avancen más rápido.</p>" +
               "<br/>" +
               "<table style='border-collapse: collapse; width: 100%; font-family: Arial, sans-serif;'>" +
               "<thead>" +
               "<tr style='background-color: #4285F4; color: white;'>" +
               "<th style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>Workload</th>" +
               "<th style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>Partner</th>" +
               "<th style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>Progress</th>" +
               "<th style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>Days to Production</th>" +
               "<th style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>Link</th>" +
               "</tr>" +
               "</thead>" +
               "<tbody>";
               
    workloads.forEach(function(wl) {
      bodyHtml += "<tr>" +
                  "<td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'><b>" + wl.name + "</b></td>" +
                  "<td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>" + wl.partner + "</td>" +
                  "<td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>" + wl.progress + "</td>" +
                  "<td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'>" + wl.daysToProd + "</td>" +
                  "<td style='border: 1px solid #dddddd; text-align: left; padding: 8px;'><a href='" + wl.link + "'>Link</a></td>" +
                  "</tr>";
    });
    
    bodyHtml += "</tbody></table>";
    
    MailApp.sendEmail({
      to: recipient,
      subject: "Actualización de Workloads de GCP - Owner: " + owner,
      htmlBody: bodyHtml
    });
    
    Logger.log("Sent email for owner: " + owner + " to " + recipient);
  }
}
