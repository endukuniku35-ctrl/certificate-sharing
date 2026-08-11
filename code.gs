/**
 * Google Apps Script for SIMATS Certificate Sharing Portal
 * 
 * Instructions:
 * 1. Open Google Sheets, create a new sheet named "Registrations".
 * 2. In Google Sheets, click Extensions -> Apps Script.
 * 3. Delete any code in the editor and paste this code.
 * 4. Click "Deploy" -> "New deployment".
 * 5. Select "Web app".
 * 6. Set "Execute as" to "Me", and "Who has access" to "Anyone".
 * 7. Copy the Web App URL and paste it in the Organizer Control Center.
 */

function doGet(e) {
  // Safety check to prevent errors when clicking "Run" inside the Apps Script Editor
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Success: Apps Script is connected! Use the Web App URL inside your HTML portal to send requests.");
  }
  
  const action = e.parameter.action;
  
  if (action === 'get') {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registrations");
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }
    
    const rows = sheet.getDataRange().getValues();
    const data = [];
    
    // Skip headers (row 0)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      data.push({
        teamName: row[0],
        college: row[1],
        leaderName: row[2],
        leaderEmail: row[3],
        leaderPhone: row[4],
        members: row[5] ? row[5].split(',').map(m => m.trim()) : [],
        status: row[7] || "Registered"
      });
    }
    
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("Hello from SIMATS Portal Backend");
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No data received." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    
    if (action === 'register') {
      const reg = postData.data;
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheet = ss.getSheetByName("Registrations");
      
      // Create sheet with headers if it doesn't exist
      if (!sheet) {
        sheet = ss.insertSheet("Registrations");
        sheet.appendRow(["Team Name", "College", "Leader Name", "Leader Email", "Leader Phone", "Members", "Timestamp", "Status"]);
      }
      
      const membersStr = Array.isArray(reg.members) ? reg.members.join(', ') : reg.members;
      sheet.appendRow([
        reg.teamName,
        reg.college,
        reg.leaderName,
        reg.leaderEmail,
        reg.leaderPhone,
        membersStr,
        new Date().toISOString(),
        "Registered"
      ]);
      SpreadsheetApp.flush();
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Action to send email with certificate attachments automatically and update status
    if (action === 'sendEmail') {
      const email = postData.email;
      const leaderName = postData.leaderName;
      const teamName = postData.teamName;
      const attachmentsData = postData.attachments || [];
      
      const emailAttachments = [];
      for (let i = 0; i < attachmentsData.length; i++) {
        const fileData = attachmentsData[i];
        let rawB64 = fileData.base64Data || "";
        
        // Remove data URI scheme prefix if present
        if (rawB64.indexOf(",") !== -1) {
          rawB64 = rawB64.split(",")[1];
        }
        
        // Clean whitespace and linebreaks
        rawB64 = rawB64.replace(/\s+/g, '');
        
        if (rawB64.length > 0) {
          const decoded = Utilities.base64Decode(rawB64);
          const mime = fileData.mimeType || 'application/pdf';
          const fileName = fileData.name || ('Certificate_' + (i + 1) + '.pdf');
          const blob = Utilities.newBlob(decoded, mime, fileName);
          emailAttachments.push(blob);
        }
      }
      
      // Send email if attachments exist
      if (emailAttachments.length > 0) {
        const subject = "SYNORA'26 Participation Certificates - Team " + teamName;
        const body = "Dear " + leaderName + ",\n\n" +
                     "Attached are the participation certificates for your team members of SYNORA'26 conducted by SIMATS Engineering.\n\n" +
                     "Kindly distribute them to the respective team members.\n\n" +
                     "Best regards,\n" +
                     "Department of Nanobiomaterials,\n" +
                     "SIMATS Engineering, SIMATS.";
                     
        MailApp.sendEmail({
          to: email,
          subject: subject,
          body: body,
          attachments: emailAttachments
        });
      }
      
      // Locate the record in sheet and update Status column to "Emailed" with normalized matching
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName("Registrations");
      if (sheet) {
        const values = sheet.getDataRange().getValues();
        const targetEmail = (email || "").toString().trim().toLowerCase();
        const targetTeam = (teamName || "").toString().trim().toLowerCase();
        
        for (let i = 1; i < values.length; i++) {
          const sheetEmail = (values[i][3] || "").toString().trim().toLowerCase();
          const sheetTeam = (values[i][0] || "").toString().trim().toLowerCase();
          
          if ((targetEmail && sheetEmail === targetEmail) || (targetTeam && sheetTeam === targetTeam)) {
            sheet.getRange(i + 1, 8).setValue("Emailed");
            SpreadsheetApp.flush(); // Force immediate commit to disk
            break;
          }
        }
      }
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
