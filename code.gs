/**
 * Google Apps Script for SIMATS Certificate Sharing Portal
 * 
 * Instructions:
 * 1. Open Google Sheets, create a new sheet named "Registrations".
 * 2. In Google Sheets, click Extensions -> Apps Script.
 * 3. Delete any code in the editor and paste this code.
 * 4. IMPORTANT: Select "testEmailAuthorization" from the function dropdown and click Run.
 *    This will trigger the authorization prompt - click "Allow" to grant email permissions.
 * 5. Click "Deploy" -> "New deployment".
 * 6. Select "Web app".
 * 7. Set "Execute as" to "Me", and "Who has access" to "Anyone".
 * 8. Copy the Web App URL and paste it in the Organizer Control Center.
 */

// ============================================================
// RUN THIS FUNCTION FIRST to authorize MailApp permissions!
// Select "testEmailAuthorization" from the dropdown above and click Run.
// ============================================================
function testEmailAuthorization() {
  // This function triggers the OAuth authorization flow for MailApp
  var remaining = MailApp.getRemainingDailyQuota();
  console.log("✅ MailApp is authorized! Remaining daily emails: " + remaining);
  
  // Also test sending a real email to confirm it works
  try {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "SYNORA'26 - Email Authorization Test",
      body: "This is a test email from your SYNORA'26 Certificate Sharing Portal.\n\nIf you received this email, your email authorization is working correctly!\n\nRemaining daily emails: " + remaining
    });
    console.log("✅ Test email sent successfully to: " + Session.getActiveUser().getEmail());
  } catch (err) {
    console.log("❌ Failed to send test email: " + err.toString());
  }
}

function doGet(e) {
  // Safety check to prevent errors when clicking "Run" inside the Apps Script Editor
  if (!e || !e.parameter) {
    return ContentService.createTextOutput("Success: Apps Script is connected! Use the Web App URL inside your HTML portal to send requests.");
  }
  
  const action = e.parameter.action;
  
  if (action === 'get') {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Registrations");
    if (!sheet && ss.getSheets().length > 0) {
      sheet = ss.getSheets()[0]; // Fallback to first sheet
    }
    
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify({ database: [], quota: -1 })).setMimeType(ContentService.MimeType.JSON);
    }
    
    const rows = sheet.getDataRange().getValues();
    const data = [];
    
    // Skip headers (row 0)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const teamName = row[0] ? row[0].toString().trim() : "";
      if (!teamName) continue; // Skip empty rows
      
      data.push({
        teamName: teamName,
        college: row[1] ? row[1].toString().trim() : "",
        leaderName: row[2] ? row[2].toString().trim() : "",
        leaderEmail: row[3] ? row[3].toString().trim() : "",
        leaderPhone: row[4] ? row[4].toString().trim() : "",
        members: row[5] ? row[5].toString().split(',').map(m => m.trim()) : [],
        status: row[7] ? row[7].toString().trim() : "Registered"
      });
    }
    
    let quota = -1;
    try {
      quota = MailApp.getRemainingDailyQuota();
    } catch (qErr) {
      console.log("⚠️ Could not get email quota: " + qErr.toString());
    }

    const payload = JSON.stringify({
      database: data,
      quota: quota
    });
    
    // Support JSONP wrapping to bypass CORS blocks on file:// origin (local runs)
    const callback = e.parameter.callback;
    if (callback) {
      return ContentService.createTextOutput(callback + "(" + payload + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return ContentService.createTextOutput(payload)
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  return ContentService.createTextOutput("Hello from SIMATS Portal Backend");
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      console.log("❌ doPost: No data received");
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No data received." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    console.log("📨 doPost: Received " + e.postData.contents.length + " bytes");
    
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    console.log("📨 doPost: action = " + action);
    
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
      console.log("✅ Registration saved for team: " + reg.teamName);
      
      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Action to send email with certificate attachments automatically and update status
    if (action === 'sendEmail') {
      const email = postData.email;
      const leaderName = postData.leaderName;
      const teamName = postData.teamName;
      const attachmentsData = postData.attachments || [];
      
      console.log("📧 sendEmail: to=" + email + ", team=" + teamName + ", attachments=" + attachmentsData.length);
      
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
        
        console.log("  📎 Attachment " + (i+1) + ": name=" + (fileData.name || "unknown") + ", base64Length=" + rawB64.length);
        
        if (rawB64.length > 0) {
          const decoded = Utilities.base64Decode(rawB64);
          const mime = fileData.mimeType || 'application/pdf';
          const fileName = fileData.name || ('Certificate_' + (i + 1) + '.pdf');
          const blob = Utilities.newBlob(decoded, mime, fileName);
          emailAttachments.push(blob);
          console.log("  ✅ Attachment " + (i+1) + " decoded: " + decoded.length + " bytes");
        } else {
          console.log("  ⚠️ Attachment " + (i+1) + " has EMPTY base64 data!");
        }
      }
      
      console.log("📧 Total valid attachments: " + emailAttachments.length);
      
      // Send email if attachments exist
      let emailSent = false;
      if (emailAttachments.length > 0) {
        const subject = "SYNORA'26 Participation Certificates - Team " + teamName;
        const body = "Dear " + leaderName + ",\n\n" +
                     "Attached are the participation certificates for your team members of SYNORA'26 conducted by SIMATS Engineering.\n\n" +
                     "Kindly distribute them to the respective team members.\n\n" +
                     "Best regards,\n" +
                     "Department of Nanobiomaterials,\n" +
                     "SIMATS Engineering, SIMATS.";
        
        try {
          MailApp.sendEmail({
            to: email,
            subject: subject,
            body: body,
            attachments: emailAttachments
          });
          emailSent = true;
          console.log("✅ EMAIL SENT SUCCESSFULLY to: " + email + " with " + emailAttachments.length + " attachments");
        } catch (mailErr) {
          console.log("❌ EMAIL FAILED: " + mailErr.toString());
          // Return error immediately so the client knows
          return ContentService.createTextOutput(JSON.stringify({ 
            status: "error", 
            message: "Email failed: " + mailErr.toString()
          })).setMimeType(ContentService.MimeType.JSON);
        }
      } else {
        console.log("⚠️ No attachments to send - skipping email");
      }
      
      // Only update sheet status if email was actually sent
      if (emailSent) {
        // Locate the EXACT record in sheet and update Status column H to "Emailed"
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName("Registrations");
        if (!sheet && ss.getSheets().length > 0) {
          sheet = ss.getSheets()[0]; // Fallback to first sheet
        }
        
        if (sheet) {
          const values = sheet.getDataRange().getValues();
          const targetEmail = (email || "").toString().trim().toLowerCase();
          const targetTeam = (teamName || "").toString().trim().toLowerCase();
          
          // Helper to normalize strings for comparison (removes spaces, punctuation, lowercase)
          const normalizeStr = function(str) {
            return str.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
          };
          
          const normTargetTeam = normalizeStr(targetTeam);
          const normTargetEmail = targetEmail.split('@')[0]; // compare local part of email as fallback
          
          let matchedRow = -1;
          
          // Pass 1: Find exact match by BOTH team name AND email (most precise)
          for (let i = 1; i < values.length; i++) {
            const sheetEmail = (values[i][3] || "").toString().trim().toLowerCase();
            const sheetTeam = (values[i][0] || "").toString().trim().toLowerCase();
            
            if (targetTeam && sheetTeam === targetTeam && targetEmail && sheetEmail === targetEmail) {
              matchedRow = i;
              break;
            }
          }
          
          // Pass 2: Normalized Team Name Match (ignoring spaces, punctuation)
          if (matchedRow === -1 && normTargetTeam) {
            for (let i = 1; i < values.length; i++) {
              const sheetTeam = normalizeStr(values[i][0] || "");
              if (sheetTeam === normTargetTeam) {
                matchedRow = i;
                break;
              }
            }
          }
          
          // Pass 3: Email Local Part Match (e.g. "ikuniku35" matches "ikuniku35@gmail.com")
          if (matchedRow === -1 && normTargetEmail) {
            for (let i = 1; i < values.length; i++) {
              const sheetEmail = (values[i][3] || "").toString().trim().toLowerCase();
              const sheetEmailLocal = sheetEmail.split('@')[0];
              if (sheetEmailLocal && sheetEmailLocal === normTargetEmail) {
                matchedRow = i;
                break;
              }
            }
          }
          
          // Pass 4: Fallback exact email match
          if (matchedRow === -1 && targetEmail) {
            for (let i = 1; i < values.length; i++) {
              const sheetEmail = (values[i][3] || "").toString().trim().toLowerCase();
              if (sheetEmail === targetEmail) {
                matchedRow = i;
                break;
              }
            }
          }

          // Pass 5: Substring matching (fuzzy match)
          if (matchedRow === -1 && normTargetTeam) {
            for (let i = 1; i < values.length; i++) {
              const sheetTeam = normalizeStr(values[i][0] || "");
              if (sheetTeam && (sheetTeam.indexOf(normTargetTeam) !== -1 || normTargetTeam.indexOf(sheetTeam) !== -1)) {
                matchedRow = i;
                break;
              }
            }
          }
          
          if (matchedRow !== -1) {
            sheet.getRange(matchedRow + 1, 8).setValue("Emailed");
            SpreadsheetApp.flush();
            console.log("✅ Sheet status updated to 'Emailed' for row " + (matchedRow + 1));
          } else {
            console.log("⚠️ Could not find matching row for team: " + teamName);
          }
        }
      }
      
      let quota = -1;
      try {
        quota = MailApp.getRemainingDailyQuota();
      } catch (qErr) {}

      return ContentService.createTextOutput(JSON.stringify({ 
        status: emailSent ? "success" : "error",
        message: emailSent ? "Email sent" : "No attachments received",
        quota: quota
      })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    console.log("❌ FATAL ERROR in doPost: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
