/**
 * Google Apps Script for Restaurant Feedback
 * 
 * INSTRUCTIONS:
 * 1. Create a new Google Sheet.
 * 2. Go to Extensions > Apps Script.
 * 3. Paste this code into Code.gs.
 * 4. Run the 'setup' function once to create headers.
 * 5. Deploy as Web App:
 *    - Click 'Deploy' > 'New deployment'.
 *    - Select type: 'Web app'.
 *    - Description: 'v1'.
 *    - Execute as: 'Me'.
 *    - Who has access: 'Anyone'.
 *    - Click 'Deploy'.
 * 6. Copy the Web App URL and paste it into index.html as SUBMISSION_ENDPOINT.
 */

const POJ_FEEDBACK_SHEET_NAME = "Feedback";
const POJ_FEEDBACK_FOLDER_NAME = "Feedback_Images";

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    // OPEN BY ID to ensure we hit the right sheet
    const doc = SpreadsheetApp.openById("1oCo-B7vc4FOI7w3qykNLN6Z1Z5X0ag7MIY3_Y40K2UU");
    let sheet = doc.getSheetByName(POJ_FEEDBACK_SHEET_NAME);

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = doc.insertSheet(POJ_FEEDBACK_SHEET_NAME);
      const headers = [
        "Timestamp", 
        "Code", 
        "Service Type", 
        "Customer Name", 
        "Order ID", 
        "Overall Rating", 
        "Food Rating", 
        "Service Rating", 
        "Speed Rating", 
        "Cleanliness Rating", 
        "Comments", 
        "Photo URL"
      ];
      sheet.appendRow(headers);
    }

    const data = JSON.parse(e.postData.contents);
    let fileUrl = "";

    // Handle Image Upload
    if (data.photoDataUrl) {
      try {
        const folder = getOrCreateFolder(POJ_FEEDBACK_FOLDER_NAME);
        const contentType = data.photoDataUrl.substring(5, data.photoDataUrl.indexOf(';'));
        const base64Data = data.photoDataUrl.substring(data.photoDataUrl.indexOf(',') + 1);
        const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), contentType, "feedback_" + data.code + ".jpg");
        const file = folder.createFile(blob);
        fileUrl = file.getUrl();
      } catch (err) {
        fileUrl = "Error saving image: " + err.toString();
      }
    }

    // Append to Sheet
    const nextRow = sheet.getLastRow() + 1;
    const newRow = [
      new Date(), // Timestamp
      data.code,
      data.serviceType,
      data.customerName,
      data.orderId,
      data.overallRating,
      data.foodRating,
      data.serviceRating,
      data.speedRating,
      data.cleanlinessRating || "", // Handle missing field gracefully
      data.comments,
      fileUrl
    ];

    sheet.appendRow(newRow);

    return ContentService
      .createTextOutput(JSON.stringify({ "result": "success", "row": nextRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (e) {
    return ContentService
      .createTextOutput(JSON.stringify({ "result": "error", "error": e.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

function setup() {
  const doc = SpreadsheetApp.openById("1oCo-B7vc4FOI7w3qykNLN6Z1Z5X0ag7MIY3_Y40K2UU");
  let sheet = doc.getSheetByName(POJ_FEEDBACK_SHEET_NAME);
  if (!sheet) {
    sheet = doc.insertSheet(POJ_FEEDBACK_SHEET_NAME);
    const headers = [
        "Timestamp", 
        "Code", 
        "Service Type", 
        "Customer Name", 
        "Order ID", 
        "Overall Rating", 
        "Food Rating", 
        "Service Rating", 
        "Speed Rating", 
        "Cleanliness Rating", 
        "Comments", 
        "Photo URL"
    ];
    sheet.appendRow(headers);
  }
}
