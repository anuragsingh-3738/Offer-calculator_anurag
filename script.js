// Handles data fetch for product suggestions
function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var data = sheet.getDataRange().getValues();
    var products = [];
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][1]) {
        products.push({
          model: data[i][0].toString(),
          price: parseFloat(data[i][1])
        });
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify(products))
                         .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// Handles unique record storage for customer responses
function doPost(e) {
  try {
    var params = JSON.parse(e.postData.contents);
    var name = params.name;
    var mobile = params.mobile;
    
    if (!mobile || !name) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Missing fields" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("CustomerResponses");
    
    // Create the responses sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet("CustomerResponses");
    }
    
    // Automatically set headers if the sheet is completely blank
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Mobile Number", "Customer Name", "Last Logged"]);
    }
    
    var data = sheet.getDataRange().getValues();
    var isUnique = true;
    
    // Check if the mobile number is already in the database
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim() === mobile.toString().trim()) {
        isUnique = false;
        // Keep it clean and updated: write over name and change date stamp
        sheet.getRange(i + 1, 2).setValue(name);
        sheet.getRange(i + 1, 3).setValue(new Date());
        break;
      }
    }
    
    // If it's a completely new customer mobile number, add a row
    if (isUnique) {
      sheet.appendRow([mobile, name, new Date()]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", unique: isUnique }))
                         .setMimeType(ContentService.MimeType.JSON);
                         
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
