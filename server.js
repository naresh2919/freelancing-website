
function doPost(e) {
  try {
    const sheet = SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName("Sheet1");

    const data = JSON.parse(e.postData.contents);

    const name = String(data.name || "").trim();
    const email = String(data.email || "").trim();
    const phone = String(data.phone || "").trim();
    const details = String(data.details || "").trim();

    if (!name || !email || !phone || !details) {
      throw new Error("All fields are required");
    }

    sheet.appendRow([
      new Date(),
      name,
      email,
      phone,
      details
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        message: "Message saved successfully"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error.message
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
