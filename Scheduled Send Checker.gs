function sendScheduledEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < data.length; i++) {
    var scheduledTime = new Date(data[i][9]); // Column J
    var draftId = data[i][10];
    var status = data[i][11];

    if (status === "Scheduled" && scheduledTime <= now) {
      try {
        var draft = GmailApp.getDraft(draftId);
        var message = draft.send();
        sheet.getRange(i + 1, 12).setValue("Sent");
      } catch (err) {
        sheet.getRange(i + 1, 12).setValue("Error: " + err.message);
      }
    }
  }
}
