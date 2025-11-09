function createDraftsWithID() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];

  for (var i = 1; i < data.length; i++) {
    if (!data[i][10]) {
      // Check if Draft ID is already filled
      var email = data[i][0];
      var subject = data[i][1];
      var body = data[i][2];

      for (var j = 3; j <= 8; j++) {
        var placeholder = "{{" + headers[j] + "}}";
        subject = subject.replace(new RegExp(placeholder, "g"), data[i][j]);
        body = body.replace(new RegExp(placeholder, "g"), data[i][j]);
      }

      var draft = GmailApp.createDraft(email, subject, "", { htmlBody: body });
      sheet.getRange(i + 1, 11).setValue(draft.getId()); // Column K: Draft ID
      sheet.getRange(i + 1, 12).setValue("Scheduled"); // Column L: Status
    }
  }
}
