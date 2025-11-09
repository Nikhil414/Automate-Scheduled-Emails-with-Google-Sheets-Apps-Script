
# 🚀 Automate Scheduled Emails with Google Sheets + Apps Script

Tired of manually sending emails at specific times?  
With this guide, you’ll automate everything—no expensive tools needed!



## ✉️ Gmail AppScript Scheduled Email Sender

Automate Gmail drafts and scheduled sending directly from Google Sheets using Apps Script.



## 🌟 Features

- **Automated Gmail Drafts:** Instantly create email drafts from sheet data—no manual copy-paste!
- **Smart Scheduling:** Send emails at custom dates/times using a “Scheduled DateTime” column.
- **Status at a Glance:** See draft status, errors, and sent confirmation right in your sheet.



## 📋 How to Set Up (Step-by-Step)

### 1️⃣ Add “Scheduled DateTime” Column

In your Google Sheet, add a column named `Scheduled DateTime` (usually Column J).  
This is where you set when each email should be sent.

-2025-07-27 20:46:00 

= J2 + TIME(0, CHOOSE(MOD(ROW()-2,3)+1, 2, 3, 2), 0) (Use this Formula for different time and drag it down)

![Sheet Setup Demo](https://github.com/Nikhil414/Automate-Scheduled-Emails-with-Google-Sheets-Apps-Script/blob/b7550e06bc5fd180c765a370f7da3760f7abe53c/image.png)

---

### 2️⃣ Open Google Apps Script

- Go to **Extensions → Apps Script** in your Google Sheet.
- Create two files:
  - `CreateGmailDrafts.gs`
  - `ScheduledSendChecker.gs`

---

### 3️⃣ Copy & Paste the Scripts

#### CreateGmailDrafts.gs

```
function createDraftsWithID() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data;
  for (var i = 1; i < data.length; i++) {
    if (!data[i]) {[1]
      var email = data[i];
      var subject = data[i];[2]
      var body = data[i];[3]
      for (var j = 3; j <= 8; j++) {
        var placeholder = '{{' + headers[j] + '}}';
        subject = subject.replace(new RegExp(placeholder, 'g'), data[i][j]);
        body = body.replace(new RegExp(placeholder, 'g'), data[i][j]);
      }
      var draft = GmailApp.createDraft(email, subject, '', {htmlBody: body});
      sheet.getRange(i + 1, 11).setValue(draft.getId()); // Column K: Draft ID
      sheet.getRange(i + 1, 12).setValue("Scheduled");    // Column L: Status
    }
  }
}
```

#### ScheduledSendChecker.gs

```
function sendScheduledEmails() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var now = new Date();
  for (var i = 1; i < data.length; i++) {
    var scheduledTime = new Date(data[i]);[4]
    var draftId = data[i];[1]
    var status = data[i];[5]
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
```

---

### 4️⃣ Setup Triggers for Automation

- On the left sidebar of Apps Script, click the ⏰ **Triggers** icon.
- Add a new trigger:
  - Set to run `sendScheduledEmails` every **5 or 10 minutes**
- Save—your automailer is live!

---

### 5️⃣ Run & Monitor

- Run `CreateGmailDrafts.gs` first (to make drafts)
- Run `ScheduledSendChecker.gs` once, or let your trigger do the work!
- **Check status in your sheet**—see if emails are “Sent” or flagged for errors

---

## 👀 Always Double-Check!

Regularly glance at the **Status** and **Draft ID** columns to confirm deliveries.






