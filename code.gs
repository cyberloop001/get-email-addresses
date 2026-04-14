function getSentEmailsWithDateFromFeb1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["Email Address", "Sent Date"]);

  // February 1 (current year)
  const startDate = new Date(new Date().getFullYear(), 1, 1);

  const query = `in:sent after:${formatDate(startDate)}`;
  const threads = GmailApp.search(query);

  const rows = [];

  threads.forEach(thread => {
    const messages = thread.getMessages();

    messages.forEach(message => {
      const date = message.getDate();

      processField(message.getTo(), date, rows);
      processField(message.getCc(), date, rows);
      processField(message.getBcc(), date, rows);
    });
  });

  // Write all rows at once (faster)
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }

  Logger.log(`Total rows: ${rows.length}`);
}

// Extract emails + attach date
function processField(field, date, rows) {
  if (!field) return;

  const matches = field.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/g);
  if (matches) {
    matches.forEach(email => {
      rows.push([email.toLowerCase(), date]);
    });
  }
}

// Format date for Gmail query
function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}
