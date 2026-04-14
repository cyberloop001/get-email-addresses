function getSentEmailAddressesFromFeb1() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["Email Address"]);

  // Set start date: February 1 (current year)
  const startDate = new Date(new Date().getFullYear(), 1, 1); // Month is 0-based → 1 = Feb

  const query = `in:sent after:${formatDate(startDate)}`;
  const threads = GmailApp.search(query);

  const emailSet = new Set();

  threads.forEach(thread => {
    const messages = thread.getMessages();
    messages.forEach(message => {
      extractEmails(message.getTo(), emailSet);
      extractEmails(message.getCc(), emailSet);
      extractEmails(message.getBcc(), emailSet);
    });
  });

  const emails = Array.from(emailSet).sort();
  emails.forEach(email => sheet.appendRow([email]));

  Logger.log(`Total unique emails: ${emails.length}`);
}

function extractEmails(field, emailSet) {
  if (!field) return;

  const matches = field.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/g);
  if (matches) {
    matches.forEach(email => emailSet.add(email.toLowerCase()));
  }
}

function formatDate(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}
