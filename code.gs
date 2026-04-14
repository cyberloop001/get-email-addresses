function getSentEmailsUniqueLatestDateFromFeb1_noSelf() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(["Email Address", "Latest Sent Date"]);

  const startDate = new Date(new Date().getFullYear(), 1, 1);

  const query = "in:sent";
  const batchSize = 100;
  let start = 0;

  const emailMap = new Map();

  // 🔴 PUT YOUR EMAIL HERE
  const myEmail = Session.getActiveUser().getEmail().toLowerCase();

  while (true) {
    const threads = GmailApp.search(query, start, batchSize);
    if (threads.length === 0) break;

    threads.forEach(thread => {
      const messages = thread.getMessages();

      messages.forEach(msg => {
        const msgDate = msg.getDate();

        if (msgDate < startDate) return;

        process(msg.getTo(), msgDate, emailMap, myEmail);
        process(msg.getCc(), msgDate, emailMap, myEmail);
        process(msg.getBcc(), msgDate, emailMap, myEmail);
      });
    });

    start += batchSize;
  }

  const rows = Array.from(emailMap.entries())
    .map(([email, date]) => [email, date])
    .sort((a, b) => new Date(b[1]) - new Date(a[1]));

  sheet.getRange(2, 1, rows.length, 2).setValues(rows);

  Logger.log("Unique external emails: " + rows.length);
}

// 🚫 exclude self email
function process(field, date, emailMap, myEmail) {
  if (!field) return;

  const matches = field.match(/[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}/g);
  if (!matches) return;

  matches.forEach(email => {
    email = email.toLowerCase();

    // 🔴 skip your own email
    if (email === myEmail) return;

    if (!emailMap.has(email)) {
      emailMap.set(email, date);
    } else {
      const existingDate = emailMap.get(email);
      if (date > existingDate) {
        emailMap.set(email, date);
      }
    }
  });
}
