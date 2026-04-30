function doPostCustom(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const payload = {
      gpt: data.gpt || '',
      founder_name: data.founder_name || 'Anonymous',
      date: data.date || new Date().toISOString().split('T')[0],
      status: data.status || '',
      customer_profile: data.customer_profile || '',
      needs_answers_chart: data.needs_answers_chart || '',
      custom_statement: data.custom_statement || '',
      tough_spots: data.tough_spots || '',
      next_gpt: data.next_gpt || ''
    };

    writeToSheetCustom(payload);
    if (EMAIL_JOEL && JOEL_EMAIL !== 'email') emailJoelCustom(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Handoff recorded via POST' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeToSheetCustom(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('d.BOBO Custom');

  if (!sheet) {
    sheet = ss.insertSheet('d.BOBO Custom');
    sheet.appendRow(['Timestamp', 'GPT', 'Founder Name', 'Status', 'Customer \nProfile', 'Needs Answers', 'Custom \nStatement', 'Tough Spots', 'Next GPT']);
    const headerRange = sheet.getRange(1, 1, 1, 9);
    headerRange.setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FFFFFF');
  }

  sheet.appendRow([
    new Date(),
    payload.gpt,
    payload.founder_name,
    payload.status,
    payload.customer_profile,
    payload.needs_answers_chart,
    payload.custom_statement,
    payload.tough_spots,
    payload.next_gpt
  ]);
}

function emailJoelCustom(payload) {
  const subject = `d.BOBO Handoff — ${payload.founder_name} (GPT ${payload.gpt})`;
  const body = `New d.BOBO handoff summary:

FOUNDER: ${payload.founder_name}
DATE: ${payload.date}
GPT: ${payload.gpt}
STATUS: ${payload.status}

CUSTOMER PROFILE:
${payload.customer_profile || '(not provided)'}

NEEDS ANSWERS:
${payload.needs_answers_chart || '(not provided)'}

CUSTOM STATEMENT:
${payload.custom_statement || '(not provided)'}

TOUGH SPOTS / OPEN QUESTIONS:
${payload.tough_spots || 'None flagged.'}

READY FOR:
${payload.next_gpt}

---
View full spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  MailApp.sendEmail(JOEL_EMAIL, subject, body);
}

function testHandoffCustom() {
  const testPayload = {
    gpt: '2-custom',
    founder_name: 'Test Founder',
    date: '2026-04-26',
    status: 'complete',
    customer_profile: 'This is a test profile.',
    needs_answers_chart: 'needs answers',
    custom_statement: 'test statement',
    tough_spots: 'None flagged.',
    next_gpt: '3-SIPOC'
  };

  writeToSheetCustom(testPayload);
  if (EMAIL_JOEL) emailJoelCustom(testPayload);
  Logger.log('Test handoff complete.');
}
