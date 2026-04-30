function doPostFunnel(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const payload = {
      gpt: data.gpt || '',
      founder_name: data.founder_name || 'Anonymous',
      date: data.date || new Date().toISOString().split('T')[0],
      status: data.status || '',
      starting_customer: data.starting_customer || '',
      awareness_channels: data.awareness_channels || '',
      consideration_assets: data.consideration_assets || '',
      sale_process: data.sale_process || '',
      account_management: data.account_management || '',
      brand_words: data.brand_words || '',
      launch_readiness: data.launch_readiness || '',
      first_48hr_action: data.first_48hr_action || '',
      tough_spots: data.tough_spots || '',
      next_gpt: data.next_gpt || ''
    };

    writeToSheetFunnel(payload);
    if (EMAIL_JOEL && JOEL_EMAIL !== 'email') emailJoelFunnel(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Handoff recorded via POST' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeToSheetFunnel(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('d.BOBO Funnel');

  if (!sheet) {
    sheet = ss.insertSheet('d.BOBO Funnel');
    sheet.appendRow(['Timestamp', 'GPT', 'Founder Name', 'Status', 'Starting \nCustomer', 'Awareness \nChannels', 'Consideration \nAssets', 'Sale Process', 'Account Mgmt', 'Brand Words', 'Launch \nReadiness', '48hr Action', 'Tough Spots', 'Next GPT']);
    const headerRange = sheet.getRange(1, 1, 1, 14);
    headerRange.setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FFFFFF');
  }

  sheet.appendRow([
    new Date(),
    payload.gpt,
    payload.founder_name,
    payload.status,
    payload.starting_customer,
    payload.awareness_channels,
    payload.consideration_assets,
    payload.sale_process,
    payload.account_management,
    payload.brand_words,
    payload.launch_readiness,
    payload.first_48hr_action,
    payload.tough_spots,
    payload.next_gpt
  ]);
}

function emailJoelFunnel(payload) {
  const subject = `d.BOBO Handoff — ${payload.founder_name} (GPT ${payload.gpt})`;
  const body = `New d.BOBO handoff summary:

FOUNDER: ${payload.founder_name}
DATE: ${payload.date}
GPT: ${payload.gpt}
STATUS: ${payload.status}

STARTING CUSTOMER:
${payload.starting_customer || '(not provided)'}

AWARENESS CHANNELS:
${payload.awareness_channels || '(not provided)'}

CONSIDERATION ASSETS:
${payload.consideration_assets || '(not provided)'}

SALE PROCESS:
${payload.sale_process || '(not provided)'}

ACCOUNT MANAGEMENT:
${payload.account_management || '(not provided)'}

BRAND WORDS:
${payload.brand_words || '(not provided)'}

LAUNCH READINESS:
${payload.launch_readiness || '(not provided)'}

FIRST 48 ACTION:
${payload.first_48hr_action || '(not provided)'}

TOUGH SPOTS / OPEN QUESTIONS:
${payload.tough_spots || 'None flagged.'}

READY FOR:
${payload.next_gpt}

---
View full spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  MailApp.sendEmail(JOEL_EMAIL, subject, body);
}

function testHandoffFunnel() {
  const testPayload = {
    gpt: '4-funnel',
    founder_name: 'Test Founder',
    date: '2026-04-29',
    status: 'complete',
    starting_customer: 'test customer',
    awareness_channels: 'test channels',
    consideration_assets: 'test assets',
    sale_process: 'test sale process',
    account_management: 'test acct mgmt',
    brand_words: 'test words',
    launch_readiness: 'test readiness',
    first_48hr_action: 'test action',
    tough_spots: 'None flagged.',
    next_gpt: 'None'
  };

  writeToSheetFunnel(testPayload);
  if (EMAIL_JOEL) emailJoelFunnel(testPayload);
  Logger.log('Test handoff complete.');
}
