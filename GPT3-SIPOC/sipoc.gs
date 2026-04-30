function doPostSIPOC(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const payload = {
      gpt: data.gpt || '',
      founder_name: data.founder_name || 'Anonymous',
      date: data.date || new Date().toISOString().split('T')[0],
      status: data.status || '',
      suppliers: data.suppliers || '',
      inputs: data.inputs || '',
      process_steps: data.process_steps || '',
      outputs: data.outputs || '',
      customers: data.customers || '',
      shop: data.shop || '',
      coq_snapshot: data.coq_snapshot || '',
      tough_spots: data.tough_spots || '',
      next_gpt: data.next_gpt || ''
    };

    writeToSheetSIPOC(payload);
    if (EMAIL_JOEL && JOEL_EMAIL !== 'email') emailJoelSIPOC(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Handoff recorded via POST' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeToSheetSIPOC(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('d.BOBO SIPOC');

  if (!sheet) {
    sheet = ss.insertSheet('d.BOBO SIPOC');
    sheet.appendRow(['Timestamp', 'GPT', 'Founder Name', 'Status', 'Suppliers', 'Inputs', 'Process Steps', 'Outputs', 'Customers', 'Shop', 'COQ Snapshot', 'Tough Spots', 'Next GPT']);
    const headerRange = sheet.getRange(1, 1, 1, 13);
    headerRange.setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FFFFFF');
  }

  sheet.appendRow([
    new Date(),
    payload.gpt,
    payload.founder_name,
    payload.status,
    payload.suppliers,
    payload.inputs,
    payload.process_steps,
    payload.outputs,
    payload.customers,
    payload.shop,
    payload.coq_snapshot,
    payload.tough_spots,
    payload.next_gpt
  ]);
}

function emailJoelSIPOC(payload) {
  const subject = `d.BOBO Handoff — ${payload.founder_name} (GPT ${payload.gpt})`;
  const body = `New d.BOBO handoff summary:

FOUNDER: ${payload.founder_name}
DATE: ${payload.date}
GPT: ${payload.gpt}
STATUS: ${payload.status}

SUPPLIERS:
${payload.suppliers || '(not provided)'}

INPUTS:
${payload.inputs || '(not provided)'}

PROCESS STEPS:
${payload.process_steps || '(not provided)'}

OUTPUTS:
${payload.outputs || '(not provided)'}

CUSTOMERS:
${payload.customers || '(not provided)'}

SHOP:
${payload.shop || '(not provided)'}

COQ SNAPSHOT:
${payload.coq_snapshot || '(not provided)'}

TOUGH SPOTS / OPEN QUESTIONS:
${payload.tough_spots || 'None flagged.'}

READY FOR:
${payload.next_gpt}

---
View full spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  MailApp.sendEmail(JOEL_EMAIL, subject, body);
}

function testHandoffSIPOC() {
  const testPayload = {
    gpt: '3-SIPOC',
    founder_name: 'Test Founder',
    date: '2026-04-29',
    status: 'complete',
    suppliers: 'test suppliers',
    inputs: 'test inputs',
    process_steps: 'test steps',
    outputs: 'test outputs',
    customers: 'test customers',
    shop: 'test shop',
    coq_snapshot: 'test number',
    tough_spots: 'None flagged.',
    next_gpt: '4-funnel'
  };

  writeToSheetSIPOC(testPayload);
  if (EMAIL_JOEL) emailJoelSIPOC(testPayload);
  Logger.log('Test handoff complete.');
}
