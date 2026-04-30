const SHEET_NAME = 'd.BOBO Handoffs';
const DO_EMAIL = true;
const CONSULTANT_EMAIL = 'cls0150@auburn.edu';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const payload = {
      gpt: data.gpt || '',
      founder_name: data.founder_name || 'Anonymous',
      date: data.date || new Date().toISOString().split('T')[0],
      status: data.status || '',
      asili_story: data.asili_story || '',
      team: data.team || '',
      proverb: data.proverb || '',
      image_metaphor: data.image_metaphor || '',
      mission_statement: data.mission_statement || '',
      vision_statement: data.vision_statement || '',
      tough_spots: data.tough_spots || '',
      next_gpt: data.next_gpt || ''
    };

    writeToSheet(payload);
    if (DO_EMAIL && CONSULTANT_EMAIL !== 'email') sendEmail(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: 'Handoff recorded via POST' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeToSheet(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Timestamp', 'GPT', 'Founder Name', 'Status', 'Asili Story', 'Team', 'Proverb', 'Image/Metaphor', 'Mission Statement', 'Vision Statement', 'Tough Spots', 'Next GPT']);
    const headerRange = sheet.getRange(1, 1, 1, 12);
    headerRange.setFontWeight('bold').setBackground('#1A1A1A').setFontColor('#FFFFFF');
  }

  sheet.appendRow([
    new Date(),
    payload.gpt,
    payload.founder_name,
    payload.status,
    payload.asili_story,
    payload.team,
    payload.proverb,
    payload.image_metaphor,
    payload.mission_statement,
    payload.vision_statement,
    payload.tough_spots,
    payload.next_gpt
  ]);
}

function sendEmail(payload) {
  const subject = `d.BOBO Handoff — ${payload.founder_name} (GPT ${payload.gpt})`;
  const body = `New d.BOBO handoff summary:

FOUNDER: ${payload.founder_name}
DATE: ${payload.date}
GPT: ${payload.gpt}
STATUS: ${payload.status}

ASILI STORY:
${payload.asili_story || '(not provided)'}

TEAM:
${payload.team || '(not provided)'}

PROVERB:
${payload.proverb || '(not provided)'}

IMAGE/METAPHOR:
${payload.image_metaphor || '(not provided)'}

MISSION STATEMENT:
${payload.mission_statement || '(not provided)'}

VISION STATEMENT:
${payload.vision_statement || '(not provided)'}

TOUGH SPOTS / OPEN QUESTIONS:
${payload.tough_spots || 'None flagged.'}

READY FOR:
${payload.next_gpt}

---
View full spreadsheet: ${SpreadsheetApp.getActiveSpreadsheet().getUrl()}`;

  MailApp.sendEmail(CONSULTANT_EMAIL, subject, body);
}

function testHandoff() {
  const testPayload = {
    gpt: '1-think',
    founder_name: 'Test Founder',
    date: '2026-02-04',
    status: 'complete',
    asili_story: 'This is a test Asili story.',
    team: 'Advisor: Jane Doe, Helper: John Smith',
    proverb: 'If you want to go fast, go alone. If you want to go far, go together.',
    image_metaphor: 'A tree with deep roots',
    mission_statement: 'We provide high-quality test services to test customers.',
    vision_statement: '5 years from now, we are a best-in-class testing organization with a reputation for excellence.',
    tough_spots: 'None flagged.',
    next_gpt: '2-custom'
  };

  writeToSheet(testPayload);
  if (DO_EMAIL) sendEmail(testPayload);
  Logger.log('Test handoff complete.');
}
