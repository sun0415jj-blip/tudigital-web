// Google Apps Script - 구독신청 고객 관리
// 사용법: 아래 코드를 Google Apps Script에 붙여넣고 웹 앱으로 배포

function doGet(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return jsonResponse({ ok: true, data: [] });
  }
  const headers = data[0];
  const rows = data.slice(1)
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      if (obj.docs && typeof obj.docs === 'string' && obj.docs !== '') {
        try { obj.docs = JSON.parse(obj.docs); } catch(e) { obj.docs = null; }
      } else {
        obj.docs = null;
      }
      return obj;
    });
  return jsonResponse({ ok: true, data: rows });
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (body.action === 'save') {
    const c = body.customer;
    sheet.appendRow([
      c.id, c.name, c.birth, c.phone, c.period, c.debit,
      c.status, String(c.confirmed || false), c.createdAt,
      c.docs ? JSON.stringify(c.docs) : ''
    ]);
  } else if (body.action === 'update') {
    updateRow(sheet, body.customer);
  } else if (body.action === 'delete') {
    deleteRow(sheet, body.id);
  }

  return jsonResponse({ ok: true });
}

function updateRow(sheet, c) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(c.id)) {
      sheet.getRange(i + 1, 1, 1, 10).setValues([[
        c.id, c.name, c.birth, c.phone, c.period, c.debit,
        c.status, String(c.confirmed || false), c.createdAt,
        c.docs ? JSON.stringify(c.docs) : ''
      ]]);
      return;
    }
  }
}

function deleteRow(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) {
      sheet.deleteRow(i + 1);
      return;
    }
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// 최초 1회 실행 - 헤더 설정
function setupSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName('고객목록');
  const headers = ['id','name','birth','phone','period','debit','status','confirmed','createdAt','docs'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
