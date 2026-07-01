// Google Apps Script - 구독신청 고객 관리
// Created: 2026-06-30

// =============================================
// 알림 설정 (담당자 정보를 여기에 입력하세요)
// =============================================
const MANAGER_EMAIL  = 'kek3171@naver.com';        // 담당자 이메일
const MANAGER_PHONE  = '01000000000';             // 담당자 핸드폰 (- 없이)

// 솔라피(solapi.com) 가입 후 아래 4개 입력
const SOLAPI_API_KEY    = '';   // 솔라피 API Key
const SOLAPI_SECRET_KEY = '';   // 솔라피 Secret Key
const SENDER_PHONE      = '';   // 발신번호 (솔라피에 등록한 번호, - 없이)

// 카카오 알림톡 설정 (솔라피에서 카카오채널 연동 후 입력)
// 비워두면 SMS로 대신 발송됩니다.
const KAKAO_PFID        = '';   // 카카오채널 pfId
const KAKAO_TEMPLATE_ID = '';   // 알림톡 템플릿 ID
// =============================================

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('고객목록')
      || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function doGet(e) {
  // proxyImage: 이미지를 서버에서 가져와 base64로 반환 (CORS 우회)
  if (e.parameter && e.parameter.action === 'proxyImage') {
    try {
      const resp = UrlFetchApp.fetch(e.parameter.url, { muteHttpExceptions: true });
      const blob = resp.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      return jsonResponse({ ok: true, data: base64, type: blob.getContentType() || 'image/jpeg' });
    } catch(err) {
      return jsonResponse({ ok: false, error: String(err) });
    }
  }

  // saveAllDocUrls: 3개 imgbb URL을 한 번에 저장 (GET 방식, CORS 안전)
  if (e.parameter && e.parameter.action === 'saveAllDocUrls') {
    let customerName = '', customerPhone = '';
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const s = getSheet();
      const allData = s.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(e.parameter.customerId)) {
          customerName  = allData[i][1] || '';
          customerPhone = allData[i][3] || '';
          const docs = {
            doc1: e.parameter.doc1 || '',
            doc2: e.parameter.doc2 || '',
            doc3: e.parameter.doc3 || ''
          };
          s.getRange(i + 1, 12).setValue(JSON.stringify(docs));
          s.getRange(i + 1, 9).setValue('서류제출');
          break;
        }
      }
    } finally {
      lock.releaseLock();
    }
    if (customerName) {
      try {
        MailApp.sendEmail({
          to: MANAGER_EMAIL,
          subject: '[서류접수] ' + customerName + ' 고객 서류 업로드 완료',
          htmlBody:
            '<h3 style="color:#0d2137;">서류가 접수되었습니다</h3>' +
            '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">' +
              '<tr><td><b>고객명</b></td><td>' + customerName + '</td></tr>' +
              '<tr><td><b>연락처</b></td><td>' + customerPhone + '</td></tr>' +
              '<tr><td><b>접수시각</b></td><td>' + new Date().toLocaleString('ko-KR') + '</td></tr>' +
            '</table>' +
            '<p style="margin-top:16px;color:#2b6cb0;font-weight:bold;">➡ 어드민에서 서류를 확인해 주세요.</p>'
        });
      } catch(err) {}
    }
    return jsonResponse({ ok: true });
  }

  const sheet = getSheet();
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
  if (!e.postData || !e.postData.contents) {
    return jsonResponse({ ok: false, error: 'no_body' });
  }
  const body = JSON.parse(e.postData.contents);

  if (body.action === 'save') {
    const sheet = getSheet();
    const c = body.customer;
    sheet.appendRow([
      c.id, c.name, c.birth, c.phone, c.product || '', c.amount || '',
      c.period, c.debit, c.status, String(c.confirmed || false), c.createdAt,
      c.docs ? JSON.stringify(c.docs) : ''
    ]);
    notifyManager(c);

  } else if (body.action === 'update') {
    updateRow(getSheet(), body.customer);

  } else if (body.action === 'delete') {
    deleteRow(getSheet(), body.id);

  } else if (body.action === 'contact') {
    sendContactEmail(body);

  } else if (body.action === 'uploadOneFile') {
    // Upload single file to Drive (slow - outside lock)
    const url = uploadFileToDrive(body.data, body.name, body.type, body.customerId, body.docKey);
    // Atomic read-modify-write only (fast - inside lock)
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const s = getSheet();
      const allData = s.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(body.customerId)) {
          let docs = {};
          try { if (allData[i][11]) docs = JSON.parse(allData[i][11]); } catch(err) {}
          docs[body.docKey] = url;
          s.getRange(i + 1, 12).setValue(JSON.stringify(docs));
          if (Object.keys(docs).length >= 3) {
            s.getRange(i + 1, 9).setValue('서류제출');
          }
          break;
        }
      }
    } finally {
      lock.releaseLock();
    }

  } else if (body.action === 'uploadAllFiles') {
    // Legacy: upload all 3 files sequentially in one execution
    const sheet = getSheet();
    const docs = {};
    for (const f of body.files) {
      docs[f.docKey] = uploadFileToDrive(f.data, f.name, f.type, body.customerId, f.docKey);
    }
    const allData = sheet.getDataRange().getValues();
    for (let i = 1; i < allData.length; i++) {
      if (String(allData[i][0]) === String(body.customerId)) {
        sheet.getRange(i + 1, 12).setValue(JSON.stringify(docs));
        sheet.getRange(i + 1, 9).setValue('서류제출');
        break;
      }
    }

  } else if (body.action === 'saveDocUrl') {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const s = getSheet();
      const allData = s.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(body.customerId)) {
          let docs = {};
          try { if (allData[i][11]) docs = JSON.parse(allData[i][11]); } catch(err) {}
          docs[body.docKey] = body.url;
          s.getRange(i + 1, 12).setValue(JSON.stringify(docs));
          if (Object.keys(docs).length >= 3) {
            s.getRange(i + 1, 9).setValue('서류제출');
          }
          break;
        }
      }
    } finally {
      lock.releaseLock();
    }
  }

  return jsonResponse({ ok: true });
}

// ── Google Drive 파일 업로드 ──────────────────────
function getOrCreateFolder(folderName) {
  const iter = DriveApp.getFoldersByName(folderName);
  return iter.hasNext() ? iter.next() : DriveApp.createFolder(folderName);
}

function uploadFileToDrive(base64Data, fileName, mimeType, customerId, docKey) {
  const folder   = getOrCreateFolder('티유디지털_구독서류');
  const base64   = base64Data.indexOf(',') > -1 ? base64Data.split(',')[1] : base64Data;
  const bytes    = Utilities.base64Decode(base64);
  const safeName = String(customerId) + '_' + docKey + '_' + fileName;
  const blob     = Utilities.newBlob(bytes, mimeType || 'application/octet-stream', safeName);
  const file     = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?export=download&id=' + file.getId();
}

// ── 파트너 문의 이메일 ────────────────────────
function sendContactEmail(data) {
  MailApp.sendEmail({
    to: MANAGER_EMAIL,
    subject: '[파트너 문의] ' + data.name + ' (' + (data.category || '기타') + ')',
    htmlBody:
      '<h3 style="color:#0d2137;">파트너 문의가 접수되었습니다</h3>' +
      '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">' +
        '<tr><td><b>이름</b></td><td>' + data.name + '</td></tr>' +
        '<tr><td><b>연락처</b></td><td>' + data.phone + '</td></tr>' +
        '<tr><td><b>이메일</b></td><td>' + (data.email || '-') + '</td></tr>' +
        '<tr><td><b>문의 유형</b></td><td>' + (data.category || '-') + '</td></tr>' +
        '<tr><td><b>문의 내용</b></td><td style="white-space:pre-line;">' + data.message + '</td></tr>' +
      '</table>'
  });
}

// ── 담당자 알림 ──────────────────────────────
function notifyManager(customer) {
  sendEmail(customer);
  if (SOLAPI_API_KEY && SENDER_PHONE) {
    sendKakaoOrSMS(customer);
  }
}

function sendEmail(customer) {
  MailApp.sendEmail({
    to: MANAGER_EMAIL,
    subject: '[신용조회 요청] ' + customer.name + ' 고객 신청 접수',
    htmlBody:
      '<h3 style="color:#0d2137;">신규 구독 신청 — 신용조회 필요</h3>' +
      '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;">' +
        '<tr><td><b>고객명</b></td><td>' + customer.name + '</td></tr>' +
        '<tr><td><b>생년월일</b></td><td>' + customer.birth + '</td></tr>' +
        '<tr><td><b>연락처</b></td><td>' + customer.phone + '</td></tr>' +
        '<tr><td><b>상품명</b></td><td>' + (customer.product || '-') + '</td></tr>' +
        '<tr><td><b>금액</b></td><td>' + (customer.amount || '-') + '원</td></tr>' +
        '<tr><td><b>구독기간</b></td><td>' + customer.period + '</td></tr>' +
        '<tr><td><b>자동이체일</b></td><td>' + customer.debit + '</td></tr>' +
        '<tr><td><b>신청일시</b></td><td>' + customer.createdAt + '</td></tr>' +
      '</table>' +
      '<p style="margin-top:16px;color:#c0392b;font-weight:bold;">➡ 신용조회를 진행해 주세요.</p>'
  });
}

function sendKakaoOrSMS(customer) {
  const date      = new Date().toISOString();
  const salt      = Utilities.getUuid();
  const signature = computeHmac(date + salt, SOLAPI_SECRET_KEY);
  const authHeader = 'HMAC-SHA256 apiKey=' + SOLAPI_API_KEY +
                     ', date=' + date +
                     ', salt=' + salt +
                     ', signature=' + signature;

  let message;
  if (KAKAO_PFID && KAKAO_TEMPLATE_ID) {
    message = {
      to:   MANAGER_PHONE,
      from: SENDER_PHONE,
      kakaoOptions: {
        pfId:       KAKAO_PFID,
        templateId: KAKAO_TEMPLATE_ID,
        variables: {
          '#{고객명}':   customer.name,
          '#{연락처}':   customer.phone,
          '#{구독기간}': customer.period,
          '#{신청일시}': customer.createdAt
        }
      }
    };
  } else {
    message = {
      to:   MANAGER_PHONE,
      from: SENDER_PHONE,
      text: '[티유디지털] 신용조회 요청\n' +
            '고객: ' + customer.name + '\n' +
            '연락처: ' + customer.phone + '\n' +
            '구독: ' + customer.period + '\n' +
            '신청: ' + customer.createdAt
    };
  }

  UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
    method:           'POST',
    headers:          { 'Authorization': authHeader, 'Content-Type': 'application/json' },
    payload:          JSON.stringify({ message: message }),
    muteHttpExceptions: true
  });
}

function computeHmac(data, secret) {
  const sig = Utilities.computeHmacSha256Signature(data, secret);
  return sig.map(function(b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
// ─────────────────────────────────────────────

function updateRow(sheet, c) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(c.id)) {
      sheet.getRange(i + 1, 1, 1, 12).setValues([[
        c.id, c.name, c.birth, c.phone, c.product || '', c.amount || '',
        c.period, c.debit, c.status, String(c.confirmed || false), c.createdAt,
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
  const sheet = getSheet();
  sheet.setName('고객목록');
  const headers = ['id','name','birth','phone','product','amount','period','debit','status','confirmed','createdAt','docs'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
