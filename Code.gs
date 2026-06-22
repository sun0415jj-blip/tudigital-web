// Google Apps Script - 구독신청 고객 관리

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
      c.id, c.name, c.birth, c.phone, c.product || '', c.amount || '',
      c.period, c.debit, c.status, String(c.confirmed || false), c.createdAt,
      c.docs ? JSON.stringify(c.docs) : ''
    ]);
    notifyManager(c);  // 담당자 알림 발송
  } else if (body.action === 'update') {
    updateRow(sheet, body.customer);
  } else if (body.action === 'delete') {
    deleteRow(sheet, body.id);
  }

  return jsonResponse({ ok: true });
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
    // 카카오 알림톡
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
    // 카카오 미설정 시 SMS 대체 발송
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
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.setName('고객목록');
  const headers = ['id','name','birth','phone','product','amount','period','debit','status','confirmed','createdAt','docs'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
}
