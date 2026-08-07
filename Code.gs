// Google Apps Script - 티유디지털 구독신청 고객 관리
// Updated: 2026-07-03

// =============================================
// 알림 설정 (담당자 정보를 여기에 입력하세요)
// =============================================
const MANAGER_EMAIL  = 'kek3171@naver.com';
const MANAGER_PHONE  = '01000000000';

const SOLAPI_API_KEY    = '';
const SOLAPI_SECRET_KEY = '';
const SENDER_PHONE      = '';

const KAKAO_PFID        = '';
const KAKAO_TEMPLATE_ID = '';
// =============================================

const CACHE_KEY = 'customers_v1';
const CACHE_TTL = 60; // 초

// =============================================
// 2단계 제출 서류 (구독절차 기준)
//   doc1 등기부등본/매매계약서 · doc2 신분증 · doc3 통장 · doc4 견적서
//   doc5 가족관계증명서 — 가족 명의 거주주택일 때만 필수
// =============================================
const DOC_LABELS = {
  doc1: '등기부등본 / 매매계약서',
  doc2: '신분증 사본',
  doc3: '통장 사본',
  doc4: '견적서',
  doc5: '가족관계증명서'
};

// ownership 이 비어 있으면 '구버전 화면(또는 구버전으로 저장된 건)'으로 본다.
// 화면 배포와 백엔드 배포 시점이 어긋나도 기존 3종 업로드가 계속 동작하도록 하위호환을 둔다.
// 신버전 화면은 항상 ownership=self|family 를 함께 보낸다.
function REQUIRED_DOCS(ownership) {
  if (!ownership) return ['doc1', 'doc2', 'doc3'];              // 구버전: 등기부/신분증/통장
  return ownership === 'family'
    ? ['doc1', 'doc2', 'doc3', 'doc4', 'doc5']                  // 가족 소유: +견적서 +가족관계증명서
    : ['doc1', 'doc2', 'doc3', 'doc4'];                         // 본인 소유: +견적서
}

// docs JSON 기준으로 필수 서류가 다 찼는지 판정
// (ownership 없는 과거 데이터는 예전 기준으로 판정되어 상태가 되돌아가지 않는다)
function docsComplete(docs) {
  docs = docs || {};
  return REQUIRED_DOCS(docs.ownership).every(function(k) { return !!docs[k]; });
}

// =============================================
// 진행 상태 (구독절차 기준)
//   서류대기 → 서류제출 → 확인완료(한국캐피탈 승인) → 약정완료(3단계 온라인 약정)
// 3단계 약정 상태는 docs JSON에 보관한다 (시트 컬럼 추가 없이):
//   docs.agreement   : '' | 'done'
//   docs.agreementAt : 완료 시각 (ISO)
// 약정은 한국캐피탈이 고객 휴대폰으로 발송하며, 반드시 시공 전에 완료되어야 한다.
// =============================================
// 4단계(시공완료 후)도 docs JSON에 보관한다:
//   docs.cdoc1  시공계약서 / docs.cdoc2 시공확인서
//   docs.photos 시공사진 URL 배열 (최소 6장)
//   docs.recording '' | 'done'   한국캐피탈 → 고객 녹취
//   docs.payment   '' | 'done'   한국캐피탈 → 판매점 입금 (녹취 후 1~2일)
const STATUS = {
  WAIT:      '서류대기',
  SUBMITTED: '서류제출',
  CONFIRMED: '확인완료',
  AGREED:    '약정완료',
  BUILT:     '시공완료',
  RECORDED:  '녹취완료',
  PAID:      '입금완료'
};

const MIN_PHOTOS = 6;

// =============================================================
// 업체 계정 · 로그인 · 열람 범위
// -------------------------------------------------------------
// 티유디지털(마스터)은 전체를 보고, 업체는 자기 업체 건만 본다.
// 거르는 일은 서버(여기)에서 한다. 화면에서만 거르면 개발자도구로 뚫린다.
//
// 계정은 같은 스프레드시트의 '업체계정' 시트에 둔다.
//   orgId | orgName | pwHash | active | createdAt
// 업체가 늘어나면 어드민 '업체 관리'에서 추가하면 되고 코드는 손대지 않는다.
//
// 비밀번호 원문은 서버로 오지 않는다.
//   화면에서  clientHash = SHA256(orgId + ':' + 비밀번호)  를 만들어 보내고
//   서버는    SHA256(clientHash + AUTH_SALT_()) 를 저장·비교한다.
//   → 주소·실행로그 어디에도 원문이 남지 않는다.
// =============================================================
const ACCOUNT_SHEET = '업체계정';
const MASTER_ID     = 'tudigital';          // 전체 열람 계정
const TOKEN_TTL_MS  = 12 * 60 * 60 * 1000;  // 12시간

// 솔트와 서명키는 코드에 두지 않는다. 이 저장소는 공개라 코드에 적으면 그대로 노출된다.
// 스크립트 속성(프로젝트 설정 → 스크립트 속성)에 보관하고, 없으면 처음 호출될 때 만들어 저장한다.
//   AUTH_SALT    : 비밀번호 해시용. 바뀌면 전 계정 비밀번호를 재설정해야 한다.
//   TOKEN_SECRET : 토큰 서명용. 바뀌면 로그인 세션이 모두 끊긴다.
function getSecret_(key) {
  const props = PropertiesService.getScriptProperties();
  let v = props.getProperty(key);
  if (!v) {
    v = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    props.setProperty(key, v);
  }
  return v;
}
function AUTH_SALT_()    { return getSecret_('AUTH_SALT'); }
function TOKEN_SECRET_() { return getSecret_('TOKEN_SECRET'); }

function sha256hex(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

function getAccountSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ACCOUNT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ACCOUNT_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['orgId', 'orgName', 'pwHash', 'active', 'createdAt']])
      .setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function readAccounts() {
  const data = getAccountSheet().getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) {
    return { orgId: String(r[0]), orgName: String(r[1]), pwHash: String(r[2]),
             active: String(r[3]) !== 'N', createdAt: r[4] };
  });
}

function findAccount(orgId) {
  const list = readAccounts();
  for (let i = 0; i < list.length; i++) if (list[i].orgId === String(orgId)) return list[i];
  return null;
}

function updateAccountField(orgId, col, value) {
  const sh = getAccountSheet();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(orgId)) { sh.getRange(i + 1, col).setValue(value); return true; }
  }
  return false;
}

function issueToken(orgId, orgName) {
  const b64 = Utilities.base64EncodeWebSafe(orgId + '|' + orgName + '|' + (Date.now() + TOKEN_TTL_MS));
  return b64 + '.' + computeHmac(b64, TOKEN_SECRET_());
}

// 유효하면 { orgId, orgName, isMaster }, 아니면 null
function verifyToken(token) {
  if (!token || String(token).indexOf('.') === -1) return null;
  const p = String(token).split('.');
  if (computeHmac(p[0], TOKEN_SECRET_()) !== p[1]) return null;
  let payload;
  try { payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString(); }
  catch (e) { return null; }
  const seg = payload.split('|');
  if (seg.length < 3 || Number(seg[2]) < Date.now()) return null;
  const acc = findAccount(seg[0]);
  if (!acc || !acc.active) return null;
  return { orgId: acc.orgId, orgName: acc.orgName, isMaster: acc.orgId === MASTER_ID };
}

// 최초 1회 실행 — 마스터와 첫 업체 계정을 만든다.
//
//  [ 사용법 ]
//   1. 아래 SETUP_PW 의 '여기에_비밀번호_입력' 을 원하는 비밀번호로 바꾼다
//   2. 함수 목록에서 setupAccounts 를 골라 실행 ▶
//   3. 실행이 끝나면 SETUP_PW 를 다시 '여기에_비밀번호_입력' 으로 되돌린다
//
//  이 저장소는 공개이므로 비밀번호를 적어둔 채로 두지 말 것.
//  계정은 시트에 해시로만 저장되므로, 여기 값을 지워도 로그인에는 지장이 없다.
//  비밀번호를 잊으면 마스터로 로그인해 '업체 관리 → 비밀번호 재설정' 으로 바꾸면 된다.
//  마스터 비밀번호를 잊으면 업체계정 시트에서 tudigital 행을 지우고 다시 실행한다.
const SETUP_PW = {
  tudigital:   '여기에_비밀번호_입력',   // 티유디지털 (전체 열람)
  changhohome: '여기에_비밀번호_입력'    // 창호홈
};

function setupAccounts() {
  const sh = getAccountSheet();
  const seed = [
    { orgId: MASTER_ID,     orgName: '티유디지털' },
    { orgId: 'changhohome', orgName: '창호홈'    }
  ];

  const notSet = seed.filter(function (s) {
    const pw = SETUP_PW[s.orgId];
    return !pw || pw === '여기에_비밀번호_입력';
  });
  if (notSet.length) {
    Logger.log('■ 실행하지 않았습니다.');
    Logger.log('  위쪽 SETUP_PW 에서 아래 계정의 비밀번호를 먼저 입력하세요:');
    notSet.forEach(function (s) { Logger.log('   - ' + s.orgId + ' (' + s.orgName + ')'); });
    return;
  }

  seed.forEach(function (s) {
    if (findAccount(s.orgId)) { Logger.log('이미 있음: ' + s.orgId); return; }
    const pw = SETUP_PW[s.orgId];
    sh.appendRow([s.orgId, s.orgName, sha256hex(sha256hex(s.orgId + ':' + pw) + AUTH_SALT_()), 'Y',
                  Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')]);
    Logger.log('생성 완료  아이디: ' + s.orgId + '  (' + s.orgName + ')');
  });

  Logger.log('');
  Logger.log('※ 이제 SETUP_PW 값을 다시 지워 주세요.');
  Logger.log('※ orgName 은 신청 링크의 ?org= 값과 정확히 같아야 그 업체 건이 보입니다.');
  Logger.log('※ 업체 추가는 어드민 → 업체 관리에서 하면 됩니다.');
}

function getSheet() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('고객목록')
      || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 캐시 무효화 ───────────────────────────────
function invalidateCache() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}

// ── 전체 고객 데이터 반환 ──────────────────────
function getAllCustomers() {
  // 캐시 확인
  try {
    const cached = CacheService.getScriptCache().get(CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  const sheet = getSheet();
  const data  = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1)
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      // docs JSON 파싱
      if (obj.docs && typeof obj.docs === 'string' && obj.docs.trim() !== '') {
        try { obj.docs = JSON.parse(obj.docs); } catch(e) { obj.docs = null; }
      } else {
        obj.docs = null;
      }
      return obj;
    });

  // 캐시 저장
  try {
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(rows), CACHE_TTL);
  } catch(e) {}

  return rows;
}

function doGet(e) {
  const action = e.parameter && e.parameter.action;

  // ── proxyImage: ImgBB CORS 우회 다운로드 ──────
  if (action === 'proxyImage') {
    const url = e.parameter.url || '';
    if (!url.startsWith('http')) {
      return jsonResponse({ ok: false, error: 'invalid_url' });
    }
    try {
      const resp   = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
      const code   = resp.getResponseCode();
      if (code !== 200) return jsonResponse({ ok: false, error: 'fetch_failed', code });
      const blob   = resp.getBlob();
      const base64 = Utilities.base64Encode(blob.getBytes());
      return jsonResponse({ ok: true, data: base64, type: blob.getContentType() || 'image/jpeg' });
    } catch(err) {
      return jsonResponse({ ok: false, error: String(err) });
    }
  }

  // ── saveAllDocUrls: 서류 URL 저장 ──────────
  // 본인 소유   : doc1 등기부등본/매매계약서, doc2 신분증, doc3 통장, doc4 견적서
  // 가족 소유   : 위 4종 + doc5 가족관계증명서 (doc1은 등기부등본만)
  if (action === 'saveAllDocUrls') {
    const customerId = String(e.parameter.customerId || '');
    const jobType    = e.parameter.jobType   || '';
    const scheduled  = e.parameter.scheduledAt || '';   // 계약서 기재 시공예정일

    // ownership 이 없으면 구버전 화면 → 예전 규칙(doc1~3)으로 받는다
    const rawOwn    = e.parameter.ownership || '';
    const isLegacy  = !rawOwn;
    const ownership = isLegacy ? '' : (rawOwn === 'family' ? 'family' : 'self');

    const incoming = {};
    ['doc1', 'doc2', 'doc3', 'doc4', 'doc5'].forEach(function(k) {
      if (e.parameter[k]) incoming[k] = e.parameter[k];
    });

    if (!customerId) return jsonResponse({ ok: false, error: 'no_customer_id' });
    const missing = REQUIRED_DOCS(ownership).filter(function(k) { return !incoming[k]; });
    if (missing.length) return jsonResponse({ ok: false, error: 'missing_docs', missing: missing });

    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch(e) {
      return jsonResponse({ ok: false, error: 'lock_timeout' });
    }

    let customerName = '', customerPhone = '', found = false;
    try {
      const s       = getSheet();
      const allData = s.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === customerId) {
          customerName  = allData[i][1] || '';
          customerPhone = allData[i][3] || '';

          let docs = {};
          try { if (allData[i][11]) docs = JSON.parse(allData[i][11]); } catch(err) {}
          Object.keys(incoming).forEach(function(k) { docs[k] = incoming[k]; });
          // 구버전 화면이면 ownership 을 기록하지 않는다 (기존 건의 판정 기준을 유지)
          if (!isLegacy) {
            docs.ownership = ownership;
            if (ownership === 'self') delete docs.doc5;   // 본인 소유면 가족관계증명서 불필요
          }
          if (jobType)   docs.jobType     = jobType;
          if (scheduled) docs.scheduledAt = scheduled;

          s.getRange(i + 1, 12).setValue(JSON.stringify(docs));
          s.getRange(i + 1, 9).setValue('서류제출');
          found = true;
          break;
        }
      }
    } finally {
      lock.releaseLock();
    }

    if (!found) return jsonResponse({ ok: false, error: 'customer_not_found' });

    invalidateCache();

    // 담당자 이메일 알림
    if (customerName && MANAGER_EMAIL) {
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

  // ── saveConstructionDocs: 4단계 시공완료 서류 ──────────
  // 온라인 약정(3단계) 완료 건에만 허용한다.
  if (action === 'saveConstructionDocs') {
    const customerId = String(e.parameter.customerId || '');
    const cdoc1  = e.parameter.cdoc1 || '';
    const cdoc2  = e.parameter.cdoc2 || '';
    const photos = String(e.parameter.photos || '').split('|').filter(function (u) { return !!u; });

    if (!customerId) return jsonResponse({ ok: false, error: 'no_customer_id' });
    if (!cdoc1 || !cdoc2) return jsonResponse({ ok: false, error: 'missing_docs' });
    if (photos.length < MIN_PHOTOS) {
      return jsonResponse({ ok: false, error: 'not_enough_photos', need: MIN_PHOTOS, got: photos.length });
    }

    const lock = LockService.getScriptLock();
    try { lock.waitLock(30000); } catch (err) { return jsonResponse({ ok: false, error: 'lock_timeout' }); }

    let customerName = '', customerPhone = '', found = false, notAgreed = false;
    try {
      const s = getSheet();
      const allData = s.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) !== customerId) continue;
        customerName  = allData[i][1] || '';
        customerPhone = allData[i][3] || '';

        let docs = {};
        try { if (allData[i][11]) docs = JSON.parse(allData[i][11]); } catch (err) {}
        if (docs.agreement !== 'done') { notAgreed = true; break; }

        docs.cdoc1  = cdoc1;
        docs.cdoc2  = cdoc2;
        docs.photos = photos;
        docs.constructionAt = new Date().toISOString();

        s.getRange(i + 1, 12).setValue(JSON.stringify(docs));
        s.getRange(i + 1, 9).setValue(STATUS.BUILT);
        found = true;
        break;
      }
    } finally {
      lock.releaseLock();
    }

    if (notAgreed) return jsonResponse({ ok: false, error: 'not_agreed' });
    if (!found)    return jsonResponse({ ok: false, error: 'customer_not_found' });

    invalidateCache();

    if (customerName && MANAGER_EMAIL) {
      try {
        MailApp.sendEmail({
          to: MANAGER_EMAIL,
          subject: '[시공완료] ' + customerName + ' 고객 시공 서류 제출',
          htmlBody:
            '<h3 style="color:#0d2137;">시공완료 서류가 접수되었습니다</h3>' +
            '<table border="1" cellpadding="8" cellspacing="0" style="border-collapse:collapse;font-size:14px;">' +
              '<tr><td><b>고객명</b></td><td>' + customerName + '</td></tr>' +
              '<tr><td><b>연락처</b></td><td>' + customerPhone + '</td></tr>' +
              '<tr><td><b>시공사진</b></td><td>' + photos.length + '장</td></tr>' +
              '<tr><td><b>접수시각</b></td><td>' + new Date().toLocaleString('ko-KR') + '</td></tr>' +
            '</table>' +
            '<p style="margin-top:16px;color:#2b6cb0;font-weight:bold;">➡ 한국캐피탈 녹취를 요청해 주세요.</p>'
        });
      } catch (err) {}
    }

    return jsonResponse({ ok: true });
  }



  // ── 로그인 ────────────────────────────────────
  if (action === 'login') {
    const orgId = String(e.parameter.orgId || '').trim();
    const ch    = String(e.parameter.h || '');
    if (!orgId || !ch) return jsonResponse({ ok: false, error: 'missing' });
    const acc = findAccount(orgId);
    // 아이디가 없는지 비번이 틀렸는지 구분해서 알려주지 않는다
    if (!acc || !acc.active || sha256hex(ch + AUTH_SALT_()) !== acc.pwHash) {
      return jsonResponse({ ok: false, error: 'invalid' });
    }
    return jsonResponse({ ok: true, token: issueToken(acc.orgId, acc.orgName),
      orgId: acc.orgId, orgName: acc.orgName, isMaster: acc.orgId === MASTER_ID });
  }

  // ── 비밀번호 변경 (본인) ───────────────────────
  if (action === 'changePw') {
    const auth = verifyToken(e.parameter.token);
    if (!auth) return jsonResponse({ ok: false, error: 'unauthorized' });
    const acc = findAccount(auth.orgId);
    if (!acc || sha256hex(String(e.parameter.oldH || '') + AUTH_SALT_()) !== acc.pwHash) {
      return jsonResponse({ ok: false, error: 'wrong_password' });
    }
    updateAccountField(auth.orgId, 3, sha256hex(String(e.parameter.newH || '') + AUTH_SALT_()));
    return jsonResponse({ ok: true });
  }

  // ── 업체 관리 (마스터 전용) ────────────────────
  if (action === 'orgList' || action === 'orgAdd' || action === 'orgSetPw' || action === 'orgToggle') {
    const auth = verifyToken(e.parameter.token);
    if (!auth || !auth.isMaster) return jsonResponse({ ok: false, error: 'unauthorized' });

    if (action === 'orgList') {
      return jsonResponse({ ok: true, data: readAccounts().map(function (a) {
        return { orgId: a.orgId, orgName: a.orgName, active: a.active, createdAt: a.createdAt };
      }) });
    }
    if (action === 'orgAdd') {
      const id = String(e.parameter.orgId || '').trim();
      const nm = String(e.parameter.orgName || '').trim();
      const h  = String(e.parameter.h || '');
      if (!id || !nm || !h)          return jsonResponse({ ok: false, error: 'missing' });
      if (!/^[a-z0-9_-]+$/.test(id)) return jsonResponse({ ok: false, error: 'bad_id' });
      if (findAccount(id))           return jsonResponse({ ok: false, error: 'duplicate' });
      getAccountSheet().appendRow([id, nm, sha256hex(h + AUTH_SALT_()), 'Y',
        Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm')]);
      return jsonResponse({ ok: true });
    }
    if (action === 'orgSetPw') {
      const id = String(e.parameter.orgId || '');
      if (!findAccount(id)) return jsonResponse({ ok: false, error: 'not_found' });
      updateAccountField(id, 3, sha256hex(String(e.parameter.h || '') + AUTH_SALT_()));
      return jsonResponse({ ok: true });
    }
    if (action === 'orgToggle') {
      const id = String(e.parameter.orgId || '');
      const acc = findAccount(id);
      if (!acc) return jsonResponse({ ok: false, error: 'not_found' });
      if (id === MASTER_ID) return jsonResponse({ ok: false, error: 'cannot_disable_master' });
      updateAccountField(id, 4, acc.active ? 'N' : 'Y');
      return jsonResponse({ ok: true });
    }
  }

  // ── 단건 조회 (서류 업로드 화면용, 로그인 없음) ──
  // 예전에는 전체 목록을 내려받아 화면에서 찾았다. 링크만 있으면 전 고객
  // 정보가 노출됐으므로 해당 고객 한 건만 반환한다.
  if (action === 'getCustomer') {
    const id = String(e.parameter.id || '');
    if (!id) return jsonResponse({ ok: false, error: 'no_id' });
    const rows = getAllCustomers();
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i].id) === id) return jsonResponse({ ok: true, data: rows[i] });
    }
    return jsonResponse({ ok: false, error: 'not_found' });
  }

  // ── 기본: 고객 목록 (로그인 필요, 업체별로 걸러서) ──
  // 토큰이 없으면 아무것도 주지 않는다. 마스터만 전체를 본다.
  try {
    const auth = verifyToken(e.parameter && e.parameter.token);
    if (!auth) return jsonResponse({ ok: false, error: 'unauthorized', data: [] });

    let rows = getAllCustomers();
    if (!auth.isMaster) {
      rows = rows.filter(function (r) { return String(r.agentOrg || '') === auth.orgName; });
    }
    return jsonResponse({ ok: true, data: rows, orgName: auth.orgName, isMaster: auth.isMaster });
  } catch(err) {
    return jsonResponse({ ok: false, error: String(err), data: [] });
  }
}


function doPost(e) {
  if (!e.postData || !e.postData.contents) {
    return jsonResponse({ ok: false, error: 'no_body' });
  }

  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch(err) {
    return jsonResponse({ ok: false, error: 'invalid_json' });
  }

  const sheet = getSheet();

  if (body.action === 'save') {
    const c = body.customer;
    sheet.appendRow([
      c.id, c.name, c.birth, c.phone, c.product || '', c.amount || '',
      c.period, c.debit, c.status, String(c.confirmed || false), c.createdAt,
      c.docs ? JSON.stringify(c.docs) : '', c.agentName || '', c.agentOrg || ''
    ]);
    invalidateCache();
    try { notifyManager(c); } catch(err) {}

  } else if (body.action === 'update') {
    updateRow(sheet, body.customer);
    invalidateCache();

  } else if (body.action === 'delete') {
    deleteRow(sheet, body.id);
    invalidateCache();

  } else if (body.action === 'contact') {
    try { sendContactEmail(body); } catch(err) {}

  } else if (body.action === 'uploadOneFile') {
    const url = uploadFileToDrive(body.data, body.name, body.type, body.customerId, body.docKey);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(body.customerId)) {
          let docs = {};
          try { if (allData[i][11]) docs = JSON.parse(allData[i][11]); } catch(err) {}
          docs[body.docKey] = url;
          sheet.getRange(i + 1, 12).setValue(JSON.stringify(docs));
          if (docsComplete(docs)) {
            sheet.getRange(i + 1, 9).setValue('서류제출');
          }
          break;
        }
      }
    } finally {
      lock.releaseLock();
    }
    invalidateCache();

  } else if (body.action === 'saveDocUrl') {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const allData = sheet.getDataRange().getValues();
      for (let i = 1; i < allData.length; i++) {
        if (String(allData[i][0]) === String(body.customerId)) {
          let docs = {};
          try { if (allData[i][11]) docs = JSON.parse(allData[i][11]); } catch(err) {}
          docs[body.docKey] = body.url;
          sheet.getRange(i + 1, 12).setValue(JSON.stringify(docs));
          if (docsComplete(docs)) {
            sheet.getRange(i + 1, 9).setValue('서류제출');
          }
          break;
        }
      }
    } finally {
      lock.releaseLock();
    }
    invalidateCache();
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

// ── 파트너 문의 이메일 ────────────────────────────
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

// ── 담당자 알림 ──────────────────────────────────
function notifyManager(customer) {
  sendEmail(customer);
  if (SOLAPI_API_KEY && SENDER_PHONE) {
    try { sendKakaoOrSMS(customer); } catch(e) {}
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
      to: MANAGER_PHONE, from: SENDER_PHONE,
      kakaoOptions: {
        pfId: KAKAO_PFID, templateId: KAKAO_TEMPLATE_ID,
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
      to: MANAGER_PHONE, from: SENDER_PHONE,
      text: '[티유디지털] 신용조회 요청\n고객: ' + customer.name +
            '\n연락처: ' + customer.phone +
            '\n구독: ' + customer.period +
            '\n신청: ' + customer.createdAt
    };
  }

  UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', {
    method: 'POST',
    headers: { 'Authorization': authHeader, 'Content-Type': 'application/json' },
    payload: JSON.stringify({ message }),
    muteHttpExceptions: true
  });
}

function computeHmac(data, secret) {
  return Utilities.computeHmacSha256Signature(data, secret)
    .map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ── Row 수정 / 삭제 ───────────────────────────────
function updateRow(sheet, c) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(c.id)) {
      sheet.getRange(i + 1, 1, 1, 14).setValues([[
        c.id, c.name, c.birth, c.phone, c.product || '', c.amount || '',
        c.period, c.debit, c.status, String(c.confirmed || false), c.createdAt,
        c.docs ? JSON.stringify(c.docs) : '', c.agentName || '', c.agentOrg || ''
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

// 최초 1회 실행 - 헤더 설정
function setupSheet() {
  const sheet   = getSheet();
  const headers = ['id','name','birth','phone','product','amount','period','debit','status','confirmed','createdAt','docs','agentName','agentOrg'];
  sheet.setName('고객목록');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  sheet.setFrozenRows(1);
}

// 최초 1회 실행 — 기존 데이터의 담당자/업체를 새 컬럼으로 옮긴다.
// 예전에는 createdAt 에 '2026. 8. 1. 오후 2:15 [이영자/현대창호]' 형태로 섞여 있었다.
function migrateAgentColumns() {
  const sh = getSheet();
  const data = sh.getDataRange().getValues();
  const head = data[0];
  let iName = head.indexOf('agentName'), iOrg = head.indexOf('agentOrg');
  if (iName === -1 || iOrg === -1) {
    sh.getRange(1, head.length + 1, 1, 2).setValues([['agentName', 'agentOrg']]).setFontWeight('bold');
    iName = head.length; iOrg = head.length + 1;
  }
  let moved = 0;
  for (let i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    if (data[i][iOrg]) continue;                    // 이미 채워진 건 건너뜀
    const m = String(data[i][10] || '').match(/\[([^\/\]]+)\/([^\]]+)\]/);
    if (!m) continue;
    sh.getRange(i + 1, iName + 1).setValue(m[1].trim());
    sh.getRange(i + 1, iOrg  + 1).setValue(m[2].trim());
    moved++;
  }
  invalidateCache();
  Logger.log('이관 완료: ' + moved + '건');
}
