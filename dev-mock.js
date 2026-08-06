/* ============================================================
   dev-mock.js — 로컬 테스트용 가짜 백엔드
   ------------------------------------------------------------
   localhost / 127.0.0.1 에서 열었을 때만 동작한다.
   실제 도메인(www.tudigital.co.kr)에서는 아무 일도 하지 않으므로
   배포되어도 운영에 영향이 없다.

   하는 일:
     · Apps Script(script.google.com) 호출을 가로채 localStorage로 처리
     · imgbb 업로드를 가로채 가짜 이미지 URL 반환
     · 구독절차 신규 규칙(견적서·가족관계증명서·소유형태)을 그대로 검증

   끄기      : 화면 우측 하단 배지 클릭 또는 콘솔에서 devMock.off()
   초기화    : devMock.reset()
   데이터보기: devMock.dump()
   ============================================================ */
(function () {
  'use strict';

  var host = location.hostname;
  var isLocal = host === 'localhost' || host === '127.0.0.1' || host === '';
  if (!isLocal) return;                       // 운영에서는 즉시 종료
  if (localStorage.getItem('devMockOff') === '1') { console.log('[devMock] 꺼져 있음'); return; }

  var KEY = 'tudigital_dev_customers';

  function load() {
    try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch (e) { return []; }
  }
  function save(list) { localStorage.setItem(KEY, JSON.stringify(list)); }

  /* ── 신규 구독절차 서류 규칙 (Code.gs 와 동일하게 유지) ── */
  // ownership 이 없으면 구버전 화면 → 예전 3종 규칙 (Code.gs 와 동일한 하위호환)
  function requiredDocs(ownership) {
    if (!ownership) return ['doc1', 'doc2', 'doc3'];
    return ownership === 'family'
      ? ['doc1', 'doc2', 'doc3', 'doc4', 'doc5']
      : ['doc1', 'doc2', 'doc3', 'doc4'];
  }
  function docsComplete(docs) {
    docs = docs || {};
    return requiredDocs(docs.ownership).every(function (k) { return !!docs[k]; });
  }

  function json(obj) {
    return Promise.resolve(new Response(JSON.stringify(obj), {
      status: 200, headers: { 'Content-Type': 'application/json' }
    }));
  }

  /* ── GET 처리 ─────────────────────────────────── */
  function handleGet(url) {
    var p = url.searchParams;
    var action = p.get('action') || '';

    if (action === 'saveAllDocUrls') {
      var id = String(p.get('customerId') || '');
      var rawOwn = p.get('ownership') || '';
      var isLegacy = !rawOwn;
      var ownership = isLegacy ? '' : (rawOwn === 'family' ? 'family' : 'self');
      var list = load();
      var row = list.filter(function (c) { return String(c.id) === id; })[0];
      if (!row) return json({ ok: false, error: 'customer_not_found' });

      var incoming = {};
      ['doc1', 'doc2', 'doc3', 'doc4', 'doc5'].forEach(function (k) {
        if (p.get(k)) incoming[k] = p.get(k);
      });
      var missing = requiredDocs(ownership).filter(function (k) { return !incoming[k]; });
      if (missing.length) {
        console.warn('[devMock] 필수 서류 누락:', missing);
        return json({ ok: false, error: 'missing_docs', missing: missing });
      }

      row.docs = row.docs || {};
      Object.keys(incoming).forEach(function (k) { row.docs[k] = incoming[k]; });
      if (!isLegacy) {
        row.docs.ownership = ownership;
        if (ownership === 'self') delete row.docs.doc5;
      }
      if (p.get('jobType')) row.docs.jobType = p.get('jobType');
      row.status = '서류제출';
      save(list);
      console.log('[devMock] 서류 저장 완료 (' + (ownership === 'family' ? '가족 소유 5종' : '본인 소유 4종') + ')', row.docs);
      return json({ ok: true });
    }

    /* 4단계 · 시공완료 서류 (Code.gs saveConstructionDocs 와 동일 규칙) */
    if (action === 'saveConstructionDocs') {
      var cid = String(p.get('customerId') || '');
      var list2 = load();
      var r2 = list2.filter(function (c) { return String(c.id) === cid; })[0];
      if (!r2) return json({ ok: false, error: 'customer_not_found' });

      var cdoc1 = p.get('cdoc1') || '', cdoc2 = p.get('cdoc2') || '';
      var photos = String(p.get('photos') || '').split('|').filter(Boolean);
      if (!cdoc1 || !cdoc2) return json({ ok: false, error: 'missing_docs' });
      if (photos.length < 6) {
        console.warn('[devMock] 시공사진 부족:', photos.length + '/6');
        return json({ ok: false, error: 'not_enough_photos', need: 6, got: photos.length });
      }
      r2.docs = r2.docs || {};
      if (r2.docs.agreement !== 'done') return json({ ok: false, error: 'not_agreed' });

      r2.docs.cdoc1 = cdoc1;
      r2.docs.cdoc2 = cdoc2;
      r2.docs.photos = photos;
      r2.docs.constructionAt = new Date().toISOString();
      r2.status = '시공완료';
      save(list2);
      console.log('[devMock] 시공완료 서류 저장 — 사진 ' + photos.length + '장');
      return json({ ok: true });
    }

    if (action === 'proxyImage') return json({ ok: false });

    // 기본: 전체 고객 목록
    return json({ data: load() });
  }

  /* ── POST 처리 ────────────────────────────────── */
  function handlePost(bodyText) {
    var body = {};
    try { body = JSON.parse(bodyText || '{}'); } catch (e) {}
    var list = load();

    if (body.action === 'save' && body.customer) {
      list.push(body.customer);
      save(list);
      console.log('[devMock] 고객 저장:', body.customer.name, '(id=' + body.customer.id + ')');

    } else if (body.action === 'update' && body.customer) {
      var i = list.findIndex(function (c) { return String(c.id) === String(body.customer.id); });
      if (i !== -1) list[i] = body.customer; else list.push(body.customer);
      save(list);
      console.log('[devMock] 고객 수정:', body.customer.name, '→', body.customer.status);

    } else if (body.action === 'delete') {
      save(list.filter(function (c) { return String(c.id) !== String(body.id); }));
      console.log('[devMock] 고객 삭제:', body.id);

    } else if (body.action === 'uploadFile') {
      var url = mockFileUrl(body.docKey || 'doc');
      var j = list.findIndex(function (c) { return String(c.id) === String(body.customerId); });
      if (j !== -1) {
        list[j].docs = list[j].docs || {};
        list[j].docs[body.docKey] = url;
        if (docsComplete(list[j].docs)) list[j].status = '서류제출';
        save(list);
      }
      return json({ ok: true, url: url });
    }
    return json({ ok: true });
  }

  var LABEL = { doc1: '등기부등본', doc2: '신분증', doc3: '통장사본', doc4: '견적서', doc5: '가족관계증명서' };
  function mockFileUrl(docKey) {
    return 'https://placehold.co/600x420/0d2137/ffffff?text='
      + encodeURIComponent((LABEL[docKey] || docKey) + ' (TEST)');
  }

  /* ── fetch 가로채기 ───────────────────────────── */
  var origFetch = window.fetch.bind(window);
  window.fetch = function (input, init) {
    var raw = (typeof input === 'string') ? input : (input && input.url) || '';
    init = init || {};

    // imgbb 업로드 → 가짜 이미지 URL
    if (raw.indexOf('api.imgbb.com') !== -1) {
      var key = 'doc';
      try {
        if (init.body && init.body.get) {
          var n = init.body.get('name');
          if (n) key = n;
        }
      } catch (e) {}
      console.log('[devMock] imgbb 업로드 가로챔');
      return json({ success: true, data: { url: mockFileUrl(key) } });
    }

    // Apps Script 호출
    if (raw.indexOf('script.google.com/macros') !== -1) {
      var method = (init.method || 'GET').toUpperCase();
      if (method === 'POST') return handlePost(init.body);
      var u;
      try { u = new URL(raw); } catch (e) { u = new URL(raw, location.origin); }
      return handleGet(u);
    }

    return origFetch(input, init);
  };

  /* ── 화면 배지 ────────────────────────────────── */
  function badge() {
    if (!document.body) { return setTimeout(badge, 100); }
    var n = load().length;
    var el = document.getElementById('dev-mock-badge');
    if (!el) {
      el = document.createElement('div');
      el.id = 'dev-mock-badge';
      el.style.cssText = 'position:fixed;right:14px;bottom:14px;z-index:99999;' +
        'background:#c0392b;color:#fff;font:700 12px/1.4 "맑은 고딕",sans-serif;' +
        'padding:9px 14px;border-radius:20px;box-shadow:0 3px 12px rgba(0,0,0,.3);' +
        'cursor:pointer;user-select:none;';
      // 실수로 꺼지지 않도록 클릭으로는 끄지 않는다. 끄려면 콘솔에서 devMock.off()
      el.title = '로컬 테스트 모드 — 저장은 이 브라우저에만 됩니다.\n끄려면 콘솔에서 devMock.off()';
      el.onclick = function () {
        console.log('[devMock] 저장된 데이터:'); window.devMock.dump();
        alert('로컬 테스트 모드입니다.\n모든 저장은 이 브라우저에만 되고 운영 서버에는 반영되지 않습니다.\n\n끄시려면 콘솔에 devMock.off() 를 입력하세요.');
      };
      document.body.appendChild(el);
    }
    el.textContent = '🧪 TEST 모드 · 고객 ' + n + '건';
  }
  badge();
  setInterval(badge, 1500);

  /* ── 콘솔 유틸 ────────────────────────────────── */
  window.devMock = {
    dump:  function () { console.table(load()); return load(); },
    reset: function () { localStorage.removeItem(KEY); console.log('[devMock] 초기화 완료'); location.reload(); },
    off:   function () { localStorage.setItem('devMockOff', '1'); location.reload(); },
    on:    function () { localStorage.removeItem('devMockOff'); location.reload(); },
    seed:  function () {                       // 테스트용 고객 1건 즉시 생성
      var list = load();
      var c = {
        id: Date.now(), name: '테스트고객', birth: '19830214', phone: '010-1234-5678',
        product: '창호', amount: '1500000', period: '60개월', debit: '25일',
        status: '신용조회중', confirmed: false,
        createdAt: new Date().toLocaleString('ko-KR') + ' [테스트]', docs: null
      };
      list.push(c); save(list);
      console.log('[devMock] 테스트 고객 생성 — upload.html?id=' + c.id);
      return c.id;
    }
  };

  console.log('%c[devMock] 로컬 테스트 모드 ON', 'color:#c0392b;font-weight:bold');
  console.log('  devMock.seed()  테스트 고객 생성 →  devMock.dump()  데이터 보기');
  console.log('  devMock.reset() 초기화          →  devMock.off()   목 모드 끄기');
})();
