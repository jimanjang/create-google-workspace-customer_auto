/***************
 * 기본 설정
 ***************/
const SHEET_NAME = 'Provisioning';
const SKU_MAP_SHEET = 'SKU_MAP';
const DEFAULT_ALT_EMAIL = 'laika.jang@netkillersoft.com';

/***************
 * SKU_MAP 보장/로드 (skuName → skuId 변환용)
 ***************/
function ensureSkuMapSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(SKU_MAP_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SKU_MAP_SHEET);
    sheet.getRange(1,1,1,2).setValues([['skuName','skuId']]);
    sheet.getRange(2,1,1,2).setValues([['Business Starter','1010020027']]);
    sheet.autoResizeColumns(1,2);
  }
  return sheet;
}
function loadSkuMap_() {
  const sheet = ensureSkuMapSheet_();
  const values = sheet.getDataRange().getValues();
  const map = new Map();
  if (values.length < 2) return map;
  const headers = values[0].map(String);
  const idxName = headers.indexOf('skuName');
  const idxId = headers.indexOf('skuId');
  for (let r=1; r<values.length; r++) {
    const name = String(values[r][idxName]||'').trim();
    const id = String(values[r][idxId]||'').trim();
    if (name && id) map.set(name.toLowerCase(), id);
  }
  return map;
}

/***************
 * 시트 로드
 ***************/
function loadRows_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_NAME}`);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { sheet, headers: [], rows: [] };
  const headers = values[0].map(h => String(h).trim());
  const rows = [];
  for (let r=1; r<values.length; r++) {
    const obj = {};
    for (let c=0; c<headers.length; c++) obj[headers[c]] = values[r][c];
    obj.__rowIndex = r+1;
    rows.push(obj);
  }
  return { sheet, headers, rows };
}

/***************
 * (추가) 결과 컬럼 보장 + 결과 기록
 ***************/
function ensureResultColumns_(sheet, headers) {
  const need = ['customerId','subscriptionId','currentPlan','currentStatus','trialEndTime'];
  const headerSet = new Set(headers);
  let changed = false;
  need.forEach(h => {
    if (!headerSet.has(h)) {
      sheet.getRange(1, headers.length + 1, 1, 1).setValue(h);
      headers.push(h);
      headerSet.add(h);
      changed = true;
    }
  });
  if (changed) sheet.autoResizeColumns(1, headers.length);
  return headers;
}

function updateRow_(sheet, headers, rowIndex, patch) {
  const headerMap = new Map(headers.map((h, i) => [h, i]));
  const lastCol = headers.length;
  const rowValues = sheet.getRange(rowIndex, 1, 1, lastCol).getValues()[0];
  Object.keys(patch).forEach(key => {
    if (!headerMap.has(key)) return; // 없는 헤더는 무시
    rowValues[headerMap.get(key)] = patch[key];
  });
  sheet.getRange(rowIndex, 1, 1, lastCol).setValues([rowValues]);
}

function writeProvisioningResult_(sheet, headers, rowIndex, customerId, sub) {
  ensureResultColumns_(sheet, headers);
  const patch = {
    customerId: customerId || '',
    subscriptionId: (sub && sub.subscriptionId) || '',
    currentPlan: (sub && sub.plan && sub.plan.planName) || '',
    currentStatus: (sub && sub.status) || '',
    trialEndTime: (sub && sub.trialSettings && sub.trialSettings.trialEndTime) || ''
  };
  updateRow_(sheet, headers, rowIndex, patch);
}

/***************
 * 행 → CONFIG 변환기
 ***************/
function makeConfigFromRow_(row, skuMap) {
  const customerDomain = String(row.customerDomain || '').trim();
  let skuId = String(row.skuId || '').trim();
  const skuName = String(row.skuName || '').trim();
  const planNameRaw = String(row.planName || '').trim();
  const planName = planNameRaw ? planNameRaw.toUpperCase() : 'TRIAL';
  const seats = Number(row.seats || 1);

  if (!skuId && skuName) {
    const mapped = skuMap.get(skuName.toLowerCase());
    if (!mapped) throw new Error(`SKU_MAP에 '${skuName}' 매핑이 없습니다.`);
    skuId = mapped;
  }
  if (!customerDomain) throw new Error('customerDomain 누락');
  if (!skuId) throw new Error('skuId 누락 (또는 skuName 매핑 실패)');
  if (!Number.isFinite(seats) || seats <= 0) throw new Error('seats는 1 이상의 정수');

  const primaryEmail = String(row.primaryEmail || '').trim();
  const givenName   = String(row.givenName   || '').trim();
  const familyName  = String(row.familyName  || '').trim();
  const password    = String(row.password    || '').trim();
  const manageCustomerUsers = !!(primaryEmail && givenName && familyName && password);

  // 🔤 언어 코드 (R열) 읽기
  //   - 헤더가 'language' 또는 '언어' 인 경우를 모두 지원
  //   - 비어 있으면 기본값 'ko'
  const languageRaw = String(row.language || row['언어'] || '').trim();
  const languageCode = languageRaw || 'ko';

  const delegatedAdmin = DEFAULT_ALT_EMAIL;

  return {
    customerDomain,
    delegatedAdmin,
    skuId,
    planName,
    seats,
    method: 'DNS_CNAME',
    manageCustomerUsers,
    languageCode, // 🔤 cfg에 언어 추가
    __admin: { primaryEmail, givenName, familyName, password }
  };
}

/***************
 * (수정) 플랜명 정규화 (Reseller API 허용 값)
 * - 이미 정규화된 값 그대로 허용
 * - 널리 쓰는 별칭 매핑
 * - 알 수 없는 값은 에러
 ***************/
function normalizePlanName_(raw) {
  const p = String(raw || '').trim().toUpperCase();

  // 이미 정규화된 값은 그대로 허용
  if (p === 'ANNUAL_MONTHLY_PAY' || p === 'ANNUAL_YEARLY_PAY' || p === 'FLEXIBLE' || p === 'TRIAL') {
    return p;
  }
  // 별칭/오타 매핑
  if (p === 'ANNUAL' || p === 'ANNUAL_MONTHLY' || p === 'ANNUAL-MONTHLY') return 'ANNUAL_MONTHLY_PAY';
  if (p === 'ANNUAL_YEARLY' || p === 'ANNUAL-YEARLY') return 'ANNUAL_YEARLY_PAY';
  if (p === 'FLEX') return 'FLEXIBLE';

  throw new Error(`알 수 없는 planName: "${raw}" → 허용값: TRIAL | FLEXIBLE | ANNUAL_MONTHLY_PAY | ANNUAL_YEARLY_PAY`);
}

/***************
 * (추가) 테넌트(고객) 기본 언어 설정
 ***************/
function setCustomerLanguage_(customerId, languageCode) {
  if (!customerId || !languageCode) {
    Logger.log('setCustomerLanguage_: customerId 또는 languageCode 누락 → 스킵');
    return;
  }

  try {
    const body = {
      language: languageCode // 예: 'ko', 'en', 'ja', 'zh-CN'
    };
    const result = AdminDirectory.Customers.update(body, customerId);
    Logger.log(`🌐 고객 기본 언어 설정 완료: ${customerId} → ${languageCode}`);
    return result;
  } catch (e) {
    Logger.log(`❌ 고객 기본 언어 설정 실패 (${customerId}, lang=${languageCode}): ${String(e)}`);
  }
}

/***************
 * cfg 한 번 실행
 ***************/
function runProvisioningOnce_(cfg) {
  try {
    const token = sv_getToken_byCfg_(cfg);
    Logger.log('Place this token: ' + token.token);
  } catch (e) {
    Logger.log('Site Verification 토큰 발급 스킵/실패: ' + String(e));
  }

  const customer = ensureCustomer_byCfg_(cfg);
  const customerId = customer.customerId;

  // 🔤 테넌트 기본 언어 설정
  if (cfg.languageCode) {
    setCustomerLanguage_(customerId, cfg.languageCode);
  }

  const sub = createSubscriptionIfAbsent_byCfg_(customerId, cfg);

  if (cfg.manageCustomerUsers) {
    const a = cfg.__admin;
    const userReq = {
      primaryEmail: a.primaryEmail || ('admin@' + cfg.customerDomain),
      name: { givenName: a.givenName || 'First', familyName: a.familyName || 'Admin' },
      password: a.password || (Utilities.getUuid().slice(0,10) + 'Aa!')
    };

    // 🔤 사용자 UI 언어 설정 (R열 기준)
    if (cfg.languageCode) {
      userReq.languages = [{
        languageCode: cfg.languageCode, // 예: 'ko', 'en', 'ja', 'zh-CN'
        preference: 'preferred'
      }];
      Logger.log('언어 설정: ' + cfg.languageCode);
    }

    const user = AdminDirectory.Users.insert(userReq);
    Logger.log('✅ 사용자 생성 완료: ' + user.primaryEmail);
    AdminDirectory.Users.makeAdmin({ status: true }, user.primaryEmail);
    Logger.log('⭐ 관리자 권한 부여 완료: ' + user.primaryEmail);
  } else {
    Logger.log('관리자 생성 스킵 (manageCustomerUsers=false)');
  }

  Logger.log({ customerId, sub });
  return { customerId, sub };
}

/***************
 * 고객 확인/생성 (cfg)
 ***************/
function ensureCustomer_byCfg_(cfg) {
  try {
    const existing = AdminReseller.Customers.get(cfg.customerDomain);
    Logger.log('기존 고객 존재: ' + existing.customerId);
    return existing;
  } catch (e) {
    const msg = String(e);
    if (msg.includes('Not Found')) {
      const req = {
        customerDomain: cfg.customerDomain,
        alternateEmail: cfg.delegatedAdmin,
        postalAddress: {
          contactName: cfg.customerDomain,
          organizationName: cfg.customerDomain,
          region: 'KR',
          postalCode: '06182',
          countryCode: 'KR',
          addressLine1: '영동대로 417'
        }
      };
      const created = AdminReseller.Customers.insert(req);
      Logger.log('신규 고객 생성 완료: ' + created.customerId);
      return created;
    }
    throw e;
  }
}

/***************
 * 핵심: TRIAL 먼저 시도 → 실패 시 FLEXIBLE 폴백
 * - FLEXIBLE은 seats 전송 금지
 ***************/
function createSubscriptionIfAbsent_byCfg_(customerId, cfg) {
  const seats = Number(cfg.seats || 1);
  if (!Number.isFinite(seats) || seats <= 0) {
    throw new Error('유효한 seats(1 이상의 정수)가 필요합니다.');
  }

  // 1) TRIAL
  const trialBody = {
    customerId,
    skuId: cfg.skuId,
    plan: { planName: 'TRIAL' },
    seats: { maximumNumberOfSeats: seats }
  };

  try {
    const trial = AdminReseller.Subscriptions.insert(trialBody, customerId);
    Logger.log('✅ TRIAL 구독 생성 완료: ' + JSON.stringify(trial));
    return trial;
  } catch (e) {
    const msg = String(e);
    if (msg.includes('already exists') || msg.includes('Conflict')) {
      const list = AdminReseller.Subscriptions.list(customerId);
      Logger.log('⚠️ 기존 구독 사용: ' + JSON.stringify(list.subscriptions || []));
      return (list.subscriptions && list.subscriptions[0]) || null;
    }
    Logger.log('ℹ️ TRIAL 생성 불가. FLEXIBLE로 폴백 시도: ' + msg);
  }

  // 2) FLEXIBLE 폴백 (seats 금지)
  const flexBody = {
    customerId,
    skuId: cfg.skuId,
    plan: { planName: 'FLEXIBLE' }
  };

  try {
    const flex = AdminReseller.Subscriptions.insert(flexBody, customerId);
    Logger.log('✅ FLEXIBLE 구독 생성 완료(폴백): ' + JSON.stringify(flex));
    return flex;
  } catch (e2) {
    const msg2 = String(e2);
    if (msg2.includes('already exists') || msg2.includes('Conflict')) {
      const list = AdminReseller.Subscriptions.list(customerId);
      Logger.log('⚠️ 기존 구독 사용(폴백 경로): ' + JSON.stringify(list.subscriptions || []));
      return (list.subscriptions && list.subscriptions[0]) || null;
    }
    throw e2;
  }
}

/***************
 * (추가) 도메인으로 customerId / subscriptionId 조회
 ***************/
function findIdsByDomain_(domain, skuIdOptional) {
  const customer = AdminReseller.Customers.get(domain); // 없으면 예외
  const customerId = customer.customerId;

  const list = AdminReseller.Subscriptions.list(customerId);
  let sub = null;
  if (list && list.subscriptions && list.subscriptions.length) {
    if (skuIdOptional) {
      sub = list.subscriptions.find(s => String(s.skuId) === String(skuIdOptional)) || null;
    }
    if (!sub) sub = list.subscriptions[0];
  }
  return {
    customerId,
    subscriptionId: sub ? sub.subscriptionId : '',
    subscription: sub || null
  };
}

/***************
 * Site Verification (cfg)
 ***************/
function getOAuthService_() {
  return OAuth2.createService('siteverification')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://oauth2.googleapis.com/token')
    .setClientId(PropertiesService.getScriptProperties().getProperty('CLIENT_ID'))
    .setClientSecret(PropertiesService.getScriptProperties().getProperty('CLIENT_SECRET'))
    .setCallbackFunction('authCallback')
    .setScope('https://www.googleapis.com/auth/siteverification')
    .setParam('access_type', 'offline')
    .setPropertyStore(PropertiesService.getUserProperties());
}

function sv_fetch_(url, payload) {
  const service = getOAuthService_();
  if (!service.hasAccess()) {
    const authUrl = service.getAuthorizationUrl();
    throw new Error('Site Verification 권한이 없습니다. 이 URL로 승인해 주세요:\n' + authUrl);
  }
  const res = UrlFetchApp.fetch(url, {
    method: payload ? 'post' : 'get',
    muteHttpExceptions: true,
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + service.getAccessToken() },
    payload: payload ? JSON.stringify(payload) : null,
  });
  const code = res.getResponseCode();
  const bodyText = res.getContentText() || '{}';
  let body;
  try { body = JSON.parse(bodyText); } catch (e) { body = { raw: bodyText }; }
  if (code >= 400) {
    throw new Error('SV API 호출 실패 (' + code + '): ' + bodyText);
  }
  return { code, body };
}

function authCallback(request) {
  const service = getOAuthService_();
  const isAuthorized = service.handleCallback(request);
  if (isAuthorized) {
    return HtmlService.createHtmlOutput('승인되었습니다. 이 창을 닫고 다시 실행하세요.');
  } else {
    return HtmlService.createHtmlOutput('승인이 취소되었습니다.');
  }
}

function sv_getToken_byCfg_(cfg) {
  const identifier = cfg.customerDomain;

  // cfg.method 값 정규화 (없으면 기본 DNS_TXT)
  let method = String(cfg.method || 'DNS_TXT').toUpperCase();
  const allowed = ['DNS_TXT', 'DNS_CNAME', 'META', 'FILE'];

  if (!allowed.includes(method)) {
    method = 'DNS_TXT';
  }

  const body = {
    site: {
      type: 'INET_DOMAIN',
      identifier
    },
    verificationMethod: method
  };

  return sv_fetch_('https://www.googleapis.com/siteVerification/v1/token', body).body;
}


/***************
 * (수정) TRIAL/FLEX에서 연간 약정으로 플랜만 지정
 * - Apps Script Advanced Service 시그니처: (resource, customerId, subscriptionId)
 ***************/
function setPlanForTrialOrFlex_(customerId, subscriptionId, targetPlanNameRaw, seatsRaw) {
  const targetPlanName = normalizePlanName_(targetPlanNameRaw);
  Logger.log(`setPlanForTrialOrFlex_: normalized planName = ${targetPlanName}`);

  // changePlan은 연간 약정(ANNUAL_*)로의 업데이트에 사용
  if (targetPlanName !== 'ANNUAL_MONTHLY_PAY' && targetPlanName !== 'ANNUAL_YEARLY_PAY') {
    throw new Error(`changePlan은 연간 약정(ANNUAL_*) 전환에만 사용합니다. 입력="${targetPlanNameRaw}", 정규화="${targetPlanName}"`);
  }

  const n = Number(seatsRaw || 1);
  if (!Number.isFinite(n) || n <= 0) {
    throw new Error('ANNUAL 전환에는 seats(정수 > 0)가 필요합니다.');
  }

  // ChangePlanRequest — renewalSettings 포함 금지
  const body = {
    planName: targetPlanName,
    seats: { numberOfSeats: n }
    // purchaseOrderId, dealCode 필요 시 추가
  };

  try {
    // ✅ 올바른 순서
    const result = AdminReseller.Subscriptions.changePlan(body, customerId, subscriptionId);
    Logger.log(`🔁 changePlan 완료: ${subscriptionId} → ${targetPlanName} (seats=${n})`);
    Logger.log(JSON.stringify(result));
    return result;
  } catch (e) {
    Logger.log(`❌ changePlan 실패 (${subscriptionId}): ${String(e)}`);
    throw e;
  }
}

/***************
 * (신규) TRIAL을 즉시 유료로 전환
 * - changePlan으로 결제 플랜을 지정한 뒤 호출
 ***************/
function startPaidService_(customerId, subscriptionId) {
  try {
    const result = AdminReseller.Subscriptions.startPaidService(customerId, subscriptionId);
    Logger.log(`🚀 startPaidService 완료: ${subscriptionId}`);
    return result;
  } catch (e) {
    Logger.log(`❌ startPaidService 실패 (${subscriptionId}): ${String(e)}`);
    throw e;
  }
}

/***************
 * (신규) 갱신 유형 설정
 * - Apps Script 시그니처: (resource, customerId, subscriptionId)
 ***************/
function setRenewalType_(customerId, subscriptionId, renewalTypeRaw) {
  const renewalType = String(renewalTypeRaw || 'AUTO_RENEW').toUpperCase();
  const body = { renewalType };
  try {
    const result = AdminReseller.Subscriptions.changeRenewalSettings(body, customerId, subscriptionId);
    Logger.log(`🔧 changeRenewalSettings 완료: ${subscriptionId} → ${renewalType}`);
    return result;
  } catch (e) {
    Logger.log(`❌ changeRenewalSettings 실패 (${subscriptionId}): ${String(e)}`);
    throw e;
  }
}

/***************
 * 메뉴
 ***************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Netkiller')
    .addItem('선택된 행을 CONFIG로 실행', 'runSelectedRowsWithConfig')
    .addItem('시트 전체를 CONFIG로 실행', 'runAllRowsWithConfig')
    .addItem('전체 구독 전환 실행', 'runChangePlanForAllRows')
    .addItem('선택된 행 설정 안내 메일 발송','sendSetupMailsForSelectedRows') // ← 추가
    .addToUi();
}


/***************
 * 실행기: 선택/전체 (프로비저닝 직후 결과 기록 추가)
 ***************/
function runSelectedRowsWithConfig() {
  Logger.log('=== ▶ 선택 영역 CONFIG 실행 시작 ===');
  const skuMap = loadSkuMap_();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error(`시트를 찾을 수 없습니다: ${SHEET_NAME}`);

  const sel = sheet.getActiveRange();
  const start = sel.getRow();
  const n = sel.getNumRows();

  // 헤더 확보 & 결과 컬럼 보장
  const headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(h=>String(h).trim());
  ensureResultColumns_(sheet, headers);

  for (let i=0; i<n; i++) {
    const rowIndex = start + i;
    if (rowIndex === 1) continue; // 헤더 스킵
    const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const row = { __rowIndex: rowIndex };
    headers.forEach((h,idx)=> row[h] = values[idx]);

    const ctx = `[row ${rowIndex} | ${row.customerDomain}]`;
    try {
      const cfg = makeConfigFromRow_(row, skuMap);
      Logger.log(`${ctx} CONFIG 실행 시작: skuId=${cfg.skuId}, plan=${cfg.planName}, seats=${cfg.seats}, lang=${cfg.languageCode}`);
      const res = runProvisioningOnce_(cfg);

      // ✅ 결과 시트 기록
      writeProvisioningResult_(sheet, headers, rowIndex, res.customerId, res.sub);

      Logger.log(`${ctx} ✅ 완료`);
    } catch (e) {
      Logger.log(`${ctx} ❌ 실패: ${String(e)}`);
    }
  }
  Logger.log('=== ✅ 선택 영역 CONFIG 실행 종료 ===');
}

function runAllRowsWithConfig() {
  Logger.log('=== ▶ 전체 시트 CONFIG 실행 시작 ===');
  const skuMap = loadSkuMap_();
  const { sheet, headers, rows } = loadRows_();
  if (!rows.length) { Logger.log('행 없음. 종료'); return; }

  // 결과 컬럼 보장
  ensureResultColumns_(sheet, headers);

  rows.forEach(row => {
    const ctx = `[row ${row.__rowIndex} | ${row.customerDomain}]`;
    try {
      const cfg = makeConfigFromRow_(row, skuMap);
      Logger.log(`${ctx} CONFIG 실행 시작: skuId=${cfg.skuId}, plan=${cfg.planName}, seats=${cfg.seats}, lang=${cfg.languageCode}`);
      const res = runProvisioningOnce_(cfg);

      // ✅ 결과 시트 기록
      writeProvisioningResult_(sheet, headers, row.__rowIndex, res.customerId, res.sub);

      Logger.log(`${ctx} ✅ 완료`);
    } catch (e) {
      Logger.log(`${ctx} ❌ 실패: ${String(e)}`);
    }
  });
  Logger.log('=== ✅ 전체 시트 CONFIG 실행 종료 ===');
}

/***************
 * TRIAL → 대상 플랜(ANNUAL_*) 전환 자동 실행
 * - ID 없으면 도메인으로 자동 보완 후 전환
 * - TRIAL이면 changePlan 후 startPaidService 호출
 * - 갱신유형은 changeRenewalSettings로 반영
 ***************/
function runChangePlanForAllRows() {
  Logger.log('=== ▶ 전체 구독 전환 실행 시작 ===');
  const { sheet, headers, rows } = loadRows_();
  if (!rows.length) { Logger.log('행 없음. 종료'); return; }

  // 결과 컬럼 보장
  ensureResultColumns_(sheet, headers);

  rows.forEach(row => {
    const ctx = `[row ${row.__rowIndex} | ${row.customerDomain}]`;
    try {
      const planNameRaw = String(row.planName || '').trim();
      const targetPlan = planNameRaw ? planNameRaw : 'FLEXIBLE';  // 기본값 유지(단, changePlan 대상은 ANNUAL_*)
      const renewalType = String(row.renewalType || 'AUTO_RENEW').trim();

      if (!targetPlan || targetPlan.toUpperCase() === 'TRIAL') {
        Logger.log(`${ctx} ⚙️ planName=TRIAL → 전환 대상 아님 (스킵)`);
        return;
      }

      let customerId = String(row.customerId || '').trim();
      let subscriptionId = String(row.subscriptionId || '').trim();

      // ✅ ID 자동 보완
      if (!customerId || !subscriptionId) {
        Logger.log(`${ctx} ID 누락 → 도메인으로 보완 시도`);
        const skuId = row.skuId ? String(row.skuId).trim() : '';
        const found = findIdsByDomain_(String(row.customerDomain).trim(), skuId || null);
        customerId = found.customerId || customerId;
        subscriptionId = found.subscriptionId || subscriptionId;

        // 보완 결과 시트 기록
        writeProvisioningResult_(sheet, headers, row.__rowIndex, customerId, found.subscription);
      }

      if (!customerId || !subscriptionId) {
        Logger.log(`${ctx} ❌ customerId/subscriptionId 여전히 누락 - 스킵`);
        return;
      }

      Logger.log(`${ctx} 플랜 전환 시도: target=${targetPlan}`);

      // 현재 플랜 조회
      const was = AdminReseller.Subscriptions.get(customerId, subscriptionId);
      const wasPlan = (was && was.plan && was.plan.planName) || '';
      const wasTrial = String(wasPlan).toUpperCase() === 'TRIAL';

      // 1) (필수) ANNUAL_*로 플랜 지정
      setPlanForTrialOrFlex_(customerId, subscriptionId, targetPlan, row.seats);

      // 2) TRIAL 이었다면 즉시 유료 전환
      if (wasTrial) {
        startPaidService_(customerId, subscriptionId);
      }

      // 3) 갱신 유형 적용 (AUTO_RENEW 등)
      setRenewalType_(customerId, subscriptionId, renewalType);

      Logger.log(`${ctx} ✅ 전환 완료`);

      // 최신 상태 반영
      try {
        const sub = AdminReseller.Subscriptions.get(customerId, subscriptionId);
        writeProvisioningResult_(sheet, headers, row.__rowIndex, customerId, sub);
      } catch (_) {}
    } catch (e) {
      Logger.log(`${ctx} ❌ 전환 실패: ${String(e)}`);
    }
  });

  Logger.log('=== ✅ 전체 구독 전환 실행 종료 ===');
}

function buildSetupMailBodyFromRow_(row) {
  const id    = String(row.primaryEmail || '').trim();
  const pw    = String(row.password     || '').trim();
  const host  = String(row.host         || '').trim();
  const value = String(row.value        || '').trim();

  return `
<div style="font-family:Apple SD Gothic Neo,Roboto,Arial,sans-serif; font-size:14px; line-height:1.6;">
안녕하세요,<br><br>
넷킬러 고객지원팀 입니다.<br><br>

귀사의 성공적인 Google Workspace 도입을 위해 필요한 활성화 절차를 안내드립니다.<br><br>

<b>1. 관리자 계정 정보</b><br>
ID : <b>${id}</b><br>
임시비밀번호 : <b>${pw}</b><br>
관리자 접속 URL : 
<a href="https://admin.google.com/" target="_blank">https://admin.google.com/</a><br>
구축이 완료될 때까지 위 임시 비밀번호를 사용하시고, 구축 완료 후에는 반드시 비밀번호를 변경해 주세요.<br><br>

<b>2. GWS 활성화를 위한 DNS 설정 (필수)</b><br>
Google Workspace 서비스를 정상적으로 이용하시려면 도메인 DNS 레코드를 변경해 주셔야 합니다.<br>
아래 설정을 완료하신 뒤 회신해 주시면 추가 지원을 도와드리겠습니다.<br>
<i>* Gmail 을 사용하지 않으실 경우 2~3단계는 생략하셔도 됩니다.</i><br><br>

<b>1단계: 도메인 소유권 확인 (CNAME 등록)</b><br>
Type : CNAME<br>
Host : <b>${host}</b><br>
TTL : 3600s (1hr)<br>
Value : <b>${value}</b><br><br>

<b>2단계: 메일 서버 설정 (MX 레코드 등록)</b><br>
메일 수신을 위해 MX 레코드를 아래 값으로 설정해 주세요.<br>
Type : MX<br>
Host : @ 또는 공란<br>
TTL : 3600s (1hr)<br>
Priority : 1<br>
Value : <b>smtp.google.com</b><br>
(공식 가이드: 
<a href="https://support.google.com/a/answer/174125" target="_blank">MX 레코드 설정 안내</a>)<br><br>

※ MX 레코드 변경 후 최대 48시간 동안 기존 서버와 병행 수신될 수 있습니다.<br>
업무 영향이 적은 <b>금요일 오후 변경</b>을 권장드립니다.<br><br>

<b>3단계: 스팸 방지 설정 (SPF 레코드 등록)</b><br>
Type : TXT<br>
Host : @ 또는 공란<br>
TTL : 3600s (1hr)<br>
Value : <b>v=spf1 include:_spf.google.com ~all</b><br><br>

<b>MSSP(보안 전문 지원) 서비스 안내</b><br>
넷킬러의 MSSP(보안 전문 지원) 서비스 를 이용하시면, DNS 설정 부터 Gmail·Drive 보안 관리 까지 전문가가 직접 체계적으로 관리해 드립니다.
담당자분께서 기술적인 부분을 신경쓰지 않으셔도 안심하고 편리하게 Google Workspace 를 운영하실 수 있도록 최적의 환경을 마련해 드리겠습니다. 관심이 있으시다면 편하게 회신 주십시오.
<br><br>

<b>3. 사용자 추가 및 관리</b><br>
필수 설정 완료 후 관리자는 아래 메뉴에서 사용자를 추가하고 관리할 수 있습니다.<br>
사용자 관리 페이지 : <a href="https://admin.google.com/ac/users" target="_blank">https://admin.google.com/ac/users</a><br>
사용자 추가방법 : <a href="https://support.google.com/a/answer/33310?hl=ko" target="_blank">Google 공식 가이드</a><br>
사용자 이메일 변경 방법 : <a href="https://support.google.com/a/answer/182084?hl=ko" target="_blank">Google 공식 주소 변경 가이드</a><br><br>

<b>4. 관리자를 위한 기타 권장사항</b><br>
최고 관리자 계정을 2개 이상 지정하여 보안과 업무 연속성을 확보하시는 것을 권장드립니다.
관리자 콘솔 및 사용자를 위한 학습자료는 아래 Netkiller 학습 센터를 참고하여 주십시오.
<br><br>
<b>사용자 학습 센터</b><br>
관리자용 : <a href="https://sites.google.com/netkiller.com/learning-center/%ED%99%88" target="_blank">Admin Learning Center</a><br>
사용자용 : <a href="https://sites.google.com/netkiller.com/learningcenter/%ED%99%88" target="_blank">User Learning Center</a><br><br>

추가 문의사항이 있으시면 언제든지 
<a href="mailto:support@netkiller.com">support@netkiller.com</a> 으로 연락 주세요.<br><br>

감사합니다.<br>
넷킬러 고객지원팀 드림.<br>
</div>
`;
}

function sendSetupMailForRow_(row) {
  const to = String(row.contactEmail || '').trim();
  const domain = String(row.customerDomain || '').trim();

  const ctx = `[row ${row.__rowIndex} | ${domain}]`;

  if (!to) {
    Logger.log(`${ctx} ❌ contactEmail 없음 → 메일 스킵`);
    return;
  }
  if (!domain) {
    Logger.log(`${ctx} ❌ customerDomain 없음 → 메일 스킵`);
    return;
  }

  const subject = `${domain} 의 GWS 구축을 위한 설정 안내`;

  // HTML 본문 생성
  const htmlBody = buildSetupMailBodyFromRow_(row);

  // 플레인텍스트 fallback (태그 제거해서 대충 뽑기)
  const plainText = htmlBody
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .trim();

  GmailApp.sendEmail(to, subject, plainText, {
    htmlBody: htmlBody,
    name: '넷킬러 고객지원팀',
    cc: 'support@netkiller.com',
    from: 'support@netkiller.com'  // alias 등록 돼 있어야 동작
  });

  Logger.log(`${ctx} ✅ 설정 안내 메일 발송 완료 → ${to}`);
}


function sendSetupMailsForSelectedRows() {
  Logger.log('=== ▶ 설정 안내 메일 발송(선택 영역) 시작 ===');

  const { sheet, headers, rows } = loadRows_();
  if (!rows.length) {
    Logger.log('행 없음. 종료');
    return;
  }

  const sel = sheet.getActiveRange();
  const start = sel.getRow();
  const end = start + sel.getNumRows() - 1;

  rows.forEach(row => {
    if (row.__rowIndex < start || row.__rowIndex > end) return;

    try {
      sendSetupMailForRow_(row);
    } catch (e) {
      const ctx = `[row ${row.__rowIndex} | ${row.customerDomain}]`;
      Logger.log(`${ctx} ❌ 메일 발송 실패: ${String(e)}`);
    }
  });

  Logger.log('=== ✅ 설정 안내 메일 발송(선택 영역) 종료 ===');
}
