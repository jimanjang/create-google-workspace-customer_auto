# Netkiller Google Workspace Provisioning Script

Google Workspace Reseller API와 Directory API를 이용해
**고객(테넌트) 생성 → 구독 생성(TRIAL 우선) → (옵션) 관리자 사용자 생성/승격 → 결과 기록 → (옵션) 플랜 전환/유료 시작/갱신설정 → 안내 메일 발송**
을 스프레드시트 기반으로 자동 처리하는 Apps Script입니다.

---

## 1. 주요 기능

### 1) 고객/구독 프로비저닝

* `Provisioning` 시트 각 행을 CONFIG로 변환하여 실행
* 고객이 없으면 생성, 있으면 재사용
* 구독은 **TRIAL 생성 시도 → 실패 시 FLEXIBLE로 폴백**

  * FLEXIBLE 폴백 시 seats 전송 금지 로직 포함

### 2) SKU 매핑 지원

* `skuId`가 없고 `skuName`만 있을 때
  `SKU_MAP` 시트에서 **skuName → skuId 자동 변환**
* SKU_MAP 시트가 없으면 자동 생성 + 기본값 삽입

### 3) 언어 설정(테넌트/사용자)

* `language` 또는 `언어` 컬럼을 읽어

  * 고객 기본 언어(`AdminDirectory.Customers.update`)
  * 생성 관리자 계정 UI 언어(`AdminDirectory.Users.insert.languages`)
    에 반영

### 4) 결과 자동 기록

프로비저닝/조회 결과를 시트에 자동 저장:

* `customerId`
* `subscriptionId`
* `currentPlan`
* `currentStatus`
* `trialEndTime`

### 5) 전체 플랜 전환 자동화

* TRIAL/FLEX → ANNUAL 전환(changePlan)
* TRIAL이면 전환 직후 startPaidService 호출
* renewalType(AUTO_RENEW 등) 설정

### 6) 설정 안내 메일 발송

선택된 고객 행에 대해 DNS/CNAME/MX/SPF + 관리자 계정 정보 포함 안내 메일 자동 발송

---

## 2. 사전 준비사항

### 2.1 Advanced Google Services 활성화

Apps Script 편집기에서
**Services(서비스)** → Advanced Google services 활성화:

* **Admin Directory API**
* **Google Workspace Reseller API**

또한 Google Cloud Console에서 동일 API가 활성화되어 있어야 합니다.

### 2.2 OAuth2 라이브러리 추가 (Site Verification)

Site Verification 토큰 발급 기능을 쓰려면 OAuth2 라이브러리를 추가합니다.

* 라이브러리 ID 예: `1B7...` (Google OAuth2 for Apps Script)
* 코드 상에서는 `OAuth2.createService(...)` 사용

### 2.3 Script Properties 설정

`CLIENT_ID`, `CLIENT_SECRET` 을 Script Properties에 저장해야 Site Verification이 동작합니다.

* Apps Script → Project Settings → Script properties

  * `CLIENT_ID`
  * `CLIENT_SECRET`

### 2.4 delegatedAdmin(대리 관리자) 계정

코드 상 기본값:

```js
const DEFAULT_ALT_EMAIL = 'laika.jang@netkillersoft.com';
```

* Reseller API 호출 가능한 Netkiller 파트너 계정이어야 함
* 고객 생성 시 alternateEmail로 등록됨

---

## 3. 스프레드시트 구조

### 3.1 필수 시트

* `Provisioning`
* `SKU_MAP` (없으면 자동 생성)

### 3.2 Provisioning 시트 헤더(필수/권장)

| 헤더             |            필수 | 설명                                 |
| -------------- | ------------: | ---------------------------------- |
| customerDomain |             ✅ | 고객 도메인 (예: example.com)            |
| skuId          | ✅(또는 skuName) | Workspace SKU ID                   |
| skuName        |   ✅(또는 skuId) | SKU 이름 (SKU_MAP으로 매핑)              |
| planName       |            권장 | TRIAL/FLEXIBLE/ANNUAL_* (기본 TRIAL) |
| seats          |            권장 | 좌석 수 (기본 1)                        |
| renewalType    |            선택 | AUTO_RENEW / CANCEL_AT_END 등       |
| language / 언어  |            선택 | 테넌트/사용자 기본 언어 (기본 ko)              |

#### 관리자 계정 생성 옵션 컬럼

아래 4개가 모두 있으면 **관리자 사용자 자동 생성 + 슈퍼관리자 승격**

| 헤더           | 설명          |
| ------------ | ----------- |
| primaryEmail | 생성할 관리자 이메일 |
| givenName    | 이름          |
| familyName   | 성           |
| password     | 임시 비밀번호     |

#### 안내 메일 발송 옵션 컬럼

| 헤더           | 설명            |
| ------------ | ------------- |
| contactEmail | 안내 메일 수신자     |
| host         | CNAME Host 값  |
| value        | CNAME Value 값 |

### 3.3 SKU_MAP 시트 헤더

| skuName          | skuId      |
| ---------------- | ---------- |
| Business Starter | 1010020027 |
| ...              | ...        |

* skuName은 **대소문자 무시** 비교
* 매핑이 없으면 오류 발생

---

## 4. 메뉴/실행 방법

스프레드시트 열면 상단에 `Netkiller` 메뉴가 생성됩니다.

### 4.1 선택된 행 프로비저닝

**Netkiller → 선택된 행을 CONFIG로 실행**

* 현재 선택한 범위의 행만 실행
* 실행 후 결과 컬럼 자동 기록

### 4.2 전체 행 프로비저닝

**Netkiller → 시트 전체를 CONFIG로 실행**

* Provisioning 시트 전체 행 순회 실행
* 실행 후 결과 컬럼 자동 기록

### 4.3 전체 구독 플랜 전환

**Netkiller → 전체 구독 전환 실행**

* 각 행의 planName을 기준으로 플랜 전환 수행
* customerId/subscriptionId 없으면 도메인으로 자동 탐색 후 기록
* TRIAL이었다면 ANNUAL 전환 후 즉시 유료 시작

### 4.4 설정 안내 메일 발송(선택행)

**Netkiller → 선택된 행 설정 안내 메일 발송**

* 선택된 행의 `contactEmail`로 안내 메일 전송
* HTML + 플레인텍스트 동시 제공

---

## 5. 동작 흐름(프로비저닝)

1. `makeConfigFromRow_()`

   * row → cfg 변환
   * skuId 없으면 SKU_MAP으로 skuName 매핑
   * planName, seats, languageCode 정리
   * 4개 관리자 컬럼이 모두 있으면 `manageCustomerUsers=true`

2. `runProvisioningOnce_(cfg)`

   1. Site Verification 토큰 발급 시도(실패해도 계속 진행)
   2. `ensureCustomer_byCfg_()`

      * 고객 없으면 생성, 있으면 재사용
   3. `setCustomerLanguage_(customerId, cfg.languageCode)`
   4. `createSubscriptionIfAbsent_byCfg_()`

      * TRIAL 생성 → 실패하면 FLEXIBLE로 재시도
   5. (옵션) 관리자 사용자 생성/승격

      * Users.insert → Users.makeAdmin
      * 언어 설정 포함
   6. `{customerId, sub}` 반환

3. `writeProvisioningResult_()`

   * 결과 컬럼 보장 후 customer/subscription 정보 기록

---

## 6. 플랜 전환 로직

### 대상

* planName이 **ANNUAL_MONTHLY_PAY** 또는 **ANNUAL_YEARLY_PAY** 일 때만 전환

### 순서

1. 현재 구독 조회 → TRIAL 여부 확인
2. `setPlanForTrialOrFlex_()`

   * changePlan(body, customerId, subscriptionId)
3. 기존이 TRIAL이면 `startPaidService_()` 실행
4. `setRenewalType_()` 로 갱신타입 반영
5. 최신 구독 재조회 후 결과 컬럼 업데이트

---

## 7. 주의사항

1. **FLEXIBLE 생성에 seats 금지**

   * 코드에서 이미 방어하지만, 외부 수정 시 주의

2. **planName은 Reseller 허용값만**

   * 허용값:

     * `TRIAL`
     * `FLEXIBLE`
     * `ANNUAL_MONTHLY_PAY`
     * `ANNUAL_YEARLY_PAY`
   * 별칭 일부 자동 매핑됨(ANNUAL, FLEX 등)

3. **언어코드 포맷**

   * 예: `ko`, `en`, `ja`, `zh-CN`
   * 비어있으면 기본 `ko`

4. **from alias 발송 조건**

   ```js
   from: 'support@netkiller.com'
   ```

   * 해당 alias가 Gmail에서 “보내는 주소”로 등록되어 있어야 동작

5. **TRIAL 생성 불가 케이스**

   * 이미 trial 종료/제한된 도메인 등
   * 이 경우 자동으로 FLEXIBLE로 폴백

---

## 8. 트러블슈팅

### Q1. `SKU_MAP에 'xxx' 매핑이 없습니다.`

* SKU_MAP 시트에 skuName/skuId 행 추가 후 재실행

### Q2. `Not Found`로 고객 조회 실패

* customerDomain 오타/공백 확인
* Reseller 파트너 계정에서 접근 가능한 도메인인지 확인

### Q3. TRIAL 생성 실패 후 FLEXIBLE도 실패

* 동일 SKU가 이미 존재하거나
* 고객이 해당 SKU를 구매할 수 없는 상태일 수 있음
* Reseller 콘솔에서 고객 상태와 SKU 권한 확인

### Q4. startPaidService 실패

* 이미 유료 상태이거나 TRIAL이 아닌데 호출된 케이스
* 로그에서 `wasPlan`, `wasTrial` 확인

### Q5. 메일 발송이 안 됨

* contactEmail 비어있음
* [support@netkiller.com](mailto:support@netkiller.com) alias 미등록
* GmailApp 권한 미승인

---

## 9. 변경 이력(예시)

* `v1` : 기본 고객+TRIAL 프로비저닝
* `v2` :

  * SKU_MAP 자동 매핑
  * 언어 설정(테넌트/사용자)
  * 결과 컬럼 자동 기록
  * TRIAL→ANNUAL 전환 자동화
  * 설정 안내 메일 발송 기능 추가

---

* Provisioning 시트 “샘플 템플릿(헤더+예시 행)”도 만들어 드릴게요.
* README를 사내 위키 양식(Confluence/Notion)으로 맞춰 재정리도 가능!
