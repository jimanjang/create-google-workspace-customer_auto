---

# Google Workspace Reseller — End-to-End Customer Provisioning (Colab)

## 개요

이 Colab 노트북은 **Google Workspace Reseller API**, **Admin SDK (Directory API)**, **Site Verification API**를 이용하여
신규 고객 도메인을 자동으로 프로비저닝(생성 및 설정)하는 과정을 단계별로 실행할 수 있게 구성되었습니다.

참고: [공식 Google Codelab - End-to-End Customer Provisioning](https://developers.google.com/workspace/admin/reseller/v1/codelab/end-to-end)

---

## 주요 기능

| 기능                    | 설명                                     |
| --------------------- | -------------------------------------- |
| Site Verification API | 신규 고객 도메인의 DNS/META 토큰 생성 및 인증         |
| Reseller API          | Customer(고객) 및 Subscription(구독) 생성     |
| Directory API         | 첫 사용자(admin 계정) 생성 및 Super Admin 권한 부여 |
| 자동화 흐름                | 고객 → 사용자 → 구독 → 인증 절차를 순차적으로 수행        |

---

## 사전 준비사항

1. **Google Workspace 리셀러 계정**

   * 파트너 자격을 가진 Workspace 리셀러만 실행 가능

2. **Google Cloud 프로젝트 설정**

   * 아래 API들을 사용 설정(Enable)해야 함:

     * Reseller API → `reseller.googleapis.com`
     * Site Verification API → `siteverification.googleapis.com`
     * Admin SDK → `admin.googleapis.com`

3. **서비스 계정 생성 및 도메인 전체 위임 (Domain-wide Delegation)**

   * Google Cloud Console → IAM & Admin → Service Accounts
   * "Enable G Suite Domain-wide Delegation" 체크
   * 키(JSON) 생성 후 Colab에 업로드

4. **관리자 콘솔에 서비스 계정 등록**

   * Admin Console → 보안 → API 컨트롤 → 도메인 전체 위임 관리
   * 클라이언트 ID 등록 후, 다음 OAuth 범위를 입력:

     ```
     https://www.googleapis.com/auth/apps.order,
     https://www.googleapis.com/auth/admin.directory.user,
     https://www.googleapis.com/auth/admin.directory.user.security,
     https://www.googleapis.com/auth/siteverification
     ```

---

## 노트북 실행 단계 요약

| 단계                                 | 설명                                                                                           |
| ---------------------------------- | -------------------------------------------------------------------------------------------- |
| 1. 라이브러리 설치                        | `google-api-python-client`, `google-auth`, `google-auth-httplib2`, `google-auth-oauthlib` 설치 |
| 2. 서비스 계정 JSON 업로드                 | 업로드 및 형식 검증 수행                                                                               |
| 3. 구성 변수 설정                        | 도메인, 이메일, 주소, SKU, 플랜 정보 입력                                                                  |
| 4. 인증 및 클라이언트 생성                   | Credentials 객체 생성, `reseller`, `sitever`, `directory` 클라이언트 빌드                               |
| 5. Site Verification 토큰 생성         | DNS TXT 또는 META 태그 방식으로 도메인 인증 토큰 발급                                                         |
| 6. Customer 생성 (Reseller API)      | 고객 도메인 등록 및 기본 연락처 정보 설정                                                                     |
| 7. 첫 사용자(Admin) 생성 (Directory API) | `admin@도메인` 생성 및 Super Admin 권한 부여                                                           |
| 8. Subscription 생성 (Reseller API)  | SKU, 플랜, 좌석수 지정하여 구독 등록                                                                      |
| 9. 도메인 최종 인증 (Site Verification)   | DNS/META 토큰 적용 후 인증 완료                                                                       |

---

## 예시 값

```python
DELEGATED_ADMIN_EMAIL = "laika.jang@netkillersoft.com"
CUSTOMER_DOMAIN = "laikacreate.netkiller-gws.com"
ALTERNATE_EMAIL = "owner@laikacreate.betanetkillersoft.com"
CUSTOMER_PHONE = "+82-2-0000-0000"
POSTAL_ADDRESS = {
    "contactName": "Laika Jang",
    "organizationName": "New Customer Inc.",
    "postalCode": "04524",
    "region": "Seoul",
    "countryCode": "KR",
    "locality": "Jung-gu",
    "addressLine1": "123 Example-ro"
}
PRIMARY_ADMIN_EMAIL = "admin@" + CUSTOMER_DOMAIN
PRIMARY_ADMIN_PASSWORD = "TempPass!234"
SKU_ID = "1010020027"  # Google Workspace Business Starter
PLAN = "TRIAL"
SEAT_NUMBER = 1
```

---

## 주의 사항

* Site Verification 후 도메인 전파에 몇 분~몇 시간이 걸릴 수 있습니다.
* `customerDomainVerified` 값이 `false`라도 실제 인증 완료 후 자동 갱신됩니다.
* Trial 구독 생성 시 `"status": "SUSPENDED"` 및 `"PENDING_TOS_ACCEPTANCE"`는 정상입니다.
  고객이 Admin Console 첫 로그인 시 약관을 수락하면 `"ACTIVE"` 상태로 전환됩니다.

---

## 오류 해결 가이드

| 오류 메시지                                                   | 원인                             | 해결                                                         |
| -------------------------------------------------------- | ------------------------------ | ---------------------------------------------------------- |
| `MalformedError: missing fields token_uri, client_email` | 업로드한 파일이 서비스 계정 JSON이 아님       | Cloud Console에서 JSON 키 재발급                                 |
| `403 accessNotConfigured`                                | API 미사용 또는 비활성화                | 해당 프로젝트에서 API 사용 설정                                        |
| `400 Required field must not be blank`                   | `postalAddress.contactName` 누락 | POSTAL_ADDRESS에 `contactName` 추가                           |
| `409 Resource already exists`                            | 이미 존재하는 고객                     | 기존 `customerId` 사용                                         |
| `400 Invalid Value`                                      | SKU ID 또는 seats 형식 오류          | SKU는 숫자 ID(`1010020027`), seats는 `maximumNumberOfSeats` 사용 |
| `SUSPENDED: PENDING_TOS_ACCEPTANCE`                      | 약관 미수락                         | 고객이 admin 계정으로 로그인 후 약관 수락                                 |

---

## 결과 확인

* 고객(Admin Console URL):
  `https://admin.google.com/ac/home?ecid=<customerId>`

* 구독 관리 URL:
  `https://admin.google.com/ac/billing/subscriptions?ecid=<customerId>`

* Colab에서 상태 확인:

  ```python
  sub = reseller.subscriptions().get(customerId=CUSTOMER_ID, subscriptionId=SUBSCRIPTION_ID).execute()
  print(sub["status"])
  ```

---

## 확장 아이디어

* 여러 고객 일괄 등록(batch provisioning)
* 구독 자동 전환 (Trial → FLEXIBLE)
* DNS TXT 자동 등록 API 연동 (Cloud DNS, Route53 등)
* 에러 로깅 및 Slack 알림 통합

---

## 라이선스

이 노트북은 Google Codelab 기반 예시를 확장한 것으로,
Google API Client Library for Python의 [Apache 2.0 License](https://www.apache.org/licenses/LICENSE-2.0)를 따릅니다.

---

이 README를 `README.md` 파일로 저장해 드릴 수도 있습니다.
원하시면 `.md` 파일로 만들어 드릴까요?
