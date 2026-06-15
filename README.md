# L-Camp Vietnam 2026 — 자동 평가서 생성기 v4

엑셀 파일을 올리면 Claude API가 IR Deck까지 읽어서 자동 심사하고
기업별 평가서 HTML + 대시보드를 생성합니다.

**v4에서는 ICAS / IBS(MYSC·ISQ) / CTS 적격 기준을 자동으로 체크하여,
탈락 기업은 평가서를 생성하지 않고 대시보드에 사유와 함께 표시합니다.**

**중복 지원 기업(L-Camp Vietnam + SW Top 100 동시 지원)은 대시보드에 한 행으로 표시하되,
파이프라인별 평가서는 각각 별도로 생성합니다.**

---

## 📁 폴더 구성

```
lcamp_2026/
├── generate.py                          ← 메인 실행 파일 (v4)
├── prompt.txt                           ← Claude 심사 지침 (회의적 평가 프레임워크)
├── template.html                        ← 평가서 HTML 템플릿
├── credentials.json                     ← Google 서비스 계정 키 (절대 외부 공유 금지)
├── ICAS_연도별_수상_기업_리스트.xlsx      ← 적격 심사용 ★ v4 신규
├── 타_IBS_사업_수상_기업_리스트.xlsx     ← 적격 심사용 ★ v4 신규
├── CTS_사업_기업_리스트.xlsx             ← 적격 심사용 ★ v4 신규
└── L-Camp__Responses_.xlsx              ← 지원서 엑셀 파일
```

---

## 🛠 최초 1회 설치 (PowerShell)

```powershell
python --version   # 3.8 이상 확인

pip install anthropic pandas openpyxl google-auth google-auth-httplib2 google-api-python-client pdfplumber
```

---

## 🚀 실행 방법

```powershell
cd C:\Users\본인이름\Documents\lcamp_2026

# API 키 설정
$env:ANTHROPIC_API_KEY = "sk-ant-여기에본인키입력"

# 전체 실행
python generate.py

# 파일명 직접 지정
python generate.py "L-Camp__Responses__0529.xlsx"

# N번째 기업부터 이어서 실행 (오류 재개용)
python generate.py --start 5

# 대시보드만 재생성 (API 호출 없이 기존 평가서 HTML 자동 수집)
python generate.py --rebuild-dashboard
python generate.py --rebuild-dashboard output   ← output 폴더 직접 지정
```

---

## 🔄 중복 지원 기업 처리 (v4 신규)

같은 기업이 **L-Camp Vietnam**과 **SW Top 100**에 모두 지원한 경우:

| 항목 | 처리 방식 |
|------|-----------|
| 대시보드 표시 | **한 행**으로 표시 (중복 없음) |
| 파이프라인 배지 | 해당하는 배지 **모두** 표시 (예: `L-Camp` `SW100`) |
| 평가서 생성 | 파이프라인별로 **각각 별도** 생성 |
| 평가서 파일명 | `기업명_LCamp_평가서.html` / `기업명_SW100_평가서.html` |
| 모달 평가서 버튼 | `L-Camp 평가서 →` / `SW100 평가서 →` 버튼 **각각** 표시 |
| 테이블 평가서 열 | `L-Camp 평가서 →` `SW100 평가서 →` **각각** 표시 |

> 엑셀에서 동일한 기업명이 두 행으로 존재하면 자동으로 중복 감지합니다.

---

## ⛔ 적격 기준 (v4 신규)

실행 시 아래 3가지 기준을 자동으로 체크합니다.
**하나라도 해당하면 평가서를 생성하지 않고** 대시보드에 지원서 정보 + 사유만 표시합니다.

| # | 기준 | 판단 방법 | 사유 표시 예시 |
|---|------|-----------|---------------|
| 1 | **ICAS 수상 여부** | ICAS_연도별_수상_기업_리스트.xlsx의 모든 시트(22·23·24년도 VN, 22 KH 포함) | `ICAS 2024 VN TOP20` |
| 2 | **타 IBS 사업(MYSC/ISQ) 수상 여부** | 타_IBS_사업_수상_기업_리스트.xlsx의 모든 시트(MYSC 1-4차년도, ISQ 1·2·3차년도) | `MYSC 1-4차년도`, `ISQ 3차년도` |
| 3 | **CTS 선정 및 수상 여부** | CTS_사업_기업_리스트.xlsx (한국 본사 기업에 한해 적용) | `CTS` |

> **CTS 적용 주의사항**  
> CTS는 한국 기업 대상 사업이므로, 지원서의 `HQ Location` 필드에 Korea/Seoul 등이 포함된 경우에만 CTS 리스트와 대조합니다.  
> BSSC 베트남 풀이더라도 한국 HQ가 있는 기업은 검토됩니다.

### 사유 표시 방식
- 1개 탈락 사유: `ICAS 2024 VN TOP20`
- 2개 이상: `ICAS 2024 VN TOP20 / MYSC 1-4차년도`
- 대시보드 테이블의 **평가 등급** 열에 빨간 태그로 표시됩니다

---

## 🌐 대시보드 열기

`output/index.html` 파일을 더블클릭하거나 브라우저로 직접 열면 됩니다.

**v4 대시보드 신규 기능:**
- 사이드바에 **"⛔ 적격 탈락"** 필터 버튼 추가
- KPI 카드에 적격 탈락 기업 수 표시
- 적격 탈락 기업 행은 회색으로 표시, 평가서 버튼 없음
- 기업 모달(지원서 탭) 상단에 적격 탈락 사유 배너 표시
- 중복 지원 기업: 파이프라인 배지 복수 표시, 평가서 버튼 파이프라인별 구분

**GitHub Pages 배포 시** `output` 폴더를 그대로 레포지토리에 올리면 됩니다.

> ⚠️ `file://`로 직접 열 경우 심사자 코멘트·등급 저장은 브라우저 보안 정책상 작동하지 않을 수 있습니다.
> ```powershell
> python -m http.server 8000 --directory output
> ```

---

## 📋 엑셀 컬럼 매핑 (v4 기준)

| 내부 키 | 엑셀 컬럼명 |
|--------|------------|
| deal_pipeline | Deal Pipeline |
| company_name | Company name (in English) |
| company_name_vn | Company name (in Vietnamese) |
| full_name | Full name |
| position | Position |
| email | Email |
| phone | Phone number |
| website | Website |
| **erc_city** | Which city in Vietnam is your ERC registered in? |
| location | Which country and city is your HQ located in? |
| founded_year | Founded year |
| product_desc | Description of your product (1-2 sentences) |
| team_background | Core competencies of the founding team |
| pitchdeck | Please upload your Company Profile (e.g. IR Deck) |
| industry | Industry/Sector |
| traction | Business traction: Number of users and key clients/partners |
| traction_awards | Business traction: Awards and other achievements |
| has_revenue | Do you currently generate revenue? |
| revenue_detail | If yes, please specify your revenue in 2024, 2025 and 2026 |
| **has_investment** | Do you have prior investment history? |
| **valuation** | If yes, please specify your current Pre-money Valuation |
| is_raising | Are you currently raising fund? |
| target_raise | If yes, please specify your target fundraising amount |
| motivation | Motivation for applying to the program |
| lotte_area | Which LOTTE business area is most relevant to your startup |
| lotte_affiliate | Which LOTTE affiliate would you like to collaborate with? |
| poc_idea | Proposed PoC idea with LOTTE |
| esg_impact | What ESG impact does your solution address |
| **heard_from** | Where did you get to know about 2026 L-Camp Vietnam? |
| recommender | (If applicable) Program Recommending Organization |

컬럼명이 다르면 `generate.py` 상단 `COLUMN_MAP`에서 수정하세요.

---

## 🆕 v4.1 버그 수정 (평가서 중복 표시 해결)

### 수정 내용
`--start/--end` 옵션이나 `--rebuild-dashboard` 명령으로 재실행 시, 이미 평가된 기업의 평가서가 대시보드에 2개 이상 중복 표시되던 버그를 수정했습니다.

**원인:** `recover_cache_from_html_files()`가 구버전 파일명(suffix 없음: `기업명_평가서.html`)을 읽을 때 `deal_pipeline`을 빈 문자열(`""`)로 복구하여, 캐시 키가 `기업명||""`와 `기업명||L-Camp Vietnam` 두 개로 분리 저장되었고, `generate_dashboard()`의 `filenames` 배열에 같은 파일이 반복 추가되었습니다.

**수정 3곳:**
1. `_infer_pipeline_from_filename()` 함수 신규 추가 — 파일명 suffix(`_LCamp_`, `_SW100_`)로 파이프라인 자동 추론
2. `recover_cache_from_html_files()` — 복구 시 동일 파일명 중복 등록 방지
3. `load_existing_summary()` — HTML 스캔 복구 시 동일 파일명이 이미 있으면 스킵
4. `generate_dashboard()` — `filenames` 배열에 동일 파일명 중복 추가 방지 (최후 방어선)

---

## 🆕 v4 주요 업데이트

### 중복 지원 기업 처리
- 동일 기업이 두 파이프라인에 지원 시 대시보드 한 행 표시
- 파이프라인별 평가서 파일 분리 생성 (`_LCamp_`, `_SW100_` suffix)
- 모달 및 테이블에서 파이프라인별 평가서 버튼 각각 표시

### 적격 심사 자동화
- 실행 시작 시 3개 엑셀 DB 자동 로드 (ICAS / IBS / CTS)
- 기업명 정규화 매칭: 법인 형태(Co.,Ltd / JSC / (주) 등) 제거 후 비교
- 한국 HQ 여부 자동 판단 (CTS 적용 대상 필터링)
- 적격 탈락 기업은 평가 루프에서 자동 스킵

### 대시보드 개선
- 적격 탈락 기업 시각화 (회색 처리, 탈락 사유 태그)
- 사이드바 적격 탈락 필터
- KPI 6개로 확장 (적격 탈락 수 추가)
- 탈락 사유 배너 (모달 지원서 탭 상단)

### 실행 로그 개선
- 시작 시 적격 탈락 기업 실시간 출력
- 중복 지원 기업 `[중복: 파이프라인명]` 표시
- 종료 시 탈락 기업 목록 + 사유 요약

---

## 🔑 Google Drive 연동 구조

```
[엑셀 피치덱 URL] → 파일 ID 추출
       ↓
[Drive API + credentials.json 인증]
       ↓
[PDF / Google Slides 다운로드 → 텍스트 추출]
       ↓
[Claude API: 지원서 + IR Deck 통합 평가]
```

서비스 계정 이메일을 드라이브 폴더 편집자로 추가 필요.

---

## 💰 예상 API 비용 (Claude Sonnet 기준)

| 기업 수 | IR Deck 없음 | IR Deck + 팀리서치 포함 |
|---------|-------------|------------------------|
| 20개    | $0.5~1      | $1.5~3                 |
| 100개   | $2~4        | $7~14                  |

※ 적격 탈락 기업은 API 호출 없으므로 비용 절감 효과 있음  
※ 중복 지원 기업은 파이프라인 수만큼 API 호출 발생

---

## ❓ 자주 묻는 문제

**Q. 중복 기업인데 대시보드에 두 행으로 나와요**  
A. 엑셀의 두 행에서 `Company name (in English)` 값이 정확히 동일해야 합니다. 띄어쓰기·대소문자 차이가 있으면 별개 기업으로 인식합니다.

**Q. 적격 탈락 기업 판단이 잘못된 것 같아요**  
A. 기업명 정규화 매칭을 사용하므로 유사한 이름은 동일 기업으로 판단할 수 있습니다.  
수동 예외 처리가 필요하다면 `generate.py`의 `check_eligibility()` 함수를 직접 수정하거나,  
해당 기업의 `final_grade`를 캐시 파일(`.summary_cache.json`)에서 수동으로 변경 후 `--rebuild-dashboard`를 실행하세요.

**Q. 적격 DB 엑셀 파일이 없어요**  
A. 해당 파일이 없으면 그 항목의 적격 심사를 건너뜁니다 (경고 메시지 출력).  
전부 없어도 나머지 로직은 정상 동작합니다.

**Q. 링크드인을 자주 못 찾아요**  
A. LinkedIn은 비로그인 크롤링을 차단합니다. `[부분 확인]` 처리가 정상입니다.

**Q. credentials.json 오류**  
A. generate.py와 같은 폴더에 있는지, 파일명이 정확한지 확인.

**Q. JSON 파싱 오류**  
A. 해당 기업 건너뛰고 계속 진행. `--start N`으로 재개 가능.

**Q. 심사자 코멘트가 다른 PC에서 안 보여요**  
A. 브라우저 로컬스토리지 기반이라 같은 브라우저/PC에서만 유지됩니다.

엑셀 파일을 올리면 Claude API가 IR Deck까지 읽어서 자동 심사하고
기업별 평가서 HTML + 대시보드를 생성합니다.

**v4에서는 ICAS / IBS(MYSC·ISQ) / CTS 적격 기준을 자동으로 체크하여,
탈락 기업은 평가서를 생성하지 않고 대시보드에 사유와 함께 표시합니다.**

---

## 📁 폴더 구성

```
lcamp_2026/
├── generate.py                          ← 메인 실행 파일 (v4)
├── prompt.txt                           ← Claude 심사 지침 (회의적 평가 프레임워크)
├── template.html                        ← 평가서 HTML 템플릿
├── credentials.json                     ← Google 서비스 계정 키 (절대 외부 공유 금지)
├── ICAS_연도별_수상_기업_리스트.xlsx      ← 적격 심사용 ★ v4 신규
├── 타_IBS_사업_수상_기업_리스트.xlsx     ← 적격 심사용 ★ v4 신규
├── CTS_사업_기업_리스트.xlsx             ← 적격 심사용 ★ v4 신규
└── L-Camp__Responses_.xlsx              ← 지원서 엑셀 파일
```

---

## 🛠 최초 1회 설치 (PowerShell)

```powershell
python --version   # 3.8 이상 확인

pip install anthropic pandas openpyxl google-auth google-auth-httplib2 google-api-python-client pdfplumber
```

---

## 🚀 실행 방법

```powershell
cd C:\Users\본인이름\Documents\lcamp_2026

# API 키 설정
$env:ANTHROPIC_API_KEY = "sk-ant-여기에본인키입력"

# 전체 실행
python generate.py

# 파일명 직접 지정
python generate.py "L-Camp__Responses__0529.xlsx"

# N번째 기업부터 이어서 실행 (오류 재개용)
python generate.py --start 5

# 대시보드만 재생성 (API 호출 없이 기존 평가서 HTML 자동 수집)
python generate.py --rebuild-dashboard
python generate.py --rebuild-dashboard output   ← output 폴더 직접 지정
```

---

## ⛔ 적격 기준 (v4 신규)

실행 시 아래 3가지 기준을 자동으로 체크합니다.
**하나라도 해당하면 평가서를 생성하지 않고** 대시보드에 지원서 정보 + 사유만 표시합니다.

| # | 기준 | 판단 방법 | 사유 표시 예시 |
|---|------|-----------|---------------|
| 1 | **ICAS 수상 여부** | ICAS_연도별_수상_기업_리스트.xlsx의 모든 시트(22·23·24년도 VN, 22 KH 포함) | `ICAS 2024 VN TOP20` |
| 2 | **타 IBS 사업(MYSC/ISQ) 수상 여부** | 타_IBS_사업_수상_기업_리스트.xlsx의 모든 시트(MYSC 1-4차년도, ISQ 1·2·3차년도) | `MYSC 1-4차년도`, `ISQ 3차년도` |
| 3 | **CTS 선정 및 수상 여부** | CTS_사업_기업_리스트.xlsx (한국 본사 기업에 한해 적용) | `CTS` |

> **CTS 적용 주의사항**  
> CTS는 한국 기업 대상 사업이므로, 지원서의 `HQ Location` 필드에 Korea/Seoul 등이 포함된 경우에만 CTS 리스트와 대조합니다.  
> BSSC 베트남 풀이더라도 한국 HQ가 있는 기업은 검토됩니다.

### 사유 표시 방식
- 1개 탈락 사유: `ICAS 2024 VN TOP20`
- 2개 이상: `ICAS 2024 VN TOP20 / MYSC 1-4차년도`
- 대시보드 테이블의 **평가 등급** 열에 빨간 태그로 표시됩니다

---

## 🌐 대시보드 열기

`output/index.html` 파일을 더블클릭하거나 브라우저로 직접 열면 됩니다.

**v4 대시보드 신규 기능:**
- 사이드바에 **"⛔ 적격 탈락"** 필터 버튼 추가
- KPI 카드에 적격 탈락 기업 수 표시
- 적격 탈락 기업 행은 회색으로 표시, 평가서 버튼 없음
- 기업 모달(지원서 탭) 상단에 적격 탈락 사유 배너 표시

**GitHub Pages 배포 시** `output` 폴더를 그대로 레포지토리에 올리면 됩니다.

> ⚠️ `file://`로 직접 열 경우 심사자 코멘트·등급 저장은 브라우저 보안 정책상 작동하지 않을 수 있습니다.
> ```powershell
> python -m http.server 8000 --directory output
> ```

---

## 📋 엑셀 컬럼 매핑 (v4 기준)

| 내부 키 | 엑셀 컬럼명 |
|--------|------------|
| deal_pipeline | Deal Pipeline |
| company_name | Company name (in English) |
| company_name_vn | Company name (in Vietnamese) |
| full_name | Full name |
| position | Position |
| email | Email |
| phone | Phone number |
| website | Website |
| **erc_city** | Which city in Vietnam is your ERC registered in? |
| location | Which country and city is your HQ located in? |
| founded_year | Founded year |
| product_desc | Description of your product (1-2 sentences) |
| team_background | Core competencies of the founding team |
| pitchdeck | Please upload your Company Profile (e.g. IR Deck) |
| industry | Industry/Sector |
| traction | Business traction: Number of users and key clients/partners |
| traction_awards | Business traction: Awards and other achievements |
| has_revenue | Do you currently generate revenue? |
| revenue_detail | If yes, please specify your revenue in 2024, 2025 and 2026 |
| **has_investment** | Do you have prior investment history? |
| **valuation** | If yes, please specify your current Pre-money Valuation |
| is_raising | Are you currently raising fund? |
| target_raise | If yes, please specify your target fundraising amount |
| motivation | Motivation for applying to the program |
| lotte_area | Which LOTTE business area is most relevant to your startup |
| lotte_affiliate | Which LOTTE affiliate would you like to collaborate with? |
| poc_idea | Proposed PoC idea with LOTTE |
| esg_impact | What ESG impact does your solution address |
| **heard_from** | Where did you get to know about 2026 L-Camp Vietnam? |
| recommender | (If applicable) Program Recommending Organization |

컬럼명이 다르면 `generate.py` 상단 `COLUMN_MAP`에서 수정하세요.

---

## 🆕 v4 주요 업데이트

### 적격 심사 자동화
- 실행 시작 시 3개 엑셀 DB 자동 로드 (ICAS / IBS / CTS)
- 기업명 정규화 매칭: 법인 형태(Co.,Ltd / JSC / (주) 등) 제거 후 비교
- 한국 HQ 여부 자동 판단 (CTS 적용 대상 필터링)
- 적격 탈락 기업은 평가 루프에서 자동 스킵

### 대시보드 개선
- 적격 탈락 기업 시각화 (회색 처리, 탈락 사유 태그)
- 사이드바 적격 탈락 필터
- KPI 6개로 확장 (적격 탈락 수 추가)
- 탈락 사유 배너 (모달 지원서 탭 상단)

### 실행 로그 개선
- 시작 시 적격 탈락 기업 실시간 출력
- 종료 시 탈락 기업 목록 + 사유 요약

---

## 🔑 Google Drive 연동 구조

```
[엑셀 피치덱 URL] → 파일 ID 추출
       ↓
[Drive API + credentials.json 인증]
       ↓
[PDF / Google Slides 다운로드 → 텍스트 추출]
       ↓
[Claude API: 지원서 + IR Deck 통합 평가]
```

서비스 계정 이메일을 드라이브 폴더 편집자로 추가 필요.

---

## 💰 예상 API 비용 (Claude Sonnet 기준)

| 기업 수 | IR Deck 없음 | IR Deck + 팀리서치 포함 |
|---------|-------------|------------------------|
| 20개    | $0.5~1      | $1.5~3                 |
| 100개   | $2~4        | $7~14                  |

※ 적격 탈락 기업은 API 호출 없으므로 비용 절감 효과 있음

---

## ❓ 자주 묻는 문제

**Q. 적격 탈락 기업 판단이 잘못된 것 같아요**  
A. 기업명 정규화 매칭을 사용하므로 유사한 이름은 동일 기업으로 판단할 수 있습니다.  
수동 예외 처리가 필요하다면 `generate.py`의 `check_eligibility()` 함수를 직접 수정하거나,  
해당 기업의 `final_grade`를 캐시 파일(`.summary_cache.json`)에서 수동으로 변경 후 `--rebuild-dashboard`를 실행하세요.

**Q. 적격 DB 엑셀 파일이 없어요**  
A. 해당 파일이 없으면 그 항목의 적격 심사를 건너뜁니다 (경고 메시지 출력).  
전부 없어도 나머지 로직은 정상 동작합니다.

**Q. 링크드인을 자주 못 찾아요**  
A. LinkedIn은 비로그인 크롤링을 차단합니다. `[부분 확인]` 처리가 정상입니다.

**Q. 지원서 탭에서 일부 필드가 안 보여요**  
A. 지원서에 해당 항목을 비워 제출한 경우 표시되지 않습니다.

**Q. credentials.json 오류**  
A. generate.py와 같은 폴더에 있는지, 파일명이 정확한지 확인.

**Q. JSON 파싱 오류**  
A. 해당 기업 건너뛰고 계속 진행. `--start N`으로 재개 가능.

**Q. 심사자 코멘트가 다른 PC에서 안 보여요**  
A. 브라우저 로컬스토리지 기반이라 같은 브라우저/PC에서만 유지됩니다.
