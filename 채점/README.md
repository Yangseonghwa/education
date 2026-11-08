# AI 자동 채점 (AI Auto-Grading for Google Sheets & Classroom)

Google 스프레드시트와 Google Apps Script 기반의 **Classroom Google Docs 제출물 AI 자동 채점 도구**입니다. Google Drive 내 `Classroom`(또는 `클래스룸`) 폴더의 학생 제출물(Google Docs)과 교사용 정답지(Google Docs)를 비교하여 **OpenAI 호환 API**, **OpenAI Responses API**, **Anthropic Claude API**, 또는 **Google(Gemini) API**로 자동 채점하고, 과목·과제별 시트에 채점 결과와 피드백을 기록합니다.

---

## 주요 기능

- **Google Drive / Classroom 폴더 탐색기**:
  - 사이드바 내에서 `Classroom` 폴더(수업 폴더 › 과제 폴더)와 `내 드라이브`(또는 공유 드라이브)의 정답지 Google Docs를 직접 탐색·선택할 수 있습니다.
  - 과제 폴더 선택 시 폴더 내 Google Docs 제출물 수를 즉시 표시하며, 빈 폴더 선택 실수를 방지합니다.
- **4종 AI API 제공자/규격 지원**:
  - `OpenAI 호환 · Chat Completions` (`/chat/completions`, 기본 Base URL: `https://api.openai.com/v1`)
  - `OpenAI · Responses` (`/responses`, 기본 Base URL: `https://api.openai.com/v1`)
  - `Claude · Messages` (`/messages`, 기본 Base URL: `https://api.anthropic.com/v1`, `x-api-key` 및 `anthropic-version: 2023-06-01` 헤더)
  - `Google(Gemini)` (`/models/{model}:generateContent` 및 `/v1beta/openai` 호환 엔드포인트 자동 판별, 기본 Base URL: `https://generativelanguage.googleapis.com/v1beta`, `x-goog-api-key` 헤더 및 `responseMimeType: "application/json"` 지원)
- **API 키 기반 실시간 동적 모델 목록 조회**:
  - 정적으로 하드코딩된 모델 목록 대신, 입력창에 입력했거나 저장된 API 키를 사용하여 각 제공자의 `/models` 엔드포인트에서 **사용 가능한 모델 목록을 실시간으로 조회(fetch)**하여 드롭다운에 표시합니다.
  - `Google(Gemini)` 선택 시 `supportedGenerationMethods`에 `generateContent`가 포함된 생성 모델만 자동으로 필터링하여 표시합니다.
  - API 규격 변경 시 해당 제공자의 기본 Base URL이 자동 설정되며, API 키가 있으면 즉시 해당 제공자의 모델 목록을 새로고침합니다.
- **정량 / 정량+정성 채점 모드**:
  - **정량평가만 반영**: 객관식·단답형·빈칸 등 정량 판정 문항만 점수에 반영하고 서술형은 참고 피드백만 기록
  - **정량 + 정성평가 모두 반영**: 정량 점수와 서술형(정성) 점수를 합산하여 최종 점수 산출
- **AI 사용 의심 신호(휴리스틱) 분석 옵션**:
  - 선택 시 서술형 답안의 AI 작성 의심 수준(`high` / `medium` / `low`), 의심 점수(`0~100`), 판단 이유 및 근거 문장을 함께 기록합니다(참고용 보조 지표).
- **안정적인 백그라운드 예약 큐 및 자동 복구**:
  - `PropertiesService`(UserProperties) 일괄 청크 저장과 `LockService` 동시성 제어를 통해 브라우저나 사이드바를 닫아도 백그라운드 트리거(`processQueue`)가 채점을 이어갑니다.
  - Apps Script 실행 시간 제한(6분)을 고려해 최대 5분 30초(`330초`) 단위로 안전하게 분할 처리합니다.
  - 전체 제출물 일괄 채점 시 응답 JSON 잘림 오류가 발생하면 자동으로 **13명 이하 소규모 묶음**으로 분할해 재시도합니다.
  - 일시적 API 오류(`408`, `409`, `425`, `429`, `5xx`)는 **10분 간격으로 최대 5회** 자동 재시도하며, 크레딧/쿼터 소진(`402`, `insufficient_quota` 등)은 즉시 감지해 불필요한 재시도를 중단합니다.
- **스프레드시트 수식 인젝션(Formula Injection) 방지 및 중복 기록 방지**:
  - 학생 답안·피드백 텍스트가 `=`, `+`, `-`, `@` 등으로 시작할 경우 자동으로 이스케이프(`'`) 처리합니다.
  - 동일 `작업ID + 학생 답안 URL` 조합은 중복 행 추가 대신 기존 행을 덮어씁니다.

---

## 설치 방법

1. Google 스프레드시트를 새로 만들거나 기존 채점용 스프레드시트를 엽니다.
2. 상단 메뉴에서 **확장 프로그램 > Apps Script**를 클릭합니다.
3. 프로젝트에 다음 두 파일을 구성합니다:
   - `Code.gs`: 이 폴더의 `Code.gs` 전체 내용을 복사하여 붙여넣습니다.
   - `sidebar.html`: 파일 추가(`+` > **HTML**)로 이름을 `sidebar`(`sidebar.html`)로 지정한 뒤, `sidebar.html` 내용을 붙여넣습니다.
4. *(권장)* 좌측 **서비스(`+`)** 메뉴에서 **Google Drive API (`Drive`, v3)** 서비스를 추가합니다.
   - 공유 드라이브 탐색 및 빠른 폴더/문서 목록 조회에 사용됩니다. 추가하지 않아도 기본 `DriveApp` 폴백으로 내 드라이브 탐색이 동작합니다.
5. 프로젝트를 저장(`Ctrl+S`)하고 스프레드시트를 새로고침하면 상단에 **AI 자동 채점** 메뉴가 나타납니다.

> **참고**: 결과 시트(`<클래스명>_<과제명>`)는 채점 완료 시 자동으로 생성·초기화되며, 수동으로 시트를 미리 만들 필요가 없습니다. (레거시 `AI채점_작업`, `AI채점_제출물` 시트가 있을 경우 자동으로 숨김 처리됩니다.)

---

## 사용 방법

1. 스프레드시트 상단 메뉴에서 **AI 자동 채점 > 채점 사이드바 열기**를 클릭합니다.
2. **API 설정 및 동적 모델 조회**:
   - 사용할 **API 규격**(`OpenAI 호환 · Chat Completions`, `OpenAI · Responses`, `Claude · Messages`, `Google(Gemini)`)을 선택합니다. 선택 시 각 제공자의 기본 **Base URL**이 자동으로 채워집니다.
   - **API 키**를 입력하고 **불러오기**(또는 **설정 저장**) 버튼을 누르면, 입력된 API 키로 해당 제공자의 사용 가능한 모델 목록을 동적으로 조회하여 **모델 선택** 드롭다운에 표시합니다.
   - 드롭다운에서 원하는 모델을 선택하거나 **모델 ID 직접 입력** 칸에 모델명(예: `gemini-2.5-flash`, `gpt-4o-mini`, `claude-3-5-sonnet-latest`)을 직접 입력할 수도 있습니다.
3. **채점 대상 선택**:
   - **과제 폴더 선택** 영역에서 `Classroom › <수업 폴더> › <과제 폴더>`로 이동한 뒤 **과제 폴더 확정** 버튼을 클릭합니다.
   - **정답지 폴더 탐색** 영역에서 정답지 문서가 있는 폴더로 이동한 후 **정답지 Google Docs 선택** 드롭다운에서 정답지 문서를 선택합니다.
4. **채점 방식 설정 및 실행**:
   - 점수 반영 범위(`정량평가만` / `정량 + 정성평가`), 만점, 최대 출력 토큰(`1,024 ~ 65,536`), AI 사용 의심 분석 여부를 선택합니다.
   - **예약 큐에 추가**를 누른 뒤, 하단 **예약 큐 확인** 영역에서 **예약 큐 전체 실행**을 클릭하면 채점이 시작됩니다.

---

## 제공자별 엔드포인트 및 인증 규격

| API 규격 (`apiFormat`) | 기본 Base URL | 모델 목록 조회 (`GET`) | 채점 요청 (`POST`) | 인증 헤더 |
|---|---|---|---|---|
| `chat_completions` | `https://api.openai.com/v1` | `{baseUrl}/models` | `{baseUrl}/chat/completions` | `Authorization: Bearer <API_KEY>` |
| `responses` | `https://api.openai.com/v1` | `{baseUrl}/models` | `{baseUrl}/responses` | `Authorization: Bearer <API_KEY>` |
| `claude` | `https://api.anthropic.com/v1` | `{baseUrl}/models` | `{baseUrl}/messages` | `x-api-key: <API_KEY>`, `anthropic-version: 2023-06-01` |
| `gemini` | `https://generativelanguage.googleapis.com/v1beta` | `{baseUrl}/models?pageSize=100` | `{baseUrl}/models/{model}:generateContent` *(또는 `/v1beta/openai` 사용 시 `/chat/completions`)* | `x-goog-api-key: <API_KEY>` *(OpenAI 호환 URL 사용 시 `Bearer`)* |

---

## 결과 시트 열 구성 (`<클래스명>_<과제명>`)

채점이 완료되면 과목·과제 이름으로 생성된 시트에 다음 17개 열이 자동 기록됩니다:

| 열 순서 | 열 이름 | 설명 |
|:---:|---|---|
| 1 | `학생명` | 제출물 Google Docs 파일명(학생 이름/문서 제목) |
| 2 | `점수` | 최종 반영 점수 (만점 기준 클램핑 및 소수 둘째 자리 반올림) |
| 3 | `만점` | 설정한 과제 만점 |
| 4 | `AI 의심` | `high` / `medium` / `low` / `not_checked` |
| 5 | `AI 의심 점수` | `0 ~ 100` (미검사 시 `0`) |
| 6 | `정성점수` | 서술형·정성평가 점수 (`정량만` 모드일 때는 `0`) |
| 7 | `피드백` | 학생별 채점 피드백 요약 |
| 8 | `신뢰도` | AI 채점 신뢰도 (`high` / `medium` / `low`) |
| 9 | `채점일시` | 채점 기록 시각 |
| 10 | `클래스` | Classroom 수업 폴더명 |
| 11 | `과제` | 과제 폴더명 |
| 12 | `학생ID` | 학생 식별자 (폴더 기반 수집 시 공란) |
| 13 | `학생 답안` | 학생 제출물 Google Docs 편집 링크 |
| 14 | `정답지` | 기준 정답지 Google Docs 편집 링크 |
| 15 | `AI 의심 이유` | AI 사용 의심 판단 근거 요약 |
| 16 | `AI 의심 근거` | 답안에서 발췌한 근거 문장 (`\|` 구분) |
| 17 | `작업ID` | 예약 큐 작업 고유 UUID |

---

## 주요 설정값 요약

| 설정 항목 | 설명 | 기본값 / 범위 |
|----------|------|---------------|
| API 규격 (`apiFormat`) | 요청 엔드포인트 및 페이로드 형식 (`chat_completions`, `responses`, `claude`, `gemini`) | `chat_completions` |
| 모델 (`model`) | API 키로 실시간 조회한 드롭다운에서 선택하거나 직접 입력 | 동적 조회 (`[불러오기]`) |
| 만점 (`maxScore`) | 과제 총점 기준 | `100` (`1 ~ 100,000`) |
| 최대 출력 토큰 (`maxOutputTokens`) | AI 응답 최대 토큰 수 (표준 모델 호환을 위해 `1,024`부터 설정 가능) | `16,384` (`1,024 ~ 65,536`) |
| 분할 재시도 묶음 크기 | 일괄 채점 시 JSON 파싱 오류가 발생하면 줄여서 재시도하는 학생 수 | `13명` (이후 절반씩 축소) |
| API 오류 재시도 | 일시적 오류(`429`, `5xx` 등) 발생 시 자동 재시도 간격 및 최대 횟수 | `10분` 간격, 최대 `5회` |
| 최대 실행 시간 (`maxRuntimeMs`) | 단일 트리거 실행당 최대 작업 시간 (Apps Script 6분 제한 보호) | `330,000ms` (5분 30초) |

---

## 라이선스

MIT License
