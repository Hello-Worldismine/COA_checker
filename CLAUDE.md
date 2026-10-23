# CLAUDE.md — AI COA 자동검증기

이 파일은 Claude Code가 매 세션 자동으로 읽는 프로젝트 지침입니다. 상세 요구사항은 `docs/기획서.md`에 있습니다.

## 프로젝트 한 줄 요약
협력사가 Outlook으로 보낸 COA(시험성적서) 메일을 즉시 감지해 첨부 PDF를 저장하고, 항목·값·단위를 추출해 사내 규격(Spec)과 자동 대조한 뒤, 부적합·검토 필요 건을 담당자에게 알리는 Python 자동화 도구.

## 마감과 우선순위
- 발표 마감: 2026-10-23 14:00 (3분 시연 영상 포함). 시간이 부족하다.
- 우선순위: **동작하는 시연 흐름 > 코드 완성도 > 확장 기능**.
- 확장 범위(OCR, ERP 직접 연동, LLM 보조 추출 실구현, Graph API)는 구현하지 않는다. 인터페이스 자리만 남겨 둔다.

## 실행 환경 (반드시 지킬 것)
- **Python 3.12.5** 고정. 가상환경 `.venv` 사용 (`py -3.12 -m venv .venv`).
- OS: **Windows**, Outlook은 **클래식 Outlook**(COM 연동). 새 Outlook(New Outlook)은 지원 안 됨.
- 의존성은 `requirements.txt`에 `==`로 버전 고정. 새 패키지를 추가하면 반드시 requirements.txt도 갱신.
- 회사 PC일 수 있으므로 외부 네트워크/외부 AI API 호출 코드를 기본 경로에 넣지 않는다.

## 기술 스택
| 용도 | 라이브러리 |
| --- | --- |
| Outlook 연동 | pywin32 (`win32com.client`, `pythoncom`) |
| PDF 추출 | pdfplumber |
| 데이터 모델·검증 | pydantic v2 |
| 엑셀 입출력 | openpyxl, pandas |
| 저장 | sqlite3 (내장) |
| GUI | streamlit |
| 대안 경로(폴더 감시) | watchdog |
| 설정 | tomllib (내장, config.toml) |
| 로그 | logging (내장) |
| 테스트 | pytest |
| 샘플 PDF 생성(개발용) | reportlab |

단위 환산은 외부 라이브러리 대신 `core/units.py`의 **자체 환산표 + Decimal**로 구현한다(결정적이고 설명 가능해야 하므로).

## 폴더 구조
```text
coa_checker/
├─ CLAUDE.md
├─ docs/기획서.md
├─ mail_watcher.py        # 백그라운드 상시 실행: 메일 감지 → 첨부 저장 → 판정 → 알림
├─ app.py                 # Streamlit 화면: 처리 현황·판정표·수동 업로드·이력 검색
├─ config.toml            # 경로, 협력사 도메인, 알림 수신자, 동작 모드
├─ core/
│  ├─ __init__.py
│  ├─ config.py           # config.toml 로딩
│  ├─ models.py           # pydantic 모델 (ExtractedItem, CoaDocument, ItemResult, CoaResult)
│  ├─ units.py            # 단위 정규화 (Decimal)
│  ├─ mapper.py           # 항목명 별칭 → 사내 표준 항목
│  ├─ extractor.py        # PDF → CoaDocument (규칙 기반, LLM 훅은 비활성)
│  ├─ judge.py            # 규격 판정: PASS / FAIL / REVIEW
│  ├─ spec.py             # spec_master.xlsx 로딩
│  ├─ reporter.py         # SQLite 저장, 결과 엑셀, 통보서 초안
│  ├─ pipeline.py         # process_pdf(): 위 단계를 묶는 단일 진입점
│  └─ outlook.py          # Outlook 메일 선별·첨부 저장·분류 표시·메일 발송
├─ scripts/
│  ├─ outlook_smoke_test.py
│  └─ make_samples.py     # 샘플 COA PDF·spec_master.xlsx 생성
├─ data/spec_master.xlsx
├─ inbox/                 # 메일에서 저장된 COA PDF
├─ output/                # 결과 엑셀, 통보서 초안
├─ samples/               # 시연용 COA PDF
├─ logs/
└─ tests/
```

## 핵심 설계 원칙 (어기지 말 것)
1. **AI/파서는 읽기만, 판정은 Python 코드가 한다.** 합부 판정은 `judge.py`의 결정적 규칙으로만 수행한다.
2. **모르면 추측하지 말고 REVIEW(검토 필요).** 항목 누락, 모르는 단위, 이상한 숫자 형식, 추출 실패는 모두 REVIEW.
3. **수치 비교는 Decimal.** float 비교 금지. 경계값은 포함(이하/이상 → 같으면 PASS). 엑셀 등에서 읽은 숫자는 `Decimal(float)`가 아니라 `Decimal(str(값))`으로 변환한다(STEP 3 spec 로딩에 반영).
4. **단일 진입점.** 메일 경로·폴더 감시·수동 업로드 모두 `core/pipeline.process_pdf()`를 호출한다.
5. **판정 근거를 남긴다.** 항목별 원문 값, 원문 단위, 환산 값, 규격 하한/상한, 판정, 사유, 적용 Spec 버전을 DB에 저장한다.
6. **협력사에 자동 발송하지 않는다.** 협력사 통보서는 초안(임시보관함 저장 또는 파일)까지만.
7. 담당자 알림 메일은 `config.toml`의 `alert.mode`로 제어: `"display"`(창만 띄움) / `"send"`(발송). 기본값 `"display"`.

## 판정 규칙 요약
- 상태: `PASS`(적합) / `FAIL`(부적합) / `REVIEW`(검토 필요). 문서 전체 판정 = FAIL이 하나라도 있으면 FAIL, 아니면 REVIEW가 있으면 REVIEW, 모두 PASS면 PASS.
- 특수 표기: `<x`, `≤x` → x가 상한 이하이면 PASS, 아니면 REVIEW. `N.D.`, `ND`, `Not Detected`, `>x`, 범위(`0.5~1.0`), 빈 값 → REVIEW.
- 하한이 있는 항목에 `<x`/`≤x`가 오면: 실제 값이 하한 미만으로 확정될 때(`<x`는 x ≤ 하한, `≤x`는 x < 하한) FAIL, 그 밖에는 REVIEW.
- 추출 경고(Lot·협력사 미확인 등)는 문서 판정을 PASS → REVIEW로만 올린다. FAIL 항목이 있으면 경고가 있어도 FAIL 유지.
- Spec에 있는데 COA에 없는 항목 → REVIEW("항목 누락"). COA에만 있는 항목 → 기록만 하고 판정 제외.
- 단위가 비어 있는 숫자 값 → REVIEW(사유 "단위 미기재 - 확인 필요"). 규격 단위로 간주하지 않는다.
- 질량분율 단위 환산(기준 ppm): `% = wt% = w/w% = 10000 ppm`, `ppm = mg/kg = μg/g = ug/g`, `ppb = μg/kg = 0.001 ppm`, `g/kg = 1000 ppm`. 무차원 단위(예: APHA)는 동일 단위끼리만 비교, 다르면 REVIEW.

## 시연 조건 (STEP 7에서 반영)
- 수신 계정 `adm1nsubin@outlook.kr`, 발신자 `jaemin9905@naver.com`(협력사 역할, `config.toml`의 `mail.test_senders`), 담당자 `alstnqls02@naver.com`.
- 메일 1통에 `samples/`의 PDF 5장을 모두 첨부해 보낸다. 제목에 `COA` 포함.
- 첨부가 여러 개면 첨부별로 각각 `process_pdf()`를 실행하고, 하나가 실패해도 나머지는 계속 처리한다.
- 담당자 알림은 첨부별이 아니라 **메일 1통당 요약 알림 1통**: 문서별 판정 요약 + FAIL/REVIEW 항목 상세.
- 기대 결과는 `samples/expected.json`(N2419는 Water `<10 ppm` → PASS, Fe `N.D.` → REVIEW, 문서 REVIEW).

## 코딩 규칙
- 함수/모듈에 한국어 docstring 한두 줄. 사용자 메시지·로그는 한국어.
- 경로는 `pathlib.Path`, 설정값 하드코딩 금지(`config.toml`).
- COM 객체는 스레드 간 공유 금지. 메일 이벤트 핸들러는 EntryID만 큐에 넣고, 메인 루프에서 처리.
- 각 단계가 끝나면: `pytest` 통과 확인 → 변경 요약 → `git commit` (메시지는 한국어로 단계명 포함).
- 단계 범위를 넘는 기능은 만들지 말고, 필요해 보이면 "다음 단계 제안"으로 적고 멈춘다.
