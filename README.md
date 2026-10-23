# AI COA 자동검증기

협력사가 Outlook으로 보낸 COA(시험성적서) 메일을 도착 즉시 감지해 첨부 PDF를 저장하고, 항목·값·단위를 읽어 사내 규격(Spec)과 자동 대조합니다.
단위 환산과 합부 판정은 정해진 규칙대로 Python 코드가 하고(AI·파서는 읽기만), 모르는 값은 추측하지 않고 "검토 필요"로 넘깁니다.
부적합·검토 필요 건은 담당자에게 알림 메일을 보내고, 결과는 DB·엑셀·통보서 초안으로 남겨 사람은 예외 건만 확인하면 됩니다.

## 설치

필요: Windows, **클래식 Outlook**(새 Outlook은 COM 미지원), **Python 3.12.5**

```bat
py -3.12 --version
py -3.12 -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.venv\Scripts\python -m pytest
```

Python 3.12.5가 없으면 [python.org](https://www.python.org/downloads/release/python-3125/)에서 Windows installer(64-bit)를 받아 설치합니다.

## 실행

| 무엇 | 실행 | 설명 |
| --- | --- | --- |
| 메일 감시 | `run_watcher.bat` | 시작 시 최근 미처리 COA 메일을 먼저 처리하고, 새 메일을 기다립니다. 종료는 `Ctrl+C` |
| 화면 | `run_app.bat` | 브라우저에 처리 현황·판정 상세·수동 업로드·이력 검색 화면이 열립니다 |
| PDF 1장 판정 | `.venv\Scripts\python -m core.pipeline samples\COA_Nippon_N2418_fail.pdf` | 판정표만 출력 (`--save`면 저장까지) |
| Outlook 연결 진단 | `.venv\Scripts\python scripts\outlook_smoke_test.py` | 연결·첨부 저장·새 메일 이벤트 확인 |
| 첨부 진단 | `.venv\Scripts\python scripts\diag_attachments.py "메일 제목"` | 첨부 속성·PDF 판별 결과(읽기 전용) |
| 시연 초기화 | `.venv\Scripts\python scripts\reset_demo.py` | DB·inbox·output 삭제 후 샘플 재생성 (Outlook 메일 분류는 유지) |

감시 프로그램과 화면은 같은 DB(SQLite, WAL 모드)를 함께 씁니다. 감시 로그 첫 줄과 화면 맨 아래에 DB 파일 경로가 나오니 같은지 확인하세요.

## config.toml

| 항목 | 설명 |
| --- | --- |
| `[paths]` inbox, output, db, spec_master, logs, samples | 저장 위치 (프로젝트 폴더 기준 상대 경로) |
| `[mail] account` | COA를 받는 Outlook 계정. 이 계정의 받은편지함만 감시 (비우면 기본 계정) |
| `[mail] subject_keyword` | 제목에 이 글자가 있으면 COA 메일 후보 (기본 `COA`) |
| `[mail] allowed_domains` | 규격 마스터의 협력사 도메인 외에 추가로 허용할 도메인 |
| `[mail] test_senders` | 시연용 발신자 주소 (협력사 역할) |
| `[mail] processed_category` | 처리 완료 메일에 붙이는 Outlook 분류 (기본 `COA검증완료`) |
| `[mail] backlog_days` | 시작 시 며칠 전 메일까지 미처리분을 훑을지 |
| `[alert] mode` | `"display"`(알림 창만 띄움) / `"send"`(발송) |
| `[alert] to` | 담당자 알림 수신 주소 목록 |
| `[watch] mode` | `"outlook"` 또는 `"folder"`(Outlook 연동이 막혔을 때 inbox 폴더 감시, 알림 메일 없음) |
| `[watch] poll_seconds` | 이벤트를 놓친 메일을 다시 훑는 주기(초) |
| `[extract] llm_fallback` | LLM 보조 추출 (확장 단계, 항상 false) |

COA 메일 조건: (제목 키워드 **또는** 협력사·허용 도메인·시연 발신자) **그리고** PDF 첨부. 우리가 보낸 알림(제목이 `[COA `로 시작)은 제외합니다.
규격·협력사·별칭은 `data/spec_master.xlsx`(시트 materials·spec·suppliers)에서 현업이 직접 고칩니다.

## 폴더 구조

```text
COA_checker/
├─ mail_watcher.py        # 메일 감시: 감지 → 첨부 저장 → 판정 → 알림 → 처리 완료 분류
├─ app.py                 # Streamlit 화면
├─ config.toml            # 설정
├─ run_watcher.bat, run_app.bat
├─ core/
│  ├─ config.py           # 설정 로딩
│  ├─ models.py           # 데이터 모델 (pydantic)
│  ├─ units.py            # 값 파싱·단위 환산 (Decimal)
│  ├─ mapper.py           # 항목명 별칭 → 표준 항목
│  ├─ spec.py             # 규격 마스터 로딩
│  ├─ extractor.py        # PDF → 항목·값·단위 (pdfplumber, 표/줄 나열 양식)
│  ├─ judge.py            # 판정 PASS / FAIL / REVIEW
│  ├─ reporter.py         # SQLite·결과 엑셀·ERP 양식·통보서 초안·원본 보관
│  ├─ pipeline.py         # process_pdf(): 모든 경로의 단일 진입점
│  └─ outlook.py          # 메일 선별·첨부 저장·분류·알림 메일
├─ scripts/               # make_samples, reset_demo, outlook_smoke_test, diag_attachments
├─ data/spec_master.xlsx  # 규격 마스터
├─ samples/               # 시연용 COA PDF 5장 + expected.json(정답지)
├─ inbox/                 # 받은 COA 원본 (inbox/YYYYMMDD/{수신일시}_{발신처}_{원본명}.pdf)
├─ output/                # coa_history.db, results/ 결과 엑셀, erp_upload_YYYYMMDD.xlsx, notices/ 통보서 초안
├─ logs/                  # mail_watcher_YYYYMMDD.log
├─ docs/                  # 기획서, 시연 체크리스트, 발표 문구
└─ tests/                 # pytest
```

## 판정 규칙

- 상태: **PASS**(적합) / **FAIL**(부적합) / **REVIEW**(검토 필요). 문서 판정은 FAIL이 하나라도 있으면 FAIL, 아니면 REVIEW가 있으면 REVIEW, 모두 PASS면 PASS.
- 수치 비교는 Decimal, 경계값 포함(규격과 같으면 PASS). 규격 마스터 숫자도 `Decimal(str(값))`으로 읽습니다.
- 단위 환산(기준 ppm): `% = wt% = w/w% = 10000 ppm`, `ppm = mg/kg = μg/g`, `ppb = μg/kg = 0.001 ppm`, `g/kg = 1000 ppm`. APHA 같은 무차원 단위는 같은 단위끼리만 비교.
- `<x`, `≤x`: x가 상한 이하면 PASS. 하한이 있는 항목은 실제 값이 하한 미만으로 확정될 때(`<x`는 x ≤ 하한, `≤x`는 x < 하한) FAIL, 그 밖에는 REVIEW.
- REVIEW: `N.D.`, `>x`, 범위(`0.5~1.0`), 빈 값·이상한 값, 단위 미기재, 모르는 단위·환산 불가, Spec 항목 누락, 같은 항목 중복, 재질 미확인, 추출 경고(스캔 PDF 등).
- 추출 경고는 PASS를 REVIEW로만 올리고, FAIL은 그대로 FAIL.
- COA에만 있는 항목은 기록만 하고 판정에서 제외. 항목별 원문 값·단위·환산 값·규격·판정·사유·Spec 버전을 DB에 남깁니다.
- 협력사에는 아무것도 자동 발송하지 않습니다(통보서는 초안 파일까지만).

## 로그온 시 메일 감시 자동 실행 (Windows 작업 스케줄러)

> Outlook이 먼저 떠 있어야 안정적입니다. 창 없이 백그라운드로 뜬 Outlook에 연결하면 멈출 수 있으므로, Outlook도 로그온 시 자동 실행되게 하고 감시는 1분 늦게 시작하세요. `run_watcher.bat` 마지막의 `pause` 때문에 종료 시 창이 남습니다(오류 확인용).

**GUI 절차**
1. 시작 메뉴 → "작업 스케줄러" 실행 → 오른쪽 **작업 만들기…**
2. **일반**: 이름 `COA 메일 감시`, "사용자가 로그온할 때만 실행" 선택 (Outlook이 사용자 세션에 있어야 함). "가장 높은 권한으로 실행"은 **체크하지 않음** (Outlook과 권한 수준이 다르면 연결 실패)
3. **트리거** → 새로 만들기 → 작업 시작: **로그온할 때**, 특정 사용자: 본인, 고급 설정 → **작업 지연 시간: 1분**
4. **동작** → 새로 만들기 → 프로그램/스크립트: `C:\Users\LG\Desktop\AI자동화프로젝트\COA_checker\run_watcher.bat`, 시작 위치: `C:\Users\LG\Desktop\AI자동화프로젝트\COA_checker`
5. **조건**: "컴퓨터의 AC 전원이 켜져 있는 경우에만 작업 시작" 해제(노트북이면)
6. **설정**: "작업이 이미 실행 중이면: 새 인스턴스 실행 안 함"
7. 확인 → 목록에서 오른쪽 클릭 **실행**으로 시험

**명령으로 등록** (명령 프롬프트, 경로는 본인 PC에 맞게)

```bat
schtasks /Create /TN "COA 메일 감시" /SC ONLOGON /DELAY 0001:00 /RL LIMITED /TR "\"C:\Users\LG\Desktop\AI자동화프로젝트\COA_checker\run_watcher.bat\""
```

- 확인: `schtasks /Query /TN "COA 메일 감시" /V /FO LIST` · 즉시 실행: `schtasks /Run /TN "COA 메일 감시"` · 삭제: `schtasks /Delete /TN "COA 메일 감시" /F`
- 회사 정책에 따라 `/SC ONLOGON` 등록에 관리자 권한이 필요할 수 있습니다. 그 경우 GUI 절차를 쓰거나 IT에 요청하세요.

## 한계와 확장 계획

| 한계 (현재) | 확장 계획 |
| --- | --- |
| 텍스트가 있는 PDF만 읽음. 스캔 PDF는 "검토 필요" | **OCR**(PaddleOCR/Tesseract, Python 3.12 호환 확인 후) |
| 결과표가 여러 페이지에 걸치면 첫 표만 사용 | 페이지 간 표 병합 추출 |
| 규칙 기반 추출이라 처음 보는 양식은 "검토 필요"가 늘 수 있음 | 사내 승인 LLM 보조 추출(`extract.llm_fallback`, 판정은 계속 코드가 수행) |
| ERP는 업로드 양식 엑셀까지 | **ERP 직접 연동**(API 또는 RPA) |
| 이 PC의 클래식 Outlook이 켜져 있어야 동작, 받은편지함만 감시(정크로 빠진 메일은 처리 안 됨) | **Microsoft Graph API** 서버형 메일 연동(IT 승인 필요) → PC가 꺼져 있어도 동작 |
| 협력사별 추세는 이력 검색으로만 확인 | 협력사별 부적합률 대시보드 |

자세한 배경과 기대 효과는 [docs/기획서.md](docs/기획서.md), 시연 절차는 [docs/시연_체크리스트.md](docs/시연_체크리스트.md)를 참고하세요.
