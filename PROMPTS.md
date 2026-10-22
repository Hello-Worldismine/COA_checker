# Claude Code 단계별 프롬프트 — AI COA 자동검증기

## 사용 방법

1. **Outlook이 설치된 Windows PC**에 프로젝트 폴더(예: `C:\coa_checker`)를 만든다.
2. 아래처럼 파일을 넣는다.
   ```text
   coa_checker/
   ├─ CLAUDE.md          ← 받은 파일 그대로 (Claude Code가 매번 자동으로 읽음)
   ├─ PROMPTS.md         ← 이 파일 (참고용)
   └─ docs/기획서.md      ← 받은 기획서
   ```
3. 그 폴더에서 Claude Code를 실행하고, **STEP 0부터 하나씩** 아래 코드 블록 안의 내용을 복사해 보낸다.
4. 각 단계가 끝나면 Claude Code의 보고를 읽고, "내가 직접 확인할 것"을 해 본 뒤 다음 STEP을 보낸다.
5. 대화가 길어져 느려지면 `/clear` 후 다음 STEP을 보내도 된다. 맥락은 CLAUDE.md·기획서·코드에 남아 있다.

| STEP | 내용 | 예상 소요 | 시간이 없을 때 |
| --- | --- | --- | --- |
| 0 | 환경·뼈대 | 15분 | 필수 |
| 1 | Outlook 연동 시험 | 10분 | 필수 (가장 먼저 위험 확인) |
| 2 | 샘플 COA·규격 마스터 | 20분 | 필수 |
| 3 | 모델·단위·매핑 | 30분 | 필수 |
| 4 | PDF 추출 | 40분 | 필수 |
| 5 | 판정·파이프라인 | 30분 | 필수 |
| 6 | 저장·엑셀·통보서 | 30분 | 통보서는 생략 가능 |
| 7 | 메일 감시 프로그램 | 40분 | 필수 (시연 핵심) |
| 8 | Streamlit 화면 | 40분 | 이력 검색 생략 가능 |
| 9 | 통합 리허설·정리 | 30분 | 필수 |

---

## STEP 0 — 개발 환경과 프로젝트 뼈대

```text
CLAUDE.md와 docs/기획서.md를 먼저 읽고 프로젝트를 이해해줘. 이번 단계는 "환경과 뼈대"만 만든다.

할 일:
1. `py -3.12 --version`으로 Python 3.12.5가 있는지 확인하고, .venv 가상환경을 만들어줘. 3.12.5가 없으면 진행하지 말고 설치 방법을 알려줘.
2. CLAUDE.md의 기술 스택 라이브러리를 설치하고 requirements.txt에 == 로 버전을 고정해줘. 각 패키지가 Python 3.12에서 설치·import 되는지 확인해줘.
3. CLAUDE.md의 폴더 구조대로 빈 모듈 파일과 폴더(inbox, output, samples, logs, data, tests)를 만들어줘. 각 모듈 맨 위에는 역할을 설명하는 한국어 docstring만 넣어줘.
4. config.toml을 만들어줘. 최소 항목:
   - [paths] inbox, output, db, spec_master, logs
   - [mail] subject_keyword = "COA", allowed_domains(=spec_master의 협력사 도메인과 별개로 추가 허용할 도메인), test_senders(시연용 사내 발신자 주소 목록, 기본 빈 목록), processed_category = "COA검증완료", backlog_days = 7
   - [alert] mode = "display", to = [] (담당자 메일 주소)
   - [watch] mode = "outlook"  # "outlook" 또는 "folder"
   - [extract] llm_fallback = false
5. core/config.py에 config.toml을 읽는 함수를 만들어줘 (tomllib 사용).
6. .gitignore(.venv, __pycache__, inbox/*, output/*, logs/*, *.db)를 만들고 git init 후 첫 커밋.

완료 기준: `.venv\Scripts\python -c "import win32com.client, pdfplumber, pydantic, openpyxl, pandas, streamlit, watchdog, reportlab"`가 오류 없이 실행되고, `pytest`가 (테스트 0개라도) 정상 종료.

끝나면 ①만든 파일 목록 ②실행 결과 ③내가 직접 확인할 것 ④다음 단계 전에 내가 정해야 할 것을 짧게 보고하고 멈춰.
```

---

## STEP 1 — Outlook 연동 시험 (가장 큰 위험을 먼저 확인)

```text
이번 단계는 회사 PC에서 Python이 Outlook에 접근할 수 있는지 확인하는 시험 스크립트만 만든다. 본 기능 코드는 아직 만들지 마.

scripts/outlook_smoke_test.py를 만들어줘:
1. win32com.client로 Outlook.Application에 연결하고 기본 받은편지함 이름과 메일 개수를 출력.
2. 받은편지함 최근 메일 20통 중 PDF 첨부가 있는 메일 1통을 찾아 첨부를 inbox/_smoke_test/ 에 저장하고 경로 출력. 없으면 "PDF 첨부 메일 없음"을 출력.
3. 그 메일의 발신자 주소를 출력. 사내(Exchange) 발신자는 SenderEmailAddress가 "/O=..." 형태로 나오므로, Sender.GetExchangeUser().PrimarySmtpAddress로 SMTP 주소를 얻는 함수도 함께 시험해줘.
4. NewMailEx 이벤트를 30초 동안 기다리면서 새 메일이 오면 제목을 출력 (DispatchWithEvents + pythoncom.PumpWaitingMessages 루프).
5. 실패 지점마다 원인 추정 메시지를 한국어로 출력: 새 Outlook(COM 미지원), Outlook 미실행, 보안 정책 차단, 프로그래밍 방식 액세스 경고 등.

실행 방법을 알려주고, 내가 실행한 결과를 붙여넣으면 판단해줘:
- 모두 성공 → config.toml의 watch.mode = "outlook" 유지
- 연결 자체가 막힘 → watch.mode = "folder"로 바꾸고, Outlook 규칙(특정 조건 메일의 첨부를 폴더에 저장)을 쓰는 대안 절차를 안내해줘.

보고 형식은 STEP 0과 같게 하고 멈춰.
```

> 실행 결과를 Claude Code에 그대로 붙여넣어 판단을 받은 뒤 STEP 2로 넘어가세요. 시험 중에는 자기 자신에게 PDF 첨부 메일을 한 통 보내 보세요.

---

## STEP 2 — 시연용 샘플 COA PDF와 규격 마스터

```text
이번 단계는 시연·테스트용 데이터만 만든다. scripts/make_samples.py 하나로 아래 파일들을 생성해줘 (reportlab, openpyxl 사용). 실행하면 언제든 다시 만들 수 있어야 해.

1) data/spec_master.xlsx — 시트 3개
- materials: material_code | material_name | name_aliases
  - NMP-EL | N-Methyl-2-pyrrolidone Electronic Grade | NMP, N-Methyl-2-pyrrolidone, 1-Methyl-2-pyrrolidinone
- spec: material_code | item_std | aliases | lower | upper | unit | spec_version
  - NMP-EL | Purity   | Purity, Assay, Purity (GC), 순도            | 99.90 |       | %    | v1.0
  - NMP-EL | Moisture | Moisture, Water, Water content, H2O, KF, 수분 |       | 0.003 | %    | v1.0
  - NMP-EL | Iron     | Iron, Fe, Iron (Fe), 철                      |       | 0.005 | %    | v1.0
  - NMP-EL | Color    | Color, Colour, APHA, 색도                    |       | 10    | APHA | v1.0
- suppliers: supplier_name | email_domain | name_aliases
  - Alpha Chemical Co., Ltd. | alpha-chem.example | Alpha Chemical
  - Nippon Solvent Co., Ltd. | nippon-solv.example | Nippon Solvent

2) samples/ — 영문 COA PDF 5장. 실제 협력사처럼 두 회사의 양식을 다르게 만들어줘.
- Alpha 양식: 상단 회사명·제품명·Lot No.·Mfg. Date, 아래 표(Test Item | Specification | Result | Unit)
- Nippon 양식: 표 없이 "Assay ........ 99.93 %" 같은 줄 나열, 단위가 값 뒤에 붙음
- 파일:
  - COA_Alpha_A2401_pass.pdf   : Purity 99.95 % / Moisture 0.002 % / Iron (Fe) 0.003 % / Color 5 APHA → 적합
  - COA_Nippon_N2417_pass.pdf  : Assay 99.93 % / Water 20 ppm / Fe 10 ppm / Colour 8 APHA → 적합 (ppm 환산 필요)
  - COA_Alpha_A2402_fail.pdf   : Iron (Fe) 0.008 % (나머지 정상) → 부적합 (기획서 사례 1)
  - COA_Nippon_N2418_fail.pdf  : Water 50 ppm (나머지 정상) → 부적합, 50 ppm = 0.005% > 0.003% (기획서 사례 2)
  - COA_Nippon_N2419_review.pdf: Fe "N.D." / Water "<10 ppm" (나머지 정상) → 검토 필요
- 각 PDF의 기대 판정을 samples/expected.json에 기록해줘 (파일명 → 문서 판정, 항목별 판정). 이후 테스트의 정답지로 쓴다.

완료 기준: 스크립트 실행 후 PDF 5장, expected.json, spec_master.xlsx가 생기고, pdfplumber로 각 PDF 텍스트가 추출되는지 확인 출력.

보고 형식은 이전과 같게 하고 멈춰.
```

---

## STEP 3 — 데이터 모델, 단위 정규화, 항목명 매핑

```text
이번 단계는 판정의 기초 부품만 만든다 (PDF 추출은 다음 단계).

1. core/models.py (pydantic v2)
   - ExtractedItem: raw_name, raw_value(str), raw_unit(str|None)
   - CoaDocument: source_path, supplier, material_code, product_name, lot_no, mfg_date, items: list[ExtractedItem], extract_warnings: list[str]
   - ItemResult: item_std, raw_name, raw_value, raw_unit, normalized_value(Decimal|None), spec_unit, lower, upper, status(PASS/FAIL/REVIEW), reason
   - CoaResult: document, item_results, overall_status, spec_version, processed_at
2. core/units.py
   - parse_value(raw: str) → (kind, Decimal|None): kind는 "number", "less_than", "not_detected", "greater_than", "range", "invalid". 쉼표 소수점·공백·전각 문자 정리 포함.
   - normalize_unit(raw_unit) → 표준 단위 문자열 또는 None (%, wt%, w/w% → "%" / mg/kg, μg/g, ug/g → "ppm" 등, 대소문자·공백 무시)
   - convert(value: Decimal, from_unit, to_unit) → Decimal|None: CLAUDE.md의 질량분율 환산표 사용. 변환 불가하면 None. 무차원 단위(APHA)는 같은 단위일 때만 그대로 반환.
3. core/spec.py: spec_master.xlsx를 읽어 재질별 규격 목록, 항목 별칭 사전, 협력사(도메인·이름 별칭), 재질 이름 별칭을 제공. Spec 버전도 함께.
4. core/mapper.py: 원문 항목명 → item_std. 소문자화·괄호·공백·특수문자 정리 후 별칭 사전으로 매칭. 못 찾으면 None.
5. tests/test_units.py, tests/test_mapper.py — 최소 포함:
   - 50 ppm → 0.005 %, 0.008 % → 80 ppm, 10 ppm → 0.001 %, 1 ppb → 0.001 ppm
   - "<10", "≤ 10", "N.D.", "ND", "0,003", " 99.95 ", "0.5~1.0", "", "abc" 파싱
   - "Iron (Fe)", "FE", "Water content", "KF" 매핑
   - Decimal 비교로 0.003 == 0.0030 이 같게 처리되는지

완료 기준: pytest 전부 통과.
보고 형식은 이전과 같게 하고, git commit 후 멈춰.
```

---

## STEP 4 — PDF 데이터 추출

```text
이번 단계는 core/extractor.py만 만든다. 판정은 아직 하지 마.

extract(pdf_path) → CoaDocument
1. pdfplumber로 텍스트와 표를 모두 읽는다.
2. 표가 있으면 헤더(Test Item/Item/Parameter, Result/Value, Unit)를 찾아 행을 ExtractedItem으로 만든다.
3. 표가 없으면 줄 단위 정규식으로 "항목명 ...(점선/공백)... 값 단위" 패턴을 추출한다. 단위가 값에 붙은 경우("20ppm")도 분리.
4. 헤더 영역에서 협력사명, 제품명, Lot No., Mfg. Date를 정규식으로 추출하고, spec 모듈의 별칭으로 supplier와 material_code를 정한다.
5. 텍스트가 거의 없으면(스캔 PDF) items를 비우고 extract_warnings에 "텍스트 없음 - 스캔 PDF 의심(OCR은 확장 단계)"을 넣는다.
6. config의 extract.llm_fallback이 true일 때만 호출되는 llm_extract() 자리만 만들고, 지금은 NotImplementedError를 던지게 해줘 (호출되지 않음).
7. "Specification" 열의 값을 측정값으로 착각하지 않도록 주의해줘.

tests/test_extractor.py: samples/ 5장에 대해 항목 수, 항목명, 값, 단위, Lot No., 협력사, 재질 코드가 expected.json·샘플 생성 스크립트와 일치하는지 검사.

완료 기준: pytest 전부 통과 + `python -m core.extractor samples\COA_Nippon_N2418_fail.pdf` 실행 시 추출 결과를 표 형태로 출력.
보고 형식은 이전과 같게 하고, git commit 후 멈춰.
```

---

## STEP 5 — 규격 판정과 파이프라인

```text
이번 단계는 core/judge.py와 core/pipeline.py를 만든다.

1. judge.judge(doc: CoaDocument, specs) → CoaResult
   - CLAUDE.md "판정 규칙 요약"을 그대로 구현 (PASS/FAIL/REVIEW, 특수 표기, 누락 항목, 경계값 포함, Decimal).
   - 각 ItemResult.reason에 사람이 읽을 한국어 사유를 넣는다. 예: "50 ppm → 0.005 % 환산, 상한 0.003 % 초과", "N.D. 표기 - 검출한계 미기재로 검토 필요".
   - material_code를 못 정했거나 items가 비었으면 문서 전체 REVIEW.
2. pipeline.process_pdf(pdf_path, meta: dict | None = None) → CoaResult
   - extract → judge 순서로 실행. meta에는 메일 경로일 때 발신자, 수신시각, 메일 EntryID가 들어온다(다음 단계에서 사용).
   - 이 단계에서는 저장 없이 결과만 반환.
   - `python -m core.pipeline <pdf>`로 실행하면 판정표를 콘솔에 출력 (FAIL은 눈에 띄게).
3. tests/test_judge.py: samples 5장 전체가 expected.json의 문서 판정·항목별 판정과 일치. 경계값(정확히 0.003 %), 상한·하한 둘 다 있는 항목, 단위 변환 불가 케이스도 단위 테스트.

완료 기준: pytest 전부 통과 + 5장 모두 기대 판정과 일치.
보고 형식은 이전과 같게 하고, git commit 후 멈춰.
```

---

## STEP 6 — 결과 저장, 엑셀, 통보서 초안

```text
이번 단계는 core/reporter.py를 만들고 pipeline.process_pdf에 연결한다.

1. SQLite (config.paths.db)
   - documents: id, processed_at, source(mail/folder/upload), mail_entry_id(UNIQUE, NULL 허용), sender, received_at, file_path, supplier, material_code, lot_no, overall_status, spec_version
   - item_results: document_id, item_std, raw_name, raw_value, raw_unit, normalized_value(TEXT로 저장), spec_unit, lower, upper, status, reason
   - is_processed(mail_entry_id) 함수 (중복 처리 방지용)
2. 결과 엑셀: output/results/{처리일시}_{협력사}_{Lot}_결과.xlsx — 항목별 판정표, FAIL 셀 빨간 배경, REVIEW 노란 배경, 상단에 문서 판정·Spec 버전.
3. ERP 업로드 양식: output/erp_upload_{YYYYMMDD}.xlsx에 문서 1건당 1행 누적 (처리일시, 협력사, 재질코드, Lot, 판정, 항목별 측정값(규격 단위로 환산)).
4. 통보서 초안: FAIL 문서만 output/notices/{Lot}_부적합통보서_초안.txt 생성 (협력사명, Lot, 부적합 항목·측정값·규격, 정중한 영문+국문 문구). 자동 발송 없음.
5. 원본 PDF는 inbox/{YYYYMMDD}/{수신일시}_{협력사}_{원본파일명}.pdf 표준 이름으로 보관 (이미 거기 있으면 이동 생략).
6. process_pdf가 저장까지 수행하고 CoaResult와 생성 파일 경로들을 반환하도록 수정.

tests/test_reporter.py: 임시 폴더·임시 DB로 5장 처리 후 DB 행 수, 엑셀 생성, FAIL 2건에만 통보서 생성, 같은 mail_entry_id 재처리 시 중복 저장 안 됨을 검사.

완료 기준: pytest 전부 통과.
보고 형식은 이전과 같게 하고, git commit 후 멈춰.
```

---

## STEP 7 — Outlook 메일 감시 프로그램 (시연 핵심)

```text
이번 단계는 core/outlook.py와 mail_watcher.py를 만든다. STEP 1 시험 결과(config.watch.mode)를 따른다.

core/outlook.py
1. get_smtp_address(mail): Exchange 사내 발신자도 SMTP 주소로 변환.
2. is_coa_mail(mail, specs, config): 아래 조건 중 (제목 키워드 OR 협력사/허용 도메인/테스트 발신자) AND PDF 첨부 1개 이상. MailItem(Class 43)만.
3. save_pdf_attachments(mail) → 저장 경로 목록 (inbox 표준 경로, 파일명 금지문자 정리, 같은 이름 충돌 처리).
4. mark_processed(mail): config의 분류(카테고리) 추가 후 Save.
5. send_alert(result, files): 담당자(config.alert.to)에게 메일. 제목 "[COA 부적합] 협력사 / Lot" 또는 "[COA 검토 필요] …". HTML 본문에 항목별 판정표(FAIL 빨강, REVIEW 노랑), 사유, 결과 엑셀 첨부. alert.mode가 "display"면 Display(), "send"면 Send(). PASS는 알림 없음(로그만).

mail_watcher.py
1. 시작 시 로그 설정(logs/, 파일+콘솔), Outlook 연결, 받은편지함에서 backlog_days 이내 미처리 COA 메일을 먼저 처리(F3).
2. DispatchWithEvents로 NewMailEx 수신 → 핸들러는 EntryID를 queue에만 넣는다.
3. 메인 루프: pythoncom.PumpWaitingMessages() + 큐에서 꺼내 처리: is_coa_mail → is_processed 확인 → 첨부 저장 → 첨부별 process_pdf(meta 포함) → send_alert → mark_processed.
4. 메일 1통 처리 중 예외가 나도 프로그램은 계속 돈다 (로그에 스택트레이스, 해당 메일은 미처리로 남김).
5. config.watch.mode == "folder"이면 Outlook 대신 watchdog으로 inbox 하위 새 PDF를 감시해 같은 파이프라인 호출 (파일 쓰기 완료 대기 포함).
6. 콘솔에 한 줄 상태 출력: "[12:03:41] COA 수신: Nippon Solvent / N2418 → 부적합 (Moisture 50 ppm = 0.005 % > 0.003 %)".
7. 실행용 run_watcher.bat 생성 (.venv 활성화 후 실행).

Outlook이 필요한 부분은 테스트하기 어려우므로, is_coa_mail과 파일명 생성 로직은 가짜 메일 객체로 단위 테스트를 만들어줘.

내가 직접 할 시험 절차도 적어줘: watcher 실행 → 다른 계정(또는 나 자신)에서 제목 "COA N2418"로 COA_Nippon_N2418_fail.pdf 첨부 발송 → 몇 초 뒤 콘솔·inbox·output·알림 메일 확인 → 같은 메일 재처리 안 되는지 확인. 시연용으로 내 사내 주소를 config의 test_senders에 넣는 방법도 알려줘.

보고 형식은 이전과 같게 하고, git commit 후 멈춰.
```

> 이 단계 후 실제로 메일을 보내 시험하고, 결과(콘솔 로그)를 Claude Code에 붙여넣어 문제를 고친 뒤 넘어가세요.

---

## STEP 8 — Streamlit 화면

```text
이번 단계는 app.py (Streamlit) 화면을 만든다. 판정 로직은 새로 만들지 말고 SQLite와 pipeline만 사용해.

탭 4개:
1. 처리 현황: 오늘 처리 건수, 적합/부적합/검토 필요 건수 지표 + 최근 처리 목록 표(처리시각, 협력사, Lot, 판정). 판정은 색 배지(부적합 빨강, 검토 노랑, 적합 초록). 10초마다 자동 새로고침(현재 streamlit 버전에서 지원하는 방법 사용, 안 되면 새로고침 버튼).
2. 판정 상세: 목록에서 문서를 고르면 항목별 판정표(원문 값·단위 → 환산 값, 규격, 판정, 사유)와 원본 PDF 미리보기(또는 열기 버튼), 결과 엑셀·통보서 초안 다운로드 버튼.
3. 수동 업로드: PDF를 끌어다 놓으면 process_pdf(source="upload") 실행 후 즉시 판정표 표시.
4. 이력 검색: 협력사·Lot·기간·판정으로 검색, 결과 표와 엑셀 다운로드 (감사 대응용).

화면 상단 제목 "AI COA 자동검증기", 부제 "메일 도착 즉시 판정 · 사람은 예외 건만 확인". 시연 화면 녹화용이므로 글자 크기와 여백을 넉넉하게.
run_app.bat 생성.

완료 기준: `streamlit run app.py` 실행 후 STEP 7에서 처리된 데이터가 보이고, 수동 업로드로 5장 모두 기대 판정이 나온다. 내가 확인할 화면 체크리스트를 적어줘.
보고 형식은 이전과 같게 하고, git commit 후 멈춰.
```

---

## STEP 9 — 통합 리허설과 마무리

```text
마지막 단계다. 시연 영상 녹화 전에 전체 흐름을 점검하고 정리해줘.

1. scripts/reset_demo.py: DB, inbox, output을 비우고(확인 질문 후) samples를 다시 생성 → 녹화를 여러 번 다시 할 수 있게.
2. 전체 pytest 실행, 실패가 있으면 고쳐줘.
3. README.md(한국어): 프로젝트 소개 3줄, 설치(Python 3.12.5, venv, requirements), 실행(run_watcher.bat, run_app.bat), config.toml 설명, 폴더 구조, 판정 규칙, 한계와 확장 계획(OCR, ERP 연동, Graph API).
4. Windows 작업 스케줄러에 "로그온 시 run_watcher.bat 실행"을 등록하는 방법(GUI 절차와 schtasks 명령 둘 다)을 README에 추가. 실제 등록은 하지 마.
5. 기획서 8장 "3분 시연 시나리오"에 맞춘 녹화 체크리스트를 docs/시연_체크리스트.md로 만들어줘: 녹화 전 준비(reset_demo, watcher·app 실행, 메일 작성 창 미리 열기, 알림 메일 mode), 장면별 클릭 순서와 말할 문장, 실패 시 대체 장면(수동 업로드).
6. 발표 양식 PPT의 "자동화 흐름" 칸과 "소요시간/수행주기" 칸에 넣을 문구를 docs/발표_문구.md로 정리해줘 (As-Is 5~10분/건 → To-Be 예외 건만 1분 이내, 수행주기: 일).
7. 마지막 커밋.

끝나면 녹화 직전에 내가 할 일을 5줄 이내로 알려줘.
```

---

## 막혔을 때 쓰는 프롬프트

```text
방금 실행에서 아래 오류가 났어. 원인을 먼저 설명하고, 이번 단계 범위 안에서 최소한으로 고쳐줘. 고친 뒤 관련 테스트를 다시 돌려 결과를 보여줘.

[오류 메시지 / 콘솔 로그 붙여넣기]
```

```text
시간이 부족해. 지금 상태에서 3분 시연이 가능한 최소 흐름(메일 수신 → 판정 → 알림 메일 → 화면 표시)만 동작하게 남은 작업 우선순위를 정리하고, 가장 중요한 것부터 처리해줘. 테스트는 판정 로직만 유지해.
```
