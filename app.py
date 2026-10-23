"""Streamlit 화면: 처리 현황·판정표·수동 업로드·이력 검색.

DB는 읽기만 한다(조회할 때만 읽기 전용으로 짧게 연결 → 감시 프로그램이 쓰는 중에도 잠금 오류 없음).
판정 로직은 새로 만들지 않고, 수동 업로드만 core.pipeline.process_pdf(source="upload")를 호출한다.
실행: run_app.bat  또는  .venv\\Scripts\\streamlit run app.py
"""

import html
import io
import os
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd
import streamlit as st

from core.config import load_config
from core.pipeline import process_pdf, short_supplier
from core.reporter import get_document, get_item_results, read_rows, safe_part, search_documents

STATUS_LABEL = {"PASS": "적합", "FAIL": "부적합", "REVIEW": "검토 필요", "EXCLUDED": "판정 제외"}
STATUS_CLASS = {"PASS": "pass", "FAIL": "fail", "REVIEW": "review", "EXCLUDED": "excluded"}
SOURCE_LABEL = {"mail": "메일", "folder": "폴더", "upload": "업로드"}
REFRESH_SECONDS = 10

CSS = """
<style>
html, body, [data-testid="stAppViewContainer"] { font-size: 18px; }
.block-container { padding-top: 2.2rem; padding-bottom: 3rem; max-width: 1500px; }
h1 { font-size: 2.6rem !important; margin-bottom: 0 !important; }
.subtitle { font-size: 1.3rem; color: #555; margin: 0.2rem 0 1.2rem 0; }
button[data-baseweb="tab"] p { font-size: 1.2rem; font-weight: 600; }
[data-testid="stMetricValue"] { font-size: 2.6rem; }
[data-testid="stMetricLabel"] p { font-size: 1.1rem; }
.badge { display: inline-block; padding: 0.15rem 0.7rem; border-radius: 999px; font-weight: 700; white-space: nowrap; }
.badge.fail { background: #FFC7CE; color: #9C0006; }
.badge.review { background: #FFEB9C; color: #7F6000; }
.badge.pass { background: #C6EFCE; color: #006100; }
.badge.excluded { background: #EEEEEE; color: #555; }
.badge.big { font-size: 1.4rem; padding: 0.3rem 1.1rem; }
table.coa { border-collapse: collapse; width: 100%; font-size: 1.05rem; }
table.coa th { background: #305496; color: #fff; padding: 0.55rem 0.7rem; text-align: left; }
table.coa td { border-bottom: 1px solid #ddd; padding: 0.5rem 0.7rem; vertical-align: top; }
table.coa tr.fail td { background: #FFF1F3; }
table.coa tr.review td { background: #FFFBEA; }
.kv { font-size: 1.1rem; line-height: 1.9; }
</style>
"""


# ── 공통 ──────────────────────────────────────────────────────
def badge(status: str | None, big: bool = False) -> str:
    """판정 색 배지 HTML."""
    status = status or "REVIEW"
    cls = STATUS_CLASS.get(status, "excluded") + (" big" if big else "")
    return f'<span class="badge {cls}">{html.escape(STATUS_LABEL.get(status, status))}</span>'


def html_table(columns: list[str], rows: list[list[str]], row_status: list[str | None] | None = None) -> str:
    """큰 글씨 HTML 표. 셀 값은 이미 이스케이프된 HTML이어야 한다(배지 포함 가능)."""
    head = "".join(f"<th>{html.escape(c)}</th>" for c in columns)
    body = []
    for i, row in enumerate(rows):
        cls = STATUS_CLASS.get((row_status or [None] * len(rows))[i] or "", "")
        body.append(f'<tr class="{cls}">' + "".join(f"<td>{c}</td>" for c in row) + "</tr>")
    return f'<table class="coa"><tr>{head}</tr>{"".join(body)}</table>'


def esc(value) -> str:
    """None은 '-'로, 나머지는 HTML 이스케이프."""
    return "-" if value in (None, "") else html.escape(str(value))


def when(iso: str | None, with_date: bool = True) -> str:
    """ISO 시각을 화면용으로."""
    if not iso:
        return "-"
    t = datetime.fromisoformat(iso)
    return f"{t:%Y-%m-%d %H:%M:%S}" if with_date else f"{t:%H:%M:%S}"


def spec_text(lower: str | None, upper: str | None, unit: str | None) -> str:
    """규격 범위 문자열."""
    unit = unit or ""
    if lower and upper:
        return f"{lower}~{upper} {unit}"
    if upper:
        return f"≤ {upper} {unit}"
    if lower:
        return f"≥ {lower} {unit}"
    return "-"


def result_xlsx_path(doc: dict, output_dir: Path) -> Path | None:
    """결과 엑셀 경로 (STEP 8 이전 이력은 표준 이름으로 찾는다)."""
    if doc.get("result_xlsx"):
        return Path(doc["result_xlsx"])
    t = datetime.fromisoformat(doc["processed_at"])
    path = output_dir / "results" / (f"{t:%Y%m%d_%H%M%S}_{safe_part(doc['supplier'], '협력사미상')}_"
                                     f"{safe_part(doc['lot_no'], 'Lot미상')}_결과.xlsx")
    return path if path.exists() else None


def notice_path(doc: dict, output_dir: Path) -> Path | None:
    """통보서 초안 경로 (FAIL 문서만)."""
    if doc.get("notice_path"):
        return Path(doc["notice_path"])
    if doc["overall_status"] != "FAIL":
        return None
    path = output_dir / "notices" / f"{safe_part(doc['lot_no'], 'Lot미상')}_부적합통보서_초안.txt"
    return path if path.exists() else None


def items_table(items: list[dict]) -> str:
    """항목별 판정표 HTML: 원문 값·단위 → 환산 값, 규격, 판정, 사유."""
    rows, statuses = [], []
    for i in items:
        raw = f"{i['raw_value'] or ''} {i['raw_unit'] or ''}".strip() if i["raw_value"] is not None else ""
        norm = f"{i['normalized_value']} {i['spec_unit']}" if i["normalized_value"] else ""
        name = i["item_std"] or "(Spec 외)"
        if i["raw_name"] and i["raw_name"] != i["item_std"]:
            name += f" <span style='color:#777'>({esc(i['raw_name'])})</span>"
        rows.append([name if i["item_std"] else esc(i["raw_name"]), esc(raw), esc(norm),
                     esc(spec_text(i["lower"], i["upper"], i["spec_unit"])), badge(i["status"]), esc(i["reason"])])
        statuses.append(i["status"])
    return html_table(["항목", "원문 값", "환산 값", "규격", "판정", "사유"], rows, statuses)


def file_button(label: str, path: Path | None, mime: str, key: str) -> None:
    """파일 다운로드 버튼 (파일이 없으면 비활성)."""
    if path and path.is_file():
        st.download_button(label, path.read_bytes(), file_name=path.name, mime=mime, key=key, width="stretch")
    else:
        st.button(f"{label} (없음)", disabled=True, key=key, width="stretch")


# ── 탭 1: 처리 현황 ────────────────────────────────────────────
@st.fragment(run_every=REFRESH_SECONDS)
def status_panel(db: Path) -> None:
    """오늘 지표와 최근 처리 목록. 10초마다 이 부분만 다시 그린다."""
    today = date.today().isoformat()
    todays = search_documents(db, date_from=today, date_to=today)
    count = {s: sum(d["overall_status"] == s for d in todays) for s in ("PASS", "FAIL", "REVIEW")}
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("오늘 처리", f"{len(todays)}건")
    c2.metric("🟥 부적합", f"{count['FAIL']}건")
    c3.metric("🟨 검토 필요", f"{count['REVIEW']}건")
    c4.metric("🟩 적합", f"{count['PASS']}건")
    st.caption(f"마지막 갱신 {datetime.now():%H:%M:%S} · {REFRESH_SECONDS}초마다 자동 새로고침")

    recent = search_documents(db, limit=30)
    if not recent:
        st.info("아직 처리 이력이 없습니다. 감시 프로그램(run_watcher.bat)을 켜고 COA 메일을 보내 보세요.")
        return
    rows = [[esc(when(d["processed_at"])), esc(SOURCE_LABEL.get(d["source"], d["source"])),
             esc(short_supplier(d["supplier"])), esc(d["lot_no"]), badge(d["overall_status"]),
             esc(d["sender"]), esc(d["attachment_name"] or Path(d["file_path"]).name)] for d in recent]
    st.markdown(html_table(["처리시각", "경로", "협력사", "Lot", "판정", "발신자", "파일"], rows,
                           [d["overall_status"] for d in recent]), unsafe_allow_html=True)


# ── 탭 2: 판정 상세 ────────────────────────────────────────────
def detail_panel(db: Path, output_dir: Path) -> None:
    """문서를 골라 항목별 판정표, 원본 PDF, 결과 엑셀·통보서 초안을 본다."""
    docs = search_documents(db, limit=300)
    if not docs:
        st.info("아직 처리 이력이 없습니다.")
        return
    labels = {d["id"]: f"#{d['id']}  ·  {when(d['processed_at'])}  ·  {short_supplier(d['supplier'])} / "
                       f"{d['lot_no'] or 'Lot 미상'}  ·  {STATUS_LABEL[d['overall_status']]}" for d in docs}
    doc_id = st.selectbox("문서 선택 (최근 처리 순)", list(labels), format_func=labels.get)
    doc = get_document(db, doc_id)
    if doc is None:
        st.warning("문서를 찾을 수 없습니다.")
        return

    st.markdown(f"### {esc(short_supplier(doc['supplier']))} / {esc(doc['lot_no'])} &nbsp; "
                f"{badge(doc['overall_status'], big=True)}", unsafe_allow_html=True)
    left, right = st.columns(2)
    left.markdown(
        f"<div class='kv'><b>협력사</b> {esc(doc['supplier'])}<br><b>재질</b> {esc(doc['material_code'])} · "
        f"{esc(doc['product_name'])}<br><b>Lot</b> {esc(doc['lot_no'])} &nbsp; <b>제조일</b> {esc(doc['mfg_date'])}"
        f"<br><b>적용 Spec</b> {esc(doc['spec_version'])}</div>", unsafe_allow_html=True)
    right.markdown(
        f"<div class='kv'><b>경로</b> {esc(SOURCE_LABEL.get(doc['source'], doc['source']))} &nbsp; "
        f"<b>발신자</b> {esc(doc['sender'])}<br><b>수신</b> {esc(when(doc['received_at']))}<br>"
        f"<b>처리</b> {esc(when(doc['processed_at']))}<br><b>파일</b> {esc(Path(doc['file_path']).name)}</div>",
        unsafe_allow_html=True)
    if doc["reasons"]:
        st.warning(f"문서 검토 사유: {doc['reasons']}")

    st.markdown(items_table(get_item_results(db, doc_id)), unsafe_allow_html=True)
    st.write("")

    pdf = Path(doc["file_path"])
    b1, b2, b3, b4 = st.columns(4)
    with b1:
        file_button("📄 원본 PDF 다운로드", pdf, "application/pdf", f"pdf_{doc_id}")
    if b2.button("🖥️ 원본 PDF 열기(이 PC)", key=f"open_{doc_id}", width="stretch",
                 disabled=not (pdf.is_file() and hasattr(os, "startfile"))):
        os.startfile(pdf)  # 화면을 띄운 PC의 기본 PDF 뷰어로 연다
    with b3:
        file_button("📊 결과 엑셀", result_xlsx_path(doc, output_dir),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"xlsx_{doc_id}")
    with b4:
        file_button("✉️ 통보서 초안", notice_path(doc, output_dir), "text/plain", f"notice_{doc_id}")


# ── 탭 3: 수동 업로드 ──────────────────────────────────────────
def upload_panel(db: Path) -> None:
    """PDF를 올리면 process_pdf(source="upload")로 판정·저장하고 바로 판정표를 보여 준다."""
    st.write("메일이 아닌 경로로 받은 COA를 끌어다 놓으면 메일과 같은 엔진으로 판정·저장합니다.")
    files = st.file_uploader("COA PDF (여러 개 가능)", type=["pdf"], accept_multiple_files=True)
    done: dict = st.session_state.setdefault("uploaded", {})
    for f in files or []:
        if f.file_id in done:
            continue  # 화면이 다시 그려져도 같은 파일을 두 번 처리하지 않는다
        with st.spinner(f"{f.name} 판정 중..."):
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    path = Path(tmp) / Path(f.name).name
                    path.write_bytes(f.getvalue())
                    out = process_pdf(path, {"source": "upload", "attachment_name": f.name})
                done[f.file_id] = {"name": f.name, "doc_id": out.saved.document_id, "errors": out.saved.errors}
            except Exception as exc:  # 화면은 계속 동작해야 한다
                done[f.file_id] = {"name": f.name, "doc_id": None, "errors": [f"처리 실패: {exc}"]}

    for f in files or []:
        info = done.get(f.file_id)
        if not info:
            continue
        doc = get_document(db, info["doc_id"]) if info["doc_id"] else None
        st.markdown(f"#### {esc(info['name'])} &nbsp; {badge(doc['overall_status']) if doc else ''}",
                    unsafe_allow_html=True)
        for err in info["errors"]:
            st.error(err)
        if doc:
            st.caption(f"{doc['supplier'] or '협력사 미상'} · Lot {doc['lot_no'] or '미상'} · 이력 #{doc['id']}"
                       + (f" · 문서 검토 사유: {doc['reasons']}" if doc["reasons"] else ""))
            st.markdown(items_table(get_item_results(db, doc["id"])), unsafe_allow_html=True)


# ── 탭 4: 이력 검색 ────────────────────────────────────────────
def history_panel(db: Path) -> None:
    """협력사·Lot·기간·판정으로 검색하고 엑셀로 내려받는다(감사 대응)."""
    suppliers = [r["supplier"] for r in read_rows(db, "SELECT DISTINCT supplier FROM documents "
                                                      "WHERE supplier IS NOT NULL ORDER BY supplier")]
    c1, c2, c3, c4 = st.columns([2, 1.3, 2, 2])
    supplier = c1.selectbox("협력사", ["전체", *suppliers])
    lot = c2.text_input("Lot (일부만 입력 가능)")
    period = c3.date_input("처리 기간", (date.today() - timedelta(days=90), date.today()))
    status_labels = c4.multiselect("판정", ["부적합", "검토 필요", "적합"])
    code = {v: k for k, v in STATUS_LABEL.items()}

    start, end = (period[0], period[-1]) if isinstance(period, (tuple, list)) and period else (None, None)
    docs = search_documents(db, supplier=None if supplier == "전체" else supplier, lot=lot or None,
                            date_from=start.isoformat() if start else None, date_to=end.isoformat() if end else None,
                            statuses=[code[s] for s in status_labels] or None)
    st.write(f"**{len(docs)}건**")
    if not docs:
        return
    table = pd.DataFrame([{
        "ID": d["id"], "처리시각": when(d["processed_at"]), "경로": SOURCE_LABEL.get(d["source"], d["source"]),
        "협력사": d["supplier"], "재질": d["material_code"], "Lot": d["lot_no"], "제조일": d["mfg_date"],
        "판정": STATUS_LABEL[d["overall_status"]], "Spec": d["spec_version"], "발신자": d["sender"],
        "수신시각": when(d["received_at"]), "원본 PDF": d["file_path"], "문서 검토 사유": d["reasons"],
    } for d in docs])
    colors = {"부적합": "background-color:#FFC7CE;color:#9C0006;font-weight:700",
              "검토 필요": "background-color:#FFEB9C;color:#7F6000;font-weight:700",
              "적합": "background-color:#C6EFCE;color:#006100;font-weight:700"}
    st.dataframe(table.style.map(lambda v: colors.get(v, ""), subset=["판정"]), hide_index=True, width="stretch")

    items = pd.DataFrame([{
        "문서 ID": i["document_id"], "항목": i["item_std"], "원문 항목명": i["raw_name"], "원문 값": i["raw_value"],
        "원문 단위": i["raw_unit"], "환산 값": i["normalized_value"], "규격 단위": i["spec_unit"], "하한": i["lower"],
        "상한": i["upper"], "판정": STATUS_LABEL.get(i["status"], i["status"]), "사유": i["reason"],
    } for i in get_item_results(db, [d["id"] for d in docs])])
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        table.to_excel(writer, sheet_name="문서", index=False)
        items.to_excel(writer, sheet_name="항목별 판정", index=False)
    st.download_button("📥 검색 결과 엑셀 다운로드 (문서 + 항목별 판정 근거)", buffer.getvalue(),
                       file_name=f"COA_이력검색_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def main() -> None:
    """화면 구성: 제목 + 탭 4개 (처리 현황이 첫 화면)."""
    st.set_page_config(page_title="AI COA 자동검증기", page_icon="🧪", layout="wide")
    st.markdown(CSS, unsafe_allow_html=True)
    cfg = load_config()
    db, output_dir = cfg["paths"]["db"], cfg["paths"]["output"]

    st.title("🧪 AI COA 자동검증기")
    st.markdown('<div class="subtitle">메일 도착 즉시 판정 · 사람은 예외 건만 확인</div>', unsafe_allow_html=True)
    tab_status, tab_detail, tab_upload, tab_history = st.tabs(["📋 처리 현황", "🔍 판정 상세", "📤 수동 업로드", "🗂️ 이력 검색"])
    with tab_status:
        status_panel(db)
    with tab_detail:
        detail_panel(db, output_dir)
    with tab_upload:
        upload_panel(db)
    with tab_history:
        history_panel(db)

    st.divider()
    st.caption(f"DB 파일: {db}")  # 감시 프로그램 로그 첫 줄의 DB 경로와 같은지 확인용


main()
