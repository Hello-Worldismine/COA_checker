"""화면(app.py)·화면용 DB 조회 테스트: 임시 설정·임시 DB로 Streamlit AppTest 실행."""

import json
import shutil
import sqlite3
from datetime import date

import pytest
from streamlit.testing.v1 import AppTest

from core.config import CONFIG_ENV, PROJECT_ROOT, load_config
from core.pipeline import process_pdf
from core.reporter import connect, get_item_results, search_documents

SAMPLES_DIR = load_config()["paths"]["samples"]
EXPECTED = json.loads((SAMPLES_DIR / "expected.json").read_text(encoding="utf-8"))["samples"]


@pytest.fixture
def demo_env(tmp_path, monkeypatch, spec_master):
    """임시 설정 파일(경로는 tmp)과 샘플 5장을 처리해 둔 DB."""
    for name in ("inbox", "output", "logs"):
        (tmp_path / name).mkdir()
    shutil.copy(PROJECT_ROOT / "data" / "spec_master.xlsx", tmp_path / "spec_master.xlsx")
    text = (PROJECT_ROOT / "config.toml").read_text(encoding="utf-8")
    text = text.replace('db = "output/coa_history.db"', 'db = "output/test.db"')
    text = text.replace('spec_master = "data/spec_master.xlsx"', 'spec_master = "spec_master.xlsx"')
    text = text.replace(f'samples = "samples"', f'samples = "{SAMPLES_DIR.as_posix()}"')
    (tmp_path / "config.toml").write_text(text, encoding="utf-8")
    monkeypatch.setenv(CONFIG_ENV, str(tmp_path / "config.toml"))
    paths = load_config()["paths"]
    for n in sorted(EXPECTED):
        meta = {"source": "mail", "mail_entry_id": "E1", "attachment_name": n, "sender": "jaemin9905@naver.com"}
        process_pdf(SAMPLES_DIR / n, meta, spec=spec_master, paths=paths)
    return paths


def test_wal_mode_and_read_while_writing(demo_env):
    """WAL 모드: 감시 프로그램이 쓰기 트랜잭션 중이어도 화면 조회가 잠금 오류 없이 된다."""
    db = demo_env["db"]
    writer = connect(db)
    assert writer.execute("PRAGMA journal_mode").fetchone()[0] == "wal"
    writer.execute("BEGIN IMMEDIATE")
    writer.execute("INSERT INTO documents (processed_at, source, file_path, overall_status) "
                   "VALUES ('2026-10-23T12:00:00', 'upload', 'x.pdf', 'PASS')")
    try:
        docs = search_documents(db)       # 아직 커밋 안 된 행은 안 보이지만 잠금 오류도 없다
        assert len(docs) == 5
        assert len(get_item_results(db, [d["id"] for d in docs])) == 20
    finally:
        writer.rollback()
        writer.close()


def test_readonly_cannot_write(demo_env):
    """화면용 연결은 읽기 전용."""
    from core.reporter import connect_readonly

    conn = connect_readonly(demo_env["db"])
    with pytest.raises(sqlite3.OperationalError):
        conn.execute("DELETE FROM documents")
    conn.close()


def test_search_filters(demo_env):
    """협력사·Lot(부분)·기간·판정 검색."""
    db = demo_env["db"]
    today = date.today().isoformat()
    assert len(search_documents(db, supplier="Nippon Solvent Co., Ltd.")) == 3
    assert [d["lot_no"] for d in search_documents(db, lot="2418")] == ["N2418"]
    assert len(search_documents(db, statuses=["FAIL"])) == 2
    assert len(search_documents(db, date_from=today, date_to=today)) == 5
    assert search_documents(db, date_from="2000-01-01", date_to="2000-01-02") == []
    assert all(d["result_xlsx"] for d in search_documents(db))  # 결과 엑셀 경로도 남는다


def test_old_db_gets_new_columns(tmp_path):
    """STEP 7까지 만든 DB에도 result_xlsx·notice_path 열이 자동으로 붙는다."""
    from core.reporter import SCHEMA

    db = tmp_path / "old.db"
    old_schema = "\n".join(line for line in SCHEMA.splitlines() if "result_xlsx" not in line and "notice_path" not in line)
    with sqlite3.connect(db) as c:
        c.executescript(old_schema)
        assert "result_xlsx" not in {r[1] for r in c.execute("PRAGMA table_info(documents)")}
    conn = connect(db)
    cols = {r[1] for r in conn.execute("PRAGMA table_info(documents)")}
    conn.close()
    assert {"result_xlsx", "notice_path"} <= cols


def test_app_renders_all_tabs(demo_env):
    """화면이 오류 없이 뜨고, 처리 현황(첫 탭)에 오늘 지표와 목록, 판정 상세에 판정표가 나온다."""
    at = AppTest.from_file(str(PROJECT_ROOT / "app.py"), default_timeout=60).run()
    assert not at.exception
    assert at.title[0].value == "🧪 AI COA 자동검증기"
    assert [t.label for t in at.tabs] == ["📋 처리 현황", "🔍 판정 상세", "📤 수동 업로드", "🗂️ 이력 검색"]
    metrics = {m.label: m.value for m in at.metric}
    assert metrics == {"오늘 처리": "5건", "🟥 부적합": "2건", "🟨 검토 필요": "1건", "🟩 적합": "2건"}
    html_text = "".join(m.value for m in at.markdown)
    assert "N2418" in html_text and "badge fail" in html_text
    assert "50 ppm" in html_text and "상한 0.003 % 초과" in html_text or "검출한계 미기재" in html_text
    assert any(c.value == f"DB 파일: {demo_env['db']}" for c in at.caption)  # 화면 하단 DB 경로


def test_app_without_db(tmp_path, monkeypatch):
    """처리 이력이 없어도 화면이 뜬다(안내 문구)."""
    (tmp_path / "config.toml").write_text(
        (PROJECT_ROOT / "config.toml").read_text(encoding="utf-8")
        .replace('db = "output/coa_history.db"', 'db = "none/empty.db"'), encoding="utf-8")
    monkeypatch.setenv(CONFIG_ENV, str(tmp_path / "config.toml"))
    at = AppTest.from_file(str(PROJECT_ROOT / "app.py"), default_timeout=60).run()
    assert not at.exception
    assert any("아직 처리 이력이 없습니다" in i.value for i in at.info)
    assert not (tmp_path / "none").exists()  # 화면은 DB를 만들지 않는다
