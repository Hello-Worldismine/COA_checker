"""scripts/reset_demo.py 테스트: 임시 폴더에서 초기화 후 처음부터 다시 처리할 수 있는지."""

from core.pipeline import process_pdf
from core.reporter import search_documents
from scripts import make_samples, reset_demo


def test_reset_then_reprocess(tmp_path, spec_master):
    """처리 이력·파일을 만든 뒤 초기화하면 DB·inbox·output이 비고 샘플이 다시 생기며, 다시 처리할 수 있다."""
    paths = {"db": tmp_path / "output" / "coa.db", "inbox": tmp_path / "inbox", "output": tmp_path / "output",
             "samples": tmp_path / "samples", "spec_master": tmp_path / "data" / "spec_master.xlsx"}
    for p in (paths["inbox"], paths["output"]):
        p.mkdir(parents=True)
        (p / ".gitkeep").write_text("")
    make_samples.build_all(paths["samples"], paths["spec_master"])
    pdf = paths["samples"] / "COA_Nippon_N2418_fail.pdf"
    process_pdf(pdf, {"source": "upload"}, spec=spec_master, paths=paths)
    (paths["inbox"] / "_smoke_test").mkdir()
    assert len(search_documents(paths["db"])) == 1
    pdf.unlink()  # 샘플이 망가져도 다시 만든다

    summary = reset_demo.reset(paths)

    assert summary["db"] is True and summary["samples"] == 5
    assert not paths["db"].exists()
    assert sorted(p.name for p in paths["inbox"].iterdir()) == [".gitkeep"]
    assert sorted(p.name for p in paths["output"].iterdir()) == [".gitkeep"]
    assert pdf.is_file() and (paths["samples"] / "expected.json").is_file() and paths["spec_master"].is_file()

    out = process_pdf(pdf, {"source": "upload"}, spec=spec_master, paths=paths)
    assert out.saved.document_id == 1 and out.result.overall_status == "FAIL"


def test_reset_when_nothing_exists(tmp_path):
    """처음 상태(아무것도 없음)에서도 오류 없이 동작한다."""
    paths = {"db": tmp_path / "output" / "coa.db", "inbox": tmp_path / "inbox", "output": tmp_path / "output",
             "samples": tmp_path / "samples", "spec_master": tmp_path / "data" / "spec_master.xlsx"}
    summary = reset_demo.reset(paths)
    assert summary == {"db": False, "inbox": 0, "output": 0, "samples": 5}
