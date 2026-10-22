"""core.extractor 테스트: 샘플 5장을 정답지(expected.json)·샘플 생성 스크립트와 대조."""

import json

import pytest
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

from core.config import load_config
from core.extractor import (
    SCAN_WARNING, extract, items_from_lines, items_from_table, llm_extract, split_value_unit,
)
from core.mapper import ItemMapper
from scripts import make_samples

SAMPLES_DIR = load_config()["paths"]["samples"]
EXPECTED = json.loads((SAMPLES_DIR / "expected.json").read_text(encoding="utf-8"))["samples"]


@pytest.mark.parametrize("file_name", sorted(EXPECTED))
def test_samples_match_expected(spec_master, file_name):
    """항목 수·항목명·값·단위·Lot·협력사·재질 코드가 정답지와 같다."""
    exp = EXPECTED[file_name]
    doc = extract(SAMPLES_DIR / file_name, spec_master, llm_fallback=False)

    got = [(i.raw_name, i.raw_value, i.raw_unit) for i in doc.items]
    want = [(i["raw_name"], i["raw_value"], i["raw_unit"] or None) for i in exp["items"].values()]
    assert got == want
    assert doc.lot_no == exp["lot"]
    assert doc.supplier == exp["supplier"]
    assert doc.material_code == exp["material_code"]
    assert doc.extract_warnings == []


@pytest.mark.parametrize("sample", make_samples.SAMPLES, ids=lambda s: s.file)
def test_samples_match_generator(spec_master, sample):
    """샘플 생성 스크립트의 정의(제조일 포함)와도 일치한다."""
    doc = extract(SAMPLES_DIR / sample.file, spec_master, llm_fallback=False)
    assert [i.raw_name for i in doc.items] == [i.name for i in sample.items]
    assert doc.mfg_date == sample.mfg_date


def test_specification_column_not_used_as_result(spec_master):
    """Alpha 표의 Specification(Max. 0.005)이 아닌 Result(0.008)를 읽는다."""
    doc = extract(SAMPLES_DIR / "COA_Alpha_A2402_fail.pdf", spec_master, llm_fallback=False)
    iron = next(i for i in doc.items if i.raw_name == "Iron (Fe)")
    assert iron.raw_value == "0.008"
    assert all("Max" not in i.raw_value and "Min" not in i.raw_value for i in doc.items)


@pytest.mark.parametrize("text, expected", [
    ("20ppm", ("20", "ppm")),
    ("20 ppm", ("20", "ppm")),
    ("<10 ppm", ("<10", "ppm")),
    ("≤ 0.5 %", ("≤ 0.5", "%")),
    ("99.93 %", ("99.93", "%")),
    ("8 APHA", ("8", "APHA")),
    ("0.5~1.0 %", ("0.5~1.0", "%")),
    ("0,003 %", ("0,003", "%")),
    ("N.D.", ("N.D.", None)),
    ("ND ppm", ("ND", "ppm")),
    ("99.95", ("99.95", None)),
    ("Conforms", ("Conforms", None)),
])
def test_split_value_unit(text, expected):
    """값에 붙은 단위를 분리한다."""
    assert split_value_unit(text) == expected


def test_table_without_unit_column_and_spec_first():
    """단위 열이 없고 Specification 열이 Result보다 앞에 있어도 Result 열에서 값·단위를 분리한다."""
    table = [
        ["Certificate", None, None],
        ["Parameter", "Specification", "Result"],
        ["Water", "Max. 30 ppm", "20ppm"],
        ["Fe", "Max. 50 ppm", "N.D."],
        [None, None, None],
    ]
    items = items_from_table(table)
    assert [(i.raw_name, i.raw_value, i.raw_unit) for i in items] == [
        ("Water", "20", "ppm"), ("Fe", "N.D.", None),
    ]


def test_table_without_header_is_ignored():
    """항목·결과 헤더가 없는 표(서명란 등)는 항목으로 보지 않는다."""
    assert items_from_table([["Approved by", "QC Manager"], ["Date", "2026-10-01"]]) == []


def test_lines_skip_header_fields(spec_master):
    """콜론 줄은 규격 별칭에 있는 항목명만 인정하고, Lot·날짜 같은 머리글은 항목으로 보지 않는다."""
    text = "Batch No. : N1\nDate of Mfg. : 2026-10-01\nWater: 20 ppm\nFe   N.D.\nPurity (GC) ...... 99.95 %\n"
    items = items_from_lines(text, ItemMapper.from_spec(spec_master))
    assert [(i.raw_name, i.raw_value, i.raw_unit) for i in items] == [
        ("Water", "20", "ppm"), ("Fe", "N.D.", None), ("Purity (GC)", "99.95", "%"),
    ]


def _pdf(path, lines):
    """주어진 줄만 있는 PDF를 만든다 (lines가 비면 글자 없는 그림만)."""
    c = canvas.Canvas(str(path), pagesize=A4)
    if lines:
        y = 800
        for line in lines:
            c.drawString(50, y, line)
            y -= 20
    else:
        c.rect(50, 500, 300, 200, fill=1)
    c.save()
    return path


def test_scanned_pdf_warns(spec_master, tmp_path):
    """글자가 없는 PDF는 항목을 비우고 스캔 PDF 경고를 남긴다."""
    doc = extract(_pdf(tmp_path / "scan.pdf", []), spec_master, llm_fallback=False)
    assert doc.items == []
    assert doc.extract_warnings == [SCAN_WARNING]


def test_unknown_layout_warns(spec_master, tmp_path):
    """항목을 못 찾으면 경고를 남기고, llm_fallback=True여도 미구현 경고만 추가된다."""
    path = _pdf(tmp_path / "odd.pdf", ["Some Unknown Company", "This document has no test results at all."])
    doc = extract(path, spec_master, llm_fallback=True)
    assert doc.items == []
    assert any(w.startswith("LLM 보조 추출 불가") for w in doc.extract_warnings)
    assert "시험 항목을 찾지 못함" in doc.extract_warnings
    assert "협력사를 찾지 못함" in doc.extract_warnings


def test_broken_pdf_warns(spec_master, tmp_path):
    """열 수 없는 파일은 예외 대신 경고."""
    path = tmp_path / "broken.pdf"
    path.write_bytes(b"not a pdf")
    doc = extract(path, spec_master, llm_fallback=False)
    assert doc.items == [] and doc.extract_warnings[0].startswith("PDF를 열 수 없음")


def test_llm_extract_not_implemented(tmp_path):
    """LLM 보조 추출은 자리만 있다."""
    with pytest.raises(NotImplementedError):
        llm_extract(tmp_path / "x.pdf", "")


def test_supplier_by_email_domain(spec_master, tmp_path):
    """회사명이 없어도 본문의 메일 도메인으로 협력사를 찾는다."""
    path = _pdf(tmp_path / "d.pdf", ["Contact: qa@nippon-solv.example", "Product: NMP", "Lot No.: X1",
                                     "Water ........ 20 ppm"])
    doc = extract(path, spec_master, llm_fallback=False)
    assert doc.supplier == "Nippon Solvent Co., Ltd."
    assert doc.material_code == "NMP-EL" and doc.lot_no == "X1"
