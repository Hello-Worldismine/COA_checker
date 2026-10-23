"""core.judge·core.pipeline 테스트: 샘플 5장 정답지 대조와 판정 규칙 단위 테스트."""

import json
from decimal import Decimal
from pathlib import Path

import pytest

from core.config import load_config
from core.judge import combine_status, fmt, judge, judge_item
from core.models import CoaDocument, ExtractedItem, Status
from core.pipeline import process_pdf
from core.spec import SpecItem, SpecMaster

SAMPLES_DIR = load_config()["paths"]["samples"]
EXPECTED = json.loads((SAMPLES_DIR / "expected.json").read_text(encoding="utf-8"))["samples"]
D = Decimal

MOISTURE = SpecItem("NMP-EL", "Moisture", (), None, D("0.003"), "%", "v1.0")
PURITY = SpecItem("NMP-EL", "Purity", (), D("99.90"), None, "%", "v1.0")
COLOR = SpecItem("NMP-EL", "Color", (), None, D("10"), "APHA", "v1.0")
BAND = SpecItem("X", "Viscosity", (), D("0.5"), D("1.0"), "%", "v1.0")  # 하한·상한 둘 다


def item(value, unit="%", name="X"):
    """테스트용 COA 항목."""
    return ExtractedItem(raw_name=name, raw_value=value, raw_unit=unit)


# ── 샘플 5장 ──────────────────────────────────────────────────
@pytest.mark.parametrize("file_name", sorted(EXPECTED))
def test_samples_match_expected(spec_master, file_name):
    """문서 판정·항목별 판정이 expected.json과 일치한다."""
    exp = EXPECTED[file_name]
    result = process_pdf(SAMPLES_DIR / file_name, spec=spec_master, save=False).result
    assert result.overall_status == exp["overall"]
    assert {r.item_std: r.status.value for r in result.item_results} == {
        k: v["status"] for k, v in exp["items"].items()
    }
    assert result.spec_version == "v1.0"
    assert result.reasons == [] and result.extra_items == []


def test_sample_reasons(spec_master):
    """사유 문구: ppm 환산 부적합, N.D. 검토."""
    fail = process_pdf(SAMPLES_DIR / "COA_Nippon_N2418_fail.pdf", spec=spec_master, save=False).result
    moisture = next(r for r in fail.item_results if r.item_std == "Moisture")
    assert moisture.reason == "50 ppm → 0.005 % 환산, 상한 0.003 % 초과"
    assert moisture.normalized_value == D("0.005")

    review = process_pdf(SAMPLES_DIR / "COA_Nippon_N2419_review.pdf", spec=spec_master, save=False).result
    iron = next(r for r in review.item_results if r.item_std == "Iron")
    assert iron.reason == "N.D. 표기 - 검출한계 미기재로 검토 필요"


def test_process_pdf_keeps_meta(spec_master):
    """메일 정보(meta)를 결과에 그대로 담는다."""
    meta = {"sender": "jaemin9905@naver.com", "entry_id": "ABC"}
    result = process_pdf(SAMPLES_DIR / "COA_Alpha_A2401_pass.pdf", meta=meta, spec=spec_master, save=False).result
    assert result.meta == meta and result.overall_status is Status.PASS


# ── 경계값·상하한 ──────────────────────────────────────────────
@pytest.mark.parametrize("spec, value, unit, status", [
    (MOISTURE, "0.003", "%", Status.PASS),      # 상한과 정확히 같음 → PASS
    (MOISTURE, "0.0030", "%", Status.PASS),
    (MOISTURE, "30", "ppm", Status.PASS),       # 환산 후 정확히 0.003 %
    (MOISTURE, "0.0031", "%", Status.FAIL),
    (MOISTURE, "31", "ppm", Status.FAIL),
    (PURITY, "99.90", "%", Status.PASS),        # 하한과 같음 → PASS
    (PURITY, "99.9", "wt%", Status.PASS),
    (PURITY, "99.89", "%", Status.FAIL),
    (BAND, "0.5", "%", Status.PASS),
    (BAND, "1.0", "%", Status.PASS),
    (BAND, "0.75", "%", Status.PASS),
    (BAND, "0.49", "%", Status.FAIL),
    (BAND, "1.01", "%", Status.FAIL),
    (BAND, "8000", "ppm", Status.PASS),         # 0.8 %
    (COLOR, "10", "APHA", Status.PASS),
    (COLOR, "11", "apha", Status.FAIL),
])
def test_limits_inclusive(spec, value, unit, status):
    """경계값 포함, 하한·상한 둘 다 있는 항목, 단위 환산 후 비교."""
    assert judge_item(spec, item(value, unit)).status is status


def test_reason_texts():
    """하한 미달·범위 이내 사유 문구."""
    assert judge_item(PURITY, item("99.80")).reason == "99.8 %, 하한 99.9 % 미달"
    assert judge_item(BAND, item("0.75")).reason == "0.75 %, 규격 0.5~1 % 이내"
    assert judge_item(MOISTURE, item("0.008", "%")).normalized_value == D("0.008")


# ── 특수 표기 ──────────────────────────────────────────────────
@pytest.mark.parametrize("spec, value, unit, status, reason_part", [
    (MOISTURE, "<10", "ppm", Status.PASS, "상한 0.003 % 이하"),
    (MOISTURE, "≤ 0.003", "%", Status.PASS, "상한 0.003 % 이하"),
    (MOISTURE, "<50", "ppm", Status.REVIEW, "실제 값 확인 필요"),        # 표기 한계가 상한보다 큼
    (BAND, "<0.8", "%", Status.REVIEW, "하한 0.5 % 이상인지 알 수 없어"),   # x > 하한
    (BAND, "<0.4", "%", Status.FAIL, "하한 0.5 % 미달 확정"),               # x < 하한
    (BAND, "<0.5", "%", Status.FAIL, "미달 확정"),                          # '<x', x = 하한 → 미만 확정
    (BAND, "≤0.5", "%", Status.REVIEW, "확인 필요"),                       # '≤x', x = 하한 → 하한과 같을 수 있음
    (BAND, "<= 0.4", "%", Status.FAIL, "≤0.4 %"),
    (PURITY, "<99.5", "%", Status.FAIL, "하한 99.9 % 미달 확정"),
    (PURITY, "<99.95", "%", Status.REVIEW, "확인 필요"),
    (MOISTURE, "N.D.", None, Status.REVIEW, "검출한계 미기재"),
    (MOISTURE, "Not Detected", "ppm", Status.REVIEW, "검출한계 미기재"),
    (MOISTURE, ">5", "ppm", Status.REVIEW, "실제 값을 알 수 없어"),
    (MOISTURE, "0.5~1.0", "ppm", Status.REVIEW, "범위 표기"),
    (MOISTURE, "", "%", Status.REVIEW, "값 없음"),
    (MOISTURE, "abc", "%", Status.REVIEW, "해석 불가"),
])
def test_special_notations(spec, value, unit, status, reason_part):
    """<x, ≤x, N.D., >x, 범위, 빈 값, 이상한 값."""
    r = judge_item(spec, item(value, unit))
    assert r.status is status and reason_part in r.reason


# ── 단위 ───────────────────────────────────────────────────────
@pytest.mark.parametrize("spec, unit, reason", [
    (MOISTURE, None, "단위 미기재 - 확인 필요"),
    (MOISTURE, "  ", "단위 미기재 - 확인 필요"),
    (MOISTURE, "mg/L", "모르는 단위('mg/L') - 확인 필요"),
    (MOISTURE, "APHA", "단위 환산 불가(APHA → %) - 확인 필요"),
    (COLOR, "%", "단위 환산 불가(% → APHA) - 확인 필요"),
])
def test_unit_problems_review(spec, unit, reason):
    """단위 없음·모르는 단위·환산 불가 → REVIEW (규격 단위로 간주하지 않음)."""
    r = judge_item(spec, item("5", unit))
    assert r.status is Status.REVIEW and r.reason == reason and r.normalized_value is None


# ── 문서 단위 ──────────────────────────────────────────────────
def _doc(items, material_code="NMP-EL", warnings=()):
    """테스트용 CoaDocument."""
    return CoaDocument(source_path=Path("x.pdf"), material_code=material_code, items=items,
                       extract_warnings=list(warnings))


GOOD = [item("99.95", "%", "Purity"), item("0.002", "%", "Moisture"),
        item("0.003", "%", "Iron"), item("5", "APHA", "Color")]


def test_all_pass(spec_master):
    """규격 4항목 모두 적합 → PASS."""
    assert judge(_doc(GOOD), spec_master).overall_status is Status.PASS


def test_missing_and_extra_items(spec_master):
    """Spec 항목 누락 → REVIEW, COA에만 있는 항목은 기록만."""
    result = judge(_doc(GOOD[:3] + [item("0.1", "%", "Viscosity")]), spec_master)
    color = next(r for r in result.item_results if r.item_std == "Color")
    assert color.status is Status.REVIEW and color.reason.startswith("항목 누락")
    assert [i.raw_name for i in result.extra_items] == ["Viscosity"]
    assert result.overall_status is Status.REVIEW


def test_fail_beats_review(spec_master):
    """FAIL과 REVIEW가 섞이면 문서는 FAIL."""
    items = [item("99.95", "%", "Purity"), item("50", "ppm", "Water"), item("N.D.", None, "Fe")]
    result = judge(_doc(items), spec_master)
    assert result.overall_status is Status.FAIL


def test_warning_does_not_downgrade_fail(spec_master):
    """추출 경고는 PASS를 REVIEW로만 올린다. FAIL 항목이 있으면 경고가 있어도 FAIL."""
    fail_items = GOOD[:2] + [item("0.008", "%", "Iron"), GOOD[3]]
    result = judge(_doc(fail_items, warnings=["Lot No.를 찾지 못함"]), spec_master)
    assert result.overall_status is Status.FAIL
    assert result.reasons == ["Lot No.를 찾지 못함"]
    assert judge(_doc(GOOD, warnings=["Lot No.를 찾지 못함"]), spec_master).overall_status is Status.REVIEW


def test_duplicate_item_review(spec_master):
    """같은 항목이 두 번 기재되면 추측하지 않고 REVIEW."""
    result = judge(_doc(GOOD + [item("20", "ppm", "Water")]), spec_master)
    moisture = next(r for r in result.item_results if r.item_std == "Moisture")
    assert moisture.status is Status.REVIEW and "2번 기재" in moisture.reason


@pytest.mark.parametrize("doc, reason", [
    (_doc(GOOD, material_code=None), "재질을 정하지 못함"),
    (_doc([]), "시험 항목을 찾지 못함"),
    (_doc(GOOD, material_code="UNKNOWN"), "규격 마스터에 재질 'UNKNOWN'의 규격이 없음"),
    (_doc(GOOD, warnings=["Lot No.를 찾지 못함"]), "Lot No.를 찾지 못함"),
])
def test_document_level_review(spec_master, doc, reason):
    """재질 미확인·항목 없음·규격 없음·추출 경고 → 문서 REVIEW, 사유 기록."""
    result = judge(doc, spec_master)
    assert result.overall_status is Status.REVIEW
    assert reason in result.reasons


def test_scanned_pdf_document_review(spec_master):
    """스캔 PDF(항목·재질 없음)는 문서 REVIEW."""
    doc = CoaDocument(source_path=Path("scan.pdf"), extract_warnings=["텍스트 없음 - 스캔 PDF 의심(OCR은 확장 단계)"])
    result = judge(doc, spec_master)
    assert result.overall_status is Status.REVIEW and result.item_results == []


def test_no_limits_spec_review():
    """규격 마스터에 하한·상한이 모두 없으면 REVIEW."""
    spec = SpecItem("X", "Odor", (), None, None, "%", "v1")
    assert judge_item(spec, item("1")).status is Status.REVIEW


def test_combine_status_and_fmt():
    """문서 판정 합치기와 숫자 표시."""
    assert combine_status([Status.PASS, Status.REVIEW]) is Status.REVIEW
    assert combine_status([Status.REVIEW, Status.FAIL, Status.PASS]) is Status.FAIL
    assert combine_status([Status.PASS]) is Status.PASS
    assert fmt(D("80.000")) == "80" and fmt(D("0.0050")) == "0.005" and fmt(D("1E+2")) == "100"


def test_custom_master_band(tmp_path):
    """하한·상한이 둘 다 있는 규격을 가진 마스터로 문서 판정."""
    master = SpecMaster(materials=(), specs=(BAND,), suppliers=())
    doc = CoaDocument(source_path=Path("x.pdf"), material_code="X", items=[item("1.2", "%", "Viscosity")])
    result = judge(doc, master)
    assert result.overall_status is Status.FAIL
    assert result.item_results[0].reason == "1.2 %, 상한 1 % 초과"
