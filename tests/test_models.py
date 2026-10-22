"""core.models 테스트: 생성·직렬화."""

from decimal import Decimal
from pathlib import Path

from core.models import CoaDocument, CoaResult, ExtractedItem, ItemResult, Status


def test_round_trip():
    """CoaResult를 JSON으로 저장했다 다시 읽어도 Decimal·상태가 유지된다."""
    doc = CoaDocument(
        source_path=Path("samples/COA_Alpha_A2402_fail.pdf"),
        lot_no="A2402",
        items=[ExtractedItem(raw_name="Iron (Fe)", raw_value="0.008", raw_unit="%")],
    )
    item = ItemResult(
        item_std="Iron", raw_name="Iron (Fe)", raw_value="0.008", raw_unit="%",
        normalized_value=Decimal("0.008"), spec_unit="%", upper=Decimal("0.005"),
        status=Status.FAIL, reason="0.008 % > 상한 0.005 %",
    )
    result = CoaResult(document=doc, item_results=[item], overall_status="FAIL", spec_version="v1.0")

    again = CoaResult.model_validate_json(result.model_dump_json())
    assert again.overall_status is Status.FAIL
    assert again.item_results[0].normalized_value == Decimal("0.008")
    assert again.document.items[0].raw_unit == "%"
    assert again.document.extract_warnings == []
