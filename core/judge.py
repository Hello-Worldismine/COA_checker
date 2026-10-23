"""규격 판정: Decimal 결정적 규칙으로 PASS / FAIL / REVIEW 판정.

CLAUDE.md "판정 규칙 요약"을 그대로 구현한다. 모르는 표기·단위·누락은 추측하지 않고 REVIEW.
"""

from decimal import Decimal

from core.mapper import ItemMapper
from core.models import CoaDocument, CoaResult, ExtractedItem, ItemResult, Status
from core.spec import SpecItem, SpecMaster
from core.units import (
    GREATER_THAN, INVALID, LESS_THAN, NOT_DETECTED, RANGE,
    convert, normalize_unit, parse_value,
)


def combine_status(statuses) -> Status:
    """문서 판정: FAIL이 하나라도 있으면 FAIL, 아니면 REVIEW가 있으면 REVIEW, 모두 PASS면 PASS."""
    statuses = set(statuses)
    if Status.FAIL in statuses:
        return Status.FAIL
    if Status.REVIEW in statuses:
        return Status.REVIEW
    return Status.PASS


def fmt(value: Decimal) -> str:
    """Decimal을 지수 표기 없이, 뒤쪽 0 없이 표시 (80.000 → '80', 0.0050 → '0.005')."""
    text = format(value.normalize(), "f")
    return "0" if text in ("-0", "") else text


def _qty(value: Decimal, unit: str) -> str:
    """'0.005 %' 같은 값+단위 문자열."""
    return f"{fmt(value)} {unit}"


def limit_text(lower: Decimal | None, upper: Decimal | None, unit: str) -> str:
    """규격 범위를 사람이 읽는 문자열로 ('≤ 0.003 %', '≥ 99.9 %', '0.5~1 %')."""
    if lower is not None and upper is not None:
        return f"{fmt(lower)}~{fmt(upper)} {unit}"
    if upper is not None:
        return f"≤ {_qty(upper, unit)}"
    if lower is not None:
        return f"≥ {_qty(lower, unit)}"
    return "규격 한계 없음"


def _compare(value: Decimal, spec: SpecItem) -> tuple[Status, str]:
    """환산된 숫자를 하한·상한과 비교한다(경계값 포함 → PASS)."""
    if spec.lower is not None and value < spec.lower:
        return Status.FAIL, f"하한 {_qty(spec.lower, spec.unit)} 미달"
    if spec.upper is not None and value > spec.upper:
        return Status.FAIL, f"상한 {_qty(spec.upper, spec.unit)} 초과"
    if spec.lower is not None and spec.upper is not None:
        return Status.PASS, f"규격 {limit_text(spec.lower, spec.upper, spec.unit)} 이내"
    if spec.upper is not None:
        return Status.PASS, f"상한 {_qty(spec.upper, spec.unit)} 이하"
    return Status.PASS, f"하한 {_qty(spec.lower, spec.unit)} 이상"


def judge_item(spec: SpecItem, item: ExtractedItem | None) -> ItemResult:
    """Spec 항목 1개를 COA 항목(없으면 None)과 대조해 ItemResult를 만든다."""
    result = ItemResult(item_std=spec.item_std, spec_unit=spec.unit, lower=spec.lower, upper=spec.upper,
                        status=Status.REVIEW)
    if item is None:
        result.reason = "항목 누락 - COA에 기재 없음"
        return result
    result.raw_name, result.raw_value, result.raw_unit = item.raw_name, item.raw_value, item.raw_unit

    def review(reason: str) -> ItemResult:
        result.status, result.reason = Status.REVIEW, reason
        return result

    if spec.lower is None and spec.upper is None:
        return review("규격 하한·상한 미설정 - 규격 마스터 확인 필요")

    kind, value = parse_value(item.raw_value)
    raw = item.raw_value.strip()
    if kind == NOT_DETECTED:
        return review(f"{raw} 표기 - 검출한계 미기재로 검토 필요")
    if kind == GREATER_THAN:
        return review(f"'{raw}' 표기 - 실제 값을 알 수 없어 검토 필요")
    if kind == RANGE:
        return review(f"범위 표기({raw}) - 단일 값이 아니어서 검토 필요")
    if kind == INVALID:
        return review("값 없음 - 확인 필요" if not raw else f"값 형식 해석 불가('{raw}') - 확인 필요")

    if not item.raw_unit or not item.raw_unit.strip():
        return review("단위 미기재 - 확인 필요")
    unit = normalize_unit(item.raw_unit)
    if unit is None:
        return review(f"모르는 단위('{item.raw_unit}') - 확인 필요")
    converted = convert(value, unit, spec.unit)
    if converted is None:
        return review(f"단위 환산 불가({unit} → {spec.unit}) - 확인 필요")
    result.normalized_value = converted

    shown = _qty(value, unit) if unit == spec.unit else f"{_qty(value, unit)} → {_qty(converted, spec.unit)} 환산"

    if kind == LESS_THAN:
        if spec.lower is not None:
            return review(f"'<{fmt(value)}' 표기 - 하한이 있는 항목이라 실제 값 확인 필요")
        if converted <= spec.upper:
            result.status = Status.PASS
            result.reason = f"<{shown}, 상한 {_qty(spec.upper, spec.unit)} 이하"
            return result
        return review(f"<{shown}, 표기 한계가 상한 {_qty(spec.upper, spec.unit)}보다 커서 실제 값 확인 필요")

    result.status, verdict = _compare(converted, spec)
    result.reason = f"{shown}, {verdict}"
    return result


def judge(doc: CoaDocument, spec: SpecMaster) -> CoaResult:
    """CoaDocument를 규격 마스터와 대조해 CoaResult를 만든다. 판정은 이 함수의 결정적 규칙으로만 한다."""
    reasons = list(doc.extract_warnings)
    if doc.material_code is None:
        if "재질을 정하지 못함" not in " ".join(reasons):
            reasons.append("재질을 정하지 못함")
        return CoaResult(document=doc, overall_status=Status.REVIEW, reasons=reasons, extra_items=doc.items)

    specs = spec.specs_for(doc.material_code)
    if not specs:
        reasons.append(f"규격 마스터에 재질 '{doc.material_code}'의 규격이 없음")
        return CoaResult(document=doc, overall_status=Status.REVIEW, reasons=reasons, extra_items=doc.items)
    if not doc.items and "시험 항목을 찾지 못함" not in reasons:
        reasons.append("시험 항목을 찾지 못함")

    mapper = ItemMapper.from_spec(spec, doc.material_code)
    matched: dict[str, list[ExtractedItem]] = {}
    extra: list[ExtractedItem] = []
    for item in doc.items:
        item_std = mapper.map(item.raw_name)
        if item_std is None:
            extra.append(item)
        else:
            matched.setdefault(item_std, []).append(item)

    item_results = []
    for s in specs:
        found = matched.get(s.item_std, [])
        if len(found) > 1:
            dup = ItemResult(item_std=s.item_std, spec_unit=s.unit, lower=s.lower, upper=s.upper,
                             status=Status.REVIEW, raw_name=found[0].raw_name, raw_value=found[0].raw_value,
                             raw_unit=found[0].raw_unit,
                             reason=f"같은 항목이 {len(found)}번 기재됨({', '.join(i.raw_name for i in found)}) - 확인 필요")
            item_results.append(dup)
        else:
            item_results.append(judge_item(s, found[0] if found else None))

    statuses = [r.status for r in item_results] + ([Status.REVIEW] if reasons else [])
    return CoaResult(
        document=doc,
        item_results=item_results,
        overall_status=combine_status(statuses),
        reasons=reasons,
        extra_items=extra,
        spec_version=spec.spec_version(doc.material_code),
    )
