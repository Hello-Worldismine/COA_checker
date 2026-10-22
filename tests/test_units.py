"""core.units 테스트: 값 파싱, 단위 정규화, 질량분율 환산."""

from decimal import Decimal

import pytest

from core.units import (
    GREATER_THAN, INVALID, LESS_THAN, NOT_DETECTED, NUMBER, RANGE,
    convert, normalize_unit, parse_value,
)

D = Decimal


@pytest.mark.parametrize("value, from_unit, to_unit, expected", [
    ("50", "ppm", "%", "0.005"),       # 기획서 사례 2: 수분 50 ppm
    ("0.008", "%", "ppm", "80"),       # 기획서 사례 1: Iron 0.008 %
    ("10", "ppm", "%", "0.001"),
    ("1", "ppb", "ppm", "0.001"),
    ("20", "mg/kg", "%", "0.002"),
    ("0.5", "wt%", "ppm", "5000"),
    ("3", "g/kg", "%", "0.3"),
    ("1000", "μg/kg", "ppm", "1"),
    ("5", "APHA", "APHA", "5"),
])
def test_convert(value, from_unit, to_unit, expected):
    """CLAUDE.md 환산표대로 Decimal 환산."""
    assert convert(D(value), from_unit, to_unit) == D(expected)


@pytest.mark.parametrize("from_unit, to_unit", [
    ("APHA", "%"),     # 무차원 ↔ 질량분율
    ("ppm", "APHA"),
    ("mg/L", "%"),     # 모르는 단위
    (None, "%"),
    ("%", ""),
])
def test_convert_impossible(from_unit, to_unit):
    """환산할 수 없으면 None."""
    assert convert(D("1"), from_unit, to_unit) is None


def test_convert_result_is_exact_decimal():
    """환산 결과는 float가 아닌 정확한 Decimal."""
    result = convert(D("50"), "ppm", "%")
    assert isinstance(result, Decimal)
    assert result > D("0.003")
    assert result == D("0.0050")


@pytest.mark.parametrize("raw, kind, value", [
    ("<10", LESS_THAN, "10"),
    ("≤ 10", LESS_THAN, "10"),
    ("<= 0.5", LESS_THAN, "0.5"),
    ("＜10", LESS_THAN, "10"),           # 전각 부등호
    ("N.D.", NOT_DETECTED, None),
    ("ND", NOT_DETECTED, None),
    ("n.d", NOT_DETECTED, None),
    ("Not Detected", NOT_DETECTED, None),
    ("<LOD", NOT_DETECTED, None),
    (">5", GREATER_THAN, "5"),
    ("≥ 99.9", GREATER_THAN, "99.9"),
    ("0,003", NUMBER, "0.003"),          # 쉼표 소수점
    (" 99.95 ", NUMBER, "99.95"),
    ("０．００３", NUMBER, "0.003"),     # 전각 숫자
    ("99.95 ", NUMBER, "99.95"),    # 줄바꿈 없는 공백
    ("1,000.5", NUMBER, "1000.5"),       # 천 단위 쉼표
    ("1.2E-3", NUMBER, "0.0012"),
    ("0.5~1.0", RANGE, None),
    ("0.5 - 1.0", RANGE, None),
    ("", INVALID, None),
    ("   ", INVALID, None),
    (None, INVALID, None),
    ("abc", INVALID, None),
    ("1,000", INVALID, None),            # 천 단위인지 소수점인지 애매 → 추측 안 함
    ("99.9 ± 0.2", INVALID, None),
    ("<", INVALID, None),
    ("10 ppm", INVALID, None),           # 단위는 값과 따로 넘겨야 함
])
def test_parse_value(raw, kind, value):
    """특수 표기·전각·쉼표·빈 값 파싱."""
    assert parse_value(raw) == (kind, None if value is None else D(value))


def test_decimal_equality_ignores_trailing_zero():
    """0.003 과 0.0030 은 Decimal 비교로 같다 (float 비교 금지)."""
    assert parse_value("0.003")[1] == parse_value("0.0030")[1]
    assert parse_value("0,0030")[1] == D("0.003")
    assert D(str(0.003)) == D("0.0030")
    assert D(0.003) != D("0.003")  # Decimal(float)는 이진 오차가 생겨 쓰면 안 되는 이유


@pytest.mark.parametrize("raw, expected", [
    ("%", "%"), ("wt%", "%"), ("w/w%", "%"), ("w/w %", "%"), ("％", "%"),
    ("ppm", "ppm"), (" MG/KG ", "ppm"), ("μg/g", "ppm"), ("µg/g", "ppm"), ("ug/g", "ppm"),
    ("ppb", "ppb"), ("μg/kg", "ppb"),
    ("g/kg", "g/kg"),
    ("APHA", "APHA"), ("apha", "APHA"),
    ("mg/L", None), ("", None), (None, None), ("Hazen", None),
])
def test_normalize_unit(raw, expected):
    """대소문자·공백·마이크로 기호 표기 차이를 무시하고 표준 단위로."""
    assert normalize_unit(raw) == expected
