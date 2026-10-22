"""core.mapper 테스트: 원문 항목명 → 사내 표준 항목."""

import pytest

from core.mapper import ItemMapper, normalize_name


@pytest.fixture(scope="module")
def mapper(spec_master):
    """NMP-EL 규격 별칭으로 만든 매퍼."""
    return ItemMapper.from_spec(spec_master, "NMP-EL")


@pytest.mark.parametrize("raw, expected", [
    ("Iron (Fe)", "Iron"),
    ("FE", "Iron"),
    ("Iron(Fe)", "Iron"),
    ("철", "Iron"),
    ("Water content", "Moisture"),
    ("WATER  CONTENT", "Moisture"),
    ("KF", "Moisture"),
    ("Moisture (KF)", "Moisture"),     # 괄호 밖·안 모두 Moisture
    ("Water", "Moisture"),
    ("Assay", "Purity"),
    ("Purity (GC)", "Purity"),
    ("Purity, %", "Purity"),
    ("Colour", "Color"),
    ("APHA", "Color"),
    ("Viscosity", None),               # 모르는 항목
    ("", None),
    (None, None),
])
def test_map(mapper, raw, expected):
    """대소문자·괄호·공백·특수문자 차이를 무시하고 매칭, 못 찾으면 None."""
    assert mapper.map(raw) == expected


def test_ambiguous_parentheses_returns_none(mapper):
    """괄호 안·밖이 서로 다른 항목을 가리키면 추측하지 않는다."""
    assert mapper.map("Water (Fe)") is None


def test_normalize_name():
    """비교용 이름 정리."""
    assert normalize_name(" Iron (Fe) ") == "ironfe"
    assert normalize_name("Ｈ２Ｏ") == "h2o"


def test_duplicate_alias_rejected():
    """같은 별칭이 서로 다른 항목에 등록되면 오류."""
    with pytest.raises(ValueError):
        ItemMapper({"Fe": "Iron", "fe ": "Moisture"})
