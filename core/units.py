"""단위 정규화: 자체 환산표와 Decimal로 질량분율 단위를 ppm 기준으로 환산.

값 파싱(parse_value)은 모르는 표기를 추측하지 않고 "invalid"로 돌려 판정 단계에서 REVIEW가 되게 한다.
"""

import re
import unicodedata
from decimal import Decimal, InvalidOperation

# parse_value가 돌려주는 값 종류
NUMBER = "number"
LESS_THAN = "less_than"
NOT_DETECTED = "not_detected"
GREATER_THAN = "greater_than"
RANGE = "range"
INVALID = "invalid"

# 표준 단위 → 원문 표기(소문자, 공백 제거, NFKC 정규화 후 기준)
UNIT_ALIASES: dict[str, set[str]] = {
    "%": {"%", "wt%", "wt.%", "w/w%", "%w/w", "%(w/w)", "mass%", "weight%"},
    "ppm": {"ppm", "ppmw", "mg/kg", "μg/g", "ug/g"},
    "ppb": {"ppb", "ppbw", "μg/kg", "ug/kg", "ng/g"},
    "g/kg": {"g/kg"},
    "APHA": {"apha"},
}
_UNIT_LOOKUP = {alias: std for std, aliases in UNIT_ALIASES.items() for alias in aliases}

# 질량분율 환산표: 1 단위 = 몇 ppm
PPM_FACTOR: dict[str, Decimal] = {
    "%": Decimal("10000"),
    "g/kg": Decimal("1000"),
    "ppm": Decimal("1"),
    "ppb": Decimal("0.001"),
}

_NUM = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?"
_UNSIGNED = r"(?:\d+(?:\.\d*)?|\.\d+)"
_ND_PATTERN = re.compile(r"n\.?d\.?|not\s*detected|불검출|<\s*(?:lod|loq|dl|mdl)", re.IGNORECASE)
_RANGE_PATTERN = re.compile(rf"({_UNSIGNED})\s*(?:~|-|–|—|to)\s*({_UNSIGNED})", re.IGNORECASE)
_LESS_PREFIX = re.compile(r"(?:<=|=<|≤|≦|⩽|<)")
_GREATER_PREFIX = re.compile(r"(?:>=|=>|≥|≧|⩾|>)")


def _clean(raw: str) -> str:
    """전각 문자·특수 공백을 정리하고 앞뒤 공백을 없앤다."""
    return unicodedata.normalize("NFKC", raw).strip()


def _to_decimal(text: str) -> Decimal | None:
    """숫자 문자열을 Decimal로 바꾼다. 쉼표 소수점을 처리하고, 해석이 애매하면 None."""
    s = re.sub(r"\s+", "", text)
    if "," in s:
        if "." in s:
            # 1,234.5 처럼 천 단위 구분 쉼표만 허용
            if not re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+\.\d+", s):
                return None
            s = s.replace(",", "")
        elif s.count(",") == 1:
            left, right = s.split(",")
            # 0,003 은 소수점 쉼표. 1,000 처럼 천 단위인지 소수점인지 애매하면 추측하지 않는다.
            if left.lstrip("+-") not in ("", "0") and len(right) == 3:
                return None
            s = f"{left}.{right}"
        else:
            if not re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+", s):
                return None
            s = s.replace(",", "")
    if not re.fullmatch(_NUM, s):
        return None
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def parse_value(raw: str | None) -> tuple[str, Decimal | None]:
    """COA 결과값 원문을 (종류, Decimal) 로 해석한다. 종류는 number/less_than/not_detected/greater_than/range/invalid."""
    if raw is None:
        return INVALID, None
    s = _clean(raw)
    if not s:
        return INVALID, None
    if _ND_PATTERN.fullmatch(s):
        return NOT_DETECTED, None

    for kind, prefix in ((LESS_THAN, _LESS_PREFIX), (GREATER_THAN, _GREATER_PREFIX)):
        m = prefix.match(s)
        if m:
            value = _to_decimal(s[m.end():])
            return (kind, value) if value is not None else (INVALID, None)

    value = _to_decimal(s)
    if value is not None:
        return NUMBER, value
    if _RANGE_PATTERN.fullmatch(s):
        return RANGE, None
    return INVALID, None


def is_inclusive_bound(raw: str) -> bool:
    """'≤x', '<=x' 처럼 x를 포함하는 부등호 표기인지 ('<x'는 False)."""
    s = _clean(raw)
    m = _LESS_PREFIX.match(s) or _GREATER_PREFIX.match(s)
    return bool(m) and m.group(0) not in ("<", ">")


def normalize_unit(raw_unit: str | None) -> str | None:
    """단위 원문을 표준 단위(%, ppm, ppb, g/kg, APHA)로 바꾼다. 모르는 단위·빈 값은 None."""
    if raw_unit is None:
        return None
    # NFKC가 마이크로 기호(µ)를 그리스 문자(μ)로, 전각 ％를 %로 바꿔 준다
    key = re.sub(r"\s+", "", _clean(raw_unit)).lower()
    return _UNIT_LOOKUP.get(key)


def convert(value: Decimal, from_unit: str | None, to_unit: str | None) -> Decimal | None:
    """value를 from_unit에서 to_unit으로 환산한다. 질량분율끼리만 환산, 무차원 단위는 같은 단위일 때만 그대로. 불가하면 None."""
    src, dst = normalize_unit(from_unit), normalize_unit(to_unit)
    if src is None or dst is None:
        return None
    if src == dst:
        return value
    if src in PPM_FACTOR and dst in PPM_FACTOR:
        return value * PPM_FACTOR[src] / PPM_FACTOR[dst]
    return None
