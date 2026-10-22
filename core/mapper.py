"""항목명 별칭(Iron, Fe, 철 등)을 사내 표준 항목명으로 매핑."""

import re
import unicodedata

from core.spec import SpecMaster


def normalize_name(name: str) -> str:
    """비교용 항목명: 전각 정리, 소문자화, 괄호·공백·특수문자 제거 ('Iron (Fe)' → 'ironfe')."""
    s = unicodedata.normalize("NFKC", name).lower()
    return re.sub(r"[\W_]+", "", s)


def _candidates(raw_name: str) -> list[str]:
    """매칭 후보: 전체 이름, 괄호를 뺀 이름, 괄호 안 내용 ('Moisture (KF)' → 'moisturekf', 'moisture', 'kf')."""
    s = unicodedata.normalize("NFKC", raw_name)
    inside = re.findall(r"\(([^()]*)\)", s)
    outside = re.sub(r"\([^()]*\)", " ", s)
    names = [s, outside, *inside]
    return [n for n in dict.fromkeys(normalize_name(x) for x in names) if n]


class ItemMapper:
    """별칭 사전으로 원문 항목명을 item_std로 바꾼다. 못 찾거나 후보가 서로 다른 항목을 가리키면 None."""

    def __init__(self, aliases: dict[str, str]):
        self._lookup: dict[str, str] = {}
        for alias, item_std in aliases.items():
            key = normalize_name(alias)
            if not key:
                continue
            if self._lookup.get(key, item_std) != item_std:
                raise ValueError(f"항목 별칭 '{alias}'이(가) 여러 항목({self._lookup[key]}, {item_std})에 등록돼 있습니다.")
            self._lookup[key] = item_std

    @classmethod
    def from_spec(cls, spec: SpecMaster, material_code: str | None = None) -> "ItemMapper":
        """규격 마스터의 항목 별칭으로 매퍼를 만든다."""
        return cls(spec.item_aliases(material_code))

    def map(self, raw_name: str | None) -> str | None:
        """원문 항목명 → item_std. 전체 이름이 맞으면 우선, 아니면 괄호 안/밖 후보가 한 항목으로 모일 때만 인정."""
        if not raw_name:
            return None
        candidates = _candidates(raw_name)
        if candidates and candidates[0] in self._lookup:
            return self._lookup[candidates[0]]
        found = {self._lookup[c] for c in candidates[1:] if c in self._lookup}
        return found.pop() if len(found) == 1 else None
