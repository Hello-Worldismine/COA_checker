"""PDF → CoaDocument 추출 (pdfplumber 규칙 기반, LLM 보조 훅은 비활성).

판정은 하지 않는다. 원문 항목명·값·단위를 그대로 담고, 읽다가 이상한 점은 extract_warnings에 남긴다.
실행: python -m core.extractor samples\\COA_Nippon_N2418_fail.pdf
"""

import re
import sys
import unicodedata
from pathlib import Path

import pdfplumber

from core.config import load_config
from core.mapper import ItemMapper
from core.models import CoaDocument, ExtractedItem
from core.spec import SpecMaster, load_spec_master
from core.units import INVALID, parse_value

MIN_TEXT_CHARS = 20  # 공백 제외 글자 수가 이보다 적으면 스캔 PDF로 본다
SCAN_WARNING = "텍스트 없음 - 스캔 PDF 의심(OCR은 확장 단계)"

# 표 헤더: 항목명 열과 결과 열을 찾는다. Specification/Limit 열은 측정값이 아니므로 쓰지 않는다.
ITEM_HEADER = re.compile(r"(?:test\s*)?items?|parameters?|tests?|characteristics?|analysis\s*items?|(?:시험\s*)?항목", re.I)
RESULT_HEADER = re.compile(r"(?:test\s*)?results?|values?|measured(?:\s*values?)?|actual(?:\s*values?)?|결과|측정값", re.I)
UNIT_HEADER = re.compile(r"units?|단위", re.I)

# 줄 나열 양식: "항목명 ....... 값 단위", "항목명: 값 단위", "항목명   값 단위"
LINE_PATTERN = re.compile(r"^(?P<name>\S.*?)\s*(?P<leader>\.{2,}|[:：]|\s{2,}|\t)\s*(?P<rest>\S.*)$")

# 값 + 단위 분리: "20ppm", "<10 ppm", "99.93 %", "0.5~1.0 %", "N.D."
VALUE_UNIT = re.compile(
    r"^(?P<value>(?:[<>≤≥≦≧]=?\s*)?[+-]?(?:\d[\d.,]*|\.\d+)(?:[eE][+-]?\d+)?"
    r"(?:\s*(?:~|-|–|to)\s*(?:\d[\d.,]*|\.\d+))?)\s*(?P<unit>.*)$",
    re.I,
)
ND_UNIT = re.compile(r"^(?P<value>n\.?d\.?|not\s+detected)\s*(?P<unit>.*)$", re.I)

LOT_PATTERN = re.compile(r"\b(?:lot|batch)\s*(?:no\.?|number|#)?\s*[:：]?\s*([A-Za-z0-9][A-Za-z0-9\-_/]*)", re.I)
DATE_PATTERN = re.compile(
    r"\b(?:mfg\.?\s*date|manufactur(?:e|ing)\s*date|date\s*of\s*(?:mfg\.?|manufacture)|"
    r"prod(?:uction|\.)?\s*date)\s*[:：]?\s*(.+)$",
    re.I | re.M,
)
PRODUCT_PATTERN = re.compile(r"\b(?:product\s*name|product|material\s*name|material)\s*[:：]\s*(.+)$", re.I | re.M)
EMAIL_PATTERN = re.compile(r"[\w.+-]+@([\w-]+(?:\.[\w-]+)+)")


def _cell(value) -> str:
    """표 셀을 한 줄 문자열로 (None은 "")."""
    return " ".join(unicodedata.normalize("NFKC", str(value)).split()) if value is not None else ""


def split_value_unit(text: str) -> tuple[str, str | None]:
    """'20ppm' → ('20', 'ppm'). 값 모양을 못 찾으면 (원문, None)."""
    s = " ".join(unicodedata.normalize("NFKC", text).split())
    m = VALUE_UNIT.match(s) or ND_UNIT.match(s)
    if not m:
        return s, None
    return m.group("value").strip(), (m.group("unit").strip() or None)


def items_from_table(table: list[list]) -> list[ExtractedItem]:
    """헤더(항목·결과·단위)를 찾아 표 행을 ExtractedItem으로. 헤더가 없으면 빈 목록."""
    for h, header in enumerate(table):
        cells = [_cell(c) for c in header]
        item_col = next((i for i, c in enumerate(cells) if ITEM_HEADER.fullmatch(c)), None)
        result_col = next((i for i, c in enumerate(cells) if RESULT_HEADER.fullmatch(c)), None)
        if item_col is None or result_col is None:
            continue
        unit_col = next((i for i, c in enumerate(cells) if UNIT_HEADER.fullmatch(c)), None)
        items = []
        for row in table[h + 1:]:
            row = [_cell(c) for c in row]
            if len(row) <= max(item_col, result_col) or not row[item_col]:
                continue
            value, unit = row[result_col], (row[unit_col] if unit_col is not None and unit_col < len(row) else "")
            if not unit:
                value, unit = split_value_unit(value)
            items.append(ExtractedItem(raw_name=row[item_col], raw_value=value, raw_unit=unit or None))
        return items
    return []


def items_from_lines(text: str, mapper: ItemMapper) -> list[ExtractedItem]:
    """줄 나열 양식에서 항목을 뽑는다. 점선 뒤 값이 읽히거나, 항목명이 규격 별칭에 있는 줄만 인정."""
    items = []
    for line in unicodedata.normalize("NFKC", text).splitlines():
        m = LINE_PATTERN.match(line.strip())
        if not m:
            continue
        name = m.group("name").strip()
        value, unit = split_value_unit(m.group("rest"))
        known = mapper.map(name) is not None
        dotted = m.group("leader").startswith("..")
        if known or (dotted and parse_value(value)[0] != INVALID):
            items.append(ExtractedItem(raw_name=name, raw_value=value, raw_unit=unit))
    return items


def _alias_regex(alias: str) -> re.Pattern:
    """별칭이 다른 영숫자에 붙지 않고 단독으로 나올 때만 맞는 정규식."""
    return re.compile(rf"(?<![a-z0-9]){re.escape(alias.lower())}(?![a-z0-9])")


def find_supplier(text: str, spec: SpecMaster) -> str | None:
    """본문에서 협력사 이름·별칭을 찾고(가장 먼저 나온 것), 없으면 메일 도메인으로 찾는다."""
    lower = " ".join(unicodedata.normalize("NFKC", text).lower().split())
    hits = []
    for s in spec.suppliers:
        for alias in (s.name, *s.aliases):
            m = _alias_regex(" ".join(alias.split())).search(lower)
            if m:
                hits.append((m.start(), s.name))
    if hits:
        return min(hits)[1]
    for domain in EMAIL_PATTERN.findall(text):
        supplier = spec.supplier_by_domain(domain)
        if supplier:
            return supplier.name
    return None


def find_material(text: str, spec: SpecMaster) -> tuple[str | None, set[str]]:
    """제품명(또는 본문)에서 재질 별칭을 찾아 material_code를 정한다. 여러 재질이 걸리면 None과 후보를 돌려준다."""
    lower = unicodedata.normalize("NFKC", text).lower()
    codes = {code for alias, code in spec.material_aliases().items() if alias and _alias_regex(alias).search(lower)}
    return (codes.pop() if len(codes) == 1 else None), codes


def _first(pattern: re.Pattern, text: str) -> str | None:
    """정규식 첫 매치의 1번 그룹 (앞뒤 공백 제거)."""
    m = pattern.search(text)
    return m.group(1).strip() if m else None


def llm_extract(pdf_path: Path, text: str) -> CoaDocument:
    """LLM 보조 추출 자리 (확장 단계). extract.llm_fallback = true 일 때만 호출되며, 지금은 미구현."""
    raise NotImplementedError("LLM 보조 추출은 확장 단계입니다(사내 승인 LLM 연결 후 구현).")


def extract(pdf_path: Path | str, spec: SpecMaster | None = None, llm_fallback: bool | None = None) -> CoaDocument:
    """PDF 1장을 읽어 CoaDocument로 만든다. 표 양식을 먼저, 없으면 줄 나열 양식을 시도한다."""
    pdf_path = Path(pdf_path)
    spec = spec or load_spec_master()
    if llm_fallback is None:
        llm_fallback = load_config()["extract"]["llm_fallback"]
    doc = CoaDocument(source_path=pdf_path)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
            tables = [t for page in pdf.pages for t in page.extract_tables()]
    except Exception as exc:  # 손상·암호화 PDF 등: 추측하지 않고 경고만 남긴다
        doc.extract_warnings.append(f"PDF를 열 수 없음: {exc}")
        return doc

    if len(re.sub(r"\s", "", text)) < MIN_TEXT_CHARS:
        doc.extract_warnings.append(SCAN_WARNING)
        return doc

    doc.product_name = _first(PRODUCT_PATTERN, text)
    doc.lot_no = _first(LOT_PATTERN, text)
    doc.mfg_date = _first(DATE_PATTERN, text)
    doc.supplier = find_supplier(text, spec)
    doc.material_code, candidates = find_material(doc.product_name or text, spec)

    for table in tables:
        doc.items = items_from_table(table)
        if doc.items:
            break
    if not doc.items:
        doc.items = items_from_lines(text, ItemMapper.from_spec(spec))

    if not doc.items and llm_fallback:
        try:
            return llm_extract(pdf_path, text)
        except NotImplementedError as exc:
            doc.extract_warnings.append(f"LLM 보조 추출 불가: {exc}")

    if not doc.items:
        doc.extract_warnings.append("시험 항목을 찾지 못함")
    if doc.lot_no is None:
        doc.extract_warnings.append("Lot No.를 찾지 못함")
    if doc.supplier is None:
        doc.extract_warnings.append("협력사를 찾지 못함")
    if doc.material_code is None:
        detail = f"(후보 여러 개: {', '.join(sorted(candidates))})" if candidates else ""
        doc.extract_warnings.append(f"재질을 정하지 못함{detail}")
    return doc


def _width(text: str) -> int:
    """콘솔 표시 폭 (한글 등 전각 문자는 2칸)."""
    return sum(2 if unicodedata.east_asian_width(ch) in "WF" else 1 for ch in text)


def _print_document(doc: CoaDocument, spec: SpecMaster) -> None:
    """추출 결과를 표 형태로 출력한다."""
    mapper = ItemMapper.from_spec(spec, doc.material_code)
    print(f"파일     : {doc.source_path}")
    print(f"협력사   : {doc.supplier}")
    print(f"제품명   : {doc.product_name}")
    print(f"재질코드 : {doc.material_code}")
    print(f"Lot No.  : {doc.lot_no}")
    print(f"제조일   : {doc.mfg_date}")
    rows = [("원문 항목명", "값", "단위", "표준 항목")]
    rows += [(i.raw_name, i.raw_value, i.raw_unit or "-", mapper.map(i.raw_name) or "(매핑 없음)") for i in doc.items]
    widths = [max(_width(r[c]) for r in rows) for c in range(4)]
    line = "+-" + "-+-".join("-" * w for w in widths) + "-+"
    print(line)
    for n, r in enumerate(rows):
        print("| " + " | ".join(v + " " * (w - _width(v)) for v, w in zip(r, widths)) + " |")
        if n == 0:
            print(line)
    print(line)
    print(f"경고     : {'; '.join(doc.extract_warnings) if doc.extract_warnings else '없음'}")


def main(argv: list[str]) -> int:
    """명령줄: PDF 경로를 받아 추출 결과를 출력한다."""
    sys.stdout.reconfigure(encoding="utf-8")
    if not argv:
        print("사용법: python -m core.extractor <PDF 경로> [<PDF 경로> ...]")
        return 2
    spec = load_spec_master()
    for n, path in enumerate(argv):
        if n:
            print()
        _print_document(extract(path, spec), spec)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
