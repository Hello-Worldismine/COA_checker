"""process_pdf(): 추출 → 매핑 → 환산 → 판정 → 저장을 묶는 단일 진입점.

메일 경로·폴더 감시·수동 업로드 모두 이 함수를 호출한다. (저장은 다음 단계에서 추가)
실행: python -m core.pipeline samples\\COA_Nippon_N2418_fail.pdf
"""

import os
import sys
from pathlib import Path
from typing import Any

from core.extractor import extract, format_table
from core.judge import fmt, judge, limit_text
from core.models import CoaResult, Status
from core.spec import SpecMaster, load_spec_master

STATUS_LABEL = {Status.PASS: "적합", Status.FAIL: "부적합", Status.REVIEW: "검토 필요"}
ANSI = {Status.FAIL: "\033[1;97;41m", Status.REVIEW: "\033[1;30;43m", Status.PASS: "\033[32m"}
ANSI_RESET = "\033[0m"


def process_pdf(pdf_path: Path | str, meta: dict[str, Any] | None = None,
                spec: SpecMaster | None = None) -> CoaResult:
    """PDF 1장을 추출하고 규격 판정한 CoaResult를 반환한다. meta는 결과에 그대로 담는다(메일 발신자·수신시각·EntryID 등)."""
    spec = spec or load_spec_master()
    result = judge(extract(pdf_path, spec), spec)
    result.meta = dict(meta or {})
    return result


def _paint(text: str, status: Status, color: bool) -> str:
    """판정에 맞는 색을 입힌다 (터미널일 때만)."""
    return f"{ANSI[status]}{text}{ANSI_RESET}" if color else text


def print_result(result: CoaResult, color: bool = False) -> None:
    """판정표를 콘솔에 출력한다. FAIL 행은 '!!'와 빨간 배경, REVIEW 행은 '?'와 노란 배경으로 강조."""
    doc = result.document
    status = result.overall_status
    print(f"파일     : {doc.source_path}")
    print(f"협력사   : {doc.supplier}   재질: {doc.material_code}   Lot: {doc.lot_no}   Spec: {result.spec_version}")
    print("문서 판정: " + _paint(f" {status.value} ({STATUS_LABEL[status]}) ", status, color))

    rows = [("", "항목", "원문", "환산값", "규격", "판정", "사유")]
    for r in result.item_results:
        raw = f"{r.raw_value} {r.raw_unit or ''}".strip() if r.raw_value is not None else "-"
        norm = f"{fmt(r.normalized_value)} {r.spec_unit}" if r.normalized_value is not None else "-"
        limit = limit_text(r.lower, r.upper, r.spec_unit)
        mark = "!!" if r.status is Status.FAIL else ("?" if r.status is Status.REVIEW else "")
        rows.append((mark, r.item_std, raw, norm, limit, r.status.value, r.reason))
    lines = format_table(rows)
    statuses = [None, None, None] + [r.status for r in result.item_results] + [None]
    for line, st in zip(lines, statuses):
        print(_paint(line, st, color) if color and st in (Status.FAIL, Status.REVIEW) else line)

    if result.reasons:
        print("문서 검토 사유: " + "; ".join(result.reasons))
    if result.extra_items:
        print("판정 제외(Spec에 없는 항목): " + ", ".join(f"{i.raw_name}={i.raw_value}" for i in result.extra_items))


def main(argv: list[str]) -> int:
    """명령줄: PDF 경로들을 판정해 판정표를 출력한다."""
    sys.stdout.reconfigure(encoding="utf-8")
    if not argv:
        print("사용법: python -m core.pipeline <PDF 경로> [<PDF 경로> ...]")
        return 2
    color = sys.stdout.isatty()
    if color and os.name == "nt":
        os.system("")  # 윈도우 콘솔에서 ANSI 색 사용 켜기
    spec = load_spec_master()
    for n, path in enumerate(argv):
        if n:
            print()
        print_result(process_pdf(path, spec=spec), color)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
