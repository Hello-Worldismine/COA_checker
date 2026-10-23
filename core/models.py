"""pydantic 데이터 모델 (ExtractedItem, CoaDocument, ItemResult, CoaResult)."""

from datetime import datetime
from decimal import Decimal
from enum import StrEnum
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field


class Status(StrEnum):
    """판정 상태: 적합 / 부적합 / 검토 필요."""

    PASS = "PASS"
    FAIL = "FAIL"
    REVIEW = "REVIEW"


class ExtractedItem(BaseModel):
    """PDF에서 읽은 시험 항목 1줄 (원문 그대로)."""

    raw_name: str
    raw_value: str
    raw_unit: str | None = None


class CoaDocument(BaseModel):
    """PDF 1장에서 추출한 COA 내용. 못 읽은 값은 None, 추출 중 이상은 extract_warnings에 남긴다."""

    source_path: Path
    supplier: str | None = None
    material_code: str | None = None
    product_name: str | None = None
    lot_no: str | None = None
    mfg_date: str | None = None
    items: list[ExtractedItem] = Field(default_factory=list)
    extract_warnings: list[str] = Field(default_factory=list)


class ItemResult(BaseModel):
    """Spec 항목 1개의 판정 근거와 결과. COA에 없는 항목이면 raw_* 는 None."""

    item_std: str
    raw_name: str | None = None
    raw_value: str | None = None
    raw_unit: str | None = None
    normalized_value: Decimal | None = None
    spec_unit: str
    lower: Decimal | None = None
    upper: Decimal | None = None
    status: Status
    reason: str = ""


class CoaResult(BaseModel):
    """COA 문서 1장의 판정 결과 전체.

    reasons: 항목이 아닌 문서 단위 검토 사유(재질 미확인, 추출 경고 등).
    extra_items: COA에만 있고 Spec에 없는 항목(기록만, 판정 제외).
    meta: 메일 경로일 때 발신자·수신시각·EntryID 등 호출자가 넘긴 정보.
    """

    document: CoaDocument
    item_results: list[ItemResult] = Field(default_factory=list)
    overall_status: Status
    reasons: list[str] = Field(default_factory=list)
    extra_items: list[ExtractedItem] = Field(default_factory=list)
    spec_version: str | None = None
    meta: dict[str, Any] = Field(default_factory=dict)
    processed_at: datetime = Field(default_factory=datetime.now)
