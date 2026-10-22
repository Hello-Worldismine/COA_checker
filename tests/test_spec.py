"""core.spec 테스트: 규격 마스터 로딩과 Decimal 변환."""

from decimal import Decimal

import pytest
from openpyxl import Workbook

from core.spec import load_spec_master


def test_spec_items(spec_master):
    """NMP-EL 규격 4항목과 하한·상한·단위."""
    specs = {s.item_std: s for s in spec_master.specs_for("NMP-EL")}
    assert set(specs) == {"Purity", "Moisture", "Iron", "Color"}
    assert specs["Purity"].lower == Decimal("99.90") and specs["Purity"].upper is None
    assert specs["Moisture"].upper == Decimal("0.003") and specs["Moisture"].unit == "%"
    assert specs["Color"].upper == Decimal("10") and specs["Color"].unit == "APHA"
    assert spec_master.spec_version("NMP-EL") == "v1.0"


def test_numbers_are_exact_decimal(spec_master):
    """엑셀 float를 Decimal(str())로 바꿔 이진 오차가 없다."""
    for s in spec_master.specs:
        for v in (s.lower, s.upper):
            assert v is None or isinstance(v, Decimal)
    moisture = next(s for s in spec_master.specs if s.item_std == "Moisture")
    assert moisture.upper.as_tuple().exponent == -3  # Decimal(0.003)였다면 0.00299999... 로 길어진다


def test_aliases(spec_master):
    """항목·재질 별칭 사전."""
    items = spec_master.item_aliases("NMP-EL")
    assert items["KF"] == "Moisture" and items["Iron"] == "Iron" and items["색도"] == "Color"
    materials = spec_master.material_aliases()
    assert materials["1-Methyl-2-pyrrolidinone"] == "NMP-EL"
    assert materials["N-Methyl-2-pyrrolidone Electronic Grade"] == "NMP-EL"


def test_suppliers(spec_master):
    """도메인으로 협력사 찾기 (메일 주소·하위 도메인·대소문자)."""
    assert spec_master.supplier_by_domain("qc@alpha-chem.example").name == "Alpha Chemical Co., Ltd."
    assert spec_master.supplier_by_domain("mail.NIPPON-SOLV.example").name == "Nippon Solvent Co., Ltd."
    assert spec_master.supplier_by_domain("someone@naver.com") is None
    assert spec_master.supplier_by_domain("evil-alpha-chem.example") is None
    assert spec_master.supplier_domains == ["alpha-chem.example", "nippon-solv.example"]


def _write(path, spec_rows):
    """최소 규격 마스터를 만든다."""
    wb = Workbook()
    wb.active.title = "materials"
    wb.active.append(["material_code", "material_name", "name_aliases"])
    ws = wb.create_sheet("spec")
    ws.append(["material_code", "item_std", "aliases", "lower", "upper", "unit", "spec_version"])
    for r in spec_rows:
        ws.append(r)
    wb.create_sheet("suppliers").append(["supplier_name", "email_domain", "name_aliases"])
    wb.save(path)


def test_text_number_keeps_trailing_zero(tmp_path):
    """문자열로 적힌 '99.90'도 Decimal로 읽힌다."""
    path = tmp_path / "s.xlsx"
    _write(path, [["X", "Purity", "", "99.90", None, "wt%", "v2"]])
    item = load_spec_master(path).specs[0]
    assert item.lower == Decimal("99.90") and item.unit == "%"


@pytest.mark.parametrize("row", [
    ["X", "Purity", "", "abc", None, "%", "v1"],     # 숫자 아님
    ["X", "Purity", "", "99", None, "mg/L", "v1"],   # 모르는 단위
])
def test_invalid_master_raises(tmp_path, row):
    """규격 마스터 오류는 판정 전에 바로 알린다."""
    path = tmp_path / "s.xlsx"
    _write(path, [row])
    with pytest.raises(ValueError):
        load_spec_master(path)


def test_default_path_loads_project_master():
    """경로를 안 주면 config.toml의 data/spec_master.xlsx를 읽는다."""
    assert load_spec_master().specs_for("NMP-EL")
