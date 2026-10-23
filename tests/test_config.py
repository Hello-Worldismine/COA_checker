"""core.config 로딩 테스트."""

from pathlib import Path

import pytest

from core.config import load_config


def test_load_default_config():
    """기본 config.toml을 읽고 경로가 절대 Path로 바뀌는지 확인."""
    cfg = load_config()
    assert isinstance(cfg["paths"]["inbox"], Path)
    assert cfg["paths"]["inbox"].is_absolute()
    assert cfg["mail"]["subject_keyword"] == "COA"
    assert cfg["mail"]["processed_category"] == "COA검증완료"
    assert cfg["alert"]["mode"] in ("display", "send")
    assert cfg["watch"]["mode"] == "outlook"
    assert cfg["extract"]["llm_fallback"] is False


def test_invalid_alert_mode(tmp_path):
    """alert.mode가 허용값이 아니면 오류."""
    p = tmp_path / "config.toml"
    p.write_text(
        '[paths]\n[mail]\n[alert]\nmode = "auto"\n[watch]\n[extract]\n',
        encoding="utf-8",
    )
    with pytest.raises(ValueError):
        load_config(p)


def test_missing_file(tmp_path):
    """설정 파일이 없으면 FileNotFoundError."""
    with pytest.raises(FileNotFoundError):
        load_config(tmp_path / "none.toml")
