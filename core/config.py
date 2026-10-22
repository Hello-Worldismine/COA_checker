"""config.toml 로딩: 설정값을 dict로 읽고 [paths]는 프로젝트 루트 기준 절대 Path로 바꾼다."""

import tomllib
from pathlib import Path
from typing import Any

PROJECT_ROOT = Path(__file__).resolve().parent.parent
DEFAULT_CONFIG_PATH = PROJECT_ROOT / "config.toml"

REQUIRED_SECTIONS = ("paths", "mail", "alert", "watch", "extract")


def load_config(path: Path | str | None = None) -> dict[str, Any]:
    """config.toml을 읽어 dict로 반환한다. [paths] 값은 설정 파일 위치 기준 절대 Path로 변환한다."""
    config_path = Path(path) if path else DEFAULT_CONFIG_PATH
    if not config_path.is_file():
        raise FileNotFoundError(f"설정 파일을 찾을 수 없습니다: {config_path}")

    with config_path.open("rb") as f:
        config = tomllib.load(f)

    missing = [s for s in REQUIRED_SECTIONS if s not in config]
    if missing:
        raise ValueError(f"설정 파일에 필수 섹션이 없습니다: {', '.join(missing)}")

    base = config_path.resolve().parent
    config["paths"] = {
        key: (base / value).resolve() for key, value in config["paths"].items()
    }

    mode = config["alert"].get("mode", "display")
    if mode not in ("display", "send"):
        raise ValueError(f'alert.mode는 "display" 또는 "send"여야 합니다: {mode!r}')

    watch_mode = config["watch"].get("mode", "outlook")
    if watch_mode not in ("outlook", "folder"):
        raise ValueError(f'watch.mode는 "outlook" 또는 "folder"여야 합니다: {watch_mode!r}')

    return config
