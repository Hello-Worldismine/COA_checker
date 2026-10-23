"""시연 초기화: DB·inbox·output을 비우고 샘플 COA와 규격 마스터를 다시 만든다 (녹화를 여러 번 반복하기 위함).

실행: .venv\\Scripts\\python scripts\\reset_demo.py        (확인 질문 후 실행)
      .venv\\Scripts\\python scripts\\reset_demo.py --yes  (묻지 않고 실행)
Outlook 메일의 'COA검증완료' 분류는 건드리지 않는다 → 예전 시험 메일은 감시 프로그램이 다시 처리하지 않는다.
logs/ 는 남긴다(문제 추적용).
"""

import argparse
import shutil
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8")

from core.config import load_config  # noqa: E402
from scripts import make_samples  # noqa: E402

KEEP = {".gitkeep"}


def clear_dir(folder: Path) -> int:
    """폴더 안을 비운다(.gitkeep은 남김). 지운 항목 수 반환."""
    if not folder.exists():
        folder.mkdir(parents=True)
        return 0
    removed = 0
    for child in folder.iterdir():
        if child.name in KEEP:
            continue
        if child.is_dir():
            shutil.rmtree(child)
        else:
            child.unlink()
        removed += 1
    return removed


def remove_db(db: Path) -> bool:
    """DB 파일과 WAL 부속 파일(-wal, -shm)을 지운다. 지운 게 있으면 True."""
    removed = False
    for path in (db, db.with_name(db.name + "-wal"), db.with_name(db.name + "-shm")):
        if path.exists():
            path.unlink()
            removed = True
    return removed


def reset(paths: dict[str, Path]) -> dict[str, int | bool]:
    """DB 삭제 → inbox·output 비우기 → 샘플·규격 마스터 재생성. 결과 요약을 돌려준다."""
    summary = {"db": remove_db(paths["db"]), "inbox": clear_dir(paths["inbox"]), "output": clear_dir(paths["output"])}
    pdfs = make_samples.build_all(paths["samples"], paths["spec_master"])
    summary["samples"] = len(pdfs)
    return summary


def main() -> int:
    """확인 질문 후 초기화한다."""
    parser = argparse.ArgumentParser(description="시연 초기화")
    parser.add_argument("--yes", action="store_true", help="확인 질문 없이 실행")
    args = parser.parse_args()
    paths = load_config()["paths"]

    print("다음을 삭제하고 샘플을 다시 만듭니다 (Outlook 메일 분류와 logs는 그대로).")
    print(f"  DB     : {paths['db']} (+ -wal, -shm)")
    print(f"  inbox  : {paths['inbox']} 안의 모든 파일·폴더")
    print(f"  output : {paths['output']} 안의 모든 파일·폴더 (결과 엑셀·ERP 양식·통보서)")
    print(f"  재생성 : {paths['samples']} 의 샘플 PDF 5장·expected.json, {paths['spec_master']}")
    print("※ 감시 프로그램(run_watcher.bat)·화면(run_app.bat)·결과 엑셀을 먼저 닫아 주세요.")
    if not args.yes and input("계속할까요? [y/N] ").strip().lower() not in ("y", "yes"):
        print("취소했습니다.")
        return 1
    try:
        summary = reset(paths)
    except PermissionError as exc:
        print(f"[실패] 파일이 사용 중이라 지울 수 없습니다: {exc.filename}")
        print("  → 감시 프로그램·화면·엑셀을 닫고 다시 실행하세요.")
        return 1
    print(f"[완료] DB {'삭제' if summary['db'] else '없음'}, inbox {summary['inbox']}개·output {summary['output']}개 "
          f"항목 삭제, 샘플 PDF {summary['samples']}장 재생성")
    return 0


if __name__ == "__main__":
    sys.exit(main())
