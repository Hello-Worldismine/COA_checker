"""백그라운드 상시 실행: Outlook 새 메일 감지 → 첨부 저장 → 판정 → 알림.

실행: run_watcher.bat  또는  .venv\\Scripts\\python mail_watcher.py [--once] [--mode outlook|folder]
- outlook 모드: 시작 시 backlog_days 이내 미처리 COA 메일을 먼저 처리하고, NewMailEx 이벤트로 새 메일을 처리한다.
  이벤트를 놓쳐도 poll_seconds마다 최근 메일을 다시 훑는다.
- folder 모드: Outlook 연동이 막혔을 때 inbox 하위에 새로 생긴 PDF를 watchdog으로 감시한다(알림 메일 없음, 로그만).
메일 1통 처리 중 오류가 나도 프로그램은 계속 돈다. 그 메일은 처리 완료 분류를 달지 않아 다음에 다시 처리된다.
"""

import argparse
import logging
import queue
import sys
import time
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from pathlib import Path
from typing import Callable

from core import outlook
from core.config import load_config
from core.pipeline import process_pdf, summary_line
from core.reporter import is_file_recorded, is_processed
from core.spec import load_spec_master

log = logging.getLogger("mail_watcher")

MAX_RETRIES = 3          # 같은 메일이 이 횟수만큼 실패하면 이번 실행 동안은 더 시도하지 않는다
ATTACH_WAIT_SECONDS = 30  # 새 메일인데 PDF가 0개로 보이면 첨부 동기화를 이 시간까지 기다린다
ATTACH_WAIT_STEP = 3
FRESH_MAIL = timedelta(minutes=5)
RESCAN_MARGIN = timedelta(minutes=10)
NEW_MAIL_IDS: "queue.Queue[str]" = queue.Queue()


def setup_logging(logs_dir: Path) -> None:
    """로그를 logs/mail_watcher_YYYYMMDD.log(상세)와 콘솔(한 줄 상태)로 남긴다."""
    logs_dir.mkdir(parents=True, exist_ok=True)
    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.handlers.clear()
    file_handler = logging.FileHandler(logs_dir / f"mail_watcher_{datetime.now():%Y%m%d}.log", encoding="utf-8")
    file_handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(logging.Formatter("[%(asctime)s] %(message)s", datefmt="%H:%M:%S"))
    root.addHandler(file_handler)
    root.addHandler(console)


# ── 메일 1통 처리 (COM 객체는 메인 스레드에서만) ──────────────────
@dataclass
class Context:
    """메일 처리에 필요한 설정과 동작. 테스트에서는 alert·process를 가짜로 바꾼다."""

    config: dict
    alert: Callable[[outlook.Alert], None]
    process: Callable = process_pdf
    load_spec: Callable = load_spec_master
    refetch: Callable = lambda mail: mail          # 메일을 다시 읽어 최신 첨부 상태를 얻는다
    sleep: Callable[[float], None] = time.sleep    # 기다리는 동안 Outlook 메시지를 처리하는 sleep으로 바꿔 끼운다
    now: Callable[[], datetime] = datetime.now
    failures: dict[str, int] = field(default_factory=dict)


def wait_for_pdf(mail, ctx: Context, spec, check: outlook.MailCheck):
    """새 메일이거나 '첨부 있음' 표시가 있는데 PDF가 0개로 보이면, 3초 간격으로 최대 30초 다시 확인한다."""
    fresh = ctx.now() - outlook.received_at(mail) <= FRESH_MAIL
    if not (fresh or outlook.has_attach_flag(mail)):
        return mail, check
    log.info("PDF 첨부가 아직 안 보여 최대 %d초 기다립니다: '%s' (%s)", ATTACH_WAIT_SECONDS, mail.Subject,
             outlook.describe_attachments(mail))
    waited = 0
    while waited < ATTACH_WAIT_SECONDS:
        ctx.sleep(ATTACH_WAIT_STEP)
        waited += ATTACH_WAIT_STEP
        mail = ctx.refetch(mail)
        check = outlook.check_mail(mail, spec, ctx.config)
        if check.pdf_count:
            log.info("%d초 뒤 PDF 첨부 %d개 확인", waited, check.pdf_count)
            break
    return mail, check


def handle_mail(mail, ctx: Context) -> str:
    """메일 1통을 처리하고 결과를 돌려준다: ignored / skipped / no_pdf / done / partial.

    첨부별로 process_pdf를 실행하고, 하나가 실패해도 나머지는 계속한다.
    알림은 메일 1통당 요약 1통. 처리 완료 분류는 모든 첨부를 처리한 뒤에만 단다.
    """
    cfg = ctx.config
    category = cfg["mail"]["processed_category"]
    if getattr(mail, "Class", None) != outlook.OL_MAIL:
        return "ignored"
    if outlook.is_marked(mail, category):
        return "skipped"

    spec = ctx.load_spec()  # 메일마다 다시 읽어 규격 마스터 수정이 바로 반영되게
    check = outlook.check_mail(mail, spec, cfg)
    if check is None or not check.is_candidate:
        return "ignored"
    subject, received = mail.Subject or "", outlook.received_at(mail)

    if check.pdf_count == 0:
        mail, check = wait_for_pdf(mail, ctx, spec, check)
    if check.pdf_count == 0:
        detail = outlook.describe_attachments(mail)
        if mail.Attachments.Count == 0 and outlook.has_attach_flag(mail):
            # 서버에는 첨부가 있는데 이 PC에 아직 안 내려온 상태: '첨부 없음'으로 확정하지 않고 다음에 다시 본다
            log.warning("첨부가 아직 동기화되지 않아 미처리로 남김(다음 확인 때 다시 처리): '%s' (%s) - %s",
                        subject, check.sender, detail)
            return "pending"
        return no_pdf(mail, ctx, check, detail)

    entry_id = mail.EntryID
    db = cfg["paths"]["db"]
    log.info("COA 메일 감지: '%s' (%s, PDF %d개)", subject, check.sender, check.pdf_count)
    report: dict[str, int] = {}
    saved = outlook.save_pdf_attachments(
        mail, cfg["paths"]["inbox"], outlook.sender_label(check.sender, spec),
        skip=lambda key: is_processed(db, entry_id, key), report=report,
    )
    if not saved and not report.get("skipped"):
        detail = outlook.describe_attachments(mail)
        if report.get("failed"):
            log.warning("PDF 첨부를 저장하지 못해 미처리로 남김: '%s' - %s", subject, detail)
            return "partial"
        return no_pdf(mail, ctx, check, detail)  # 후보가 모두 PDF가 아니었음

    outputs, failures = [], []
    for key, path in saved:
        meta = {"source": "mail", "mail_entry_id": entry_id, "attachment_name": key,
                "sender": check.sender, "received_at": received}
        try:
            out = ctx.process(path, meta, spec=spec)
        except Exception:
            log.exception("첨부 처리 실패: %s", key)
            failures.append(key)
            continue
        outputs.append((key, out))
        if out.skipped:
            log.info("이미 처리한 첨부: %s", key)
            continue
        log.info(summary_line(out.result))
        for err in out.saved.errors if out.saved else []:
            log.warning(err)

    alert = outlook.build_alert(outputs, failures, subject, check.sender, received)
    if alert:
        try:
            ctx.alert(alert)
        except Exception:
            log.exception("알림 메일 만들기 실패 (판정 결과는 저장됨)")
    elif outputs:
        log.info("모두 적합 - 알림 메일 없음")

    if failures:
        log.warning("첨부 %d개 처리 실패 - 처리 완료 분류를 달지 않음(다음에 남은 첨부만 다시 처리)", len(failures))
        return "partial"
    outlook.mark_processed(mail, category)
    return "done"


def no_pdf(mail, ctx: Context, check: outlook.MailCheck, detail: str) -> str:
    """PDF 없는 COA 메일: 등록 발신자면 담당자 알림, 제목 키워드만 맞으면 로그만. 둘 다 처리 완료 분류."""
    subject = mail.Subject or ""
    if check.known_sender:
        log.info("COA 메일에 PDF 첨부 없음 → 담당자 알림: '%s' (%s) - %s", subject, check.sender, detail)
        ctx.alert(outlook.build_no_pdf_alert(subject, check.sender, outlook.received_at(mail)))
    else:
        log.info("제목 키워드만 맞고 PDF 없음 → 로그만 남김: '%s' (%s) - %s", subject, check.sender, detail)
    outlook.mark_processed(mail, ctx.config["mail"]["processed_category"])
    return "no_pdf"


def safe_handle(mail, ctx: Context) -> str:
    """handle_mail을 예외로부터 보호한다. 같은 메일이 계속 실패하면 MAX_RETRIES 후 이번 실행에서는 포기."""
    entry_id = getattr(mail, "EntryID", "?")
    if ctx.failures.get(entry_id, 0) >= MAX_RETRIES:
        return "gave_up"
    try:
        status = handle_mail(mail, ctx)
    except Exception:
        log.exception("메일 처리 중 오류 - 미처리로 남김: '%s'", getattr(mail, "Subject", "?"))
        status = "error"
    if status in ("error", "partial", "pending"):
        ctx.failures[entry_id] = ctx.failures.get(entry_id, 0) + 1
        if ctx.failures[entry_id] >= MAX_RETRIES:
            log.error("같은 메일이 %d번 실패해 이번 실행에서는 더 시도하지 않습니다. 로그를 확인하세요: '%s'",
                      MAX_RETRIES, getattr(mail, "Subject", "?"))
    return status


def scan_inbox(ns, inbox, ctx: Context, since: datetime) -> int:
    """받은편지함에서 since 이후 메일 중 미처리 COA 메일을 처리한다. 처리한 메일 수 반환."""
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    entry_ids = []
    item = items.GetFirst()
    while item is not None:
        if getattr(item, "Class", None) == outlook.OL_MAIL:
            if outlook.received_at(item) < since:
                break
            entry_ids.append(item.EntryID)
        item = items.GetNext()
    handled = 0
    for entry_id in reversed(entry_ids):  # 오래된 메일부터
        status = safe_handle(ns.GetItemFromID(entry_id), ctx)
        handled += status in ("done", "partial", "no_pdf", "pending")
    return handled


# ── Outlook 모드 ──────────────────────────────────────────────
class OutlookEvents:
    """NewMailEx 핸들러: EntryID만 큐에 넣는다(처리는 메인 루프에서)."""

    def OnNewMailEx(self, entry_id_collection):
        for entry_id in str(entry_id_collection).split(","):
            if entry_id:
                NEW_MAIL_IDS.put(entry_id)


def run_outlook(cfg: dict, once: bool = False) -> int:
    """Outlook 이벤트로 새 COA 메일을 처리한다."""
    import pythoncom
    import win32com.client

    try:
        app = win32com.client.DispatchWithEvents("Outlook.Application", OutlookEvents)
        ns = app.GetNamespace("MAPI")
        inbox = outlook.get_inbox(ns, cfg["mail"].get("account", ""))
    except Exception:
        log.exception("Outlook 연결 실패 - 클래식 Outlook 실행 여부를 확인하고 scripts\\outlook_smoke_test.py로 진단하세요. "
                      "연동이 막혔다면 config.toml의 watch.mode를 \"folder\"로 바꾸세요.")
        return 1
    account = outlook.find_account(ns, cfg["mail"].get("account", ""))
    outlook.ensure_category(ns, cfg["mail"]["processed_category"])
    def pump_sleep(seconds: float) -> None:
        """기다리는 동안에도 Outlook 메시지(새 메일 이벤트 등)를 처리한다."""
        end = time.monotonic() + seconds
        while time.monotonic() < end:
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)

    ctx = Context(cfg, alert=lambda a: outlook.send_alert(app, a, cfg, account),
                  refetch=lambda m: ns.GetItemFromID(m.EntryID), sleep=pump_sleep)

    alert_to = ", ".join(cfg["alert"].get("to", [])) or "(없음)"
    log.info("메일 감시 시작: 계정=%s, 받은편지함='%s', 알림=%s → %s",
             cfg["mail"].get("account") or "기본", inbox.Name, cfg["alert"]["mode"], alert_to)
    days = cfg["mail"].get("backlog_days", 7)
    started = datetime.now()
    n = scan_inbox(ns, inbox, ctx, started - timedelta(days=days))
    log.info("최근 %d일 미처리 COA 메일 %d통 처리 완료. 새 메일을 기다립니다 (Ctrl+C로 종료)", days, n)
    if once:
        return 0

    store_id = inbox.StoreID
    poll = cfg["watch"].get("poll_seconds", 60)
    last_scan, next_poll = started, time.monotonic() + poll
    try:
        while True:
            pythoncom.PumpWaitingMessages()
            while not NEW_MAIL_IDS.empty():
                entry_id = NEW_MAIL_IDS.get()
                try:
                    item = ns.GetItemFromID(entry_id)
                    if getattr(item.Parent, "StoreID", store_id) != store_id:
                        continue  # 감시 대상이 아닌 다른 계정의 메일
                except Exception:
                    log.exception("새 메일을 열 수 없음: %s", entry_id)
                    continue
                safe_handle(item, ctx)
            if poll and time.monotonic() >= next_poll:
                now = datetime.now()
                scan_inbox(ns, inbox, ctx, last_scan - RESCAN_MARGIN)  # 놓친 이벤트 보완
                last_scan, next_poll = now, time.monotonic() + poll
            time.sleep(0.2)
    except KeyboardInterrupt:
        log.info("메일 감시 종료")
    return 0


# ── 폴더 모드 ─────────────────────────────────────────────────
def wait_until_ready(path: Path, timeout: float = 30) -> bool:
    """파일 쓰기가 끝날 때까지 기다린다(크기가 1초간 그대로이고 쓰기 잠금이 풀리면 완료)."""
    end = time.monotonic() + timeout
    last = -1
    while time.monotonic() < end:
        try:
            size = path.stat().st_size
            if size > 0 and size == last:
                with path.open("ab"):
                    return True
            last = size
        except (FileNotFoundError, PermissionError):
            pass
        time.sleep(1)
    return False


def _folder_candidates(inbox: Path, since: datetime) -> list[Path]:
    """inbox 하위 PDF 중 since 이후 수정된 것 ('_'로 시작하는 폴더는 제외)."""
    found = []
    for p in sorted(inbox.rglob("*.pdf")):
        if any(part.startswith("_") for part in p.relative_to(inbox).parts):
            continue
        if datetime.fromtimestamp(p.stat().st_mtime) >= since:
            found.append(p)
    return found


def process_file(path: Path, cfg: dict) -> None:
    """폴더 모드에서 PDF 1개를 처리한다(이미 이력에 있으면 건너뜀)."""
    if is_file_recorded(cfg["paths"]["db"], path):
        return
    if not wait_until_ready(path):
        log.warning("파일 쓰기가 끝나지 않아 건너뜀: %s", path)
        return
    received = datetime.fromtimestamp(path.stat().st_mtime)
    out = process_pdf(path, {"source": "folder", "received_at": received})
    log.info(summary_line(out.result))
    for err in out.saved.errors if out.saved else []:
        log.warning(err)


def run_folder(cfg: dict, once: bool = False) -> int:
    """watchdog으로 inbox 하위에 새로 생긴 PDF를 처리한다."""
    from watchdog.events import FileSystemEventHandler
    from watchdog.observers import Observer

    inbox: Path = cfg["paths"]["inbox"]
    inbox.mkdir(parents=True, exist_ok=True)
    pending: "queue.Queue[Path]" = queue.Queue()

    class PdfHandler(FileSystemEventHandler):
        """새 PDF 경로만 큐에 넣는다(감시 스레드에서 실행)."""

        def on_created(self, event):
            if not event.is_directory and event.src_path.lower().endswith(".pdf"):
                pending.put(Path(event.src_path))

        def on_moved(self, event):
            if not event.is_directory and event.dest_path.lower().endswith(".pdf"):
                pending.put(Path(event.dest_path))

    days = cfg["mail"].get("backlog_days", 7)
    log.info("폴더 감시 시작: %s (Outlook 알림 메일 없음, 결과는 로그·output 폴더 확인)", inbox)
    for path in _folder_candidates(inbox, datetime.now() - timedelta(days=days)):
        try:
            process_file(path, cfg)
        except Exception:
            log.exception("파일 처리 실패: %s", path)
    if once:
        return 0

    observer = Observer()
    observer.schedule(PdfHandler(), str(inbox), recursive=True)
    observer.start()
    try:
        while True:
            try:
                path = pending.get(timeout=1)
            except queue.Empty:
                continue
            if any(part.startswith("_") for part in path.relative_to(inbox).parts):
                continue
            try:
                process_file(path, cfg)
            except Exception:
                log.exception("파일 처리 실패: %s", path)
    except KeyboardInterrupt:
        log.info("폴더 감시 종료")
    finally:
        observer.stop()
        observer.join()
    return 0


def main(argv: list[str] | None = None) -> int:
    """설정을 읽고 watch.mode에 따라 Outlook 또는 폴더 감시를 시작한다."""
    sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description="COA 메일 감시")
    parser.add_argument("--once", action="store_true", help="미처리 메일(파일)만 처리하고 종료")
    parser.add_argument("--mode", choices=["outlook", "folder"], help="config.toml의 watch.mode 대신 사용")
    args = parser.parse_args(argv)

    cfg = load_config()
    setup_logging(cfg["paths"]["logs"])
    log.info("DB 파일: %s", cfg["paths"]["db"])  # 화면 하단의 DB 경로와 같은지 확인용
    mode = args.mode or cfg["watch"]["mode"]
    return run_folder(cfg, args.once) if mode == "folder" else run_outlook(cfg, args.once)


if __name__ == "__main__":
    sys.exit(main())
