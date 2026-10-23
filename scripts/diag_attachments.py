"""첨부 진단(읽기 전용): 받은편지함에서 제목으로 메일을 찾아 첨부 속성과 PDF 판별 결과를 출력한다.

실행: .venv\\Scripts\\python scripts\\diag_attachments.py ["COA N2418"]
      .venv\\Scripts\\python scripts\\diag_attachments.py --watch 120   (새 메일 도착 직후와 10초 뒤 첨부 개수 비교)
메일을 바꾸지 않는다(저장·분류·읽음 표시 없음).
"""

import argparse
import sys
import time
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8")

import pythoncom  # noqa: E402
import win32com.client  # noqa: E402

from core import outlook  # noqa: E402
from core.config import load_config  # noqa: E402

TYPE_NAME = {1: "일반(ByValue)", 4: "참조/링크(ByReference)", 5: "포함 항목(EmbeddedItem)", 6: "OLE"}
PROPS = {
    "PR_ATTACH_LONG_FILENAME": "http://schemas.microsoft.com/mapi/proptag/0x3707001F",
    "PR_ATTACH_MIME_TAG": "http://schemas.microsoft.com/mapi/proptag/0x370E001F",
    "PR_ATTACH_FLAGS": "http://schemas.microsoft.com/mapi/proptag/0x37140003",
    "PR_ATTACHMENT_HIDDEN": "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B",
}


def _get(obj, name):
    """속성을 읽고 실패하면 오류 문구를 돌려준다."""
    try:
        return getattr(obj, name)
    except Exception as exc:
        return f"(읽기 실패: {exc})"


def _prop(att, tag):
    """PropertyAccessor로 MAPI 속성을 읽는다."""
    try:
        return att.PropertyAccessor.GetProperty(tag)
    except Exception as exc:
        return f"(없음/읽기 실패: {str(exc)[:60]})"


def current_rule(att) -> tuple[bool, str]:
    """수정 전 코드의 PDF 판별 조건(Type == 1 AND FileName이 .pdf로 끝남)을 재현한다."""
    reasons = []
    if att.Type != outlook.OL_BY_VALUE:
        reasons.append(f"Type이 {att.Type}(일반=1 아님)")
    if not str(att.FileName).lower().endswith(".pdf"):
        reasons.append(f"FileName '{att.FileName}'이 .pdf로 끝나지 않음")
    return (not reasons), ("; ".join(reasons) or "조건 충족")


def describe(mail) -> None:
    """메일 1통의 첨부를 모두 출력한다."""
    print(f"제목: {mail.Subject!r}  수신: {outlook.received_at(mail)}  발신: {outlook.get_smtp_address(mail)}")
    print(f"분류: {mail.Categories!r}  Attachments.Count = {mail.Attachments.Count}  "
          f"PR_HASATTACH = {outlook.has_attach_flag(mail)}  폴더 = {mail.Parent.Name}")
    for i in range(1, mail.Attachments.Count + 1):
        att = mail.Attachments.Item(i)
        t = _get(att, "Type")
        print(f"  [{i}] Index={_get(att, 'Index')}  FileName={_get(att, 'FileName')!r}  "
              f"DisplayName={_get(att, 'DisplayName')!r}  Type={t} {TYPE_NAME.get(t, '')}  Size={_get(att, 'Size')}")
        for name, tag in PROPS.items():
            print(f"       {name} = {_prop(att, tag)!r}")
        ok, why = current_rule(att)
        print(f"       수정 전 코드 PDF 판별 → {ok} ({why})")
        kind, why = outlook.attachment_kind(att)
        print(f"       수정 후 코드 PDF 판별 → {kind} ({why})")


class Events:
    """새 메일 EntryID를 모은다."""

    ids: list[str] = []

    def OnNewMailEx(self, collection):
        Events.ids.extend(i for i in str(collection).split(",") if i)


def main() -> int:
    """제목으로 찾아 출력하거나(--watch 없음), 새 메일을 기다려 도착 직후·10초 뒤 첨부 수를 비교한다."""
    parser = argparse.ArgumentParser()
    parser.add_argument("subject", nargs="?", default="COA N2418", help="찾을 메일 제목(정확히 일치)")
    parser.add_argument("--watch", type=int, default=0, help="새 메일을 기다릴 시간(초)")
    args = parser.parse_args()
    cfg = load_config()

    if args.watch:
        app = win32com.client.DispatchWithEvents("Outlook.Application", Events)
        ns = app.GetNamespace("MAPI")
        print(f"새 메일 {args.watch}초 대기 중... 지금 PDF 첨부 메일을 보내세요.")
        end = time.monotonic() + args.watch
        while time.monotonic() < end:
            pythoncom.PumpWaitingMessages()
            while Events.ids:
                mail = ns.GetItemFromID(Events.ids.pop(0))
                print(f"도착 직후: 제목={mail.Subject!r} Attachments.Count={mail.Attachments.Count}")
                t10 = time.monotonic() + 10
                while time.monotonic() < t10:
                    pythoncom.PumpWaitingMessages()
                    time.sleep(0.2)
                mail = ns.GetItemFromID(mail.EntryID)
                print(f"10초 뒤  : Attachments.Count={mail.Attachments.Count}")
                describe(mail)
            time.sleep(0.2)
        return 0

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.get_inbox(ns, cfg["mail"].get("account", ""))
    found = search(inbox, args.subject, "받은 편지함")
    if not found:
        print(f"받은편지함에 제목 {args.subject!r} 메일이 없어 감시 계정의 모든 폴더를 찾습니다.")
        found = search_tree(inbox.Store.GetRootFolder(), args.subject, inbox.Store.DisplayName)
    if not found:
        print(f"제목 {args.subject!r} 메일을 찾지 못했습니다.")
    return 0


def search(folder, subject: str, path: str) -> int:
    """폴더에서 제목이 정확히 같은 메일을 찾아 출력한다."""
    found = 0
    items = folder.Items
    for i in range(1, items.Count + 1):
        item = items.Item(i)
        if getattr(item, "Class", None) == outlook.OL_MAIL and (item.Subject or "").strip() == subject:
            found += 1
            print("=" * 70)
            print(f"폴더: {path}")
            describe(item)
    return found


def search_tree(folder, subject: str, path: str) -> int:
    """하위 폴더까지 모두 찾는다."""
    found = search(folder, subject, path)
    for j in range(1, folder.Folders.Count + 1):
        sub = folder.Folders.Item(j)
        found += search_tree(sub, subject, f"{path}/{sub.Name}")
    return found


if __name__ == "__main__":
    sys.exit(main())
