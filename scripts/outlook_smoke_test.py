"""Outlook COM 연결 점검: 클래식 Outlook 접근과 첨부 저장 가능 여부 확인.

실행: .venv\\Scripts\\python scripts\\outlook_smoke_test.py [--wait 30] [--count 20]
"""

import argparse
import re
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.stdout.reconfigure(encoding="utf-8")  # 파일로 저장(> result.txt)해도 한글이 깨지지 않게

from core.config import load_config  # noqa: E402

try:
    import pythoncom
    import pywintypes
    import win32com.client
except ImportError as exc:
    print(f"[실패] pywin32를 불러올 수 없습니다: {exc}")
    print("  → .venv의 파이썬으로 실행했는지 확인하세요: .venv\\Scripts\\python scripts\\outlook_smoke_test.py")
    sys.exit(1)

OL_FOLDER_INBOX = 6
OL_MAIL = 43
OL_BY_VALUE = 1
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"

# HRESULT(부호 없는 32비트) → 원인 추정
HRESULT_HINTS = {
    0x800401F3: "Outlook.Application 클래스를 찾을 수 없음 → 클래식 Outlook이 설치돼 있지 않거나 새 Outlook(New Outlook)만 있음. 새 Outlook은 COM을 지원하지 않습니다.",
    0x80040154: "COM 클래스가 등록되지 않음 → 클래식 Outlook 미설치 또는 Office 설치 손상(새 Outlook만 쓰는 경우 포함).",
    0x80080005: "Outlook 서버 실행 실패 → 터미널과 Outlook의 권한 수준이 다를 가능성이 큽니다(관리자 권한 터미널 ↔ 일반 Outlook). 같은 권한으로 실행하세요.",
    0x80070005: "접근 거부 → 회사 보안 정책(그룹 정책/백신)이 외부 프로그램의 Outlook 접근을 막았거나, 권한 수준이 다릅니다.",
    0x80004004: "작업 취소됨 → '프로그래밍 방식 액세스' 보안 경고에서 거부를 눌렀거나, 정책이 자동 거부하도록 설정돼 있습니다.",
    0x800401E3: "실행 중인 Outlook을 찾지 못함 → Outlook을 먼저 실행하세요.",
    0x80010001: "Outlook이 호출을 거부(바쁨) → Outlook에 열린 대화상자(보안 경고, 프로필 선택 등)를 닫고 다시 실행하세요.",
    0x800706BA: "RPC 서버 사용 불가 → 실행 도중 Outlook이 종료됐습니다. Outlook을 켠 상태로 다시 실행하세요.",
    0x80010108: "COM 개체 연결 끊김 → 실행 도중 Outlook이 재시작됐습니다. 다시 실행하세요.",
    0x8004010F: "MAPI 항목을 찾지 못함 → Outlook 프로필/메일 계정이 설정돼 있지 않을 수 있습니다.",
    0x80040111: "MAPI 로그온 실패 → Outlook 프로필 로그인에 실패했습니다. Outlook을 직접 열어 로그인 상태를 확인하세요.",
}

# NewMailEx 핸들러가 EntryID만 넣는 큐 (처리는 메인 루프에서)
NEW_MAIL_IDS: list[str] = []


def explain_com_error(exc: pywintypes.com_error) -> str:
    """COM 오류를 HRESULT 기반의 한국어 원인 추정 문장으로 바꾼다."""
    hr = exc.hresult & 0xFFFFFFFF
    inner, desc = None, ""
    if exc.excepinfo:
        desc = exc.excepinfo[2] or ""
        if exc.excepinfo[5]:
            inner = exc.excepinfo[5] & 0xFFFFFFFF
    hint = HRESULT_HINTS.get(inner) or HRESULT_HINTS.get(hr)
    if hint is None and re.search(r"denied|거부|security|보안", desc, re.IGNORECASE):
        hint = HRESULT_HINTS[0x80004004]
    if hint is None:
        hint = "알 수 없는 COM 오류 → 아래 코드와 설명을 그대로 붙여넣어 주세요."
    code = f"HRESULT=0x{hr:08X}" + (f", 내부=0x{inner:08X}" if inner else "")
    return f"{hint}\n    ({code}, 설명: {desc or exc.strerror})"


def report_error(step: str, exc: Exception) -> None:
    """단계 실패 원인을 출력한다."""
    print(f"[실패] {step}")
    if isinstance(exc, pywintypes.com_error):
        print(f"  → {explain_com_error(exc)}")
    elif isinstance(exc, AttributeError) and "CLSIDToClassMap" in str(exc):
        print("  → win32com 캐시(gen_py)가 손상됐습니다. %TEMP%\\gen_py 폴더를 지우고 다시 실행하세요.")
        print(f"    ({exc!r})")
    else:
        print(f"  → 예상하지 못한 오류: {exc!r}")


def check_processes() -> None:
    """클래식 Outlook(OUTLOOK.EXE)과 새 Outlook(olk.exe) 실행 여부를 출력한다."""
    try:
        out = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"], capture_output=True, text=True, errors="ignore"
        ).stdout.lower()
    except OSError:
        print("[참고] 프로세스 목록을 확인하지 못했습니다.")
        return
    classic = '"outlook.exe"' in out
    new = '"olk.exe"' in out
    print(f"[0] 실행 중: 클래식 Outlook={'예' if classic else '아니오'}, 새 Outlook={'예' if new else '아니오'}")
    if new and not classic:
        print("  → 새 Outlook만 실행 중입니다. 새 Outlook은 COM을 지원하지 않으므로 클래식 Outlook으로 전환하세요.")
    elif not classic:
        print("  → 클래식 Outlook이 꺼져 있습니다. 연결 시 자동 실행을 시도하지만, 먼저 켜 두는 것을 권장합니다.")


def get_sender_smtp(mail) -> tuple[str, str]:
    """발신자 SMTP 주소와 얻은 방법을 반환. 사내(Exchange) 발신자는 GetExchangeUser로 변환한다."""
    if mail.SenderEmailType != "EX":
        return mail.SenderEmailAddress, "SenderEmailAddress"
    sender = mail.Sender
    if sender is not None:
        ex_user = sender.GetExchangeUser()
        if ex_user is not None and ex_user.PrimarySmtpAddress:
            return ex_user.PrimarySmtpAddress, "Sender.GetExchangeUser().PrimarySmtpAddress"
        try:
            return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS), "AddressEntry PR_SMTP_ADDRESS"
        except pywintypes.com_error:
            pass
    return mail.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS), "MailItem PR_SENDER_SMTP_ADDRESS"


def safe_filename(name: str) -> str:
    """Windows 파일명에 쓸 수 없는 문자를 _로 바꾼다."""
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "attachment.pdf"


def step1_connect():
    """Outlook에 연결해 받은편지함 이름과 메일 개수를 출력한다."""
    print("[1] Outlook 연결 중... (30초 넘게 멈추면 Outlook의 프로필 선택 창·보안 경고 창을 확인하세요)")
    app = win32com.client.Dispatch("Outlook.Application")
    ns = app.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    print(f"[성공] 연결됨: Outlook {app.Version}, 받은편지함='{inbox.Name}', 메일 {inbox.Items.Count}통")
    return ns, inbox


def step2_find_pdf(inbox, save_dir: Path, count: int):
    """최근 메일 count통에서 PDF 첨부 메일 1통을 찾아 첨부를 저장한다. (PDF 메일, 사내 발신 메일) 반환."""
    print(f"[2] 최근 메일 {count}통에서 PDF 첨부 찾는 중...")
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    pdf_mail, ex_mail = None, None
    checked = 0
    item = items.GetFirst()
    while item is not None and checked < count:
        if item.Class == OL_MAIL:
            checked += 1
            if ex_mail is None and item.SenderEmailType == "EX":
                ex_mail = item
            if pdf_mail is None:
                for i in range(1, item.Attachments.Count + 1):
                    att = item.Attachments.Item(i)
                    if att.Type == OL_BY_VALUE and att.FileName.lower().endswith(".pdf"):
                        save_dir.mkdir(parents=True, exist_ok=True)
                        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        path = save_dir / f"{stamp}_{safe_filename(att.FileName)}"
                        att.SaveAsFile(str(path))
                        print(f"[성공] 제목='{item.Subject}' 첨부 저장: {path}")
                        pdf_mail = item
                        break
        item = items.GetNext()
    if pdf_mail is None:
        print("[결과] PDF 첨부 메일 없음 (자신에게 PDF 첨부 메일을 보낸 뒤 다시 실행하세요)")
    return pdf_mail, ex_mail


def step3_sender(pdf_mail, ex_mail) -> None:
    """PDF 메일의 발신자 주소를 출력하고, 사내(Exchange) 발신자 SMTP 변환도 시험한다."""
    print("[3] 발신자 주소 확인 (보안 경고 창이 뜨면 '허용'을 눌러야 진행됩니다)")
    if pdf_mail is not None:
        smtp, how = get_sender_smtp(pdf_mail)
        print(f"  PDF 메일: 유형={pdf_mail.SenderEmailType}, 원본='{pdf_mail.SenderEmailAddress}'")
        print(f"[성공] SMTP 주소={smtp} (방법: {how})")
    if ex_mail is None:
        print("[결과] 최근 메일에 사내(Exchange, 유형 EX) 발신자가 없어 GetExchangeUser 변환은 시험하지 못했습니다.")
        return
    if ex_mail is not pdf_mail:
        smtp, how = get_sender_smtp(ex_mail)
        print(f"  사내 발신 메일: 제목='{ex_mail.Subject}', 원본='{ex_mail.SenderEmailAddress[:60]}...'")
        print(f"[성공] SMTP 주소={smtp} (방법: {how})")


class OutlookEvents:
    """NewMailEx 이벤트 핸들러: EntryID만 큐에 넣고 처리는 메인 루프에서 한다."""

    def OnNewMailEx(self, entry_id_collection):
        NEW_MAIL_IDS.extend(i for i in entry_id_collection.split(",") if i)


def step4_events(wait: int) -> int:
    """NewMailEx 이벤트를 wait초 동안 기다리며 새 메일 제목을 출력한다. 받은 메일 수 반환."""
    print(f"[4] 새 메일 이벤트 {wait}초 대기 중... 지금 자신에게 메일을 한 통 보내 보세요.")
    app = win32com.client.DispatchWithEvents("Outlook.Application", OutlookEvents)
    ns = app.GetNamespace("MAPI")
    received = 0
    end = time.monotonic() + wait
    next_notice = time.monotonic() + 10
    while time.monotonic() < end:
        pythoncom.PumpWaitingMessages()
        while NEW_MAIL_IDS:
            item = ns.GetItemFromID(NEW_MAIL_IDS.pop(0))
            received += 1
            print(f"[성공] 새 메일 감지: 제목='{item.Subject}'")
        if time.monotonic() >= next_notice:
            print(f"  ... 남은 시간 {int(end - time.monotonic())}초")
            next_notice += 10
        time.sleep(0.1)
    if received == 0:
        print("[결과] 대기 중 새 메일 이벤트 0건 (메일을 보냈다면: 받은편지함 규칙으로 다른 폴더로 옮겨졌는지, 동기화가 늦은지 확인)")
    return received


def main() -> int:
    """점검 단계를 순서대로 실행하고 요약과 판정을 출력한다."""
    parser = argparse.ArgumentParser(description="Outlook COM 연결 점검")
    parser.add_argument("--wait", type=int, default=30, help="새 메일 이벤트 대기 시간(초)")
    parser.add_argument("--count", type=int, default=20, help="PDF 첨부를 찾을 최근 메일 수")
    args = parser.parse_args()

    save_dir = load_config()["paths"]["inbox"] / "_smoke_test"
    results: dict[str, str] = {}

    print("=" * 60)
    print(f"Outlook 연동 시험  ({datetime.now():%Y-%m-%d %H:%M:%S}, Python {sys.version.split()[0]})")
    print("=" * 60)
    check_processes()

    try:
        ns, inbox = step1_connect()
        results["1 연결·받은편지함"] = "성공"
    except Exception as exc:
        report_error("Outlook 연결", exc)
        results["1 연결·받은편지함"] = "실패"
        ns = inbox = None

    if inbox is not None:
        pdf_mail = ex_mail = None
        try:
            pdf_mail, ex_mail = step2_find_pdf(inbox, save_dir, args.count)
            results["2 PDF 첨부 저장"] = "성공" if pdf_mail is not None else "PDF 메일 없음"
        except Exception as exc:
            report_error("PDF 첨부 찾기·저장", exc)
            results["2 PDF 첨부 저장"] = "실패"

        try:
            step3_sender(pdf_mail, ex_mail)
            if ex_mail is not None:
                results["3 발신자 주소"] = "성공"
            else:
                results["3 발신자 주소"] = "성공(사내 발신자 미시험)" if pdf_mail is not None else "미시험"
        except Exception as exc:
            report_error("발신자 주소 확인", exc)
            results["3 발신자 주소"] = "실패"

        try:
            received = step4_events(args.wait)
            results["4 새 메일 이벤트"] = f"성공({received}건)" if received else "0건"
        except Exception as exc:
            report_error("NewMailEx 이벤트 연결", exc)
            results["4 새 메일 이벤트"] = "실패"

    print("=" * 60)
    print("요약")
    for step, result in results.items():
        print(f"  {step}: {result}")
    if results["1 연결·받은편지함"] == "실패":
        print('판정 제안: Outlook 연결 불가 → config.toml의 watch.mode = "folder" 검토')
    elif all(r.startswith("성공") for r in results.values()):
        print('판정 제안: 모두 성공 → watch.mode = "outlook" 유지')
    else:
        print("판정 제안: 일부 미완료 → 위 결과를 그대로 붙여넣어 판단을 받으세요")
    print("=" * 60)
    return 0 if results["1 연결·받은편지함"] == "성공" else 1


if __name__ == "__main__":
    sys.exit(main())
