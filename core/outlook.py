"""Outlook COM 연동: COA 메일 선별·첨부 저장·분류 표시·담당자 알림 메일.

COM 객체는 만든 스레드(메인 스레드)에서만 쓴다. 알림 본문 생성(build_*)은 COM 없이 동작해 테스트할 수 있다.
협력사에는 아무것도 보내지 않는다. 알림은 config.toml의 alert.to(담당자)에게만.
"""

import html
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any

from core.judge import fmt, limit_text
from core.models import ProcessOutput, Status
from core.pipeline import STATUS_LABEL, short_supplier, summary_line
from core.reporter import safe_part
from core.spec import SpecMaster

log = logging.getLogger(__name__)

OL_MAIL_ITEM = 0       # CreateItem 종류: 메일
OL_FOLDER_INBOX = 6
OL_MAIL = 43           # MailItem.Class
OL_BY_VALUE = 1        # 일반 첨부 파일
PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E"
PR_SENDER_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"

CELL_COLOR = {Status.FAIL: "#FFC7CE", Status.REVIEW: "#FFEB9C", Status.PASS: "#C6EFCE"}
ALERT_SUBJECT_PREFIX = "[COA "  # 우리가 보내는 알림 제목 머리. 이 제목의 메일은 다시 감지하지 않는다.


# ── 메일 정보 ─────────────────────────────────────────────────
def get_smtp_address(mail) -> str:
    """발신자 SMTP 주소. 사내(Exchange, 유형 EX) 발신자는 '/O=...' 대신 실제 메일 주소로 바꾼다."""
    if getattr(mail, "SenderEmailType", "SMTP") != "EX":
        return (mail.SenderEmailAddress or "").strip()
    sender = mail.Sender
    if sender is not None:
        ex_user = sender.GetExchangeUser()
        if ex_user is not None and ex_user.PrimarySmtpAddress:
            return ex_user.PrimarySmtpAddress
        try:
            return sender.PropertyAccessor.GetProperty(PR_SMTP_ADDRESS)
        except Exception:
            pass
    return mail.PropertyAccessor.GetProperty(PR_SENDER_SMTP_ADDRESS)


def received_at(mail) -> datetime:
    """수신시각(로컬, 시간대 없음). pywin32는 로컬 시각에 UTC 표시를 붙여 주므로 값만 옮긴다."""
    t = mail.ReceivedTime
    return datetime(t.year, t.month, t.day, t.hour, t.minute, t.second)


def pdf_attachments(mail) -> list:
    """일반 첨부 중 .pdf 파일만."""
    result = []
    for i in range(1, mail.Attachments.Count + 1):
        att = mail.Attachments.Item(i)
        if att.Type == OL_BY_VALUE and str(att.FileName).lower().endswith(".pdf"):
            result.append(att)
    return result


def _domain_match(domain: str, allowed: str) -> bool:
    """도메인이 같거나 하위 도메인인지."""
    allowed = allowed.lower().lstrip("@").strip()
    return bool(allowed) and (domain == allowed or domain.endswith("." + allowed))


def is_known_sender(address: str, spec: SpecMaster, config: dict) -> bool:
    """협력사 도메인(spec_master)·추가 허용 도메인·시연용 발신자에 해당하는지."""
    address = (address or "").strip().lower()
    if not address:
        return False
    if address in {s.strip().lower() for s in config["mail"].get("test_senders", [])}:
        return True
    domain = address.rsplit("@", 1)[-1]
    allowed = list(spec.supplier_domains) + list(config["mail"].get("allowed_domains", []))
    return any(_domain_match(domain, d) for d in allowed)


@dataclass
class MailCheck:
    """COA 메일 판별 결과."""

    sender: str
    keyword_hit: bool
    known_sender: bool
    pdf_count: int
    own_alert: bool = False

    @property
    def is_candidate(self) -> bool:
        """제목 키워드 또는 등록 발신자 중 하나라도 맞음 (PDF 여부 무관). 우리가 보낸 알림 메일은 제외."""
        return not self.own_alert and (self.keyword_hit or self.known_sender)

    @property
    def is_coa(self) -> bool:
        """(제목 키워드 OR 등록 발신자) AND PDF 첨부 1개 이상."""
        return self.is_candidate and self.pdf_count > 0


def check_mail(mail, spec: SpecMaster, config: dict) -> MailCheck | None:
    """메일이 COA 메일인지 판별한다. MailItem(Class 43)이 아니면 None."""
    if getattr(mail, "Class", None) != OL_MAIL:
        return None
    keyword = config["mail"].get("subject_keyword", "COA")
    subject = mail.Subject or ""
    sender = get_smtp_address(mail)
    return MailCheck(
        sender=sender,
        keyword_hit=bool(keyword) and keyword.lower() in subject.lower(),
        known_sender=is_known_sender(sender, spec, config),
        pdf_count=len(pdf_attachments(mail)),
        own_alert=subject.strip().startswith(ALERT_SUBJECT_PREFIX),
    )


def is_coa_mail(mail, spec: SpecMaster, config: dict) -> bool:
    """(제목 키워드 OR 협력사/허용 도메인/테스트 발신자) AND PDF 첨부 1개 이상. MailItem만."""
    check = check_mail(mail, spec, config)
    return check is not None and check.is_coa


def sender_label(address: str, spec: SpecMaster) -> str:
    """파일명용 발신처 이름: 등록 협력사면 협력사명, 아니면 메일 주소."""
    supplier = spec.supplier_by_domain(address) if address else None
    return safe_part(supplier.name if supplier else address, "발신자미상")


# ── 첨부 저장 ─────────────────────────────────────────────────
def attachment_filename(received: datetime, label: str, original: str, occurrence: int = 1) -> str:
    """표준 파일명 '{수신일시}_{발신처}_{원본명}.pdf'. 금지문자 정리, 같은 메일 안 동명 첨부는 _2, _3."""
    stem = original[:-4] if original.lower().endswith(".pdf") else original
    stem = re.sub(r'[\\/:*?"<>|\x00-\x1f]+', "_", stem).strip(" ._") or "COA"
    stem = stem[:80]
    suffix = f"_{occurrence}" if occurrence > 1 else ""
    return f"{received:%Y%m%d_%H%M%S}_{label}_{stem}{suffix}.pdf"


def attachment_key(original: str, occurrence: int) -> str:
    """DB에 남기는 첨부 이름 (같은 메일 안 동명 첨부는 '#2' 를 붙여 구분)."""
    return original if occurrence == 1 else f"{original}#{occurrence}"


def save_pdf_attachments(mail, inbox_dir: Path, label: str, skip=None) -> list[tuple[str, Path]]:
    """PDF 첨부를 inbox/{YYYYMMDD}/ 표준 이름으로 저장하고 [(첨부 키, 경로)]를 반환한다.

    skip(첨부 키)가 True인 첨부(이미 처리됨)는 저장하지 않는다. 같은 메일을 다시 처리하면 같은 경로에 덮어쓴다.
    """
    received = received_at(mail)
    folder = inbox_dir / f"{received:%Y%m%d}"
    seen: dict[str, int] = {}
    saved = []
    for att in pdf_attachments(mail):
        original = str(att.FileName)
        seen[original.lower()] = seen.get(original.lower(), 0) + 1
        key = attachment_key(original, seen[original.lower()])
        if skip and skip(key):
            log.info("이미 처리한 첨부라 건너뜀: %s", key)
            continue
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / attachment_filename(received, label, original, seen[original.lower()])
        att.SaveAsFile(str(path))
        saved.append((key, path))
    return saved


# ── 처리 완료 분류 ────────────────────────────────────────────
def _categories(mail) -> list[str]:
    """메일의 분류 목록."""
    return [c.strip() for c in (mail.Categories or "").split(",") if c.strip()]


def is_marked(mail, category: str) -> bool:
    """이미 처리 완료 분류가 붙어 있는지."""
    return category in _categories(mail)


def mark_processed(mail, category: str) -> None:
    """처리 완료 분류를 추가하고 저장한다."""
    cats = _categories(mail)
    if category not in cats:
        mail.Categories = ", ".join(cats + [category])
        mail.Save()


def ensure_category(ns, category: str) -> None:
    """Outlook 분류 목록에 처리 완료 분류가 없으면 만든다(색 표시용). 실패해도 무시."""
    try:
        names = [ns.Categories.Item(i).Name for i in range(1, ns.Categories.Count + 1)]
        if category not in names:
            ns.Categories.Add(category)
    except Exception:
        log.debug("분류 목록 추가 실패(무시)", exc_info=True)


# ── 계정·받은편지함 ───────────────────────────────────────────
def find_account(ns, address: str):
    """SMTP 주소로 Outlook 계정을 찾는다. 없으면 None."""
    if not address:
        return None
    for i in range(1, ns.Accounts.Count + 1):
        acc = ns.Accounts.Item(i)
        if str(acc.SmtpAddress).lower() == address.lower():
            return acc
    return None


def get_inbox(ns, account_address: str):
    """감시할 받은편지함. config mail.account가 있으면 그 계정, 없으면 기본 계정."""
    if account_address:
        acc = find_account(ns, account_address)
        if acc is None:
            raise RuntimeError(f"Outlook 프로필에 '{account_address}' 계정이 없습니다. config.toml의 mail.account를 확인하세요.")
        return acc.DeliveryStore.GetDefaultFolder(OL_FOLDER_INBOX)
    return ns.GetDefaultFolder(OL_FOLDER_INBOX)


# ── 알림 메일 ─────────────────────────────────────────────────
@dataclass
class Alert:
    """담당자 알림 메일 내용."""

    subject: str
    html: str
    attachments: list[Path] = field(default_factory=list)


def _e(value: Any) -> str:
    """HTML 이스케이프."""
    return html.escape("" if value is None else str(value))


def _status_cell(status: Status | None, label: str | None = None) -> str:
    """판정 색을 입힌 표 칸 (status가 None이면 회색, label로 문구 지정)."""
    color = CELL_COLOR[status] if status else "#D9D9D9"
    text = label or f"{status.value} ({STATUS_LABEL[status]})"
    return f'<td style="background:{color};font-weight:bold">{_e(text)}</td>'


TABLE = 'style="border-collapse:collapse;font-family:Malgun Gothic,Arial;font-size:10pt" border="1" cellpadding="4"'
HEAD = 'style="background:#305496;color:#fff"'


def build_alert(outputs: list[tuple[str, ProcessOutput]], failures: list[str], mail_subject: str,
                sender: str, received: datetime | None) -> Alert | None:
    """메일 1통의 처리 결과로 담당자 요약 알림 1통을 만든다. 모두 PASS면 None(알림 없음).

    문서별 판정 요약 + FAIL/REVIEW 항목 상세 + 결과 엑셀 첨부. 처리 실패한 첨부도 표에 남긴다.
    """
    docs = [(name, o) for name, o in outputs if o.result is not None and not o.skipped]
    problems = [(name, o) for name, o in docs if o.result.overall_status is not Status.PASS]
    if not problems and not failures:
        return None

    has_fail = any(o.result.overall_status is Status.FAIL for _, o in problems)
    tag = "[COA 부적합]" if has_fail else "[COA 검토 필요]"
    if problems:
        first = problems[0][1].result.document
        label = f"{short_supplier(first.supplier)} / {first.lot_no or 'Lot 미상'}"
        extra = len(problems) + len(failures) - 1
        if extra > 0:
            label += f" 외 {extra}건"
    else:
        label = f"첨부 처리 실패 {len(failures)}건"
    subject = f"{tag} {label}"

    counts = {s: sum(o.result.overall_status is s for _, o in docs) for s in Status}
    parts = [
        '<div style="font-family:Malgun Gothic,Arial;font-size:10pt">',
        f"<p><b>COA 자동검증 결과</b> — 부적합 {counts[Status.FAIL]}건, 검토 필요 {counts[Status.REVIEW]}건, "
        f"적합 {counts[Status.PASS]}건" + (f", 처리 실패 {len(failures)}건" if failures else "") + "</p>",
        f"<p>수신 메일: {_e(mail_subject)}<br>발신자: {_e(sender)}<br>"
        f"수신시각: {_e(f'{received:%Y-%m-%d %H:%M:%S}' if received else '-')}</p>",
        f"<table {TABLE}><tr {HEAD}><th>첨부 파일</th><th>협력사</th><th>재질</th><th>Lot</th><th>판정</th><th>요약</th></tr>",
    ]
    for name, o in docs:
        d = o.result.document
        parts.append(f"<tr><td>{_e(name)}</td><td>{_e(d.supplier or '-')}</td><td>{_e(d.material_code or '-')}</td>"
                     f"<td>{_e(d.lot_no or '-')}</td>{_status_cell(o.result.overall_status)}"
                     f"<td>{_e(summary_line(o.result).split('→', 1)[-1].strip())}</td></tr>")
    for name in failures:
        parts.append(f"<tr><td>{_e(name)}</td><td>-</td><td>-</td><td>-</td>{_status_cell(None, '처리 실패')}"
                     "<td>처리 중 오류 - 로그 확인, 다음 실행 때 다시 처리</td></tr>")
    parts.append("</table>")

    attachments = []
    for name, o in problems:
        r = o.result
        parts.append(f"<h3 style=\"margin-bottom:4px\">{_e(name)} — {_e(r.overall_status.value)} "
                     f"({_e(STATUS_LABEL[r.overall_status])})</h3>")
        if r.reasons:
            parts.append(f"<p>문서 검토 사유: {_e('; '.join(r.reasons))}</p>")
        rows = [i for i in r.item_results if i.status is not Status.PASS]
        if rows:
            parts.append(f"<table {TABLE}><tr {HEAD}><th>항목</th><th>원문 값</th><th>환산값</th><th>규격</th>"
                         "<th>판정</th><th>사유</th></tr>")
            for i in rows:
                raw = f"{i.raw_value} {i.raw_unit or ''}".strip() if i.raw_value is not None else "-"
                norm = f"{fmt(i.normalized_value)} {i.spec_unit}" if i.normalized_value is not None else "-"
                parts.append(f"<tr><td>{_e(i.item_std)} ({_e(i.raw_name or '-')})</td><td>{_e(raw)}</td>"
                             f"<td>{_e(norm)}</td><td>{_e(limit_text(i.lower, i.upper, i.spec_unit))}</td>"
                             f"{_status_cell(i.status)}<td>{_e(i.reason)}</td></tr>")
            parts.append("</table>")
        if o.saved and o.saved.notice:
            parts.append(f"<p>협력사 통보서 초안(발송 안 됨): {_e(o.saved.notice)}</p>")
        if o.saved and o.saved.result_xlsx and Path(o.saved.result_xlsx).is_file():
            attachments.append(Path(o.saved.result_xlsx))

    parts.append('<p style="color:#666">이 메일은 COA 자동검증기가 생성했습니다. 판정 근거는 첨부 결과 엑셀을 참고하세요. '
                 "협력사에는 자동으로 아무것도 발송되지 않습니다.</p></div>")
    return Alert(subject=subject, html="\n".join(parts), attachments=attachments)


def build_no_pdf_alert(mail_subject: str, sender: str, received: datetime | None) -> Alert:
    """등록 협력사·시연 발신자가 PDF 없이 보낸 COA 메일 알림."""
    body = (
        '<div style="font-family:Malgun Gothic,Arial;font-size:10pt">'
        "<p><b>COA 메일로 보이지만 PDF 첨부가 없습니다.</b> 협력사에 PDF 재송부를 요청할지 확인하세요.</p>"
        f"<p>수신 메일: {_e(mail_subject)}<br>발신자: {_e(sender)}<br>"
        f"수신시각: {_e(f'{received:%Y-%m-%d %H:%M:%S}' if received else '-')}</p>"
        '<p style="color:#666">이 메일은 COA 자동검증기가 생성했습니다.</p></div>'
    )
    return Alert(subject=f"[COA 확인 필요] PDF 첨부 없음 - {sender}", html=body)


def send_alert(app, alert: Alert, config: dict, account=None) -> None:
    """담당자(alert.to)에게 알림 메일. alert.mode가 "send"면 발송, "display"면 창만 띄운다."""
    to = [a for a in config["alert"].get("to", []) if a]
    mode = config["alert"].get("mode", "display")
    mail = app.CreateItem(OL_MAIL_ITEM)
    if account is not None:
        try:
            mail.SendUsingAccount = account
        except Exception:
            log.warning("보내는 계정 지정 실패 - Outlook 기본 계정으로 보냅니다.", exc_info=True)
    mail.To = "; ".join(to)
    mail.Subject = alert.subject
    mail.HTMLBody = alert.html
    for path in alert.attachments:
        mail.Attachments.Add(str(path))
    if mode == "send" and to:
        mail.Send()
        log.info("알림 메일 발송: %s → %s", alert.subject, ", ".join(to))
    else:
        if mode == "send":
            log.warning("alert.to가 비어 있어 발송하지 않고 창만 띄웁니다.")
        mail.Display(False)
        log.info("알림 메일 창 표시(발송 안 함): %s", alert.subject)
