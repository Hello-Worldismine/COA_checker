"""core.outlook·mail_watcher 테스트: Outlook 없이 가짜 메일 객체로 선별·파일명·처리 흐름 확인."""

import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import pytest

import mail_watcher
from core import outlook
from core.config import load_config
from core.models import Status
from core.pipeline import process_pdf, summary_line

SAMPLES_DIR = load_config()["paths"]["samples"]
RECEIVED = datetime(2026, 10, 23, 9, 30, 15, tzinfo=timezone.utc)  # pywin32처럼 로컬 시각에 UTC 표시


class FakeAttachment:
    """Outlook Attachment 흉내."""

    def __init__(self, file_name, source=None, att_type=outlook.OL_BY_VALUE):
        self.FileName, self.Type, self.source = file_name, att_type, source

    def SaveAsFile(self, path):
        Path(path).write_bytes(Path(self.source).read_bytes() if self.source else b"%PDF-1.4 fake")


class FakeAttachments:
    """Outlook Attachments 컬렉션(1부터 시작) 흉내."""

    def __init__(self, items):
        self.items = items
        self.Count = len(items)

    def Item(self, i):
        return self.items[i - 1]


class FakeExchangeUser:
    PrimarySmtpAddress = "kim@company.co.kr"


class FakeSender:
    def GetExchangeUser(self):
        return FakeExchangeUser()


class FakeMail:
    """Outlook MailItem 흉내."""

    def __init__(self, subject="COA N2418", sender="jaemin9905@naver.com", attachments=(), sender_type="SMTP",
                 cls=outlook.OL_MAIL, entry_id="ENTRY-1", categories=""):
        self.Class, self.Subject, self.EntryID = cls, subject, entry_id
        self.SenderEmailType = sender_type
        self.SenderEmailAddress = sender if sender_type != "EX" else "/O=COMPANY/OU=EXCHANGE/CN=RECIPIENTS/CN=KIM"
        self.Sender = FakeSender() if sender_type == "EX" else None
        self.Attachments = FakeAttachments(list(attachments))
        self.ReceivedTime = RECEIVED
        self.Categories = categories
        self.saved = 0

    def Save(self):
        self.saved += 1


def pdf(name="COA_Nippon_N2418_fail.pdf"):
    """샘플 PDF를 내용으로 갖는 가짜 첨부."""
    return FakeAttachment(name, SAMPLES_DIR / name)


@pytest.fixture
def cfg(tmp_path):
    """임시 경로를 쓰는 설정."""
    config = load_config()
    config["paths"] = {**config["paths"], "inbox": tmp_path / "inbox", "output": tmp_path / "output",
                       "db": tmp_path / "output" / "t.db"}
    config["mail"] = {**config["mail"], "test_senders": ["jaemin9905@naver.com"], "allowed_domains": ["partner.example"]}
    return config


# ── 선별 ──────────────────────────────────────────────────────
@pytest.mark.parametrize("subject, sender, attachments, expected", [
    ("COA N2418", "someone@gmail.com", [pdf()], True),             # 제목 키워드 + PDF
    ("성적서 송부", "qc@alpha-chem.example", [pdf()], True),        # 협력사 도메인 + PDF
    ("자료", "qa@mail.nippon-solv.example", [pdf()], True),         # 협력사 하위 도메인
    ("자료", "x@partner.example", [pdf()], True),                   # 추가 허용 도메인
    ("자료", "JAEMIN9905@naver.com", [pdf()], True),                # 시연 발신자(대소문자 무시)
    ("coa 송부", "someone@gmail.com", [pdf()], True),               # 제목 대소문자 무시
    ("자료", "someone@gmail.com", [pdf()], False),                  # 키워드·발신자 모두 아님
    ("COA N2418", "qc@alpha-chem.example", [], False),              # PDF 없음
    ("COA", "qc@alpha-chem.example", [FakeAttachment("a.xlsx")], False),
    ("COA", "qc@alpha-chem.example", [FakeAttachment("logo.pdf", att_type=6)], False),  # 본문 삽입 개체
    ("자료", "evil-alpha-chem.example@x.com", [pdf()], False),
    ("[COA 부적합] Nippon Solvent / N2418", "adm1nsubin@outlook.kr", [pdf()], False),   # 우리가 보낸 알림
    ("[COA 검토 필요] Nippon Solvent / N2419", "jaemin9905@naver.com", [pdf()], False),
    ("Fwd: [COA 부적합] N2418", "jaemin9905@naver.com", [pdf()], True),               # 전달 메일은 대상
])
def test_is_coa_mail(spec_master, cfg, subject, sender, attachments, expected):
    """(제목 키워드 OR 협력사/허용 도메인/시연 발신자) AND PDF 첨부."""
    assert outlook.is_coa_mail(FakeMail(subject, sender, attachments), spec_master, cfg) is expected


def test_non_mail_item_ignored(spec_master, cfg):
    """MailItem(Class 43)이 아니면(회의 요청 등) COA 메일이 아니다."""
    mail = FakeMail(attachments=[pdf()], cls=53)
    assert outlook.check_mail(mail, spec_master, cfg) is None
    assert not outlook.is_coa_mail(mail, spec_master, cfg)


def test_exchange_sender_smtp(spec_master, cfg):
    """사내(Exchange) 발신자는 '/O=...' 대신 SMTP 주소로 판별한다."""
    mail = FakeMail("자료", sender_type="EX", attachments=[pdf()])
    assert outlook.get_smtp_address(mail) == "kim@company.co.kr"
    cfg["mail"]["test_senders"] = ["kim@company.co.kr"]
    assert outlook.is_coa_mail(mail, spec_master, cfg)


# ── 파일명 ────────────────────────────────────────────────────
def test_attachment_filename():
    """표준 파일명, 금지문자 정리, 동명 첨부 번호."""
    when = datetime(2026, 10, 23, 9, 30, 15)
    assert outlook.attachment_filename(when, "Nippon_Solvent_Co_Ltd", "COA_N2418.pdf") == \
        "20261023_093015_Nippon_Solvent_Co_Ltd_COA_N2418.pdf"
    assert outlook.attachment_filename(when, "X", 'a/b:c*?"<>|.PDF') == "20261023_093015_X_a_b_c.pdf"
    assert outlook.attachment_filename(when, "X", "COA.pdf", 2) == "20261023_093015_X_COA_2.pdf"
    assert outlook.attachment_filename(when, "X", "  .pdf") == "20261023_093015_X_COA.pdf"
    assert len(outlook.attachment_filename(when, "X", "a" * 300 + ".pdf")) < 120


def test_sender_label(spec_master):
    """협력사 도메인이면 협력사명, 아니면 메일 주소를 파일명용으로."""
    assert outlook.sender_label("qc@alpha-chem.example", spec_master) == "Alpha_Chemical_Co_Ltd"
    assert outlook.sender_label("jaemin9905@naver.com", spec_master) == "jaemin9905_naver_com"


def test_save_pdf_attachments(tmp_path):
    """PDF만 inbox/YYYYMMDD/에 저장, 동명 첨부 구분, 이미 처리된 첨부는 건너뜀."""
    mail = FakeMail(attachments=[pdf(), FakeAttachment("spec.xlsx"), FakeAttachment("COA_Nippon_N2418_fail.pdf"),
                                 pdf("COA_Alpha_A2401_pass.pdf")])
    saved = outlook.save_pdf_attachments(mail, tmp_path, "X", skip=lambda key: key == "COA_Alpha_A2401_pass.pdf")
    assert [k for k, _ in saved] == ["COA_Nippon_N2418_fail.pdf", "COA_Nippon_N2418_fail.pdf#2"]
    assert [p.name for _, p in saved] == ["20261023_093015_X_COA_Nippon_N2418_fail.pdf",
                                         "20261023_093015_X_COA_Nippon_N2418_fail_2.pdf"]
    assert all(p.parent == tmp_path / "20261023" and p.is_file() for _, p in saved)


def test_mark_processed():
    """처리 완료 분류를 기존 분류 뒤에 추가하고 저장, 이미 있으면 그대로."""
    mail = FakeMail(categories="중요")
    outlook.mark_processed(mail, "COA검증완료")
    assert mail.Categories == "중요, COA검증완료" and mail.saved == 1
    assert outlook.is_marked(mail, "COA검증완료")
    outlook.mark_processed(mail, "COA검증완료")
    assert mail.saved == 1


# ── 알림 본문 ─────────────────────────────────────────────────
def _outputs(spec_master, names, tmp_path):
    paths = {"inbox": tmp_path / "inbox", "output": tmp_path / "output", "db": tmp_path / "o.db"}
    return [(n, process_pdf(SAMPLES_DIR / n, {"source": "upload"}, spec=spec_master, paths=paths)) for n in names]


def test_build_alert_summary(spec_master, tmp_path):
    """메일 1통 → 요약 알림 1통: 제목, 문서별 판정, FAIL·REVIEW 항목 상세, 결과 엑셀 첨부."""
    names = ["COA_Alpha_A2401_pass.pdf", "COA_Nippon_N2418_fail.pdf", "COA_Nippon_N2419_review.pdf"]
    alert = outlook.build_alert(_outputs(spec_master, names, tmp_path), [], "COA 3건", "jaemin9905@naver.com",
                                datetime(2026, 10, 23, 9, 30))
    assert alert.subject == "[COA 부적합] Nippon Solvent / N2418 외 1건"
    assert "부적합 1건, 검토 필요 1건, 적합 1건" in alert.html
    assert "50 ppm → 0.005 % 환산, 상한 0.003 % 초과" in alert.html
    assert "N.D. 표기 - 검출한계 미기재로 검토 필요" in alert.html
    assert "#FFC7CE" in alert.html and "#FFEB9C" in alert.html
    assert len(alert.attachments) == 2 and all(p.suffix == ".xlsx" for p in alert.attachments)


def test_build_alert_review_only_and_failures(spec_master, tmp_path):
    """REVIEW만 있으면 [COA 검토 필요], 처리 실패 첨부도 표에 남는다."""
    alert = outlook.build_alert(_outputs(spec_master, ["COA_Nippon_N2419_review.pdf"], tmp_path), ["bad.pdf"],
                                "COA", "a@b.c", None)
    assert alert.subject == "[COA 검토 필요] Nippon Solvent / N2419 외 1건"
    assert "bad.pdf" in alert.html and "처리 실패" in alert.html


def test_build_alert_all_pass_is_none(spec_master, tmp_path):
    """모두 PASS면 알림 없음."""
    assert outlook.build_alert(_outputs(spec_master, ["COA_Alpha_A2401_pass.pdf"], tmp_path), [], "COA", "a@b.c",
                               None) is None


def test_summary_line(spec_master):
    """콘솔 한 줄 상태."""
    def line(name):
        return summary_line(process_pdf(SAMPLES_DIR / name, spec=spec_master, save=False).result)

    assert line("COA_Nippon_N2418_fail.pdf") == \
        "COA 수신: Nippon Solvent / N2418 → 부적합 (Moisture 50 ppm = 0.005 % > 0.003 %)"
    assert line("COA_Alpha_A2402_fail.pdf") == "COA 수신: Alpha Chemical / A2402 → 부적합 (Iron 0.008 % > 0.005 %)"
    assert line("COA_Nippon_N2419_review.pdf").endswith("→ 검토 필요 (Iron: N.D. 표기 - 검출한계 미기재로 검토 필요)")
    assert line("COA_Alpha_A2401_pass.pdf").endswith("→ 적합 (규격 항목 모두 적합)")


# ── 메일 1통 처리 흐름 (mail_watcher.handle_mail) ─────────────────
def _ctx(cfg, spec_master, alerts, process=process_pdf):
    def run(path, meta, spec):
        return process(path, meta, spec=spec, paths=cfg["paths"])

    return mail_watcher.Context(cfg, alert=alerts.append, process=run, load_spec=lambda: spec_master)


def _docs(cfg):
    with sqlite3.connect(cfg["paths"]["db"]) as conn:
        return conn.execute("SELECT COUNT(*) FROM documents").fetchone()[0]


def test_handle_mail_five_attachments(cfg, spec_master):
    """시연 조건: 메일 1통·PDF 5장 → 5건 저장, 요약 알림 1통, 처리 완료 분류. 다시 와도 재처리 안 함."""
    names = sorted(p.name for p in SAMPLES_DIR.glob("*.pdf"))
    mail = FakeMail(attachments=[pdf(n) for n in names])
    alerts = []
    ctx = _ctx(cfg, spec_master, alerts)

    assert mail_watcher.handle_mail(mail, ctx) == "done"
    assert _docs(cfg) == 5
    assert len(alerts) == 1 and alerts[0].subject.startswith("[COA 부적합]")
    assert mail.Categories == "COA검증완료"
    assert len(list((cfg["paths"]["inbox"] / "20261023").glob("*.pdf"))) == 5

    assert mail_watcher.handle_mail(mail, ctx) == "skipped"          # 분류가 붙어 있음
    mail.Categories = ""                                              # 분류가 지워져도
    assert mail_watcher.handle_mail(mail, ctx) == "done"              # DB 기준으로 첨부를 모두 건너뜀
    assert _docs(cfg) == 5 and len(alerts) == 1


def test_handle_mail_partial_failure_retries_rest(cfg, spec_master):
    """첨부 하나가 실패하면 나머지는 처리하고 분류를 달지 않는다. 다음 실행 때 실패한 것만 다시 처리."""
    names = ["COA_Alpha_A2401_pass.pdf", "COA_Nippon_N2418_fail.pdf"]
    mail = FakeMail(attachments=[pdf(n) for n in names])
    alerts = []

    def flaky(path, meta, spec, paths):
        if "N2418" in meta["attachment_name"]:
            raise RuntimeError("일시 오류")
        return process_pdf(path, meta, spec=spec, paths=paths)

    assert mail_watcher.handle_mail(mail, _ctx(cfg, spec_master, alerts, flaky)) == "partial"
    assert _docs(cfg) == 1 and mail.Categories == ""
    assert "처리 실패" in alerts[0].html

    assert mail_watcher.handle_mail(mail, _ctx(cfg, spec_master, alerts)) == "done"
    assert _docs(cfg) == 2 and mail.Categories == "COA검증완료"
    assert alerts[-1].subject == "[COA 부적합] Nippon Solvent / N2418"


def test_handle_mail_no_pdf(cfg, spec_master):
    """PDF 없는 COA 메일: 등록 발신자면 알림, 키워드만이면 로그만. 둘 다 처리 완료 분류."""
    alerts = []
    known = FakeMail("COA 송부", "jaemin9905@naver.com")
    assert mail_watcher.handle_mail(known, _ctx(cfg, spec_master, alerts)) == "no_pdf"
    assert alerts[0].subject == "[COA 확인 필요] PDF 첨부 없음 - jaemin9905@naver.com"
    assert known.Categories == "COA검증완료"

    reply = FakeMail("RE: COA 문의", "colleague@company.co.kr")
    assert mail_watcher.handle_mail(reply, _ctx(cfg, spec_master, alerts)) == "no_pdf"
    assert len(alerts) == 1 and reply.Categories == "COA검증완료"


def test_own_alert_mail_not_reprocessed(cfg, spec_master):
    """알림 수신 주소가 감시 계정이어도, 우리가 보낸 알림('[COA '로 시작)은 다시 감지하지 않는다."""
    alerts = []
    for subject in ("[COA 부적합] Nippon Solvent / N2418", "[COA 확인 필요] PDF 첨부 없음 - jaemin9905@naver.com"):
        mail = FakeMail(subject, "adm1nsubin@outlook.kr")
        assert mail_watcher.handle_mail(mail, _ctx(cfg, spec_master, alerts)) == "ignored"
        assert mail.Categories == ""
    assert alerts == []


def test_handle_mail_ignores_other_mail(cfg, spec_master):
    """COA와 무관한 메일은 손대지 않는다(분류도 안 붙임)."""
    mail = FakeMail("점심 메뉴", "friend@gmail.com", [pdf()])
    assert mail_watcher.handle_mail(mail, _ctx(cfg, spec_master, [])) == "ignored"
    assert mail.Categories == "" and mail.saved == 0


def test_safe_handle_survives_errors(cfg, spec_master):
    """처리 중 예외가 나도 멈추지 않고, 같은 메일은 정해진 횟수만 재시도한다."""
    def broken(path, meta, spec, paths):
        raise RuntimeError("고장")

    ctx = _ctx(cfg, spec_master, [], broken)
    mail = FakeMail(attachments=[pdf()])
    results = [mail_watcher.safe_handle(mail, ctx) for _ in range(mail_watcher.MAX_RETRIES + 1)]
    assert results == ["partial"] * mail_watcher.MAX_RETRIES + ["gave_up"]
    assert mail.Categories == ""


def test_received_at_drops_timezone():
    """pywin32 수신시각(로컬 시각+UTC 표시)을 시간대 없는 로컬 시각으로."""
    assert outlook.received_at(FakeMail()) == datetime(2026, 10, 23, 9, 30, 15)


def test_folder_candidates_skip_underscore(tmp_path):
    """폴더 모드: '_'로 시작하는 폴더(_smoke_test 등)는 감시 대상에서 뺀다."""
    (tmp_path / "20261023").mkdir()
    (tmp_path / "_smoke_test").mkdir()
    (tmp_path / "20261023" / "a.pdf").write_bytes(b"x")
    (tmp_path / "_smoke_test" / "b.pdf").write_bytes(b"x")
    found = mail_watcher._folder_candidates(tmp_path, datetime(2000, 1, 1))
    assert [p.name for p in found] == ["a.pdf"]


def test_status_label_used():
    """Status 값과 한글 라벨."""
    assert Status.FAIL.value == "FAIL"
