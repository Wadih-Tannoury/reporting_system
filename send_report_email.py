"""
Email the cancellation report
=============================

Sends the workbook produced by kibo_cancellation_report.py as an attachment.

The SMTP handling deliberately mirrors main/main.py of the ERP Alerting System,
so the same repository secrets and variables drive both: SMTP_HOST, SMTP_PORT,
SMTP_STARTTLS, SMTP_USERNAME, SMTP_PASSWORD, EMAIL_FROM, SMTP_AUTH_MECHANISM,
SMTP_REMOVE_PASSWORD_SPACES and EMAIL_DRY_RUN.

    python scripts/send_report_email.py --file output/report.xlsx \
        --window-start 2026-11-04T08:00:00Z --window-end 2026-11-05T08:00:00Z

    python scripts/send_report_email.py --login-test   # auth only, sends nothing
"""

from __future__ import annotations

import argparse
import logging
import mimetypes
import os
import re
import smtplib
import ssl
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

DEFAULT_RECIPIENTS: Sequence[str] = (
    "giulia.silvestro@ladoublej.com",
    "sabrina.baraldi@ladoublej.com",
    "sara.mattioli@ladoublej.com",
    "giulia.piroli@ladoublej.com",
    "beatrice.cabrini@ladoublej.com",
)

DEFAULT_SUBJECT = "DJ daily cancellation report"

XLSX_MIMETYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%H:%M:%S",
    stream=sys.stdout,
)
log = logging.getLogger("send_report_email")


def env_value(name: str, default: str = "") -> str:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    return value.strip()


def env_bool(name: str, default: bool = False) -> bool:
    value = os.getenv(name)
    if value is None or value.strip() == "":
        return default
    return value.strip().lower() in {"1", "true", "yes", "y", "on"}


def env_int(name: str, default: int) -> int:
    value = env_value(name)
    if not value:
        return default
    try:
        return int(value)
    except ValueError as exc:
        raise ValueError(f"Environment variable {name} must be an integer; got {value!r}") from exc


@dataclass(frozen=True)
class Settings:
    smtp_host: str
    smtp_port: int
    smtp_username: str
    smtp_password: str
    smtp_from_email: str
    smtp_use_tls: bool
    smtp_auth_mechanism: str
    smtp_remove_password_spaces: bool
    dry_run: bool


def load_settings() -> Settings:
    """Same variable names and aliases as the ERP Alerting System."""
    smtp_username = env_value("SMTP_USERNAME")
    smtp_from_email = env_value(
        "SMTP_FROM_EMAIL", env_value("EMAIL_FROM", smtp_username)
    )
    return Settings(
        smtp_host=env_value("SMTP_HOST", "smtp.office365.com"),
        smtp_port=env_int("SMTP_PORT", 587),
        smtp_username=smtp_username,
        smtp_password=env_value("SMTP_PASSWORD"),
        smtp_from_email=smtp_from_email,
        smtp_use_tls=env_bool("SMTP_USE_TLS", env_bool("SMTP_STARTTLS", True)),
        smtp_auth_mechanism=env_value("SMTP_AUTH_MECHANISM", "AUTO").upper(),
        smtp_remove_password_spaces=env_bool("SMTP_REMOVE_PASSWORD_SPACES", False),
        dry_run=env_bool("DRY_RUN", env_bool("EMAIL_DRY_RUN", False)),
    )


def recipients_from_env() -> List[str]:
    raw = env_value("EMAIL_RECIPIENTS")
    if not raw:
        return list(DEFAULT_RECIPIENTS)
    return [part.strip() for part in re.split(r"[,;\s]+", raw) if part.strip()]


# --------------------------------------------------------------------------- #
# SMTP, as in the ERP Alerting System
# --------------------------------------------------------------------------- #

def normalize_smtp_password(settings: Settings) -> str:
    password = settings.smtp_password
    if settings.smtp_remove_password_spaces:
        password = re.sub(r"\s+", "", password)
    return password


def smtp_error_text(exc: smtplib.SMTPException) -> str:
    """SMTP errors as text, without echoing credentials."""
    smtp_error = getattr(exc, "smtp_error", b"")
    if isinstance(smtp_error, bytes):
        text = smtp_error.decode("utf-8", errors="replace")
    else:
        text = str(smtp_error)
    smtp_code = getattr(exc, "smtp_code", None)
    if smtp_code:
        return f"{smtp_code} {text}".strip()
    return text.strip() or repr(exc)


def authenticate_smtp(smtp: smtplib.SMTP, settings: Settings) -> None:
    mechanism = settings.smtp_auth_mechanism.upper()
    password = normalize_smtp_password(settings)

    if mechanism in {"", "AUTO"}:
        smtp.login(settings.smtp_username, password)
        return

    if mechanism not in {"LOGIN", "PLAIN"}:
        raise RuntimeError("Unsupported SMTP_AUTH_MECHANISM. Use AUTO, LOGIN, or PLAIN.")

    advertised = smtp.esmtp_features.get("auth", "")
    if mechanism not in advertised.upper().split():
        log.warning(
            "SMTP_AUTH_MECHANISM=%s was requested, but server advertised AUTH mechanisms: %s",
            mechanism,
            advertised or "none",
        )

    smtp.user = settings.smtp_username
    smtp.password = password
    if mechanism == "LOGIN":
        smtp.auth("LOGIN", smtp.auth_login, initial_response_ok=False)
    else:
        smtp.auth("PLAIN", smtp.auth_plain)


def raise_smtp_auth_error(exc: smtplib.SMTPAuthenticationError, settings: Settings) -> None:
    detail = smtp_error_text(exc)
    raise RuntimeError(
        "SMTP authentication failed for user "
        f"{settings.smtp_username!r} against {settings.smtp_host}:{settings.smtp_port}. "
        f"Microsoft returned: {detail}. For Microsoft 365, verify: "
        "1) the GitHub secret SMTP_PASSWORD is the real mailbox password or an app password, not placeholder text; "
        "2) Authenticated SMTP is enabled for this mailbox; "
        "3) Security Defaults, Conditional Access, and the Exchange authentication policy are not blocking SMTP AUTH; "
        "4) if MFA is enabled, use an app password if the tenant allows them. "
        "If the app password was copied with spaces, set SMTP_REMOVE_PASSWORD_SPACES=true."
    ) from exc


def validate_smtp_settings(settings: Settings) -> None:
    if not settings.smtp_host:
        raise RuntimeError("Missing SMTP_HOST environment variable")
    if not settings.smtp_username:
        raise RuntimeError("Missing SMTP_USERNAME secret")
    if not settings.smtp_password:
        raise RuntimeError("Missing SMTP_PASSWORD secret")
    stripped = settings.smtp_password.strip()
    if stripped.startswith("<") and stripped.endswith(">"):
        raise RuntimeError(
            "SMTP_PASSWORD still looks like placeholder text. Store only the actual "
            "password or app password as the GitHub secret value."
        )


def open_smtp_session(settings: Settings) -> smtplib.SMTP:
    validate_smtp_settings(settings)
    context = ssl.create_default_context()
    smtp = smtplib.SMTP(settings.smtp_host, settings.smtp_port, timeout=60)
    try:
        smtp.ehlo()
        if settings.smtp_use_tls:
            smtp.starttls(context=context)
            smtp.ehlo()
        log.info(
            "Connected to SMTP host=%s port=%s starttls=%s advertised_auth=%s",
            settings.smtp_host,
            settings.smtp_port,
            settings.smtp_use_tls,
            smtp.esmtp_features.get("auth", "none"),
        )
        authenticate_smtp(smtp, settings)
        return smtp
    except smtplib.SMTPAuthenticationError as exc:
        try:
            smtp.quit()
        except Exception:
            smtp.close()
        raise_smtp_auth_error(exc, settings)
    except Exception:
        try:
            smtp.quit()
        except Exception:
            smtp.close()
        raise


# --------------------------------------------------------------------------- #
# Message
# --------------------------------------------------------------------------- #

def sheet_counts(path: Path) -> Optional[dict]:
    """Row counts per sheet, for the body. None if the file cannot be read."""
    try:
        import pandas as pd

        with pd.ExcelFile(path) as book:
            return {name: len(pd.read_excel(book, sheet_name=name)) for name in book.sheet_names}
    except Exception as exc:  # a readable body matters less than sending the file
        log.warning("Could not count rows in %s: %s", path, exc)
        return None


def build_body(path: Path, window_start: str, window_end: str) -> str:
    counts = sheet_counts(path)
    lines = ["Daily cancellation report for tenant 30018, attached as an Excel file.", ""]

    if window_start and window_end:
        lines.append(f"Window (UTC): {window_start} to {window_end}")

    if counts is not None:
        total = counts.get("cancellations", 0)
        if total:
            lines.append(
                f"Canceled items: {total} "
                f"({counts.get('full', 0)} full, {counts.get('partial', 0)} partial)"
            )
        else:
            lines.append("No cancellations in this window. The workbook is empty.")

    lines += [
        "",
        "Sheets: cancellations (every canceled item), full (shipments canceled "
        "outright), partial (shipments where only some lines were canceled).",
        "",
        "Sent automatically by the Kibo cancellation report workflow.",
    ]
    return "\n".join(lines) + "\n"


def send_report(
    settings: Settings,
    recipients: Sequence[str],
    subject: str,
    path: Path,
    window_start: str,
    window_end: str,
) -> None:
    from email.message import EmailMessage

    validate_smtp_settings(settings)
    if not recipients:
        raise RuntimeError("Cannot send the report with an empty recipient list")
    if not path.exists():
        raise RuntimeError(f"Attachment not found: {path}")

    payload = path.read_bytes()
    mimetype, _ = mimetypes.guess_type(path.name)
    maintype, _, subtype = (mimetype or XLSX_MIMETYPE).partition("/")

    msg = EmailMessage()
    msg["From"] = settings.smtp_from_email
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = subject
    msg.set_content(build_body(path, window_start, window_end))
    msg.add_attachment(payload, maintype=maintype, subtype=subtype, filename=path.name)

    if settings.dry_run:
        log.info(
            "DRY_RUN=true: would send subject=%r from=%s to=%s attachment=%s bytes=%d",
            subject,
            settings.smtp_from_email,
            list(recipients),
            path.name,
            len(payload),
        )
        return

    with open_smtp_session(settings) as smtp:
        smtp.send_message(msg)

    log.info(
        "Sent %r from %s to %s with %s (%.1f KB)",
        subject,
        settings.smtp_from_email,
        ", ".join(recipients),
        path.name,
        len(payload) / 1024,
    )


def login_test(settings: Settings) -> None:
    if settings.dry_run:
        log.info("DRY_RUN=true: skipping the SMTP login test")
        return
    with open_smtp_session(settings) as smtp:
        smtp.noop()
    log.info("SMTP login test succeeded for %s", settings.smtp_username)


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Email the cancellation report.")
    parser.add_argument("--file", type=Path, default=None, help="The .xlsx to attach.")
    parser.add_argument("--subject", default=env_value("EMAIL_SUBJECT", DEFAULT_SUBJECT))
    parser.add_argument("--to", default=None,
                        help="Recipients, comma separated. Defaults to EMAIL_RECIPIENTS "
                             "or the built-in list.")
    parser.add_argument("--window-start", default="", help="Shown in the body.")
    parser.add_argument("--window-end", default="", help="Shown in the body.")
    parser.add_argument("--login-test", action="store_true",
                        help="Connect and authenticate only, then exit.")
    return parser.parse_args(argv)


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    settings = load_settings()

    if args.login_test:
        login_test(settings)
        return 0

    if args.file is None:
        raise SystemExit("--file is required unless --login-test is used.")

    recipients = (
        [part.strip() for part in re.split(r"[,;\s]+", args.to) if part.strip()]
        if args.to
        else recipients_from_env()
    )

    send_report(
        settings,
        recipients,
        args.subject,
        args.file,
        args.window_start,
        args.window_end,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
