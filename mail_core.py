from __future__ import annotations

import mimetypes
import smtplib
import webbrowser
from dataclasses import dataclass
from email.message import EmailMessage
from pathlib import Path
from typing import Iterable
from urllib.parse import quote


@dataclass(slots=True)
class SMTPSettings:
    host: str = ""
    port: int = 587
    username: str = ""
    password: str = ""
    sender: str = ""
    default_to: str = ""
    use_ssl: bool = False
    use_starttls: bool = True
    timeout_seconds: int = 12
    save_password: bool = False

    @classmethod
    def from_dict(cls, data: dict[str, object] | None) -> "SMTPSettings":
        payload = data or {}
        port_value = payload.get("port", 587)
        try:
            port = int(port_value) if str(port_value).strip() else 587
        except Exception:
            port = 587
        timeout_value = payload.get("timeout_seconds", 12)
        try:
            timeout_seconds = max(5, int(timeout_value))
        except Exception:
            timeout_seconds = 12
        return cls(
            host=str(payload.get("host", "")).strip(),
            port=port,
            username=str(payload.get("username", "")).strip(),
            password=str(payload.get("password", "")),
            sender=str(payload.get("sender", "")).strip(),
            default_to=str(payload.get("default_to", "")).strip(),
            use_ssl=bool(payload.get("use_ssl", False)),
            use_starttls=bool(payload.get("use_starttls", True)),
            timeout_seconds=timeout_seconds,
            save_password=bool(payload.get("save_password", False)),
        )

    def to_state_dict(self) -> dict[str, object]:
        return {
            "host": self.host,
            "port": int(self.port or 587),
            "username": self.username,
            "password": self.password if self.save_password else "",
            "sender": self.sender,
            "default_to": self.default_to,
            "use_ssl": bool(self.use_ssl),
            "use_starttls": bool(self.use_starttls),
            "timeout_seconds": int(self.timeout_seconds or 12),
            "save_password": bool(self.save_password),
        }

    def sanitized_dict(self) -> dict[str, object]:
        data = self.to_state_dict()
        data["password"] = "***" if self.password else ""
        return data

    def validate_for_send(self) -> None:
        if not self.host:
            raise ValueError("SMTP host is required.")
        if not self.sender:
            raise ValueError("Sender email is required.")
        if not self.port or int(self.port) <= 0:
            raise ValueError("SMTP port must be a positive integer.")
        if self.use_ssl and self.use_starttls:
            raise ValueError("Use SSL or STARTTLS, not both together.")



def _parse_recipients(recipients: str | Iterable[str]) -> list[str]:
    if isinstance(recipients, str):
        raw_items = recipients.replace(";", ",").split(",")
    else:
        raw_items = list(recipients)
    cleaned = [str(item).strip() for item in raw_items if str(item).strip()]
    if not cleaned:
        raise ValueError("At least one recipient email address is required.")
    return cleaned



def build_email_message(
    sender: str,
    recipients: str | Iterable[str],
    subject: str,
    body: str,
    attachments: Iterable[str | Path] = (),
) -> EmailMessage:
    resolved_recipients = _parse_recipients(recipients)
    if not sender.strip():
        raise ValueError("Sender email is required.")
    message = EmailMessage()
    message["From"] = sender.strip()
    message["To"] = ", ".join(resolved_recipients)
    message["Subject"] = subject.strip() or "Message from Gokul Omni Convert Lite"
    message.set_content(body or "")

    for item in attachments:
        path = Path(item).expanduser()
        if not path.exists() or not path.is_file():
            continue
        mime_type, _ = mimetypes.guess_type(str(path))
        if mime_type:
            maintype, subtype = mime_type.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        with path.open("rb") as handle:
            message.add_attachment(handle.read(), maintype=maintype, subtype=subtype, filename=path.name)

    return message



def _connect(settings: SMTPSettings) -> tuple[smtplib.SMTP, str]:
    settings.validate_for_send()
    if settings.use_ssl:
        client: smtplib.SMTP = smtplib.SMTP_SSL(settings.host, int(settings.port), timeout=settings.timeout_seconds)
        negotiated = "SSL"
    else:
        client = smtplib.SMTP(settings.host, int(settings.port), timeout=settings.timeout_seconds)
        negotiated = "PLAIN"
    client.ehlo()
    if settings.use_starttls and not settings.use_ssl:
        client.starttls()
        client.ehlo()
        negotiated = "STARTTLS"
    if settings.username:
        client.login(settings.username, settings.password)
    return client, negotiated



def test_smtp_connection(settings: SMTPSettings) -> str:
    client, negotiated = _connect(settings)
    try:
        return (
            f"Connected to {settings.host}:{settings.port} using {negotiated}. "
            f"Authentication {'used' if settings.username else 'not required'} for this test."
        )
    finally:
        try:
            client.quit()
        except Exception:
            client.close()



def send_email(
    settings: SMTPSettings,
    recipients: str | Iterable[str],
    subject: str,
    body: str,
    attachments: Iterable[str | Path] = (),
) -> list[Path]:
    resolved_recipients = _parse_recipients(recipients)
    message = build_email_message(settings.sender, resolved_recipients, subject, body, attachments)
    client, _negotiated = _connect(settings)
    try:
        client.send_message(message)
    finally:
        try:
            client.quit()
        except Exception:
            client.close()
    delivered_attachments = [Path(item).expanduser() for item in attachments if Path(item).expanduser().exists()]
    return delivered_attachments


def create_mailto_url(
    recipients: str | Iterable[str] = (),
    *,
    subject: str = "",
    body: str = "",
    cc: str | Iterable[str] = (),
) -> str:
    to_values = _parse_recipients(recipients) if recipients else []
    cc_values = _parse_recipients(cc) if cc else []
    query_parts: list[str] = []
    if cc_values:
        query_parts.append(f"cc={quote(', '.join(cc_values))}")
    if subject:
        query_parts.append(f"subject={quote(subject)}")
    if body:
        query_parts.append(f"body={quote(body)}")
    query = "&".join(query_parts)
    to_segment = quote(", ".join(to_values)) if to_values else ""
    return f"mailto:{to_segment}?{query}" if query else f"mailto:{to_segment}"


def open_mailto_draft(
    recipients: str | Iterable[str] = (),
    *,
    subject: str = "",
    body: str = "",
    cc: str | Iterable[str] = (),
) -> str:
    url = create_mailto_url(recipients, subject=subject, body=body, cc=cc)
    webbrowser.open(url)
    return url


def build_eml_draft(
    output_path: str | Path,
    *,
    sender: str,
    recipients: str | Iterable[str],
    subject: str,
    body: str,
    attachments: Iterable[str | Path] = (),
    cc: str | Iterable[str] = (),
) -> Path:
    message = build_email_message(sender, recipients, subject, body, attachments)
    cc_values = _parse_recipients(cc) if cc else []
    if cc_values:
        message["Cc"] = ", ".join(cc_values)
    destination = Path(output_path).expanduser()
    destination.parent.mkdir(parents=True, exist_ok=True)
    destination.write_bytes(message.as_bytes())
    return destination
