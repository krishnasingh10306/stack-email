"""SMTP settings storage and Windows-user-scoped password protection."""

from __future__ import annotations

import base64
import ctypes
import json
import os
import ssl
from contextlib import contextmanager
from ctypes import wintypes
from email.message import EmailMessage
from pathlib import Path
from smtplib import SMTP, SMTP_SSL

from cryptography.fernet import Fernet, InvalidToken


BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("APP_DATA_DIR", str(BASE_DIR))).resolve()
SETTINGS_PATH = DATA_DIR / "smtp_settings.json"


class _DataBlob(ctypes.Structure):
    _fields_ = [("cbData", wintypes.DWORD), ("pbData", ctypes.POINTER(ctypes.c_byte))]


def _windows_crypto():
    if os.name != "nt":
        raise RuntimeError("SMTP passwords are protected with Windows encryption; run this dashboard on Windows.")
    crypt32 = ctypes.WinDLL("Crypt32.dll", use_last_error=True)
    kernel32 = ctypes.WinDLL("Kernel32.dll", use_last_error=True)
    kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    kernel32.LocalFree.restype = ctypes.c_void_p
    return crypt32, kernel32


def _blob(data: bytes):
    buffer = ctypes.create_string_buffer(data)
    blob = _DataBlob(len(data), ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)))
    return blob, buffer


def _protect(value: str) -> str:
    encryption_key = os.environ.get("SMTP_ENCRYPTION_KEY", "").strip()
    if encryption_key:
        return "fernet:" + Fernet(encryption_key.encode("ascii")).encrypt(value.encode("utf-8")).decode("ascii")
    raw = value.encode("utf-8")
    source, source_buffer = _blob(raw)
    destination = _DataBlob()
    crypt32, kernel32 = _windows_crypto()
    protect = crypt32.CryptProtectData
    protect.argtypes = [
        ctypes.POINTER(_DataBlob), wintypes.LPCWSTR, ctypes.POINTER(_DataBlob),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_DataBlob),
    ]
    protect.restype = wintypes.BOOL
    if not protect(ctypes.byref(source), "Training Reminder SMTP", None, None, None, 1, ctypes.byref(destination)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        encrypted = ctypes.string_at(destination.pbData, destination.cbData)
        return base64.b64encode(encrypted).decode("ascii")
    finally:
        kernel32.LocalFree(destination.pbData)


def _unprotect(value: str) -> str:
    if value.startswith("fernet:"):
        encryption_key = os.environ.get("SMTP_ENCRYPTION_KEY", "").strip()
        if not encryption_key:
            raise RuntimeError("SMTP_ENCRYPTION_KEY is required to decrypt the saved SMTP password.")
        try:
            return Fernet(encryption_key.encode("ascii")).decrypt(value.removeprefix("fernet:").encode("ascii")).decode("utf-8")
        except (InvalidToken, ValueError) as error:
            raise RuntimeError("The saved SMTP password cannot be decrypted. Check SMTP_ENCRYPTION_KEY.") from error
    encrypted = base64.b64decode(value.encode("ascii"))
    source, source_buffer = _blob(encrypted)
    destination = _DataBlob()
    crypt32, kernel32 = _windows_crypto()
    unprotect = crypt32.CryptUnprotectData
    unprotect.argtypes = [
        ctypes.POINTER(_DataBlob), ctypes.POINTER(wintypes.LPWSTR), ctypes.POINTER(_DataBlob),
        ctypes.c_void_p, ctypes.c_void_p, wintypes.DWORD, ctypes.POINTER(_DataBlob),
    ]
    unprotect.restype = wintypes.BOOL
    if not unprotect(ctypes.byref(source), None, None, None, None, 1, ctypes.byref(destination)):
        raise ctypes.WinError(ctypes.get_last_error())
    try:
        return ctypes.string_at(destination.pbData, destination.cbData).decode("utf-8")
    finally:
        kernel32.LocalFree(destination.pbData)


def load_settings(settings_path: Path | None = None) -> dict:
    path = Path(settings_path) if settings_path is not None else SETTINGS_PATH
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    password = _unprotect(data["password_protected"]) if data.get("password_protected") else ""
    return {**data, "password": password}


def public_settings(settings_path: Path | None = None) -> dict:
    path = Path(settings_path) if settings_path is not None else SETTINGS_PATH
    if not path.exists():
        return {"configured": False, "passwordSet": False}
    data = json.loads(path.read_text(encoding="utf-8"))
    return {
        "configured": bool(data.get("host") and data.get("from_email")),
        "host": data.get("host", ""),
        "port": data.get("port", 587),
        "security": data.get("security", "STARTTLS"),
        "username": data.get("username", ""),
        "fromEmail": data.get("from_email", ""),
        "passwordSet": bool(data.get("password_protected")),
    }


def save_settings(payload: dict, settings_path: Path | None = None) -> None:
    path = Path(settings_path) if settings_path is not None else SETTINGS_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    host = str(payload.get("host", "")).strip()
    from_email = str(payload.get("fromEmail", "")).strip()
    username = str(payload.get("username", "")).strip()
    security = str(payload.get("security", "STARTTLS")).upper().strip()
    try:
        port = int(payload.get("port", 587))
    except (TypeError, ValueError) as error:
        raise ValueError("SMTP port must be a number.") from error
    if not host or len(host) > 255:
        raise ValueError("Enter your SMTP server host name.")
    if not (1 <= port <= 65535):
        raise ValueError("SMTP port must be between 1 and 65535.")
    if security not in {"STARTTLS", "SSL"}:
        raise ValueError("Choose STARTTLS or SSL/TLS for the connection security.")
    if "@" not in from_email or any(char.isspace() for char in from_email):
        raise ValueError("Enter the sender email address.")

    current = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    password = str(payload.get("password", ""))
    protected_password = _protect(password) if password else current.get("password_protected", "")
    if username and not protected_password:
        raise ValueError("Enter the SMTP password or app password for this username.")
    data = {
        "host": host,
        "port": port,
        "security": security,
        "username": username,
        "from_email": from_email,
        "password_protected": protected_password,
    }
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(data, indent=2), encoding="utf-8")
    temporary.replace(SETTINGS_PATH)


@contextmanager
def smtp_client(settings: dict | None = None, settings_path: Path | None = None):
    settings = settings or load_settings(settings_path)
    host = str(settings.get("host", "")).strip()
    if not host:
        raise ValueError("Set up SMTP in the dashboard before sending reminders.")
    port = int(settings.get("port", 587))
    security = str(settings.get("security", "STARTTLS")).upper()
    username = str(settings.get("username", "")).strip()
    password = str(settings.get("password", ""))
    if username and not password:
        raise ValueError("SMTP password is missing. Save your SMTP settings again.")

    context = ssl.create_default_context()
    if security == "SSL":
        server = SMTP_SSL(host, port, timeout=30, context=context)
    else:
        server = SMTP(host, port, timeout=30)
        server.ehlo()
        if security == "STARTTLS":
            server.starttls(context=context)
            server.ehlo()
    try:
        if username:
            server.login(username, password)
        yield server
    finally:
        try:
            server.quit()
        except Exception:
            server.close()


def send_message(message: EmailMessage, settings: dict | None = None, settings_path: Path | None = None) -> None:
    settings = settings or load_settings(settings_path)
    from_email = str(settings.get("from_email", "")).strip()
    if not from_email:
        raise ValueError("Set the sender email address in SMTP settings.")
    if "From" not in message:
        message["From"] = from_email
    with smtp_client(settings, settings_path) as server:
        server.send_message(message, from_addr=from_email, to_addrs=[message["To"]])


def check_connection(settings_path: Path | None = None) -> str:
    settings = load_settings(settings_path)
    with smtp_client(settings, settings_path):
        return f"Connected to {settings['host']}:{settings['port']} using {settings['security']}. No email was sent."
