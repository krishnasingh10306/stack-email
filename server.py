import csv
import base64
import html
import hmac
import io
import json
import os
import re
import sqlite3
import threading
import time
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen
from uuid import uuid4

from cryptography.fernet import Fernet, InvalidToken


ROOT = Path(__file__).parent
MAX_UPLOAD = 15 * 1024 * 1024
MAX_REQUEST_BYTES = MAX_UPLOAD * 2 + 1024 * 1024
MAX_RECIPIENTS = 500
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
DATA_DIR = Path(os.environ.get("DATA_DIR", ROOT / "data"))
DATABASE = DATA_DIR / "letterdrop.sqlite3"
_cipher = None


def get_cipher():
    global _cipher
    if _cipher is None:
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        key = os.environ.get("APP_ENCRYPTION_KEY")
        if key:
            _cipher = Fernet(key.encode("ascii"))
        else:
            key_file = DATA_DIR / ".app_key"
            if not key_file.exists():
                key_file.write_bytes(Fernet.generate_key())
                try:
                    key_file.chmod(0o600)
                except OSError:
                    pass
            _cipher = Fernet(key_file.read_bytes().strip())
    return _cipher


def database_connection():
    connection = sqlite3.connect(DATABASE, timeout=30)
    connection.row_factory = sqlite3.Row
    return connection


def initialize_database():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    with database_connection() as connection:
        connection.execute("CREATE TABLE IF NOT EXISTS campaigns (id TEXT PRIMARY KEY, run_at REAL NOT NULL, status TEXT NOT NULL, recipient_count INTEGER NOT NULL, payload BLOB, error TEXT, created_at REAL NOT NULL)")
        connection.execute("UPDATE campaigns SET status = 'failed', payload = NULL, error = 'The server restarted during sending. Review delivery before scheduling again.' WHERE status = 'sending'")
        connection.execute("DELETE FROM campaigns WHERE status IN ('sent', 'partial', 'failed', 'cancelled') AND created_at < ?", (time.time() - 90 * 24 * 60 * 60,))


def validate_campaign(payload):
    contacts = payload.get("contacts", [])
    if not contacts or len(contacts) > MAX_RECIPIENTS:
        raise ValueError(f"Choose between 1 and {MAX_RECIPIENTS} valid contacts.")
    api_key = os.environ.get("RESEND_API_KEY", "").strip()
    sender = str(payload.get("sender", "")).strip()
    subject, body = str(payload.get("subject", "")).strip(), str(payload.get("body", ""))
    if not all((api_key, sender, subject, body)):
        raise ValueError("Configure RESEND_API_KEY on the server, then enter a verified sender email, subject, and message.")
    if not api_key.startswith("re_"):
        raise ValueError("Enter a valid Resend API key.")
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", sender):
        raise ValueError("Enter a valid sender email address.")
    return contacts, api_key, sender, subject, body


def deliver_campaign(payload):
    contacts, api_key, sender, subject, body = validate_campaign(payload)
    results = []
    for contact in contacts:
        recipient = next((v for k, v in contact.items() if k.lower() in ("email", "email address", "work email", "e-mail")), "")
        personalized = re.sub(r"\{\{\s*([^}]+?)\s*\}\}", lambda m: next((v for k, v in contact.items() if k.lower() == m.group(1).lower()), m.group(0)), body)
        message = {
            "from": sender,
            "to": [recipient],
            "subject": subject,
            "text": personalized,
            "html": "<div style=\"white-space: pre-wrap\">" + html.escape(personalized) + "</div>",
        }
        request = Request(
            "https://api.resend.com/emails",
            data=json.dumps(message).encode("utf-8"),
            headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urlopen(request, timeout=30) as response:
                result = json.loads(response.read())
            results.append({"email": recipient, "ok": True, "id": result.get("id")})
        except HTTPError as exc:
            try:
                details = json.loads(exc.read()).get("message", str(exc))
            except (ValueError, AttributeError):
                details = str(exc)
            results.append({"email": recipient, "ok": False, "error": details})
        except (URLError, TimeoutError, OSError) as exc:
            results.append({"email": recipient, "ok": False, "error": f"Resend API request failed: {exc}"})
    return results


def send_scheduled_campaign(campaign_id, payload_blob):
    try:
        payload = json.loads(get_cipher().decrypt(payload_blob))
        results = deliver_campaign(payload)
        sent = sum(1 for result in results if result["ok"])
        status = "sent" if sent == len(results) else "partial" if sent else "failed"
        error = None if status == "sent" else f"{len(results) - sent} recipient(s) could not be delivered."
    except Exception as exc:
        status, error = "failed", str(exc)
    with database_connection() as connection:
        connection.execute("UPDATE campaigns SET status = ?, payload = NULL, error = ? WHERE id = ? AND status = 'sending'", (status, error, campaign_id))


def scheduler_loop():
    while True:
        try:
            with database_connection() as connection:
                connection.execute("BEGIN IMMEDIATE")
                row = connection.execute("SELECT id, payload FROM campaigns WHERE status = 'scheduled' AND run_at <= ? ORDER BY run_at LIMIT 1", (time.time(),)).fetchone()
                if row:
                    connection.execute("UPDATE campaigns SET status = 'sending' WHERE id = ? AND status = 'scheduled'", (row["id"],))
                connection.commit()
            if row:
                send_scheduled_campaign(row["id"], row["payload"])
            else:
                time.sleep(5)
        except Exception as exc:
            print(f"Scheduler error: {exc}")
            time.sleep(5)


def read_xlsx(data):
    with zipfile.ZipFile(io.BytesIO(data)) as book:
        shared = []
        if "xl/sharedStrings.xml" in book.namelist():
            root = ET.fromstring(book.read("xl/sharedStrings.xml"))
            shared = ["".join(node.itertext()) for node in root.findall("m:si", NS)]
        workbook = ET.fromstring(book.read("xl/workbook.xml"))
        sheet = workbook.find("m:sheets/m:sheet", NS)
        if sheet is None:
            raise ValueError("The workbook does not contain a worksheet.")
        rel_id = sheet.attrib[f"{{{NS['r']}}}id"]
        rels = ET.fromstring(book.read("xl/_rels/workbook.xml.rels"))
        target = next(item.attrib["Target"] for item in rels if item.attrib["Id"] == rel_id)
        path = target.lstrip("/") if target.startswith("/") else "xl/" + target
        path = os.path.normpath(path).replace("\\", "/")
        root = ET.fromstring(book.read(path))
        rows = []
        for row in root.findall(".//m:sheetData/m:row", NS):
            values = []
            for cell in row.findall("m:c", NS):
                match = re.match(r"([A-Z]+)", cell.attrib.get("r", "A1"))
                col = 0
                for char in match.group(1):
                    col = col * 26 + ord(char) - 64
                col -= 1
                while len(values) <= col:
                    values.append("")
                value = cell.find("m:v", NS)
                text = "" if value is None else value.text or ""
                if cell.attrib.get("t") == "s" and text:
                    text = shared[int(text)]
                elif cell.attrib.get("t") == "inlineStr":
                    inline = cell.find("m:is", NS)
                    text = "".join(inline.itertext()) if inline is not None else ""
                values[col] = text.strip()
            rows.append(values)
        return rows


def parse_contacts(filename, data):
    if len(data) > MAX_UPLOAD:
        raise ValueError("File is larger than 15 MB.")
    if filename.lower().endswith(".csv"):
        rows = list(csv.reader(io.StringIO(data.decode("utf-8-sig"))))
    elif filename.lower().endswith(".xlsx"):
        rows = read_xlsx(data)
    else:
        raise ValueError("Upload an .xlsx or .csv file.")
    if not rows:
        raise ValueError("The file is empty.")
    headers = [str(value).strip() for value in rows[0]]
    email_idx = next((i for i, h in enumerate(headers) if h.lower() in ("email", "email address", "work email", "e-mail")), None)
    if email_idx is None:
        raise ValueError("Could not find an Email column. Name it Email or Email Address.")
    contacts, seen = [], set()
    for row in rows[1:]:
        record = {header or f"Column {i + 1}": str(row[i]).strip() if i < len(row) else "" for i, header in enumerate(headers)}
        email = record.get(headers[email_idx], "")
        if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email):
            continue
        if email.lower() in seen:
            continue
        seen.add(email.lower())
        contacts.append(record)
    if not contacts:
        raise ValueError("No valid email addresses found under the Email column.")
    if len(contacts) > MAX_RECIPIENTS:
        raise ValueError(f"This local version supports up to {MAX_RECIPIENTS} recipients per campaign.")
    return contacts


class Handler(BaseHTTPRequestHandler):
    def is_authorized(self):
        username = os.environ.get("APP_USERNAME", "")
        password = os.environ.get("APP_PASSWORD", "")
        if not username and not password:
            return os.environ.get("HOST", "127.0.0.1") != "0.0.0.0"
        if not username or not password:
            return False
        authorization = self.headers.get("Authorization", "")
        if not authorization.startswith("Basic "):
            return False
        try:
            provided = base64.b64decode(authorization[6:], validate=True).decode("utf-8")
            given_username, given_password = provided.split(":", 1)
        except (ValueError, UnicodeDecodeError):
            return False
        return hmac.compare_digest(given_username, username) and hmac.compare_digest(given_password, password)

    def require_auth(self):
        if self.path.split("?", 1)[0] == "/healthz":
            return True
        if self.is_authorized():
            return True
        body = b"Authentication required"
        self.send_response(401)
        self.send_header("WWW-Authenticate", 'Basic realm="Letterdrop", charset="UTF-8"')
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)
        return False

    def send_json(self, payload, status=200):
        body = json.dumps(payload).encode()
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/healthz":
            return self.send_json({"status": "ok"})
        if not self.require_auth():
            return
        if path == "/api/schedules":
            with database_connection() as connection:
                rows = connection.execute("SELECT id, run_at, status, recipient_count, error FROM campaigns ORDER BY run_at DESC LIMIT 50").fetchall()
            return self.send_json({"schedules": [{"id": row["id"], "run_at": datetime.fromtimestamp(row["run_at"], timezone.utc).isoformat(), "status": row["status"], "recipient_count": row["recipient_count"], "error": row["error"]} for row in rows]})
        if path == "/":
            body = (ROOT / "index.html").read_bytes()
            mime = "text/html; charset=utf-8"
        elif path in ("/app.js", "/styles.css"):
            body = (ROOT / path[1:]).read_bytes()
            mime = "text/javascript; charset=utf-8" if path.endswith(".js") else "text/css; charset=utf-8"
        else:
            self.send_error(404)
            return
        self.send_response(200)
        self.send_header("Content-Type", mime)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_POST(self):
        if not self.require_auth():
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            if length > MAX_REQUEST_BYTES:
                return self.send_json({"error": "Request too large."}, 413)
            payload = json.loads(self.rfile.read(length))
            if self.path == "/api/contacts":
                raw = __import__("base64").b64decode(payload["data"])
                return self.send_json({"contacts": parse_contacts(payload["filename"], raw)})
            if self.path == "/api/send":
                return self.send_campaign(payload)
            if self.path == "/api/schedule":
                return self.schedule_campaign(payload)
            if self.path.startswith("/api/schedules/") and self.path.endswith("/cancel"):
                campaign_id = self.path.removeprefix("/api/schedules/").removesuffix("/cancel").strip("/")
                with database_connection() as connection:
                    result = connection.execute("UPDATE campaigns SET status = 'cancelled', payload = NULL WHERE id = ? AND status = 'scheduled'", (campaign_id,))
                if result.rowcount != 1:
                    return self.send_json({"error": "This send has already started or is no longer scheduled."}, 409)
                return self.send_json({"ok": True})
            self.send_error(404)
        except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError, UnicodeDecodeError) as exc:
            self.send_json({"error": str(exc)}, 400)

    def send_campaign(self, payload):
        results = []
        try:
            results = deliver_campaign(payload)
        except ValueError as exc:
            return self.send_json({"error": str(exc)}, 400)
        except RuntimeError as exc:
            return self.send_json({"error": str(exc)}, 502)
        self.send_json({"results": results, "sent": sum(1 for result in results if result["ok"])})

    def schedule_campaign(self, payload):
        try:
            contacts, *_ = validate_campaign(payload)
            run_at = datetime.fromisoformat(payload.get("run_at", "").replace("Z", "+00:00"))
            if run_at.tzinfo is None:
                raise ValueError("Choose a date and time with a time zone.")
            run_timestamp = run_at.astimezone(timezone.utc).timestamp()
            if run_timestamp < time.time() + 10:
                raise ValueError("Choose a send time at least 10 seconds in the future.")
            if run_timestamp > time.time() + 365 * 24 * 60 * 60:
                raise ValueError("Schedule sends no more than one year in advance.")
            campaign_id = uuid4().hex
            encrypted_payload = get_cipher().encrypt(json.dumps(payload).encode("utf-8"))
            with database_connection() as connection:
                connection.execute("INSERT INTO campaigns (id, run_at, status, recipient_count, payload, error, created_at) VALUES (?, ?, 'scheduled', ?, ?, NULL, ?)", (campaign_id, run_timestamp, len(contacts), encrypted_payload, time.time()))
            return self.send_json({"id": campaign_id, "status": "scheduled", "run_at": run_at.astimezone(timezone.utc).isoformat(), "recipient_count": len(contacts)}, 201)
        except (ValueError, TypeError) as exc:
            return self.send_json({"error": str(exc)}, 400)

    def log_message(self, format, *args):
        print(f"{self.address_string()} - {format % args}")


if __name__ == "__main__":
    host, port = os.environ.get("HOST", "127.0.0.1"), int(os.environ.get("PORT", "8000"))
    if host == "0.0.0.0" and not all((os.environ.get("APP_USERNAME"), os.environ.get("APP_PASSWORD"))):
        raise SystemExit("Set APP_USERNAME and APP_PASSWORD before exposing Letterdrop publicly.")
    get_cipher()
    initialize_database()
    threading.Thread(target=scheduler_loop, daemon=True, name="letterdrop-scheduler").start()
    print(f"TeamMail is ready at http://{host}:{port}")
    ThreadingHTTPServer((host, port), Handler).serve_forever()
