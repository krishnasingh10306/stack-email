import csv
import base64
import hmac
import io
import json
import os
import re
import smtplib
import threading
import xml.etree.ElementTree as ET
import zipfile
from email.message import EmailMessage
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


ROOT = Path(__file__).parent
MAX_UPLOAD = 15 * 1024 * 1024
MAX_RECIPIENTS = 500
NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}


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
            if length > MAX_UPLOAD + 1024 * 1024:
                return self.send_json({"error": "Request too large."}, 413)
            payload = json.loads(self.rfile.read(length))
            if self.path == "/api/contacts":
                raw = __import__("base64").b64decode(payload["data"])
                return self.send_json({"contacts": parse_contacts(payload["filename"], raw)})
            if self.path == "/api/send":
                return self.send_campaign(payload)
            self.send_error(404)
        except (ValueError, KeyError, zipfile.BadZipFile, ET.ParseError, UnicodeDecodeError) as exc:
            self.send_json({"error": str(exc)}, 400)

    def send_campaign(self, payload):
        contacts = payload.get("contacts", [])
        if not contacts or len(contacts) > MAX_RECIPIENTS:
            return self.send_json({"error": f"Choose between 1 and {MAX_RECIPIENTS} valid contacts."}, 400)
        host, port = payload.get("host", "").strip(), int(payload.get("port", 587))
        username, password = payload.get("username", ""), payload.get("password", "")
        sender = payload.get("sender", "").strip()
        subject, body = payload.get("subject", "").strip(), payload.get("body", "")
        if not all((host, username, password, sender, subject, body)):
            return self.send_json({"error": "Fill in all SMTP and message fields."}, 400)
        if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", sender):
            return self.send_json({"error": "Enter a valid sender email address."}, 400)
        results = []
        try:
            with smtplib.SMTP(host, port, timeout=30) as smtp:
                smtp.ehlo()
                smtp.starttls()
                smtp.ehlo()
                smtp.login(username, password)
                for contact in contacts:
                    recipient = next((v for k, v in contact.items() if k.lower() in ("email", "email address", "work email", "e-mail")), "")
                    message = EmailMessage()
                    message["From"], message["To"] = sender, recipient
                    message["Subject"] = subject
                    message.set_content(re.sub(r"\{\{\s*([^}]+?)\s*\}\}", lambda m: next((v for k, v in contact.items() if k.lower() == m.group(1).lower()), m.group(0)), body))
                    try:
                        smtp.send_message(message)
                        results.append({"email": recipient, "ok": True})
                    except Exception as exc:
                        results.append({"email": recipient, "ok": False, "error": str(exc)})
        except Exception as exc:
            return self.send_json({"error": f"SMTP connection failed: {exc}"}, 502)
        self.send_json({"results": results, "sent": sum(1 for result in results if result["ok"])})

    def log_message(self, format, *args):
        print(f"{self.address_string()} - {format % args}")


if __name__ == "__main__":
    host, port = os.environ.get("HOST", "127.0.0.1"), int(os.environ.get("PORT", "8000"))
    if host == "0.0.0.0" and not all((os.environ.get("APP_USERNAME"), os.environ.get("APP_PASSWORD"))):
        raise SystemExit("Set APP_USERNAME and APP_PASSWORD before exposing Letterdrop publicly.")
    print(f"TeamMail is ready at http://{host}:{port}")
    ThreadingHTTPServer((host, port), Handler).serve_forever()
