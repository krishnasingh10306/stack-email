"""Serve a small, local-only dashboard for training_reminders.xlsx."""

from __future__ import annotations

import json
import io
import base64
import hmac
import os
import re
import subprocess
import sys
import threading
import time
from contextlib import redirect_stderr, redirect_stdout
from datetime import date, datetime
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.request import urlopen
from urllib.parse import unquote, urlparse
import webbrowser

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Side

import reminder_automation as automation
import smtp_settings


BASE_DIR = Path(__file__).resolve().parent
PAGE_PATH = BASE_DIR / "web" / "index.html"
CLOUD_MODE = os.environ.get("APP_ENV", "").casefold() == "cloud" or bool(os.environ.get("RENDER"))
DATA_DIR = Path(os.environ.get("APP_DATA_DIR", str(BASE_DIR))).resolve()
SCHEDULER_PATH = DATA_DIR / "auto_send_enabled.json"
HOST = os.environ.get("HOST", "0.0.0.0" if CLOUD_MODE else "127.0.0.1")
PORT = int(os.environ.get("PORT", "10000" if CLOUD_MODE else "8765"))
AUTOMATION_LOCK = threading.Lock()


class LocalDashboardServer(ThreadingHTTPServer):
    allow_reuse_address = False

SOURCE_FIELDS = {
    "serial number": "serialNo",
    "department": "department",
    "training topic": "topic",
    "plan status": "planStatus",
    "actual status": "actualStatus",
    "replan status": "replanStatus",
    "training time": "trainingTime",
    "trainer name": "trainer",
}
SOURCE_HEADERS = {
    "serial number": "Serial No.",
    "department": "Department",
    "training topic": "Training Topic",
    "plan status": "Plan Status",
    "actual status": "Actual Status",
    "replan status": "Replan Status",
    "training time": "Training Time",
    "trainer name": "Trainer Name",
}


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(timespec="minutes")
    if isinstance(value, date):
        return value.isoformat()
    return str(value)


def workbook_rows() -> tuple[list[dict], dict]:
    if not automation.WORKBOOK_PATH.exists():
        raise FileNotFoundError("training_reminders.xlsx is missing from this folder.")
    workbook = load_workbook(automation.WORKBOOK_PATH)
    sheet = workbook.active
    header_row, headers = automation.find_header_row(sheet)
    bind_source_headers(sheet, header_row, headers)
    now = datetime.now(automation.TIME_ZONE)
    rows = []
    for row_number in range(header_row + 1, sheet.max_row + 1):
        values = [sheet.cell(row_number, col).value for col in range(1, sheet.max_column + 1)]
        if not any(value is not None for value in values):
            continue
        status = str(automation.row_value(values, headers, "status") or "Sample").strip()
        reminder = None
        training = None
        issue = ""
        try:
            reminder = automation.due_at(values, headers, row_number)
            training = automation.as_date(
                automation.row_value(values, headers, "actual training date"),
                "Actual Training Date",
                row_number,
            )
        except ValueError as error:
            issue = str(error)
        rows.append(
            {
                "row": row_number,
                "recipient": cell_text(automation.row_value(values, headers, "recipient email")),
                "participant": cell_text(automation.row_value(values, headers, "participant name")),
                **{key: cell_text(values[headers[name] - 1]) if name in headers else "" for name, key in SOURCE_FIELDS.items()},
                "reminderAt": reminder.isoformat() if reminder else "",
                "trainingDate": training.isoformat() if training else "",
                "status": status,
                "sentAt": cell_text(automation.row_value(values, headers, "sent at")),
                "messageId": cell_text(automation.row_value(values, headers, "gmail message id")),
                "notes": cell_text(automation.row_value(values, headers, "notes")) or issue,
                "isDue": bool(status.lower() == "ready" and reminder and reminder <= now),
            }
        )
    workbook.close()
    rows.sort(key=lambda item: (not item["isDue"], item["reminderAt"], item["row"]))
    ready = [item for item in rows if item["status"].lower() == "ready"]
    due_count = sum(1 for item in ready if item["isDue"])
    summary = {
        "due": due_count,
        "upcoming": sum(1 for item in ready if not item["isDue"]),
        "sent": sum(1 for item in rows if item["status"].lower() == "sent"),
        "needsAttention": sum(1 for item in rows if item["status"].lower() in {"paused", "error", "sending"}),
        "smtp": smtp_settings.public_settings(),
        "scheduler": scheduler_state(),
        "schedulerMode": "cloud" if CLOUD_MODE else "windows",
        "timeZone": "Asia/Kolkata",
    }
    return rows, summary


def scheduler_state() -> str:
    try:
        enabled = json.loads(SCHEDULER_PATH.read_text(encoding="utf-8")).get("enabled")
        return "Installed" if enabled else "Not installed"
    except (OSError, json.JSONDecodeError):
        return "Not installed"


def install_schedule() -> tuple[bool, str]:
    try:
        temporary = SCHEDULER_PATH.with_suffix(".tmp")
        temporary.write_text(json.dumps({"enabled": True}), encoding="utf-8")
        temporary.replace(SCHEDULER_PATH)
    except OSError as error:
        return False, f"Could not enable automatic sending: {error}"
    mode = "cloud service" if CLOUD_MODE else "dashboard server"
    return True, f"Automatic sending is enabled. The {mode} checks reminders every minute while it is running."


def remove_legacy_scheduled_task() -> None:
    """Remove the old task that launched a separate Python process each minute."""
    if sys.platform != "win32":
        return
    hidden = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        query = subprocess.run(
            ["schtasks", "/Query", "/TN", "Training Email Reminders"],
            capture_output=True, text=True, timeout=10, check=False, creationflags=hidden,
        )
        if query.returncode == 0:
            subprocess.run(
                ["schtasks", "/Delete", "/TN", "Training Email Reminders", "/F"],
                capture_output=True, text=True, timeout=15, check=False, creationflags=hidden,
            )
    except (OSError, subprocess.TimeoutExpired) as error:
        print(f"[dashboard] Could not remove the old scheduled task: {error}")


UPLOAD_ALIASES = {
    "recipient email": {"email", "emailid", "emailaddress", "recipientemail", "recipientemailid", "traineeemail", "participantemail", "gmailid"},
    "participant name": {"name", "participantname", "employeename", "traineename", "studentname", "candidate name".replace(" ", ""), "traineeemployee"},
    "reminder date": {"reminderdate", "remainderdate", "reminderon", "remainderon"},
    "reminder time": {"remindertime", "remaindertime", "remainder time".replace(" ", ""), "remindertimeist"},
    "reminder date time": {"reminderdatetime", "reminderdateandtime", "reminderdateime", "remainderdateandtime", "reminderdate/time".replace("/", "")},
    "actual training date": {"actualtrainingdate", "actualtraingdate", "actualtraniningdate", "trainingdate", "traingdate", "actualtrainingday", "trainingday"},
    "serial number": {"srno", "serialno", "serialnumber"},
    "department": {"department", "dept"},
    "training topic": {"topicname", "trainingtopic", "topic"},
    "plan status": {"plan", "planstatus"},
    "actual status": {"actual", "actualstatus"},
    "replan status": {"replanstatus", "replan"},
    "training time": {"time", "trainingtime", "sessiontime"},
    "trainer name": {"trainername", "trainnername", "trainer"},
}


def normalize_heading(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value or "").casefold())


def bind_source_headers(sheet, header_row: int, headers: dict, create_missing: bool = False) -> None:
    normalized = {
        normalize_heading(sheet.cell(header_row, column).value): column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(header_row, column).value is not None
    }
    aliases = {"serial number": ("serialnumber", "serialno", "srno")}
    for name, label in SOURCE_HEADERS.items():
        column = next((normalized[key] for key in (normalize_heading(label), *aliases.get(name, ())) if key in normalized), None)
        if column is None and create_missing:
            column = sheet.max_column + 1
            sheet.cell(header_row, column, label)
            normalized[normalize_heading(label)] = column
        if column is not None:
            headers[name] = column


def find_upload_header(sheet):
    for row_index in range(1, min(sheet.max_row, 25) + 1):
        cells = {normalize_heading(cell.value): cell.column for cell in sheet[row_index] if cell.value is not None}
        mapped = {}
        for field, aliases in UPLOAD_ALIASES.items():
            for alias in aliases:
                key = normalize_heading(alias)
                if key in cells:
                    mapped[field] = cells[key]
                    break
        if "recipient email" not in mapped:
            for next_row in range(row_index + 1, min(sheet.max_row, row_index + 2) + 1):
                email_column = next((cell.column for cell in sheet[next_row] if normalize_heading(cell.value) in UPLOAD_ALIASES["recipient email"]), None)
                if email_column is not None:
                    mapped["recipient email"] = email_column
                    break
        has_separate = "reminder date" in mapped and "reminder time" in mapped
        if "recipient email" in mapped and "actual training date" in mapped and (has_separate or "reminder date time" in mapped):
            return row_index, mapped
    raise ValueError(
        "I could not identify the columns. Include email/email ID, Reminder Date and Reminder Time "
        "(or Reminder Date Time), and Actual Training Date in the header row."
    )


def parse_combined_datetime(value, row_number: int) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        raise ValueError(f"Excel row {row_number}: the combined reminder field also needs a time.")
    if isinstance(value, str):
        try:
            return datetime.fromisoformat(value.strip().replace("Z", "+00:00")).replace(tzinfo=None)
        except ValueError:
            for pattern in ("%d/%m/%Y %H:%M", "%m/%d/%Y %I:%M %p", "%Y-%m-%d %H:%M", "%d.%m.%Y %H.%M%p", "%d.%m.%Y %I.%M%p"):
                try:
                    return datetime.strptime(value.strip().replace(" ", ""), pattern.replace(" ", ""))
                except ValueError:
                    pass
    raise ValueError(f"Excel row {row_number}: reminder date/time could not be read.")


def extract_upload(file_name: str, file_bytes: bytes) -> tuple[list[dict], dict]:
    if not file_name.casefold().endswith(".xlsx"):
        raise ValueError("Please upload an Excel .xlsx file.")
    if not file_bytes or len(file_bytes) > 10 * 1024 * 1024:
        raise ValueError("The file is empty or larger than 10 MB.")
    try:
        workbook = load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
    except Exception as error:
        raise ValueError("This file could not be opened as an .xlsx workbook.") from error
    found = None
    for sheet in workbook.worksheets:
        try:
            header_row, headers = find_upload_header(sheet)
            found = (sheet, header_row, headers)
            break
        except ValueError:
            continue
    if not found:
        workbook.close()
        raise ValueError("No worksheet has the required reminder columns. Check the first 25 rows for headers.")
    sheet, header_row, headers = found
    extracted = []
    errors = []
    for row_number, values in enumerate(sheet.iter_rows(min_row=header_row + 1, values_only=True), header_row + 1):
        if not any(value is not None and str(value).strip() for value in values):
            continue
        if "serial number" in headers and "training topic" in headers:
            serial = values[headers["serial number"] - 1]
            topic = values[headers["training topic"] - 1]
            if serial in (None, "") or not str(topic or "").strip():
                continue
        try:
            email = str(values[headers["recipient email"] - 1] or "").strip()
            if normalize_heading(email) == "email":
                email = ""
            valid_email = bool(re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", email))
            if email and not valid_email:
                raise ValueError("email address is invalid")
            has_schedule = any(
                values[headers[field] - 1] not in (None, "")
                for field in ("reminder date", "reminder time", "reminder date time", "actual training date")
                if field in headers
            )
            if not valid_email and not has_schedule:
                continue
            if "reminder date time" in headers:
                when = parse_combined_datetime(values[headers["reminder date time"] - 1], row_number)
                reminder_date, reminder_time = when.date(), when.time()
            else:
                reminder_date = automation.as_date(values[headers["reminder date"] - 1], "Reminder Date", row_number)
                reminder_time = automation.as_time(values[headers["reminder time"] - 1], row_number)
            training_date = automation.as_date(values[headers["actual training date"] - 1], "Actual Training Date", row_number)
            participant = str(values[headers["participant name"] - 1] or "").strip() if "participant name" in headers else ""
            extras = {
                key: values[headers[name] - 1] if name in headers else ""
                for name, key in SOURCE_FIELDS.items()
            }
            extracted.append({
                "recipient": email,
                "participant": participant,
                "reminder_date": reminder_date,
                "reminder_time": reminder_time,
                "training_date": training_date,
                "status": "Paused",
                "missing_email": not valid_email,
                **extras,
            })
        except (ValueError, TypeError, IndexError) as error:
            errors.append(f"Row {row_number}: {error}")
    workbook.close()
    if errors:
        preview = "; ".join(errors[:5])
        if len(errors) > 5:
            preview += f"; and {len(errors) - 5} more issue(s)"
        raise ValueError(f"No reminders were imported. Fix these rows and upload again: {preview}")
    if not extracted:
        raise ValueError("No training rows with usable reminder and actual training dates were found.")
    return extracted, {
        "worksheet": sheet.title,
        "headerRow": header_row,
        "count": len(extracted),
        "paused": sum(1 for record in extracted if record["status"] == "Paused"),
        "missingEmail": sum(1 for record in extracted if record["missing_email"]),
    }


def replace_reminders(records: list[dict], source_name: str) -> None:
    workbook = load_workbook(automation.WORKBOOK_PATH)
    sheet = workbook.active
    header_row, headers = automation.find_header_row(sheet)
    bind_source_headers(sheet, header_row, headers, create_missing=True)
    for row_index in range(header_row + 1, sheet.max_row + 1):
        for column in range(1, sheet.max_column + 1):
            sheet.cell(row_index, column).value = None
    start_row = header_row + 1
    for offset, record in enumerate(records):
        row_number = start_row + offset
        values = {
            "recipient email": record["recipient"],
            "participant name": record["participant"],
            "reminder date": record["reminder_date"],
            "reminder time": record["reminder_time"],
            "actual training date": record["training_date"],
            "status": record.get("status", "Ready"),
            "sent at": None,
            "gmail message id": None,
            "notes": f"Imported from {Path(source_name).name}" + (" · Add a valid email before setting Ready." if record.get("missing_email") else " · Review, then set Ready to activate."),
            **{name: record.get(key, "") for name, key in SOURCE_FIELDS.items()},
        }
        for field, value in values.items():
            sheet.cell(row_number, headers[field]).value = value
        sheet.cell(row_number, headers["reminder date"]).number_format = "dd-mmm-yyyy"
        sheet.cell(row_number, headers["reminder time"]).number_format = "hh:mm"
        sheet.cell(row_number, headers["actual training date"]).number_format = "dd-mmm-yyyy"
    workbook.save(automation.WORKBOOK_PATH)
    workbook.close()


def read_multipart_upload(handler) -> tuple[str, bytes]:
    length = int(handler.headers.get("Content-Length", "0"))
    if length <= 0 or length > 10 * 1024 * 1024 + 100_000:
        raise ValueError("The file is empty or larger than 10 MB.")
    body = handler.rfile.read(length)
    envelope = (
        f"Content-Type: {handler.headers.get('Content-Type', '')}\r\n"
        "MIME-Version: 1.0\r\n\r\n"
    ).encode("ascii") + body
    message = BytesParser(policy=email_policy).parsebytes(envelope)
    if not message.is_multipart():
        raise ValueError("Choose an Excel file to upload.")
    for part in message.iter_parts():
        if part.get_content_disposition() == "form-data" and part.get_param("name", header="content-disposition") == "file":
            file_name = Path(part.get_filename() or "reminders.xlsx").name
            return file_name, part.get_payload(decode=True) or b""
    raise ValueError("Choose an Excel file to upload.")


def update_status(row_number: int, new_status: str) -> None:
    if new_status not in {"Ready", "Paused"}:
        raise ValueError("Status must be Ready or Paused.")
    workbook = load_workbook(automation.WORKBOOK_PATH)
    sheet = workbook.active
    header_row, headers = automation.find_header_row(sheet)
    bind_source_headers(sheet, header_row, headers)
    if row_number <= header_row or row_number > sheet.max_row:
        raise ValueError("Reminder row was not found.")
    if new_status == "Ready":
        recipient = str(sheet.cell(row_number, headers["recipient email"]).value or "").strip()
        if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", recipient):
            workbook.close()
            raise ValueError("Add a valid Email before setting this row to Ready.")
    current = str(sheet.cell(row_number, headers["status"]).value or "").strip().lower()
    if current in {"sent", "sending"}:
        raise ValueError("Sent or in-progress reminders cannot be changed here.")
    sheet.cell(row_number, headers["status"]).value = new_status
    if new_status == "Ready":
        sheet.cell(row_number, headers["notes"]).value = ""
    workbook.save(automation.WORKBOOK_PATH)
    workbook.close()


def add_reminder(payload: dict) -> None:
    _write_reminder(payload, None)


def update_reminder(row_number: int, payload: dict) -> None:
    _write_reminder(payload, row_number)


def _write_reminder(payload: dict, row_number: int | None) -> None:
    recipient = str(payload.get("recipient", "")).strip()
    participant = str(payload.get("participant", "")).strip()
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]+", recipient):
        raise ValueError("Enter a valid recipient email address.")
    try:
        reminder_date = date.fromisoformat(str(payload.get("reminderDate", "")))
        reminder_time = datetime.strptime(str(payload.get("reminderTime", "")), "%H:%M").time()
        training_date = date.fromisoformat(str(payload.get("trainingDate", "")))
    except ValueError as error:
        raise ValueError("Enter valid reminder date, time, and training date values.") from error

    workbook = load_workbook(automation.WORKBOOK_PATH)
    sheet = workbook.active
    header_row, headers = automation.find_header_row(sheet)
    new_row = max(sheet.max_row + 1, header_row + 1) if row_number is None else row_number
    bind_source_headers(sheet, header_row, headers, create_missing=True)
    if row_number is not None:
        if row_number <= header_row or row_number > sheet.max_row:
            workbook.close()
            raise ValueError("Reminder row was not found.")
        current = str(sheet.cell(row_number, headers["status"]).value or "").strip().lower()
        if current in {"sent", "sending"}:
            workbook.close()
            raise ValueError("Sent or in-progress reminders cannot be edited.")
    entries = {
        "recipient email": recipient,
        "participant name": participant,
        "reminder date": reminder_date,
        "reminder time": reminder_time,
        "actual training date": training_date,
        "status": "Ready",
        "sent at": sheet.cell(new_row, headers["sent at"]).value if row_number is not None else None,
        "gmail message id": sheet.cell(new_row, headers["gmail message id"]).value if row_number is not None else None,
        "notes": None,
        **{name: str(payload.get(key, "")).strip() for name, key in SOURCE_FIELDS.items()},
    }
    for name, value in entries.items():
        cell = sheet.cell(new_row, headers[name], value)
        cell.alignment = Alignment(vertical="top", wrap_text=name in {"notes", "gmail message id"})
        cell.border = Border(bottom=Side(style="hair", color="D9E2F3"))
    sheet.cell(new_row, headers["reminder date"]).number_format = "dd-mmm-yyyy"
    sheet.cell(new_row, headers["reminder time"]).number_format = "hh:mm"
    sheet.cell(new_row, headers["actual training date"]).number_format = "dd-mmm-yyyy"
    workbook.save(automation.WORKBOOK_PATH)
    workbook.close()


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "TrainingReminders/1.0"

    def _send_json(self, status: int, body: dict) -> None:
        data = json.dumps(body, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def _read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        if length <= 0 or length > 32_768:
            raise ValueError("Request body is empty or too large.")
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def _valid_host(self) -> bool:
        if CLOUD_MODE:
            return True
        return self.headers.get("Host", "").lower() in {f"{HOST}:{PORT}", f"localhost:{PORT}"}

    def _authorized(self) -> bool:
        if not CLOUD_MODE:
            return True
        expected_user = os.environ.get("DASHBOARD_USERNAME", "")
        expected_password = os.environ.get("DASHBOARD_PASSWORD", "")
        supplied = self.headers.get("Authorization", "")
        try:
            scheme, encoded = supplied.split(" ", 1)
            username, password = base64.b64decode(encoded, validate=True).decode("utf-8").split(":", 1)
        except (ValueError, UnicodeDecodeError):
            username, password, scheme = "", "", ""
        authorized = (
            scheme.casefold() == "basic"
            and hmac.compare_digest(username.encode("utf-8"), expected_user.encode("utf-8"))
            and hmac.compare_digest(password.encode("utf-8"), expected_password.encode("utf-8"))
        )
        if not authorized:
            self.send_response(401)
            self.send_header("WWW-Authenticate", 'Basic realm="Training Reminders", charset="UTF-8"')
            self.send_header("Content-Length", "0")
            self.end_headers()
        return authorized

    def do_GET(self) -> None:  # noqa: N802
        path = urlparse(self.path).path
        if path == "/healthz":
            self._send_json(200, {"status": "ok"})
            return
        if not self._authorized():
            return
        if path == "/":
            try:
                page = PAGE_PATH.read_bytes()
            except OSError as error:
                self._send_json(500, {"error": str(error)})
                return
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(page)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.send_header("Referrer-Policy", "no-referrer")
            self.end_headers()
            self.wfile.write(page)
            return
        if path == "/api/dashboard":
            try:
                rows, summary = workbook_rows()
                self._send_json(200, {"rows": rows, "summary": summary})
            except Exception as error:
                self._send_json(500, {"error": str(error)})
            return
        self._send_json(404, {"error": "Not found"})

    def do_POST(self) -> None:  # noqa: N802
        if not self._valid_host():
            self._send_json(403, {"error": "This local dashboard only accepts requests from this computer."})
            return
        if not self._authorized():
            return
        path = unquote(urlparse(self.path).path)
        try:
            if path == "/api/upload":
                file_name, file_bytes = read_multipart_upload(self)
                records, details = extract_upload(file_name, file_bytes)
                replace_reminders(records, file_name)
                self._send_json(200, {"message": f"Imported {details['count']} training rows from {file_name}; all are paused for review, and {details['missingEmail']} need an email address.", "count": details["count"], "worksheet": details["worksheet"]})
                return
            payload = self._read_json()
            if path == "/api/settings":
                smtp_settings.save_settings(payload)
                self._send_json(200, {"message": "SMTP settings saved securely.", "smtp": smtp_settings.public_settings()})
                return
            if path == "/api/test-smtp":
                try:
                    result = smtp_settings.check_connection()
                except PermissionError as error:
                    if getattr(error, "winerror", None) == 10013:
                        self._send_json(
                            503,
                            {"error": "Windows denied this process permission to open the Gmail SMTP socket (WinError 10013). Stop this server and run `python web_app.py` from a normal Windows PowerShell or Terminal window, signed in to the same Windows account. If it still fails there, allow python.exe outbound access in Windows Firewall or your network policy."},
                        )
                        return
                    raise
                self._send_json(200, {"message": result})
                return
            if path == "/api/enable-schedule":
                success, output = install_schedule()
                self._send_json(200 if success else 400, {"message": output, "success": success})
                return
            if path == "/api/reminders":
                add_reminder(payload)
                self._send_json(201, {"message": "Reminder added to the workbook."})
                return
            match = re.fullmatch(r"/api/reminders/(\d+)/status", path)
            if match:
                update_status(int(match.group(1)), str(payload.get("status", "")))
                self._send_json(200, {"message": "Reminder status updated."})
                return
            match = re.fullmatch(r"/api/reminders/(\d+)/update", path)
            if match:
                update_reminder(int(match.group(1)), payload)
                self._send_json(200, {"message": "Reminder updated and set to Ready."})
                return
            if path in {"/api/preview", "/api/send-due"}:
                output = io.StringIO()
                with AUTOMATION_LOCK, redirect_stdout(output), redirect_stderr(output):
                    return_code = automation.process(
                        send=path == "/api/send-due",
                        workbook_path=automation.WORKBOOK_PATH,
                    )
                self._send_json(
                    200 if return_code == 0 else 400,
                    {"output": output.getvalue().strip(), "success": return_code == 0},
                )
                return
            self._send_json(404, {"error": "Not found"})
        except subprocess.TimeoutExpired:
            self._send_json(504, {"error": "The operation timed out. Check the automation log and your email Sent folder."})
        except (ValueError, json.JSONDecodeError) as error:
            self._send_json(400, {"error": str(error)})
        except Exception as error:
            self._send_json(500, {"error": str(error)})

    def log_message(self, format: str, *args) -> None:
        print(f"[dashboard] {self.address_string()} - {format % args}")


def _create_empty_workbook() -> None:
    if automation.WORKBOOK_PATH.exists():
        return
    from openpyxl import Workbook

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Training Reminders"
    sheet.append([
        "Recipient Email", "Participant Name", "Reminder Date", "Reminder Time",
        "Actual Training Date", "Status", "Sent At", "Gmail Message ID", "Notes",
        *SOURCE_HEADERS.values(),
    ])
    workbook.save(automation.WORKBOOK_PATH)
    workbook.close()


def _scheduler_worker() -> None:
    while True:
        if scheduler_state() == "Installed":
            try:
                with AUTOMATION_LOCK:
                    return_code = automation.process(send=True, workbook_path=automation.WORKBOOK_PATH)
                if return_code:
                    print("[scheduler] An automatic reminder check failed; see reminder_automation.log.")
            except Exception as error:
                print(f"[scheduler] Automatic check failed: {error}")
        time.sleep(60)


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if CLOUD_MODE:
        missing = [name for name in ("DASHBOARD_USERNAME", "DASHBOARD_PASSWORD", "SMTP_ENCRYPTION_KEY") if not os.environ.get(name)]
        if missing:
            raise SystemExit("Missing required cloud environment variables: " + ", ".join(missing))
    _create_empty_workbook()
    automation.log_setup()
    if not CLOUD_MODE:
        remove_legacy_scheduled_task()
        try:
            with urlopen(f"http://{HOST}:{PORT}/api/dashboard", timeout=2) as response:
                payload = json.loads(response.read().decode("utf-8"))
            if response.status == 200 and isinstance(payload.get("rows"), list):
                print(f"Training Reminders dashboard is already running: http://{HOST}:{PORT}")
                webbrowser.open(f"http://{HOST}:{PORT}")
                return
        except Exception:
            pass
    threading.Thread(target=_scheduler_worker, name="reminder-scheduler", daemon=True).start()
    try:
        server = LocalDashboardServer((HOST, PORT), DashboardHandler)
    except OSError as error:
        try:
            if not CLOUD_MODE:
                with urlopen(f"http://{HOST}:{PORT}/api/dashboard", timeout=2) as response:
                    payload = json.loads(response.read().decode("utf-8"))
                if response.status == 200 and isinstance(payload.get("rows"), list):
                    print(f"Training Reminders dashboard is already running: http://{HOST}:{PORT}")
                    webbrowser.open(f"http://{HOST}:{PORT}")
                    return
            if CLOUD_MODE:
                raise error
        except Exception:
            pass
        raise SystemExit(
            f"Could not open http://{HOST}:{PORT} ({error}). "
            "The port may be occupied or blocked; close the other process or choose another port."
        ) from error
    print(f"Training Reminders dashboard listening on {HOST}:{PORT}")
    if not CLOUD_MODE:
        print("Keep this window open while using the dashboard. Press Ctrl+C to stop.")
        try:
            webbrowser.open(f"http://{HOST}:{PORT}")
        except Exception:
            print(f"Open http://{HOST}:{PORT} in your browser.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nDashboard stopped.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
