from __future__ import annotations

import csv
import io
import json
import os
from dotenv import load_dotenv
load_dotenv(override=True)
from datetime import datetime, timezone, timedelta
from pathlib import Path
from typing import Any, Dict, List, Tuple
from uuid import uuid4
from urllib.parse import urljoin, urlparse

from flask import Flask, Response, flash, redirect, render_template, request, session, url_for, jsonify
from flask_login import (
    LoginManager,
    UserMixin,
    current_user,
    login_required,
    login_user,
    logout_user,
)
from werkzeug.security import check_password_hash
from werkzeug.utils import secure_filename
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter
from sqlalchemy.orm import sessionmaker
from sqlalchemy import func, and_, or_

from dotenv import load_dotenv
load_dotenv(override=True)

# Import our database models and configuration
from models import Goal, TimeEntry
from db import engine, SessionLocal, init_db

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

# Initialize database
init_db()

ADMIN_EMAIL = (os.environ.get("ADMIN_EMAIL") or "").strip()
ADMIN_HASH = (os.environ.get("ADMIN_HASH") or "").strip()

login_manager = LoginManager(app)
login_manager.login_view = "login"
login_manager.login_message_category = "warning"


class AdminUser(UserMixin):
    def __init__(self, email: str) -> None:
        self._email = email

    def get_id(self) -> str:
        return self._email


@login_manager.user_loader
def load_user(user_id: str) -> AdminUser | None:
    if ADMIN_EMAIL and user_id == ADMIN_EMAIL:
        return AdminUser(user_id)
    return None

DATA_DIR = Path("data")
DATA_FILE = DATA_DIR / "vocab.json"
EXPORT_FIELDS = [
    "word",
    "sentence",
    "synonym",
    "native_meaning",
    "type",
    "base_word",
    "created_at",
]
IMPORT_FIELDS = EXPORT_FIELDS
DEDUP_FIELDS = ("word", "sentence", "synonym", "type", "base_word", "native_meaning")
ALLOWED_IMPORT_EXTENSIONS = {".csv", ".json", ".xlsx"}
PENDING_IMPORT_SESSION_KEY = "pending_import_rows"
PENDING_IMPORT_META_SESSION_KEY = "pending_import_meta"

app.config["MAX_CONTENT_LENGTH"] = 2 * 1024 * 1024


def ensure_store() -> None:
    """Make sure the data directory and file exist before use."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not DATA_FILE.exists():
        DATA_FILE.write_text("[]\n", encoding="utf-8")


def load_items() -> List[Dict[str, Any]]:
    """Load vocabulary entries, normalise fields, and backfill missing IDs."""
    ensure_store()
    try:
        raw = DATA_FILE.read_text(encoding="utf-8")
    except OSError:
        return []

    if not raw.strip():
        return []

    try:
        data = json.loads(raw)
    except json.JSONDecodeError:
        save_items([])
        return []

    if not isinstance(data, list):
        return []

    cleaned: List[Dict[str, Any]] = []
    needs_persist = False

    for item in data:
        if not isinstance(item, dict):
            continue

        entry_id = str(item.get("id") or "").strip()
        if not entry_id:
            entry_id = uuid4().hex
            needs_persist = True

        entry = {
            "id": entry_id,
            "word": str(item.get("word", "") or "").strip(),
            "sentence": str(item.get("sentence", "") or "").strip(),
            "synonym": str(item.get("synonym", "") or "").strip(),
            "type": str(item.get("type", "") or "").strip(),
            "base_word": str(item.get("base_word", "") or "").strip(),
            "native_meaning": str(item.get("native_meaning", "") or "").strip(),
            "created_at": str(item.get("created_at", "") or ""),
        }
        if "native_meaning" not in item:
            needs_persist = True

        cleaned.append(entry)

    if needs_persist:
        save_items(cleaned)

    return cleaned


def save_items(items: List[Dict[str, Any]]) -> None:
    """Persist entries safely with UTF-8 encoding."""
    ensure_store()
    temp_path = DATA_FILE.with_suffix(".tmp")
    payload = json.dumps(items, ensure_ascii=False, indent=2)
    temp_path.write_text(f"{payload}\n", encoding="utf-8")
    temp_path.replace(DATA_FILE)


def get_sorted_items() -> List[Dict[str, Any]]:
    items = load_items()
    items.sort(key=lambda item: item.get("created_at") or "", reverse=True)
    return items


def export_timestamp() -> str:
    return datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%SZ")


def normalise_field_key(key: Any) -> str:
    if key is None:
        return ""
    return str(key).strip().lower().replace(" ", "_")


def stringify_value(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def normalise_import_row(raw: Dict[str, Any]) -> Dict[str, str]:
    prepared = {normalise_field_key(k): v for k, v in raw.items()}
    entry: Dict[str, str] = {}
    for field in IMPORT_FIELDS:
        entry[field] = stringify_value(prepared.get(field, ""))
    return entry


def parse_import_payload(data: bytes, extension: str) -> List[Dict[str, str]]:
    ext = extension.lower()
    if ext == ".csv":
        return _parse_csv(data)
    if ext == ".json":
        return _parse_json(data)
    if ext == ".xlsx":
        return _parse_xlsx(data)
    raise ValueError("Unsupported file type.")


def _parse_csv(data: bytes) -> List[Dict[str, str]]:
    try:
        text_stream = data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("CSV file must be UTF-8 encoded.") from exc
    reader = csv.DictReader(io.StringIO(text_stream))
    if reader.fieldnames is None:
        raise ValueError("CSV file is missing a header row.")
    rows: List[Dict[str, str]] = []
    for raw_row in reader:
        normalised = normalise_import_row(raw_row)
        if not any(normalised[field] for field in DEDUP_FIELDS):
            continue
        rows.append(normalised)
    return rows


def _parse_json(data: bytes) -> List[Dict[str, str]]:
    try:
        text_stream = data.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("JSON file must be UTF-8 encoded.") from exc
    try:
        payload: Any = json.loads(text_stream)
    except json.JSONDecodeError as exc:
        raise ValueError("JSON file is not valid JSON.") from exc
    if isinstance(payload, dict):
        if "items" in payload and isinstance(payload["items"], list):
            payload = payload["items"]
        else:
            payload = [payload]
    if not isinstance(payload, list):
        raise ValueError("JSON file must contain a list of entries.")
    rows: List[Dict[str, str]] = []
    for item in payload:
        if not isinstance(item, dict):
            continue
        normalised = normalise_import_row(item)
        if not any(normalised[field] for field in DEDUP_FIELDS):
            continue
        rows.append(normalised)
    return rows


def _parse_xlsx(data: bytes) -> List[Dict[str, str]]:
    buffer = io.BytesIO(data)
    try:
        workbook = load_workbook(buffer, read_only=True, data_only=True)
    except Exception as exc:
        raise ValueError("XLSX file could not be read.") from exc
    try:
        sheet = workbook.active
        header_row = next(sheet.iter_rows(values_only=True), None)
        if header_row is None:
            return []
        headers = [normalise_field_key(value) for value in header_row]
        rows: List[Dict[str, str]] = []
        for row_values in sheet.iter_rows(values_only=True, min_row=2):
            raw_entry: Dict[str, Any] = {}
            for index, value in enumerate(row_values):
                if index >= len(headers):
                    continue
                header_key = headers[index]
                if not header_key:
                    continue
                raw_entry[header_key] = value
            normalised = normalise_import_row(raw_entry)
            if not any(normalised[field] for field in DEDUP_FIELDS):
                continue
            rows.append(normalised)
        return rows
    finally:
        workbook.close()


def deduplication_key(item: Dict[str, Any]) -> Tuple[str, str, str, str, str, str]:
    return tuple(stringify_value(item.get(field)) for field in DEDUP_FIELDS)


def is_safe_redirect(target: str) -> bool:
    if not target:
        return False
    host_url = urlparse(request.host_url)
    redirect_url = urlparse(urljoin(request.host_url, target))
    return redirect_url.scheme in ("http", "https") and host_url.netloc == redirect_url.netloc


@app.route("/login", methods=["GET"])
def login() -> Response | str:
    next_url = request.args.get("next")
    if next_url and not is_safe_redirect(next_url):
        next_url = None

    if current_user.is_authenticated:
        if next_url:
            return redirect(next_url)
        return redirect(url_for("index"))

    return render_template("login.html", next_url=next_url)


@app.route("/login", methods=["POST"])
def login_submit() -> Response:
    email = (request.form.get("email") or "").strip()
    password = request.form.get("password") or ""
    next_url = request.form.get("next")
    if next_url and not is_safe_redirect(next_url):
        next_url = None

    if not ADMIN_EMAIL or not ADMIN_HASH:
        flash("Authentication is not configured.", "error")
        return redirect(url_for("login", next=next_url) if next_url else url_for("login"))

    try:
        password_matches = check_password_hash(ADMIN_HASH, password)
    except ValueError:
        flash("Authentication is misconfigured.", "error")
        return redirect(url_for("login", next=next_url) if next_url else url_for("login"))

    if email == ADMIN_EMAIL and password_matches:
        login_user(AdminUser(ADMIN_EMAIL))
        flash("Logged in successfully.", "success")
        return redirect(next_url or url_for("index"))

    flash("Invalid email or password.", "error")
    return redirect(url_for("login", next=next_url) if next_url else url_for("login"))


@app.route("/logout", methods=["POST"])
@login_required
def logout() -> Response:
    logout_user()
    flash("Logged out.", "success")
    return redirect(url_for("index"))


@app.route("/healthz", methods=["GET"])
def healthz() -> Response:
    """Health check endpoint for deployment platforms."""
    return Response("OK", status=200, mimetype="text/plain")


# Dashboard route
@app.route("/dashboard", methods=["GET"])
@login_required
def dashboard() -> str:
    """Main dashboard showing vocabulary and goals summary."""
    db = SessionLocal()
    try:
        # Get vocabulary summary
        vocab_items = get_sorted_items()
        vocab_total = len(vocab_items)
        recent_vocab = vocab_items[:5]  # Last 5 entries
        
        # Get goals summary
        active_goals = db.query(Goal).filter(Goal.is_active == True).all()
        
        # Get time tracking summary
        today = datetime.now().date()
        week_start = today - timedelta(days=today.weekday())
        month_start = today.replace(day=1)
        
        # Today's time
        today_time = db.query(func.sum(TimeEntry.minutes)).join(Goal).filter(
            func.date(TimeEntry.started_at) == today
        ).scalar() or 0
        
        # This week's time
        week_time = db.query(func.sum(TimeEntry.minutes)).join(Goal).filter(
            func.date(TimeEntry.started_at) >= week_start
        ).scalar() or 0
        
        # This month's time
        month_time = db.query(func.sum(TimeEntry.minutes)).join(Goal).filter(
            func.date(TimeEntry.started_at) >= month_start
        ).scalar() or 0
        
        return render_template("dashboard.html", 
                             vocab_total=vocab_total,
                             recent_vocab=recent_vocab,
                             active_goals=active_goals,
                             today_time=today_time,
                             week_time=week_time,
                             month_time=month_time)
    finally:
        db.close()


# Goals CRUD routes
@app.route("/goals", methods=["GET"])
@login_required
def goals_list() -> str:
    """List all goals."""
    db = SessionLocal()
    try:
        goals = db.query(Goal).order_by(Goal.created_at.desc()).all()
        return render_template("goals.html", goals=goals)
    finally:
        db.close()


@app.route("/goals", methods=["POST"])
@login_required
def create_goal() -> Response:
    """Create a new goal."""
    title = request.form.get("title", "").strip()
    category = request.form.get("category", "").strip()
    notes = request.form.get("notes", "").strip()
    
    if not title:
        flash("Title is required.", "error")
        return redirect(url_for("goals_list"))
    
    db = SessionLocal()
    try:
        goal = Goal(title=title, category=category, notes=notes)
        db.add(goal)
        db.commit()
        flash("Goal created successfully!", "success")
        return redirect(url_for("goals_list"))
    finally:
        db.close()


@app.route("/goals/<int:goal_id>/edit", methods=["POST"])
@login_required
def update_goal(goal_id: int) -> Response:
    """Update a goal."""
    title = request.form.get("title", "").strip()
    category = request.form.get("category", "").strip()
    notes = request.form.get("notes", "").strip()
    is_active = request.form.get("is_active") == "on"
    
    if not title:
        flash("Title is required.", "error")
        return redirect(url_for("goals_list"))
    
    db = SessionLocal()
    try:
        goal = db.query(Goal).filter(Goal.id == goal_id).first()
        if not goal:
            flash("Goal not found.", "error")
            return redirect(url_for("goals_list"))
        
        goal.title = title
        goal.category = category
        goal.notes = notes
        goal.is_active = is_active
        db.commit()
        flash("Goal updated successfully!", "success")
        return redirect(url_for("goals_list"))
    finally:
        db.close()


@app.route("/goals/<int:goal_id>/delete", methods=["POST"])
@login_required
def delete_goal(goal_id: int) -> Response:
    """Delete a goal."""
    db = SessionLocal()
    try:
        goal = db.query(Goal).filter(Goal.id == goal_id).first()
        if not goal:
            flash("Goal not found.", "error")
            return redirect(url_for("goals_list"))
        
        db.delete(goal)
        db.commit()
        flash("Goal deleted successfully!", "success")
        return redirect(url_for("goals_list"))
    finally:
        db.close()


# Time logging routes
@app.route("/goals/<int:goal_id>/log", methods=["GET"])
@login_required
def log_time_form(goal_id: int) -> str:
    """Form to log time for a specific goal."""
    db = SessionLocal()
    try:
        goal = db.query(Goal).filter(Goal.id == goal_id).first()
        if not goal:
            flash("Goal not found.", "error")
            return redirect(url_for("goals_list"))
        
        # Get recent time entries for this goal
        recent_entries = db.query(TimeEntry).filter(TimeEntry.goal_id == goal_id)\
            .order_by(TimeEntry.started_at.desc()).limit(10).all()
        
        return render_template("goal_detail.html", goal=goal, recent_entries=recent_entries)
    finally:
        db.close()


@app.route("/goals/<int:goal_id>/log", methods=["POST"])
@login_required
def log_time(goal_id: int) -> Response:
    """Log time for a specific goal."""
    started_at_str = request.form.get("started_at", "").strip()
    minutes = request.form.get("minutes", "").strip()
    note = request.form.get("note", "").strip()
    
    if not started_at_str or not minutes:
        flash("Started at and minutes are required.", "error")
        return redirect(url_for("log_time_form", goal_id=goal_id))
    
    try:
        started_at = datetime.fromisoformat(started_at_str.replace('Z', '+00:00'))
        minutes_int = int(minutes)
    except (ValueError, TypeError):
        flash("Invalid date or minutes format.", "error")
        return redirect(url_for("log_time_form", goal_id=goal_id))
    
    db = SessionLocal()
    try:
        goal = db.query(Goal).filter(Goal.id == goal_id).first()
        if not goal:
            flash("Goal not found.", "error")
            return redirect(url_for("goals_list"))
        
        time_entry = TimeEntry(
            goal_id=goal_id,
            started_at=started_at,
            minutes=minutes_int,
            note=note
        )
        db.add(time_entry)
        db.commit()
        flash("Time logged successfully!", "success")
        return redirect(url_for("log_time_form", goal_id=goal_id))
    finally:
        db.close()


# API routes
@app.route("/api/goals/summary", methods=["GET"])
@login_required
def goals_summary_api() -> Response:
    """API endpoint for goals summary data."""
    period = request.args.get("period", "day")  # day, week, month
    
    db = SessionLocal()
    try:
        today = datetime.now().date()
        
        if period == "day":
            start_date = today
        elif period == "week":
            start_date = today - timedelta(days=today.weekday())
        elif period == "month":
            start_date = today.replace(day=1)
        else:
            return jsonify({"error": "Invalid period"}), 400
        
        # Get goals with their time totals
        goals_data = db.query(
            Goal.id,
            Goal.title,
            Goal.category,
            func.sum(TimeEntry.minutes).label('total_minutes')
        ).outerjoin(TimeEntry).filter(
            func.date(TimeEntry.started_at) >= start_date
        ).group_by(Goal.id, Goal.title, Goal.category).all()
        
        result = []
        for goal_id, title, category, total_minutes in goals_data:
            result.append({
                "id": goal_id,
                "title": title,
                "category": category,
                "total_minutes": total_minutes or 0
            })
        
        return jsonify({"goals": result, "period": period})
    finally:
        db.close()


@app.before_request
def require_auth():
    """Require authentication for all routes except login, logout, healthz, and static files."""
    # Allow access to login, logout, healthz, and static files without authentication
    if request.endpoint in ['login', 'login_submit', 'logout', 'healthz']:
        return None
    
    # Allow access to static files
    if request.endpoint == 'static':
        return None
    
    # Redirect to login if not authenticated
    if not current_user.is_authenticated:
        return redirect(url_for('login', next=request.url))


@app.route("/", methods=["GET"])
@login_required
def index() -> str:
    """Vocabulary page as home."""
    items = get_sorted_items()
    return render_template("index.html", items=items, total=len(items), edit_entry_id=None)


@app.route("/vocabulary", methods=["GET"])
@login_required
def vocabulary() -> str:
    """Vocabulary management page."""
    items = get_sorted_items()
    return render_template("index.html", items=items, total=len(items), edit_entry_id=None)


@app.route("/add", methods=["POST"])
@login_required
def add_word() -> Response:
    word = (request.form.get("word") or "").strip()
    sentence = (request.form.get("sentence") or "").strip()
    synonym = (request.form.get("synonym") or "").strip()
    type_ = (request.form.get("type") or "").strip()
    base_word = (request.form.get("base_word") or "").strip()
    native_meaning = (request.form.get("native_meaning") or "").strip()

    if not word:
        flash("Field 'word' is required.", "error")
        return redirect(url_for("vocabulary"))

    items = load_items()
    existing_keys = {deduplication_key(item) for item in items}

    new_entry = {
        "id": uuid4().hex,
        "word": word,
        "sentence": sentence,
        "synonym": synonym,
        "type": type_,
        "base_word": base_word,
        "native_meaning": native_meaning,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }

    if deduplication_key(new_entry) in existing_keys:
        flash("This exact entry already exists.", "warning")
        return redirect(url_for("vocabulary"))

    items.append(new_entry)
    save_items(items)
    flash("Word added successfully!", "success")
    return redirect(url_for("vocabulary"))


@app.route("/edit/<entry_id>", methods=["GET"])
@login_required
def edit_entry_form(entry_id: str) -> Response | str:
    items = get_sorted_items()
    if not any(item.get("id") == entry_id for item in items):
        flash("Entry not found.", "error")
        return redirect(url_for("vocabulary"))
    return render_template("index.html", items=items, total=len(items), edit_entry_id=entry_id)


@app.route("/edit/<entry_id>", methods=["POST"])
@login_required
def edit_entry(entry_id: str) -> Response:
    items = load_items()
    target = next((item for item in items if item.get("id") == entry_id), None)

    if target is None:
        flash("Entry not found.", "error")
        return redirect(url_for("vocabulary"))

    word = (request.form.get("word") or "").strip()
    sentence = (request.form.get("sentence") or "").strip()
    synonym = (request.form.get("synonym") or "").strip()
    type_ = (request.form.get("type") or "").strip()
    base_word = (request.form.get("base_word") or "").strip()
    native_meaning = (request.form.get("native_meaning") or "").strip()

    if not word:
        flash("Field 'word' is required.", "error")
        return redirect(url_for("edit_entry_form", entry_id=entry_id))

    candidate = {
        "word": word,
        "sentence": sentence,
        "synonym": synonym,
        "type": type_,
        "base_word": base_word,
        "native_meaning": native_meaning,
    }
    candidate_key = deduplication_key(candidate)

    for other in items:
        if other.get("id") == entry_id:
            continue
        if deduplication_key(other) == candidate_key:
            flash("This exact entry already exists.", "warning")
            return redirect(url_for("edit_entry_form", entry_id=entry_id))

    target.update(candidate)
    save_items(items)
    flash("Entry updated.", "success")
    return redirect(url_for("vocabulary"))


@app.route("/delete/<entry_id>", methods=["POST"])
@login_required
def delete_entry(entry_id: str) -> Response:
    items = load_items()
    remaining = [item for item in items if item.get("id") != entry_id]

    if len(remaining) == len(items):
        flash("Entry not found.", "error")
        return redirect(url_for("vocabulary"))

    save_items(remaining)
    flash("Entry deleted.", "success")
    return redirect(url_for("vocabulary"))


@app.route("/export/csv", methods=["GET"])
@login_required
def export_csv() -> Response:
    items = get_sorted_items()

    buffer = io.StringIO()
    writer = csv.DictWriter(buffer, fieldnames=EXPORT_FIELDS, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()

    for item in items:
        writer.writerow({field: item.get(field, "") for field in EXPORT_FIELDS})

    csv_data = buffer.getvalue()
    buffer.close()

    filename = "vocab_export_" + export_timestamp() + ".csv"
    response = Response(csv_data, mimetype="text/csv; charset=utf-8")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/export/html", methods=["GET"])
@login_required
def export_html() -> Response:
    items = get_sorted_items()
    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")
    html_output = render_template("export.html", items=items, generated_at=generated_at)

    filename = f"vocab_export_{export_timestamp()}.html"
    response = Response(html_output, mimetype="text/html; charset=utf-8")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/export/json", methods=["GET"])
@login_required
def export_json() -> Response:
    items = get_sorted_items()
    payload_data = [{field: item.get(field, "") for field in EXPORT_FIELDS} for item in items]
    payload = json.dumps(payload_data, ensure_ascii=False, indent=2)

    filename = "vocab_export_" + export_timestamp() + ".json"
    response = Response(payload, mimetype="application/json; charset=utf-8")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/export/xlsx", methods=["GET"])
@login_required
def export_xlsx() -> Response:
    items = get_sorted_items()
    rows_for_export = [{field: item.get(field, "") for field in EXPORT_FIELDS} for item in items]

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Vocabulary"

    headers = list(EXPORT_FIELDS)
    sheet.append(headers)

    header_font = Font(bold=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    for cell in sheet[1]:
        cell.font = header_font
        cell.border = thin_border

    for row_data in rows_for_export:
        row = [row_data.get(field, "") for field in headers]
        sheet.append(row)
        for cell in sheet[sheet.max_row]:
            cell.border = thin_border

    for column_index, column_cells in enumerate(sheet.columns, start=1):
        max_length = 0
        for cell in column_cells:
            value = cell.value
            if value is None:
                continue
            max_length = max(max_length, len(str(value)))
        adjusted_width = min(max_length + 2, 60) if max_length else 12
        sheet.column_dimensions[get_column_letter(column_index)].width = adjusted_width

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)

    filename = "vocab_export_" + export_timestamp() + ".xlsx"
    response = Response(buffer.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/import", methods=["POST"])
@login_required
def import_words() -> Response:
    file = request.files.get("file")
    if file is None or not file.filename:
        flash("Please choose a file to import.", "error")
        return redirect(url_for("vocabulary"))

    filename = secure_filename(file.filename)
    extension = Path(filename).suffix.lower()
    if extension not in ALLOWED_IMPORT_EXTENSIONS:
        flash("Unsupported file type. Please upload a CSV, JSON, or XLSX file.", "error")
        return redirect(url_for("vocabulary"))

    data = file.read()
    if not data:
        flash("Uploaded file is empty.", "warning")
        return redirect(url_for("vocabulary"))

    try:
        parsed_rows = parse_import_payload(data, extension)
    except ValueError as exc:
        flash(str(exc), "error")
        return redirect(url_for("vocabulary"))

    if not parsed_rows:
        flash("No valid rows found in the uploaded file.", "warning")
        return redirect(url_for("vocabulary"))

    items = load_items()
    existing_keys = {deduplication_key(item) for item in items}

    new_rows: List[Dict[str, str]] = []
    duplicate_count = 0
    missing_word_count = 0

    for row in parsed_rows:
        if not row.get("word"):
            missing_word_count += 1
            continue
        key = deduplication_key(row)
        if key in existing_keys:
            duplicate_count += 1
            continue
        existing_keys.add(key)
        new_rows.append(row)

    total_rows = len(parsed_rows)
    if not new_rows:
        flash("All rows in the uploaded file already exist or are invalid.", "warning")
        session.pop(PENDING_IMPORT_SESSION_KEY, None)
        session.pop(PENDING_IMPORT_META_SESSION_KEY, None)
        return redirect(url_for("vocabulary"))

    session[PENDING_IMPORT_SESSION_KEY] = new_rows
    session[PENDING_IMPORT_META_SESSION_KEY] = {
        "total_rows": total_rows,
        "duplicate_rows": duplicate_count,
        "skipped_missing_word": missing_word_count,
        "filename": filename,
    }
    session.modified = True

    return redirect(url_for("import_preview"))


@app.route("/import/preview", methods=["GET"])
@login_required
def import_preview() -> Response | str:
    pending_rows = session.get(PENDING_IMPORT_SESSION_KEY)
    meta = session.get(PENDING_IMPORT_META_SESSION_KEY)

    if not pending_rows or not meta:
        flash("No import in progress.", "warning")
        return redirect(url_for("vocabulary"))

    total_rows = int(meta.get("total_rows", len(pending_rows)))
    duplicate_count = int(meta.get("duplicate_rows", 0))
    skipped_missing_word = int(meta.get("skipped_missing_word", 0))
    filename = meta.get("filename", "")
    preview_limit = 50
    preview_rows = pending_rows[:preview_limit]

    return render_template(
        "import_preview.html",
        filename=filename,
        total_rows=total_rows,
        new_count=len(pending_rows),
        duplicate_count=duplicate_count,
        skipped_missing_word=skipped_missing_word,
        preview_rows=preview_rows,
        preview_limit=preview_limit,
    )


@app.route("/import/confirm", methods=["POST"])
@login_required
def confirm_import() -> Response:
    pending_rows = session.get(PENDING_IMPORT_SESSION_KEY)
    if not pending_rows:
        flash("No rows pending import.", "warning")
        return redirect(url_for("vocabulary"))

    items = load_items()
    existing_keys = {deduplication_key(item) for item in items}

    imported_count = 0
    for row in pending_rows:
        key = deduplication_key(row)
        if key in existing_keys:
            continue
        existing_keys.add(key)
        created_at = stringify_value(row.get("created_at")) or datetime.now(timezone.utc).isoformat()
        entry = {
            "id": uuid4().hex,
            "word": stringify_value(row.get("word")),
            "sentence": stringify_value(row.get("sentence")),
            "synonym": stringify_value(row.get("synonym")),
            "type": stringify_value(row.get("type")),
            "base_word": stringify_value(row.get("base_word")),
            "native_meaning": stringify_value(row.get("native_meaning")),
            "created_at": created_at,
        }
        items.append(entry)
        imported_count += 1

    session.pop(PENDING_IMPORT_SESSION_KEY, None)
    session.pop(PENDING_IMPORT_META_SESSION_KEY, None)
    session.modified = True

    if imported_count:
        save_items(items)
        entry_word = "entry" if imported_count == 1 else "entries"
        flash(f"Imported {imported_count} new {entry_word}.", "success")
    else:
        flash("No new entries were imported.", "warning")

    return redirect(url_for("vocabulary"))


if __name__ == "__main__":
    ensure_store()
    app.run(host="127.0.0.1", port=5000, debug=False)
