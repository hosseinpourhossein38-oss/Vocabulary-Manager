from __future__ import annotations

import csv
import io
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List
from uuid import uuid4

from flask import Flask, Response, flash, redirect, render_template, request, url_for
from openpyxl import Workbook
from openpyxl.styles import Border, Font, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-change-me")

DATA_DIR = Path("data")
DATA_FILE = DATA_DIR / "vocab.json"


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
            "created_at": str(item.get("created_at", "") or ""),
        }
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


@app.route("/", methods=["GET"])
def index() -> str:
    items = get_sorted_items()
    return render_template("index.html", items=items, total=len(items), edit_entry_id=None)


@app.route("/add", methods=["POST"])
def add_word() -> Response:
    word = (request.form.get("word") or "").strip()
    sentence = (request.form.get("sentence") or "").strip()
    synonym = (request.form.get("synonym") or "").strip()
    type_ = (request.form.get("type") or "").strip()
    base_word = (request.form.get("base_word") or "").strip()

    if not word:
        flash("Field 'word' is required.", "error")
        return redirect(url_for("index"))

    items = load_items()

    new_entry = {
        "id": uuid4().hex,
        "word": word,
        "sentence": sentence,
        "synonym": synonym,
        "type": type_,
        "base_word": base_word,
        "created_at": datetime.now(timezone.utc).isoformat(),
    }

    is_duplicate = any(
        existing.get("word") == new_entry["word"]
        and existing.get("sentence") == new_entry["sentence"]
        and existing.get("synonym") == new_entry["synonym"]
        and existing.get("type") == new_entry["type"]
        and existing.get("base_word") == new_entry["base_word"]
        for existing in items
    )

    if is_duplicate:
        flash("This exact entry already exists.", "warning")
        return redirect(url_for("index"))

    items.append(new_entry)
    save_items(items)
    flash("Word added successfully!", "success")
    return redirect(url_for("index"))


@app.route("/edit/<entry_id>", methods=["GET"])
def edit_entry_form(entry_id: str) -> Response | str:
    items = get_sorted_items()
    if not any(item.get("id") == entry_id for item in items):
        flash("Entry not found.", "error")
        return redirect(url_for("index"))
    return render_template("index.html", items=items, total=len(items), edit_entry_id=entry_id)


@app.route("/edit/<entry_id>", methods=["POST"])
def edit_entry(entry_id: str) -> Response:
    items = load_items()
    target = next((item for item in items if item.get("id") == entry_id), None)

    if target is None:
        flash("Entry not found.", "error")
        return redirect(url_for("index"))

    word = (request.form.get("word") or "").strip()
    sentence = (request.form.get("sentence") or "").strip()
    synonym = (request.form.get("synonym") or "").strip()
    type_ = (request.form.get("type") or "").strip()
    base_word = (request.form.get("base_word") or "").strip()

    if not word:
        flash("Field 'word' is required.", "error")
        return redirect(url_for("edit_entry_form", entry_id=entry_id))

    is_duplicate = any(
        other.get("id") != entry_id
        and other.get("word") == word
        and other.get("sentence") == sentence
        and other.get("synonym") == synonym
        and other.get("type") == type_
        and other.get("base_word") == base_word
        for other in items
    )

    if is_duplicate:
        flash("This exact entry already exists.", "warning")
        return redirect(url_for("edit_entry_form", entry_id=entry_id))

    target.update(
        {
            "word": word,
            "sentence": sentence,
            "synonym": synonym,
            "type": type_,
            "base_word": base_word,
        }
    )
    save_items(items)
    flash("Entry updated.", "success")
    return redirect(url_for("index"))


@app.route("/delete/<entry_id>", methods=["POST"])
def delete_entry(entry_id: str) -> Response:
    items = load_items()
    remaining = [item for item in items if item.get("id") != entry_id]

    if len(remaining) == len(items):
        flash("Entry not found.", "error")
        return redirect(url_for("index"))

    save_items(remaining)
    flash("Entry deleted.", "success")
    return redirect(url_for("index"))


@app.route("/export/csv", methods=["GET"])
def export_csv() -> Response:
    items = load_items()
    if not items:
        items = []

    buffer = io.StringIO()
    fieldnames = ["word", "sentence", "synonym", "type", "base_word", "created_at"]
    writer = csv.DictWriter(buffer, fieldnames=fieldnames, extrasaction="ignore", lineterminator="\n")
    writer.writeheader()

    for item in sorted(items, key=lambda entry: entry.get("created_at") or "", reverse=True):
        writer.writerow({key: item.get(key, "") for key in fieldnames})

    csv_data = buffer.getvalue()
    buffer.close()

    filename = f"vocab_export_{export_timestamp()}.csv"
    response = Response(csv_data, mimetype="text/csv; charset=utf-8")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/export/html", methods=["GET"])
def export_html() -> Response:
    items = get_sorted_items()
    generated_at = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%SZ")
    html_output = render_template("export.html", items=items, generated_at=generated_at)

    filename = f"vocab_export_{export_timestamp()}.html"
    response = Response(html_output, mimetype="text/html; charset=utf-8")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/export/json", methods=["GET"])
def export_json() -> Response:
    items = get_sorted_items()
    payload = json.dumps(items, ensure_ascii=False, indent=2)

    filename = f"vocab_export_{export_timestamp()}.json"
    response = Response(payload, mimetype="application/json; charset=utf-8")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


@app.route("/export/xlsx", methods=["GET"])
def export_xlsx() -> Response:
    items = get_sorted_items()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Vocabulary"

    headers = ["word", "sentence", "synonym", "type", "base_word", "created_at"]
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

    for item in items:
        row = [item.get(key, "") for key in headers]
        sheet.append(row)
        for cell in sheet[sheet.max_row]:
            cell.border = thin_border

    # Adjust column widths based on content length.
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

    filename = f"vocab_export_{export_timestamp()}.xlsx"
    response = Response(buffer.getvalue(), mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers["Content-Disposition"] = f"attachment; filename=\"{filename}\""
    return response


if __name__ == "__main__":
    ensure_store()
    app.run(host="127.0.0.1", port=5000, debug=False)
