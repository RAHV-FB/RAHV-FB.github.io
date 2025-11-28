"""
Simple Flask backend for IB Question Entry web app.

This backend is stateless: the browser sends all exam + question data
as JSON, and the backend returns an Excel file built from that data.
"""

from __future__ import annotations

from io import BytesIO
from typing import Any, Dict, List

from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import pandas as pd


app = Flask(__name__)
CORS(app)  # Allow calls from GitHub Pages / other origins


EXPECTED_COLUMNS = [
    "uniqueid",
    "path",
    "text_body",
    "answer_type",
    "mark_scheme",
    "needs_context",
    "exam",
    "section",
    "topic",
    "order",
    "marks",
]


def _normalise_questions(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Normalise incoming JSON into rows for the Excel export."""
    subject = (payload.get("subject") or "").strip()
    exam = (payload.get("exam") or "").strip()
    questions = payload.get("questions") or []

    rows: List[Dict[str, Any]] = []
    for idx, q in enumerate(questions, start=1):
        # Basic defensive defaults
        row = {
            "uniqueid": str(q.get("uniqueid") or f"q-{idx}"),
            "path": q.get("path") or "",
            "text_body": q.get("text_body") or "",
            "answer_type": q.get("answer_type", 0),
            "mark_scheme": q.get("mark_scheme") or "",
            "needs_context": bool(q.get("needs_context", False)),
            "exam": exam,
            "section": q.get("section") or "",
            "topic": q.get("topic") or "",
            "order": q.get("order", idx),
            "marks": q.get("marks", 0),
        }
        # Optionally store subject in topic if not provided
        if not row["topic"] and subject:
            row["topic"] = subject
        rows.append(row)

    return rows


@app.post("/export_excel")
def export_excel():
    """
    Accept JSON with exam + questions and return an .xlsx file.

    Expected JSON body:
    {
      "subject": "Computer Science",
      "exam": "M25",
      "questions": [
        {
          "uniqueid": "optional-id",
          "path": "10.a.i",
          "text_body": "Question text",
          "answer_type": 1,
          "mark_scheme": "Marks...",
          "needs_context": true,
          "section": "A",
          "topic": "",
          "order": 1,
          "marks": 4
        },
        ...
      ]
    }
    """
    try:
        payload = request.get_json(force=True, silent=False)
    except Exception:
        return jsonify({"error": "Invalid JSON body"}), 400

    if not isinstance(payload, dict):
        return jsonify({"error": "JSON body must be an object"}), 400

    exam = (payload.get("exam") or "").strip()
    if not exam:
        return jsonify({"error": "Field 'exam' is required"}), 400

    questions = payload.get("questions")
    if not isinstance(questions, list) or not questions:
        return jsonify({"error": "Field 'questions' must be a non-empty list"}), 400

    rows = _normalise_questions(payload)

    # Build DataFrame with consistent column ordering
    df = pd.DataFrame(rows)
    # Ensure columns exist and in expected order
    for col in EXPECTED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    df = df[EXPECTED_COLUMNS]

    # Write to in-memory Excel file
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="questions")

    output.seek(0)
    exam_clean = exam.replace(" ", "_")
    filename = f"{exam_clean}_questions.xlsx" if exam_clean else "questions.xlsx"

    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
def health():
    """Simple health check endpoint."""
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    # For local testing; in production use gunicorn/uvicorn/etc.
    app.run(host="0.0.0.0", port=8000, debug=True)


