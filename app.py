from __future__ import annotations

import io
import json
import tempfile
from pathlib import Path
from typing import Dict, Optional

from flask import Flask, flash, redirect, render_template, request, send_file, url_for

from generate_captions import (
    CaptionTheme,
    ProgramLexicon,
    build_document,
    parse_clips,
)

app = Flask(__name__)
app.secret_key = "variety-caption-demo"


def _load_json_from_upload(field: str) -> Optional[Dict]:
    file = request.files.get(field)
    if not file or not file.filename:
        return None
    try:
        return json.load(file.stream)
    except json.JSONDecodeError:
        flash(f"{field} 파일이 JSON 형식이 아닙니다.")
        return None


def _load_json_from_path(path: Path) -> Optional[Dict]:
    if not path.exists():
        return None
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def _resolve_dataset(sample: bool) -> Dict[str, Optional[Dict]]:
    base = Path(__file__).parent / "sample_data"
    return {
        "analysis": _load_json_from_path(base / "analysis.json") if sample else None,
        "lexicon": _load_json_from_path(base / "lexicon.json") if sample else None,
        "theme": _load_json_from_path(base / "theme.json") if sample else None,
    }


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    use_sample = bool(request.form.get("use_sample"))
    sample_payloads = _resolve_dataset(use_sample)

    analysis_data = _load_json_from_upload("analysis") or sample_payloads["analysis"]
    if not analysis_data:
        flash("분석 JSON(analysis)을 업로드하거나 샘플 데이터를 사용하세요.")
        return redirect(url_for("index"))

    lexicon_data = _load_json_from_upload("lexicon") or sample_payloads["lexicon"]
    theme_data = _load_json_from_upload("theme") or sample_payloads["theme"]

    clips = parse_clips(analysis_data)
    theme = CaptionTheme.from_dict(theme_data)
    lexicon = ProgramLexicon.from_dict(lexicon_data)
    document = build_document(clips, theme, lexicon)

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        document.save(tmp.name)
        tmp.flush()
        tmp.seek(0)
        buffer = io.BytesIO(tmp.read())

    buffer.seek(0)
    filename = request.form.get("filename") or "captions.docx"
    flash(f"{len(clips)}개의 자막을 워드 파일로 생성했습니다.")
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
