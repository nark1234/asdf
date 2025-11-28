from __future__ import annotations

import io
import json
import tempfile
from pathlib import Path
from typing import Dict, Optional

from flask import Flask, flash, redirect, render_template, request, send_file, url_for

from docx.enum.text import WD_ALIGN_PARAGRAPH

from generate_captions import (
    DEFAULT_STYLES,
    SIZE_MAP,
    CaptionTheme,
    ProgramLexicon,
    build_document,
    build_metadata_text,
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


def _load_json_override(field: str) -> Optional[Dict]:
    raw = request.form.get(f"{field}_json")
    if not raw:
        return None
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        flash(f"{field} JSON 문자열을 다시 확인하세요.")
        return None


def _resolve_inputs_from_form() -> Optional[Dict[str, Optional[Dict]]]:
    use_sample = bool(request.form.get("use_sample"))
    sample_payloads = _resolve_dataset(use_sample)

    analysis_data = _load_json_override("analysis") or _load_json_from_upload("analysis") or sample_payloads["analysis"]
    if not analysis_data:
        flash("분석 JSON(analysis)을 업로드하거나 샘플 데이터를 사용하세요.")
        return None

    lexicon_data = _load_json_override("lexicon") or _load_json_from_upload("lexicon") or sample_payloads["lexicon"]
    theme_data = _load_json_override("theme") or _load_json_from_upload("theme") or sample_payloads["theme"]

    return {
        "analysis": analysis_data,
        "lexicon": lexicon_data,
        "theme": theme_data,
        "filename": request.form.get("filename") or "captions.docx",
    }


def _align_label(style_info: Dict[str, str], style_name: str) -> str:
    align = style_info.get("align")
    if align in {"left", "center", "right"}:
        return align

    default_align = DEFAULT_STYLES.get(style_name, {}).get("align")
    if default_align == WD_ALIGN_PARAGRAPH.CENTER:
        return "center"
    if default_align == WD_ALIGN_PARAGRAPH.RIGHT:
        return "right"
    return "left"


def _size_label(style_info: Dict[str, str], style_name: str) -> str:
    size_key = style_info.get("size")
    if size_key:
        return size_key

    default_size = DEFAULT_STYLES.get(style_name, {}).get("size")
    for name, pt in SIZE_MAP.items():
        if pt == default_size:
            return name

    return "default"


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    inputs = _resolve_inputs_from_form()
    if not inputs:
        return redirect(url_for("index"))

    clips = parse_clips(inputs["analysis"])
    theme = CaptionTheme.from_dict(inputs["theme"])
    lexicon = ProgramLexicon.from_dict(inputs["lexicon"])
    document = build_document(clips, theme, lexicon)

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        document.save(tmp.name)
        tmp.flush()
        tmp.seek(0)
        buffer = io.BytesIO(tmp.read())

    buffer.seek(0)
    filename = inputs["filename"]
    flash(f"{len(clips)}개의 자막을 워드 파일로 생성했습니다.")
    return send_file(
        buffer,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/preview", methods=["POST"])
def preview():
    inputs = _resolve_inputs_from_form()
    if not inputs:
        return redirect(url_for("index"))

    clips = parse_clips(inputs["analysis"])
    theme = CaptionTheme.from_dict(inputs["theme"])
    lexicon = ProgramLexicon.from_dict(inputs["lexicon"])

    preview_rows = []
    for clip in clips:
        clip.notes = lexicon.annotate(clip)
        style_info = theme.resolve_style(clip)
        style_name = style_info["style_name"]
        preview_rows.append(
            {
                "text": clip.text,
                "subtitle_type": clip.subtitle_type,
                "style_name": style_name,
                "align": _align_label(style_info, style_name),
                "size": _size_label(style_info, style_name),
                "metadata": build_metadata_text(clip),
                "notes": clip.notes,
            }
        )

    return render_template(
        "preview.html",
        rows=preview_rows,
        style_overrides=theme.style_overrides,
        style_id=theme.style_id,
        filename=inputs["filename"],
        clip_count=len(clips),
        analysis_payload=inputs["analysis"],
        lexicon_payload=inputs["lexicon"],
        theme_payload=inputs["theme"],
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
