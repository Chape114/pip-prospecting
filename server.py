#!/usr/bin/env python3
"""
Servidor local - PiP Studio Buscador de Leads
Incluye SSE (Server-Sent Events) para progreso en tiempo real.
"""

from __future__ import annotations

import json
import queue
import re
import threading
import uuid
from pathlib import Path

import requests
from flask import Flask, Response, abort, jsonify, request, send_file, stream_with_context

import scraper

APP_DIR    = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
OUTPUT_DIR = scraper.OUTPUT_DIR

app = Flask(__name__)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

SAFE_FILENAME = re.compile(r"^[\w\-.]+$")

_jobs      = {}
_jobs_lock = threading.Lock()


def build_app_config():
    return {
        "categories": [
            {"id": i, "label": cat.label}
            for i, cat in enumerate(scraper.CATEGORIES, start=1)
        ],
        "province":   scraper.PROVINCE,
        "cities": [
            {"id": i, "name": city}
            for i, city in enumerate(scraper.CITIES, start=1)
        ],
        "max_cities": scraper.MAX_CITIES,
    }


@app.get("/")
def index():
    html = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    return Response(html, mimetype="text/html; charset=utf-8")


@app.get("/api/config")
def api_config():
    return jsonify(build_app_config())


@app.post("/api/scrape/start")
def api_scrape_start():
    data = request.get_json(silent=True) or {}

    # Categoria
    category = None
    raw_id   = data.get("category_id")
    if raw_id is not None and str(raw_id).strip().isdigit():
        try:
            category = scraper.get_category(int(raw_id))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
    if category is None:
        rubro = str(data.get("rubro", "")).strip()
        if rubro:
            category = scraper.get_category_by_label(rubro)
    if category is None:
        return jsonify({"error": "Categoria invalida."}), 400

    # Ciudades
    cities   = []
    city_ids = data.get("city_ids")
    city_id  = data.get("city_id")

    if city_ids and isinstance(city_ids, list):
        if len(city_ids) > scraper.MAX_CITIES:
            return jsonify({"error": "Maximo {} ciudades.".format(scraper.MAX_CITIES)}), 400
        for cid in city_ids:
            if not str(cid).strip().isdigit():
                return jsonify({"error": "ID de ciudad invalido: {}".format(cid)}), 400
            try:
                cities.append(scraper.get_city(int(cid)))
            except ValueError as exc:
                return jsonify({"error": str(exc)}), 400
    elif city_id is not None and str(city_id).strip().isdigit():
        try:
            cities.append(scraper.get_city(int(city_id)))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400

    if not cities:
        return jsonify({"error": "Selecciona al menos una ciudad."}), 400

    # Crear job
    job_id = str(uuid.uuid4())
    q      = queue.Queue()

    with _jobs_lock:
        _jobs[job_id] = {"queue": q, "result": None, "error": None}

    def progress_cb(msg, current, total):
        pct = int((current / total) * 100) if total > 0 else 0
        q.put({"type": "progress", "message": msg, "current": current, "total": total, "pct": pct})

    def run_job():
        try:
            if len(cities) == 1:
                result = scraper.scrape_leads(
                    category, cities[0],
                    output_dir=OUTPUT_DIR,
                    verbose=False,
                    progress=progress_cb,
                )
            else:
                result = scraper.scrape_leads_multi_city(
                    category, cities,
                    output_dir=OUTPUT_DIR,
                    verbose=False,
                    progress=progress_cb,
                )
            with _jobs_lock:
                _jobs[job_id]["result"] = result
            q.put({
                "type":         "done",
                "count":        result.count,
                "emails_found": result.emails_found,
                "filename":     result.filename,
                "download_url": "/api/download/{}".format(result.filename),
                "cities":       cities,
            })
        except Exception as exc:
            with _jobs_lock:
                _jobs[job_id]["error"] = str(exc)
            q.put({"type": "error", "message": str(exc)})

    thread = threading.Thread(target=run_job, daemon=True)
    thread.start()

    return jsonify({"job_id": job_id})


@app.get("/api/scrape/progress/<job_id>")
def api_scrape_progress(job_id):
    with _jobs_lock:
        job = _jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job no encontrado."}), 404

    q = job["queue"]

    def generate():
        while True:
            try:
                event = q.get(timeout=60)
                yield "data: {}\n\n".format(json.dumps(event))
                if event.get("type") in ("done", "error"):
                    with _jobs_lock:
                        _jobs.pop(job_id, None)
                    break
            except queue.Empty:
                yield 'data: {"type":"ping"}\n\n'

    return Response(
        stream_with_context(generate()),
        mimetype="text/event-stream",
        headers={
            "Cache-Control":     "no-cache",
            "X-Accel-Buffering": "no",
        },
    )


@app.get("/api/download/<filename>")
def api_download(filename):
    if not SAFE_FILENAME.match(filename) or ".." in filename:
        abort(404)
    path = OUTPUT_DIR / filename
    if not path.is_file():
        abort(404)
    return send_file(
        path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    print("PiP Studio - Servidor de Leads")
    print("Abri en el navegador: http://127.0.0.1:5000")
    print("Detener con Ctrl+C\n")
    app.run(host="127.0.0.1", port=5000, debug=False, threaded=True)

# ---------------------------------------------------------------------------
#   python -m pip install flask requests openpyxl beautifulsoup4
#   python server.py
# ---------------------------------------------------------------------------