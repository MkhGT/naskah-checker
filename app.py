"""Aplikasi web pemeriksa naskah (Flask)."""
import json
import logging
import math
import os
import re
import threading
import time
from pathlib import Path

from flask import Flask, abort, redirect, render_template, request, send_file
from werkzeug.exceptions import RequestEntityTooLarge

from checker import m1_upload
from checker.m3_rules import load_rules
from checker.pipeline import run_pipeline

JOB_RE = re.compile(r"^[0-9a-f]{32}$")
DOWNLOADS = {"docx": ("berkomentar.docx", "naskah-berkomentar.docx"),
             "json": ("laporan.json", "laporan-pemeriksaan.json"),
             "asli": ("asli.docx", "naskah-asli.docx")}
STATUS = {"pass": ("Sesuai", "ok"), "warn": ("Perbaikan kecil", "warn"),
          "fail": ("Tidak sesuai", "fail"), "skip": ("Dilewati", "skip")}
CR_LABEL = {"terverifikasi": ("Terverifikasi", "ok"), "tahun_berbeda": ("Tahun berbeda", "warn"),
            "ditemukan_tanpa_doi": ("Ditemukan (tanpa DOI)", "ok"), "tidak_cocok": ("Judul tidak cocok", "fail"),
            "doi_tidak_terdaftar": ("DOI tidak terdaftar", "fail"), "tidak_ditemukan": ("Tidak ditemukan", "warn"),
            "error": ("Tidak terverifikasi", "skip")}


def create_app(jobs_dir=None, use_crossref=None, retention_minutes=None, sweep_interval=60.0):
    """retention_minutes: lama naskah dan hasil disimpan (bawaan 30 menit, env RETENTION_MINUTES).
    sweep_interval: jeda penyapu job kedaluwarsa dalam detik (0 = penyapu tidak dijalankan)."""
    app = Flask(__name__)
    app.logger.setLevel(logging.INFO)                           # agar log waktu per pemeriksaan tampil
    app.config["MAX_CONTENT_LENGTH"] = m1_upload.MAX_BYTES + 1024 * 1024
    app.config["JOBS_DIR"] = Path(jobs_dir or os.environ.get("JOBS_DIR") or Path(app.instance_path) / "jobs")
    app.config["USE_CROSSREF"] = (os.environ.get("CROSSREF_ENABLED", "1") != "0") if use_crossref is None else use_crossref
    app.config["JOBS_DIR"].mkdir(parents=True, exist_ok=True)
    minutes = m1_upload.parse_retention_minutes(
        retention_minutes if retention_minutes is not None else os.environ.get("RETENTION_MINUTES"))
    app.config["RETENTION_SECONDS"] = minutes * 60
    # alamat unduhan template jurnal, diisi lewat env TEMPLATE_URL (kosong = tombol tidak aktif)
    app.config["TEMPLATE_URL"] = os.environ.get("TEMPLATE_URL", "").strip()
    app.jinja_env.globals.update(STATUS=STATUS, CR_LABEL=CR_LABEL, RETENTION_MINUTES=math.ceil(minutes),
                                 TEMPLATE_URL=app.config["TEMPLATE_URL"])

    # hapus sisa job lama saat start, lalu cek berkala di latar belakang
    m1_upload.purge_expired(app.config["JOBS_DIR"], app.config["RETENTION_SECONDS"])
    if sweep_interval and sweep_interval > 0:
        def _sweeper():
            stop = threading.Event()
            while not stop.wait(sweep_interval):
                try:
                    m1_upload.purge_expired(app.config["JOBS_DIR"], app.config["RETENTION_SECONDS"])
                except Exception:                                    # satu galat tidak boleh mematikan penyapu
                    app.logger.exception("Penyapuan job kedaluwarsa gagal")
        threading.Thread(target=_sweeper, name="sapu-job", daemon=True).start()

    @app.get("/")
    def index():
        rules = load_rules()
        notice = "Naskah dan hasil pemeriksaan sudah dihapus dari server." if request.args.get("dihapus") else None
        return render_template("index.html", kriteria=rules["kriteria"], error=None, notice=notice)

    @app.post("/periksa")
    def periksa():
        t_req = time.perf_counter()
        f = request.files.get("berkas")
        jenis = request.form.get("jenis", "orisinal")
        if jenis not in ("orisinal", "review"):
            jenis = "orisinal"
        if f is None or not f.filename:
            return render_template("index.html", kriteria=load_rules()["kriteria"],
                                   error="Pilih berkas .docx terlebih dahulu."), 400
        try:                                                    # M1
            job_id, path, original = m1_upload.save_upload(f, app.config["JOBS_DIR"])
        except m1_upload.ValidationError as exc:
            return render_template("index.html", kriteria=load_rules()["kriteria"], error=str(exc)), 400
        try:                                                    # M2-M7
            report = run_pipeline(path, path.parent, {"jenis": jenis}, use_crossref=app.config["USE_CROSSREF"],
                                  meta={"nama_berkas": original, "pembuat": m1_upload.detect_producer(path)})
        except Exception:
            app.logger.exception("Pipeline gagal")
            m1_upload.purge_job(app.config["JOBS_DIR"], job_id)        # gagal: naskah tidak perlu disimpan
            return render_template("index.html", kriteria=load_rules()["kriteria"],
                                   error="Naskah tidak dapat dibaca oleh sistem. Jika dokumen diekspor dari Google Docs atau aplikasi selain Microsoft Word, buka lalu simpan ulang dengan Microsoft Word (format .docx), kemudian coba lagi."), 500
        w = report["meta"].get("waktu_detik", {})              # catat waktu per tahap (tanpa nama berkas / isi naskah)
        app.logger.info("job=%s selesai: permintaan=%.3fs rule_based=%.3fs crossref=%.3fs (%s referensi) anotasi=%.3fs",
                        job_id, time.perf_counter() - t_req, w.get("rule_based", 0), w.get("crossref_m5", 0),
                        report["meta"].get("crossref_referensi", 0), w.get("anotasi_m6", 0))
        return app.redirect(f"/laporan/{job_id}")

    def job_path(job_id, name):
        if not JOB_RE.match(job_id):
            abort(404)
        job_dir = app.config["JOBS_DIR"] / job_id
        age = m1_upload.job_age_seconds(job_dir)
        if age is not None and age > app.config["RETENTION_SECONDS"]:      # sudah kedaluwarsa: hapus, jangan dilayani
            m1_upload.purge_job(app.config["JOBS_DIR"], job_id)
            abort(404)
        p = job_dir / name
        if not p.is_file():
            abort(404)
        return p

    def remaining_minutes(job_id):
        age = m1_upload.job_age_seconds(app.config["JOBS_DIR"] / job_id)
        left = app.config["RETENTION_SECONDS"] - (age or 0)
        return max(1, math.ceil(left / 60))

    @app.get("/laporan/<job_id>")
    def laporan(job_id):
        report = json.loads(job_path(job_id, "laporan.json").read_text(encoding="utf-8"))
        return render_template("report.html", r=report, job_id=job_id, sisa_menit=remaining_minutes(job_id))

    @app.post("/hapus/<job_id>")
    def hapus(job_id):
        """Hapus naskah dan hasil pemeriksaannya sekarang."""
        if not JOB_RE.match(job_id):
            abort(404)
        m1_upload.purge_job(app.config["JOBS_DIR"], job_id)
        return redirect("/?dihapus=1", code=303)

    @app.get("/unduh/<job_id>/<kind>")
    def unduh(job_id, kind):
        if kind not in DOWNLOADS:
            abort(404)
        stored, name = DOWNLOADS[kind]
        return send_file(job_path(job_id, stored), as_attachment=True, download_name=name)

    @app.errorhandler(RequestEntityTooLarge)
    def too_large(_):
        return render_template("index.html", kriteria=load_rules()["kriteria"],
                               error=f"Ukuran berkas melebihi {m1_upload.MAX_BYTES // 1048576} MB."), 413

    return app


app = create_app()

if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG") == "1", host="127.0.0.1", port=5000)
