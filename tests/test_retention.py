"""Tes retensi: naskah yang diunggah dan seluruh hasil pemeriksaan hanya disimpan selama masa retensi.
Tes ini adalah bukti teknis bahwa tidak ada salinan yang tersisa di server oleh aplikasi
(tidak mencakup cadangan server, log server web, atau penyimpanan di luar aplikasi)."""
import io
import os
import re
import tempfile
import time
from pathlib import Path

import pytest

from app import create_app
from checker import m1_upload
from conftest import make_docx

OLD = 60 * 60 * 24            # 1 hari dalam detik


def app_client(tmp_path, **kw):
    kw.setdefault("sweep_interval", 0)
    kw.setdefault("use_crossref", False)
    app = create_app(jobs_dir=tmp_path / "jobs", **kw)
    return app, app.test_client(), tmp_path / "jobs"


def upload(client, tmp_path):
    src = make_docx(tmp_path / "naskah.docx")
    with open(src, "rb") as fh:
        r = client.post("/periksa", data={"berkas": (fh, "naskah.docx"), "jenis": "orisinal"},
                        content_type="multipart/form-data")
    assert r.status_code in (302, 303), r.status_code
    return re.search(r"/laporan/([0-9a-f]{32})", r.headers["Location"]).group(1)


def age(path, seconds):
    t = time.time() - seconds
    os.utime(path, (t, t))


def all_files(root):
    return sorted(str(p.relative_to(root)) for p in Path(root).rglob("*"))


# ---------------------------------------------------------------- hapus manual
def test_files_exist_after_check_and_report_is_reachable(tmp_path):
    """Pra-syarat: sebelum dihapus, naskah dan hasilnya memang ada (supaya tes penghapusan bermakna)."""
    _, c, jobs = app_client(tmp_path)
    job = upload(c, tmp_path)
    assert sorted(p.name for p in (jobs / job).iterdir()) == ["asli.docx", "berkomentar.docx", "laporan.json"]
    assert c.get(f"/laporan/{job}").status_code == 200


def test_delete_now_removes_everything_and_links_stop_working(tmp_path):
    _, c, jobs = app_client(tmp_path)
    job = upload(c, tmp_path)
    r = c.post(f"/hapus/{job}")
    assert r.status_code == 303 and r.headers["Location"].endswith("/?dihapus=1")
    assert list(jobs.iterdir()) == []                                           # tidak ada folder/berkas tersisa
    for url in (f"/laporan/{job}", f"/unduh/{job}/docx", f"/unduh/{job}/json", f"/unduh/{job}/asli"):
        assert c.get(url).status_code == 404, url
    assert "sudah dihapus dari server" in c.get(r.headers["Location"]).get_data(as_text=True)


def test_delete_is_idempotent_and_rejects_malformed_ids(tmp_path):
    _, c, jobs = app_client(tmp_path)
    job = upload(c, tmp_path)
    assert c.post(f"/hapus/{job}").status_code == 303
    assert c.post(f"/hapus/{job}").status_code == 303                           # klik ulang tidak error
    for bad in ("abc", "..", "0" * 31, "../" + "0" * 29):
        assert c.post(f"/hapus/{bad}").status_code == 404
    assert c.get(f"/hapus/{job}").status_code == 405                            # hanya POST (tidak terhapus lewat tautan)


def test_deleting_one_job_leaves_other_jobs_and_foreign_files_alone(tmp_path):
    _, c, jobs = app_client(tmp_path)
    a, b = upload(c, tmp_path), upload(c, tmp_path)
    (jobs / "catatan.txt").write_text("bukan job")
    (jobs / "folder-lain").mkdir()
    c.post(f"/hapus/{a}")
    assert not (jobs / a).exists() and (jobs / b).is_dir()
    assert (jobs / "catatan.txt").exists() and (jobs / "folder-lain").is_dir()


# ---------------------------------------------------------------- kedaluwarsa
def test_expired_job_is_refused_and_deleted_on_access(tmp_path):
    """Walau penyapu belum jalan, berkas kedaluwarsa tidak boleh dilayani dan langsung dihapus."""
    _, c, jobs = app_client(tmp_path, retention_minutes=30)
    job = upload(c, tmp_path)
    age(jobs / job, 31 * 60)
    assert c.get(f"/laporan/{job}").status_code == 404
    assert not (jobs / job).exists()


def test_job_inside_retention_window_is_still_served(tmp_path):
    _, c, jobs = app_client(tmp_path, retention_minutes=30)
    job = upload(c, tmp_path)
    age(jobs / job, 29 * 60)
    assert c.get(f"/laporan/{job}").status_code == 200
    assert c.get(f"/unduh/{job}/docx").status_code == 200


def test_expired_download_is_refused_too(tmp_path):
    _, c, jobs = app_client(tmp_path, retention_minutes=10)
    job = upload(c, tmp_path)
    age(jobs / job, 11 * 60)
    assert c.get(f"/unduh/{job}/docx").status_code == 404
    assert not (jobs / job).exists()


def test_purge_expired_removes_only_expired_job_dirs(tmp_path):
    jobs = tmp_path / "jobs"
    jobs.mkdir()
    old, fresh = "a" * 32, "b" * 32
    for j in (old, fresh):
        (jobs / j).mkdir()
        (jobs / j / "asli.docx").write_bytes(b"x")
    (jobs / "bukan-job").mkdir()
    (jobs / "catatan.txt").write_text("x")
    age(jobs / old, OLD)
    age(jobs / "bukan-job", OLD)
    age(jobs / "catatan.txt", OLD)
    assert m1_upload.purge_expired(jobs, 30 * 60) == 1
    assert not (jobs / old).exists() and (jobs / fresh).is_dir()
    assert (jobs / "bukan-job").is_dir() and (jobs / "catatan.txt").exists()      # bukan job: tidak disentuh


def test_startup_sweeps_leftover_jobs_from_previous_runs(tmp_path):
    jobs = tmp_path / "jobs"
    jobs.mkdir()
    leftover = "c" * 32
    (jobs / leftover).mkdir()
    (jobs / leftover / "asli.docx").write_bytes(b"naskah lama")
    age(jobs / leftover, OLD)
    create_app(jobs_dir=jobs, use_crossref=False, sweep_interval=0)
    assert not (jobs / leftover).exists()


def test_background_sweeper_deletes_without_any_request(tmp_path):
    """Tanpa ada yang mengakses: penyapu latar belakang menghapus job setelah masa retensi (periksa sistem berkas)."""
    app, c, jobs = app_client(tmp_path, retention_minutes=0.02, sweep_interval=0.1)       # 1,2 detik
    job = upload(c, tmp_path)
    assert (jobs / job).is_dir()
    deadline = time.time() + 8
    while (jobs / job).exists() and time.time() < deadline:
        time.sleep(0.2)
    assert not (jobs / job).exists()
    assert list(jobs.iterdir()) == []


# ---------------------------------------------------------------- tidak ada sisa di folder sementara
def test_no_leftover_files_in_system_temp_dir(tmp_path, monkeypatch):
    """Setelah pemeriksaan + hapus, tidak ada berkas tersisa di folder sementara OS maupun di folder job."""
    temp = tmp_path / "tmp"
    temp.mkdir()
    monkeypatch.setattr(tempfile, "tempdir", str(temp))
    _, c, jobs = app_client(tmp_path)
    job = upload(c, tmp_path)
    c.get(f"/laporan/{job}")
    c.get(f"/unduh/{job}/docx")
    c.post(f"/hapus/{job}")
    assert all_files(temp) == [], all_files(temp)
    assert all_files(jobs) == []


def test_failed_upload_leaves_nothing_behind(tmp_path):
    _, c, jobs = app_client(tmp_path)
    r = c.post("/periksa", data={"berkas": (io.BytesIO(b"bukan docx"), "x.docx"), "jenis": "orisinal"},
               content_type="multipart/form-data")
    assert r.status_code == 400
    assert all_files(jobs) == []


# ---------------------------------------------------------------- konfigurasi
def test_retention_minutes_parsing_and_env(tmp_path, monkeypatch):
    assert m1_upload.parse_retention_minutes("15") == 15
    assert m1_upload.parse_retention_minutes("0.5") == 0.5
    for bad in (None, "", "abc", "0", "-5"):
        assert m1_upload.parse_retention_minutes(bad) == m1_upload.DEFAULT_RETENTION_MINUTES    # retensi tidak bisa dimatikan
    monkeypatch.setenv("RETENTION_MINUTES", "12")
    app = create_app(jobs_dir=tmp_path / "j", use_crossref=False, sweep_interval=0)
    assert app.config["RETENTION_SECONDS"] == 12 * 60
    assert "12 menit" in app.test_client().get("/").get_data(as_text=True)


def test_report_page_shows_retention_notice_and_delete_button(tmp_path):
    _, c, _ = app_client(tmp_path, retention_minutes=30)
    job = upload(c, tmp_path)
    html = c.get(f"/laporan/{job}").get_data(as_text=True)
    assert f'action="/hapus/{job}"' in html and "Hapus sekarang" in html
    assert re.search(r"dihapus otomatis dalam\s+sekitar\s+<strong>\d+ menit", html)
