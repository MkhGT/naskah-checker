"""Tes pencatatan waktu eksekusi per tahap (M2/M3/M4 rule-based, M5 Crossref, M6 anotasi) dan statistiknya."""
import csv
import logging
import re
import sys
import time
from pathlib import Path

import pytest

from app import create_app
from checker import m5_crossref, m6_annotate
from checker.pipeline import run_pipeline
from conftest import make_docx

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import run_batch  # noqa: E402

KEYS = {"parse_m2", "aturan_m3", "sitasi_m4", "crossref_m5", "anotasi_m6", "laporan_m7", "rule_based", "total"}


def run(tmp_path, **kw):
    src = make_docx(tmp_path / "n.docx")
    return run_pipeline(src, tmp_path / "job", use_crossref=kw.pop("use_crossref", False), **kw)


def test_report_has_per_stage_timings(tmp_path):
    rep = run(tmp_path)
    w = rep["meta"]["waktu_detik"]
    assert set(w) == KEYS
    assert all(isinstance(v, float) and v >= 0 for v in w.values())
    assert w["parse_m2"] > 0 and w["aturan_m3"] > 0
    assert abs(w["rule_based"] - (w["parse_m2"] + w["aturan_m3"] + w["sitasi_m4"])) < 0.001
    parts = w["rule_based"] + w["crossref_m5"] + w["anotasi_m6"] + w["laporan_m7"]
    assert w["total"] >= parts - 0.002                                           # total mencakup seluruh tahap
    assert w["total"] - parts < 0.5                                              # dan tidak ada waktu besar yang tak terhitung


def test_timings_are_saved_in_laporan_json(tmp_path):
    import json
    run(tmp_path)
    saved = json.loads((tmp_path / "job" / "laporan.json").read_text(encoding="utf-8"))
    assert set(saved["meta"]["waktu_detik"]) == KEYS and saved["meta"]["crossref_referensi"] == 0


def test_crossref_time_is_attributed_to_the_crossref_stage(tmp_path, monkeypatch):
    real = m5_crossref.run

    def slow(*a, **k):
        time.sleep(0.4)
        return real(*a, **k)
    monkeypatch.setattr(m5_crossref, "run", slow)
    w = run(tmp_path)["meta"]["waktu_detik"]
    assert w["crossref_m5"] >= 0.4
    assert w["anotasi_m6"] < 0.4 and w["aturan_m3"] < 0.4                        # tidak salah catat ke tahap lain


def test_annotation_time_is_attributed_to_the_annotation_stage(tmp_path, monkeypatch):
    real = m6_annotate.annotate

    def slow(*a, **k):
        time.sleep(0.4)
        return real(*a, **k)
    monkeypatch.setattr(m6_annotate, "annotate", slow)
    w = run(tmp_path)["meta"]["waktu_detik"]
    assert w["anotasi_m6"] >= 0.4 and w["crossref_m5"] < 0.4


def test_app_logs_stage_timings_without_filename_or_content(tmp_path, caplog):
    app = create_app(jobs_dir=tmp_path / "jobs", use_crossref=False, sweep_interval=0)
    c = app.test_client()
    src = make_docx(tmp_path / "rahasia-skripsi-budi.docx")
    with caplog.at_level(logging.INFO, logger="app"), open(src, "rb") as fh:
        c.post("/periksa", data={"berkas": (fh, "rahasia-skripsi-budi.docx"), "jenis": "orisinal"},
               content_type="multipart/form-data")
    line = next(r.getMessage() for r in caplog.records if "selesai" in r.getMessage())
    assert re.search(r"permintaan=\d+\.\d{3}s rule_based=\d+\.\d{3}s crossref=\d+\.\d{3}s \(0 referensi\) anotasi=\d+\.\d{3}s", line)
    assert "rahasia" not in line and "budi" not in line


# ---------------------------------------------------------------- run_batch
def test_describe_statistics():
    d = run_batch.describe([1, 2, 3, 4, 100])
    assert d["n"] == 5 and d["median"] == 3 and d["min"] == 1 and d["maks"] == 100
    assert d["p95"] == 100                                                      # nearest-rank: ceil(0,95 x 5) = 5
    assert run_batch.describe([2.5])["sd"] == ""                                # SD tak terdefinisi untuk n = 1
    assert run_batch.describe([None, ""]) is None


@pytest.mark.parametrize("warm", [True, False])
def test_run_batch_writes_timing_columns_and_stats(tmp_path, warm):
    folder = tmp_path / "naskah"
    folder.mkdir()
    for i in range(3):
        make_docx(folder / f"n{i}.docx")
    argv = [str(folder), "--out", str(tmp_path / "hasil")] + ([] if warm else ["--tanpa-pemanasan"])
    assert run_batch.main(argv) == 0
    rows = list(csv.DictReader(open(tmp_path / "hasil" / "hasil_per_berkas.csv", encoding="utf-8-sig")))
    assert len(rows) == 3
    for r in rows:
        assert float(r["waktu_total"]) >= float(r["waktu_rule_based"]) > 0
        assert r["waktu_crossref_m5"] != "" and r["jumlah_referensi_crossref"] != ""
    stats = list(csv.DictReader(open(tmp_path / "hasil" / "evaluasi_waktu.csv", encoding="utf-8-sig")))
    semua = {r["tahap"]: r for r in stats if r["kelompok"] == "semua"}
    assert set(semua) == KEYS and all(r["n"] == "3" for r in semua.values())
    assert {r["kelompok"] for r in stats} == {"semua", "crossref_mati"}           # tanpa manifest: Crossref mati
    assert float(semua["rule_based"]["median_detik"]) > 0


def test_run_batch_keeps_old_column_order(tmp_path):
    """Kolom waktu ditambahkan di akhir; urutan kolom lama tidak bergeser (agar skrip analisis lama tetap jalan)."""
    folder = tmp_path / "n"
    folder.mkdir()
    make_docx(folder / "a.docx")
    run_batch.main([str(folder), "--out", str(tmp_path / "h")])
    head = next(csv.reader(open(tmp_path / "h" / "hasil_per_berkas.csv", encoding="utf-8-sig")))
    assert head[:8] == ["file", "basis", "tipe", "kode_cacat", "kriteria_disisipi", "crossref", "detik", "galat"]
    assert head[8:26] == [f"K{k:02d}" for k in range(1, 19)]
    assert head[26:32] == ["sesuai", "perbaikan_kecil", "tidak_sesuai", "kesalahan", "peringatan", "saran"]
    assert head[32:] == [f"waktu_{k}" for k, _ in run_batch.STAGES] + ["jumlah_referensi_crossref"]
