"""Menjalankan tahap M2 sampai M7 berurutan. M1 sudah dijalankan saat berkas diunggah.

Waktu tiap tahap dicatat di report["meta"]["waktu_detik"]; rule_based = parse_m2 + aturan_m3 + sitasi_m4
(semua proses lokal tanpa jaringan), total = seluruh pipeline sampai laporan tersusun.
"""
import time
from pathlib import Path

from . import m2_parser, m3_rules, m4_citations, m5_crossref, m6_annotate, m7_report


def run_pipeline(src_path, job_dir, options=None, rules_path=None, use_crossref=True, meta=None):
    options = options or {}
    job_dir = Path(job_dir)
    job_dir.mkdir(parents=True, exist_ok=True)
    clock = time.perf_counter
    t_all, t = clock(), {}

    t0 = clock()
    parsed = m2_parser.parse(src_path)                                  # M2
    t["parse_m2"] = clock() - t0

    t0 = clock()
    engine = m3_rules.RuleEngine(m3_rules.load_rules(rules_path))
    results = engine.run(parsed, options)                               # M3
    t["aturan_m3"] = clock() - t0

    t0 = clock()
    crit14 = next(c for c in engine.db["kriteria"] if c["modul"] == "M4")
    res14 = m4_citations.run(parsed, crit14, options)                   # M4
    t["sitasi_m4"] = clock() - t0

    t0 = clock()
    cr_rows = m5_crossref.run(parsed, res14, enabled=use_crossref)      # M5
    t["crossref_m5"] = clock() - t0
    results.append(res14)

    t0 = clock()
    annotated = job_dir / "berkomentar.docx"
    ann = m6_annotate.annotate(src_path, annotated, results, parsed)    # M6
    t["anotasi_m6"] = clock() - t0

    meta = {**(meta or {}), "jenis": options.get("jenis", "orisinal"),
            "crossref_aktif": use_crossref, "komentar": ann["comments"]}
    t0 = clock()
    report = m7_report.build_report(results, parsed, cr_rows, meta)     # M7
    t["laporan_m7"] = clock() - t0

    t["rule_based"] = t["parse_m2"] + t["aturan_m3"] + t["sitasi_m4"]
    t["total"] = clock() - t_all
    report["meta"]["waktu_detik"] = {k: round(v, 4) for k, v in t.items()}
    report["meta"]["crossref_referensi"] = len(cr_rows or [])
    m7_report.save_json(report, job_dir / "laporan.json")
    return report
