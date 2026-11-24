"""M7 - menyusun laporan hasil pemeriksaan."""
import json
from datetime import datetime
from pathlib import Path


def build_report(results, parsed, crossref_rows, meta):
    results = sorted(results, key=lambda r: r.id)
    counts = {"error": 0, "warning": 0, "info": 0}
    for r in results:
        for f in r.findings:
            if f.para_index is not None and parsed.para(f.para_index):
                t = parsed.para(f.para_index).text
                f.excerpt = t if len(t) <= 110 else t[:107].rstrip() + "..."
            counts[f.severity] += 1
    stat = {s: sum(r.status == s for r in results) for s in ("pass", "warn", "fail", "skip")}
    total = len(results)
    cr_stat = {}
    for row in crossref_rows:
        cr_stat[row["status"]] = cr_stat.get(row["status"], 0) + 1
    return {
        "meta": {**meta, "dibuat": datetime.now().strftime("%d-%m-%Y %H:%M")},
        "ringkasan": {
            "kriteria_total": total, "sesuai": stat["pass"], "perbaikan_kecil": stat["warn"],
            "tidak_sesuai": stat["fail"], "dilewati": stat["skip"],
            "persen_sesuai": round(100 * stat["pass"] / total) if total else 0,
            "temuan": counts, "crossref": cr_stat,
            "judul": parsed.title.text if parsed.title else None,
            "jumlah_referensi": len(parsed.references),
        },
        "kriteria": [r.to_dict() for r in results],
        "crossref": crossref_rows,
    }


def save_json(report, path):
    Path(path).write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

