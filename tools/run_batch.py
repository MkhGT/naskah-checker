#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run_batch.py - jalankan sistem pemeriksa pada SEMUA .docx dalam satu folder sekaligus, lalu hitung metrik
evaluasi kuantitatif per kriteria dari varian hasil tools/make_variants.py (controlled error injection).

Pemakaian:
    python tools/run_batch.py varian                       # folder berisi varian + manifest.json
    python tools/run_batch.py varian --out hasil --workers 4
    python tools/run_batch.py varian --crossref off        # paksa Crossref mati untuk SEMUA berkas
    python tools/run_batch.py varian --lanjutkan           # lewati berkas yang laporannya sudah ada (lanjut setelah putus)
    python tools/run_batch.py varian --tanpa-pemanasan     # jangan lakukan eksekusi pemanasan (lihat catatan waktu)
    python tools/run_batch.py folder_bebas --pola "*.docx" # tanpa manifest: hanya hasil per berkas (tanpa metrik)

Crossref (bawaan --crossref auto):
    HIDUP hanya untuk berkas yang menguji kriteria sitasi & daftar pustaka (K18), yaitu varian yang memuat cacat K18
    menurut manifest (K18a-e, R01, dan varian majemuk/EA/EB yang memuat K18). Semua berkas lain (V00, K01-K17, netral)
    dijalankan dengan Crossref MATI. Tanpa manifest: berkas bernama *-K18x / *-R01 yang dihidupkan.
    Atur surel sopan santun Crossref dengan variabel lingkungan CROSSREF_MAILTO.

Keluaran (folder --out, bawaan <folder>/hasil_evaluasi):
    hasil_per_berkas.csv        status 18 kriteria per berkas + jumlah temuan + durasi + apakah Crossref dipakai
    evaluasi_per_kriteria.csv   TP/FN/FP/TN, recall, spesifisitas, presisi, F1, akurasi, ketepatan lokasi (per kriteria + total mikro + rata-rata makro)
    evaluasi_per_cacat.csv      tingkat keterbacaan per jenis cacat (K02a, K06b, ...)
    evaluasi_per_basis.csv      metrik per naskah basis (T01 ... T05) dan alarm palsu pada kontrol V00
    evaluasi_netral.csv         varian netral (N14a/b): apakah sistem memberi alarm palsu
    evaluasi_crossref.csv       varian R01: kecocokan status Crossref terhadap label cacat (hanya bila Crossref hidup)
    evaluasi_waktu.csv          statistik waktu eksekusi per tahap (rata-rata, SD, median, min, maks, P95): rule-based (M2+M3+M4),
                                Crossref (M5; juga per referensi), anotasi komentar (M6), total; dipisah Crossref mati/hidup
    detail_cacat.csv            satu baris per cacat tersisip (terbaca? lokasi tepat? temuan utama)
    alarm_lain.csv              kriteria yang berubah status padahal tidak disisipi cacat (alarm palsu / efek lintas kriteria)
    ringkasan_evaluasi.json     semua angka di atas dalam satu berkas
    laporan/<berkas>.json       laporan lengkap sistem untuk tiap berkas

Catatan waktu: tiap berkas dicatat waktu per tahap dari pipeline (laporan.json -> meta.waktu_detik) dan waktu dinding total
    ("detik", termasuk penyalinan berkas). Eksekusi PERTAMA pada sebuah proses lebih lambat (memuat leksikon/pustaka bahasa;
    pada uji ~0,25 dtk vs ~0,03 dtk untuk M3), maka secara bawaan satu eksekusi pemanasan dibuang di tiap proses sebelum pengukuran.
    Waktu Crossref bergantung jaringan: laporkan median dan P95, ulangi pada jam berbeda. Gunakan --workers 1 untuk pengukuran
    waktu yang bersih (proses paralel saling berebut CPU/jaringan).

Definisi metrik (per kriteria k, status dihitung dari temuan NON-Crossref agar K18 tidak bercampur hasil jaringan):
    terbaca (TP)  : k disisipi cacat dan status k = warn/fail          tak terbaca (FN): k disisipi cacat, status k = pass/skip
    alarm palsu (FP): k TIDAK disisipi cacat (dan bukan efek turunan yang dinyatakan operator), status k berbeda dari kontrol V00
                      basisnya; pada V00 sendiri: status k bukan pass.     TN: kebalikan FP.
    recall = TP/(TP+FN)   spesifisitas = TN/(TN+FP)   presisi = TP/(TP+FP)   F1 = 2TP/(2TP+FP+FN)   akurasi = (TP+TN)/semua
    ketepatan lokasi: temuan kriteria k menunjuk paragraf yang sama dengan lokasi cacat (hanya untuk cacat berlokasi paragraf)
Catatan: ini mengukur sistem terhadap cacat BUATAN yang diketahui; varian dari operator yang sama bukan sampel independen.
"""
import argparse
import csv
import glob
import json
import re
import math
import shutil
import statistics
import sys
import tempfile
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))

CRIT = list(range(1, 19))
TAG_RE = re.compile(r"^(?P<prefix>.+)-(?P<tag>V00|K\d{2}[a-z]|M\d{2}|R\d{2}|E[AB]|N\d{2}[a-z])$", re.I)
XREF_EXPECTED = {"R01a": {"doi_tidak_terdaftar"}, "R01b": {"tidak_cocok"}, "R01c": {"tidak_ditemukan"}}
XREF_UNDECIDED = {"error", "tidak_terverifikasi"}
XREF_DESC = {"R01a": "DOI tidak terdaftar", "R01b": "judul tidak cocok dengan DOI", "R01c": "referensi fiktif tanpa DOI"}


STAGES = [("parse_m2", "M2 parser"), ("aturan_m3", "M3 aturan"), ("sitasi_m4", "M4 sitasi"),
          ("rule_based", "RULE-BASED (M2+M3+M4)"), ("crossref_m5", "M5 Crossref"), ("anotasi_m6", "M6 anotasi komentar"),
          ("laporan_m7", "M7 laporan"), ("total", "TOTAL pipeline")]


def warmup(path):
    """Eksekusi pemanasan (hasil dibuang) agar biaya sekali-jalan (memuat leksikon dll.) tidak masuk statistik."""
    from checker.pipeline import run_pipeline
    tmp = Path(tempfile.mkdtemp(prefix="rb_warm_"))
    try:
        run_pipeline(str(path), tmp, {"jenis": "orisinal"}, use_crossref=False)
    except Exception:
        pass
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def describe(xs):
    xs = sorted(x for x in xs if x is not None and x != "")
    if not xs:
        return None
    n = len(xs)
    return {"n": n, "rata2": round(statistics.mean(xs), 4), "sd": round(statistics.stdev(xs), 4) if n > 1 else "",
            "median": round(statistics.median(xs), 4), "min": round(xs[0], 4), "maks": round(xs[-1], 4),
            "p95": round(xs[max(0, math.ceil(0.95 * n) - 1)], 4)}


def time_stats(res, meta):
    """Statistik waktu per tahap untuk kelompok semua / Crossref mati / Crossref hidup (satuan detik)."""
    groups = {"semua": list(res), "crossref_mati": [f for f in res if not meta[f]["crossref"]],
              "crossref_hidup": [f for f in res if meta[f]["crossref"]]}
    rows, out = [], {}
    for gname, files in groups.items():
        for key, label in STAGES:
            d = describe([res[f]["waktu"].get(key) for f in files if res[f].get("waktu")])
            if d:
                rows.append([gname, key, label, d["n"], d["rata2"], d["sd"], d["median"], d["min"], d["maks"], d["p95"]])
                out.setdefault(gname, {})[key] = d
        per_ref = [res[f]["waktu"]["crossref_m5"] / res[f]["n_ref"] for f in files
                   if res[f].get("waktu") and res[f].get("n_ref") and "crossref_m5" in res[f]["waktu"]]
        d = describe(per_ref) if gname == "crossref_hidup" else None
        if d:
            rows.append([gname, "crossref_per_referensi", "M5 Crossref per referensi", d["n"], d["rata2"], d["sd"], d["median"],
                         d["min"], d["maks"], d["p95"]])
            out[gname]["crossref_per_referensi"] = d
    return rows, out


# ============================================================================ penamaan & manifest
def infer_from_name(name):
    """Tanpa manifest: prefix/tipe/kode dari nama berkas (T01-K02a.docx)."""
    m = TAG_RE.match(Path(name).stem)
    if not m:
        return {"base": Path(name).stem, "type": "unknown", "tag": "", "defects": None, "expected": None}
    tag = m.group("tag").upper() if m.group("tag")[0] in "VMREvmre" else m.group("tag")
    tag = tag[0].upper() + tag[1:]
    typ = ("control" if tag == "V00" else "ref" if tag[0] == "R" else "multi" if tag[0] == "M"
           else "eff" if tag[0] == "E" else "benign" if tag[0] == "N" else "single")
    return {"base": m.group("prefix"), "type": typ, "tag": tag, "defects": None, "expected": None}


def load_manifest(path):
    if not path or not Path(path).is_file():
        return {}, {}
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    index = {}
    for b in data.get("bases", []):
        for v in b["variants"]:
            index[v["file"]] = {"base": b["prefix"], "type": v["type"], "defects": v["defects"],
                                "expected": v.get("expected"), "tag": v["file"].rsplit("-", 1)[-1].rsplit(".", 1)[0]}
    return index, data.get("catalog", {})


def wants_crossref(meta, mode):
    if mode == "on":
        return True
    if mode == "off":
        return False
    if meta["defects"] is not None and meta["type"] != "benign":
        return any(d["criterion"] == 18 for d in meta["defects"])
    return meta["type"] == "ref" or meta["tag"].upper().startswith("K18")


# ============================================================================ menjalankan sistem
def core_status(k):
    """Status kriteria dari temuan NON-Crossref (sumber M5 dikecualikan)."""
    if k["status"] == "skip":
        return "skip"
    sev = {f["severity"] for f in k["findings"] if f.get("source") != "M5"}
    return "fail" if "error" in sev else "warn" if "warning" in sev else "pass"


def run_one(job):
    """Dijalankan di proses terpisah bila --workers > 1 (harus fungsi tingkat-atas agar dapat di-pickle)."""
    path, out_dir, jenis, rules, crossref, keep_annotated = job
    from checker.pipeline import run_pipeline
    path = Path(path)
    t0 = time.perf_counter()
    tmp = Path(tempfile.mkdtemp(prefix="rb_"))
    try:
        rep = run_pipeline(str(path), tmp, {"jenis": jenis}, rules_path=rules, use_crossref=crossref,
                           meta={"nama_berkas": path.name})
        (Path(out_dir) / "laporan").mkdir(parents=True, exist_ok=True)
        shutil.copy(tmp / "laporan.json", Path(out_dir) / "laporan" / f"{path.stem}.json")
        if keep_annotated:
            (Path(out_dir) / "berkomentar").mkdir(parents=True, exist_ok=True)
            shutil.copy(tmp / "berkomentar.docx", Path(out_dir) / "berkomentar" / f"{path.stem}.docx")
        return {"file": path.name, "error": None, "seconds": round(time.perf_counter() - t0, 2), "report": rep}
    except Exception as exc:                                   # satu berkas gagal tidak boleh menghentikan batch
        return {"file": path.name, "error": f"{type(exc).__name__}: {exc}", "seconds": round(time.perf_counter() - t0, 2),
                "report": None}
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def digest(rep):
    """Ringkas laporan menjadi struktur yang dipakai evaluasi."""
    status, core, findings = {}, {}, {}
    for k in rep["kriteria"]:
        status[k["id"]] = k["status"]
        core[k["id"]] = core_status(k)
        findings[k["id"]] = [f for f in k["findings"] if f.get("source") != "M5"]
    return {"status": status, "core": core, "findings": findings, "crossref": rep.get("crossref", []),
            "summary": rep["ringkasan"], "names": {k["id"]: f"{k['code']} {k['name']}" for k in rep["kriteria"]},
            "crossref_on": rep["meta"].get("crossref_aktif", False),
            "waktu": rep["meta"].get("waktu_detik", {}), "n_ref": rep["meta"].get("crossref_referensi", 0)}


# ============================================================================ evaluasi
def ratio(a, b):
    return round(a / b, 4) if b else ""


def f1_score(tp, fn, fp):
    """F1 = 2*presisi*recall/(presisi+recall) = 2TP/(2TP+FP+FN)."""
    return ratio(2 * tp, 2 * tp + fp + fn)


def accuracy(tp, fn, fp, tn):
    return ratio(tp + tn, tp + fn + fp + tn)


def evaluate(res, meta, catalog):
    """res: file -> digest; meta: file -> info (base/type/defects). Kembalikan dict tabel hasil evaluasi."""
    ref_of = {}
    for f, m in meta.items():
        if m["type"] == "control" and f in res:
            ref_of[m["base"]] = res[f]["core"]
    conf = {k: {"TP": 0, "FN": 0, "FP": 0, "TN": 0, "TPs": 0, "Ns": 0, "TPm": 0, "Nm": 0, "FP_ctl": 0, "N_ctl": 0,
                "loc_ok": 0, "loc_n": 0} for k in CRIT}
    by_code, by_base, detail, alarms, neutral, ctl_bad = {}, {}, [], [], [], []

    def base_row(b):
        return by_base.setdefault(b, {"base": b, "variants": 0, "injected": 0, "detected": 0, "FP": 0, "TN": 0,
                                      "control_nonpass": ""})

    for f in sorted(res):
        m, d = meta[f], res[f]
        ref = ref_of.get(m["base"], {k: "pass" for k in CRIT})
        typ = m["type"]
        if typ == "control":
            bad = [k for k in CRIT if d["core"].get(k) not in ("pass",)]
            base_row(m["base"])["control_nonpass"] = ";".join(f"K{k:02d}={d['core'][k]}" for k in bad)
            for k in CRIT:
                conf[k]["N_ctl"] += 1
                if k in bad:
                    conf[k]["FP_ctl"] += 1
                    conf[k]["FP"] += 1
                    alarms.append([f, f"K{k:02d}", d["core"][k], "pass", "kontrol V00 tidak 100% sesuai",
                                   " | ".join(x["message"][:100] for x in d["findings"][k][:2])])
                else:
                    conf[k]["TN"] += 1
            continue
        if typ == "benign":
            changed = [k for k in CRIT if d["core"].get(k) != ref.get(k)]       # alarm di kriteria MANA PUN
            for dd in m["defects"] or []:
                k = dd["criterion"]
                neutral.append([f, m["base"], dd["code"], f"K{k:02d}", d["core"].get(k), ref.get(k),
                                ";".join(f"K{x:02d}" for x in changed), "YA" if changed else "tidak"])
        if typ not in ("single", "multi", "eff", "benign") or m["defects"] is None:
            continue
        injected = {dd["criterion"] for dd in m["defects"]} if typ != "benign" else set()
        side = {n for dd in m["defects"] for n in dd.get("side_effects", [])} if typ != "benign" else set()
        brow = base_row(m["base"])
        brow["variants"] += 1
        for k in CRIT:
            c = conf[k]
            if k in injected:
                if d["core"].get(k) in ("warn", "fail"):
                    c["TP"] += 1
                    c["TPs" if typ == "single" else "TPm"] += 1
                else:
                    c["FN"] += 1
                c["Ns" if typ == "single" else "Nm"] += 1
            elif k in side:
                continue
            elif d["core"].get(k) != ref.get(k):
                c["FP"] += 1
                brow["FP"] += 1
                alarms.append([f, f"K{k:02d}", d["core"].get(k), ref.get(k),
                               "varian netral" if typ == "benign" else "kriteria tidak disisipi cacat",
                               " | ".join(x["message"][:100] for x in d["findings"][k][:2])])
            else:
                c["TN"] += 1
                brow["TN"] += 1
        if typ == "benign":
            continue
        for dd in m["defects"]:
            k = dd["criterion"]
            detected = d["core"].get(k) in ("warn", "fail")
            brow["injected"] += 1
            brow["detected"] += int(detected)
            locs = [x["idx_final"] for x in dd.get("locations", []) if x.get("idx_final") is not None]
            fnd = [x for x in d["findings"][k] if x["severity"] != "info"]
            if not locs or not detected or not any(x.get("para_index") is not None for x in fnd):
                located = "n/a"
            else:
                located = any(x.get("para_index") in locs for x in fnd)
                conf[k]["loc_n"] += 1
                conf[k]["loc_ok"] += int(located)
            cr = by_code.setdefault(dd["code"], {"code": dd["code"], "criterion": k, "n": 0, "ok": 0, "ns": 0, "oks": 0})
            cr["n"] += 1
            cr["ok"] += int(detected)
            if typ == "single":
                cr["ns"] += 1
                cr["oks"] += int(detected)
            detail.append([f, m["base"], typ, dd["code"], f"K{k:02d}", "YA" if detected else "TIDAK", located,
                           d["core"].get(k), " | ".join(x["message"][:110] for x in fnd[:2])])

    crit_rows, tot = [], {x: 0 for x in ("TP", "FN", "FP", "TN", "TPs", "Ns", "TPm", "Nm", "loc_ok", "loc_n")}
    for k in CRIT:
        c = conf[k]
        for x in tot:
            tot[x] += c[x]
        crit_rows.append([f"K{k:02d}", c["TP"] + c["FN"], c["TP"], c["FN"], ratio(c["TP"], c["TP"] + c["FN"]),
                          ratio(c["TPs"], c["Ns"]), ratio(c["TPm"], c["Nm"]), c["FP"], c["TN"],
                          ratio(c["TN"], c["TN"] + c["FP"]), ratio(c["TP"], c["TP"] + c["FP"]),
                          c["FP_ctl"], ratio(c["loc_ok"], c["loc_n"]), c["loc_ok"], c["loc_n"],
                          f1_score(c["TP"], c["FN"], c["FP"]), accuracy(c["TP"], c["FN"], c["FP"], c["TN"])])
    recalls = [r[4] for r in crit_rows if r[4] != ""]
    specs = [r[9] for r in crit_rows if r[9] != ""]
    precs = [r[10] for r in crit_rows if r[10] != ""]
    f1s = [r[15] for r in crit_rows if r[15] != ""]
    accs = [r[16] for r in crit_rows if r[16] != ""]

    def mean(xs):
        return round(sum(xs) / len(xs), 4) if xs else ""
    crit_rows.append(["TOTAL (mikro)", tot["TP"] + tot["FN"], tot["TP"], tot["FN"], ratio(tot["TP"], tot["TP"] + tot["FN"]),
                      ratio(tot["TPs"], tot["Ns"]), ratio(tot["TPm"], tot["Nm"]), tot["FP"], tot["TN"],
                      ratio(tot["TN"], tot["TN"] + tot["FP"]), ratio(tot["TP"], tot["TP"] + tot["FP"]),
                      sum(conf[k]["FP_ctl"] for k in CRIT), ratio(tot["loc_ok"], tot["loc_n"]), tot["loc_ok"], tot["loc_n"],
                      f1_score(tot["TP"], tot["FN"], tot["FP"]), accuracy(tot["TP"], tot["FN"], tot["FP"], tot["TN"])])
    crit_rows.append(["RATA-RATA (makro)", "", "", "", mean(recalls), "", "", "", "",
                      mean(specs), mean(precs), "", "", "", "", mean(f1s), mean(accs)])

    code_rows = [[v["code"], f"K{v['criterion']:02d}", catalog.get(v["code"], {}).get("description", ""),
                  catalog.get(v["code"], {}).get("kind", ""), v["n"], v["ok"], ratio(v["ok"], v["n"]), v["ns"], v["oks"],
                  ratio(v["oks"], v["ns"])] for v in sorted(by_code.values(), key=lambda x: (x["criterion"], x["code"]))]
    base_rows = [[b["base"], b["variants"], b["injected"], b["detected"], ratio(b["detected"], b["injected"]), b["FP"], b["TN"],
                  ratio(b["TN"], b["TN"] + b["FP"]), b["control_nonpass"] or "0 (18/18 sesuai)"] for b in
                 sorted(by_base.values(), key=lambda x: x["base"])]

    # ---- Crossref (varian R01)
    xl, intact = {}, {}
    for f, m in meta.items():
        if m["type"] != "ref" or f not in res or not m["defects"]:
            continue
        labels = {str(k): v for k, v in (m["defects"][0].get("reference_labels") or {}).items()}
        for row in res[f]["crossref"]:
            lab = labels.get(str(row.get("num")))
            st = row.get("status")
            if lab:
                e = xl.setdefault(lab, {"n": 0, "ok": 0, "miss": 0, "undec": 0})
                e["n"] += 1
                if st in XREF_UNDECIDED:
                    e["undec"] += 1
                elif st in XREF_EXPECTED.get(lab, set()):
                    e["ok"] += 1
                else:
                    e["miss"] += 1
            else:
                intact[st] = intact.get(st, 0) + 1
    xl_rows = [[lab, XREF_DESC.get(lab, ""), "/".join(sorted(XREF_EXPECTED.get(lab, []))), e["n"], e["ok"], e["miss"], e["undec"],
                ratio(e["ok"], e["ok"] + e["miss"])] for lab, e in sorted(xl.items())]
    intact_rows = [[st, n] for st, n in sorted(intact.items())]

    summary = {"mikro": dict(zip(["n_cacat", "terbaca", "tak_terbaca", "recall", "recall_tunggal", "recall_majemuk", "FP", "TN",
                                  "spesifisitas", "presisi", "FP_kontrol", "lokasi_tepat", "lokasi_ok", "lokasi_n", "f1", "akurasi"],
                                 crit_rows[-2][1:])),
               "recall_makro": crit_rows[-1][4], "spesifisitas_makro": crit_rows[-1][9],
               "presisi_makro": crit_rows[-1][10], "f1_makro": crit_rows[-1][15], "akurasi_makro": crit_rows[-1][16],
               "kontrol_tidak_100": {b["base"]: b["control_nonpass"] for b in by_base.values() if b["control_nonpass"]}}
    return {"crit_rows": crit_rows, "code_rows": code_rows, "base_rows": base_rows, "detail": detail, "alarms": alarms,
            "neutral": neutral, "xl_rows": xl_rows, "intact_rows": intact_rows, "summary": summary}


# ============================================================================ keluaran
def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(rows)


def main(argv=None):
    ap = argparse.ArgumentParser(description="Jalankan sistem pemeriksa pada semua .docx dalam folder + hitung metrik evaluasi")
    ap.add_argument("folder", help="folder berisi naskah (.docx); idealnya keluaran make_variants.py (dengan manifest.json)")
    ap.add_argument("--out", help="folder keluaran (bawaan: <folder>/hasil_evaluasi)")
    ap.add_argument("--manifest", help="manifest.json (bawaan: <folder>/manifest.json bila ada)")
    ap.add_argument("--pola", default="*.docx", help="pola berkas dalam folder (bawaan *.docx)")
    ap.add_argument("--jenis", default="orisinal", choices=["orisinal", "review"],
                    help="jenis artikel untuk batas halaman K04 (berlaku untuk semua berkas)")
    ap.add_argument("--crossref", default="auto", choices=["auto", "on", "off"],
                    help="auto = hidup hanya untuk berkas penguji K18 (bawaan); on/off = paksa semua")
    ap.add_argument("--rules", help="rules.json alternatif")
    ap.add_argument("--workers", type=int, default=1, help="jumlah proses paralel (bawaan 1)")
    ap.add_argument("--lanjutkan", action="store_true", help="pakai laporan yang sudah ada di <out>/laporan, jangan dijalankan ulang")
    ap.add_argument("--tanpa-pemanasan", dest="pemanasan", action="store_false",
                    help="jangan jalankan eksekusi pemanasan sebelum pengukuran waktu (bawaan: pemanasan satu kali per proses)")
    ap.add_argument("--simpan-berkomentar", action="store_true", help="simpan juga dokumen berkomentar tiap berkas")
    ap.set_defaults(pemanasan=True)
    args = ap.parse_args(argv)

    folder = Path(args.folder)
    if not folder.is_dir():
        ap.error(f"folder tidak ditemukan: {folder}")
    out = Path(args.out) if args.out else folder / "hasil_evaluasi"
    out.mkdir(parents=True, exist_ok=True)
    files = [Path(f) for f in sorted(glob.glob(str(folder / args.pola))) if not Path(f).name.startswith("~$")]
    if not files:
        ap.error(f"tidak ada berkas yang cocok dengan {folder / args.pola}")

    man_path = args.manifest or (folder / "manifest.json")
    index, catalog = load_manifest(man_path)
    meta = {}
    for f in files:
        meta[f.name] = index.get(f.name) or infer_from_name(f.name)
    n_man = sum(f.name in index for f in files)
    print(f"{len(files)} berkas di {folder} ({n_man} ada di manifest"
          + ("" if index else "; TANPA manifest: hanya hasil per berkas, tanpa metrik") + ")")

    jobs, results, errors = [], {}, {}
    for f in files:
        cr = wants_crossref(meta[f.name], args.crossref)
        meta[f.name]["crossref"] = cr
        cached = out / "laporan" / f"{f.stem}.json"
        if args.lanjutkan and cached.is_file():
            rep = json.loads(cached.read_text(encoding="utf-8"))
            meta[f.name]["crossref"] = rep["meta"].get("crossref_aktif", cr)
            results[f.name] = {**digest(rep), "seconds": ""}
            continue
        jobs.append((str(f), str(out), args.jenis, args.rules, cr, args.simpan_berkomentar))
    n_cr = sum(1 for j in jobs if j[4])
    print(f"Menjalankan {len(jobs)} berkas (Crossref hidup untuk {n_cr}, mati untuk {len(jobs) - n_cr}); "
          f"{len(results)} dipakai ulang dari laporan sebelumnya")

    def collect(r, i, total):
        if r["error"]:
            errors[r["file"]] = r["error"]
            print(f"[{i}/{total}] {r['file']}: GAGAL ({r['error']})")
        else:
            results[r["file"]] = {**digest(r["report"]), "seconds": r["seconds"]}
            s = r["report"]["ringkasan"]
            print(f"[{i}/{total}] {r['file']}: {s['sesuai']}/{s['kriteria_total']} sesuai"
                  f"{'  (Crossref)' if meta[r['file']]['crossref'] else ''}  {r['seconds']}s")

    warm_src = jobs[0][0] if (jobs and args.pemanasan) else None
    if warm_src and not (args.workers > 1 and len(jobs) > 1):
        warmup(warm_src)                                          # proses utama: buang biaya eksekusi pertama
    if args.workers > 1 and len(jobs) > 1:
        with ProcessPoolExecutor(max_workers=args.workers, initializer=warmup if warm_src else None,
                                 initargs=(warm_src,) if warm_src else ()) as ex:
            for i, fut in enumerate(as_completed([ex.submit(run_one, j) for j in jobs]), 1):
                collect(fut.result(), i, len(jobs))
    else:
        for i, j in enumerate(jobs, 1):
            collect(run_one(j), i, len(jobs))

    # ---- hasil per berkas
    names = {}
    for d in results.values():
        names = d["names"] or names
    header = (["file", "basis", "tipe", "kode_cacat", "kriteria_disisipi", "crossref", "detik", "galat"]
              + [f"K{k:02d}" for k in CRIT] + ["sesuai", "perbaikan_kecil", "tidak_sesuai", "kesalahan", "peringatan", "saran"]
              + [f"waktu_{k}" for k, _ in STAGES] + ["jumlah_referensi_crossref"])
    rows = []
    for f in files:
        m, d = meta[f.name], results.get(f.name)
        defs = m["defects"] or []
        base = [f.name, m["base"], m["type"], ";".join(x["code"] for x in defs),
                ";".join(sorted({f"K{x['criterion']:02d}" for x in defs})), "ya" if m["crossref"] else "tidak"]
        if not d:
            rows.append(base + ["", errors.get(f.name, "tidak dijalankan")] + [""] * (len(header) - len(base) - 2))
            continue
        s = d["summary"]
        rows.append(base + [d["seconds"], ""] + [d["core"].get(k, "") for k in CRIT]
                    + [s["sesuai"], s["perbaikan_kecil"], s["tidak_sesuai"], s["temuan"]["error"], s["temuan"]["warning"], s["temuan"]["info"]]
                    + [d["waktu"].get(k, "") for k, _ in STAGES] + [d["n_ref"] if d["waktu"] else ""])
    write_csv(out / "hasil_per_berkas.csv", header, rows)

    summary = {"folder": str(folder), "jenis": args.jenis, "crossref_mode": args.crossref, "berkas": len(files),
               "berhasil": len(results), "gagal": errors, "pemanasan": bool(args.pemanasan), "workers": args.workers}
    trows, tstats = time_stats(results, meta)
    if trows:
        write_csv(out / "evaluasi_waktu.csv", ["kelompok", "tahap", "keterangan", "n", "rata2_detik", "sd", "median_detik", "min", "maks", "p95"], trows)
        summary["waktu_detik"] = tstats
        print("\nWaktu eksekusi (detik; median / P95)  [kelompok semua, kecuali dinyatakan lain]")
        for key, label in STAGES:
            d = tstats.get("semua", {}).get(key)
            if d:
                print(f"  {label:<26}{d['median']:>9}{d['p95']:>9}   (n={d['n']})")
        d = tstats.get("crossref_hidup", {}).get("crossref_m5")
        if d:
            print(f"  Crossref (hanya berkas Crossref hidup): median {d['median']} dtk, P95 {d['p95']} dtk, n={d['n']}")
    if index:
        ev = evaluate({f: results[f] for f in results}, meta, catalog)
        write_csv(out / "evaluasi_per_kriteria.csv",
                  ["kriteria", "cacat_disisipkan", "terbaca_TP", "tak_terbaca_FN", "recall", "recall_tunggal", "recall_majemuk",
                   "alarm_palsu_FP", "TN", "spesifisitas", "presisi", "FP_pada_kontrol", "ketepatan_lokasi", "lokasi_tepat", "lokasi_dinilai", "F1", "akurasi"],
                  ev["crit_rows"])
        write_csv(out / "evaluasi_per_cacat.csv",
                  ["kode", "kriteria", "deskripsi", "jenis_operator", "n", "terbaca", "tingkat", "n_tunggal", "terbaca_tunggal", "tingkat_tunggal"],
                  ev["code_rows"])
        write_csv(out / "evaluasi_per_basis.csv",
                  ["basis", "varian", "cacat_disisipkan", "terbaca", "recall", "alarm_palsu_FP", "TN", "spesifisitas", "kontrol_V00_tidak_sesuai"],
                  ev["base_rows"])
        write_csv(out / "detail_cacat.csv",
                  ["file", "basis", "tipe", "kode", "kriteria", "terbaca", "lokasi_tepat", "status_kriteria", "temuan_utama"], ev["detail"])
        write_csv(out / "alarm_lain.csv", ["file", "kriteria", "status", "status_kontrol", "jenis", "temuan"], ev["alarms"])
        if ev["neutral"]:
            write_csv(out / "evaluasi_netral.csv", ["file", "basis", "kode", "kriteria_sasaran", "status_sasaran", "status_kontrol", "kriteria_berubah", "alarm_palsu"], ev["neutral"])
        if ev["xl_rows"]:
            write_csv(out / "evaluasi_crossref.csv", ["label", "deskripsi", "status_diharapkan", "n", "cocok", "meleset", "tak_terverifikasi",
                                                      "tingkat_cocok"], ev["xl_rows"])
            write_csv(out / "evaluasi_crossref_referensi_utuh.csv", ["status_crossref", "jumlah"], ev["intact_rows"])
        summary.update(ev["summary"])
        summary["per_kriteria"] = [dict(zip(["kriteria", "cacat", "TP", "FN", "recall", "recall_tunggal", "recall_majemuk", "FP", "TN",
                                             "spesifisitas", "presisi", "FP_kontrol", "ketepatan_lokasi", "lokasi_ok", "lokasi_n"], r))
                                   for r in ev["crit_rows"]]
        summary["netral_alarm"] = sum(1 for r in ev["neutral"] if r[-1] == "YA")
        summary["netral_total"] = len(ev["neutral"])
        print("\nKriteria   cacat  terbaca  recall  FP   spesifisitas")
        for r in ev["crit_rows"]:
            print(f"{str(r[0]):<18}{str(r[1]):>5}{str(r[2]):>8}{str(r[4]):>8}{str(r[7]):>5}{str(r[9]):>10}")
        print(f"\nVarian netral dengan alarm palsu: {summary['netral_alarm']}/{summary['netral_total']}")
        if summary["kontrol_tidak_100"]:
            print("PERINGATAN: kontrol V00 tidak 100% sesuai (basis tercemar; hasil kriteria itu tidak sahih): "
                  + "; ".join(f"{b}: {v}" for b, v in summary["kontrol_tidak_100"].items()))
    (out / "ringkasan_evaluasi.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\nSelesai: {len(results)}/{len(files)} berkas berhasil" + (f", {len(errors)} gagal" if errors else "")
          + f". Keluaran di {out.resolve()}")
    return 1 if errors else 0


if __name__ == "__main__":
    sys.exit(main())