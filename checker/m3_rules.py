"""M3 - Rule Engine: mengevaluasi ParsedDoc terhadap kriteria di data/rules.json.

Setiap kriteria punya fungsi pemeriksa yang dipetakan lewat field "check".
Ambang batas dan istilah dibaca dari `params`, bukan ditulis mati di kode.
"""
import json
import math
import re
from pathlib import Path

from .models import CriterionResult

DEFAULT_RULES = Path(__file__).resolve().parent.parent / "data" / "rules.json"

SECTION_LABEL = {
    "pendahuluan": "Pendahuluan", "metodologi": "Metodologi/Metode Penelitian",
    "hasil_pembahasan": "Hasil dan Pembahasan", "kesimpulan": "Kesimpulan",
    "referensi": "Referensi", "ucapan": "Ucapan Terima Kasih",
}


def load_rules(path=None):
    with open(path or DEFAULT_RULES, encoding="utf-8") as fh:
        return json.load(fh)


def _close(a, b, tol):
    return a is not None and abs(a - b) <= tol


def _words(text):
    return re.findall(r"[\w'\u2019-]+", text.lower())


# Pemeriksa per kriteria. Signature: fn(parsed, params, res, options)
def _sec_label(secs, i):
    return f"Section {i + 1}: " if len(secs) > 1 else ""


def check_paper_size(pd, prm, res, opt):
    secs = pd.page["sections_info"]
    res.note = secs[0]["paper"] if len({x["paper"] for x in secs}) == 1 else \
        "; ".join(f"section {i + 1} {x['paper']}" for i, x in enumerate(secs))
    bad = [f"{_sec_label(secs, i)}ukuran kertas {x['paper']}" for i, x in enumerate(secs) if x["paper"] != prm["paper"]]
    if bad:
        res.add("; ".join(bad) + f"; seharusnya {prm['paper']}.", suggestion="Layout > Size > A4 (semua section).")


def check_margins(pd, prm, res, opt):
    secs, tol, bad = pd.page["sections_info"], prm["tolerance_cm"], []
    for i, p in enumerate(secs):
        sides = []
        for side, label in (("top", "atas"), ("bottom", "bawah"), ("left", "kiri"), ("right", "kanan")):
            want = prm[f"{side}_cm"]
            if not _close(p[f"{side}_cm"], want, tol):
                sides.append(f"{label} {p[side + '_cm']} cm (seharusnya {want} cm)")
        if sides:
            bad.append(_sec_label(secs, i) + ", ".join(sides))
    p0 = secs[0]
    res.note = f"atas {p0['top_cm']}, bawah {p0['bottom_cm']}, kiri {p0['left_cm']}, kanan {p0['right_cm']} cm"
    if bad:
        res.add("Margin tidak sesuai: " + "; ".join(bad) + ".", suggestion="Layout > Margins > Custom Margins (semua section).")


def check_columns(pd, prm, res, opt):
    """Bagian identitas (sebelum Pendahuluan) harus 1 kolom; tata letak kolom isi dimulai dari Pendahuluan."""
    secs = pd.page["sections_info"]
    head = pd.headings[0] if pd.headings else None
    cut = head.idx if head else None
    ident = sorted({p.section_index for p in pd.paras if p.text and (cut is None or p.idx < cut)})
    body = sorted({p.section_index for p in pd.paras if p.text and cut is not None and p.idx >= cut})
    want_i, want_b = prm["identity_columns"], prm["body_columns"]
    got_i, got_b = sorted({secs[i]["columns"] for i in ident}), sorted({secs[i]["columns"] for i in body})
    fmt = lambda g: "/".join(map(str, g)) if g else "-"
    res.note = f"identitas {fmt(got_i)} kolom, isi {fmt(got_b)} kolom"
    title_idx = pd.title.idx if pd.title else None
    if set(ident) & set(body) and want_i != want_b:
        res.add("Bagian identitas dan bagian isi berada dalam satu section sehingga jumlah kolomnya sama. "
                f"Identitas harus {want_i} kolom dan isi (mulai Pendahuluan) {want_b} kolom.",
                para=head.idx if head else None,
                suggestion="Sisipkan Layout > Breaks > Continuous tepat sebelum 'Pendahuluan', lalu atur kolom "
                           "tiap bagian.")
        return
    if any(secs[i]["columns"] != want_i for i in ident):
        res.add(f"Bagian identitas memakai {fmt(got_i)} kolom; seharusnya {want_i} kolom.", para=title_idx,
                suggestion="Atur bagian identitas (judul sampai kata kunci) menjadi 1 kolom.")
    if any(secs[i]["columns"] != want_b for i in body):
        res.add(f"Bagian isi (mulai Pendahuluan) memakai {fmt(got_b)} kolom; seharusnya {want_b} kolom.",
                para=head.idx if head else None,
                suggestion=f"Atur bagian isi menjadi {want_b} kolom (Layout > Columns) setelah Continuous Section Break.")


def check_body_font(pd, prm, res, opt):
    bad_font, bad_size, bad_ls = [], [], []
    body = [p for p in pd.paras if p.role in ("body", "list") and p.words >= 3]
    for p in body:
        if (p.font or "").lower() != prm["font"].lower():
            bad_font.append((p.idx, f"Font '{p.font or 'tidak terdeteksi (font tema)'}'; seharusnya {prm['font']}."))
        if not _close(p.size_pt, prm["size_pt"], prm["size_tolerance"]):
            bad_size.append((p.idx, f"Ukuran huruf {p.size_pt} pt; seharusnya {prm['size_pt']} pt."))
        want = prm["line_spacing"]
        if isinstance(p.line_spacing, str):
            bad_ls.append((p.idx, f"Spasi baris tetap {p.line_spacing}; seharusnya kelipatan {want}."))
        elif abs(p.line_spacing - want) > prm["line_spacing_tolerance"]:
            bad_ls.append((p.idx, f"Spasi baris {p.line_spacing}; seharusnya {want}."))
    res.note = f"{len(body)} paragraf isi diperiksa"
    res.add_many(bad_font, "{n} paragraf lain juga memakai jenis huruf yang salah.",
                 suggestion=f"Gunakan {prm['font']} pada seluruh isi naskah.")
    res.add_many(bad_size, "{n} paragraf lain juga berukuran huruf salah.",
                 suggestion=f"Gunakan ukuran {prm['size_pt']} pt.")
    res.add_many(bad_ls, "{n} paragraf lain juga salah spasi baris.",
                 suggestion=f"Paragraph > Line spacing: Multiple {prm['line_spacing']}.")


def check_title(pd, prm, res, opt):
    t = pd.title
    if t is None:
        res.add("Judul artikel tidak ditemukan.", suggestion="Gunakan style 'Title' untuk judul.")
        return
    n = len(t.text.split())
    res.note = f"{n} kata"
    if n > prm["max_words"]:
        res.add(f"Judul {n} kata; maksimal {prm['max_words']} kata.", para=t.idx,
                suggestion="Ringkas judul tanpa menghilangkan unsur masalah, metode, dan hasil.")
    low = " " + " ".join(_words(t.text)) + " "
    hits = [w for w in prm["forbidden_terms"] if f" {w} " in low]
    if hits:
        res.add("Judul memuat kata yang tidak dianjurkan: " + ", ".join(hits) + ".", para=t.idx,
                suggestion="Hapus atau ganti kata tersebut.")
    f = prm.get("format")
    if f:
        bad = []
        if (t.font or "").lower() != f["font"].lower():
            bad.append(f"font '{t.font}' (seharusnya {f['font']})")
        if not _close(t.size_pt, f["size_pt"], f["size_tolerance"]):
            bad.append(f"ukuran {t.size_pt} pt (seharusnya {f['size_pt']} pt)")
        if f["bold"] and not t.bold:
            bad.append("tidak tebal")
        if t.alignment != f["alignment"]:
            bad.append(f"perataan {t.alignment} (seharusnya rata tengah)")
        if not _close(t.space_after_pt, f["space_after_pt"], f["space_tolerance_pt"]):
            bad.append(f"spasi sesudah {t.space_after_pt} pt (seharusnya {f['space_after_pt']} pt)")
        if bad:
            res.add("Format judul tidak sesuai: " + "; ".join(bad) + ".", para=t.idx, severity="warning",
                    suggestion="Gunakan style judul dari template (Cambria 16 pt, tebal, rata tengah, spasi sesudah 6 pt).")


IDENTITY_TABLE_HINT = ("Salin kotak identitas (tabel) dari template, lalu isi ulang di dalam tabel; "
                       "jangan diketik sebagai paragraf biasa.")


def _identity_in_table(pd, res, keys, severity=None):
    """Elemen identitas (Informasi Artikel, Abstrak, Kata Kunci, Penulis Korespondensi) harus berada di dalam
    tabel, sebagaimana kotak identitas pada template. Elemen yang tidak ditemukan dilaporkan oleh pemeriksa
    masing-masing, bukan di sini."""
    for key in keys:
        for e in pd.identity.get(key, []):
            if e["table_index"] is None:
                res.add(f"{e['label']} tidak berada di dalam tabel; pada template bagian ini berada di tabel identitas.",
                        para=e["idx"], suggestion=IDENTITY_TABLE_HINT, severity=severity)


ALIGN_LABEL = {"left": "rata kiri", "center": "rata tengah", "right": "rata kanan", "justify": "rata kiri-kanan (justify)"}


def _check_format(pd, idxs, fmt, res, what):
    """Periksa jenis huruf, ukuran, dan perataan paragraf pada daftar indeks (abstrak / kata kunci)."""
    bad_font, bad_size, bad_align = [], [], []
    for idx in idxs:
        p = pd.para(idx)
        if p is None or not p.text:
            continue
        if fmt.get("font") and (p.font or "").lower() != fmt["font"].lower():
            bad_font.append((idx, f"{what}: font '{p.font or 'tidak terdeteksi (font tema)'}'; seharusnya {fmt['font']}."))
        if fmt.get("size_pt") and not _close(p.size_pt, fmt["size_pt"], fmt.get("size_tolerance", 0.1)):
            bad_size.append((idx, f"{what}: ukuran huruf {p.size_pt} pt; seharusnya {fmt['size_pt']} pt."))
        if fmt.get("alignment") and p.alignment != fmt["alignment"]:
            bad_align.append((idx, f"{what}: perataan {ALIGN_LABEL.get(p.alignment, p.alignment)}; "
                                   f"seharusnya {ALIGN_LABEL.get(fmt['alignment'], fmt['alignment'])}."))
    res.add_many(bad_font, f"{{n}} paragraf {what.lower()} lain juga salah jenis huruf.",
                 suggestion=f"Gunakan {fmt.get('font')}.", severity="warning")
    res.add_many(bad_size, f"{{n}} paragraf {what.lower()} lain juga salah ukuran huruf.",
                 suggestion=f"Gunakan ukuran {fmt.get('size_pt')} pt.", severity="warning")
    res.add_many(bad_align, f"{{n}} paragraf {what.lower()} lain juga salah perataan.",
                 suggestion=f"Atur perataan: {ALIGN_LABEL.get(fmt.get('alignment'), fmt.get('alignment'))}.", severity="warning")


def _check_corresponding_lines(c, res):
    """Penulis korespondensi: label, lalu nama, afiliasi/instansi, dan email, masing-masing di baris sendiri."""
    from .m2_parser import EMAIL_RE
    hint = "Susun: 'Penulis Korespondensi:' lalu nama penulis, afiliasi/instansi, dan email, masing-masing di baris sendiri."
    lines = c.get("lines", [])
    if c.get("inline"):
        res.add("Data penulis korespondensi ditulis satu baris dengan labelnya.", para=c["idx"],
                suggestion=hint, severity="warning")
        return
    if len(lines) < 3:
        res.add(f"Penulis korespondensi hanya {len(lines)} baris; seharusnya 3 baris terpisah "
                "(nama, afiliasi/instansi, email).", para=c["idx"], suggestion=hint, severity="warning")
        return
    if not c.get("email"):
        return                                   # kekurangan email sudah dilaporkan terpisah
    for ln in lines[:-1]:
        if EMAIL_RE.search(ln["text"]):
            res.add("Email harus ditulis di baris terakhir, setelah nama dan afiliasi.", para=ln["idx"],
                    suggestion=hint, severity="warning")
            return
    last = lines[-1]
    sisa = re.sub(r"(?i)^e-?mail\s*:?\s*", "", last["text"]).strip()
    em = EMAIL_RE.search(last["text"])
    if not em or sisa != em.group(0):
        res.add("Baris terakhir harus berisi email saja, tanpa nama atau afiliasi.", para=last["idx"],
                suggestion=hint, severity="warning")


def check_abstract(pd, prm, res, opt):
    if not pd.abstract_text:
        res.add("Abstrak tidak ditemukan.", suggestion="Tambahkan abstrak (100-250 kata) tepat setelah judul/penulis.")
        return
    n = len(pd.abstract_text.split())
    res.note = f"{n} kata"
    if prm.get("require_in_table"):
        _identity_in_table(pd, res, ["abstrak"])
    if n < prm["min_words"] or n > prm["max_words"]:
        res.add(f"Abstrak {n} kata; seharusnya {prm['min_words']}-{prm['max_words']} kata.",
                para=pd.abstract_idx, suggestion="Sesuaikan panjang abstrak.")
    if not prm["allow_citations"]:
        from .m2_parser import extract_citations
        if extract_citations(pd.abstract_text):
            res.add("Abstrak memuat sitasi bernomor; abstrak tidak boleh memuat sitasi.",
                    para=pd.abstract_idx, suggestion="Hapus sitasi dari abstrak.")
    if prm.get("format"):
        _check_format(pd, pd.abstract_idxs or [pd.abstract_idx], prm["format"], res, "Isi abstrak")


def check_keywords(pd, prm, res, opt):
    kws = pd.keywords
    if not kws:
        res.add("Kata kunci tidak ditemukan.", suggestion="Cantumkan 3-5 kata kunci setelah abstrak.")
        return
    res.note = f"{len(kws)} kata kunci"
    if prm.get("require_in_table"):
        _identity_in_table(pd, res, ["kata_kunci"])
    if not prm["min"] <= len(kws) <= prm["max"]:
        res.add(f"Jumlah kata kunci {len(kws)}; seharusnya {prm['min']}-{prm['max']}.",
                para=pd.keywords_idx, suggestion="Sesuaikan jumlah kata kunci.")
    if prm["alphabetical"] and [k.lower() for k in kws] != sorted(k.lower() for k in kws):
        res.add("Kata kunci belum terurut abjad.", para=pd.keywords_idx,
                suggestion="Urutkan: " + "; ".join(sorted(kws, key=str.lower)) + ".")
    for k in kws:
        if set(_words(k)) & set(prm["forbidden_words"]):
            res.add(f"Kata kunci '{k}' memuat kata penghubung; hindari 'dan'/'dari'.",
                    para=pd.keywords_idx, severity="warning")
    if prm.get("format"):
        _check_format(pd, pd.keywords_idxs or [pd.keywords_idx], prm["format"], res, "Kata kunci")


def check_structure(pd, prm, res, opt):
    found = [h for h in pd.headings if h.level == 1]
    keys = [(h, _key(h)) for h in found]
    present = [k for _, k in keys if k]
    res.note = f"{len(found)} bagian level-1"
    for req in prm["required"]:
        if req not in present:
            res.add(f"Bagian wajib '{SECTION_LABEL[req]}' tidak ditemukan.",
                    suggestion="Tambahkan bagian dengan judul standar dan style Heading 1.")
    if prm["enforce_order"]:
        order = [k for k in prm["required"] if k in present]
        seq = list(dict.fromkeys(k for k in present if k in prm["required"]))      # tanpa duplikat
        if seq != order:
            res.add("Urutan bagian tidak sesuai: seharusnya " +
                    " > ".join(SECTION_LABEL[k] for k in prm["required"]) + ".",
                    suggestion="Susun ulang bagian naskah.")
    allowed = set(prm["required"]) | set(prm.get("optional", []))
    imrad = "Pendahuluan, Metode, Hasil dan Pembahasan, Kesimpulan, Referensi"
    for h, k in keys:                                  # bagian di luar IMRaD (mis. Tinjauan Pustaka) tidak sesuai
        if any(re.search(pt, h.text, re.I) for pt in prm.get("forbidden_patterns", [])):
            res.add(f"Bagian '{h.text}' tidak termasuk struktur IMRaD ({imrad}).", para=h.idx,
                    suggestion="Hapus bagian ini; ringkas kajian pustaka ke dalam Pendahuluan.")
        elif prm.get("reject_unknown") and k not in allowed:
            res.add(f"Bagian '{h.text}' tidak dikenali dalam struktur IMRaD ({imrad}).", para=h.idx,
                    suggestion="Gunakan judul bagian baku atau gabungkan isinya ke bagian yang sesuai.")


def _key(h):
    from .m2_parser import canonical_section
    return canonical_section(h.text)


def check_headings(pd, prm, res, opt):
    res.note = f"{len(pd.headings)} judul bagian"
    prev, fmt_items = 0, []
    for h in pd.headings:
        key = _key(h)
        if h.fake_heading:
            res.add(f"'{h.text}' tidak memakai style Heading.", para=h.idx,
                    suggestion="Terapkan style Heading 1/2/3 agar penomoran otomatis.")
        rule = prm["levels"].get(str(h.level), {})      # aturan per level; level tak terdaftar = tidak diperiksa
        is_caps = h.all_caps or h.text.isupper()
        for attr, actual, adj in (("bold", h.bold, "tebal"), ("caps", is_caps, "berhuruf kapital")):
            want = rule.get(attr)                        # true = wajib | "forbid" = dilarang | false = tidak diperiksa
            if want is True and not actual:
                res.add(f"Heading {h.level} '{h.text}' belum {adj}.", para=h.idx,
                        suggestion=f"Gunakan style Heading {h.level} dari template ({adj}).")
            elif want == "forbid" and actual:
                res.add(f"Heading {h.level} '{h.text}' tidak boleh {adj}.", para=h.idx,
                        suggestion=f"Hilangkan format {adj} pada Heading {h.level}; gunakan style Heading {h.level} dari template.")
        if prm["require_numbering"] and not h.numbered and not re.match(r"^\d+(\.\d+)*\.?\s", h.text) \
                and key not in prm["unnumbered_ok"]:
            res.add(f"Heading '{h.text}' belum bernomor urut.", para=h.idx,
                    suggestion="Aktifkan penomoran multilevel pada heading.")
        if prev and h.level > prev + 1:
            res.add(f"Level heading melompat (Heading {prev} -> Heading {h.level}) pada '{h.text}'.",
                    para=h.idx, severity="info")
        prev = h.level
        bad = []                                          # spasi paragraf dan indentasi (gantung) subjudul
        tol = prm.get("spacing_tolerance_pt", 0.5)
        for key_, label in (("space_before_pt", "sebelum"), ("space_after_pt", "sesudah")):
            want = rule.get(key_)
            have = getattr(h, key_)
            if want is not None and not _close(have, want, tol):
                bad.append(f"spasi {label} {have} pt (seharusnya {want} pt)")
        ind = prm.get("indent")
        if ind:
            left_in = (h.left_indent_cm or 0) / 2.54
            hang_in = -(h.first_indent_cm or 0) / 2.54
            t = ind.get("tolerance_in", 0.03)
            if abs(left_in - ind["left_in"]) > t or abs(hang_in - ind["hanging_in"]) > t:
                bad.append(f'indentasi kiri {left_in:.2f}" gantung {hang_in:.2f}" (seharusnya kiri {ind["left_in"]:.2f}" '
                           f'gantung {ind["hanging_in"]:.2f}", nomor rata di 0)')
        if bad:
            fmt_items.append((h.idx, f"Heading {h.level} '{h.text[:40]}': " + "; ".join(bad) + "."))
    res.add_many(fmt_items, "{n} subjudul lain juga salah spasi/indentasi.",
                 suggestion="Terapkan style Heading 1/2/3 dari template (spasi 6/6, 6/3, 6/2 pt; nomor rata di 0, "
                            "teks menjorok 0,25 in dengan gantung 0,25 in).")


def check_intro_paragraphs(pd, prm, res, opt):
    paras = [p for p in pd.section_paras("pendahuluan") if p.words >= prm["min_words_per_paragraph"]]
    n = len(paras)
    res.note = f"{n} paragraf"
    head = next((h for h in pd.headings if _key(h) == "pendahuluan"), None)
    if head is None:
        res.add("Bagian Pendahuluan tidak ditemukan.", suggestion="Tambahkan bagian Pendahuluan.")
        return
    if n < prm["min"]:
        res.add(f"Pendahuluan hanya {n} paragraf; minimal {prm['min']} (latar belakang, state of the art, gap & tujuan).",
                para=head.idx, severity="error")
    elif n > prm["max"]:
        res.add(f"Pendahuluan {n} paragraf; disarankan {prm['min']}-{prm['max']} paragraf.",
                para=head.idx, severity="warning")


def _prev_text_para(pd, idx):
    for p in reversed(pd.paras[:idx]):
        if p.text or p.has_drawing:
            return p
    return None


def _next_text_para(pd, idx):
    for p in pd.paras[idx + 1:]:
        if p.text or p.has_drawing:
            return p
    return None


def _table_content_items(pd, cfg):
    """Isi tabel (di luar kotak identitas): font, ukuran, perataan, dan baris pertama tebal. Satu temuan per jenis
    pelanggaran per tabel, dengan jumlah sel yang menyimpang."""
    items = []
    body_tables = [t for t in pd.tables if not t["front"]]
    for n, t in enumerate(body_tables, 1):
        cells = [p for p in pd.paras if p.table_index == t["index"] and p.text]
        if not cells:
            continue
        bad_font = [p for p in cells if cfg.get("font") and (p.font or "").lower() != cfg["font"].lower()]
        bad_size = [p for p in cells if cfg.get("size_pt") and not _close(p.size_pt, cfg["size_pt"], cfg.get("size_tolerance", 0.1))]
        bad_al = [p for p in cells if cfg.get("alignment") and p.alignment != cfg["alignment"]]
        bad_head = [p for p in cells if cfg.get("header_bold") and p.table_row == 0 and not p.bold]
        if bad_font:
            items.append((bad_font[0].idx, f"Tabel ke-{n}: {len(bad_font)} sel tidak memakai font {cfg['font']}."))
        if bad_size:
            items.append((bad_size[0].idx, f"Tabel ke-{n}: {len(bad_size)} sel tidak berukuran {cfg['size_pt']} pt "
                                           f"(contoh: {bad_size[0].size_pt} pt)."))
        if bad_al:
            items.append((bad_al[0].idx, f"Tabel ke-{n}: {len(bad_al)} sel tidak rata tengah."))
        if bad_head:
            items.append((bad_head[0].idx, f"Tabel ke-{n}: teks pada baris pertama (judul kolom) harus tebal."))
    return items


def check_captions(pd, prm, res, opt):
    caps = {c["idx"]: c for c in pd.captions}
    items = []
    # gambar: caption harus di bawah
    for f in pd.figures:
        nxt, prv = _next_text_para(pd, f.idx), _prev_text_para(pd, f.idx)
        if f.alignment != "center":
            items.append((f.idx, "Gambar tidak rata tengah."))
        if nxt is not None and nxt.idx in caps and caps[nxt.idx]["kind"] == "gambar":
            continue
        if prv is not None and prv.idx in caps and caps[prv.idx]["kind"] == "gambar":
            items.append((prv.idx, "Caption gambar harus berada di bawah gambar."))
        else:
            items.append((f.idx, "Gambar tidak memiliki caption 'Gambar N. Judul' di bawahnya."))
    # tabel (selain kotak informasi di awal naskah): caption di atas
    for t in (t for t in pd.tables if not t["front"]):
        prv, nxt = _prev_text_para(pd, t["first_idx"]), _next_text_para(pd, t["last_idx"])
        if t.get("alignment") != "center":
            items.append((t["first_idx"], "Tabel tidak rata tengah."))
        if prv is not None and prv.idx in caps and caps[prv.idx]["kind"] == "tabel":
            continue
        if nxt is not None and nxt.idx in caps and caps[nxt.idx]["kind"] == "tabel":
            items.append((nxt.idx, "Caption tabel harus berada di atas tabel."))
        else:
            items.append((t["first_idx"], "Tabel tidak memiliki caption 'Tabel N. Judul' di atasnya."))
    for c in pd.captions:
        if c["alignment"] != "center":
            items.append((c["idx"], "Caption tidak rata tengah."))
        if not c["well_formed"]:
            items.append((c["idx"], "Format caption harus 'Gambar N. Judul' / 'Tabel N. Judul' (titik setelah nomor)."))
        if not _close(c["size_pt"], prm["caption_size_pt"], prm["size_tolerance"]):
            items.append((c["idx"], f"Ukuran caption {c['size_pt']} pt; seharusnya {prm['caption_size_pt']} pt."))
    tc_items = _table_content_items(pd, prm["table_content"]) if prm.get("table_content") else []
    for kind in ("gambar", "tabel"):
        nums = [c["number"] for c in pd.captions if c["kind"] == kind]
        if nums and nums != list(range(1, len(nums) + 1)):
            items.append((None, f"Penomoran {kind} tidak berurutan dari 1: {nums}."))
    res.note = f"{len(pd.figures)} gambar, {len([t for t in pd.tables if not t['front']])} tabel"
    seen = set()
    uniq = [i for i in items if not (i in seen or seen.add(i))]
    res.add_many(uniq, "{n} temuan caption lain tidak ditampilkan.")
    res.add_many(tc_items, "{n} temuan isi tabel lain tidak ditampilkan.", limit=8)


def check_figure_citations(pd, prm, res, opt):
    body = [p for p in pd.paras if p.role in ("body", "list") or
            (p.role == "table" and not any(t["front"] and t["index"] == p.table_index for t in pd.tables))]
    for c in pd.captions:
        word = r"(?:Gambar|Gbr\.?)" if c["kind"] == "gambar" else r"Tabel"
        rx = re.compile(rf"\b{word}\s*{c['number']}\b", re.I)
        if not any(rx.search(p.text) for p in body if p.idx != c["idx"]):
            label = f"{'Gambar' if c['kind'] == 'gambar' else 'Tabel'} {c['number']}"
            res.add(f"{label} tidak dirujuk di dalam teks naskah.", para=c["idx"],
                    suggestion=f"Rujuk dalam kalimat, mis. 'seperti pada {label}'.")
    res.note = f"{len(pd.captions)} caption"


def _check_blank_space(pd, cfg, res):
    """Spasi/baris kosong yang tidak perlu pada paragraf isi."""
    dbl,  empty = [], []
    for p in pd.paras:
        if p.role not in ("body", "list") or not p.text:
            continue
        t = p.raw_text or p.text
        if cfg.get("double_space") and re.search(r"(?<=\S) {2,}(?=\S)", t):
            dbl.append((p.idx, "Terdapat spasi ganda di dalam paragraf."))
    if cfg.get("empty_paragraphs"):
        paras = pd.paras
        i = 0
        while i < len(paras):
            p = paras[i]
            ok = (p.table_index is None and not p.text and not p.has_drawing and not p.has_break
                  and p.section is not None and p.role != "heading")
            if not ok:
                i += 1
                continue
            j = i
            while j + 1 < len(paras) and paras[j + 1].table_index is None and not paras[j + 1].text \
                    and not paras[j + 1].has_drawing and not paras[j + 1].has_break:
                j += 1
            prev = paras[i - 1] if i > 0 else None
            nxt = paras[j + 1] if j + 1 < len(paras) else None
            akhir_section = nxt is not None and nxt.section_index > paras[j].section_index
            setelah_tabel = prev is not None and prev.table_index is not None
            if not akhir_section and not setelah_tabel and nxt is not None:
                n = j - i + 1
                empty.append((p.idx, f"Terdapat {n} baris kosong; atur jarak dengan Spacing Before/After pada "
                                     "Paragraph, bukan dengan menekan Enter."))
            i = j + 1
    res.add_many(dbl, "{n} paragraf lain juga memuat spasi ganda.", suggestion="Hapus spasi berlebih (Ctrl+H: dua spasi diganti satu).")
    res.add_many(empty, "{n} baris kosong lain tidak ditampilkan.", suggestion="Hapus baris kosong; atur Spacing After pada paragraf.")
    return len(dbl) + len(empty)


def check_body_paragraph(pd, prm, res, opt):
    bad_ind, bad_al = [], []
    body = [p for p in pd.paras if p.role == "body" and p.words >= prm["min_words"]]
    for p in body:
        if not _close(p.first_indent_cm, prm["first_line_indent_cm"], prm["indent_tolerance_cm"]):
            bad_ind.append((p.idx, f"Baris pertama menjorok {p.first_indent_cm or 0} cm; seharusnya {prm['first_line_indent_cm']} cm."))
        if p.alignment != prm["alignment"]:
            bad_al.append((p.idx, "Paragraf belum rata kiri-kanan (justify)."))
    res.note = f"{len(body)} paragraf diperiksa"
    res.add_many(bad_ind, "{n} paragraf lain juga salah indentasi.", suggestion="Paragraph > Special > First line 1,27 cm.")
    res.add_many(bad_al, "{n} paragraf lain juga belum justify.", suggestion="Ctrl+J.")
    if prm.get("foreign_italic"):
        res.note += "; " + _foreign_italic(pd, prm["foreign_italic"], res, opt)
    if prm.get("blank_space"):
        n = _check_blank_space(pd, prm["blank_space"], res)
        res.note += f"; {n} temuan spasi kosong"


def _sentences(text):
    return [s for s in re.split(r"(?<=[.!?])\s+", text) if len(s.split()) >= 5]


def check_conclusion(pd, prm, res, opt):
    paras = pd.section_paras("kesimpulan")
    head = next((h for h in pd.headings if _key(h) == "kesimpulan"), None)
    if not paras:
        res.status = "skip" if head is None else res.status
        if head is not None:
            res.add("Bagian Kesimpulan kosong.", para=head.idx)
        return
    concl = " ".join(p.text for p in paras)
    if prm["require_suggestion"] and not re.search(
            r"\b(saran|rekomendasi|penelitian (?:selanjutnya|lanjutan|mendatang|berikutnya)|future work)\b", concl, re.I):
        res.add("Kesimpulan belum memuat saran untuk penelitian selanjutnya.", para=head.idx if head else None,
                suggestion="Tambahkan saran/prospek penelitian lanjutan.")
    if not pd.abstract_text:
        return
    abs_s = [set(_words(s)) for s in _sentences(pd.abstract_text)]
    for p in paras:
        for s in _sentences(p.text):
            ws = set(_words(s))
            if any(a and len(ws & a) / len(ws | a) >= prm["sentence_similarity"] for a in abs_s):
                res.add("Kalimat kesimpulan hampir identik dengan kalimat pada abstrak.", para=p.idx,
                        suggestion="Tulis ulang; kesimpulan tidak boleh mengulang abstrak.")
                break
    g = lambda ws: {tuple(ws[i:i + 4]) for i in range(len(ws) - 3)}
    c4, a4 = g(_words(concl)), g(_words(pd.abstract_text))
    ratio = len(c4 & a4) / len(c4) if c4 else 0
    res.note = f"kemiripan frasa dengan abstrak {ratio:.0%}"
    if ratio >= prm["overlap_ratio"] and not any("identik" in f.message for f in res.findings):
        res.add(f"{ratio:.0%} frasa kesimpulan sama dengan abstrak.", para=paras[0].idx,
                suggestion="Parafrase dan fokus pada temuan, kontribusi, implikasi, serta keterbatasan.")


def check_page_count(pd, prm, res, opt):
    kind = opt.get("jenis", "orisinal")
    limit = prm["max_pages"].get(kind, prm["max_pages"]["orisinal"])
    if pd.pages_meta:
        pages, how = pd.pages_meta, "metadata berkas"
    else:
        pages, how = math.ceil(pd.total_words / prm["words_per_page_estimate"]), "estimasi dari jumlah kata"
    res.note = f"{pages} halaman ({how}); batas {limit}"
    if pages > limit:
        res.add(f"Naskah sekitar {pages} halaman ({how}); batas {limit} halaman untuk artikel "
                f"{'penelitian orisinal' if kind == 'orisinal' else 'review/survei'}.",
                suggestion="Padatkan naskah. Buka dan simpan ulang di Word agar jumlah halaman akurat.")
    elif not pd.pages_meta:
        res.add("Jumlah halaman hanya estimasi (metadata tidak tersedia).", severity="info", comment=False)


def check_authors(pd, prm, res, opt):
    """K07: baris penulis, afiliasi (ukuran, perataan, penanda angka), penulis korespondensi."""
    tol = prm["size_tolerance"]
    a = pd.authors
    n_aff = pd.affiliation_count
    res.note = f"{len(pd.author_names)} penulis, {n_aff} afiliasi"
    if a is None:
        res.add("Baris nama penulis tidak ditemukan di bawah judul.",
                suggestion="Tulis nama penulis (rata tengah, 10 pt) tepat di bawah judul.")
    else:
        if (a.font or "").lower() != prm["font"].lower():
            res.add(f"Font nama penulis '{a.font}'; seharusnya {prm['font']}.", para=a.idx, severity="warning")
        if not _close(a.size_pt, prm["author_size_pt"], tol):
            res.add(f"Ukuran huruf nama penulis {a.size_pt} pt; seharusnya {prm['author_size_pt']} pt.",
                    para=a.idx, severity="warning")
        if a.alignment != prm["alignment"]:
            res.add("Baris nama penulis harus rata tengah.", para=a.idx, severity="warning")
        if not _close(a.space_after_pt, prm["space_after_pt"], prm["space_tolerance_pt"]):
            res.add(f"Spasi sesudah nama penulis {a.space_after_pt} pt; seharusnya {prm['space_after_pt']} pt.",
                    para=a.idx, severity="warning")
    if not pd.affiliations:
        res.add("Afiliasi penulis tidak ditemukan di bawah nama penulis.", para=a.idx if a else None,
                suggestion="Tulis afiliasi (program studi, institusi) rata tengah 9 pt dengan penanda angka.")
    for f in pd.affiliations:
        if not _close(f.size_pt, prm["affiliation_size_pt"], tol):
            res.add(f"Ukuran huruf afiliasi {f.size_pt} pt; seharusnya {prm['affiliation_size_pt']} pt.",
                    para=f.idx, severity="warning")
        if f.alignment != prm["alignment"]:
            res.add("Afiliasi harus rata tengah.", para=f.idx, severity="warning")
        if (f.font or "").lower() != prm["font"].lower():
            res.add(f"Font afiliasi '{f.font}'; seharusnya {prm['font']}.", para=f.idx, severity="warning")
        if not _close(f.space_after_pt, prm["space_after_pt"], prm["space_tolerance_pt"]):
            res.add(f"Spasi sesudah afiliasi {f.space_after_pt} pt; seharusnya {prm['space_after_pt']} pt.",
                    para=f.idx, severity="warning")

    # penanda angka penulis <-> afiliasi
    if a is not None and pd.affiliations:
        used = {m for n in pd.author_names for m in n["markers"]}
        listed = set(pd.affiliation_markers)
        if n_aff >= 2:
            for n in pd.author_names:
                if not n["markers"]:
                    res.add(f"Penulis '{n['name']}' belum diberi penanda afiliasi (angka superskrip).", para=a.idx,
                            severity="warning", suggestion="Beri angka superskrip yang sesuai dengan nomor afiliasi.")
        for m in sorted(used - listed):
            res.add(f"Penanda afiliasi {m} dipakai pada nama penulis, tetapi tidak ada afiliasi bernomor {m}.",
                    para=pd.affiliations[0].idx, severity="error")
        for m in sorted(listed - used):
            res.add(f"Afiliasi bernomor {m} tidak dipakai oleh satu pun penulis.", para=pd.affiliations[0].idx,
                    severity="warning")

    # penulis korespondensi
    c = pd.corresponding
    if prm["require_corresponding"]:
        if c is None:
            res.add("Penulis korespondensi tidak dicantumkan.",
                    suggestion="Isi kolom 'Penulis Korespondensi' (nama, afiliasi, email).")
        elif not c["email"]:
            res.add("Penulis korespondensi belum mencantumkan alamat email.", para=c["idx"],
                    suggestion="Tambahkan 'Email: nama@institusi.ac.id'.")

    if c is not None and prm.get("corresponding_lines"):
        _check_corresponding_lines(c, res)

    # letak di tabel identitas
    _identity_in_table(pd, res, prm.get("identity_in_table", []))
    if prm.get("require_article_info") and not pd.identity.get("informasi"):
        res.add("Blok 'Informasi Artikel' (riwayat artikel) tidak ditemukan.", severity="warning",
                suggestion="Pertahankan kotak 'Informasi Artikel' dari template; kolom tanggal diisi editor.")
    if prm.get("identity_single_table"):
        found = [e for k in ("informasi", "abstrak", "kata_kunci", "korespondensi") for e in pd.identity.get(k, [])
                 if e["table_index"] is not None]
        if found:
            ref = found[0]["table_index"]
            for e in [e for e in found if e["table_index"] != ref][:3]:
                res.add(f"{e['label']} berada di tabel yang berbeda dari elemen identitas lain; template memakai satu "
                        "tabel identitas.", para=e["idx"], severity="warning", suggestion=IDENTITY_TABLE_HINT)

    # sisa teks bawaan template
    blob = [(a.idx, a.text)] if a else []
    blob += [(f.idx, f.text) for f in pd.affiliations]
    if c:
        blob.append((c["idx"], c["text"]))
    for idx, text in blob:
        hits = [ph for ph in prm["placeholders"] if ph.lower() in text.lower()]
        if hits:
            res.add("Masih memuat teks bawaan template: " + ", ".join(f"'{h}'" for h in hits) + ".", para=idx,
                    suggestion="Ganti dengan data penulis yang sebenarnya.")


def _flat(text):
    return re.sub(r"[^a-z0-9]", "", (text or "").lower())


HF_LABEL = {"header_first": "Header halaman pertama", "header_odd": "Header halaman ganjil",
            "header_even": "Header halaman genap", "footer_first": "Footer halaman pertama",
            "footer_odd": "Footer halaman ganjil", "footer_even": "Footer halaman genap"}


def check_header_footer(pd, prm, res, opt):
    """K05: pembedaan halaman pertama / ganjil / genap, isi wajib, nomor halaman, sisa teks template."""
    fl, hf = pd.hf_flags, pd.headers_footers
    if prm["require_different_first_page"] and not fl.get("different_first"):
        res.add("Opsi 'Different First Page' tidak aktif; header/footer halaman pertama tidak akan tampil berbeda.",
                suggestion="Layout > Page Setup > Layout > centang 'Different first page'.")
    if prm["require_odd_even"] and not fl.get("odd_even"):
        res.add("Opsi 'Different Odd & Even Pages' tidak aktif; header/footer halaman genap diabaikan.",
                suggestion="Layout > Page Setup > Layout > centang 'Different odd and even'.")
    for key, label in HF_LABEL.items():
        r = hf[key]
        needs = prm["contains"].get(key, [])
        if not r["defined"] or (not r["text"] and not r["page_field"]):
            if needs or key in prm["require_page_field"]:
                res.add(f"{label} kosong atau tidak terdefinisi.",
                        suggestion="Pulihkan header/footer dari template jurnal.")
            continue
        if key in prm.get("require_table", []) and not r["in_table"]:
            where = ("sebagian isinya (teks atau nomor halaman) berada di luar tabel" if r["has_table"]
                     else "tidak berada di dalam tabel")
            res.add(f"{label} {where}; template menempatkan isi header/footer di dalam tabel.",
                    suggestion="Gunakan header/footer bawaan template (salin tabelnya), jangan mengetik ulang sebagai paragraf biasa.")
        for tok in needs:
            if _flat(tok) not in _flat(r["text"]):
                res.add(f"{label} tidak memuat '{tok}'.", suggestion="Gunakan header/footer bawaan template.")
        if key in prm["require_page_field"] and not r["page_field"]:
            res.add(f"{label} tidak memuat nomor halaman otomatis (field PAGE).",
                    suggestion="Insert > Page Number, bukan mengetik angka manual.")
        if key in prm["require_text_beyond_page"] and not r["content"]:
            res.add(f"{label} hanya berisi nomor halaman; seharusnya memuat judul singkat dan nama penulis.",
                    suggestion="Isi judul singkat dan nama penulis pada footer halaman genap.")
        for ph in prm["placeholders"].get(key, []):
            if ph.lower() in r["text"].lower():
                res.add(f"{label} masih memuat teks bawaan template ('{ph}').",
                        suggestion="Ganti dengan judul singkat dan nama penulis naskah ini.", severity="warning")
                break
        for ph in prm["editor_fields"].get(key, []):
            if ph.lower() in r["text"].lower():
                res.add(f"{label}: kolom volume/nomor/halaman/bulan-tahun diisi oleh editor; abaikan jika naskah belum terbit.",
                        severity="info", comment=False)
                break
    res.note = ("halaman pertama berbeda" if fl.get("different_first") else "halaman pertama sama") + \
               ", " + ("ganjil-genap berbeda" if fl.get("odd_even") else "ganjil-genap sama")


def _italic_coverage(p, start, end):
    """Bagian karakter bukan-spasi pada [start, end) yang bercetak miring (0..1)."""
    mask = bytearray(len(p.raw_text))
    for a, b in p.italic_ranges:
        mask[a:b] = b"\x01" * (b - a)
    chars = [k for k in range(start, end) if not p.raw_text[k].isspace()]
    return sum(mask[k] for k in chars) / len(chars) if chars else 1.0


def foreign_kwargs(pd, prm):
    """Argumen langid.foreign_spans yang berlaku untuk dokumen ini (dipakai M3 dan tools/make_variants.py)."""
    from . import langid
    texts = [p.text for p in pd.paras if p.role not in ("heading", "title") and p.text]
    return dict(allowlist=prm["allowlist"], force_terms=prm["force_terms"], extra_patterns=prm["extra_patterns"],
                title_case_phrases=prm["title_case_phrases"], sentence_detection=prm["sentence_detection"],
                abbreviations=prm.get("abbreviations", []), doc_acronyms=langid.document_acronyms(texts))


def _foreign_italic(pd, prm, res, opt):
    """Bagian K14: pada naskah berbahasa Indonesia, istilah asing di badan teks dicetak miring.
    Menambahkan temuan ke `res` dan mengembalikan ringkasan teks. Tidak pernah mengubah status K14 menjadi
    'dilewati': pada naskah berbahasa asing hanya bagian miring ini yang dilewati, indentasi/perataan tetap diperiksa."""
    from . import langid
    body = [p for p in pd.paras if p.role in ("body", "list") and p.words >= 3 and p.raw_text]
    forced = opt.get("bahasa")
    if forced:
        lang, conf, how = forced, 1.0, "pilihan pengguna"
    else:
        lang, conf, how = langid.detect_language(" ".join(p.raw_text for p in body))
    if lang != prm["apply_to_language"]:
        return (f"istilah asing miring dilewati (naskah berbahasa {lang or 'tidak terdeteksi'}, {how}; "
                "aturan hanya untuk naskah berbahasa Indonesia)")
    if not langid.lexicon()[0]:
        res.add("Leksikon istilah asing (data/foreign_en_*.txt) tidak ditemukan; hanya pola dan kalimat berbahasa "
                "Inggris yang diperiksa.", severity="info", comment=False)
    total, missing, items = 0, 0, []
    kw = foreign_kwargs(pd, prm)
    for p in body:
        spans = langid.foreign_spans(p.raw_text, **kw)
        bad = []
        for s, e, kind in spans:
            total += 1
            if _italic_coverage(p, s, e) < prm["min_italic_ratio"]:
                bad.append(p.raw_text[s:e].strip())
        if bad:
            missing += len(bad)
            shown = ", ".join(f"\u2018{t[:40]}\u2019" for t in bad[:5]) + (f" (+{len(bad) - 5} lainnya)" if len(bad) > 5 else "")
            items.append((p.idx, f"Istilah asing belum dicetak miring: {shown}."))
    pct = f"{(total - missing) / total:.0%}" if total else "-"
    res.add_many(items, "{n} paragraf lain juga memuat istilah asing yang belum dicetak miring.",
                 suggestion="Cetak miring (Ctrl+I) pada istilah asing. Deteksi bersifat otomatis; abaikan jika istilah "
                            "sudah baku (terdaftar di KBBI) atau berupa nama diri.")
    return f"{total} istilah asing, {total - missing} sudah miring ({pct}); bahasa naskah {lang} ({how})"


CHECKS = {
    "paper_size": check_paper_size, "margins": check_margins, "columns": check_columns,
    "body_font": check_body_font, "title": check_title, "abstract": check_abstract,
    "keywords": check_keywords, "structure": check_structure, "headings": check_headings,
    "intro_paragraphs": check_intro_paragraphs, "captions": check_captions,
    "figure_citations": check_figure_citations, "body_paragraph": check_body_paragraph,
    "conclusion": check_conclusion, "page_count": check_page_count,
    "authors": check_authors, "header_footer": check_header_footer,
}


class RuleEngine:
    def __init__(self, rules_db=None):
        self.db = rules_db or load_rules()

    def criteria(self, module):
        return [c for c in self.db["kriteria"] if c["modul"] == module]

    def run(self, parsed, options=None):
        options, results = options or {}, []
        for crit in self.criteria("M3"):
            res = CriterionResult(crit["id"], crit["kode"], crit["nama"], "M3", crit["severity"],
                                  group=crit.get("kelompok", ""))
            CHECKS[crit["check"]](parsed, crit["params"], res, options)
            results.append(res.finalize())
        return results
