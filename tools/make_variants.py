#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_variants.py - pembangkit varian naskah (controlled error injection) secara batch.

Masukan : satu atau lebih naskah dasar BEBAS CACAT (golden), mis. T01-Base.docx, T02-Base.docx
Keluaran: folder berisi varian ber-nama pendek + manifest.json/manifest.csv (kunci jawaban) + skipped.csv

Penamaan (awalan diambil dari nama basis sebelum "-Base"):
    T01-V00   kontrol: salinan basis tanpa cacat (untuk mengukur alarm palsu)
    T01-K02a  varian TUNGGAL: satu cacat; K02 = nomor kriteria, huruf = jenis cacat (lihat KATALOG di bawah)
    T01-M03   varian MAJEMUK: 3 cacat dari kriteria berbeda (M06 = 6 cacat, M09 = 9 cacat; 3+6+9 = 18 kriteria)
    T01-R01   varian REFERENSI untuk verifikasi Crossref (9 referensi diberi cacat: DOI/judul/fiktif)
    T01-EA / T01-EB  sepasang varian setara untuk uji efisiensi (jenis & jumlah cacat sama, lokasi berbeda)
    T01-N14a  varian NETRAL (kontrol negatif): perubahan yang BUKAN cacat (mis. singkatan dfl/pkg/src tidak dimiringkan);
              sistem HARUS tetap menilai sesuai. Mengukur alarm palsu pada kriteria terkait.
  Operator tambahan (v1.1, aktif secara bawaan; matikan dengan --no-ext / --no-benign):
    K05d  header/footer diketik sebagai paragraf biasa (bukan tabel)     K05e  teks liar di luar tabel footer ganjil
    K07e  Penulis Korespondensi di luar tabel                            K07f  Informasi/Riwayat Artikel di luar tabel
    K08d  Abstrak di luar tabel                                          K09c  Kata Kunci di luar tabel
    N14a  singkatan huruf kecil (dfl, pkg, src, cfg, env) tanpa miring   N14b  nama berkas/identifier/URL tanpa miring
  Operator tambahan dibangkitkan sebagai varian tunggal dan TIDAK masuk kolam varian majemuk/EA-EB, sehingga varian
  M03/M06/M09/EA/EB yang pernah dibangkitkan dengan seed yang sama tidak berubah (gunakan --pool-ext bila ingin
  memasukkannya; hasil majemuk/EA/EB lalu berbeda).

Pemakaian:
    python tools/make_variants.py naskah/T01-Base.docx naskah/T02-Base.docx --out varian --seed 42
    python tools/make_variants.py "naskah/*-Base.docx" --out varian
    python tools/make_variants.py T01-Base.docx --only K02a,K06b --no-multi --no-ref --no-eff   # uji sebagian
    python tools/make_variants.py T01-Base.docx --selfcheck    # jalankan sistem pemeriksa pada hasilnya (sanity check)
    python tools/make_variants.py "naskah/*-Base.docx" --out varian --no-ext --no-benign   # perilaku v1.0 (tanpa operator tambahan)
  Lalu evaluasi seluruh folder sekaligus:  python tools/run_batch.py varian

Catatan:
  * Hasil bersifat deterministik: seed sama + basis sama -> berkas identik.
  * Varian jumlah halaman (K04) menulis <Pages> pada docProps/app.xml (yang dibaca sistem). Jumlah halaman diukur
    dengan LibreOffice bila tersedia (--soffice atau variabel SOFFICE), jika tidak diasumsikan sama dengan target.
  * Basis harus memenuhi 18 kriteria, termasuk header/footer dan kotak identitas (Informasi Artikel, Abstrak, Kata Kunci,
    Penulis Korespondensi) di dalam TABEL seperti template; basis dicek otomatis di awal (--no-check-base untuk melewati).
    (K03: bagian identitas 1 kolom dan isi sesuai body_columns). Operator yang tidak berlaku pada basis (mis. tak ada Heading 3) dicatat di
    skipped.csv, bukan dipaksakan.
"""
import argparse
import copy
import csv
import difflib
import glob
import hashlib
import json
import math
import random
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from types import SimpleNamespace

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

from docx import Document                                   # noqa: E402
from docx.enum.text import WD_ALIGN_PARAGRAPH               # noqa: E402
from docx.oxml import OxmlElement                            # noqa: E402
from docx.oxml.ns import qn                                 # noqa: E402
from docx.shared import Cm, Pt                              # noqa: E402
from docx.text.paragraph import Paragraph                   # noqa: E402

import docx_edit as E                                       # noqa: E402
from checker import m2_parser, m3_rules                     # noqa: E402

VERSION = "1.1"


class NotApplicable(Exception):
    """Operator tidak berlaku pada basis ini (alasan dicatat)."""


@dataclass
class Op:
    code: str
    criterion: int
    desc: str
    func: object
    side: tuple = ()
    kind: str = "std"          # std | ref


OPS = {}


def op(code, crit, desc, side=(), kind="std"):
    def deco(fn):
        OPS[code] = Op(code, crit, desc, fn, tuple(side), kind)
        return fn
    return deco


# Urutan penerapan pada varian majemuk: yang mengubah struktur dulu, properti belakangan, halaman paling akhir.
ORDER = [10, 13, 15, 16, 18, 17, 8, 6, 9, 11, 14, 12, 7, 5, 3, 2, 1, 4]
# Pasangan kriteria yang tidak boleh berada dalam satu varian majemuk (saling meniadakan).
CONFLICTS = {frozenset(p) for p in [(10, 15), (10, 16), (15, 16)]}


class Ctx:
    def __init__(self, work, rng, soffice, params):
        self.work, self.rng, self.soffice, self.params = str(work), rng, soffice, params
        self.doc = Document(self.work)
        self.pd = m2_parser.parse(self.work)
        self.paras = {i: p for i, p, _ in m2_parser.iter_paragraphs(self.doc)}
        self.marks, self.state = [], {}

    def mark(self, par, role="target"):
        self.marks.append((par, E.par_text(par)[:80], role))

    def save(self):
        self.doc.save(self.work)

    def body(self, min_words=15, section=None, roles=("body",)):
        return [p for p in self.pd.paras if p.role in roles and p.words >= min_words
                and (section is None or p.section == section)]

    def pick(self, items, frac=0.3, n_min=1):
        if not items:
            return []
        n = min(len(items), max(n_min, round(len(items) * frac)))
        return sorted(self.rng.sample(items, n), key=lambda x: getattr(x, "idx", 0))


def _strip_non_field_text(container_el):
    """Buang teks yang BUKAN bagian field (mis. nomor halaman) di semua paragraf, termasuk di dalam tabel.
    Struktur (tabel, sel, paragraf, field) dipertahankan. Mandiri: tidak memerlukan versi docx_edit tertentu."""
    for p in container_el.iter(qn("w:p")):
        depth = 0
        for ch in list(p):
            if ch.tag in (qn("w:pPr"), qn("w:fldSimple")):
                continue
            if ch.tag == qn("w:r"):
                fc = ch.find(qn("w:fldChar"))
                if fc is not None:
                    kind = fc.get(qn("w:fldCharType"))
                    depth += 1 if kind == "begin" else 0
                    depth -= 1 if kind == "end" else 0
                    continue
                if depth > 0:
                    continue
            p.remove(ch)


def _foreign_kwargs(pd, prm):
    """Argumen langid.foreign_spans; memakai m3_rules.foreign_kwargs bila ada (checker v1.1), jika tidak bentuk lama."""
    fk = getattr(m3_rules, "foreign_kwargs", None)
    if fk:
        return fk(pd, prm)
    return dict(allowlist=prm["allowlist"], force_terms=prm["force_terms"], extra_patterns=prm["extra_patterns"],
                title_case_phrases=prm["title_case_phrases"], sentence_detection=prm["sentence_detection"])


def _prev_text(c, idx):
    return next((p for p in reversed(c.pd.paras[:idx]) if p.text or p.has_drawing), None)


# KATALOG OPERATOR
# ---- K01 kertas, K02 margin, K03 kolom (identitas vs isi) ----------------------------------------------------------
@op("K01a", 1, "kertas Letter")
def k01a(c):
    for s in c.doc.sections:
        s.page_width, s.page_height = Cm(21.59), Cm(27.94)


@op("K01b", 1, "kertas A5")
def k01b(c):
    for s in c.doc.sections:
        s.page_width, s.page_height = Cm(14.8), Cm(21.0)


@op("K02a", 2, "margin atas 3,5 cm")
def k02a(c):
    for s in c.doc.sections:
        s.top_margin = Cm(3.5)


@op("K02b", 2, "margin kiri 1,5 cm")
def k02b(c):
    for s in c.doc.sections:
        s.left_margin = Cm(1.5)


def _identity_body_sections(c):
    heads = c.pd.headings
    if not heads:
        raise NotApplicable("tidak ada heading (batas identitas/isi tidak diketahui)")
    cut = heads[0].idx
    ident = sorted({p.section_index for p in c.pd.paras if p.text and p.idx < cut})
    body = sorted({p.section_index for p in c.pd.paras if p.text and p.idx >= cut})
    if set(ident) & set(body):
        raise NotApplicable("basis tidak memiliki section terpisah untuk bagian identitas")
    return ident, body


@op("K03a", 3, "bagian identitas 2 kolom")
def k03a(c):
    ident, _ = _identity_body_sections(c)
    for i in ident:
        E.set_columns(c.doc.sections[i]._sectPr, 2)


@op("K03b", 3, "bagian isi bukan jumlah kolom yang diwajibkan")
def k03b(c):
    _, body = _identity_body_sections(c)
    wrong = 1 if c.params[3]["body_columns"] != 1 else 2
    for i in body:
        E.set_columns(c.doc.sections[i]._sectPr, wrong)


# ---- K12 jenis dan ukuran huruf, spasi ---------------------------------------------------------------------------
def _on_body(c, fn):
    chosen = c.pick(c.body())
    if not chosen:
        raise NotApplicable("tidak ada paragraf isi")
    for info in chosen:
        par = c.paras[info.idx]
        c.mark(par)
        fn(par)


@op("K12a", 12, "font Times New Roman pada 30% paragraf isi")
def k12a(c):
    def f(par):
        for r in E.inline_runs(par):
            r.font.name = "Times New Roman"
    _on_body(c, f)


@op("K12b", 12, "ukuran huruf 12 pt pada 30% paragraf isi")
def k12b(c):
    def f(par):
        for r in E.inline_runs(par):
            r.font.size = Pt(12)
    _on_body(c, f)


@op("K12c", 12, "spasi baris 1,5 pada 30% paragraf isi")
def k12c(c):
    _on_body(c, lambda par: setattr(par.paragraph_format, "line_spacing", 1.5))


# ---- K06 judul (panjang, kata terlarang, format) -----------------------------------------------------------------------------------
NEUTRAL = ["pada", "berbagai", "jenis", "data", "yang", "berbeda", "secara", "bersama", "dalam", "konteks",
           "nyata", "terbaru", "untuk", "tujuan", "umum"]


def _title_par(c):
    if c.pd.title is None:
        raise NotApplicable("judul tidak ditemukan")
    return c.paras[c.pd.title.idx]


@op("K06a", 6, "judul 20 kata")
def k06a(c):
    par = _title_par(c)
    text = E.par_text(par)
    n = len(text.split())
    target = 20 if n <= 15 else n + 5
    extra = [NEUTRAL[i % len(NEUTRAL)] for i in range(target - n)]
    c.mark(par)
    E.replace_range(par, len(text), len(text), " " + " ".join(extra))


@op("K06b", 6, "judul memuat 'Analisis ... Studi Kasus'")
def k06b(c):
    par = _title_par(c)
    text = E.par_text(par)
    words = text.split()[:12]
    c.mark(par)
    E.replace_range(par, 0, len(text), "Analisis " + " ".join(words) + " Studi Kasus")


@op("K06c", 6, "font dan ukuran judul salah (Arial 12 pt, tidak tebal)")
def k06c(c):
    par = _title_par(c)
    c.mark(par)
    for r in E.inline_runs(par):
        r.font.name, r.font.size, r.font.bold = "Arial", Pt(12), False


@op("K06d", 6, "judul rata kiri dan tanpa spasi sesudah")
def k06d(c):
    par = _title_par(c)
    c.mark(par)
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    par.paragraph_format.space_after = Pt(0)


# ---- K08 abstrak ---------------------------------------------------------------------------------
def _abstract_pars(c):
    pd = c.pd
    if pd.abstract_idx is None or not pd.abstract_text:
        raise NotApplicable("abstrak tidak ditemukan")
    first = c.paras[pd.abstract_idx]
    parent = first._p.getparent()
    if parent.tag == qn("w:tc"):
        return [Paragraph(p, first._parent) for p in parent.findall(qn("w:p"))
                if "".join(t.text or "" for t in p.iter(qn("w:t"))).strip()]
    return [first]


def _set_abstract(c, new_body):
    pars = _abstract_pars(c)
    first = pars[0]
    text = E.par_text(first)
    m = re.match(r"(?i)^(abstrak|abstract)\b[\s:.\u2014-]*", text)
    start = len(m.group(0)) if m else 0                    # pertahankan label "Abstrak:" bila ada
    c.mark(first)
    E.replace_range(first, start, len(text), new_body)
    for p in pars[1:]:
        E.delete_paragraph(p)


@op("K08a", 8, "abstrak sekitar 60 kata")
def k08a(c):
    _abstract_pars(c)
    _set_abstract(c, " ".join(c.pd.abstract_text.split()[:60]))


@op("K08b", 8, "abstrak sekitar 320 kata")
def k08b(c):
    _abstract_pars(c)
    words = c.pd.abstract_text.split()
    out = []
    while len(out) < 320:
        out += words
    _set_abstract(c, " ".join(out[:320]))


@op("K08c", 8, "abstrak memuat sitasi [1]")
def k08c(c):
    first = _abstract_pars(c)[0]
    text = E.par_text(first)
    m = re.search(r"[.!?]", text[30:])
    pos = 30 + m.start() if m else len(text)
    c.mark(first)
    E.replace_range(first, pos, pos, " [1]")


@op("K07d", 7, "ukuran huruf, perataan, dan spasi penulis/afiliasi salah")
def k07d(c):
    if c.pd.authors is None:
        raise NotApplicable("baris penulis tidak ditemukan")
    for info in [c.pd.authors] + list(c.pd.affiliations):
        par = c.paras[info.idx]
        c.mark(par)
        for r in E.inline_runs(par):
            r.font.size = Pt(12)
        par.alignment = WD_ALIGN_PARAGRAPH.LEFT
        par.paragraph_format.space_after = Pt(0)


# ---- K09 kata kunci ------------------------------------------------------------------------------
EXTRA_KW = ["klasifikasi teks", "evaluasi model", "analisis data", "pembelajaran mesin", "pemrosesan bahasa",
            "kualitas data", "validasi silang"]


def _kw_info(c):
    pd = c.pd
    if pd.keywords_idx is None or not pd.keywords:
        raise NotApplicable("kata kunci tidak ditemukan")
    label = c.paras[pd.keywords_idx]
    text = E.par_text(label)
    m = re.match(r"^\s*(kata kunci|keywords?)\s*:?\s*(.*)$", text, re.I)
    if m and m.group(2).strip():
        return "inline", label, m, []
    parent = label._p.getparent()
    if parent.tag != qn("w:tc"):
        raise NotApplicable("struktur kata kunci tidak dikenali")
    allp = [Paragraph(p, label._parent) for p in parent.findall(qn("w:p"))]
    i = next(k for k, p in enumerate(allp) if p._p is label._p)
    return "list", label, m, [p for p in allp[i + 1:] if E.par_text(p).strip()]


def _kw_sep(text):
    return "; " if ";" in text else ", "


@op("K09a", 9, "7 kata kunci")
def k09a(c):
    mode, label, m, kwp = _kw_info(c)
    kws = list(c.pd.keywords)
    have = {k.lower() for k in kws}
    add = []
    for e in EXTRA_KW:
        if len(kws) + len(add) >= max(7, len(kws) + 2):
            break
        if e.lower() not in have:
            add.append(e)
    c.mark(label)
    if mode == "inline":
        text = E.par_text(label)
        E.replace_range(label, m.start(2), len(text), _kw_sep(text).join(kws + add))
    else:
        last = kwp[-1]
        for kw in add:
            new = E.clone_after(last)
            E.replace_range(new, 0, len(E.par_text(new)), kw)
            last = new


@op("K09b", 9, "kata kunci tidak urut abjad")
def k09b(c):
    mode, label, m, kwp = _kw_info(c)
    kws = list(c.pd.keywords)
    if len(kws) < 2:
        raise NotApplicable("kata kunci kurang dari 2")
    new_order = sorted(kws, key=str.lower, reverse=True)
    c.mark(label)
    if mode == "inline":
        text = E.par_text(label)
        E.replace_range(label, m.start(2), len(text), _kw_sep(text).join(new_order))
    else:
        for p, kw in zip(kwp, new_order):
            E.replace_range(p, 0, len(E.par_text(p)), kw)


# ---- K10 struktur IMRaD --------------------------------------------------------------------------------
def _l1(c):
    return [h for h in c.pd.headings if h.level == 1]


def _block(c, key):
    heads = _l1(c)
    body_el = c.doc.element.body
    for i, h in enumerate(heads):
        if h.section == key:
            start = c.paras[h.idx]._p
            if start.getparent() is not body_el:
                raise NotApplicable("heading berada di struktur bersarang")
            nxt = c.paras[heads[i + 1].idx]._p if i + 1 < len(heads) else None
            els, el = [], start
            while el is not None and el is not nxt and el.tag != qn("w:sectPr"):
                els.append(el)
                el = el.getnext()
            if any(e.find(".//" + qn("w:sectPr")) is not None for e in els):
                raise NotApplicable("blok memuat section break")
            return h, els
    raise NotApplicable(f"bagian '{key}' tidak ditemukan")


@op("K10a", 10, "bagian Metodologi dihapus", side=(18,))
def k10a(c):
    h, els = _block(c, "metodologi")
    c.mark(c.paras[h.idx])
    for e in els:
        e.getparent().remove(e)


@op("K10b", 10, "urutan Hasil dan Kesimpulan ditukar", side=(18,))
def k10b(c):
    hh, hels = _block(c, "hasil_pembahasan")
    kh, kels = _block(c, "kesimpulan")
    if kh.idx < hh.idx:
        raise NotApplicable("Kesimpulan sudah mendahului Hasil")
    c.mark(c.paras[hh.idx])
    c.mark(c.paras[kh.idx])
    anchor = hels[0]
    for e in kels:
        anchor.addprevious(e)


@op("K10c", 10, "bagian 'Tinjauan Pustaka' ditambahkan")
def k10c(c):
    h, _ = _block(c, "metodologi")
    src = [i for i in _intro(c)]
    if not src:
        raise NotApplicable("tidak ada paragraf Pendahuluan untuk digandakan")
    heading = c.paras[h.idx]
    text = E.par_text(heading)
    pre = re.match(r"^\d+(?:\.\d+)*\.?\s*", text)
    name = "TINJAUAN PUSTAKA" if text.isupper() else "Tinjauan Pustaka"
    new_h = copy.deepcopy(heading._p)
    heading._p.addprevious(new_h)
    nh = Paragraph(new_h, heading._parent)
    E.replace_range(nh, len(pre.group(0)) if pre else 0, len(E.par_text(nh)), name)
    new_p = copy.deepcopy(c.paras[c.rng.choice(src).idx]._p)
    new_h.addnext(new_p)
    c.mark(nh)


# ---- K11 format subjudul ---------------------------------------------------------------------------------
def _heads(c, level):
    hs = [h for h in c.pd.headings if h.level == level]
    if not hs:
        raise NotApplicable(f"tidak ada Heading {level}")
    return hs


@op("K11a", 11, "Heading 1 tidak kapital")
def k11a(c):
    for h in _heads(c, 1):
        par = c.paras[h.idx]
        c.mark(par)
        t = E.par_text(par)
        E.replace_range(par, 0, len(t), t.title())
        for r in E.inline_runs(par):
            r.font.all_caps = False


@op("K11b", 11, "Heading 2 tidak tebal")
def k11b(c):
    for h in _heads(c, 2):
        par = c.paras[h.idx]
        c.mark(par)
        for r in E.inline_runs(par):
            r.font.bold = False


@op("K11c", 11, "Heading 3 tebal")
def k11c(c):
    for h in _heads(c, 3):
        par = c.paras[h.idx]
        c.mark(par)
        for r in E.inline_runs(par):
            r.font.bold = True


def _all_heads(c):
    if not c.pd.headings:
        raise NotApplicable("tidak ada heading")
    return c.pd.headings


@op("K11d", 11, "indentasi subjudul salah (kiri 1,27 cm tanpa gantung)")
def k11d(c):
    for h in _all_heads(c):
        par = c.paras[h.idx]
        c.mark(par)
        par.paragraph_format.left_indent = Cm(1.27)
        par.paragraph_format.first_line_indent = Cm(0)


@op("K11e", 11, "spasi subjudul salah (sebelum 0 pt, sesudah 12 pt)")
def k11e(c):
    for h in _all_heads(c):
        par = c.paras[h.idx]
        c.mark(par)
        par.paragraph_format.space_before = Pt(0)
        par.paragraph_format.space_after = Pt(12)


# ---- K13 paragraf Pendahuluan --------------------------------------------------------------------
def _intro(c):
    return c.body(15, "pendahuluan", ("body", "list"))


def _merge_into(dst, src):
    """Pindahkan isi `src` ke akhir `dst` (format run terjaga), lalu hapus `src`."""
    runs = list(E.inline_runs(dst))
    if runs:
        dst._p.append(E._new_run_like(runs[-1], " "))
    for ch in list(src._p):
        if ch.tag != qn("w:pPr"):
            dst._p.append(ch)
    src._p.getparent().remove(src._p)


@op("K13a", 13, "Pendahuluan 2 paragraf")
def k13a(c):
    objs = [c.paras[i.idx] for i in _intro(c)]
    if len(objs) < 3:
        raise NotApplicable(f"Pendahuluan sudah {len(objs)} paragraf")
    c.mark(objs[1])
    for p in objs[2:]:
        _merge_into(objs[1], p)


@op("K13b", 13, "Pendahuluan 8 paragraf")
def k13b(c):
    objs = [c.paras[i.idx] for i in _intro(c)]
    if not objs:
        raise NotApplicable("Pendahuluan tidak ditemukan")
    count = len(objs)
    for _ in range(60):
        if count >= 8:
            break
        best = None
        for p in objs:
            t = E.par_text(p)
            sp = E.sentence_spans(t)
            for k in range(1, len(sp)):
                left, right = t[:sp[k][0]], t[sp[k][0]:]
                if len(left.split()) >= 15 and len(right.split()) >= 15:
                    score = abs(len(left.split()) - len(right.split()))
                    if best is None or score < best[0]:
                        best = (score, p, sp[k][0], len(right) - len(right.lstrip()))
        if not best:
            break
        _, p, cut, ws = best
        t = E.par_text(p)
        new = E.clone_after(p)
        E.replace_range(p, cut, len(t), "")                 # potong ekor (format sisa terjaga)
        E.replace_range(new, 0, cut + ws, "")               # potong kepala pada salinan
        objs.insert(objs.index(p) + 1, new)
        c.mark(p)
        c.mark(new)
        count += 1
    if count < 8:
        raise NotApplicable(f"hanya dapat dipecah menjadi {count} paragraf")


# ---- K15 caption ---------------------------------------------------------------------------------
@op("K15a", 15, "caption gambar dihapus")
def k15a(c):
    caps = [cp for cp in c.pd.captions if cp["kind"] == "gambar"]
    if not caps:
        raise NotApplicable("tidak ada caption gambar")
    cap = c.rng.choice(caps)
    fig = _prev_text(c, cap["idx"])
    if fig is not None and fig.has_drawing:
        c.mark(c.paras[fig.idx], "figure")
    par = c.paras[cap["idx"]]
    c.mark(par, "caption")
    E.delete_paragraph(par)


@op("K15b", 15, "caption tabel dipindah ke bawah tabel")
def k15b(c):
    caps = {cp["idx"]: cp for cp in c.pd.captions if cp["kind"] == "tabel"}
    for t in (t for t in c.pd.tables if not t["front"]):
        prev = _prev_text(c, t["first_idx"])
        if prev is not None and prev.idx in caps:
            cap = c.paras[prev.idx]
            tbl = c.paras[t["first_idx"]]._p.xpath("ancestor::w:tbl")[-1]
            c.mark(cap, "caption")
            tbl.addnext(cap._p)
            return
    raise NotApplicable("tidak ada tabel dengan caption di atasnya")


@op("K15c", 15, "format caption 'Gambar 1: ...'")
def k15c(c):
    caps = [cp for cp in c.pd.captions if cp["kind"] == "gambar"] or [cp for cp in c.pd.captions if cp["kind"] == "tabel"]
    if not caps:
        raise NotApplicable("tidak ada caption")
    par = c.paras[caps[0]["idx"]]
    text = E.par_text(par)
    m = re.match(r"^(Gambar|Tabel)\s+\d+\.", text)
    if not m:
        raise NotApplicable("format caption asal tidak standar")
    c.mark(par, "caption")
    E.replace_range(par, m.end() - 1, m.end(), ":")


# ---- K16 rujukan gambar/tabel di teks ------------------------------------------------------------
def _uncite(c, kind):
    caps = [cp for cp in c.pd.captions if cp["kind"] == kind and cp["number"] == 1]
    if not caps:
        raise NotApplicable(f"tidak ada caption {kind} 1")
    rx = re.compile(r"\b(Gambar|Gbr\.?)\s*1\b" if kind == "gambar" else r"\bTabel\s*1\b", re.I)
    repl = "gambar berikut" if kind == "gambar" else "tabel berikut"
    hit = False
    for info in c.pd.paras:
        if info.role in ("body", "list") and info.raw_text:
            ms = list(rx.finditer(info.raw_text))
            if ms:
                par = c.paras[info.idx]
                hit = True
                for m in reversed(ms):
                    E.replace_range(par, m.start(), m.end(), repl)
    if not hit:
        raise NotApplicable("tidak ada rujukan di teks")
    c.mark(c.paras[caps[0]["idx"]], "caption")


@op("K16a", 16, "rujukan 'Gambar 1' dihapus dari teks")
def k16a(c):
    _uncite(c, "gambar")


@op("K16b", 16, "rujukan 'Tabel 1' dihapus dari teks")
def k16b(c):
    _uncite(c, "tabel")


# ---- K14 format paragraf isi -------------------------------------------------------------------------
@op("K14a", 14, "tanpa indentasi baris pertama pada 30% paragraf")
def k14a(c):
    _on_body(c, lambda par: setattr(par.paragraph_format, "first_line_indent", Cm(0)))


@op("K14b", 14, "rata kiri pada 30% paragraf")
def k14b(c):
    _on_body(c, lambda par: setattr(par, "alignment", WD_ALIGN_PARAGRAPH.LEFT))


# ---- K18 sitasi & daftar pustaka (M4) ------------------------------------------------------------
def _refs(c):
    if not c.pd.references:
        raise NotApplicable("tidak ada referensi")
    return c.pd.references


def _cite_edits(text, keep):
    edits, pairs = [], []
    for m in re.finditer(r"\[(\d+)\]\s*[\u2013\u2014-]\s*\[(\d+)\]", text):
        a, b = int(m.group(1)), int(m.group(2))
        pairs.append((m.start(), m.end()))
        if a > keep:
            edits.append((m.start(), m.end(), ""))
        elif b > keep:
            edits.append((m.start(), m.end(), f"[{a}]" if a == keep else f"[{a}]\u2013[{keep}]"))
    for m in m2_parser.CITE_RE.finditer(text):
        if any(s <= m.start() and m.end() <= e for s, e in pairs) or not m2_parser._is_citation_bracket(text, m):
            continue
        nums = m2_parser.extract_citations(m.group(0))
        kept = [n for n in nums if n <= keep]
        if len(kept) != len(nums):
            edits.append((m.start(), m.end(), "[" + ", ".join(map(str, kept)) + "]" if kept else ""))
    return sorted(edits)


@op("K18a", 18, "hanya 15 referensi (sitasi ikut dihapus)")
def k18a(c):
    refs = _refs(c)
    keep = 15
    if len(refs) <= keep:
        raise NotApplicable(f"referensi sudah {len(refs)}")
    for cit in c.pd.citations:
        if max(cit["nums"]) > keep:
            par = c.paras[cit["idx"]]
            teks = E.par_text(par)
            for s, e, r in reversed(_cite_edits(teks, keep)):
                if not r:                                    # sitasi dibuang: rapikan spasi dan koma di sekitarnya
                    if teks[e:e + 1] in ("", ".", ",", ";", ":", ")"):
                        while s > 0 and teks[s - 1] in " ,":
                            s -= 1
                    else:
                        while e < len(teks) and teks[e] == " ":
                            e += 1
                E.replace_range(par, s, e, r)
    first = True
    for r in refs:
        if r["num"] is not None and r["num"] > keep:
            par = c.paras[r["idx"]]
            if first:
                c.mark(par)
                first = False
            E.delete_paragraph(par)


@op("K18b", 18, "sitasi ke nomor yang tidak ada")
def k18b(c):
    n = len(_refs(c))
    cands = list(c.pd.citations)
    c.rng.shuffle(cands)
    for cit in cands:
        par = c.paras[cit["idx"]]
        text = E.par_text(par)
        for m in re.finditer(r"\[(\d+)\]", text):
            if m2_parser._is_citation_bracket(text, m):
                c.mark(par)
                E.replace_range(par, m.start(1), m.end(1), str(n + 5))
                return
    raise NotApplicable("tidak ada sitasi tunggal")


@op("K18c", 18, "entri bukan format IEEE (judul tanpa tanda kutip)")
def k18c(c):
    cands = [r for r in _refs(c) if r["has_quoted_title"]]
    if not cands:
        raise NotApplicable("tidak ada entri bertanda kutip")
    for r in sorted(c.rng.sample(cands, min(3, len(cands))), key=lambda x: x["idx"]):
        par = c.paras[r["idx"]]
        c.mark(par)
        text = E.par_text(par)
        for m in reversed(list(re.finditer(r"[\u201c\u201d\"]", text))):
            E.replace_range(par, m.start(), m.end(), "")


@op("K18d", 18, "satu referensi tidak pernah disitasi")
def k18d(c):
    refs = _refs(c)

    def removable(num):
        plain = total = 0
        for cit in c.pd.citations:
            text = E.par_text(c.paras[cit["idx"]])
            if re.search(rf"[\u2013\u2014-]\s*\[{num}\]|\[{num}\]\s*[\u2013\u2014-]", text):
                return False
            plain += len(re.findall(rf"\[{num}\]", text))
            total += cit["nums"].count(num)
        return plain > 0 and plain == total

    cands = [r for r in refs if r["num"] is not None and removable(r["num"])]
    if not cands:
        raise NotApplicable("tidak ada referensi yang aman dihapus sitasinya")
    r = c.rng.choice(cands)
    for cit in c.pd.citations:
        par = c.paras[cit["idx"]]
        text = E.par_text(par)
        for m in reversed(list(re.finditer(rf"\[{r['num']}\]", text))):
            s, e = m.start(), m.end()
            if text[e:e + 2] == ", ":
                e += 2
            elif text[max(0, s - 2):s] == ", ":
                s -= 2
            E.replace_range(par, s, e, "")
    c.mark(c.paras[r["idx"]])


@op("K18e", 18, "huruf entri referensi 10 pt pada 30% entri")
def k18e(c):
    infos = [SimpleNamespace(idx=r["idx"]) for r in _refs(c)]
    for info in c.pick(infos):
        par = c.paras[info.idx]
        c.mark(par)
        for run in E.inline_runs(par):
            run.font.size = Pt(10)


# ---- K17 kesimpulan ------------------------------------------------------------------------------
def _concl(c):
    pars = c.pd.section_paras("kesimpulan")
    if not pars:
        raise NotApplicable("Kesimpulan tidak ditemukan")
    return pars


SUGGEST_RX = re.compile(r"\b(saran|rekomendasi|penelitian (?:selanjutnya|lanjutan|mendatang|berikutnya)|future work)\b", re.I)


@op("K17a", 17, "dua kalimat abstrak disalin ke Kesimpulan")
def k17a(c):
    pars = _concl(c)
    sents = [s for s in re.split(r"(?<=[.!?])\s+", c.pd.abstract_text) if len(s.split()) >= 8][:2]
    if not sents:
        raise NotApplicable("abstrak tidak punya kalimat yang cukup panjang")
    par = c.paras[pars[-1].idx]
    t = E.par_text(par)
    c.mark(par)
    E.replace_range(par, len(t), len(t), " " + " ".join(sents))


@op("K17b", 17, "saran dihapus dari Kesimpulan")
def k17b(c):
    hit = False
    for info in _concl(c):
        par = c.paras[info.idx]
        t = E.par_text(par)
        spans = [(s, e) for s, e in E.sentence_spans(t) if SUGGEST_RX.search(t[s:e])]
        if spans:
            hit = True
            c.mark(par)
            for s, e in reversed(spans):
                if e >= len(t.rstrip()):                    # kalimat terakhir: buang spasi sebelumnya
                    while s > 0 and t[s - 1] == " ":
                        s -= 1
                else:                                       # kalimat di tengah: buang spasi sesudahnya
                    while e < len(t) and t[e] == " ":
                        e += 1
                E.replace_range(par, s, e, "")
            if not E.par_text(par).strip():                 # paragraf jadi kosong: hapus, bukan dibiarkan sebagai baris kosong
                E.delete_paragraph(par)
    if not hit:
        raise NotApplicable("tidak ada kalimat saran pada Kesimpulan")


# ---- K04 batas halaman --------------------------------------------------------------------------
def _k16(c, target):
    pages, source = E.measure_pages(c.work, c.soffice), "libreoffice"
    if not pages:
        pages, source = c.pd.pages_meta, "metadata"
    if not pages:
        pages, source = max(1, math.ceil(c.pd.total_words / 600)), "estimasi"
    if pages >= target:
        raise NotApplicable(f"naskah dasar sudah {pages} halaman")
    wpp = max(150.0, c.pd.total_words / pages)
    pool = c.body(15)
    anchors = c.body(15, "hasil_pembahasan") or pool
    if not pool:
        raise NotApplicable("tidak ada paragraf untuk digandakan")
    cur = c.paras[anchors[-1].idx]
    c.mark(cur)
    added_total = 0

    def grow(words):
        nonlocal cur, added_total
        added = 0
        while added < words:
            src = c.paras[c.rng.choice(pool).idx]
            new = copy.deepcopy(src._p)
            cur._p.addnext(new)
            cur = Paragraph(new, src._parent)
            added += len(E.par_text(cur).split())
        added_total += added

    grow((target - pages + 0.6) * wpp)
    final = target
    if source == "libreoffice":
        reached = False
        for _ in range(8):
            c.save()
            measured = E.measure_pages(c.work, c.soffice)
            if measured is None:
                break
            final = measured
            if measured >= target:
                reached = True
                break
            if measured > pages:                       # laju kata/halaman sebenarnya dari hasil ukur
                wpp = max(100.0, added_total / (measured - pages))
            grow((target - measured + 0.6) * wpp)
        if not reached and final < target:
            raise NotApplicable(f"target {target} halaman tidak tercapai (terukur {final})")
    else:
        source = "diasumsikan"
    c.state["pages"], c.state["pages_source"] = final, source


@op("K04a", 4, "13 halaman")
def k04a(c):
    _k16(c, 13)


@op("K04b", 4, "16 halaman")
def k04b(c):
    _k16(c, 16)


# ---- K07 penulis & afiliasi ----------------------------------------------------------------------
@op("K07a", 7, "afiliasi dihapus")
def k07a(c):
    if not c.pd.affiliations:
        raise NotApplicable("tidak ada afiliasi")
    for info in c.pd.affiliations:
        par = c.paras[info.idx]
        c.mark(par)
        E.delete_paragraph(par)


@op("K07b", 7, "penanda afiliasi penulis tidak cocok")
def k07b(c):
    if c.pd.authors is None:
        raise NotApplicable("baris penulis tidak ditemukan")
    par = c.paras[c.pd.authors.idx]
    norm_text = m2_parser._with_superscripts(par)
    groups = list(m2_parser.MARKER_RE.finditer(norm_text))
    if not groups:
        raise NotApplicable("penulis tanpa penanda afiliasi")
    g = groups[-1]
    new_n = max(c.pd.affiliation_markers or [0]) + 3
    orig = E.par_text(par)[g.start():g.end()]
    new = str(new_n) if re.fullmatch(r"\d+", orig) else str(new_n).translate(m2_parser.DIGIT_TO_SUP)
    c.mark(par)
    E.replace_range(par, g.start(), g.end(), new)


@op("K07c", 7, "email penulis korespondensi dihapus")
def k07c(c):
    corr = c.pd.corresponding
    if not corr or not corr["email"]:
        raise NotApplicable("tidak ada email penulis korespondensi")
    for off in range(0, 4):
        par = c.paras.get(corr["idx"] + off)
        if par is None:
            break
        text = E.par_text(par)
        m = m2_parser.EMAIL_RE.search(text)
        if m:
            c.mark(par)
            E.replace_range(par, m.start(), m.end(), "")
            return
    raise NotApplicable("email tidak ditemukan pada paragraf")


# ---- K05 header & footer -------------------------------------------------------------------------
@op("K05a", 5, "opsi halaman ganjil-genap dimatikan")
def k05a(c):
    if not c.doc.settings.odd_and_even_pages_header_footer:
        raise NotApplicable("opsi sudah nonaktif")
    c.doc.settings.odd_and_even_pages_header_footer = False


@op("K05b", 5, "footer halaman genap hanya berisi nomor halaman")
def k05b(c):
    ft = c.doc.sections[0].even_page_footer
    if ft.is_linked_to_previous:
        raise NotApplicable("footer genap tidak terdefinisi")
    el = ft._element
    paras = list(el.iter(qn("w:p")))
    if not paras:
        raise NotApplicable("footer genap tidak memuat paragraf")
    _strip_non_field_text(el)                        # tabel footer dipertahankan: hanya satu cacat (isi), bukan letak
    if not m2_parser._has_page_field(el):
        E.add_page_field(Paragraph(paras[0], ft))


@op("K05c", 5, "nomor halaman diketik manual (field PAGE dihapus)")
def k05c(c):
    count = 0
    for s in c.doc.sections:
        for part in (s.first_page_footer, s.footer, s.even_page_footer):
            if not part.is_linked_to_previous:
                count += E.replace_page_fields(part._element, "1")
    if not count:
        raise NotApplicable("tidak ada field PAGE di footer")


# ---- K14c/d istilah asing miring (bagian K14) --------------------------------------------------------------------
def _foreign_spans(c):
    from checker import langid
    prm, out = c.params[14]["foreign_italic"], []
    for info in c.body(3):
        spans = langid.foreign_spans(info.raw_text, **_foreign_kwargs(c.pd, prm))
        ital = [(s, e) for s, e, _ in spans if m3_rules._italic_coverage(info, s, e) >= prm["min_italic_ratio"]]
        if ital:
            out.append((info, ital))
    return out


@op("K14c", 14, "istilah asing tidak miring pada 30% paragraf")
def k14c(c):
    items = _foreign_spans(c)
    if not items:
        raise NotApplicable("tidak ada istilah asing bercetak miring pada basis")
    by = {i.idx: s for i, s in items}
    for info in c.pick([i for i, _ in items]):
        par = c.paras[info.idx]
        c.mark(par)
        for s, e in by[info.idx]:
            E.set_italic_range(par, s, e, False)


@op("K14d", 14, "istilah asing majemuk hanya sebagian miring")
def k14d(c):
    items = [(i, [(s, e) for s, e in sp if re.search(r"[\s\-]", E.par_text(c.paras[i.idx])[s:e])]) for i, sp in _foreign_spans(c)]
    items = [(i, sp) for i, sp in items if sp]
    if not items:
        raise NotApplicable("tidak ada istilah majemuk bercetak miring")
    by = {i.idx: s for i, s in items}
    for info in c.pick([i for i, _ in items]):
        par = c.paras[info.idx]
        c.mark(par)
        text = E.par_text(par)
        for s, e in by[info.idx]:
            m = re.search(r"[\s\-]+", text[s:e])
            E.set_italic_range(par, s + m.end(), e, False)


# OPERATOR TAMBAHAN v1.1 (kind="ext" dan "benign")
# Cek letak di tabel (K05, K07, K08, K09) dan kontrol negatif istilah asing (K14).
# Varian tunggal saja; tidak masuk kolam majemuk/EA-EB kecuali --pool-ext (lihat process_base).
HF_PARTS = ("first_page_header", "header", "even_page_header", "first_page_footer", "footer", "even_page_footer")


def _unwrap_tables(root):
    """Ganti tiap tabel tingkat-atas pada `root` dengan paragraf isi selnya (tampilan mirip, bukan tabel lagi)."""
    n = 0
    for tbl in list(root.findall(qn("w:tbl"))):
        for tc in tbl.iter(qn("w:tc")):
            for p in tc.findall(qn("w:p")):
                tbl.addprevious(copy.deepcopy(p))
        tbl.getparent().remove(tbl)
        n += 1
    return n


@op("K05d", 5, "header/footer diketik sebagai paragraf biasa (bukan tabel)", kind="ext")
def k05d(c):
    n = 0
    for sec in c.doc.sections:
        for name in HF_PARTS:
            part = getattr(sec, name)
            if not part.is_linked_to_previous:
                n += _unwrap_tables(part._element)
    if not n:
        raise NotApplicable("header/footer pada basis tidak memakai tabel")


@op("K05e", 5, "teks tambahan di luar tabel pada footer halaman ganjil", kind="ext")
def k05e(c):
    ft = c.doc.sections[0].footer
    if ft.is_linked_to_previous or ft._element.find(qn("w:tbl")) is None:
        raise NotApplicable("footer ganjil pada basis tidak memakai tabel")
    ft.add_paragraph("Halaman")


def _text_of(p_el):
    return "".join(t.text or "" for t in p_el.iter(qn("w:t"))).strip()


def _need_identity(c):
    if not hasattr(c.pd, "identity"):
        raise NotApplicable("parser tidak memuat letak identitas (butuh checker v1.1)")


def _tc_pars(par):
    """Semua paragraf pada sel tabel yang memuat `par`; NotApplicable bila `par` tidak berada di sel tabel."""
    tc = par._p.getparent()
    if tc.tag != qn("w:tc"):
        raise NotApplicable("elemen tidak berada di dalam tabel pada basis")
    return [Paragraph(p, par._parent) for p in tc.findall(qn("w:p"))]


def _move_out(c, pars, primary):
    """Pindahkan `pars` (paragraf dari sel tabel identitas) menjadi paragraf biasa tepat setelah tabel itu, urut dokumen.
    Sel asal tetap valid (diberi paragraf kosong bila habis). `primary` = paragraf yang ditandai sebagai lokasi cacat."""
    uniq, seen = [], set()
    for p in pars:
        if id(p._p) not in seen:
            seen.add(id(p._p))
            uniq.append(p)
    tbls = [a for a in uniq[0]._p.iterancestors() if a.tag == qn("w:tbl")]
    if not tbls or tbls[-1].getparent().tag != qn("w:body"):
        raise NotApplicable("tabel identitas tidak berada langsung di badan dokumen")
    tbl = tbls[-1]
    order = {id(x): i for i, x in enumerate(tbl.iter(qn("w:p")))}
    uniq.sort(key=lambda p: order.get(id(p._p), 0))
    anchor, moved = tbl, None
    for p in uniq:
        el = copy.deepcopy(p._p)
        anchor.addnext(el)
        anchor = el
        tc = p._p.getparent()
        tc.remove(p._p)
        if tc.find(qn("w:p")) is None:
            tc.append(OxmlElement("w:p"))
        if p._p is primary._p:
            moved = Paragraph(el, c.doc._body)
    c.mark(moved if moved is not None else Paragraph(anchor, c.doc._body))


@op("K08d", 8, "abstrak diketik di luar tabel identitas", kind="ext")
def k08d(c):
    _need_identity(c)
    body = _abstract_pars(c)
    _tc_pars(body[0])                                         # harus berada di sel tabel
    tbls = [a for a in body[0]._p.iterancestors() if a.tag == qn("w:tbl")]
    label = next((p for p in tbls[-1].iter(qn("w:p"))
                  if re.fullmatch(r"(?i)(abstrak|abstract)\s*:?", _text_of(p))), None)
    if label is None:
        raise NotApplicable("label 'ABSTRAK' pada tabel tidak ditemukan")
    lab = Paragraph(label, body[0]._parent)
    _move_out(c, [lab] + body, lab)             # parser menunjuk paragraf label 'ABSTRAK' sebagai lokasi abstrak di luar tabel


@op("K09c", 9, "kata kunci diketik di luar tabel identitas", kind="ext")
def k09c(c):
    _need_identity(c)
    _, label, _, _ = _kw_info(c)
    pars = [p for p in _tc_pars(label) if E.par_text(p).strip()]
    _move_out(c, pars, label)


@op("K07e", 7, "penulis korespondensi diketik di luar tabel identitas [catatan: paragraf hasil pindahan ikut terbaca parser sebagai baris afiliasi; temuan format afiliasi tambahan pada K07 wajar]", kind="ext")
def k07e(c):
    _need_identity(c)
    corr = c.pd.corresponding
    if not corr or c.paras.get(corr["idx"]) is None:
        raise NotApplicable("penulis korespondensi tidak ditemukan")
    label = c.paras[corr["idx"]]
    _move_out(c, [p for p in _tc_pars(label) if E.par_text(p).strip()], label)


@op("K07f", 7, "Informasi Artikel/Riwayat Artikel diketik di luar tabel identitas [catatan: paragraf hasil pindahan ikut terbaca parser sebagai baris afiliasi; temuan format afiliasi tambahan pada K07 wajar]", kind="ext")
def k07f(c):
    _need_identity(c)
    entries = c.pd.identity.get("informasi") or []
    if not entries:
        raise NotApplicable("blok Informasi Artikel tidak ditemukan pada basis")
    pars = []
    for e in entries:
        pars += [p for p in _tc_pars(c.paras[e["idx"]]) if E.par_text(p).strip()]
    _move_out(c, pars, c.paras[entries[0]["idx"]])


# ---- varian NETRAL (kontrol negatif): bukan cacat, sistem harus tetap menilai sesuai -----------------------------
BENIGN_ABBR = " Kolom dfl dan pkg dibaca dari berkas src, lalu parameter cfg dan env diatur."
BENIGN_CODE = " Berkas model.pkl dan pkg_1 disimpan di src/utils; contoh tersedia di https://github.com/user/pkg/dfl."


def _append_benign(c, text):
    items = c.body(15)
    if not items:
        raise NotApplicable("tidak ada paragraf isi")
    for info in c.pick(items, frac=0.2, n_min=2):
        par = c.paras[info.idx]
        t = E.par_text(par)
        c.mark(par)
        E.replace_range(par, len(t), len(t), text)


@op("N14a", 14, "NETRAL: singkatan huruf kecil (dfl, pkg, src, cfg, env) tanpa miring; seharusnya BUKAN istilah asing",
    kind="benign")
def n14a(c):
    _append_benign(c, BENIGN_ABBR)


@op("N14b", 14, "NETRAL: nama berkas, identifier kode, jalur, dan URL tanpa miring; seharusnya BUKAN istilah asing",
    kind="benign")
def n14b(c):
    _append_benign(c, BENIGN_CODE)


# ---- R01 verifikasi Crossref (varian referensi) --------------------------------------------------
def _ref_title_span(text):
    m = re.search(r"[\u201c\"](.+?)[\u201d\"]", text)
    return (m.start(1), m.end(1)) if m else None


@op("R01", 18, "verifikasi Crossref: DOI tak terdaftar / judul tak cocok / referensi fiktif", kind="ref")
def r01(c):
    refs = [r for r in c.pd.references if r["doi"] and r["has_quoted_title"] and r["num"] is not None]
    if len(refs) < 3:
        raise NotApplicable("kurang dari 3 referensi ber-DOI dan berjudul")
    sel = c.rng.sample(refs, min(9, len(refs)))
    ga, gb, gc = sel[0::3], sel[1::3], sel[2::3]
    spare = [r for r in refs if r not in sel]
    labels = {}
    for r in ga:                                                  # (a) DOI tidak terdaftar
        par = c.paras[r["idx"]]
        text = E.par_text(par)
        m = m2_parser.DOI_RE.search(text)
        end = m.start() + len(m.group(0).rstrip(".,;)"))
        c.mark(par)
        E.replace_range(par, end, end, "zz9")
        labels[r["num"]] = "R01a"
    titles = []
    for r in gb:
        s, e = _ref_title_span(E.par_text(c.paras[r["idx"]]))
        titles.append(E.par_text(c.paras[r["idx"]])[s:e])
    if len(gb) == 1 and spare:
        s, e = _ref_title_span(E.par_text(c.paras[spare[0]["idx"]]))
        titles = [E.par_text(c.paras[spare[0]["idx"]])[s:e]]
    elif len(gb) > 1:
        titles = titles[1:] + titles[:1]
    for r, new_title in zip(gb, titles):                          # (b) judul tidak cocok dengan DOI
        par = c.paras[r["idx"]]
        s, e = _ref_title_span(E.par_text(par))
        c.mark(par)
        E.replace_range(par, s, e, new_title)
        labels[r["num"]] = "R01b"
    for n, r in enumerate(gc, 1):                                 # (c) referensi fiktif tanpa DOI
        par = c.paras[r["idx"]]
        text = E.par_text(par)
        pre = re.match(r"^\[\d+\]\s*", text)
        fake = (f"A. Fiktif, \u201cPenerapan Metode Hipotetis untuk Verifikasi Tidak Terdaftar {n},\u201d "
                "Jurnal Fiktif Nusantara, vol. 99, no. 9, pp. 1\u201310, 2024.")
        c.mark(par)
        E.replace_range(par, len(pre.group(0)) if pre else 0, len(text), fake)
        labels[r["num"]] = "R01c"
    c.state["ref_labels"] = labels


# PEMBANGUNAN VARIAN
def sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def prefix_of(base):
    stem = Path(base).stem
    m = re.match(r"^(.*?)[-_ ]?base$", stem, re.I)
    return m.group(1) if m and m.group(1) else stem


def para_texts(path):
    return [E.par_text(p) for _, p, _ in m2_parser.iter_paragraphs(Document(str(path)))]


def build_variant(base, codes, key, seed, workdir, soffice, params):
    """Terapkan `codes` (urut ORDER) pada salinan basis. Kembalikan (path_kerja, defects, skipped, state)."""
    work = Path(workdir) / f"{key}.docx"
    shutil.copy(base, work)
    rng = random.Random(f"{seed}:{key}")
    defects, skipped, state = [], [], {}
    for code in sorted(codes, key=lambda cd: ORDER.index(OPS[cd].criterion)):
        o = OPS[code]
        ctx = Ctx(work, rng, soffice, params)
        try:
            o.func(ctx)
        except NotApplicable as exc:
            skipped.append({"code": code, "reason": str(exc)})
            continue
        except Exception as exc:                      # satu operator gagal tidak boleh menghentikan batch
            skipped.append({"code": code, "reason": f"galat penyuntingan ({type(exc).__name__}: {exc})"})
            print(f"  ! {key}: {code} dilewati karena galat: {type(exc).__name__}: {exc}")
            continue
        ctx.save()
        locs = []
        for par, before, role in ctx.marks:
            attached = par._p.getparent() is not None
            locs.append({"role": role, "before": before, "after": E.par_text(par)[:80] if attached else None})
        rec = {"code": code, "criterion": o.criterion, "description": o.desc,
               "side_effects": list(o.side), "locations": locs}
        if "ref_labels" in ctx.state:
            rec["reference_labels"] = ctx.state["ref_labels"]
        defects.append(rec)
        state.update(ctx.state)
    return work, defects, skipped, state


def finalize(work, out_path, defects, state):
    """Tulis metadata halaman (bila ada), salin ke folder keluaran, dan resolusikan indeks paragraf akhir."""
    if "pages" in state:
        E.patch_pages(work, state["pages"])
    E.normalize_zip(work)                                   # cap waktu ZIP seragam -> hasil deterministik
    shutil.copy(work, out_path)
    texts = para_texts(out_path)
    for d in defects:
        for loc in d["locations"]:
            key = (loc["after"] or "")[:40]
            loc["idx_final"] = next((i for i, t in enumerate(texts) if key and t.startswith(key)), None)
    return texts


def changed_paragraphs(base_texts, texts):
    sm = difflib.SequenceMatcher(None, base_texts, texts, autojunk=False)
    return sum(max(i2 - i1, j2 - j1) for tag, i1, i2, j1, j2 in sm.get_opcodes() if tag != "equal")


def partition(rng, ok_by_crit, sizes):
    crits = sorted(ok_by_crit)
    for _ in range(3000):
        rng.shuffle(crits)
        groups, i = [], 0
        for k in sizes:
            groups.append(crits[i:i + k])
            i += k
        if all(frozenset((a, b)) not in CONFLICTS for g in groups for n, a in enumerate(g) for b in g[n + 1:]):
            return groups
    raise RuntimeError("tidak menemukan pembagian kriteria tanpa konflik")


def pick_eff_set(rng, ok_by_crit, n=6):
    crits = sorted(ok_by_crit)
    for _ in range(3000):
        s = rng.sample(crits, min(n, len(crits)))
        if all(frozenset((a, b)) not in CONFLICTS for i, a in enumerate(s) for b in s[i + 1:]):
            return s
    raise RuntimeError("tidak menemukan himpunan efisiensi tanpa konflik")


def process_base(base, out, args, soffice, params):
    base = Path(base)
    prefix = prefix_of(base)
    work = Path(tempfile.mkdtemp(prefix=f"mv_{prefix}_"))
    base_texts = para_texts(base)
    entries, skipped, ok_by_crit = [], [], {}

    def emit(name, vtype, defects, sk, state, work_path):
        path = Path(out) / f"{prefix}-{name}.docx"
        texts = finalize(work_path, path, defects, state)
        try:
            Document(str(path))
            valid = True
        except Exception:
            valid = False
        entries.append({"file": path.name, "type": vtype, "defects": defects, "sha256": sha256(path),
                        "changed_paragraphs": changed_paragraphs(base_texts, texts), "valid_docx": valid,
                        "pages": state.get("pages"), "pages_source": state.get("pages_source")})
        skipped.extend({"variant": path.name, **s} for s in sk)
        return path

    # kontrol
    ctl = Path(out) / f"{prefix}-V00.docx"
    shutil.copy(base, ctl)
    entries.append({"file": ctl.name, "type": "control", "defects": [], "sha256": sha256(ctl),
                    "changed_paragraphs": 0, "valid_docx": True, "pages": None, "pages_source": None})

    # tunggal (std; + ext bila aktif). --only boleh menyebut kode std/ext/benign apa pun.
    only = set(args.only.split(",")) if args.only else None
    use_ext = getattr(args, "ext", False)            # tanpa atribut (pemanggilan programatik lama) = perilaku v1.0
    use_benign = getattr(args, "benign", False)
    pool_ext = getattr(args, "pool_ext", False)
    for code in [c for c, o in sorted(OPS.items(), key=lambda kv: (kv[1].criterion, kv[0]))
                 if o.kind == "std" or (o.kind == "ext" and (use_ext or only))]:
        if only and code not in only:
            continue
        w, d, sk, st = build_variant(base, [code], code, args.seed + f":{prefix}", work, soffice, params)
        if not d:
            skipped.extend({"variant": f"{prefix}-{code}", **s} for s in sk)
            print(f"  - {prefix}-{code}: dilewati ({sk[0]['reason']})")
            continue
        emit(code, "single", d, sk, st, w)
        if OPS[code].kind == "std" or pool_ext:      # operator ext di luar kolam -> varian majemuk/EA-EB lama tak berubah
            ok_by_crit.setdefault(OPS[code].criterion, []).append(code)
    # netral (kontrol negatif): perubahan yang bukan cacat; sistem harus tetap menilai sesuai
    for code in [c for c, o in sorted(OPS.items(), key=lambda kv: (kv[1].criterion, kv[0])) if o.kind == "benign"]:
        if (only and code not in only) or (not only and not use_benign):
            continue
        w, d, sk, st = build_variant(base, [code], code, args.seed + f":{prefix}", work, soffice, params)
        if not d:
            skipped.extend({"variant": f"{prefix}-{code}", **s} for s in sk)
            print(f"  - {prefix}-{code}: dilewati ({sk[0]['reason']})")
            continue
        emit(code, "benign", d, sk, st, w)
        entries[-1]["expected"] = "unchanged"        # kriteria terkait harus berstatus sama dengan kontrol V00
    full = not only

    rng = random.Random(f"{args.seed}:{prefix}:plan")
    if full and not args.no_multi and ok_by_crit:
        sizes = [int(x) for x in args.sizes.split(",")]
        for k, group in zip(sizes, partition(rng, ok_by_crit, sizes)):
            codes = [rng.choice(ok_by_crit[cr]) for cr in group]
            w, d, sk, st = build_variant(base, codes, f"M{k:02d}", args.seed + f":{prefix}", work, soffice, params)
            emit(f"M{k:02d}", "multi", d, sk, st, w)
    if full and not args.no_ref and "R01" in OPS:
        w, d, sk, st = build_variant(base, ["R01"], "R01", args.seed + f":{prefix}", work, soffice, params)
        if d:
            emit("R01", "ref", d, sk, st, w)
        else:
            skipped.extend({"variant": f"{prefix}-R01", **s} for s in sk)
    if full and not args.no_eff and ok_by_crit:
        crits = pick_eff_set(rng, ok_by_crit)
        codes = [rng.choice(ok_by_crit[cr]) for cr in crits]
        for name in ("EA", "EB"):
            w, d, sk, st = build_variant(base, codes, name, args.seed + f":{prefix}", work, soffice, params)
            emit(name, "eff", d, sk, st, w)
    shutil.rmtree(work, ignore_errors=True)
    return {"base": base.name, "prefix": prefix, "sha256": sha256(base), "variants": entries}, skipped


# KELUARAN & SELFCHECK
def write_outputs(out, bases, skipped, args):
    out = Path(out)
    catalog = {c: {"criterion": o.criterion, "description": o.desc, "side_effects": list(o.side), "kind": o.kind}
               for c, o in OPS.items()}
    manifest = {"generator": "make_variants.py", "version": VERSION, "seed": args.seed,
                "created": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "catalog": catalog, "bases": bases}
    (out / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    with open(out / "manifest.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["file", "base", "type", "defect_code", "criterion", "description", "idx_final", "before", "after",
                    "side_effects", "changed_paragraphs", "sha256"])
        for b in bases:
            for v in b["variants"]:
                rows = v["defects"] or [{"code": "V00", "criterion": "", "description": "kontrol (tanpa cacat)",
                                         "locations": [], "side_effects": []}]
                for d in rows:
                    loc = d["locations"][0] if d["locations"] else {}
                    w.writerow([v["file"], b["base"], v["type"], d["code"], d["criterion"], d["description"],
                                loc.get("idx_final", ""), loc.get("before", ""), loc.get("after", ""),
                                ";".join(f"K{n:02d}" for n in d["side_effects"]), v["changed_paragraphs"], v["sha256"]])
    with open(out / "skipped.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["variant", "code", "reason"])
        for s in skipped:
            w.writerow([s["variant"], s["code"], s["reason"]])


def _status_runner(rules_path=None):
    """Kembalikan fungsi path -> {id kriteria: status} (M3 + M4, tanpa Crossref)."""
    from checker import m4_citations
    db = m3_rules.load_rules(rules_path)
    eng = m3_rules.RuleEngine(db)
    crit14 = next(c for c in db["kriteria"] if c["modul"] == "M4")

    def run(path):
        pd = m2_parser.parse(str(path))
        res = eng.run(pd, {"jenis": "orisinal"}) + [m4_citations.run(pd, crit14)]
        return {r.id: r.status for r in res}
    return run


def check_bases(files, rules_path=None):
    """Peringatkan bila basis tidak 100% sesuai (basis harus bebas cacat agar pengukuran sah). Tanpa Crossref."""
    run = _status_runner(rules_path)
    for f in files:
        bad = {k: v for k, v in run(f).items() if v != "pass"}
        if bad:
            print(f"  ! BASIS {Path(f).name} belum 100% sesuai: " + ", ".join(f"K{k:02d}={v}" for k, v in sorted(bad.items()))
                  + "  (hasil evaluasi pada kriteria ini tidak sahih)")
        else:
            print(f"  basis {Path(f).name}: 18/18 sesuai")


def selfcheck(out, bases, rules_path=None):
    """Jalankan sistem pemeriksa pada hasil (tanpa Crossref). BUKAN evaluasi independen: hanya memeriksa bahwa
    cacat tersisip terbaca, varian netral tidak memicu alarm, dan memperlihatkan perubahan status pada kriteria lain
    (isolasi). Evaluasi menyeluruh per kriteria: tools/run_batch.py."""
    run = _status_runner(rules_path)
    rows = []
    for b in bases:
        base_status = run(Path(out) / f"{b['prefix']}-V00.docx")
        for v in b["variants"]:
            if v["type"] in ("control", "ref"):
                continue
            st = run(Path(out) / v["file"])
            if v["type"] == "benign":
                want = {d["criterion"] for d in v["defects"]}
                for d in v["defects"]:
                    rows.append([v["file"], d["code"], d["criterion"], st.get(d["criterion"]) == base_status[d["criterion"]], "netral"])
                extra = sorted(k for k in st if st[k] != base_status[k] and k not in want)
                if extra:
                    rows.append([v["file"], "EFEK-LAIN", extra, False, "efek-lain"])
                continue
            want = {d["criterion"] for d in v["defects"]}
            side = {n for d in v["defects"] for n in d["side_effects"]}
            for d in v["defects"]:
                rows.append([v["file"], d["code"], d["criterion"], st.get(d["criterion"]) != "pass", "cacat"])
            extra = sorted(k for k in st if st[k] != base_status[k] and k not in want and k not in side)
            if extra:
                rows.append([v["file"], "EFEK-LAIN", extra, False, "efek-lain"])
    with open(Path(out) / "selfcheck.csv", "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(["file", "code", "criterion", "terbaca", "jenis"])      # netral: terbaca=True berarti TIDAK ada alarm palsu
        w.writerows(rows)
    det = [r for r in rows if r[4] == "cacat"]
    neu = [r for r in rows if r[4] == "netral"]
    print(f"\nSelfcheck: {sum(r[3] for r in det)}/{len(det)} cacat terbaca oleh sistem")
    if neu:
        print(f"Selfcheck: {sum(r[3] for r in neu)}/{len(neu)} varian netral tanpa alarm palsu")
    for r in rows:
        if r[4] == "efek-lain":
            print(f"  ! {r[0]}: kriteria lain ikut berubah status: {['K%02d' % k for k in r[2]]}")
        elif not r[3] and r[4] == "cacat":
            print(f"  ! {r[0]}: {r[1]} TIDAK terbaca (operator kurang kuat atau celah sistem; periksa manual)")
        elif not r[3]:
            print(f"  ! {r[0]}: {r[1]} memicu ALARM PALSU (sistem salah menandai perubahan netral)")


def main(argv=None):
    ap = argparse.ArgumentParser(description="Pembangkit varian naskah (controlled error injection)")
    ap.add_argument("bases", nargs="+", help="naskah dasar bebas cacat (.docx), boleh pola glob")
    ap.add_argument("--out", default="varian", help="folder keluaran")
    ap.add_argument("--seed", default="42", help="seed (teks atau angka)")
    ap.add_argument("--only", help="hanya kode tunggal ini, mis. K02a,K06b (menonaktifkan multi/ref/eff)")
    ap.add_argument("--sizes", default="3,6,9", help="jumlah cacat varian majemuk (jumlahnya = 18 kriteria)")
    ap.add_argument("--no-multi", action="store_true")
    ap.add_argument("--no-ref", action="store_true")
    ap.add_argument("--no-eff", action="store_true")
    ap.add_argument("--soffice", help="path LibreOffice (soffice) untuk mengukur jumlah halaman")
    ap.add_argument("--rules", help="rules.json alternatif (untuk parameter, mis. istilah asing di K14)")
    ap.add_argument("--selfcheck", action="store_true", help="jalankan sistem pemeriksa pada hasil (sanity check)")
    ap.add_argument("--no-ext", dest="ext", action="store_false",
                    help="tanpa operator tambahan v1.1 (K05d/e, K07e/f, K08d, K09c): perilaku v1.0")
    ap.add_argument("--no-benign", dest="benign", action="store_false", help="tanpa varian netral (N14a/b)")
    ap.add_argument("--pool-ext", action="store_true",
                    help="masukkan operator tambahan ke kolam varian majemuk/EA-EB (MENGUBAH M03/M06/M09/EA/EB vs v1.0)")
    ap.add_argument("--no-check-base", dest="check_base", action="store_false", help="lewati pengecekan basis 18/18")
    ap.set_defaults(ext=True, benign=True, check_base=True)
    args = ap.parse_args(argv)

    files = []
    for pat in args.bases:
        files += sorted(glob.glob(pat)) or [pat]
    missing = [f for f in files if not Path(f).is_file()]
    if missing:
        ap.error(f"berkas tidak ditemukan: {missing}")
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    params = {c["id"]: c["params"] for c in m3_rules.load_rules(args.rules)["kriteria"]}

    if args.check_base:
        print("Pengecekan basis (tanpa Crossref):")
        check_bases(files, args.rules)
        print()
    bases, skipped = [], []
    for f in files:
        print(f"Basis: {f}")
        b, sk = process_base(f, out, args, args.soffice, params)
        bases.append(b)
        skipped += sk
        print(f"  -> {len(b['variants'])} berkas, {len(sk)} operator dilewati")
    write_outputs(out, bases, skipped, args)
    print(f"\nSelesai: {sum(len(b['variants']) for b in bases)} varian di {out.resolve()} (manifest.json, manifest.csv, skipped.csv)")
    if args.selfcheck:
        selfcheck(out, bases, args.rules)
    return 0


if __name__ == "__main__":
    sys.exit(main())
