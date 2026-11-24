import struct
import sys
import zlib
from pathlib import Path

import pytest
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

LOREM = ("Penelitian ini membahas penerapan metode pembelajaran mesin untuk mengklasifikasikan dokumen "
         "berbahasa Indonesia secara otomatis dengan akurasi yang tinggi dan waktu komputasi yang efisien")


def _png(path):
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * 8 for _ in range(8))
    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    data = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 8, 8, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))
    Path(path).write_bytes(data)


def _add_page_field(par):
    run = par.add_run()
    for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = text
        run._r.append(el)


def _hf_table(part, *fillers):
    """Isi header/footer di dalam tabel satu baris (seperti template). Tiap `filler(paragraf)` mengisi satu sel.
    Paragraf awal part tetap ada di bawah tabel sebagai paragraf kosong penutup (syarat Word)."""
    part.is_linked_to_previous = False
    tbl = part.add_table(1, len(fillers), Cm(16.5))
    part._element.insert(0, tbl._tbl)
    for cell, fill in zip(tbl.rows[0].cells, fillers):
        fill(cell.paragraphs[0])


def unwrap_tables(root_el):
    """Ganti tiap tabel tingkat-atas pada `root_el` (body/header/footer) dengan paragraf isi selnya:
    tampilan mirip, tetapi isi tidak lagi berada di dalam tabel."""
    import copy
    for tbl in list(root_el.findall(qn("w:tbl"))):
        for tc in tbl.iter(qn("w:tc")):
            for p in tc.findall(qn("w:p")):
                tbl.addprevious(copy.deepcopy(p))
        tbl.getparent().remove(tbl)


def _identity_two_tables(d, abstract, keywords, corr_lines):
    """Kotak identitas terpecah: tabel 1 = Informasi/Riwayat + Abstrak, tabel 2 = Kata Kunci + Korespondensi
    (tiap elemen di dalam tabel, tetapi bukan satu tabel seperti template)."""
    t1 = d.add_table(rows=2, cols=2)
    t1.cell(0, 0).text, t1.cell(0, 1).text = "Informasi Artikel", "ABSTRAK"
    t1.cell(1, 0).text = "Riwayat Artikel:"
    t1.cell(1, 1).text = abstract
    d.add_paragraph()                                   # pemisah agar Word tidak menyatukan kedua tabel
    t2 = d.add_table(rows=2, cols=1)
    kw = t2.cell(0, 0)
    kw.paragraphs[0].text = "Kata Kunci:"
    for k in keywords:
        kw.add_paragraph(k)
    co = t2.cell(1, 0)
    co.paragraphs[0].text = "Penulis Korespondensi:"
    for line in corr_lines:
        co.add_paragraph(line)


def _identity_table(d, abstract, keywords, corr_lines):
    """Kotak identitas seperti template JISINTA: Informasi/Riwayat Artikel, Kata Kunci, Penulis Korespondensi
    di kolom kiri; Abstrak di kolom kanan (sel digabung vertikal)."""
    t = d.add_table(rows=4, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Informasi Artikel", "ABSTRAK"
    left = t.cell(1, 0)
    left.paragraphs[0].text = "Riwayat Artikel:"
    for line in ("Diajukan 9 Juni 2026", "Direvisi 20 Juni 2026", "Diterima (dikosongkan)"):
        left.add_paragraph(line)
    kw = t.cell(2, 0)
    kw.paragraphs[0].text = "Kata Kunci:"
    for k in keywords:
        kw.add_paragraph(k)
    co = t.cell(3, 0)
    co.paragraphs[0].text = "Penulis Korespondensi:"
    for line in corr_lines:
        co.add_paragraph(line)
    merged = t.cell(1, 1).merge(t.cell(3, 1))
    merged.paragraphs[0].text = abstract
    for extra in merged.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)
    return t


def _setup_headers(d, variant):
    """variant: ok | none | no_flags | no_even_text | plain.
    Seperti template: header pertama/genap dan semua footer berada di dalam tabel, header ganjil paragraf biasa.
    'plain' = semua paragraf biasa (naskah yang tampak sama tetapi tidak memakai tabel)."""
    s = d.sections[0]
    if variant == "none":
        return
    if variant != "no_flags":
        s.different_first_page_header_footer = True
        d.settings.odd_and_even_pages_header_footer = True
    plain = variant == "plain"
    T_FIRST = "JURNAL INFORMATIKA DAN SISTEM INFORMASI UNTIRTA"
    T_ODD = "JISINTA | Volume 1 No. 1 Page 1-8 | Januari-2026"
    T_EVEN = "JISINTA: Jurnal Informatika dan Sistem Informasi UNTIRTA E-ISSN 3163-7434"
    even_text = "Klasifikasi Dokumen Bahasa Indonesia \u2014 Aminah dan Santoso  " if variant != "no_even_text" else ""

    def text(t):
        return lambda par: par.add_run(t)

    s.header.is_linked_to_previous = False
    s.header.paragraphs[0].text = T_ODD                                   # header ganjil: paragraf biasa (template)
    for part, t in ((s.first_page_header, T_FIRST), (s.even_page_header, T_EVEN)):
        if plain:
            part.is_linked_to_previous = False
            part.paragraphs[0].text = t
        else:
            _hf_table(part, text(t))
    for part in (s.first_page_footer, s.footer):
        if plain:
            part.is_linked_to_previous = False
            _add_page_field(part.paragraphs[0])
        else:
            _hf_table(part, lambda par: _add_page_field(par))
    if plain:
        s.even_page_footer.is_linked_to_previous = False
        fp = s.even_page_footer.paragraphs[0]
        fp.add_run(even_text)
        _add_page_field(fp)
    else:
        _hf_table(s.even_page_footer, text(even_text), lambda par: _add_page_field(par))


def make_docx(path, *, margin_top=2.5, title=None, abstract_words=130, keywords=("akurasi", "klasifikasi", "pembelajaran mesin"),
              n_refs=20, bad_cite=False, with_conclusion=True, caption=True,
              authors="ok", hf="ok", identity="table"):
    d = Document()
    for st in ("Normal",):
        d.styles[st].font.name, d.styles[st].font.size = "Cambria", Pt(10)
    d.styles["Normal"].paragraph_format.line_spacing = 1.15      # sesuai style Normal pada template
    s = d.sections[0]
    s.page_width, s.page_height = Cm(21.0), Cm(29.7)
    s.top_margin, s.left_margin, s.right_margin, s.bottom_margin = Cm(margin_top), Cm(2.5), Cm(2), Cm(2)

    d.add_paragraph(title or "Klasifikasi Dokumen Berbahasa Indonesia Menggunakan Pembelajaran Mesin", style="Title")
    a = d.add_paragraph("Siti Aminah\u00b9, Budi Santoso\u00b2" if authors != "bad" else "Siti Aminah\u00b9, Budi Santoso\u00b3, Dewi Lestari")
    a.alignment = WD_ALIGN_PARAGRAPH.CENTER if authors != "bad" else WD_ALIGN_PARAGRAPH.LEFT
    a.runs[0].font.size = Pt(10 if authors != "bad" else 12)
    a.paragraph_format.space_after = Pt(6)
    if authors != "no_aff":
        f = d.add_paragraph("\u00b9Program Studi Informatika, Universitas Contoh \u00b2Program Studi Matematika, Universitas Contoh")
        f.alignment = WD_ALIGN_PARAGRAPH.CENTER
        f.paragraph_format.space_after = Pt(6)
        f.runs[0].font.size = Pt(9 if authors != "bad" else 10)
    words = (LOREM + " ").split() * 10
    abstract = " ".join(words[:abstract_words])
    corr = (["Nama Penulis Korespondensi", "Universitas Contoh"] if authors == "bad"
            else ["Siti Aminah", "Program Studi Informatika, Universitas Contoh", "siti.aminah@contoh.ac.id"])
    if identity == "table":                                  # kotak identitas dalam tabel (sesuai template)
        _identity_table(d, abstract, keywords, corr)
    elif identity == "split":                                # tiap elemen di tabel, tetapi tidak satu tabel
        _identity_two_tables(d, abstract, keywords, corr)
    else:                                                    # identitas diketik sebagai paragraf biasa (salah)
        d.add_paragraph("Abstrak: " + abstract)
        d.add_paragraph("Kata kunci: " + "; ".join(keywords))
        d.add_paragraph("Penulis Korespondensi:")
        for line in corr:
            d.add_paragraph(line)

    def body(text):
        p = d.add_paragraph(text); p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.27); p.paragraph_format.space_after = Pt(0)
        return p

    body_sec = d.add_section(WD_SECTION.CONTINUOUS)          # identitas 1 kolom | isi 2 kolom (K03)
    body_sec.different_first_page_header_footer = False
    cols = body_sec._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        body_sec._sectPr.find(qn("w:pgMar")).addnext(cols)
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), "708")
    d.add_heading("1. PENDAHULUAN", 1)
    body("Latar belakang masalah klasifikasi teks sangat penting dalam banyak bidang aplikasi modern saat ini [1], [2].")
    body("Penelitian terdahulu telah mengusulkan berbagai pendekatan klasifikasi dokumen yang cukup beragam, relevan, dan banyak dirujuk oleh peneliti lain [3]\u2013[7].")
    body("Kesenjangan penelitian terletak pada kurangnya studi untuk dokumen berbahasa Indonesia sehingga tujuan penelitian ini jelas [8].")
    d.add_heading("2. METODE PENELITIAN", 1)
    body("Metode yang digunakan adalah eksperimen kuantitatif dengan data dari sumber terbuka yang dikumpulkan pada tahun terakhir [9], [10] dan hasilnya ditampilkan pada Gambar 1 dan Tabel 1.")
    if caption:
        png = Path(path).with_suffix(".png")
        _png(png)
        d.add_picture(str(png), width=Cm(3))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    if caption:
        c = d.add_paragraph("Gambar 1. Contoh diagram alur"); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.runs[0].font.size = Pt(8)
        c = d.add_paragraph("Tabel 1. Hasil pengujian"); c.alignment = WD_ALIGN_PARAGRAPH.CENTER
        c.runs[0].font.size = Pt(8)
    t = d.add_table(rows=2, cols=2); t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.cell(0, 0).text, t.cell(0, 1).text = "Model", "Akurasi"
    t.cell(1, 0).text, t.cell(1, 1).text = "SVM", "0,91"
    d.add_heading("3. HASIL DAN PEMBAHASAN", 1)
    body("Hasil pengujian menunjukkan peningkatan akurasi dibandingkan penelitian terdahulu yang dilaporkan dalam literatur [11]\u2013[15].")
    body(f"Pembahasan menunjukkan bahwa fitur tambahan memberi kontribusi signifikan terhadap performa model secara keseluruhan [16], [17], [19], [20], [{99 if bad_cite else 20}].")
    if with_conclusion:
        d.add_heading("4. KESIMPULAN", 1)
        body("Model yang diusulkan efektif untuk tugas ini dan memberi kontribusi praktis. Penelitian selanjutnya disarankan menguji korpus yang lebih besar [18].")
    d.add_heading("REFERENSI", 1)
    for i in range(1, n_refs + 1):
        p = d.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.13); p.paragraph_format.first_line_indent = Cm(-1.13)
        r1 = p.add_run(f"[{i}] A. Penulis, \u201cJudul Artikel Nomor {i} tentang Klasifikasi,\u201d ")
        r2 = p.add_run("Jurnal Informatika"); r2.italic = True
        r3 = p.add_run(f", vol. {i}, no. 2, pp. 1\u201310, 2025, doi: 10.1234/jurnal.{i}.")
        for r in (r1, r2, r3):
            r.font.size = Pt(8)
    _setup_headers(d, hf)
    d.save(path)
    return path


def crossref_ok(url, params=None):
    """Mock Crossref: semua DOI 'terdaftar' dengan judul yang cocok."""
    import re
    m = re.search(r"jurnal\.(\d+)$", url)
    if m:
        n = m.group(1)
        return 200, {"message": {"title": [f"Judul Artikel Nomor {n} tentang Klasifikasi"],
                                 "DOI": f"10.1234/jurnal.{n}", "issued": {"date-parts": [[2025]]}}}
    return 404, None


@pytest.fixture
def good_docx(tmp_path):
    return make_docx(tmp_path / "bagus.docx")


@pytest.fixture
def mock_crossref(monkeypatch):
    from checker import m5_crossref
    monkeypatch.setattr(m5_crossref, "http_get", crossref_ok)
