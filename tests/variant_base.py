"""Dokumen uji BEBAS CACAT (memenuhi 19 kriteria) untuk menguji pembangkit varian. Hanya untuk pengujian."""
import re
import struct
import zlib
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def _png(path):
    raw = b"".join(b"\x00" + b"\x20\x60\xc0" * 8 for _ in range(8))

    def chunk(t, d):
        c = struct.pack(">I", len(d)) + t + d
        return c + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    Path(path).write_bytes(b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 8, 8, 8, 2, 0, 0, 0))
                           + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def _field(par):
    run = par.add_run()
    for kind, text in (("begin", None), (None, " PAGE "), ("end", None)):
        if kind:
            el = OxmlElement("w:fldChar"); el.set(qn("w:fldCharType"), kind)
        else:
            el = OxmlElement("w:instrText"); el.set(qn("xml:space"), "preserve"); el.text = text
        run._r.append(el)


def marked(par, text, size=None):
    """Teks dengan penanda *istilah miring*."""
    for i, part in enumerate(re.split(r"\*", text)):
        if part:
            r = par.add_run(part)
            r.italic = bool(i % 2)
            if size:
                r.font.size = Pt(size)
    return par


INTRO = [
    "Perkembangan platform digital membuat jumlah ulasan pengguna meningkat pesat sehingga pengelolaan umpan balik "
    "secara manual menjadi tidak efisien bagi pengembang. Karena itu diperlukan pendekatan *machine learning* yang "
    "andal untuk mengelompokkan opini pengguna secara otomatis dan konsisten [1], [2].",
    "Sejumlah penelitian terdahulu telah menerapkan klasifikasi sentimen pada berbagai domain dengan memanfaatkan "
    "*dataset* publik dan fitur berbasis frekuensi kata. Hasil yang dilaporkan menunjukkan kinerja cukup baik, "
    "meskipun sebagian besar masih terbatas pada teks berbahasa Inggris [3]\u2013[6].",
    "Kesenjangan penelitian terletak pada minimnya studi untuk ulasan berbahasa Indonesia yang memuat singkatan "
    "dan bahasa gaul. Ulasan semacam itu memerlukan tahap *preprocessing* yang lebih teliti agar variasi penulisan "
    "tidak mengurangi kualitas fitur yang dihasilkan [7].",
    "Tujuan penelitian ini adalah membandingkan beberapa algoritma klasifikasi pada ulasan aplikasi seluler berbahasa Indonesia yang dikumpulkan dari toko aplikasi. "
    "Penelitian ini juga mengevaluasi pengaruh pemilihan fitur terhadap akurasi, presisi, serta waktu komputasi "
    "model yang diuji pada seluruh skenario [8], [9].",
]
METHOD = [
    "Data penelitian berupa ulasan pengguna yang dikumpulkan melalui proses *crawling* dari toko aplikasi selama "
    "tiga bulan, kemudian dilabeli secara manual oleh dua anotator dengan tingkat kesepakatan yang tinggi [10].",
    "Alur penelitian secara keseluruhan ditunjukkan pada Gambar 1, sedangkan ringkasan hasil pengujian awal "
    "disajikan pada Tabel 1 untuk memudahkan pembaca membandingkan setiap skenario pengujian [11], [12].",
    "Tahap pra-pemrosesan meliputi pembersihan simbol, pengubahan huruf menjadi huruf kecil, penghapusan kata henti, "
    "dan normalisasi kata tidak baku sesuai kamus yang disusun peneliti sebelum proses pelatihan dimulai [13].",
    "Normalisasi dilakukan dengan memetakan kata tidak baku ke bentuk bakunya menggunakan kamus yang disusun dari "
    "korpus ulasan, sehingga variasi penulisan yang sama dapat diperlakukan sebagai satu fitur yang konsisten [14].",
]
RESULT = [
    "Pengujian dilakukan dengan validasi silang lima lipatan dan menunjukkan bahwa model *support vector machine* "
    "memberi akurasi tertinggi dibandingkan algoritma lain pada seluruh skenario pembagian data [15], [16].",
    "Peningkatan kinerja paling besar terjadi setelah normalisasi kata diterapkan, yang konsisten dengan temuan "
    "penelitian sebelumnya bahwa kualitas teks masukan sangat memengaruhi performa klasifikasi [17], [18].",
    "Analisis kesalahan memperlihatkan bahwa ulasan bernada sarkasme masih sulit diklasifikasikan, sehingga "
    "keterbatasan ini perlu dibahas lebih lanjut dalam kaitannya dengan konteks penggunaan aplikasi [19], [20].",
]
CONCL = [
    "Model yang diusulkan terbukti efektif untuk mengklasifikasikan sentimen ulasan dan memberi kontribusi praktis "
    "bagi pengembang aplikasi dalam memantau kepuasan pengguna secara berkelanjutan.",
    "Penelitian selanjutnya disarankan menguji korpus yang lebih besar serta menambahkan model berbasis konteks "
    "untuk menangani ulasan bernada sarkasme yang belum teratasi.",
]
ABSTRACT = (
    "Ulasan pengguna aplikasi seluler berkembang pesat dan memuat opini yang berharga bagi pengembang. Penelitian ini "
    "bertujuan membandingkan beberapa algoritma untuk mengklasifikasikan sentimen ulasan berbahasa Indonesia. Data "
    "dikumpulkan dari toko aplikasi, dilabeli oleh dua anotator, lalu melalui tahap pembersihan dan normalisasi kata. "
    "Pengujian memakai validasi silang lima lipatan dengan metrik akurasi, presisi, dan waktu komputasi. Hasil "
    "menunjukkan bahwa algoritma terbaik mencapai akurasi yang lebih tinggi dibandingkan pembanding, dan normalisasi "
    "memberi peningkatan paling besar. Temuan ini menegaskan pentingnya kualitas teks masukan pada klasifikasi "
    "sentimen untuk bahasa dengan banyak variasi penulisan. Penelitian lanjutan dapat memperluas cakupan data dan "
    "menguji model berbasis konteks agar kinerja pada ulasan yang bernada sarkasme dapat ditingkatkan lagi.")


def build_base(path):
    path = Path(path)
    d = Document()
    d.styles["Normal"].font.name, d.styles["Normal"].font.size = "Cambria", Pt(10)
    d.styles["Normal"].paragraph_format.line_spacing = 1.15
    d.styles["Heading 3"].font.bold = False
    sec = d.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.top_margin, sec.left_margin, sec.right_margin, sec.bottom_margin = Cm(2.5), Cm(2.5), Cm(2.0), Cm(2.0)
    sec.different_first_page_header_footer = True
    d.settings.odd_and_even_pages_header_footer = True

    # header & footer: pertama / ganjil / genap (seperti template: tabel, kecuali header ganjil)
    def hf_table(part, *fillers):
        part.is_linked_to_previous = False
        tbl = part.add_table(1, len(fillers), Cm(16.5))
        part._element.insert(0, tbl._tbl)
        for cell, fill in zip(tbl.rows[0].cells, fillers):
            fill(cell.paragraphs[0])

    hf_table(sec.first_page_header, lambda p: p.add_run("JURNAL INFORMATIKA DAN SISTEM INFORMASI UNTIRTA"))
    sec.header.is_linked_to_previous = False
    sec.header.paragraphs[0].text = "JISINTA | Volume 1 No. 1 Page 1-8 | Januari-2026"
    hf_table(sec.even_page_header,
             lambda p: p.add_run("JISINTA: Jurnal Informatika dan Sistem Informasi UNTIRTA E-ISSN 3163-7434"))
    hf_table(sec.first_page_footer, _field)
    hf_table(sec.footer, _field)
    hf_table(sec.even_page_footer,
             lambda p: p.add_run("Klasifikasi Sentimen Ulasan Aplikasi \u2014 Aminah dan Santoso  "), _field)

    tp = d.add_paragraph("Klasifikasi Sentimen Ulasan Aplikasi Menggunakan Pembelajaran Mesin", style="Title")
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tp.paragraph_format.space_after = Pt(6)
    for r in tp.runs:
        r.font.name, r.font.size, r.font.bold = "Cambria", Pt(16), True
    a = d.add_paragraph("Siti Aminah\u00b9, Budi Santoso\u00b2")
    a.alignment = WD_ALIGN_PARAGRAPH.CENTER
    a.paragraph_format.space_after = Pt(6)
    a.runs[0].font.size = Pt(10)
    f = d.add_paragraph("\u00b9Program Studi Informatika, Universitas Contoh \u00b2Program Studi Sistem Informasi, Universitas Contoh")
    f.alignment = WD_ALIGN_PARAGRAPH.CENTER
    f.paragraph_format.space_after = Pt(6)
    f.runs[0].font.size = Pt(9)

    # kotak informasi (seperti template JISINTA)
    t = d.add_table(rows=4, cols=2)
    t.cell(0, 0).text, t.cell(0, 1).text = "Informasi Artikel", "ABSTRAK"
    left = t.cell(1, 0)
    left.paragraphs[0].text = "Riwayat Artikel:"
    for line in ("Diajukan 9 Juni 2026", "Direvisi 20 Juni 2026", "Diterima (dikosongkan)"):
        left.add_paragraph(line)
    kw = t.cell(2, 0)
    kw.paragraphs[0].text = "Kata Kunci:"
    for k in ("akurasi", "klasifikasi", "pembelajaran mesin", "sentimen", "ulasan"):
        kw.add_paragraph(k)
    co = t.cell(3, 0)
    co.paragraphs[0].text = "Penulis Korespondensi:"
    co.add_paragraph("Siti Aminah")
    co.add_paragraph("Program Studi Informatika, Universitas Contoh")
    co.add_paragraph("siti.aminah@contoh.ac.id")
    merged = t.cell(1, 1).merge(t.cell(3, 1))
    merged.paragraphs[0].text = ABSTRACT
    merged.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for extra in merged.paragraphs[1:]:
        extra._p.getparent().remove(extra._p)

    def body(text):
        p = d.add_paragraph()
        marked(p, text)
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Cm(1.27)
        p.paragraph_format.space_after = Pt(0)
        return p

    body_sec = d.add_section(WD_SECTION.CONTINUOUS)           # identitas 1 kolom | isi 2 kolom
    body_sec.different_first_page_header_footer = False
    cols = body_sec._sectPr.find(qn("w:cols"))
    if cols is None:
        cols = OxmlElement("w:cols")
        body_sec._sectPr.find(qn("w:pgMar")).addnext(cols)
    cols.set(qn("w:num"), "2")
    cols.set(qn("w:space"), "708")

    SPACING = {1: (6, 6), 2: (6, 3), 3: (6, 2)}

    def heading(text, level):
        h = d.add_heading(text, level)
        pf = h.paragraph_format
        pf.space_before, pf.space_after = Pt(SPACING[level][0]), Pt(SPACING[level][1])
        pf.left_indent, pf.first_line_indent = Cm(0.635), Cm(-0.635)          # nomor rata di 0, teks 0,25 in
        return h

    heading("1. PENDAHULUAN", 1)
    for x in INTRO:
        body(x)
    heading("2. METODE PENELITIAN", 1)
    heading("2.1. Pengumpulan Data", 2)
    body(METHOD[0])
    body(METHOD[1])
    png = path.with_suffix(".png")
    _png(png)
    d.add_picture(str(png), width=Cm(3))
    d.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    c = d.add_paragraph("Gambar 1. Alur penelitian")
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.runs[0].font.size = Pt(8)
    c = d.add_paragraph("Tabel 1. Ringkasan hasil pengujian awal")
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    c.runs[0].font.size = Pt(8)
    tb = d.add_table(rows=2, cols=2)
    tb.alignment = WD_TABLE_ALIGNMENT.CENTER
    tb.cell(0, 0).text, tb.cell(0, 1).text = "Model", "Akurasi"
    tb.cell(1, 0).text, tb.cell(1, 1).text = "SVM", "0,91"
    for ri, row in enumerate(tb.rows):                       # isi tabel: Cambria 8 pt, rata tengah, baris pertama tebal
        for cell in row.cells:
            for par in cell.paragraphs:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for r in par.runs:
                    r.font.name, r.font.size, r.font.bold = "Cambria", Pt(8), (ri == 0)
    heading("2.2. Pra-pemrosesan", 2)
    body(METHOD[2])
    heading("2.2.1. Normalisasi Kata", 3)
    body(METHOD[3])
    heading("3. HASIL DAN PEMBAHASAN", 1)
    heading("3.1. Kinerja Model", 2)
    body(RESULT[0])
    body(RESULT[1])
    heading("3.2. Analisis Kesalahan", 2)
    body(RESULT[2])
    heading("4. KESIMPULAN", 1)
    for x in CONCL:
        body(x)
    heading("REFERENSI", 1)
    for i in range(1, 21):
        p = d.add_paragraph()
        p.paragraph_format.left_indent = Cm(1.13)
        p.paragraph_format.first_line_indent = Cm(-1.13)
        r1 = p.add_run(f"[{i}] A. Penulis, \u201cJudul Artikel Nomor {i} tentang Klasifikasi Sentimen,\u201d ")
        r2 = p.add_run("Jurnal Informatika"); r2.italic = True
        r3 = p.add_run(f", vol. {i}, no. 2, pp. 1\u201310, {2022 + i % 4}, doi: 10.1234/jurnal.{i}.")
        for r in (r1, r2, r3):
            r.font.size = Pt(8)
    d.save(path)
    return path
