"""Isi tabel (K15), baris penulis korespondensi (K07), dan spasi kosong (K14)."""
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from checker import m2_parser, m3_rules
from variant_base import build_base


def hasil(path, kode):
    pd = m2_parser.parse(path)
    engine = m3_rules.RuleEngine(m3_rules.load_rules())
    return next(x for x in engine.run(pd, {"jenis": "orisinal"}) if x.code == kode)


def ubah(tmp_path, fungsi):
    src = tmp_path / "base.docx"
    build_base(src)
    d = Document(src)
    fungsi(d)
    out = tmp_path / "ubah.docx"
    d.save(out)
    return out


def tabel_isi(d):
    return d.tables[-1]


def sel_korespondensi(d):
    return d.tables[0].cell(3, 0)


def paragraf_isi(d):
    return next(p for p in d.paragraphs if p.text.startswith("Perkembangan platform digital"))


# ---------------------------------------------------------------- K15: isi tabel
def test_dasar_sesuai(tmp_path):
    src = tmp_path / "base.docx"
    build_base(src)
    for k in ("K07", "K14", "K15"):
        assert hasil(src, k).status == "pass", k


def test_tabel_ukuran_salah(tmp_path):
    def f(d):
        for r in tabel_isi(d).cell(1, 0).paragraphs[0].runs:
            r.font.size = Pt(10)
    r = hasil(ubah(tmp_path, f), "K15")
    assert any("tidak berukuran 8" in x.message for x in r.findings)


def test_tabel_font_salah(tmp_path):
    def f(d):
        for r in tabel_isi(d).cell(1, 1).paragraphs[0].runs:
            r.font.name = "Arial"
    assert any("font Cambria" in x.message for x in hasil(ubah(tmp_path, f), "K15").findings)


def test_tabel_tidak_rata_tengah(tmp_path):
    def f(d):
        tabel_isi(d).cell(1, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    assert any("rata tengah" in x.message for x in hasil(ubah(tmp_path, f), "K15").findings)


def test_baris_pertama_tidak_tebal(tmp_path):
    def f(d):
        for r in tabel_isi(d).cell(0, 0).paragraphs[0].runs:
            r.font.bold = False
    assert any("baris pertama" in x.message for x in hasil(ubah(tmp_path, f), "K15").findings)


def test_baris_data_tebal_tidak_masalah(tmp_path):
    def f(d):
        for r in tabel_isi(d).cell(1, 0).paragraphs[0].runs:
            r.font.bold = True
    assert hasil(ubah(tmp_path, f), "K15").status == "pass"


# ---------------------------------------------------------------- K07: penulis korespondensi per baris
def test_korespondensi_digabung_satu_baris(tmp_path):
    def f(d):
        c = sel_korespondensi(d)
        c.paragraphs[1].text = "Siti Aminah, Program Studi Informatika, Universitas Contoh Email: siti.aminah@contoh.ac.id"
        for q in c.paragraphs[2:]:
            q._p.getparent().remove(q._p)
    r = hasil(ubah(tmp_path, f), "K07")
    assert any("3 baris terpisah" in x.message for x in r.findings)


def test_korespondensi_isi_satu_baris_dengan_label(tmp_path):
    def f(d):
        c = sel_korespondensi(d)
        c.paragraphs[0].text = "Penulis Korespondensi: Siti Aminah"
    r = hasil(ubah(tmp_path, f), "K07")
    assert any("satu baris dengan labelnya" in x.message for x in r.findings)


def test_korespondensi_email_bukan_baris_terakhir(tmp_path):
    def f(d):
        c = sel_korespondensi(d)
        c.paragraphs[2].text = "siti.aminah@contoh.ac.id"
        c.paragraphs[3].text = "Program Studi Informatika, Universitas Contoh"
    r = hasil(ubah(tmp_path, f), "K07")
    assert any("baris terakhir" in x.message for x in r.findings)


def test_korespondensi_email_dengan_awalan_boleh(tmp_path):
    def f(d):
        sel_korespondensi(d).paragraphs[3].text = "Email: siti.aminah@contoh.ac.id"
    assert hasil(ubah(tmp_path, f), "K07").status == "pass"


# ---------------------------------------------------------------- K14: spasi kosong
def test_spasi_ganda(tmp_path):
    def f(d):
        p = paragraf_isi(d)
        p.runs[0].text = p.runs[0].text.replace("platform digital", "platform  digital", 1)
    r = hasil(ubah(tmp_path, f), "K14")
    assert any("spasi ganda" in x.message for x in r.findings)


def test_spasi_di_awal_paragraf(tmp_path):
    def f(d):
        p = paragraf_isi(d)
        p.runs[0].text = "  " + p.runs[0].text
    assert any("diawali atau diakhiri" in x.message for x in hasil(ubah(tmp_path, f), "K14").findings)


def test_spasi_sebelum_koma(tmp_path):
    def f(d):
        p = paragraf_isi(d)
        p.runs[0].text = p.runs[0].text.replace("platform digital", "platform digital ,", 1)
    assert any("sebelum tanda baca" in x.message for x in hasil(ubah(tmp_path, f), "K14").findings)


def test_baris_kosong_pengatur_jarak(tmp_path):
    def f(d):
        paragraf_isi(d).insert_paragraph_before("")
    r = hasil(ubah(tmp_path, f), "K14")
    assert any("baris kosong" in x.message for x in r.findings)


def test_jeda_halaman_bukan_baris_kosong(tmp_path):
    def f(d):
        from docx.enum.text import WD_BREAK
        p = paragraf_isi(d).insert_paragraph_before("")
        p.add_run().add_break(WD_BREAK.PAGE)
    assert hasil(ubah(tmp_path, f), "K14").status == "pass"
