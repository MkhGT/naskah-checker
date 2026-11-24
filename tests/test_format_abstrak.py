"""Format abstrak (Cambria 10 pt, justify) dan perataan kata kunci (kiri)."""
import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

from checker import m2_parser, m3_rules
from variant_base import build_base


def hasil(path, kode):
    pd = m2_parser.parse(path)
    engine = m3_rules.RuleEngine(m3_rules.load_rules())
    r = next(x for x in engine.run(pd, {"jenis": "orisinal"}) if x.code == kode)
    return r


def ubah(tmp_path, fungsi):
    src = tmp_path / "base.docx"
    build_base(src)
    d = Document(src)
    fungsi(d.tables[0])
    out = tmp_path / "ubah.docx"
    d.save(out)
    return out


def sel_abstrak(t):
    return t.cell(1, 1).paragraphs[0]


def sel_kata_kunci(t):
    return t.cell(2, 0)


def test_dokumen_dasar_sesuai(tmp_path):
    src = tmp_path / "base.docx"
    build_base(src)
    assert hasil(src, "K08").status == "pass"
    assert hasil(src, "K09").status == "pass"


def test_abstrak_ukuran_salah(tmp_path):
    def f(t):
        for r in sel_abstrak(t).runs:
            r.font.size = Pt(9)
    r = hasil(ubah(tmp_path, f), "K08")
    assert r.status == "warn"
    assert any("ukuran huruf" in x.message for x in r.findings)


def test_abstrak_font_salah(tmp_path):
    def f(t):
        for r in sel_abstrak(t).runs:
            r.font.name = "Arial"
    r = hasil(ubah(tmp_path, f), "K08")
    assert any("font" in x.message for x in r.findings)


def test_abstrak_rata_kiri_dianggap_salah(tmp_path):
    r = hasil(ubah(tmp_path, lambda t: setattr(sel_abstrak(t), "alignment", WD_ALIGN_PARAGRAPH.LEFT)), "K08")
    assert r.status == "warn"
    assert any("perataan" in x.message for x in r.findings)


def test_kata_kunci_justify_dianggap_salah(tmp_path):
    def f(t):
        for p in sel_kata_kunci(t).paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r = hasil(ubah(tmp_path, f), "K09")
    assert r.status == "warn"
    assert any("perataan" in x.message for x in r.findings)
