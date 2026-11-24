"""Batasan pengolah dokumen: sistem dirancang untuk .docx hasil Microsoft Word.
M1 menolak ukuran desimal (lazim pada ekspor Google Docs) dengan pesan jelas dan mengenali aplikasi pembuat;
berkas non-Word tanpa desimal tetap diperiksa tetapi laporan memuat peringatan."""
import json
import re
import zipfile

import pytest
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import app as app_mod
from app import create_app
from checker import m1_upload
from conftest import make_docx

GDOCS_MARKER = (b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                b'<go:gDocsCustomXmlDataStorage xmlns:go="http://customooxmlschemas.google.com/"/>')


def rewrite(src, dst, edits=None, drop=(), add=None):
    """Salin zip dengan mengubah isi bagian tertentu (edits: nama -> fungsi bytes->bytes), membuang, atau menambah bagian."""
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            if item.filename in drop:
                continue
            data = zi.read(item.filename)
            if edits and item.filename in edits:
                data = edits[item.filename](data)
            zo.writestr(item, data)
        for name, data in (add or {}).items():
            zo.writestr(name, data)
    return dst


def with_app(tmp_path, application, name):
    """Salinan dokumen dengan <Application> diganti."""
    src = make_docx(tmp_path / "w.docx")
    return rewrite(src, tmp_path / name,
                   edits={"docProps/app.xml": lambda b: re.sub(rb"<Application>.*?</Application>",
                                                                 b"<Application>" + application.encode() + b"</Application>", b, flags=re.S)})


def drop_app_xml(src, dst, add=None):
    """Hapus docProps/app.xml BESERTA relasi dan deklarasi tipenya (seperti berkas yang memang tidak memilikinya)."""
    return rewrite(
        src, dst, drop={"docProps/app.xml"}, add=add,
        edits={"_rels/.rels": lambda b: re.sub(rb'<Relationship [^>]*Target="docProps/app\.xml"[^>]*/>', b"", b),
               "[Content_Types].xml": lambda b: re.sub(rb'<Override [^>]*PartName="/docProps/app\.xml"[^>]*/>', b"", b)})


def gdocs_like(tmp_path, name="g.docx"):
    """Tanpa app.xml dan dengan penanda penyimpanan Google Docs (seperti ekspor Google Docs), tanpa nilai desimal."""
    return drop_app_xml(make_docx(tmp_path / "w.docx"), tmp_path / name, add={"customXML/item1.xml": GDOCS_MARKER})


def decimal_docx(tmp_path, where="ind", name="desimal.docx"):
    """Dokumen yang sah menurut Word tetapi memuat ukuran desimal pada satu tempat (where)."""
    d = Document(make_docx(tmp_path / "w.docx"))
    if where == "pgMar":
        d.sections[0]._sectPr.find(qn("w:pgMar")).set(qn("w:left"), "1700.7874015748032")
    elif where == "ind":                                                    # persis nilai pada galat yang dilaporkan
        pPr = d.paragraphs[0]._p.get_or_add_pPr()
        ind = OxmlElement("w:ind")
        ind.set(qn("w:left"), "708.6614173228347")
        pPr.append(ind)
    elif where == "tblW":
        d.element.body.find(qn("w:tbl")).find(qn("w:tblPr")).find(qn("w:tblW")).set(qn("w:w"), "9000.75")
    out = tmp_path / name
    d.save(out)
    if where == "styles":
        return rewrite(out, tmp_path / ("s_" + name),
                       edits={"word/styles.xml": lambda b: b.replace(b"</w:styles>",
                              b'<w:style w:type="paragraph" w:styleId="X"><w:pPr><w:spacing w:before="12.5"/></w:pPr></w:style></w:styles>')})
    return out


def client(tmp_path):
    a = create_app(jobs_dir=tmp_path / "jobs", use_crossref=False, sweep_interval=0)
    return a.test_client(), tmp_path / "jobs"


def post(c, path, name="naskah.docx"):
    with open(path, "rb") as fh:
        return c.post("/periksa", data={"berkas": (fh, name), "jenis": "orisinal"}, content_type="multipart/form-data")


# ---------------------------------------------------------------- pengenalan aplikasi pembuat
def test_detect_producer_categories(tmp_path):
    assert m1_upload.detect_producer(make_docx(tmp_path / "a.docx"))["kategori"] == "word"             # python-docx: Microsoft Macintosh Word
    assert m1_upload.detect_producer(with_app(tmp_path, "Microsoft Office Word", "b.docx"))["word"] is True
    lo = m1_upload.detect_producer(with_app(tmp_path, "LibreOffice/24.2.7.2$Linux_X86_64", "c.docx"))
    assert lo["kategori"] == "libreoffice" and lo["word"] is False and lo["nama"] == "LibreOffice"
    g = m1_upload.detect_producer(gdocs_like(tmp_path))
    assert g["kategori"] == "google_docs" and g["aplikasi"] is None and g["nama"] == "Google Docs"
    wps = m1_upload.detect_producer(with_app(tmp_path, "WPS Office", "d.docx"))
    assert wps["kategori"] == "lainnya" and wps["nama"] == "WPS Office"
    none = m1_upload.detect_producer(drop_app_xml(make_docx(tmp_path / "e0.docx"), tmp_path / "e.docx"))
    assert none["kategori"] == "tidak_diketahui" and none["word"] is False


def test_detect_producer_never_raises_on_bad_files(tmp_path):
    bad = tmp_path / "x.docx"
    bad.write_bytes(b"bukan zip")
    assert m1_upload.detect_producer(bad)["kategori"] == "tidak_diketahui"
    assert m1_upload.detect_producer(tmp_path / "tidak-ada.docx")["kategori"] == "tidak_diketahui"


# ---------------------------------------------------------------- deteksi ukuran desimal
@pytest.mark.parametrize("where, part, tag", [("pgMar", "word/document.xml", "pgMar"), ("ind", "word/document.xml", "ind"),
                                              ("tblW", "word/document.xml", "tblW"), ("styles", "word/styles.xml", "spacing")])
def test_decimal_measure_is_found_in_each_part(tmp_path, where, part, tag):
    with zipfile.ZipFile(decimal_docx(tmp_path, where)) as z:
        hit = m1_upload.find_decimal_measure(z)
    assert hit and hit["bagian"] == part and hit["elemen"] == tag and "." in hit["nilai"]


def test_decimal_in_header_is_found(tmp_path):
    src = make_docx(tmp_path / "w.docx")
    with zipfile.ZipFile(src) as z:
        hdr = next(n for n in z.namelist() if re.match(r"word/header\d*\.xml$", n))
    out = rewrite(src, tmp_path / "h.docx", edits={hdr: lambda b: re.sub(rb"(<w:tblW )", rb'\1w:w="8788.0" ', b, count=1) if b"<w:tblW " in b
                                                   else b.replace(b"<w:p>", b'<w:p><w:pPr><w:ind w:left="100.5"/></w:pPr>', 1)})
    with zipfile.ZipFile(out) as z:
        assert m1_upload.find_decimal_measure(z)["bagian"] == hdr


def test_clean_word_document_has_no_decimal_measure(tmp_path):
    with zipfile.ZipFile(make_docx(tmp_path / "w.docx")) as z:
        assert m1_upload.find_decimal_measure(z) is None


def test_no_false_rejection_of_legitimate_decimals(tmp_path):
    """Nilai desimal yang SAH pada atribut non-ukuran (tag/nama/w:val umum) tidak boleh membuat dokumen Word ditolak."""
    src = make_docx(tmp_path / "w.docx")
    extra = (b'<w:p><w:sdt><w:sdtPr><w:tag w:val="1.5"/><w:alias w:val="2.0"/></w:sdtPr><w:sdtContent/></w:sdt>'
             b'<w:pPr><w:pStyle w:val="3.1"/></w:pPr></w:p>')
    out = rewrite(src, tmp_path / "ok.docx", edits={"word/document.xml": lambda b: b.replace(b"<w:body>", b"<w:body>" + extra, 1)})
    with zipfile.ZipFile(out) as z:
        assert m1_upload.find_decimal_measure(z) is None
    assert m1_upload.validate_file(out, "ok.docx") is True


def test_decimal_font_size_is_flagged_but_integer_is_not(tmp_path):
    src = make_docx(tmp_path / "w.docx")
    out = rewrite(src, tmp_path / "sz.docx", edits={"word/styles.xml": lambda b: b.replace(
        b"</w:styles>", b'<w:style w:type="character" w:styleId="Y"><w:rPr><w:sz w:val="21.5"/></w:rPr></w:style></w:styles>')})
    with zipfile.ZipFile(out) as z:
        assert m1_upload.find_decimal_measure(z)["elemen"] == "sz"


def test_validate_file_rejects_decimal_with_clear_message(tmp_path):
    with pytest.raises(m1_upload.ValidationError) as e:
        m1_upload.validate_file(decimal_docx(tmp_path, "ind"), "x.docx")
    msg = str(e.value)
    assert "bilangan desimal" in msg and "Microsoft Word" in msg and "simpan ulang" in msg


# ---------------------------------------------------------------- melalui aplikasi web
def test_web_rejects_decimal_document_with_clear_message_and_stores_nothing(tmp_path, monkeypatch):
    called = []
    monkeypatch.setattr(app_mod, "run_pipeline", lambda *a, **k: called.append(1))
    c, jobs = client(tmp_path)
    r = post(c, decimal_docx(tmp_path, "ind"), "naskah-gdocs.docx")
    html = r.get_data(as_text=True)
    assert r.status_code == 400 and "bilangan desimal" in html and "simpan ulang" in html
    assert called == []                                                     # ditolak di M1: pipeline tidak dijalankan
    assert list(jobs.iterdir()) == []                                       # dan naskahnya tidak tersimpan


def test_web_accepts_non_word_document_without_decimals_and_warns(tmp_path):
    c, jobs = client(tmp_path)
    r = post(c, with_app(tmp_path, "LibreOffice/24.2.7.2", "lo.docx"))
    assert r.status_code in (302, 303)
    html = c.get(r.headers["Location"]).get_data(as_text=True)
    assert "tampaknya dibuat dengan LibreOffice" in html and "bukan Microsoft Word" in html
    job = re.search(r"/laporan/([0-9a-f]{32})", r.headers["Location"]).group(1)
    meta = json.loads((jobs / job / "laporan.json").read_text(encoding="utf-8"))["meta"]
    assert meta["pembuat"]["kategori"] == "libreoffice" and meta["pembuat"]["word"] is False


def test_web_warns_for_google_docs_marker(tmp_path):
    c, _ = client(tmp_path)
    r = post(c, gdocs_like(tmp_path))
    assert r.status_code in (302, 303)
    assert "tampaknya dibuat dengan Google Docs" in c.get(r.headers["Location"]).get_data(as_text=True)


def test_web_word_document_has_no_warning(tmp_path):
    c, jobs = client(tmp_path)
    r = post(c, make_docx(tmp_path / "w.docx"))
    assert r.status_code in (302, 303)
    html = c.get(r.headers["Location"]).get_data(as_text=True)
    assert "bukan Microsoft Word" not in html
    job = re.search(r"/laporan/([0-9a-f]{32})", r.headers["Location"]).group(1)
    assert json.loads((jobs / job / "laporan.json").read_text(encoding="utf-8"))["meta"]["pembuat"]["word"] is True


def test_index_states_the_word_boundary(tmp_path):
    c, _ = client(tmp_path)
    assert "dari Microsoft Word" in c.get("/").get_data(as_text=True)


# ---------------------------------------------------------------- pembuat berkas di laporan
def test_producer_is_recorded_in_report(tmp_path):
    from checker.pipeline import run_pipeline
    src = with_app(tmp_path, "LibreOffice/24.2", "lo.docx")
    rep = run_pipeline(src, tmp_path / "out", {}, use_crossref=False,
                       meta={"nama_berkas": "lo.docx", "pembuat": m1_upload.detect_producer(src)})
    assert rep["meta"]["pembuat"]["kategori"] == "libreoffice"
