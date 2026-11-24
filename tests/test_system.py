import json
import zipfile
from pathlib import Path

import pytest
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Cm, Pt

from checker import m1_upload, m2_parser, m3_rules, m4_citations, m5_crossref
from checker.m2_parser import extract_citations, iter_paragraphs
from checker.pipeline import run_pipeline
from conftest import make_docx, crossref_ok


def statuses(report):
    return {k["id"]: k["status"] for k in report["kriteria"]}


def finding_msgs(report, cid):
    return [f["message"] for k in report["kriteria"] if k["id"] == cid for f in k["findings"]]


# ---------------------------------------------------------------- M1
class FakeUpload:
    def __init__(self, data, name):
        self.data, self.filename = data, name
    def save(self, dest):
        open(dest, "wb").write(self.data)


def test_m1_rejects_wrong_extension(tmp_path):
    p = tmp_path / "x.pdf"; p.write_bytes(b"%PDF-1.4")
    with pytest.raises(m1_upload.ValidationError, match="\\.docx"):
        m1_upload.validate_file(p)


def test_m1_rejects_fake_docx_and_empty(tmp_path):
    p = tmp_path / "x.docx"; p.write_bytes(b"ini bukan zip")
    with pytest.raises(m1_upload.ValidationError, match="OOXML"):
        m1_upload.validate_file(p)
    p.write_bytes(b"")
    with pytest.raises(m1_upload.ValidationError, match="kosong"):
        m1_upload.validate_file(p)


def test_m1_rejects_zip_without_word_parts_and_too_large(tmp_path, good_docx):
    z = tmp_path / "z.docx"
    with zipfile.ZipFile(z, "w") as zf:
        zf.writestr("a.txt", "x")
    with pytest.raises(m1_upload.ValidationError, match="utuh"):
        m1_upload.validate_file(z)
    with pytest.raises(m1_upload.ValidationError, match="melebihi"):
        m1_upload.validate_file(good_docx, max_bytes=100)
    assert m1_upload.validate_file(good_docx)


def test_m1_save_upload_cleans_up_invalid(tmp_path):
    with pytest.raises(m1_upload.ValidationError):
        m1_upload.save_upload(FakeUpload(b"bukan docx", "a.docx"), tmp_path)
    assert list(tmp_path.iterdir()) == []


# ---------------------------------------------------------------- M2
@pytest.mark.parametrize("text,expected", [
    ("lihat [1], [3] dan [2]", [1, 3, 2]),
    ("rentang [5]\u2013[7]", [5, 6, 7]),
    ("rentang [2]-[4], [9]", [2, 3, 4, 9]),
    ("gabungan [1,2]", [1, 2]),
    ("bukan sitasi [x] atau (3)", []),
    # --- notasi non-sitasi yang dulu salah terbaca ---
    ("normalisasi ke rentang [0,1]", []),
    ("nilai dalam [0, 1] saja", []),
    ("elemen x[1] dan A[2]", []),
    ("pada tahun [2024]", []),
    ("dalam interval [1,5]", []),
    ("jika x \u2208 [1,2]", []),
    # --- sitasi sah tetap terbaca ---
    ("(lihat [1])", [1]),
    ("perbedaan antara [3] dan [5]", [3, 5]),
    ("rentang nilai berbeda, seperti dilaporkan pada [4]", [4]),
])
def test_extract_citations(text, expected):
    assert extract_citations(text) == expected


def test_parser_finds_front_matter_and_refs(good_docx):
    pd = m2_parser.parse(good_docx)
    assert pd.title.text.startswith("Klasifikasi")
    assert len(pd.abstract_text.split()) == 130
    assert pd.keywords == ["akurasi", "klasifikasi", "pembelajaran mesin"]
    assert len(pd.references) == 20 and pd.references[0]["doi"] == "10.1234/jurnal.1"
    assert [c["kind"] for c in pd.captions] == ["gambar", "tabel"]
    assert [h.section for h in pd.headings][-1] == "referensi"


def test_parser_reads_template_table_layout():
    """Template JISINTA: abstrak & kata kunci berada di kotak tabel."""
    from pathlib import Path
    tpl = Path("/mnt/user-data/uploads/Template-Final.docx")
    if not tpl.exists():
        pytest.skip("template tidak tersedia")
    pd = m2_parser.parse(tpl)
    assert len(pd.keywords) == 5 and len(pd.abstract_text.split()) > 90


# ---------------------------------------------------------------- M3 + M4 (dokumen baik)
def test_good_document_has_no_failures(good_docx, tmp_path, mock_crossref):
    rep = run_pipeline(good_docx, tmp_path / "job", meta={"nama_berkas": "bagus.docx"})
    st = statuses(rep)
    assert len(st) == 18 and "fail" not in st.values(), json.dumps(
        [k for k in rep["kriteria"] if k["status"] == "fail"], ensure_ascii=False, indent=1)
    assert rep["ringkasan"]["crossref"].get("terverifikasi") == 20


# ---------------------------------------------------------------- M3 + M4 (dokumen salah)
def test_bad_document_triggers_expected_criteria(tmp_path, mock_crossref):
    bad = make_docx(tmp_path / "salah.docx", margin_top=3.5,
                    title="Analisis Komparatif Penerapan Berbagai Metode Pembelajaran Mesin untuk Klasifikasi Dokumen "
                          "Berbahasa Indonesia pada Studi Kasus di Beberapa Perguruan Tinggi Negeri",
                    abstract_words=40, keywords=("zeta", "alfa", "beta", "gamma", "delta", "kappa"),
                    n_refs=12, bad_cite=True, with_conclusion=False, caption=False)
    rep = run_pipeline(bad, tmp_path / "job")
    st = statuses(rep)
    for cid in (2, 6, 8, 9, 10, 18):
        assert st[cid] == "fail", (cid, st)
    assert any("3.5" in m for m in finding_msgs(rep, 2))
    assert any("kata" in m and "maksimal 15" in m for m in finding_msgs(rep, 6))
    assert any("analisis" in m for m in finding_msgs(rep, 6))
    assert any("terurut abjad" in m for m in finding_msgs(rep, 9))
    assert any("Kesimpulan" in m for m in finding_msgs(rep, 10))
    msgs14 = finding_msgs(rep, 18)
    assert any("minimal 20" in m for m in msgs14) and any("[99]" in m for m in msgs14)


def test_missing_caption_and_uncited_figure(tmp_path):
    d = make_docx(tmp_path / "c.docx", caption=False)
    pd = m2_parser.parse(d)
    res = m3_rules.RuleEngine().run(pd, {})
    k15 = next(r for r in res if r.id == 15)
    assert k15.status in ("warn", "fail") or not pd.figures   # tabel tanpa caption
    assert any("caption" in f.message.lower() for f in k15.findings)


def test_page_limit_depends_on_article_type(tmp_path):
    pd = m2_parser.parse(make_docx(tmp_path / "p.docx"))
    pd.pages_meta = 14
    eng = m3_rules.RuleEngine()
    k = lambda o: next(r for r in eng.run(pd, o) if r.id == 4)
    assert k({"jenis": "orisinal"}).status == "warn"
    assert k({"jenis": "review"}).status == "pass"


def test_rules_json_drives_thresholds(tmp_path, good_docx):
    db = m3_rules.load_rules()
    next(c for c in db["kriteria"] if c["id"] == 6)["params"]["max_words"] = 3
    pd = m2_parser.parse(good_docx)
    res = m3_rules.RuleEngine(db).run(pd, {})
    assert next(r for r in res if r.id == 6).status == "fail"


# ---------------------------------------------------------------- M5
@pytest.fixture
def ref():
    return {"num": 1, "idx": 0, "doi": "10.1/x", "title": "Deep Learning for Text", "text": "A. B, ...",
            "year": 2024, "journal_like": True}


def run_with(monkeypatch, handler, ref):
    monkeypatch.setattr(m5_crossref, "http_get", handler)
    return m5_crossref.verify_reference(ref)


def test_crossref_verified(monkeypatch, ref):
    r = run_with(monkeypatch, lambda u, p=None: (200, {"message": {"title": ["Deep learning for text"],
                 "DOI": "10.1/x", "issued": {"date-parts": [[2024]]}}}), ref)
    assert r["status"] == "terverifikasi" and r["similarity"] >= 0.85


def test_crossref_title_mismatch_year_mismatch_and_404(monkeypatch, ref):
    r = run_with(monkeypatch, lambda u, p=None: (200, {"message": {"title": ["Quantum gravity survey"],
                 "issued": {"date-parts": [[2024]]}}}), ref)
    assert r["status"] == "tidak_cocok"
    r = run_with(monkeypatch, lambda u, p=None: (200, {"message": {"title": ["Deep Learning for Text"],
                 "issued": {"date-parts": [[2021]]}}}), ref)
    assert r["status"] == "tahun_berbeda"
    assert run_with(monkeypatch, lambda u, p=None: (404, None), ref)["status"] == "doi_tidak_terdaftar"


def test_crossref_search_without_doi_and_network_error(monkeypatch, ref):
    ref = {**ref, "doi": None}
    r = run_with(monkeypatch, lambda u, p=None: (200, {"message": {"items": [
        {"title": ["Deep Learning for Text"], "DOI": "10.9/found"}]}}), ref)
    assert r["status"] == "ditemukan_tanpa_doi" and r["crossref_doi"] == "10.9/found"
    def boom(u, p=None): raise ConnectionError("offline")
    r = run_with(monkeypatch, boom, ref)
    assert r["status"] == "error" and "ConnectionError" in r["message"]


def test_pipeline_survives_crossref_outage(good_docx, tmp_path, monkeypatch):
    def boom(u, p=None): raise TimeoutError()
    monkeypatch.setattr(m5_crossref, "http_get", boom)
    rep = run_pipeline(good_docx, tmp_path / "j")
    assert rep["ringkasan"]["crossref"] == {"error": 20}
    assert any("Verifikasi Crossref gagal" in m for m in finding_msgs(rep, 18))


# ---------------------------------------------------------------- M6
def test_annotation_adds_comments_without_changing_text(tmp_path, mock_crossref):
    bad = make_docx(tmp_path / "salah.docx", margin_top=3.5, abstract_words=40)
    rep = run_pipeline(bad, tmp_path / "job")
    a, b = Document(bad), Document(tmp_path / "job" / "berkomentar.docx")
    assert [p.text for _, p, _ in iter_paragraphs(a)] == [p.text for _, p, _ in iter_paragraphs(b)]
    comments = list(b.comments)
    assert len(comments) == rep["meta"]["komentar"] > 0
    assert any("[K08]" in c.text for c in comments)
    assert comments[0].author == "Pemeriksa Naskah"


# ---------------------------------------------------------------- Web
@pytest.fixture
def client(tmp_path):
    from app import create_app
    app = create_app(jobs_dir=tmp_path / "jobs", use_crossref=False)
    app.config["TESTING"] = True
    return app.test_client()


def post(client, path, name="n.docx", jenis="orisinal"):
    import io
    return client.post("/periksa", data={"berkas": (io.BytesIO(open(path, "rb").read()), name), "jenis": jenis},
                       content_type="multipart/form-data")


def test_web_flow(client, good_docx):
    assert client.get("/").status_code == 200
    r = post(client, good_docx)
    assert r.status_code == 302 and r.location.startswith("/laporan/")
    page = client.get(r.location)
    assert page.status_code == 200 and "Laporan pemeriksaan" in page.get_data(as_text=True)
    job = r.location.rsplit("/", 1)[1]
    for kind, ctype in (("docx", "officedocument"), ("json", "json"), ("asli", "officedocument")):
        d = client.get(f"/unduh/{job}/{kind}")
        assert d.status_code == 200 and ctype in d.content_type


def test_web_rejects_bad_input_and_bad_ids(client, tmp_path):
    bad = tmp_path / "x.docx"; bad.write_bytes(b"bukan docx")
    r = post(client, bad)
    assert r.status_code == 400 and "OOXML" in r.get_data(as_text=True)
    assert client.post("/periksa", data={}).status_code == 400
    assert client.get("/laporan/../../etc/passwd").status_code == 404
    assert client.get("/laporan/" + "0" * 32).status_code == 404
    assert client.get("/unduh/" + "0" * 32 + "/rahasia").status_code == 404


# ---------------------------------------------------------------- K17 penulis & afiliasi
def run_k(tmp_path, cid, **kw):
    pd = m2_parser.parse(make_docx(tmp_path / "k.docx", **kw))
    res = next(r for r in m3_rules.RuleEngine().run(pd, {}) if r.id == cid)
    return pd, res, [f.message for f in res.findings]


def test_parser_reads_authors_affiliations_corresponding(good_docx):
    pd = m2_parser.parse(good_docx)
    assert [(n["name"], n["markers"]) for n in pd.author_names] == [("Siti Aminah", [1]), ("Budi Santoso", [2])]
    assert pd.affiliation_markers == [1, 2] and pd.affiliation_count == 2
    assert pd.corresponding["email"] == "siti.aminah@contoh.ac.id"
    assert len(pd.affiliations) == 1          # paragraf penulis korespondensi bukan afiliasi


def test_k17_passes_for_correct_front_matter(tmp_path):
    _, res, msgs = run_k(tmp_path, 7)
    assert res.status == "pass", msgs


def test_k17_detects_wrong_format_marker_mismatch_missing_email_and_placeholder(tmp_path):
    _, res, msgs = run_k(tmp_path, 7, authors="bad")
    assert res.status == "fail"
    joined = " | ".join(msgs)
    assert "rata tengah" in joined                      # penulis rata kiri
    assert "12.0 pt" in joined                          # ukuran penulis salah
    assert "10.0 pt; seharusnya 9.0" in joined          # ukuran afiliasi salah
    assert "Penanda afiliasi 3" in joined               # penanda 3 tanpa afiliasi 3
    assert "Dewi Lestari" in joined                     # penulis tanpa penanda
    assert "email" in joined                            # korespondensi tanpa email
    assert "teks bawaan template" in joined             # 'Nama Penulis Korespondensi'


def test_k17_missing_affiliation(tmp_path):
    _, res, msgs = run_k(tmp_path, 7, authors="no_aff")
    assert res.status == "fail" and any("Afiliasi penulis tidak ditemukan" in m for m in msgs)


# ---------------------------------------------------------------- K18 header & footer
def test_k18_passes_with_first_odd_even(tmp_path):
    pd, res, msgs = run_k(tmp_path, 5)
    assert res.status == "pass", msgs
    assert pd.hf_flags == {"different_first": True, "odd_even": True}
    assert all(r["page_field"] for k, r in pd.headers_footers.items() if k.startswith("footer"))


def test_k18_fails_without_headers(tmp_path):
    _, res, msgs = run_k(tmp_path, 5, hf="none")
    assert res.status == "fail"
    assert any("Different First Page" in m for m in msgs) and any("Odd & Even" in m for m in msgs)
    assert any("Header halaman genap kosong" in m for m in msgs)
    assert any("Footer halaman pertama kosong" in m for m in msgs)


def test_k18_flags_off_even_though_parts_exist(tmp_path):
    _, res, msgs = run_k(tmp_path, 5, hf="no_flags")
    assert res.status == "fail" and sum("tidak aktif" in m for m in msgs) == 2


def test_k18_even_footer_must_hold_title_and_authors(tmp_path):
    _, res, msgs = run_k(tmp_path, 5, hf="no_even_text")
    assert any("hanya berisi nomor halaman" in m for m in msgs)


def test_k18_template_placeholders_in_even_footer():
    from pathlib import Path
    tpl = Path("/mnt/user-data/uploads/Template-Final.docx")
    if not tpl.exists():
        pytest.skip("template tidak tersedia")
    res = next(r for r in m3_rules.RuleEngine().run(m2_parser.parse(tpl), {}) if r.id == 5)
    assert any("bawaan template" in f.message for f in res.findings)


# ---------------------------------------------------------------- K09 aturan heading per level
def _fmt_heading(h, level):
    pf = h.paragraph_format
    pf.space_before, pf.space_after = Pt(6), Pt({1: 6, 2: 3, 3: 2}[level])
    pf.left_indent, pf.first_line_indent = Cm(0.635), Cm(-0.635)
    return h


def _heading_doc(tmp_path, h2_bold=True, h3_bold=False, h1_caps=True):
    d = Document()
    _fmt_heading(d.add_heading("1. PENDAHULUAN" if h1_caps else "1. Pendahuluan", 1), 1)
    h2 = _fmt_heading(d.add_heading("1.1. Latar Belakang", 2), 2); h2.runs[0].font.bold = h2_bold
    h3 = _fmt_heading(d.add_heading("1.1.1. Rincian", 3), 3); h3.runs[0].font.bold = h3_bold
    d.add_paragraph("Isi paragraf.")
    path = tmp_path / "h.docx"; d.save(path)
    return m2_parser.parse(path)


def _k09(pd, db=None):
    return next(r for r in m3_rules.RuleEngine(db).run(pd, {}) if r.id == 11)


def test_heading_rules_per_level_pass(tmp_path):
    res = _k09(_heading_doc(tmp_path))
    msgs = [f.message for f in res.findings if f.severity != "info"]
    assert not msgs, msgs


def test_heading2_not_bold_is_flagged_and_level3_unchecked(tmp_path):
    res = _k09(_heading_doc(tmp_path, h2_bold=False))
    msgs = [f.message for f in res.findings]
    assert any("Heading 2" in m and "belum tebal" in m for m in msgs)
    assert not any("Heading 3" in m for m in msgs)          # H3 tidak bold = sesuai aturan


def test_heading3_bold_is_forbidden(tmp_path):
    res = _k09(_heading_doc(tmp_path, h3_bold=True))
    assert any("Heading 3" in f.message and "tidak boleh tebal" in f.message for f in res.findings)
    assert res.status == "warn"


def test_forbid_false_and_true_semantics(tmp_path):
    pd = _heading_doc(tmp_path, h3_bold=True)
    db = m3_rules.load_rules()
    lv3 = next(c for c in db["kriteria"] if c["id"] == 11)["params"]["levels"]["3"]
    lv3["bold"] = False                                   # tidak diperiksa
    assert not any("Heading 3" in f.message for f in _k09(pd, db).findings)


def test_heading3_bold_required_only_when_rules_say_so(tmp_path):
    pd = _heading_doc(tmp_path, h3_bold=False)
    db = m3_rules.load_rules()
    next(c for c in db["kriteria"] if c["id"] == 11)["params"]["levels"]["3"]["bold"] = True
    assert any("Heading 3" in f.message and "belum tebal" in f.message for f in _k09(pd, db).findings)


def test_heading1_still_requires_caps(tmp_path):
    res = _k09(_heading_doc(tmp_path, h1_caps=False))
    assert any("Heading 1" in f.message and "kapital" in f.message for f in res.findings)


def test_range_notation_not_counted_as_citation_in_parser(tmp_path):
    d = Document()
    d.add_heading("1. PENDAHULUAN", 1)
    d.add_paragraph("Skor dinormalisasi ke rentang [0,1] seperti pada [3], dengan x[1] sebagai indeks.")
    d.save(tmp_path / "r.docx")
    pd = m2_parser.parse(tmp_path / "r.docx")
    assert [c["nums"] for c in pd.citations] == [[3]]


# ---------------------------------------------------------------- spasi baris (K04)
def _k04_for(tmp_path, spacing):
    d = Document()
    d.styles["Normal"].font.name, d.styles["Normal"].font.size = "Cambria", Pt(10)
    d.add_heading("1. PENDAHULUAN", 1)
    p = d.add_paragraph("Paragraf isi naskah dengan jumlah kata yang cukup untuk diperiksa oleh sistem pemeriksa.")
    if spacing is not None:
        p.paragraph_format.line_spacing = spacing
    d.save(tmp_path / "s.docx")
    pd = m2_parser.parse(tmp_path / "s.docx")
    return pd, next(r for r in m3_rules.RuleEngine().run(pd, {}) if r.id == 12)


def test_line_spacing_1_15_is_accepted(tmp_path):
    _, res = _k04_for(tmp_path, 1.15)
    assert not [f for f in res.findings if "spasi" in f.message.lower()]


def test_single_and_double_spacing_are_flagged(tmp_path):
    for sp in (1.0, 2.0):
        _, res = _k04_for(tmp_path, sp)
        assert any(f"Spasi baris {sp}; seharusnya 1.15" in f.message for f in res.findings), sp


def test_exact_point_spacing_is_flagged_not_treated_as_single(tmp_path):
    _, res = _k04_for(tmp_path, Pt(14))
    assert any("Spasi baris tetap 14.0 pt" in f.message for f in res.findings)


def test_doc_default_spacing_is_used_when_style_has_none(tmp_path):
    from docx.oxml import OxmlElement
    d = Document()
    dd = d.styles.element.find(qn("w:docDefaults"))
    ppr_def = dd.find(qn("w:pPrDefault"))
    if ppr_def is None:
        ppr_def = OxmlElement("w:pPrDefault"); dd.append(ppr_def)
    ppr = ppr_def.find(qn("w:pPr"))
    if ppr is None:
        ppr = OxmlElement("w:pPr"); ppr_def.append(ppr)
    for old in ppr.findall(qn("w:spacing")):
        ppr.remove(old)
    sp = OxmlElement("w:spacing"); sp.set(qn("w:line"), "259"); sp.set(qn("w:lineRule"), "auto"); ppr.append(sp)
    d.add_paragraph("Teks tanpa spasi eksplisit pada style maupun paragraf.")
    d.save(tmp_path / "dd.docx")
    pd = m2_parser.parse(tmp_path / "dd.docx")
    assert pd.paras[0].line_spacing == 1.08


# ---------------------------------------------------------------- M5: HTTP stdlib, tanpa 'requests'
@pytest.fixture
def fake_crossref(monkeypatch):
    import http.server, json as _json, threading, urllib.parse
    class H(http.server.BaseHTTPRequestHandler):
        def log_message(self, *a): pass
        def do_GET(self):
            u = urllib.parse.urlparse(self.path)
            if u.path == "/works/10.1/ada":
                body, code = {"message": {"title": ["Deep Learning for Text"], "DOI": "10.1/ada",
                                          "issued": {"date-parts": [[2024]]}}}, 200
            elif u.path == "/works" and "query.bibliographic" in u.query:
                body, code = {"message": {"items": [{"title": ["Deep Learning for Text"], "DOI": "10.9/cari"}]}}, 200
            else:
                body, code = {"status": "not found"}, 404
            raw = _json.dumps(body).encode()
            self.send_response(code); self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(raw))); self.end_headers(); self.wfile.write(raw)
    srv = http.server.HTTPServer(("127.0.0.1", 0), H)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    monkeypatch.setattr(m5_crossref, "API", f"http://127.0.0.1:{srv.server_port}/works")
    yield
    srv.shutdown()


def test_http_get_works_without_requests_package(fake_crossref, monkeypatch):
    import sys
    monkeypatch.setitem(sys.modules, "requests", None)      # import requests -> ModuleNotFoundError
    ref = {"num": 1, "idx": 0, "doi": "10.1/ada", "title": "Deep Learning for Text", "text": "...",
           "year": 2024, "journal_like": True}
    assert m5_crossref.verify_reference(ref)["status"] == "terverifikasi"
    assert m5_crossref.verify_reference({**ref, "doi": "10.1/tidakada"})["status"] == "doi_tidak_terdaftar"
    assert m5_crossref.verify_reference({**ref, "doi": None})["status"] == "ditemukan_tanpa_doi"


def test_unreachable_crossref_reports_real_cause(monkeypatch):
    monkeypatch.setattr(m5_crossref, "API", "http://127.0.0.1:9/works")      # port tertutup
    r = m5_crossref.verify_reference({"num": 1, "idx": 0, "doi": "10.1/x", "title": "T", "text": "T",
                                      "year": None, "journal_like": True})
    assert r["status"] == "error" and "Tidak dapat menghubungi Crossref" in r["message"]


# ---------------------------------------------------------------- M5: judul beda bahasa
ID_REF_TEXT = ("A. Santoso dan B. Lestari, \u201cPenerapan metode klasifikasi untuk analisis sentimen pada ulasan produk,\u201d "
               "Jurnal Informatika, vol. 5, no. 2, pp. 1\u201310, 2024, doi: 10.1/bahasa.")


def _id_ref(**kw):
    base = {"num": 3, "idx": 7, "doi": "10.1/bahasa", "text": ID_REF_TEXT, "year": 2024, "journal_like": True,
            "title": "Penerapan metode klasifikasi untuk analisis sentimen pada ulasan produk"}
    base.update(kw)
    return base


def _cr_item(**kw):
    item = {"title": ["Application of classification methods for sentiment analysis on product reviews"],
            "DOI": "10.1/bahasa", "issued": {"date-parts": [[2024]]},
            "author": [{"given": "Budi", "family": "Santoso"}, {"given": "Dewi", "family": "Lestari"}]}
    item.update(kw)
    return item


def _serve(item):
    return lambda url, params=None: (200, {"message": item if "/works/" in url else {"items": [item]}})


def test_indonesian_title_vs_english_crossref_is_verified_when_author_and_year_match(monkeypatch):
    r = run_with(monkeypatch, _serve(_cr_item()), _id_ref())
    assert r["status"] == "terverifikasi" and r["corroborated"] and r["cross_language"]
    assert r["similarity"] < 0.85 and "berbeda bahasa" in r["message"]


def test_cross_language_but_wrong_authors_stays_mismatch(monkeypatch):
    other = _cr_item(author=[{"given": "John", "family": "Smith"}])
    r = run_with(monkeypatch, _serve(other), _id_ref())
    assert r["status"] == "tidak_cocok" and "berbeda bahasa" in r["message"]


def test_cross_language_but_year_contradicts_stays_mismatch(monkeypatch):
    r = run_with(monkeypatch, _serve(_cr_item(issued={"date-parts": [[2019]]})), _id_ref())
    assert r["status"] == "tidak_cocok"


def test_no_author_data_cannot_corroborate(monkeypatch):
    item = _cr_item(); item.pop("author")
    r = run_with(monkeypatch, _serve(item), _id_ref())
    assert r["status"] == "tidak_cocok" and r["author_match"] is None


def test_same_language_unrelated_title_with_same_authors_is_not_accepted(monkeypatch):
    item = _cr_item(title=["Quantum error correction in superconducting circuits"])
    ref = _id_ref(title="Deep learning for sentiment analysis of product reviews",
                  text=ID_REF_TEXT.replace("Penerapan metode klasifikasi untuk analisis sentimen pada ulasan produk",
                                           "Deep learning for sentiment analysis of product reviews"))
    r = run_with(monkeypatch, _serve(item), ref)
    assert r["status"] == "tidak_cocok" and not r["cross_language"]


def test_crossref_bilingual_title_list_matches_indonesian_title(monkeypatch):
    item = _cr_item(title=["Application of classification methods for sentiment analysis on product reviews",
                           "Penerapan metode klasifikasi untuk analisis sentimen pada ulasan produk"])
    r = run_with(monkeypatch, _serve(item), _id_ref())
    assert r["status"] == "terverifikasi" and not r["corroborated"] and r["similarity"] >= 0.85


def test_no_doi_indonesian_reference_searches_with_full_text_and_is_found(monkeypatch):
    seen = {}
    def handler(url, params=None):
        seen["q"] = (params or {}).get("query.bibliographic", "")
        return 200, {"message": {"items": [_cr_item()]}}
    r = run_with(monkeypatch, handler, _id_ref(doi=None))
    assert "Santoso" in seen["q"] and "Jurnal Informatika" in seen["q"]      # seluruh teks, bukan hanya judul
    assert r["status"] == "ditemukan_tanpa_doi" and r["corroborated"]


def test_corroborated_match_is_reported_as_info_not_warning(good_docx, monkeypatch):
    pd = m2_parser.parse(good_docx)
    pd.references = pd.references[:1]
    pd.references[0].update(_id_ref(num=1, idx=pd.references[0]["idx"]))
    from checker.models import CriterionResult
    res = CriterionResult(18, "K18", "x", "M4")
    monkeypatch.setattr(m5_crossref, "http_get", _serve(_cr_item()))
    rows = m5_crossref.run(pd, res)
    assert rows[0]["status"] == "terverifikasi"
    f = [x for x in res.findings if x.source == "M5"]
    assert len(f) == 1 and f[0].severity == "info" and not f[0].comment


def test_mailto_env_overrides_default_and_is_sent_in_user_agent(monkeypatch):
    captured = {}
    class FakeResp:
        status = 200
        def __enter__(self): return self
        def __exit__(self, *a): return False
        def read(self): return b"{}"
    def fake_urlopen(req, timeout=None):
        captured["ua"] = req.get_header("User-agent")
        return FakeResp()
    monkeypatch.setattr(m5_crossref.urllib.request, "urlopen", fake_urlopen)
    monkeypatch.setattr(m5_crossref, "DEFAULT_MAILTO", "default@contoh.ac.id")
    monkeypatch.delenv("CROSSREF_MAILTO", raising=False)
    m5_crossref.http_get("https://api.crossref.org/works/10.1/x")
    assert "mailto:default@contoh.ac.id" in captured["ua"]
    monkeypatch.setenv("CROSSREF_MAILTO", "env@contoh.ac.id")
    m5_crossref.http_get("https://api.crossref.org/works/10.1/x")
    assert "mailto:env@contoh.ac.id" in captured["ua"] and "default@" not in captured["ua"]


# ---------------------------------------------------------------- K19: istilah asing miring
from checker import langid

ID_PARA = ("Penelitian ini menggunakan {t} untuk mengklasifikasikan dokumen berbahasa Indonesia secara otomatis, "
           "kemudian hasilnya dievaluasi dengan beberapa metrik pada kumpulan data yang berbeda.")


def _lang_doc(tmp_path, paragraphs, name="l.docx"):
    """paragraphs: daftar list potongan (teks, miring?) per paragraf."""
    d = Document()
    d.add_heading("1. PENDAHULUAN", 1)
    for parts in paragraphs:
        p = d.add_paragraph()
        for text, it in parts:
            r = p.add_run(text); r.italic = it
    d.save(tmp_path / name)
    return m2_parser.parse(tmp_path / name)


def _k19(pd, opt=None, db=None):
    """Bagian istilah-asing-miring dari K13 (dahulu K19). Hanya temuan terkait italic yang dinilai, karena
    dokumen uji kecil tidak berindentasi sehingga K13 wajar memberi peringatan indentasi."""
    from types import SimpleNamespace
    r = next(x for x in m3_rules.RuleEngine(db).run(pd, opt or {}) if x.id == 14)
    fnd = [f for f in r.findings if "Istilah asing" in f.message]
    status = "skip" if "miring dilewati" in r.note else ("warn" if fnd else "pass")
    return SimpleNamespace(status=status, findings=fnd, note=r.note)


def _para(term_italic, term="machine learning"):
    before, after = ID_PARA.split("{t}")
    return [(before, False), (term, term_italic), (after, False)]


def test_parser_records_italic_ranges_incl_character_style(tmp_path):
    d = Document()
    p = d.add_paragraph()
    p.add_run("awal "); r = p.add_run("miring"); r.italic = True
    p.add_run(" tengah "); p.add_run("emfasis", style="Emphasis"); p.add_run(" akhir")
    d.save(tmp_path / "i.docx")
    info = m2_parser.parse(tmp_path / "i.docx").paras[0]
    assert info.raw_text == "awal miring tengah emfasis akhir"
    assert [info.raw_text[a:b] for a, b in info.italic_ranges] == ["miring", "emfasis"]


def test_langid_marks_terms_not_names_acronyms_or_indonesian_words():
    t = "Kami memakai machine learning dan dataset dari SVM, TensorFlow, Python, serta Google Scholar di sistem ini."
    got = [t[s:e] for s, e, _ in langid.foreign_spans(t)]
    assert got == ["machine learning", "dataset"]


def test_langid_title_case_phrase_and_product_name_split():
    t = "Natural Language Processing kini dipakai bersama framework Flask di banyak aplikasi."
    got = [t[s:e] for s, e, _ in langid.foreign_spans(t)]
    assert any("Language Processing" in g for g in got) and "framework" in got and "Flask" not in " ".join(got)


def test_langid_hyphenated_and_function_words_inside_terms():
    t = "Proses fine-tuning mengikuti pendekatan state of the art pada learning rate kecil."
    assert [t[s:e] for s, e, _ in langid.foreign_spans(t)] == ["fine-tuning", "state of the art", "learning rate"]


def test_langid_rare_indonesian_words_and_loanwords_are_not_flagged():
    t = "Acuan yang baur ini adar dan cena serta memakai data internet pada server."
    assert langid.foreign_spans(t) == []


def test_langid_allowlist_force_terms_and_patterns():
    t = "Pengguna melakukan download berkas sesuai Smith et al pada software tersebut."
    base = [t[s:e] for s, e, _ in langid.foreign_spans(t)]
    forced = [t[s:e] for s, e, _ in langid.foreign_spans(t, force_terms=["download", "software"],
                                                         extra_patterns=[r"\bet al\b"])]
    assert "download" not in base and {"download", "software", "et al"} <= set(forced)
    assert "software" not in [t[s:e] for s, e, _ in langid.foreign_spans(t, force_terms=["software"], allowlist=["software"])]


@pytest.mark.skipif(not langid.HAVE_LANGDETECT, reason="langdetect belum terpasang")
def test_langid_full_english_sentence_is_one_span():
    t = "Penulis menulis: the quick results show that the model clearly outperforms every other baseline method here."
    spans = langid.foreign_spans(t)
    assert any(k == "kalimat" for _, _, k in spans)


def test_k19_flags_non_italic_term_and_passes_when_italic(tmp_path):
    bad = _k19(_lang_doc(tmp_path, [_para(False)] * 3))
    assert bad.status == "warn" and "machine learning" in bad.findings[0].message
    ok = _k19(_lang_doc(tmp_path, [_para(True)] * 3, "ok.docx"))
    assert ok.status == "pass" and "3 sudah miring (100%)" in ok.note


def test_k19_partial_italic_counts_as_missing(tmp_path):
    parts = [(ID_PARA.split("{t}")[0], False), ("machine", True), (" learning", False), (ID_PARA.split("{t}")[1], False)]
    assert _k19(_lang_doc(tmp_path, [parts] * 2)).status == "warn"           # 7 dari 15 huruf miring (< 80%)


def test_k19_skips_english_manuscripts_and_respects_override(tmp_path):
    en = ("This study uses machine learning to classify documents written in English and the results are evaluated "
          "with several metrics on different datasets across all experiments in this paper.")
    pd = _lang_doc(tmp_path, [[(en, False)]] * 3)
    res = _k19(pd)
    assert res.status == "skip" and "berbahasa en" in res.note
    forced = _k19(_lang_doc(tmp_path, [_para(False)] * 3, "f.docx"), {"bahasa": "en"})
    assert forced.status == "skip"                                          # override pengguna


def test_k19_policy_lists_come_from_rules_json(tmp_path):
    para = [("Pengguna dapat melakukan download berkas hasil pemeriksaan melalui antarmuka web yang disediakan "
             "oleh sistem pemeriksa naskah ini setiap saat.", False)]
    pd = _lang_doc(tmp_path, [para] * 3)
    assert _k19(pd).status == "warn"                                        # 'download' ada di force_terms bawaan
    db = m3_rules.load_rules()
    prm = next(c for c in db["kriteria"] if c["id"] == 14)["params"]["foreign_italic"]
    prm["force_terms"].remove("download")
    assert _k19(pd, db=db).status == "pass"


def test_k19_does_not_touch_headings_tables_or_references(tmp_path):
    d = Document()
    d.add_heading("1. PENDAHULUAN", 1)
    d.add_paragraph(ID_PARA.format(t="metode biasa") + " Tambahan kalimat agar naskah terdeteksi Indonesia.")
    d.add_heading("2. METODE machine learning", 1)
    t = d.add_table(rows=1, cols=1); t.cell(0, 0).text = "overfitting dataset parser"
    d.save(tmp_path / "x.docx")
    assert _k19(m2_parser.parse(tmp_path / "x.docx")).findings == []


# ---------------------------------------------------------------- M5 memakai langdetect untuk bahasa judul
@pytest.mark.skipif(not langid.HAVE_LANGDETECT, reason="langdetect belum terpasang")
def test_title_language_handles_titles_the_heuristic_cannot():
    assert langid.title_language("Klasifikasi Sentimen Ulasan Aplikasi") == "id"
    assert langid.title_language("Deep learning for text") == "en"
    assert m5_crossref.guess_lang("Klasifikasi Sentimen Ulasan Aplikasi") == "id"


def test_title_language_falls_back_to_heuristic_without_langdetect(monkeypatch):
    monkeypatch.setattr(langid, "HAVE_LANGDETECT", False)
    assert langid.title_language("Penerapan metode klasifikasi untuk analisis sentimen") == "id"
    assert langid.title_language("Application of classification methods for analysis") == "en"
    assert langid.title_language("Sistem Informasi Akademik") is None


def test_k13_still_checks_indent_and_alignment_when_italic_part_is_skipped(tmp_path):
    en = ("This study uses machine learning to classify documents written in English and the results are evaluated "
          "with several metrics on different datasets across all experiments in this paper.")
    pd = _lang_doc(tmp_path, [[(en, False)]] * 3)                       # naskah Inggris, tanpa indentasi
    r = next(x for x in m3_rules.RuleEngine().run(pd, {}) if x.id == 14)
    assert "miring dilewati" in r.note                                  # bagian italic dilewati ...
    assert r.status == "warn" and any("menjorok" in f.message for f in r.findings)   # ... indentasi tetap diperiksa


def test_k13_combines_both_kinds_of_findings(tmp_path):
    pd = _lang_doc(tmp_path, [_para(False)] * 3)
    r = next(x for x in m3_rules.RuleEngine().run(pd, {}) if x.id == 14)
    msgs = " | ".join(f.message for f in r.findings)
    assert "menjorok" in msgs and "Istilah asing belum dicetak miring" in msgs


# ================================================================ KRITERIA FINAL (urutan A-D)
from variant_base import build_base                               # noqa: E402


def _mutated(tmp_path, fn, name="m.docx"):
    """Basis patuh (variant_base) + satu perubahan; kembalikan hasil parse."""
    path = build_base(tmp_path / name)
    d = Document(path)
    fn(d)
    d.save(path)
    return m2_parser.parse(path)


def _crit(pd, cid, db=None):
    return next(r for r in m3_rules.RuleEngine(db).run(pd, {}) if r.id == cid)


def test_rules_json_follows_final_order_and_groups():
    db = m3_rules.load_rules()
    ks = db["kriteria"]
    assert [c["id"] for c in ks] == list(range(1, 19))
    assert [c["kode"] for c in ks] == [f"K{i:02d}" for i in range(1, 19)]
    groups = [c["kelompok"] for c in ks]
    assert groups == (["A. Pengaturan dokumen"] * 4 + ["B. Identitas dokumen"] * 5
                      + ["C. Isi dokumen"] * 8 + ["D. Referensi"])
    assert [c["check"] for c in ks][:5] == ["paper_size", "margins", "columns", "page_count", "header_footer"]
    assert ks[-1]["modul"] == "M4" and all(c["modul"] == "M3" for c in ks[:-1])


def test_results_carry_group_label(tmp_path):
    res = m3_rules.RuleEngine().run(m2_parser.parse(build_base(tmp_path / "g.docx")), {})
    assert {r.id: r.group[:1] for r in res}[1] == "A" and {r.id: r.group[:1] for r in res}[17] == "C"
    assert res[0].to_dict()["group"] == "A. Pengaturan dokumen"


def test_parser_tracks_sections_and_columns(tmp_path):
    pd = m2_parser.parse(build_base(tmp_path / "s.docx"))
    assert [s["columns"] for s in pd.page["sections_info"]] == [1, 2]
    first_h = pd.headings[0].idx
    assert {p.section_index for p in pd.paras if p.text and p.idx < first_h} == {0}
    assert {p.section_index for p in pd.paras if p.text and p.idx >= first_h} == {1}


# ---------------------------------------------------------------- K01/K02 semua section
def test_k01_k02_check_every_section(tmp_path):
    def f(d):
        s2 = d.sections[1]
        s2.page_width, s2.page_height = Cm(21.59), Cm(27.94)
        s2.left_margin = Cm(1.0)
    pd = _mutated(tmp_path, f)
    k1, k2 = _crit(pd, 1), _crit(pd, 2)
    assert k1.status == "fail" and "Section 2" in k1.findings[0].message and "Letter" in k1.findings[0].message
    assert k2.status == "fail" and "Section 2" in k2.findings[0].message and "kiri" in k2.findings[0].message


# ---------------------------------------------------------------- K03 layout kolom
def test_k03_passes_identity_one_body_two(tmp_path):
    r = _crit(m2_parser.parse(build_base(tmp_path / "ok.docx")), 3)
    assert r.status == "pass" and r.note == "identitas 1 kolom, isi 2 kolom"


def test_k03_single_section_needs_continuous_break(tmp_path):
    pd = m2_parser.parse(make_docx(tmp_path / "single.docx"))
    d = Document(tmp_path / "single.docx")
    # hapus section kedua: gabungkan kembali ke satu section 1 kolom
    body = d.element.body
    for p in body.iter(qn("w:p")):
        sp = p.find(qn("w:pPr") + "/" + qn("w:sectPr"))
        if sp is not None:
            sp.getparent().remove(sp)
    d.sections[0]._sectPr.find(qn("w:cols")).set(qn("w:num"), "1")
    d.save(tmp_path / "single.docx")
    r = _crit(m2_parser.parse(tmp_path / "single.docx"), 3)
    assert r.status == "fail" and "Continuous" in r.findings[0].suggestion
    assert "satu section" in r.findings[0].message


def _set_cols(d, idx, n):
    d.sections[idx]._sectPr.find(qn("w:cols")).set(qn("w:num"), str(n))


def test_k03_identity_two_columns_and_wrong_body_columns(tmp_path):
    r = _crit(_mutated(tmp_path, lambda d: _set_cols(d, 0, 2)), 3)
    assert r.status == "fail" and "Bagian identitas memakai 2 kolom" in r.findings[0].message
    r = _crit(_mutated(tmp_path, lambda d: _set_cols(d, 1, 3), "m3.docx"), 3)
    assert r.status == "fail" and "Bagian isi" in r.findings[0].message and "3 kolom" in r.findings[0].message


def test_k03_body_column_count_is_a_parameter(tmp_path):
    pd = _mutated(tmp_path, lambda d: _set_cols(d, 1, 1))              # isi 1 kolom
    assert _crit(pd, 3).status == "fail"                               # bawaan: 2 kolom
    db = m3_rules.load_rules()
    next(c for c in db["kriteria"] if c["id"] == 3)["params"]["body_columns"] = 1
    assert _crit(pd, 3, db).status == "pass"


# ---------------------------------------------------------------- K06 format judul
def test_k06_title_format(tmp_path):
    assert _crit(m2_parser.parse(build_base(tmp_path / "t.docx")), 6).status == "pass"

    def bad(d):
        t = next(p for p in d.paragraphs if p.style.name == "Title")
        for r in t.runs:
            r.font.name, r.font.size, r.font.bold = "Arial", Pt(12), False
        t.alignment = WD_ALIGN_PARAGRAPH.LEFT
        t.paragraph_format.space_after = Pt(0)
    r = _crit(_mutated(tmp_path, bad, "t2.docx"), 6)
    msg = r.findings[0].message
    assert r.status == "warn" and all(w in msg for w in ("font 'Arial'", "ukuran 12.0 pt", "tidak tebal", "rata tengah", "0.0 pt"))


# ---------------------------------------------------------------- K07 spasi & font penulis/afiliasi
def test_k07_spacing_and_affiliation_font(tmp_path):
    def bad(d):
        paras = [p for p in d.paragraphs if p.text.startswith(("Siti Aminah\u00b9", "\u00b9Program"))]
        paras[0].paragraph_format.space_after = Pt(0)
        for r in paras[1].runs:
            r.font.name = "Arial"
    msgs = " | ".join(f.message for f in _crit(_mutated(tmp_path, bad), 7).findings)
    assert "Spasi sesudah nama penulis 0.0 pt" in msgs and "Font afiliasi 'Arial'" in msgs


# ---------------------------------------------------------------- K10 struktur IMRaD
def _add_h1_before_method(d, text):
    h = next(p for p in d.paragraphs if p.style.name == "Heading 1" and "METODE" in p.text)
    new = copy_el(h._p)
    h._p.addprevious(new)
    from docx.text.paragraph import Paragraph
    from docx.text.run import Run
    p = Paragraph(new, h._parent)
    p.runs[0].text = text
    for r in p.runs[1:]:
        r._r.getparent().remove(r._r)


def copy_el(el):
    import copy
    return copy.deepcopy(el)


def test_k10_rejects_literature_review_and_unknown_sections(tmp_path):
    r = _crit(_mutated(tmp_path, lambda d: _add_h1_before_method(d, "2. TINJAUAN PUSTAKA")), 10)
    assert r.status == "fail" and "tidak termasuk struktur IMRaD" in r.findings[0].message
    r = _crit(_mutated(tmp_path, lambda d: _add_h1_before_method(d, "2. LANDASAN TEORI"), "m2.docx"), 10)
    assert r.status == "fail" and "tidak termasuk struktur IMRaD" in r.findings[0].message
    r = _crit(_mutated(tmp_path, lambda d: _add_h1_before_method(d, "2. Analisis Kebutuhan Sistem"), "m3.docx"), 10)
    assert r.status == "fail" and "tidak dikenali" in r.findings[0].message


def test_k10_optional_acknowledgement_is_allowed(tmp_path):
    def f(d):
        ref = next(p for p in d.paragraphs if p.style.name == "Heading 1" and p.text == "REFERENSI")
        new = copy_el(ref._p)
        ref._p.addprevious(new)
        from docx.text.paragraph import Paragraph
        Paragraph(new, ref._parent).runs[0].text = "UCAPAN TERIMA KASIH"
    assert _crit(_mutated(tmp_path, f), 10).status == "pass"


# ---------------------------------------------------------------- K11 spasi & indentasi subjudul
def test_k11_spacing_and_indent_rules(tmp_path):
    assert _crit(m2_parser.parse(build_base(tmp_path / "h.docx")), 11).status == "pass"

    def spacing(d):
        h2 = next(p for p in d.paragraphs if p.style.name == "Heading 2")
        h2.paragraph_format.space_before, h2.paragraph_format.space_after = Pt(0), Pt(12)
    msgs = " | ".join(f.message for f in _crit(_mutated(tmp_path, spacing, "s.docx"), 11).findings)
    assert "Heading 2" in msgs and "spasi sebelum 0.0 pt (seharusnya 6.0 pt)" in msgs and "sesudah 12.0 pt" in msgs

    def indent(d):
        h1 = next(p for p in d.paragraphs if p.style.name == "Heading 1")
        h1.paragraph_format.left_indent, h1.paragraph_format.first_line_indent = Cm(1.27), Cm(0)
    r = _crit(_mutated(tmp_path, indent, "i.docx"), 11)
    assert r.status == "warn" and 'indentasi kiri 0.50" gantung 0.00"' in r.findings[0].message


def test_k11_reads_indent_from_numbering_definition_like_the_template():
    tpl = Path("/mnt/user-data/uploads/Template-Final.docx")
    if not tpl.exists():
        pytest.skip("template tidak tersedia")
    pd = m2_parser.parse(tpl)                          # indentasi subjudul ada di numbering.xml, bukan di paragraf
    assert {(h.left_indent_cm, h.first_indent_cm) for h in pd.headings} == {(0.64, -0.64)}
    assert _crit(pd, 11).status == "pass"              # dan spasi 6/6, 6/3, 6/2 pt sesuai aturan
    assert _crit(pd, 6).status == "pass"               # judul Cambria 16 pt tebal tengah, spasi 6 pt


def test_template_is_flagged_by_k03_because_it_has_a_single_section():
    tpl = Path("/mnt/user-data/uploads/Template-Final.docx")
    if not tpl.exists():
        pytest.skip("template tidak tersedia")
    r = _crit(m2_parser.parse(tpl), 3)
    assert r.status == "fail" and "Continuous" in r.findings[0].suggestion


def test_web_report_shows_groups(client, good_docx):
    r = post(client, good_docx)
    html = client.get(r.location).get_data(as_text=True)
    for g in ("A. Pengaturan dokumen", "B. Identitas dokumen", "C. Isi dokumen", "D. Referensi"):
        assert g in html
    assert html.index("A. Pengaturan dokumen") < html.index("B. Identitas dokumen") < html.index("D. Referensi")
    idx = client.get("/").get_data(as_text=True)
    assert "K01" in idx and "K18" in idx and "C. Isi dokumen" in idx
