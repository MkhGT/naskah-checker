"""Tes: (1) header/footer dan identitas artikel harus berada di dalam tabel (seperti template);
(2) singkatan (dfl, pkg, src, ...) tidak boleh dianggap istilah asing berbahasa Inggris."""
import copy
from pathlib import Path
from types import SimpleNamespace

import pytest
from docx import Document
from docx.oxml.ns import qn

from checker import langid, m2_parser, m3_rules
from checker.pipeline import run_pipeline
from conftest import make_docx, unwrap_tables

HF_PARTS = {                      # nama bagian python-docx -> kunci wilayah parser + label pesan
    "first_page_header": ("header_first", "Header halaman pertama"),
    "even_page_header": ("header_even", "Header halaman genap"),
    "first_page_footer": ("footer_first", "Footer halaman pertama"),
    "footer": ("footer_odd", "Footer halaman ganjil"),
    "even_page_footer": ("footer_even", "Footer halaman genap"),
}


def table_msgs(path, crit, opts=None, db=None):
    """Pesan temuan kriteria `crit` yang menyangkut tabel."""
    pd = m2_parser.parse(path)
    res = next(r for r in m3_rules.RuleEngine(db).run(pd, opts or {}) if r.id == crit)
    return pd, res, [f.message for f in res.findings if "tabel" in f.message]


def resave(src, dst, mutate):
    d = Document(src)
    mutate(d)
    d.save(dst)
    return dst


# ======================================================================= header & footer (K05)
def test_good_document_headers_and_footers_are_in_tables(tmp_path):
    pd, res, msgs = table_msgs(make_docx(tmp_path / "ok.docx"), 5)
    assert msgs == [] and res.status == "pass"
    hf = pd.headers_footers
    assert all(hf[k]["in_table"] for k, _ in HF_PARTS.values())
    assert hf["header_odd"]["has_table"] is False          # di template, header ganjil memang paragraf biasa


def test_plain_paragraph_headers_footers_are_flagged(tmp_path):
    """Isi header/footer benar, tetapi tidak dalam tabel: dulu lolos, sekarang salah."""
    _, res, msgs = table_msgs(make_docx(tmp_path / "plain.docx", hf="plain"), 5)
    assert res.status == "fail" and len(msgs) == 5
    for _, label in HF_PARTS.values():
        assert any(m.startswith(label) and "tidak berada di dalam tabel" in m for m in msgs)
    assert not any("ganjil" in m and m.startswith("Header") for m in msgs)     # header ganjil tidak dituntut


@pytest.mark.parametrize("part", list(HF_PARTS))
def test_only_the_unwrapped_region_is_flagged(tmp_path, part):
    good = make_docx(tmp_path / "ok.docx")
    bad = resave(good, tmp_path / "bad.docx", lambda d: unwrap_tables(getattr(d.sections[0], part)._element))
    _, res, msgs = table_msgs(bad, 5)
    assert len(msgs) == 1 and msgs[0].startswith(HF_PARTS[part][1]), msgs
    assert res.status == "fail"


def test_text_outside_the_table_is_flagged_as_partial(tmp_path):
    good = make_docx(tmp_path / "ok.docx")
    bad = resave(good, tmp_path / "bad.docx", lambda d: d.sections[0].even_page_footer.add_paragraph("teks liar"))
    pd, _, msgs = table_msgs(bad, 5)
    r = pd.headers_footers["footer_even"]
    assert r["has_table"] and not r["in_table"]
    assert len(msgs) == 1 and "sebagian isinya" in msgs[0] and msgs[0].startswith("Footer halaman genap")


def test_empty_trailing_paragraph_after_table_is_not_outside_content(tmp_path):
    pd = m2_parser.parse(make_docx(tmp_path / "ok.docx"))
    el = Document(tmp_path / "ok.docx").sections[0].footer._element
    assert [c.tag.split("}")[1] for c in el.iterchildren()] == ["tbl", "p"]      # tabel + paragraf kosong penutup
    assert pd.headers_footers["footer_odd"]["in_table"]


def test_require_table_can_be_switched_off_in_rules(tmp_path):
    db = m3_rules.load_rules()
    k5 = next(k for k in db["kriteria"] if k["id"] == 5)
    k5["params"]["require_table"] = []
    _, res, msgs = table_msgs(make_docx(tmp_path / "plain.docx", hf="plain"), 5, db=db)
    assert msgs == [] and res.status == "pass"


def test_require_table_can_include_odd_header(tmp_path):
    db = m3_rules.load_rules()
    next(k for k in db["kriteria"] if k["id"] == 5)["params"]["require_table"].append("header_odd")
    _, _, msgs = table_msgs(make_docx(tmp_path / "ok.docx"), 5, db=db)
    assert len(msgs) == 1 and msgs[0].startswith("Header halaman ganjil")


# ======================================================================= identitas artikel (K07-K09)
def identity_findings(path):
    pd = m2_parser.parse(path)
    by = {r.id: r for r in m3_rules.RuleEngine().run(pd, {})}
    return pd, {k: [f.message for f in by[k].findings if "tabel" in f.message or "Informasi Artikel" in f.message]
                for k in (7, 8, 9)}


def test_good_document_identity_is_in_one_table(tmp_path):
    pd, msgs = identity_findings(make_docx(tmp_path / "ok.docx"))
    assert msgs == {7: [], 8: [], 9: []}
    tables = {e["table_index"] for v in pd.identity.values() for e in v}
    assert tables == {0}
    assert {k for k, v in pd.identity.items() if v} == {"informasi", "abstrak", "kata_kunci", "korespondensi"}


def test_identity_typed_as_plain_paragraphs_is_flagged(tmp_path):
    pd, msgs = identity_findings(make_docx(tmp_path / "plain.docx", identity="plain"))
    assert any("Abstrak tidak berada di dalam tabel" in m for m in msgs[8])
    assert any("Kata Kunci tidak berada di dalam tabel" in m for m in msgs[9])
    assert any("Penulis Korespondensi tidak berada di dalam tabel" in m for m in msgs[7])
    assert any("'Informasi Artikel'" in m and "tidak ditemukan" in m for m in msgs[7])     # blok hilang sama sekali


def test_identity_unwrapped_from_template_style_table_is_flagged(tmp_path):
    """Tata letak persis template (label sendiri, tiap kata kunci di baris baru) tetapi bukan tabel."""
    good = make_docx(tmp_path / "ok.docx")
    bad = resave(good, tmp_path / "bad.docx", lambda d: unwrap_tables(d.element.body))
    pd, msgs = identity_findings(bad)
    assert pd.keywords == ["akurasi", "klasifikasi", "pembelajaran mesin"]            # tetap terbaca (bukan 'tidak ditemukan')
    labels = [m.split(" tidak berada")[0] for m in msgs[7]]
    assert {"Informasi Artikel", "Riwayat Artikel", "Penulis Korespondensi"} <= set(labels)
    assert any(m.startswith("Abstrak tidak berada") for m in msgs[8])
    assert any(m.startswith("Kata Kunci tidak berada") for m in msgs[9])


def test_identity_split_across_two_tables_warns_but_each_is_in_table(tmp_path):
    pd, msgs = identity_findings(make_docx(tmp_path / "split.docx", identity="split"))
    assert pd.keywords == ["akurasi", "klasifikasi", "pembelajaran mesin"]            # kata kunci di tabel kedua tetap ditemukan
    assert not any("tidak berada di dalam tabel" in m for k in (7, 8, 9) for m in msgs[k])
    assert any("tabel yang berbeda" in m for m in msgs[7])
    k7 = next(r for r in m3_rules.RuleEngine().run(pd, {}) if r.id == 7)
    assert all(f.severity == "warning" for f in k7.findings if "tabel yang berbeda" in f.message)


def test_identity_flags_point_to_the_offending_paragraph_and_become_comments(tmp_path):
    src = make_docx(tmp_path / "plain.docx", identity="plain")
    report = run_pipeline(src, tmp_path / "job", use_crossref=False)
    flat = [f for c in report["kriteria"] for f in c["findings"] if "tidak berada di dalam tabel" in f["message"]]
    assert {f["message"].split(" tidak")[0] for f in flat} >= {"Abstrak", "Kata Kunci", "Penulis Korespondensi"}
    assert all(f["para_index"] is not None for f in flat)               # ada lokasi -> komentar menempel di paragraf itu
    texts = [c.text for c in Document(tmp_path / "job" / "berkomentar.docx").comments]
    assert any("[K08]" in t and "tidak berada di dalam tabel" in t for t in texts)
    assert any("[K09]" in t and "tidak berada di dalam tabel" in t for t in texts)


def test_identity_table_requirement_can_be_switched_off(tmp_path):
    db = m3_rules.load_rules()
    for k in db["kriteria"]:
        if k["id"] == 8 or k["id"] == 9:
            k["params"]["require_in_table"] = False
        if k["id"] == 7:
            k["params"]["identity_in_table"] = []
            k["params"]["require_article_info"] = False
            k["params"]["identity_single_table"] = False
    pd = m2_parser.parse(make_docx(tmp_path / "plain.docx", identity="plain"))
    res = m3_rules.RuleEngine(db).run(pd, {})
    assert not any("tabel" in f.message for r in res if r.id in (7, 8, 9) for f in r.findings)


def test_real_template_has_no_table_findings():
    tpl = Path("/mnt/user-data/uploads/Template-Final.docx")
    if not tpl.exists():
        pytest.skip("template tidak tersedia")
    for crit in (5, 7, 8, 9):
        assert table_msgs(tpl, crit)[2] == []


# ======================================================================= langid: singkatan bukan istilah asing
def spans(text, **kw):
    return [text[s:e] for s, e, _ in langid.foreign_spans(text, **kw)]


@pytest.mark.parametrize("text", [
    "Kolom dfl dan pkg dipakai sebagai fitur pada tahap awal",
    "Parameter src, cfg, env, img, tbl, msg dan btn diatur sebelum pengujian",
    "Nilai fig dan vol dibaca dari berkas konfigurasi tersebut",
    "Algoritma lstm, knn, svm, mse dan avg dibandingkan satu per satu",
])
def test_abbreviations_are_not_foreign_terms(text):
    assert spans(text) == []


def test_lexicon_no_longer_contains_abbreviations():
    strong, weak = langid.lexicon()
    for w in "dfl pkg src cfg tbl img env mse lstm knn svm avg num def func obj init json dfd uml ocr nlp".split():
        assert w not in strong and w not in weak, w


def test_real_english_terms_are_still_detected():
    t = "Model dioptimasi dengan overfitting pada dataset melalui machine learning dan backend yang baru"
    assert spans(t) == ["overfitting", "dataset", "machine learning", "backend"]


def test_code_and_path_context_is_not_foreign():
    assert spans("Berkas model.pkl dan pkg_1 disimpan di src/utils pada folder kerja") == []
    assert spans("Tautan https://github.com/user/overfitting/dataset dan nama@backend.id dicantumkan") == []


def test_document_acronyms_free_their_lowercase_form():
    """Bila dokumen memuat akronim huruf besar 'TASK', bentuk huruf kecil 'task' bukan lagi istilah asing."""
    assert "task" in langid.lexicon()[0]                                # prasyarat: kata ini memang ada di leksikon
    t = "Setiap task dijalankan berurutan oleh penjadwal"
    assert spans(t) == ["task"]                                          # tanpa bukti: terdeteksi sebagai istilah asing
    acr = langid.document_acronyms(["Modul TASK menjadwalkan pekerjaan"])
    assert acr == {"task"}
    assert spans(t, doc_acronyms=acr) == []                              # dengan bukti akronim di dokumen: bebas
    assert langid.document_acronyms(["Metode SVM dan KNN dipakai", "tanpa singkatan"]) == {"svm", "knn"}
    assert langid.document_acronyms(["DATASET PENELITIAN"]) == set()           # kata panjang berhuruf kapital (judul) bukan bukti


def test_doc_acronym_evidence_only_applies_to_short_tokens():
    # 'dataset' (7 huruf) tidak pernah bebas hanya karena ada 'DATASET' di judul
    assert "dataset" not in langid.document_acronyms(["DATASET PENELITIAN"])


def test_force_terms_override_the_abbreviation_guard():
    assert spans("Kolom pkg dipakai", force_terms=["pkg"]) == ["pkg"]


def test_custom_abbreviations_list_from_rules_is_honoured():
    t = "Proses download dan framework dijalankan"
    assert "framework" in spans(t)
    assert "framework" not in spans(t, abbreviations=["framework"])


def test_looks_like_abbreviation_rules():
    for w in ("pkg", "dfl", "lstm", "dbms", "SRC", "fig", "x"):
        assert langid.looks_like_abbreviation(w), w
    for w in ("dataset", "overfitting", "framework", "learning", "cloud"):
        assert not langid.looks_like_abbreviation(w), w


class _Lang:
    def __init__(self, lang, prob):
        self.lang, self.prob = lang, prob


def test_indonesian_sentence_misread_as_english_is_not_marked(monkeypatch):
    monkeypatch.setattr(langid, "detect_langs", lambda s: [_Lang("en", 0.99)])
    monkeypatch.setattr(langid, "HAVE_LANGDETECT", True)
    indo = "Penelitian ini menggunakan dfl dan pkg yang dapat digunakan untuk analisis secara efektif dan efisien"
    assert spans(indo) == []                       # >= 2 kata fungsi Indonesia -> bukan kalimat Inggris
    eng = "The proposed model achieves higher accuracy than the baseline on every benchmark tested"
    assert spans(eng) == [eng]                     # kalimat Inggris sungguhan tetap ditandai utuh


def test_foreign_kwargs_collects_acronyms_but_ignores_headings():
    paras = [SimpleNamespace(role="heading", text="METODE SVM DATASET"),
             SimpleNamespace(role="body", text="Kami memakai KNN dan PKG untuk uji"),
             SimpleNamespace(role="title", text="JUDUL XYZ")]
    prm = {"allowlist": [], "force_terms": [], "extra_patterns": [], "title_case_phrases": True,
           "sentence_detection": False, "abbreviations": ["abc"]}
    kw = m3_rules.foreign_kwargs(SimpleNamespace(paras=paras), prm)
    assert kw["doc_acronyms"] == {"knn", "pkg"} and kw["abbreviations"] == ["abc"]
