"""Pengujian pembangkit varian (tools/make_variants.py)."""
import csv
import hashlib
import json
import sys
import zipfile
from argparse import Namespace
from pathlib import Path

import pytest
from docx import Document
from docx.oxml import OxmlElement

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
sys.path.insert(0, str(ROOT / "tests"))

import docx_edit as E                          # noqa: E402
import make_variants as mv                     # noqa: E402
from checker import m2_parser, m3_rules, m4_citations   # noqa: E402
from variant_base import build_base            # noqa: E402

STD_CODES = sorted(c for c, o in mv.OPS.items() if o.kind == "std")


@pytest.fixture(scope="module")
def base(tmp_path_factory):
    d = tmp_path_factory.mktemp("vb")
    return build_base(d / "T01-Base.docx")


@pytest.fixture(scope="module")
def rules():
    db = m3_rules.load_rules()
    return db, {c["id"]: c["params"] for c in db["kriteria"]}, \
        next(c for c in db["kriteria"] if c["modul"] == "M4")


def statuses(path, rules):
    db, _, crit14 = rules
    pd = m2_parser.parse(path)
    res = m3_rules.RuleEngine(db).run(pd, {"jenis": "orisinal"}) + [m4_citations.run(pd, crit14)]
    return {r.id: r.status for r in res}


def args(**kw):
    base_args = dict(seed="42", only=None, sizes="3,6,9", no_multi=False, no_ref=False, no_eff=False)
    base_args.update(kw)
    return Namespace(**base_args)


# ----------------------------------------------------------------- dasar: basis harus bersih
def test_test_base_is_clean(base, rules):
    assert set(statuses(base, rules).values()) == {"pass"}


def test_catalog_has_47_single_operators_plus_crossref():
    assert len(STD_CODES) == 53 and "R01" in mv.OPS
    assert {mv.OPS[c].criterion for c in STD_CODES} == set(range(1, 19))        # 18 kriteria tercakup


# ----------------------------------------------------------------- tiap operator: terbaca & terisolasi
@pytest.mark.slow
@pytest.mark.parametrize("code", STD_CODES)
def test_each_operator_injects_its_defect_and_nothing_else(code, base, rules, tmp_path):
    work, defects, skipped, state = mv.build_variant(base, [code], code, "42:T01", tmp_path, None, rules[1])
    assert defects and not skipped, skipped
    out = tmp_path / "v.docx"
    mv.finalize(work, out, defects, state)
    Document(str(out))                                                          # berkas valid
    st, st0 = statuses(out, rules), statuses(base, rules)
    target = mv.OPS[code].criterion
    assert st[target] != "pass", f"{code}: cacat tidak terbaca"
    collateral = {k for k in st if st[k] != st0[k]} - {target} - set(mv.OPS[code].side)
    assert not collateral, f"{code}: kriteria lain ikut berubah: {sorted(collateral)}"


# ----------------------------------------------------------------- perilaku batch
def test_naming_prefix():
    assert mv.prefix_of("naskah/T01-Base.docx") == "T01"
    assert mv.prefix_of("T12_base.docx") == "T12"
    assert mv.prefix_of("Naskah Saya.docx") == "Naskah Saya"


def test_batch_output_naming_manifest_and_control(base, tmp_path):
    out = tmp_path / "out"
    out.mkdir()
    k19 = {c["id"]: c["params"] for c in m3_rules.load_rules()["kriteria"]}
    b, sk = mv.process_base(base, out, args(only="K02a,K06b,K09b"), None, k19)
    names = sorted(v["file"] for v in b["variants"])
    assert names == ["T01-K02a.docx", "T01-K06b.docx", "T01-K09b.docx", "T01-V00.docx"]
    ctl = next(v for v in b["variants"] if v["type"] == "control")
    assert ctl["sha256"] == b["sha256"] and ctl["defects"] == []               # kontrol = basis persis
    k02 = next(v for v in b["variants"] if v["file"] == "T01-K02a.docx")
    assert k02["defects"][0]["criterion"] == 2 and k02["valid_docx"]
    mv.write_outputs(out, [b], sk, args())
    rows = list(csv.DictReader(open(out / "manifest.csv", encoding="utf-8-sig")))
    assert {r["defect_code"] for r in rows} == {"V00", "K02a", "K06b", "K09b"}
    assert json.load(open(out / "manifest.json"))["bases"][0]["prefix"] == "T01"


def test_deterministic_and_seed_sensitive(base, tmp_path):
    k19 = {c["id"]: c["params"] for c in m3_rules.load_rules()["kriteria"]}
    def run(seed, name):
        out = tmp_path / name
        out.mkdir()
        b, _ = mv.process_base(base, out, args(seed=seed, only="K12a,K14a,K02a"), None, k19)
        return {v["file"]: v["sha256"] for v in b["variants"]}
    a1, a2, b1 = run("42", "a1"), run("42", "a2"), run("7", "b1")
    assert a1 == a2                                                             # seed sama -> byte identik
    assert a1["T01-K02a.docx"] == b1["T01-K02a.docx"]                           # tanpa unsur acak -> sama
    assert a1["T01-K12a.docx"] != b1["T01-K12a.docx"]                           # unsur acak -> beda


@pytest.mark.slow
def test_full_batch_composition(base, rules, tmp_path):
    out = tmp_path / "full"
    out.mkdir()
    b, sk = mv.process_base(base, out, args(), None, rules[1])
    by = {t: [v for v in b["variants"] if v["type"] == t] for t in ("control", "single", "multi", "ref", "eff")}
    assert [len(by[t]) for t in ("control", "single", "multi", "ref", "eff")] == [1, 53, 3, 1, 2]
    assert not sk
    crit = [d["criterion"] for v in by["multi"] for d in v["defects"]]
    assert sorted(crit) == list(range(1, 19))                                   # 3+6+9: tiap kriteria tepat sekali
    assert [len(v["defects"]) for v in by["multi"]] == [3, 6, 9]
    for v in by["multi"]:                                                       # tanpa pasangan yang bertentangan
        cs = {d["criterion"] for d in v["defects"]}
        assert not any(p <= cs for p in mv.CONFLICTS)
    ea, eb = by["eff"]
    assert [d["code"] for d in ea["defects"]] == [d["code"] for d in eb["defects"]]
    assert ea["sha256"] != eb["sha256"]
    labels = by["ref"][0]["defects"][0]["reference_labels"]
    assert sorted(labels.values()) == ["R01a"] * 3 + ["R01b"] * 3 + ["R01c"] * 3


def test_partition_respects_conflicts_for_many_seeds():
    import random
    ok = {c: ["x"] for c in range(1, 19)}
    for s in range(30):
        groups = mv.partition(random.Random(s), ok, [3, 6, 9])
        assert sorted(c for g in groups for c in g) == list(range(1, 19))
        assert not any(p <= set(g) for g in groups for p in mv.CONFLICTS)


def test_inapplicable_operators_are_skipped_not_forced(tmp_path):
    from conftest import make_docx                                              # basis sederhana: tanpa H2/H3
    simple = make_docx(tmp_path / "S-Base.docx")
    out = tmp_path / "o"
    out.mkdir()
    k19 = {c["id"]: c["params"] for c in m3_rules.load_rules()["kriteria"]}
    b, sk = mv.process_base(simple, out, args(only="K11b,K11c,K02a"), None, k19)
    reasons = {s["code"]: s["reason"] for s in sk}
    assert "tidak ada Heading 2" in reasons["K11b"] and "tidak ada Heading 3" in reasons["K11c"]
    assert [v["file"] for v in b["variants"] if v["type"] == "single"] == ["S-K02a.docx"]


def test_operator_failure_does_not_stop_batch(base, tmp_path, monkeypatch, rules):
    def boom(c):
        raise ValueError("percobaan galat")
    monkeypatch.setitem(mv.OPS, "K02a", mv.Op("K02a", 2, "x", boom))
    work, defects, skipped, _ = mv.build_variant(base, ["K02a", "K01a"], "t", "1", tmp_path, None, rules[1])
    assert [d["code"] for d in defects] == ["K01a"]
    assert "galat penyuntingan (ValueError" in skipped[0]["reason"]


# ----------------------------------------------------------------- K16 & metadata halaman
def test_page_variant_without_libreoffice_assumes_target_and_patches_metadata(base, rules, tmp_path):
    work, defects, skipped, state = mv.build_variant(base, ["K04a"], "p", "1", tmp_path, "/tidak/ada", rules[1])
    assert state["pages"] == 13 and state["pages_source"] == "diasumsikan"
    mv.finalize(work, tmp_path / "hasil.docx", defects, state)
    assert E.read_pages_meta(tmp_path / "hasil.docx") == 13


# ----------------------------------------------------------------- docx_edit
def _mixed_par():
    d = Document()
    p = d.add_paragraph()
    r = p.add_run("abc")
    r.italic = True
    r._r.append(OxmlElement("w:tab"))
    t = OxmlElement("w:t"); t.text = "defg"; r._r.append(t)
    r._r.append(OxmlElement("w:br"))
    t2 = OxmlElement("w:t"); t2.text = "hij"; r._r.append(t2)
    return p


@pytest.mark.parametrize("k", [2, 3, 4, 6, 8, 9, 11])
def test_split_run_handles_tabs_breaks_and_multiple_text_nodes(k):
    full = E.par_text(_mixed_par())
    p = _mixed_par()
    assert E.split_at(p, k) and E.par_text(p) == full
    assert all(r.italic for r in E.inline_runs(p))


def test_replace_range_keeps_surrounding_formatting():
    d = Document()
    p = d.add_paragraph()
    p.add_run("awal ")
    r = p.add_run("miring"); r.italic = True
    p.add_run(" akhir")
    E.replace_range(p, 2, 4, "XX")                          # ganti di tengah run pertama
    E.replace_range(p, len(E.par_text(p)), len(E.par_text(p)), "!")
    assert E.par_text(p) == "awXX miring akhir!"
    assert [r.text for r in E.inline_runs(p) if r.italic] == ["miring"]


def test_normalize_zip_makes_bytes_independent_of_time(tmp_path):
    a, b = tmp_path / "a.docx", tmp_path / "b.docx"
    for f in (a, b):
        Document().save(f)
    for f in (a, b):
        with zipfile.ZipFile(f) as z:
            pass
    E.normalize_zip(a); E.normalize_zip(b)
    assert hashlib.sha256(a.read_bytes()).hexdigest() == hashlib.sha256(b.read_bytes()).hexdigest()
