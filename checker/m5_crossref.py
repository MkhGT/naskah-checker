"""M5 - Verifikasi Crossref: mencocokkan metadata dan menilai kemiripan tiap referensi.

- Referensi ber-DOI  -> GET /works/{doi}; bandingkan judul, penulis, dan tahun.
- Referensi tanpa DOI -> GET /works?query.bibliographic=...; jika cocok, sarankan DOI-nya.
- Judul yang berbeda bahasa (mis. judul Indonesia vs judul asli Inggris di Crossref) tidak
  langsung dianggap tidak cocok: bila nama penulis dan tahun terbit cocok, referensi dinyatakan
  terverifikasi dan kemiripan judul hanya dicatat.
Jaringan bisa gagal/dimatikan: hasil jatuh ke status "tidak_terverifikasi" tanpa
menggagalkan pipeline. `http_get` dapat diganti pada pengujian.
"""
import json
import os
import re
import unicodedata
import urllib.error
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from difflib import SequenceMatcher

from . import langid

API = "https://api.crossref.org/works"
TIMEOUT = 10
MAX_REFS = 60
SIM_VERIFIED = 0.85       # judul cocok
SIM_SEARCH = 0.90         # pencarian tanpa DOI dianggap ditemukan
SIM_PARTIAL = 0.50        # kemiripan minimum agar penulis+tahun cocok dapat menguatkan (jika bahasa sama)

# Email untuk "polite pool" Crossref (permintaan lebih stabil). Isi di sini, ATAU lebih baik lewat
# variabel lingkungan CROSSREF_MAILTO (nilai variabel lingkungan didahulukan).
DEFAULT_MAILTO = "3337220060@untirta.ac.id"


def http_get(url, params=None):
    """Kembalikan (status_code, json|None). Memakai pustaka standar (urllib), jadi tidak butuh
    paket tambahan. Ganti fungsi ini untuk mock di tes."""
    if params:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(params)
    mailto = os.environ.get("CROSSREF_MAILTO") or DEFAULT_MAILTO
    ua = f"NaskahChecker/1.0 (mailto:{mailto})" if mailto else "NaskahChecker/1.0"
    req = urllib.request.Request(url, headers={"User-Agent": ua, "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=TIMEOUT) as resp:
            code, body = resp.status, resp.read()
    except urllib.error.HTTPError as err:          # 404, 429, 5xx: tetap kembalikan kodenya
        code, body = err.code, err.read()
    try:
        return code, json.loads(body.decode("utf-8"))
    except ValueError:
        return code, None


def _tokens(s):
    return re.findall(r"[a-z0-9]+", (s or "").lower())


def similarity(a, b):
    """Kemiripan 0..1: maks(rasio urutan karakter, Jaccard token)."""
    ta, tb = _tokens(a), _tokens(b)
    if not ta or not tb:
        return 0.0
    seq = SequenceMatcher(None, " ".join(ta), " ".join(tb)).ratio()
    jac = len(set(ta) & set(tb)) / len(set(ta) | set(tb))
    return round(max(seq, jac), 3)


def _containment(title, text):
    """Berapa bagian token judul Crossref yang muncul di teks referensi."""
    tt, tx = _tokens(title), set(_tokens(text))
    return round(sum(t in tx for t in tt) / len(tt), 3) if tt else 0.0


def _cr_year(item):
    for key in ("issued", "published-print", "published-online", "published"):
        parts = (item.get(key) or {}).get("date-parts") or [[None]]
        if parts and parts[0] and parts[0][0]:
            return parts[0][0]
    return None


def _ascii(s):
    """Huruf kecil tanpa aksen/tanda baca, kata dipisah satu spasi."""
    s = unicodedata.normalize("NFKD", s or "")
    s = "".join(c for c in s if not unicodedata.combining(c)).lower()
    return " ".join(re.findall(r"[a-z0-9]+", s))


def guess_lang(title):
    """'id' / 'en' / None untuk judul. Memakai langdetect bila terpasang (jauh lebih akurat untuk judul
    tanpa kata fungsi), jika tidak memakai heuristik kata fungsi (lihat langid.title_language)."""
    return langid.title_language(title)


def _cr_titles(item):
    """Semua judul yang tercatat di Crossref (judul utama, judul asli, dan gabungan judul+subjudul)."""
    titles = [t for t in (item.get("title") or []) if t] + [t for t in (item.get("original-title") or []) if t]
    subs = [t for t in (item.get("subtitle") or []) if t]
    main = (item.get("title") or [""])[0]
    titles += [f"{main} {sub}" for sub in subs if main]
    return titles


def author_match(ref_text, item):
    """True bila >= 50% nama keluarga penulis Crossref (maks. 3 pertama) muncul di teks referensi.
    None bila Crossref tidak menyimpan data penulis (tidak dapat dinilai)."""
    fams = [a.get("family") or a.get("name") or "" for a in (item.get("author") or [])[:3]]
    fams = [_ascii(f) for f in fams if _ascii(f)]
    if not fams:
        return None
    text = f" {_ascii(ref_text)} "
    return sum(f" {f} " in text for f in fams) / len(fams) >= 0.5


def year_match(ref_year, cr_year):
    """True bila selisih tahun <= 1 (terbit daring vs cetak). None bila salah satu tidak ada."""
    if not ref_year or not cr_year:
        return None
    return abs(ref_year - cr_year) <= 1



def verify_reference(ref):
    out = {"num": ref["num"], "idx": ref["idx"], "doi": ref["doi"], "status": "error",
           "similarity": None, "crossref_title": None, "crossref_doi": None,
           "crossref_year": None, "message": "", "author_match": None, "year_match": None,
           "cross_language": False, "corroborated": False}
    try:
        if ref["doi"]:
            code, js = http_get(f"{API}/{urllib.parse.quote(ref['doi'], safe='/')}")
            if code == 404:
                out.update(status="doi_tidak_terdaftar", message="DOI tidak terdaftar di Crossref.")
                return out
            if code != 200 or not js:
                out["message"] = f"Crossref membalas status {code}."
                return out
            item = js.get("message", {})
        else:
            # judul Indonesia hampir tak pernah cocok dengan metadata berbahasa Inggris -> cari memakai
            # seluruh teks referensi (penulis, jurnal, tahun tidak bergantung bahasa)
            q = ref["text"] if (not ref["title"] or guess_lang(ref["title"]) == "id") else ref["title"]
            code, js = http_get(API, {"query.bibliographic": q[:300], "rows": 1,
                                      "select": "DOI,title,original-title,issued,container-title,author"})
            if code != 200 or not js:
                out["message"] = f"Crossref membalas status {code}."
                return out
            items = js.get("message", {}).get("items", [])
            if not items:
                out.update(status="tidak_ditemukan", message="Tidak ada hasil di Crossref.")
                return out
            item = items[0]
    except Exception as exc:    # jaringan mati, timeout, dsb.
        out["message"] = f"Tidak dapat menghubungi Crossref ({type(exc).__name__}: {str(exc)[:80]})."
        return out

    cr_titles = _cr_titles(item)
    cr_title = cr_titles[0] if cr_titles else ""
    out.update(crossref_title=cr_title, crossref_doi=item.get("DOI"), crossref_year=_cr_year(item))
    if ref["title"]:
        sim = max((similarity(ref["title"], t) for t in cr_titles), default=0.0)   # banding ke semua judul Crossref
    else:
        sim = max((_containment(t, ref["text"]) for t in cr_titles), default=0.0)
    out["similarity"] = sim

    lang_ref = guess_lang(ref["title"]) if ref["title"] else None
    lang_cr = (item.get("language") or "")[:2].lower() or guess_lang(cr_title)
    cross = bool(lang_ref and lang_cr and lang_ref != lang_cr)
    auth, yr = author_match(ref["text"], item), year_match(ref["year"], out["crossref_year"])
    corroborated = auth is True and yr is not False       # penulis cocok, tahun tidak bertentangan
    out.update(author_match=auth, year_match=yr, cross_language=cross)

    def same_work_note():
        why = "berbeda bahasa" if cross else "tidak identik"
        return f"Judul {why} dengan Crossref (kemiripan {sim:.0%}), tetapi penulis dan tahun cocok."

    if ref["doi"]:
        if sim >= SIM_VERIFIED:
            out.update(status="terverifikasi", message="Judul cocok dengan metadata Crossref.")
            if ref["year"] and out["crossref_year"] and ref["year"] != out["crossref_year"]:
                out.update(status="tahun_berbeda",
                           message=f"Judul cocok, tetapi tahun {ref['year']} berbeda dari Crossref ({out['crossref_year']}).")
        elif corroborated and (cross or sim >= SIM_PARTIAL):
            out.update(status="terverifikasi", corroborated=True, message=same_work_note())
        else:
            hint = (" Kemungkinan judul berbeda bahasa, tetapi penulis/tahun tidak dapat dikonfirmasi."
                    if cross else "")
            out.update(status="tidak_cocok", message="DOI terdaftar, tetapi judul berbeda dari metadata Crossref." + hint)
    else:
        if sim >= SIM_SEARCH:
            out.update(status="ditemukan_tanpa_doi", message="Ditemukan di Crossref; DOI dapat ditambahkan.")
        elif corroborated and (cross or sim >= SIM_PARTIAL):
            out.update(status="ditemukan_tanpa_doi", corroborated=True,
                       message=same_work_note() + " DOI dapat ditambahkan; periksa kecocokannya.")
        else:
            out.update(status="tidak_ditemukan", message="Tidak ditemukan padanan yang cukup mirip di Crossref.")
    return out


def run(pd, res, enabled=True, workers=4):
    """Verifikasi semua referensi; tambahkan temuan ke `res` (kriteria 14). Kembalikan tabel hasil."""
    refs = pd.references[:MAX_REFS]
    if not enabled or not refs:
        return []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        rows = list(ex.map(verify_reference, refs))
    by_idx = {r["idx"]: r for r in refs}
    failed = [r for r in rows if r["status"] == "error"]
    for r in rows:
        ref, lab = by_idx[r["idx"]], f"[{r['num'] or '?'}]"
        st = r["status"]
        if st == "doi_tidak_terdaftar":
            res.add(f"Referensi {lab}: DOI {r['doi']} tidak terdaftar di Crossref.", para=r["idx"],
                    severity="warning", source="M5", suggestion="Periksa kembali DOI; kemungkinan salah ketik.")
        elif st == "tidak_cocok":
            res.add(f"Referensi {lab}: judul tidak cocok dengan metadata Crossref (kemiripan {r['similarity']:.0%}). "
                    f"Crossref: \u201c{r['crossref_title']}\u201d."
                    + (" Kemungkinan perbedaan bahasa judul; penulis atau tahun tidak dapat dikonfirmasi." if r["cross_language"] else ""),
                    para=r["idx"], severity="warning", source="M5",
                    suggestion="Pastikan DOI dan judul merujuk pada karya yang sama.")
        elif st == "terverifikasi" and r["corroborated"]:
            res.add(f"Referensi {lab}: {r['message']}", para=r["idx"], severity="info", source="M5", comment=False)
        elif st == "tahun_berbeda":
            res.add(f"Referensi {lab}: {r['message']}", para=r["idx"], severity="info", source="M5")
        elif st == "ditemukan_tanpa_doi":
            res.add(f"Referensi {lab} ditemukan di Crossref; tambahkan DOI: {r['crossref_doi']}.", para=r["idx"],
                    severity="info", source="M5")
        elif st == "tidak_ditemukan":
            sev = "warning" if ref["journal_like"] else "info"
            res.add(f"Referensi {lab} tidak ditemukan di Crossref"
                    + ("" if ref["journal_like"] else " (buku/situs web mungkin tidak terindeks)") + ".",
                    para=r["idx"], severity=sev, source="M5",
                    suggestion="Periksa keberadaan dan ejaan referensi secara manual.")
    if failed:
        res.add(f"Verifikasi Crossref gagal untuk {len(failed)} dari {len(rows)} referensi "
                f"({failed[0]['message']}). Hasil bagian ini tidak lengkap.", severity="info",
                source="M5", comment=False)
    res.note += f"; Crossref: {sum(r['status'] == 'terverifikasi' for r in rows)}/{len(rows)} terverifikasi"
    res.finalize()
    return rows
