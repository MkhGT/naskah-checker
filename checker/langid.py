"""Utilitas bahasa: bahasa naskah/judul dan penanda istilah asing (Inggris) di dalam teks Indonesia.

Pembagian tugas (berdasarkan uji akurasi):
  - langdetect  : andal untuk teks panjang (naskah), kalimat penuh, dan judul >= 3 kata; TIDAK andal untuk
                  kata/frasa pendek ('dataset' -> id, 'overfitting' -> no). Dipakai hanya pada konteks andal.
  - leksikon    : istilah pendek dikenali dari daftar kata Inggris (data/foreign_en_*.txt, dibangun dari
                  wordfreq oleh tools/build_lexicon.py) sehingga TIDAK perlu wordfreq saat dijalankan.
langdetect bersifat opsional: bila belum terpasang, dipakai heuristik kata fungsi (lebih kasar).
"""
import re
from functools import lru_cache
from pathlib import Path

try:                                                    # opsional
    from langdetect import DetectorFactory, detect_langs
    DetectorFactory.seed = 0                            # hasil deterministik
    HAVE_LANGDETECT = True
except ImportError:                                     # pragma: no cover
    HAVE_LANGDETECT = False

DATA = Path(__file__).resolve().parent.parent / "data"
INDONESIAN = {"id", "ms"}          # langdetect sering menjawab 'ms' (Melayu) untuk teks Indonesia

ID_STOP = {"dan", "yang", "di", "untuk", "dengan", "pada", "dari", "dalam", "ke", "sebagai", "atau", "oleh",
           "terhadap", "antara", "adalah", "ini", "itu", "tidak", "dapat", "secara", "juga", "karena", "akan",
           "telah", "sebuah", "melalui", "serta", "tersebut", "menggunakan", "berbasis", "analisis", "penerapan",
           "pengaruh", "implementasi", "sistem", "metode", "penelitian", "studi", "kajian", "tentang",
           "pengembangan", "perancangan", "pemanfaatan", "klasifikasi", "prediksi", "deteksi"}
EN_STOP = {"the", "of", "and", "for", "in", "on", "with", "using", "a", "an", "to", "based", "by", "from", "is",
           "are", "as", "at", "that", "this", "which", "be", "can", "was", "were", "has", "have", "analysis",
           "system", "method", "approach", "study", "towards", "via", "into", "its", "their", "classification",
           "prediction", "detection"}
# kata fungsi Inggris yang ikut dianggap asing bila bersebelahan dengan istilah asing ('state of the art')
EN_FUNC = {"of", "the", "and", "for", "in", "on", "to", "a", "an", "by", "with", "from", "as", "at", "is", "are",
           "via", "per", "vs", "or", "not", "its", "be"}
WORD_RE = re.compile(r"[A-Za-z]+")
JOINER_RE = re.compile(r"[ \u00a0\-\u2010\u2011\u2013]{1,2}")      # spasi / tanda hubung antar kata satu istilah
SENT_RE = re.compile(r"[^.!?\n]+[.!?]?")

# --- penjaga singkatan / bukan-kata (mencegah 'dfl', 'pkg', 'src' dibaca sebagai bahasa Inggris) -----------
# Singkatan, nama berkas/variabel, dan akronim yang lazim di naskah teknis. Ini kebijakan bawaan; jurnal bisa
# menambah lewat `abbreviations` di rules.json (K14 -> foreign_italic).
ABBREVIATIONS = frozenset("""
    pkg src cfg conf config tbl img msg btn env fig figs eq eqs ref refs vol vols sec avg std var vars num
    def func obj lib libs dir dirs dll exe tmp temp doc docs init bool int str dfl cpu gpu tpu ram rom api sdk
    ide sql dbms rdbms crud html css xml json yaml csv pdf url uri http https ftp ssh git svn
    mse rmse mae mape mpe smape lstm gru rnn cnn knn svm pca lda nlp ocr tfidf bert gpt iot gps wifi
    dfd erd uml dtd ui ux db os pc id ip
""".split())
VOWEL_RE = re.compile(r"[aeiouy]")
CONSONANT_RUN_RE = re.compile(r"[^aeiouy]{5,}")
URL_RE = re.compile(r"(?:https?://|ftp://|www\.)\S+|[\w.+-]+@[\w-]+(?:\.[\w-]+)+|\b10\.\d{4,9}/\S+", re.I)
CODE_CHARS = set("_=\\`{}<>$#@")        # karakter yang menandai identifier/kode bila menempel pada kata
# kata fungsi Indonesia (tanpa 'di'/'ke'/'ini' yang rawan bentrok): kalimat yang memuatnya bukan kalimat Inggris utuh
ID_FUNC = frozenset({"dan", "yang", "untuk", "dengan", "pada", "dari", "dalam", "sebagai", "atau", "oleh", "tidak",
                     "dapat", "akan", "adalah", "secara", "karena", "juga", "serta", "bahwa", "telah", "sudah",
                     "tersebut", "antara", "terhadap", "melalui", "sehingga", "namun", "yaitu", "dalam"})


def looks_like_abbreviation(word):
    """True bila `word` bukan kata Inggris wajar: tanpa vokal (pkg, dfl, lstm), gugus >= 5 konsonan, atau
    termasuk daftar singkatan bawaan. Dipakai juga oleh tools/build_lexicon.py agar leksikon bersih."""
    w = (word or "").lower()
    return (not w) or w in ABBREVIATIONS or not VOWEL_RE.search(w) or bool(CONSONANT_RUN_RE.search(w))


def document_acronyms(texts, max_len=4):
    """Himpunan token (huruf kecil) yang di dokumen ini muncul sebagai akronim huruf besar ('SVM', 'PKG').
    Bentuk huruf kecilnya ('svm') lalu bukan istilah asing. Hanya token pendek (<= max_len) agar judul
    berhuruf kapital ('DATASET') tidak membebaskan kata biasa."""
    out = set()
    for t in texts:
        for m in re.finditer(r"\b[A-Z]{2,%d}\b" % max_len, t or ""):
            out.add(m.group(0).lower())
    return out


@lru_cache(maxsize=1)
def lexicon():
    def load(name):
        f = DATA / name
        return set(f.read_text(encoding="utf-8").split()) if f.exists() else set()
    return load("foreign_en_strong.txt"), load("foreign_en_weak.txt")


def _tokens(text):
    return [t.lower() for t in WORD_RE.findall(text or "")]


def _heuristic_lang(text):
    toks = _tokens(text)
    idh, enh = sum(t in ID_STOP for t in toks), sum(t in EN_STOP for t in toks)
    if idh >= 2 and idh > enh:
        return "id", min(1.0, idh / max(len(toks), 1) * 3)
    if enh >= 2 and enh > idh:
        return "en", min(1.0, enh / max(len(toks), 1) * 3)
    return None, 0.0


def detect_language(text, sample=20000):
    """(kode_bahasa, keyakinan, metode). 'ms' dipetakan ke 'id'. Dipakai untuk teks panjang (naskah)."""
    text = (text or "")[:sample]
    if HAVE_LANGDETECT and len(text.split()) >= 8:
        try:
            top = detect_langs(text)[0]
            return ("id" if top.lang in INDONESIAN else top.lang), round(top.prob, 3), "langdetect"
        except Exception:
            pass
    lang, conf = _heuristic_lang(text)
    return lang, conf, "heuristik"


def title_language(title):
    """'id' / 'en' / None untuk judul (kalimat pendek). Hanya mengembalikan id/en; selain itu None."""
    toks = _tokens(title)
    if HAVE_LANGDETECT and len(toks) >= 3:
        try:
            top = detect_langs(title)[0]
            code = "id" if top.lang in INDONESIAN else top.lang
            if code in ("id", "en") and top.prob >= 0.7:
                return code
        except Exception:
            pass
    return _heuristic_lang(title)[0]


def _merge(spans):
    out = []
    for s, e, kind in sorted(spans):
        if out and s <= out[-1][1]:
            out[-1] = (out[-1][0], max(e, out[-1][1]), out[-1][2])
        else:
            out.append((s, e, kind))
    return out


def _id_function_hits(sentence):
    return len({t for t in _tokens(sentence) if t in ID_FUNC})


def foreign_spans(text, allowlist=(), force_terms=(), extra_patterns=(), title_case_phrases=True,
                  sentence_detection=True, min_sentence_words=8, sentence_prob=0.95,
                  abbreviations=(), doc_acronyms=()):
    """Daftar (mulai, akhir, jenis) istilah asing dalam `text` (naskah berbahasa Indonesia).

    Aturan: kata dari leksikon Inggris; kata kuat cukup sendiri, kata lemah hanya bila bersebelahan dengan
    kata kuat; kata fungsi (of, the, ...) ikut bila di tengah istilah. Dilewati: akronim (SVM), camelCase,
    kata <3 huruf, kata berhuruf kapital tunggal (nama diri / awal kalimat), dan kata di `allowlist`.
    Kata di `force_terms` selalu dianggap asing (untuk istilah yang lazim di web Indonesia sehingga tak
    terbedakan dari frekuensi, mis. download, software).
    Kalimat penuh berbahasa Inggris (langdetect, >= 8 kata) ditandai utuh.

    Penjaga bukan-istilah (hanya untuk kata yang lolos leksikon; `force_terms` tetap menang): singkatan bawaan +
    `abbreviations` (kebijakan jurnal), kata tanpa vokal/gugus konsonan tak wajar (pkg, dfl), akronim huruf
    besar yang ada di dokumen (`doc_acronyms`, mis. SVM -> svm), kata di dalam URL/email/DOI, dan kata yang
    menempel pada karakter kode (pkg_1, model.pkl, `dfl`). Kalimat yang memuat >= 2 kata fungsi Indonesia
    tidak ditandai sebagai kalimat Inggris."""
    strong, weak = lexicon()
    allow = {a.lower() for a in allowlist}
    force = {f.lower() for f in force_terms}
    abbr = {a.lower() for a in abbreviations} | {a.lower() for a in doc_acronyms}
    skip_zones = [(m.start(), m.end()) for m in URL_RE.finditer(text)]
    toks = [(m.start(), m.end(), m.group(0)) for m in WORD_RE.finditer(text)]

    def guarded(s, e, lw):
        """Kata yang menurut leksikon 'Inggris' tetapi sebenarnya singkatan / identifier / bagian URL."""
        if lw in abbr or looks_like_abbreviation(lw):
            return True
        if any(a <= s and e <= b for a, b in skip_zones):
            return True
        before, after = text[s - 1] if s > 0 else "", text[e] if e < len(text) else ""
        if before in CODE_CHARS or after in CODE_CHARS or before.isdigit() or after.isdigit():
            return True
        if after == "." and e + 1 < len(text) and text[e + 1].isalpha():              # model.pkl, pkg.version
            return True
        if before == "." and s >= 2 and text[s - 2].isalpha():                          # df.dfl
            return True
        if after == "/":                                                                # jalur: src/utils, pkg/dfl
            m = re.match(r"/([A-Za-z]+)", text[e:])
            if m and (m.group(1).lower() in abbr or looks_like_abbreviation(m.group(1))):
                return True
        if before == "/":
            m = re.search(r"([A-Za-z]+)/$", text[:s])
            if m and (m.group(1).lower() in abbr or looks_like_abbreviation(m.group(1))):
                return True
        return False

    label = []
    for s_, e_, w in toks:
        lw = w.lower()
        cap = w[0].isupper() and w[1:].islower() and len(w) > 1
        if lw in allow or (w.isupper() and len(w) > 1) or (not w.islower() and not cap):   # akronim / camelCase
            label.append(("N", cap))
        elif len(lw) < 3 and lw not in EN_FUNC:
            label.append(("N", cap))
        elif lw in force:                                  # kebijakan eksplisit jurnal: selalu asing
            label.append(("S", cap))
        elif lw in strong or lw in weak:
            if guarded(s_, e_, lw):
                label.append(("N", cap))
            else:
                label.append(("S" if lw in strong else "W", cap))
        elif lw in EN_FUNC:
            label.append(("F", cap))
        else:
            label.append(("N", cap))

    spans, i = [], 0
    while i < len(toks):
        if label[i][0] == "N":
            i += 1
            continue
        j = i                                              # rantai kata asing yang bersebelahan
        while (j + 1 < len(toks) and label[j + 1][0] != "N"
               and JOINER_RE.fullmatch(text[toks[j][1]:toks[j + 1][0]] or "x")
               and not (label[j + 1][1] and not label[j][1])):     # 'framework Flask': nama diri memutus rantai
            j += 1
        chain = list(range(i, j + 1))
        while chain and label[chain[0]][0] == "F":         # buang kata fungsi di tepi
            chain.pop(0)
        while chain and label[chain[-1]][0] == "F":
            chain.pop()
        if chain and any(label[k][0] == "S" for k in chain):
            caps = [label[k][1] for k in chain]
            ok = True
            if len(chain) == 1 and caps[0]:                # satu kata berhuruf kapital = nama diri/awal kalimat
                ok = False
            elif all(caps) and not title_case_phrases:
                ok = False
            if ok:
                spans.append((toks[chain[0]][0], toks[chain[-1]][1], "istilah"))
        i = j + 1

    for pat in extra_patterns:                            # mis. 'et al', 'i.e.'
        for m in re.finditer(pat, text, re.I):
            spans.append((m.start(), m.end(), "pola"))

    if sentence_detection and HAVE_LANGDETECT:
        for m in SENT_RE.finditer(text):
            sent = m.group(0)
            if len(sent.split()) >= min_sentence_words:
                try:
                    top = detect_langs(sent)[0]
                except Exception:
                    continue
                if top.lang == "en" and top.prob >= sentence_prob and _id_function_hits(sent) < 2:
                    s = m.start() + (len(sent) - len(sent.lstrip()))
                    spans.append((s, m.start() + len(sent.rstrip()), "kalimat"))
    return _merge(spans)
