#!/usr/bin/env python3
"""Membangun leksikon istilah Inggris untuk K14 (istilah asing harus miring).

Dijalankan SEKALI oleh pengembang (butuh `pip install wordfreq`); hasilnya berupa berkas teks di data/
sehingga pengguna sistem tidak perlu memasang wordfreq (57 MB).

Kriteria (skor Zipf wordfreq):
  kuat  : Inggris >= 1.5 dan tidak ada di korpus Indonesia, atau Inggris - Indonesia >= 0.75 (Inggris >= 2.3)
  lemah : selisih 0.5-0.75 dan Inggris >= 3.0 (hanya dianggap asing bila bersebelahan dengan kata kuat,
          mis. 'vector' pada 'support vector machine')
Kata dasar Indonesia (kamus Sastrawi, memuat serapan KBBI) DIKELUARKAN: kata yang sudah baku di Indonesia
tidak perlu dimiringkan, dan banyak kata Indonesia jarang kebetulan ada juga di korpus Inggris.
Kebersihan leksikon (agar singkatan seperti dfl, pkg, src tidak dianggap bahasa Inggris):
  - kata <= 4 huruf hanya dimasukkan bila ada di kamus Inggris sungguhan (gcide + web2 dari paket english-words);
    korpus web penuh singkatan, nama, dan kode yang frekuensinya tinggi tetapi bukan kata;
  - kata tanpa vokal, bergugus >= 5 konsonan, atau terdaftar di checker.langid.ABBREVIATIONS dibuang.
Pembangunan ulang membutuhkan:  pip install wordfreq Sastrawi english-words
Kata yang selisihnya ~0 (data, model, internet, online, server, ...) sengaja tidak dimasukkan karena
sudah lazim di korpus Indonesia.
"""
import re
import sys
from pathlib import Path
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
from english_words import get_english_words_set
from wordfreq import top_n_list, zipf_frequency as z

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from checker.langid import looks_like_abbreviation                    # satu sumber kebenaran dengan runtime

ID_ROOTS = set(StemmerFactory().get_words())
EN_DICT = get_english_words_set(["gcide"], lower=True, alpha=True) | get_english_words_set(["web2"], lower=True, alpha=True)
SHORT_MAX = 4                                                          # kata sependek ini wajib ada di kamus

OUT = Path(__file__).resolve().parent.parent / "data"
strong, weak = [], []
for w in top_n_list("en", 200000):
    if not re.fullmatch(r"[a-z]{3,}", w) or w in ID_ROOTS:
        continue
    if looks_like_abbreviation(w) or (len(w) <= SHORT_MAX and w not in EN_DICT):
        continue
    en, idn = z(w, "en"), z(w, "id")
    if (idn == 0 and en >= 1.5) or (en >= 2.3 and en - idn >= 0.75):
        strong.append(w)
    elif en >= 3.0 and 0.5 <= en - idn < 0.75:
        weak.append(w)
(OUT / "foreign_en_strong.txt").write_text("\n".join(sorted(strong)) + "\n", encoding="utf-8")
(OUT / "foreign_en_weak.txt").write_text("\n".join(sorted(weak)) + "\n", encoding="utf-8")
print(f"kuat={len(strong)} lemah={len(weak)}")
