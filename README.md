# Pemeriksa Naskah JISINTA

Aplikasi web untuk memeriksa apakah naskah `.docx` sudah sesuai dengan template jurnal. Penulis mengunggah naskah,
lalu mendapat laporan per kriteria dan salinan naskah yang diberi komentar Word pada bagian yang perlu diperbaiki.
Naskah asli tidak diubah.

Aturan pemeriksaan dibaca dari `data/rules.json`, yang disusun dari `Template-Final.docx`.

## Menjalankan di komputer sendiri

```bash
pip install -r requirements.txt
python app.py        # buka http://127.0.0.1:5000
```

## Deploy ke server

Aplikasi ini hanya berupa web (Flask), jadi cukup dijalankan dengan gunicorn di belakang nginx atau Apache.

```bash
python3 -m venv venv
venv/bin/pip install -r requirements.txt
venv/bin/gunicorn wsgi:app --workers 2 --bind 127.0.0.1:8000 --timeout 120
```

`--timeout 120` dipasang karena pengecekan Crossref bisa memakan waktu beberapa detik.

Contoh blok nginx:

```nginx
location / {
    proxy_pass http://127.0.0.1:8000;
    proxy_set_header Host $host;
    proxy_set_header X-Forwarded-For $remote_addr;
    client_max_body_size 11m;
}
```

Pengaturan lewat environment variable:

| Variabel | Fungsi | Bawaan |
|---|---|---|
| `TEMPLATE_URL` | alamat unduhan template, tampil di kotak "Belum punya template?" pada halaman upload | kosong (tautan nonaktif) |
| `RETENTION_MINUTES` | lama naskah dan hasil disimpan di server | 30 |
| `JOBS_DIR` | folder penyimpanan sementara | `instance/jobs` |
| `CROSSREF_ENABLED` | `0` untuk mematikan verifikasi Crossref | `1` |
| `CROSSREF_MAILTO` | email untuk antrean "polite" Crossref | kosong |
| `FLASK_DEBUG` | `1` hanya untuk pengembangan lokal | `0` |

Folder `tests/` dan `tools/` dipakai untuk pengujian dan penelitian, tidak dibutuhkan di server.

## Alamat template

Tautan template di halaman upload bisa diisi dengan dua cara: set `TEMPLATE_URL`, atau ubah langsung bagian
`tpl-box` di `templates/index.html`.

## Struktur

| Modul | Berkas | Tugas |
|---|---|---|
| M1 | `checker/m1_upload.py` | validasi unggahan (ekstensi, isi zip, ukuran, makro, ukuran desimal), pengenalan aplikasi pembuat, penghapusan berkas |
| M2 | `checker/m2_parser.py` | membaca paragraf, heading, tabel, gambar, caption, section, header/footer, abstrak, kata kunci, referensi |
| M3 | `checker/m3_rules.py` | memeriksa kriteria 1-17 berdasarkan `rules.json` |
| M4 | `checker/m4_citations.py` | sitasi dan daftar pustaka (kriteria 18) |
| M5 | `checker/m5_crossref.py` | mencocokkan referensi dengan Crossref |
| M6 | `checker/m6_annotate.py` | menulis komentar Word pada salinan naskah |
| M7 | `checker/m7_report.py` | menyusun laporan JSON |
| Web | `app.py`, `templates/`, `static/` | halaman upload, laporan, unduhan |

## Kriteria pemeriksaan

| | Kriteria | Yang diperiksa |
|---|---|---|
| **A. Pengaturan dokumen** | | |
| K01 | Ukuran kertas | A4, pada semua section |
| K02 | Margin | 2,5 / 2 / 2,5 / 2 cm (atas / bawah / kiri / kanan), semua section |
| K03 | Layout kolom | bagian identitas (sebelum Pendahuluan) 1 kolom; bagian isi sesuai `body_columns` (bawaan 2) |
| K04 | Batas halaman | maks. 12 (orisinal) / 16 (review) |
| **B. Identitas dokumen** | | |
| K05 | Header dan footer | halaman pertama, ganjil, genap; field nomor halaman; isi harus berada di dalam tabel |
| K06 | Judul | Cambria 16 pt, tebal, tengah, spasi sesudah 6 pt; maks. 15 kata; tanpa kata terlarang |
| K07 | Penulis, afiliasi, korespondensi | penulis Cambria 10 pt, afiliasi Cambria 9 pt, tengah, spasi sesudah 6 pt; penanda angka cocok; email; Informasi Artikel dan Penulis Korespondensi harus di dalam tabel identitas; penulis korespondensi ditulis per baris (nama, afiliasi/instansi, email) |
| K08 | Abstrak | 100-250 kata, tanpa sitasi; harus di dalam tabel identitas; isi Cambria 10 pt rata kiri-kanan |
| K09 | Kata kunci | 3-5, urut abjad; harus di dalam tabel identitas; rata kiri |
| **C. Isi dokumen** | | |
| K10 | Struktur IMRaD | bagian wajib dan urutannya; bagian lain (mis. Tinjauan Pustaka) tidak dianggap sesuai |
| K11 | Format subjudul | nomor rata di 0, teks menjorok 0,25 in (gantung 0,25 in); spasi H1 6/6, H2 6/3, H3 6/2 pt; tebal/kapital per level |
| K12 | Jenis dan ukuran huruf, spasi | Cambria 10 pt, spasi 1,15 |
| K13 | Jumlah paragraf Pendahuluan | 3-6 |
| K14 | Format paragraf isi | indentasi 1,27 cm, rata kiri-kanan, istilah asing miring (naskah berbahasa Indonesia); tanpa spasi ganda, spasi/Tab di awal atau akhir paragraf, spasi sebelum tanda baca, dan baris kosong pengatur jarak |
| K15 | Caption, posisi, perataan gambar/tabel, isi tabel | caption tabel di atas, gambar di bawah, tengah, 8 pt; isi tabel Cambria 8 pt rata tengah, baris pertama tebal |
| K16 | Gambar dan tabel disitasi | dirujuk di teks |
| K17 | Kesimpulan tidak mengulang abstrak | kemiripan kalimat/frasa; memuat saran |
| **D. Referensi** | | |
| K18 | Sitasi, daftar pustaka IEEE, Crossref | M4 + M5 |

Ambang dan istilah tiap kriteria ada di `params` pada `data/rules.json`, jadi bisa diubah tanpa menyentuh kode.

Catatan:
- K03: template contoh hanya punya satu section satu kolom, jadi jumlah kolom isi tidak bisa dibaca dari template.
  Nilai `body_columns` diset 2; ubah ke 1 bila isi naskah tetap satu kolom.
- K11: "left 0 hanging 0,25 in" dibaca sebagai nomor rata di 0 dan teks menjorok 0,25 in, sama seperti template.
- Pembagian kriteria ke kelompok A-D adalah rancangan sendiri karena isi Tabel 3.3 asli tidak tersedia. Sesuaikan
  `rules.json` bila tabel aslinya berbeda.

## Pemeriksaan tabel (K05, K07, K08, K09)

Di template, header, footer, dan kotak identitas artikel dibuat dengan tabel. Naskah yang tampilannya sama tetapi
diketik sebagai paragraf biasa dianggap tidak sesuai.

- K05: wilayah yang wajib bertabel ada di `require_table`. Bawaannya header halaman pertama, header genap, dan ketiga
  footer. Header ganjil tidak termasuk karena di template berupa paragraf biasa; tambahkan `"header_odd"` bila perlu.
- K07: label "Informasi Artikel", "Riwayat Artikel", dan "Penulis Korespondensi" harus berada di dalam tabel (`identity_in_table`).
- K08 dan K09: abstrak dan kata kunci harus di dalam tabel (`require_in_table`).
- Format abstrak (font, ukuran, perataan) dan perataan kata kunci diatur di `params.format` pada K08 dan K09 di `rules.json`. Hasilnya berupa peringatan, bukan kesalahan.
- K07: penulis korespondensi harus berbentuk label, lalu nama, afiliasi/instansi, dan email, masing-masing di paragraf sendiri (`corresponding_lines`). Email boleh diawali "Email:" tetapi tidak boleh digabung dengan teks lain.
- K15: isi tabel di luar kotak identitas diperiksa lewat `table_content` (font, ukuran, perataan, `header_bold` untuk baris pertama). Temuan diringkas per tabel dengan jumlah sel yang menyimpang.
- K14: `blank_space` memeriksa spasi ganda, spasi/Tab di awal atau akhir paragraf, spasi sebelum tanda baca, dan baris kosong yang dipakai sebagai jarak. Baris kosong yang memuat jeda halaman, akhir section, atau tepat setelah tabel diabaikan. Tiap pemeriksaan bisa dimatikan satu per satu di `rules.json`.
- Yang dicek hanya letaknya di dalam tabel, bukan gaya garis, lebar kolom, atau jumlah baris.
- Hanya header/footer section pertama yang dibaca.
- Tiap pemeriksaan bisa dimatikan dari `rules.json` tanpa mengubah kode.

## Istilah asing dicetak miring (bagian K14)

Hanya berlaku untuk naskah berbahasa Indonesia (bahasa dideteksi otomatis). Yang diperiksa adalah paragraf isi dan daftar,
bukan heading, caption, tabel, abstrak, atau daftar pustaka.

- `langdetect` dipakai untuk menentukan bahasa naskah dan kalimat Inggris yang utuh. Pada kata atau frasa pendek ia kurang andal,
  jadi istilah pendek dicocokkan dengan leksikon di `data/foreign_en_strong.txt` dan `foreign_en_weak.txt`.
- Leksikon dibuat dengan `tools/build_lexicon.py` (hanya untuk pengembang, butuh `wordfreq` dan `Sastrawi`).
- Singkatan seperti `dfl` atau `pkg` tidak dianggap bahasa Inggris: kata <= 4 huruf hanya masuk leksikon bila ada di kamus
  Inggris, dan saat berjalan dicek lagi lewat daftar singkatan, bentuk kata yang tidak wajar, akronim di dokumen, URL/email/DOI,
  serta kata yang menempel pada karakter kode.
- Daftar singkatan khas bidang bisa ditambah di `rules.json` pada K14 -> `foreign_italic.abbreviations`.
- Kebijakan jurnal diatur lewat `force_terms` (selalu dianggap asing), `allowlist` (boleh tanpa miring), dan
  `extra_patterns` (pola tambahan seperti `et al`).
- Hasil deteksi bisa meleset, jadi temuannya berstatus peringatan.

## Privasi dan masa simpan

Naskah dan hasil pemeriksaan disimpan sementara di `JOBS_DIR/<id>/` (`asli.docx`, `berkomentar.docx`, `laporan.json`)
karena halaman laporan dan tombol unduh membacanya kembali.

- Masa simpan 30 menit sejak pemeriksaan selesai, bisa diubah dengan `RETENTION_MINUTES`. Tidak bisa dimatikan.
- Tombol "Hapus sekarang" di halaman laporan menghapus folder job seketika.
- Berkas yang sudah kedaluwarsa tidak dilayani (langsung dihapus dan dijawab 404). Penyapu di latar belakang
  membersihkan job kedaluwarsa tiap 60 detik, dan sisa job dari proses sebelumnya dihapus saat server dinyalakan.
- Hanya folder ber-ID job yang dihapus; berkas lain di `JOBS_DIR` tidak disentuh.
- Verifikasi Crossref hanya mengirim judul/DOI referensi, bukan isi naskah.
- Penghapusan berlaku di sistem berkas aplikasi. Cadangan server, snapshot disk, dan log web server/proxy ada di luar kendali aplikasi.

## Batasan

- Sistem dibuat dan diuji untuk `.docx` dari Microsoft Word. Berkas dengan ukuran desimal (umum pada ekspor Google Docs)
  ditolak dengan pesan yang jelas. Berkas dari aplikasi lain tanpa desimal tetap diperiksa, tetapi laporan memberi
  peringatan bahwa hasilnya bisa berbeda.
- Pada LibreOffice, 17 dari 18 kriteria menghasilkan status yang sama dengan Word dalam uji sampel kecil; perbedaannya ada di
  pembacaan daftar pustaka Mendeley (K18).
- Jumlah halaman (K04) dibaca dari `docProps/app.xml`; bila tidak ada, diperkirakan dari jumlah kata.
- "Artikel jurnal" pada K18 dikenali dari pola `vol.` dan `no./pp.`, jadi entri non-IEEE bisa salah dikenali.
- Kurung siku dianggap sitasi hanya bila isinya angka 1-999 dan tidak terlihat seperti notasi matematika. Sitasi yang
  menempel pada kata tanpa spasi (`kata[1]`) tidak terbaca.
- Crossref membutuhkan internet. Bila gagal, referensi berstatus "Tidak terverifikasi" dan pemeriksaan lain tetap jalan.
  Referensi tanpa DOI dicari lewat pencarian bibliografis, jadi hasilnya perkiraan.
- Belum ada login dan perlindungan CSRF. Hasil hanya dilindungi ID acak dan masa simpan, jadi tambahkan
  pembatasan akses bila dipakai publik.

## Pengujian

```bash
pip install -r requirements-dev.txt
python -m pytest -q                  # semua tes
python -m pytest -q -m "not slow"    # tanpa tes varian yang lama
```

`tools/make_variants.py` membuat varian naskah berisi cacat yang disisipkan terkendali, dan `tools/run_batch.py`
mengukur hasil serta waktu pemeriksaan pada banyak berkas. Keduanya dipakai untuk evaluasi, bukan untuk server.
