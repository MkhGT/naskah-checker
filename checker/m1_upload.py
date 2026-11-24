"""M1 - validasi berkas unggahan dan pengelolaan masa simpan."""
import re
import uuid
import zipfile
from pathlib import Path

MAX_BYTES = 10 * 1024 * 1024          # 10 MB
MAX_UNCOMPRESSED = 150 * 1024 * 1024  # batas anti zip-bomb
REQUIRED_PARTS = ("[Content_Types].xml", "word/document.xml")


class ValidationError(Exception):
    """Berkas ditolak; pesan aman ditampilkan ke pengguna."""


def validate_file(path, original_name=None, max_bytes=MAX_BYTES):
    path = Path(path)
    name = original_name or path.name
    if not name.lower().endswith(".docx"):
        raise ValidationError("Ekstensi berkas harus .docx (Microsoft Word).")
    size = path.stat().st_size
    if size == 0:
        raise ValidationError("Berkas kosong.")
    if size > max_bytes:
        raise ValidationError(f"Ukuran berkas {size / 1048576:.1f} MB melebihi batas {max_bytes // 1048576} MB.")
    with open(path, "rb") as fh:
        if fh.read(4) != b"PK\x03\x04":
            raise ValidationError("Isi berkas bukan format .docx yang valid (bukan arsip OOXML).")
    try:
        with zipfile.ZipFile(path) as z:
            names = set(z.namelist())
            missing = [p for p in REQUIRED_PARTS if p not in names]
            if missing:
                raise ValidationError("Berkas bukan dokumen Word yang utuh (bagian wajib tidak ada).")
            if sum(i.file_size for i in z.infolist()) > MAX_UNCOMPRESSED:
                raise ValidationError("Isi berkas terlalu besar setelah diekstrak.")
            if any(n.startswith("word/vbaProject") for n in names):
                raise ValidationError("Berkas mengandung makro; gunakan .docx tanpa makro.")
            if z.testzip() is not None:
                raise ValidationError("Berkas rusak (arsip tidak utuh).")
            if find_decimal_measure(z):
                raise ValidationError(MSG_DECIMAL)
    except zipfile.BadZipFile:
        raise ValidationError("Berkas rusak atau bukan arsip .docx yang valid.")
    return True


def save_upload(file_storage, jobs_dir, max_bytes=MAX_BYTES):
    """Simpan unggahan ke folder job baru, validasi, kembalikan (job_id, path, nama_asli)."""
    original = file_storage.filename or "naskah.docx"
    job_id = uuid.uuid4().hex
    job_dir = Path(jobs_dir) / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    dest = job_dir / "asli.docx"
    file_storage.save(dest)
    try:
        validate_file(dest, original, max_bytes)
        from docx import Document   # pastikan python-docx bisa membukanya
        Document(dest)
    except ValidationError:
        _cleanup(job_dir)
        raise
    except Exception:
        _cleanup(job_dir)
        raise ValidationError("Berkas tidak dapat dibaca sebagai dokumen Word.")
    return job_id, dest, original


def _cleanup(job_dir):
    for p in Path(job_dir).glob("*"):
        p.unlink(missing_ok=True)
    Path(job_dir).rmdir()


# --- masa simpan: naskah dan hasil dihapus setelah waktunya habis ---
# Umur job dihitung dari waktu terakhir folder job diubah. Hanya folder bernama ID job
# (32 karakter heksadesimal) yang dihapus.
import shutil
import time

JOB_ID_RE = re.compile(r"^[0-9a-f]{32}$")
DEFAULT_RETENTION_MINUTES = 30.0


def parse_retention_minutes(value, default=DEFAULT_RETENTION_MINUTES):
    """Ambil masa simpan (menit); nilai kosong, salah, atau <= 0 memakai bawaan."""
    try:
        v = float(value)
    except (TypeError, ValueError):
        return default
    return v if v > 0 else default


def job_age_seconds(job_dir, now=None):
    """Umur job dalam detik; None bila folder tidak ada."""
    try:
        mtime = Path(job_dir).stat().st_mtime
    except OSError:
        return None
    return (time.time() if now is None else now) - mtime


def purge_job(jobs_dir, job_id):
    """Hapus folder satu job. True bila berhasil dihapus; ID yang tidak valid diabaikan."""
    if not JOB_ID_RE.match(job_id or ""):
        return False
    job_dir = Path(jobs_dir) / job_id
    if not job_dir.is_dir():
        return False
    shutil.rmtree(job_dir, ignore_errors=True)
    return not job_dir.exists()


def purge_expired(jobs_dir, ttl_seconds, now=None):
    """Hapus semua job yang umurnya lebih dari ttl_seconds; mengembalikan jumlahnya."""
    removed = 0
    base = Path(jobs_dir)
    if not base.is_dir():
        return 0
    for d in base.iterdir():
        if d.is_dir() and JOB_ID_RE.match(d.name):
            age = job_age_seconds(d, now)
            if age is not None and age > ttl_seconds and purge_job(base, d.name):
                removed += 1
    return removed


# --- batasan pengolah dokumen ---
# Sistem diuji untuk .docx buatan Microsoft Word. Ukuran bernilai desimal (mis. w:left="708.66") biasa muncul
# pada ekspor Google Docs dan membuat parser berhenti, jadi berkas seperti itu ditolak di awal.
# Berkas dari aplikasi lain tanpa desimal tetap diperiksa, tetapi laporan memberi peringatan.
MSG_DECIMAL = ("Dokumen ini menyimpan ukuran dalam bilangan desimal, yang lazim pada dokumen hasil ekspor Google Docs atau "
               "aplikasi selain Microsoft Word, sehingga tidak dapat diperiksa oleh sistem. Buka dokumen di Microsoft Word, "
               "simpan ulang sebagai .docx, lalu unggah kembali.")

_XML_PARTS = re.compile(r"^word/(document|styles|numbering|footnotes|endnotes|header\d*|footer\d*)\.xml$")
_VAL_MEASURE_TAGS = {"sz", "szCs", "trHeight"}         # w:val berisi ukuran hanya pada elemen ini
_DEC_ANY = re.compile(r'\bw:[A-Za-z]+="-?\d+\.\d+"')
_TAG = re.compile(r"<w:([A-Za-z]+)\b([^>]*)>")
_DEC_ATTR = re.compile(r'\bw:([A-Za-z]+)="(-?\d+\.\d+)"')


def find_decimal_measure(zf):
    """Cari atribut ukuran w:* bernilai desimal; kembalikan temuan pertama atau None."""
    for name in zf.namelist():
        if not _XML_PARTS.match(name):
            continue
        text = zf.read(name).decode("utf-8", "ignore")
        if not _DEC_ANY.search(text):                  # kebanyakan berkas tidak punya desimal
            continue
        for m in _TAG.finditer(text):
            for a in _DEC_ATTR.finditer(m.group(2)):
                if a.group(1) != "val" or m.group(1) in _VAL_MEASURE_TAGS:
                    return {"bagian": name, "elemen": m.group(1), "atribut": a.group(1), "nilai": a.group(2)}
    return None


_PRODUCER_NAMES = {"word": "Microsoft Word", "libreoffice": "LibreOffice", "google_docs": "Google Docs",
                   "tidak_diketahui": "aplikasi yang tidak diketahui"}


def detect_producer(path):
    """Kenali aplikasi pembuat berkas (word, libreoffice, google_docs, lainnya, tidak_diketahui)."""
    app, gdocs = "", False
    try:
        with zipfile.ZipFile(path) as z:
            names = z.namelist()
            if "docProps/app.xml" in names:
                m = re.search(r"<Application>(.*?)</Application>", z.read("docProps/app.xml").decode("utf-8", "ignore"), re.S)
                app = m.group(1).strip() if m else ""
            gdocs = any(n.lower().startswith("customxml/item") and n.endswith(".xml")
                        and b"gDocsCustomXmlDataStorage" in z.read(n) for n in names)
    except (zipfile.BadZipFile, OSError):
        pass
    low = app.lower()
    if gdocs or "google" in low:
        kat = "google_docs"
    elif "libreoffice" in low or "openoffice" in low:
        kat = "libreoffice"
    elif re.search(r"microsoft.*word", app, re.I):
        kat = "word"
    elif app:
        kat = "lainnya"
    else:
        kat = "tidak_diketahui"
    return {"kategori": kat, "aplikasi": app or None, "nama": app if kat == "lainnya" else _PRODUCER_NAMES[kat],
            "word": kat == "word"}
