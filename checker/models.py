"""Struktur data bersama antar-modul."""
from dataclasses import dataclass, field
from typing import List, Optional

SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}


@dataclass
class Finding:
    criterion: int
    severity: str                  # error | warning | info
    message: str
    para_index: Optional[int] = None   # indeks paragraf (lihat m2_parser.iter_paragraphs)
    suggestion: str = ""
    source: str = "M3"             # modul penghasil temuan
    excerpt: str = ""              # cuplikan teks lokasi (diisi pipeline)
    comment: bool = True           # False = tidak dijadikan komentar di dokumen

    def to_dict(self):
        return self.__dict__.copy()


@dataclass
class CriterionResult:
    id: int
    code: str
    name: str
    module: str
    default_severity: str = "warning"
    group: str = ""                # kelompok kriteria, mis. "A. Pengaturan dokumen"
    status: str = "pass"           # pass | warn | fail | skip
    findings: List[Finding] = field(default_factory=list)
    note: str = ""                 # ringkasan hasil (mis. "12 halaman")

    def add(self, message, para=None, suggestion="", severity=None,
            source=None, comment=True):
        self.findings.append(Finding(
            criterion=self.id, severity=severity or self.default_severity,
            message=message, para_index=para, suggestion=suggestion,
            source=source or self.module, comment=comment))

    def add_many(self, items, summary, limit=5, suggestion="", severity=None):
        """items = [(para_index, pesan)]; tampilkan `limit` pertama, sisanya diringkas."""
        for idx, msg in items[:limit]:
            self.add(msg, para=idx, suggestion=suggestion, severity=severity)
        rest = len(items) - limit
        if rest > 0:
            self.add(summary.format(n=rest), suggestion=suggestion,
                     severity=severity, comment=False)

    def finalize(self):
        sev = {f.severity for f in self.findings}
        if "error" in sev:
            self.status = "fail"
        elif "warning" in sev:
            self.status = "warn"
        elif self.status != "skip":
            self.status = "pass"
        return self

    def to_dict(self):
        return {"id": self.id, "code": self.code, "name": self.name, "module": self.module, "group": self.group,
                "status": self.status, "note": self.note,
                "findings": [f.to_dict() for f in self.findings]}
