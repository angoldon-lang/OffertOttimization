from __future__ import annotations

import hashlib
from dataclasses import dataclass

from .docx_utils import Change, Location

STATUSES = ["Conforme", "Da verificare", "Errore", "Non verificabile", "Non applicabile"]
SEVERITY = {"Errore": 4, "Da verificare": 3, "Non verificabile": 2, "Conforme": 1, "Non applicabile": 0}

CONTROLS = {
    "impaginazione": "Impaginazione Word",
    "template": "Confronto con il template",
    "segnaposto": "Segnaposto e campi non compilati",
    "campi": "Campi richiesti dal profilo",
    "testo": "Refusi e grammatica",
    "coerenza": "Coerenza di denominazioni, date e riferimenti",
    "calcoli": "Calcoli economici",
    "margine": "Margine (quotazione fornitore)",
    "riservatezza": "Riservatezza della copia cliente",
}


@dataclass
class Finding:
    check: str
    title: str
    status: str
    location: Location
    evidence: str = ""
    proposal: str = ""
    change: Change | None = None
    auto_fixed: bool = False
    kind: str = ""            # sostanziale | grafica | ""
    needs_input: bool = False

    @property
    def id(self) -> str:
        raw = f"{self.check}|{self.title}|{self.location.label}|{self.evidence[:80]}"
        return hashlib.sha1(raw.encode()).hexdigest()[:10]


def worst(statuses: list[str]) -> str:
    if not statuses:
        return "Non applicabile"
    return max(statuses, key=lambda s: SEVERITY[s])
