# -*- coding: utf-8 -*-
"""Agent « Data Extractor » : transforme les fichiers déposés + la demande en contexte exact pour les générateurs.
Les fichiers tabulaires sont lus directement (aucun risque d'invention) ; le texte (PDF, Word) est transmis tel quel."""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from utils import files as F


@dataclass
class Extraction:
    files: list[dict] = field(default_factory=list)
    context: str = ""
    wants_excel: bool = False
    wants_pdf: bool = False
    notes: list[str] = field(default_factory=list)


EXCEL_WORDS = re.compile(r"excel|xlsx|tableur|feuille|classeur|spreadsheet|formule|tableau", re.I)
PDF_WORDS = re.compile(r"\bpdf\b|rapport|report|document|facture|devis|attestation|courrier|lettre|fiche", re.I)


def extract(request: str, uploads: list[tuple[str, bytes]]) -> Extraction:
    ex = Extraction()
    for name, data in uploads:
        f = F.read_upload(name, data)
        ex.files.append(f)
        if f["truncated"]:
            ex.notes.append(f"« {name} » est long : seules les {F.MAX_ROWS} premières lignes / {F.MAX_TEXT:,} premiers caractères sont lus.")
        if f["text"].startswith("[Fichier illisible"):
            ex.notes.append(f"« {name} » : {f['text']}")
    ex.context = F.uploads_to_context(ex.files)
    ex.wants_excel = bool(EXCEL_WORDS.search(request))
    ex.wants_pdf = bool(re.search(r"\bpdf\b|rapport|report|attestation|courrier|lettre|fiche|devis", request, re.I))
    if not ex.wants_excel and not ex.wants_pdf:                       # rien de précisé : selon les données déposées
        has_tables = any(f["tables"] for f in ex.files)
        ex.wants_excel, ex.wants_pdf = has_tables, not has_tables
    return ex
