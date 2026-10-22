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


def _pdf_tables(pdfs: list[tuple[str, bytes]]) -> list[dict]:
    """Lecture hors ligne d'un PDF de BL / facture avec le moteur du site (règles, sans IA) -> tables « BL References » et « Factures »."""
    try:
        import core_export
        import export_engine as ee

        class _F:
            def __init__(s, n, d): s.name, s._d = n, d
            def getvalue(s): return s._d

        class _P:
            def progress(s, *a, **k): pass
        r = ee.run_export_pipeline([_F(n, d) for n, d in pdfs], core_export.OCREngine(), "client", _P())
    except Exception:  # noqa: BLE001
        return []
    out = []
    bl, fa = r.get("blref"), r.get("factures")
    if bl is not None and len(bl):
        b = bl.drop(columns=[c for c in ("_group", "_currency") if c in bl.columns])
        out.append({"name": "BL References", "columns": [str(c) for c in b.columns],
                    "rows": [[None if v != v else v for v in row] for row in b.itertuples(index=False)]})
    if fa is not None and len(fa):
        out.append({"name": "Factures", "columns": [str(c) for c in fa.columns],
                    "rows": [[None if v != v else (v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else v) for v in row] for row in fa.itertuples(index=False)]})
    return out


def extract(request: str, uploads: list[tuple[str, bytes]], offline: bool = False) -> Extraction:
    ex = Extraction()
    for name, data in uploads:
        f = F.read_upload(name, data)
        if f["text"].startswith("[Fichier illisible"):
            ex.notes.append(f"« {name} » : {f['text']}")
        ex.files.append(f)
        if f["truncated"]:
            ex.notes.append(f"« {name} » est long : seules les {F.MAX_ROWS} premières lignes / {F.MAX_TEXT:,} premiers caractères sont lus.")
    pdfs = [(n, d) for n, d in uploads if n.lower().endswith(".pdf")]
    if offline and pdfs:                                              # tous les PDF ensemble : BL et factures se rapprochent entre eux
        tabs = _pdf_tables(pdfs)
        if tabs:
            first = next(f for f in ex.files if f["name"].lower().endswith(".pdf"))
            first["tables"], first["kind"] = tabs, "table"
            ex.notes.append(f"{len(pdfs)} PDF lu(s) ensemble par le moteur local du site (BL / factures) : "
                            + ", ".join(f"{t['name']} ({len(t['rows'])} lignes)" for t in tabs) + ".")
    ex.context = F.uploads_to_context(ex.files)
    ex.wants_excel = bool(EXCEL_WORDS.search(request))
    ex.wants_pdf = bool(re.search(r"\bpdf\b|rapport|report|attestation|courrier|lettre|fiche|devis", request, re.I))
    return ex                                                         # rien de précisé : l'orchestrateur choisit (fichier précédent, sinon Excel si tableau)
