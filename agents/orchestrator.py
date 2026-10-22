# -*- coding: utf-8 -*-
"""Orchestrateur : Extracteur -> générateur(s) (Excel / PDF) -> QA -> correction automatique (jusqu'à MAX_TRIES essais)."""

from __future__ import annotations

from dataclasses import dataclass, field

from agents import excel_agent, extractor, local_engine, pdf_agent, qa_agent
from utils import files as F
from utils import llm as L

MAX_TRIES = 3


@dataclass
class Deliverable:
    kind: str                       # "excel" | "pdf"
    filename: str
    data: bytes
    spec: dict
    validated: bool
    qa_log: list[str] = field(default_factory=list)      # une ligne par essai
    errors: list[str] = field(default_factory=list)      # erreurs restantes si non validé
    attempts: int = 1


@dataclass
class Result:
    deliverables: list[Deliverable] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)


def _run_one(kind: str, llm: L.LLM | None, request: str, ex, previous: dict | None, history: str, progress=None) -> Deliverable:
    local = llm is None or not llm.available
    build, check, ext = ((excel_agent.build_xlsx, qa_agent.validate_excel, "xlsx") if kind == "excel" else (pdf_agent.build_pdf, qa_agent.validate_pdf, "pdf"))
    label = "Excel" if kind == "excel" else "PDF"
    errors: list[str] = []
    log: list[str] = []
    spec, data = previous, b""
    for attempt in range(1, (1 if local else MAX_TRIES) + 1):
        if progress:
            progress(f"{label} — essai {attempt}: génération{' (moteur local)' if local else ''}…")
        try:
            if local:
                spec = (local_engine.excel_spec if kind == "excel" else local_engine.pdf_spec)(request, ex, previous)
            else:
                gen = excel_agent.generate_spec if kind == "excel" else pdf_agent.generate_spec
                spec = gen(llm, request, ex.context, spec if errors else previous, errors or None, history)
            data = build(spec)
            if progress:
                progress(f"{label} — essai {attempt}: contrôle qualité…")
            errors = check(data, spec)
        except L.LLMError:
            raise
        except Exception as exc:  # noqa: BLE001  (spec invalide, rendu impossible, donnée manquante…) : renvoyé au modèle pour correction
            errors = [str(exc) if isinstance(exc, ValueError) else f"Erreur de construction du fichier : {type(exc).__name__}: {exc}"]
        log.append(f"Essai {attempt}{' (moteur local)' if local else ''} : " + ("✔ validé" if not errors else f"{len(errors)} erreur(s) — " + " | ".join(errors[:3])))
        if not errors:
            break
    name = F.safe_filename((spec or {}).get("filename") or (spec or {}).get("title") or "fichier", ext) if spec else f"fichier.{ext}"
    return Deliverable(kind=kind, filename=name, data=data, spec=spec or {}, validated=not errors and bool(data), qa_log=log, errors=errors, attempts=len(log))


def run(request: str, uploads: list[tuple[str, bytes]], llm: L.LLM | None, previous: dict | None = None, history: str = "", progress=None) -> Result:
    """previous : {"excel": spec, "pdf": spec} des fichiers déjà générés (pour les modifications)."""
    previous = previous or {}
    local = llm is None or not llm.available
    ex = extractor.extract(request, uploads, offline=local)
    res = Result(notes=list(ex.notes))
    kinds = [k for k, w in (("excel", ex.wants_excel), ("pdf", ex.wants_pdf)) if w]
    if not kinds:                                                      # rien de précisé : même type(s) que le fichier précédent, sinon selon les données
        kinds = [k for k in ("excel", "pdf") if k in previous] or (["excel"] if any(f["tables"] for f in ex.files) else ["pdf"])
    for kind in kinds:
        res.deliverables.append(_run_one(kind, llm, request, ex, previous.get(kind), history, progress))
    if local:
        res.notes.append("Moteur local (sans clé API) : le fichier est construit à partir de vos données déposées, sans rien inventer.")
    return res
