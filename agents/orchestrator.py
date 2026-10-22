# -*- coding: utf-8 -*-
"""Orchestrateur : Extracteur -> générateur(s) (Excel / PDF) -> QA -> correction automatique (jusqu'à MAX_TRIES essais)."""

from __future__ import annotations

from dataclasses import dataclass, field

from agents import excel_agent, extractor, pdf_agent, qa_agent
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


def _run_one(kind: str, llm: L.LLM, request: str, context: str, previous: dict | None, history: str, progress=None) -> Deliverable:
    gen, build, check, ext = ((excel_agent.generate_spec, excel_agent.build_xlsx, qa_agent.validate_excel, "xlsx") if kind == "excel"
                              else (pdf_agent.generate_spec, pdf_agent.build_pdf, qa_agent.validate_pdf, "pdf"))
    errors: list[str] = []
    log: list[str] = []
    spec, data = previous, b""
    for attempt in range(1, MAX_TRIES + 1):
        if progress:
            progress(f"{'Excel' if kind == 'excel' else 'PDF'} — essai {attempt}/{MAX_TRIES} : génération…")
        try:
            spec = gen(llm, request, context, spec if errors else previous, errors or None, history)
            data = build(spec)
            if progress:
                progress(f"{'Excel' if kind == 'excel' else 'PDF'} — essai {attempt}/{MAX_TRIES} : contrôle qualité…")
            errors = check(data, spec)
        except L.LLMError:
            raise
        except Exception as exc:  # noqa: BLE001  (spec invalide, rendu impossible…) : renvoyé au modèle pour correction
            errors = [f"Erreur de construction du fichier : {type(exc).__name__}: {exc}"]
        log.append(f"Essai {attempt} : " + ("✔ validé" if not errors else f"{len(errors)} erreur(s) — " + " | ".join(errors[:3])))
        if not errors:
            break
    name = F.safe_filename((spec or {}).get("filename") or (spec or {}).get("title") or "fichier", ext) if spec else f"fichier.{ext}"
    return Deliverable(kind=kind, filename=name, data=data, spec=spec or {}, validated=not errors, qa_log=log, errors=errors, attempts=len(log))


def run(request: str, uploads: list[tuple[str, bytes]], llm: L.LLM, previous: dict | None = None, history: str = "", progress=None) -> Result:
    """previous : {"excel": spec, "pdf": spec} des fichiers déjà générés (pour les modifications)."""
    previous = previous or {}
    ex = extractor.extract(request, uploads)
    res = Result(notes=list(ex.notes))
    kinds = [k for k, w in (("excel", ex.wants_excel), ("pdf", ex.wants_pdf)) if w]
    if not kinds:
        kinds = ["excel" if "excel" in previous else "pdf" if "pdf" in previous else "excel"]
    for kind in kinds:
        res.deliverables.append(_run_one(kind, llm, request, ex.context, previous.get(kind), history, progress))
    return res
