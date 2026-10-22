# -*- coding: utf-8 -*-
"""Agent PDF : le modèle décrit le document (titre, sections, tableaux) en JSON ; le rendu est fait avec ReportLab
(pur Python, aucune bibliothèque système — fonctionne sur Streamlit Cloud) et le thème `templates/pdf/theme.json`."""

from __future__ import annotations

import io
import json
from pathlib import Path
from xml.sax.saxutils import escape

from utils import llm as L

THEME_FILE = Path(__file__).resolve().parent.parent / "templates" / "pdf" / "theme.json"

SCHEMA = {
    "type": "object",
    "properties": {
        "filename": {"type": "string"},
        "title": {"type": "string"},
        "subtitle": {"type": "string"},
        "author": {"type": "string"},
        "footer": {"type": "string"},
        "sections": {"type": "array", "minItems": 1, "items": {
            "type": "object",
            "properties": {
                "heading": {"type": "string"},
                "paragraphs": {"type": "array", "items": {"type": "string"}},
                "bullets": {"type": "array", "items": {"type": "string"}},
                "table": {"type": "object", "properties": {
                    "columns": {"type": "array", "items": {"type": "string"}},
                    "rows": {"type": "array", "items": {"type": "array", "items": {"type": ["string", "number", "null"]}}},
                    "total_row": {"type": "boolean", "description": "true si la dernière ligne est une ligne de total"}},
                    "required": ["columns", "rows"]},
            },
            "required": ["heading"]}},
    },
    "required": ["filename", "title", "sections"],
}

SYSTEM = """Tu es l'agent PDF d'une application de transit / douane. Tu décris un document PDF professionnel en JSON (schéma imposé).
RÈGLES :
- Un titre, éventuellement un sous-titre, puis des sections (heading + paragraphes / puces / tableau).
- Recopie EXACTEMENT les données fournies (montants, références, dates, noms). N'invente rien ; si une information manque, écris-le clairement.
- Les totaux affichés doivent être exacts (recalcule-les à partir des lignes du tableau).
- Texte clair, phrases complètes, pas de balises HTML. Chaque section doit avoir du contenu.
- Si l'utilisateur demande une modification, pars du document précédent (fourni) et ne change que ce qui est demandé."""


def generate_spec(llm: L.LLM, request: str, context: str, previous: dict | None, errors: list[str] | None, history: str = "") -> dict:
    parts = []
    if history:
        parts.append(f"Historique de la conversation :\n{history}")
    if context:
        parts.append(f"Données fournies par l'utilisateur :\n{context}")
    if previous:
        parts.append("Document précédent (JSON) à modifier :\n" + json.dumps(previous, ensure_ascii=False))
    if errors:
        parts.append("Le contrôle qualité a trouvé ces erreurs dans ta version précédente — corrige-les toutes :\n- " + "\n- ".join(errors))
    parts.append(f"Demande : {request}")
    return llm.structured(SYSTEM, "\n\n".join(parts), SCHEMA, "pdf_document")


def load_theme() -> dict:
    try:
        return json.loads(THEME_FILE.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return {"font": "Helvetica", "primary": "#1F3864", "accent": "#3B82F6", "text": "#1B2333", "muted": "#6B7280",
                "table_header_bg": "#1F3864", "table_header_fg": "#FFFFFF", "table_band": "#EEF3FB", "margin_mm": 18}


def _fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float):
        return f"{v:,.2f}".replace(",", " ").replace(".", ",") if abs(v) < 1e15 else str(v)
    return str(v)


def build_pdf(spec: dict) -> bytes:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.lib.units import mm
    from reportlab.platypus import ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    th = load_theme()
    C = lambda k: colors.HexColor(th[k])  # noqa: E731
    base = getSampleStyleSheet()
    font = th.get("font", "Helvetica")
    st_title = ParagraphStyle("t", parent=base["Title"], fontName=f"{font}-Bold", fontSize=22, textColor=C("primary"), alignment=0, spaceAfter=4)
    st_sub = ParagraphStyle("s", parent=base["Normal"], fontName=font, fontSize=11, textColor=C("muted"), spaceAfter=10)
    st_h = ParagraphStyle("h", parent=base["Heading2"], fontName=f"{font}-Bold", fontSize=14, textColor=C("primary"), spaceBefore=12, spaceAfter=6)
    st_p = ParagraphStyle("p", parent=base["BodyText"], fontName=font, fontSize=10, leading=14, textColor=C("text"))
    st_c = ParagraphStyle("c", parent=st_p, fontSize=8.5, leading=10.5)
    st_ch = ParagraphStyle("ch", parent=st_c, fontName=f"{font}-Bold", textColor=C("table_header_fg"))
    margin = th.get("margin_mm", 18) * mm
    footer = str(spec.get("footer") or "")
    title = str(spec.get("title") or "")

    def deco(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 8)
        canvas.setFillColor(C("muted"))
        canvas.drawString(margin, 10 * mm, footer[:110])
        canvas.drawRightString(A4[0] - margin, 10 * mm, f"Page {doc.page}")
        canvas.setStrokeColor(C("accent"))
        canvas.setLineWidth(1.2)
        canvas.line(margin, A4[1] - 12 * mm, A4[0] - margin, A4[1] - 12 * mm)
        canvas.restoreState()

    story = [Paragraph(escape(title), st_title)]
    if spec.get("subtitle"):
        story.append(Paragraph(escape(str(spec["subtitle"])), st_sub))
    for sec in spec.get("sections", []):
        story.append(Paragraph(escape(str(sec.get("heading", ""))), st_h))
        for p in sec.get("paragraphs") or []:
            story += [Paragraph(escape(str(p)), st_p), Spacer(1, 4)]
        if sec.get("bullets"):
            story.append(ListFlowable([ListItem(Paragraph(escape(str(b)), st_p)) for b in sec["bullets"]], bulletType="bullet", leftIndent=14))
            story.append(Spacer(1, 4))
        t = sec.get("table")
        if t and t.get("columns"):
            ncol = len(t["columns"])
            data = [[Paragraph(escape(str(c)), st_ch) for c in t["columns"]]]
            for r in t.get("rows") or []:
                data.append([Paragraph(escape(_fmt(r[j]) if j < len(r) else ""), st_c) for j in range(ncol)])
            tb = Table(data, repeatRows=1, colWidths=[(A4[0] - 2 * margin) / ncol] * ncol)
            style = [("BACKGROUND", (0, 0), (-1, 0), C("table_header_bg")), ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#B7C2D6")),
                     ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                     ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, C("table_band")])]
            if t.get("total_row") and len(data) > 1:
                style += [("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#D9E1F2")), ("LINEABOVE", (0, -1), (-1, -1), 1, C("primary"))]
            tb.setStyle(TableStyle(style))
            story += [tb, Spacer(1, 8)]
    bio = io.BytesIO()
    doc = SimpleDocTemplate(bio, pagesize=A4, leftMargin=margin, rightMargin=margin, topMargin=margin + 4 * mm, bottomMargin=margin,
                            title=title, author=str(spec.get("author") or ""))
    doc.build(story, onFirstPage=deco, onLaterPages=deco)
    return bio.getvalue()
