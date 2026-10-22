# -*- coding: utf-8 -*-
"""Agent QA : contrôle chaque fichier avant livraison.

EXCEL  1) structure du JSON  2) contrôle statique de TOUTES les formules (syntaxe, parenthèses, fonctions connues, feuilles et plages)
       3) recalcul LibreOffice (si installé) : aucune cellule en #REF! #VALUE! #NAME? #DIV/0! #N/A #NUM!
PDF    contenu extrait du PDF généré : titre, chaque section, chaque tableau (1re et dernière ligne), aucune page vide.
Chaque erreur est une phrase précise renvoyée au générateur pour correction automatique."""

from __future__ import annotations

import io
import re

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string

from utils import files as F

KNOWN_FUNCS = {
    "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA", "COUNTBLANK", "COUNTIF", "COUNTIFS", "SUMIF", "SUMIFS", "AVERAGEIF", "AVERAGEIFS",
    "IF", "IFS", "IFERROR", "IFNA", "AND", "OR", "NOT", "SWITCH", "ROUND", "ROUNDUP", "ROUNDDOWN", "INT", "MOD", "ABS", "POWER", "SQRT",
    "VLOOKUP", "HLOOKUP", "XLOOKUP", "XMATCH", "LOOKUP", "INDEX", "MATCH", "CHOOSE", "CONCAT", "CONCATENATE", "TEXTJOIN", "TEXT", "VALUE",
    "LEFT", "RIGHT", "MID", "LEN", "UPPER", "LOWER", "PROPER", "TRIM", "SUBSTITUTE", "TODAY", "NOW", "DATE", "YEAR", "MONTH", "DAY",
    "SUMPRODUCT", "PRODUCT", "MEDIAN", "LARGE", "SMALL", "RANK", "MAXIFS", "MINIFS", "ISNUMBER", "ISBLANK", "ISERROR", "FILTER", "UNIQUE", "SORT",
}
ERRORS = ("#REF!", "#VALUE!", "#NAME?", "#DIV/0!", "#N/A", "#NUM!", "#NULL!", "Err:")
MIN_ARGS = {"VLOOKUP": 3, "HLOOKUP": 3, "XLOOKUP": 3, "INDEX": 2, "MATCH": 2, "IF": 2, "SUMIF": 2, "COUNTIF": 2, "IFERROR": 2, "ROUND": 1}
CELL_RE = re.compile(r"(?<![A-Za-z_.])(?:(?:'([^']+)'|([A-Za-z_][\w.]*))!)?\$?([A-Z]{1,3})\$?(\d{1,7})(?::\$?([A-Z]{1,3})\$?(\d{1,7}))?(?![\w(])")


def _strip_strings(f: str) -> str:
    return re.sub(r'"(?:[^"]|"")*"', '""', f)


def _split_args(inner: str) -> list[str]:
    args, depth, cur = [], 0, ""
    for ch in inner:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            args.append(cur)
            cur = ""
        else:
            cur += ch
    args.append(cur)
    return [a for a in args if a.strip() != ""] if inner.strip() else []


def check_formula(sheet: str, addr: str, formula: str, wb) -> list[str]:
    errs = []
    where = f"{sheet}!{addr}"
    f = _strip_strings(formula)
    if f.count("(") != f.count(")"):
        return [f"{where} : parenthèses non équilibrées dans {formula}"]
    if re.search(r"\(\s*\)", f) and not re.search(r"\b(TODAY|NOW)\(\s*\)", f):
        errs.append(f"{where} : appel de fonction sans argument dans {formula}")
    if re.search(r"[+\-*/^&]\s*$", f) or re.search(r"=\s*[*/^&]", f):
        errs.append(f"{where} : formule incomplète {formula}")
    for m in re.finditer(r"(?<![\w.])(?:_xlfn\.)?([A-Za-z][A-Za-z0-9.]*)\(", f):
        fn = m.group(1).upper()
        if fn not in KNOWN_FUNCS:
            errs.append(f"{where} : fonction inconnue « {fn} » dans {formula}")
        elif fn in MIN_ARGS:                                               # nombre d'arguments minimal
            i, depth = m.end(), 1
            j = i
            while j < len(f) and depth:
                depth += (f[j] == "(") - (f[j] == ")")
                j += 1
            if len(_split_args(f[i:j - 1])) < MIN_ARGS[fn]:
                errs.append(f"{where} : {fn} attend au moins {MIN_ARGS[fn]} arguments dans {formula}")
    names = {s.lower(): s for s in wb.sheetnames}
    for m in CELL_RE.finditer(f):
        sh = (m.group(1) or m.group(2) or "").strip()
        if sh and sh.lower() not in names:
            errs.append(f"{where} : la feuille « {sh} » n'existe pas (formule {formula})")
            continue
        ws = wb[names[sh.lower()]] if sh else wb[sheet]
        c1, r1, c2, r2 = m.group(3), int(m.group(4)), m.group(5), m.group(6)
        if column_index_from_string(c1) > 16384 or r1 < 1 or r1 > 1048576:
            errs.append(f"{where} : référence {m.group(0)} hors limites")
            continue
        if c2:
            if r2 and int(r2) < r1 or column_index_from_string(c2) < column_index_from_string(c1):
                errs.append(f"{where} : plage {m.group(0)} à l'envers")
            if int(r2) > ws.max_row + 1 and ws.max_row >= 2:
                errs.append(f"{where} : plage {m.group(0)} dépasse la dernière ligne ({ws.max_row}) de « {ws.title} »")
        elif r1 > ws.max_row + 1 and ws.max_row >= 2:
            errs.append(f"{where} : référence {m.group(0)} sur une ligne inexistante (« {ws.title} » s'arrête à la ligne {ws.max_row})")
    return errs


def validate_excel(data: bytes, spec: dict | None = None) -> list[str]:
    errs: list[str] = []
    try:
        wb = load_workbook(io.BytesIO(data))
    except Exception as exc:  # noqa: BLE001
        return [f"Le fichier Excel est illisible : {type(exc).__name__}: {exc}"]
    n_formulas = 0
    for ws in wb:
        if ws.max_row < 1 or ws.max_column < 1:
            errs.append(f"La feuille « {ws.title} » est vide.")
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("="):
                    n_formulas += 1
                    errs += check_formula(ws.title, c.coordinate, c.value.replace("_xlfn.", ""), wb)
    if spec:
        for sh in spec.get("sheets", []):
            ncol = len(sh.get("columns", []))
            for i, r in enumerate(sh.get("rows", []), 2):
                if len(r) > ncol:
                    errs.append(f"Feuille « {sh.get('name')} » ligne {i} : {len(r)} valeurs pour {ncol} colonnes.")
                    break
    if not errs:                                                           # recalcul réel (LibreOffice), si disponible
        calc = F.recalc_xlsx(data)
        if calc:
            wv = load_workbook(io.BytesIO(calc), data_only=True)
            wf = load_workbook(io.BytesIO(data))
            for ws in wv:
                for row in ws.iter_rows():
                    for c in row:
                        if isinstance(c.value, str) and c.value.startswith(ERRORS):
                            f = str(wf[ws.title][c.coordinate].value)
                            if c.value.startswith(("#NAME?", "Err:5")) and re.search(r"XLOOKUP|XMATCH|FILTER|UNIQUE|SORT", f, re.I):
                                continue                                   # LibreOffice ancien : fonction récente non évaluée
                            errs.append(f"{ws.title}!{c.coordinate} : le calcul renvoie {c.value} (formule {f})")
    return errs[:25]


def validate_pdf(data: bytes, spec: dict) -> list[str]:
    errs: list[str] = []
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            pages = [(p.extract_text() or "") for p in pdf.pages]
    except Exception as exc:  # noqa: BLE001
        return [f"Le PDF est illisible : {type(exc).__name__}: {exc}"]
    if not pages:
        return ["Le PDF n'a aucune page."]
    if not any(p.strip() for p in pages):
        return ["Le PDF est vide (aucun texte extrait)."]
    flat = re.sub(r"\s+", " ", " ".join(pages)).lower()
    norm = lambda s: re.sub(r"\s+", " ", str(s if s is not None else "")).strip().lower()  # noqa: E731
    if norm(spec.get("title")) and norm(spec["title"]) not in flat:
        errs.append("Le titre est absent du PDF.")
    for sec in spec.get("sections", []):
        h = norm(sec.get("heading"))
        if h and h not in flat:
            errs.append(f"La section « {sec.get('heading')} » est absente du PDF.")
        if not (sec.get("paragraphs") or sec.get("bullets") or (sec.get("table") or {}).get("rows")):
            errs.append(f"La section « {sec.get('heading')} » n'a aucun contenu (ni texte, ni puces, ni tableau).")
        t = sec.get("table") or {}
        for lab, row in (("première", (t.get("rows") or [None])[0]), ("dernière", (t.get("rows") or [None])[-1])):
            if row:
                first = next((norm(v) for v in row if v not in (None, "")), "")
                if first and first[:25] not in flat:
                    errs.append(f"Tableau de « {sec.get('heading')} » : la {lab} ligne ({first[:30]}) est absente du PDF (tableau tronqué ?).")
    return errs
