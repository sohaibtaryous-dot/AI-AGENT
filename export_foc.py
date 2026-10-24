# -*- coding: utf-8 -*-
"""EXPORT — documents « FOC » d'Alstom Railway Maroc (envois gratuits / remboursements) :

  • CUSTOMS INVOICE/FOC : « Invoice Number FCM3625 … BOL FCAB3625 », client = « Shipping Address »,
                          tableau « Site · DESCRIPTION · Qty · PRIX · EXTENDED PRICE », « Total Value 40,00 € ».
  • Delivery Note       : « Shipper ID FCAB3625 », tableau « Master N° · Site · DESCRIPTION · Qty · POSTE »,
                          bloc « Packing-List » : N°-Packaging (= colis), Net Weight, Gross Weight.

Ces pages ne ressemblent pas aux BL / factures habituels : sans traitement dédié, le bon de livraison était collé au BL précédent
et la facture restait sans client ni BL. Ici chaque page devient son propre document et ses champs sont lus directement."""

from __future__ import annotations

import re

DN_RE = re.compile(r"Delivery\s+Note\s+ID", re.I)
SHIPPER_RE = re.compile(r"Shipper\s+ID\s+([A-Z0-9][A-Z0-9\-/_]*\d[A-Z0-9\-/_]*)", re.I)
INV_RE = re.compile(r"CUSTOMS\s+INVOICE\s*/\s*FOC", re.I)
INV_NO_RE = re.compile(r"Invoice\s+Number\s+([A-Z0-9][A-Z0-9\-/_]*\d[A-Z0-9\-/_]*)", re.I)
BOL_RE = re.compile(r"\bBOL\s+([A-Z0-9][A-Z0-9\-/_]*\d[A-Z0-9\-/_]*)", re.I)
DATE_RE = re.compile(r"Ship\s+Date\s+(\d{1,2})/(\d{1,2})/(\d{4})", re.I)
NUM = r"\d[\d\s.,]*"


def foc_type(text: str) -> str:
    t = text or ""
    if DN_RE.search(t) and SHIPPER_RE.search(t):
        return "DN"
    if INV_RE.search(t) or (re.search(r"Invoice-ID", t) and BOL_RE.search(t) and INV_NO_RE.search(t)):
        return "INV"
    return ""


def split_documents(docs, Document):
    """Chaque page FOC devient un document à part (les pages suivantes « AUTRE » restent avec elle).
    Le reste du document d'origine garde le même objet (identifiant conservé pour la lecture IA)."""
    out = []
    for d in docs:
        groups: list[list] = []                                       # [type FOC ou "", [pages]]
        for p in d.pages:
            t = foc_type(p.text)
            if t:
                groups.append([t, [p]])
            elif groups and (getattr(p, "kind", "") == "AUTRE" or groups[-1][0] == ""):
                groups[-1][1].append(p)
            else:
                groups.append(["", [p]])
        if not any(t for t, _ in groups):
            out.append(d)
            continue
        reused = False
        for t, pages in groups:
            if not t and not reused:
                d.pages = pages
                out.append(d)
                reused = True
                continue
            nd = Document("PACKING" if t == "DN" else ("FACTURE" if t == "INV" else d.kind), d.file, pages)
            nd._foc = t
            out.append(nd)
    return out


def _num(s, parse_number, dec=","):
    s = re.sub(r"[^\d.,]", "", str(s or ""))
    if not s:
        return None
    if "," in s and "." not in s and re.fullmatch(r"\d{1,3}(?:,\d{3})+", s):
        return float(s.replace(",", ""))
    return parse_number(s, "," if "," in s else ".")


def _client(text: str) -> str:
    m = re.search(r"Shipping\s+Address[^\n]*\n(.*?)(?:\n\s*Details\b)", text, re.S | re.I)
    if not m:
        return ""
    for ln in m.group(1).split("\n"):
        s = ln.strip()
        if s and not re.match(r"(?i)(emilie|to the attention|attention)", s) and re.search(r"[A-Za-z]{3}", s) \
                and not re.search(r"(?i)plichard|dordain", s):
            return s
    return ""


def parse_dn(text: str, parse_number) -> dict:
    t = text or ""
    out = {"number": (SHIPPER_RE.search(t) or [None, ""])[1] if SHIPPER_RE.search(t) else "", "lines": [],
           "packages": None, "net": None, "gross": None}
    m = re.search(r"N[°º]?\s*-?\s*Packaging\s+(\d+)", t, re.I)
    if m:
        out["packages"] = int(m.group(1))
    m = re.search(r"Net\s+Weight\s+(" + NUM + r")", t, re.I)
    if m:
        out["net"] = _num(m.group(1).split("\n")[0], parse_number)
    m = re.search(r"Gross\s+Weight\s+(" + NUM + r")", t, re.I)
    if m:
        out["gross"] = _num(m.group(1).split("\n")[0], parse_number)
    blk = re.search(r"Details\s+D-?N(.*?)(?:Packing-?List|$)", t, re.S | re.I)
    for ln in (blk.group(1).split("\n") if blk else []):
        mm = re.match(r"^\s*(\d+)\s+([A-Z0-9]{3,})\s+(.+?)\s+(\d+(?:[.,]\d+)?)\s+(\d+)\b", ln.strip())
        if mm:
            out["lines"].append({"pos": mm.group(5), "ref": "", "desc": mm.group(3).strip(), "qty": _num(mm.group(4), parse_number),
                                 "unit": "", "po": "", "_master": mm.group(1), "_site": mm.group(2)})
    return out


def parse_invoice(text: str, parse_number) -> dict:
    t = text or ""
    m = INV_NO_RE.search(t)
    b = BOL_RE.search(t)
    dm = DATE_RE.search(t)
    out = {"number": m.group(1) if m else "", "bl_refs": [b.group(1)] if b else [], "client": _client(t), "lines": [], "total": None,
           "currency": "EUR" if "€" in t or re.search(r"\bEUR\b", t) else "", "date": None}
    if dm:
        from datetime import date
        try:
            out["date"] = date(int(dm.group(3)), int(dm.group(2)), int(dm.group(1)))
        except ValueError:
            pass
    blk = re.search(r"Details\s+Invoice(.*?)(?:Total\s+Value|$)", t, re.S | re.I)
    for ln in (blk.group(1).split("\n") if blk else []):
        mm = re.match(r"^\s*([A-Z0-9]{3,})\s+(.+?)\s+(\d+(?:[.,]\d+)?)\s+(\d[\d.,]*)\s+(\d[\d.,]*)\s*€?\s*$", ln.strip())
        if mm and not re.match(r"(?i)site\b", ln.strip()):
            qty, price, amount = (_num(mm.group(k), parse_number) for k in (3, 4, 5))
            out["lines"].append({"pos": "", "ref": "", "desc": mm.group(2).strip(), "qty": qty, "unit": "", "amount": amount,
                                 "price": price, "price_per": 1, "po": "", "_used": False, "_site": mm.group(1)})
    m = re.search(r"Total\s+Value\s+(" + NUM + r")", t, re.I)
    if m:
        out["total"] = _num(m.group(1).split("\n")[0], parse_number)
    return out
