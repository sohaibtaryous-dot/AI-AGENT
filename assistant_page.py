# -*- coding: utf-8 -*-
"""Module ASSISTANT — conversation pour générer / modifier des fichiers Excel et PDF (agents IA + contrôle qualité)."""

from __future__ import annotations

import io

import pandas as pd
import streamlit as st

import ui
from agents import orchestrator
from utils import files as F
from utils import llm as L

K_MSGS, K_DELIV, K_SPECS, K_UP = "asst_msgs", "asst_deliv", "asst_specs", "asst_uploads"
EXAMPLES = ["Fais un tableau Excel des factures avec total et TVA 20 % en formules.",
            "Génère un rapport PDF de ce dossier avec un tableau récapitulatif.",
            "Ajoute une colonne « Écart » et colore l'en-tête en vert."]


def _history(msgs: list[dict], n: int = 6) -> str:
    return "\n".join(f"{'Utilisateur' if m['role'] == 'user' else 'Assistant'} : {m['content']}" for m in msgs[-n:])


def _preview_excel(d: orchestrator.Deliverable):
    calc = F.recalc_xlsx(d.data) if F.libreoffice_path() else None
    src = calc or d.data
    xl = pd.ExcelFile(io.BytesIO(src))
    tabs = st.tabs(xl.sheet_names[:8])
    for tab, name in zip(tabs, xl.sheet_names[:8]):
        with tab:
            st.dataframe(xl.parse(name, dtype=object).fillna("").astype(str), use_container_width=True, hide_index=True)
    if not calc:
        st.caption("Aperçu : valeurs saisies ; les formules se calculent à l'ouverture dans Excel.")


def _preview_pdf(d: orchestrator.Deliverable):
    png = F.pdf_first_page_png(d.data)
    if png:
        st.image(png, caption="Aperçu — page 1", use_container_width=True)


def _render_deliverable(d: orchestrator.Deliverable, i: int):
    icon = ":material/table_view:" if d.kind == "excel" else ":material/picture_as_pdf:"
    with st.container(border=True):
        st.markdown(f"**{d.filename}**")
        if d.validated:
            st.success(f"Contrôle qualité réussi ({d.attempts} essai{'s' if d.attempts > 1 else ''}).", icon=":material/verified:")
            (_preview_excel if d.kind == "excel" else _preview_pdf)(d)
            st.download_button(f"Télécharger {d.filename}", d.data, file_name=d.filename, icon=icon, type="primary", key=f"asst_dl_{i}",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" if d.kind == "excel" else "application/pdf",
                               use_container_width=True)
        else:
            if d.data:
                st.error("Le fichier n'a pas passé le contrôle qualité : il n'est pas proposé au téléchargement.", icon=":material/gpp_bad:")
            else:
                st.warning("Fichier non généré.", icon=":material/info:")
            for e in d.errors[:8]:
                st.markdown(f"- {e}")
            st.caption("Reformulez votre demande ou précisez les données, puis réessayez.")
        with st.expander("Journal du contrôle qualité"):
            for line in d.qa_log:
                st.markdown(f"- {line}")


def render(ai_cfg):
    ui.hero_3d({"agents": 4, "formats": "Excel · PDF", "contrôle": "QA auto", "essais": orchestrator.MAX_TRIES},
               badge="TIMAR TANGER · ASSISTANT", title="Assistant fichiers", highlight="Excel & PDF",
               desc="Décrivez le fichier voulu (ou déposez vos données) : les agents extraient, génèrent, puis un contrôle qualité "
                    "corrige automatiquement avant le téléchargement.", ai_label="")
    has_key = L.LLM(ai_cfg).available
    use_claude = has_key and st.toggle("Utiliser Claude (clé API détectée) — sinon moteur local", value=True, key="asst_use_claude",
                                       help="Claude comprend n'importe quelle demande ; le moteur local (sans clé) couvre les demandes courantes sur vos données déposées.")
    llm = L.LLM(ai_cfg) if use_claude else None
    if not use_claude:
        st.info("**Moteur local (sans clé API)** : déposez vos données (Excel, CSV, ou PDF de BL / factures) puis demandez, par exemple : "
                "« tableau Excel avec total et TVA 20 % », « trié par Montant, récapitulatif par Client, en-tête vert », "
                "« rapport PDF », puis « ajoute une colonne écart », « supprime la colonne Date », « renomme X en Y ».", icon=":material/smart_toy:")
    st.session_state.setdefault(K_MSGS, [])
    st.session_state.setdefault(K_DELIV, [])
    st.session_state.setdefault(K_SPECS, {})
    left, right = st.columns([1.05, 1], gap="large")
    with left:
        up = st.file_uploader("Données (optionnel) : Excel, CSV, PDF, Word, texte", accept_multiple_files=True, key="asst_files",
                              type=["xlsx", "xlsm", "xls", "csv", "pdf", "docx", "txt"])
        box = st.container(height=470, border=True)
        with box:
            if not st.session_state[K_MSGS]:
                st.caption("Exemples de demandes :")
                for e in EXAMPLES:
                    st.markdown(f"- {e}")
            for m in st.session_state[K_MSGS]:
                with st.chat_message(m["role"]):
                    st.markdown(m["content"])
        prompt = st.chat_input("Décrivez le fichier à créer ou la modification à faire…")
        if st.session_state[K_MSGS] and st.button("Nouvelle conversation", icon=":material/restart_alt:", key="asst_reset"):
            for k in (K_MSGS, K_DELIV, K_SPECS):
                st.session_state.pop(k, None)
            st.rerun()
    if prompt:
        st.session_state[K_MSGS].append({"role": "user", "content": prompt})
        uploads = [(f.name, f.getvalue()) for f in (up or [])]
        sig = tuple((n, len(d)) for n, d in uploads)
        same_files = bool(st.session_state[K_SPECS]) and sig == st.session_state.get("asst_up_sig")
        st.session_state["asst_up_sig"] = sig
        if same_files and llm is None:                         # mêmes fichiers qu'avant : le moteur local MODIFIE le fichier précédent
            uploads = []
        status = left.empty()
        try:
            res = orchestrator.run(prompt, uploads, llm, previous=st.session_state[K_SPECS], history=_history(st.session_state[K_MSGS][:-1]),
                                   progress=lambda t: status.info(t, icon=":material/smart_toy:"))
            status.empty()
            st.session_state[K_DELIV] = res.deliverables
            for d in res.deliverables:
                if d.validated:
                    st.session_state[K_SPECS][d.kind] = d.spec
            ok = [d for d in res.deliverables if d.validated]
            msg = (f"Fichier(s) prêt(s) : {', '.join(d.filename for d in ok)}." if ok else "Je n'ai pas réussi à produire un fichier valide.")
            for d in res.deliverables:
                if not d.validated and d.errors:
                    msg += f"\n\n**{d.filename}** : {d.errors[0]}"
            if res.notes:
                msg += "\n\n" + "\n".join(f"⚠ {n}" for n in res.notes)
            st.session_state[K_MSGS].append({"role": "assistant", "content": msg})
        except L.LLMError as exc:
            status.empty()
            st.session_state[K_MSGS].append({"role": "assistant", "content": f"Erreur : {exc}"})
        st.rerun()
    with right:
        ui.section("Aperçu & téléchargement", "download")
        dels = st.session_state[K_DELIV]
        if not dels:
            st.info("Le fichier généré apparaîtra ici avec son aperçu ; le bouton de téléchargement n'apparaît qu'après validation.")
        for i, d in enumerate(dels):
            _render_deliverable(d, i)
