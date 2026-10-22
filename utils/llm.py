# -*- coding: utf-8 -*-
"""Appel LLM (Anthropic) avec sortie structurée forcée (tool use) : le modèle renvoie toujours un JSON conforme au schéma."""

from __future__ import annotations

import json
from typing import Callable


class LLMError(Exception):
    pass


class LLM:
    """`cfg` : ai_engine.AIConfig (clé + modèle). `fn` : fonction de remplacement (tests / hors ligne) :
    fn(system, user, schema, tool_name) -> dict."""

    def __init__(self, cfg=None, fn: Callable | None = None):
        self.cfg, self.fn = cfg, fn

    @property
    def available(self) -> bool:
        return self.fn is not None or bool(self.cfg and getattr(self.cfg, "api_key", ""))

    def structured(self, system: str, user: str, schema: dict, tool_name: str = "emit", max_tokens: int = 16000) -> dict:
        if self.fn is not None:
            return self.fn(system, user, schema, tool_name)
        if not self.available:
            raise LLMError("Clé API manquante : renseignez-la dans la barre latérale (ou ANTHROPIC_API_KEY).")
        import ai_engine
        client = ai_engine.make_client(self.cfg)
        kwargs = dict(model=self.cfg.model, max_tokens=max_tokens, system=system,
                      messages=[{"role": "user", "content": user}],
                      tools=[{"name": tool_name, "description": "Renvoie le résultat structuré.", "input_schema": schema}],
                      tool_choice={"type": "tool", "name": tool_name})
        try:
            with client.messages.stream(**kwargs) as stream:
                msg = stream.get_final_message()
        except Exception as exc:  # noqa: BLE001
            raise LLMError(ai_engine.explain_error(exc) if hasattr(ai_engine, "explain_error") else str(exc)) from exc
        for block in msg.content:
            if getattr(block, "type", "") == "tool_use":
                data = block.input
                if isinstance(data, str):
                    data = json.loads(data)
                return data
        raise LLMError("Le modèle n'a renvoyé aucun résultat structuré.")
