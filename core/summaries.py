"""Utilitários para blocos de resumo, recomendações e próximos passos."""

from __future__ import annotations


def build_action_bullets(diagnostico: str | None = None, risco: str | None = None, economia: str | None = None, proximo_passo: str | None = None) -> list[str]:
    bullets = []
    if diagnostico:
        bullets.append(f"💡 Diagnóstico: {diagnostico}")
    if risco:
        bullets.append(f"⚠️ Risco: {risco}")
    if economia:
        bullets.append(f"💰 Economia: {economia}")
    if proximo_passo:
        bullets.append(f"✅ Próximo passo: {proximo_passo}")
    return bullets
