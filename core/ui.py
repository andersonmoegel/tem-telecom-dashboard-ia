"""Componentes visuais reutilizáveis para páginas Streamlit."""

from __future__ import annotations

import html
from typing import Iterable
import streamlit as st


def _safe(value: object) -> str:
    return html.escape(str(value if value is not None else ""))


def page_header(title: str, subtitle: str | None = None) -> None:
    """Cabeçalho padronizado de página."""
    st.markdown(f"<div class='ds-page-title'>{_safe(title)}</div>", unsafe_allow_html=True)
    if subtitle:
        st.markdown(f"<div class='ds-page-subtitle'>{_safe(subtitle)}</div>", unsafe_allow_html=True)


def criar_card_metrica(titulo: str, valor: object, delta: object | None = None, status: str = "info") -> None:
    """Card KPI customizado, substituto visual do st.metric.

    status: info | success | warning | danger
    """
    delta_html = ""
    if delta not in (None, ""):
        delta_class = "neutral"
        text_delta = str(delta)
        if text_delta.strip().startswith("-") or "↓" in text_delta:
            delta_class = "negative"
        elif text_delta.strip().startswith("+") or "↑" in text_delta or "▲" in text_delta:
            delta_class = ""
        delta_html = f"<div class='ds-metric-delta {delta_class}'>{_safe(delta)}</div>"

    st.markdown(
        f"""
        <div class='ds-metric-card {_safe(status)}'>
            <div class='ds-metric-label'>{_safe(titulo)}</div>
            <div class='ds-metric-value'>{_safe(valor)}</div>
            {delta_html}
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_grid(items: Iterable[dict], columns: int = 4) -> None:
    """Renderiza uma linha/grid de cards KPI."""
    items = list(items)
    if not items:
        return
    cols = st.columns(min(columns, len(items)))
    for idx, item in enumerate(items):
        with cols[idx % len(cols)]:
            criar_card_metrica(
                item.get("titulo", "Indicador"),
                item.get("valor", "-"),
                item.get("delta"),
                item.get("status", "info"),
            )


def filter_chips(filters: dict) -> None:
    """Exibe filtros ativos como chips."""
    chips = []
    for key, value in (filters or {}).items():
        if value in (None, "", [], "Todos", "Todas"):
            continue
        label = ", ".join(map(str, value)) if isinstance(value, (list, tuple, set)) else str(value)
        chips.append(f"<span class='ds-filter-chip'>{_safe(key)}: <strong>{_safe(label)}</strong></span>")
    if chips:
        st.markdown("".join(chips), unsafe_allow_html=True)


def resumo_proativo(score: object, bullets: list[str], titulo: str = "Resumo proativo") -> None:
    """Bloco padrão para score + lista de achados/ações."""
    left, right = st.columns([1, 3])
    with left:
        criar_card_metrica("Score", score, status="info")
    with right:
        bullet_html = "".join(f"<li>{_safe(item)}</li>" for item in bullets)
        st.markdown(
            f"""
            <div class='ds-proactive-card'>
                <div class='ds-metric-label'>{_safe(titulo)}</div>
                <ul>{bullet_html}</ul>
            </div>
            """,
            unsafe_allow_html=True,
        )
