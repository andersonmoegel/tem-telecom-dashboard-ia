"""Componentes reutilizáveis para sidebar."""

from __future__ import annotations

import html
import streamlit as st


def _safe(value: object) -> str:
    return html.escape(str(value if value is not None else ""))


def status_card(operadoras: int | str = "-", contratos: int | str = "-", status: str = "Base ativa") -> None:
    """Card de status do topo da sidebar."""
    st.sidebar.markdown(
        f"""
        <div class='ds-status-card'>
            <div class='ds-status-title'>Status do ambiente</div>
            <div class='ds-status-line'><span>Operadoras</span><strong>{_safe(operadoras)}</strong></div>
            <div class='ds-status-line'><span>Contratos</span><strong>{_safe(contratos)}</strong></div>
            <div class='ds-status-line'><span>Condição</span><strong>{_safe(status)}</strong></div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def group_label(label: str) -> None:
    st.sidebar.markdown(f"<div class='ds-sidebar-group'>{_safe(label)}</div>", unsafe_allow_html=True)
