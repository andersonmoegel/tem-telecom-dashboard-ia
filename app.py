
from __future__ import annotations

import os
import re
import json
from pathlib import Path
import inspect
import html
import hashlib
import difflib
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

try:
    from streamlit_plotly_events import plotly_events
except Exception:
    plotly_events = None

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None

try:
    from google import genai
except Exception:
    genai = None

APP_TITLE = "TEM Telecom FinOps"
BASE_DIR = Path(__file__).resolve().parent
UPLOADS_DIR = BASE_DIR / "uploads"
UPLOADS_DIR.mkdir(exist_ok=True)
ACTION_HISTORY_FILE = BASE_DIR / "dados_atualizados" / "historico_acoes.json"
DEFAULT_FILES = [UPLOADS_DIR / "TelecomDB_v1_0.xlsx", BASE_DIR / "TelecomDB_exemplo.xlsx"]

VALUE_COL = "Valor_Realizado"
DATE_COL = "Periodo_Data"
MONTH_COL = "Periodo"
REF_DATE_COL = "Periodo_Data_Referencia"
DUE_DATE_COL = "Periodo_Data_Vencimento"
REF_MONTH_COL = "Periodo_Referencia"
DUE_MONTH_COL = "Periodo_Vencimento"
PERIOD_BASIS_KEY = "period_basis"
CATEGORY_COL = "Categoria"
SUPPLIER_COL = "Fornecedor"
SERVICE_COL = "Servico"
CONTRACT_COL = "Contrato"
BRANCH_COL = "Filial"
REGION_COL = "Regiao"
CC_COL = "Centro_Custo"
CC_ID_COL = "Centro_Custo_ID"
STATUS_COL = "Status_Contrato"
INVOICE_COL = "Fatura"
DIFF_COL = "Diferenca"
CONTESTED_COL = "Contestado"

# Aliases aceitos para pequenas mudanças futuras na estrutura das planilhas.
# A base continua preferindo os nomes originais, mas tenta reconhecer variações comuns.
COLUMN_ALIASES = {
    VALUE_COL: ["Financeiro_Servico_Valor", "Valor_Realizado", "Valor Realizado", "Valor Pago", "Valor_Pago", "Valor", "Valor Total", "Valor_Total", "Total", "Total Pago", "Total_Pago", "Custo", "Despesa", "Amount", "Price", "valor"],
    DIFF_COL: ["Financeiro_Calc_diffValor", "Diferenca", "Diferença", "Diff", "Divergencia", "Divergência"],
    DATE_COL: ["Financeiro_Vencimento", "Financeiro_Data_Leitura", "Data_Vencimento", "Vencimento", "Data", "Periodo", "Período", "Mes", "Mês", "Financeiro_mesReferencia", "Referencia", "Referência"],
    SUPPLIER_COL: ["Fornecedor", "Operadora", "Operadoras", "Carrier", "Prestador", "Prestadora", "Empresa", "Vendor", "Fornecedor_Nome", "Fornecedor_Nome_Fantasia", "Fornecedor_Razao_Social", "Nome Fornecedor", "Nome_Fornecedor", "Razao Social", "Razão Social", "Nome Fantasia"],
    SERVICE_COL: ["Servico", "Serviço", "Servicos", "Serviços", "Produto", "Item", "Plano", "Linha", "Servico_Descricao", "Serviço Descrição", "Descricao Servico", "Descrição Serviço", "Descrição", "Descricao"],
    CATEGORY_COL: ["Categoria", "Servico_C_Categoria", "Categoria Serviço", "Tipo", "Tipo_Servico"],
    CONTRACT_COL: ["Contrato", "Contratos", "Financeiro_Contrato_ID", "Contrato_ID", "Numero Contrato", "Número Contrato", "Nº Contrato", "N Contrato", "Contrato Nº", "Contrato Numero"],
    INVOICE_COL: ["Fatura", "Faturas", "Financeiro_Codigo_daFatura", "Codigo Fatura", "Código Fatura", "Numero Fatura", "Número Fatura", "Nº Fatura", "Nota", "Nota Fiscal", "NF", "Invoice"],
    BRANCH_COL: ["Filial", "Filial_Nome", "Filial_Descricao", "Unidade", "Site"],
    REGION_COL: ["Regiao", "Região", "Regional", "Macro_Regiao", "Macro Região", "UF", "Estado", "Cidade", "Municipio", "Município"],
    CC_COL: ["Centro_Custo", "Centro de Custo", "Centro_deCusto_Descricao", "Centro_deCusto_Nome", "Centro_Custo_Descricao", "Centro_Custo_Nome", "CC", "C.C."],
    CC_ID_COL: ["Centro_Custo_ID", "Centro_deCusto_ID", "Codigo CC", "Código CC", "CC_ID"],
    STATUS_COL: ["Status_Contrato", "Financeiro_Contrato_Status", "Status", "Situacao", "Situação"],
    CONTESTED_COL: ["Contestado", "Financeiro_Contestado", "Contestacao", "Contestação"],
}

SHEET_ALIASES = {
    "tbFinanceiro": ["tbFinanceiro", "Financeiro", "Faturas", "Pagamentos", "Lancamentos", "Lançamentos"],
    "tbFornecedor": ["tbFornecedor", "Fornecedores", "Operadoras", "Fornecedor", "Operadora"],
    "tbServicos": ["tbServicos", "Servicos", "Serviços", "tbServiços", "Services"],
    "tbFilial": ["tbFilial", "Filiais", "Unidades", "Sites"],
    "tbCentroCusto": ["tbCentroCusto", "CentroCusto", "Centro de Custo", "Centros de Custo", "CC"],
}

PALETTE = [
    "#2563EB", "#06B6D4", "#10B981", "#F59E0B", "#EF4444", "#64748B",
    "#1D4ED8", "#0E7490", "#059669", "#D97706", "#DC2626", "#475569",
]

if load_dotenv is not None:
    load_dotenv(BASE_DIR / ".env")

st.set_page_config(page_title=APP_TITLE, page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# -----------------------------
# Design
# -----------------------------
def inject_css() -> None:
    st.markdown(
        """
        <style>
            :root { --bg:#060B1A; --card:#0F172A; --muted:#B6C7E6; --txt:#F8FAFC; --cyan:#06B6D4; --blue:#2563EB; --green:#10B981; --amber:#F59E0B; --red:#EF4444; }
            [data-testid="stAppViewContainer"] {
                background:
                    radial-gradient(circle at 10% 0%, rgba(37,99,235,.20), transparent 26%),
                    radial-gradient(circle at 88% 8%, rgba(6,182,212,.14), transparent 30%),
                    linear-gradient(180deg, #060B1A 0%, #0B1020 55%, #070B15 100%) !important;
            }
            .block-container { padding-top: 1rem; padding-bottom: 2rem; max-width: 1580px; }
            .hero { padding: 1.15rem 1.3rem; border-radius: 22px; border: 1px solid rgba(6,182,212,.24); background: linear-gradient(135deg, rgba(15,23,42,.96), rgba(30,64,175,.34), rgba(8,47,73,.30)); box-shadow: 0 18px 48px rgba(37,99,235,.10); margin-bottom: .85rem; }
            .hero h1 { font-size: 1.9rem; margin:0; letter-spacing:-.04em; }
            .hero p { color: var(--muted); margin:.35rem 0 0 0; line-height:1.45; }
            .kpi { min-height:128px; padding: 1rem; border-radius:18px; border:1px solid rgba(148,163,184,.20); background:linear-gradient(150deg, rgba(15,23,42,.97), rgba(30,41,59,.92)); box-shadow: inset 0 1px 0 rgba(255,255,255,.05), 0 12px 28px rgba(0,0,0,.18); border-left: 5px solid var(--blue); }
            .kpi small { color:var(--muted); font-weight:900; letter-spacing:.05em; text-transform:uppercase; }
            .kpi strong { display:block; color:#fff; font-size:1.45rem; margin-top:.35rem; line-height:1.1; }
            .kpi span { display:block; color:var(--muted); margin-top:.35rem; font-size:.8rem; }
            .markdown-panel { padding:1rem 1.1rem; border-radius:18px; border:1px solid rgba(139,92,246,.28); background:linear-gradient(135deg, rgba(139,92,246,.14), rgba(0,212,255,.08)); color:#F8FAFC; line-height:1.55; }
            .markdown-panel.readable-text { font-size:.98rem; line-height:1.65; }
            .markdown-panel.readable-text p { margin:.25rem 0 .65rem 0; }
            .markdown-panel.readable-text ul { margin:.35rem 0 .7rem 1.15rem; padding:0; }
            .markdown-panel.readable-text li { margin:.28rem 0; }
            .markdown-panel.readable-text h3 { margin:.05rem 0 .7rem 0; font-size:1.05rem; letter-spacing:-.01em; }
            .text-currency { font-weight:900; color:#38BDF8; white-space:nowrap; }
            .text-number { font-weight:900; color:#F8FAFC; white-space:nowrap; }

            .filter-chip { display:inline-flex; align-items:center; gap:.35rem; padding:.35rem .55rem; margin:.12rem .15rem; border-radius:999px; background:rgba(0,212,255,.15); border:1px solid rgba(0,212,255,.32); color:#CFFAFE; font-size:.78rem; font-weight:800; }
            .smart-summary-card { padding:1.05rem 1.15rem; border-radius:18px; border:1px solid rgba(148,163,184,.22); background:linear-gradient(145deg, rgba(15,23,42,.96), rgba(30,41,59,.72)); box-shadow: inset 0 1px 0 rgba(255,255,255,.04), 0 16px 36px rgba(0,0,0,.18); }
            .smart-summary-intro { color:#E2E8F0; font-size:1rem; line-height:1.55; margin:0 0 .9rem 0; }
            .smart-summary-grid { display:grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap:.75rem; margin:.9rem 0; }
            .smart-mini-kpi { padding:.85rem .95rem; border-radius:16px; background:rgba(2,6,23,.38); border:1px solid rgba(148,163,184,.18); }
            .smart-mini-kpi small { display:block; color:#94A3B8; font-size:.73rem; font-weight:900; letter-spacing:.06em; text-transform:uppercase; margin-bottom:.25rem; }
            .smart-mini-kpi strong { color:#F8FAFC; font-size:1.12rem; letter-spacing:-.02em; }
            .smart-section-title { color:#F8FAFC; font-weight:900; margin:.95rem 0 .48rem 0; font-size:.92rem; letter-spacing:.01em; }
            .smart-ranking { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:.55rem; }
            .smart-rank-item { border:1px solid rgba(148,163,184,.16); background:rgba(15,23,42,.62); border-radius:14px; padding:.72rem .8rem; min-height:74px; }
            .smart-rank-item .name { color:#E2E8F0; font-weight:850; font-size:.86rem; line-height:1.22; margin-bottom:.3rem; }
            .smart-rank-item .value { color:#38BDF8; font-weight:900; font-size:.95rem; }
            .smart-hint { margin-top:.95rem; padding:.65rem .8rem; border-radius:14px; color:#CBD5E1; background:rgba(6,182,212,.08); border:1px solid rgba(6,182,212,.18); font-size:.88rem; }
            @media (max-width: 900px) { .smart-summary-grid, .smart-ranking { grid-template-columns: 1fr; } }

            div[data-testid="stDialog"] div[role="dialog"] { width: min(92vw, 460px) !important; }
            div[data-testid="stDataFrame"] { border-radius: 14px; overflow:hidden; }
            .stButton>button { border-radius: 13px; font-weight: 800; border: 1px solid rgba(0,212,255,.30); }
            .stButton>button:hover { border-color:#00D4FF; box-shadow: 0 0 0 2px rgba(0,212,255,.12); }
            /* Botão flutuante da IA em link fixo: não cria campo/botão vazio no fluxo do dashboard. */
            .ai-fab-link {
                position: fixed !important; right: 26px !important; bottom: 26px !important; z-index: 2147483647 !important;
                width:64px !important; height:64px !important; border-radius:999px !important; display:flex !important;
                align-items:center !important; justify-content:center !important; text-decoration:none !important; font-size:30px !important; line-height:1 !important;
                background: linear-gradient(135deg, #06B6D4, #7C3AED, #F97316) !important; color:#FFFFFF !important;
                border:1px solid rgba(255,255,255,.46) !important; box-shadow:0 18px 48px rgba(124,58,237,.48), 0 0 0 7px rgba(6,182,212,.12) !important;
                cursor:pointer !important; transition: transform .18s ease, box-shadow .18s ease, filter .18s ease !important;
            }
            .ai-fab-link:hover { transform: translateY(-2px) scale(1.07) !important; filter:saturate(1.18) !important; color:#FFFFFF !important; }
            @media (max-width:768px) { .ai-fab-link { right:16px !important; bottom:16px !important; width:56px !important; height:56px !important; font-size:27px !important; } }
            .st-key-ai_fab_button {
                position: fixed !important; right: 26px !important; bottom: 26px !important;
                z-index: 2147483647 !important; width: 64px !important; height: 64px !important;
                min-width: 64px !important; min-height: 64px !important; max-width: 64px !important; max-height: 64px !important;
                margin: 0 !important; padding: 0 !important; display: flex !important; align-items: center !important; justify-content: center !important;
            }
            .st-key-ai_fab_button button {
                width: 64px !important; height: 64px !important; min-width: 64px !important; min-height: 64px !important; max-width: 64px !important; max-height: 64px !important;
                padding: 0 !important; margin: 0 !important; border-radius: 999px !important; display: inline-flex !important; align-items: center !important; justify-content: center !important;
                background: linear-gradient(135deg, #06B6D4, #7C3AED, #F97316) !important; color: #FFFFFF !important;
                border: 1px solid rgba(255,255,255,.46) !important;
                box-shadow: 0 18px 48px rgba(124,58,237,.48), 0 0 0 7px rgba(6,182,212,.12) !important;
                cursor: pointer !important; transition: transform .18s ease, box-shadow .18s ease, filter .18s ease !important;
            }
            .st-key-ai_fab_button button:hover { transform: translateY(-2px) scale(1.07) !important; filter: saturate(1.18) !important; }
            .ai-loading-note { padding:.58rem .8rem; border-radius:14px; border:1px solid rgba(148,163,184,.16); background:rgba(15,23,42,.46); color:#CBD5E1; font-weight:700; font-size:.84rem; margin:.25rem 0 .55rem; }
            .ai-opening-toast { position:fixed; right:102px; bottom:34px; z-index:2147483646; padding:.58rem .82rem; border-radius:999px; background:rgba(15,23,42,.92); border:1px solid rgba(148,163,184,.20); color:#E2E8F0; box-shadow:0 12px 34px rgba(0,0,0,.28); font-weight:800; font-size:.84rem; }
            .ai-loader-dot { display:inline-block; width:.55rem; height:.55rem; margin-right:.42rem; border-radius:999px; background:#94A3B8; animation: aiPulse 1.05s infinite ease-in-out; vertical-align:-.04rem; }
            @keyframes aiPulse { 0% { opacity:.45; transform:scale(.84); } 50% { opacity:1; transform:scale(1); } 100% { opacity:.45; transform:scale(.84); } }
            .ai-mode-badge { display:inline-flex; align-items:center; gap:.35rem; padding:.28rem .55rem; border-radius:999px; background:rgba(16,185,129,.10); border:1px solid rgba(16,185,129,.20); color:#D1FAE5; font-size:.78rem; font-weight:800; margin-bottom:.35rem; }

            .st-key-ai_fab_button button p, .st-key-ai_fab_button button span, .st-key-ai_fab_button button div {
                font-size: 30px !important; line-height: 1 !important; color: #FFFFFF !important; margin: 0 !important; padding: 0 !important;
            }
            @media (max-width: 768px) {
                .st-key-ai_fab_button, .st-key-ai_fab_button button { right: 16px !important; bottom: 16px !important; width: 56px !important; height: 56px !important; min-width: 56px !important; min-height: 56px !important; max-width: 56px !important; max-height: 56px !important; }
                .st-key-ai_fab_button button p, .st-key-ai_fab_button button span, .st-key-ai_fab_button button div { font-size: 27px !important; }
            }
            .ai-note { font-size:.84rem; color:#B6C7E6; margin-bottom:.5rem; }

            /* IA flutuante reformulada: overlay real, conversa com scroll próprio e input sempre fixo. */
            .gemini-shell {
                height: 100% !important;
                min-height: 0 !important;
                display: flex !important;
                flex-direction: column !important;
                gap: .55rem !important;
                padding: 0 !important;
                overflow: hidden !important;
            }
            .gemini-hero {
                flex: 0 0 auto !important;
                display:flex; align-items:flex-start; justify-content:space-between; gap:.55rem;
                padding:.2rem .1rem .5rem; border-bottom:1px solid rgba(148,163,184,.12); margin:0;
            }
            .gemini-title { font-size:1rem; font-weight:900; color:#F8FAFC; margin:0; letter-spacing:-.01em; }
            .gemini-subtitle { display:block; margin-top:.12rem; color:#94A3B8; font-size:.76rem; line-height:1.25; }
            .gemini-context {
                flex:0 0 auto !important; display:flex; flex-wrap:wrap; gap:.32rem; margin:0;
                padding:.42rem .48rem; border-radius:16px; background:rgba(2,6,23,.30); border:1px solid rgba(148,163,184,.12);
            }
            .gemini-context-chip {
                display:inline-flex; max-width:100%; align-items:center; gap:.25rem; padding:.22rem .45rem; border-radius:999px;
                background:rgba(14,165,233,.10); color:#DFF7FF; border:1px solid rgba(14,165,233,.18); font-size:.72rem; font-weight:800;
                white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
            }
            .gemini-context-chip.muted { background:rgba(148,163,184,.08); color:#CBD5E1; border-color:rgba(148,163,184,.14); }
            .gemini-status-row { margin:.15rem 0 .55rem; color:#94A3B8; font-size:.76rem; }
            .gemini-status-pill { display:none; }
            .gemini-thread {
                flex: 1 1 auto !important;
                height: 100% !important;
                min-height: 0 !important;
                max-height: none !important;
                overflow-y: auto !important;
                overflow-x: hidden !important;
                padding:.28rem .18rem .75rem;
                scroll-behavior:smooth; overscroll-behavior:contain; -webkit-overflow-scrolling:touch;
            }
            .gemini-thread::-webkit-scrollbar { width:7px; }
            .gemini-thread::-webkit-scrollbar-thumb { background:rgba(148,163,184,.30); border-radius:999px; }
            .gemini-row { display:flex; gap:.4rem; align-items:flex-end; width:100%; margin-bottom:.55rem; }
            .gemini-row.user { justify-content:flex-end; }
            .gemini-row.assistant { justify-content:flex-start; }
            .gemini-avatar { display:none; }
            .gemini-bubble { max-width:90%; overflow:visible; padding:.62rem .78rem; border-radius:18px; line-height:1.45; border:1px solid rgba(148,163,184,.12); white-space:normal; overflow-wrap:anywhere; font-size:.91rem; }
            .gemini-row.assistant .gemini-bubble { border-bottom-left-radius:6px; background:rgba(15,23,42,.58); color:#EAF5FF; }
            .gemini-row.user .gemini-bubble { border-bottom-right-radius:6px; background:rgba(37,99,235,.86); color:#FFFFFF; }
            .gemini-bubble p { margin:.08rem 0 .35rem; }
            .gemini-bubble p:last-child { margin-bottom:0; }
            .gemini-bubble ul, .gemini-bubble ol { margin:.3rem 0 .3rem 1rem; padding:0; }
            .gemini-bubble li { margin:.15rem 0; }
            .gemini-bubble h1, .gemini-bubble h2, .gemini-bubble h3 { margin:.08rem 0 .35rem; line-height:1.2; color:#FFFFFF; font-size:1rem; }
            .gemini-bubble code { padding:.06rem .22rem; border-radius:6px; background:rgba(148,163,184,.16); }
            .gemini-input-note { display:none; }
            .gemini-quick-grid { flex:0 0 auto !important; display:grid; grid-template-columns:1fr 1fr; gap:.35rem; margin:.1rem 0 .15rem; }
            .gemini-quick-grid .stButton>button { min-height:32px !important; padding:.32rem .48rem !important; font-size:.76rem !important; border-radius:999px !important; color:#DFF7FF !important; background:rgba(14,165,233,.08) !important; border-color:rgba(14,165,233,.20) !important; }
            .gemini-input-wrap { flex:0 0 auto !important; padding:.42rem .45rem .48rem; border-radius:18px; background:rgba(15,23,42,.96); border:1px solid rgba(14,165,233,.25); box-shadow:0 -8px 22px rgba(15,23,42,.62); }
            .gemini-input-wrap div[data-testid="stTextInput"] { margin-bottom:.35rem !important; }
            .gemini-input-wrap div[data-testid="stTextInput"] input { border-radius:14px !important; min-height:42px !important; }
            .gemini-input-wrap div[data-testid="InputInstructions"] { display:none !important; }
            .st-key-floating_ai_panel {
                position: fixed !important;
                top: 72px !important;
                right: 18px !important;
                bottom: 18px !important;
                z-index: 2147483646 !important;
                width: min(430px, calc(100vw - 36px)) !important;
                height: calc(100dvh - 90px) !important;
                max-height: calc(100dvh - 90px) !important;
                overflow: hidden !important;
                padding: .82rem !important;
                border-radius: 24px !important;
                background: rgba(15,23,42,.98) !important;
                border:1px solid rgba(148,163,184,.22) !important;
                box-shadow:0 24px 70px rgba(0,0,0,.44) !important;
                backdrop-filter: blur(14px) !important;
            }
            .st-key-floating_ai_panel > div[data-testid="stVerticalBlock"],
            .st-key-floating_ai_panel div[data-testid="stVerticalBlock"]:first-child {
                height:100% !important;
                min-height:0 !important;
                display:flex !important;
                flex-direction:column !important;
                overflow:hidden !important;
            }
            .st-key-floating_ai_panel .element-container { min-height:0 !important; }
            .st-key-floating_ai_panel .element-container:has(.gemini-shell) {
                height:100% !important;
                min-height:0 !important;
                display:flex !important;
                flex-direction:column !important;
                overflow:hidden !important;
            }
            .st-key-floating_ai_panel .element-container:has(.gemini-thread) {
                flex:1 1 auto !important;
                min-height:0 !important;
                overflow:hidden !important;
                display:flex !important;
                flex-direction:column !important;
            }
            .st-key-floating_ai_panel .element-container:has(.gemini-input-wrap),
            .st-key-floating_ai_panel .element-container:has(.ai-mini-actions),
            .st-key-floating_ai_panel form,
            .st-key-floating_ai_panel div[data-testid="stForm"] {
                flex:0 0 auto !important;
                min-height:0 !important;
            }
            .st-key-floating_ai_panel form,
            .st-key-floating_ai_panel div[data-testid="stForm"] {
                margin:0 !important;
                padding:0 !important;
                background:transparent !important;
                box-shadow:none !important;
            }
            .ai-mini-actions { flex:0 0 auto !important; padding-top:.28rem; }
            .ai-mini-actions .stButton>button { padding:.34rem .6rem !important; font-size:.82rem !important; border-radius:999px !important; min-height:36px !important; }
            .st-key-floating_ai_panel .stCaption { display:none !important; }
            .st-key-floating_ai_panel iframe { width:0 !important; height:0 !important; min-height:0 !important; display:block !important; opacity:0 !important; pointer-events:none !important; }
            .st-key-floating_ai_panel div[data-testid="stElementToolbar"] { display:none !important; }
            @media (max-width: 768px) {
                .st-key-floating_ai_panel { left: 10px !important; right: 10px !important; top: auto !important; bottom: 78px !important; width: calc(100vw - 20px) !important; height: min(76dvh, 640px) !important; max-height: min(76dvh, 640px) !important; padding:.72rem !important; }
                .gemini-hero { padding:.18rem .06rem .38rem !important; }
                .gemini-context { padding:.34rem .4rem !important; }
                .gemini-context-chip { font-size:.68rem !important; }
                .gemini-bubble { max-width:91% !important; font-size:.9rem !important; padding:.6rem .72rem !important; }
                .gemini-quick-grid { grid-template-columns:1fr !important; }
                .gemini-input-wrap { padding:.38rem !important; }
            }
            

            /* Página dedicada da IA: layout estilo ChatGPT, sem overlay quebrado. */
            .ai-page-shell {
                max-width: 1080px;
                margin: 0 auto;
                min-height: calc(100dvh - 220px);
                display: grid;
                grid-template-columns: minmax(0, 1fr) 280px;
                gap: 1rem;
            }
            .ai-page-main, .st-key-ai_page_main {
                min-height: 68dvh;
                border-radius: 28px;
                border: 1px solid rgba(148,163,184,.20);
                background: linear-gradient(180deg, rgba(15,23,42,.98), rgba(15,23,42,.90));
                box-shadow: 0 24px 70px rgba(0,0,0,.28), inset 0 1px 0 rgba(255,255,255,.04);
                padding: 1rem;
            }
            .ai-page-side {
                border-radius: 24px;
                border: 1px solid rgba(148,163,184,.18);
                background: rgba(15,23,42,.72);
                padding: 1rem;
                height: fit-content;
                position: sticky;
                top: 92px;
            }
            .ai-page-title { font-size: 1.55rem; font-weight: 950; color:#F8FAFC; letter-spacing:-.03em; margin:0; }
            .ai-page-subtitle { color:#94A3B8; margin:.25rem 0 1rem; font-size:.95rem; }
            .ai-page-main .gemini-shell, .st-key-ai_page_main .gemini-shell { min-height: 64dvh !important; height: 64dvh !important; }
            .ai-page-main .gemini-thread, .st-key-ai_page_main .gemini-thread {
                min-height: 0 !important;
                height: auto !important;
                max-height: none !important;
                padding: .6rem .35rem 1rem;
            }
            .ai-page-main .gemini-bubble, .st-key-ai_page_main .gemini-bubble { max-width: 78%; font-size: .96rem; padding: .75rem .92rem; }
            .ai-page-main .gemini-input-wrap, .st-key-ai_page_main .gemini-input-wrap {
                position: sticky;
                bottom: 0;
                z-index: 5;
                margin-top: .4rem;
                background: rgba(15,23,42,.98);
            }
            .ai-page-main .ai-mini-actions, .st-key-ai_page_main .ai-mini-actions { padding-top:.45rem; }
            .ai-side-title { color:#E2E8F0; font-size:.82rem; font-weight:900; letter-spacing:.05em; text-transform:uppercase; margin:.1rem 0 .65rem; }
            .ai-side-note { color:#AAB7CF; font-size:.86rem; line-height:1.45; margin:.35rem 0 .75rem; }
            .ai-page-main .gemini-quick-grid, .st-key-ai_page_main .gemini-quick-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            @media (max-width: 980px) {
                .ai-page-shell { grid-template-columns: 1fr; }
                .ai-page-side { position: static; }
                .ai-page-main .gemini-shell, .st-key-ai_page_main .gemini-shell { height: 68dvh !important; }
                .ai-page-main .gemini-bubble, .st-key-ai_page_main .gemini-bubble { max-width: 88%; }
            }
            @media (max-width: 640px) {
                .ai-page-main, .st-key-ai_page_main { padding:.75rem; border-radius:22px; }
                .ai-page-main .gemini-shell, .st-key-ai_page_main .gemini-shell { height: 72dvh !important; }
                .ai-page-main .gemini-quick-grid, .st-key-ai_page_main .gemini-quick-grid { grid-template-columns:1fr; }
                .ai-page-main .gemini-bubble, .st-key-ai_page_main .gemini-bubble { max-width: 92%; font-size:.92rem; }
            }


            /* Chat IA dedicado: visual limpo, sem botão flutuante e sem painel lateral pesado. */
            .st-key-ai_fab_button, .ai-fab-link { display: none !important; }
            div[data-testid="stChatMessage"] {
                border-radius: 22px !important;
                border: 1px solid rgba(148,163,184,.16) !important;
                background: rgba(15,23,42,.54) !important;
                padding: .85rem 1rem !important;
                margin: .55rem 0 !important;
                box-shadow: inset 0 1px 0 rgba(255,255,255,.03) !important;
            }
            div[data-testid="stChatMessage"] p { line-height: 1.55 !important; }
            div[data-testid="stChatInput"] {
                border-radius: 24px !important;
                border: 1px solid rgba(14,165,233,.24) !important;
                background: rgba(15,23,42,.96) !important;
                box-shadow: 0 -12px 36px rgba(2,6,23,.36) !important;
            }
            .ai-clean-header {
                display:flex; align-items:flex-start; justify-content:space-between; gap:1rem;
                padding: 1.2rem 1.35rem; margin: .25rem 0 1rem;
                border-radius: 26px;
                background: radial-gradient(circle at top left, rgba(6,182,212,.18), transparent 36%), linear-gradient(145deg, rgba(15,23,42,.96), rgba(2,6,23,.80));
                border: 1px solid rgba(148,163,184,.18);
                box-shadow: 0 22px 60px rgba(0,0,0,.24), inset 0 1px 0 rgba(255,255,255,.05);
            }
            .ai-clean-title { color:#F8FAFC; font-size:1.65rem; font-weight:950; letter-spacing:-.04em; margin:0; }
            .ai-clean-subtitle { color:#AAB7CF; font-size:.94rem; line-height:1.45; margin:.25rem 0 0; max-width:740px; }
            .ai-clean-status { display:flex; flex-wrap:wrap; justify-content:flex-end; gap:.35rem; min-width:260px; }
            .ai-clean-stage {
                max-width: 1040px; margin: 0 auto 1rem;
                padding: 1rem; border-radius: 28px;
                border: 1px solid rgba(148,163,184,.18);
                background: linear-gradient(180deg, rgba(15,23,42,.82), rgba(15,23,42,.46));
                box-shadow: 0 18px 48px rgba(0,0,0,.18);
            }
            .ai-empty-state {
                text-align:center; padding: 3.2rem 1rem 2.2rem; color:#AAB7CF;
            }
            .ai-empty-state .icon { font-size:2.35rem; margin-bottom:.7rem; }
            .ai-empty-state strong { display:block; color:#F8FAFC; font-size:1.15rem; margin-bottom:.25rem; }
            .ai-suggestion-title { color:#CBD5E1; font-size:.8rem; font-weight:900; text-transform:uppercase; letter-spacing:.07em; margin:.15rem 0 .55rem; }
            .ai-suggestion-grid { display:grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap:.55rem; margin:.3rem 0 1rem; }
            .ai-suggestion-grid .stButton>button { min-height:44px !important; border-radius:18px !important; white-space:normal !important; font-size:.84rem !important; background:rgba(14,165,233,.08) !important; border-color:rgba(14,165,233,.22) !important; color:#DFF7FF !important; }
            @media (max-width: 980px) {
                .ai-clean-header { flex-direction:column; }
                .ai-clean-status { justify-content:flex-start; min-width:0; }
                .ai-suggestion-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
            }
            @media (max-width: 640px) {
                .ai-clean-header, .ai-clean-stage { border-radius:20px; padding:.9rem; }
                .ai-clean-title { font-size:1.35rem; }
                .ai-suggestion-grid { grid-template-columns:1fr; }
            }

            .kpi.delta-up { border-left-color: var(--green); }
            .kpi.delta-down { border-left-color: var(--red); }
            .kpi.delta-warn { border-left-color: var(--amber); }
            .kpi .meta { display:flex; justify-content:space-between; gap:.6rem; margin-top:.55rem; font-size:.78rem; color:var(--muted); }
            .kpi .delta { display:inline-flex; align-items:center; gap:.25rem; margin-top:.5rem; padding:.22rem .48rem; border-radius:999px; font-size:.78rem; font-weight:900; background:rgba(16,185,129,.12); color:#86EFAC; }
            .kpi .delta.neg { background:rgba(239,68,68,.12); color:#FCA5A5; }
            .kpi .delta.neu { background:rgba(6,182,212,.12); color:#67E8F9; }
            .executive-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:.75rem; margin:.75rem 0 1rem; }
            .readme-card { padding:.82rem .95rem; border-radius:15px; border:1px solid rgba(6,182,212,.22); background:rgba(8,47,73,.30); color:#CFFAFE; font-size:.88rem; line-height:1.48; margin:.4rem 0 .7rem; }
            .section-note { color:#B6C7E6; font-size:.9rem; line-height:1.45; margin:-.2rem 0 .7rem; }
            .exec-card { padding:.9rem 1rem; border-radius:16px; border:1px solid rgba(148,163,184,.18); background:rgba(15,23,42,.76); }
            .exec-card strong { display:block; color:#fff; margin-bottom:.28rem; }
            .exec-card span { color:var(--muted); font-size:.88rem; line-height:1.42; }
            .visual-shell { padding:1.2rem 1.25rem; border-radius:22px; border:1px solid rgba(148,163,184,.22); background:linear-gradient(150deg, rgba(15,23,42,.94), rgba(8,47,73,.28)); box-shadow: 0 18px 44px rgba(2,6,23,.22), inset 0 1px 0 rgba(255,255,255,.045); margin:1rem 0 1.35rem; width:100%; overflow:visible; }
            .visual-shell [data-testid='stPlotlyChart'] { min-height: 360px; }
            .visual-shell [data-testid='stExpander'] { border:1px solid rgba(6,182,212,.22); border-radius:16px; background:rgba(15,23,42,.52); margin-top:1rem; width:100%; }
            .visual-shell [data-testid='stExpander'] summary { font-weight:900; color:#DFFBFF; }
            .director-grid { display:grid; grid-template-columns: minmax(0,1.35fr) minmax(330px,.75fr); gap:.9rem; align-items:start; }
            .director-grid-3 { display:grid; grid-template-columns: repeat(3, minmax(0,1fr)); gap:.9rem; align-items:start; }
            .compact-section-title { margin:.95rem 0 .25rem; font-size:1.05rem; font-weight:950; color:#F8FAFC; }
            @media (max-width: 1100px) { .director-grid, .director-grid-3 { grid-template-columns:1fr; } }
            .visual-head { display:flex; justify-content:space-between; align-items:flex-start; gap:1rem; margin-bottom:.75rem; }
            .visual-title { color:#FFFFFF; font-size:1.22rem; font-weight:950; letter-spacing:-.02em; margin:0; }
            .visual-subtitle { color:#AFC4E6; font-size:.9rem; line-height:1.45; margin:.22rem 0 0; }
            .visual-analysis { margin:.62rem 0 .85rem; padding:.82rem .92rem; border-radius:18px; background:rgba(2,6,23,.34); border:1px solid rgba(6,182,212,.22); border-left:5px solid var(--blue); color:#EAF5FF; line-height:1.5; }
            .visual-analysis .badge { display:inline-flex; padding:.18rem .48rem; border-radius:999px; border:1px solid rgba(6,182,212,.28); background:rgba(6,182,212,.12); color:#CFFAFE; font-size:.74rem; font-weight:950; margin-bottom:.45rem; }
            .visual-analysis p { margin:.24rem 0; font-size:.91rem; }
            .visual-analysis strong { color:#FFFFFF; }
            .attention-card { border-left-color: var(--amber) !important; background:linear-gradient(135deg, rgba(245,158,11,.13), rgba(15,23,42,.78)) !important; }
            .risk-card { border-left-color: var(--red) !important; background:linear-gradient(135deg, rgba(239,68,68,.13), rgba(15,23,42,.78)) !important; }
            .opportunity-card { border-left-color: var(--green) !important; background:linear-gradient(135deg, rgba(16,185,129,.13), rgba(15,23,42,.78)) !important; }
            .cc-summary { padding:.95rem 1rem; border-radius:18px; border:1px solid rgba(99,102,241,.28); background:linear-gradient(135deg, rgba(37,99,235,.18), rgba(15,23,42,.82)); margin:.65rem 0 1rem; }
            .cc-summary h4 { margin:0 0 .35rem; color:#fff; }
            .cc-summary p { margin:.25rem 0; color:#DDEBFF; line-height:1.45; font-size:.92rem; }
            .cc-summary .mini { display:inline-flex; margin:.18rem .35rem .18rem 0; padding:.2rem .5rem; border-radius:999px; background:rgba(255,255,255,.08); color:#CFFAFE; font-size:.78rem; font-weight:900; border:1px solid rgba(255,255,255,.10); }
            .insufficient-card { padding:.9rem 1rem; border-radius:16px; border:1px dashed rgba(245,158,11,.40); background:rgba(245,158,11,.08); color:#FDE68A; line-height:1.45; margin:.5rem 0 .7rem; }
            .insufficient-card strong { color:#FEF3C7; }
            .action-list { display:grid; grid-template-columns: 1fr; gap:.72rem; margin:.45rem 0 .75rem; }
            .action-card { padding:.9rem 1rem; border-radius:18px; border:1px solid rgba(148,163,184,.18); background:linear-gradient(135deg, rgba(15,23,42,.93), rgba(30,41,59,.72)); border-left:5px solid var(--blue); }
            .action-card.alta { border-left-color: var(--red); }
            .action-card.media, .action-card.média { border-left-color: var(--amber); }
            .action-card.baixa, .action-card.ok { border-left-color: var(--green); }
            .action-card h4 { margin:0 0 .4rem 0; color:#fff; font-size:1rem; }
            .action-card .tag { display:inline-flex; padding:.18rem .48rem; border-radius:999px; background:rgba(6,182,212,.14); color:#CFFAFE; border:1px solid rgba(6,182,212,.22); font-size:.73rem; font-weight:900; margin-right:.35rem; }
            .action-card p { margin:.28rem 0; color:#DDEBFF; line-height:1.42; font-size:.88rem; }
            .action-card strong { color:#CFFAFE; }
            .mode-card { padding:.9rem 1rem; border-radius:18px; border:1px solid rgba(148,163,184,.18); background:rgba(15,23,42,.82); margin:.65rem 0 .9rem; }
            .mode-card h4 { margin:0 0 .35rem; color:#fff; }
            .mode-card p { margin:.22rem 0; color:#DDEBFF; line-height:1.45; font-size:.9rem; }
            .priority-badge { display:inline-flex; align-items:center; gap:.35rem; padding:.2rem .55rem; border-radius:999px; font-weight:950; font-size:.74rem; border:1px solid rgba(255,255,255,.12); }
            .priority-alta { background:rgba(239,68,68,.16); color:#FCA5A5; }
            .priority-media, .priority-média { background:rgba(245,158,11,.16); color:#FDE68A; }
            .priority-baixa { background:rgba(16,185,129,.14); color:#BBF7D0; }
            .validation-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:.7rem; margin:.55rem 0 .9rem; }
            .validation-card { padding:.75rem .85rem; border-radius:16px; background:rgba(15,23,42,.82); border:1px solid rgba(148,163,184,.18); }
            .validation-card strong { display:block; color:#fff; margin-bottom:.25rem; }
            .validation-card span { color:#B6C7E6; font-size:.86rem; line-height:1.4; }
            .comparison-card { padding:.95rem 1rem; border-radius:18px; border:1px solid rgba(6,182,212,.24); background:linear-gradient(135deg, rgba(6,182,212,.12), rgba(15,23,42,.84)); margin:.6rem 0 1rem; }
            .comparison-card h4 { margin:0 0 .35rem; color:#fff; }
            .comparison-card p { margin:.25rem 0; color:#DDEBFF; line-height:1.45; font-size:.92rem; }

            .director-minutes { padding:1rem 1.1rem; border-radius:20px; border:1px solid rgba(6,182,212,.25); background:linear-gradient(135deg, rgba(15,23,42,.94), rgba(30,64,175,.22)); margin:.75rem 0 1rem; border-left:5px solid var(--blue); }
            .director-minutes .minutes-head { display:flex; justify-content:space-between; gap:.75rem; align-items:center; margin-bottom:.45rem; }
            .director-minutes .badge { display:inline-flex; padding:.2rem .55rem; border-radius:999px; background:rgba(6,182,212,.14); color:#CFFAFE; border:1px solid rgba(6,182,212,.26); font-weight:950; font-size:.76rem; }
            .director-minutes strong { color:#fff; }
            .director-minutes p { margin:.28rem 0; color:#DDEBFF; line-height:1.43; font-size:.92rem; }

            .radar-grid { display:grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap:.75rem; margin:.8rem 0 1rem; }
            .radar-card { min-height:128px; padding:.9rem 1rem; border-radius:18px; border:1px solid rgba(148,163,184,.18); background:linear-gradient(135deg, rgba(15,23,42,.92), rgba(30,41,59,.82)); border-left:5px solid var(--blue); box-shadow:0 12px 32px rgba(0,0,0,.16); }
            .radar-card strong { display:block; color:#fff; margin-bottom:.35rem; font-size:.95rem; }
            .radar-card span { display:block; color:#CFFAFE; font-weight:950; line-height:1.22; }
            .radar-card p { color:#DDEBFF; margin:.45rem 0 0; font-size:.86rem; line-height:1.36; }
            .radar-card.risk { border-left-color: var(--red); background:linear-gradient(135deg, rgba(127,29,29,.28), rgba(15,23,42,.88)); }
            .radar-card.attention { border-left-color: var(--amber); background:linear-gradient(135deg, rgba(120,53,15,.26), rgba(15,23,42,.88)); }
            .radar-card.opportunity { border-left-color: var(--green); background:linear-gradient(135deg, rgba(6,95,70,.24), rgba(15,23,42,.88)); }
            @media (max-width: 900px) { .executive-grid, .visual-action-grid, .validation-grid, .radar-grid { grid-template-columns: 1fr; } }

            h2, h3 { letter-spacing:-.02em; }
            hr { border-color: rgba(148,163,184,.18); }
        
.detail-summary-note {
    margin: 0.15rem 0 0.85rem 0;
    color: rgba(229, 231, 235, 0.78);
    font-size: 0.92rem;
    line-height: 1.45;
}
</style>
        """,
        unsafe_allow_html=True,
    )

# -----------------------------
# Helpers
# -----------------------------
BAD_TEXT_TOKENS = {"", "nan", "none", "nat", "undefined", "null", "<na>"}
BAD_TEXT_PATTERN = re.compile(r"(?i)\b(undefined|none|nan|null|nat|<na>)\b")

def clean_text(value: Any, default: str = "Não informado") -> str:
    """Normaliza texto de exibição para evitar undefined/None/nan na interface.

    Alguns gráficos do Plotly herdavam valores como ``undefined`` não como título
    direto, mas dentro de labels/customdata, por exemplo ``Filial undefined`` ou
    ``undefined / undefined``. Por isso a limpeza remove também ocorrências
    internas desses tokens e não apenas quando o valor inteiro é igual a eles.
    """
    try:
        if value is None or pd.isna(value):
            return default
    except Exception:
        if value is None:
            return default
    txt = str(value).replace("\u00a0", " ").strip()
    txt = re.sub(r"\s+", " ", txt)
    if txt.lower() in BAD_TEXT_TOKENS:
        return default
    cleaned = BAD_TEXT_PATTERN.sub("", txt)
    cleaned = re.sub(r"\s*[-–—|/\\]+\s*", " - ", cleaned)
    cleaned = re.sub(r"(?:^\s*[-–—|/\\]+\s*|\s*[-–—|/\\]+\s*$)", "", cleaned).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    if cleaned.lower() in BAD_TEXT_TOKENS or not cleaned:
        return default
    return cleaned

def normalize_key(value: Any) -> str:
    return re.sub(r"\s+", " ", clean_text(value, "").lower()).strip()

def short_label(value: Any, limit: int = 28) -> str:
    txt = clean_text(value, "Não informado")
    txt = re.sub(r"\s+", " ", txt).strip()
    if len(txt) <= limit:
        return txt
    return txt[: max(8, limit - 1)].rstrip() + "…"

def money(v: Any) -> str:
    try: x = float(v)
    except Exception: x = 0.0
    return f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

def compact_money(v: Any) -> str:
    try: x = float(v)
    except Exception: x = 0.0
    if abs(x) >= 1_000_000: return f"R$ {x/1_000_000:.1f} mi".replace(".", ",")
    if abs(x) >= 1_000: return f"R$ {x/1_000:.1f} mil".replace(".", ",")
    return money(x)

def pct(v: Any) -> str:
    try: x = float(v)
    except Exception: x = 0.0
    return f"{x:.1f}%".replace(".", ",")



def safe_numeric_series(df: pd.DataFrame, col: str) -> pd.Series:
    """Retorna coluna numérica estável para gráficos e agregações."""
    if col not in df.columns:
        return pd.Series([0.0] * len(df), index=df.index, dtype="float64")
    return to_number_safe(df[col]).fillna(0.0)

def normalize_dimension_series(series: pd.Series, default: str) -> pd.Series:
    """Normaliza dimensões para evitar filtros/gráficos vazios por NaN, espaços ou IDs quebrados."""
    return series.map(lambda x: clean_text(x, default)).astype(str).str.replace(r"\s+", " ", regex=True).str.strip().replace({"": default})

def period_sort_key(value: Any) -> str:
    txt = clean_text(value, "Sem período")
    try:
        if re.fullmatch(r"\d{1,2}/\d{2}", txt):
            month, year = txt.split("/")
            return f"20{int(year):02d}-{int(month):02d}-01"
        if re.fullmatch(r"\d{4}-\d{2}", txt):
            return f"{txt}-01"
        dt = pd.to_datetime(txt, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return txt


def format_period_label(value: Any) -> str:
    """Exibe períodos no padrão brasileiro MM/AA, mantendo ordenação por data."""
    if pd.isna(value):
        return "Sem período"
    try:
        dt = pd.to_datetime(value, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%m/%y")
    except Exception:
        pass
    txt = clean_text(value, "Sem período")
    if re.fullmatch(r"\d{4}-\d{2}", txt):
        year, month = txt.split("-")
        return f"{month}/{year[-2:]}"
    return txt


def choose_period_date_column(df: pd.DataFrame) -> Optional[str]:
    """Escolhe a melhor coluna para representar o período padrão do dashboard."""
    return choose_due_date_column(df) or choose_reference_date_column(df) or find_column(df, DATE_COL)


def choose_reference_date_column(df: pd.DataFrame) -> Optional[str]:
    """Localiza a melhor coluna de mês/data de referência da fatura."""
    preferred = [
        "Financeiro_mesReferencia", "mesReferencia", "Mês Referência", "Mes Referencia",
        "Referencia", "Referência", "Competencia", "Competência", "Periodo", "Período", "Mes", "Mês",
    ]
    for col in preferred:
        if col in df.columns and _series_has_useful_values(df[col]):
            parsed = to_datetime_safe(df[col])
            if parsed.notna().any():
                return col
    return None


def choose_due_date_column(df: pd.DataFrame) -> Optional[str]:
    """Localiza a melhor coluna de vencimento/leitura para visão financeira."""
    preferred = [
        "Financeiro_Vencimento", "Data_Vencimento", "Vencimento",
        "Financeiro_Data_Leitura", "Data_Leitura", "Data Leitura", "Data",
    ]
    for col in preferred:
        if col in df.columns and _series_has_useful_values(df[col]):
            parsed = to_datetime_safe(df[col])
            if parsed.notna().any():
                return col
    return None


def apply_period_basis(df: pd.DataFrame) -> pd.DataFrame:
    """Ativa o período exibido/filtrado conforme escolha da sidebar."""
    if df is None or df.empty:
        return df
    basis = st.session_state.get(PERIOD_BASIS_KEY, "Vencimento")
    if basis == "Mês de referência" and REF_DATE_COL in df.columns:
        df[DATE_COL] = df[REF_DATE_COL]
        df[MONTH_COL] = df[REF_MONTH_COL] if REF_MONTH_COL in df.columns else df[DATE_COL].map(format_period_label)
        df["Periodo_Origem_Ativa"] = df.get("Periodo_Origem_Referencia", "Mês de referência")
    else:
        df[DATE_COL] = df[DUE_DATE_COL] if DUE_DATE_COL in df.columns else df.get(DATE_COL, pd.NaT)
        df[MONTH_COL] = df[DUE_MONTH_COL] if DUE_MONTH_COL in df.columns else df[DATE_COL].map(format_period_label)
        df["Periodo_Origem_Ativa"] = df.get("Periodo_Origem_Vencimento", "Vencimento")
    df[MONTH_COL] = normalize_dimension_series(df[MONTH_COL], "Sem período")
    return df

def ensure_model_quality(df: pd.DataFrame) -> pd.DataFrame:
    """Camada final de qualidade: todos os visuais passam a usar as mesmas dimensões/valores."""
    out = df.copy()
    for col, default in [
        (SUPPLIER_COL, "Sem fornecedor"), (SERVICE_COL, "Sem serviço"), (CATEGORY_COL, "Sem categoria"),
        (CONTRACT_COL, "Sem contrato"), (BRANCH_COL, "Sem filial"), (REGION_COL, "Sem região"), (CC_COL, "Sem centro de custo"),
        (CC_ID_COL, "Sem código"), (STATUS_COL, "Não informado"), (INVOICE_COL, "Sem fatura"),
        (CONTESTED_COL, "Não informado"), (MONTH_COL, "Sem período"),
    ]:
        if col not in out.columns:
            out[col] = default
        out[col] = normalize_dimension_series(out[col], default)
    out[VALUE_COL] = safe_numeric_series(out, VALUE_COL)
    out[DIFF_COL] = safe_numeric_series(out, DIFF_COL)
    if "Valor_Contratado" in out.columns:
        out["Valor_Contratado"] = safe_numeric_series(out, "Valor_Contratado")
    if "Linha_ID" not in out.columns:
        out["Linha_ID"] = range(1, len(out) + 1)
    return out

def filtered_options(df: pd.DataFrame, field: str) -> List[str]:
    if field not in df.columns:
        return []
    vals = [clean_text(v, default_for_field(field)) for v in df[field].dropna().astype(str).unique().tolist()]
    sort_fn = period_sort_key if field == MONTH_COL else (lambda x: normalize_key(x))
    return sorted([v for v in vals if normalize_key(v) not in {"", "nan", "none", "nat", "undefined", "null"}], key=sort_fn)

def normalize_markdown(text: Any) -> str:
    """Normaliza respostas em Markdown preservando quebras de linha.

    Importante: não usar clean_text aqui, porque ele compacta todos os espaços
    e transforma listas/títulos em uma única linha, quebrando o Markdown do chat.
    """
    if text is None:
        return ""
    try:
        if pd.isna(text):
            return ""
    except Exception:
        pass

    txt = str(text).replace("\u00a0", " ").replace("\r\n", "\n").replace("\r", "\n").strip()
    txt = txt.replace("\\*", "*").replace("\\~", "~")
    txt = re.sub(r"<br\s*/?>", "\n", txt, flags=re.I)
    txt = re.sub(r"</?(div|span|p|strong|b|em|i|code)[^>]*>", "", txt, flags=re.I)

    # Limpa espaços por linha sem destruir parágrafos/listas.
    txt = "\n".join(re.sub(r"[ \t]+", " ", line).strip() for line in txt.split("\n"))

    # Corrige moedas e remove artefatos que faziam o Streamlit renderizar valores como código/math.
    txt = re.sub(r"`+\s*(R\$?\s*[\d\.]+,\d{2})\s*`+", r"\1", txt)
    txt = re.sub(r"\bR\s+(?=\d)", "R$ ", txt)
    txt = re.sub(r"R\$\s*\$+", "R$", txt)
    txt = re.sub(r"R\$\$+", "R$", txt)

    # Remove ênfase solta ao redor de moedas/números, preservando headings e listas.
    txt = re.sub(r"\*{2,}([^\n]*?R\$?\s*[\d\.]+,\d{2}[^\n]*?)\*{2,}", r"\1", txt)
    txt = re.sub(r"(?<!\*)\*{2,}(?!\*)", "", txt)

    # Remove títulos/frases genéricas que poluem respostas objetivas.
    noise = [
        r"^Resumo analítico:\s*",
        r"Use os filtros do painel ou selecione pontos nos gráficos para refinar a análise\.?",
        r"### Próximos passos sugeridos\s*",
    ]
    for pat in noise:
        txt = re.sub(pat, "", txt, flags=re.I | re.M).strip()

    txt = re.sub(r"\n{3,}", "\n\n", txt)
    return txt.strip()

def text_block_to_html(text: Any) -> str:
    """Renderiza texto de painel sem depender de Markdown do Streamlit.

    Isso evita que valores como R$ 70.719,34 sejam interpretados como LaTeX,
    código inline ou markdown quebrado. Mantém parágrafos, títulos e listas.
    """
    clean = normalize_markdown(text)
    if not clean:
        return ""
    lines = clean.split("\n")
    html_parts: List[str] = []
    in_ul = False

    def close_ul() -> None:
        nonlocal in_ul
        if in_ul:
            html_parts.append("</ul>")
            in_ul = False

    def inline_safe(value: str) -> str:
        safe = html.escape(value, quote=True)
        # Destaque de moeda sem usar markdown.
        safe = re.sub(r"(R\$\s*[0-9.]+,[0-9]{2})", r"<span class='text-currency'>\1</span>", safe)
        safe = re.sub(r"\b([0-9]+)\s+(lançamento\(s\)|lançamentos|faturas|contratos|serviços)\b", r"<span class='text-number'>\1</span> \2", safe, flags=re.I)
        return safe

    for raw in lines:
        line = raw.strip()
        if not line:
            close_ul()
            continue
        heading = re.match(r"^(#{1,4})\s+(.+)$", line)
        if heading:
            close_ul()
            level = min(4, len(heading.group(1)) + 2)
            html_parts.append(f"<h{level}>{inline_safe(heading.group(2))}</h{level}>")
            continue
        bullet = re.match(r"^[-•]\s+(.+)$", line)
        if bullet:
            if not in_ul:
                html_parts.append("<ul>")
                in_ul = True
            html_parts.append(f"<li>{inline_safe(bullet.group(1))}</li>")
            continue
        close_ul()
        html_parts.append(f"<p>{inline_safe(line)}</p>")
    close_ul()
    return "".join(html_parts)


def render_markdown_box(text: Any) -> None:
    body = text_block_to_html(text)
    with st.container(border=True):
        st.markdown(f"<div class='markdown-panel readable-text'>{body}</div>", unsafe_allow_html=True)


def dashboard_local_summary(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> str:
    """Resumo leve para a tela principal sem chamar IA externa em todo rerun."""
    df = full_df if isinstance(full_df, pd.DataFrame) else pd.DataFrame()
    visible = visible_df if isinstance(visible_df, pd.DataFrame) else df
    if df.empty:
        return "Sem dados carregados para gerar resumo."
    total = visible[VALUE_COL].sum() if VALUE_COL in visible.columns else 0
    diff = visible[DIFF_COL].sum() if DIFF_COL in visible.columns else 0
    top_supplier = _top_items_text(visible, SUPPLIER_COL, 3) if SUPPLIER_COL in visible.columns else "não disponível"
    top_branch = _top_items_text(visible, BRANCH_COL, 3) if BRANCH_COL in visible.columns else "não disponível"
    return "\n".join([
        f"A operação filtrada reúne {len(visible)} lançamento(s) e soma {money(total)}.",
        f"Diferença acumulada: {money(diff)}.",
        "",
        f"Top fornecedores: {top_supplier}.",
        f"Top filiais: {top_branch}.",
        "",
        "➡ Para perguntas detalhadas, use o botão flutuante da IA."
    ])

def render_reading_note(text: str) -> None:
    """Card curto para explicar como ler gráficos/tabelas analíticas."""
    clean = normalize_markdown(text)
    if clean:
        st.markdown(f"<div class='readme-card'>{inline_markdown_to_html(clean)}</div>", unsafe_allow_html=True)

def trace_supports(trace: Any, prop: str) -> bool:
    """Confere se o tipo de trace aceita a propriedade antes de aplicar update.

    Alguns traces do Plotly, como Treemap, Heatmap e Indicator, não aceitam
    propriedades comuns em barras/linhas, como showlegend ou legendgroup.
    Essa checagem evita quebrar a tela inteira por causa de um visual específico.
    """
    try:
        return prop in (getattr(trace, "_valid_props", set()) or set())
    except Exception:
        return False


def hover_money_label(label: Any, value: Any, label_name: str = "", value_name: str = "Gasto") -> str:
    """Tooltip pré-formatado para evitar divergência entre valor da barra e valor exibido no hover."""
    prefix = f"{clean_text(label_name)}: " if clean_text(label_name, "") else ""
    return f"{prefix}{clean_text(label, 'Não informado')}<br>{value_name}: {money(value)}"

def hover_texts(df: pd.DataFrame, label_col: str, value_col: str, label_name: str = "", value_name: str = "Gasto") -> list[str]:
    if df is None or df.empty or label_col not in df.columns or value_col not in df.columns:
        return []
    return [hover_money_label(row[label_col], row[value_col], label_name, value_name) for _, row in df.iterrows()]


def clean_plotly_template(value: Any) -> str:
    """Limpa templates do Plotly sem quebrar tags como <extra></extra>."""
    if value is None:
        return ""
    txt = str(value).replace("\u00a0", " ").strip()
    # Corrige templates quebrados por versões anteriores da limpeza.
    txt = txt.replace("<extra>< - extra>", "<extra></extra>")
    txt = txt.replace("<extra><- extra>", "<extra></extra>")
    txt = txt.replace("<extra>< -extra>", "<extra></extra>")
    txt = txt.replace("<extra><-extra>", "<extra></extra>")
    txt = txt.replace("< extra>< - extra>", "<extra></extra>")
    # Remove apenas tokens técnicos, preservando sintaxe %{x}, <br> e <extra>.
    txt = BAD_TEXT_PATTERN.sub("", txt)
    txt = re.sub(r"(?:<br>\s*){2,}", "<br>", txt, flags=re.I)
    txt = re.sub(r"^\s*<br>\s*|\s*<br>\s*$", "", txt, flags=re.I)
    # Se existir tag extra aberta e não houver fechamento correto, corrige para ocultar box secundário.
    if "<extra>" in txt and "</extra>" not in txt:
        txt = re.sub(r"<extra>.*$", "<extra></extra>", txt, flags=re.I)
    return txt

def safe_trace_update(trace: Any, **kwargs: Any) -> None:
    """Atualiza apenas propriedades suportadas pelo trace atual."""
    supported = {k: v for k, v in kwargs.items() if trace_supports(trace, k)}
    if supported:
        try:
            trace.update(**supported)
        except Exception:
            for k, v in supported.items():
                try:
                    setattr(trace, k, v)
                except Exception:
                    pass

def _clean_nested_display_value(value: Any, default: str = "Não informado") -> Any:
    """Limpa valores textuais dentro de listas/arrays/DataFrames usados por Plotly e tabelas."""
    if isinstance(value, str):
        return clean_text(value, default)
    try:
        if isinstance(value, pd.DataFrame):
            return sanitize_dataframe_display(value).astype(object).values
        if isinstance(value, pd.Series):
            return value.map(lambda v: _clean_nested_display_value(v, default)).astype(object).values
    except Exception:
        pass
    if isinstance(value, (list, tuple)):
        return type(value)(_clean_nested_display_value(v, default) for v in value)
    try:
        import numpy as _np
        if isinstance(value, _np.ndarray):
            cleaned = [_clean_nested_display_value(v, default) for v in value.tolist()]
            return _np.array(cleaned, dtype=object)
    except Exception:
        pass
    try:
        if pd.isna(value):
            return default
    except Exception:
        pass
    return value

def _clean_plotly_array(value: Any) -> Any:
    """Limpa arrays de labels/textos do Plotly sem converter números úteis."""
    if value is None:
        return value
    return _clean_nested_display_value(value, "Não informado")

def sanitize_figure(fig: go.Figure) -> go.Figure:
    """Remove títulos, legendas e textos ruins antes de enviar o gráfico para a tela."""
    if fig is None:
        return fig
    try:
        fig.update_layout(title_text="")
    except Exception:
        pass

    # Limpa títulos de eixos/legendas/colorbars quando Plotly herdou valores vazios/undefined.
    try:
        layout = fig.layout
        for axis_name in ["xaxis", "yaxis", "xaxis2", "yaxis2", "xaxis3", "yaxis3"]:
            axis = getattr(layout, axis_name, None)
            if axis is not None and getattr(axis, "title", None) is not None:
                current = getattr(axis.title, "text", None)
                if normalize_key(current) in {"undefined", "none", "nan", "null", ""}:
                    axis.title.text = None
                else:
                    axis.title.text = clean_display_label(current, "")
        if getattr(layout, "legend", None) is not None and getattr(layout.legend, "title", None) is not None:
            current = getattr(layout.legend.title, "text", None)
            if normalize_key(current) in {"undefined", "none", "nan", "null", ""}:
                layout.legend.title.text = None
        if getattr(layout, "coloraxis", None) is not None and getattr(layout.coloraxis, "colorbar", None) is not None:
            cb = layout.coloraxis.colorbar
            if getattr(cb, "title", None) is not None:
                current = getattr(cb.title, "text", None)
                if normalize_key(current) in {"undefined", "none", "nan", "null", ""}:
                    cb.title.text = None
    except Exception:
        pass

    try:
        if getattr(fig.layout, "annotations", None):
            for ann in fig.layout.annotations:
                current = getattr(ann, "text", None)
                cleaned = clean_display_label(current, "")
                ann.text = cleaned
    except Exception:
        pass

    for trace in fig.data:
        name = clean_display_label(getattr(trace, "name", ""), "")
        if normalize_key(name) in {"undefined", "none", "nan", "null", ""}:
            safe_trace_update(trace, name="", showlegend=False, legendgroup="")
            if trace_supports(trace, "name"):
                try:
                    trace.name = ""
                except Exception:
                    pass
        else:
            safe_trace_update(trace, name=name)
        for attr in ["x", "y", "labels", "parents", "ids", "text", "hovertext", "customdata"]:
            try:
                value = getattr(trace, attr, None)
                if value is not None:
                    setattr(trace, attr, _clean_plotly_array(value))
            except Exception:
                pass
        for attr in ["hovertemplate", "texttemplate"]:
            try:
                value = getattr(trace, attr, None)
                if value is None:
                    continue
                setattr(trace, attr, clean_plotly_template(value))
            except Exception:
                pass
    try:
        fig = strip_plotly_undefined(fig)
    except Exception:
        pass
    return fig


def style_fig(fig: go.Figure, height: int = 460, showlegend: bool = False) -> go.Figure:
    """Compatibilidade para gráficos antigos.

    Algumas funções do projeto ainda chamam ``style_fig``. A V76 passou a usar
    ``sanitize_figure`` + layout no renderizador, mas manter este wrapper evita
    NameError e aplica um padrão seguro sem forçar propriedades incompatíveis
    em traces como Treemap, Heatmap e Indicator.
    """
    if fig is None:
        fig = go.Figure()
    fig = sanitize_figure(fig)
    try:
        fig.update_layout(
            height=max(360, int(height or 460)),
            showlegend=bool(showlegend),
            margin=dict(t=26, b=72, l=56, r=86),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#E5E7EB"),
            hoverlabel=dict(bgcolor="#0F172A", font_size=12, font_color="#F8FAFC"),
        )
    except Exception:
        pass
    for trace in getattr(fig, "data", []):
        # Aplica propriedades apenas quando o tipo de trace suporta.
        safe_trace_update(trace, showlegend=bool(showlegend))
        if getattr(trace, "hovertemplate", None):
            try:
                trace.hovertemplate = clean_plotly_template(trace.hovertemplate)
            except Exception:
                pass
    return fig

def safe_title(text: Any, default: str = "") -> str:
    """Evita que títulos vazios ou 'undefined' apareçam na interface."""
    value = clean_text(text, default)
    return "" if normalize_key(value) in {"undefined", "none", "nan", "null"} else value

def clean_display_label(value: Any, default: str = "") -> str:
    """Limpa rótulos de eixos, legendas, tabelas e títulos para nunca exibir undefined/null."""
    txt = clean_text(value, default)
    if normalize_key(txt) in {"undefined", "none", "nan", "null", ""}:
        return default
    return txt


def _deep_clean_plotly_payload(obj: Any, path: tuple = ()) -> Any:
    """Limpa o JSON final do Plotly antes de renderizar.

    O `undefined` que ainda aparecia em alguns gráficos vinha de pontos mais
    profundos do payload serializado (ex.: layout.title.text, legend title,
    arrays de label/hover/customdata). A limpeza por trace não pegava 100%
    desses casos. Esta função percorre o JSON final e remove qualquer token
    técnico antes do Streamlit/Plotly renderizar.
    """
    if isinstance(obj, dict):
        return {k: _deep_clean_plotly_payload(v, path + (str(k),)) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_deep_clean_plotly_payload(v, path) for v in obj]
    if isinstance(obj, tuple):
        return tuple(_deep_clean_plotly_payload(v, path) for v in obj)
    if isinstance(obj, str):
        raw = obj.replace('\u00a0', ' ').strip()
        low = raw.lower()
        key = path[-1].lower() if path else ''
        joined = '.'.join(path).lower()

        # Títulos/eixos/legendas: se vierem vazios ou técnicos, ficam realmente sem título.
        title_like = key in {'title', 'name'} or joined.endswith('title.text') or joined.endswith('legend.title.text')
        template_like = key in {'hovertemplate', 'texttemplate'}

        if low in BAD_TEXT_TOKENS or low in {'undefined', 'none', 'nan', 'null'}:
            return '' if title_like or template_like else 'Não informado'

        if template_like:
            return clean_plotly_template(raw)

        if BAD_TEXT_PATTERN.search(raw):
            cleaned = BAD_TEXT_PATTERN.sub('', raw)
            cleaned = re.sub(r'\s*[-–—|/\\]+\s*', ' - ', cleaned)
            cleaned = re.sub(r'(?:^\s*[-–—|/\\]+\s*|\s*[-–—|/\\]+\s*$)', '', cleaned).strip()
            cleaned = re.sub(r'\s+', ' ', cleaned)
            if not cleaned:
                return '' if title_like else 'Não informado'
            return cleaned
        return raw
    try:
        if pd.isna(obj):
            return 'Não informado'
    except Exception:
        pass
    return obj

def strip_plotly_undefined(fig: go.Figure) -> go.Figure:
    """Reconstrói a figura a partir do JSON já limpo para eliminar `undefined`."""
    if fig is None:
        return fig
    try:
        payload = fig.to_plotly_json()
        payload = _deep_clean_plotly_payload(payload)
        # Força ausência real de título no layout, evitando fallback JS como `undefined`.
        payload.setdefault('layout', {})
        payload['layout']['title'] = {'text': ''}
        for axis_name in ['xaxis', 'yaxis', 'xaxis2', 'yaxis2', 'xaxis3', 'yaxis3']:
            axis = payload['layout'].setdefault(axis_name, {}) if axis_name in payload.get('layout', {}) else payload['layout'].get(axis_name)
            if isinstance(axis, dict):
                title = axis.get('title')
                if title is None or isinstance(title, str):
                    axis['title'] = {'text': '' if normalize_key(title) in BAD_TEXT_TOKENS else clean_text(title, '')}
                elif isinstance(title, dict) and normalize_key(title.get('text')) in BAD_TEXT_TOKENS:
                    title['text'] = ''
        if isinstance(payload.get('layout', {}).get('legend'), dict):
            payload['layout']['legend']['title'] = {'text': ''}
        return go.Figure(payload)
    except Exception:
        return fig


def sanitize_dataframe_display(df: pd.DataFrame) -> pd.DataFrame:
    """Prepara tabelas para exibição sem cabeçalhos/valores técnicos como undefined."""
    if df is None:
        return pd.DataFrame()
    out = df.copy()
    new_cols = []
    used = set()
    for idx, col in enumerate(out.columns):
        label = clean_display_label(col, f"Campo {idx + 1}")
        if not label:
            label = f"Campo {idx + 1}"
        base = label
        counter = 2
        while label in used:
            label = f"{base} {counter}"
            counter += 1
        used.add(label)
        new_cols.append(label)
    out.columns = new_cols
    try:
        out = out.replace({"undefined": "Não informado", "Undefined": "Não informado", "UNDEFINED": "Não informado", "None": "Não informado", "none": "Não informado", "nan": "Não informado", "NaN": "Não informado", "null": "Não informado", "NULL": "Não informado"})
        for col in out.columns:
            if out[col].dtype == "object" or str(out[col].dtype).startswith("string"):
                out[col] = out[col].map(lambda v: clean_text(v, "Não informado"))
    except Exception:
        pass
    return out

def first_existing(df: pd.DataFrame, candidates: Iterable[str]) -> Optional[str]:
    lower_map = {c.lower(): c for c in df.columns}
    for c in candidates:
        if c in df.columns: return c
        if c.lower() in lower_map: return lower_map[c.lower()]
    return None

def norm_col_name(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())

def find_column(df: pd.DataFrame, canonical: str, extra: Optional[Iterable[str]] = None) -> Optional[str]:
    """Procura uma coluna usando alias exato e fuzzy leve."""
    if df is None or df.empty:
        return None
    candidates = list(COLUMN_ALIASES.get(canonical, [])) + list(extra or [])
    found = first_existing(df, candidates)
    if found:
        return found
    normalized = {norm_col_name(c): c for c in df.columns}
    for cand in candidates + [canonical]:
        key = norm_col_name(cand)
        if key in normalized:
            return normalized[key]
    return None

def get_table(tables: Dict[str, pd.DataFrame], canonical: str) -> Optional[pd.DataFrame]:
    """Localiza abas mesmo quando o usuário altera levemente o nome."""
    for name in SHEET_ALIASES.get(canonical, [canonical]):
        if name in tables and tables[name] is not None and not tables[name].empty:
            return tables[name]
    normalized = {norm_col_name(k): k for k in tables.keys()}
    for name in SHEET_ALIASES.get(canonical, [canonical]):
        key = norm_col_name(name)
        if key in normalized:
            table = tables[normalized[key]]
            if table is not None and not table.empty:
                return table
    return None


def _series_has_useful_values(series: pd.Series) -> bool:
    try:
        normalized = series.map(lambda x: clean_text(x, "")).astype(str).str.strip()
        return bool((normalized != "").any())
    except Exception:
        return False


def looks_like_identifier(value: Any) -> bool:
    """Detecta CNPJ/códigos puros para evitar uso como rótulo principal do usuário."""
    txt = clean_text(value, "")
    if not txt:
        return True
    digits = re.sub(r"\D", "", txt)
    letters = re.sub(r"[^A-Za-zÀ-ÿ]", "", txt)
    # CNPJ/CPF ou sequências numéricas longas não são boas para exibição principal.
    if len(digits) >= 11 and len(letters) <= 3:
        return True
    # IDs técnicos como 12345, F00012, CNPJ mascarado, etc.
    if len(digits) >= 5 and len(letters) <= 2:
        return True
    return False


def prefer_display_series(primary: pd.Series, fallback: pd.Series, default: str) -> pd.Series:
    """Combina duas séries priorizando nomes amigáveis e evitando CNPJ/código como rótulo."""
    pser = primary if primary is not None else pd.Series([default] * len(fallback), index=fallback.index)
    fser = fallback if fallback is not None else pd.Series([default] * len(pser), index=pser.index)
    out = []
    for pval, fval in zip(pser.tolist(), fser.tolist()):
        ptxt = clean_text(pval, "")
        ftxt = clean_text(fval, "")
        if ptxt and not looks_like_identifier(ptxt):
            out.append(ptxt)
        elif ftxt and not looks_like_identifier(ftxt):
            out.append(ftxt)
        elif ptxt:
            out.append(ptxt)
        elif ftxt:
            out.append(ftxt)
        else:
            out.append(default)
    return pd.Series(out, index=pser.index if hasattr(pser, "index") else None).map(lambda x: clean_text(x, default))


def coalesce_display_column(df: pd.DataFrame, target: str, candidates: List[str], default: str) -> pd.DataFrame:
    """Garante target amigável usando candidates, sem deixar CNPJ/códigos dominarem quando há nome."""
    base = df[target] if target in df.columns else pd.Series([default] * len(df), index=df.index)
    result = base
    for cand in candidates:
        if cand in df.columns:
            result = prefer_display_series(df[cand], result, default)
    df[target] = result.map(lambda x: clean_text(x, default))
    return df


def assign_canonical_column(df: pd.DataFrame, canonical: str, default: Any = None, numeric: bool = False) -> pd.DataFrame:
    """Preenche uma coluna canônica a partir de aliases, sem sobrescrever dado útil.

    Isso evita que pequenas mudanças de nome na planilha derrubem gráficos/filtros.
    """
    existing_ok = canonical in df.columns and _series_has_useful_values(df[canonical])
    found = find_column(df, canonical)
    if found is not None and (not existing_ok or found != canonical):
        df[canonical] = to_number_safe(df[found]) if numeric else df[found]
    elif canonical not in df.columns and default is not None:
        df[canonical] = default
    if numeric and canonical in df.columns:
        df[canonical] = to_number_safe(df[canonical])
    return df


def to_number_safe(series: pd.Series) -> pd.Series:
    """Converte valores BR/US para número sem quebrar gráficos.

    Aceita exemplos como: R$ 1.234,56, 1,234.56, 1234.56, -1.234,56 e
    números já tipados. Quando não reconhece, retorna 0.0 para manter o app estável.
    """
    if series is None:
        return pd.Series(dtype="float64")
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0.0)

    def parse_one(value: Any) -> float:
        txt = clean_text(value, "")
        if not txt:
            return 0.0
        txt = re.sub(r"[^0-9,.-]", "", txt)
        if not txt or txt in {"-", ",", "."}:
            return 0.0
        # Se houver vírgula e ponto, o último separador costuma ser o decimal.
        if "," in txt and "." in txt:
            if txt.rfind(",") > txt.rfind("."):
                txt = txt.replace(".", "").replace(",", ".")
            else:
                txt = txt.replace(",", "")
        elif "," in txt:
            txt = txt.replace(".", "").replace(",", ".")
        try:
            return float(txt)
        except Exception:
            return 0.0

    return series.map(parse_one).astype("float64").fillna(0.0)

def to_datetime_safe(series: pd.Series) -> pd.Series:
    """Converte datas BR e ISO sem deslocar mês/dia.

    Evita o caso comum em que `2026-02-01` vira janeiro quando `dayfirst=True`.
    """
    if series is None:
        return pd.Series(dtype="datetime64[ns]")

    def parse_one(value: Any) -> Any:
        txt = clean_text(value, "")
        if not txt:
            return pd.NaT
        # YYYY-MM ou YYYY/MM
        if re.fullmatch(r"\d{4}[-/]\d{1,2}", txt):
            txt = txt.replace("/", "-") + "-01"
            return pd.to_datetime(txt, errors="coerce", yearfirst=True)
        # YYYY-MM-DD ou YYYY/MM/DD
        if re.fullmatch(r"\d{4}[-/]\d{1,2}[-/]\d{1,2}", txt):
            return pd.to_datetime(txt.replace("/", "-"), errors="coerce", yearfirst=True)
        return pd.to_datetime(txt, errors="coerce", dayfirst=True)

    return series.map(parse_one)

@st.cache_data(show_spinner=False)
def read_excel_bytes(raw: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(BytesIO(raw))
    return {sheet: pd.read_excel(BytesIO(raw), sheet_name=sheet) for sheet in xls.sheet_names}


def canonical_sheet_name(sheet_name: str) -> str:
    """Agrupa abas equivalentes para permitir upload de arquivos de meses diferentes."""
    key = norm_col_name(sheet_name)
    for canonical, aliases in SHEET_ALIASES.items():
        if key in {norm_col_name(a) for a in aliases + [canonical]}:
            return canonical
    return sheet_name


@st.cache_data(show_spinner=False, ttl=900)
def combine_excel_payloads(payloads: List[Tuple[bytes, str]]) -> Tuple[Dict[str, pd.DataFrame], str]:
    """Lê uma ou várias planilhas e consolida abas equivalentes.

    Ex.: Janeiro.xlsx + Fevereiro.xlsx com aba `Financeiro` viram uma única
    tbFinanceiro. Isso deixa o dashboard pronto para uploads mensais ou base
    completa substituída pela interface.
    """
    combined: Dict[str, List[pd.DataFrame]] = {}
    source_names: List[str] = []
    for raw, name in payloads:
        if not raw:
            continue
        source_names.append(name)
        tables = read_excel_bytes(raw)
        for sheet, table in tables.items():
            if table is None or table.empty:
                continue
            canon = canonical_sheet_name(sheet)
            temp = table.copy()
            temp["Arquivo_Origem"] = name
            temp["Aba_Origem"] = sheet
            combined.setdefault(canon, []).append(temp)
    if not combined:
        raise FileNotFoundError("Nenhuma planilha Excel válida foi encontrada.")
    out = {sheet: pd.concat(frames, ignore_index=True, sort=False) for sheet, frames in combined.items()}
    label = ", ".join(source_names[:3]) + (f" +{len(source_names)-3}" if len(source_names) > 3 else "")
    return out, label


def load_default_payloads() -> List[Tuple[bytes, str]]:
    """Carrega todas as planilhas da pasta uploads; se não houver, usa a base exemplo."""
    upload_files = sorted(
        [p for p in UPLOADS_DIR.glob("*.xls*") if not p.name.startswith("~$")],
        key=lambda p: p.name.lower(),
    )
    if upload_files:
        return [(p.read_bytes(), p.name) for p in upload_files]
    for path in DEFAULT_FILES:
        if path.exists():
            return [(path.read_bytes(), path.name)]
    raise FileNotFoundError("Nenhuma planilha Excel encontrada em /uploads.")


def load_default_bytes() -> Tuple[bytes, str]:
    """Compatibilidade: retorna o primeiro arquivo padrão disponível."""
    payloads = load_default_payloads()
    return payloads[0]


def save_uploaded_files(uploaded_files) -> List[Tuple[bytes, str]]:
    if uploaded_files is None:
        return []
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]
    saved: List[Tuple[bytes, str]] = []
    for uploaded in uploaded_files:
        raw = uploaded.getvalue()
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", uploaded.name)
        (UPLOADS_DIR / safe).write_bytes(raw)
        saved.append((raw, safe))
    return saved


def save_uploaded_file(uploaded) -> Tuple[bytes, str]:
    """Compatibilidade com versões antigas do fluxo de upload."""
    saved = save_uploaded_files([uploaded])
    if not saved:
        raise ValueError("Arquivo não enviado.")
    return saved[0]

# -----------------------------
# Data model
# -----------------------------
@st.cache_data(show_spinner=False, ttl=900)
def enrich_model(tables: Dict[str, pd.DataFrame]) -> Tuple[pd.DataFrame, List[str]]:
    issues: List[str] = []
    fin = get_table(tables, "tbFinanceiro")
    if fin is None or fin.empty:
        for name, d in tables.items():
            if find_column(d, VALUE_COL):
                fin = d.copy(); issues.append(f"A aba tbFinanceiro não foi encontrada; usando {name}."); break
    if fin is None or fin.empty: raise ValueError("Não encontrei uma aba financeira válida.")

    df = fin.copy()
    df["Fornecedor_ID"] = df.get("Fornecedor_ID", pd.Series([None] * len(df))).astype(str)
    df["Servico_ID"] = df.get("Servico_ID", pd.Series([None] * len(df))).astype(str)
    df["Financeiro_Contrato_ID"] = df.get("Financeiro_Contrato_ID", df.get("Contrato_ID", pd.Series([None] * len(df)))).astype(str)

    # Normalização canônica já na aba financeira. Assim o painel continua funcionando
    # quando a planilha traz Operadora em vez de Fornecedor, Valor Pago em vez de
    # Valor_Realizado, etc. As dimensões enriquecidas por cadastro ainda prevalecem
    # quando houver tabelas auxiliares.
    assign_canonical_column(df, VALUE_COL, default=0.0, numeric=True)
    assign_canonical_column(df, DIFF_COL, default=0.0, numeric=True)
    for canonical, default in [
        (SUPPLIER_COL, "Sem fornecedor"), (SERVICE_COL, "Sem serviço"),
        (CATEGORY_COL, "Sem categoria"), (CONTRACT_COL, "Sem contrato"),
        (INVOICE_COL, "Sem fatura"), (BRANCH_COL, "Sem filial"), (REGION_COL, "Sem região"),
        (CC_COL, "Sem centro de custo"), (CC_ID_COL, "Sem código"),
        (STATUS_COL, "Não informado"), (CONTESTED_COL, "Não informado"),
    ]:
        assign_canonical_column(df, canonical, default=default)
    required_dynamic = [(VALUE_COL, "valor"), (SUPPLIER_COL, "operadora/fornecedor")]
    for col, label in required_dynamic:
        if col not in df.columns or not _series_has_useful_values(df[col]):
            issues.append(f"Não encontrei uma coluna útil de {label}; alguns gráficos podem ficar limitados.")

    ref_col = choose_reference_date_column(df)
    due_col = choose_due_date_column(df)
    date_col = choose_period_date_column(df)

    df[REF_DATE_COL] = to_datetime_safe(df[ref_col]) if ref_col else pd.NaT
    df[DUE_DATE_COL] = to_datetime_safe(df[due_col]) if due_col else pd.NaT
    if df[REF_DATE_COL].isna().all() and ref_col:
        df[REF_DATE_COL] = to_datetime_safe(df[ref_col].astype(str) + "-01")
    if df[DUE_DATE_COL].isna().all() and due_col:
        df[DUE_DATE_COL] = to_datetime_safe(df[due_col].astype(str) + "-01")

    if df[DUE_DATE_COL].notna().any():
        df[DATE_COL] = df[DUE_DATE_COL]
        df["Periodo_Origem"] = due_col or "Vencimento"
    elif df[REF_DATE_COL].notna().any():
        df[DATE_COL] = df[REF_DATE_COL]
        df["Periodo_Origem"] = ref_col or "Mês de referência"
    else:
        df[DATE_COL] = to_datetime_safe(df[date_col]) if date_col else pd.NaT
        df["Periodo_Origem"] = date_col or "Não encontrada"

    df[REF_MONTH_COL] = df[REF_DATE_COL].map(format_period_label)
    df[DUE_MONTH_COL] = df[DUE_DATE_COL].map(format_period_label)
    df[MONTH_COL] = df[DATE_COL].map(format_period_label)
    df["Periodo_Origem_Referencia"] = ref_col or "Não encontrada"
    df["Periodo_Origem_Vencimento"] = due_col or "Não encontrada"

    df[INVOICE_COL] = df[INVOICE_COL].map(lambda x: clean_text(x, "Sem fatura"))
    df[CONTESTED_COL] = df[CONTESTED_COL].map(lambda x: clean_text(x, "Não informado"))
    df[STATUS_COL] = df[STATUS_COL].map(lambda x: clean_text(x, "Não informado"))
    df[CONTRACT_COL] = df[CONTRACT_COL].map(lambda x: clean_text(x, "Sem contrato"))

    # Suppliers
    # Exibição amigável: prioriza nome fantasia/operadora/razão social e evita CNPJ/código
    # como rótulo principal dos gráficos. O ID/CNPJ fica apenas como apoio de auditoria.
    supplier_financial = df[SUPPLIER_COL].copy() if SUPPLIER_COL in df.columns else pd.Series(["Sem fornecedor"] * len(df), index=df.index)
    forn = get_table(tables, "tbFornecedor")
    if forn is not None and not forn.empty and "Fornecedor_ID" in forn.columns:
        use = forn.copy(); use["Fornecedor_ID"] = use["Fornecedor_ID"].astype(str)
        name = first_existing(use, [
            "Fornecedor_Nome_Fantasia", "Operadora", "Nome_Fantasia", "Nome Fantasia",
            "Fornecedor_Razao_Social", "Razao_Social", "Razão Social", "Fornecedor", "Nome"
        ])
        if name:
            df["Fornecedor_ID"] = df["Fornecedor_ID"].astype(str)
            dim = use[["Fornecedor_ID", name]].drop_duplicates("Fornecedor_ID").rename(columns={name: "__Fornecedor_Nome_Dim"})
            df = df.merge(dim, on="Fornecedor_ID", how="left")
    dim_supplier = df["__Fornecedor_Nome_Dim"] if "__Fornecedor_Nome_Dim" in df.columns else pd.Series([""] * len(df), index=df.index)
    df[SUPPLIER_COL] = prefer_display_series(dim_supplier, supplier_financial, "Sem fornecedor")
    # Fallback final: só usa Fornecedor_ID quando realmente não houver nome amigável.
    df[SUPPLIER_COL] = prefer_display_series(df[SUPPLIER_COL], df["Fornecedor_ID"], "Sem fornecedor")
    df[SUPPLIER_COL] = df[SUPPLIER_COL].map(lambda x: clean_text(x, "Sem fornecedor"))
    for col in ["__Fornecedor_Nome_Dim"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Services
    service_financial = df[SERVICE_COL].copy() if SERVICE_COL in df.columns else pd.Series(["Sem serviço"] * len(df), index=df.index)
    category_financial = df[CATEGORY_COL].copy() if CATEGORY_COL in df.columns else pd.Series(["Sem categoria"] * len(df), index=df.index)
    serv = get_table(tables, "tbServicos")
    if serv is not None and not serv.empty and "Servico_ID" in serv.columns:
        use = serv.copy(); use["Servico_ID"] = use["Servico_ID"].astype(str)
        cols, rename = ["Servico_ID"], {}
        for src, dst in [
            ("Servico_Descricao", "__Servico_Nome_Dim"), ("Servico_Nome", "__Servico_Nome_Dim"), ("Descricao", "__Servico_Nome_Dim"),
            ("Servico_C_Categoria", "__Categoria_Dim"), ("Categoria", "__Categoria_Dim"),
            ("Servico_C_Tipo", "Tipo_Servico"), ("Servico_Valor", "Valor_Contratado"),
            ("Centro_deCusto_ID", "__Centro_deCusto_ID_Dim"), ("Filial_ID", "__Filial_ID_Dim"), ("Contrato_ID", "Contrato_ID_Servico")
        ]:
            if src in use.columns and src not in cols:
                cols.append(src); rename[src] = dst
        df = df.merge(use[cols].drop_duplicates("Servico_ID").rename(columns=rename), on="Servico_ID", how="left")
    service_dim = df["__Servico_Nome_Dim"] if "__Servico_Nome_Dim" in df.columns else pd.Series([""] * len(df), index=df.index)
    category_dim = df["__Categoria_Dim"] if "__Categoria_Dim" in df.columns else pd.Series([""] * len(df), index=df.index)
    df[SERVICE_COL] = prefer_display_series(service_dim, service_financial, "Sem serviço")
    df[SERVICE_COL] = prefer_display_series(df[SERVICE_COL], df["Servico_ID"], "Sem serviço")
    df[CATEGORY_COL] = prefer_display_series(category_dim, category_financial, "Sem categoria")
    df[SERVICE_COL] = df[SERVICE_COL].map(lambda x: clean_text(x, "Sem serviço"))
    df[CATEGORY_COL] = df[CATEGORY_COL].map(lambda x: clean_text(x, "Sem categoria"))
    if "Valor_Contratado" in df.columns: df["Valor_Contratado"] = to_number_safe(df["Valor_Contratado"])
    # Complementa IDs vindos do cadastro de serviços sem quebrar nomes de colunas existentes.
    if "__Filial_ID_Dim" in df.columns:
        if "Filial_ID" not in df.columns:
            df["Filial_ID"] = df["__Filial_ID_Dim"]
        else:
            df["Filial_ID"] = prefer_display_series(df["Filial_ID"], df["__Filial_ID_Dim"], "")
    if "__Centro_deCusto_ID_Dim" in df.columns:
        if "Centro_deCusto_ID" not in df.columns:
            df["Centro_deCusto_ID"] = df["__Centro_deCusto_ID_Dim"]
        else:
            df["Centro_deCusto_ID"] = prefer_display_series(df["Centro_deCusto_ID"], df["__Centro_deCusto_ID_Dim"], "")
    for col in ["__Servico_Nome_Dim", "__Categoria_Dim", "__Filial_ID_Dim", "__Centro_deCusto_ID_Dim"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Branch / region / cost center
    branch_financial = df[BRANCH_COL].copy() if BRANCH_COL in df.columns else pd.Series(["Sem filial"] * len(df), index=df.index)
    region_financial = df[REGION_COL].copy() if REGION_COL in df.columns else pd.Series(["Sem região"] * len(df), index=df.index)
    filial = get_table(tables, "tbFilial")
    if filial is not None and not filial.empty and "Filial_ID" in filial.columns and "Filial_ID" in df.columns:
        use = filial.copy(); use["Filial_ID"] = use["Filial_ID"].astype(str)
        name = first_existing(use, ["Filial_Nome", "Filial_Descricao", "Filial", "Unidade", "Site", "Nome"])
        region_name = first_existing(use, ["Regiao", "Região", "Regional", "Macro_Regiao", "Macro Região", "UF", "Estado", "Cidade", "Municipio", "Município"])
        cols = ["Filial_ID"]
        rename = {}
        if name:
            cols.append(name); rename[name] = "__Filial_Nome_Dim"
        if region_name and region_name not in cols:
            cols.append(region_name); rename[region_name] = "__Regiao_Nome_Dim"
        if len(cols) > 1:
            df["Filial_ID"] = df["Filial_ID"].astype(str)
            df = df.merge(use[cols].drop_duplicates("Filial_ID").rename(columns=rename), on="Filial_ID", how="left")
    dim_branch = df["__Filial_Nome_Dim"] if "__Filial_Nome_Dim" in df.columns else pd.Series([""] * len(df), index=df.index)
    dim_region = df["__Regiao_Nome_Dim"] if "__Regiao_Nome_Dim" in df.columns else pd.Series([""] * len(df), index=df.index)
    df[BRANCH_COL] = prefer_display_series(dim_branch, branch_financial, "Sem filial")
    df[REGION_COL] = prefer_display_series(dim_region, region_financial, "Sem região")
    df[BRANCH_COL] = df[BRANCH_COL].map(lambda x: clean_text(x, "Sem filial"))
    df[REGION_COL] = df[REGION_COL].map(lambda x: clean_text(x, "Sem região"))
    for col in ["__Filial_Nome_Dim", "__Regiao_Nome_Dim"]:
        if col in df.columns:
            df.drop(columns=[col], inplace=True)

    # Centro de custo: usa a descrição/nome no gráfico quando existir,
    # mas preserva o código para auditoria no detalhes e na tabela.
    if "Centro_deCusto_ID" in df.columns and (CC_ID_COL not in df.columns or not _series_has_useful_values(df[CC_ID_COL])):
        df[CC_ID_COL] = df["Centro_deCusto_ID"].map(lambda x: clean_text(x, "Sem código"))
    elif CC_ID_COL not in df.columns:
        df[CC_ID_COL] = "Sem código"

    cc = get_table(tables, "tbCentroCusto")
    if cc is not None and not cc.empty and "Centro_deCusto_ID" in cc.columns and "Centro_deCusto_ID" in df.columns:
        use = cc.copy(); use["Centro_deCusto_ID"] = use["Centro_deCusto_ID"].astype(str)
        name = first_existing(use, [
            "Centro_deCusto_Descricao", "Centro_deCusto_Nome", "Centro_Custo_Descricao",
            "Centro_Custo_Nome", "Centro_Custo", "Centro de Custo", "Descricao", "Descrição", "Nome"
        ])
        if name:
            df["Centro_deCusto_ID"] = df["Centro_deCusto_ID"].astype(str)
            df = df.merge(
                use[["Centro_deCusto_ID", name]].drop_duplicates("Centro_deCusto_ID").rename(columns={name: "__CC_Nome_Dim"}),
                on="Centro_deCusto_ID",
                how="left",
            )
    cc_financial = df[CC_COL].copy() if CC_COL in df.columns else pd.Series(["Sem centro de custo"] * len(df), index=df.index)
    cc_dim = df["__CC_Nome_Dim"] if "__CC_Nome_Dim" in df.columns else pd.Series([""] * len(df), index=df.index)
    df[CC_COL] = prefer_display_series(cc_dim, cc_financial, "Sem centro de custo")
    # Se não houver descrição real, usa o código, mas apenas como fallback.
    df[CC_COL] = prefer_display_series(df[CC_COL], df[CC_ID_COL], "Sem centro de custo")
    df[CC_COL] = df[CC_COL].map(lambda x: clean_text(x, "Sem centro de custo"))
    df[CC_ID_COL] = df[CC_ID_COL].map(lambda x: clean_text(x, "Sem código"))
    if "__CC_Nome_Dim" in df.columns:
        df.drop(columns=["__CC_Nome_Dim"], inplace=True)

    df["Linha_ID"] = range(1, len(df) + 1)
    df = ensure_model_quality(df)
    return df, issues

# -----------------------------
# State
# -----------------------------
def init_state() -> None:
    st.session_state.setdefault("panel_filters", {})
    st.session_state.setdefault("detail", None)
    st.session_state.setdefault("last_selection", {})
    st.session_state.setdefault("visual_scopes", {})
    st.session_state.setdefault("filter_version", 0)
    st.session_state.setdefault("filter_reset_token", 0)
    st.session_state.setdefault(PERIOD_BASIS_KEY, "Vencimento")
    st.session_state.setdefault("ai_opening", False)
    st.session_state.setdefault("ai_busy", False)


def _manual_filter_key(field: str) -> str:
    """Chave do widget de filtro manual.

    Importante: o Streamlit não permite alterar st.session_state de um
    widget depois que ele foi instanciado no mesmo ciclo. Por isso a chave
    recebe um token de reset. Para limpar selectboxes, incrementamos o token
    e forçamos um rerun, criando novos widgets já no valor padrão.
    """
    token = st.session_state.get("filter_reset_token", 0)
    return f"manual_{field}_{token}"


def _detail_filter_prefix() -> str:
    return "detail_filter_"


def _reset_filter_widget_keys(fields: Optional[List[str]] = None) -> None:
    """Agenda reset dos widgets de filtro sem alterar keys já instanciadas."""
    st.session_state.filter_reset_token = st.session_state.get("filter_reset_token", 0) + 1
    # Não apaga nem altera chaves de widgets já instanciados no ciclo atual.
    # A troca de token muda a key do próximo widget e evita erro de Session State.


def _clear_detail_filter_keys() -> None:
    for key in list(st.session_state.keys()):
        text_key = str(key)
        if text_key.startswith(_detail_filter_prefix()):
            st.session_state.pop(key, None)


def _clear_chart_and_selection_keys() -> None:
    prefixes = (
        "chart_", "plot_", "selected_", "selection_", "cross_", "click_",
        "last_click_", "active_chart_", "active_filter_", "clear_filter_",
        "remove_",
    )
    exact_keys = {"last_selection", "visual_scopes", "detail"}
    for key in list(st.session_state.keys()):
        text_key = str(key)
        if text_key in exact_keys or text_key.startswith(prefixes):
            st.session_state.pop(key, None)


def set_filter(field: str, value: Any, source: str) -> None:
    if field and value not in (None, "", "Todos"):
        old = st.session_state.panel_filters.get(field, {}).get("value")
        clean_value = clean_text(value)
        st.session_state.panel_filters[field] = {"value": clean_value, "source": source}
        # Não altere a key do widget aqui. O próprio selectbox controla o estado
        # manual; alterar st.session_state após a criação do widget gera
        # StreamlitAPIException em versões novas do Streamlit.
        if normalize_key(old) != normalize_key(clean_value):
            st.session_state.filter_version = st.session_state.get("filter_version", 0) + 1


def clear_filter(field: str, reset_widget: bool = True) -> None:
    changed = False
    if field in st.session_state.panel_filters:
        st.session_state.panel_filters.pop(field, None)
        changed = True
    if reset_widget:
        _reset_filter_widget_keys([field])
        changed = True
    # Remove detalhes/filtros derivados para evitar a sensação de filtro preso.
    st.session_state.pop("detail", None)
    st.session_state.pop("last_selection", None)
    if changed:
        st.session_state.filter_version = st.session_state.get("filter_version", 0) + 1


def clear_all_filters() -> None:
    """Limpa filtros manuais, filtros por visual e filtros internos.

    Esta função é usada como callback do botão, então roda antes da nova
    renderização da página. Isso impede que os selectboxes restaurem valores
    antigos e recriem os filtros logo depois da limpeza.
    """
    st.session_state.panel_filters = {}
    _reset_filter_widget_keys()
    _clear_detail_filter_keys()
    _clear_chart_and_selection_keys()
    st.session_state.last_selection = {}
    st.session_state.visual_scopes = {}
    st.session_state.detail = None
    st.session_state.filter_version = st.session_state.get("filter_version", 0) + 1

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    """Filtro global único: todos os KPIs, gráficos, tabelas, IA e detalhes usam esta saída."""
    out = ensure_model_quality(df)
    invalid: List[str] = []
    for field, meta in list(st.session_state.panel_filters.items()):
        if field not in out.columns:
            invalid.append(field)
            continue
        value = meta.get("value") if isinstance(meta, dict) else meta
        if value in (None, "", "Todos"):
            invalid.append(field)
            continue
        mask = out[field].map(normalize_key) == normalize_key(value)
        # Evita tela aparentemente quebrada quando o filtro antigo não existe mais na nova planilha.
        if not mask.any():
            invalid.append(field)
            continue
        out = out[mask].copy()
    for field in invalid:
        st.session_state.panel_filters.pop(field, None)
    return out

def build_detail_payload(
    title: str,
    subtitle: str,
    df: pd.DataFrame,
    group_field: Optional[str] = None,
    visual_key: str = "",
    visible_values: Optional[List[Any]] = None,
    selected_value: Optional[Any] = None,
) -> Dict[str, Any]:
    """Monta o payload do detalhamento usando exatamente o escopo exibido no visual."""
    detail_df = df.copy()
    clean_visible = [clean_text(v) for v in (visible_values or []) if clean_text(v, "") not in ("", "None", "nan")]
    if group_field and group_field in detail_df.columns and clean_visible:
        allowed = {normalize_key(v) for v in clean_visible}
        scoped = detail_df[detail_df[group_field].map(normalize_key).isin(allowed)].copy()
        if not scoped.empty:
            detail_df = scoped
    return {
        "title": title,
        "subtitle": subtitle,
        "data": detail_df,
        "source_rows": len(df),
        "visible_values": clean_visible,
        "selected_value": clean_text(selected_value, "") if selected_value not in (None, "") else "",
        "group_field": group_field,
        "visual_key": visual_key,
    }


def open_detail(
    title: str,
    subtitle: str,
    df: pd.DataFrame,
    group_field: Optional[str] = None,
    visual_key: str = "",
    visible_values: Optional[List[Any]] = None,
    selected_value: Optional[Any] = None,
) -> None:
    """Compatibilidade: mantém abertura antiga quando necessário."""
    st.session_state.detail = build_detail_payload(title, subtitle, df, group_field, visual_key, visible_values, selected_value)

def resolve_selected_value(data: pd.DataFrame, field: Optional[str], raw_value: Any) -> Optional[str]:
    """Resolve valor clicado mesmo quando o gráfico usa rótulo abreviado."""
    if not field or field not in data.columns or raw_value in (None, ""):
        return None
    valid_values = data[field].dropna().astype(str).unique().tolist()
    selected_norm = normalize_key(raw_value)
    for item in valid_values:
        candidates = {normalize_key(item)}
        for lim in (18, 22, 28, 34, 42, 72):
            candidates.add(normalize_key(short_label(item, lim)))
        if selected_norm in candidates:
            return item
    return clean_text(raw_value, "")

def visible_values_from_figure(fig: go.Figure, field: Optional[str]) -> List[Any]:
    """Extrai os itens que realmente aparecem no visual a partir do customdata do Plotly."""
    if not field:
        return []
    values: List[Any] = []
    try:
        for trace in fig.data:
            customdata = getattr(trace, "customdata", None)
            if customdata is not None:
                for row in customdata:
                    try:
                        value = row[0] if isinstance(row, (list, tuple)) else row
                    except Exception:
                        value = row
                    if value not in (None, ""):
                        values.append(value)
            else:
                # fallback para gráficos simples sem customdata. Evita usar "or"
                # com arrays/Series, pois isso pode gerar ambiguidade booleana.
                axis_values = getattr(trace, "y", None)
                if axis_values is None or (hasattr(axis_values, "__len__") and len(axis_values) == 0):
                    axis_values = getattr(trace, "x", None)
                if axis_values is not None:
                    values.extend([v for v in list(axis_values) if v not in (None, "")])
    except Exception:
        return []

    seen = set()
    out: List[Any] = []
    for value in values:
        key_norm = normalize_key(value)
        if key_norm and key_norm not in seen:
            seen.add(key_norm)
            out.append(value)
    return out


def _plain(value: Any) -> str:
    return html.escape(normalize_markdown(value), quote=True)


def inline_markdown_to_html(text: Any) -> str:
    """Converte markdown inline simples para HTML seguro em cards customizados."""
    safe = html.escape(normalize_markdown(text), quote=True)
    safe = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", safe)
    safe = re.sub(r"__(.+?)__", r"<strong>\1</strong>", safe)
    safe = safe.replace("\n", "<br>")
    return safe


def _chart_point_count(fig: go.Figure) -> int:
    count = 0
    for trace in getattr(fig, "data", []) or []:
        for attr in ("x", "y", "labels", "values", "z"):
            val = getattr(trace, attr, None)
            if val is None:
                continue
            try:
                if attr == "z" and hasattr(val, "__len__"):
                    count = max(count, sum(len(r) for r in val if hasattr(r, "__len__")))
                else:
                    count = max(count, len(val))
            except Exception:
                pass
    return count


def has_enough_visual_data(fig: go.Figure, data: pd.DataFrame, filter_field: Optional[str], key: str = "") -> Tuple[bool, str]:
    """Decide se vale renderizar gráfico/tabela ou apenas uma nota executiva.

    Evita telas com gráficos ocultos, vazios ou sem massa mínima para análise.
    """
    if data is None or data.empty:
        return False, "não há registros no filtro atual."
    if VALUE_COL in data.columns and float(data[VALUE_COL].sum()) == 0 and key not in {"trend"}:
        return False, "o valor financeiro está zerado no contexto atual."
    if key in {"trend", "heat_supplier_month"} and MONTH_COL in data.columns:
        periods = data[MONTH_COL].dropna().astype(str).nunique()
        if periods < 2:
            return False, "há apenas um período disponível; a tendência ou sazonalidade não seria confiável."
    if key in {"pareto", "anomaly_service", "treemap_supplier_service"} and SERVICE_COL in data.columns:
        if data[SERVICE_COL].dropna().astype(str).nunique() < 2:
            return False, "há poucos serviços distintos para formar uma comparação útil."
    if filter_field and filter_field in data.columns:
        if data[filter_field].dropna().astype(str).nunique() < 1:
            return False, f"a coluna {filter_field} não possui valores válidos."
    if _chart_point_count(fig) <= 0:
        return False, "o gráfico não encontrou pontos válidos para desenhar."
    return True, ""


def _main_finding(data: pd.DataFrame, group: Optional[str]) -> str:
    if data is None or data.empty or VALUE_COL not in data.columns:
        return "Sem massa de dados suficiente para apontar concentração."
    total = float(data[VALUE_COL].sum()) or 0.0
    if group and group in data.columns and total:
        rank = data.groupby(group, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
        if not rank.empty:
            return f"{clean_text(rank.index[0])} responde por {pct(float(rank.iloc[0]) / total * 100)} do valor exibido ({money(rank.iloc[0])})."
    if SUPPLIER_COL in data.columns and total:
        rank = data.groupby(SUPPLIER_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
        if not rank.empty:
            return f"{clean_text(rank.index[0])} é o maior bloco financeiro deste contexto ({money(rank.iloc[0])})."
    return f"O contexto soma {money(total)} em {len(data)} lançamento(s)."


def _visual_goal(title: str) -> str:
    """Objetivo específico do visual, sem texto genérico repetido."""
    t = normalize_key(title)
    if "fornecedor" in t or "operadora" in t:
        return "Identificar concentração por operadora e priorizar negociação."
    if "categoria" in t:
        return "Ver quais naturezas de gasto explicam o total do período."
    if "servi" in t or "pareto" in t:
        return "Apontar os serviços que mais pesam no custo mensal."
    if "contrato" in t:
        return "Localizar contratos relevantes para revisão, renovação ou auditoria."
    if "fatura" in t:
        return "Conferir faturas com maior exposição financeira ou diferença."
    if "diferen" in t or "variacao" in t or "varia" in t:
        return "Separar variações que merecem validação financeira."
    if "filial" in t or "unidade" in t:
        return "Mostrar onde o consumo/custo está concentrado por unidade."
    if "centro" in t or "custo" in t:
        return "Apoiar rateio, orçamento e cobrança por centro de custo."
    if "tend" in t or "evolu" in t or "period" in t:
        return "Acompanhar evolução mensal e antecipar desvios."
    if "anomalia" in t or "alerta" in t:
        return "Destacar pontos fora do padrão para conferência rápida."
    if "oportun" in t or "econom" in t:
        return "Priorizar oportunidades de economia com maior impacto."
    return "Transformar o visual em uma decisão prática de análise."


def _suggested_action(title: str, data: pd.DataFrame, group: Optional[str]) -> str:
    t = normalize_key(title)
    diff = float(data[DIFF_COL].sum()) if data is not None and not data.empty and DIFF_COL in data.columns else 0.0
    if "pareto" in t or "servi" in t:
        return "Priorize os itens do topo e valide contrato, fatura e recorrência."
    if "fornecedor" in t or "operadora" in t:
        return "Comece pela maior operadora e compare dependência, reajustes e alternativas."
    if "categoria" in t:
        return "Confira a categoria dominante e veja se há cobrança fora do perfil esperado."
    if "diferen" in t or abs(diff) > 0:
        return "Investigue as maiores diferenças antes de aprovar ou contestar a fatura."
    if "contrato" in t:
        return "Revise vencimento, recorrência e vínculo com serviços de maior valor."
    if "fatura" in t:
        return "Abra as faturas de maior valor e confirme período, serviço e contrato."
    if "anomalia" in t or "alerta" in t:
        return "Valide picos contra histórico, consumo e reajustes aplicados."
    if "oportun" in t or "econom" in t:
        return "Priorize a oportunidade com maior valor e menor esforço de execução."
    return "Clique no visual para filtrar e abrir os lançamentos que sustentam a leitura."

def render_visual_action_context(title: str, desc: str, data: pd.DataFrame, filter_field: Optional[str]) -> None:
    """Leitura objetiva e específica para cada visual, sem repetir textos genéricos."""
    total = float(data[VALUE_COL].sum()) if data is not None and not data.empty and VALUE_COL in data.columns else 0.0
    linhas = len(data) if data is not None else 0
    finding = _main_finding(data, filter_field)
    action = _suggested_action(title, data, filter_field)
    goal = _visual_goal(title)
    emphasis = ""
    if data is not None and not data.empty and filter_field and filter_field in data.columns and VALUE_COL in data.columns:
        try:
            rank = data.groupby(filter_field, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
            share = float(rank.iloc[0] / total * 100) if total and not rank.empty else 0
            if share >= 50:
                emphasis = "risk-card"
            elif share >= 30:
                emphasis = "attention-card"
            elif "oportun" in normalize_key(title) or "econom" in normalize_key(title):
                emphasis = "opportunity-card"
        except Exception:
            emphasis = ""
    st.markdown(f"""
    <div class='visual-analysis {emphasis}'>
      <span class='badge'>Leitura rápida</span>
      <p><strong>Objetivo:</strong> {_plain(goal)}</p>
      <p><strong>Destaque:</strong> {_plain(finding)}</p>
      <p><strong>Próxima ação:</strong> {_plain(action)}</p>
      <small>{_plain(str(linhas))} lançamento(s) · {_plain(money(total))}</small>
    </div>
    """, unsafe_allow_html=True)

def render_insufficient_visual_info(title: str, reason: str, data: pd.DataFrame, filter_field: Optional[str]) -> None:
    total = float(data[VALUE_COL].sum()) if data is not None and not data.empty and VALUE_COL in data.columns else 0.0
    info = _main_finding(data, filter_field) if data is not None and not data.empty else "Sem registros disponíveis para este filtro."
    detalhe = "Abra o bloco abaixo para ver tabelas, contratos, faturas e registros disponíveis." if data is not None and not data.empty else "Ajuste os filtros ou carregue mais dados para habilitar o visual."
    st.markdown(
        f"<div class='insufficient-card'><strong>Visual omitido:</strong> {_plain(title)} não foi exibido porque {_plain(reason)}<br>"
        f"<strong>Informação útil:</strong> {_plain(info)} Total do contexto: {_plain(money(total))}.<br>"
        f"<strong>Próximo passo:</strong> {_plain(detalhe)}</div>",
        unsafe_allow_html=True,
    )





_PLOT_RENDER_COUNTER = 0


def _next_plot_key(base_key: str) -> str:
    """Gera chave única para cada gráfico Plotly renderizado na execução atual."""
    global _PLOT_RENDER_COUNTER
    _PLOT_RENDER_COUNTER += 1
    safe_base = normalize_key(str(base_key or "visual")).replace(" ", "_")[:80] or "visual"
    return f"plot_{safe_base}_{_PLOT_RENDER_COUNTER}"


def safe_plotly_chart(fig: go.Figure, key: str, height: int = 460) -> None:
    """Renderização única e estável para todos os gráficos Plotly.

    Centraliza width, config, sanitização e principalmente a geração de `key`
    única. Isso evita StreamlitDuplicateElementId quando o mesmo gráfico aparece
    no card principal e nos detalhes expansíveis.
    """
    if fig is None:
        st.info("ℹ️ Visual indisponível para o filtro atual.")
        return
    plot_key = _next_plot_key(key)
    try:
        fig = sanitize_figure(fig)
        fig.update_layout(
            height=max(420, int(height or 460)),
            autosize=True,
            margin=dict(t=26, b=78, l=58, r=96),
            title={"text": ""},
            legend_title_text="",
        )
        # Evita rótulos cortados em barras horizontais/verticais.
        for trace in fig.data:
            trace_type = getattr(trace, "type", "")
            if trace_type == "bar":
                try:
                    trace.update(cliponaxis=False)
                except Exception:
                    pass
        fig = strip_plotly_undefined(fig)
        st.plotly_chart(
            fig,
            width="stretch",
            config={"displayModeBar": False, "responsive": True},
            key=plot_key,
        )
    except TypeError:
        # Compatibilidade com Streamlit antigo que ainda não aceita width.
        st.plotly_chart(fig, config={"displayModeBar": False, "responsive": True}, key=plot_key)
    except Exception as exc:
        st.warning(f"Não foi possível renderizar este visual. Motivo técnico: {clean_text(exc)}")


def render_clickable_plotly(fig: go.Figure, key: str) -> list[dict]:
    """Compatibilidade: mantém a função antiga sem depender de eventos de clique.

    Alguns trechos antigos ainda podem chamar esta função. Ela agora apenas renderiza
    com segurança e retorna lista vazia, evitando NameError e falhas da dependência
    opcional streamlit-plotly-events.
    """
    safe_plotly_chart(fig, key, int(getattr(fig.layout, "height", 460) or 460))
    return []


def selected_value_from_points(points: Any) -> Optional[str]:
    """Compatibilidade com versões anteriores: sem evento de clique, não há seleção."""
    return None


def render_filter_hint(filter_field: Optional[str]) -> None:
    if filter_field:
        st.caption("🔎 Use os filtros laterais para cruzar o painel. Abaixo do visual há evidências, contratos, faturas e composição dos dados.")


def render_active_filter_chip(filter_field: Optional[str], key: str) -> None:
    if not filter_field or filter_field not in st.session_state.panel_filters:
        return
    active_value = st.session_state.panel_filters.get(filter_field, {}).get("value")
    if not active_value:
        return
    c_info, c_clear = st.columns([.82, .18], vertical_alignment="center")
    c_info.markdown(f"<span class='filter-chip'>🎯 Filtro ativo: {_plain(filter_field)} = {_plain(active_value)}</span>", unsafe_allow_html=True)
    if c_clear.button("Limpar", key=f"clear_filter_{key}", width="stretch"):
        clear_filter(filter_field)
        st.rerun()


def render_native_chart(title: str, desc: str, fig: go.Figure, data: pd.DataFrame, filter_field: Optional[str], key: str, height: int = 460, show_details: bool = True) -> None:
    """Renderiza gráfico em card analítico, com detalhes auditáveis e sem dependência frágil de clique.

    Padrão usado em Diretoria e Analítica:
    - título e descrição limpos;
    - resumo executivo compacto;
    - gráfico com tamanho mínimo confortável;
    - detalhe expansível com composição, contratos, faturas e registros;
    - ocultação amigável quando não há massa de dados suficiente.
    """
    title = safe_title(title, "Visual analítico") or "Visual analítico"
    desc = normalize_markdown(desc or "")
    chart_key = f"chart_{key}_{st.session_state.get('filter_version', 0)}"
    try:
        fig = sanitize_figure(fig)
    except Exception:
        pass
    visible_values = visible_values_from_figure(fig, filter_field) if fig is not None else []
    st.session_state.visual_scopes[key] = {"field": filter_field, "values": [clean_text(v) for v in visible_values]}

    with st.container():
        st.markdown("<div class='visual-shell'>", unsafe_allow_html=True)
        st.markdown(
            f"<div class='visual-title'>{_plain(title)}</div>"
            f"<div class='visual-subtitle'>{inline_markdown_to_html(desc)}</div>",
            unsafe_allow_html=True,
        )
        render_filter_hint(filter_field)
        render_visual_action_context(title, desc, data, filter_field)

        enough, reason = has_enough_visual_data(fig, data, filter_field, key)
        if not enough:
            render_insufficient_visual_info(title, reason, data, filter_field)
            if show_details and data is not None and not data.empty:
                payload = build_detail_payload(title, desc, data, filter_field, key, visible_values, None)
                with st.expander("▾ Ver mais informações dos dados disponíveis", expanded=False):
                    detail_content(payload)
            st.markdown("</div>", unsafe_allow_html=True)
            return

        safe_plotly_chart(fig, chart_key, height)
        render_active_filter_chip(filter_field, key)

        if show_details:
            payload = build_detail_payload(title, desc, data, filter_field, key, visible_values, None)
            with st.expander("▾ Ver mais detalhes, evidências e composição do visual", expanded=False):
                detail_content(payload)
        st.markdown("</div>", unsafe_allow_html=True)




def render_chart_detail_expander(title: str, desc: str, fig: go.Figure, data: pd.DataFrame, filter_field: Optional[str], key: str) -> None:
    """Renderiza o detalhamento em largura total, fora de colunas estreitas."""
    title = safe_title(title, "Visual analítico") or "Visual analítico"
    desc = normalize_markdown(desc or "")
    try:
        fig_clean = sanitize_figure(fig) if fig is not None else fig
    except Exception:
        fig_clean = fig
    visible_values = visible_values_from_figure(fig_clean, filter_field) if fig_clean is not None else []
    payload = build_detail_payload(title, desc, data, filter_field, key, visible_values, None)
    with st.expander(f"▾ Ver mais detalhes — {title}", expanded=False):
        detail_content(payload)


def render_chart_grid(chart_specs: List[Dict[str, Any]], columns: int = 2) -> None:
    """Mostra gráficos menores em até duas colunas, com detalhes logo abaixo do próprio visual.

    Regra de UX da V83:
    - se o gráfico está em uma coluna, o expansor fica na mesma coluna, imediatamente abaixo;
    - detalhes genéricos/redundantes foram removidos para manter a leitura leve;
    - gráficos pesados podem passar columns=1 para ganhar largura total.
    """
    if not chart_specs:
        return
    columns = max(1, min(2, int(columns or 2)))
    for start in range(0, len(chart_specs), columns):
        row = chart_specs[start:start + columns]
        cols = st.columns(len(row)) if len(row) > 1 else [st.container()]
        for col, spec in zip(cols, row):
            with col:
                render_native_chart(
                    spec["title"],
                    spec.get("desc", ""),
                    spec["fig"],
                    spec["data"],
                    spec.get("filter_field"),
                    spec["key"],
                    spec.get("height", 460),
                    show_details=True,
                )

# -----------------------------
# Aggregations / charts
# -----------------------------
def default_for_field(field: str) -> str:
    defaults = {
        SUPPLIER_COL: "Sem fornecedor",
        SERVICE_COL: "Sem serviço",
        CATEGORY_COL: "Sem categoria",
        CONTRACT_COL: "Sem contrato",
        INVOICE_COL: "Sem fatura",
        BRANCH_COL: "Sem filial",
        REGION_COL: "Sem região",
        CC_COL: "Sem centro de custo",
        CC_ID_COL: "Sem código",
        MONTH_COL: "Sem período",
        STATUS_COL: "Não informado",
    }
    return defaults.get(field, "Não informado")

def topn(df: pd.DataFrame, field: str, n: int = 10) -> pd.DataFrame:
    if field not in df.columns or df.empty:
        return pd.DataFrame(columns=[field, VALUE_COL])
    work = df.copy()
    work[field] = work[field].map(lambda v: clean_text(v, default_for_field(field)))
    work = work[work[field].map(lambda v: normalize_key(v) not in {"", "undefined", "none", "nan", "null", "nat"})]
    if work.empty:
        return pd.DataFrame(columns=[field, VALUE_COL])
    return work.groupby(field, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(n).reset_index()

def make_bar(df: pd.DataFrame, field: str, title: str, n: int = 10, horizontal: bool = True, label_limit: int = 42) -> go.Figure:
    """Barra robusta para rankings: não depende de px/color por categoria e preserva customdata real."""
    agg = topn(df, field, n)
    if not agg.empty:
        agg[VALUE_COL] = to_number_safe(agg[VALUE_COL]).fillna(0.0)
        agg["_Label"] = agg[field].map(lambda x: short_label(x, label_limit))
    fig = go.Figure()
    if horizontal:
        if not agg.empty:
            agg = agg.sort_values(VALUE_COL, ascending=True)
        fig.add_bar(
            x=agg[VALUE_COL] if not agg.empty else [],
            y=agg["_Label"] if not agg.empty else [],
            orientation="h",
            customdata=agg[[field, VALUE_COL]] if not agg.empty else None,
            hovertext=hover_texts(agg, field, VALUE_COL),
            marker_color=PALETTE[0],
            text=[compact_money(v) for v in agg[VALUE_COL]] if not agg.empty else [],
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{hovertext}<extra></extra>",
            name="Gasto",
        )
        fig.update_yaxes(title_text="", categoryorder="array", categoryarray=agg["_Label"].tolist() if not agg.empty else [])
    else:
        fig.add_bar(
            x=agg["_Label"] if not agg.empty else [],
            y=agg[VALUE_COL] if not agg.empty else [],
            customdata=agg[[field, VALUE_COL]] if not agg.empty else None,
            hovertext=hover_texts(agg, field, VALUE_COL),
            marker_color=PALETTE[0],
            text=[compact_money(v) for v in agg[VALUE_COL]] if not agg.empty else [],
            textposition="outside",
            cliponaxis=False,
            hovertemplate="%{hovertext}<extra></extra>",
            name="Gasto",
        )
        fig.update_xaxes(title_text="", tickangle=-12)
    fig.update_layout(title_text="")
    return style_fig(fig, height=460, showlegend=False)

def make_donut(df: pd.DataFrame, field: str) -> go.Figure:
    agg = topn(df, field, 8)
    fig = px.pie(agg, names=field, values=VALUE_COL, hole=.58, custom_data=[field], color=field, color_discrete_sequence=PALETTE)
    fig.update_traces(textinfo="percent", textposition="inside", sort=False, marker=dict(line=dict(color="rgba(255,255,255,.22)", width=1.2)))
    return style_fig(fig, height=410, showlegend=True)

def make_trend(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        agg = pd.DataFrame({MONTH_COL: [], VALUE_COL: []})
    else:
        agg = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().reset_index()
        agg["_sort"] = agg[MONTH_COL].map(period_sort_key)
        agg = agg.sort_values("_sort").drop(columns=["_sort"])
    fig = px.line(agg, x=MONTH_COL, y=VALUE_COL, markers=True, custom_data=[MONTH_COL], color_discrete_sequence=[PALETTE[1]])
    fig.update_traces(
        line=dict(width=4), marker=dict(size=10), fill="tozeroy",
        text=agg[VALUE_COL].map(compact_money) if not agg.empty else None,
        hovertemplate="%{x}<br>Gasto: %{text}<extra></extra>"
    )
    return style_fig(fig, height=360, showlegend=False)

def make_pareto(df: pd.DataFrame, field: str) -> go.Figure:
    # Serviço costuma ter descrição grande. A barra usa rótulo curto, mas o customdata
    # preserva o nome completo para filtro, tooltip e janela de detalhe.
    agg = topn(df, field, 12)
    if not agg.empty:
        agg = agg.sort_values(VALUE_COL, ascending=False).copy()
        agg["_Label"] = agg[field].map(lambda x: short_label(x, 22))
        total = agg[VALUE_COL].sum() or 1
        agg["Acumulado_%"] = agg[VALUE_COL].cumsum() / total * 100
    fig = go.Figure()
    fig.add_bar(
        x=agg["_Label"] if not agg.empty else [],
        y=agg[VALUE_COL] if not agg.empty else [],
        customdata=agg[[field]] if not agg.empty else None,
        hovertext=hover_texts(agg, field, VALUE_COL),
        name="Gasto",
        marker_color=PALETTE[0],
        text=[compact_money(v) for v in agg[VALUE_COL]] if not agg.empty else [],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{hovertext}<extra></extra>",
    )
    fig.add_scatter(
        x=agg["_Label"] if not agg.empty else [],
        y=agg["Acumulado_%"] if not agg.empty else [],
        customdata=agg[[field]] if not agg.empty else None,
        name="Acumulado %",
        yaxis="y2",
        mode="lines+markers+text",
        text=[pct(v) for v in agg["Acumulado_%"]] if not agg.empty else [],
        textposition="top center",
        line=dict(width=3, color=PALETTE[3]),
        hovertemplate="%{customdata[0]}<br>Acumulado: %{y:.1f}%<extra></extra>",
    )
    fig.update_layout(yaxis2=dict(overlaying="y", side="right", range=[0, 108], ticksuffix="%", showgrid=False), xaxis_tickangle=-28, margin=dict(t=34, b=112, l=44, r=48))
    return style_fig(fig, height=510, showlegend=True)

def make_diff_bar(df: pd.DataFrame) -> go.Figure:
    """Barras divergentes para diferença por fornecedor, separando crédito/débito visualmente."""
    if DIFF_COL not in df.columns or df.empty:
        agg = pd.DataFrame(columns=[SUPPLIER_COL, DIFF_COL, "_Label"])
    else:
        agg = df.groupby(SUPPLIER_COL, dropna=False)[DIFF_COL].sum().reset_index()
        agg["_Abs"] = agg[DIFF_COL].abs()
        agg = agg.sort_values("_Abs", ascending=False).head(10).sort_values(DIFF_COL, ascending=True)
        agg["_Label"] = agg[SUPPLIER_COL].map(lambda x: short_label(x, 34))
    colors = [PALETTE[4] if v < 0 else PALETTE[2] for v in agg[DIFF_COL]] if not agg.empty else []
    fig = go.Figure()
    fig.add_bar(
        x=agg[DIFF_COL] if not agg.empty else [],
        y=agg["_Label"] if not agg.empty else [],
        orientation="h",
        customdata=agg[[SUPPLIER_COL]] if not agg.empty else None,
        hovertext=hover_texts(agg, SUPPLIER_COL, DIFF_COL, value_name="Diferença"),
        marker_color=colors,
        text=[money(v) for v in agg[DIFF_COL]] if not agg.empty else [],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{hovertext}<extra></extra>",
        name="Diferença",
    )
    fig.add_vline(x=0, line_width=1, line_dash="dash", line_color="rgba(255,255,255,.45)")
    fig.update_yaxes(title_text="")
    return style_fig(fig, height=440, showlegend=False)

def make_costcenter_bar(df: pd.DataFrame) -> go.Figure:
    agg = topn(df, CC_COL, 8)
    if not agg.empty:
        ids = df[[CC_COL, CC_ID_COL]].drop_duplicates(CC_COL) if CC_ID_COL in df.columns else pd.DataFrame({CC_COL: [], CC_ID_COL: []})
        agg = agg.merge(ids, on=CC_COL, how="left")
        agg["_Label"] = agg[CC_COL].map(lambda x: short_label(x, 34))
        agg = agg.sort_values(VALUE_COL, ascending=True)
    fig = go.Figure()
    fig.add_bar(
        x=agg[VALUE_COL] if not agg.empty else [],
        y=agg["_Label"] if not agg.empty else [],
        orientation="h",
        customdata=agg[[CC_COL, CC_ID_COL, VALUE_COL]] if not agg.empty and CC_ID_COL in agg.columns else None,
        hovertext=[f"{clean_text(r[CC_COL])}<br>Código CC: {clean_text(r.get(CC_ID_COL, '-'))}<br>Gasto: {money(r[VALUE_COL])}" for _, r in agg.iterrows()] if not agg.empty else [],
        marker_color=PALETTE[5],
        text=[compact_money(v) for v in agg[VALUE_COL]] if not agg.empty else [],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{hovertext}<extra></extra>",
        name="Gasto",
    )
    fig.update_yaxes(title_text="")
    return style_fig(fig, height=450, showlegend=False)

def make_status_donut(df: pd.DataFrame) -> go.Figure:
    field = STATUS_COL if STATUS_COL in df.columns else CONTESTED_COL
    return make_donut(df, field)


def make_share_bar(df: pd.DataFrame, field: str, n: int = 8) -> go.Figure:
    """Substitui roscas/pizzas por barras horizontais com participação percentual."""
    agg = topn(df, field, n)
    if not agg.empty:
        total = agg[VALUE_COL].sum() or 1
        agg["Participacao"] = agg[VALUE_COL] / total * 100
        agg["_Label"] = agg[field].map(lambda x: short_label(x, 34))
        agg = agg.sort_values(VALUE_COL, ascending=True)
    fig = go.Figure()
    fig.add_bar(
        x=agg[VALUE_COL] if not agg.empty else [],
        y=agg["_Label"] if not agg.empty else [],
        orientation="h",
        customdata=agg[[field, "Participacao", VALUE_COL]] if not agg.empty else None,
        hovertext=[f"{clean_text(r[field])}<br>Gasto: {money(r[VALUE_COL])}<br>Participação: {pct(r["Participacao"])}" for _, r in agg.iterrows()] if not agg.empty else [],
        marker_color=PALETTE[0],
        text=[f"{compact_money(v)} • {pct(p)}" for v, p in zip(agg[VALUE_COL], agg["Participacao"])] if not agg.empty else [],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{hovertext}<extra></extra>",
        name="Gasto",
    )
    fig.update_yaxes(title_text="")
    return style_fig(fig, height=470, showlegend=False)

def make_waterfall(df: pd.DataFrame) -> go.Figure:
    """Cascata FinOps: contratado + diferenças/ajustes = realizado."""
    realizado = float(df[VALUE_COL].sum()) if VALUE_COL in df.columns and not df.empty else 0.0
    contratado = float(df["Valor_Contratado"].sum()) if "Valor_Contratado" in df.columns and not df.empty else 0.0
    diferenca = float(df[DIFF_COL].sum()) if DIFF_COL in df.columns and not df.empty else 0.0
    if contratado <= 0:
        contratado = max(0.0, realizado - diferenca)
    ajustes = max(0.0, realizado - contratado - diferenca)
    fig = go.Figure(go.Waterfall(
        name="FinOps", orientation="v",
        measure=["absolute", "relative", "relative", "total"],
        x=["Contratado", "Ajustes/Serviços", "Diferenças", "Realizado"],
        y=[contratado, ajustes, diferenca, realizado],
        text=[compact_money(v) for v in [contratado, ajustes, diferenca, realizado]],
        textposition="outside",
        connector={"line": {"color": "rgba(148,163,184,.45)"}},
        increasing={"marker": {"color": PALETTE[2]}},
        decreasing={"marker": {"color": PALETTE[4]}},
        totals={"marker": {"color": PALETTE[0]}},
        hovertemplate="%{x}<br>Valor: %{text}<extra></extra>",
    ))
    return style_fig(fig, height=420, showlegend=False)



def make_heatmap_supplier_month(df: pd.DataFrame) -> go.Figure:
    if df.empty or SUPPLIER_COL not in df.columns or MONTH_COL not in df.columns:
        pivot = pd.DataFrame()
    else:
        top_suppliers = topn(df, SUPPLIER_COL, 8)[SUPPLIER_COL].tolist()
        base = df[df[SUPPLIER_COL].isin(top_suppliers)].copy() if top_suppliers else df.copy()
        pivot = base.pivot_table(index=SUPPLIER_COL, columns=MONTH_COL, values=VALUE_COL, aggfunc="sum", fill_value=0.0)
        pivot = pivot.reindex(index=top_suppliers).sort_index(axis=1, key=lambda x: [period_sort_key(v) for v in x])
    fig = go.Figure(data=go.Heatmap(
        z=pivot.values if not pivot.empty else [],
        x=list(pivot.columns) if not pivot.empty else [],
        y=[short_label(v, 30) for v in pivot.index] if not pivot.empty else [],
        customdata=[[supplier for _ in pivot.columns] for supplier in pivot.index] if not pivot.empty else None,
        colorscale="Blues",
        hovertemplate="Fornecedor: %{customdata}<br>Período: %{x}<br>Gasto: R$ %{z:,.2f}<extra></extra>",
    ))
    return style_fig(fig, height=430, showlegend=False)

def make_supplier_service_treemap(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        agg = pd.DataFrame(columns=[SUPPLIER_COL, SERVICE_COL, VALUE_COL])
    else:
        agg = df.groupby([SUPPLIER_COL, SERVICE_COL], dropna=False)[VALUE_COL].sum().reset_index()
        top_pairs = agg.sort_values(VALUE_COL, ascending=False).head(40)
        agg = top_pairs.copy()
    if agg.empty:
        fig = go.Figure(go.Treemap(labels=[], parents=[], values=[]))
    else:
        fig = px.treemap(agg, path=[SUPPLIER_COL, SERVICE_COL], values=VALUE_COL, custom_data=[SUPPLIER_COL, SERVICE_COL], color=VALUE_COL, color_continuous_scale="Blues")
        fig.update_traces(hovertemplate="Fornecedor: %{customdata[0]}<br>Serviço: %{customdata[1]}<br>Gasto: R$ %{value:,.2f}<extra></extra>")
    return style_fig(fig, height=460, showlegend=False)

def make_anomaly_bar(df: pd.DataFrame) -> go.Figure:
    """Serviços cujo último mês está acima da média histórica do próprio serviço."""
    rows = []
    if not df.empty and SERVICE_COL in df.columns and MONTH_COL in df.columns:
        by = df.groupby([SERVICE_COL, MONTH_COL], dropna=False)[VALUE_COL].sum().reset_index()
        months = sorted(by[MONTH_COL].unique().tolist(), key=period_sort_key)
        if len(months) >= 2:
            last = months[-1]
            for service, g in by.groupby(SERVICE_COL):
                hist = g[g[MONTH_COL] != last][VALUE_COL]
                current = float(g.loc[g[MONTH_COL] == last, VALUE_COL].sum())
                avg = float(hist.mean()) if not hist.empty else 0.0
                delta = current - avg
                delta_pct = (delta / avg * 100) if avg else 0.0
                if current > 0 and delta > 0:
                    rows.append({SERVICE_COL: service, "Atual": current, "Media": avg, "Aumento": delta, "Aumento_%": delta_pct})
    agg = pd.DataFrame(rows).sort_values("Aumento", ascending=False).head(10) if rows else pd.DataFrame(columns=[SERVICE_COL,"Atual","Media","Aumento","Aumento_%"])
    if not agg.empty:
        agg["_Label"] = agg[SERVICE_COL].map(lambda x: short_label(x, 34))
        agg = agg.sort_values("Aumento", ascending=True)
    fig = go.Figure()
    fig.add_bar(
        x=agg["Aumento"] if not agg.empty else [],
        y=agg["_Label"] if not agg.empty else [],
        orientation="h",
        customdata=agg[[SERVICE_COL, "Atual", "Media", "Aumento_%", "Aumento"]] if not agg.empty else None,
        hovertext=[f"{clean_text(r[SERVICE_COL])}<br>Aumento: {money(r["Aumento"])}<br>Atual: {money(r["Atual"])}<br>Média: {money(r["Media"])}<br>Variação: {pct(r["Aumento_%"])}" for _, r in agg.iterrows()] if not agg.empty else [],
        marker_color=PALETTE[3],
        text=[compact_money(v) for v in agg["Aumento"]] if not agg.empty else [],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{hovertext}<extra></extra>",
        name="Aumento",
    )
    fig.update_yaxes(title_text="")
    return style_fig(fig, height=430, showlegend=False)

def make_contract_exposure(df: pd.DataFrame) -> go.Figure:
    field = CONTRACT_COL if CONTRACT_COL in df.columns else INVOICE_COL
    agg = topn(df, field, 12)
    if not agg.empty:
        total = float(agg[VALUE_COL].sum()) or 1.0
        agg["Participacao"] = agg[VALUE_COL] / total * 100
        agg["_Label"] = agg[field].map(lambda x: short_label(x, 34))
        agg = agg.sort_values(VALUE_COL, ascending=True)
    fig = go.Figure()
    fig.add_bar(
        x=agg[VALUE_COL] if not agg.empty else [],
        y=agg["_Label"] if not agg.empty else [],
        orientation="h",
        customdata=agg[[field, "Participacao", VALUE_COL]] if not agg.empty else None,
        hovertext=[f"{clean_text(r[field])}<br>Gasto: {money(r[VALUE_COL])}<br>Participação: {pct(r["Participacao"])}" for _, r in agg.iterrows()] if not agg.empty else [],
        marker_color=PALETTE[2],
        text=[f"{compact_money(v)} • {pct(p)}" for v, p in zip(agg[VALUE_COL], agg["Participacao"])] if not agg.empty else [],
        textposition="outside",
        cliponaxis=False,
        hovertemplate="%{hovertext}<extra></extra>",
        name="Gasto",
    )
    fig.update_yaxes(title_text="")
    return style_fig(fig, height=430, showlegend=False)

def render_trend_strip(df: pd.DataFrame) -> None:
    if df.empty or MONTH_COL not in df.columns or VALUE_COL not in df.columns:
        return
    trend = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().reset_index()
    trend["_sort"] = trend[MONTH_COL].map(period_sort_key)
    trend = trend.sort_values("_sort").drop(columns=["_sort"]).tail(12)
    if trend.empty:
        return
    fig = go.Figure()
    fig.add_scatter(x=trend[MONTH_COL], y=trend[VALUE_COL], mode="lines+markers", line=dict(width=3, color=PALETTE[1]), marker=dict(size=7), fill="tozeroy", hovertemplate="%{x}<br>Gasto: R$ %{y:,.2f}<extra></extra>")
    fig.update_layout(height=180, margin=dict(t=10,b=30,l=10,r=10), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#E5EEF9", size=11), showlegend=False)
    fig.update_yaxes(visible=False)
    fig.update_xaxes(gridcolor="rgba(148,163,184,.12)")
    with st.container(border=True):
        st.markdown("### 📈 Tendência dos últimos períodos")
        st.caption("Faixa executiva para enxergar rapidamente se o gasto está subindo, caindo ou estabilizado.")
        safe_plotly_chart(fig, "trend_strip", 220)

def render_executive_insights(df: pd.DataFrame) -> None:
    if df.empty:
        return
    total = df[VALUE_COL].sum() if VALUE_COL in df.columns else 0
    top_supplier = df.groupby(SUPPLIER_COL)[VALUE_COL].sum().sort_values(ascending=False).head(1) if SUPPLIER_COL in df.columns else pd.Series(dtype=float)
    top_service = df.groupby(SERVICE_COL)[VALUE_COL].sum().sort_values(ascending=False).head(1) if SERVICE_COL in df.columns else pd.Series(dtype=float)
    diff = df[DIFF_COL].sum() if DIFF_COL in df.columns else 0
    supplier_txt = "Sem fornecedor relevante"
    if not top_supplier.empty and total:
        supplier_txt = f"{clean_text(top_supplier.index[0])} concentra {pct(top_supplier.iloc[0] / total * 100)} do gasto."
    service_txt = "Sem serviço relevante"
    if not top_service.empty and total:
        service_txt = f"{clean_text(top_service.index[0])} é o serviço de maior impacto financeiro."
    diff_txt = "Diferenças sob controle." if abs(diff) < 1 else f"Há {money(diff)} em diferenças para validar."
    st.markdown(f"""
    <div class='executive-grid'>
      <div class='exec-card'><strong>🏢 Concentração</strong><span>{supplier_txt}</span></div>
      <div class='exec-card'><strong>🧱 Maior impacto</strong><span>{service_txt}</span></div>
      <div class='exec-card'><strong>⚠️ Auditoria</strong><span>{diff_txt}</span></div>
    </div>
    """, unsafe_allow_html=True)


def visual_description(title: str, data: pd.DataFrame, group: Optional[str]) -> str:
    """Texto curto e específico para a janela detalhes."""
    if data.empty:
        return "Sem dados disponíveis para este visual."
    total = data[VALUE_COL].sum() if VALUE_COL in data.columns else 0
    linhas = len(data)
    if group and group in data.columns:
        ranking = data.groupby(group, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
        if not ranking.empty:
            top_name, top_value = ranking.index[0], ranking.iloc[0]
            share = (top_value / total * 100) if total else 0
            return (
                f"Este detalhamento analisa **{linhas} lançamento(s)** ligados ao visual **{safe_title(title, 'selecionado')}**. "
                f"O principal item em **{group}** é **{top_name}**, com **{money(top_value)}** "
                f"(**{pct(share)}** do total exibido nesta janela)."
            )
    return f"Este detalhamento reúne **{linhas} lançamento(s)** e soma **{money(total)}** no contexto do visual **{safe_title(title, 'selecionado')}**."

# -----------------------------
# Detail dialog
# -----------------------------
def _format_detail_table(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if c in [VALUE_COL, DIFF_COL, "Valor_Contratado", "Gasto", "Diferenca", "Valor", "Media", "Maior_Valor", "Menor_Valor"]:
            out[c] = out[c].map(money)
        elif str(c).endswith("_%") or c in ["Participacao_%"]:
            out[c] = out[c].map(lambda v: pct(float(v)) if pd.notna(v) else "0,0%")
    return sanitize_dataframe_display(out)


def _agg_detail(df: pd.DataFrame, group: str, limit: int = 20) -> pd.DataFrame:
    if df.empty or group not in df.columns:
        return pd.DataFrame()
    agg_map = {
        "Gasto": (VALUE_COL, "sum"),
        "Lancamentos": ("Linha_ID", "count") if "Linha_ID" in df.columns else (VALUE_COL, "count"),
        "Media": (VALUE_COL, "mean"),
        "Maior_Valor": (VALUE_COL, "max"),
    }
    if DIFF_COL in df.columns:
        agg_map["Diferenca"] = (DIFF_COL, "sum")
    if SUPPLIER_COL in df.columns and group != SUPPLIER_COL:
        agg_map["Fornecedores"] = (SUPPLIER_COL, "nunique")
    if SERVICE_COL in df.columns and group != SERVICE_COL:
        agg_map["Servicos"] = (SERVICE_COL, "nunique")
    if INVOICE_COL in df.columns and group != INVOICE_COL:
        agg_map["Faturas"] = (INVOICE_COL, "nunique")
    if CONTRACT_COL in df.columns and group != CONTRACT_COL:
        agg_map["Contratos"] = (CONTRACT_COL, "nunique")
    rank = df.groupby(group, dropna=False).agg(**agg_map).reset_index().sort_values("Gasto", ascending=False)
    total = float(rank["Gasto"].sum()) or 1.0
    rank["Participacao_%"] = rank["Gasto"] / total * 100
    return rank.head(limit)


def _render_rank_table(df: pd.DataFrame, group: str, title: str, limit: int = 20, note: str = "") -> None:
    rank = _agg_detail(df, group, limit)
    if rank.empty:
        st.info(f"Sem dados consolidados por {safe_title(group, 'campo')} neste contexto. Isso geralmente indica que a coluna não existe na base carregada ou que o filtro atual deixou somente lançamentos sem esse campo preenchido.")
        return
    st.markdown(f"### {safe_title(title, 'Tabela analítica') or 'Tabela analítica'}")
    if note:
        render_reading_note(note)
    else:
        render_reading_note(f"Leia esta tabela como um ranking por **{safe_title(group, 'campo')}**: a coluna **Gasto** mostra o total financeiro, **Participacao_%** mostra o peso dentro deste detalhamento e as demais colunas ajudam a avaliar concentração e volume operacional.")
    show = rank.copy()
    show[group] = show[group].map(lambda x: short_label(x, 90))
    st.dataframe(_format_detail_table(show), width="stretch", hide_index=True, height=min(520, 96 + len(show) * 36))


def _render_detail_insights(df: pd.DataFrame, group: Optional[str]) -> None:
    total = float(df[VALUE_COL].sum()) if VALUE_COL in df.columns and not df.empty else 0.0
    linhas = len(df)
    diff = float(df[DIFF_COL].sum()) if DIFF_COL in df.columns and not df.empty else 0.0
    maior = None
    if group and group in df.columns and total:
        g = df.groupby(group, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(1)
        if not g.empty:
            maior = f"{clean_text(g.index[0])} concentra {pct(float(g.iloc[0]) / total * 100)}."
    if not maior and SUPPLIER_COL in df.columns and total:
        g = df.groupby(SUPPLIER_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(1)
        if not g.empty:
            maior = f"{clean_text(g.index[0])} é o maior fornecedor do contexto, com {money(float(g.iloc[0]))}."
    diff_txt = "Sem diferença relevante." if abs(diff) < 0.01 else f"Diferença acumulada de {money(diff)} para auditar."
    st.markdown(f"""
    <div class='executive-grid'>
      <div class='exec-card'><strong>📌 Escopo</strong><span>{linhas} lançamento(s) somando {money(total)}.</span></div>
      <div class='exec-card'><strong>🏆 Concentração</strong><span>{maior or 'Sem concentração relevante.'}</span></div>
      <div class='exec-card'><strong>🧾 Auditoria</strong><span>{diff_txt}</span></div>
    </div>
    """, unsafe_allow_html=True)


def render_detail_summary_table(data: pd.DataFrame) -> None:
    """Resumo compacto para detalhamentos.

    Evita usar cards/KPIs grandes dentro de containers estreitos, prevenindo
    cortes de valores monetários longos em telas menores ou quando o detalhe
    está aberto abaixo de gráficos em coluna.
    """
    total = data[VALUE_COL].sum() if VALUE_COL in data.columns else 0
    diff = data[DIFF_COL].sum() if DIFF_COL in data.columns else 0
    fornecedores = data[SUPPLIER_COL].nunique() if SUPPLIER_COL in data.columns else 0
    contratos = data[CONTRACT_COL].nunique() if CONTRACT_COL in data.columns else 0
    faturas = data[INVOICE_COL].nunique() if INVOICE_COL in data.columns else 0
    servicos = data[SERVICE_COL].nunique() if SERVICE_COL in data.columns else 0

    resumo = pd.DataFrame(
        [
            {"Indicador": "Gasto analisado", "Valor": money(total), "Leitura": "Soma financeira dos lançamentos que compõem este visual."},
            {"Indicador": "Lançamentos", "Valor": f"{len(data):,}".replace(",", "."), "Leitura": "Quantidade de registros considerados no detalhamento."},
            {"Indicador": "Fornecedores", "Valor": str(fornecedores), "Leitura": "Operadoras/fornecedores presentes no recorte aberto."},
            {"Indicador": "Contratos", "Valor": str(contratos), "Leitura": "Contratos relacionados ao visual."},
            {"Indicador": "Faturas", "Valor": str(faturas), "Leitura": "Faturas relacionadas ao visual."},
            {"Indicador": "Serviços", "Valor": str(servicos), "Leitura": "Serviços distintos envolvidos."},
            {"Indicador": "Diferença", "Valor": money(diff), "Leitura": "Valor acumulado para revisar/validar."},
        ]
    )

    st.markdown("### 📌 Resumo do detalhamento")
    st.markdown(
        "<div class='detail-summary-note'>Resumo do recorte aberto, em formato compacto para evitar valores cortados e facilitar a leitura antes da análise das tabelas.</div>",
        unsafe_allow_html=True,
    )
    st.dataframe(sanitize_dataframe_display(resumo), width="stretch", hide_index=True)


def detail_content(payload: Dict[str, Any]) -> None:
    title = safe_title(payload.get("title"), "Detalhamento do visual") or "Detalhamento do visual"
    data_original = payload["data"].copy()
    group = payload.get("group_field")
    visual_key = payload.get("visual_key", "visual")
    visible_values = payload.get("visible_values", []) or []
    selected_value = payload.get("selected_value", "")

    st.markdown(f"## 🔎 {title}")
    subtitle = safe_title(payload.get("subtitle"), "Detalhamento auditável do visual selecionado.")
    if subtitle:
        st.caption(subtitle)

    scope_note = visual_description(title, data_original, group)
    if group and visible_values:
        scope_note += f"\n\n**Escopo real do gráfico:** {len(visible_values)} item(ns) visível(is) em **{group}**. O detalhamento usa somente os lançamentos que formam esse visual."
    if selected_value:
        scope_note += f"\n\n**Última seleção por clique:** {group or 'visual'} = **{selected_value}**."
    render_markdown_box(scope_note)
    render_reading_note("Caminho recomendado de drilldown: **Fornecedor → Contrato → Serviço → Fatura → Lançamento**. Comece pelo ranking, depois valide contratos/faturas e finalize na tabela linha a linha.")

    if data_original.empty:
        st.info("Nenhum lançamento encontrado para este contexto.")
        return

    with st.expander("🎛️ Filtros internos do detalhamento", expanded=False):
        filter_candidates = [c for c in [MONTH_COL, SUPPLIER_COL, CATEGORY_COL, SERVICE_COL, STATUS_COL, CONTRACT_COL, INVOICE_COL, REGION_COL, BRANCH_COL, CC_COL] if c in data_original.columns]
        cols = st.columns(min(2, max(1, len(filter_candidates))))
        internal_filters: Dict[str, str] = {}
        for i, field in enumerate(filter_candidates):
            opts = ["Todos"] + sorted(data_original[field].dropna().astype(str).unique().tolist())
            key = f"detail_filter_{visual_key}_{field}"
            with cols[i % len(cols)]:
                choice = st.selectbox(field, opts, key=key)
            if choice != "Todos":
                internal_filters[field] = choice
        if st.button("🧹 Limpar filtros internos", key=f"detail_clear_{visual_key}"):
            for field in filter_candidates:
                st.session_state.pop(f"detail_filter_{visual_key}_{field}", None)
            st.rerun()

    data = data_original.copy()
    for field, value in internal_filters.items():
        data = data[data[field].map(normalize_key) == normalize_key(value)]

    if data.empty:
        st.info("Nenhum lançamento encontrado com os filtros internos desta janela.")
        return

    render_detail_summary_table(data)

    _render_detail_insights(data, group)

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Dados do gráfico", "🔍 Quebras complementares", "🧾 Contratos/Faturas", "🚨 Auditoria", "📋 Lançamentos"])

    with tab1:
        render_reading_note("Esta aba mostra exatamente a composição do visual aberto. Use o ranking para descobrir quem ou o que forma a maior parte do custo, e a evolução para conferir se o gasto é recorrente, sazonal ou pontual.")
        if group and group in data.columns:
            _render_rank_table(data, group, f"Ranking que compõe o visual por {group}", limit=30)
        else:
            st.markdown("### Composição executiva do visual")
            for fallback in [SUPPLIER_COL, CATEGORY_COL, SERVICE_COL, MONTH_COL]:
                if fallback in data.columns:
                    _render_rank_table(data, fallback, f"Composição por {fallback}", limit=12)
                    break
        # V83: removido o bloco genérico “🏷️ Top Serviço/Fornecedor”.
        # A informação específica do visual fica no ranking principal acima;
        # quebras por Serviço, Fornecedor, CC, Filial e Região permanecem na aba
        # “Quebras complementares”, evitando repetição no detalhamento.

    with tab2:
        render_reading_note("Quebras complementares servem para investigação: elas mostram o mesmo conjunto de lançamentos por mês, fornecedor, categoria, serviço, status, filial e centro de custo quando esses campos existem na planilha.")
        complementares = [MONTH_COL, SUPPLIER_COL, CATEGORY_COL, SERVICE_COL, STATUS_COL, REGION_COL, BRANCH_COL, CC_COL]
        complementares = [c for c in complementares if c in data.columns and c != group]
        if not complementares:
            st.info("Não há campos complementares disponíveis nesta base para detalhar mais este visual.")
        else:
            for field in complementares[:5]:
                _render_rank_table(data, field, f"Quebra por {field}", limit=12)

        if DIFF_COL in data.columns and data[DIFF_COL].abs().sum() > 0:
            st.markdown("### ⚠️ Maiores diferenças para auditoria")
            diff_cols = [c for c in [MONTH_COL, SUPPLIER_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, VALUE_COL, DIFF_COL, STATUS_COL] if c in data.columns]
            diff_show = data.loc[data[DIFF_COL].abs().sort_values(ascending=False).index, diff_cols].head(20).copy()
            st.dataframe(_format_detail_table(diff_show), width="stretch", hide_index=True, height=420)

    with tab3:
        render_reading_note("Aqui ficam os vínculos administrativos do gasto. Use contratos para avaliar exposição e vencimento; use faturas para conferir cobranças recorrentes, duplicidades ou valores fora do padrão.")
        if CONTRACT_COL in data.columns:
            _render_rank_table(data, CONTRACT_COL, "Contratos relacionados ao visual", limit=30)
        if INVOICE_COL in data.columns:
            _render_rank_table(data, INVOICE_COL, "Faturas relacionadas ao visual", limit=30)
        if CONTRACT_COL not in data.columns and INVOICE_COL not in data.columns:
            st.info("Esta base não trouxe campos de contrato/fatura para este detalhamento.")

    with tab4:
        st.markdown("### 🚨 Leituras de auditoria do contexto")
        st.markdown("#### ⚖️ Diferenças por fornecedor")
        safe_plotly_chart(make_diff_bar(data), f"detail_{visual_key}_diff", 500)
        st.markdown("#### 📊 Cascata FinOps")
        safe_plotly_chart(make_waterfall(data), f"detail_{visual_key}_waterfall", 500)
        st.markdown("#### 🔎 Possíveis anomalias")
        safe_plotly_chart(make_anomaly_bar(data), f"detail_{visual_key}_anomaly", 500)
        if DIFF_COL in data.columns:
            diff_cols = [c for c in [MONTH_COL, SUPPLIER_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, VALUE_COL, DIFF_COL, STATUS_COL] if c in data.columns]
            audit = data.assign(_abs_diff=data[DIFF_COL].abs()).sort_values("_abs_diff", ascending=False).drop(columns=["_abs_diff"]).head(30)
            if not audit.empty:
                st.markdown("#### 🧾 Linhas com maior diferença")
                st.dataframe(_format_detail_table(audit[diff_cols]), width="stretch", hide_index=True, height=360)

    with tab5:
        render_reading_note("Tabela linha a linha usada para rastreabilidade. Ela mostra os registros brutos que sustentam os gráficos, já respeitando os filtros do painel e os filtros internos desta janela.")
        st.markdown("### 📋 Lançamentos que formam o visual")
        cols = [c for c in [MONTH_COL, SUPPLIER_COL, CATEGORY_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, VALUE_COL, DIFF_COL, STATUS_COL, REGION_COL, BRANCH_COL, CC_COL, CC_ID_COL, "Linha_ID"] if c in data.columns]
        show = data[cols].copy()
        st.dataframe(_format_detail_table(show), width="stretch", hide_index=True, height=520)

if hasattr(st, "dialog"):
    @st.dialog("Detalhamento do visual")
    def show_detail_dialog() -> None:
        detail_content(st.session_state.detail)
        if st.button("Fechar", width="stretch"):
            visual_key = st.session_state.detail.get("visual_key", "visual") if st.session_state.detail else "visual"
            for key in list(st.session_state.keys()):
                if str(key).startswith(f"detail_filter_{visual_key}_"):
                    st.session_state.pop(key, None)
            st.session_state.detail = None
            st.rerun()
else:
    def show_detail_dialog() -> None:
        with st.expander("Detalhamento do visual", expanded=True):
            detail_content(st.session_state.detail)
            if st.button("Fechar", width="stretch"):
                visual_key = st.session_state.detail.get("visual_key", "visual") if st.session_state.detail else "visual"
                for key in list(st.session_state.keys()):
                    if str(key).startswith(f"detail_filter_{visual_key}_"):
                        st.session_state.pop(key, None)
                st.session_state.detail = None
                st.rerun()



# -----------------------------
# V63 - Diagnóstico executivo, alertas, oportunidades e qualidade dos dados
# -----------------------------
def _dimension_top(df: pd.DataFrame, field: str, limit: int = 5) -> pd.DataFrame:
    if df.empty or field not in df.columns or VALUE_COL not in df.columns:
        return pd.DataFrame(columns=[field, VALUE_COL, "Participacao_%", "Lançamentos"])
    agg = df.groupby(field, dropna=False).agg(
        **{VALUE_COL: (VALUE_COL, "sum"), "Lançamentos": (VALUE_COL, "size")}
    ).sort_values(VALUE_COL, ascending=False).head(limit).reset_index()
    total = agg[VALUE_COL].sum() or 1
    agg["Participacao_%"] = agg[VALUE_COL] / total * 100
    return agg


def build_smart_alerts(df: pd.DataFrame) -> pd.DataFrame:
    """Gera alertas automáticos simples, auditáveis e independentes de IA externa."""
    rows: List[Dict[str, Any]] = []
    if df.empty or VALUE_COL not in df.columns:
        return pd.DataFrame(columns=["Prioridade", "Tipo", "Achado", "Impacto", "Ação sugerida"])

    total = float(df[VALUE_COL].sum())
    if SUPPLIER_COL in df.columns and total:
        top_supplier = _dimension_top(df, SUPPLIER_COL, 1)
        if not top_supplier.empty:
            share = float(top_supplier.iloc[0].get("Participacao_%", 0))
            supplier = clean_text(top_supplier.iloc[0][SUPPLIER_COL])
            if share >= 50:
                rows.append({"Prioridade": "Alta", "Tipo": "Dependência", "Achado": f"{supplier} concentra {pct(share)} do gasto do contexto.", "Impacto": money(top_supplier.iloc[0][VALUE_COL]), "Ação sugerida": "Avaliar risco de dependência e preparar comparação comercial com fornecedores alternativos."})
            elif share >= 35:
                rows.append({"Prioridade": "Média", "Tipo": "Concentração", "Achado": f"{supplier} é o principal fornecedor, com {pct(share)} do gasto.", "Impacto": money(top_supplier.iloc[0][VALUE_COL]), "Ação sugerida": "Monitorar tendência mensal e negociar pacotes/contratos de maior volume."})

    if MONTH_COL in df.columns:
        trend = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().sort_index()
        if len(trend) >= 2 and trend.iloc[-2] > 0:
            var = (trend.iloc[-1] - trend.iloc[-2]) / trend.iloc[-2] * 100
            if abs(var) >= 15:
                rows.append({"Prioridade": "Alta" if var > 0 else "Média", "Tipo": "Variação mensal", "Achado": f"Último período variou {pct(var)} em relação ao período anterior.", "Impacto": f"{money(trend.iloc[-2])} → {money(trend.iloc[-1])}", "Ação sugerida": "Abrir por fornecedor, serviço e fatura para identificar causa da oscilação."})

    if DIFF_COL in df.columns and df[DIFF_COL].abs().sum() > 0:
        diff = float(df[DIFF_COL].sum())
        rows.append({"Prioridade": "Alta" if abs(diff) > total * .02 else "Média", "Tipo": "Diferença", "Achado": f"Há diferença acumulada de {money(diff)} nos lançamentos filtrados.", "Impacto": money(abs(diff)), "Ação sugerida": "Priorizar fornecedores/contratos com maior diferença acumulada para conferência de fatura."})

    if CONTRACT_COL in df.columns:
        no_contract = df[df[CONTRACT_COL].map(normalize_key).isin({"sem contrato", "nao informado", "não informado", "nan", "none", ""})]
        if not no_contract.empty:
            rows.append({"Prioridade": "Alta", "Tipo": "Governança", "Achado": f"{len(no_contract)} lançamento(s) sem contrato identificado.", "Impacto": money(no_contract[VALUE_COL].sum()), "Ação sugerida": "Regularizar vínculo contrato-serviço-fatura antes de usar o dado para auditoria executiva."})

    if SERVICE_COL in df.columns and MONTH_COL in df.columns:
        pivot = df.groupby([SERVICE_COL, MONTH_COL])[VALUE_COL].sum().reset_index()
        service_counts = pivot.groupby(SERVICE_COL)[MONTH_COL].nunique()
        recurring = service_counts[service_counts >= 3].index
        if len(recurring):
            recent_month = sorted(df[MONTH_COL].dropna().astype(str).unique())[-1]
            rec = pivot[pivot[SERVICE_COL].isin(recurring)].copy()
            base_avg = rec[rec[MONTH_COL] != recent_month].groupby(SERVICE_COL)[VALUE_COL].mean()
            recent = rec[rec[MONTH_COL] == recent_month].set_index(SERVICE_COL)[VALUE_COL]
            anomaly = ((recent - base_avg) / base_avg.replace(0, pd.NA) * 100).dropna().sort_values(ascending=False)
            if not anomaly.empty and anomaly.iloc[0] >= 25:
                serv = clean_text(anomaly.index[0])
                rows.append({"Prioridade": "Alta", "Tipo": "Anomalia", "Achado": f"{short_label(serv, 50)} subiu {pct(float(anomaly.iloc[0]))} contra sua média histórica.", "Impacto": money(float(recent.loc[anomaly.index[0]])), "Ação sugerida": "Conferir fatura, franquia, reajuste e consumo associado ao serviço."})

    if not rows:
        rows.append({"Prioridade": "Baixa", "Tipo": "Leitura geral", "Achado": "Nenhum alerta crítico automático foi encontrado no contexto filtrado.", "Impacto": money(total), "Ação sugerida": "Use os rankings e o detalhes para validar concentração, tendência e qualidade dos dados."})
    return enrich_priority_columns(pd.DataFrame(rows), "alerta")


def build_opportunities(df: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    if df.empty or VALUE_COL not in df.columns:
        return pd.DataFrame(columns=["Prioridade", "Oportunidade", "Estimativa", "Base", "Próximo passo"])
    total = float(df[VALUE_COL].sum())

    if SUPPLIER_COL in df.columns:
        sup = _dimension_top(df, SUPPLIER_COL, 5)
        if not sup.empty:
            top = sup.iloc[0]
            share = float(top["Participacao_%"])
            if share >= 30:
                est = float(top[VALUE_COL]) * .05
                rows.append({"Prioridade": "Alta" if share >= 45 else "Média", "Oportunidade": f"Renegociar fornecedor {clean_text(top[SUPPLIER_COL])}", "Estimativa": money(est), "Base": f"{pct(share)} do gasto analisado", "Próximo passo": "Abrir contratos e serviços do fornecedor; simular redução de 3% a 8%."})

    if SERVICE_COL in df.columns:
        serv = _dimension_top(df, SERVICE_COL, 10)
        if not serv.empty:
            top_serv = serv.iloc[0]
            est = float(top_serv[VALUE_COL]) * .03
            rows.append({"Prioridade": "Média", "Oportunidade": f"Revisar serviço de maior custo: {short_label(top_serv[SERVICE_COL], 42)}", "Estimativa": money(est), "Base": money(top_serv[VALUE_COL]), "Próximo passo": "Conferir uso, contrato, fatura e histórico antes de renegociar ou cancelar."})

    if DIFF_COL in df.columns and df[DIFF_COL].abs().sum() > 0:
        diff_pot = float(df[DIFF_COL].abs().sum())
        rows.append({"Prioridade": "Alta", "Oportunidade": "Contestar/validar diferenças acumuladas", "Estimativa": money(diff_pot), "Base": "Soma absoluta das diferenças", "Próximo passo": "Ordenar linhas por maior diferença e abrir evidências de fatura/contrato."})

    if CONTRACT_COL in df.columns:
        no_contract = df[df[CONTRACT_COL].map(normalize_key).isin({"sem contrato", "nao informado", "não informado", "nan", "none", ""})]
        if not no_contract.empty:
            rows.append({"Prioridade": "Alta", "Oportunidade": "Regularizar lançamentos sem contrato", "Estimativa": money(no_contract[VALUE_COL].sum() * .02), "Base": f"{len(no_contract)} lançamento(s)", "Próximo passo": "Vincular serviço/fatura ao contrato correto para habilitar auditoria confiável."})

    if total and not rows:
        rows.append({"Prioridade": "Baixa", "Oportunidade": "Monitoramento mensal preventivo", "Estimativa": money(total * .01), "Base": "1% do gasto filtrado", "Próximo passo": "Acompanhar Pareto, anomalias e diferenças a cada fechamento mensal."})
    return enrich_priority_columns(pd.DataFrame(rows), "oportunidade")


def build_data_quality(base: pd.DataFrame) -> pd.DataFrame:
    rows: List[Dict[str, Any]] = []
    required = [VALUE_COL, MONTH_COL, SUPPLIER_COL, SERVICE_COL, CATEGORY_COL, CONTRACT_COL, INVOICE_COL, CC_COL, CC_ID_COL]
    for col in required:
        if col not in base.columns:
            rows.append({"Severidade": "Alta", "Item": f"Coluna ausente: {col}", "Registros": len(base), "Impacto": "Pode deixar gráficos/detalhamentos incompletos."})
        else:
            blanks = int(base[col].isna().sum() + (base[col].astype(str).str.strip().isin(["", "nan", "None"]).sum()))
            if blanks:
                rows.append({"Severidade": "Média", "Item": f"Valores vazios em {col}", "Registros": blanks, "Impacto": "Pode agrupar dados em 'Não informado' e reduzir precisão da análise."})

    if VALUE_COL in base.columns:
        zeros = int((to_number_safe(base[VALUE_COL]) == 0).sum())
        if zeros:
            rows.append({"Severidade": "Média", "Item": "Valores financeiros zerados", "Registros": zeros, "Impacto": "Pode distorcer ranking, Pareto e médias."})
    if DATE_COL in base.columns:
        invalid = int(base[DATE_COL].isna().sum())
        if invalid:
            rows.append({"Severidade": "Média", "Item": "Datas/períodos inválidos", "Registros": invalid, "Impacto": "Pode afetar tendência, heatmap e comparações mensais."})
    if INVOICE_COL in base.columns and SERVICE_COL in base.columns and VALUE_COL in base.columns:
        dup_cols = [c for c in [INVOICE_COL, SERVICE_COL, VALUE_COL, MONTH_COL] if c in base.columns]
        dups = int(base.duplicated(subset=dup_cols).sum()) if dup_cols else 0
        if dups:
            rows.append({"Severidade": "Baixa", "Item": "Possíveis lançamentos duplicados", "Registros": dups, "Impacto": "Verificar se são parcelas legítimas ou duplicidade de importação."})
    if not rows:
        rows.append({"Severidade": "OK", "Item": "Nenhum problema estrutural crítico encontrado", "Registros": len(base), "Impacto": "Base apta para análise executiva com os campos atuais."})
    return pd.DataFrame(rows)



def _money_to_float(value: Any) -> float:
    """Converte textos como R$ 1.234,56 em número para ordenação de oportunidades."""
    if isinstance(value, (int, float)):
        return float(value)
    txt = clean_text(value, "0")
    txt = re.sub(r"[^0-9,.-]", "", txt).replace(".", "").replace(",", ".")
    try:
        return float(txt)
    except Exception:
        return 0.0


def _top_row(df: pd.DataFrame, group: str) -> Dict[str, Any]:
    if df is None or df.empty or group not in df.columns or VALUE_COL not in df.columns:
        return {}
    g = df.groupby(group, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
    if g.empty:
        return {}
    total = float(g.sum()) or 1.0
    return {"name": clean_text(g.index[0]), "value": float(g.iloc[0]), "share": float(g.iloc[0]) / total * 100}


def build_risk_scores(df: pd.DataFrame, group: Optional[str] = None) -> pd.DataFrame:
    """Score simples e auditável de risco por contrato/fornecedor/serviço."""
    if df is None or df.empty or VALUE_COL not in df.columns:
        return pd.DataFrame(columns=["Item", "Dimensão", "Score_Risco", "Gasto", "Diferença", "Prioridade", "Ação"])
    group = group or (CONTRACT_COL if CONTRACT_COL in df.columns else SUPPLIER_COL)
    if group not in df.columns:
        group = SUPPLIER_COL if SUPPLIER_COL in df.columns else None
    if not group:
        return pd.DataFrame(columns=["Item", "Dimensão", "Score_Risco", "Gasto", "Diferença", "Prioridade", "Ação"])
    agg_map = {"Gasto": (VALUE_COL, "sum"), "Lancamentos": (VALUE_COL, "count")}
    if DIFF_COL in df.columns:
        agg_map["Diferença"] = (DIFF_COL, "sum")
    risk = df.groupby(group, dropna=False).agg(**agg_map).reset_index().rename(columns={group: "Item"})
    if risk.empty:
        return risk
    max_gasto = float(risk["Gasto"].abs().max()) or 1.0
    max_diff = float(risk.get("Diferença", pd.Series([0])).abs().max()) or 1.0
    risk["Score_Risco"] = (risk["Gasto"].abs() / max_gasto * 60) + (risk.get("Diferença", 0).abs() / max_diff * 30)
    if group == CONTRACT_COL:
        mask_no_contract = risk["Item"].map(normalize_key).isin({"sem contrato", "nao informado", "não informado", "none", "nan", ""})
        risk.loc[mask_no_contract, "Score_Risco"] += 20
    risk["Score_Risco"] = risk["Score_Risco"].clip(0, 100).round(0).astype(int)
    risk["Dimensão"] = group
    risk["Prioridade"] = risk["Score_Risco"].map(lambda x: "Alta" if x >= 75 else ("Média" if x >= 45 else "Baixa"))
    risk["Ação"] = risk["Prioridade"].map({"Alta": "Auditar primeiro e abrir evidências no detalhes.", "Média": "Monitorar tendência e validar faturas relacionadas.", "Baixa": "Manter acompanhamento periódico."})
    return risk.sort_values(["Score_Risco", "Gasto"], ascending=False)


def build_economy_scores(df: pd.DataFrame, group: Optional[str] = None) -> pd.DataFrame:
    """Score de economia por fornecedor/serviço com estimativa conservadora."""
    if df is None or df.empty or VALUE_COL not in df.columns:
        return pd.DataFrame(columns=["Item", "Dimensão", "Score_Economia", "Gasto", "Economia_Estimada", "Prioridade", "Ação"])
    group = group or (SUPPLIER_COL if SUPPLIER_COL in df.columns else SERVICE_COL)
    if group not in df.columns:
        return pd.DataFrame(columns=["Item", "Dimensão", "Score_Economia", "Gasto", "Economia_Estimada", "Prioridade", "Ação"])
    eco = df.groupby(group, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).reset_index().rename(columns={group: "Item", VALUE_COL: "Gasto"})
    if eco.empty:
        return eco
    total = float(eco["Gasto"].sum()) or 1.0
    eco["Participacao_%"] = eco["Gasto"] / total * 100
    eco["Economia_Estimada"] = eco["Gasto"] * eco["Participacao_%"].map(lambda p: .06 if p >= 45 else (.04 if p >= 25 else .02))
    max_est = float(eco["Economia_Estimada"].max()) or 1.0
    eco["Score_Economia"] = (eco["Economia_Estimada"] / max_est * 100).clip(0, 100).round(0).astype(int)
    eco["Dimensão"] = group
    eco["Prioridade"] = eco["Score_Economia"].map(lambda x: "Alta" if x >= 70 else ("Média" if x >= 40 else "Baixa"))
    eco["Ação"] = eco["Prioridade"].map({"Alta": "Priorizar negociação e comparação comercial.", "Média": "Revisar contrato, consumo e reajustes.", "Baixa": "Acompanhar no fechamento mensal."})
    return eco.sort_values(["Score_Economia", "Economia_Estimada"], ascending=False)


def make_forecast_line(df: pd.DataFrame, periods: int = 3) -> go.Figure:
    """Forecast simples baseado na variação média recente; serve como sinalização, não previsão oficial."""
    fig = go.Figure()
    if df is None or df.empty or MONTH_COL not in df.columns or VALUE_COL not in df.columns:
        return style_fig(fig, height=430, showlegend=True)
    trend = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().reset_index()
    trend["_sort"] = trend[MONTH_COL].map(period_sort_key)
    trend = trend.sort_values("_sort")
    if trend.empty:
        return style_fig(fig, height=430, showlegend=True)
    x = trend[MONTH_COL].astype(str).tolist()
    y = trend[VALUE_COL].astype(float).tolist()
    fig.add_trace(go.Scatter(x=x, y=y, mode="lines+markers", name="Realizado", hovertemplate="Período: %{x}<br>Gasto: R$ %{y:,.2f}<extra></extra>"))
    if len(y) >= 2:
        recent = y[-4:] if len(y) >= 4 else y
        deltas = [recent[i] - recent[i-1] for i in range(1, len(recent))]
        avg_delta = sum(deltas) / len(deltas) if deltas else 0.0
        fx, fy = [], []
        last = y[-1]
        for i in range(1, periods + 1):
            last = max(0.0, last + avg_delta)
            fx.append(f"Forecast +{i}")
            fy.append(last)
        fig.add_trace(go.Scatter(x=[x[-1]] + fx, y=[y[-1]] + fy, mode="lines+markers", name="Projeção", line=dict(dash="dash"), hovertemplate="Período: %{x}<br>Projetado: R$ %{y:,.2f}<extra></extra>"))
    fig.update_yaxes(tickprefix="R$ ", title_text="")
    return style_fig(fig, height=460, showlegend=True)


def make_score_bar(score_df: pd.DataFrame, score_col: str, title_col: str = "Item", n: int = 8) -> go.Figure:
    fig = go.Figure()
    if score_df is None or score_df.empty or score_col not in score_df.columns or title_col not in score_df.columns:
        return style_fig(fig, height=430, showlegend=False)
    data = score_df.head(n).copy().sort_values(score_col, ascending=True)
    data["_Label"] = data[title_col].map(lambda x: short_label(x, 42))
    fig.add_bar(
        x=data[score_col], y=data["_Label"], orientation="h", marker_color=PALETTE[3],
        customdata=data[[title_col, score_col]], text=data[score_col].map(lambda x: f"{int(x)}"), textposition="outside", cliponaxis=False,
        hovertemplate="%{customdata[0]}<br>Score: %{customdata[1]}<extra></extra>", name="Score",
    )
    fig.update_xaxes(range=[0, 110], title_text="")
    fig.update_yaxes(title_text="")
    return style_fig(fig, height=460, showlegend=False)


def make_cross_heatmap(df: pd.DataFrame, row_field: str, col_field: str, n_rows: int = 8, n_cols: int = 8) -> go.Figure:
    fig = go.Figure()
    if df is None or df.empty or row_field not in df.columns or col_field not in df.columns or VALUE_COL not in df.columns:
        return style_fig(fig, height=520, showlegend=False)
    rows = topn(df, row_field, n_rows)[row_field].tolist()
    cols = topn(df, col_field, n_cols)[col_field].tolist()
    sub = df[df[row_field].isin(rows) & df[col_field].isin(cols)]
    if sub.empty:
        return style_fig(fig, height=520, showlegend=False)
    piv = sub.pivot_table(index=row_field, columns=col_field, values=VALUE_COL, aggfunc="sum", fill_value=0.0)
    fig.add_heatmap(
        z=piv.values, x=[short_label(c, 24) for c in piv.columns], y=[short_label(i, 28) for i in piv.index],
        colorbar=dict(title="Gasto"), hovertemplate="Linha: %{y}<br>Coluna: %{x}<br>Gasto: R$ %{z:,.2f}<extra></extra>",
    )
    return style_fig(fig, height=560, showlegend=False)


def build_tomorrow_actions(df: pd.DataFrame, limit: int = 6) -> pd.DataFrame:
    """Fila única: impacto + risco + urgência = prioridade prática."""
    rows: List[Dict[str, Any]] = []
    alerts = build_smart_alerts(df)
    opps = build_opportunities(df)
    if alerts is not None and not alerts.empty:
        for _, r in alerts.head(4).iterrows():
            impacto = _money_to_float(r.get("Impacto", 0))
            prioridade = clean_text(r.get("Prioridade", "Média"), "Média")
            score = impacto + ({"Alta": 3, "Média": 2, "Baixa": 1}.get(prioridade, 2) * 1000000)
            rows.append({"Prioridade": prioridade, "Ação": clean_text(r.get("Ação sugerida", "Auditar evidências.")), "Por quê": clean_text(r.get("Achado", "Alerta identificado.")), "Onde": clean_text(r.get("Tipo", "Alerta")), "Impacto": clean_text(r.get("Impacto", "Não estimado")), "Score": score})
    if opps is not None and not opps.empty:
        for _, r in opps.head(4).iterrows():
            impacto = _money_to_float(r.get("Estimativa", 0))
            prioridade = clean_text(r.get("Prioridade", "Média"), "Média")
            score = impacto + ({"Alta": 3, "Média": 2, "Baixa": 1}.get(prioridade, 2) * 1000000)
            rows.append({"Prioridade": prioridade, "Ação": clean_text(r.get("Próximo passo", "Validar oportunidade.")), "Por quê": clean_text(r.get("Oportunidade", "Oportunidade identificada.")), "Onde": clean_text(r.get("Base", "Oportunidade")), "Impacto": clean_text(r.get("Estimativa", "Não estimado")), "Score": score})
    out = pd.DataFrame(rows)
    if out.empty:
        return pd.DataFrame([{"Prioridade": "Baixa", "Ação": "Manter acompanhamento mensal.", "Por quê": "Nenhuma ação crítica foi identificada no filtro atual.", "Onde": "Painel", "Impacto": "Monitoramento", "Score": 0}])
    out = out.sort_values("Score", ascending=False).drop(columns=["Score"]).head(limit).reset_index(drop=True)
    out.insert(0, "Ordem", range(1, len(out) + 1))
    return out


def render_radar_executive(df: pd.DataFrame) -> None:
    st.markdown("## 🛰️ Radar executivo")
    render_reading_note("Resumo de decisão: mostra onde está o maior risco, a maior oportunidade, o maior crescimento e o principal bloco financeiro. Use esta seção para decidir o que investigar primeiro.")
    total = float(df[VALUE_COL].sum()) if df is not None and not df.empty and VALUE_COL in df.columns else 0.0
    top_supplier = _top_row(df, SUPPLIER_COL)
    top_contract = _top_row(df, CONTRACT_COL)
    top_service = _top_row(df, SERVICE_COL)
    risk = build_risk_scores(df).head(1)
    eco = build_economy_scores(df).head(1)
    trend_txt = "Sem histórico suficiente"
    if df is not None and MONTH_COL in df.columns and VALUE_COL in df.columns:
        trend = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().reset_index()
        trend["_sort"] = trend[MONTH_COL].map(period_sort_key)
        trend = trend.sort_values("_sort")
        if len(trend) >= 2 and float(trend.iloc[-2][VALUE_COL]) != 0:
            delta = float(trend.iloc[-1][VALUE_COL] - trend.iloc[-2][VALUE_COL])
            trend_txt = f"{'Subiu' if delta >= 0 else 'Caiu'} {money(abs(delta))} no último período"
    def card(icon: str, title: str, value: str, action: str, cls: str = "info") -> str:
        return f"<div class='radar-card {cls}'><strong>{icon} {html.escape(title)}</strong><span>{html.escape(value)}</span><p>{html.escape(action)}</p></div>"
    cards = []
    cards.append(card("💰", "Maior fornecedor", f"{top_supplier.get('name','Sem dado')} · {money(top_supplier.get('value',0))}", "Abrir detalhes e priorizar contratos/serviços desse fornecedor.", "opportunity" if top_supplier.get('share',0) < 35 else "attention"))
    cards.append(card("⚠️", "Maior risco", f"{clean_text(risk.iloc[0]['Item']) if not risk.empty else 'Sem risco'} · score {int(risk.iloc[0]['Score_Risco']) if not risk.empty else 0}", "Auditar evidências, faturas e vínculo contratual primeiro.", "risk"))
    cards.append(card("✅", "Maior oportunidade", f"{clean_text(eco.iloc[0]['Item']) if not eco.empty else 'Sem oportunidade'} · {money(float(eco.iloc[0]['Economia_Estimada'])) if not eco.empty else money(0)}", "Validar economia e negociar com base no volume financeiro.", "opportunity"))
    cards.append(card("📈", "Variação recente", trend_txt, "Explicar a causa por fornecedor, CC, filial e serviço.", "attention"))
    cards.append(card("📄", "Maior contrato", f"{top_contract.get('name','Sem contrato')} · {money(top_contract.get('value',0))}", "Conferir vigência, reajuste e aderência ao gasto realizado.", "info"))
    cards.append(card("🧱", "Serviço mais relevante", f"{top_service.get('name','Sem serviço')} · {money(top_service.get('value',0))}", "Verificar consumo, necessidade e alternativas de otimização.", "info"))
    st.markdown("<div class='radar-grid'>" + "".join(cards) + "</div>", unsafe_allow_html=True)


def render_tomorrow_actions(df: pd.DataFrame, compact: bool = False) -> None:
    st.markdown("## ✅ O que fazer amanhã")
    render_reading_note("Fila prática de decisão. A ordem combina impacto financeiro, risco e urgência para transformar o dashboard em plano de trabalho.")
    actions = build_tomorrow_actions(df, limit=4 if compact else 7)
    st.dataframe(sanitize_dataframe_display(actions), width="stretch", hide_index=True, height=220 if compact else 300)


def render_score_panels(df: pd.DataFrame) -> None:
    st.markdown("## 🎯 Scores de risco e economia")
    render_reading_note("Os scores não substituem auditoria, mas ajudam a priorizar: risco aponta onde conferir primeiro; economia aponta onde negociar ou revisar consumo.")
    risk = build_risk_scores(df)
    eco = build_economy_scores(df)
    c1, c2 = st.columns(2)
    with c1:
        render_native_chart("Score de risco", "Ranking dos itens com maior combinação de gasto, diferença e fragilidade cadastral.", make_score_bar(risk, "Score_Risco"), df, CONTRACT_COL if CONTRACT_COL in df.columns else SUPPLIER_COL, "risk_score", 470)
        with st.expander("Tabela de score de risco", expanded=False):
            show = risk.head(25).copy()
            for col in ["Gasto", "Diferença"]:
                if col in show.columns:
                    show[col] = show[col].map(money)
            st.dataframe(sanitize_dataframe_display(show), width="stretch", hide_index=True)
    with c2:
        render_native_chart("Score de economia", "Ranking das maiores oportunidades conservadoras por volume e concentração.", make_score_bar(eco, "Score_Economia"), df, SUPPLIER_COL if SUPPLIER_COL in df.columns else SERVICE_COL, "economy_score", 470)
        with st.expander("Tabela de score de economia", expanded=False):
            show = eco.head(25).copy()
            for col in ["Gasto", "Economia_Estimada"]:
                if col in show.columns:
                    show[col] = show[col].map(money)
            if "Participacao_%" in show.columns:
                show["Participacao_%"] = show["Participacao_%"].map(pct)
            st.dataframe(sanitize_dataframe_display(show), width="stretch", hide_index=True)


def render_forecast_panel(df: pd.DataFrame) -> None:
    st.markdown("## 🔮 Tendência e forecast")
    render_reading_note("Projeção simples dos próximos períodos usando a variação média recente. Use como alerta preventivo para orçamento, não como previsão contábil oficial.")
    render_native_chart("Forecast de gastos", "Compara realizado com uma projeção curta para antecipar pressão de orçamento.", make_forecast_line(df), df, MONTH_COL, "forecast", 500)


def render_cross_analysis(df: pd.DataFrame) -> None:
    st.markdown("## 🔗 Correlação CC × Filial × Serviço")
    render_reading_note("Cruzamentos mostram onde o gasto nasce: área, localidade e tipo de serviço. Use para separar problema financeiro de problema operacional.")
    c1, c2 = st.columns(2)
    with c1:
        render_native_chart("Centro de custo × serviço", "Mapa de calor entre CC e serviço para identificar áreas que mais consomem cada tipo de telecom.", make_cross_heatmap(df, CC_COL, SERVICE_COL), df, CC_COL, "cc_service_heat", 560)
    with c2:
        render_native_chart("Filial × serviço", "Mapa de calor entre filial e serviço para identificar concentração local e possíveis desvios operacionais.", make_cross_heatmap(df, BRANCH_COL, SERVICE_COL), df, BRANCH_COL, "branch_service_heat", 560)


def render_underutilized_contracts(df: pd.DataFrame) -> None:
    st.markdown("## 🧾 Contratos subutilizados ou sem lastro")
    if "Valor_Contratado" not in df.columns or CONTRACT_COL not in df.columns:
        render_reading_note("A base atual não trouxe valor contratado suficiente para medir subutilização. Ainda assim, use exposição por contrato e diferenças para priorizar revisão contratual.")
        return
    agg = df.groupby(CONTRACT_COL, dropna=False).agg(Contratado=("Valor_Contratado", "sum"), Realizado=(VALUE_COL, "sum"), Lancamentos=(VALUE_COL, "count")).reset_index()
    agg["Uso_%"] = agg.apply(lambda r: (float(r["Realizado"]) / float(r["Contratado"]) * 100) if float(r["Contratado"] or 0) else 0, axis=1)
    low = agg[(agg["Contratado"] > 0) & (agg["Uso_%"] < 60)].sort_values("Contratado", ascending=False).head(20)
    if low.empty:
        render_reading_note("Nenhum contrato com uso abaixo de 60% foi identificado no filtro atual. Continue monitorando após novos uploads.")
        return
    render_reading_note("Contratos com baixo realizado frente ao contratado podem indicar folga, serviço não utilizado ou oportunidade de renegociação.")
    show = low.copy()
    show["Contratado"] = show["Contratado"].map(money)
    show["Realizado"] = show["Realizado"].map(money)
    show["Uso_%"] = show["Uso_%"].map(pct)
    st.dataframe(sanitize_dataframe_display(show), width="stretch", hide_index=True, height=360)

def render_action_cards(rows: pd.DataFrame, kind: str = "alerta", limit: int = 4) -> None:
    """Renderiza alertas/oportunidades em cards explicativos, no padrão plano de ação."""
    if rows is None or rows.empty:
        st.info("Nenhum item relevante foi identificado para este filtro. Continue monitorando a base ou amplie o período para encontrar padrões.")
        return
    st.markdown("<div class='action-list'>", unsafe_allow_html=True)
    for _, row in rows.head(limit).iterrows():
        prioridade = clean_text(row.get("Prioridade", "Média"), "Média")
        prioridade_class = normalize_key(prioridade).replace("é", "e")
        if kind == "oportunidade":
            titulo = clean_text(row.get("Oportunidade", "Oportunidade identificada"), "Oportunidade identificada")
            base = clean_text(row.get("Base", "Sem base informada"), "Sem base informada")
            estimativa = clean_text(row.get("Estimativa", "Sem estimativa"), "Sem estimativa")
            acao = clean_text(row.get("Próximo passo", row.get("Ação sugerida", "Validar evidências no detalhes.")), "Validar evidências no detalhes.")
            corpo = f"""
            <div class='action-card {prioridade_class}'>
              <h4>💰 {_plain(titulo)}</h4>
              <span class='tag'>Prioridade: {_plain(prioridade)}</span><span class='tag'>Estimativa: {_plain(estimativa)}</span>
              <p><strong>Por que apareceu?</strong> {_plain(base)}</p>
              <p><strong>Como agir?</strong> {_plain(acao)}</p>
              <p><strong>Onde validar?</strong> Abra os detalhes do gráfico relacionado e confira fornecedor, contrato, serviço, fatura e lançamento.</p>
            </div>
            """
        else:
            titulo = clean_text(row.get("Achado", "Alerta identificado"), "Alerta identificado")
            tipo = clean_text(row.get("Tipo", "Alerta"), "Alerta")
            impacto = clean_text(row.get("Impacto", "Impacto não calculado"), "Impacto não calculado")
            acao = clean_text(row.get("Ação sugerida", "Validar evidências no detalhes."), "Validar evidências no detalhes.")
            corpo = f"""
            <div class='action-card {prioridade_class}'>
              <h4>🚨 {_plain(titulo)}</h4>
              <span class='tag'>{_plain(tipo)}</span><span class='tag'>Prioridade: {_plain(prioridade)}</span><span class='tag'>Impacto: {_plain(impacto)}</span>
              <p><strong>Por que importa?</strong> Pode indicar risco financeiro, dependência de fornecedor, cobrança fora do padrão ou fragilidade cadastral.</p>
              <p><strong>Como agir?</strong> {_plain(acao)}</p>
              <p><strong>Onde validar?</strong> Use os rankings e o detalhes para chegar até contratos, faturas e lançamentos que sustentam o alerta.</p>
            </div>
            """
        st.markdown(corpo, unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)


def render_region_branch_summary(df: pd.DataFrame) -> None:
    """Card executivo de região/filial para análise geográfica e operacional."""
    if df is None or df.empty or VALUE_COL not in df.columns:
        st.info("ℹ️ Região/filial: não há dados suficientes neste filtro.")
        return
    total = float(df[VALUE_COL].sum()) or 0.0
    parts = []
    if REGION_COL in df.columns:
        rrank = df.groupby(REGION_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
        if not rrank.empty:
            reg = clean_text(rrank.index[0], "Sem região")
            val = float(rrank.iloc[0])
            share = (val / total * 100) if total else 0.0
            parts.append(f"Região líder: <strong>{_plain(reg)}</strong> com {_plain(pct(share))} ({_plain(money(val))}).")
    if BRANCH_COL in df.columns:
        brank = df.groupby(BRANCH_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
        if not brank.empty:
            filial = clean_text(brank.index[0], "Sem filial")
            val = float(brank.iloc[0])
            share = (val / total * 100) if total else 0.0
            parts.append(f"Filial líder: <strong>{_plain(filial)}</strong> com {_plain(pct(share))} ({_plain(money(val))}).")
    if not parts:
        st.info("ℹ️ Região/filial: campos não encontrados ou sem dados úteis.")
        return
    st.markdown(f"""
    <div class='cc-summary'>
      <h4>🗺️ Região e filial em destaque</h4>
      <p>{' '.join(parts)}</p>
      <span class='mini'>Regiões: {_plain(str(df[REGION_COL].nunique() if REGION_COL in df.columns else 0))}</span>
      <span class='mini'>Filiais: {_plain(str(df[BRANCH_COL].nunique() if BRANCH_COL in df.columns else 0))}</span>
      <span class='mini'>Total: {_plain(money(total))}</span>
      <p><strong>Ação:</strong> usar os filtros de Região e Filial para comparar concentração, rateio e possíveis variações locais.</p>
    </div>
    """, unsafe_allow_html=True)


def render_cost_center_summary(df: pd.DataFrame) -> None:
    """Card executivo de Centro de Custo para rateio e cobrança interna."""
    if df is None or df.empty or CC_COL not in df.columns or VALUE_COL not in df.columns:
        st.info("ℹ️ Centro de custo: não há dados suficientes para montar análise de CC neste filtro.")
        return
    total = float(df[VALUE_COL].sum()) or 0.0
    rank = df.groupby(CC_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
    if rank.empty:
        st.info("ℹ️ Centro de custo: nenhum CC válido foi encontrado no filtro atual.")
        return
    top_cc = clean_text(rank.index[0], "Sem centro de custo")
    top_val = float(rank.iloc[0])
    share = (top_val / total * 100) if total else 0.0
    cc_count = int(df[CC_COL].nunique())
    if CC_ID_COL in df.columns:
        id_map = df[[CC_COL, CC_ID_COL]].drop_duplicates(CC_COL)
        cc_id = clean_text(id_map[id_map[CC_COL].map(normalize_key) == normalize_key(top_cc)][CC_ID_COL].iloc[0], "Sem código") if not id_map[id_map[CC_COL].map(normalize_key) == normalize_key(top_cc)].empty else "Sem código"
    else:
        cc_id = "Sem código"
    action = "Validar rateio interno e revisar contratos/serviços do CC líder." if share >= 30 else "Monitorar distribuição por CC e investigar variações mensais relevantes."
    level = "risk-card" if share >= 50 else "attention-card" if share >= 30 else ""
    st.markdown(f"""
    <div class='cc-summary {level}'>
      <h4>🏷️ Centro de custo em destaque</h4>
      <p><strong>{_plain(top_cc)}</strong> concentra {_plain(pct(share))} do gasto filtrado ({_plain(money(top_val))}).</p>
      <span class='mini'>Código: {_plain(cc_id)}</span><span class='mini'>{cc_count} CC(s) no filtro</span><span class='mini'>Total: {_plain(money(total))}</span>
      <p><strong>Ação:</strong> {_plain(action)}</p>
    </div>
    """, unsafe_allow_html=True)



def parse_money_value(value: Any) -> float:
    """Converte textos como R$ 1.234,56 em float para priorização."""
    if isinstance(value, (int, float)):
        return float(value)
    txt = clean_text(value, "0")
    txt = re.sub(r"[^0-9,.-]", "", txt)
    if "," in txt and "." in txt:
        txt = txt.replace(".", "").replace(",", ".")
    elif "," in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except Exception:
        return 0.0


def priority_score(priority: Any, impact: Any = 0, urgency: float = 0) -> int:
    base = {"alta": 70, "média": 45, "media": 45, "baixa": 20, "ok": 5}.get(normalize_key(priority), 25)
    val = abs(parse_money_value(impact))
    financial = 20 if val >= 50000 else 14 if val >= 10000 else 8 if val >= 1000 else 3
    return int(min(100, base + financial + urgency))


def enrich_priority_columns(rows: pd.DataFrame, kind: str) -> pd.DataFrame:
    """Adiciona score e responsáveis sugeridos para transformar achados em gestão."""
    if rows is None or rows.empty:
        return rows
    out = rows.copy()
    if kind == "oportunidade":
        impact_col = "Estimativa"
        owner = "Compras / Telecom"
    else:
        impact_col = "Impacto"
        owner = "Telecom / Financeiro"
    out["Score"] = out.apply(lambda r: priority_score(r.get("Prioridade", "Baixa"), r.get(impact_col, 0)), axis=1)
    out["Responsável sugerido"] = out.apply(
        lambda r: "Cadastro / Governança" if "contrato" in normalize_key(r.to_dict()) else owner,
        axis=1,
    )
    out["Status sugerido"] = "Novo"
    return out.sort_values(["Score"], ascending=False)


def load_action_history() -> Dict[str, Dict[str, str]]:
    try:
        if ACTION_HISTORY_FILE.exists():
            return json.loads(ACTION_HISTORY_FILE.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def save_action_history(history: Dict[str, Dict[str, str]]) -> None:
    try:
        ACTION_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
        ACTION_HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        st.warning(f"Não foi possível salvar o histórico de ações: {exc}")


def action_key(kind: str, row: pd.Series) -> str:
    raw = kind + "|" + "|".join([clean_text(row.get(c, ""), "") for c in ["Tipo", "Achado", "Oportunidade", "Base", "Impacto", "Estimativa"]])
    return re.sub(r"[^a-zA-Z0-9_-]+", "_", raw)[:110]


def render_action_history_controls(rows: pd.DataFrame, kind: str) -> None:
    """Permite registrar acompanhamento local dos achados principais."""
    if rows is None or rows.empty:
        return
    st.markdown("### 🗂️ Histórico de decisões")
    st.caption("Use para acompanhar o que já foi analisado. O registro fica salvo localmente em `dados_atualizados/historico_acoes.json`.")
    history = load_action_history()
    for idx, row in rows.head(8).iterrows():
        key = action_key(kind, row)
        current = history.get(key, {}).get("status", "Novo")
        title = clean_text(row.get("Achado", row.get("Oportunidade", "Item")), "Item")
        c1, c2, c3 = st.columns([.54, .23, .23], vertical_alignment="center")
        c1.markdown(f"**{short_label(title, 72)}**")
        status = c2.selectbox("Status", ["Novo", "Em análise", "Resolvido", "Ignorado"], index=["Novo", "Em análise", "Resolvido", "Ignorado"].index(current) if current in ["Novo", "Em análise", "Resolvido", "Ignorado"] else 0, key=f"hist_status_{kind}_{idx}_{key}")
        owner = clean_text(row.get("Responsável sugerido", "Telecom"), "Telecom")
        if c3.button("Salvar", key=f"hist_save_{kind}_{idx}_{key}", width="stretch"):
            history[key] = {"status": status, "responsavel": owner, "titulo": title}
            save_action_history(history)
            st.toast("Histórico atualizado.", icon="✅")


def render_month_comparison(df: pd.DataFrame) -> None:
    """Comparativo mês atual x anterior com causa provável."""
    if df is None or df.empty or MONTH_COL not in df.columns or VALUE_COL not in df.columns:
        st.info("ℹ️ Comparativo mensal indisponível: faltam período ou valor.")
        return
    trend = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().reset_index()
    trend["_sort"] = trend[MONTH_COL].map(period_sort_key)
    trend = trend.sort_values("_sort")
    if len(trend) < 2:
        st.markdown("<div class='comparison-card'><h4>📅 Comparativo mensal</h4><p>Há apenas um período disponível. O painel mostra o gasto atual, mas ainda não há base para comparar evolução.</p></div>", unsafe_allow_html=True)
        return
    prev_m, curr_m = trend.iloc[-2][MONTH_COL], trend.iloc[-1][MONTH_COL]
    prev_v, curr_v = float(trend.iloc[-2][VALUE_COL]), float(trend.iloc[-1][VALUE_COL])
    delta = curr_v - prev_v
    var = (delta / prev_v * 100) if prev_v else 0.0
    cause = "sem dimensão suficiente para explicar a variação"
    for field in [SUPPLIER_COL, SERVICE_COL, CC_COL, CATEGORY_COL]:
        if field in df.columns:
            comp = df[df[MONTH_COL].isin([prev_m, curr_m])].pivot_table(index=field, columns=MONTH_COL, values=VALUE_COL, aggfunc="sum", fill_value=0)
            if prev_m in comp.columns and curr_m in comp.columns and not comp.empty:
                comp["_delta"] = comp[curr_m] - comp[prev_m]
                top = comp["_delta"].abs().sort_values(ascending=False)
                if not top.empty:
                    item = top.index[0]
                    cause = f"principal variação em {field}: {clean_text(item)} ({money(comp.loc[item, '_delta'])})"
                    break
    cls = "risk-card" if delta > 0 and abs(var) >= 10 else "opportunity-card" if delta < 0 else ""
    st.markdown(f"""
    <div class='comparison-card {cls}'>
      <h4>📅 Mês atual x mês anterior</h4>
      <p><strong>{_plain(curr_m)}</strong>: {_plain(money(curr_v))} | <strong>{_plain(prev_m)}</strong>: {_plain(money(prev_v))}</p>
      <p><strong>Variação:</strong> {_plain(money(delta))} ({_plain(pct(var))}). {_plain(cause)}.</p>
      <p><strong>Ação:</strong> abrir o detalhes da tendência mensal e validar fornecedor, serviço e CC que explicam a mudança.</p>
    </div>
    """, unsafe_allow_html=True)


def render_cost_center_deep_panel(df: pd.DataFrame) -> None:
    """Visão executiva ampliada de Centro de Custo."""
    st.markdown("## 🏷️ Centro de custo")
    render_reading_note("Mostra como os gastos estão distribuídos por CC, qual área concentra custo, onde houve aumento e onde pode existir oportunidade de revisão/rateio.")
    if df is None or df.empty or CC_COL not in df.columns:
        st.info("ℹ️ Não há centro de custo reconhecido para este filtro.")
        return
    total = float(df[VALUE_COL].sum()) if VALUE_COL in df.columns else 0
    rank = _dimension_top(df, CC_COL, 8)
    top = rank.iloc[0] if not rank.empty else None
    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("CCs ativos", str(df[CC_COL].nunique()), "Centros de custo no filtro", status="neutral")
    with c2: kpi_card("Maior CC", short_label(top[CC_COL], 24) if top is not None else "—", money(top[VALUE_COL]) if top is not None else "R$ 0,00", status="delta-warn" if top is not None and float(top.get("Participacao_%",0)) >= 30 else "neutral")
    with c3: kpi_card("Participação líder", pct(top["Participacao_%"]) if top is not None else "0%", "Concentração do principal CC", status="delta-warn" if top is not None and float(top.get("Participacao_%",0)) >= 30 else "delta-up")
    with c4: kpi_card("Base CC", money(total), "Gasto total considerado", status="neutral")
    fig = make_costcenter_bar(df)
    render_native_chart("Top centros de custo", "Ranking dos CCs por gasto total. Use para cobrança interna, rateio e priorização de revisão por área.", fig, df, CC_COL, "cc_deep", 520)


def render_upload_validator(base: pd.DataFrame, tables: Dict[str, pd.DataFrame], issues: List[str]) -> None:
    """Validador executivo da estrutura recebida."""
    st.markdown("## 🧾 Validador de upload")
    render_reading_note("Confere se a planilha possui as colunas que o dashboard reconhece. Colunas novas não são erro; elas só indicam possíveis campos para enriquecer análises futuras.")
    expected = [MONTH_COL, SUPPLIER_COL, CATEGORY_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, VALUE_COL, DIFF_COL, STATUS_COL, REGION_COL, BRANCH_COL, CC_COL, CC_ID_COL]
    recognized = [c for c in expected if c in base.columns]
    missing = [c for c in expected if c not in base.columns]
    known = set(expected + [DATE_COL, "Valor_Contratado", "Linha_ID"])
    new_cols = [c for c in base.columns if c not in known][:20]
    st.markdown(f"""
    <div class='validation-grid'>
      <div class='validation-card'><strong>✅ Reconhecidas</strong><span>{len(recognized)} coluna(s): {_plain(', '.join(recognized[:8]) or 'nenhuma')}</span></div>
      <div class='validation-card'><strong>⚠️ Ausentes</strong><span>{len(missing)} coluna(s): {_plain(', '.join(missing[:8]) or 'nenhuma crítica')}</span></div>
      <div class='validation-card'><strong>ℹ️ Novas</strong><span>{len(new_cols)} campo(s): {_plain(', '.join(new_cols[:8]) or 'nenhuma')}</span></div>
    </div>
    """, unsafe_allow_html=True)
    if issues:
        with st.expander("Avisos de leitura/modelagem", expanded=False):
            for issue in issues:
                st.warning(issue)



def render_director_minutes_card(df: pd.DataFrame, source_name: str) -> None:
    """Card-resumo de diretoria: uma ata visual compacta da situação atual."""
    total = float(df[VALUE_COL].sum()) if df is not None and not df.empty and VALUE_COL in df.columns else 0.0
    linhas = len(df) if df is not None else 0
    fornecedores = int(df[SUPPLIER_COL].nunique()) if df is not None and SUPPLIER_COL in df.columns else 0
    servicos = int(df[SERVICE_COL].nunique()) if df is not None and SERVICE_COL in df.columns else 0
    diff = float(df[DIFF_COL].sum()) if df is not None and DIFF_COL in df.columns else 0.0

    top_supplier_txt = "Sem fornecedor dominante."
    if df is not None and not df.empty and SUPPLIER_COL in df.columns and VALUE_COL in df.columns and total:
        rank = df.groupby(SUPPLIER_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
        if not rank.empty:
            top_supplier_txt = f"{clean_text(rank.index[0])} lidera com {pct(float(rank.iloc[0]) / total * 100)} do gasto."

    trend_txt = "Sem histórico suficiente para comparar períodos."
    if df is not None and not df.empty and MONTH_COL in df.columns and VALUE_COL in df.columns:
        trend = df.groupby(MONTH_COL, dropna=False)[VALUE_COL].sum().reset_index()
        trend["_sort"] = trend[MONTH_COL].map(period_sort_key)
        trend = trend.sort_values("_sort")
        if len(trend) >= 2:
            prev, curr = float(trend.iloc[-2][VALUE_COL]), float(trend.iloc[-1][VALUE_COL])
            delta = curr - prev
            trend_txt = f"Último período {'subiu' if delta >= 0 else 'caiu'} {money(abs(delta))} vs período anterior."

    alerts = build_smart_alerts(df)
    opportunities = build_opportunities(df)
    main_alert = "Sem alerta crítico no filtro atual."
    if not alerts.empty:
        r = alerts.iloc[0]
        main_alert = f"{clean_text(r.get('Prioridade',''))}: {clean_text(r.get('Achado',''))}"
    main_opp = "Sem oportunidade relevante no filtro atual."
    if not opportunities.empty:
        r = opportunities.iloc[0]
        main_opp = f"{clean_text(r.get('Oportunidade',''))} — {clean_text(r.get('Estimativa',''))}"

    status_cls = "risk-card" if abs(diff) > 0 or len(alerts) else "opportunity-card"
    st.markdown(f"""
    <div class='director-minutes {status_cls}'>
      <div class='minutes-head'>
        <span class='badge'>📌 Ata executiva</span>
        <strong>{_plain(source_name)}</strong>
      </div>
      <p><strong>Situação atual:</strong> {_plain(money(total))} analisados em {_plain(str(linhas))} lançamento(s), com {_plain(str(fornecedores))} fornecedor(es) e {_plain(str(servicos))} serviço(s).</p>
      <p><strong>Leitura principal:</strong> {_plain(top_supplier_txt)} {_plain(trend_txt)}</p>
      <p><strong>Ponto de atenção:</strong> {_plain(main_alert)}</p>
      <p><strong>Próxima ação:</strong> {_plain(main_opp)}</p>
    </div>
    """, unsafe_allow_html=True)


def render_director_action_brief(df: pd.DataFrame) -> None:
    """Resumo enxuto de alertas e oportunidades para a visão Diretoria."""
    alerts = build_smart_alerts(df)
    opportunities = build_opportunities(df)
    a1, a2 = st.columns(2)
    with a1:
        st.markdown("### 🚨 Atenções prioritárias")
        render_action_cards(alerts, kind="alerta", limit=2)
    with a2:
        st.markdown("### 💰 Ações recomendadas")
        render_action_cards(opportunities, kind="oportunidade", limit=2)


def render_director_view(df: pd.DataFrame, base: pd.DataFrame, tables: Dict[str, pd.DataFrame], issues: List[str], source_name: str) -> None:
    """Visão Diretoria enxuta: decisão, tendência, concentração e ação.

    Evita repetir diagnóstico, radar, rankings e tabelas auditáveis que já existem
    na visão Analítica. A diretoria vê apenas o que ajuda a decidir.
    """
    st.markdown("<div class='mode-card'><h4>👔 Visão Diretoria</h4><p>Leitura executiva com foco em decisão: gasto, tendência, concentração e próximos passos.</p></div>", unsafe_allow_html=True)

    render_director_minutes_card(df, source_name)

    st.markdown("<div class='compact-section-title'>✅ Prioridades de decisão</div>", unsafe_allow_html=True)
    render_director_action_brief(df)

    st.markdown("<div class='compact-section-title'>📈 Tendência financeira</div>", unsafe_allow_html=True)
    render_native_chart(
        "Evolução e projeção",
        "Realizado por período e projeção curta para apoiar orçamento.",
        make_forecast_line(df),
        df,
        MONTH_COL,
        "director_forecast",
        520,
    )

    st.markdown("<div class='compact-section-title'>🏢 Concentração principal</div>", unsafe_allow_html=True)
    render_chart_grid([
        {"title": "Top fornecedores", "desc": "Quem mais concentra gasto no recorte.", "fig": make_bar(df, SUPPLIER_COL, "", n=5, horizontal=True), "data": df, "filter_field": SUPPLIER_COL, "key": "director_supplier_essential", "height": 440},
        {"title": "Top categorias", "desc": "Quais grupos de serviço mais pesam no custo.", "fig": make_share_bar(df, CATEGORY_COL), "data": df, "filter_field": CATEGORY_COL, "key": "director_category_essential", "height": 440},
    ])

    with st.expander("🗺️ Ver região, filial e centro de custo", expanded=False):
        st.caption("Use quando precisar sair da decisão executiva e localizar onde o custo está alocado.")
        render_chart_grid([
            {"title": "Gasto por filial", "desc": "Unidades com maior gasto.", "fig": make_bar(df, BRANCH_COL, "", n=7, horizontal=True), "data": df, "filter_field": BRANCH_COL, "key": "director_branch", "height": 420},
            {"title": "Top centros de custo", "desc": "Áreas que absorvem maior custo.", "fig": make_costcenter_bar(df), "data": df, "filter_field": CC_COL, "key": "director_cc_essential", "height": 420},
        ])

    with st.expander("📄 Ata executiva pronta para reunião", expanded=False):
        report = generate_executive_report(df, source_name)
        st.download_button("⬇️ Baixar ata/resumo em Markdown", report.encode("utf-8"), file_name="ata_executiva_telecom.md", mime="text/markdown", width="stretch")
        render_markdown_box(report)

def render_analyst_view(df: pd.DataFrame, base: pd.DataFrame, tables: Dict[str, pd.DataFrame], issues: List[str], source_name: str) -> None:
    st.markdown("<div class='mode-card'><h4>🧑‍💻 Visão Analítica</h4><p>Investigação operacional: composição do gasto, serviços, tendência, contratos e base auditável. Visões avançadas ficam recolhidas para reduzir ruído.</p></div>", unsafe_allow_html=True)

def render_diagnostic_executive(df: pd.DataFrame) -> None:
    st.markdown("## 🩺 Diagnóstico executivo")
    render_reading_note("Esta seção resume automaticamente o que mais merece atenção no período filtrado: concentração, variação, diferenças, governança e oportunidades de economia.")
    alerts = build_smart_alerts(df)
    opportunities = build_opportunities(df)
    total = df[VALUE_COL].sum() if VALUE_COL in df.columns and not df.empty else 0
    est = 0.0
    if not opportunities.empty and "Estimativa" in opportunities.columns:
        est = sum(to_number_safe(opportunities["Estimativa"])) if False else 0.0
    # estimativa numérica simples baseada em oportunidades calculadas acima
    est = abs(df[DIFF_COL].sum()) if DIFF_COL in df.columns else 0
    if SUPPLIER_COL in df.columns and not df.empty:
        top_sup = _dimension_top(df, SUPPLIER_COL, 1)
        if not top_sup.empty:
            est += float(top_sup.iloc[0][VALUE_COL]) * .03

    c1, c2, c3, c4 = st.columns(4)
    with c1: kpi_card("Alertas gerados", str(len(alerts)), "Regras automáticas sobre o contexto", meta="Alta/Média/Baixa", status="delta-warn" if len(alerts) else "neutral")
    with c2: kpi_card("Economia estimada", money(est), "Estimativa conservadora inicial", meta="Validar antes de usar", status="delta-up")
    with c3: kpi_card("Concentração", money(total), "Base financeira analisada", meta=f"{len(df):,} lançamentos".replace(",", "."), status="neutral")
    with c4: kpi_card("Qualidade de leitura", "Auditável", "Dados rastreáveis até lançamento", meta="Use os detalhes", status="delta-up")

    d1, d2 = st.columns([.54, .46])
    with d1:
        st.markdown("### 🚨 Alertas inteligentes")
        render_reading_note("Leia como uma fila de investigação. **Prioridade alta não significa erro confirmado**; significa que o ponto merece conferência primeiro. Cada card explica por que apareceu, o impacto e onde validar.")
        render_action_cards(alerts, kind="alerta", limit=3)
        with st.expander("Tabela auditável de alertas", expanded=False):
            st.dataframe(sanitize_dataframe_display(alerts), width="stretch", hide_index=True, height=260)
            render_action_history_controls(alerts, "alerta")
    with d2:
        st.markdown("### 💰 Oportunidades")
        render_reading_note("Estimativas são conservadoras e servem para priorizar negociação, contestação ou saneamento de cadastro. Use os cards como plano de ação e a tabela como evidência exportável.")
        render_action_cards(opportunities, kind="oportunidade", limit=3)
        with st.expander("Tabela auditável de oportunidades", expanded=False):
            st.dataframe(sanitize_dataframe_display(opportunities), width="stretch", hide_index=True, height=260)
            render_action_history_controls(opportunities, "oportunidade")


def render_opportunity_panel(df: pd.DataFrame) -> None:
    st.markdown("## 💰 Plano de ação")
    render_reading_note("Área focada em **decisão e execução**: renegociar, contestar, regularizar, revisar ou monitorar. Os cards explicam **por que**, **onde validar** e **como agir**.")
    opportunities = build_opportunities(df)
    render_action_cards(opportunities, kind="oportunidade", limit=6)
    if opportunities.empty:
        return
    with st.expander("📋 Evidências e exportação do plano", expanded=False):
        st.dataframe(sanitize_dataframe_display(opportunities), width="stretch", hide_index=True, height=280)
        render_action_history_controls(opportunities, "oportunidade_plano")
        st.download_button(
            "⬇️ Baixar oportunidades em CSV",
            opportunities.to_csv(index=False).encode("utf-8-sig"),
            file_name="oportunidades_finops_telecom.csv",
            mime="text/csv",
            width="stretch",
        )


def render_data_quality_panel(base: pd.DataFrame, issues: List[str]) -> None:
    st.markdown("## 🧪 Qualidade dos dados")
    render_reading_note("Esta leitura ajuda a saber se a planilha está pronta para auditoria. Problemas de qualidade podem explicar gráficos vazios, filtros estranhos ou detalhamentos incompletos.")
    quality = build_data_quality(base)
    if issues:
        with st.expander("Avisos técnicos da modelagem", expanded=False):
            for i in issues:
                st.warning(i)
    st.dataframe(sanitize_dataframe_display(quality), width="stretch", hide_index=True, height=260)
    st.download_button(
        "⬇️ Baixar diagnóstico de qualidade",
        quality.to_csv(index=False).encode("utf-8-sig"),
        file_name="qualidade_dados_telecom.csv",
        mime="text/csv",
        width="stretch",
    )


def generate_executive_report(df: pd.DataFrame, source_name: str) -> str:
    total = df[VALUE_COL].sum() if VALUE_COL in df.columns and not df.empty else 0
    alerts = build_smart_alerts(df)
    opportunities = build_opportunities(df)
    top_sup = _dimension_top(df, SUPPLIER_COL, 5) if SUPPLIER_COL in df.columns else pd.DataFrame()
    top_serv = _dimension_top(df, SERVICE_COL, 5) if SERVICE_COL in df.columns else pd.DataFrame()
    lines = [
        f"# Relatório executivo — {APP_TITLE}",
        "",
        f"Base: {source_name}",
        f"Gasto analisado: {money(total)}",
        f"Lançamentos: {len(df):,}".replace(",", "."),
        "",
        "## 5 principais alertas",
    ]
    for _, row in alerts.head(5).iterrows():
        lines.append(f"- **{row.get('Prioridade','')} / {row.get('Tipo','')}** — {row.get('Achado','')} Impacto: {row.get('Impacto','')}. Ação: {row.get('Ação sugerida','')}")
    lines += ["", "## 5 oportunidades recomendadas"]
    for _, row in opportunities.head(5).iterrows():
        lines.append(f"- **{row.get('Prioridade','')}** — {row.get('Oportunidade','')} | Estimativa: {row.get('Estimativa','')} | Próximo passo: {row.get('Próximo passo','')}")
    if not top_sup.empty:
        lines += ["", "## Top fornecedores"]
        for _, r in top_sup.iterrows():
            lines.append(f"- {r[SUPPLIER_COL]}: {money(r[VALUE_COL])} ({pct(r['Participacao_%'])})")
    if not top_serv.empty:
        lines += ["", "## Top serviços"]
        for _, r in top_serv.iterrows():
            lines.append(f"- {r[SERVICE_COL]}: {money(r[VALUE_COL])} ({pct(r['Participacao_%'])})")
    lines += ["", "## Próximas ações", "1. Abrir detalhes nos fornecedores e serviços de maior impacto.", "2. Validar diferenças e anomalias com faturas/contratos.", "3. Registrar ações de negociação, contestação ou saneamento cadastral."]
    return "\n".join(lines)


def render_export_center(df: pd.DataFrame, source_name: str) -> None:
    st.markdown("## 📤 Exportação executiva")
    render_reading_note("Gere um resumo em Markdown para enviar por e-mail, colar em ata ou transformar em PDF. O conteúdo respeita os filtros atuais do painel.")
    report = generate_executive_report(df, source_name)
    st.download_button(
        "⬇️ Baixar relatório executivo (.md)",
        report.encode("utf-8"),
        file_name="relatorio_executivo_tem_telecom.md",
        mime="text/markdown",
        width="stretch",
    )
    with st.expander("Prévia do relatório", expanded=False):
        render_markdown_box(report)

# -----------------------------
# AI
# -----------------------------
def local_summary(df: pd.DataFrame) -> str:
    if df.empty: return "Não há dados para resumir com os filtros atuais."
    total = df[VALUE_COL].sum(); linhas = len(df)
    top_f = df.groupby(SUPPLIER_COL)[VALUE_COL].sum().sort_values(ascending=False).head(1)
    top_c = df.groupby(CATEGORY_COL)[VALUE_COL].sum().sort_values(ascending=False).head(1)
    parts = [f"Foram analisados {linhas} lançamento(s), somando {money(total)}."]
    if not top_f.empty: parts.append(f"Maior fornecedor: {clean_text(top_f.index[0], 'Sem fornecedor')} ({money(top_f.iloc[0])}).")
    if not top_c.empty: parts.append(f"Categoria mais relevante: {clean_text(top_c.index[0], 'Sem categoria')} ({money(top_c.iloc[0])}).")
    if DIFF_COL in df and abs(df[DIFF_COL].sum()) > 0: parts.append(f"Diferença acumulada: {money(df[DIFF_COL].sum())}; validar faturas e contratos.")
    parts.append("Use os filtros do painel ou selecione pontos nos gráficos para refinar a análise.")
    return " ".join(parts)

def _top_items_text(df: pd.DataFrame, field: str, limit: int = 5) -> str:
    if df is None or df.empty or field not in df.columns or VALUE_COL not in df.columns:
        return "sem dados"
    try:
        temp = df[[field, VALUE_COL]].copy()
        temp[field] = temp[field].map(lambda x: clean_text(x, "Sem identificação"))
        g = temp.groupby(field, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(limit)
        if g.empty:
            return "sem dados"
        return "; ".join([f"{clean_text(idx, 'Sem identificação')}: {money(val)}" for idx, val in g.items()])
    except Exception:
        return "sem dados"



def _top_items_records(df: pd.DataFrame, field: str, limit: int = 3) -> List[Tuple[str, float]]:
    """Retorna ranking estruturado para renderização visual, sem depender de Markdown."""
    if df is None or df.empty or field not in df.columns or VALUE_COL not in df.columns:
        return []
    try:
        temp = df[[field, VALUE_COL]].copy()
        temp[field] = temp[field].map(lambda x: clean_text(x, "Sem identificação"))
        temp[VALUE_COL] = safe_numeric_series(temp, VALUE_COL)
        g = temp.groupby(field, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(limit)
        return [(clean_text(idx, "Sem identificação"), float(val)) for idx, val in g.items()]
    except Exception:
        return []


def _ranking_cards_html(items: List[Tuple[str, float]]) -> str:
    if not items:
        return '<div class="smart-rank-item"><div class="name">Sem dados</div><div class="value">—</div></div>'
    cards = []
    for name, value in items:
        cards.append(
            '<div class="smart-rank-item">'
            f'<div class="name">{html.escape(short_label(name, 42))}</div>'
            f'<div class="value">{html.escape(money(value))}</div>'
            '</div>'
        )
    return "".join(cards)


def render_dashboard_smart_summary(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> None:
    """Renderiza o resumo do painel com HTML controlado, evitando Markdown quebrado e R$ como LaTeX."""
    df = full_df if isinstance(full_df, pd.DataFrame) else pd.DataFrame()
    visible = visible_df if isinstance(visible_df, pd.DataFrame) else df
    if visible is None or visible.empty:
        with st.container(border=True):
            st.info("Sem dados no recorte atual para gerar o resumo inteligente.")
        return

    total = float(visible[VALUE_COL].sum()) if VALUE_COL in visible.columns else 0.0
    diff = float(visible[DIFF_COL].sum()) if DIFF_COL in visible.columns else 0.0
    fornecedores = _top_items_records(visible, SUPPLIER_COL, 3) if SUPPLIER_COL in visible.columns else []
    filiais = _top_items_records(visible, BRANCH_COL, 3) if BRANCH_COL in visible.columns else []

    html_block = f"""
    <div class="smart-summary-card">
        <div class="smart-summary-grid">
            <div class="smart-mini-kpi"><small>Lançamentos analisados</small><strong>{len(visible)}</strong></div>
            <div class="smart-mini-kpi"><small>Total filtrado</small><strong>{html.escape(money(total))}</strong></div>
            <div class="smart-mini-kpi"><small>Diferença acumulada</small><strong>{html.escape(money(diff))}</strong></div>
            <div class="smart-mini-kpi"><small>Recorte atual</small><strong>{html.escape(_period_range_text(visible) or 'Base completa')}</strong></div>
        </div>
        <div class="smart-section-title">Top fornecedores</div>
        <div class="smart-ranking">{_ranking_cards_html(fornecedores)}</div>
        <div class="smart-section-title">Top filiais</div>
        <div class="smart-ranking">{_ranking_cards_html(filiais)}</div>
        <div class="smart-hint">Abra a IA flutuante para detalhar por fornecedor, filial, contrato, serviço ou período.</div>
    </div>
    """
    st.markdown(html_block, unsafe_allow_html=True)




def _question_tokens(question: Any) -> List[str]:
    """Tokens normalizados da pergunta para tolerar pequenos erros de digitação."""
    return [t for t in re.findall(r"[a-z0-9]{3,}", normalize_key(question))]


STOPWORD_TOKENS = {
    "quanto", "quantos", "quantas", "gastei", "paguei", "valor", "total", "gasto", "gastos",
    "com", "por", "para", "dos", "das", "que", "qual", "quais", "fale", "sobre", "listar", "mostre",
    "mes", "mês", "periodo", "período", "referencia", "referência", "vencimento", "contrato", "contratos",
    "fatura", "faturas", "servico", "serviço", "servicos", "serviços", "operadora", "operadoras"
}


def _known_entities(df: pd.DataFrame, limit_per_field: int = 1200) -> Dict[str, List[str]]:
    """Dicionário vivo da base: nada de lista fixa de UNIFIQUE/GNET/ALGAR."""
    out: Dict[str, List[str]] = {}
    if df is None or df.empty:
        return out
    for field in [SUPPLIER_COL, BRANCH_COL, REGION_COL, CC_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, CATEGORY_COL]:
        if field not in df.columns:
            continue
        vals = []
        for v in df[field].dropna().astype(str).unique().tolist():
            txt = clean_text(v, "")
            if txt and txt.lower() not in {"nan", "none", "null"}:
                vals.append(txt)
        out[field] = sorted(set(vals), key=lambda x: (-len(x), normalize_key(x)))[:limit_per_field]
    return out


def _df_signature_for_ai(df: pd.DataFrame) -> str:
    """Assinatura leve para cachear entidades sem reprocessar a cada pergunta."""
    if df is None or df.empty:
        return "empty"
    cols = [c for c in [SUPPLIER_COL, BRANCH_COL, REGION_COL, CC_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, CATEGORY_COL, VALUE_COL, MONTH_COL] if c in df.columns]
    raw = f"{len(df)}|{','.join(cols)}"
    try:
        raw += "|" + str(pd.util.hash_pandas_object(df[cols].head(250), index=True).sum())
        raw += "|" + str(pd.util.hash_pandas_object(df[cols].tail(250), index=True).sum())
    except Exception:
        pass
    return hashlib.md5(raw.encode("utf-8", errors="ignore")).hexdigest()


@st.cache_data(show_spinner=False, ttl=600)
def _known_entities_cached(signature: str, payload: Dict[str, List[str]]) -> Dict[str, List[str]]:
    # O payload já vem pronto; o cache evita reconstruir e resortear em reruns rápidos.
    return payload


def _ai_entities(df: pd.DataFrame) -> Dict[str, List[str]]:
    return _known_entities_cached(_df_signature_for_ai(df), _known_entities(df, limit_per_field=450))


def _has_fuzzy_word(question: Any, words: Iterable[str], cutoff: float = 0.78) -> bool:
    """Confere termos por similaridade: contratus≈contratos, unifuque≈unifique."""
    q = normalize_key(question)
    normalized_words = [normalize_key(w) for w in words if clean_text(w, "")]
    if any(w and w in q for w in normalized_words):
        return True
    tokens = _question_tokens(question)
    for token in tokens:
        for word in normalized_words:
            if not word:
                continue
            # Para expressões, compara cada parte relevante.
            word_tokens = re.findall(r"[a-z0-9]{3,}", word) or [word]
            if any(difflib.SequenceMatcher(None, token, wt).ratio() >= cutoff for wt in word_tokens):
                return True
    return False


def _format_dates_in_answer(text: Any) -> str:
    """Padroniza datas sem destruir Markdown/quebras de linha da resposta da IA."""
    if text is None:
        return ""
    out = str(text).replace("\r\n", "\n").replace("\r", "\n")

    def repl_full(match: re.Match) -> str:
        year, month, day = match.group(1), match.group(2), match.group(3)
        return f"{int(day):02d}/{int(month):02d}/{year[-2:]}"

    def repl_month(match: re.Match) -> str:
        year, month = match.group(1), match.group(2)
        return f"{int(month):02d}/{year[-2:]}"

    # 2026-02-15 -> 15/02/26; 2026-02 -> 02/26.
    out = re.sub(r"\b(20\d{2})-(0?[1-9]|1[0-2])-(0?[1-9]|[12]\d|3[01])\b", repl_full, out)
    out = re.sub(r"\b(20\d{2})-(0?[1-9]|1[0-2])\b", repl_month, out)
    # 15/02/2026 -> 15/02/26.
    out = re.sub(r"\b(0?[1-9]|[12]\d|3[01])/(0?[1-9]|1[0-2])/(20\d{2})\b", lambda m: f"{int(m.group(1)):02d}/{int(m.group(2)):02d}/{m.group(3)[-2:]}", out)
    # 1/26 -> 01/26 quando for período.
    out = re.sub(r"(?<!\d)([1-9])/(\d{2})(?!\d)", lambda m: f"0{m.group(1)}/{m.group(2)}", out)
    # Corrige artefatos gerados por limpeza antiga: 01 - 26 / 01-26 -> 01/26.
    out = re.sub(r"(?<!\d)(0?[1-9]|1[0-2])\s*-\s*(\d{2})(?!\d)", lambda m: f"{int(m.group(1)):02d}/{m.group(2)}", out)
    return out


def _final_ai_text(text: Any) -> str:
    """Pós-processamento único das respostas: limpas, objetivas e com datas BR."""
    return normalize_markdown(_format_dates_in_answer(text))

def _ai_engine_label() -> str:
    """Indica a camada usada pela IA sem expor segredo/chave."""
    use_gemini = str(os.getenv("AI_USE_GEMINI", "false")).strip().lower() in {"1", "true", "sim", "yes"}
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if use_gemini and api_key and genai is not None:
        return "🟢 IA Premium"
    return "🟡 IA Local rápida"


def _field_alias_from_question(question: str) -> Optional[str]:
    """Detecta a dimensão principal citada na pergunta, com tolerância a digitação."""
    mapping = [
        (SUPPLIER_COL, ["fornecedor", "fornecedores", "operadora", "operadoras", "empresa", "unifique", "gnet", "algar"]),
        (BRANCH_COL, ["filial", "filiais", "unidade", "unidades", "site", "sites"]),
        (REGION_COL, ["regiao", "região", "regional", "estado", "cidade"]),
        (CC_COL, ["centro de custo", "centro custo", "cc", "c.c", "custo"]),
        (SERVICE_COL, ["servico", "serviço", "servicos", "serviços", "produto"]),
        (CONTRACT_COL, ["contrato", "contratos", "contratus", "contratu"]),
        (INVOICE_COL, ["fatura", "faturas", "nota", "nf", "nfs"]),
        (CATEGORY_COL, ["categoria", "tipo"]),
    ]
    for field, words in mapping:
        if field and _has_fuzzy_word(question, words):
            return field
    return None

def _find_entity_context(df: pd.DataFrame, question: str, history: Optional[List[Dict[str, str]]] = None) -> Tuple[Optional[str], Optional[str]]:
    """Identifica entidade citada ou herdada do histórico, aceitando erro de digitação."""
    if df is None or df.empty:
        return None, None
    history_text = " ".join(clean_text(m.get("content"), "") for m in (history or [])[-8:])
    q_text = clean_text(question, "")
    q_norm = normalize_key(q_text)
    pronouns = {"dela", "dele", "disso", "desse", "dessa", "nesse", "nessa", "ele", "ela"}
    use_history = bool(pronouns.intersection(set(q_norm.split())))
    search_norm = normalize_key(q_text + (" " + history_text if use_history else ""))
    q_tokens = [t for t in re.findall(r"[a-z0-9]{3,}", search_norm) if t not in STOPWORD_TOKENS]
    found: List[Tuple[float, str, str]] = []
    for field, values in _ai_entities(df).items():
        if field not in [SUPPLIER_COL, BRANCH_COL, REGION_COL, CC_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, CATEGORY_COL]:
            continue
        for val in values[:450]:
            key = normalize_key(val)
            if not key or len(key) < 3:
                continue
            if key in q_norm:
                found.append((1.0, field, val))
                break
            if use_history and key in search_norm:
                found.append((0.92, field, val))
                break
            value_tokens = re.findall(r"[a-z0-9]{3,}", key)
            score = 0.0
            for qt in q_tokens[:12]:
                for vt in value_tokens[:8]:
                    score = max(score, difflib.SequenceMatcher(None, qt, vt).ratio())
            if score >= 0.78:
                found.append((score, field, val))
                break
    if not found:
        return None, None
    found.sort(reverse=True, key=lambda x: (x[0], len(x[2])))
    return found[0][1], found[0][2]


def _filter_for_entity(df: pd.DataFrame, field: Optional[str], value: Optional[str]) -> pd.DataFrame:
    if df is None or df.empty or not field or not value or field not in df.columns:
        return df if isinstance(df, pd.DataFrame) else pd.DataFrame()
    key = normalize_key(value)
    mask = df[field].map(lambda x: normalize_key(x) == key)
    return df.loc[mask].copy()


def _entity_match_score(question_norm: str, q_tokens: List[str], value: Any) -> float:
    """Pontua entidade citada na pergunta sem deixar um único token genérico vencer.

    Ex.: "Empresa Cliente equipamentos" deve ganhar de "Empresa Cliente CESTARI", porque a entidade
    correta compartilha Empresa Cliente + equipamentos. Isso evita misturar filiais quando a
    pergunta combina operadora + filial + contrato.
    """
    key = normalize_key(value)
    if not key or len(key) < 3:
        return 0.0
    if key in question_norm:
        return 1.0
    value_tokens = [t for t in re.findall(r"[a-z0-9]{3,}", key) if t not in STOPWORD_TOKENS]
    if not value_tokens or not q_tokens:
        return 0.0
    best_by_value_token: List[float] = []
    exact_hits = 0
    for vt in value_tokens[:10]:
        best = 0.0
        for qt in q_tokens[:16]:
            if qt == vt:
                best = 1.0
            elif len(qt) >= 4 and len(vt) >= 4:
                best = max(best, difflib.SequenceMatcher(None, qt, vt).ratio())
        if best >= 0.96:
            exact_hits += 1
        best_by_value_token.append(best)
    coverage = sum(1 for x in best_by_value_token if x >= 0.82) / max(1, min(len(value_tokens), 4))
    exact_bonus = min(exact_hits, 3) * 0.08
    # Penaliza entidades longas em que só "Empresa Cliente" bateu.
    if exact_hits == 1 and len(value_tokens) >= 3 and coverage < 0.5:
        return 0.45
    return min(1.0, coverage + exact_bonus)


def _find_question_entities(df: pd.DataFrame, question: str, history: Optional[List[Dict[str, str]]] = None) -> Dict[str, str]:
    """Encontra múltiplas entidades na mesma pergunta.

    Diferente de _find_entity_context, esta função não escolhe apenas uma coisa.
    Ela permite perguntas como: "contratos da unifique com a Empresa Cliente equipamentos".
    """
    if df is None or df.empty:
        return {}
    q_text = clean_text(question, "")
    q_norm = normalize_key(q_text)
    history_text = " ".join(clean_text(m.get("content"), "") for m in (history or [])[-4:])
    pronouns = {"dela", "dele", "disso", "desse", "dessa", "nesse", "nessa", "ele", "ela"}
    use_history = bool(pronouns.intersection(set(q_norm.split())))
    search_norm = normalize_key(q_text + (" " + history_text if use_history else ""))
    q_tokens = [t for t in re.findall(r"[a-z0-9]{3,}", search_norm) if t not in STOPWORD_TOKENS]

    fields = [SUPPLIER_COL, BRANCH_COL, REGION_COL, CC_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, CATEGORY_COL]
    found: Dict[str, str] = {}
    for field in fields:
        if field not in df.columns:
            continue
        best_score = 0.0
        best_value = None
        for val in _ai_entities(df).get(field, [])[:600]:
            score = _entity_match_score(search_norm, q_tokens, val)
            if score > best_score:
                best_score = score
                best_value = val
        # Campos curtos aceitam um pouco menos; filiais/serviços exigem mais para evitar falso positivo.
        threshold = 0.72 if field in {SUPPLIER_COL, CONTRACT_COL, INVOICE_COL, CATEGORY_COL} else 0.82
        if best_value and best_score >= threshold:
            found[field] = clean_text(best_value)
    return found


def _apply_entity_filters(df: pd.DataFrame, entities: Dict[str, str]) -> pd.DataFrame:
    scoped = df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    for field, value in (entities or {}).items():
        if scoped.empty or field not in scoped.columns:
            continue
        key = normalize_key(value)
        scoped = scoped.loc[scoped[field].map(lambda x: normalize_key(x) == key)].copy()
    return scoped


def _entity_scope_title(entities: Dict[str, str]) -> str:
    if not entities:
        return "base analisada"
    order = [SUPPLIER_COL, BRANCH_COL, CONTRACT_COL, INVOICE_COL, SERVICE_COL, CATEGORY_COL, CC_COL, REGION_COL]
    labels = []
    for field in order:
        if field in entities:
            labels.append(clean_text(entities[field]))
    return " + ".join(labels) if labels else "base analisada"


def _contracts_answer(scoped: pd.DataFrame, entities: Dict[str, str]) -> str:
    if scoped is None or scoped.empty:
        filtros = ", ".join(f"{k}: {v}" for k, v in (entities or {}).items()) or "filtros informados"
        return f"Não encontrei contratos para **{clean_text(filtros)}**.\n\n➡ Confira se a filial/operadora está escrita igual à base ou remova um filtro para validar."
    if CONTRACT_COL not in scoped.columns:
        return "A base não possui coluna de contratos para responder essa pergunta."
    group_cols = [CONTRACT_COL]
    for col in [SUPPLIER_COL, BRANCH_COL, SERVICE_COL, INVOICE_COL]:
        if col in scoped.columns and col not in group_cols:
            group_cols.append(col)
    agg = scoped.groupby(group_cols, dropna=False)[VALUE_COL].agg(["sum", "count"]).reset_index().sort_values("sum", ascending=False)
    if agg.empty:
        return f"Não encontrei contratos para **{_entity_scope_title(entities)}**."
    contratos = agg[CONTRACT_COL].map(lambda x: clean_text(x, "Sem contrato")).replace("Sem contrato", pd.NA).dropna().nunique()
    total = float(scoped[VALUE_COL].sum()) if VALUE_COL in scoped.columns else 0.0
    periodo = _period_range_text(scoped)
    lines = [
        f"### Contratos — {_entity_scope_title(entities)}",
        "",
        f"**Contratos encontrados:** {contratos}",
        f"**Total vinculado:** {money(total)}",
        f"**Lançamentos:** {len(scoped)}",
    ]
    if periodo:
        lines.append(f"**Período:** {periodo}")
    lines += ["", "**Principais contratos**"]
    for _, r in agg.head(8).iterrows():
        contrato = clean_text(r.get(CONTRACT_COL), "Sem contrato")
        filial = clean_text(r.get(BRANCH_COL), "") if BRANCH_COL in agg.columns else ""
        servico = clean_text(r.get(SERVICE_COL), "") if SERVICE_COL in agg.columns else ""
        extra = ""
        if filial and filial != "Sem filial":
            extra += f" · {filial}"
        if servico and servico != "Sem serviço":
            extra += f" · {short_label(servico, 48)}"
        lines.append(f"- **{contrato}**: {money(r['sum'])} · {int(r['count'])} lançamento(s){extra}")
    return "\n".join(lines)


def _local_table_lines(df: pd.DataFrame, field: str, n: int = 5) -> List[str]:
    if df is None or df.empty or field not in df.columns or VALUE_COL not in df.columns:
        return []
    temp = df[[field, VALUE_COL]].copy()
    temp[field] = temp[field].map(lambda x: clean_text(x, "Sem identificação"))
    g = temp.groupby(field, dropna=False)[VALUE_COL].agg(["sum", "count"]).sort_values("sum", ascending=False).head(n)
    total = float(temp[VALUE_COL].sum()) or 0.0
    out = []
    for idx, row in g.iterrows():
        share = (float(row["sum"]) / total * 100) if total else 0
        out.append(f"- **{clean_text(idx, 'Sem identificação')}**: {money(row['sum'])} · {int(row['count'])} lançamento(s) · {pct(share)}")
    return out


def _local_dimension_answer(df: pd.DataFrame, field: str, label: str, question: str = "") -> str:
    if df is None or df.empty:
        return "Não encontrei lançamentos para responder no contexto atual.\n\n➡ Revise os filtros ou carregue uma base com dados."
    lines = _local_table_lines(df, field, 7)
    if not lines:
        return f"Não há dados suficientes de **{label}** para montar ranking.\n\n➡ Verifique se a coluna está presente e preenchida na planilha."
    total = df[VALUE_COL].sum() if VALUE_COL in df.columns else 0
    return "\n".join([
        f"### {label}",
        f"**Total:** {money(total)} · **Lançamentos:** {len(df)}",
        "",
        *lines,
    ])



def _question_intent(question: str) -> str:
    """Classifica a pergunta aceitando pequenos erros de digitação."""
    q = normalize_key(question)
    if any(x in q for x in [
        "quanto gastei", "quanto ja gastei", "quanto já gastei", "quanto paguei",
        "valor gasto", "total gasto", "gasto total", "quanto custou", "soma", "total pago"
    ]) or (_has_fuzzy_word(question, ["quanto", "valor", "total", "soma", "gastei", "paguei"]) and _has_fuzzy_word(question, ["gasto", "pago", "custou", "valor"])):
        return "direct_value"
    if _has_fuzzy_word(question, ["quantas", "quantos", "qtd", "quantidade"]):
        return "direct_count"
    if _has_fuzzy_word(question, ["compare", "comparar", "comparativo", "diferença", "diferenca"]):
        return "comparison"
    if _has_fuzzy_word(question, ["resumo", "analise", "análise", "explique", "economizar", "oportunidade", "risco"]):
        return "analysis"
    return "general"

def _best_entity_match(df: pd.DataFrame, question: str) -> Tuple[Optional[str], Optional[str]]:
    """Busca entidade citada, aceitando pequenos erros de digitação como Gbet -> GNET."""
    if df is None or df.empty:
        return None, None
    q_norm = normalize_key(question)
    q_tokens = [t for t in re.findall(r"[a-z0-9]{3,}", q_norm) if t not in STOPWORD_TOKENS]
    best = (0.0, None, None)
    for field, values in _ai_entities(df).items():
        for val in values:
            key = normalize_key(val)
            if not key or len(key) < 3:
                continue
            score = 0.0
            if key in q_norm:
                score = 1.0
            else:
                value_tokens = [t for t in re.findall(r"[a-z0-9]{3,}", key)]
                for qt in q_tokens:
                    for vt in value_tokens:
                        score = max(score, difflib.SequenceMatcher(None, qt, vt).ratio())
            if score > best[0]:
                best = (score, field, val)
    if best[0] >= 0.74:
        return best[1], best[2]
    return None, None


def _period_range_text(df: pd.DataFrame) -> str:
    if df is None or df.empty or MONTH_COL not in df.columns:
        return ""
    vals = [clean_text(v, "") for v in df[MONTH_COL].dropna().astype(str).tolist() if clean_text(v, "")]
    if not vals:
        return ""
    ordered = sorted(set(vals), key=period_sort_key)
    if len(ordered) == 1:
        return ordered[0]
    return f"{ordered[0]} a {ordered[-1]}"


def _direct_value_answer(df: pd.DataFrame, field: Optional[str], value: Optional[str]) -> str:
    """Resposta objetiva para pergunta de valor, com Markdown limpo para o chat."""
    if df is None or df.empty or VALUE_COL not in df.columns:
        return "Não encontrei lançamentos para calcular esse valor."

    scoped = _filter_for_entity(df, field, value) if field and value else df
    if scoped.empty:
        return f"Não encontrei lançamentos para **{clean_text(value, 'esse item')}**."

    title = clean_text(value, "Total analisado")
    total = float(scoped[VALUE_COL].sum())
    periodo = _period_range_text(scoped)

    lines = [
        f"### {title}",
        "",
        f"**Total gasto:** {money(total)}",
        f"**Lançamentos:** {len(scoped)}",
    ]
    if periodo:
        lines.append(f"**Período:** {periodo}")

    if CATEGORY_COL in scoped.columns:
        cat = scoped.groupby(CATEGORY_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(3)
        if not cat.empty:
            lines += ["", "**Principais categorias**"]
            lines += [f"- {clean_text(k, 'Sem categoria')}: {money(v)}" for k, v in cat.items()]

    if SERVICE_COL in scoped.columns:
        serv = scoped.groupby(SERVICE_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False).head(3)
        if not serv.empty:
            lines += ["", "**Serviços de maior valor**"]
            lines += [f"- {clean_text(k, 'Sem serviço')}: {money(v)}" for k, v in serv.items()]

    return "\n".join(lines)

def _direct_count_answer(df: pd.DataFrame, question: str, field: Optional[str], value: Optional[str]) -> str:
    scoped = _filter_for_entity(df, field, value) if field and value else df
    q = normalize_key(question)
    target = "lançamentos"
    count = len(scoped)
    for col, label, words in [
        (INVOICE_COL, "faturas", ["fatura", "faturas"]),
        (CONTRACT_COL, "contratos", ["contrato", "contratos"]),
        (SERVICE_COL, "serviços", ["servico", "serviço", "servicos", "serviços"]),
    ]:
        if col in scoped.columns and any(w in q for w in words):
            target = label
            count = scoped[col].nunique(dropna=True)
            break
    ctx = f" para **{clean_text(value)}**" if value else ""
    return f"### Resultado\n**{count} {target}**{ctx}."

def local_ai_answer(full_df: pd.DataFrame, question: str = "", visible_df: Optional[pd.DataFrame] = None, history: Optional[List[Dict[str, str]]] = None) -> str:
    """Motor local sempre disponível: responde com Pandas/regras mesmo sem API Key."""
    df = full_df if isinstance(full_df, pd.DataFrame) else pd.DataFrame()
    visible_df = visible_df if isinstance(visible_df, pd.DataFrame) else df
    question = clean_text(question, "")
    q = normalize_key(question)
    if df.empty:
        return "Não há dados carregados para análise.\n\n➡ Adicione uma planilha válida na pasta `uploads` e recarregue o painel."

    intent = _question_intent(question)

    # Primeiro tenta entender combinações de entidades na mesma pergunta.
    # Ex.: "contratos da unifique com a Empresa Cliente equipamentos" = fornecedor + filial.
    entities = _find_question_entities(df, question, history)
    scoped = _apply_entity_filters(df, entities) if entities else df
    if scoped.empty:
        # Se a combinação exata não retornar dados, não mistura com outra filial/operadora.
        if entities and any(w in q for w in ["contrato", "contratos", "fatura", "faturas", "servico", "serviço", "servicos", "serviços"]):
            return _contracts_answer(scoped, entities)
        scoped = df
        entities = {}

    entity_field, entity_value = None, None
    if len(entities) == 1:
        entity_field, entity_value = next(iter(entities.items()))
    elif not entities:
        entity_field, entity_value = _find_entity_context(df, question, history)
        if not entity_value:
            entity_field, entity_value = _best_entity_match(df, question)
        scoped = _filter_for_entity(df, entity_field, entity_value)
        if scoped.empty:
            scoped = df
            entity_field, entity_value = None, None

    if intent == "direct_value":
        if entities:
            return _direct_value_answer(scoped, None, None).replace("### Total analisado", f"### {_entity_scope_title(entities)}", 1)
        return _direct_value_answer(df, entity_field, entity_value)
    if intent == "direct_count":
        if entities and any(w in q for w in ["contrato", "contratos"]):
            return _contracts_answer(scoped, entities)
        return _direct_count_answer(scoped if entities else df, question, entity_field, entity_value)

    scope_name = _entity_scope_title(entities) if entities else clean_text(entity_value, "")
    scope_note = f"Contexto considerado: **{scope_name}**.\n\n" if scope_name else ""

    # Perguntas sobre detalhes da entidade/combinação citada.
    if (entities or entity_value) and any(w in q for w in ["contrato", "contratos", "fatura", "faturas", "servico", "serviço", "servicos", "serviços"]):
        if any(w in q for w in ["contrato", "contratos"]):
            return _contracts_answer(scoped, entities or ({entity_field: entity_value} if entity_field and entity_value else {}))
        parts = [f"### Detalhes de {scope_name}", scope_note.strip()]
        if CONTRACT_COL in scoped.columns:
            parts += ["", "**Contratos principais**", *_local_table_lines(scoped, CONTRACT_COL, 5)]
        if INVOICE_COL in scoped.columns:
            parts += ["", "**Faturas principais**", *_local_table_lines(scoped, INVOICE_COL, 5)]
        if SERVICE_COL in scoped.columns:
            parts += ["", "**Serviços principais**", *_local_table_lines(scoped, SERVICE_COL, 5)]
        parts += ["", "➡ Use esses itens para auditar os maiores valores e conferir divergências." ]
        return "\n".join([p for p in parts if p])

    # Plano de trabalho.
    if any(w in q for w in ["priorizar", "amanha", "amanhã", "primeiro", "agir", "acao", "ação", "proximo passo", "próximo passo"]):
        actions = build_tomorrow_actions(scoped, limit=5)
        if not actions.empty:
            lines = ["### O que priorizar", scope_note + "A fila abaixo combina impacto financeiro, risco e urgência.", ""]
            for _, r in actions.head(5).iterrows():
                lines.append(f"- **{clean_text(r.get('Prioridade'), 'Média')}** · {clean_text(r.get('Ação'), 'Auditar evidências.')} — {clean_text(r.get('Impacto'), 'Impacto não estimado')}")
            lines += ["", "➡ Comece pelos itens de prioridade alta e maior impacto financeiro."]
            return "\n".join(lines)

    if any(w in q for w in ["risco", "riscos", "critico", "crítico", "venc", "dependencia", "dependência"]):
        group = CONTRACT_COL if CONTRACT_COL in scoped.columns else SUPPLIER_COL
        risk = build_risk_scores(scoped, group=group).head(5)
        if not risk.empty:
            lines = ["### Principais riscos", scope_note + "Os itens abaixo merecem validação primeiro.", ""]
            for _, r in risk.iterrows():
                lines.append(f"- **{clean_text(r.get('Item'))}** · score {int(r.get('Score_Risco', 0))} · {money(r.get('Gasto', 0))} · prioridade {clean_text(r.get('Prioridade'))}")
            lines += ["", "➡ Audite faturas, contrato vinculado e diferenças dos itens com maior score."]
            return "\n".join(lines)

    if any(w in q for w in ["economia", "economizar", "oportunidade", "oportunidades", "reduzir", "negociar", "renegociar"]):
        eco = build_economy_scores(scoped, group=SUPPLIER_COL if SUPPLIER_COL in scoped.columns else SERVICE_COL).head(5)
        if not eco.empty:
            lines = ["### Oportunidades de economia", scope_note + "Ranking conservador por volume e concentração financeira.", ""]
            for _, r in eco.iterrows():
                lines.append(f"- **{clean_text(r.get('Item'))}** · economia estimada {money(r.get('Economia_Estimada', 0))} · score {int(r.get('Score_Economia', 0))}")
            lines += ["", "➡ Valide reajustes, serviços com maior custo e contratos negociáveis."]
            return "\n".join(lines)

    if any(w in q for w in ["comparar", "compare", "comparativo", "filtro", "filtrado", "tela"]):
        total_full = df[VALUE_COL].sum() if VALUE_COL in df.columns else 0
        total_visible = visible_df[VALUE_COL].sum() if isinstance(visible_df, pd.DataFrame) and VALUE_COL in visible_df.columns else 0
        share = (total_visible / total_full * 100) if total_full else 0
        return "\n".join([
            "### Comparativo base completa × tela atual",
            f"- Base completa: **{money(total_full)}** em **{len(df)} lançamento(s)**.",
            f"- Tela filtrada: **{money(total_visible)}** em **{len(visible_df)} lançamento(s)**.",
            f"- O recorte representa **{pct(share)}** do total financeiro.",
            "",
            "➡ Use essa diferença para avaliar se o filtro atual representa bem o cenário geral ou apenas um recorte específico."
        ])

    field = _field_alias_from_question(question)
    if field:
        labels = {SUPPLIER_COL: "Fornecedores/operadoras", BRANCH_COL: "Filiais", REGION_COL: "Regiões", CC_COL: "Centros de custo", SERVICE_COL: "Serviços", CONTRACT_COL: "Contratos", INVOICE_COL: "Faturas", CATEGORY_COL: "Categorias"}
        return _local_dimension_answer(scoped, field, labels.get(field, field), question)

    if any(w in q for w in ["maior", "top", "ranking", "gastou", "gasto", "custo", "valor"]):
        return _local_dimension_answer(scoped, SUPPLIER_COL if SUPPLIER_COL in scoped.columns else CATEGORY_COL, "Ranking financeiro", question)

    # Resposta padrão, mas útil.
    summary = local_summary(scoped)
    actions = build_tomorrow_actions(scoped, limit=3)
    if not actions.empty:
        action_lines = [f"- {clean_text(r.get('Ação'))} — {clean_text(r.get('Impacto'))}" for _, r in actions.head(3).iterrows()]
        return summary + "\n\n### Próximos passos sugeridos\n" + "\n".join(action_lines)
    return summary


def _data_scope_context(full_df: pd.DataFrame, visible_df: pd.DataFrame) -> str:
    """Resumo agregado para a IA consultar a base completa sem enviar a planilha inteira."""
    full_df = full_df if isinstance(full_df, pd.DataFrame) else pd.DataFrame()
    visible_df = visible_df if isinstance(visible_df, pd.DataFrame) else pd.DataFrame()

    def block(label: str, frame: pd.DataFrame) -> str:
        if frame.empty:
            return f"### {label}\n- Sem lançamentos."
        total = frame[VALUE_COL].sum() if VALUE_COL in frame.columns else 0
        diff = frame[DIFF_COL].sum() if DIFF_COL in frame.columns else 0
        lines = [
            f"### {label}",
            f"- Lançamentos: {len(frame)}",
            f"- Total financeiro: {money(total)}",
            f"- Diferença acumulada: {money(diff)}",
        ]
        for field, name in [
            (SUPPLIER_COL, "fornecedores"), (REGION_COL, "regiões"), (BRANCH_COL, "filiais"),
            (CC_COL, "centros de custo"), (SERVICE_COL, "serviços"), (CONTRACT_COL, "contratos")
        ]:
            if field in frame.columns:
                lines.append(f"- Top {name}: {_top_items_text(frame, field, 5)}")
        return "\n".join(lines)

    filtros = []
    for col, meta in st.session_state.get("panel_filters", {}).items():
        if isinstance(meta, dict) and meta.get("value") not in (None, "", "Todos"):
            filtros.append(f"{col}: {meta.get('value')}")
    filtros_txt = "; ".join(filtros) if filtros else "sem filtros ativos"
    return "\n\n".join([
        "A IA deve responder usando a base completa como referência e comparar com o recorte filtrado quando fizer sentido.",
        f"Filtros ativos no painel: {filtros_txt}.",
        block("Base completa disponível", full_df),
        block("Recorte visível na tela", visible_df),
    ])


def _conversation_context(messages: Optional[List[Dict[str, str]]], limit: int = 8) -> str:
    """Converte o histórico recente em contexto curto para a IA manter continuidade."""
    if not messages:
        return "Sem histórico anterior."
    recent = messages[-limit:]
    lines = []
    for msg in recent:
        role = "Usuário" if msg.get("role") == "user" else "Assistente"
        text = clean_text(msg.get("content"), "")
        if text:
            lines.append(f"{role}: {text[:1200]}")
    return "\n".join(lines) if lines else "Sem histórico anterior."


def ai_summary(df: pd.DataFrame, question: str = "", visible_df: Optional[pd.DataFrame] = None, history: Optional[List[Dict[str, str]]] = None) -> str:
    """Consulta IA usando base completa + contexto agregado + histórico da conversa."""
    full_df = df if isinstance(df, pd.DataFrame) else pd.DataFrame()
    visible_df = visible_df if isinstance(visible_df, pd.DataFrame) else full_df
    intent = _question_intent(question)
    # Perguntas objetivas são respondidas pela própria base, sem IA externa.
    # Isso evita respostas genéricas/prolixas e garante números auditáveis.
    if intent in {"direct_value", "direct_count", "comparison"}:
        return _final_ai_text(local_ai_answer(full_df, question, visible_df=visible_df, history=history))

    # Performance: por padrão, a IA responde localmente com Pandas/regras.
    # Para usar Gemini nas perguntas abertas, defina AI_USE_GEMINI=true no .env.
    use_gemini = str(os.getenv("AI_USE_GEMINI", "false")).strip().lower() in {"1", "true", "sim", "yes"}
    api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
    if not use_gemini or not api_key or genai is None:
        return _final_ai_text(local_ai_answer(full_df, question, visible_df=visible_df, history=history))
    try:
        context = _data_scope_context(full_df, visible_df)
        convo = _conversation_context(history)
        sample_cols = [c for c in [MONTH_COL, SUPPLIER_COL, REGION_COL, BRANCH_COL, CC_COL, CATEGORY_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, VALUE_COL, DIFF_COL] if c in full_df.columns]
        amostra = sanitize_dataframe_display(full_df[sample_cols].head(30)).to_dict("records") if sample_cols else []
        prompt = f"""
Você é um analista FinOps de Telecom dentro de um SaaS executivo. Responda em português, com estilo direto, útil e parecido com um assistente moderno: claro, contextual e orientado à decisão.

Pergunta atual do usuário:
{question or 'Gere resumo executivo.'}

Histórico recente da conversa:
{convo}

Contexto agregado da base:
{context}

Amostra auditável da base completa:
{amostra}

Regras de resposta:
- Seja direto: responda primeiro exatamente o que foi perguntado.
- Não use abertura genérica como "Resumo analítico".
- Não diga "use os filtros do painel" em toda resposta.
- Use Markdown limpo: títulos curtos, bullets e negrito somente quando ajudar.
- Para pergunta objetiva, limite a resposta a no máximo 80 palavras.
- Responda de forma objetiva, com no máximo 1 título curto e bullets essenciais.
- Interprete erros de digitação comuns do usuário. Exemplos: unifuque = UNIFIQUE; contratus = contratos.
- Datas devem sair em DD/MM/AA; períodos devem sair em MM/AA.
- Só inclua próximo passo quando a pergunta pedir análise, economia, risco ou decisão.
- Use a base completa como referência principal; use o recorte filtrado apenas quando fizer sentido.
- Se a base não tiver uma informação, diga isso claramente e sugira como validar.
""".strip()
        client = genai.Client(api_key=api_key)
        resp = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)
        return _final_ai_text(getattr(resp, "text", None) or local_summary(full_df))
    except Exception as exc:
        fallback = local_ai_answer(full_df, question, visible_df=visible_df, history=history)
        return _final_ai_text(fallback + f"\n\n⚠️ Respondi em modo local porque a IA externa não respondeu agora ({clean_text(type(exc).__name__)}).")

# -----------------------------
# Main
# -----------------------------


def _is_ai_open() -> bool:
    """Estado interno do assistente IA, sem depender de HTML/link na página."""
    return bool(st.session_state.get("ai_drawer_open", False))


def _open_ai() -> None:
    """Atalho do botão flutuante: leva para a aba dedicada do chat."""
    st.session_state["view_mode"] = "Chat IA"
    st.session_state["ai_opening"] = True
    st.rerun()


def _close_ai() -> None:
    """Sai da página de chat e volta para a visão executiva."""
    st.session_state["ai_drawer_open"] = False
    st.session_state["view_mode"] = "Diretoria"
    try:
        if "ai" in st.query_params:
            del st.query_params["ai"]
    except Exception:
        pass
    st.rerun()


def _ai_messages() -> List[Dict[str, str]]:
    msgs = st.session_state.setdefault("floating_ai_messages", [])
    if not isinstance(msgs, list):
        st.session_state["floating_ai_messages"] = []
    return st.session_state["floating_ai_messages"]


def _add_ai_message(role: str, content: str) -> None:
    content = _final_ai_text(content)
    if not content:
        return
    msgs = _ai_messages()
    msgs.append({"role": role, "content": content})
    # Mantém o histórico útil sem deixar a tela pesada.
    st.session_state["floating_ai_messages"] = msgs[-16:]


def _clear_ai_chat() -> None:
    st.session_state["floating_ai_messages"] = []
    st.session_state.pop("floating_ai_answer", None)
    st.rerun()


def _ai_context_markdown(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> str:
    """Monta contexto curto e auditável sobre base completa + recorte visível."""
    visible_df = visible_df if visible_df is not None else full_df
    if full_df is None or full_df.empty:
        return "A base não possui lançamentos disponíveis."
    total_base = full_df[VALUE_COL].sum() if VALUE_COL in full_df.columns else 0
    total_visivel = visible_df[VALUE_COL].sum() if isinstance(visible_df, pd.DataFrame) and VALUE_COL in visible_df.columns else 0
    return (
        f"**Contexto da IA**\n\n"
        f"<span class='ai-mode-badge'>{_ai_engine_label()} · Consulta: base completa + filtros atuais</span>\n\n"
        f"- Base completa: **{len(full_df)} lançamento(s)** / **{money(total_base)}**\n"
        f"- Tela filtrada: **{len(visible_df)} lançamento(s)** / **{money(total_visivel)}**\n"
        f"- Top fornecedores da base: {_top_items_text(full_df, SUPPLIER_COL, 3)}\n"
        f"- Top filiais da base: {_top_items_text(full_df, BRANCH_COL, 3)}\n"
        f"- Top centros de custo da base: {_top_items_text(full_df, CC_COL, 3)}"
    )






def _compact_ai_context_html(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> str:
    """Contexto visível no topo do chat, curto para não poluir a IA."""
    visible_df = visible_df if isinstance(visible_df, pd.DataFrame) else full_df
    chips: List[str] = []
    filtros = st.session_state.get("panel_filters", {}) or {}
    for label_col in [SUPPLIER_COL, BRANCH_COL, CC_COL, CATEGORY_COL, SERVICE_COL]:
        meta = filtros.get(label_col)
        if isinstance(meta, dict) and meta.get("value") not in (None, "", "Todos"):
            chips.append(f"{label_col}: {clean_text(meta.get('value'))}")
    basis = "Vencimento" if st.session_state.get(PERIOD_BASIS_KEY, "vencimento") == "vencimento" else "Referência"
    if isinstance(visible_df, pd.DataFrame) and not visible_df.empty:
        total = visible_df[VALUE_COL].sum() if VALUE_COL in visible_df.columns else 0
        chips.append(f"{len(visible_df)} lanç. · {money(total)}")
        if MONTH_COL in visible_df.columns:
            months = sorted([m for m in visible_df[MONTH_COL].dropna().astype(str).unique() if m])
            if months:
                chips.append(f"Período: {months[0]}" + (f" a {months[-1]}" if len(months) > 1 else ""))
    chips.append(f"Base: {basis}")
    chips.append(_ai_engine_label())
    if not chips:
        chips = ["Sem contexto carregado"]
    html_chips = []
    for i, chip in enumerate(chips[:6]):
        cls = "gemini-context-chip muted" if i >= 3 else "gemini-context-chip"
        html_chips.append(f"<span class='{cls}'>{html.escape(clean_text(chip))}</span>")
    return "<div class='gemini-context'>" + "".join(html_chips) + "</div>"


def _suggested_ai_questions(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> List[str]:
    """Sugestões curtas e clicáveis, baseadas nos dados disponíveis."""
    frame = visible_df if isinstance(visible_df, pd.DataFrame) and not visible_df.empty else full_df
    supplier = "operadora"
    branch = "filial"
    try:
        if isinstance(frame, pd.DataFrame) and not frame.empty:
            if SUPPLIER_COL in frame.columns:
                top_sup = frame.groupby(SUPPLIER_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
                if not top_sup.empty:
                    supplier = clean_text(top_sup.index[0], "operadora")
            if BRANCH_COL in frame.columns:
                top_branch = frame.groupby(BRANCH_COL, dropna=False)[VALUE_COL].sum().sort_values(ascending=False)
                if not top_branch.empty:
                    branch = clean_text(top_branch.index[0], "filial")
    except Exception:
        pass
    return [
        f"Qual o total da {supplier}?",
        f"Quais contratos da {supplier}?",
        "Quais contratos vencem primeiro?",
        f"Resumo da filial {branch}",
    ]


def _render_quick_questions(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> None:
    if _ai_messages():
        return
    st.markdown("<div class='gemini-quick-grid'>", unsafe_allow_html=True)
    suggestions = _suggested_ai_questions(full_df, visible_df)
    cols = st.columns(2)
    for idx, text in enumerate(suggestions[:4]):
        with cols[idx % 2]:
            if st.button(text, key=f"ai_quick_{idx}", width="stretch"):
                _process_ai_question(full_df, visible_df if isinstance(visible_df, pd.DataFrame) else full_df, text)
    st.markdown("</div>", unsafe_allow_html=True)


def _inline_markdown_to_html(text: str) -> str:
    """Markdown inline mínimo e seguro para bolhas do chat."""
    text = html.escape(clean_text(text, ""))
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)
    text = re.sub(r"`([^`]+)`", r"<code>\1</code>", text)
    return text


def _markdown_to_chat_html(markdown_text: str) -> str:
    """Converte um subconjunto de Markdown em HTML seguro para manter balões bonitos.

    Evita misturar st.markdown externo com <div> aberto, que no Streamlit pode
    quebrar listas e compactar tudo em uma linha.
    """
    text = normalize_markdown(markdown_text)
    if not text:
        return ""

    lines = text.splitlines()
    parts: List[str] = []
    list_open = False

    def close_list() -> None:
        nonlocal list_open
        if list_open:
            parts.append("</ul>")
            list_open = False

    paragraph: List[str] = []

    def flush_paragraph() -> None:
        nonlocal paragraph
        if paragraph:
            parts.append("<p>" + "<br>".join(_inline_markdown_to_html(x) for x in paragraph) + "</p>")
            paragraph = []

    for raw in lines:
        line = raw.rstrip()
        stripped = line.strip()
        if not stripped:
            flush_paragraph()
            close_list()
            continue

        heading = re.match(r"^(#{1,3})\s+(.+)$", stripped)
        bullet = re.match(r"^[-•]\s+(.+)$", stripped)

        if heading:
            flush_paragraph()
            close_list()
            level = min(len(heading.group(1)), 3)
            parts.append(f"<h{level}>{_inline_markdown_to_html(heading.group(2))}</h{level}>")
        elif bullet:
            flush_paragraph()
            if not list_open:
                parts.append("<ul>")
                list_open = True
            parts.append("<li>" + _inline_markdown_to_html(bullet.group(1)) + "</li>")
        else:
            close_list()
            paragraph.append(stripped)

    flush_paragraph()
    close_list()
    return "".join(parts)


def _render_ai_messages() -> None:
    """Renderiza a conversa como um único bloco HTML.

    Importante: manter a thread em um único st.markdown evita que o Streamlit
    quebre o container em vários wrappers. Isso estabiliza o scroll da IA e
    impede que o campo de digitação seja empurrado para fora da janela.
    """
    msgs = _ai_messages()
    html_parts: List[str] = ["<div class='gemini-thread' id='gemini-thread-bottom'>"]
    if not msgs:
        intro = _markdown_to_chat_html("Olá! Pergunte algo sobre a base de telecom. ✨")
        html_parts.append(
            f"<div class='gemini-row assistant'><div class='gemini-avatar'>🤖</div><div class='gemini-bubble'>{intro}</div></div>"
        )
    else:
        for msg in msgs:
            role = msg.get("role", "assistant")
            css_role = "user" if role == "user" else "assistant"
            avatar = "👤" if role == "user" else "🤖"
            content = _markdown_to_chat_html(msg.get("content", ""))
            if not content:
                continue
            if css_role == "user":
                html_parts.append(
                    f"<div class='gemini-row user'><div class='gemini-bubble'>{content}</div><div class='gemini-avatar'>{avatar}</div></div>"
                )
            else:
                html_parts.append(
                    f"<div class='gemini-row assistant'><div class='gemini-avatar'>{avatar}</div><div class='gemini-bubble'>{content}</div></div>"
                )
    html_parts.append("</div>")
    st.markdown("".join(html_parts), unsafe_allow_html=True)
    components.html("""
    <script>
    const doc = window.parent.document;
    function scrollChat(){
      const threads = doc.querySelectorAll('.gemini-thread');
      const thread = threads[threads.length - 1];
      if (thread) {
        thread.scrollTop = thread.scrollHeight;
      }
    }
    requestAnimationFrame(scrollChat);
    setTimeout(scrollChat, 80);
    setTimeout(scrollChat, 260);
    </script>
    """, height=0)


def _process_ai_question(full_df: pd.DataFrame, visible_df: pd.DataFrame, question: str) -> None:
    question = clean_text(question, "")
    if not question:
        return
    _add_ai_message("user", question)
    st.session_state["ai_busy"] = True
    try:
        with st.spinner("Analisando dados..."):
            answer = ai_summary(full_df, question, visible_df=visible_df, history=_ai_messages())
        _add_ai_message("assistant", answer)
    finally:
        st.session_state["ai_busy"] = False
    st.rerun()


def _ai_chat_body(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> None:
    visible_df = visible_df if visible_df is not None else full_df
    st.markdown("<div class='gemini-shell'>", unsafe_allow_html=True)
    st.markdown(
        "<div class='gemini-hero'><div>"
        "<div class='gemini-title'>Assistente FinOps</div>"
        "<div class='gemini-subtitle'>Pergunte sobre contratos, faturas, operadoras e oportunidades.</div>"
        "</div></div>",
        unsafe_allow_html=True,
    )
    st.markdown(_compact_ai_context_html(full_df, visible_df), unsafe_allow_html=True)

    if st.session_state.get("ai_busy"):
        st.markdown("<div class='ai-loading-note'><span class='ai-loader-dot'></span> Preparando resposta...</div>", unsafe_allow_html=True)

    _render_quick_questions(full_df, visible_df)
    _render_ai_messages()

    st.markdown("<div class='gemini-input-wrap'>", unsafe_allow_html=True)
    with st.form("floating_ai_form", clear_on_submit=True):
        pergunta = st.text_input(
            "Mensagem para a IA",
            key="floating_ai_input",
            placeholder="Pergunte sobre gastos, faturas, contratos...",
            label_visibility="collapsed",
        )
        enviar = st.form_submit_button("Enviar", type="primary", width="stretch")
    st.markdown("</div>", unsafe_allow_html=True)

    if enviar:
        _process_ai_question(full_df, visible_df, pergunta)

    st.markdown("<div class='ai-mini-actions'>", unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Limpar", key="ai_clear_button", width="stretch"):
            _clear_ai_chat()
    with c2:
        if st.button("Fechar", key="ai_close_button", width="stretch"):
            _close_ai()
    st.markdown("</div>", unsafe_allow_html=True)

    components.html("""
    <script>
    const doc = window.parent.document;
    function focusInput(){
      const input = doc.querySelector('input[aria-label="Mensagem para a IA"], input[placeholder*="Pergunte"]');
      if (input && !doc.activeElement?.matches('input, textarea')) input.focus({preventScroll:true});
    }
    setTimeout(focusInput, 80);
    </script>
    """, height=0)
    st.markdown("</div>", unsafe_allow_html=True)

def render_ai_shortcut(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> None:
    """Compatibilidade: o botão flutuante foi removido; navegação fica na sidebar."""
    try:
        if st.query_params.get("ai") == "open":
            st.session_state["view_mode"] = "Chat IA"
            del st.query_params["ai"]
            st.rerun()
    except Exception:
        pass
    return


def render_floating_ai(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> None:
    """Compatibilidade: mantém o nome antigo, mas agora só renderiza o atalho."""
    render_ai_shortcut(full_df, visible_df)


def _render_chat_messages_native() -> None:
    """Renderiza mensagens no componente nativo do Streamlit, mais estável e sem espaço vazio."""
    msgs = _ai_messages()
    if not msgs:
        st.markdown("""
        <div class='ai-empty-state'>
          <div class='icon'>✨</div>
          <strong>Pronto para analisar a base</strong>
          Faça uma pergunta sobre gastos, contratos, faturas, operadoras, filiais ou períodos.
        </div>
        """, unsafe_allow_html=True)
        return
    for msg in msgs:
        role = "user" if msg.get("role") == "user" else "assistant"
        avatar = "👤" if role == "user" else "🤖"
        with st.chat_message(role, avatar=avatar):
            st.markdown(normalize_markdown(msg.get("content", "")))


def render_ai_page(full_df: pd.DataFrame, visible_df: Optional[pd.DataFrame] = None) -> None:
    """Página exclusiva do Chat IA, estilo GPT, sem overlay e sem áreas vazias."""
    visible_df = visible_df if isinstance(visible_df, pd.DataFrame) else full_df

    st.markdown(
        "<div class='ai-clean-header'>"
        "<div>"
        "<div class='ai-clean-title'>Chat IA FinOps</div>"
        "<div class='ai-clean-subtitle'>Converse com a base completa de telecom. A IA entende operadoras, filiais, contratos, faturas, períodos e perguntas de continuidade.</div>"
        "</div>"
        f"<div class='ai-clean-status'>{_compact_ai_context_html(full_df, visible_df)}</div>"
        "</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='ai-clean-stage'>", unsafe_allow_html=True)

    if not _ai_messages():
        st.markdown("<div class='ai-suggestion-title'>Sugestões rápidas</div>", unsafe_allow_html=True)
        st.markdown("<div class='ai-suggestion-grid'>", unsafe_allow_html=True)
        suggestions = _suggested_ai_questions(full_df, visible_df)
        cols = st.columns(4)
        for idx, text in enumerate(suggestions[:4]):
            with cols[idx % 4]:
                if st.button(text, key=f"ai_page_quick_clean_{idx}", width="stretch"):
                    _process_ai_question(full_df, visible_df, text)
        st.markdown("</div>", unsafe_allow_html=True)

    if st.session_state.get("ai_busy"):
        st.markdown("<div class='ai-loading-note'><span class='ai-loader-dot'></span> Analisando dados...</div>", unsafe_allow_html=True)

    _render_chat_messages_native()
    st.markdown("</div>", unsafe_allow_html=True)

    question = st.chat_input("Pergunte sobre gastos, faturas, contratos, operadoras ou períodos...")
    if question:
        _process_ai_question(full_df, visible_df, question)

    components.html("""
    <script>
    const doc = window.parent.document;
    function scrollToBottom(){ window.parent.scrollTo({ top: doc.body.scrollHeight, behavior: 'smooth' }); }
    setTimeout(scrollToBottom, 80);
    setTimeout(scrollToBottom, 260);
    </script>
    """, height=0)



def kpi_card(label: str, value: str, help_text: str, delta: Optional[str] = None, meta: Optional[str] = None, status: str = "") -> None:
    """Card KPI compatível com as chamadas antigas do dashboard."""
    delta_html = ""
    if delta:
        delta_class = "neg" if str(delta).strip().startswith(("-", "▼")) else ("neu" if status == "neutral" else "")
        delta_html = f"<div class='delta {delta_class}'>{html.escape(str(delta))}</div>"
    meta_html = f"<div class='meta'><span>{html.escape(str(meta))}</span></div>" if meta else ""
    st.markdown(
        f"<div class='kpi {html.escape(str(status))}'>"
        f"<small>{html.escape(str(label))}</small>"
        f"<strong>{html.escape(str(value))}</strong>"
        f"<span>{html.escape(str(help_text))}</span>"
        f"{delta_html}{meta_html}</div>",
        unsafe_allow_html=True,
    )


def render_filters(df: pd.DataFrame) -> None:
    """Renderiza a sidebar principal com upload, filtro de período, filtros globais e modo de uso."""
    st.sidebar.markdown("## 🎛️ Painel de controle")
    st.sidebar.caption("Base, filtros e navegação do dashboard.")

    with st.sidebar.expander("📦 Base de dados", expanded=True):
        uploaded = st.file_uploader(
            "Atualizar base",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="Envie uma planilha completa ou várias planilhas mensais. O dashboard consolida os arquivos e se adapta às colunas disponíveis.",
        )
        if uploaded:
            saved = save_uploaded_files(uploaded)
            st.session_state.uploaded_payloads = saved
            st.session_state.uploaded_names = [name for _, name in saved]
            if len(saved) == 1:
                st.session_state.uploaded_bytes, st.session_state.uploaded_name = saved[0]
            else:
                st.session_state.pop("uploaded_bytes", None)
                st.session_state.uploaded_name = f"{len(saved)} arquivos enviados"
            st.cache_data.clear()
            st.success(f"Base carregada: {len(saved)} arquivo(s)")
            st.rerun()

        available_period_modes = []
        if DUE_DATE_COL in df.columns and df[DUE_DATE_COL].notna().any():
            available_period_modes.append("Vencimento")
        if REF_DATE_COL in df.columns and df[REF_DATE_COL].notna().any():
            available_period_modes.append("Mês de referência")
        if not available_period_modes:
            available_period_modes = ["Vencimento"]

        current_basis = st.session_state.get(PERIOD_BASIS_KEY, available_period_modes[0])
        if current_basis not in available_period_modes:
            current_basis = available_period_modes[0]
            st.session_state[PERIOD_BASIS_KEY] = current_basis

        chosen_basis = st.radio(
            "Filtrar período por",
            available_period_modes,
            index=available_period_modes.index(current_basis),
            horizontal=False,
            help="Define se o filtro Período usa vencimento financeiro ou mês de referência/competência.",
        )
        if chosen_basis != st.session_state.get(PERIOD_BASIS_KEY):
            st.session_state[PERIOD_BASIS_KEY] = chosen_basis
            clear_filter(MONTH_COL, reset_widget=True)
            st.rerun()

        try:
            apply_period_basis(df)
            origem_col = "Periodo_Origem_Ativa" if "Periodo_Origem_Ativa" in df.columns else "Periodo_Origem"
            if origem_col in df.columns and df[origem_col].dropna().any():
                origem = clean_text(df[origem_col].dropna().astype(str).iloc[0], "coluna de data")
                st.caption(f"Período ativo: {chosen_basis} · coluna: {origem}")
        except Exception:
            st.caption(f"Período ativo: {chosen_basis}")

    st.sidebar.markdown("### 🧭 Modo de uso")
    st.sidebar.radio(
        "Escolha a visão",
        ["Diretoria", "Analítica", "Chat IA"],
        key="view_mode",
        horizontal=False,
        label_visibility="collapsed",
        help="Diretoria resume decisões; Analítica investiga dados; Chat IA conversa com a base.",
    )
    st.sidebar.divider()

    with st.sidebar.expander("🔎 Filtros", expanded=True):
        apply_period_basis(df)
        filter_fields = [
            (MONTH_COL, f"Período ({st.session_state.get(PERIOD_BASIS_KEY, 'Vencimento').lower()})"),
            (SUPPLIER_COL, "Fornecedor"),
            (CATEGORY_COL, "Categoria"),
            (REGION_COL, "Região"),
            (BRANCH_COL, "Filial"),
            (CC_COL, "Centro de custo"),
        ]
        for field, label in filter_fields:
            if field not in df.columns:
                continue
            opts = ["Todos"] + filtered_options(df, field)
            current = st.session_state.panel_filters.get(field, {}).get("value", "Todos")
            idx = opts.index(current) if current in opts else 0
            choice = st.selectbox(label, opts, index=idx, key=_manual_filter_key(field))
            if choice == "Todos":
                if field in st.session_state.panel_filters:
                    clear_filter(field, reset_widget=False)
            else:
                set_filter(field, choice, "Filtro manual")

    with st.sidebar.expander("🎯 Filtros ativos", expanded=bool(st.session_state.panel_filters)):
        if st.session_state.panel_filters:
            for field, meta in list(st.session_state.panel_filters.items()):
                value = meta.get("value") if isinstance(meta, dict) else meta
                c1, c2 = st.columns([0.78, 0.22])
                c1.markdown(f"<span class='filter-chip'>{field}: {value}</span>", unsafe_allow_html=True)
                if c2.button("✕", key=f"remove_{field}", help="Remover este filtro"):
                    clear_filter(field)
                    st.rerun()
            st.button("🧹 Limpar todos", width="stretch", on_click=clear_all_filters)
        else:
            st.caption("Nenhum filtro ativo.")

    with st.sidebar.expander("♿ Acessibilidade", expanded=False):
        fonte = st.radio("Tamanho da fonte", ["Normal", "Grande"], horizontal=True)
        alto = st.toggle("Alto contraste")
        if fonte == "Grande" or alto:
            st.markdown(
                f"""<style>{'.block-container {font-size: 1.08rem;}' if fonte == 'Grande' else ''}{'[data-testid="stAppViewContainer"] {filter: contrast(1.12) saturate(1.15);}' if alto else ''}</style>""",
                unsafe_allow_html=True,
            )

def main() -> None:
    inject_css(); init_state()

    try:
        with st.spinner("📦 Carregando e preparando a base de dados..."):
            if "uploaded_payloads" in st.session_state:
                payloads = st.session_state.uploaded_payloads
            elif "uploaded_bytes" in st.session_state:
                payloads = [(st.session_state.uploaded_bytes, st.session_state.get("uploaded_name", "arquivo enviado"))]
            else:
                payloads = load_default_payloads()
            tables, source_name = combine_excel_payloads(payloads)
            base, issues = enrich_model(tables)
            base = ensure_model_quality(base)
    except Exception as e:
        st.error(f"Não foi possível carregar a base: {e}")
        return

    render_filters(base)
    base = apply_period_basis(base)
    view_mode = st.session_state.get("view_mode", "Diretoria")
    df = apply_filters(base)

    if view_mode == "Chat IA":
        render_ai_page(base, df)
        return

    st.markdown(f"""
    <div class='hero'>
      <h1>📊 {APP_TITLE}</h1>
      <p>Dashboard financeiro de telecomunicações baseado na planilha <strong>{source_name}</strong>. Use filtros, clique em barras/pontos dos gráficos para filtrar o painel e abra a seção <strong>▾ Ver mais detalhes</strong> abaixo de cada visual para auditar contratos, faturas, valores e lançamentos.</p>
    </div>
    """, unsafe_allow_html=True)

    with st.expander("🧾 Validador de upload", expanded=False):
        render_upload_validator(base, tables, issues)

    if issues:
        with st.expander("Avisos de leitura da base", expanded=False):
            for i in issues: st.warning(i)

    if st.session_state.panel_filters:
        st.markdown(" ".join([f"<span class='filter-chip'>🎯 {k}: {v['value']}</span>" for k, v in st.session_state.panel_filters.items()]), unsafe_allow_html=True)

    # KPIs executivos — valor + contexto + comparação
    total = df[VALUE_COL].sum() if not df.empty else 0
    diff = df[DIFF_COL].sum() if DIFF_COL in df else 0
    fornecedores = df[SUPPLIER_COL].nunique() if SUPPLIER_COL in df else 0
    trend = pd.Series(dtype=float)
    if not df.empty and MONTH_COL in df:
        trend = df.groupby(MONTH_COL)[VALUE_COL].sum().sort_index()
    var = 0
    delta_txt = "Sem histórico"
    if len(trend) >= 2 and trend.iloc[-2]:
        var = (trend.iloc[-1] - trend.iloc[-2]) / trend.iloc[-2] * 100
        delta_txt = ("▲ " if var >= 0 else "▼ ") + pct(abs(var)) + " vs período anterior"
    meta = total * 0.95 if total else 0
    k1, k2, k3, k4 = st.columns(4)
    with k1: kpi_card("Gasto filtrado", money(total), "Soma financeira dos lançamentos", delta_txt, f"Meta sugerida: {money(meta)}", "delta-up" if var <= 0 else "delta-warn")
    with k2: kpi_card("Lançamentos", f"{len(df):,}".replace(",", "."), "Registros auditáveis", meta=f"{fornecedores} fornecedor(es) na seleção", status="neutral")
    with k3: kpi_card("Variação recente", pct(var), "Último período comparado ao anterior", delta_txt if len(trend) >= 2 else None, status="delta-up" if var <= 0 else "delta-warn")
    with k4: kpi_card("Diferença acumulada", money(diff), "Valor a revisar/validar", "⚠️ revisar" if abs(diff) > 0 else "✓ ok", status="delta-warn" if abs(diff) > 0 else "delta-up")

    if view_mode == "Diretoria":
        render_director_view(df, base, tables, issues, source_name)
        render_ai_shortcut(base, df)
        if st.session_state.detail:
            show_detail_dialog()
        return

    render_analyst_view(df, base, tables, issues, source_name)

    st.markdown("---")
    st.markdown("## 🧭 Visão analítica principal")
    st.caption("Conjunto essencial: composição do gasto, fornecedores, serviços e tendência. Diagnóstico, scores e auditoria ficam recolhidos abaixo.")

    # Primeira faixa: visão de composição e concentração por fornecedor em duas colunas, com detalhes em largura total.
    render_chart_grid([
        {"title": "Gasto por categoria", "desc": "Ranking com participação percentual por tipo/categoria de serviço.", "fig": make_share_bar(df, CATEGORY_COL), "data": df, "filter_field": CATEGORY_COL, "key": "cat", "height": 480},
        {"title": "Top fornecedores", "desc": "Fornecedores com maior valor financeiro no período filtrado.", "fig": make_bar(df, SUPPLIER_COL, "", n=10, horizontal=True), "data": df, "filter_field": SUPPLIER_COL, "key": "supplier", "height": 500},
    ])

    # Pareto e evolução ficam empilhados para evitar leitura espremida.
    st.markdown("### 🧱 Serviços e tendência")
    fig = make_pareto(df, SERVICE_COL)
    render_native_chart("Pareto de serviços", "Serviços com maior impacto no custo. Os rótulos foram abreviados apenas no eixo; o nome completo aparece no tooltip, no filtro e no detalhes.", fig, df, SERVICE_COL, "pareto", 520)

    fig = make_trend(df)
    render_native_chart("Evolução mensal", "Tendência do gasto ao longo dos períodos da base.", fig, df, MONTH_COL, "trend", 500)

    with st.expander("🧩 Auditoria operacional", expanded=False):
        st.caption("Abra apenas quando precisar explicar origem do gasto, diferenças ou alocação interna.")
        field_status = STATUS_COL if STATUS_COL in df.columns else CONTESTED_COL
        render_chart_grid([
            {"title": "Gasto por centro de custo", "desc": "Áreas/códigos que concentram custo.", "fig": make_costcenter_bar(df), "data": df, "filter_field": CC_COL, "key": "cc", "height": 440},
            {"title": "Status financeiro/contratual", "desc": "Valor agrupado pelo status disponível na base.", "fig": make_share_bar(df, field_status), "data": df, "filter_field": field_status, "key": "status", "height": 440},
            {"title": "Gasto por filial", "desc": "Unidades com maior gasto no recorte.", "fig": make_bar(df, BRANCH_COL, "", n=10, horizontal=True), "data": df, "filter_field": BRANCH_COL, "key": "branch", "height": 440},
            {"title": "Diferenças por fornecedor", "desc": "Divergências acumuladas para revisão de faturas e contratos.", "fig": make_diff_bar(df), "data": df, "filter_field": SUPPLIER_COL, "key": "diff_supplier", "height": 440},
        ])
        render_native_chart("Contratado x realizado", "Cascata FinOps entre contratado, ajustes/diferenças e realizado.", make_waterfall(df), df, None, "waterfall", 460)

    with st.expander("🔬 Visuais analíticos avançados", expanded=False):
        st.caption("Visões especializadas para investigação: sazonalidade, exposição por contrato, anomalias e concentração hierárquica.")
        render_chart_grid([
            {"title": "Heatmap fornecedor x mês", "desc": "Mapa de calor para encontrar sazonalidade e concentração mensal por fornecedor.", "fig": make_heatmap_supplier_month(df), "data": df, "filter_field": SUPPLIER_COL, "key": "heat_supplier_month", "height": 520},
            {"title": "Exposição por contrato", "desc": "Ranking dos contratos/faturas que mais concentram gasto no contexto filtrado.", "fig": make_contract_exposure(df), "data": df, "filter_field": CONTRACT_COL if CONTRACT_COL in df.columns else INVOICE_COL, "key": "contract_exposure", "height": 520},
        ])
        render_native_chart("Anomalias de serviços", "Serviços cujo último período ficou acima da média histórica do próprio serviço.", make_anomaly_bar(df), df, SERVICE_COL, "anomaly_service", 520)
        render_native_chart("Treemap fornecedor → serviço", "Composição hierárquica dos maiores blocos de custo por fornecedor e serviço.", make_supplier_service_treemap(df), df, SUPPLIER_COL, "treemap_supplier_service", 600)

    with st.expander("🧠 Diagnóstico, scores e oportunidades", expanded=False):
        render_diagnostic_executive(df)
        render_tomorrow_actions(df, compact=True)
        render_score_panels(df)
        render_opportunity_panel(df)

    with st.expander("🔗 Cruzamentos estratégicos e contratos", expanded=False):
        render_cross_analysis(df)
        render_underutilized_contracts(df)

    with st.expander("🧪 Qualidade dos dados e exportação", expanded=False):
        render_data_quality_panel(base, issues)
        render_export_center(df, source_name)

    st.markdown("## 📋 Base auditável")
    st.caption("Tabela final já respeitando filtros aplicados manualmente ou por seleção nos gráficos.")
    show_cols = [c for c in [MONTH_COL, SUPPLIER_COL, CATEGORY_COL, SERVICE_COL, CONTRACT_COL, INVOICE_COL, VALUE_COL, DIFF_COL, STATUS_COL, REGION_COL, BRANCH_COL, CC_COL, CC_ID_COL] if c in df.columns]
    table = df[show_cols].copy()
    for c in [VALUE_COL, DIFF_COL]:
        if c in table: table[c] = table[c].map(money)
    st.dataframe(sanitize_dataframe_display(table), width="stretch", hide_index=True, height=420)

    with st.expander("🤖 Resumo inteligente do recorte", expanded=False):
        render_dashboard_smart_summary(base, visible_df=df)
    render_ai_shortcut(base, df)

    if st.session_state.detail:
        show_detail_dialog()

if __name__ == "__main__":
    main()
