"""CSS e tema visual global do app.

A ideia é centralizar os ajustes de aparência em uma camada só, evitando CSS espalhado
pelas abas. O CSS usa variáveis nativas do Streamlit para respeitar Light/Dark mode.
"""

import streamlit as st


def inject_design_system_css() -> None:
    """Injeta o CSS premium comum a todas as telas."""
    st.markdown(
        """
        <style>
        :root {
            color-scheme: light dark;
            --ds-primary: var(--primary-color, #7C3AED);
            --ds-bg: var(--background-color, #0B1020);
            --ds-surface: var(--secondary-background-color, #151B2E);
            --ds-text: var(--text-color, #F8FAFC);
            --ds-muted: color-mix(in srgb, var(--text-color) 58%, transparent);
            --ds-border: color-mix(in srgb, var(--text-color) 14%, transparent);
            --ds-border-strong: color-mix(in srgb, var(--primary-color) 42%, transparent);
            --ds-cyan: #38BDF8;
            --ds-emerald: #10B981;
            --ds-amber: #F59E0B;
            --ds-coral: #EF4444;
            --ds-purple: #8B5CF6;
            --ds-radius: 16px;
            --ds-radius-sm: 10px;
            --ds-shadow: 0 14px 34px color-mix(in srgb, #000 20%, transparent);
            --ds-shadow-soft: 0 8px 20px color-mix(in srgb, #000 14%, transparent);
        }

        .block-container {
            max-width: 1500px;
        }

        .ds-page-title {
            font-size: clamp(1.55rem, 2.2vw, 2.25rem);
            line-height: 1.1;
            font-weight: 850;
            letter-spacing: -0.03em;
            margin: 0 0 .25rem 0;
            color: var(--ds-text) !important;
        }
        .ds-page-subtitle {
            font-size: .96rem;
            color: var(--ds-muted) !important;
            margin: 0 0 1.1rem 0;
        }

        .ds-card,
        .ds-metric-card,
        .ds-proactive-card,
        .ds-status-card {
            background: linear-gradient(145deg,
                color-mix(in srgb, var(--ds-surface) 94%, var(--ds-primary) 6%),
                color-mix(in srgb, var(--ds-surface) 92%, var(--ds-cyan) 8%)
            ) !important;
            border: 1px solid var(--ds-border) !important;
            border-radius: var(--ds-radius) !important;
            box-shadow: var(--ds-shadow-soft);
            color: var(--ds-text) !important;
        }

        .ds-metric-card {
            padding: 16px 18px;
            min-height: 126px;
            border-left: 5px solid var(--ds-primary) !important;
        }
        .ds-metric-card.success { border-left-color: var(--ds-emerald) !important; }
        .ds-metric-card.warning { border-left-color: var(--ds-amber) !important; }
        .ds-metric-card.danger { border-left-color: var(--ds-coral) !important; }
        .ds-metric-card.info { border-left-color: var(--ds-cyan) !important; }

        .ds-metric-label {
            font-size: .78rem;
            color: var(--ds-muted) !important;
            text-transform: uppercase;
            letter-spacing: .07em;
            font-weight: 800;
            margin-bottom: 8px;
        }
        .ds-metric-value {
            font-size: clamp(1.35rem, 2vw, 2rem);
            line-height: 1.05;
            color: var(--ds-text) !important;
            font-weight: 900;
            word-break: break-word;
        }
        .ds-metric-delta {
            margin-top: 9px;
            font-size: .86rem;
            color: var(--ds-emerald) !important;
            font-weight: 700;
        }
        .ds-metric-delta.negative { color: var(--ds-coral) !important; }
        .ds-metric-delta.neutral { color: var(--ds-cyan) !important; }

        .ds-status-card {
            padding: 13px 14px;
            margin-bottom: 14px;
        }
        .ds-status-title {
            font-size: .78rem;
            color: var(--ds-cyan) !important;
            font-weight: 850;
            letter-spacing: .07em;
            text-transform: uppercase;
            margin-bottom: 8px;
        }
        .ds-status-line {
            display: flex;
            justify-content: space-between;
            gap: 10px;
            font-size: .86rem;
            padding: 4px 0;
            color: var(--ds-muted) !important;
        }
        .ds-status-line strong { color: var(--ds-text) !important; }

        .ds-filter-chip {
            display: inline-flex;
            align-items: center;
            gap: 6px;
            border-radius: 999px;
            padding: 6px 11px;
            margin: 3px 4px 3px 0;
            background: color-mix(in srgb, var(--ds-primary) 18%, var(--ds-surface));
            border: 1px solid color-mix(in srgb, var(--ds-primary) 38%, transparent);
            color: var(--ds-text) !important;
            font-size: .78rem;
            font-weight: 750;
        }

        .ds-proactive-card {
            padding: 16px 18px;
        }
        .ds-proactive-card ul {
            margin: 0;
            padding-left: 1.15rem;
        }
        .ds-proactive-card li {
            margin: .36rem 0;
            color: var(--ds-text) !important;
        }

        .ds-sidebar-group {
            color: var(--ds-muted) !important;
            font-size: .72rem;
            font-weight: 850;
            letter-spacing: .10em;
            text-transform: uppercase;
            margin: 18px 0 6px 0;
        }

        div[data-testid="stDataFrame"] {
            border-radius: var(--ds-radius) !important;
            overflow: hidden;
            border: 1px solid var(--ds-border) !important;
            box-shadow: var(--ds-shadow-soft);
        }

        div[data-testid="stPlotlyChart"] {
            border-radius: var(--ds-radius) !important;
            overflow: hidden;
            border: 1px solid var(--ds-border) !important;
            box-shadow: var(--ds-shadow-soft);
            background: color-mix(in srgb, var(--ds-surface) 94%, var(--ds-primary) 6%) !important;
            padding: 6px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def inject_visual_emphasis_css() -> None:
    """Refinamento visual adicional: mais diferenciação entre sidebar, filtros, cards e botões."""
    st.markdown(
        """
        <style>
        :root {
            --ds-primary: var(--primary-color, #7C3AED);
            --ds-bg: var(--background-color, #0B1020);
            --ds-surface: var(--secondary-background-color, #151B2E);
            --ds-text: var(--text-color, #F8FAFC);
            --ds-cyan: #22D3EE;
            --ds-blue: #3B82F6;
            --ds-indigo: #6366F1;
            --ds-purple: #8B5CF6;
            --ds-emerald: #10B981;
            --ds-amber: #F59E0B;
            --ds-coral: #EF4444;
            --ds-muted: color-mix(in srgb, var(--text-color) 62%, transparent);
            --ds-border: color-mix(in srgb, var(--text-color) 14%, transparent);
            --ds-strong-border: color-mix(in srgb, var(--primary-color) 46%, transparent);
            --ds-glass: color-mix(in srgb, var(--secondary-background-color) 82%, var(--background-color) 18%);
            --ds-raised: color-mix(in srgb, var(--secondary-background-color) 76%, #7C3AED 10%);
            --ds-shadow-xl: 0 18px 44px color-mix(in srgb, #000 24%, transparent);
            --ds-shadow-glow: 0 14px 34px color-mix(in srgb, var(--primary-color) 20%, transparent);
        }

        /* Fundo com profundidade sutil, sem ficar carnavalesco */
        [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(circle at 12% 0%, color-mix(in srgb, var(--ds-purple) 18%, transparent), transparent 30%),
                radial-gradient(circle at 92% 8%, color-mix(in srgb, var(--ds-cyan) 12%, transparent), transparent 28%),
                linear-gradient(180deg, var(--ds-bg), color-mix(in srgb, var(--ds-bg) 88%, #020617 12%)) !important;
        }

        /* Sidebar com identidade própria */
        [data-testid="stSidebar"] {
            background:
                linear-gradient(180deg,
                    color-mix(in srgb, var(--ds-surface) 92%, var(--ds-primary) 8%),
                    color-mix(in srgb, var(--ds-surface) 96%, #020617 4%)
                ) !important;
            border-right: 1px solid color-mix(in srgb, var(--ds-primary) 28%, transparent) !important;
            box-shadow: 14px 0 38px color-mix(in srgb, #000 18%, transparent) !important;
        }
        [data-testid="stSidebar"] .stButton > button {
            min-height: 46px !important;
            border-radius: 14px !important;
            border: 1px solid color-mix(in srgb, var(--ds-text) 12%, transparent) !important;
            background: linear-gradient(135deg,
                color-mix(in srgb, var(--ds-surface) 84%, var(--ds-blue) 6%),
                color-mix(in srgb, var(--ds-surface) 88%, var(--ds-purple) 5%)
            ) !important;
            color: var(--ds-text) !important;
            font-weight: 850 !important;
            letter-spacing: -.01em !important;
            box-shadow: 0 5px 14px color-mix(in srgb, #000 10%, transparent) !important;
        }
        [data-testid="stSidebar"] .stButton > button:hover {
            transform: translateY(-1px);
            border-color: color-mix(in srgb, var(--ds-cyan) 50%, transparent) !important;
            background: linear-gradient(135deg,
                color-mix(in srgb, var(--ds-indigo) 28%, var(--ds-surface)),
                color-mix(in srgb, var(--ds-purple) 20%, var(--ds-surface))
            ) !important;
            box-shadow: var(--ds-shadow-glow) !important;
        }
        [data-testid="stSidebar"] .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, var(--ds-indigo), var(--ds-purple), var(--ds-cyan)) !important;
            color: #fff !important;
            border-color: color-mix(in srgb, var(--ds-cyan) 54%, transparent) !important;
            box-shadow: 0 16px 34px color-mix(in srgb, var(--ds-primary) 32%, transparent) !important;
        }
        [data-testid="stSidebar"] .stButton > button[kind="primary"] * { color: #fff !important; }

        .sidebar-brand, .ds-status-card {
            background:
                radial-gradient(circle at 0% 0%, color-mix(in srgb, var(--ds-cyan) 16%, transparent), transparent 42%),
                linear-gradient(135deg, color-mix(in srgb, var(--ds-surface) 82%, var(--ds-primary) 14%), color-mix(in srgb, var(--ds-surface) 90%, #020617 10%)) !important;
            border: 1px solid color-mix(in srgb, var(--ds-cyan) 30%, transparent) !important;
            box-shadow: var(--ds-shadow-xl) !important;
        }
        .nav-group-title, .ds-sidebar-group {
            color: color-mix(in srgb, var(--ds-cyan) 78%, var(--ds-text) 22%) !important;
            letter-spacing: .13em !important;
            margin-top: 18px !important;
        }

        /* Filtros com superfície distinta do fundo */
        .filter-panel, .filter-panel-mini,
        [data-testid="stSidebar"] [data-testid="stExpander"] {
            background:
                linear-gradient(135deg,
                    color-mix(in srgb, var(--ds-surface) 78%, var(--ds-cyan) 7%),
                    color-mix(in srgb, var(--ds-surface) 86%, var(--ds-primary) 8%)
                ) !important;
            border: 1px solid color-mix(in srgb, var(--ds-cyan) 22%, var(--ds-text) 8%) !important;
            border-radius: 18px !important;
            box-shadow: 0 12px 28px color-mix(in srgb, #000 15%, transparent) !important;
        }
        .active-filter-chip, .ds-filter-chip, .global-search-chip, .badge, .status-badge {
            background: linear-gradient(135deg,
                color-mix(in srgb, var(--ds-cyan) 18%, var(--ds-surface)),
                color-mix(in srgb, var(--ds-primary) 18%, var(--ds-surface))
            ) !important;
            border: 1px solid color-mix(in srgb, var(--ds-cyan) 44%, transparent) !important;
            color: var(--ds-text) !important;
            box-shadow: 0 6px 16px color-mix(in srgb, var(--ds-cyan) 12%, transparent) !important;
        }

        /* Inputs e selects com realce claro */
        input, textarea, [data-baseweb="select"] > div, [data-baseweb="input"] > div {
            background: color-mix(in srgb, var(--ds-surface) 82%, var(--ds-bg) 18%) !important;
            border: 1px solid color-mix(in srgb, var(--ds-text) 18%, transparent) !important;
            color: var(--ds-text) !important;
            border-radius: 12px !important;
        }
        input:focus, textarea:focus, [data-baseweb="select"] > div:focus-within, [data-baseweb="input"] > div:focus-within {
            border-color: color-mix(in srgb, var(--ds-cyan) 62%, transparent) !important;
            box-shadow: 0 0 0 3px color-mix(in srgb, var(--ds-cyan) 16%, transparent) !important;
        }

        /* Cards mais identificáveis */
        [data-testid="stMetric"], .metric-card, .ds-metric-card, .premium-card, .compact-card,
        .alert-card, .insight-card, .answer-card, .ai-action-card, .ds-card, .ds-proactive-card {
            background:
                linear-gradient(145deg,
                    color-mix(in srgb, var(--ds-surface) 82%, var(--ds-primary) 8%),
                    color-mix(in srgb, var(--ds-surface) 90%, var(--ds-cyan) 5%)
                ) !important;
            border: 1px solid color-mix(in srgb, var(--ds-text) 13%, transparent) !important;
            box-shadow: var(--ds-shadow-xl) !important;
        }
        [data-testid="stMetric"], .metric-card, .ds-metric-card {
            border-left-width: 6px !important;
            border-left-color: var(--ds-primary) !important;
        }
        [data-testid="stHorizontalBlock"] > div:nth-child(1) [data-testid="stMetric"],
        [data-testid="stHorizontalBlock"] > div:nth-child(1) .metric-card,
        [data-testid="stHorizontalBlock"] > div:nth-child(1) .ds-metric-card { border-left-color: var(--ds-cyan) !important; }
        [data-testid="stHorizontalBlock"] > div:nth-child(2) [data-testid="stMetric"],
        [data-testid="stHorizontalBlock"] > div:nth-child(2) .metric-card,
        [data-testid="stHorizontalBlock"] > div:nth-child(2) .ds-metric-card { border-left-color: var(--ds-emerald) !important; }
        [data-testid="stHorizontalBlock"] > div:nth-child(3) [data-testid="stMetric"],
        [data-testid="stHorizontalBlock"] > div:nth-child(3) .metric-card,
        [data-testid="stHorizontalBlock"] > div:nth-child(3) .ds-metric-card { border-left-color: var(--ds-amber) !important; }
        [data-testid="stHorizontalBlock"] > div:nth-child(4) [data-testid="stMetric"],
        [data-testid="stHorizontalBlock"] > div:nth-child(4) .metric-card,
        [data-testid="stHorizontalBlock"] > div:nth-child(4) .ds-metric-card { border-left-color: var(--ds-coral) !important; }

        /* Botões principais mais claros e consistentes */
        .stButton > button, .stDownloadButton > button, button[kind="secondary"], button[kind="primary"] {
            border-radius: 999px !important;
            font-weight: 850 !important;
            border: 1px solid color-mix(in srgb, var(--ds-primary) 35%, transparent) !important;
        }
        button[kind="primary"], .stButton > button[kind="primary"] {
            background: linear-gradient(135deg, var(--ds-indigo), var(--ds-purple)) !important;
            color: #fff !important;
            box-shadow: 0 12px 28px color-mix(in srgb, var(--ds-primary) 25%, transparent) !important;
        }
        button[kind="primary"] *, .stButton > button[kind="primary"] * { color: #fff !important; }

        /* Gráficos e tabelas com container mais definido */
        div[data-testid="stPlotlyChart"], div[data-testid="stDataFrame"] {
            background: color-mix(in srgb, var(--ds-surface) 88%, var(--ds-bg) 12%) !important;
            border: 1px solid color-mix(in srgb, var(--ds-cyan) 18%, var(--ds-border)) !important;
            box-shadow: 0 14px 32px color-mix(in srgb, #000 18%, transparent) !important;
            border-radius: 18px !important;
        }
        div[data-testid="stDataFrame"] [role="columnheader"] {
            background: color-mix(in srgb, var(--ds-primary) 20%, var(--ds-surface)) !important;
            color: var(--ds-text) !important;
        }

        /* Abas em formato de botão/pílula */
        .stTabs [data-baseweb="tab-list"] {
            gap: 10px !important;
            border-bottom: 0 !important;
            background: color-mix(in srgb, var(--ds-surface) 78%, var(--ds-bg) 22%) !important;
            border: 1px solid color-mix(in srgb, var(--ds-text) 10%, transparent) !important;
            border-radius: 999px !important;
            padding: 6px !important;
            box-shadow: 0 8px 20px color-mix(in srgb, #000 10%, transparent) !important;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 999px !important;
            padding: 10px 16px !important;
            color: var(--ds-muted) !important;
            font-weight: 850 !important;
        }
        .stTabs [aria-selected="true"] {
            background: linear-gradient(135deg, color-mix(in srgb, var(--ds-indigo) 72%, transparent), color-mix(in srgb, var(--ds-purple) 70%, transparent)) !important;
            color: #fff !important;
            box-shadow: 0 10px 22px color-mix(in srgb, var(--ds-primary) 25%, transparent) !important;
        }
        .stTabs [aria-selected="true"] * { color: #fff !important; }


        /* Botão flutuante real da IA: link fixo independente dos componentes do Streamlit. */
        .ai-fab-link {
            position: fixed !important;
            right: 24px !important;
            bottom: 24px !important;
            z-index: 9999999 !important;
            width: 58px !important;
            height: 58px !important;
            min-width: 58px !important;
            min-height: 58px !important;
            border-radius: 999px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            text-decoration: none !important;
            font-size: 29px !important;
            line-height: 1 !important;
            background: linear-gradient(135deg, #06B6D4, #7C3AED, #F97316) !important;
            color: #fff !important;
            border: 1px solid rgba(255,255,255,.38) !important;
            box-shadow:
                0 16px 42px rgba(124,58,237,.40),
                0 0 0 6px rgba(6,182,212,.10) !important;
            cursor: pointer !important;
            transition: transform .18s ease, box-shadow .18s ease, filter .18s ease !important;
        }
        .ai-fab-link:hover {
            transform: translateY(-2px) scale(1.07) !important;
            filter: saturate(1.16) !important;
            box-shadow:
                0 22px 56px rgba(124,58,237,.50),
                0 0 0 8px rgba(6,182,212,.14) !important;
        }
        @media (max-width: 768px) {
            .ai-fab-link {
                right: 16px !important;
                bottom: 16px !important;
                width: 54px !important;
                height: 54px !important;
                min-width: 54px !important;
                min-height: 54px !important;
                font-size: 27px !important;
            }
        }
        /* Assistente IA flutuante: FAB real, fixo e pequeno no canto inferior direito.
           A regra abaixo mira qualquer popover usado no app. Como o dashboard usa apenas
           o popover do assistente, isso evita a barra comprida no rodapé. */
        div[data-testid="stPopover"] {
            position: fixed !important;
            right: 24px !important;
            bottom: 24px !important;
            z-index: 999999 !important;
            width: 60px !important;
            min-width: 60px !important;
            max-width: 60px !important;
            height: 60px !important;
            min-height: 60px !important;
            max-height: 60px !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        div[data-testid="stPopover"] > button,
        div[data-testid="stPopover"] button {
            width: 60px !important;
            min-width: 60px !important;
            max-width: 60px !important;
            height: 60px !important;
            min-height: 60px !important;
            max-height: 60px !important;

            padding: 0 !important;
            margin: 0 !important;
            border-radius: 999px !important;
            overflow: hidden !important;

            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;

            font-size: 0 !important;
            line-height: 1 !important;

            background: linear-gradient(135deg, #06B6D4, #7C3AED, #F97316) !important;
            color: #fff !important;
            border: 1px solid rgba(255,255,255,.35) !important;
            box-shadow:
                0 16px 42px rgba(124,58,237,.38),
                0 0 0 6px rgba(6,182,212,.08) !important;
            cursor: pointer !important;
            transition: transform .18s ease, box-shadow .18s ease, filter .18s ease !important;
        }
        div[data-testid="stPopover"] > button::before,
        div[data-testid="stPopover"] button::before {
            content: "🤖";
            font-size: 29px !important;
            line-height: 1 !important;
            display: block !important;
            filter: drop-shadow(0 2px 6px rgba(0,0,0,.28));
        }
        div[data-testid="stPopover"] > button:hover,
        div[data-testid="stPopover"] button:hover {
            transform: translateY(-2px) scale(1.06) !important;
            box-shadow:
                0 22px 56px rgba(124,58,237,.48),
                0 0 0 8px rgba(6,182,212,.12) !important;
            filter: saturate(1.15) !important;
        }
        div[data-testid="stPopover"] button p,
        div[data-testid="stPopover"] button span,
        div[data-testid="stPopover"] button div {
            font-size: 0 !important;
            line-height: 0 !important;
        }
        div[data-testid="stPopover"] button svg {
            display: none !important;
        }
        @media (max-width: 768px) {
            div[data-testid="stPopover"] {
                right: 16px !important;
                bottom: 16px !important;
                width: 54px !important;
                min-width: 54px !important;
                max-width: 54px !important;
                height: 54px !important;
                min-height: 54px !important;
                max-height: 54px !important;
            }
            div[data-testid="stPopover"] > button,
            div[data-testid="stPopover"] button {
                width: 54px !important;
                min-width: 54px !important;
                max-width: 54px !important;
                height: 54px !important;
                min-height: 54px !important;
                max-height: 54px !important;
            }
        }
        .floating-ai-card {
            background:
                linear-gradient(color-mix(in srgb, var(--ds-bg) 74%, var(--ds-surface) 26%), color-mix(in srgb, var(--ds-bg) 78%, var(--ds-primary) 8%)) padding-box,
                linear-gradient(135deg, var(--ds-cyan), var(--ds-purple)) border-box !important;
            border: 1px solid transparent !important;
            box-shadow: 0 22px 60px color-mix(in srgb, #000 34%, transparent) !important;
        }
        

        /* FAB IA definitivo: botão nativo do Streamlit, pequeno e fixo à direita. */
        .st-key-ai_fab_button {
            position: fixed !important;
            right: 26px !important;
            bottom: 26px !important;
            z-index: 2147483647 !important;
            width: 64px !important;
            height: 64px !important;
            min-width: 64px !important;
            min-height: 64px !important;
            max-width: 64px !important;
            max-height: 64px !important;
            margin: 0 !important;
            padding: 0 !important;
            display: flex !important;
            align-items: center !important;
            justify-content: center !important;
            pointer-events: auto !important;
        }
        .st-key-ai_fab_button button {
            width: 64px !important;
            height: 64px !important;
            min-width: 64px !important;
            min-height: 64px !important;
            max-width: 64px !important;
            max-height: 64px !important;
            padding: 0 !important;
            margin: 0 !important;
            border-radius: 999px !important;
            display: inline-flex !important;
            align-items: center !important;
            justify-content: center !important;
            background: linear-gradient(135deg, #06B6D4, #7C3AED, #F97316) !important;
            border: 1px solid rgba(255,255,255,.46) !important;
            color: #FFFFFF !important;
            box-shadow: 0 18px 48px rgba(124,58,237,.48), 0 0 0 7px rgba(6,182,212,.12) !important;
            cursor: pointer !important;
            transition: transform .18s ease, box-shadow .18s ease, filter .18s ease !important;
        }
        .st-key-ai_fab_button button:hover {
            transform: translateY(-2px) scale(1.07) !important;
            filter: saturate(1.18) !important;
            box-shadow: 0 22px 56px rgba(124,58,237,.58), 0 0 0 8px rgba(6,182,212,.16) !important;
        }
        .st-key-ai_fab_button button p,
        .st-key-ai_fab_button button span,
        .st-key-ai_fab_button button div {
            font-size: 30px !important;
            line-height: 1 !important;
            color: #FFFFFF !important;
            margin: 0 !important;
            padding: 0 !important;
        }
        @media (max-width: 768px) {
            .st-key-ai_fab_button {
                right: 16px !important;
                bottom: 16px !important;
                width: 56px !important;
                height: 56px !important;
                min-width: 56px !important;
                min-height: 56px !important;
                max-width: 56px !important;
                max-height: 56px !important;
            }
            .st-key-ai_fab_button button {
                width: 56px !important;
                height: 56px !important;
                min-width: 56px !important;
                min-height: 56px !important;
                max-width: 56px !important;
                max-height: 56px !important;
            }
            .st-key-ai_fab_button button p,
            .st-key-ai_fab_button button span,
            .st-key-ai_fab_button button div {
                font-size: 27px !important;
            }
        }
</style>
        """,
        unsafe_allow_html=True,
    )
