"""Padronização visual para gráficos Plotly."""

from __future__ import annotations

from .colors import CHART_COLOSP, PALETTE


def configurar_tema_grafico(fig, height: int | None = None):
    """Aplica tema visual comum a qualquer figura Plotly."""
    fig.update_layout(
        colorway=CHART_COLOSP,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color=PALETTE["text"], family="Inter, Segoe UI, Arial, sans-serif"),
        title=dict(font=dict(color=PALETTE["text"], size=18), x=0.01, xanchor="left"),
        margin=dict(l=28, r=28, t=72, b=44),
        hovermode="closest",
        legend=dict(
            title=dict(text="Legenda", font=dict(color=PALETTE["muted"], size=12)),
            bgcolor="rgba(15,23,42,.04)",
            bordercolor="rgba(148,163,184,.22)",
            borderwidth=1,
            font=dict(color=PALETTE["text"], size=12),
            orientation="h",
            yanchor="bottom",
            y=1.03,
            xanchor="right",
            x=1,
            itemclick="toggleothers",
            itemdoubleclick="toggle",
        ),
        uniformtext_minsize=10,
        uniformtext_mode="hide",
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="rgba(148,163,184,.16)",
        zerolinecolor="rgba(148,163,184,.24)",
        tickfont=dict(color=PALETTE["text"], size=11),
        title_font=dict(color=PALETTE["muted"], size=12),
    )
    fig.update_yaxes(
        showgrid=True,
        gridcolor="rgba(148,163,184,.16)",
        zerolinecolor="rgba(148,163,184,.24)",
        tickfont=dict(color=PALETTE["text"], size=11),
        title_font=dict(color=PALETTE["muted"], size=12),
    )
    try:
        fig.update_traces(marker_line_color="rgba(255,255,255,.16)", marker_line_width=.7, selector=dict(type="bar"))
    except Exception:
        pass
    try:
        fig.update_traces(textfont_color=PALETTE["text"], selector=dict(type="treemap"))
    except Exception:
        pass
    if height:
        fig.update_layout(height=height)
    return fig


def area_chart_defaults(fig):
    """Aplica padrão para gráficos de evolução por período."""
    try:
        fig.update_traces(line_shape="spline")
    except Exception:
        pass
    return fig
