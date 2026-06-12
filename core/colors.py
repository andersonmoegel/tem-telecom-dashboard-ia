"""Paleta centralizada do sistema.

Use este arquivo para manter gráficos, cards, alertas e IA com a mesma identidade visual.
"""

PALETTE = {
    "background": "#0B1020",
    "surface": "#151B2E",
    "surface_elevated": "#1D2640",
    "text": "#F8FAFC",
    "muted": "#94A3B8",
    "primary": "#7C3AED",
    "indigo": "#6366F1",
    "cyan": "#38BDF8",
    "success": "#10B981",
    "warning": "#F59E0B",
    "danger": "#EF4444",
    "neutral_ai": "#8B5CF6",
}

CHART_COLOSP = [
    PALETTE["primary"],
    PALETTE["cyan"],
    PALETTE["success"],
    PALETTE["warning"],
    PALETTE["danger"],
    PALETTE["indigo"],
    "#14B8A6",
    "#A855F7",
]

STATUS_COLOSP = {
    "sucesso": PALETTE["success"],
    "ativo": PALETTE["success"],
    "alerta": PALETTE["warning"],
    "medio": PALETTE["warning"],
    "critico": PALETTE["danger"],
    "erro": PALETTE["danger"],
    "ia": PALETTE["neutral_ai"],
    "neutro": PALETTE["cyan"],
}
