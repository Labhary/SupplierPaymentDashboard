from __future__ import annotations

import pandas as pd

from src.metrics import currency


def percent(value: float) -> str:
    return f"{value * 100:.0f}%"


def generate_executive_insights(
    df: pd.DataFrame,
    monthly_forecast: pd.DataFrame,
    executive_metrics: dict[str, object],
) -> list[str]:
    if df.empty:
        return ["Aucune donnée exploitable après application des filtres."]

    insights: list[str] = []
    highest_month = executive_metrics["highest_month"]
    if highest_month["label"] != "N/A":
        insights.append(
            f"{highest_month['label']} représente le plus fort besoin de décaissement "
            f"avec {currency(highest_month['amount'])}."
        )

    top5_share = executive_metrics["top5_supplier_share"]
    insights.append(
        f"Les 5 principaux fournisseurs représentent {percent(top5_share)} des engagements filtrés."
    )

    project = executive_metrics["most_expensive_project"]
    if project["label"] != "N/A":
        insights.append(
            f"Le projet / code affaire {project['label']} concentre la plus grande exposition financière "
            f"avec {currency(project['amount'])}."
        )

    average_invoice = executive_metrics["average_invoice_amount"]
    insights.append(f"Le montant moyen par facture est de {currency(average_invoice)}.")

    return insights
