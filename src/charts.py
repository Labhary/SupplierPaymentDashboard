from __future__ import annotations

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go


TEMPLATE = "plotly_dark"
PRIMARY = "#2f6fed"
ACCENT = "#20a39e"


def empty_chart(message: str) -> go.Figure:
    fig = go.Figure()
    fig.add_annotation(
        text=message,
        xref="paper",
        yref="paper",
        x=0.5,
        y=0.5,
        showarrow=False,
        font={"size": 15, "color": "#aebbd0"},
    )
    fig.update_layout(
        template=TEMPLATE,
        height=320,
        margin=dict(l=20, r=20, t=35, b=20),
        paper_bgcolor="#172033",
        plot_bgcolor="#172033",
    )
    return fig


def _clean_layout(fig: go.Figure, height: int = 380) -> go.Figure:
    fig.update_layout(
        template=TEMPLATE,
        height=height,
        margin=dict(l=20, r=20, t=58, b=25),
        title_font=dict(size=17, color="#f8fafc"),
        font=dict(color="#dbe4f0"),
        paper_bgcolor="#172033",
        plot_bgcolor="#172033",
        legend_title_text="",
    )
    return fig


def monthly_bar(forecast: pd.DataFrame) -> go.Figure:
    if forecast.empty:
        return empty_chart("Aucune échéance valide à afficher.")
    fig = px.bar(
        forecast,
        x="Mois",
        y="Montant",
        text_auto=".2s",
        title="Décaissements prévisionnels par mois",
        color_discrete_sequence=[PRIMARY],
        template=TEMPLATE,
    )
    fig.update_layout(yaxis_title="Montant (MAD)", xaxis_title="")
    return _clean_layout(fig)


def cumulative_line(forecast: pd.DataFrame) -> go.Figure:
    if forecast.empty:
        return empty_chart("Le cumul ne peut pas être calculé sans dates d'échéance.")
    fig = px.line(
        forecast,
        x="Mois",
        y="Cumul",
        markers=True,
        title="Cumul des décaissements prévisionnels",
        color_discrete_sequence=[ACCENT],
        template=TEMPLATE,
    )
    fig.update_layout(yaxis_title="Cumul (MAD)", xaxis_title="")
    return _clean_layout(fig)


def top_suppliers_bar(top_df: pd.DataFrame) -> go.Figure:
    if top_df.empty:
        return empty_chart("Aucun fournisseur disponible.")
    fig = px.bar(
        top_df.sort_values("Montant"),
        x="Montant",
        y="Nom du fournisseur",
        orientation="h",
        title="Top 10 fournisseurs par montant",
        color_discrete_sequence=[PRIMARY],
        template=TEMPLATE,
    )
    fig.update_layout(xaxis_title="Montant (MAD)", yaxis_title="")
    return _clean_layout(fig, height=420)


def project_bar(project_df: pd.DataFrame) -> go.Figure:
    if project_df.empty:
        return empty_chart("Aucun projet ou code affaire disponible.")
    fig = px.bar(
        project_df.sort_values("Montant"),
        x="Montant",
        y="Projet / Code Affaire",
        orientation="h",
        title="Montant par projet / code affaire",
        color_discrete_sequence=[ACCENT],
        template=TEMPLATE,
    )
    fig.update_layout(xaxis_title="Montant (MAD)", yaxis_title="")
    return _clean_layout(fig, height=420)


def suppliers_treemap(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return empty_chart("Aucun fournisseur disponible pour la treemap.")
    supplier_amounts = df.groupby("Nom du fournisseur", as_index=False)["Montant"].sum()
    fig = px.treemap(
        supplier_amounts,
        path=["Nom du fournisseur"],
        values="Montant",
        title="Treemap des fournisseurs par montant",
        color="Montant",
        color_continuous_scale="Blues",
        template=TEMPLATE,
    )
    return _clean_layout(fig, height=420)


def supplier_pareto_chart(pareto: pd.DataFrame) -> go.Figure:
    if pareto.empty:
        return empty_chart("Aucune contribution fournisseur disponible.")
    plot_df = pareto.sort_values("Montant")
    fig = go.Figure()
    fig.add_bar(
        x=plot_df["Montant"],
        y=plot_df["Nom du fournisseur"],
        orientation="h",
        marker_color=PRIMARY,
        name="Montant",
        text=(plot_df["Contribution"] * 100).round(1).astype(str) + "%",
        textposition="auto",
    )
    fig.add_scatter(
        x=plot_df["Contribution cumulée"] * plot_df["Montant"].max(),
        y=plot_df["Nom du fournisseur"],
        mode="lines+markers",
        name="Contribution cumulée",
        line=dict(color=ACCENT, width=2),
        hovertemplate="Contribution cumulée: %{customdata:.1%}<extra></extra>",
        customdata=plot_df["Contribution cumulée"],
    )
    fig.update_layout(
        title="Contribution fournisseur type Pareto",
        xaxis_title="Montant (MAD)",
        yaxis_title="",
    )
    return _clean_layout(fig, height=430)
