from __future__ import annotations

import pandas as pd


def currency(value: float) -> str:
    return f"{value:,.2f} MAD".replace(",", " ")


def compute_kpis(df: pd.DataFrame, monthly_forecast: pd.DataFrame) -> dict[str, str]:
    if df.empty:
        return {
            "Total à payer": currency(0),
            "Nombre de factures": "0",
            "Nombre de fournisseurs": "0",
            "Mois le plus élevé": "N/A",
        }

    biggest_month = "N/A"
    if not monthly_forecast.empty:
        row = monthly_forecast.loc[monthly_forecast["Montant"].idxmax()]
        biggest_month = f"{row['Mois']}\n{currency(row['Montant'])}"

    return {
        "Total à payer": currency(float(df["Montant"].sum())),
        "Nombre de factures": f"{len(df):,}".replace(",", " "),
        "Nombre de fournisseurs": f"{df['Nom du fournisseur'].nunique():,}".replace(",", " "),
        "Mois le plus élevé": biggest_month,
    }


def top_suppliers(df: pd.DataFrame, limit: int = 10) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Nom du fournisseur", "Montant", "Nombre de factures"])
    return (
        df.groupby("Nom du fournisseur", as_index=False)
        .agg(Montant=("Montant", "sum"), **{"Nombre de factures": ("Montant", "size")})
        .sort_values("Montant", ascending=False)
        .head(limit)
    )


def amount_by_project(df: pd.DataFrame, limit: int = 15) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Projet / Code Affaire", "Montant"])
    return (
        df.groupby("Projet / Code Affaire", as_index=False)["Montant"]
        .sum()
        .sort_values("Montant", ascending=False)
        .head(limit)
    )


def severity_from_percentage(value: float, medium: float, high: float) -> str:
    if value >= high:
        return "Élevé"
    if value >= medium:
        return "Moyen"
    return "Faible"


def compute_executive_metrics(df: pd.DataFrame, monthly_forecast: pd.DataFrame) -> dict[str, object]:
    total = float(df["Montant"].sum()) if not df.empty and "Montant" in df.columns else 0.0

    top5_share = 0.0
    if total > 0 and not df.empty:
        supplier_amounts = df.groupby("Nom du fournisseur")["Montant"].sum().sort_values(ascending=False)
        top5_share = float(supplier_amounts.head(5).sum() / total)

    average_invoice = float(df["Montant"].mean()) if not df.empty and "Montant" in df.columns else 0.0

    highest_month = {"label": "N/A", "amount": 0.0, "share": 0.0}
    if not monthly_forecast.empty:
        row = monthly_forecast.loc[monthly_forecast["Montant"].idxmax()]
        highest_month = {
            "label": row["Mois"],
            "amount": float(row["Montant"]),
            "share": float(row["Montant"] / total) if total > 0 else 0.0,
        }

    most_expensive_project = {"label": "N/A", "amount": 0.0, "share": 0.0}
    if total > 0 and not df.empty:
        project_amounts = df.groupby("Projet / Code Affaire")["Montant"].sum().sort_values(ascending=False)
        if not project_amounts.empty:
            most_expensive_project = {
                "label": project_amounts.index[0],
                "amount": float(project_amounts.iloc[0]),
                "share": float(project_amounts.iloc[0] / total),
            }

    return {
        "highest_month": highest_month,
        "top5_supplier_share": top5_share,
        "top5_supplier_severity": severity_from_percentage(top5_share, 0.45, 0.65),
        "most_expensive_project": most_expensive_project,
        "average_invoice_amount": average_invoice,
        "large_month_severity": severity_from_percentage(highest_month["share"], 0.30, 0.45),
    }


def supplier_pareto(df: pd.DataFrame, limit: int = 15) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Nom du fournisseur", "Montant", "Contribution", "Contribution cumulée"])

    total = float(df["Montant"].sum())
    result = (
        df.groupby("Nom du fournisseur", as_index=False)["Montant"]
        .sum()
        .sort_values("Montant", ascending=False)
        .head(limit)
    )
    result["Contribution"] = result["Montant"] / total if total else 0
    result["Contribution cumulée"] = result["Contribution"].cumsum()
    return result
