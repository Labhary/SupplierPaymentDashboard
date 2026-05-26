from __future__ import annotations

from io import BytesIO

import pandas as pd
import streamlit as st

from src.metrics import currency
from src.transformations import parse_excel_date


def configure_page() -> None:
    st.set_page_config(
        page_title="Dashboard de prévision des décaissements fournisseurs",
        page_icon=":bar_chart:",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    st.markdown(
        """
        <style>
        .stApp {background: #0f172a; color: #e5e7eb;}
        .block-container {padding-top: 1.2rem; padding-bottom: 2.2rem;}
        [data-testid="stMetric"] {
            background: #172033;
            border: 1px solid #273449;
            border-radius: 8px;
            padding: 16px 18px;
            box-shadow: 0 1px 8px rgba(0, 0, 0, 0.22);
        }
        [data-testid="stMetricLabel"] {color: #aebbd0;}
        [data-testid="stMetricValue"] {
            font-size: 1.35rem;
            color: #f8fafc;
            line-height: 1.2;
            white-space: pre-line;
            overflow-wrap: anywhere;
        }
        [data-testid="stSidebar"] {background: #111827;}
        [data-testid="stSidebar"] .block-container {padding-top: 1.1rem;}
        div[data-testid="stDataFrame"] {border: 1px solid #273449; border-radius: 8px;}
        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: #273449;
            border-radius: 8px;
        }
        .section-title {
            font-size: 1.15rem;
            font-weight: 750;
            margin: 1.1rem 0 0.15rem;
            color: #f8fafc;
        }
        .section-caption {
            color: #aebbd0;
            font-size: 0.9rem;
            margin-bottom: 0.55rem;
        }
        .exec-card {
            border: 1px solid #273449;
            border-radius: 8px;
            padding: 14px 15px;
            background: #172033;
            min-height: 118px;
            box-shadow: 0 1px 8px rgba(0, 0, 0, 0.2);
        }
        .exec-label {font-size: 0.78rem; color: #aebbd0; font-weight: 700; text-transform: uppercase;}
        .exec-value {font-size: 1.15rem; font-weight: 760; color: #f8fafc; margin-top: 0.3rem;}
        .exec-caption {font-size: 0.82rem; color: #aebbd0; margin-top: 0.25rem;}
        .badge {
            display: inline-block;
            border-radius: 999px;
            padding: 3px 9px;
            font-size: 0.75rem;
            font-weight: 700;
            margin-top: 0.45rem;
        }
        .badge-low {background: #dcfce7; color: #166534;}
        .badge-medium {background: #fef3c7; color: #92400e;}
        .badge-high {background: #fee2e2; color: #991b1b;}
        .insight-box {
            border-left: 4px solid #2f6fed;
            background: #172033;
            padding: 10px 13px;
            margin-bottom: 8px;
            border-radius: 6px;
            color: #e5e7eb;
        }
        .footer {
            border-top: 1px solid #273449;
            margin-top: 2rem;
            padding-top: 1rem;
            color: #aebbd0;
            font-size: 0.86rem;
            text-align: center;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_header() -> None:
    st.title("Dashboard de prévision des décaissements fournisseurs")
    st.caption(
        "Visualisation locale des paiements fournisseurs prévisionnels à partir d'un fichier Excel."
    )


def render_scope_box() -> None:
    st.info(
        "Périmètre actuel: ce tableau de bord couvre uniquement les factures fournisseurs "
        "et les décaissements prévisionnels associés. Il ne constitue pas encore une "
        "prévision complète de trésorerie, car il n'intègre pas les soldes bancaires, "
        "les encaissements clients ni les scénarios de paiement."
    )


def render_upload_box():
    return st.file_uploader(
        "Importer le fichier Excel",
        type=["xlsx", "xlsm", "xls"],
        help="Le fichier attendu contient Feuil1 pour les factures et Feuil2 pour la synthèse mensuelle.",
    )


def render_sidebar_filters(df: pd.DataFrame) -> dict[str, object]:
    st.sidebar.header("Filtres")
    st.sidebar.caption("Affinez le périmètre d'analyse")

    if df.empty:
        return {
            "suppliers": [],
            "codes": [],
            "project_names": [],
            "months": [],
            "min_amount": 0.0,
            "active_count": 0,
        }

    def reset_filters() -> None:
        st.session_state["filter_suppliers"] = []
        st.session_state["filter_codes"] = []
        st.session_state["filter_project_names"] = []
        st.session_state["filter_months"] = []
        st.session_state["filter_min_amount"] = 0.0

    st.sidebar.button("Réinitialiser les filtres", on_click=reset_filters, use_container_width=True)
    st.sidebar.write("")
    st.sidebar.divider()

    suppliers = st.sidebar.multiselect(
        "Fournisseur",
        sorted(df["Nom du fournisseur"].dropna().unique().tolist()),
        key="filter_suppliers",
    )
    codes = st.sidebar.multiselect(
        "Code Affaire",
        sorted(df["Code Affaire"].dropna().unique().tolist()),
        key="filter_codes",
    )
    project_names = st.sidebar.multiselect(
        "Nom Projet",
        sorted(df["Nom Projet"].dropna().unique().tolist()),
        key="filter_project_names",
    )
    months = st.sidebar.multiselect(
        "Mois calculé",
        sorted([month for month in df["Mois"].dropna().unique().tolist() if month != "Date invalide"]),
        key="filter_months",
    )
    st.sidebar.write("")
    max_amount = float(df["Montant"].max()) if not df.empty else 0.0
    if st.session_state.get("filter_min_amount", 0.0) > max_amount:
        st.session_state["filter_min_amount"] = 0.0
    min_amount = st.sidebar.slider(
        "Montant minimum",
        min_value=0.0,
        max_value=max_amount,
        value=0.0,
        step=max(max_amount / 100, 1.0) if max_amount else 1.0,
        key="filter_min_amount",
    )
    active_count = sum(
        [
            bool(suppliers),
            bool(codes),
            bool(project_names),
            bool(months),
            min_amount > 0,
        ]
    )
    st.sidebar.info("Aucun filtre actif" if active_count == 0 else f"{active_count} filtre(s) actif(s)")

    return {
        "suppliers": suppliers,
        "codes": codes,
        "project_names": project_names,
        "months": months,
        "min_amount": min_amount,
        "active_count": active_count,
    }


def render_kpis(kpis: dict[str, str]) -> None:
    captions = {
        "Total à payer": "Somme des factures après filtres.",
        "Nombre de factures": "Volume de lignes retenues.",
        "Nombre de fournisseurs": "Fournisseurs distincts.",
        "Mois le plus élevé": "Pic de décaissement mensuel.",
    }
    columns = st.columns(5)
    for column, (label, value) in zip(columns, kpis.items()):
        display_value = str(value).replace("\n", "<br>")
        column.markdown(
            f"""
            <div class="exec-card">
                <div class="exec-label">{label}</div>
                <div class="exec-value">{display_value}</div>
                <div class="exec-caption">{captions.get(label, "")}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )


def _badge_class(severity: str) -> str:
    return {
        "Faible": "badge-low",
        "Moyen": "badge-medium",
        "Élevé": "badge-high",
    }.get(severity, "badge-low")


def render_executive_analysis(metrics: dict[str, object]) -> None:
    highest_month = metrics["highest_month"]
    project = metrics["most_expensive_project"]
    cards = [
        (
            "Mois le plus exposé",
            f"{highest_month['label']}",
            currency(highest_month["amount"]),
            metrics["large_month_severity"],
        ),
        (
            "Concentration fournisseurs",
            f"{metrics['top5_supplier_share'] * 100:.0f}%",
            "Part des 5 premiers fournisseurs",
            metrics["top5_supplier_severity"],
        ),
        (
            "Projet / code affaire clé",
            f"{project['label']}",
            currency(project["amount"]),
            "Élevé" if project["share"] >= 0.45 else "Moyen" if project["share"] >= 0.25 else "Faible",
        ),
        (
            "Facture moyenne",
            currency(metrics["average_invoice_amount"]),
            "Ticket moyen des factures filtrées",
            "Faible",
        ),
    ]
    columns = st.columns(len(cards))
    for column, (label, value, caption, severity) in zip(columns, cards):
        column.markdown(
            f"""
            <div class="exec-card">
                <div class="exec-label">{label}</div>
                <div class="exec-value">{value}</div>
                <div class="exec-caption">{caption}</div>
                <span class="badge {_badge_class(severity)}">Risque {severity.lower()}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_insights(insights: list[str]) -> None:
    for insight in insights:
        st.markdown(f'<div class="insight-box">{insight}</div>', unsafe_allow_html=True)


def _dataframe_to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(
        buffer,
        engine="openpyxl",
        date_format="yyyy-mm-dd",
        datetime_format="yyyy-mm-dd",
    ) as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
    return buffer.getvalue()


def render_downloads(invoices: pd.DataFrame, forecast: pd.DataFrame) -> None:
    export_invoices = prepare_invoice_export(invoices)
    export_forecast = prepare_forecast_export(forecast)
    left, middle, right = st.columns(3)
    left.download_button(
        "Factures filtrées (CSV)",
        data=export_invoices.to_csv(index=False).encode("utf-8-sig"),
        file_name="factures_fournisseurs_filtrees.csv",
        mime="text/csv",
        disabled=invoices.empty,
    )
    middle.download_button(
        "Prévision mensuelle (CSV)",
        data=export_forecast.to_csv(index=False).encode("utf-8-sig"),
        file_name="prevision_decaissements_mensuelle.csv",
        mime="text/csv",
        disabled=forecast.empty,
    )
    right.download_button(
        "Export Excel",
        data=_dataframe_to_excel_bytes({"Factures filtrées": export_invoices, "Prévision mensuelle": export_forecast}),
        file_name="export_decaissements_fournisseurs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        disabled=invoices.empty and forecast.empty,
    )


def format_money_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    output = df.copy()
    for column in columns:
        if column in output.columns:
            output[column] = output[column].apply(lambda value: currency(float(value)) if pd.notna(value) else "")
    return output


def prepare_display_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    for column in output.columns:
        if pd.api.types.is_datetime64_any_dtype(output[column]):
            output[column] = output[column].dt.strftime("%d/%m/%Y")
    return output.replace({pd.NaT: "", pd.NA: "", None: ""}).fillna("")


DETAIL_TABLE_COLUMNS = [
    "Nom du fournisseur",
    "Numéro réf. fournisseur",
    "Date d'échéance",
    "Montant",
    "Date comptable",
    "Code Affaire",
    "Nom Projet",
    "Mois source",
    "Mois",
    "Année",
]


def prepare_invoice_display(df: pd.DataFrame) -> pd.DataFrame:
    columns = [column for column in DETAIL_TABLE_COLUMNS if column in df.columns]
    output = prepare_display_dataframe(df[columns])
    output = output.rename(columns={"Mois": "Mois calculé"})
    return format_money_columns(output, ["Montant"])


def prepare_invoice_export(df: pd.DataFrame) -> pd.DataFrame:
    columns = [column for column in DETAIL_TABLE_COLUMNS if column in df.columns]
    output = df[columns].copy() if columns else pd.DataFrame()
    output = output.rename(columns={"Mois": "Mois calculé"})
    if "Montant" in output.columns:
        output["Montant"] = pd.to_numeric(output["Montant"], errors="coerce")
    for column in ["Date d'échéance", "Date comptable"]:
        if column in output.columns:
            output[column] = parse_excel_date(output[column])
    if "Mois calculé" in output.columns:
        output["Mois calculé"] = parse_excel_date(output["Mois calculé"])
    return output


def prepare_forecast_export(df: pd.DataFrame) -> pd.DataFrame:
    output = df.copy()
    for column in ["Montant", "Cumul"]:
        if column in output.columns:
            output[column] = pd.to_numeric(output[column], errors="coerce")
    return output


def render_metadata(invoices: pd.DataFrame) -> None:
    valid_months = sorted([month for month in invoices.get("Mois", pd.Series(dtype=str)).dropna().unique() if month != "Date invalide"])
    period = "N/A"
    if valid_months:
        period = f"{valid_months[0]} à {valid_months[-1]}"
    detected = [
        "Nom du fournisseur",
        "Numéro réf. fournisseur",
        "Date d'échéance",
        "Montant",
        "Date comptable",
        "Code Affaire",
        "Nom Projet",
        "MOIS",
    ]
    cols = st.columns(3)
    cols[0].metric("Nombre de factures", f"{len(invoices):,}".replace(",", " "))
    cols[1].metric("Période détectée", period)
    cols[2].metric("Colonnes source détectées", f"{len(detected)}/8")


def render_footer() -> None:
    st.markdown(
        '<div class="footer">Dashboard local de prévision des décaissements fournisseurs</div>',
        unsafe_allow_html=True,
    )
