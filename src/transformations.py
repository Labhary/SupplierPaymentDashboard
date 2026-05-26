from __future__ import annotations

import re
import unicodedata
from datetime import date, datetime

import pandas as pd


INVOICE_COLUMNS = [
    "Nom du fournisseur",
    "Numéro réf. fournisseur",
    "Date d'échéance",
    "Montant",
    "Date comptable",
    "Code Affaire",
    "Nom Projet",
    "MOIS",
]


def normalize_column_label(value: object) -> str:
    """Normalize Excel labels for tolerant matching."""
    if pd.isna(value):
        return ""
    text = str(value).strip().lower().replace("\n", " ")
    text = "".join(
        char
        for char in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(char)
    )
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


CANONICAL_INVOICE_COLUMNS = {
    normalize_column_label("Nom du fournisseur"): "Nom du fournisseur",
    normalize_column_label("Fournisseur"): "Nom du fournisseur",
    normalize_column_label("Numéro réf. fournisseur"): "Numéro réf. fournisseur",
    normalize_column_label("Numero ref fournisseur"): "Numéro réf. fournisseur",
    normalize_column_label("Numéro réf. Fournisseur"): "Numéro réf. fournisseur",
    normalize_column_label("Reference fournisseur"): "Numéro réf. fournisseur",
    normalize_column_label("Date d'échéance"): "Date d'échéance",
    normalize_column_label("Date echeance"): "Date d'échéance",
    normalize_column_label("Echeance"): "Date d'échéance",
    normalize_column_label("Montant"): "Montant",
    normalize_column_label("Date comptable"): "Date comptable",
    normalize_column_label("Code Affaire"): "Code Affaire",
    normalize_column_label("Nom Projet"): "Nom Projet",
    normalize_column_label("Projet"): "Nom Projet",
    normalize_column_label("MOIS"): "MOIS",
    normalize_column_label("Mois source"): "MOIS",
}


def canonicalize_invoice_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Map common Excel column variants to internal invoice column names."""
    output = pd.DataFrame(index=df.index)
    used_original_columns: set[object] = set()

    for expected in INVOICE_COLUMNS:
        matches = [
            column
            for column in df.columns
            if CANONICAL_INVOICE_COLUMNS.get(normalize_column_label(column)) == expected
        ]
        if not matches:
            continue

        values = df[matches].bfill(axis=1).iloc[:, 0] if len(matches) > 1 else df[matches[0]]
        output[expected] = values
        used_original_columns.update(matches)

    for column in df.columns:
        if column in used_original_columns:
            continue
        candidate = str(column).strip()
        if candidate and candidate not in output.columns:
            output[candidate] = df[column]

    return output


def _ensure_columns(df: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    output = df.copy()
    for column in columns:
        if column not in output.columns:
            output[column] = pd.NA
    return output


def _format_source_month(value: object) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    numeric = pd.to_numeric(text, errors="coerce")
    if pd.notna(numeric) and float(numeric).is_integer() and 1 <= int(numeric) <= 12:
        return f"{int(numeric):02d}"
    return text


def parse_amount(series: pd.Series) -> pd.Series:
    """Convert mixed French/Excel amount values to numeric values."""
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce").fillna(0)

    normalized = (
        series.astype(str)
        .str.replace("\u00a0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace("MAD", "", case=False, regex=False)
        .str.replace("DH", "", case=False, regex=False)
        .str.replace(",", ".", regex=False)
        .str.replace(r"[^0-9.\-]", "", regex=True)
    )
    return pd.to_numeric(normalized, errors="coerce").fillna(0)


def parse_excel_date(series: pd.Series) -> pd.Series:
    """Parse Excel datetime cells, pandas timestamps, serials, and mixed strings."""

    def parse_one(value: object) -> pd.Timestamp:
        if pd.isna(value):
            return pd.NaT
        if isinstance(value, pd.Timestamp):
            return value.normalize()
        if isinstance(value, datetime):
            return pd.Timestamp(value).normalize()
        if isinstance(value, date):
            return pd.Timestamp(value)
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            # Excel serial dates are typically in this range for operational data.
            if 1 <= float(value) <= 80000:
                return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
            return pd.NaT

        text = str(value).strip()
        if not text or text.lower() in {"nat", "nan", "none", "null"}:
            return pd.NaT

        numeric = pd.to_numeric(text, errors="coerce")
        if pd.notna(numeric) and 1 <= float(numeric) <= 80000:
            return pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")

        if re.match(r"^\d{4}[-/]\d{1,2}[-/]\d{1,2}", text):
            parsed = pd.to_datetime(text, dayfirst=False, errors="coerce")
            if pd.notna(parsed):
                return parsed.normalize()

        normalized = text.replace(".", "/").replace("-", "/")
        parsed = pd.to_datetime(normalized, dayfirst=True, errors="coerce")
        if pd.notna(parsed):
            return parsed.normalize()
        return pd.to_datetime(text, errors="coerce")

    return pd.Series([parse_one(value) for value in series], index=series.index, dtype="datetime64[ns]")


def clean_invoices(df: pd.DataFrame) -> pd.DataFrame:
    output = canonicalize_invoice_columns(df)
    output = _ensure_columns(output, INVOICE_COLUMNS)
    output = output[INVOICE_COLUMNS].copy()

    output["Nom du fournisseur"] = output["Nom du fournisseur"].fillna("Non renseigné").astype(str).str.strip()
    output["Numéro réf. fournisseur"] = output["Numéro réf. fournisseur"].fillna("").astype(str).str.strip()
    output["Code Affaire"] = output["Code Affaire"].fillna("Non renseigné").astype(str).str.strip()
    output["Nom Projet"] = output["Nom Projet"].fillna("Non renseigné").astype(str).str.strip()
    output["Mois source"] = output["MOIS"].apply(_format_source_month)

    output["Montant"] = parse_amount(output["Montant"])
    output["Date d'échéance"] = parse_excel_date(output["Date d'échéance"])
    output["Date comptable"] = parse_excel_date(output["Date comptable"])
    output["Date d'échéance valide"] = output["Date d'échéance"].notna()

    output = output[output["Montant"].notna()]
    output = output[output["Montant"] != 0]
    valid_due_dates = output["Date d'échéance"].notna()
    output["Mois"] = "Date invalide"
    output.loc[valid_due_dates, "Mois"] = output.loc[valid_due_dates, "Date d'échéance"].dt.to_period("M").astype(str)
    output["Année"] = "Date invalide"
    output.loc[valid_due_dates, "Année"] = output.loc[valid_due_dates, "Date d'échéance"].dt.year.astype(int).astype(str)
    output["Projet / Code Affaire"] = (
        output["Nom Projet"].replace("", "Non renseigné")
        + " / "
        + output["Code Affaire"].replace("", "Non renseigné")
    )
    output = output.drop(columns=["MOIS"])
    return output.reset_index(drop=True)


def clean_monthly_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()

    output = df.copy()
    output.columns = [str(column).strip() for column in output.columns]
    for column in [
        "dépenses prévisionnelles",
        "Dépenses réelles",
        "ecarts",
        "Retard en jours",
        "Montant de pénalité",
    ]:
        if column in output.columns:
            output[column] = parse_amount(output[column])
    return output.dropna(how="all").reset_index(drop=True)


def build_monthly_forecast(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame(columns=["Mois", "Montant", "Nombre de factures", "Cumul"])

    valid = df[df["Date d'échéance"].notna() & (df["Mois"] != "Date invalide")].copy()
    if valid.empty:
        return pd.DataFrame(columns=["Mois", "Montant", "Nombre de factures", "Cumul"])

    forecast = (
        valid.groupby("Mois", as_index=False)
        .agg(Montant=("Montant", "sum"), **{"Nombre de factures": ("Montant", "size")})
        .sort_values("Mois")
    )
    forecast["Cumul"] = forecast["Montant"].cumsum()
    return forecast


def date_quality_summary(df: pd.DataFrame) -> dict[str, float | int]:
    if df.empty or "Date d'échéance valide" not in df.columns:
        return {"total_rows": 0, "invalid_rows": 0, "invalid_share": 0.0}
    total_rows = int(len(df))
    invalid_rows = int((~df["Date d'échéance valide"]).sum())
    invalid_share = invalid_rows / total_rows if total_rows else 0.0
    return {
        "total_rows": total_rows,
        "invalid_rows": invalid_rows,
        "invalid_share": invalid_share,
    }


def filter_invoices(
    df: pd.DataFrame,
    suppliers: list[str],
    codes: list[str],
    project_names: list[str],
    months: list[str],
    min_amount: float,
) -> pd.DataFrame:
    output = df.copy()
    if suppliers:
        output = output[output["Nom du fournisseur"].isin(suppliers)]
    if codes:
        output = output[output["Code Affaire"].isin(codes)]
    if project_names:
        output = output[output["Nom Projet"].isin(project_names)]
    if months:
        output = output[output["Mois"].isin(months)]
    output = output[output["Montant"] >= min_amount]
    return output
