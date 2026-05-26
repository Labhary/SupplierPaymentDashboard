from __future__ import annotations

from dataclasses import dataclass
from typing import BinaryIO

import pandas as pd

from src.transformations import canonicalize_invoice_columns, normalize_column_label


REQUIRED_FEUIL1_COLUMNS = [
    "Nom du fournisseur",
    "Numéro réf. fournisseur",
    "Date d'échéance",
    "Montant",
    "Date comptable",
    "Code Affaire",
    "Nom Projet",
    "MOIS",
]

FEUIL2_COLUMNS = [
    "mois",
    "dépenses prévisionnelles",
    "Dépenses réelles",
    "ecarts",
    "Retard en jours",
    "Montant de pénalité",
]


@dataclass
class WorkbookData:
    invoices: pd.DataFrame
    monthly_summary: pd.DataFrame
    warnings: list[str]


class ExcelStructureError(ValueError):
    """Raised when the workbook cannot be interpreted as an invoice workbook."""


def _normalize_label(value: object) -> str:
    return normalize_column_label(value)


def _find_header_row(raw: pd.DataFrame, required_columns: list[str]) -> int | None:
    required = {_normalize_label(column) for column in required_columns}
    best_index: int | None = None
    best_score = 0

    for index, row in raw.iterrows():
        labels = {_normalize_label(value) for value in row.tolist()}
        score = len(required.intersection(labels))
        if score > best_score:
            best_score = score
            best_index = int(index)

    return best_index if best_score >= 4 else None


def _read_sheet_with_detected_header(
    excel_file: BinaryIO,
    sheet_name: str,
    required_columns: list[str],
) -> pd.DataFrame:
    raw = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, engine="openpyxl")
    header_row = _find_header_row(raw, required_columns)
    if header_row is None:
        raise ExcelStructureError(
            f"Impossible de trouver la ligne d'en-tête dans la feuille '{sheet_name}'."
        )

    excel_file.seek(0)
    df = pd.read_excel(
        excel_file,
        sheet_name=sheet_name,
        header=header_row,
        engine="openpyxl",
    )
    df.columns = [str(column).strip() for column in df.columns]
    if sheet_name == "Feuil1":
        df = canonicalize_invoice_columns(df)
    return df.dropna(how="all")


def _read_optional_summary(excel_file: BinaryIO) -> tuple[pd.DataFrame, list[str]]:
    warnings: list[str] = []
    try:
        excel_file.seek(0)
        summary = _read_sheet_with_detected_header(excel_file, "Feuil2", FEUIL2_COLUMNS)
        return summary, warnings
    except Exception as exc:
        warnings.append(
            "La feuille Feuil2 n'a pas pu être lue. Le tableau de synthèse mensuelle "
            f"sera reconstruit depuis Feuil1. Détail: {exc}"
        )
        return pd.DataFrame(), warnings


def load_workbook(uploaded_file: BinaryIO) -> WorkbookData:
    """Load Feuil1 and Feuil2 with tolerant header detection."""
    warnings: list[str] = []

    try:
        uploaded_file.seek(0)
        invoices = _read_sheet_with_detected_header(
            uploaded_file,
            "Feuil1",
            REQUIRED_FEUIL1_COLUMNS,
        )
    except ValueError as exc:
        raise ExcelStructureError(
            "Le fichier Excel doit contenir une feuille 'Feuil1' avec la base des factures."
        ) from exc

    missing = [column for column in REQUIRED_FEUIL1_COLUMNS if column not in invoices.columns]
    if missing:
        warnings.append(
            "Certaines colonnes attendues sont absentes de Feuil1: " + ", ".join(missing)
        )

    monthly_summary, summary_warnings = _read_optional_summary(uploaded_file)
    warnings.extend(summary_warnings)

    return WorkbookData(
        invoices=invoices,
        monthly_summary=monthly_summary,
        warnings=warnings,
    )
