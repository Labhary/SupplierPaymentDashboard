from __future__ import annotations

from pathlib import Path

import pandas as pd


OUTPUT_PATH = Path(__file__).resolve().parents[1] / "sample_supplier_payments.xlsx"


def main() -> None:
    invoices = pd.DataFrame(
        [
            ["Atlas Travaux", "AT-2026-001", 46139, 125000, "27.01.2026", "CA-100", "Résidence Palmier", "04"],
            ["Maroc Electric", 8841, 46161, 48500.75, "18.02.2026", "RB-210", "Tour Agdal", "05"],
            ["Bureau Conseil Nord", "BCN-332", 46202, 31500, "02.03.2026", "SG-001", "Audit technique", "06"],
            ["Logistique Atlas", "LA-771", 46220, 62400, "18.03.2026", "CA-100", "Résidence Palmier", "06"],
            ["Steel Maghreb", "SM-1902", 46247, 210000, "14.04.2026", "TG-450", "Port Extension", "07"],
            ["Maroc Electric", "ME-8990", 46265, 73250, "02.05.2026", "RB-210", "Tour Agdal", "07"],
            ["Office Supplies Pro", "OSP-104", 46292, 9800, "29.05.2026", "SG-001", "Frais généraux", "08"],
            ["Atlas Travaux", "AT-2026-019", 46310, 158900, "16.06.2026", "TG-450", "Port Extension", "08"],
        ],
        columns=[
            "Nom du fournisseur",
            "Numéro réf. fournisseur",
            "Date d'échéance",
            "Montant",
            "Date comptable",
            "Code Affaire",
            "Nom Projet",
            "MOIS",
        ],
    )

    monthly = (
        invoices.assign(
            mois=pd.to_datetime(
                invoices["Date d'échéance"],
                unit="D",
                origin="1899-12-30",
                errors="coerce",
            ).dt.to_period("M").astype(str)
        )
        .groupby("mois", as_index=False)
        .agg(**{"dépenses prévisionnelles": ("Montant", "sum")})
    )
    monthly["Dépenses réelles"] = (monthly["dépenses prévisionnelles"] * 0.82).round(2)
    monthly["ecarts"] = monthly["dépenses prévisionnelles"] - monthly["Dépenses réelles"]
    monthly["Retard en jours"] = [index % 6 for index in range(len(monthly))]
    monthly["Montant de pénalité"] = monthly["Retard en jours"].apply(lambda value: round(max(value, 0) * 250, 2))

    with pd.ExcelWriter(OUTPUT_PATH, engine="openpyxl") as writer:
        invoices.to_excel(writer, sheet_name="Feuil1", index=False, startrow=1, startcol=1)
        monthly.to_excel(writer, sheet_name="Feuil2", index=False)

    print(f"Sample Excel created: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
