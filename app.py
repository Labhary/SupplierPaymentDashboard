from __future__ import annotations

import streamlit as st

from src.charts import (
    cumulative_line,
    monthly_bar,
    project_bar,
    supplier_pareto_chart,
    suppliers_treemap,
    top_suppliers_bar,
)
from src.data_loader import ExcelStructureError, load_workbook
from src.insights import generate_executive_insights
from src.manager_exports import (
    build_executive_pdf,
    build_full_pdf,
    build_interactive_html_report,
    build_manager_excel,
    build_visuals_pdf,
)
from src.metrics import (
    amount_by_project,
    compute_executive_metrics,
    compute_kpis,
    supplier_pareto,
    top_suppliers,
)
from src.transformations import (
    build_monthly_forecast,
    clean_invoices,
    clean_monthly_summary,
    date_quality_summary,
    filter_invoices,
)
from src.ui import (
    configure_page,
    format_money_columns,
    prepare_invoice_display,
    prepare_display_dataframe,
    render_downloads,
    render_executive_analysis,
    render_footer,
    render_header,
    render_insights,
    render_kpis,
    render_metadata,
    render_scope_box,
    render_sidebar_filters,
    render_upload_box,
)


def render_empty_state() -> None:
    st.info(
        "Importez un fichier Excel contenant Feuil1 pour afficher la prévision des "
        "décaissements fournisseurs."
    )


def main() -> None:
    configure_page()
    render_header()
    render_scope_box()

    uploaded_file = render_upload_box()
    if uploaded_file is None:
        render_empty_state()
        return

    with st.spinner("Traitement du fichier Excel en cours..."):
        try:
            workbook = load_workbook(uploaded_file)
            invoices = clean_invoices(workbook.invoices)
            source_summary = clean_monthly_summary(workbook.monthly_summary)
        except ExcelStructureError as exc:
            st.error(str(exc))
            st.stop()
        except Exception as exc:
            st.error(
                "Le fichier n'a pas pu être traité. Vérifiez qu'il s'agit d'un Excel valide "
                f"et que les feuilles attendues existent. Détail: {exc}"
            )
            st.stop()

    for warning in workbook.warnings:
        st.warning(warning)

    quality = date_quality_summary(invoices)
    if quality["invalid_rows"]:
        level = st.warning if quality["invalid_share"] >= 0.10 else st.info
        level(
            f"{quality['invalid_rows']} ligne(s) sur {quality['total_rows']} ont une date d'échéance invalide "
            "ou vide. Elles restent visibles dans les données mais sont ignorées dans les graphiques mensuels."
        )

    filters = render_sidebar_filters(invoices)
    filtered = filter_invoices(
        invoices,
        suppliers=filters["suppliers"],
        codes=filters["codes"],
        project_names=filters["project_names"],
        months=filters["months"],
        min_amount=filters["min_amount"],
    )
    forecast = build_monthly_forecast(filtered)
    supplier_ranking = top_suppliers(filtered)
    project_amounts = amount_by_project(filtered)
    executive_metrics = compute_executive_metrics(filtered, forecast)
    kpis = compute_kpis(filtered, forecast)
    insights = generate_executive_insights(filtered, forecast, executive_metrics)
    pareto = supplier_pareto(filtered)

    if filtered.empty:
        st.info("Aucune facture ne correspond aux filtres actifs. Réinitialisez ou élargissez les filtres.")

    executive_tab, detail_tab, data_tab = st.tabs(["Vue exécutive", "Analyse détaillée", "Données & exports"])

    with executive_tab:
        st.markdown('<div class="section-title">Métadonnées du fichier</div>', unsafe_allow_html=True)
        render_metadata(invoices)
        st.info(
            "Le fichier source Feuil1 ne contient pas de retard par facture. "
            "L'analyse des retards est limitée à la synthèse mensuelle Feuil2 si elle est renseignée."
        )

        st.markdown('<div class="section-title">Indicateurs clés</div>', unsafe_allow_html=True)
        render_kpis(kpis)

        st.markdown('<div class="section-title">Analyse exécutive</div>', unsafe_allow_html=True)
        render_executive_analysis(executive_metrics)
        if executive_metrics["large_month_severity"] == "Élevé":
            st.warning("Un mois concentre une part élevée des décaissements prévisionnels.")
        if executive_metrics["top5_supplier_severity"] == "Élevé":
            st.warning("Concentration fournisseur élevée: les 5 premiers fournisseurs dominent les engagements.")

        st.markdown('<div class="section-title">Insights automatiques</div>', unsafe_allow_html=True)
        render_insights(insights)

        st.markdown('<div class="section-title">Décaissements et fournisseurs</div>', unsafe_allow_html=True)
        left, right = st.columns(2)
        left.plotly_chart(monthly_bar(forecast), use_container_width=True)
        right.plotly_chart(top_suppliers_bar(supplier_ranking), use_container_width=True)

    with detail_tab:
        st.markdown('<div class="section-title">Analyse mensuelle avancée</div>', unsafe_allow_html=True)
        left, right = st.columns(2)
        left.plotly_chart(cumulative_line(forecast), use_container_width=True)
        right.plotly_chart(project_bar(project_amounts), use_container_width=True)

        st.markdown('<div class="section-title">Concentration fournisseurs</div>', unsafe_allow_html=True)
        left, right = st.columns(2)
        left.plotly_chart(suppliers_treemap(filtered), use_container_width=True)
        right.plotly_chart(supplier_pareto_chart(pareto), use_container_width=True)

    with data_tab:
        st.markdown('<div class="section-title">Centre d’export manager</div>', unsafe_allow_html=True)
        st.caption("Les exports utilisent uniquement les données actuellement filtrées.")
        excel_bytes = build_manager_excel(filtered, forecast, supplier_ranking)
        pdf_summary = build_executive_pdf(filters, kpis, executive_metrics, insights, forecast, supplier_ranking)
        pdf_visuals = build_visuals_pdf(forecast, supplier_ranking, project_amounts, pareto)
        pdf_full = build_full_pdf(
            filters,
            kpis,
            executive_metrics,
            insights,
            forecast,
            supplier_ranking,
            project_amounts,
            pareto,
            filtered,
        )
        html_report = build_interactive_html_report(filters, filtered, kpis, executive_metrics, insights)
        col1, col2 = st.columns(2)
        col1.download_button(
            "Télécharger Excel BI-ready",
            data=excel_bytes,
            file_name="manager_export_bi_ready.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            disabled=filtered.empty,
            use_container_width=True,
        )
        col2.download_button(
            "Télécharger PDF synthèse exécutive",
            data=pdf_summary,
            file_name="synthese_executive_decaissements.pdf",
            mime="application/pdf",
            disabled=filtered.empty,
            use_container_width=True,
        )
        col3, col4 = st.columns(2)
        col3.download_button(
            "Télécharger PDF visualisations",
            data=pdf_visuals,
            file_name="visualisations_decaissements.pdf",
            mime="application/pdf",
            disabled=filtered.empty,
            use_container_width=True,
        )
        col4.download_button(
            "Télécharger PDF rapport complet",
            data=pdf_full,
            file_name="rapport_complet_decaissements.pdf",
            mime="application/pdf",
            disabled=filtered.empty,
            use_container_width=True,
        )
        st.download_button(
            "Télécharger rapport HTML interactif",
            data=html_report,
            file_name="Rapport.html",
            mime="text/html",
            disabled=filtered.empty,
            use_container_width=True,
        )

        st.markdown('<div class="section-title">Exports techniques</div>', unsafe_allow_html=True)
        render_downloads(filtered, forecast)

        table_tabs = st.tabs(["Factures détaillées", "Synthèse mensuelle", "Top fournisseurs"])
        with table_tabs[0]:
            display = prepare_invoice_display(filtered)
            st.dataframe(display, use_container_width=True, hide_index=True)

        with table_tabs[1]:
            if not source_summary.empty:
                st.caption("Synthèse issue de Feuil2")
                st.dataframe(
                    prepare_display_dataframe(
                        format_money_columns(
                            source_summary,
                            ["dépenses prévisionnelles", "Dépenses réelles", "ecarts", "Montant de pénalité"],
                        )
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
            st.caption("Prévision recalculée depuis les factures filtrées")
            st.dataframe(
                prepare_display_dataframe(format_money_columns(forecast, ["Montant", "Cumul"])),
                use_container_width=True,
                hide_index=True,
            )

        with table_tabs[2]:
            st.dataframe(
                prepare_display_dataframe(format_money_columns(supplier_ranking, ["Montant"])),
                use_container_width=True,
                hide_index=True,
            )

    render_footer()


if __name__ == "__main__":
    main()
