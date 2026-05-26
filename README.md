# Supplier Payment Dashboard

Application locale de pilotage des décaissements fournisseurs construite avec Streamlit.

Le projet permet d’analyser des prévisions de paiements fournisseurs à partir de fichiers Excel, avec :
- visualisation exécutive ;
- KPI financiers ;
- analyse par fournisseur, projet et mois ;
- exports Excel/PDF ;
- rapport HTML interactif offline destiné aux managers ;
- packaging Windows portable sans installation Python.

---

# Fonctionnalités principales

## Analyse financière

- Prévision mensuelle des décaissements fournisseurs.
- Analyse des montants à payer.
- Répartition par fournisseur.
- Répartition par projet / code affaire.
- Suivi des échéances fournisseurs.
- KPI exécutifs dynamiques.

## Dashboard interactif

- Interface Streamlit locale.
- Thème sombre orienté reporting exécutif.
- Filtres dynamiques :
  - fournisseur ;
  - projet ;
  - mois ;
  - montant minimum.
- Graphiques interactifs Plotly.
- Table détaillée des factures.
- Top fournisseurs.
- Prévision mensuelle cumulée.

## Exports manager

Le projet inclut un centre d’export manager permettant de générer :

- Excel BI-ready ;
- PDF synthèse exécutive ;
- PDF visualisations ;
- PDF rapport complet ;
- rapport HTML interactif offline.

## Rapport HTML interactif offline

Le rapport HTML exporté :

- fonctionne sans internet ;
- ne nécessite ni Python ni Streamlit ;
- s’ouvre directement dans un navigateur ;
- contient :
  - KPI ;
  - graphiques ;
  - filtres ;
  - recherche ;
  - pagination ;
  - détails fournisseurs ;
  - export CSV local.

Le manager peut interagir avec les données exportées sans accès à la plateforme analyste.

---

# Architecture générale

## Plateforme analyste

Technologies principales :

- Python
- Streamlit
- Pandas
- Plotly
- OpenPyXL
- ReportLab
- Matplotlib

## Livraison manager

Exports autonomes :

- HTML offline interactif ;
- PDF ;
- Excel BI-ready.

## Packaging Windows

Le projet peut être distribué comme application Windows portable via PyInstaller.

---

# Installation

## Cloner le projet

```bash
git clone https://github.com/Labhary/SupplierPaymentDashboard.git
cd SupplierPaymentDashboard
```

## Installer les dépendances

```bash
pip install -r requirements.txt
```

---

# Lancement

```bash
streamlit run app.py
```

L’application s’ouvre ensuite automatiquement dans le navigateur local.

---

# Tests

```bash
python -m pytest
```

Les tests couvrent notamment :

- nettoyage des montants ;
- parsing des dates Excel ;
- formats français ;
- agrégations mensuelles ;
- exports ;
- rapport HTML interactif ;
- compatibilité des filtres ;
- absence de dépendances externes ;
- régressions JavaScript.

---

# Structure Excel attendue

## Feuil1 — Factures fournisseurs

Colonnes attendues :

- Nom du fournisseur
- Numéro réf. fournisseur
- Date d'échéance
- Montant
- Date comptable
- Code Affaire
- Nom Projet
- MOIS

## Feuil2 — Synthèse mensuelle (optionnelle)

Colonnes possibles :

- mois
- dépenses prévisionnelles
- dépenses réelles
- ecarts
- Retard en jours
- Montant de pénalité

---

# Tolérance des données Excel

L’application gère automatiquement :

- dates Excel sérialisées ;
- dates texte françaises ;
- dates ISO ;
- variations d’accents ;
- espaces supplémentaires ;
- différences majuscules/minuscules ;
- lignes d’en-tête décalées ;
- valeurs manquantes ;
- formats numériques hétérogènes.

---

# Générer un fichier d’exemple

```bash
python scripts/create_sample_excel.py
```

Le script génère :

```text
sample_supplier_payments.xlsx
```

avec :
- Feuil1 ;
- Feuil2 ;
- données réalistes de démonstration.

---

# Packaging Windows portable

Le projet peut être transformé en application Windows autonome.

L’utilisateur final n’a pas besoin :
- de Python ;
- de VS Code ;
- de Streamlit.

## Construire l’application

```bat
build_exe.bat
```

Le script :
- crée un environnement virtuel ;
- installe les dépendances ;
- construit l’exécutable PyInstaller.

## Résultat

```text
dist\SupplierPaymentDashboard\
```

Lancement :

```text
SupplierPaymentDashboard.exe
```

L’application démarre automatiquement sur localhost dans le navigateur.

---

# Exports disponibles

## Excel BI-ready

Compatible Power BI et outils BI classiques.

## PDF

- synthèse exécutive ;
- visualisations ;
- rapport complet.

## HTML interactif

Rapport autonome contenant :
- KPI ;
- graphiques ;
- filtres ;
- détails fournisseurs ;
- recherche ;
- pagination ;
- export CSV.

---

# Cas d’usage

Le projet est adapté pour :

- équipes finance ;
- contrôle de gestion ;
- pilotage fournisseurs ;
- reporting exécutif ;
- prévision de décaissements ;
- analyse projet / affaire.

---

# Limites actuelles

Le périmètre actuel couvre principalement :

- les factures fournisseurs ;
- les décaissements prévisionnels.

Une plateforme complète de prévision de trésorerie nécessiterait également :

- soldes bancaires ;
- encaissements clients ;
- statuts de paiement ;
- scénarios de trésorerie ;
- projections multi-sources.

---

# Améliorations futures possibles

- gestion des paiements réalisés ;
- scénarios de trésorerie ;
- gestion multi-entités ;
- authentification utilisateurs ;
- génération PowerPoint ;
- base de données ;
- synchronisation ERP ;
- API d’import automatique.

---

# Auteur

Projet développé par Labhary.