# Dashboard de prévision des décaissements fournisseurs

Application locale Streamlit pour visualiser les paiements fournisseurs prévisionnels à partir d'un fichier Excel.

## Objectif

Ce projet aide à analyser les décaissements fournisseurs prévus par mois: montants à payer, échéances, retards, fournisseurs principaux et répartition par projet ou code affaire.

Ce n'est pas encore une plateforme complète de prévision de trésorerie. Le périmètre actuel couvre uniquement les factures fournisseurs et les comptes à payer. Une vision de trésorerie complète nécessiterait aussi les soldes bancaires, les encaissements clients, les statuts de paiement et des scénarios.

## Installation

```bash
pip install -r requirements.txt
```

## Lancement

```bash
streamlit run app.py
```

L'application s'ouvre ensuite dans le navigateur local.

## Tests

```bash
python -m pytest
```

Les tests couvrent le nettoyage des montants, les dates Excel, les dates texte, les colonnes manquantes, les agrégations mensuelles, les KPI et les filtres.

## Générer un fichier Excel d'exemple

```bash
python scripts/create_sample_excel.py
```

Le script crée `sample_supplier_payments.xlsx` à la racine du projet avec deux feuilles:

- `Feuil1`: factures fournisseurs réalistes.
- `Feuil2`: synthèse mensuelle.

## Format Excel attendu

Le fichier doit contenir:

- `Feuil1`: base des factures fournisseurs.
- `Feuil2`: synthèse mensuelle, optionnelle mais prise en charge.

Colonnes attendues dans `Feuil1`:

- Nom du fournisseur
- Retard en jours
- Numéro réf. fournisseur
- Date d'échéance
- Montant
- Montant d'origine
- Date comptable
- Projet partenaire
- Code Affaire
- Nom Projet

Colonnes attendues dans `Feuil2`:

- mois
- dépenses prévisionnelles
- Dépenses réelles
- ecarts
- Retard en jours
- Montant de pénalité

L'application détecte automatiquement la ligne d'en-tête dans `Feuil1`, même si les données commencent autour de `B2:K1328`.

## Problèmes fréquents de colonnes Excel

L'application tolère plusieurs variations courantes:

- accents présents ou absents;
- majuscules / minuscules;
- espaces supplémentaires;
- retours à la ligne dans les en-têtes;
- variantes comme `Numéro réf. Fournisseur`, `Numéro réf. fournisseur` ou `Numero ref fournisseur`.

Si une colonne attendue est absente, l'application crée une valeur vide ou `Non renseigné` quand c'est possible, puis affiche un avertissement au lieu de bloquer le tableau de bord.

## Fonctionnalités

- Import d'un fichier Excel.
- Lecture automatique de `Feuil1` et `Feuil2`.
- Nettoyage des montants, dates Excel, dates texte et valeurs manquantes.
- KPI: total à payer, nombre de factures, fournisseurs, mois le plus élevé, montant en retard.
- Graphiques interactifs Plotly:
  - décaissements prévisionnels par mois;
  - cumul des décaissements;
  - top 10 fournisseurs;
  - montant par projet / code affaire;
  - répartition en retard / non en retard.
- Filtres: fournisseur, projet / code affaire, mois, statut retard et montant minimum.
- Tables: factures détaillées, synthèse mensuelle et top fournisseurs.
- Export CSV des factures nettoyées et de la prévision mensuelle.

## Améliorations futures possibles

- Ajouter les soldes bancaires.
- Ajouter les créances clients.
- Ajouter un statut payé / non payé.
- Ajouter des scénarios de paiement.
- Ajouter un export PDF du rapport.

## Packaging Windows portable

Le projet peut etre transforme en application Windows portable avec PyInstaller. L'utilisateur final n'a pas besoin de VS Code et ne tape pas de commande Streamlit.

### Construire l'application

Depuis le dossier du projet, double-cliquer sur:

```bat
build_exe.bat
```

Le script:

- cree un environnement virtuel `.venv` si necessaire;
- installe `requirements.txt`;
- installe PyInstaller;
- genere une application en mode one-folder.

### Emplacement du resultat

L'executable est genere ici:

```text
dist\SupplierPaymentDashboard\SupplierPaymentDashboard.exe
```

### Lancer l'application packagee

Envoyer le dossier complet:

```text
dist\SupplierPaymentDashboard
```

L'utilisateur lance:

```text
SupplierPaymentDashboard.exe
```

L'application demarre un serveur local Streamlit et ouvre automatiquement le navigateur sur `localhost`.

### Note Windows Defender

Comme l'executable n'est pas signe numeriquement, Windows Defender ou SmartScreen peut afficher un avertissement au premier lancement. C'est normal pour un executable local non signe.
