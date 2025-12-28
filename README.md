# Automated Financial Portfolio Analyzer (VBA)

![Dashboard Preview](dashboard.png)

---

## 🇬🇧 Project Description
This project implements a VBA-based tool for financial data analysis and portfolio risk assessment. The objective is to automate the full data processing workflow—from raw CSV import to final investment reporting.

The algorithm processes historical stock market data to calculate key risk-adjusted performance metrics, facilitating comparative analysis between assets (e.g., LVMH vs. Air Liquide) and a benchmark index (CAC40).

### Key Features
* **Data ETL (Extract, Transform, Load):** Algorithms for parsing raw CSV files, cleaning time-series data, and handling missing values.
* **Financial Modeling:** Calculation of Daily/Annual Returns, Volatility, Beta (systematic risk), and Sharpe Ratio.
* **Decision Logic:** Implementation of conditional logic to generate investment recommendations based on computed risk metrics.
* **Automated Reporting:** Programmatic generation of a PDF report containing analysis results and visualizations.

> **Output Example:**
> A sample of the automatically generated report is available for download:
> 📄 **[View Sample Report (PDF)](Sample_Report.pdf)**

---

## 🇫🇷 Description du Projet
Ce projet implémente un outil VBA dédié à l'analyse de données financières et à l'évaluation des risques de portefeuille. L'objectif est d'automatiser le flux de traitement des données, de l'importation de fichiers CSV bruts jusqu'au reporting d'investissement.

L'algorithme traite des données boursières historiques pour calculer des indicateurs de performance ajustés au risque, facilitant l'analyse comparative entre des actifs (ex: LVMH vs Air Liquide) et un indice de référence (CAC40).

### Fonctionnalités Principales
* **Traitement des Données (ETL) :** Algorithmes de lecture de fichiers CSV, nettoyage des séries temporelles et gestion des valeurs manquantes.
* **Modélisation Financière :** Calcul des Rendements (Journaliers/Annuels), de la Volatilité, du Bêta (risque systématique) et du Ratio de Sharpe.
* **Logique Décisionnelle :** Implémentation de logique conditionnelle pour générer des recommandations basées sur les métriques de risque.
* **Reporting Automatisé :** Génération programmatique d'un rapport PDF incluant les résultats de l'analyse et les visualisations.

> **Exemple de Résultat :**
> Un exemple du rapport généré automatiquement est disponible ici :
> 📄 **[Voir le Rapport (PDF)](Sample_Report.pdf)**

---

### Technical Stack
* **Language:** VBA (Visual Basic for Applications)
* **Environment:** Microsoft Excel
* **Input Data:** Historical stock quotes (CSV format)
