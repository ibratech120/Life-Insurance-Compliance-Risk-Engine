# Adveez — Life-Insurance Compliance Engine

> **Moteur de conformité et de contrôle des investissements pour l'Assurance-Vie**
> Solvency II • Pre/Post-Trade • ESG Scoring

---

## Table of Contents

1. [Overview](#overview)
2. [Regulatory Context](#regulatory-context)
3. [Architecture & Components](#architecture--components)
4. [Data Model](#data-model)
5. [Compliance Rules Implemented](#compliance-rules-implemented)
6. [Getting Started](#getting-started)
7. [Usage](#usage)
8. [Sample Output](#sample-output)
9. [Project Structure](#project-structure)
10. [Roadmap](#roadmap)
11. [License](#license)

---

## Overview

**Adveez** is a proof-of-concept compliance engine designed for **French life-insurance asset management**. It demonstrates how regulatory concentration limits, credit-quality constraints, and ESG scoring can be computed programmatically across multiple fund types:

| Fund Type | Description |
|---|---|
| **Fonds en Euros** | Capital-guaranteed general account — strict regulatory limits apply (Solvency II dispersion, minimum bond quota, High-Yield cap). |
| **Unités de Compte (UC / FID)** | Dedicated investment funds linked to policyholder accounts — subject to issuer-concentration limits. |

The engine is delivered in **three complementary implementations** that cover the full technology stack commonly found in insurance back-office environments:

| Layer | File | Purpose |
|---|---|---|
| **Database** | `schema_and_checks.sql` | Relational schema + standalone compliance queries (portable SQL) |
| **Application** | `Mod_Compliance_Engine.bas` | VBA module for MS Access / Excel — daily batch controls & breach logging |
| **Analytics / Demo** | `compliance_engine_demo.py` | Self-contained Python demo with in-memory SQLite, Pandas output, and synthetic data |

---

## Regulatory Context

### Solvency II — Directive 2009/138/EC

The controls implemented in this engine are inspired by the **Solvency II** prudential framework governing European insurance undertakings. Key principles addressed:

- **Prudent Person Principle (Art. 132)** — Assets must be invested in a manner appropriate to the nature and duration of insurance liabilities.
- **Issuer Dispersion Rule** — Maximum **5 %** exposure to any single corporate issuer (sovereign debt and UCITS funds are exempt).
- **Asset-Liability Matching** — Fonds en Euros must maintain a minimum proportion in fixed-income instruments (≥ 75 % bonds).
- **Credit Quality Monitoring** — Speculative-grade (High-Yield: BB, B, CCC) exposure is capped at **10 %** for guaranteed funds.

### ESG Integration

Each security carries a proprietary **ESG score** (0–100). The engine computes a weighted-average ESG score per fund, supporting future SFDR (Sustainable Finance Disclosure Regulation) reporting.

---

## Architecture & Components

```
┌─────────────────────────────────────────────────────────────┐
│                        DATA LAYER                           │
│  schema_and_checks.sql                                      │
│  ┌───────────────────┐ ┌──────────────────────────────────┐ │
│  │ Referentiel_Titres│ │ Portefeuilles_AssuranceVie       │ │
│  │ (Securities)      │ │ (Insurance Funds)                │ │
│  └────────┬──────────┘ └──────────────┬───────────────────┘ │
│           │         ┌─────────────────┤                     │
│           ▼         ▼                 ▼                     │
│  ┌────────────────────────┐  ┌────────────────────────────┐ │
│  │ Positions_Inventaire   │  │ Historique_Breaches        │ │
│  │ (Portfolio Holdings)   │  │ (Audit Trail)              │ │
│  └────────────────────────┘  └────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    APPLICATION LAYER                         │
│  Mod_Compliance_Engine.bas  (VBA / MS Access)               │
│  ┌──────────────────┐  ┌──────────────────────────────────┐ │
│  │ Data Cleansing   │  │ Breach Detection & Logging       │ │
│  │ (ISIN normaliz.) │  │ (INSERT → Historique_Breaches)   │ │
│  └──────────────────┘  └──────────────────────────────────┘ │
│  ┌──────────────────────────────────────────────────────────┐│
│  │ User Notification (MsgBox alerts)                        ││
│  └──────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────┐
│                    ANALYTICS LAYER                           │
│  compliance_engine_demo.py  (Python / Pandas)               │
│  ┌──────────────────┐  ┌──────────────────────────────────┐ │
│  │ In-memory SQLite  │  │ Concentration Analysis          │ │
│  │ (synthetic data)  │  │ Duration / HY / ESG Metrics     │ │
│  └──────────────────┘  └──────────────────────────────────┘ │
└─────────────────────────────────────────────────────────────┘
```

---

## Data Model

### `Referentiel_Titres` — Securities Master

| Column | Type | Description |
|---|---|---|
| `ISIN` | `VARCHAR(12)` PK | International Securities Identification Number |
| `Nom_Titre` | `VARCHAR(100)` | Security name |
| `Classe_Actif` | `VARCHAR(30)` | Asset class: `Obligation`, `Action`, `OPCVM`, `Derive_Taux` |
| `Emetteur` | `VARCHAR(50)` | Issuer name |
| `Notation_Credit` | `VARCHAR(10)` | Credit rating (AAA → CCC) |
| `Score_ESG` | `INT` | ESG score (0–100) |
| `Duration_Modifiee` | `DECIMAL(5,2)` | Modified duration in years |

### `Portefeuilles_AssuranceVie` — Insurance Funds

| Column | Type | Description |
|---|---|---|
| `ID_Fonds` | `VARCHAR(20)` PK | Fund identifier |
| `Nom_Fonds` | `VARCHAR(50)` | Fund name |
| `Type_Support` | `VARCHAR(20)` | `Fonds_Euros` or `Unite_Compte` |
| `AUM_Total` | `DECIMAL(18,2)` | Total assets under management (EUR) |

### `Positions_Inventaire` — Portfolio Holdings

| Column | Type | Description |
|---|---|---|
| `ID_Position` | `INT` PK | Position identifier |
| `Date_Arrete` | `DATE` | Reporting date |
| `ID_Fonds` | `VARCHAR(20)` FK | Fund reference |
| `ISIN` | `VARCHAR(12)` FK | Security reference |
| `Quantite` | `DECIMAL(18,4)` | Quantity held |
| `Valeur_Marche_EUR` | `DECIMAL(18,2)` | Market value in EUR |

### `Historique_Breaches` — Breach Audit Log

| Column | Type | Description |
|---|---|---|
| `ID_Breach` | `INT` PK | Breach identifier |
| `Date_Controle` | `DATE` | Control date |
| `ID_Fonds` | `VARCHAR(20)` | Fund in breach |
| `Regle_Violee` | `VARCHAR(100)` | Description of the violated rule |
| `Exposition_Actuelle_Pct` | `DECIMAL(6,2)` | Actual exposure (%) |
| `Limite_Reglementaire_Pct` | `DECIMAL(6,2)` | Regulatory limit (%) |
| `Statut` | `VARCHAR(20)` | `OPEN_BREACH` / `RESOLVED` |

---

## Compliance Rules Implemented

### Rule 1 — Issuer Concentration (Solvency II Dispersion)

| Parameter | Value |
|---|---|
| **Scope** | All funds (Fonds Euros + Unités de Compte) |
| **Limit** | ≤ 5 % of AUM per corporate issuer |
| **Exemptions** | Sovereign debt (e.g. OAT France), UCITS/OPCVM, CCP (Clearing Houses) |
| **Status** | ✅ Fully implemented in all three layers |

### Rule 2 — Minimum Bond Quota (Fonds en Euros)

| Parameter | Value |
|---|---|
| **Scope** | Fonds en Euros only |
| **Limit** | ≥ 75 % of AUM in fixed-income (`Obligation`) |
| **Purpose** | Asset-liability matching for capital-guaranteed products |
| **Status** | ✅ Implemented in SQL layer |

### Rule 3 — High-Yield Cap (Fonds en Euros)

| Parameter | Value |
|---|---|
| **Scope** | Fonds en Euros only |
| **Limit** | ≤ 10 % of AUM in speculative-grade bonds (BB, B, CCC) |
| **Purpose** | Credit-risk containment for guaranteed funds |
| **Status** | ✅ Implemented in VBA layer |

### Market Risk Indicators (Informational)

| Metric | Description |
|---|---|
| **Portfolio Duration** | Weighted-average modified duration — measures interest-rate sensitivity |
| **High-Yield Weight** | Percentage of sub-investment-grade exposure |
| **ESG Score** | Weighted-average ESG score per fund |

---

## Getting Started

### Prerequisites

| Dependency | Version | Purpose |
|---|---|---|
| **Python** | 3.8+ | Run the analytics demo |
| **pandas** | ≥ 1.3 | DataFrame operations and tabular output |
| **sqlite3** | (stdlib) | In-memory relational database |
| **MS Access** *(optional)* | 2016+ | Run the VBA module natively |

### Installation

```bash
# Clone the repository
git clone <repository-url>
cd Adveez

# Install Python dependencies
pip install pandas
```

No additional configuration is required — the Python demo is entirely self-contained with synthetic data.

---

## Usage

### Python Demo (recommended for quick testing)

```bash
python compliance_engine_demo.py
```

This will:
1. Create an in-memory SQLite database with the full schema
2. Populate it with synthetic securities, funds, and positions (including deliberate breaches)
3. Run issuer-concentration checks and print compliance status
4. Compute market-risk indicators (duration, high-yield weight, ESG) and print the results

### SQL Queries (standalone)

The file `schema_and_checks.sql` can be executed against any SQL-compliant database (SQL Server, PostgreSQL, MySQL, SQLite) after creating the schema and inserting data. It contains:
- **DDL** — Full table-creation statements
- **Control 1** — Issuer concentration query (returns only breaches)
- **Control 2** — Bond-quota query for Fonds en Euros (returns only breaches)

### VBA Module (MS Access)

1. Import `Mod_Compliance_Engine.bas` into your MS Access database
2. Ensure the four tables (`Referentiel_Titres`, `Portefeuilles_AssuranceVie`, `Positions_Inventaire`, `Historique_Breaches`) exist and are populated
3. Run `Run_Daily_Compliance_Checks` from the VBA editor or bind it to a macro button

The module will:
- Normalize ISIN codes (trim + uppercase)
- Purge same-day duplicate checks
- Detect breaches and log them to `Historique_Breaches`
- Display a summary `MsgBox` alert

---

## Sample Output

```
=== 1. CONTROLE DE CONCENTRATION EMETTEUR (SOLVABILITE II) ===
     ID_Fonds Type_Support       Emetteur  Exposition_EUR  Exposition_Pct  Limite_Max_Pct            Statut_Compliance
FE_GARANTI_01  Fonds_Euros  Etat Francais      65000000.0            65.0             5.0 OK (Exempté Souverain/OPCVM)
FE_GARANTI_01  Fonds_Euros         Amundi      13500000.0            13.5             5.0 OK (Exempté Souverain/OPCVM)
FE_GARANTI_01  Fonds_Euros         Altice      12000000.0            12.0             5.0                BREACH (> 5%)
FE_GARANTI_01  Fonds_Euros    BNP Paribas       7500000.0             7.5             5.0                BREACH (> 5%)
FE_GARANTI_01  Fonds_Euros Clearing_Eurex       2000000.0             2.0             5.0 OK (Exempté Souverain/OPCVM)
  UC_DYNAM_02 Unite_Compte         Amundi      45500000.0            91.0             5.0 OK (Exempté Souverain/OPCVM)
  UC_DYNAM_02 Unite_Compte           LVMH       4500000.0             9.0             5.0                BREACH (> 5%)

=== 2. INDICATEURS DE RISQUE DE MARCHE (DURATION, CREDIT HY, ESG) ===
     ID_Fonds Type_Support  Duration_Portefeuille_Yrs  Poids_High_Yield_Pct  Score_ESG_Moyen
FE_GARANTI_01  Fonds_Euros                       5.06                  12.0             80.3
  UC_DYNAM_02 Unite_Compte                       0.18                   0.0             89.3
```

### Reading the Results

- **Altice** at 12 % exposure on the Fonds Euros → **BREACH** (limit: 5 %)
- **BNP Paribas** at 7.5 % (combined bond + equity) → **BREACH** (limit: 5 %)
- **LVMH** at 9 % on the Unité de Compte → **BREACH** (limit: 5 %)
- **Etat Français**, **Amundi** (OPCVM), and **Clearing Eurex** (CCP) are **exempt** from the concentration rule
- Fonds Euros portfolio duration: **5.06 years** — reflecting significant OAT exposure
- Fonds Euros High-Yield weight: **12 %** — above the 10 % internal limit (breach candidate)
- ESG scores: **80.3** (Fonds Euros) and **89.3** (UC) — strong overall ESG positioning

---

## Project Structure

```
Adveez/
├── README.md                       # This file
├── compliance_engine_demo.py       # Python demo — self-contained analytics & compliance checks
├── schema_and_checks.sql           # SQL DDL + compliance queries (portable SQL)
└── Mod_Compliance_Engine.bas       # VBA module — daily batch controls for MS Access / Excel
```

---

## Roadmap

- [ ] **Pre-trade simulation** — Evaluate compliance impact of proposed trades before execution
- [ ] **Dashboard (Streamlit / Power BI)** — Interactive breach visualization and drill-down
- [ ] **SCR Market Module** — Compute Solvency Capital Requirement shocks (equity, spread, interest rate)
- [ ] **SFDR Reporting** — Automated PAI (Principal Adverse Impact) indicators and Article 8/9 classification
- [ ] **Automated remediation workflows** — Suggest rebalancing trades to cure open breaches
- [ ] **Multi-currency support** — FX exposure monitoring and hedging-ratio checks
- [ ] **Historical trend analysis** — Track compliance metrics over time with charting

---

## License

This project is provided as a **proof-of-concept / portfolio demonstration**. All data is synthetic and does not represent real portfolio positions or investment advice.

---

*Built with ❤️ for insurance risk & compliance professionals.*
