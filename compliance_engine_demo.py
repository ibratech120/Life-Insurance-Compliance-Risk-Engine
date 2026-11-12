import sqlite3
import pandas as pd

def run_life_insurance_compliance_engine():
    # 1. Connexion à la base de données SQL (simule le serveur SQL / MS Access)
    conn = sqlite3.connect(":memory:")
    cursor = conn.cursor()

    # 2. Création des tables relationnelles
    cursor.executescript("""
        CREATE TABLE Referentiel_Titres (
            ISIN TEXT PRIMARY KEY,
            Nom_Titre TEXT,
            Classe_Actif TEXT,
            Emetteur TEXT,
            Notation_Credit TEXT,
            Score_ESG INTEGER,
            Duration_Modifiee REAL
        );

        CREATE TABLE Portefeuilles_AssuranceVie (
            ID_Fonds TEXT PRIMARY KEY,
            Nom_Fonds TEXT,
            Type_Support TEXT,
            AUM_Total REAL
        );

        CREATE TABLE Positions_Inventaire (
            ID_Position INTEGER PRIMARY KEY AUTOINCREMENT,
            ID_Fonds TEXT,
            ISIN TEXT,
            Valeur_Marche_EUR REAL,
            FOREIGN KEY (ID_Fonds) REFERENCES Portefeuilles_AssuranceVie(ID_Fonds),
            FOREIGN KEY (ISIN) REFERENCES Referentiel_Titres(ISIN)
        );
    """)

    # 3. Alimentation du Référentiel Produits de Marché (Obligations, Actions, OPCVM, Dérivés)
    titres = [
        ("FR0014007L00", "OAT France 2.5% 2032", "Obligation", "Etat Francais", "AA", 85, 6.8),
        ("FR0013250560", "BNP Paribas Snr 3.1% 2028", "Obligation", "BNP Paribas", "A", 78, 3.9),
        ("XS2010039999", "Altice France High Yield 5.8%", "Obligation", "Altice", "B", 45, 4.2),
        ("FR0000121014", "LVMH Moet Hennessy", "Action", "LVMH", "AA", 82, 0.0),
        ("FR0000131104", "BNP Paribas Action", "Action", "BNP Paribas", "A", 78, 0.0),
        ("LU0996182563", "Amundi Euro Liquidity UCITS", "OPCVM", "Amundi", "AAA", 90, 0.2),
        ("EU0009658456", "Swap de Taux Payeur Fixe 5Y", "Derive_Taux", "Clearing_Eurex", "AAA", 80, -4.5)
    ]
    cursor.executemany("INSERT INTO Referentiel_Titres VALUES (?, ?, ?, ?, ?, ?, ?);", titres)

    # 4. Création des Fonds d'Assurance-Vie (1 Fonds en Euros & 1 Unité de Compte FID)
    fonds = [
        ("FE_GARANTI_01", "Fonds Euros Sécurité Lux", "Fonds_Euros", 100_000_000.0),
        ("UC_DYNAM_02", "FID Actions & Multi-Asset", "Unite_Compte", 50_000_000.0)
    ]
    cursor.executemany("INSERT INTO Portefeuilles_AssuranceVie VALUES (?, ?, ?, ?);", fonds)

    # 5. Injection des positions (avec création volontaire de Breaches pour le test)
    positions = [
        # Fonds en Euros (100M€) :
        # - BNP Paribas (Oblig 5M€ + Action 2.5M€ = 7.5M€ -> 7.5% > limite de 5% par émetteur corporate !)
        # - Altice High Yield (12M€ -> 12% > limite High Yield de 10% !)
        ("FE_GARANTI_01", "FR0014007L00", 65_000_000.0),
        ("FE_GARANTI_01", "FR0013250560", 5_000_000.0),
        ("FE_GARANTI_01", "FR0000131104", 2_500_000.0),
        ("FE_GARANTI_01", "XS2010039999", 12_000_000.0),
        ("FE_GARANTI_01", "LU0996182563", 13_500_000.0),
        ("FE_GARANTI_01", "EU0009658456", 2_000_000.0),

        # Unité de Compte (50M€) :
        # - LVMH (4.5M€ -> 9.0% > limite de 5% par émetteur !)
        ("UC_DYNAM_02", "FR0000121014", 4_500_000.0),
        ("UC_DYNAM_02", "LU0996182563", 45_500_000.0)
    ]
    cursor.executemany(
        "INSERT INTO Positions_Inventaire (ID_Fonds, ISIN, Valeur_Marche_EUR) VALUES (?, ?, ?);",
        positions
    )
    conn.commit()

    # 6. REQUETE SQL D'ANALYSE DES RISQUES & CONFORMITE (Dispersion Emetteur Solvabilité II)
    query_concentration = """
    SELECT 
        p.ID_Fonds,
        f.Type_Support,
        r.Emetteur,
        SUM(p.Valeur_Marche_EUR) AS Exposition_EUR,
        ROUND((SUM(p.Valeur_Marche_EUR) / f.AUM_Total) * 100, 2) AS Exposition_Pct,
        5.00 AS Limite_Max_Pct,
        CASE 
            WHEN r.Emetteur IN ('Etat Francais', 'Amundi', 'Clearing_Eurex') THEN 'OK (Exempté Souverain/OPCVM)'
            WHEN (SUM(p.Valeur_Marche_EUR) / f.AUM_Total) * 100 > 5.00 THEN 'BREACH (> 5%)'
            ELSE 'CONFORME'
        END AS Statut_Compliance
    FROM Positions_Inventaire p
    JOIN Referentiel_Titres r ON p.ISIN = r.ISIN
    JOIN Portefeuilles_AssuranceVie f ON p.ID_Fonds = f.ID_Fonds
    GROUP BY p.ID_Fonds, r.Emetteur
    ORDER BY p.ID_Fonds, Exposition_Pct DESC;
    """

    df_concentration = pd.read_sql_query(query_concentration, conn)
    print("=== 1. CONTROLE DE CONCENTRATION EMETTEUR (SOLVABILITE II) ===")
    print(df_concentration.to_string(index=False))

    # 7. REQUETE SQL : RISQUE DE TAUX (DURATION) & QUOTA HIGH YIELD SUR FONDS EUROS
    query_metrics = """
    SELECT 
        p.ID_Fonds,
        f.Type_Support,
        ROUND(SUM(p.Valeur_Marche_EUR * r.Duration_Modifiee) / f.AUM_Total, 2) AS Duration_Portefeuille_Yrs,
        ROUND(SUM(CASE WHEN r.Notation_Credit IN ('BB', 'B') THEN p.Valeur_Marche_EUR ELSE 0 END) / f.AUM_Total * 100, 2) AS Poids_High_Yield_Pct,
        ROUND(SUM(p.Valeur_Marche_EUR * r.Score_ESG) / f.AUM_Total, 1) AS Score_ESG_Moyen
    FROM Positions_Inventaire p
    JOIN Referentiel_Titres r ON p.ISIN = r.ISIN
    JOIN Portefeuilles_AssuranceVie f ON p.ID_Fonds = f.ID_Fonds
    GROUP BY p.ID_Fonds;
    """

    df_metrics = pd.read_sql_query(query_metrics, conn)
    print("\n=== 2. INDICATEURS DE RISQUE DE MARCHE (DURATION, CREDIT HY, ESG) ===")
    print(df_metrics.to_string(index=False))

    conn.close()

if __name__ == "__main__":
    run_life_insurance_compliance_engine()