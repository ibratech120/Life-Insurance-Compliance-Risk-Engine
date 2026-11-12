-- =========================================================================
-- 1. CREATION DE L'ARCHITECTURE DES TABLES (Compatible MS Access / SQL)
-- =========================================================================

CREATE TABLE Referentiel_Titres (
    ISIN VARCHAR(12) PRIMARY KEY,
    Nom_Titre VARCHAR(100),
    Classe_Actif VARCHAR(30),     -- Obligation, Action, OPCVM, Derive_Taux
    Emetteur VARCHAR(50),
    Notation_Credit VARCHAR(10),  -- AAA, AA, A, BBB, BB, B (High Yield < BBB)
    Score_ESG INT,                -- Score sur 100
    Duration_Modifiee DECIMAL(5,2)
);

CREATE TABLE Portefeuilles_AssuranceVie (
    ID_Fonds VARCHAR(20) PRIMARY KEY,
    Nom_Fonds VARCHAR(50),
    Type_Support VARCHAR(20),     -- 'Fonds_Euros' (Capital Garanti) ou 'Unite_Compte' (UC/FID)
    AUM_Total DECIMAL(18,2)       -- Actif Net sous gestion
);

CREATE TABLE Positions_Inventaire (
    ID_Position INT PRIMARY KEY,
    Date_Arrete DATE,
    ID_Fonds VARCHAR(20),
    ISIN VARCHAR(12),
    Quantite DECIMAL(18,4),
    Valeur_Marche_EUR DECIMAL(18,2),
    FOREIGN KEY (ID_Fonds) REFERENCES Portefeuilles_AssuranceVie(ID_Fonds),
    FOREIGN KEY (ISIN) REFERENCES Referentiel_Titres(ISIN)
);

CREATE TABLE Historique_Breaches (
    ID_Breach INT PRIMARY KEY,
    Date_Controle DATE,
    ID_Fonds VARCHAR(20),
    Regle_Violee VARCHAR(100),
    Exposition_Actuelle_Pct DECIMAL(6,2),
    Limite_Reglementaire_Pct DECIMAL(6,2),
    Statut VARCHAR(20)
);

-- =========================================================================
-- 2. REQUETES SQL COMPLEXES DE CONTROLE DE CONFORMITE (PRE/POST-TRADE)
-- =========================================================================

-- CONTROLE 1 : Règle de dispersion Solvabilité II (Max 5% d'exposition sur un même émetteur)
SELECT 
    p.ID_Fonds,
    f.Nom_Fonds,
    r.Emetteur,
    SUM(p.Valeur_Marche_EUR) AS Exposition_EUR,
    ROUND((SUM(p.Valeur_Marche_EUR) / f.AUM_Total) * 100, 2) AS Poids_Pct,
    5.00 AS Limite_Max_Pct,
    'BREACH - Concentration Emetteur (> 5%)' AS Alerte_Conformite
FROM Positions_Inventaire p
INNER JOIN Referentiel_Titres r ON p.ISIN = r.ISIN
INNER JOIN Portefeuilles_AssuranceVie f ON p.ID_Fonds = f.ID_Fonds
GROUP BY p.ID_Fonds, f.Nom_Fonds, r.Emetteur, f.AUM_Total
HAVING (SUM(p.Valeur_Marche_EUR) / f.AUM_Total) * 100 > 5.00;

-- CONTROLE 2 : Fonds en Euros - Minimum 75% en Obligations (Adossement Actif-Passif)
SELECT 
    p.ID_Fonds,
    f.Nom_Fonds,
    ROUND((SUM(CASE WHEN r.Classe_Actif = 'Obligation' THEN p.Valeur_Marche_EUR ELSE 0 END) / f.AUM_Total) * 100, 2) AS Poids_Obligations_Pct,
    75.00 AS Seuil_Min_Pct,
    'BREACH - Quota Obligataire Fonds Euros insuffisant (< 75%)' AS Alerte_Conformite
FROM Positions_Inventaire p
INNER JOIN Referentiel_Titres r ON p.ISIN = r.ISIN
INNER JOIN Portefeuilles_AssuranceVie f ON p.ID_Fonds = f.ID_Fonds
WHERE f.Type_Support = 'Fonds_Euros'
GROUP BY p.ID_Fonds, f.Nom_Fonds, f.AUM_Total
HAVING (SUM(CASE WHEN r.Classe_Actif = 'Obligation' THEN p.Valeur_Marche_EUR ELSE 0 END) / f.AUM_Total) * 100 < 75.00;