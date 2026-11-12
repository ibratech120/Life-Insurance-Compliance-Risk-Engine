Option Explicit

' ========================================================================================
' MODULE VBA : MOTEUR DE CONFORMITE & CONTROLE DES INVESTISSEMENTS (MS ACCESS / EXCEL)
' Objectif : Réconciliation de flux, contrôle des limites Solvabilité II et log des Breaches
' ========================================================================================

Public Sub Run_Daily_Compliance_Checks()
    On Error GoTo ErrorHandler
    
    Dim db As Object ' DAO.Database
    Dim rs As Object ' DAO.Recordset
    Dim strSQL As String
    Dim breachCount As Integer
    
    Set db = CurrentDb() ' Fonctionne nativement dans MS Access
    breachCount = 0
    
    ' 1. DATA MANAGEMENT : Nettoyage des codes ISIN et suppression des doublons d'import
    db.Execute "UPDATE Positions_Inventaire SET ISIN = TRIM(UCase(ISIN));"
    
    ' Purge des anciens contrôles de la journée pour éviter les doublons
    db.Execute "DELETE FROM Historique_Breaches WHERE Date_Controle = Date();"
    
    ' 2. CONTROLE 1 : Limite de concentration par émetteur (> 5% de l'AUM)
    strSQL = "SELECT p.ID_Fonds, r.Emetteur, " & _
             "ROUND((SUM(p.Valeur_Marche_EUR) / FIRST(f.AUM_Total)) * 100, 2) AS Exp_Pct " & _
             "FROM (Positions_Inventaire p " & _
             "INNER JOIN Referentiel_Titres r ON p.ISIN = r.ISIN) " & _
             "INNER JOIN Portefeuilles_AssuranceVie f ON p.ID_Fonds = f.ID_Fonds " & _
             "GROUP BY p.ID_Fonds, r.Emetteur " & _
             "HAVING (SUM(p.Valeur_Marche_EUR) / FIRST(f.AUM_Total)) * 100 > 5.0;"
             
    Set rs = db.OpenRecordset(strSQL)
    
    Do While Not rs.EOF
        breachCount = breachCount + 1
        Call Log_Breach(db, rs!ID_Fonds, "Concentration Emetteur > 5% (" & rs!Emetteur & ")", rs!Exp_Pct, 5#)
        rs.MoveNext
    Loop
    rs.Close
    
    ' 3. CONTROLE 2 : Limite High Yield (Titres spéculatifs BB/B <= 10% sur Fonds en Euros)
    strSQL = "SELECT p.ID_Fonds, " & _
             "ROUND((SUM(p.Valeur_Marche_EUR) / FIRST(f.AUM_Total)) * 100, 2) AS HY_Pct " & _
             "FROM (Positions_Inventaire p " & _
             "INNER JOIN Referentiel_Titres r ON p.ISIN = r.ISIN) " & _
             "INNER JOIN Portefeuilles_AssuranceVie f ON p.ID_Fonds = f.ID_Fonds " & _
             "WHERE f.Type_Support = 'Fonds_Euros' AND r.Notation_Credit IN ('BB', 'B', 'CCC') " & _
             "GROUP BY p.ID_Fonds " & _
             "HAVING (SUM(p.Valeur_Marche_EUR) / FIRST(f.AUM_Total)) * 100 > 10.0;"
             
    Set rs = db.OpenRecordset(strSQL)
    
    Do While Not rs.EOF
        breachCount = breachCount + 1
        Call Log_Breach(db, rs!ID_Fonds, "Limite High Yield Fonds Euros > 10%", rs!HY_Pct, 10#)
        rs.MoveNext
    Loop
    rs.Close
    
    ' 4. NOTIFICATION UTILISATEUR (Interface de pilotage)
    If breachCount > 0 Then
        MsgBox "ALERTE CONFORMITE : " & breachCount & " dépassement(s) détecté(s) !" & vbCrLf & _
               "Consultez le tableau de bord des Breaches.", vbExclamation, "Risk & Compliance Engine"
    Else
        MsgBox "Contrôles pré/post-trade terminés : 100% des fonds sont conformes.", vbInformation, "Risk & Compliance Engine"
    End If

    Set rs = Nothing
    Set db = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Erreur critique dans le moteur de contrôle VBA : " & Err.Description, vbCritical, "Erreur Exécution"
End Sub

' Fonction d'insertion dans la table d'audit
Private Sub Log_Breach(db As Object, fondsID As String, regle As String, expActuelle As Double, limite As Double)
    Dim insertSQL As String
    insertSQL = "INSERT INTO Historique_Breaches (Date_Controle, ID_Fonds, Regle_Violee, Exposition_Actuelle_Pct, Limite_Reglementaire_Pct, Statut) " & _
                "VALUES (Date(), '" & fondsID & "', '" & regle & "', " & Replace(CStr(expActuelle), ",", ".") & ", " & Replace(CStr(limite), ",", ".") & ", 'OPEN_BREACH');"
    db.Execute insertSQL
End Sub