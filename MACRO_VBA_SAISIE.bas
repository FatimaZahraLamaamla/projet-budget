' ============================================================
' MODULE VBA - Formulaire Saisie Engagements SNRT 2026
' A coller dans : Alt+F11 > Insertion > Module
' ============================================================

Sub AjouterLigne()

    Dim ws_saisie  As Worksheet
    Dim ws_eng     As Worksheet
    Dim next_row   As Long
    Dim statut     As String
    Dim statut_sup As String
    Dim nature     As String

    Set ws_saisie = ThisWorkbook.Sheets("SAISIE")
    Set ws_eng    = ThisWorkbook.Sheets("ENGAGEMENTS 2026")

    ' ── Validation champs obligatoires ──────────────────────────────────────
    Dim erreurs As String
    erreurs = ""
    If Trim(ws_saisie.Range("B5").Value) = ""  Then erreurs = erreurs & Chr(10) & "- Année lancement"
    If Trim(ws_saisie.Range("B6").Value) = ""  Then erreurs = erreurs & Chr(10) & "- Direction"
    If Trim(ws_saisie.Range("B8").Value) = ""  Then erreurs = erreurs & Chr(10) & "- Type d'achat"
    If Trim(ws_saisie.Range("B11").Value) = "" Then erreurs = erreurs & Chr(10) & "- Objet du marché"
    If Trim(ws_saisie.Range("B14").Value) = "" Then erreurs = erreurs & Chr(10) & "- Nature (Fonct./Invest.)"
    If Trim(ws_saisie.Range("B18").Value) = "" Then erreurs = erreurs & Chr(10) & "- Statut d'engagement"

    If erreurs <> "" Then
        MsgBox "Champs obligatoires manquants :" & Chr(10) & erreurs, vbExclamation, "Validation"
        Exit Sub
    End If

    ' ── Lecture des valeurs ──────────────────────────────────────────────────
    statut     = UCase(Trim(ws_saisie.Range("B18").Value))
    statut_sup = UCase(Trim(ws_saisie.Range("B20").Value))
    nature     = UCase(Trim(ws_saisie.Range("B14").Value))

    ' ── Validation logique montant ───────────────────────────────────────────
    Dim montant_n  As Variant
    Dim montant_ad As Variant
    montant_n  = ws_saisie.Range("E8").Value
    montant_ad = ws_saisie.Range("E11").Value

    If statut = "EN COURS DE PREENGAGEMENT" Then
        If IsEmpty(montant_n) Or montant_n = "" Or montant_n = 0 Then
            If MsgBox("Le Montant Préengagé Initial est vide." & Chr(10) & "Continuer quand même ?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    ElseIf statut = "PREENGAGE" Then
        If InStr(statut_sup, "ATTRIBU") > 0 Then
            If IsEmpty(montant_ad) Or montant_ad = "" Or montant_ad = 0 Then
                If MsgBox("Le Total Marché TTC est vide." & Chr(10) & "Continuer quand même ?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
        Else
            If IsEmpty(montant_n) Or montant_n = "" Or montant_n = 0 Then
                If MsgBox("Le Montant Préengagé Initial est vide." & Chr(10) & "Continuer quand même ?", vbQuestion + vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    ElseIf statut = "EN COURS D'ENGAGEMENT" Or statut = "ENGAGE" Then
        If IsEmpty(montant_ad) Or montant_ad = "" Or montant_ad = 0 Then
            If MsgBox("Le Total Marché TTC est vide." & Chr(10) & "Continuer quand même ?", vbQuestion + vbYesNo) = vbNo Then
                Exit Sub
            End If
        End If
    End If

    ' ── Trouver la première ligne vide dans ENGAGEMENTS 2026 ─────────────────
    next_row = ws_eng.Cells(ws_eng.Rows.Count, 2).End(xlUp).Row + 1

    ' ── Écriture dans le fichier (correspondance colonnes) ──────────────────
    ' A=1  Année lancement
    ws_eng.Cells(next_row, 1).Value  = ws_saisie.Range("B5").Value
    ' B=2  Direction
    ws_eng.Cells(next_row, 2).Value  = ws_saisie.Range("B6").Value
    ' C=3  Référence
    ws_eng.Cells(next_row, 3).Value  = ws_saisie.Range("B7").Value
    ' D=4  Type d'achat (BCI)
    ws_eng.Cells(next_row, 4).Value  = ws_saisie.Range("B8").Value
    ' E=5  Statut 1
    ws_eng.Cells(next_row, 5).Value  = ws_saisie.Range("B19").Value
    ' F=6  Statut support d'engagement
    ws_eng.Cells(next_row, 6).Value  = ws_saisie.Range("B20").Value
    ' G=7  N° support d'engagement
    ws_eng.Cells(next_row, 7).Value  = ws_saisie.Range("B21").Value
    ' H=8  Catégorie
    ws_eng.Cells(next_row, 8).Value  = ws_saisie.Range("B9").Value
    ' I=9  Lots
    ws_eng.Cells(next_row, 9).Value  = ws_saisie.Range("B10").Value
    ' J=10 Objet
    ws_eng.Cells(next_row, 10).Value = ws_saisie.Range("B11").Value
    ' K=11 Type budget
    ws_eng.Cells(next_row, 11).Value = ws_saisie.Range("B15").Value
    ' L=12 Nature (FONCTIONNEMENT/INVESTISSEMENT)
    ws_eng.Cells(next_row, 12).Value = ws_saisie.Range("B14").Value
    ' M=13 Ligne budgétaire
    ws_eng.Cells(next_row, 13).Value = ws_saisie.Range("B16").Value
    ' N=14 Montant préengagé initial
    If Not IsEmpty(montant_n) And montant_n <> "" Then
        ws_eng.Cells(next_row, 14).Value = montant_n
    End If
    ' O=15 Estimation TTC
    If Not IsEmpty(ws_saisie.Range("E9").Value) And ws_saisie.Range("E9").Value <> "" Then
        ws_eng.Cells(next_row, 15).Value = ws_saisie.Range("E9").Value
    End If
    ' P=16 Estimation HT
    If Not IsEmpty(ws_saisie.Range("E10").Value) And ws_saisie.Range("E10").Value <> "" Then
        ws_eng.Cells(next_row, 16).Value = ws_saisie.Range("E10").Value
    End If
    ' Q=17 Nature besoin
    ws_eng.Cells(next_row, 17).Value = ws_saisie.Range("B12").Value
    ' T=20 Adjudicataire
    ws_eng.Cells(next_row, 20).Value = ws_saisie.Range("E13").Value
    ' U=21 Année approbation
    ws_eng.Cells(next_row, 21).Value = ws_saisie.Range("E14").Value
    ' V=22 Date de la fiche
    If Trim(ws_saisie.Range("E15").Value) <> "" Then
        ws_eng.Cells(next_row, 22).Value = CDate(ws_saisie.Range("E15").Value)
        ws_eng.Cells(next_row, 22).NumberFormat = "DD/MM/YYYY"
    End If
    ' W=23 Date visa
    If Trim(ws_saisie.Range("E16").Value) <> "" Then
        ws_eng.Cells(next_row, 23).Value = CDate(ws_saisie.Range("E16").Value)
        ws_eng.Cells(next_row, 23).NumberFormat = "DD/MM/YYYY"
    End If
    ' X=24 Date approbation
    If Trim(ws_saisie.Range("E17").Value) <> "" Then
        ws_eng.Cells(next_row, 24).Value = CDate(ws_saisie.Range("E17").Value)
        ws_eng.Cells(next_row, 24).NumberFormat = "DD/MM/YYYY"
    End If
    ' AD=30 Total Marché TTC
    If Not IsEmpty(montant_ad) And montant_ad <> "" Then
        ws_eng.Cells(next_row, 30).Value = montant_ad
    End If
    ' AE=31 Statut d'engagement
    ws_eng.Cells(next_row, 31).Value = ws_saisie.Range("B18").Value
    ' AM=39 Remarque
    ws_eng.Cells(next_row, 39).Value = ws_saisie.Range("E18").Value

    ' ── Confirmer et effacer ─────────────────────────────────────────────────
    MsgBox "Ligne " & next_row & " ajoutée avec succès !" & Chr(10) & Chr(10) & _
           "N'oubliez pas de lancer le script Actualiser pour mettre à jour le tableau de bord.", _
           vbInformation, "Ajout réussi"

    Call EffacerFormulaire

End Sub

' ─────────────────────────────────────────────────────────────────────────────
Sub EffacerFormulaire()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SAISIE")

    ' Effacer toutes les cellules de saisie
    Dim cellules_saisie As Variant
    cellules_saisie = Array("B5","B6","B7","B8","B9","B10","B11","B12", _
                            "B14","B15","B16","B18","B19","B20","B21", _
                            "E8","E9","E10","E11","E13","E14","E15","E16","E17","E18")
    Dim i As Integer
    For i = 0 To UBound(cellules_saisie)
        ws.Range(cellules_saisie(i)).ClearContents
    Next i
End Sub

' ─────────────────────────────────────────────────────────────────────────────
Sub ActualiserListeLignesBudg()
    ' Met à jour la liste déroulante Ligne Budgétaire selon la Nature choisie
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("SAISIE")

    Dim nature As String
    nature = UCase(Trim(ws.Range("B14").Value))

    Dim rng As Range
    Set rng = ws.Range("B16")

    ' Supprimer les anciennes validations sur B16
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0

    Dim formula As String
    If nature = "FONCTIONNEMENT" Then
        formula = "=LIGNES_FONCT"
    ElseIf nature = "INVESTISSEMENT" Then
        formula = "=LIGNES_INVEST"
    Else
        Exit Sub
    End If

    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=formula
    rng.Validation.ShowDropDown = True
    rng.ClearContents
End Sub
