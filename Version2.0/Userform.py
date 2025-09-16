' UserForm Ajout_rapport
' Contrôles nécessaires :
' - TextBox : NomBien, Localisation, DefaillanceConst, CauseDefaillance, ActionMenee, Intervenant
' - ComboBox : EtatActuel, Resultat, Observation, FiltreAnnee, FiltreMois, FiltreJour
' - ListBox : ListeRapports
' - Labels pour chaque champ
' - CommandButton : btnAjouter, btnModifier, btnSupprimer, btnFiltrer, btnEnvoyer

Option Explicit
Dim modifMode As Boolean
Dim selectedID As Long

Private Sub UserForm_Initialize()
    ' Initialisation des ComboBox
    InitialiserCombobox
    
    ' Configuration de la date actuelle
    DateActuelle.Value = Format(Now, "dd/mm/yyyy")
    
    ' Charger les rapports existants
    ChargerRapports
    
    ' Configuration des filtres de date
    ConfigurerFiltresDate
    
    modifMode = False
    btnModifier.Enabled = False
    btnSupprimer.Enabled = False
End Sub

Private Sub ConfigurerFiltresDate()
    ' Remplir les années (par exemple, de 2020 à l'année en cours)
    Dim annee As Integer
    FiltreAnnee.Clear
    FiltreAnnee.AddItem "Toutes"
    For annee = 2020 To Year(Now)
        FiltreAnnee.AddItem annee
    Next annee
    FiltreAnnee.Value = "Toutes"
    
    ' Remplir les mois
    FiltreMois.Clear
    FiltreMois.AddItem "Tous"
    Dim i As Integer
    For i = 1 To 12
        FiltreMois.AddItem MonthName(i)
    Next i
    FiltreMois.Value = "Tous"
    
    ' Remplir les jours
    FiltreJour.Clear
    FiltreJour.AddItem "Tous"
    For i = 1 To 31
        FiltreJour.AddItem i
    Next i
    FiltreJour.Value = "Tous"
End Sub

Private Sub InitialiserCombobox()
    ' Valeurs pour l'état actuel
    With EtatActuel
        .AddItem "Bon état"
        .AddItem "État moyen"
        .AddItem "Mauvais état"
        .AddItem "Hors service"
    End With
    
    ' Valeurs pour le résultat
    With Resultat
        .AddItem "Résolu"
        .AddItem "En cours"
        .AddItem "En attente de pièces"
        .AddItem "Reporté"
        .AddItem "Non résolu"
    End With
    
    ' Valeurs pour les observations
    With Observation
        .AddItem "Aucune"
        .AddItem "Contrôle nécessaire"
        .AddItem "Maintenance préventive requise"
        .AddItem "À remplacer"
        .AddItem "Urgent"
    End With
End Sub

Private Sub ChargerRapports()
    ' Charger les rapports depuis la feuille "suivi_intervention"
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("suivi_intervention")
    
    ' Vider la ListBox
    ListeRapports.Clear
    
    ' Configurer les colonnes de la ListBox (ajuster selon vos colonnes)
    With ListeRapports
        .ColumnCount = 12 ' Ajuster selon le nombre de colonnes
        .ColumnWidths = "40;80;60;80;100;100;100;80;80;50;80;100" ' Ajuster les largeurs
    End With
    
    ' Remplir la ListBox
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim rowData As Variant
        rowData = tbl.ListRows(i).Range.Value
        
        ' Ajouter la ligne à la ListBox
        ListeRapports.AddItem rowData(1, 1) ' ID
        For j = 2 To UBound(rowData, 2)
            ListeRapports.List(ListeRapports.ListCount - 1, j - 1) = rowData(1, j)
        Next j
    Next i
End Sub

Private Sub AppliquerFiltre()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("suivi_intervention")
    
    ' Vider la ListBox
    ListeRapports.Clear
    
    ' Appliquer le filtre selon les sélections
    Dim filtreAnnee As String
    Dim filtreMois As String
    Dim filtreJour As String
    
    filtreAnnee = FiltreAnnee.Value
    filtreMois = FiltreMois.Value
    filtreJour = FiltreJour.Value
    
    ' Parcourir toutes les lignes et appliquer le filtre
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        Dim dateDef As Date
        dateDef = tbl.ListRows(i).Range.Cells(1, 7).Value ' Colonne date défaillance
        
        Dim ajouterLigne As Boolean
        ajouterLigne = True
        
        ' Filtrer par année
        If filtreAnnee <> "Toutes" And Year(dateDef) <> CInt(filtreAnnee) Then
            ajouterLigne = False
        End If
        
        ' Filtrer par mois
        If filtreMois <> "Tous" And MonthName(Month(dateDef)) <> filtreMois Then
            ajouterLigne = False
        End If
        
        ' Filtrer par jour
        If filtreJour <> "Tous" And Day(dateDef) <> CInt(filtreJour) Then
            ajouterLigne = False
        End If
        
        If ajouterLigne Then
            Dim rowData As Variant
            rowData = tbl.ListRows(i).Range.Value
            
            ListeRapports.AddItem rowData(1, 1) ' ID
            For j = 2 To UBound(rowData, 2)
                ListeRapports.List(ListeRapports.ListCount - 1, j - 1) = rowData(1, j)
            Next j
        End If
    Next i
End Sub

Private Sub btnFiltrer_Click()
    AppliquerFiltre
End Sub

Private Sub ListeRapports_Click()
    If ListeRapports.ListIndex <> -1 Then
        ' Activer les boutons modification et suppression
        btnModifier.Enabled = True
        btnSupprimer.Enabled = True
        
        ' Stocker l'ID sélectionné
        selectedID = ListeRapports.Value
    End If
End Sub

Private Sub btnAjouter_Click()
    If ValiderFormulaire Then
        AjouterRapport
        ReinitialiserFormulaire
        ChargerRapports
        MsgBox "Rapport ajouté avec succès!", vbInformation
    End If
End Sub

Private Function ValiderFormulaire() As Boolean
    ' Validation des champs obligatoires
    If NomBien.Value = "" Then
        MsgBox "Veuillez saisir le nom du bien.", vbExclamation
        NomBien.SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    If Localisation.Value = "" Then
        MsgBox "Veuillez saisir la localisation.", vbExclamation
        Localisation.SetFocus
        ValiderFormulaire = False
        Exit Function
    End If
    
    ' Ajouter d'autres validations si nécessaire
    
    ValiderFormulaire = True
End Function

Private Sub AjouterRapport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("suivi_intervention")
    
    ' Ajouter une nouvelle ligne
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    ' Remplir la nouvelle ligne
    With newRow
        .Range(1, 1) = TrouverNouvelID() ' ID
        .Range(1, 2) = NomBien.Value
        .Range(1, 3) = Localisation.Value
        .Range(1, 4) = EtatActuel.Value
        .Range(1, 5) = DefaillanceConst.Value
        .Range(1, 6) = CauseDefaillance.Value
        .Range(1, 7) = ActionMenee.Value
        .Range(1, 8) = CDate(DateDefaillance.Value)
        .Range(1, 9) = CDate(DateActuelle.Value)
        .Range(1, 10) = Temps.Value
        .Range(1, 11) = Resultat.Value
        .Range(1, 12) = Intervenant.Value
        .Range(1, 13) = Observation.Value
    End With
End Sub

Private Function TrouverNouvelID() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("suivi_intervention")
    
    If tbl.ListRows.Count = 0 Then
        TrouverNouvelID = 1
    Else
        TrouverNouvelID = Application.WorksheetFunction.Max(tbl.ListColumns(1).DataBodyRange) + 1
    End If
End Function

Private Sub btnModifier_Click()
    If ListeRapports.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un rapport à modifier.", vbExclamation
        Exit Sub
    End If
    
    If ValiderFormulaire Then
        ModifierRapport
        ReinitialiserFormulaire
        ChargerRapports
        MsgBox "Rapport modifié avec succès!", vbInformation
    End If
End Sub

Private Sub ModifierRapport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("suivi_intervention")
    
    ' Trouver la ligne correspondant à l'ID sélectionné
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range(1, 1).Value = selectedID Then
            ' Modifier la ligne
            With tbl.ListRows(i)
                .Range(1, 2) = NomBien.Value
                .Range(1, 3) = Localisation.Value
                .Range(1, 4) = EtatActuel.Value
                .Range(1, 5) = DefaillanceConst.Value
                .Range(1, 6) = CauseDefaillance.Value
                .Range(1, 7) = ActionMenee.Value
                .Range(1, 8) = CDate(DateDefaillance.Value)
                .Range(1, 9) = CDate(DateActuelle.Value)
                .Range(1, 10) = Temps.Value
                .Range(1, 11) = Resultat.Value
                .Range(1, 12) = Intervenant.Value
                .Range(1, 13) = Observation.Value
            End With
            Exit For
        End If
    Next i
End Sub

Private Sub btnSupprimer_Click()
    If ListeRapports.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un rapport à supprimer.", vbExclamation
        Exit Sub
    End If
    
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Êtes-vous sûr de vouloir supprimer ce rapport?", vbYesNo + vbQuestion, "Confirmation de suppression")
    
    If reponse = vbYes Then
        SupprimerRapport
        ReinitialiserFormulaire
        ChargerRapports
        MsgBox "Rapport supprimé avec succès!", vbInformation
    End If
End Sub

Private Sub SupprimerRapport()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("suivi_intervention")
    
    ' Trouver la ligne correspondant à l'ID sélectionné
    Dim i As Long
    For i = 1 To tbl.ListRows.Count
        If tbl.ListRows(i).Range(1, 1).Value = selectedID Then
            tbl.ListRows(i).Delete
            Exit For
        End If
    Next i
End Sub

Private Sub btnEnvoyer_Click()
    ' Vérifier l'heure actuelle
    If Time() < #5:30:00 PM# Then
        MsgBox "L'envoi du rapport n'est autorisé qu'après 17h30.", vbExclamation
        Exit Sub
    End If
    
    Dim reponse As VbMsgBoxResult
    reponse = MsgBox("Êtes-vous sûr de vouloir envoyer le rapport journalier?", vbYesNo + vbQuestion, "Confirmation d'envoi")
    
    If reponse = vbYes Then
        ' Code pour envoyer le rapport (à adapter selon votre méthode d'envoi)
        EnvoyerRapport
        MsgBox "Rapport envoyé avec succès!", vbInformation
    End If
End Sub

Private Sub EnvoyerRapport()
    ' Cette procédure doit être adaptée selon votre méthode d'envoi
    ' (email, impression, sauvegarde, etc.)
    
    ' Exemple: créer une copie du rapport pour envoi
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RJ")
    
    ' Créer une nouvelle feuille pour l'envoi
    Dim newSheet As Worksheet
    On Error Resume Next
    Set newSheet = ThisWorkbook.Sheets("Rapport_" & Format(Date, "ddmmyy"))
    On Error GoTo 0
    
    If newSheet Is Nothing Then
        Set newSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newSheet.Name = "Rapport_" & Format(Date, "ddmmyy")
    End If
    
    ' Copier les données filtrées
    ws.ListObjects("suivi_intervention").Range.Copy
    newSheet.Range("A1").PasteSpecial xlPasteAll
    
    ' Formater la nouvelle feuille
    Application.CutCopyMode = False
    
    MsgBox "Rapport préparé pour envoi dans la feuille '" & newSheet.Name & "'", vbInformation
End Sub

Private Sub ReinitialiserFormulaire()
    NomBien.Value = ""
    Localisation.Value = ""
    EtatActuel.Value = ""
    DefaillanceConst.Value = ""
    CauseDefaillance.Value = ""
    ActionMenee.Value = ""
    DateDefaillance.Value = ""
    DateActuelle.Value = Format(Now, "dd/mm/yyyy")
    Temps.Value = ""
    Resultat.Value = ""
    Intervenant.Value = ""
    Observation.Value = ""
    
    modifMode = False
    btnModifier.Enabled = False
    btnSupprimer.Enabled = False
    btnAjouter.Caption = "Ajouter"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Sauvegarder avant de fermer
    ThisWorkbook.Save
End Sub
