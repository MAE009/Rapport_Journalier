Private Sub UserForm_Initialize()
    ' Date actuelle auto
    Me.TB_DateActuel.Value = Format(Date, "dd/mm/yyyy")

    ' Remplir la ListBox
    Call Charger_ListBox

    ' Remplir les filtres
    Call Charger_Filtres
End Sub


Private Sub ChargerRapports(Optional ByVal filtreAnnee As String = "", _
                            Optional ByVal filtreMois As String = "", _
                            Optional ByVal filtreJour As String = "")
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Set ws = ThisWorkbook.Sheets("RJ")
    Set tbl = ws.ListObjects("suivi_interventio")
    
    Me.LstRapports.Clear
    
    ' Charger les titres dans la ListBox (optionnel si MultiColumn)
    Me.LstRapports.ColumnCount = tbl.ListColumns.Count
    
    For i = 1 To tbl.ListRows.Count
        Dim ligne() As Variant
        ligne = tbl.ListRows(i).Range.Value
        
        ' Filtrage par date de défaillance
        Dim dateDef As Date
        dateDef = ligne(1, 7) ' 7ème colonne = Date de défaillance
        
        If (filtreAnnee = "" Or Year(dateDef) = filtreAnnee) And _
           (filtreMois = "" Or Month(dateDef) = filtreMois) And _
           (filtreJour = "" Or Day(dateDef) = filtreJour) Then
           
            Me.LstRapports.AddItem
            Dim j As Long
            For j = 1 To UBound(ligne, 2)
                Me.LstRapports.List(Me.LstRapports.ListCount - 1, j - 1) = ligne(1, j)
            Next j
        End If
    Next i
End Sub



Private Sub ChargerFiltres()
    Dim ws As Worksheet, tbl As ListObject, i As Long
    Dim dateDef As Date
    Dim dictAnnee As Object, dictMois As Object, dictJour As Object
    
    Set ws = ThisWorkbook.Sheets("RJ")
    Set tbl = ws.ListObjects("suivi_interventio")
    
    ' Création des dictionnaires pour éviter les doublons
    Set dictAnnee = CreateObject("Scripting.Dictionary")
    Set dictMois = CreateObject("Scripting.Dictionary")
    Set dictJour = CreateObject("Scripting.Dictionary")
    
    ' Balayer toutes les lignes du tableau
    For i = 1 To tbl.ListRows.Count
        On Error Resume Next
        dateDef = tbl.ListRows(i).Range.Cells(1, 7).Value '⚠️ 7ème colonne = "Date de défaillance"
        On Error GoTo 0
        
        If IsDate(dateDef) Then
            If Not dictAnnee.exists(Year(dateDef)) Then dictAnnee.Add Year(dateDef), Year(dateDef)
            If Not dictMois.exists(Month(dateDef)) Then dictMois.Add Month(dateDef), Month(dateDef)
            If Not dictJour.exists(Day(dateDef)) Then dictJour.Add Day(dateDef), Day(dateDef)
        End If
    Next i
    
    ' Remplir les ComboBox
    Me.CboAnnee.Clear
    Me.CboMois.Clear
    Me.CboJour.Clear
    
    If dictAnnee.Count > 0 Then Me.CboAnnee.List = dictAnnee.Items
    If dictMois.Count > 0 Then Me.CboMois.List = dictMois.Items
    If dictJour.Count > 0 Then Me.CboJour.List = dictJour.Items
End Sub
