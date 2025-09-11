# charger le listbox au démarrage 
Private Sub UserForm_Initialize()
    Call ChargerListBox
End Sub

Private Sub ChargerListBox(Optional ByVal FiltreDate As String = "")
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long, ligne As ListRow
    
    Set ws = ThisWorkbook.Sheets("RJ")
    Set tbl = ws.ListObjects("suivi_interventio")
    
    Me.ListBox1.Clear
    ' Ajouter les titres en en-tête de la ListBox
    Me.ListBox1.ColumnCount = tbl.ListColumns.Count
    Me.ListBox1.ColumnHeads = False
    
    For Each ligne In tbl.ListRows
        If FiltreDate = "" Or Format(ligne.Range(1, 7).Value, "mm/yyyy") = FiltreDate Then
            Me.ListBox1.AddItem
            For i = 1 To tbl.ListColumns.Count
                Me.ListBox1.List(Me.ListBox1.ListCount - 1, i - 1) = ligne.Range(1, i).Value
            Next i
        End If
    Next ligne
End Sub
