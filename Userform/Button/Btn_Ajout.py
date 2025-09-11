Private Sub BtnAjouter_Click()
    Dim ws As Worksheet, tbl As ListObject
    Set ws = ThisWorkbook.Sheets("RJ")
    Set tbl = ws.ListObjects("suivi_interventio")
    
    ' Ajout d'une nouvelle ligne
    Dim newRow As ListRow
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, 1).Value = Me.TxtNomBien.Value
        .Cells(1, 2).Value = Me.TxtLocalisation.Value
        .Cells(1, 3).Value = Me.CboEtat.Value
        .Cells(1, 4).Value = Me.TxtDefaillance.Value
        .Cells(1, 5).Value = Me.TxtCause.Value
        .Cells(1, 6).Value = Me.TxtAction.Value
        .Cells(1, 7).Value = Me.TxtDateDefaillance.Value
        .Cells(1, 8).Value = Me.TxtDateActuel.Value
        .Cells(1, 9).Value = Me.TxtTemps.Value
        .Cells(1, 10).Value = Me.CboResultat.Value
        .Cells(1, 11).Value = Me.TxtIntervenant.Value
        .Cells(1, 12).Value = Me.CboObservation.Value
    End With
    
    Call ChargerRapports
    MsgBox "Rapport ajouté avec succès !", vbInformation
End Sub
