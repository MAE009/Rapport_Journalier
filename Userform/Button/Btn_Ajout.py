# Ajouter un rapport 
Private Sub Btn_Ajouter_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    
    Set ws = ThisWorkbook.Sheets("RJ")
    Set tbl = ws.ListObjects("suivi_interventio")
    
    Set newRow = tbl.ListRows.Add
    
    With newRow
        .Range(1, 1).Value = Me.TxtNomBien.Value
        .Range(1, 2).Value = Me.TxtLocalisation.Value
        .Range(1, 3).Value = Me.CboEtat.Value
        .Range(1, 4).Value = Me.TxtDefaillance.Value
        .Range(1, 5).Value = Me.TxtCause.Value
        .Range(1, 6).Value = Me.TxtAction.Value
        .Range(1, 7).Value = Me.TxtDateDefaillance.Value
        .Range(1, 8).Value = Me.TxtDateActuel.Value
        .Range(1, 9).Value = Me.TxtTemps.Value
        .Range(1, 10).Value = Me.CboResultat.Value
        .Range(1, 11).Value = Me.TxtIntervenant.Value
        .Range(1, 12).Value = Me.CboObservation.Value
    End With
    
    Call ChargerListBox
    MsgBox "Rapport ajouté avec succès.", vbInformation
End Sub
