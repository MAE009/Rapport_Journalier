Private Sub Btn_Modifier_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim ligne As Range
    Dim idx As Long
    Dim rowTable As Long
    
    If Me.ListBox1.ListIndex = -1 Then
        MsgBox "Veuillez sélectionner un rapport à modifier.", vbExclamation
        Exit Sub
    End If
    
    Set ws = ThisWorkbook.Sheets("RJ")
    Set tbl = ws.ListObjects("suivi_interventio")
    
    ' ⚠️ On récupère le numéro de ligne du tableau stocké dans la colonne cachée de la ListBox
    rowTable = CLng(Me.ListBox1.List(Me.ListBox1.ListIndex, 0))  ' supposons que la col(0) contient l'index
    
    Set ligne = tbl.DataBodyRange.Rows(rowTable)
    
    With ligne
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
    
    Call ChargerRapports  ' recharge la liste
    MsgBox "Rapport modifié avec succès.", vbInformation
End Sub
