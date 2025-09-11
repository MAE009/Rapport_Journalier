Private Sub ListBox1_Click()
    Dim i As Integer
    If Me.ListBox1.ListIndex = -1 Then Exit Sub
    
    Me.TxtNomBien.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 0)
    Me.TxtLocalisation.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 1)
    Me.CboEtat.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 2)
    Me.TxtDefaillance.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 3)
    Me.TxtCause.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 4)
    Me.TxtAction.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 5)
    Me.TxtDateDefaillance.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 6)
    Me.TxtDateActuel.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 7)
    Me.TxtTemps.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 8)
    Me.CboResultat.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 9)
    Me.TxtIntervenant.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 10)
    Me.CboObservation.Value = Me.ListBox1.List(Me.ListBox1.ListIndex, 11)
End Sub
