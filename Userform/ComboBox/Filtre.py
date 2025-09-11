Private Sub CboAnnee_Change()
    Call ChargerRapports(Me.CboAnnee.Value, Me.CboMois.Value, Me.CboJour.Value)
End Sub

Private Sub CboMois_Change()
    Call ChargerRapports(Me.CboAnnee.Value, Me.CboMois.Value, Me.CboJour.Value)
End Sub

Private Sub CboJour_Change()
    Call ChargerRapports(Me.CboAnnee.Value, Me.CboMois.Value, Me.CboJour.Value)
End Sub
