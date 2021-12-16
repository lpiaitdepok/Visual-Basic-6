Private Sub cmdClose_Click()
'-- untuk menutup semua form
Dim iLoop As Integer
Dim iHighestForm As Integer
iHighestForm = Forms.Count - 1
For iLoop = 1 To iHighestForm
Unload Forms(iLoop)
Next iLoop
End Sub
