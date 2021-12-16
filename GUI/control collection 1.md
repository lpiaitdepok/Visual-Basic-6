Private Sub Command1_Click()
Dim ctl As Control

For Each ctl In Form1.Controls
If TypeOf ctl Is OptionButton Then
ctl.Value = False
End If
Next

End Sub
