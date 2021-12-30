```
Private Sub MSFlexGrid1_RowColChange()
'Taruh isi grid ke Text1
Text1.Text = MSFlexGrid1.TextMatrix(MSFlexGrid1.RowSel, 1)
End Sub
```
