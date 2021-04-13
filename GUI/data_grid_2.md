' untuk menghilangkan tampilan fixed column
DataGrid1.Columns(0).Visible = False

' untuk menampilkan tampilan column dengan nama
DataGrid1.Columns("AlamatLengkap").Visible = True

' untuk merefresh datagrid
DataGrid1.Referesh

' untuk masukan data pada datagrid
Set DataGrid1.DataSource = Adodc1

' untuk masukan data dari datagrid
Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Text1.Text = DataGrid1.Columns(0).Text

DataGrid1.AllowDelete = False
DataGrid1.AllowUpdate = False
DataGrid1.AllowAddNew = False
End Sub
