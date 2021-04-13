'make sure the search path to the db is always in the right spot
Data1.DatabaseName = App.Path & "\Student.mdb"

'set up the recordsource for the datasource and flexgrid control
'in this case, it's just a raw SQL query, simple simple.
Data1.RecordSource = "SELECT * FROM STUDENT ORDER BY StudentName"

'End Of File
Data1.Recordset.EOF

'Begin Of file
Data1.Recordset.BOF

'
Data1.Recordset.MoveFirst

'add a new entry to our table.
Data1.Recordset.AddNew 
'Data1.Recordset!namakolom1 = Text1.Text
Data1.Recordset!StudentName = txtStudentName.Text
'or you can use this
Data1.Recordset("namakolom1") = Text1.Text
Data1.Update
Data1.Refresh

'delete an entry from the database
' Kode ini sebaiknya dijalankan setelah kode pencarian dijalankan terlebih dahulu.
Data1.Recordset.Move ( MSFlexgrid1.Row - 1) ' we minus one because row zero is the header row
Data1.Delete
Data1.Refresh

'Pencarian Data :
Data1.Recordset.Index = "KodeIdx"
Data1.Recordset.Seek "=", Textcari.Text
If Not Data1.Recordset.NoMatch Then
Text1.Text = Data1.Recordset!namakolom1
Else
MsgBox "Maaf, Data Tidak Ditemukan!"
End if

' Edit Data :
' Kode ini sebaiknya dijalankan setelah kode pencarian dijalankan terlebih dahulu.
Data1.Recordset.Edit
Data1.Recordset!namakolom1=Text1.Text
Data1.Recordset!namakolom2=Text2.Text
Data1.Recordset.Update
Data1.Refresh
