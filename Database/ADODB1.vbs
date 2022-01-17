• Simpan Data :
ado.Execute "INSERT INTO [nama tabel] VALUES ('" + Text1.Text + "','" + Text2.Text + "')"

• Pencarian Data
Set Rs = New Adodb.Recordset
Rs.Open "SELECT * FROM [nama table1] WHERE [nama kolom1]='" + Text3.Text + "'", ado
If Not rs.EOF Then
Text1.Text = rs("namakolom1")
Text2.Text = rs("namakolom2")
Else
MsgBox "Maaf, Data Tidak Ditemukan!"
End if

• Edit Data
ado.Execute "UPDATE [nama tabel] Set [namakolom1]='" + Text1.Text + _
"',[namakolom2]='" + Text2.Text + _
"' WHERE [nama kolom1]='" + Text3.Text + "'"

'Code diatas tidak memerlukan lagi kode pencarian seperti code edit untuk DATA dan Adodc

• Hapus Data
ado.Execute "DELETE * FROM [nama tabel] WHERE [nama kolom1]='" + Text3.Text + "'"
'Code diatas tidak memerlukan lagi kode pencarian seperti code hapus untuk DATA dan Adodc

# in ASP
Note: The returned Recordset is always a read-only, forward-only Recordset!

Tip: To create a Recordset with more functionality, first create a Recordset object. Set the desired properties, and then use the Recordset object's Open method to execute the query.

# in Visual Basic Classic
Note Use the ExecuteOptionEnum value adExecuteNoRecords to improve performance by minimizing internal processing and for applications that you are porting from Visual Basic 6.0.
