• Simpan Data :
ado.Execute "INSERT INTO [nama tabel] VALUES ('" + Text1.Text + "','" + Text2.Text + "')"

• Pencarian Data
Set Rs = New Adodb.Recordset
Rs.Open "SELECT * FROM [nama table1] WHERE [nama kolom1]='" + TextCari.Text + "'", ado
If Not rs.EOF Then
Text1.Text = rs("namakolom1")
Text2.Text = rs("namakolom2")
Else
MsgBox "Maaf, Data Tidak Ditemukan!"
End if

• Edit Data
ado.Execute "UPDATE [nama tabel] Set [namakolom1]='" + Text1.Text + _
"',[namakolom2]='" + Text2.Text + _
"' WHERE [nama kolom1]='" + TextCari.Text + "'"

Code diatas tidak memerlukan lagi kode pencarian seperti code edit untuk DATA dan Adodc

• Hapus Data
ado.Execute "DELETE * FROM [nama tabel] WHERE [nama kolom1]='" + TextCari.Text + "'"
Code diatas tidak memerlukan lagi kode pencarian seperti code hapus untuk DATA dan Adodc
