'ADODC

' Simpan Data :
Adodc1.Recordset.AddNew
Adodc1.Recordset!namakolom1 = Text1.Text
Adodc1.Recordset!namakolom2 = Text2.Text
Adodc1.Recordset.Update
Adodc1.Refresh

' Pencarian Data :
Adodc1.Recordset.Find "namakolom1='" + Text1.Text + "'", , adSearchForward, 1
If Not Adodc1.Recordset.EOF Then
Text1.Text = Adodc1.Recordset!namakolom1
Text2.Text = Adodc1.Recordset!namakolom2
Else
MsgBox "Maaf, Data Tidak Ditemukan!"
End if

' Edit Data :
'Kode ini sebaiknya dijalankan setelah kode pencarian dijalankan terlebih dahulu.

Adodc1.Recordset!namakolom1=Text1.Text
Adodc1.Recordset!namakolom2=Text2.Text
Adodc1.Recordset.Update
Adodc1.Refresh

' Hapus Data :
' Kode ini sebaiknya dijalankan setelah kode pencarian dijalankan terlebih dahulu.
Adodc1.Recordset.Delete
Adodc1.Refresh
