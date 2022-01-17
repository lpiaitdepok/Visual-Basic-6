
```vb
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim strConn As String 
Dim SQL As String

'Create the connection
Set con = New ADODB.Connection

strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & App.Path & "\nama-database.mdb;" & _
          "Persist Security Info=False"

con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & strDataSource
con.Open

Set rs = New ADODB.Recordset

rs.CursorLocation = adUseClient

SQL = "SELECT * FROM DataSiswa"

rs.Open SQL, con, adOpenDynamic, adLockOptimistic ', adCmdText

Set DataGrid1.DataSource = rs("Nama")
```
