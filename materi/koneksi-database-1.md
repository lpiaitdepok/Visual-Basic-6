```vb
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Qy As New ADODB.Command
Dim sql As String

Const ConnectString = "Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & App.Path & "\DBPembelian.mdb"

cn.Open ConnectString
```

