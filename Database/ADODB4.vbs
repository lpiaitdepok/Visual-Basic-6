Public Conn As ADODB.Connection
Public Rs As ADODB.Recordset
Set Conn = New ADODB.Connection
Set Rs = New ADODB.Recordset
Conn.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DatabaseName.mdb"
Rs.Open "SELECT * FROM TableName", Conn
