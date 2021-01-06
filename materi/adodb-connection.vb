Dim con As New ADODB.Connection  
Dim rs As New ADODB.Recordset

con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=nama_database.mdb;Mode=readwrite"
con.Open
