Option Explicit
Public conn As ADODB.Connection 'nama koneksi Conn

'recordset
Set rs = New ADODB.Recordset
rs.ActiveConnection = conn
rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.LockType = adLockBatchOptimistic
rs.Source = "SELECT * FROM table"
rs.Open
