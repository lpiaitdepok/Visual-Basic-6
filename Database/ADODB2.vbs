Option Explicit
Private cn As ADODB.Connection  'this is the connection object
Private rs As ADODB.Recordset   'this is the recordset object

Private Sub Form_Load()
    'turn MousePointer to HourGlass to show that we are busy processing
    Me.MousePointer = vbHourglass
    
    'instantiate the connection object
    Set cn = New ADODB.Connection
    'specify the connectionstring
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source=" & App.Path & "\DatabaseName.mdb"
    'open the connection
    cn.Open
    
    'instantiate the recordset object
    Set rs = New ADODB.Recordset
    'open the recordset
    With rs
        .Open "tbl_master", cn, adOpenKeyset, adLockPessimistic, adCmdTable
           
        'loop through the records until reaching the end or last record
        Do While Not .EOF
            Combo1.AddItem rs.Fields("field1")
            rs.MoveNext 'moves next record
        Loop
        
        If Not (.EOF And .BOF) Then
            rs.MoveFirst    'go to the first record if there are existing records
            FillFields      'to reflect the current record in the controls
        End If
        
    End With
    
    Me.MousePointer = vbNormal 'sets the mouse pointer to the normal arrow
End Sub
