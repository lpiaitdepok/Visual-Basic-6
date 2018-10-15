'Example for adding items into combo by using Win32 API.
' faster when calling after first time
' www.nirsoft.net

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const CB_ADDSTRING = &H143


Private Sub cmdComboAPI_Click()
    Dim strItemText     As String
    Dim intIndex        As Integer
    Dim dblTimer        As Double
    
    cmbTest.Clear
    dblTimer = Timer
    'Adding items with Win32 API
    For intIndex = 1 To 5000
        strItemText = "item number " & CStr(intIndex)
        SendMessage cmbTest.hWnd, CB_ADDSTRING, 0, ByVal strItemText
    Next
    MsgBox Format(Timer - dblTimer, "0.000") & " seconds"

End Sub

Private Sub cmdComboVB_Click()
    Dim strItemText     As String
    Dim intIndex        As Integer
    Dim dblTimer        As Double
    
    cmbTest.Clear
    dblTimer = Timer
    'Adding items with standard AddItem method.
    For intIndex = 1 To 5000
        strItemText = "item number " & CStr(intIndex)
        cmbTest.AddItem strItemText
    Next
    MsgBox Format(Timer - dblTimer, "0.000") & " seconds"

End Sub
