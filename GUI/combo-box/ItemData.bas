'happycodings.com
Private Sub Form_Load()
    'Fill List1 and ItemData array with
    'corresponding items in sorted order.
    With List1
        .AddItem "Mallik Murthy"
        .ItemData(.NewIndex) = 42310
        .AddItem "Chien Lieu"
        .ItemData(.NewIndex) = 52855
        .AddItem "Mauro Sorrento"
        .ItemData(.NewIndex) = 64932
        .AddItem "Cynthia Bennet"
        .ItemData(.NewIndex) = 39227
    End With
End Sub

Private Sub List1_Click()
    With List1
        'Append the employee number and the employee name.
        MsgBox .ItemData(.ListIndex) & " " & .List(.ListIndex), vbInformation
    End With
End Sub
