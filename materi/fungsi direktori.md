```bas
Public Function direktori() As String 
 If Right$(App.Path, 1) = "\" Then 
 direktori$ = App.Path 
 Else 
 direktori$ = App.Path & "\" 
 End If 
End Function
```
