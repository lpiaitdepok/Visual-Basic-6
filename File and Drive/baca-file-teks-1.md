```VB
Private Sub Form_Click()
Dim LinesFromFile, NextLine As String
Dim FileNum As Integer

FileNum = FreeFile
Open App.Path & "\FormMenuUtama.frm" For Input As FileNum

Do Until EOF(FileNum)
   Line Input #FileNum, NextLine
   LinesFromFile = LinesFromFile + NextLine + Chr(13) + Chr(10)
   Print LinesFromFile
Loop
End Sub
```
