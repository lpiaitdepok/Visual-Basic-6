'happycodings.com
'Check if file already exists

'---------------- first version 

Function FileExist (Path$) as Integer
    dim x
    x = FreeFile
    on Error Resume Next
    open Path$ For Input as x
    FileExist = (Err = 0)
    Close x
End Function

'---------------- second version 

'thanks for modifications: Lynton 

'The function above assumes that the file you are checking for is
'not locked (in use). In that case, fileexists would return false because
'you are attempting to open a locked file.

Function FileExists%(ByVal sPath$)
  ' Check for the existence of a file.
  dim rc%
  FileExists = False
  on Error Resume Next
  If Len(sPath$) Then
    rc% = Len(Dir$(sPath$))
    If rc% And Not Err Then FileExists% = True
  end If
End Function

'---------------- third version 
'George Toft 


'This is much easier and quicker than the ones you have.  I used to
'use code almost identical to the ones you have until I learned about
'the DIR function.

Public Function FileExist(parmPath as String) as Integer

    FileExist = Not (Dir(parmPath) = "")

End Function' FileExist

'---------------- fourth adjustment
'dayak 


'Using a Form, containing a Textbox, and a Command button, the following code
'works for creating and checking the existence of a directory.
'============================Code Follows===================================



Private sub Command1_Click()

Dim sFname as String
sFname = App.Path & "\" & "mydir"

If Not FileExist(sFname) Then
    MsgBox ("Creating 'mydir' Directory in App.Path")
    MkDir (sFname)
    Text1.Text = "Directory 'mydir' has been created"
Else
    Text1.Text = "Directory 'mydir' already exists"
End If


End Sub


Private Function FileExist(ByRef sFname) as Boolean

        If Len(Dir(sFname, 16)) Then FileExist = True Else FileExist = False

End Function
