'List all files in sub folders
Sub LoopAllSubFolders(ByVal folderPath As String)

Dim fileName As String
Dim fullFilePath As String
Dim numFolders As Long
Dim folders() As String
Dim i As Long

If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
fileName = Dir(folderPath & "*.*", vbDirectory)

While Len(fileName) <> 0

    If Left(fileName, 1) <> "." Then
 
        fullFilePath = folderPath & fileName
 
        If (GetAttr(fullFilePath) And vbDirectory) = vbDirectory Then
            ReDim Preserve folders(0 To numFolders) As String
            folders(numFolders) = fullFilePath
            numFolders = numFolders + 1
        Else
            'Insert the actions to be performed on each file
            'This example will print the full file path to the immediate window
            Debug.Print folderPath & fileName
        End If
 
    End If
 
    fileName = Dir()

Wend

For i = 0 To numFolders - 1

    LoopAllSubFolders folders(i)
 
Next i

End Sub
