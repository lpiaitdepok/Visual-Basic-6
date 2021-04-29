Sub LoopAllFilesInAFolder()

'Loop through all files in a folder
Dim fileName As Variant
fileName = Dir("C:\Users\marks\Documents\")

While fileName <> ""
    
    'Insert the actions to be performed on each file
    'This example will print the file name to the immediate window
    Debug.Print fileName

    'Set the fileName to the next file
    fileName = Dir
Wend

End Sub
