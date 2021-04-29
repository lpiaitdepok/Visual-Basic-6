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

'For example:

'Loop through each file with an extension of ".xlsx"
'fileName = Dir("C:\Users\marks\Documents\*.xlsx")
'Loop through each file containing the word "January" in the filename
'fileName = Dir("C:\Users\marks\Documents\*January*")
'Loop through each text file in a folder
'fileName = Dir("C:\Users\marks\Documents\*.txt")
End Sub
