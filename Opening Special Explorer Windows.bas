'Opening Special Explorer Windows
'
' www.nirsoft.net

Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" _
(ByVal hwndOwner As Long, ByVal lpszPath As String, ByVal nFolder As Long, ByVal fCreate As Long) As Long

Private Const CSIDL_FONTS = &H14
Private Const CSIDL_DESKTOP = &H0
Private Const CSIDL_FAVORITES = &H6
Private Const CSIDL_RECENT = &H8
Private Const CSIDL_COOKIES = &H21
Private Const CSIDL_HISTORY = &H22

Private Const NameSpace_MyComputer = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
Private Const NameSpace_RecycleBin = "::{645FF040-5081-101B-9F08-00AA002F954E}"
Private Const NameSpace_NetworkNeighborhood = "::{208D2C60-3AEA-1069-A2D7-08002B30309D}"
Private Const NameSpace_Dialup = "::{a4d92740-67cd-11cf-96f2-00aa00a11dd9}"
Private Const NameSpace_ControlPanel = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Private Const NameSpace_Printers = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{2227A280-3AEA-1069-A2DE-08002B30309D}"
Private Const NameSpace_ScheduledTasks = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{D6277990-4C6A-11CF-8D87-00AA0060F5BF}"

Private Const MAX_PATH = 260


Private Sub OpenExplorerWindow(FolderName As String)
    Shell "explorer " & FolderName, vbNormalFocus
End Sub

Private Function TrimNull(Str1 As String) As String
    Dim Loc         As Integer
    
    Loc = InStr(Str1, Chr$(0))
    If Loc <> 0 Then
        TrimNull = Mid$(Str1, 1, Loc - 1)
    Else
        TrimNull = Str1
    End If
End Function

Private Function GetSpecialFolder(Folder As Long) As String
    Dim FolderPath          As String * MAX_PATH
    SHGetSpecialFolderPath 0, FolderPath, Folder, 0
    GetSpecialFolder = TrimNull(FolderPath)
End Function

Private Sub cmdControlPanel_Click()
    OpenExplorerWindow NameSpace_ControlPanel
End Sub

Private Sub cmdDialup_Click()
    OpenExplorerWindow NameSpace_Dialup
End Sub

Private Sub cmdCookies_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_COOKIES)
End Sub

Private Sub cmdDesktop_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_DESKTOP)
End Sub

Private Sub cmdFavorites_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_FAVORITES)
End Sub

Private Sub cmdFonts_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_FONTS)
End Sub

Private Sub cmdHistory_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_HISTORY)
End Sub

Private Sub cmdMyComputer_Click()
    OpenExplorerWindow NameSpace_MyComputer
End Sub

Private Sub cmdNetworkNeighborhood_Click()
    OpenExplorerWindow NameSpace_NetworkNeighborhood
End Sub

Private Sub cmdPrinters_Click()
    OpenExplorerWindow NameSpace_Printers
End Sub

Private Sub cmdRecent_Click()
    OpenExplorerWindow GetSpecialFolder(CSIDL_RECENT)
End Sub

Private Sub cmdRecycleBin_Click()
    OpenExplorerWindow NameSpace_RecycleBin
End Sub

Private Sub cmdScheduledTasks_Click()
    OpenExplorerWindow NameSpace_ScheduledTasks
End Sub
