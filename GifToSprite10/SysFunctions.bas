Attribute VB_Name = "SysFunctions"
'**************************************
'Windows API/Global Declarations for :Br
'     owse Folder Dialog
'**************************************
Private Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000

Public Const MAX_PATH = 260


Public Declare Function SHBrowseForFolder Lib "shell32" _
    (lpbi As BrowseInfo) As Long

Public Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, _
    ByVal lpBuffer As String) As Long


Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
    (ByVal lpString1 As String, ByVal _
    lpString2 As String) As Long


Public Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
'**************************************

