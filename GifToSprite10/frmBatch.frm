VERSION 5.00
Begin VB.Form frmBatch 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Batch Processing"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4785
   Icon            =   "frmBatch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin GifToSprite.AnimGif AnimGif 
      Height          =   3600
      Left            =   4320
      TabIndex        =   12
      Top             =   7200
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   6350
   End
   Begin VB.PictureBox SaveStrip 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   360
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   11
      Top             =   7260
      Width           =   3495
   End
   Begin VB.PictureBox FullStrip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   300
      ScaleHeight     =   525
      ScaleWidth      =   3465
      TabIndex        =   9
      Top             =   5100
      Width           =   3495
      Begin VB.PictureBox StripFrame 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         FillColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   1035
         Index           =   0
         Left            =   0
         ScaleHeight     =   1035
         ScaleWidth      =   1875
         TabIndex        =   10
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2100
      TabIndex        =   8
      Top             =   3660
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "GO!"
      Default         =   -1  'True
      Height          =   375
      Left            =   3300
      TabIndex        =   7
      Top             =   3660
      Width           =   1155
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   300
      TabIndex        =   6
      Top             =   2520
      Width           =   4155
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Resume on error"
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   2160
      Value           =   1  'Checked
      Width           =   4095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Quality (Only JPG format):"
      Height          =   195
      Left            =   300
      TabIndex        =   4
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SaveFormat:"
      Height          =   195
      Left            =   300
      TabIndex        =   3
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files found:"
      Height          =   195
      Left            =   300
      TabIndex        =   2
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PATH"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   300
      TabIndex        =   1
      Top             =   540
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folder to process:"
      Height          =   195
      Left            =   300
      TabIndex        =   0
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frmBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CF As String
Private m_cDib As New cDIBSection

Private Sub ChangeFolder()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "Select the destination folder"


    With tBrowseInfo
        .hWndOwner = Me.hWnd
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
        
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)


    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        CF = sBuffer
    End If

End Sub

Private Sub Command1_Click()
ChangeFolder
If CF = nullstring Then
    Exit Sub
Else
On Local Error Resume Next
Dim SavePath As String

SavePath = "C:\Documents and Settings\Administrator\My Documents\VB\Gif-To-Sprite New\testbatch"
'Ask for destination folder

Dim i As Integer
    
    List1.AddItem "Batch process started at " & Time & " on " & Date
Command1.Enabled = False

For m = 0 To frmMain.File1.ListCount - 1
    Me.Refresh
    List1.Refresh
    'load the sprite
    AnimGif.GifPath = frmMain.File1.Path & "\" & frmMain.File1.List(m)
    List1.AddItem "Loading frmaes (" & frmMain.File1.List(m) & ")"
    GenerateAllFrames
    FullStrip.Height = StripFrame(0).Height
    FullStrip.Width = StripFrame(0).Width * (StripFrame.Count)
    SaveStrip.Width = FullStrip.Width
    SaveStrip.Height = FullStrip.Height
    'prints frames to savestrip
    For i = 0 To StripFrame.UBound
        SaveStrip.PaintPicture StripFrame(i).Image, StripFrame(i).left, 0
    Next i
    
    'Save sprite in the desired folder with desired format and same name as source
    If SaveToFile(SavePath, Mid(frmMain.File1.List(m), 1, Len(frmMain.File1.List(m)) - 4) & "." & SaveMode) = False Then
        If Check1.Value = 0 Then
            GoTo stopOperations
        End If
    End If
    'unload previously used frames
    For u = 1 To StripFrame.UBound
        Unload StripFrame(u)
    Next u
Next m
List1.AddItem "Batch Process finished at " & Time & " on " & Date
Command1.Enabled = False
Exit Sub

stopOperations:
List1.AddItem "Operation stopped because of a runtime error"
Command1.Enabled = False
End If
End Sub

Private Function SaveToFile(DesiredPath As String, DesiredFileName As String) As Boolean
Dim sI As String
sI = DesiredPath & "\" & DesiredFileName

If SaveMode = "JPG" Then
SavePicture SaveStrip.Image, "tmp.dat"

Set m_cDib = New cDIBSection
m_cDib.CreateFromPicture LoadPicture("tmp.dat")

      If SaveJPG(m_cDib, sI, JPEGQuality) Then
         List1.AddItem "File " & sI & " was successfuly saved."
         Kill "tmp.dat"
         SaveToFile = True
         Exit Function
      Else
         List1.AddItem "Failed to save the picture to the file: '" & sI & "'"
         SaveToFile = False
         Exit Function
      End If

Else
         On Error GoTo Failed
         SavePicture SaveStrip.Image, sI
         List1.AddItem "File " & sI & " was successfuly saved."
         SaveToFile = True
         Exit Function
Failed:
         List1.AddItem "Failed to save the picture to the file: '" & sI & "'"
         SaveToFile = False
         Exit Function
End If

End Function

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
If frmMain.File1.ListCount = 0 Then
    Unload Me
Else
MsgBox "Thank you for trying new features. Batch Processing is not 100% bug-free, but it works." & vbNewLine & "Don't forget to report bugs, errors and request features at mariodivece@hotmail.com", vbInformation, "Notice"
    
    List1.AddItem "Gif-To-Sprite Batch Processing Log"
    Label2.Caption = frmMain.File1.Path
    Label3.Caption = "Files Found: " & frmMain.File1.ListCount
    Label4.Caption = "Save Format: " & SaveMode
    Label5.Caption = "Quality (Only JPG format): " & JPEGQuality
End If
End Sub

Private Sub GenerateAllFrames()
AnimGif.StartGif
StripFrame(0).Picture = AnimGif.Picture(0)

Dim i As Integer

    For i = 1 To AnimGif.TotalFrames
        Load StripFrame(i)
        StripFrame(i).Picture = AnimGif.Picture(i)
        StripFrame(i).left = StripFrame(0).left + (StripFrame(0).Width * i)
    Next i
AnimGif.StopGif
End Sub
