VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gif-To-Sprite 1.0 by Mario Di Vece"
   ClientHeight    =   5550
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   7350
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   4320
      Left            =   120
      Pattern         =   "*.gif"
      TabIndex        =   8
      Top             =   1020
      Width           =   2355
   End
   Begin VB.CommandButton Command9 
      Height          =   555
      Left            =   3000
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":27A2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "About this program..."
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Height          =   555
      Left            =   2400
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":37E4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Generate Sprite"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Height          =   555
      Left            =   1800
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":4826
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Batch process"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Height          =   555
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":5868
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Options and Settings"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Height          =   555
      Left            =   600
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":68AA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Change animation background color"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Height          =   555
      Left            =   0
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmMain.frx":78EC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Change Folder"
      Top             =   0
      UseMaskColor    =   -1  'True
      Width           =   615
   End
   Begin VB.PictureBox GifFrame 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   2520
      ScaleHeight     =   4305
      ScaleWidth      =   4665
      TabIndex        =   10
      Top             =   1020
      Width           =   4695
      Begin VB.CommandButton Command1 
         Height          =   435
         Left            =   3420
         Picture         =   "frmMain.frx":892E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Paly Animation"
         Top             =   3900
         Width           =   435
      End
      Begin VB.CommandButton Command3 
         Height          =   435
         Left            =   3840
         Picture         =   "frmMain.frx":9970
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Stop Animation"
         Top             =   3900
         Width           =   435
      End
      Begin VB.CommandButton Command2 
         Height          =   435
         Left            =   4260
         Picture         =   "frmMain.frx":A9B2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Continue Animation"
         Top             =   3900
         Width           =   435
      End
      Begin GifToSprite.AnimGif AnimGif 
         Height          =   1155
         Left            =   1620
         TabIndex        =   11
         Top             =   1440
         Width           =   1335
         _ExtentX        =   8467
         _ExtentY        =   6350
      End
   End
   Begin VB.Label lblGifs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Gifs in Folder:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   780
      Width           =   1650
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "If you want new versions of this program, please send feedback. (mariodivece@hotmail.com)"
      Height          =   435
      Left            =   3720
      TabIndex        =   7
      Top             =   60
      Width           =   3495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   7320
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Animation Preview"
      Height          =   195
      Left            =   2580
      TabIndex        =   0
      Top             =   780
      Width           =   1305
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu smnuChangeDirectory 
         Caption         =   "Change Directory"
      End
      Begin VB.Menu smnuBatchProcessing 
         Caption         =   "Batch Processing"
      End
      Begin VB.Menu smnuGeneralOptions 
         Caption         =   "General Options"
      End
      Begin VB.Menu smnuAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AGifIsLoaded As Boolean
Dim CurrentFolder As String

Private Sub Command1_Click()
    AnimGif.StartGif
End Sub

Private Sub Command2_Click()
    AnimGif.ContinueGif
End Sub

Private Sub Command3_Click()
    AnimGif.StopGif
End Sub

Private Sub Command4_Click()
ChangeFolder
End Sub

Private Sub ChangeFolder()
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim szTitle As String
    Dim tBrowseInfo As BrowseInfo
    szTitle = "Select the folder where your GIFs are"


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
        CurrentFolder = sBuffer
        File1.Path = sBuffer
    End If

End Sub

Private Sub Command5_Click()
Dim Color As Long
Color = GetColor(AnimGif.BackColor)
'If cancel was selected then the returnvalue = &h7FFFFFFF
If Color > &HFFFFFF Then Exit Sub 'cancel was selected
AnimGif.BackColor = Color
GifFrame.BackColor = Color
End Sub

Private Sub Command6_Click()
    frmOptions.Show 1, Me
End Sub

Private Sub Command7_Click()
On Error Resume Next
    
    frmBatch.Show 1, Me
End Sub

Private Sub Command8_Click()
On Error Resume Next
If AGifIsLoaded Then
    AnimGif.StopGif
    frmStrip.Show 1, Me
End If
End Sub

Private Sub Command9_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub File1_DblClick()
On Local Error GoTo ErrLoad
    AnimGif.StopGif
    AnimGif.GifPath = File1.Path & "\" & File1.Filename
    AnimGif.Visible = False
    CenterGif
    CenterGif
    AnimGif.StartGif
    CenterGif
    AnimGif.Visible = True
    'MsgBox AnimGif.TotalFrames
    AGifIsLoaded = True
Exit Sub
ErrLoad:
    MsgBox "Error loading. Gif format invalid or object already loaded.", vbCritical, "Load Error"
    AGifIsLoaded = False
End Sub
Private Sub LoadSettings()
    File1.Path = App.Path
    JPEGQuality = 100
    SaveMode = "JPG"
    AGifIsLoaded = False
End Sub

Private Sub Form_Load()
AnimGif.SetBackColor vbWhite
LoadSettings
End Sub

Private Sub CenterGif()
    AnimGif.left = (GifFrame.Width / 2) - (AnimGif.Width / 2)
    AnimGif.top = (GifFrame.Height / 2) - (AnimGif.Height / 2)
End Sub

Private Sub smnuAbout_Click()
    frmAbout.Show 1, Me
End Sub

Private Sub smnuBatchProcessing_Click()
On Error Resume Next
    
    frmBatch.Show 1, Me
End Sub

Private Sub smnuChangeDirectory_Click()
ChangeFolder
End Sub

Private Sub smnuGeneralOptions_Click()
    frmOptions.Show 1, Me
End Sub
