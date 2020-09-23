VERSION 5.00
Begin VB.Form frmStrip 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sprite Preview"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   8085
   Icon            =   "frmStrip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   8085
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox SaveStrip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   -5580
      ScaleHeight     =   1425
      ScaleWidth      =   5505
      TabIndex        =   3
      Top             =   2340
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.PictureBox FullStrip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      FillColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   3465
      TabIndex        =   1
      Top             =   0
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
         TabIndex        =   2
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.HScrollBar Scroller 
      Height          =   255
      Left            =   0
      Max             =   1
      TabIndex        =   0
      Top             =   1200
      Width           =   7215
   End
   Begin VB.Menu mnuSvaeSprite 
      Caption         =   "&Save Sprite As..."
   End
   Begin VB.Menu mnuBackground 
      Caption         =   "&Change Background Color..."
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "Sprite &Info..."
   End
End
Attribute VB_Name = "frmStrip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Color As Variant
Private m_cDib As New cDIBSection

Private Sub PrintFrames()
SaveStrip.Visible = True
SaveStrip.Width = FullStrip.Width
SaveStrip.Height = FullStrip.Height

    For i = 0 To StripFrame.UBound
        SaveStrip.PaintPicture StripFrame(i).Image, StripFrame(i).left, 0
    Next i
End Sub

Private Sub Form_Load()
On Local Error GoTo ShowErr
    GenerateAllFrames
    Scroller.left = 0
    Scroller.Width = Me.ScaleWidth
    Me.Height = StripFrame(0).Height + Scroller.Height + 700
    Scroller.top = StripFrame(0).Height + 25
    Scroller.Min = 0
    Scroller.Max = frmMain.AnimGif.TotalFrames
    FullStrip.top = 0
    FullStrip.Height = StripFrame(0).Height
    FullStrip.Width = StripFrame(0).Width * (StripFrame.Count)
Exit Sub
ShowErr:
MsgBox "Error: Please load a file first.", vbCritical, "No loaded file found"
Unload Me

End Sub


Private Sub mnuBackground_Click()

Color = GetColor(StripFrame(0).BackColor)
'If cancel was selected then the returnvalue = &h7FFFFFFF
If Color > &HFFFFFF Then Exit Sub 'cancel was selected
For i = 0 To StripFrame.UBound
    StripFrame(i).BackColor = Color
Next i
End Sub

Private Sub mnuInfo_Click()
MsgBox "Sprite Information: " & vbNewLine & _
       "Frame Width  :  " & TwipsToPixels_width(frmMain.AnimGif.Width) & vbNewLine & _
       "Frame Height : " & TwipsToPixels_height(frmMain.AnimGif.Height) & vbNewLine & _
       "Total Frames : " & frmMain.AnimGif.TotalFrames, vbInformation, "Sprite Information"
End Sub

Function TwipsToPixels_height(pxls)
    TwipsToPixels_height = pxls \ Screen.TwipsPerPixelY
End Function


Function TwipsToPixels_width(pxls)
    TwipsToPixels_width = pxls \ Screen.TwipsPerPixelX
End Function


Private Sub mnuSvaeSprite_Click()

Dim sI As String
PrintFrames

If SaveMode = "JPG" Then

SavePicture SaveStrip.Image, "tmp.dat"
SaveStrip.Visible = False
Set m_cDib = New cDIBSection
m_cDib.CreateFromPicture LoadPicture("tmp.dat")

   If VBGetSaveFileName(sI, , , "JPEG Files (*.JPG)|*.JPG|All Files (*.*)|*.*", , , , "JPG", Me.hWnd) Then
      If SaveJPG(m_cDib, sI, JPEGQuality) Then
         MsgBox "File " & sI & " was successfuly saved.", vbInformation, "Operation Succeeded"
         Kill "tmp.dat"
         Unload Me
      Else
         MsgBox "Failed to save the picture to the file: '" & sI & "'", vbCritical, "Save Error"
         Unload Me
      End If
   End If
Else

   If VBGetSaveFileName(sI, , , "Bitmap Files (*.bmp)|*.bmp|All Files (*.*)|*.*", , , , "BMP", Me.hWnd) Then
         On Error GoTo Failed
         SavePicture SaveStrip.Image, sI
         MsgBox "File " & sI & " was successfuly saved.", vbInformation, "Operation Succeeded"
         Unload Me
         Exit Sub
Failed:
         MsgBox "Failed to save the picture to the file: '" & sI & "'", vbCritical, "Save Error"
         Unload Me
         Exit Sub
   End If
End If

End Sub

Private Sub Scroller_Change()
    FullStrip.left = -1 * (Scroller.Value * StripFrame(0).Width)
End Sub

Private Sub GenerateAllFrames()
StripFrame(0).Picture = frmMain.AnimGif.Picture(0)

Dim i As Integer

    For i = 1 To frmMain.AnimGif.TotalFrames
        Load StripFrame(i)
        StripFrame(i).Picture = frmMain.AnimGif.Picture(i)
        StripFrame(i).left = StripFrame(0).left + (StripFrame(0).Width * i)
        StripFrame(i).Visible = True
    Next i

End Sub
