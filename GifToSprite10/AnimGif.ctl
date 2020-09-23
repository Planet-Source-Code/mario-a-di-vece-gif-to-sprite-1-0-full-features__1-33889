VERSION 5.00
Begin VB.UserControl AnimGif 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox imgSource 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2655
      Index           =   0
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   3435
      TabIndex        =   0
      Top             =   0
      Width           =   3435
   End
   Begin VB.Timer Timer 
      Left            =   1080
      Top             =   3120
   End
End
Attribute VB_Name = "AnimGif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim mTotalFrames As Long
Dim mRepeatTimes As Long
Dim mGifPath As String
Dim FrameCount As Long
Private Sub Timer_Timer()
On Error Resume Next
Dim i As Long
    If FrameCount < TotalFrames Then
        imgSource(FrameCount).Visible = False
        FrameCount = FrameCount + 1
        imgSource(FrameCount).Visible = True
        Timer.Interval = CLng(imgSource(FrameCount).Tag)
    Else
        On Error Resume Next
        FrameCount = 0
        For i = 1 To imgSource.Count - 1
            imgSource(i).Visible = False
        Next i
        imgSource(FrameCount).Visible = True
        Timer.Interval = CLng(imgSource(FrameCount).Tag)
    End If
End Sub
Private Sub UserControl_Initialize()
imgSource(0).Move 0, 0, ScaleWidth, ScaleHeight

End Sub

Private Sub UserControl_Resize()
UserControl.Width = imgSource(0).Width
UserControl.Height = imgSource(0).Height
End Sub

Public Property Get TotalFrames() As Long
    TotalFrames = mTotalFrames
End Property

Public Property Let TotalFrames(ByVal vNewValue As Long)
    mTotalFrames = vNewValue
End Property

Public Property Let BackColor(ByVal vNewValue As Long)
Dim i As Integer
    UserControl.BackColor = vNewValue
    For i = 0 To imgSource.UBound
        imgSource(i).BackColor = vNewValue
    Next i
End Property

Public Property Get BackColor() As Long
    BackColor = imgSource(0).BackColor
End Property

Public Property Get RepeatTimes() As Long
    RepeatTimes = mRepeatTimes
End Property

Public Property Let RepeatTimes(ByVal vNewValue As Long)
    mRepeatTimes = vNewValue
End Property

Public Property Get GifPath() As String
    GifPath = mGifPath
End Property

Public Property Let GifPath(ByVal vNewValue As String)
On Error Resume Next
    If Dir(vNewValue) = "" Then
        Err.Raise vbObjectError + 1, , "File not found"
        Exit Property
    End If
    If right(vNewValue, 3) <> "gif" Then
        Err.Raise vbObjectError + 2, , "File format is not supported"
        Exit Property
    End If
    mGifPath = vNewValue
End Property
Private Function LoadGif(sFile As String, aImg As Variant) As Boolean
On Error GoTo ShowErr
    LoadGif = False
    If Dir$(sFile) = "" Or sFile = "" Then
       Err.Raise vbObjectError + 1, , "File not found"
       Exit Function
    End If
    On Error GoTo ErrHandler
    Dim fNum As Integer
    Dim imgHeader As String, fileHeader As String
    Dim buf$, picbuf$
    Dim imgCount As Integer
    Dim i&, j&, xOff&, yOff&, TimeWait&
    Dim GifEnd As String
    GifEnd = Chr(0) & Chr(33) & Chr(249)
    For i = 1 To aImg.Count - 1
        Unload aImg(i)
    Next i
    fNum = FreeFile
    Open sFile For Binary Access Read As fNum
        buf = String(LOF(fNum), Chr(0))
        Get #fNum, , buf 'Get GIF File into buffer
    Close fNum
    
    i = 1
    imgCount = 0
    j = InStr(1, buf, GifEnd) + 1
    fileHeader = left(buf, j)
    If left$(fileHeader, 3) <> "GIF" Then
       Err.Raise vbObjectError + 2, , "File format is not supported"
       Exit Function
    End If
    LoadGif = True
    i = j + 2
    If Len(fileHeader) >= 127 Then
        mRepeatTimes = Asc(Mid(fileHeader, 126, 1)) + (Asc(Mid(fileHeader, 127, 1)) * 256&)
    Else
        mRepeatTimes = 0
    End If

    Do ' Split GIF Files at separate pictures
       ' and load them into Image Array
        imgCount = imgCount + 1
        j = InStr(i, buf, GifEnd) + 3
        If j > Len(GifEnd) Then
            fNum = FreeFile
            Open "temp.gif" For Binary As fNum
                picbuf = String(Len(fileHeader) + j - i, Chr(0))
                picbuf = fileHeader & Mid(buf, i - 1, j - i)
                Put #fNum, 1, picbuf
                imgHeader = left(Mid(buf, i - 1, j - i), 16)
            Close fNum
            TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256&)) * 10&
            If imgCount > 1 Then
                xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256&)
                yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256&)
                Load aImg(imgCount - 1)
                aImg(imgCount - 1).left = aImg(0).left + (xOff * Screen.TwipsPerPixelX)
                aImg(imgCount - 1).top = aImg(0).top + (yOff * Screen.TwipsPerPixelY)
            End If
            ' Use .Tag Property to save TimeWait interval for separate Image
            aImg(imgCount - 1).Tag = TimeWait
            aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
            Kill ("temp.gif")
            i = j
        End If
        DoEvents
    Loop Until j = 3
' If there are one more Image - Load it
    If i < Len(buf) Then
        fNum = FreeFile
        Open "temp.gif" For Binary As fNum
            picbuf = String(Len(fileHeader) + Len(buf) - i, Chr(0))
            picbuf = fileHeader & Mid(buf, i - 1, Len(buf) - i)
            Put #fNum, 1, picbuf
            imgHeader = left(Mid(buf, i - 1, Len(buf) - i), 16)
        Close fNum
        TimeWait = ((Asc(Mid(imgHeader, 4, 1))) + (Asc(Mid(imgHeader, 5, 1)) * 256)) * 10
        If imgCount > 1 Then
            xOff = Asc(Mid(imgHeader, 9, 1)) + (Asc(Mid(imgHeader, 10, 1)) * 256)
            yOff = Asc(Mid(imgHeader, 11, 1)) + (Asc(Mid(imgHeader, 12, 1)) * 256)
            Load aImg(imgCount - 1)
            aImg(imgCount - 1).left = aImg(0).left + (xOff * Screen.TwipsPerPixelX)
            aImg(imgCount - 1).top = aImg(0).top + (yOff * Screen.TwipsPerPixelY)
        End If
        aImg(imgCount - 1).Tag = TimeWait
        aImg(imgCount - 1).Picture = LoadPicture("temp.gif")
        Kill ("temp.gif")
    End If
    TotalFrames = aImg.Count - 1
    UserControl.Width = imgSource(0).Width
    UserControl.Height = imgSource(0).Height
    Exit Function
ShowErr:
    MsgBox "Error: File not found/loaded. Plaese try again.", vbCritical, "File not found/loaded"
    Exit Function
ErrHandler:
    MsgBox "Error loading. Gif format invalid or object already loaded.", vbCritical, "Load Error"
    'Exit Function
    'Err.Raise Err.Number, Err.Source, Err.Description
    LoadGif = False
    On Error GoTo 0
End Function


Public Sub StartGif()
    Timer.Enabled = False
    If LoadGif(mGifPath, imgSource) Then
       FrameCount = 0
       Timer.Interval = CLng(imgSource(0).Tag)
       Timer.Enabled = True
    End If
End Sub

Public Sub StopGif()
    Timer.Enabled = False
End Sub

Public Sub ContinueGif()
    Timer.Enabled = True
End Sub
Public Sub SetBackColor(BackColor As Long)
    imgSource(0).BackColor = BackColor
    UserControl.BackColor = BackColor
End Sub
Public Property Get Picture(FrameNumber As Integer) As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
On Error Resume Next
    Set Picture = imgSource(FrameNumber).Picture
End Property
