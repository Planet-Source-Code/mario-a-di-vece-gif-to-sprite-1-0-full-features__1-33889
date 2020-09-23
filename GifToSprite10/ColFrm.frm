VERSION 5.00
Begin VB.Form ColFrm 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "    Color Picker"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   194
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   374
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Cancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   4860
      TabIndex        =   17
      Top             =   2475
      Width           =   690
   End
   Begin VB.CommandButton Pick 
      Caption         =   "&OK"
      Height          =   330
      Left            =   4095
      TabIndex        =   16
      Top             =   2475
      Width           =   690
   End
   Begin VB.PictureBox CPic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1950
      Left            =   90
      ScaleHeight     =   130
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   135
      Width           =   3840
   End
   Begin VB.Label InfoLabel 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Color Properties"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4140
      TabIndex        =   22
      Top             =   180
      Width           =   1320
   End
   Begin VB.Label OldLabel 
      BackColor       =   &H00C0C0C0&
      Height          =   285
      Left            =   90
      TabIndex        =   21
      Top             =   2475
      Width           =   1410
   End
   Begin VB.Label OLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Old Color"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   90
      TabIndex        =   20
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label PLabel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Color"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1665
      TabIndex        =   19
      Top             =   2160
      Width           =   1410
   End
   Begin VB.Label PickLabel 
      Height          =   285
      Left            =   1665
      TabIndex        =   18
      Top             =   2475
      Width           =   1410
   End
   Begin VB.Label D 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dec:"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4140
      TabIndex        =   15
      Top             =   1800
      Width           =   510
   End
   Begin VB.Label DLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "00000000"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4635
      TabIndex        =   14
      Top             =   1800
      Width           =   825
   End
   Begin VB.Label H 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hex:"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4140
      TabIndex        =   13
      Top             =   1530
      Width           =   510
   End
   Begin VB.Label HLabel 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000000"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4635
      TabIndex        =   12
      Top             =   1530
      Width           =   825
   End
   Begin VB.Label Clabel 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4140
      TabIndex        =   11
      Top             =   1260
      Width           =   1320
   End
   Begin VB.Label B 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   4140
      TabIndex        =   9
      Top             =   990
      Width           =   555
   End
   Begin VB.Label G 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   4140
      TabIndex        =   8
      Top             =   720
      Width           =   555
   End
   Begin VB.Label R 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   4140
      TabIndex        =   7
      Top             =   450
      Width           =   555
   End
   Begin VB.Label ValB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4680
      TabIndex        =   6
      Top             =   990
      Width           =   510
   End
   Begin VB.Label ValG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4680
      TabIndex        =   5
      Top             =   720
      Width           =   510
   End
   Begin VB.Label VAlR 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "255"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4680
      TabIndex        =   4
      Top             =   450
      Width           =   510
   End
   Begin VB.Label BLabel 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5175
      TabIndex        =   3
      Top             =   990
      Width           =   285
   End
   Begin VB.Label GLabel 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5175
      TabIndex        =   2
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Rlabel 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5175
      TabIndex        =   1
      Top             =   450
      Width           =   285
   End
   Begin VB.Label Label1 
      Height          =   1950
      Left            =   4095
      TabIndex        =   10
      Top             =   135
      Width           =   1410
   End
End
Attribute VB_Name = "ColFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Code by Stephan Swertvaegher - Belgium, Europe
'stephan.swertvaegher@ planetinternet.be
'or
'gumming@compaqnet.be

Dim CPxx%, CPyy%, NewColor&
Dim RetR%, RetG%, RetB%
'if already declared as public, remove these API's
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long)

Private Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Private Type POINT
    x As Long
    y As Long
End Type

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
GetColorReturn = 0
ColFrm.Hide
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClipCursor ByVal 0&
End Sub

Private Sub SetBorders(OB As Object)
Line (OB.left - 1, OB.top - 1)-(OB.left + OB.Width, OB.top + OB.Height), &H808080, B
Line (OB.left - 1, OB.top + OB.Height)-(OB.left + OB.Width, OB.top + OB.Height), &HE0E0E0, B
Line (OB.left + OB.Width, OB.top)-(OB.left + OB.Width, OB.top + OB.Height), &HE0E0E0, B

Line (OB.left - 5, OB.top - 5)-(OB.left + OB.Width + 4, OB.top + OB.Height + 4), &HE0E0E0, B
Line (OB.left - 5, OB.top + OB.Height + 4)-(OB.left + OB.Width + 4, OB.top + OB.Height + 4), &H808080, B
Line (OB.left + OB.Width + 4, OB.top - 4)-(OB.left + OB.Width + 4, OB.top + OB.Height + 4), &H808080, B
End Sub

Private Sub Cancel_Click()
GetColorReturn = 0
ColFrm.Hide
End Sub

Private Sub GetColors()
RetR = NewColor Mod 256&
Rlabel.BackColor = RGB(RetR, 0, 0)
VAlR.Caption = Format(RetR, "000")
RetG = ((NewColor And &HFF00&) / 256&) Mod 256&
GLabel.BackColor = RGB(0, RetG, 0)
ValG.Caption = Format(RetG, "000")
RetB = (NewColor And &HFF0000) / 65536
BLabel.BackColor = RGB(0, 0, RetB)
ValB.Caption = Format(RetB, "000")
Clabel.BackColor = RGB(RetR, RetG, RetB)
HLabel.Caption = Hex(NewColor)
DLabel.Caption = NewColor
End Sub

Private Sub CPic1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Down_error
If Button = 1 Then
    Dim CL As RECT
    Dim U_Left As POINT
    GetClientRect CPic1.hWnd, CL
    U_Left.x = CL.left
    U_Left.y = CL.top
    ClientToScreen CPic1.hWnd, U_Left
    OffsetRect CL, U_Left.x, U_Left.y
    ClipCursor CL
NewColor = GetPixel(CPic1.hdc, x, y)
PickLabel.BackColor = &HC0C0C0
GetColors
End If
Down_error:
End Sub

Private Sub CPic1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Move_error
If Button = 1 Then
NewColor = GetPixel(CPic1.hdc, x, y)
GetColors
End If
Move_error:
End Sub

Private Sub CPic1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo Up_Error
PickLabel.BackColor = NewColor
Up_Error:
    ClipCursor ByVal 0&
End Sub

Private Sub Form_Activate()
CPic1.SetFocus
PickLabel.BackColor = &HC0C0C0
End Sub

Private Sub Form_Load()
CPic1.Move 6, 9, 256, 129
Label1.Height = 129

    SetBorders CPic1
    SetBorders Label1
    SetBorders PickLabel
    SetBorders OldLabel
        
For CPxx = 0 To 63
    For CPyy = 0 To 128
    SetPixel CPic1.hdc, CPxx, CPyy, RGB(CPxx * 4, 0, 2 * CPyy)
    SetPixel CPic1.hdc, CPxx + 64, CPyy, RGB(255, 4 * CPxx, 2 * CPyy)
    SetPixel CPic1.hdc, CPxx + 128, CPyy, RGB(255 - (4 * CPxx), 255, 2 * CPyy)
    SetPixel CPic1.hdc, CPxx + 192, CPyy, RGB(0, 255 - (4 * CPxx), 2 * CPyy)
Next CPyy
Next CPxx
End Sub

Private Sub Pick_Click()
GetColorReturn = 1
ColFrm.Hide
End Sub
