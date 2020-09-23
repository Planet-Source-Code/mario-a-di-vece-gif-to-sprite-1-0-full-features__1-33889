VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  General Options and Settings"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   3435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1980
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   540
      TabIndex        =   6
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sprite Saving Options"
      Height          =   2055
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   3135
      Begin VB.HScrollBar HScroll1 
         Height          =   255
         Left            =   900
         Max             =   100
         Min             =   30
         TabIndex        =   3
         Top             =   1140
         Value           =   90
         Width           =   1875
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "90%"
         Height          =   195
         Left            =   1740
         TabIndex        =   5
         Top             =   1500
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quality:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Format:"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   540
         Width           =   525
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
    If Combo1.ListIndex = 0 Then
        HScroll1.Enabled = True
    Else
        HScroll1.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    SaveMode = Combo1.Text
    JPEGQuality = HScroll1.Value
    Unload Me
End Sub

Private Sub Form_Load()
    Combo1.AddItem "JPG"
    Combo1.AddItem "BMP"
    Combo1.ListIndex = 0
    HScroll1.Value = JPEGQuality
    If SaveMode = "JPG" Then
        Combo1.ListIndex = 0
    Else
        Combo1.ListIndex = 1
    End If
End Sub

Private Sub HScroll1_Change()
    Label3.Caption = HScroll1.Value & "%"
End Sub
