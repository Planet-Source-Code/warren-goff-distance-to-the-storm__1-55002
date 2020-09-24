VERSION 5.00
Begin VB.Form Distancestorm 
   BackColor       =   &H001047EA&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Simpleton's Storm calculator"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Caption         =   "Reset"
      Height          =   255
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2055
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   -675
      Top             =   1965
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   90
      Picture         =   "Distancestorm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   75
      Width           =   2970
   End
   Begin VB.Image Image2 
      Height          =   4800
      Left            =   4215
      Picture         =   "Distancestorm.frx":2B85A
      Top             =   1020
      Visible         =   0   'False
      Width           =   6405
   End
   Begin VB.Image Image1 
      Height          =   2970
      Left            =   3945
      Picture         =   "Distancestorm.frx":8FD9C
      Top             =   105
      Visible         =   0   'False
      Width           =   4485
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click the above at the instant of the lightning flash"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1230
      Left            =   -15
      TabIndex        =   1
      Top             =   1185
      Width           =   3180
   End
End
Attribute VB_Name = "Distancestorm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Single
Private Sub Command1_Click()
If Label1.Caption = "Click the above at the instant of the lightning flash" Then
    Timer1.Enabled = True
    Label1.Caption = "Click the above at the instant of the Thunder"
    Command1.Picture = Image2.Picture
    Distancestorm.Picture = Image2.Picture
    Exit Sub
End If
If Label1.Caption = "Click the above at the instant of the Thunder" Then
    Label1.Caption = "The Storm is " & X & " mile(s) away"
    Distancestorm.Picture = Image1.Picture
    Command1.Picture = Image1.Picture
    Command2.Visible = True
    Timer1.Enabled = False
    X = 0
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
    Command2.Visible = False
    Label1.Caption = "Click the above at the instant of the lightning flash"
End Sub

Private Sub Form_Load()
Distancestorm.Picture = Image1.Picture
Me.Top = (Screen.Height - Me.Height) / 2
Me.Left = (Screen.Width - Me.Width) / 2
X = 0
End Sub

Private Sub Timer1_Timer()
    X = X + 0.21
    'Me.Caption = Str(X)
End Sub
