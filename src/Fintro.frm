VERSION 5.00
Begin VB.Form Fintro 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2100
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3780
   ControlBox      =   0   'False
   Icon            =   "Fintro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tiempo 
      Interval        =   100
      Left            =   3240
      Top             =   4680
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Punto de Información http://pdeinfo.com"
      Height          =   495
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Derechos reservados Piura 1999"
      Height          =   495
      Index           =   0
      Left            =   1680
      TabIndex        =   2
      Top             =   600
      Width           =   1815
   End
   Begin VB.Line Line3 
      X1              =   3600
      X2              =   3600
      Y1              =   480
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   1680
      X2              =   3600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   3600
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Por Favor Espere"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   3015
      WordWrap        =   -1  'True
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "InfoChat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   6
      Left            =   4920
      Picture         =   "Fintro.frx":0442
      Top             =   4080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   5
      Left            =   4920
      Picture         =   "Fintro.frx":0D23
      Top             =   4080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   4
      Left            =   4920
      Picture         =   "Fintro.frx":1706
      Top             =   4080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   3
      Left            =   4920
      Picture         =   "Fintro.frx":230A
      Top             =   4080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   2
      Left            =   4920
      Picture         =   "Fintro.frx":30FF
      Top             =   4080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   1
      Left            =   4920
      Picture         =   "Fintro.frx":41E4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image fig 
      Height          =   1545
      Index           =   0
      Left            =   0
      Picture         =   "Fintro.frx":5C9B
      Top             =   0
      Width           =   1500
   End
End
Attribute VB_Name = "Fintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub Form_Load()
  i = 1
End Sub

Private Sub Form_LostFocus()
  If lprimero = 1 Then
    SetFocus
  Else
    Unload Me
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'End
End Sub

Private Sub tiempo_Timer()
  fig(0).Picture = fig(i).Picture
  i = i + 1
  If i = 7 Then i = 1
  If lprimero = 0 Then Unload Me
  DoEvents
End Sub
