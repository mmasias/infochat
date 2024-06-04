VERSION 5.00
Begin VB.Form Fdata 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1965
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5115
   ControlBox      =   0   'False
   Icon            =   "Fdata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   5115
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Información de usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   10
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   2775
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   4920
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   4920
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   2
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Fdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Caption = "Quién es " & Finfo.Lista.SelectedItem
    say(1) = Finfo.Lista.SelectedItem.SubItems(1)
    say(3) = Finfo.Lista.SelectedItem.SubItems(2)
End Sub

Private Sub Ok_Click()
    Unload Me
End Sub

