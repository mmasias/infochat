VERSION 5.00
Begin VB.Form Fguest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2685
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4440
   ControlBox      =   0   'False
   Icon            =   "Fguest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4440
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Ok 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   2160
      Width           =   975
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4320
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4320
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      Height          =   255
      Index           =   9
      Left            =   1200
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "País"
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
      Index           =   8
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      Height          =   255
      Index           =   7
      Left            =   1200
      TabIndex        =   7
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Ciudad"
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
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      Height          =   255
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Top             =   1200
      Width           =   2535
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
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      Height          =   255
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombres"
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
      TabIndex        =   2
      Top             =   960
      Width           =   735
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   975
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Nick"
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
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "Fguest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  Say(1) = usuarios(1)
  Say(3) = usuarios(2)
  Say(5) = usuarios(3)
  Say(7) = usuarios(5)
  Say(9) = usuarios(6)
End Sub

Private Sub Ok_Click()
  Unload Me
End Sub

