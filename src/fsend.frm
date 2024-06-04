VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Fmail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar correo"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   Icon            =   "fsend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Ok 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   4200
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   105
      Top             =   4980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4440
      TabIndex        =   7
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   1605
      Index           =   3
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2325
      Width           =   5535
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   1080
      TabIndex        =   0
      Top             =   1680
      Width           =   4695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1080
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   4695
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   5880
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   5880
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   5880
      Y1              =   480
      Y2              =   480
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
      TabIndex        =   9
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      Caption         =   "Asunto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   6
      Top             =   1725
      Width           =   660
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      Caption         =   "Para:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1245
      Width           =   465
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      Caption         =   "De:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   765
      Width           =   315
   End
End
Attribute VB_Name = "Fmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sw As Integer
Private Sub Command1_Click()
Dim cad As String
  If Len(txt(2)) = 0 Then Exit Sub
  If Len(txt(3)) = 0 Then Exit Sub
  Screen.MousePointer = 11
  sw = False
  ws.Connect
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  sw = False
  cad = "mail from:<" & txt(0) & ">" & Chr(13) & Chr(10)
  ws.SendData cad
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  sw = False
  cad = "rcpt to:<" & txt(1) & ">" & Chr(13) & Chr(10)
  ws.SendData cad
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  sw = False
  cad = "data" & Chr(13) & Chr(10)
  ws.SendData cad
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  sw = False
  cad = "From: " & txt(0) & Chr(13) & Chr(10)
  cad = cad & "To: " & txt(1) & Chr(13) & Chr(10)
  cad = cad & "Subject : " & txt(2) & Chr(13) & Chr(10)
  cad = cad & txt(3) & Chr(13) & Chr(10)
  cad = cad & "Mensaje enviado usando Infochat http://pdeinfo.com"
  cad = cad & Chr(13) & Chr(10) & "." & Chr(13) & Chr(10)
  ws.SendData cad
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  Screen.MousePointer = 0
  MsgBox "Su mensaje fue enviado"
  ws.Close
  Unload Me
End Sub

Private Sub Form_Load()
  sw = False
  ws.RemoteHost = "pdeinfo.com"
  ws.RemotePort = 25
  txt(0) = cemail
  txt(1) = Finfo.Lista.SelectedItem.SubItems(2)
End Sub

Private Sub Ok_Click()
  Unload Me
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim cad As String
  ws.GetData cad, vbString
  'MsgBox cad
  sw = True
End Sub

