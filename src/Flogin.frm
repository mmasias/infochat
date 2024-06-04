VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Flogin 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   -150
   ClientWidth     =   4095
   ControlBox      =   0   'False
   Icon            =   "Flogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Ok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Registrar"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   870
   End
   Begin VB.CommandButton Ok 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   1920
      Width           =   840
   End
   Begin InetCtlsObjects.Inet perl 
      Left            =   4440
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Ok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3060
      TabIndex        =   4
      Top             =   1920
      Width           =   870
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1680
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox txt 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   3960
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   3960
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Verificar Identidad"
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
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1215
      Width           =   1335
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Ifc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   735
      Width           =   1095
   End
End
Attribute VB_Name = "Flogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rpta As String
Dim wstate As Integer
Dim nintento As Integer
Function Separar(cadEnt As String, sep As String, ctodos() As String) As Integer
'Dim sep As String
Dim j As Integer
Dim cad As String
Dim i As Integer, p As Integer
'Dim cotro As Collection
  'sep = "_"
  i = 1
  j = 1
  cad = cadEnt
  Do While True
    p = InStr(j, cad, sep)
    ctodos(i) = Mid(cad, j, p - j)
    j = p + 1
    i = i + 1
    If j >= Len(cad) Then Exit Do
  Loop
  Separar = i - 1
End Function

Private Sub Form_Load()
 nintento = 1
 lprimero = 0
End Sub

Private Sub Ok_Click(Index As Integer)
Select Case Index

Case 0
  Dim ctodos(30) As String
  Dim cbus As String
    Ok(0).Enabled = False
    Ok(2).Enabled = False
    'curl = "http://pdeinfo.com"
    If Len(curl) = 0 Then curl = "http://127.0.0.1"
    If Len(txt(0)) = 0 Then Exit Sub
    If Len(txt(1)) = 0 Then Exit Sub
    Screen.MousePointer = 11
    cbus = "nick=" & txt(0) & "&paswd=" & txt(1)
    Perl.Execute curl & "/cgi-local/valida.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
    'Exit Sub
    wstate = 0
    Do While True
      If wstate = 1 Then Exit Do
      If (Perl.ResponseCode = 12029) Or (Perl.ResponseCode = 12007) Then
        MsgBox "No se puede encontrar el servidor infochat " & vbCrLf & "Necesita estar conectado a internet para usar este programa"
        Ok(0).Enabled = True
        Ok(2).Enabled = True
        Screen.MousePointer = 0
        Exit Sub
      End If
    DoEvents
    Loop
   Screen.MousePointer = 0
   If rpta = "Error" Then
    MsgBox "Lo siento " & txt(0) & " la autentificación falló" & vbCrLf & "Su nick o contraseña no son correctos " & vbCrLf & "Reíntente nuevamente"
    If nintento = 3 Then
      MsgBox txt(0) & " Quiza no esté registrado Ejecute El programa registrar "
      End
    End If
    nintento = nintento + 1
    Ok(0).Enabled = True
    txt(0).SetFocus
    Exit Sub
   Else
    din = Separar(rpta, ":", ctodos())
    Open App.Path & "\infochat.ini" For Output As #1
      Print #1, ctodos(1) & ":" & ctodos(2) & ":" & ctodos(3) & ":" & ctodos(4) & ":" & ctodos(5) & ":" & ctodos(6) & ":";
    Close #1
    lprimero = 2
   End If
   Unload Me
Case 1
  End
Case 2
  Fregistrar.Show 1
End Select

End Sub

Private Sub Perl_StateChanged(ByVal State As Integer)
 If State = 12 Then
    rpta = Perl.GetChunk(1024)
    wstate = 1
    'say = rpta
    'say.Refresh
    'MsgBox rpta
  End If
End Sub

