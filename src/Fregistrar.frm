VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Fregistrar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InfoChat - Registro"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   Icon            =   "Fregistrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   4980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chk 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Permitir que otros usuarios vean mi email"
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
      Left            =   120
      TabIndex        =   20
      Tag             =   "no"
      Top             =   3720
      Value           =   1  'Checked
      Width           =   4695
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Tag             =   "si"
      Top             =   1185
      Width           =   3615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   1200
      TabIndex        =   2
      Tag             =   "si"
      Top             =   1605
      Width           =   3615
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Tag             =   "si"
      ToolTipText     =   "El ifc debe ser al menos de 4 caracteres"
      Top             =   765
      Width           =   1095
   End
   Begin VB.TextBox txt 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   1200
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   6
      Tag             =   "si"
      Top             =   3285
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   7
      Left            =   3600
      MaxLength       =   5
      PasswordChar    =   "*"
      TabIndex        =   7
      Tag             =   "si"
      Top             =   3285
      Width           =   1215
   End
   Begin InetCtlsObjects.Inet Perl 
      Left            =   120
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton Ok 
      Appearance      =   0  'Flat
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   2460
      TabIndex        =   9
      Top             =   4440
      Width           =   1155
   End
   Begin VB.CommandButton Ok 
      Appearance      =   0  'Flat
      Caption         =   "&Registrar"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   8
      Top             =   4440
      Width           =   1155
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   5
      Left            =   1200
      TabIndex        =   5
      Tag             =   "si"
      Top             =   2865
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   4
      Left            =   1200
      TabIndex        =   4
      Tag             =   "si"
      Top             =   2445
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   3
      Left            =   1200
      TabIndex        =   3
      Top             =   2025
      Width           =   1215
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   945
      Left            =   3960
      Picture         =   "Fregistrar.frx":030A
      Stretch         =   -1  'True
      Tag             =   "http://pdeinfo.com"
      ToolTipText     =   "DobleClick aquí"
      Top             =   60
      Width           =   960
   End
   Begin VB.Line Line2 
      X1              =   -1680
      X2              =   3960
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro en el servidor InfoChat"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   60
      TabIndex        =   19
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   -120
      X2              =   4980
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Label Say 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   18
      ToolTipText     =   "El Ifc es el equivalente al nick en otros programas"
      Top             =   5040
      Width           =   4815
   End
   Begin VB.Label Say 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmar Contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   2520
      TabIndex        =   17
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label Say 
      BackStyle       =   0  'Transparent
      Caption         =   "Contraseña"
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
      Left            =   120
      TabIndex        =   16
      Top             =   3300
      Width           =   1095
   End
   Begin VB.Label Say 
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
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Say 
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
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   2460
      Width           =   735
   End
   Begin VB.Label Say 
      BackStyle       =   0  'Transparent
      Caption         =   "Teléfono"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   13
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Say 
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
      Left            =   120
      TabIndex        =   12
      Top             =   1620
      Width           =   735
   End
   Begin VB.Label Say 
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
      Index           =   1
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   735
   End
   Begin VB.Label Say 
      BackStyle       =   0  'Transparent
      Caption         =   "Ifc"
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
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "El Ifc es el equivalente al nick en otros programas"
      Top             =   780
      Width           =   495
   End
End
Attribute VB_Name = "Fregistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rpta As String
Dim wstate As String
Dim curl As String

Private Sub chk_Click()
  If chk.Value = 1 Then
    chk.Tag = "no"
  Else
    chk.Tag = "si"
  End If
End Sub

Private Sub Form_Load()
  'curl = "http://127.0.0.1"
  curl = "http://pdeinfo.com"
End Sub

Private Sub Ok_Click(Index As Integer)
Dim n As Integer, blanco As Integer
Dim cbus As String
Select Case Index
  Case 0
    blanco = False
    n = txt.Count - 1
    If txt(6) <> txt(7) Then
      MsgBox "Las contraseñas no coinciden"
      Exit Sub
    End If
    For i = 0 To n
      If (Len(txt(i).Tag) > 0 And Len(txt(i)) = 0) Then blanco = True: Exit For
    Next
    If blanco Then MsgBox "Debe llenar todos los campos cuyos nombres se encuentren en negrita": Exit Sub
    Screen.MousePointer = 11
    cbus = "nick=" & txt(0) & "&nombres=" & txt(1) & "&email=" & txt(2) & "&tel=" & txt(3) & "&ciudad=" & txt(4) & "&pais=" & txt(5) & "&paswd=" & txt(6) & "&oculto=" & chk.Tag
    Perl.Execute curl & "/cgi-local/regnick.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
    wstate = 0
    Do While True
      If wstate = 1 Then Exit Do
      If (Perl.ResponseCode = 12029) Or (Perl.ResponseCode = 12007) Then
        MsgBox "No se puede conectar con el servidor" & vbCrLf & "Reinténtelo más tarde"
        Ok(0).Enabled = True
        Screen.MousePointer = 0
        Exit Sub
      End If
    DoEvents
    Loop
    If rpta = "Error" Then
      MsgBox "El ifc " & txt(0) & " Ya está en uso elija otro"
      Screen.MousePointer = 0
    Else
      MsgBox "Fue registrado Satisfactoriamente en el servidor, Gracias Por su preferencia"
      Screen.MousePointer = 0
      Open App.Path & "\Infochat.ini" For Output As #1
        Print #1, txt(0) & ":" & txt(1) & ":" & txt(2) & ":" & txt(3) & ":" & txt(4) & ":" & txt(5) & ":"
      Close #1
      Unload Me
    
    End If
  Case 1
    Unload Me
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

Private Sub txt_GotFocus(Index As Integer)
 Select Case Index
    Case 0
      Say(8) = "El ifc es su código de usuario (al menos de 4 caracteres)"
    Case 1
      Say(8) = "Ingrese sus nombres completos"
    Case 2
      Say(8) = "Ingrese su email"
    Case 3
      Say(8) = "Ingrese su número de teléfono si lo desea"
    Case 4
      Say(8) = "Ingrese la ciudad desde donde se conecta"
    Case 5
      Say(8) = "Ingrese el país desde donde se conecta"
    Case 6
      Say(8) = "Ingrese su contraseña"
    Case 7
      Say(8) = "Reescriba su contraseña"
 End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
  Select Case Index
    Case 0, 6, 7
      KeyAscii = Asc(LCase(Chr(KeyAscii)))
  End Select
End Sub

Private Sub txt_LostFocus(Index As Integer)
  Select Case Index
  Case 0
    If Len(txt(0)) < 4 Then txt(0) = ""
  Case 2
    If InStr(txt(Index), "@") = 0 Then txt(Index) = ""
    If InStr(txt(Index), ".") = 0 Then txt(Index) = ""
End Select

End Sub
