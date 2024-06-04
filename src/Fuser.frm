VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Fuser 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   6195
   ControlBox      =   0   'False
   Icon            =   "Fuser.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock ws 
      Left            =   480
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lemail 
      Height          =   1035
      Left            =   1680
      TabIndex        =   18
      Top             =   3600
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   4920
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Cerrar"
      Height          =   495
      Index           =   3
      Left            =   4320
      TabIndex        =   15
      Top             =   3480
      Width           =   1695
   End
   Begin InetCtlsObjects.Inet perl 
      Left            =   4680
      Top             =   5280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ComctlLib.ListView Lusuarios 
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   2143
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ifc"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Nombres"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Email"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.TextBox txtbusca 
      Height          =   315
      Left            =   2160
      TabIndex        =   10
      Top             =   1170
      Width           =   2535
   End
   Begin VB.ComboBox Lista 
      Height          =   315
      ItemData        =   "Fuser.frx":030A
      Left            =   2160
      List            =   "Fuser.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5280
      TabIndex        =   8
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Cancelar"
      Height          =   375
      Index           =   1
      Left            =   6720
      TabIndex        =   7
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   2
      Left            =   4920
      TabIndex        =   6
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   7920
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   1
      Left            =   840
      TabIndex        =   3
      Top             =   6240
      Width           =   2895
   End
   Begin VB.TextBox txt 
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   1
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Haga Doble click sobre el ifc para agregarlo"
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
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Top             =   3120
      Width           =   3840
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   120
      X2              =   6120
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   120
      X2              =   6120
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   120
      X2              =   6120
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Contactos"
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
      TabIndex        =   14
      Top             =   120
      Width           =   2715
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor"
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
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label say 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo de búsqueda"
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
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   1740
   End
   Begin VB.Label say 
      Caption         =   "Email"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label say 
      Caption         =   "Nombres"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   6360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label say 
      Caption         =   "Nick"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   0
      Top             =   6000
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "Fuser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rpta As String
Dim wstate As String
Dim ctodos(10) As String
Dim sw As Integer
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

Function valida() As Integer
Dim cad As String
Dim enc As Integer
  'If Len(txt(0)) = 0 Then valida = 0: Exit Function
  'Verificar duplicados
  enc = False
  Open App.Path & "\usuarios.dat" For Input As #1
  Do While Not EOF(1)
    Line Input #1, cad
    din = Separa(cad, ":")
    If usuario(1) = txt(0) Then enc = True: Exit Do
  Loop
  Close #1
  If enc Then valida = 0: Exit Function
  'If Len(txt(1)) = 0 Then valida = 0: Exit Function
  'If (Len(txt(2)) = 0 Or InStr(txt(2), "@") = 0) Then valida = 0: Exit Function
  valida = 1
End Function

Private Sub cmdbuscar_Click()
  Call txtbusca_KeyPress(13)
End Sub

Private Sub Form_Load()
  Lista.AddItem "Ifc"
  Lista.AddItem "nombres"
  Lista.AddItem "email"
  Lista.ListIndex = 0
  
End Sub



Private Sub Lusuarios_BeforeLabelEdit(Cancel As Integer)
  Cancel = True
End Sub

Private Sub Lusuarios_DblClick()
Dim cad As String
Dim x As ListItem
Dim l As Integer
  
  If Lusuarios.ListItems.Count = 0 Then Exit Sub
  If MsgBox("Desea Adicionar", vbYesNo) <> vbYes Then Exit Sub
  txt(0) = Lusuarios.SelectedItem.Text
  If valida() = 0 Then MsgBox "Ya existe en su lista de contactos :-(": Exit Sub
  Open App.Path & "\usuarios.dat" For Append As #1
  Print #1, Lusuarios.SelectedItem.Text & ":" & Lusuarios.SelectedItem.SubItems(1) & ":" & Lusuarios.SelectedItem.SubItems(2) & ":"
  Close #1
  l = Finfo.Lista.ListItems.Count + 1
  If l > 10 Then l = 1
  Set x = Finfo.Lista.ListItems.Add(, , txt(0), l, l)
  x.SubItems(1) = Lusuarios.SelectedItem.SubItems(1)
  x.SubItems(2) = Lusuarios.SelectedItem.SubItems(2)
  x.Ghosted = True
  Screen.MousePointer = 11
  sw = False
  ws.RemoteHost = "pdeinfo.com"
  ws.RemotePort = 25
  ws.Connect
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  sw = False
  cad = "mail from:<infochat@pdeinfo.com>" & Chr(13) & Chr(10)
  ws.SendData cad
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  sw = False
  cad = "rcpt to:<" & lemail.List(Lusuarios.SelectedItem.Index - 1) & ">" & Chr(13) & Chr(10)
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
  cad = "From: infochat@pdeinfo.com" & Chr(13) & Chr(10)
  cad = cad & "To: " & lemail.List(Lusuarios.SelectedItem.Index - 1) & Chr(13) & Chr(10)
  cad = cad & "Subject : Notificación infochat" & Chr(13) & Chr(10)
  cad = cad & "El usuario con Ifc " & login & " lo ha agregado como uno de sus contactos " & Chr(13) & Chr(10)
  cad = cad & "Mensaje enviado usando Infochat http://pdeinfo.com/infochat" & vbCrLf
  cad = cad & "Este es un mensaje autogenerado por favor no responder"
  cad = cad & Chr(13) & Chr(10) & "." & Chr(13) & Chr(10)
  ws.SendData cad
  Do While Not sw
    If sw Then Exit Do
  DoEvents
  Loop
  ws.Close
  Screen.MousePointer = 0
  MsgBox txt(0) & " Fue agregado a su base de datos"
End Sub



Private Sub Ok_Click(Index As Integer)
Dim x As ListItem
Dim l As Integer
  Select Case Index
    Case 0 'Grabar
     If valida() = 0 Then MsgBox "Verifique que ha llenado correctamente TODOS los campos" & vbCrLf & "El nick debe ser único para cada contacto": Exit Sub
     Open App.Path & "\usuarios.dat" For Append As #1
      Print #1, txt(0) & ":" & txt(1) & ":" & txt(2) & ":"
     Close #1
      l = Finfo.Lista.ListItems.Count + 1
      Set x = Finfo.Lista.ListItems.Add(, , txt(0), l, l)
      x.SubItems(1) = txt(1)
      x.SubItems(2) = txt(2)
      x.Ghosted = True
    Case 1 'Cancelar
    Case 3
      Unload Me
  End Select
  Unload Me
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

Private Sub txtbusca_KeyPress(KeyAscii As Integer)
Dim cbus As String
Dim x As ListItem
Dim cpartes(10) As String
  If Len(txtbusca) >= 4 Then
    If KeyAscii = 13 Then
      If Screen.MousePointer = 11 Then Exit Sub
      Screen.MousePointer = 11
      Dim cBusqueda As String
      cBusqueda = IIf(Lista.ListIndex = 0, "nick", IIf(Lista.ListIndex = 1, "nombres", "email"))
      cbus = "campo=" & cBusqueda & "&valor=" & txtbusca
      Perl.Execute curl & "/cgi-local/encuentra.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
      wstate = 0
      Do While True
        If wstate = 1 Then Exit Do
      DoEvents
      Loop
      Screen.MousePointer = 0
        If rpta = "Error" Then
          MsgBox "No se encontró ninguna coincidencia"
          lemail.Clear
          Lusuarios.ListItems.Clear
        Else
          Lusuarios.ListItems.Clear
          lemail.Clear
          din = Separar(rpta, "^", ctodos())
          For i = 1 To din
            If Right(ctodos(i), 1) <> ":" Then ctodos(i) = ctodos(i) & ":"
            din = Separar(ctodos(i), ":", cpartes())
            Set x = Lusuarios.ListItems.Add(, , cpartes(1))
              x.SubItems(1) = cpartes(2)
              lemail.AddItem cpartes(3)
              If cpartes(8) = "si" Then
                x.SubItems(2) = cpartes(3)
              Else
                x.SubItems(2) = ""
              End If
          Next
          'MsgBox rpta
        End If
    End If
  End If
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
Dim cad As String
  ws.GetData cad, vbString
  'MsgBox cad
  sw = True
End Sub

