VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{D6EEA3C0-6216-11CF-BE62-0080C72EDD2D}#1.0#0"; "MARQUEE.OCX"
Begin VB.Form Fpostal 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minipostal de Infochat"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   Icon            =   "fpower.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   324
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   405
   StartUpPosition =   1  'CenterOwner
   Begin MarqueeObjectsCtl.Marquee mrq 
      Height          =   1335
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   2355
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   2760
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Enviar"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ver lista"
      Height          =   375
      Index           =   0
      Left            =   1200
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
   Begin ComctlLib.TreeView Arbol 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   5530
      _Version        =   327682
      Indentation     =   265
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin InetCtlsObjects.Inet perl 
      Left            =   240
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label say 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Texto (Máximo 64 letras)"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   168
      X2              =   168
      Y1              =   16
      Y2              =   280
   End
   Begin VB.Line Line2 
      X1              =   384
      X2              =   384
      Y1              =   16
      Y2              =   280
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   8
      X2              =   384
      Y1              =   16
      Y2              =   16
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   8
      X2              =   384
      Y1              =   280
      Y2              =   280
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Texto (Máximo 64 letras)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lblpostal 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   3720
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Vista previa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   4
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label say 
      BackStyle       =   0  'Transparent
      Caption         =   "Postales"
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
      TabIndex        =   3
      Top             =   420
      Width           =   975
   End
End
Attribute VB_Name = "Fpostal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lprim As Integer
Dim rpta As String
Dim wstate As String
Dim categorias(30) As String
Dim postales(15) As String
Dim selec As Integer
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
Function verpostal(cpostal As String)
  mrq.insertURL 0, cpostal
End Function


Private Sub Arbol_BeforeLabelEdit(Cancel As Integer)
  Cancel = True
End Sub

Private Sub Arbol_Click()
  If Arbol.Nodes.Count = 0 Then Exit Sub
  If Arbol.SelectedItem.Children = 0 Then
    mrq.insertURL 0, curl & "/infochat/p" & Arbol.SelectedItem.Text & ".jpg"
    lblpostal = "<p>" & Arbol.SelectedItem.Text
  End If
End Sub


Private Sub cmdzoom_Click(Index As Integer)
Select Case Index
 Case 0
  If mrq.Zoom = 200 Then mrq.Zoom = 90
  mrq.Zoom = mrq.Zoom + 10
 Case 1
  If mrq.Zoom = 50 Then mrq.Zoom = 110
  mrq.Zoom = mrq.Zoom - 10
End Select
End Sub

Private Sub Form_Load()
Dim nodX As Node
  Caption = "Minipostal Infochat para " & Finfo.Lista.SelectedItem.Text
  Tag = Finfo.Lista.SelectedItem.Tag
  selec = Val(Tag)
  mrq.LoopsX = -3
  mrq.LoopsY = 0
  'mrq.Zoom = 10
  'mrq.insertURL 0, "http://127.0.0.1/infochat/matrimonio.gif"
  'nav.Navigate "http://127.0.0.1/infochat/matrimonio.gif"
  
  
End Sub

Private Sub mrq_OnStartOfImage()
  'say.Caption = "Cargando vista preliminar de postal" & Arbol.SelectedItem.Text
  'say.Refresh
End Sub



Private Sub Ok_Click(Index As Integer)
Dim cbus As String
Dim npos As Integer
Dim nodX As Node
Select Case Index
 Case 0
  Ok(0).Visible = False
  If Len(curl) = 0 Then curl = "http://127.0.0.1"
  Screen.MousePointer = 11
  cbus = "campo=hola"
  Perl.Execute curl & "/cgi-local/minipostal.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  Screen.MousePointer = 0
  din = Separar(rpta, "-", categorias())
  For i = 1 To din
    If Right(categorias(i), 1) <> ":" Then categorias(i) = categorias(i) & ":"
    npos = Separar(categorias(i), ":", postales())
    Set nodX = Arbol.Nodes.Add(, , postales(1), postales(1))
    nodX.Expanded = True
    For j = 2 To npos
     Set nodX = Arbol.Nodes.Add(postales(1), tvwChild, , postales(j))
    Next
  Next
  Ok(1).Visible = True
 Case 1
  If Len(lblpostal) = 0 Then Exit Sub
  If Finfo.ws(selec).State <> 7 Then
      MsgBox "Se perdió la conexión"
      Finfo.ws(selec).Close
      Unload Finfo.ws(selec)
      npos = Finfo.Buscatag(selec)
      Finfo.Lista.ListItems(npos).Ghosted = True
      Finfo.Lista.ListItems(npos).Tag = ""
     Else
       lblpostal = lblpostal & ":" & txt
       Finfo.ws(selec).SendData lblpostal
     End If
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

Private Sub txt_KeyPress(KeyAscii As Integer)
  If KeyAscii = 34 Then KeyAscii = 0
  Say(3) = Len(txt) & " Letra(s)"
  Say(3).Refresh
End Sub

Private Sub txt_LostFocus()
  If Len(txt) >= 64 Then
    MsgBox "Parte de su texto ha sido eliminado"
    Say(3) = "64 Letra(s)"
    Say(3).Refresh
    txt = Mid(txt, 1, 64)
  End If
End Sub
