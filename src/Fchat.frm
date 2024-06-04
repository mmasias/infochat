VERSION 5.00
Begin VB.Form Fchat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventana de Chat"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Fchat.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   318
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   312
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton carita 
      Caption         =   ":o"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   3360
      TabIndex        =   12
      ToolTipText     =   "Sorprendido"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton carita 
      Caption         =   ";)"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2880
      TabIndex        =   11
      ToolTipText     =   "Guiño"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton carita 
      Caption         =   ":/"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2400
      TabIndex        =   10
      ToolTipText     =   "Molesto"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton carita 
      Caption         =   ":D"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   9
      ToolTipText     =   "Riendo"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton carita 
      Caption         =   ":Þ"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1440
      TabIndex        =   8
      ToolTipText     =   "Sacando la lengua"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton carita 
      Caption         =   ":("
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   7
      ToolTipText     =   "Triste"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton carita 
      Caption         =   ":)"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "Sonriendo"
      Top             =   4320
      Width           =   435
   End
   Begin VB.CommandButton Ok 
      Caption         =   "&Abrir en Editor"
      Height          =   315
      Index           =   0
      Left            =   3180
      TabIndex        =   4
      Top             =   255
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   300
      Top             =   3285
   End
   Begin VB.TextBox txt 
      Height          =   300
      Index           =   2
      Left            =   0
      TabIndex        =   0
      Top             =   3960
      Width           =   4680
   End
   Begin VB.TextBox txt 
      ForeColor       =   &H80000002&
      Height          =   285
      Index           =   1
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   4680
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Index           =   0
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   4680
   End
   Begin VB.Label say 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1155
      TabIndex        =   5
      Top             =   315
      Width           =   3420
   End
   Begin VB.Label pdeinfo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   4680
      TabIndex        =   3
      Top             =   0
      Width           =   7995
   End
End
Attribute VB_Name = "Fchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim selec As Integer
Dim hora
Dim banner(11) As String

Private Sub carita_Click(Index As Integer)
    txt(2) = txt(2) & " " & carita(Index).Caption & " "
    txt(2).SelStart = Len(txt(2))
    txt(2).SetFocus
End Sub

Private Sub Form_Activate()
 selec = Val(Tag)
End Sub



Private Sub Ok_Click(Index As Integer)
Dim nf As Integer
Select Case Index
 Case 0
  nf = FreeFile
  Screen.MousePointer = 11
    Open App.Path & "\" & Caption & ".txt" For Output As #nf
      Print #1, txt(0)
    Close #nf
  Screen.MousePointer = 0
  din = Shell("notepad " & App.Path & "\" & Caption & ".txt", vbNormalFocus)
  Kill App.Path & "\" & Caption & ".txt"
  'say = "Se creo " & App.Path & "\" & Caption & ".txt"
  'say.Refresh
End Select
End Sub

Private Sub Timer1_Timer()
Dim cquien  As String
Dim minutos As String
Dim nmin As Integer

  If Len(hora) = 0 Then
    hora = Now
    For i = 1 To 10
      banner(i) = anuncios(i)
      If InStr(anuncios(i), vbLf) Then
        banner(i) = Mid(anuncios(i), 2)
        d = 2
      End If
    Next
    banner(11) = Finfo.pdeinfo
    pdeinfo = banner(Int((11 * Rnd) + 1))
  End If
  If Len(cchat) > 0 Then
    cquien = Mid(cchat, 1, InStr(cchat, ":") - 1)
    If cquien <> Caption Then Exit Sub
    cchat = Mid(cchat, InStr(cchat, ":") + 1)
    txt(0) = txt(0) & "[" & Caption & "] " & cchat & Chr(13) & Chr(10)
    txt(0).SelStart = Len(txt(0))
    'txt(1) = txt(1) & cchat & Chr(13) & Chr(10)
    'txt(1).SelStart = Len(txt(1))
    'txt(2).SetFocus
    cchat = ""
  End If
  pdeinfo.Left = pdeinfo.Left - 1
  'pdeinfo.Refresh
  If pdeinfo.Left = -pdeinfo.Width Then
    pdeinfo.Left = pdeinfo.Width
    pdeinfo = banner(Int((11 * Rnd) + 1))
  End If
  minutos = Format(Now - horaini, "hh:mm:ss")
  nmin = Val(Mid(minutos, 4, 2))
  If ((nmin Mod 5) = 0) And nmin > 0 Then
    For i = 1 To 10
      If InStr(anuncios(i), vbLf) Then
        banner(i) = Mid(anuncios(i), 2)
        d = 2
      End If
      banner(i) = anuncios(i)
    Next
    banner(11) = Finfo.pdeinfo
    horaini = Now
  End If
  DoEvents
End Sub


Private Sub txt_GotFocus(Index As Integer)
Select Case Index
  Case 0
   
    'txt(2).SetFocus
End Select
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
Dim npos As Integer
Select Case Index
 Case 2
  If KeyAscii = 60 Or KeyAscii = 62 Then KeyAscii = 0
  If KeyAscii = 13 And Len(Trim(txt(2))) > 0 Then
    If Finfo.ws(selec).State <> 7 Then
      MsgBox "Se perdió la conexión"
      txt(2).Enabled = False
      Finfo.ws(selec).Close
      Unload Finfo.ws(selec)
      npos = Finfo.Buscatag(selec)
      Finfo.Lista.ListItems(npos).Ghosted = True
      Finfo.Lista.ListItems(npos).Tag = ""
      'Unload Me
      'Exit Sub
     Else
       Finfo.ws(selec).SendData txt(2)
       txt(0) = txt(0) & "[" & login & "] " & txt(2) & Chr(13) & Chr(10)
       txt(0).SelStart = Len(txt(0))
       txt(2) = ""
     End If
  End If
End Select
End Sub
