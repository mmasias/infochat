VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Finfo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "InfoChat"
   ClientHeight    =   5730
   ClientLeft      =   495
   ClientTop       =   330
   ClientWidth     =   6465
   Icon            =   "Finfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   382
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   431
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView Lista 
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4471
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   2385038
      BackColor       =   14089214
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Buu"
         Object.Width           =   2469
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Completo"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "email"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "ip"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ComboBox Lhistoria 
      Height          =   315
      Left            =   2175
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   4860
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
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
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   5400
      Width           =   6330
   End
   Begin VB.CommandButton cmdpdeinfo 
      Caption         =   "Pdeinfo"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer tiempo 
      Left            =   3000
      Top             =   5640
   End
   Begin VB.CommandButton cmdpubli 
      Caption         =   "Publicidad"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2520
      TabIndex        =   9
      Top             =   5880
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frame 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1680
      TabIndex        =   7
      Top             =   1320
      Width           =   4695
      Begin VB.Label publicidad 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Soluciones creativas en el web"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   840
         Width           =   4695
      End
   End
   Begin MSWinsockLib.Winsock ws 
      Index           =   0
      Left            =   4200
      Top             =   7320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdlogin 
      Caption         =   "Logearse"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2640
      TabIndex        =   6
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton ok 
      Caption         =   "chequear"
      Enabled         =   0   'False
      Height          =   255
      Left            =   3720
      TabIndex        =   5
      Top             =   5640
      Visible         =   0   'False
      Width           =   975
   End
   Begin InetCtlsObjects.Inet perl 
      Left            =   3600
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Image imgboton 
      Height          =   1140
      Left            =   240
      Picture         =   "Finfo.frx":030A
      Top             =   6000
      Width           =   1425
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0080C0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   -540
      Top             =   5280
      Width           =   7215
   End
   Begin VB.Label lbladic 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Adicionar Contacto"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      MouseIcon       =   "Finfo.frx":58CC
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Tag             =   "Envía Postales Virtuales con motivos regionales (acompañados con música de fondo de la región) a todo el mundo"
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblinfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2280
      TabIndex        =   22
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "Postales"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   2400
      MouseIcon       =   "Finfo.frx":5D0E
      MousePointer    =   99  'Custom
      TabIndex        =   21
      Tag             =   "Envía Postales Virtuales con motivos regionales (acompañados con música de fondo de la región) a todo el mundo"
      ToolTipText     =   "http://www.pdeinfo.com/postales"
      Top             =   2640
      Width           =   1155
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Anuncios"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   1
      Left            =   3720
      MouseIcon       =   "Finfo.frx":6150
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Tag             =   "Publica tus anuncios en cualquiera de nuestras secciones clasificadas, de manera gratuita y durante el tiempo que sea necesario."
      ToolTipText     =   "http://pdeinfo.com/anuncios"
      Top             =   2640
      Width           =   1080
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Currículos"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   4920
      MouseIcon       =   "Finfo.frx":6592
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Tag             =   $"Finfo.frx":69D4
      ToolTipText     =   "http://pdeinfo.com/curriculos"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H0000C0C0&
      Caption         =   "Diseño de Sitios Web"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   3
      Left            =   2400
      MouseIcon       =   "Finfo.frx":6AAD
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Tag             =   "Deja que nosotros nos encarguemos de tu presencia en el web"
      ToolTipText     =   "http://pdeinfo.com/web"
      Top             =   2940
      Width           =   1845
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H002B8CAD&
      Caption         =   "Compras vía Internet"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   4
      Left            =   4290
      MouseIcon       =   "Finfo.frx":6EEF
      MousePointer    =   99  'Custom
      TabIndex        =   17
      Tag             =   "Disfruta comprando desde la comodidad del lugar en el cual estes accediendo al web y obtén tus productos en la puerta de tu casa."
      ToolTipText     =   "http://pdeinfo.com/compras"
      Top             =   2940
      Width           =   1845
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Listas de Interés"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   5
      Left            =   2400
      MouseIcon       =   "Finfo.frx":7331
      MousePointer    =   99  'Custom
      TabIndex        =   16
      Tag             =   "Intégrese y discuta alrededor de temas comunes por medio de nuestras listas de interés"
      ToolTipText     =   "http://pdeinfo.com/listas"
      Top             =   3240
      Width           =   1845
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Novedades"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   6
      Left            =   3720
      MouseIcon       =   "Finfo.frx":7773
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Tag             =   $"Finfo.frx":7BB5
      ToolTipText     =   "http://pdeinfo.com/novedades"
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Correo Electrónico"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   7
      Left            =   4290
      MouseIcon       =   "Finfo.frx":7C5C
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Tag             =   "Si aún no tienes correo gratuito intégrate en un nuevo correo gratuito via web"
      ToolTipText     =   "http://pdeinfo.com/email"
      Top             =   3240
      Width           =   1845
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H0024648E&
      Caption         =   "Ezine - p.a."
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   8
      Left            =   2400
      MouseIcon       =   "Finfo.frx":809E
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Tag             =   "La primera magazine electrónica diseñada por jóvenes del norte del país"
      ToolTipText     =   "http://pdeinfo.com/pa"
      Top             =   3540
      Width           =   1155
   End
   Begin VB.Label lblservicios 
      Alignment       =   2  'Center
      BackColor       =   &H00800080&
      Caption         =   "Familia PdeInfo"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   9
      Left            =   4920
      MouseIcon       =   "Finfo.frx":84E0
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Tag             =   "Inscríbete y participa de forma activa y continua en las actividades que desarrollamos"
      ToolTipText     =   "http://pdeinfo.com/familiapdeinfo"
      Top             =   3540
      Width           =   1215
   End
   Begin VB.Image fig 
      Height          =   480
      Index           =   4
      Left            =   5760
      Picture         =   "Finfo.frx":8922
      Top             =   6840
      Visible         =   0   'False
      Width           =   480
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
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7995
   End
   Begin VB.Image fig 
      Height          =   480
      Index           =   5
      Left            =   6240
      Picture         =   "Finfo.frx":8C2C
      Top             =   6360
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fig 
      Height          =   480
      Index           =   2
      Left            =   5040
      Picture         =   "Finfo.frx":8F36
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fig 
      Height          =   480
      Index           =   3
      Left            =   5700
      Picture         =   "Finfo.frx":9240
      Top             =   6315
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fig 
      Height          =   480
      Index           =   1
      Left            =   4440
      Picture         =   "Finfo.frx":954A
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image fig 
      Height          =   480
      Index           =   0
      Left            =   3840
      Picture         =   "Finfo.frx":998C
      Tag             =   "Esperando conexión"
      Top             =   6240
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label say 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   6135
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":9DCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":9EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":9FF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A104
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A216
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A328
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A43A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A54C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A65E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A770
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Finfo.frx":A882
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0080C0FF&
      X1              =   -12
      X2              =   448
      Y1              =   152
      Y2              =   152
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Soluciones creativas en el web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   3495
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Punto de Información"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Descubra nuestra concepción del Web"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   0
      Top             =   240
      Width           =   6495
   End
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   0
      Picture         =   "Finfo.frx":A994
      Tag             =   "http://pdeinfo.com"
      ToolTipText     =   "DobleClick aquí"
      Top             =   720
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      X1              =   84
      X2              =   464
      Y1              =   80
      Y2              =   80
   End
   Begin VB.Menu Chat 
      Caption         =   "Chat"
      Visible         =   0   'False
      Begin VB.Menu ArcEmail 
         Caption         =   "&Email"
      End
      Begin VB.Menu ChatMes 
         Caption         =   "&Mensaje"
      End
      Begin VB.Menu Chatpostal 
         Caption         =   "&Minipostal"
      End
      Begin VB.Menu ChatChat 
         Caption         =   "&Chat"
      End
      Begin VB.Menu ChatConect 
         Caption         =   "&Conectar"
      End
      Begin VB.Menu raya 
         Caption         =   "-"
      End
      Begin VB.Menu Delcontacto 
         Caption         =   "&Eliminar Contacto"
      End
   End
   Begin VB.Menu Arch 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu ArcRes 
         Caption         =   "&Restaurar"
      End
      Begin VB.Menu ArcSal 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "Finfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rpta As String
Dim wstate As Integer
Dim iret As Long

Dim horaini As Date
'Dim usuarios(20) As String
Private Const NIM_ADD = &H0             'Adds an icon to the taskbar notification area
Private Const NIM_MODIFY = &H1          'Changes the icon, tooltip text or notification message for an icon in the notification area
Private Const NIM_DELETE = &H2          'Deletes an icon from the taskbar notification area
'Flags
Private Const NIF_MESSAGE = &H1         'hIcon is valid
Private Const NIF_ICON = &H2            'uCallbackMessage is valid
Private Const NIF_TIP = &H4             'szTip is valid

Private Const WM_MOUSEMOVE = &H200      'MouseMove message identifier
                                    
                                        'Messages sent to the form's MouseMove event
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205


Private Type NOTIFYICONDATA
    cbSize              As Long
    hwnd                As Long         'Handle of window that receives notification messages
    uID                 As Long         'Application-defined identifier of the taskbar icon
    uFlags              As Long         'Flags indicating which structure members contain valid data
    uCallbackMessage    As Long         'Application defined callback message
    hIcon               As Long         'Handle of taskbar icon
    szTip               As String * 64  'Tooltip text to display for icon
End Type

Dim mtIconData          As NOTIFYICONDATA
Dim mnLight             As Integer

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long


Private Sub AddIconToTray() 'Adds an icon to the taskbar notification area

    With mtIconData
        .cbSize = Len(mtIconData)
        .hwnd = Me.hwnd                                     'Use the form to receive callback messages.
        .uCallbackMessage = WM_MOUSEMOVE                    'Tell icon to send MouseMove messages.
        .uID = 1&                                           'Application defined identifier
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .hIcon = fig(0).Picture                     'Initial icon
        .szTip = fig(0).Tag & Chr$(0)               'Initial tooltip for icon
        If Shell_NotifyIcon(NIM_ADD, mtIconData) = 0 Then   'Create icon in tray
            MsgBox "Unable to add icon to system tray!"
        End If
    End With
    
End Sub
Private Sub DeleteIconFromTray()
    If Shell_NotifyIcon(NIM_DELETE, mtIconData) = 0 Then
        MsgBox "Unable to delete icon from system tray!"
    End If
End Sub
Function buscanick(cnic As String) As Integer
Dim cbus As String
  Screen.MousePointer = 11
  cbus = "nick=" & Lista.SelectedItem.Text & "&fec=" & Format(Now, "ddmmyyyy") & ".txt"
  If Perl.StillExecuting Then buscanick = 0: Exit Function
  Perl.Execute curl & "/cgi-local/busnick.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  din = Separa(rpta & ":", ":")
  buscanick = din
  Screen.MousePointer = 0
End Function
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
Function Separa2(cadEnt As String, sep)
'Dim sep As String
Dim j As Integer
Dim cad As String
Dim i As Integer, p As Integer
  
  'sep = "_"
  i = 1
  j = 1
  cad = Mid(cadEnt, 1, InStr(cadEnt, vbLf) - 1)
  Do While True
    p = InStr(j, cad, sep)
    usuarios(i) = Mid(cad, j, p - j)
    j = p + 1
    i = i + 1
    If j >= Len(cad) Then Exit Do
  Loop
  Separa2 = i - 1
End Function
Function quienes(cnick As String) As String
Dim cbus As String
  Screen.MousePointer = 11
  cbus = "nick=" & cnick & "&fec=" & Format(Now, "ddmmyyyy") & ".txt"
  Perl.Execute curl & "/cgi-local/devuelve.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  quienes = rpta
  Screen.MousePointer = 0
  
End Function
Function NuevoSock() As Integer
Dim ns As Integer, l As Integer
  ns = ws.Count
  If ns = 1 Then
    ns = 1
  Else
    l = 0
    For Each soquet In ws
      If soquet.Index <> l Then
        ns = l
        Exit For
      End If
      l = l + 1
    Next
  End If
  NuevoSock = ns
End Function
Function jalapubli()
Dim cbus As String
  cbus = "nick=narf"
  Perl.Execute curl & "/cgi-local/publicidad.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
End Function
Function Buscatag(ntag As Integer) As Integer
Dim cad As String
Dim n As Integer
  'Dim n As Integer
  n = Lista.ListItems.Count
  cad = Trim(Str(ntag))
  For i = 1 To n
    If Lista.ListItems(i).Tag = cad Then
      Buscatag = i
      Exit Function
    End If
  Next
End Function
Function Encuentra(cnick As String) As Integer
Dim n As Integer
  n = Lista.ListItems.Count
  For i = 1 To n
    If Lista.ListItems(i).Text = cnick Then
      Encuentra = i
      Exit Function
    End If
  Next
  Encuentra = 0
End Function
Function llaman(cip As String) As String
Dim cbus As String
  cbus = "nick=" & cip & "&fec=" & Format(Now, "ddmmyyyy") & ".txt"
  Perl.Execute curl & "/cgi-local/busip.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  din = Separa(rpta & ":", ":")
  llaman = usuario(2)
End Function

Private Sub AddContacto_Click()
  'Fuser.Show 1
End Sub

Private Sub ArcEmail_Click()
  Fmail.Show 1
End Sub

Private Sub ArcRes_Click()
  WindowState = 0
  Visible = True
  Call DeleteIconFromTray
End Sub

Private Sub ArcSal_Click()
  DeleteIconFromTray
  End
End Sub


Private Sub ChatChat_Click()
Dim f As Fchat
Dim n As Integer
Dim enc As Integer
Dim nactual As Integer
  enc = False
  n = Forms.Count
  For i = 0 To n - 1
    If Forms(i).Caption = Lista.SelectedItem.Text Then
      enc = True
      Exit For
    End If
  Next
  If enc Then Exit Sub
    nactual = Val(Lista.SelectedItem.Tag)
    If ws(nactual).State <> 7 Then
      MsgBox "Se perdió la conexión"
      ws(nactual).Close
      Unload ws(nactual)
      Lista.SelectedItem.Ghosted = True
      Lista.SelectedItem.Tag = ""
      Exit Sub
    End If
    ws(nactual).SendData "<c>"
  Set f = New Fchat
  selec = Lista.SelectedItem.Tag
  f.Caption = Lista.SelectedItem.Text
  f.Tag = selec
  f.Show
End Sub

Private Sub ChatConect_Click()
Dim nsock As Integer
  din = buscanick(Lista.SelectedItem.Text)
  Select Case din
   Case Is > 1
    'Conectar
    Lista.SelectedItem.SubItems(3) = usuario(2)
    Lista.SelectedItem.Ghosted = False
    nsock = NuevoSock()
    Load ws(nsock)
    ws(nsock).LocalPort = 0
    ws(nsock).RemoteHost = usuario(2)
    ws(nsock).RemotePort = 75
    ws(nsock).Connect
    Lista.SelectedItem.Tag = nsock
   Case 1
    MsgBox "Parece que " & Lista.SelectedItem & " No está en línea"
    Lista.SelectedItem.Ghosted = True
   Case -1
    MsgBox "Actualmente no se está recibiendo información del servidor" & vbCrLf & "Reinténtelo más tarde"
  End Select
End Sub

Private Sub ChatMes_Click()
On Error GoTo problema
Dim cad As String
Dim nactual As Integer
  cad = Trim(InputBox("Mensaje"))
  If InStr(cad, "<p>") > 0 Then MsgBox "Su mensaje no puede contener los simbolos <> :-(": Exit Sub
  If Len(cad) > 0 Then
    nactual = Val(Lista.SelectedItem.Tag)
    If ws(nactual).State <> 7 Then
      MsgBox "Se perdió la conexión"
      ws(nactual).Close
      Unload ws(nactual)
      Lista.SelectedItem.Ghosted = True
      Lista.SelectedItem.Tag = ""
      Exit Sub
    End If
    ws(nactual).SendData cad
  End If
problema:
If Err > 0 Then
  MsgBox "Hubo un problema con la conexión reintente más tarde"
  Resume Next
End If
End Sub

Private Sub Chatpostal_Click()
  Fpostal.Show 1
End Sub

Private Sub cmdlogin_Click()
Dim cbus As String
Dim cad As String
'Dim StartupData As WSADataType
Dim cip As String
'Dim socketbuffer As sockaddr
  cad = String(256, 0)
  'rc = WSAStartup(&H101, StartupData)
  'sock = socket(AF_INET, SOCK_STREAM, 0)
  cip = GetIPAddress()
  'rc = WSACleanup
  Fintro.Say(1) = "Buscando Servidor"
  Fintro.Say(1).Refresh
  'Aquí hay que arreglar
  
  cbus = "nick=" & login & "&ip=" & cip & "&fec=" & Format(Now, "ddmmyyyy") & ".txt"
  Perl.Execute curl & "/cgi-local/addnick.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  'Exit Sub
  Screen.MousePointer = 11
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  Fintro.Say(1) = "Servidor Ok"
  Fintro.Say(1).Refresh
  Screen.MousePointer = 0
End Sub


Private Sub cmdpdeinfo_Click()
Dim cbus As String
Dim cpdeinfo(5) As String
  Screen.MousePointer = 11
  cbus = "nick=narf"
  Perl.Execute curl & "/cgi-local/pdeinfo.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  If InStr(rpta, vbCr) > 0 Then rpta = Mid(rpta, 1, InStr(rpta, vbCr) - 1)
  din = Separar(rpta, ":", cpdeinfo())
  If (cpdeinfo(1) = ciudad Or cpdeinfo(1) = "todos") Then pdeinfo.Caption = cpdeinfo(2) Else pdeinfo.Caption = "Gracias por usar InfoChat"
  Screen.MousePointer = 0
End Sub

Private Sub cmdpubli_Click()
  Screen.MousePointer = 11
  din = jalapubli()
  din = Separa(rpta & vbCrLf, vbLf)
  For i = 1 To din
    anuncios(i) = usuario(i)
  Next
  Screen.MousePointer = 0
End Sub
Private Sub Delcontacto_Click()
Dim cad As String


If MsgBox("Está seguro", vbYesNo) = vbYes Then
  If Lista.SelectedItem.Tag <> "" Then MsgBox "No puedes eliminar un usuario con el cual estas conectado :)": Exit Sub
  Screen.MousePointer = 11
  Open App.Path & "\usuarios.dat" For Input As #1
  Open App.Path & "\tmp.000" For Output As #2
  Do While Not EOF(1)
    Line Input #1, cad
    din = Separa(cad, ":")
    If usuario(1) = Lista.SelectedItem.Text Then
    Else
      Print #2, cad
    End If
  Loop
  Close #1
  Close #2
  Kill App.Path & "\usuarios.dat"
  Name App.Path & "\tmp.000" As App.Path & "\usuarios.dat"
  Lista.ListItems.Remove Lista.SelectedItem.Index
  Screen.MousePointer = 0
End If
End Sub

Private Sub Form_Activate()
  If lprimero = 1 Then
    tipomsg = 3
    Visible = False
    Fintro.Show
    Call cmdlogin_Click
    Call Ok_Click
    Call cmdpubli_Click
    Call cmdpdeinfo_Click
    lprimero = 0
    'Unload Fintro
    Visible = True
    tiempo.Interval = 50
  End If
End Sub

Private Sub Form_Load()
Dim x As ListItem
Dim l As Integer
Dim nimage As Integer
Dim cad As String
  'curl = "http://127.0.0.1"
  curl = "http://pdeinfo.com"
  Flogin.Show 1
  If lprimero = 0 Then End
  'Fintro.Show
  'curl = "http://pdeinfo.com"
  horaini = Now
  lprimero = 1
  ws(0).LocalPort = 75
  ws(0).Listen
  Open App.Path & "\infochat.ini" For Input As #1
  Line Input #1, cad
  Close #1
  din = Separa(cad, ":")
  cemail = usuario(3)
  login = usuario(1)
  ciudad = usuario(5)
  Open App.Path & "\usuarios.dat" For Input As #1
  l = 1
  nimage = 1
  Do While Not EOF(1)
    Line Input #1, cad
    usuario(3) = ""
    din = Separa(cad, ":")
    'Buscar al infeliz
    'wstate = 0
    nimage = l
    If l >= 10 Then nimage = 1
    Set x = Lista.ListItems.Add(, , usuario(1), nimage, nimage)
    x.SubItems(1) = usuario(2)
    x.SubItems(2) = usuario(3)
    x.Ghosted = True
    l = l + 1
  Loop
  Close #1

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Static bBusy As Boolean

    If bBusy = False Then           'Do one thing at a time
        bBusy = True
        
        
        Select Case CLng(x)
            Case WM_LBUTTONDBLCLK   'Double-click left mouse button: same as selecting About
                WindowState = 0
                Visible = True
                Call DeleteIconFromTray
            Case WM_LBUTTONDOWN     'Left mouse button pressed: change traffic light icon & tip
               
                
            Case WM_LBUTTONUP       'Left mouse button released
            
            Case WM_RBUTTONDBLCLK   'Double-click right mouse button
            
            Case WM_RBUTTONDOWN     'Right mouse button pressed
            
            Case WM_RBUTTONUP       'Right mouse button released: display popup menu
              PopupMenu Arch
        End Select
        
        bBusy = False
    End If

End Sub

Private Sub Form_Resize()
  If Me.WindowState = 1 Then
    Me.Visible = False
    'tiempo = False
    AddIconToTray
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  End
End Sub
Private Sub Image1_DblClick()
  iret = ShellExecute(Me.hwnd, vbNullString, Image1.Tag, vbNullString, "c:\", SW_SHOWNORMAL)
''  'Fabout.Show 1
End Sub

Private Sub imgboton_Click()
  imgboton.Visible = False
  imgboton.Width = 0
  imgboton.Height = 0
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 And Shift = 7 Then
    imgboton.Top = 0
    imgboton.Left = 0
    imgboton.Width = 95
    imgboton.Height = 76
    imgboton.Visible = True
  End If
End Sub

Private Sub lbladic_Click()
  Fuser.Show 1
End Sub

Private Sub lblservicios_Click(Index As Integer)
  iret = ShellExecute(Me.hwnd, vbNullString, lblservicios(Index).ToolTipText, vbNullString, "c:\", SW_SHOWNORMAL)
End Sub

Private Sub lblservicios_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
  lblinfo.Caption = lblservicios(Index).Tag
End Sub

Private Sub Lista_BeforeLabelEdit(Cancel As Integer)
  Cancel = True
End Sub

Private Sub Lista_DblClick()
  If Len(Lista.SelectedItem.Tag) = 0 Then Exit Sub
  Fdata.Show 1
End Sub

Private Sub Lista_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If Button = 2 Then
    ChatMes.Visible = (Lista.SelectedItem.Tag <> "")
    Chatpostal.Visible = (Lista.SelectedItem.Tag <> "")
    ChatChat.Visible = (Lista.SelectedItem.Tag <> "")
    ChatConect.Visible = (Lista.SelectedItem.Tag = "")
    Delcontacto.Visible = (Lista.SelectedItem.Tag = "")
    Delcontacto.Visible = (Lista.ListItems.Count > 1)
    ArcEmail.Visible = Len(Lista.SelectedItem.SubItems(2)) > 0
    PopupMenu Chat
  End If
End Sub

Private Sub Ok_Click()
Dim cbus As String
Dim n As Integer
  Screen.MousePointer = 11
  n = Lista.ListItems.Count
  For i = 1 To n
  Fintro.Say(1) = "Buscando usuario " & Lista.ListItems(i).Text
  Fintro.Say(1).Refresh
  cbus = "nick=" & Lista.ListItems(i).Text & "&fec=" & Format(Now, "ddmmyyyy") & ".txt"
  Perl.Execute curl & "/cgi-local/busnick.pl", "POST", cbus, "Content-Type: application/x-www-form-urlencoded"
  'Exit Sub
  wstate = 0
  Do While True
    If wstate = 1 Then Exit Do
  DoEvents
  Loop
  din = Separa(rpta & ":", ":")
  If din > 1 Then
    Lista.ListItems(i).SubItems(3) = usuario(2)
    Lista.ListItems(i).Ghosted = False
    Fintro.Say(1) = "Usuario Ok"
    Fintro.Say(1).Refresh
  Else
    Lista.ListItems(i).Ghosted = True
    Fintro.Say(1) = "Usuario No conectado"
    Fintro.Say(1).Refresh
  End If
  Next
  Fintro.Say(1) = "Por favor espere..."
  Fintro.Say(1).Refresh
  Screen.MousePointer = 0
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


Private Sub tiempo_Timer()
Dim minutos As String
Dim nmin As Integer
Dim wflags As Integer
Dim x As Integer
Static nf As Integer
  DoEvents
If WindowState = 1 Then
  Select Case tipomsg
  Case 0 'Llaman
    If nf > 2 Then
      nf = 0
    End If
    wflags = SND_ASYNC Or SND_NODEFAULT
    x = sndPlaySound(sonido, wflags)
    If Len(sonido) > 0 Then sonido = ""
    With mtIconData
          .hIcon = fig(nf).Picture                     'Initial icon
          .szTip = "Alquién lo está llamando " & Chr(0)               'Initial tooltip for icon
    End With
    If Shell_NotifyIcon(NIM_MODIFY, mtIconData) = 0 Then
          MsgBox "Unable to change icon in system tray!"
    End If
    nf = nf + 1
  Case 1 'Escriben
    wflags = SND_ASYNC Or SND_NODEFAULT
    x = sndPlaySound(sonido, wflags)
     If Len(sonido) > 0 Then sonido = ""
     If (nf > 5) Or (nf < 3) Then
      nf = 3
     End If
     With mtIconData
          .hIcon = fig(nf).Picture                     'Initial icon
          .szTip = "Mensaje nuevo" & Chr(0)               'Initial tooltip for icon
     End With
     If Shell_NotifyIcon(NIM_MODIFY, mtIconData) = 0 Then
          MsgBox "Unable to change icon in system tray!"
     End If
     nf = nf + 1
  End Select
Else
  tipomsg = 3
End If
  If publicidad.Top = -860 Then
    Randomize
    publicidad.Top = 840
    publicidad = anuncios(Int((10 * Rnd) + 1))
  End If
  pdeinfo.Left = pdeinfo.Left - 1
  'pdeinfo.Refresh
  If pdeinfo.Left = -pdeinfo.Width Then pdeinfo.Left = pdeinfo.Width
  publicidad.Top = publicidad.Top - 10
  minutos = Format(Now - horaini, "hh:mm:ss")
  nmin = Val(Mid(minutos, 4, 2))
  If ((nmin Mod 10) = 0) And nmin > 0 Then
    If Perl.StillExecuting Then Exit Sub
    Call cmdpubli_Click
    Call cmdpdeinfo_Click
    horaini = Now
  End If
  'say = minutos
  'say.Refresh
End Sub



Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
Dim ntag As Integer
Dim ntecla As Integer
  ntecla = KeyCode
  KeyCode = 0
  ntag = Val(txt.Tag)
  If Len(txt.Tag) = 0 Then Exit Sub
 Select Case ntecla
  Case 38
    If ntag - 1 < 0 Then Exit Sub
    txt.Tag = ntag - 1
    txt = Lhistoria.List(txt.Tag)
    txt.Refresh
  Case 40
    If ntag = (Lhistoria.ListCount - 1) Then Exit Sub
    txt.Tag = ntag + 1
    txt = Lhistoria.List(txt.Tag)
    txt.Refresh
 End Select
End Sub

Private Sub ws_ConnectionRequest(Index As Integer, ByVal requestID As Long)
Dim nsock As Integer, nenc As Integer
Dim cad As String
Dim cmes As String
Dim l As Integer
Dim x As ListItem
If lprimero = 1 Then Exit Sub
Select Case Index
  Case 0
    tipomsg = 0
    sonido = App.Path & "\llaman.wav"
    cad = llaman(ws(0).RemoteHostIP)
    Say = cad & " Esta tratando de ubicarlo"
    Say.Refresh
    nenc = Encuentra(cad)
    If nenc = 0 Then
      cmes = cad & " Está tratando de contactarle pero no es uno de sus contactos" & vbCrLf & "Desea más información"
      If MsgBox(cmes, vbYesNo) = vbYes Then
        cmes = quienes(cad)
        din = Separa2(cmes, ":")
        Fguest.Show 1
        If MsgBox("Desea agregarlo a su lista de contactos", vbYesNo) = vbYes Then
          Open App.Path & "\usuarios.dat" For Append As #1
            Print #1, usuarios(1) & ":" & usuarios(2) & ":" & usuarios(3) & ":"
          Close #1
          nsock = NuevoSock()
          Load ws(nsock)
          ws(nsock).Accept requestID
          l = Finfo.Lista.ListItems.Count + 1
          Set x = Finfo.Lista.ListItems.Add(, , usuarios(1), l, l)
          x.SubItems(1) = usuarios(2)
          x.SubItems(2) = usuarios(3)
          x.Ghosted = False
          x.Tag = nsock
        End If
      End If
      Exit Sub
    End If
    nsock = NuevoSock()
    Load ws(nsock)
    ws(nsock).Accept requestID
    Lista.ListItems(nenc).Ghosted = False
    Lista.ListItems(nenc).Tag = nsock
    Say = ""
    Say.Refresh
  Case Else
End Select
End Sub

Private Sub ws_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error GoTo Errmodal
Dim cad As String
Dim ctrans As String
Dim cmpostal As String
Dim nf As Integer
Dim cfile As String
Dim curlpostal As String
Dim ctxtpostal As String
  
  Select Case Index
    Case 0
    Case Else
      tipomsg = 1
      sonido = App.Path & "\mensaje.wav"
      ws(Index).GetData cad
      Select Case Mid(cad, 1, 3)
       Case "<s>"
        MsgBox Lista.ListItems(Buscatag(Index)).Text & " Recibió la postal", vbSystemModal
       Case "<c>" ' Avisar que ha entrado al chat
        MsgBox Lista.ListItems(Buscatag(Index)).Text & " Está en modo chat", vbSystemModal
       Case "<p>" ' Minipostal
        cmpostal = Lista.ListItems(Buscatag(Index)).Text & " Le está enviando una postal" & vbCrLf & "Desea Recibirla"
        If MsgBox(cmpostal, vbYesNo) = vbYes Then
          curlpostal = Mid(cad, InStr(cad, ">") + 1)
          nf = FreeFile
          ctxtpostal = "v" & Format(Now, "hhmmss")
          curlpostal = Mid(cad, InStr(cad, ">") + 1)
          nf = FreeFile
            cfile = App.Path & "\" & ctxtpostal & ".htm"
            Open cfile For Output As #nf
            Print #1, "<html><head><title>InfoChat - Ha recibido una postal virtual!</title><SCRIPT>"
            Print #1, "function MuestraPostal(foto){"
            Print #1, ctxtpostal & " = window.open(" & Chr(34) & Chr(34) & "," & Chr(34) & Chr(34) & "," & Chr(34) & "width=450,height=290" & Chr(34) & ")" & ";"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "<html><title>Servicio de Minipostales Infochat</title><body><center>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "<table border=1 CELLSPACING=0 CELLPADDING=5 width=430><tr><td width=40%>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "<img src='http://pdeinfo.com/infochat/" & Chr(34) & "+foto+" & Chr(34) & "'>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "</td><td width=60%><font face='Times New Roman' size=4 color=green><B><I>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "<center>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "<font face='Times New Roman' size=4><b><i>" & Mid(curlpostal, InStr(curlpostal, ":") + 1) & "</b></i></font>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.write(" & Chr(34) & "</center></I></B></font></td></tr></table></center></body></html>" & Chr(34) & ");"
            Print #1, ctxtpostal & ".document.close;return true}</SCRIPT></head><body>"
            Print #1, "<h1 align=center><font size=6><i><b>Acaba de recibir una minipostal de " & Lista.ListItems(Buscatag(Index)).Text & "</b></i></font></h1><form>"
            Print #1, "<p align=center><input type=button value='Pulse aqui para verla' onclick='MuestraPostal(" & Chr(34) & Mid(curlpostal, 1, InStr(curlpostal, ":") - 1) & ".jpg" & Chr(34) & ")' name=miMiniPostal>"
            Print #1, "</form><hr size=1 color=#FF9966><div align=right><font face='Ms Sans Serif' size=1 color=#FF9966>"
            Print #1, "Postal generada con el <a href='http://pdeinfo.com/infochat'><font color='#FF9966'><b>InfoChat</b>"
            Print #1, "</font></a> - Este es un servicio gratuito de <a href='http://pdeinfo.com'><font color=#FF9966><B>Punto de Información</B></font></a></a></font></div></body></html>"
            Close #nf
            iret = ShellExecute(Me.hwnd, vbNullString, "file://" & cfile & vbNullString, vbNullString, "c:\", 3)
            ws(Index).SendData "<s>"
        Else
          ws(Index).SendData "<n>" ' rechazo postal
        End If
       Case "<n>" ' Mensaje Rechazó postal
        MsgBox Lista.ListItems(Buscatag(Index)).Text & " Rechazó la postal", vbSystemModal
        'cchat = Lista.ListItems(Buscatag(Index)).Text & ":" & cad
        'Say = Lista.ListItems(Buscatag(Index)).Text & " rechazo la postal "
        'Lhistoria.AddItem Say
        'txt = Say
        'txt.Tag = Lhistoria.ListCount - 1
       Case Else
        cchat = Lista.ListItems(Buscatag(Index)).Text & ":" & cad
        Say = Lista.ListItems(Buscatag(Index)).Text & " escribe: " & cad
        Lhistoria.AddItem Say
        txt = Say
        txt.Tag = Lhistoria.ListCount - 1
      End Select
  End Select
Errmodal:
If Err > 0 Then
  If Err = 400 Then
    ' MsgBox "Para poder ver la nueva postal se cerrara la postal actual"
    'Unload Frecibe
    Resume
  End If
End If
End Sub
