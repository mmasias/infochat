Attribute VB_Name = "power"
Global curl As String
Global sonido As String
Global lprimero As Integer
Global anuncios(10) As String
Global usuario(15) As String
Global usuarios(15) As String
Global cdir As String
Global arconly As String
Global cemail As String
Global login As String
Global cchat As String
Global selec As Integer
Global tipomsg As Integer
Global ldblclick As Integer
Global ciudad As String
Global curlpostal As String
Global ctrans As String
Global Const SND_SYNC = &H0
Global Const SND_ASYNC = &H1
Global Const SND_NODEFAULT = &H2
Global Const SND_LOOP = &H8
Global Const SND_NOSTOP = &H10
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Function Separa(cadEnt As String, sep)
'Dim sep As String
Dim j As Integer
Dim cad As String
Dim i As Integer, p As Integer
  
  'sep = "_"
  i = 1
  j = 1
  cad = cadEnt
  Do While True
    p = InStr(j, cad, sep)
    usuario(i) = Mid(cad, j, p - j)
    j = p + 1
    i = i + 1
    If j >= Len(cad) Then Exit Do
  Loop
  Separa = i - 1
End Function
