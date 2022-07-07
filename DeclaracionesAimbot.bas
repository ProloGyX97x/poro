Attribute VB_Name = "DeclaracionesAimbot"
'Begin Code

    Option Explicit










Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Const LB_FINDSTRINGEXACT As Long = &H1A2
Public Const LB_FINDSTRING As Long = &H18F
Public Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Const MAX_PLAYERS As Long = 10000
Public Const SIZE_PLAYERS As Long = &H1CDC
Public Const BASE_PLAYERS As Long = &H781300 '0077E2F4 75DD64
Public Const OFFSET_POSX As Long = &H8  ' Cambiar por 8 si ta mal xd
Public Const OFFSET_POSY As Long = &HA
Public Const OFFSET_NAME As Long = &H1CB4
Public Const MagicAddress As Long = &H75D4B0
Public Const ADDRESS_MY_POSX As Long = &H781216 '
Public Const ADDRESS_MY_POSY As Long = ADDRESS_MY_POSX + 2
Public Const MY_POSX As Long = -1
Public Const MY_POSY As Long = -1
Public MOUSE_POSITION_ME_X As Long ' = 736 / 2 - 16 '352 ' este es mi pj desde el formulario de tp
Public MOUSE_POSITION_ME_Y As Long ' = 480 / 2 - 16 '224
Public Declare Function WindowFromPoint Lib "User32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public tHwnd As Long
Public MyPosX As Long, MyPosY As Long
Public MyPosX2 As Long, MyPosY2 As Long
Public CharList(MAX_PLAYERS) As tCharlist

Public Type tCharlist
    idPlayer As Long
    address As Long
    
    Active As Long
    Heading As Long
    
    PosX As Integer
    PosY As Integer
    
    addressName As Long
    name As String
    
    ScrollDirectionX As Integer
    ScrollDirectionY As Integer
    Moving As Boolean
    
    Target As Boolean
End Type

Public RenderTop As Long 'NEW
Public RenderLeft As Long
Public RenderRight As Long 'NEW
Public RenderBottom As Long

Public Const RANGE_X As Long = 11
Public Const RANGE_Y As Long = 7
Private Const BM_SETSTATE = &HF3
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Public mm As New MemoryManager
Public resX As Single
Public resY As Single
Public Type tCustomList
    name As String
    Delete As Boolean
End Type

Public InRangeList() As tCustomList

Public LastCharTargeted As Integer

Public MainAOhWnd As Long
Public PictureRenderhWnd As Long
Public Sub EsperaSeg(ByVal Segundos As Long)
  Dim Hora As Double
  On Local Error Resume Next
  Hora = Timer
  Do Until Timer >= Hora + Segundos
    DoEvents
  Loop
End Sub
Public Sub EsperaMs(ByVal Tiempo As Double)
  Dim HoraActual As Double
  On Local Error Resume Next
  HoraActual = Timer
  Do Until Timer >= HoraActual + (Tiempo / 1000)
    DoEvents
  Loop
End Sub
Private Function TransformarCors(ByVal x As Long, ByVal y As Long) As Long
TransformarCors = (y * &H10000) + x
End Function
Public Function GetCaption(ByVal hwnd As Long)
Dim hWndlength As Long, hWndTitle As String, a As Long

'Get the length of the caption
hWndlength = GetWindowTextLength(hwnd)

'Fill up a string with that amount of characters
hWndTitle = String$(hWndlength, 0)

'Fill the string with the real caption
a = GetWindowText(hwnd, hWndTitle, (hWndlength + 1))
GetCaption = hWndTitle
End Function

Public Sub SetUserByChar(ByVal idPlayer As Long, ByRef tmpChar As tCharlist)
With CharList(idPlayer)
    .idPlayer = tmpChar.idPlayer
    .address = tmpChar.address
    
    .Active = tmpChar.Active
    .Heading = tmpChar.Heading
    .PosX = tmpChar.PosX
    .PosY = tmpChar.PosY
    
    .addressName = tmpChar.addressName
    .name = tmpChar.name
    
    .ScrollDirectionX = tmpChar.ScrollDirectionX
    .ScrollDirectionY = tmpChar.ScrollDirectionY
    .Moving = tmpChar.Moving
End With
End Sub

Public Sub SetUser(ByVal idPlayer As Long, ByVal address As Long, ByVal addressName As Long, ByVal name As String, ByVal PosX As Long, ByVal PosY As Long)
With CharList(idPlayer)
    .idPlayer = idPlayer
    .address = address
    .addressName = addressName
    .name = name
    .PosX = PosX
    .PosY = PosY
End With
End Sub

Public Function GetUserByName(ByVal name As String) As Long
Dim i As Integer

'Loopeo todos los usuarios hasta encontrar el nombre por parámetro
For i = 1 To MAX_PLAYERS
    If CharList(i - 1).name = name Then
        GetUserByName = (i - 1)
        Exit Function
    End If
Next i

GetUserByName = -1
End Function

Public Function GetUserNameByIndex(ByVal idPlayer As Long) As String
GetUserNameByIndex = CharList(idPlayer).name
End Function

Public Function GetPosXByIndex(ByVal PosX As Long) As String
GetPosXByIndex = CharList(PosX).PosX
End Function

Public Function GetPosYByIndex(ByVal PosY As Long) As String
GetPosYByIndex = CharList(PosY).PosY
End Function

Public Function GetAddressByIndex(ByVal CharIndex As Long) As String
GetAddressByIndex = CharList(CharIndex).address
End Function

'Seteo target por id de usuario
Public Sub SetUserTarget(ByVal idPlayer As Long, ByVal Target As Boolean)
With CharList(idPlayer)
    .Target = Target
End With
End Sub

Public Function AddToList(ByVal name As String) As Boolean
Dim i As Integer

For i = LBound(InRangeList) To UBound(InRangeList)
    If InRangeList(i).name = name Then
        AddToList = False
        Exit Function
    End If
Next i

i = GetFreeIndex()

If i >= 0 Then
    InRangeList(i).name = name
    InRangeList(i).Delete = False
Else
    ReDim Preserve InRangeList(UBound(InRangeList) + 1)
End If

AddToList = True
End Function

Public Function GetFreeIndex() As Integer
Dim i As Integer

For i = LBound(InRangeList) To UBound(InRangeList)
    If Len(InRangeList(i).name) = 0 And Not InRangeList(i).Delete Then
        GetFreeIndex = i
        Exit Function
    End If
Next i

GetFreeIndex = -1
End Function

Public Sub CleanList()
Dim i As Integer

For i = LBound(InRangeList) To UBound(InRangeList)
    If InRangeList(i).Delete Then
        InRangeList(i).name = ""
        InRangeList(i).Delete = False
    End If
Next i
End Sub
Public Sub Aimbot()
If LastCharTargeted <> -1 Then
    Inicial.Mouse2.Enabled = True
End If
End Sub



