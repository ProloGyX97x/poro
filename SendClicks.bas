Attribute VB_Name = "SendClicks"
'Begin Code

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, _
Source As Any, ByVal Length As Long)
Private Declare Function SendInput Lib "user32.dll" (ByVal nInputs As Long, pInputs As INPUT_TYPE, ByVal cbSize As Long) As Long

Public Const WM_SETFOCUS As Long = &H7
Public Const WM_MOUSEMOVE As Long = &H200
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const INPUT_MOUSE = 0
Private Const INPUT_KEYBOARD = 1
Private Const INPUT_HARDWARE = 2

Private Type INPUT_TYPE
  dwType As Long
  xi(0 To 23) As Byte
End Type

Private Type MOUSEINPUT
  dx As Long
  dy As Long
  mouseData As Long
  dwFlags As Long
  time As Long
  dwExtraInfo As Long
End Type

Public Declare Function PostMessageA Lib "user32" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function AttachThreadInput Lib "user32.dll" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function GetWindowThreadProcessId_New Lib "user32.dll" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Public Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Function MakeDWord(ByVal LoWord As Integer, ByVal HiWord As Integer) As Long
    MakeDWord = (HiWord * &H10000) Or (LoWord And &HFFFF&)
End Function

'INPUT    Input = { 0 };
'// left down
'Input.type = INPUT_MOUSE;
'Input.mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
'::SendInput(1, &Input, sizeof(INPUT));
'// left up
'::ZeroMemory(&Input, sizeof(INPUT));
'Input.type = INPUT_MOUSE;
'Input.mi.dwFlags = MOUSEEVENTF_LEFTUP;
'::SendInput(1, &Input, sizeof(INPUT));
Public Sub LeftClick()
    Dim inputevents(2) As INPUT_TYPE
    Dim mouseevent As MOUSEINPUT  ' temporarily hold mouse input info
       
    mouseevent.dx = 0  ' no horizontal movement
    mouseevent.dy = 0 ' no vertical movement
    mouseevent.mouseData = 0  ' not needed
    mouseevent.dwFlags = MOUSEEVENTF_LEFTDOWN  ' right button up
    mouseevent.dwExtraInfo = 0  ' not needed
    CopyMemory inputevents(0).xi(0), mouseevent, Len(mouseevent)
    SendInput 1, inputevents(0), Len(inputevents(0))

    mouseevent.dx = 0 ' no horizontal movement
    mouseevent.dy = 0  ' no vertical movement
    mouseevent.mouseData = 0  ' not needed
    mouseevent.dwFlags = MOUSEEVENTF_LEFTUP  ' right button up
    mouseevent.dwExtraInfo = 0  ' not needed
    CopyMemory inputevents(0).xi(0), mouseevent, Len(mouseevent)

    SendInput 1, inputevents(0), Len(inputevents(0))
End Sub



