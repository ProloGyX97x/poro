Attribute VB_Name = "Mouse"
'Begin Code


Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, _
ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Sub ClickIzquierdo(xP As Long, yP As Long)
Dim junk As Long
    junk = SetCursorPos(xP, yP)
    mouse_event MOUSEEVENTF_LEFTDOWN, xP, yP, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, xP, yP, 0, 0
End Sub


