Attribute VB_Name = "Memoria"
    'Begin Code


Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal Classname As String, ByVal WindowName As String) As Long
Private Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
        
        Public hwndRender As Long
        Public hWndMain As Long

Private Const BM_SETSTATE = &HF3
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_LBUTTONDBLCLK As Long = &H203
Private Declare Function GetClassName _
Lib "User32" _
Alias "GetClassNameA" ( _
ByVal hwnd As Long, _
ByVal lpClassName As String, _
ByVal nMaxCount As Long) As Long
Private Declare Function GetParent _
Lib "User32" ( _
ByVal hwnd As Long) As Long
Private Declare Function GetWindowText _
Lib "User32" _
Alias "GetWindowTextA" ( _
ByVal hwnd As Long, _
ByVal lpString As String, _
ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength _
Lib "User32" _
Alias "GetWindowTextLengthA" ( _
ByVal hwnd As Long) As Long
Private Const VK_RETURN = &HD
Private Const WM_KEYDOWN = &H100
Private Const WM_KEYUP = &H101
Private Const VK_KEY_U = &H55
Private Const VK_KEY_R = &H52
Private Const VK_SPACE = &H20
Private Const VK_CONTROL = &H11
Private Const VK_RCONTROL = &HA2

Private Function TransformarCors(ByVal x As Long, ByVal y As Long) As Long
TransformarCors = (y * &H10000) + x
End Function
Sub UsarItem(ByVal Slot As String)
On Error Resume Next
Dim hwnd As Long, pid As Long
GetWindowThreadProcessId hwndexe, pid
mm.pid = pid

    Select Case Slot

        Case "Slot01"
        Call SendMessage(hwndInventario, WM_LBUTTONUP, 0, ByVal TransformarCors(10, 16))
        Call SendMessage(hwndInventario, WM_LBUTTONDBLCLK, 0, ByVal TransformarCors(10, 16))
       ' If FrmTeclas.modComun.Value = vbChecked Then
                
                'Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_KEY_U, 0&)
                'Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_KEY_U, 0&)
        
       ' End If
        
       ' If FrmTeclas.modWASD.Enabled = vbChecked Then
                'Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_SPACE, 0&)
        'End If
            

        

        Case "Slot02"
        
        Call SendMessage(hwndInventario, WM_LBUTTONUP, 0, ByVal TransformarCors(48, 16))
        Call SendMessage(hwndInventario, WM_LBUTTONDBLCLK, 0, ByVal TransformarCors(48, 16))

                    
        'If FrmTeclas.modComun.Value = vbChecked Then
                
                'Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_KEY_U, 0&)
                'Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_KEY_U, 0&)
        
        'End If
        'If FrmTeclas.modWASD.Value = vbChecked Then
        
                Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_SPACE, 0&)
                Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_SPACE, 0&)

        'End If
                    
        Case "Rojas"
            Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_SPACE, 0&)
            Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_SPACE, 0&)
            
        Case "Azules"
        
        Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_SPACE, 0&)
        Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_SPACE, 0&)
        
                    
        Case "ARCO"
            'Call PostMessage(PictureRenderhWnd, WM_KEYDOWN, VK_CONTROL, 0&)
            Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_CONTROL, 0&)
                    
        Case "Slot03"
        
                    Call SendMessage(hwndInventario, WM_LBUTTONUP, 0, ByVal TransformarCors(80, 16))
                    Call SendMessage(hwndInventario, WM_LBUTTONDBLCLK, 0, ByVal TransformarCors(80, 16))
                    Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_KEY_U, 0&)
        Case "Slot04"
        
                    Call SendMessage(hwndInventario, WM_LBUTTONUP, 0, ByVal TransformarCors(116, 16))
                    Call SendMessage(hwndInventario, WM_LBUTTONDBLCLK, 0, ByVal TransformarCors(116, 16))
                    Call PostMessage(PictureRenderhWnd, WM_KEYUP, VK_KEY_U, 0&)
                    
        'Case "setEspecial2"
                    'Call SendMessage(hwndHechizos, WM_LBUTTONDOWN, 0, ByVal TransformarCors(42, 151))
                    'Call SendMessage(hwndHechizos, WM_LBUTTONUP, 0, ByVal TransformarCors(42, 151))
                    
        Case "setEspecial"
                    Call SendMessage(hwndHechizos, WM_LBUTTONDOWN, 0, ByVal TransformarCors(29, 164))
                    Call SendMessage(hwndHechizos, WM_LBUTTONUP, 0, ByVal TransformarCors(29, 164))
                    
        Case "setParalizar"
                    Call SendMessage(hwndHechizos, WM_LBUTTONDOWN, 0, ByVal TransformarCors(14, 174))
                    Call SendMessage(hwndHechizos, WM_LBUTTONUP, 0, ByVal TransformarCors(14, 174))
                    
        Case "setRemover"
                    Call SendMessage(hwndHechizos, WM_LBUTTONDOWN, 0, ByVal TransformarCors(67, 190))
                    Call SendMessage(hwndHechizos, WM_LBUTTONUP, 0, ByVal TransformarCors(67, 190))
    End Select
End Sub
Public Function LeerMemoria(OffSet As Long, WindowName As String) As Integer
    Dim hwnd As Long
    Dim ProcessID As Long
    Dim ProcessHandle As Long
    Dim Value As Integer
    hwnd = FindWindow(vbNullString, WindowName)
    If hwnd = 0 Then
        'MsgBox "ERROR:" & vbNewLine & "[Ejecutar una vez logeado]", vbInformation
    'SetWindowPos FrmMain.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
                            SWP_NOMOVE Or SWP_NOSIZE


        Exit Function
    End If
    GetWindowThreadProcessId hwnd, ProcessID
    ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
    If ProcessHandle = 0 Then
        Call FormNoPrimerPlano(FrmMain)
        MsgBox "ERROR:" & vbNewLine & "[Ejecutar una vez logeado]", vbInformation


        End
        Exit Function
    End If
    ReadProcessMem ProcessHandle, OffSet, Value, 2, 0&
    LeerMemoria = Value
    CloseHandle ProcessHandle
End Function
Public Sub Inject()
Dim DllPath As String
Dim ExeName As Integer
ExeName = 1
Dim Direct As String
Direct = App.Path & "\Memory.dll"
ProsH = GetHProcExe(EXEAO)
If ProsH = 0 Then Exit Sub
DllPath = Direct
InjectDll DllPath, ProsH
End Sub


