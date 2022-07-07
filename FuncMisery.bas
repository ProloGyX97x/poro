Attribute VB_Name = "FuncMisery"
'Begin Code

Option Explicit

Private Type tOffset
    HexValue As Variant
    DecimValue As Long
    HexAddress As Variant
    DecimAddress As Long
End Type

Private NumOffsets As Long
Private Offsets() As tOffset
Private ActualOffsetVal As Variant
'CONSTANTS CHO SENMESSAGE
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_KEYDOWN = &H100
Public Const WM_SETTEXT                As Long = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_ENABLE                 As Long = &HA
Public Const WM_CHAR                   As Long = &H102
Public Const WM_LBUTTONUP              As Long = &H202


Private FAddressHex As Variant
Private FAddressDecim As Long
Private FValueHex As Variant
Private FValueDecim As Long
Private Const EM_GETLINECOUNT As Long = &HBA
Private Const EM_LINEFROMCHAR As Long = &HC9
Private Const EM_LINELENGTH As Long = &HC1
Private Const EM_LINEINDEX As Long = &HBB
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "User32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long
Public Declare Function ReadProcessMem Lib "kernel32" Alias "ReadProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMem Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByRef lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Private Declare Function Hotkey Lib "User32" Alias "GetAsyncKeyState" (ByVal key As Long) As Integer

Private Declare Function FindWindowEx Lib "User32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public RichTextBoxHwnd As Long
Public L(1 To 8) As Long, lR(1 To 8) As Long, lS(1 To 8) As Long
Public v(1 To 8) As Variant
Public OffSet(1 To 8) As Variant

Public Pass As String
Public Sub GetRichTextBoxHwnd(ByVal hwnd As Long)
RichTextBoxHwnd = FindWindowEx(hwnd, 0&, "RichTextWndClass", vbNullString)
'MsgBox RichTextBoxHwnd
'RichTextBoxHwnd = WindowFromPoint(50, 50) 'RICHTEXTBOX HWND
End Sub
Public Sub Buscar_ListBox(Frase As String, _
                   Optional Frase_Completa As Boolean = False)
Dim Indice As Long
      
    ' Tipo de búsqueda
    If Frase_Completa Then
       Indice = SendMessage(Inicial.lstInRange.hwnd, LB_FINDSTRINGEXACT, -1, Frase)
    Else
       Indice = SendMessage(Inicial.lstInRange.hwnd, LB_FINDSTRING, -1, Frase)
    End If
      
      
    If Indice < 0 Then
        'no se encontró
        'MsgBox " No se encontró la cadena en el ListBox ", vbExclamation
    Else
        ' se encontró la frase entonces la selecciona
        Inicial.lstInRange.ListIndex = Indice
        Call Inicial.lstInRange_DblClick
        'End
    End If
End Sub
Public Sub ClearRichTextBox(ByVal hwnd As Long)
'Call SendMessage(Hwnd, WM_SETTEXT, 0, "")
End Sub
Public Sub msgREMO(ByVal hwnd As Long)
Call SendMessage(hwnd, WM_SETTEXT, 0, "Has Lanzado Remover sobre ti mismo")
Call SendMessage(hwnd, WM_SETTEXT, 1, "Has Lanzado Remover sobre ti mismo")
Call SendMessage(hwnd, WM_SETTEXT, 2, "Has Lanzado Remover sobre ti mismo")
Call SendMessage(hwnd, WM_SETTEXT, 3, "Has Lanzado Remover sobre ti mismo")
Call SendMessage(hwnd, WM_SETTEXT, 0, "")
End Sub
Public Function WindowTextGetLastLine(ByVal hwnd As Long) As String
On Local Error Resume Next
Dim i As Long, k As Long
Dim L1 As Long, L2 As Long
Dim Texto As String

Texto = WindowTextGet(hwnd)

'K retorna el Número de líneas del Box
k = SendMessage(hwnd, EM_GETLINECOUNT, 0, 0&) - 1

'L1 el Primer carácter de la línea actual
L1 = SendMessage(hwnd, EM_LINEINDEX, k - 1, 0&) + 1

'Longitud de la línea actual (Cantidad de caracteres )
L2 = SendMessage(hwnd, EM_LINELENGTH, L1, 0&)

'Mostramos la ultima línea del box
WindowTextGetLastLine = Mid$(Texto, L1, L2)
End Function
Public Function WindowTextGetPreviousLine(ByVal hwnd As Long) As String
On Local Error Resume Next
Dim i As Long, k As Long
Dim L1 As Long, L2 As Long
Dim Texto As String

Texto = WindowTextGet(hwnd)

'K retorna el Número de líneas del Box
k = SendMessage(hwnd, EM_GETLINECOUNT, 0, 0&) - 1

'L1 el Primer carácter de la línea actual
L1 = SendMessage(hwnd, EM_LINEINDEX, k - 2, 0&) + 1

'Longitud de la línea actual (Cantidad de caracteres )
L2 = SendMessage(hwnd, EM_LINELENGTH, L1, 0&)

'Mostramos la ultima línea del box
WindowTextGetPreviousLine = Mid$(Texto, L1, L2)
End Function
Public Function WindowTextGet(ByVal hwnd As Long) As String
Dim strBuff As String, lngLen As Long

lngLen = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lngLen > 0 Then
    lngLen = lngLen + 1
    strBuff = String(lngLen, vbNullChar)
    lngLen = SendMessage(hwnd, WM_GETTEXT, lngLen, ByVal strBuff)
    WindowTextGet = Left(strBuff, lngLen)
End If
End Function



Public Function MiseryCalc(ByVal address As Long, ParamArray TheOffsets() As Variant) As Variant
'On Error GoTo Err:

Dim i As Byte
Dim handle As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
Dim PointerValue As Long
Dim AddressDec As Long
Dim AddressHex As String

'MsgBox UBound(TheOffsets) '0, 1
NumOffsets = UBound(TheOffsets) + 1
'MsgBox NumOffsets
'Exit Sub

ReDim Offsets(NumOffsets)

For i = 1 To NumOffsets
    ActualOffsetVal = TheOffsets(i - 1)
    'MsgBox ActualOffsetVal
    
    Offsets(i).HexValue = "&H" & ActualOffsetVal
    Offsets(i).DecimValue = "&H" & ActualOffsetVal
Next i

'handle = FindWindow(vbNullString, "Argentum Online")
'GetWindowThreadProcessId handle, ProcessID
'ProcessHandle = OpenProcess(&H1F0FFF, True, ProcessID)
'aca habia puesto lo mismo que formload del form
'ProcessHandle = myHandle

ProcessHandle = mm.hProcess

For i = 1 To NumOffsets
    If i = 1 Then
        ReadProcessMem ProcessHandle, CLng(address), PointerValue, 4&, 0
    Else
        ReadProcessMem ProcessHandle, Offsets(i - 1).DecimAddress, PointerValue, 4&, 0
    End If
    AddressDec = PointerValue + Offsets(i).DecimValue
    Offsets(i).DecimAddress = AddressDec
    Offsets(i).HexAddress = Hex(AddressDec)
Next i

FAddressDecim = Offsets(NumOffsets).DecimAddress
ReadProcessMem ProcessHandle, FAddressDecim, FValueDecim, 4&, 0

FValueDecim = FValueDecim + 0

FAddressHex = Hex(AddressDec)
FValueHex = Hex(FValueDecim)

MiseryCalc = FAddressDecim

'Exit Function
'Err:
'    Exit Function
End Function

Public Function MiseryCalc2(ByVal address As Long, ParamArray TheOffsets() As Variant) As Variant
'On Error GoTo Err:

Dim i As Byte
Dim handle As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
Dim PointerValue As Long
Dim AddressDec As Long
Dim AddressHex As String

'MsgBox UBound(TheOffsets) '0, 1
NumOffsets = UBound(TheOffsets) + 1
'MsgBox NumOffsets
'Exit Sub

ReDim Offsets(NumOffsets)

For i = 1 To NumOffsets
    ActualOffsetVal = TheOffsets(i - 1)
    'MsgBox ActualOffsetVal
    
    Offsets(i).HexValue = ActualOffsetVal
    Offsets(i).DecimValue = ActualOffsetVal
Next i

'handle = FindWindow(vbNullString, "Argentum Online")
'GetWindowThreadProcessId handle, ProcessID
'ProcessHandle = OpenProcess(&H1F0FFF, True, ProcessID)
'ProcessHandle = myHandle

ProcessHandle = mm.hProcess

For i = 1 To NumOffsets
    If i = 1 Then
        ReadProcessMem ProcessHandle, address, PointerValue, 4&, 0
    Else
        ReadProcessMem ProcessHandle, Offsets(i - 1).DecimAddress, PointerValue, 4&, 0
    End If
    AddressDec = PointerValue + Offsets(i).DecimValue
    Offsets(i).DecimAddress = AddressDec
    Offsets(i).HexAddress = Hex(AddressDec)
Next i

FAddressDecim = Offsets(NumOffsets).DecimAddress
ReadProcessMem ProcessHandle, FAddressDecim, FValueDecim, 4&, 0

FValueDecim = FValueDecim + 0

FAddressHex = Hex(AddressDec)
FValueHex = Hex(FValueDecim)

MiseryCalc2 = FAddressDecim

'Exit Function
'Err:
'    Exit Function
End Function

Public Function CalcularBytes(ByVal address As Long) As String
Dim i As Byte
Dim AddressHex As Variant
Dim NAH As Variant

AddressHex = Hex(address)

AddressHex = "0000000" & (AddressHex)

NAH = Right(AddressHex, 8)

'jne 12345678
'XX -XX - L2, L1 - L4, L3 - L6, L5 - L8, L7

For i = 1 To 8
    v(9 - i) = "&H" & Mid(NAH, i, 1)
Next i

OffSet(1) = &H3
'OffSet(2) = &H6

'OffSet(3) = &HA
'OffSet(4) = &H2

OffSet(5) = &H7
OffSet(6) = &HB

OffSet(7) = &HF
OffSet(8) = &HF

For i = 1 To 8
    L(i) = L(i) + v(i) + OffSet(i)
    
    If L(i) > &HF Then
        lR(i) = (L(i) - &H10)
        
        
        If i <> 8 Then
            lS(i) = (L(i) - lR(i))
            L(i + 1) = L(i + 1) + (lS(i) / &H10)
        End If
        
        '//FIX
        L(i) = lR(i)
    End If
Next i

'XX - XX - L2, L1 - L4, L3 - L6, L5 - L8, L7
'CalcularBytes = "0F - " & _
                "85 - " & _
                Hex(L(2)) & Hex(L(1)) & " - " & _
                Hex(L(4)) & Hex(L(3)) & " - " & _
                Hex(L(6)) & Hex(L(5)) & " - " & _
                Hex(L(8)) & Hex(L(7))

'0F 85 FC 04 00 00
'0x0 4  F C  85 0F
'  4,3  2,1  85 0F
CalcularBytes = Hex(L(4)) & Hex(L(3)) & Hex(L(2)) & Hex(L(1)) & "850F"
End Function

Public Function CalcularBytes2(ByVal address As Long) As String
Dim i As Byte
Dim AddressHex As Variant
Dim NAH As Variant

AddressHex = Hex(address)

AddressHex = "0000000" & (AddressHex)

NAH = Right(AddressHex, 8)

'jne 12345678
'XX -XX - L2, L1 - L4, L3 - L6, L5 - L8, L7

For i = 1 To 8
    v(9 - i) = "&H" & Mid(NAH, i, 1)
Next i

OffSet(1) = &H3
'OffSet(2) = &H6

'OffSet(3) = &HA
'OffSet(4) = &H2

OffSet(5) = &H7
OffSet(6) = &HB

OffSet(7) = &HF
OffSet(8) = &HF

For i = 1 To 8
    L(i) = L(i) + v(i) + OffSet(i)
    
    If L(i) > &HF Then
        lR(i) = (L(i) - &H10)
        
        
        If i <> 8 Then
            lS(i) = (L(i) - lR(i))
            L(i + 1) = L(i + 1) + (lS(i) / &H10)
        End If
        
        '//FIX
        L(i) = lR(i)
    End If
Next i

'XX - XX - L2, L1 - L4, L3 - L6, L5 - L8, L7
'CalcularBytes = "0F - " & _
                "85 - " & _
                Hex(L(2)) & Hex(L(1)) & " - " & _
                Hex(L(4)) & Hex(L(3)) & " - " & _
                Hex(L(6)) & Hex(L(5)) & " - " & _
                Hex(L(8)) & Hex(L(7))

'0F 85 FC 04 00 00
'0x0 4  F C  85 0F
'  4,3  2,1  85 0F
CalcularBytes2 = Hex(L(4)) & Hex(L(3)) & Hex(L(2)) & Hex(L(1)) & "850F"
End Function




