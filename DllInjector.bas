Attribute VB_Name = "DllInjector"
'Begin Code


Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal fAllocType As Long, FlProtect As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, ByVal lpStartAddress As Any, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public CustomKeys As New clsCustomKeys
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Public Function GetHProcExe(strExeName As String) As Long
Dim hSnap As Long
    hSnap = CreateToolhelpSnapshot(2, 0)
Dim peProcess As PROCESSENTRY32
    peProcess.dwSize = LenB(peProcess)
Dim nProcess As Long
    nProcess = Process32First(hSnap, peProcess)
    Do While nProcess
        If StrComp(Trim$(peProcess.szExeFile), strExeName, vbTextCompare) _
            = 0 Then
            GetHProcExe = OpenProcess(PROCESS_ALL_ACCESS, False, peProcess.th32ProcessID)
            Exit Function
        End If
        peProcess.szExeFile = vbNullString
        nProcess = Process32Next(hSnap, peProcess)
    Loop
    CloseHandle hSnap
End Function
Public Function FindProc(ProcName As String) As Long
Dim hwnd As Long
Dim ProcessID As Long
Dim ProcessHandle As Long
hwnd = FindWindow(vbNullString, ProcName)
GetWindowThreadProcessId hwnd, ProcessID
ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, ProcessID)
FindProc = ProcessHandle
End Function
Public Function InjectDll(DllPath As String, ProsH As Long)
Dim DLLVirtLoc As Long, DllLength, Inject As Long, LibAddress As Long
Dim CreateThread As Long, ThreadID As Long
DllLength = Len(DllPath)
DLLVirtLoc = VirtualAllocEx(ProsH, ByVal 0, DllLength, &H1000, ByVal &H4)
If DLLVirtLoc = 0 Then q.Label1.Caption = "VirtualAllocEx API failed!": Exit Function
Inject = WriteProcessMemory(ProsH, DLLVirtLoc, ByVal DllPath, DllLength, vbNull)
If Inject = 0 Then q.Label1.Caption = "Failed to Write DLL to Process!"
LibAddress = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")
If LibAddress = 0 Then q.Label1.Caption = "Can"
CreateThread = CreateRemoteThread(ProsH, vbNull, 0, LibAddress, DLLVirtLoc, 0, ThreadID)
If CreateThread = 0 Then q.Label1.Caption = "Failed to Create Thead!"
End Function


