Attribute VB_Name = "modProcessManager"
'进程管理模块

Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Public Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Public Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF
Public Const TH32CS_SNAPPROCESS As Long = 2&

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type

Public Function GetFileName(TargetPID As Long) As String            '获取指定进程路径的过程
    Dim lpid As Long
    Dim sBuffer As String
    Dim hHandle As Long
    Dim hInst As Long
    Dim lRet As Long
    
    sBuffer = Space(255)
    hHandle = OpenProcess(PROCESS_ALL_ACCESS, 0, TargetPID)
    lRet = GetModuleFileNameExA(hHandle, 0, sBuffer, 255)
    CloseHandle hHandle
    GetFileName = Trim(sBuffer)
    If lRet = 0 Then
        GetFileName = "获取失败"
    End If
End Function

Public Function KillPID(Process_Id As Long) As Boolean              '结束拥有指定PID的进程的过程
    Dim lProcess As Long
    Dim lExitCode As Long
    Dim x As Long
    '=============================================
    lProcess = OpenProcess(1, False, Process_Id)
    If bKillTasks = True Then
        KillPID = CBool(TerminateProcess(lProcess, lExitCode))
    End If
    CloseHandle lProcess
End Function

