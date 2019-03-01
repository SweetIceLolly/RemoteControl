VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCMD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'=========================================================================================================================
'延时函数部分
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&
Private Const INFINITE = &HFFFF
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Const UNITS = 4294967296#
Private Const MAX_LONG = -2147483648#

Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long
Private mlTimer As Long
                                                                                '=========================================================================================================================
Private Type STARTUPINFO                                                        '(createprocess)
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hfile As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long

Private Type PROCESS_INFORMATION                                                '(creteprocess)
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadID As Long
End Type
Private Type SECURITY_ATTRIBUTES                                                '(createprocess)
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
                                                                                '常数声明
Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const STARTF_USESTDHANDLES = &H100&
Private Const STARTF_USESHOWWINDOW = &H1
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION = &H400
                                                                                '函数声明
Private Declare Function CreateProcessA Lib "kernel32" ( _
    ByVal lpApplicationName As Long, _
    ByVal lpCommandLine As String, _
    lpProcessAttributes As SECURITY_ATTRIBUTES, _
    lpThreadAttributes As SECURITY_ATTRIBUTES, _
    ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, _
    ByVal lpEnvironment As Long, _
    ByVal lpCurrentDirectory As Long, _
    lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" _
    (ByVal hNamedPipe As Long, _
    ByVal lpBuffer As Long, _
    ByVal nBufferSize As Long, _
    ByRef lpBytesRead As Long, _
    ByRef lpTotalBytesAvail As Long, _
    ByRef lpBytesLeftThisMessage As Long _
    ) As Long

Private Declare Function CreatePipe Lib "kernel32" ( _
    phReadPipe As Long, _
    phWritePipe As Long, _
    lpPipeAttributes As Any, _
    ByVal nSize As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hfile As Long, _
    ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Any) As Long

                                                                                'Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hHandle As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hfile As Long, _
    ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToWrite As Long, _
    ByRef lpNumberOfBytesWritten As Long, _
    lpOverlapped As Any) As Long

Private Declare Function DuplicateHandle Lib "kernel32" _
    (ByVal hSourceProcessHandle As Long, _
    ByVal hSourceHandle As Long, _
    ByVal hTargetProcessHandle As Long, _
    lpTargetHandle As Long, _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwOptions As Long) As Long
 
Private Const DUPLICATE_SAME_ACCESS = &H2
Private PipeR4InputChannel As Long, PipeW4InputChannel As Long, hInputHandle As Long
Private PipeR4OutputChannel As Long, PipeW4OutputChannel As Long, hOutputHandle As Long
Private Proc As PROCESS_INFORMATION
Public Enum InitResult
    ERROR_OK = 0
    ERROR_INIT_INPUT_HANDLE = 1
    ERROR_INIT_OUTPUT_HANDLE = 2
    ERROR_DUP_READ_HANDLE = 3
    ERROR_DUP_WRITE_HANDLE = 4
    ERROR_CREATE_CHILD_PROCESS = 5
End Enum

Public Enum TermResult
    ERROR_OK = 0
End Enum
Public Enum InputResult
    ERROR_OK = 0
    ERROR_QUERY_WRITE_INFO_SIZE = 1
    ERROR_DATA_TO_LARGE = 2
    ERROR_WRITE_INFO = 3
    ERROR_WRITE_UNEXPECTED = 5
End Enum
Public Enum OutputResult
    ERROR_OK = 0
    ERROR_QUERY_READ_INFO_SIZE = 1
    ERROR_ZERO_INFO_SIZE = 2
    ERROR_READ_INFO = 3
    ERROR_UNEQUAL_INFO_SIZE = 4
    ERROR_READ_UNEXPECTED = 5
    ERROR_QUERY_INFO_SIZE = 6
End Enum

Private Declare Function timeGetTime Lib "winmm.dll" () As Long

Dim OutPutData As String, ReadData As Boolean, OutData As String
Private IsExit As Boolean
Public Event Events()

'===========================================================================================
Private Sub Class_Initialize()
    If InitDosIO <> 0 Then
    'MsgBox "DOS通道创建失败！"
    End If
End Sub


Private Sub Class_Terminate()
    On Error Resume Next
    IsExit = True
    Call EndDosIo
    If mlTimer <> 0 Then CloseHandle mlTimer
End Sub

Public Sub Wait(MilliSeconds As Long)                                           '延时函数
    On Error GoTo ErrHandler
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    
    mlTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer" & Format$(Now(), "NNSS"))
    
    If Err.LastDllError <> ERROR_ALREADY_EXISTS Then
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If
    
    dblDelay = CDbl(MilliSeconds) * 10000#
    
    ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    
    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    
    Do
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0 And IsExit = False
    
    CloseHandle mlTimer
    mlTimer = 0
    Exit Sub
    
ErrHandler:
    Err.Raise Err.Number, Err.Source, "[clsWaitableTimer.Wait]" & Err.Description
End Sub

                                                                                '========================================================================================

Private Function InitDosIO() As InitResult
    Dim Sa As SECURITY_ATTRIBUTES, ret As Long
    
    With Sa
        .nLength = Len(Sa)
        .bInheritHandle = 1&
        .lpSecurityDescriptor = 0&
    End With
    
    ret = CreatePipe(PipeR4InputChannel, PipeW4InputChannel, Sa, 1024&)
    If ret = 0 Then                                                             '建立进程输入管道
        InitDosIO = ERROR_INIT_INPUT_HANDLE
        Exit Function
    End If
    
    ret = CreatePipe(PipeR4OutputChannel, PipeW4OutputChannel, Sa, 4096&)       '建立输出通道,若建立失败，则关闭管道，退出
    If ret = 0 Then                                                             '建立进程的输出管道
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        InitDosIO = ERROR_INIT_OUTPUT_HANDLE
        Exit Function
    End If
    
    
    ret = DuplicateHandle(GetCurrentProcess(), PipeW4InputChannel, GetCurrentProcess(), hInputHandle, 0, True, DUPLICATE_SAME_ACCESS)
    If ret = 0 Then                                                             '转换写句柄
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        CloseHandle PipeR4OutputChannel
        CloseHandle PipeW4OutputChannel
        InitDosIO = ERROR_DUP_WRITE_HANDLE
        Exit Function
    End If
    ret = CloseHandle(PipeW4InputChannel)
    If ret = 0 Then
                                                                                'MsgBox "close handle eerr"
    End If
    ret = DuplicateHandle(GetCurrentProcess(), PipeR4OutputChannel, GetCurrentProcess(), hOutputHandle, 0, True, DUPLICATE_SAME_ACCESS)
    If ret = 0 Then                                                             '转换读句柄
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        CloseHandle PipeR4OutputChannel
        CloseHandle PipeW4OutputChannel
        InitDosIO = ERROR_DUP_READ_HANDLE
        Exit Function
    End If
    ret = CloseHandle(PipeR4OutputChannel)
    If ret = 0 Then
                                                                                'MsgBox "close handle 2 er"
    End If
    
    
    Dim Start As STARTUPINFO, CmdStr As String
    Start.cb = Len(Start)
    Start.dwFlags = STARTF_USESTDHANDLES Or STARTF_USESHOWWINDOW
    Start.hStdOutput = PipeW4OutputChannel
    Start.hStdError = PipeW4OutputChannel
    Start.hStdInput = PipeR4InputChannel
    CmdStr = "CMD"                                                              '需要打开的命令行外壳(CMD.EXE)
    ret& = CreateProcessA(0&, CmdStr, Sa, Sa, True, NORMAL_PRIORITY_CLASS, 0&, 0&, Start, Proc)
    
    If ret <> 1 Then                                                            '建立控制进程
        CloseHandle PipeR4InputChannel
        CloseHandle PipeW4InputChannel
        CloseHandle PipeR4OutputChannel
        CloseHandle PipeW4OutputChannel
        InitDosIO = ERROR_CREATE_CHILD_PROCESS
        Exit Function
    End If
    ReadData = True
End Function

Public Function DosInput(ByVal Str As String) As InputResult
    Dim Buflen As Long, BtWritten As Long, rtn As Long
    Dim BtTest() As Byte
    ReadData = True
    BtTest = StrConv(Str + vbCrLf, vbFromUnicode)
    Buflen = UBound(BtTest) + 1
    rtn = WriteFile(hInputHandle, StrPtr(BtTest), Buflen, BtWritten, ByVal 0&)
    
    If BtWritten = 0 Then
        DosInput = ERROR_WRITE_INFO
        Exit Function
    End If
    DosInput = 0
End Function

Public Function DosOutPutEx(Optional TimeOut As Long = 20000) As String         '默认命令超时20秒
    Dim OutPt As String, OldTime As Long, NewTime As Long, tmp() As String
    If ReadData = False Then Exit Function
    OldTime = timeGetTime
    OutPutData = ""
    Do
        DoEvents
        If DosOutput(OutPt) = 0 Then
            OutData = Left$(OutPt, Len(OutPt) - 1)
            OutPutData = OutPutData & Left$(OutPt, Len(OutPt) - 1)
            Call NowEvents
            tmp = Split(OutPt, Chr$(13))
            If Len(tmp(UBound(tmp))) <> 0 Then
                If InStr(tmp(UBound(tmp)), ":\") = 3 And Right(tmp(UBound(tmp)), 2) = ">" & Chr$(0) Then Exit Do
            End If
        End If
        NewTime = timeGetTime
        Call Wait(10)                                                           '如需要更高实时性可以改为1
        If NewTime - OldTime >= TimeOut Then Exit Do
    Loop
    DosOutPutEx = OutPutData
    ReadData = False
End Function

Public Function GetOut() As String '实时数据
    GetOut = OutData
End Function

Friend Sub NowEvents() '接口
    RaiseEvent Events
End Sub

Private Function DosOutput(ByRef StrOutput As String) As OutputResult
    Dim ret As Long, TmpBuf As String * 128, BtRead As Long, BtTotal As Long, BtLeft As Long
    Dim rtn As Long, lngbytesread As Long
    
    rtn = PeekNamedPipe(hOutputHandle, StrPtr(TmpBuf), 128, BtRead, BtTotal, BtLeft)
    If rtn = 0 Then                                                             '查询信息量
        DosOutput = ERROR_QUERY_INFO_SIZE
        Exit Function
    End If
    
    If BtTotal = 0 Then                                                         '若信息为空，则退出
        DosOutput = ERROR_ZERO_INFO_SIZE
        Exit Function
    End If
    
    Dim Btbuf() As Byte, BtReaded As Long
    ReDim Btbuf(BtTotal)
    ret = ReadFile(hOutputHandle, VarPtr(Btbuf(0)), BtTotal, lngbytesread, 0&)
    
    If ret = 0 Then
        DosOutput = ERROR_READ_INFO
        Exit Function
    End If
    If BtTotal <> lngbytesread Then
        DosOutput = ERROR_UNEQUAL_INFO_SIZE
    End If
    
    Dim strBuf As String
    strBuf = StrConv(Btbuf, vbUnicode)
                                                                                
    StrOutput = strBuf
    
End Function

Private Function EndDosIo() As Long
 Dim ret As Long
 CloseHandle PipeR4InputChannel
 CloseHandle PipeW4InputChannel
 CloseHandle PipeR4OutputChannel
 CloseHandle PipeW4OutputChannel
 CloseHandle Proc.hThread
 CloseHandle Proc.hProcess
If EndProcess(Proc.dwProcessId) = False Then
   'MsgBox "主服务程序[CMD.EXE]没有关闭，请您手动关闭 ", vbInformation, "不好意思"
End If
End Function

Private Function EndProcess(ByVal ProcessID As Long) As Boolean
    Dim hProcess As Long, ExitCode As Long, Rst As Long
    hProcess = OpenProcess(PROCESS_TERMINATE Or PROCESS_QUERY_INFORMATION, True, ProcessID)
    If hProcess <> 0 Then
        GetExitCodeProcess hProcess, ExitCode
        If ExitCode <> 0 Then
            Rst = TerminateProcess(hProcess, 0)
            CloseHandle hProcess
            If Rst = 0 Then
                
                EndProcess = False
            Else
                EndProcess = True
            End If
        Else
            EndProcess = False
        End If
    Else
        EndProcess = False
    End If
    
End Function
 


