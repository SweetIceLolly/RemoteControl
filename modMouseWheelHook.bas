Attribute VB_Name = "modMouseWheelHook"
'���������������¼���ģ��

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long

Public Const GWL_WNDPROC = (-4)
Public Const WM_MOUSEWHEEL = &H20A

Public PrevWndProc As Long

Public Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_MOUSEWHEEL Then                '����ǹ����¼�
        If wParam < 0 Then                                  '����
            If IsControlling Then
                wsSendData frmRemoteControl.wsMessage, "MWD"
            End If
        Else                                                '����
            If IsControlling Then
                wsSendData frmRemoteControl.wsMessage, "MWU"
            End If
        End If
    Else
        WndProc = CallWindowProc(PrevWndProc, hwnd, uMsg, wParam, lParam)
    End If
End Function

