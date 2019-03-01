Attribute VB_Name = "modSubs"
Public Const EncryptNumber = 13                 '�ܳ� (*^-^*)

'==========================================================================================
'ע�⣺���¿����ڱ���ǰ����ȫ���趨ΪTrue������������п��ܲ�������
Public Const bCompress = True                   '�����ÿ��أ��Ƿ�ѹ������
Public Const bEncrypt = True                    '�����ÿ��أ��Ƿ������������
Public Const bKillTasks = True                  '�����ÿ��أ��Ƿ��������
Public Const bKeepSending = True                '�����ÿ��أ��Ƿ���������ͼƬ
Public Const bMouseWheelHook = True             '�����ÿ��أ��Ƿ�������������Ϣ
Public Const bBlockInput = True                 '�����ÿ��أ��Ƿ��������������
Public Const bDeleteFiles = True                '�����ÿ��أ��Ƿ�����ɾ���ļ�

'==========================================================================================
'��������Ϊ�����ļ��������õ�����
Public bAutoStart As Boolean                    '�Ƿ񿪻��Զ�����
Public bTrayWhenStart As Boolean                '�������Ƿ���С����������
Public bHideMode As Boolean                     '�Ƿ�������ģʽ����
Public sUserPassword As String                  '�û�����
Public bUseUserPassword As Boolean              '�Ƿ�ʹ���û�����
Public bAutoRecord As Boolean                   '�Ƿ��Զ���¼IP������
Public bHideWhenStart As Boolean                '�Ƿ��Զ��������Ժ�̨ģʽ����
Public bNoExit As Boolean                       '�Ƿ��ֹ�˳�
Public bExitWhenClose As Boolean                '�ر�ʱ�Ƿ��˳� ��True���ر�ʱ�˳���� False���ر�ʱ��С������������

'==========================================================================================
'����ΪԶ��ʱ��״̬
Public IsRemoteControl As Boolean               '�Ƿ�ΪԶ�̿���״̬ ��True:Զ�̿���ģʽ False���ļ�����ģʽ��
Public IsControlling As Boolean                 '�Ƿ���ƶԷ�������
Public IsWorking As Boolean                     '��ǰ�Ƿ����ϴ�������������������ڽ���
Public AutoResize As Boolean                    '�Ƿ����컭��
Public ScreenScaleW As Double                   '��С��Ļ��ԭʼ��Ļ�Ŀ�ȱ�����
Public ScreenScaleH As Double                   '��С��Ļ��ԭʼ��Ļ�ĸ߶ȱ�����

Public Function Encrypt(strEncrypt As String) As String     '���ܣ����Խ����������Ŷ~��
    Dim tmp As String
    '==============================
    If bEncrypt = False Then                    '���ݵ��Կ�����ֵ�ж��Ƿ���мӽ���
        Encrypt = strEncrypt
        Exit Function
    End If
    For i = 1 To Len(strEncrypt)
        tmp = tmp & Chr(Asc(Left(Mid(strEncrypt, i), 1)) Xor EncryptNumber)   '������Xor���ܼ򵥴ֱ�������������
    Next i
    Encrypt = tmp
End Function

Public Function Decrypt(strDecrypt As String) As String     '����
    If bEncrypt = False Then                                    '���ݵ��Կ�����ֵ�ж��Ƿ���мӽ���
        Decrypt = strDecrypt
        Exit Function
    End If
    Decrypt = Encrypt(strDecrypt)                               '�ٵ���һ�μ��ܽ�����������ǽ�����~
End Function

Public Sub wsSendData(TargetWinsock As Winsock, DataToSend As String)         '�������ݹ���
    On Error Resume Next                                    '�������������磡��
    TargetWinsock.SendData Encrypt(DataToSend & "{S}{END}")
End Sub

Public Sub LoadConfig()                                     '���������ļ�����
    On Error Resume Next
    Dim tmpString As String                                 'ÿ������
    Dim tmpSplit() As String                                '�и����ַ�������
    Dim sIP() As String                                     'IP�Ͷ�Ӧ������Ļ���
    '----------------------------
    Open App.Path + "\Config.ini" For Input As #1               '���ļ�
        If Err.Number <> 0 Then                                 '���޷����ļ����ճ����Ĭ��ֵ������
            bAutoStart = False
            bTrayWhenStart = False
            bHideMode = False
            sUserPassword = "123456"
            '----------------------------
            Call SaveConfig                                         '�Զ����������ļ�
            Exit Sub                                                '�˳�����
        End If
        '----------------------------
        Do While Not EOF(1)
            Line Input #1, tmpString                            '���ж�ȡ����
            tmpSplit = Split(tmpString, "=")                    '�ָ�
            Select Case LCase(tmpSplit(0))                      '�������� ��ת��Сд��Ϊ����Ӧ��ǿ��������
                Case "autostart"                                    '�Ƿ񿪻��Զ�����
                    bAutoStart = CBool(tmpSplit(1))                 '����������ת��Ϊ������
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bAutoStart = False
                    End If
                    
                Case "traywhenstart"                                '�������Ƿ���С����������
                    bTrayWhenStart = CBool(tmpSplit(1))             '����������ת��Ϊ������
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bTrayWhenStart = False
                    End If
                
                Case "hidemode"                                     '�Ƿ�������ģʽ����
                    bHideMode = CBool(tmpSplit(1))                  '����������ת��Ϊ������
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bHideMode = False
                    End If
                
                Case "userpassword"                                 '�û�����
                    sUserPassword = Decrypt(tmpSplit(1))                '��ȡ��ʱ�����
                    
                Case "autoresize"                                   '�Ƿ��Զ�����ͼ��
                    AutoResize = CBool(tmpSplit(1))                 '����������ת��Ϊ������
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        AutoResize = False
                    End If
                    
                Case "useuserpassword"                              '�Ƿ�ʹ���û�����
                    bUseUserPassword = CBool(tmpSplit(1))           '����������ת��Ϊ������
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bUseUserPassword = False
                    End If
                
                Case "autorecord"                                   '�Ƿ��Զ���¼����
                    bAutoRecord = CBool(tmpSplit(1))                '����������ת��Ϊ������
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bAutoRecord = False
                    End If
                
                Case "hidewhenstart"                                '�Ƿ��Զ��������Ժ�̨ģʽ����
                    bHideWhenStart = CBool(tmpSplit(1))
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bHideWhenStart = False
                    End If
                            
                Case "noexit"                                       '�Ƿ��ֹ�˳�
                    bNoExit = CBool(tmpSplit(1))
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bNoExit = False
                    End If
                    
                Case "exitwhenclose"                                '�ر�ʱ�Ƿ��˳�
                    bExitWhenClose = CBool(tmpSplit(1))
                    If Err.Number <> 0 Then                         '�����ȡ����Ч���ݾ���Ĭ�ϵ�����
                        bExitWhenClose = False
                    End If
                
                Case Else
                    If InStr(tmpSplit(0), "[") = 0 Then             '������Ǳ�Ǿͼ�¼���б���
                        sIP = Split(tmpSplit(0), "����")                '�ָ�IP������ �������������ȡ��Ǹ�С�ʵ�Ӵ~�á����ȡ���Ϊ�ָ�������������ó���
                        frmMain.comIP.AddItem sIP(0)                    '���IP
                        frmMain.lstPassword.AddItem Decrypt(sIP(1))     '�������
                    End If
                    
            End Select
        Loop
    Close #1
End Sub

Public Sub SaveConfig()                                     '���������ļ�����
    Open App.Path + "\Config.ini" For Output As #1              '���ļ�
        Print #1, "[Config]"                                        '�ļ���һ�У���Ȼ�Ӳ������ж��޹ؽ�Ҫ�����Ǽ�������ÿ�~
        Print #1, "AutoStart=" & bAutoStart                         '�Ƿ񿪻�����
        Print #1, "TrayWhenStart=" & bTrayWhenStart                 '�������Ƿ���С����������
        Print #1, "HideMode=" & bHideMode                           '�Ƿ�������ģʽ����
        Print #1, "UserPassword=" & Encrypt(sUserPassword)          '�û����� ��Ϊ�˰�ȫ��������������ٱ��桿
        Print #1, "AutoResize=" & AutoResize                        '�Ƿ�����ͼ��
        Print #1, "UseUserPassword=" & bUseUserPassword             '�Ƿ�ʹ���û�����
        Print #1, "AutoRecord=" & bAutoRecord                       '�Ƿ��Զ���¼IP
        Print #1, "HideWhenStart=" & bHideWhenStart                 '�Ƿ��Զ��������Ժ�̨ģʽ����
        Print #1, "NoExit=" & bNoExit                               '�Ƿ��ֹ�˳�
        Print #1, "ExitWhenClose=" & bExitWhenClose                 '�ر�ʱ�Ƿ��˳�
        Print #1, "[List]"                                          '�����ļ��������ֵı�־��ͬ����Ϊ�����ۣ�����ȥ��Ҳ����ν~ �r(�s_�t)�q
        For i = 0 To frmMain.comIP.ListCount - 1                    '�б����IP�Ͷ�Ӧ������ ��ͬ����Ϊ�˰�ȫ��������������ٱ��桿
            Print #1, frmMain.comIP.List(i) & "����" & Encrypt(frmMain.lstPassword.List(i))
        Next i
    Close #1
End Sub
