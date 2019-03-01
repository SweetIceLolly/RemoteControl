VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmDownload 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���� (1/1)"
   ClientHeight    =   1344
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   5508
   Icon            =   "frmDownload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1344
   ScaleWidth      =   5508
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3840
      Top             =   840
   End
   Begin MSWinsockLib.Winsock wsFile 
      Left            =   3360
      Top             =   840
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBar 
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5292
      _Version        =   786432
      _ExtentX        =   9334
      _ExtentY        =   444
      _StockProps     =   93
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton cmdOpenLocalDir 
      Height          =   372
      Left            =   4320
      TabIndex        =   3
      Top             =   864
      Width           =   1092
      _Version        =   786432
      _ExtentX        =   1926
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "�ж�����"
      ForeColor       =   16777215
      BackColor       =   16416003
      Appearance      =   3
      DrawFocusRect   =   0   'False
      ImageGap        =   1
   End
   Begin VB.Label labStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2333 GB/s 2333 GB/2333 GB 100%"
      Height          =   180
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   2880
   End
   Begin VB.Label labFile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "�������أ�C:\Windows\System32\I Love CXT.exe"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4164
   End
End
Attribute VB_Name = "frmDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public FileList As String                                   '�ļ��б�
Public IsUpload As Boolean                                  '�Ƿ�Ϊ�ϴ�״̬ ��True:�ϴ� False:���ء�
Public CurrentSize As Long                                  '��ǰ�ļ���С
Public CurrentFileIndex As Integer                          '��ǰ���ļ����
Public TotalFiles As Integer                                '�ļ�����
Public DownloadPath As String                               '�ļ����ص�·��
Public TotalSent As Integer                                 '�ļ�������

Public lSent As Long                                        '�Ѿ����͵������ֽ���
Public oldSent As Long                                      '��һ�뷢�͵������ֽ���
Public BytesPerSec As Long                                  'ÿ�������ֽ���

Public lRec As Long                                         '���յ��ֽ���
Public lOldRec As Long                                      '��һ����յ��ֽ���

Dim Temp() As Byte                                          '�򿪵��ļ�������
Dim EachFile() As String                                    'ÿ���ļ���·��

Public Function LoadFile(sFileName As String) As Boolean    '��ȡ�ļ�����
    On Error Resume Next
    '-----------------------------------
    LoadFile = True
    Open sFileName For Binary As #1                             '���ļ�
        If Err.Number <> 0 Then                                 '�ļ�������
            If sFileName <> "" Then                                 '�������Split���µĿ��ļ���
                frmFileTransfer.AddLog "�ļ� " & sFileName & " ��ȡʧ�ܡ��������ϴ����ļ���"
            End If
            Close #1                                            '�ر��ļ�
            LoadFile = False
            Exit Function
        End If
        '--------------------
        ReDim Temp(LOF(1))                                      '�����ڴ�
        CurrentSize = LOF(1)
        If Err.Number <> 0 Then                                 '�ļ�����
            frmFileTransfer.AddLog "�ļ�����" & sFileName & " �������ϴ����ļ���"
            LoadFile = False
            Exit Function
        End If
        If UBound(Temp) = 0 Then                                '�����ȡ�����ֽ���Ϊ0
            frmFileTransfer.AddLog "�ļ� " & sFileName & " ��ȡʧ�ܣ����ļ������ܱ�����Ϊ���ļ����������ϴ����ļ���"
            Close #1                                            '�ر��ļ�
            LoadFile = False
            Exit Function
        End If
        Get #1, , Temp                                          '��ȡ�ļ�
    Close #1
    lSent = 0                                               '����ļ������ֽ���
    oldSent = 0
End Function

Public Function NextFile() As Boolean                       '��һ���ļ�����
    CurrentFileIndex = CurrentFileIndex + 1                     '�ļ���� + 1
    If CurrentFileIndex > TotalFiles Then                       '�����ų����ļ�����
        NextFile = False
        frmFileTransfer.AddLog "�ϴ����������������" & TotalSent & "���ļ���"
        Unload Me
        Exit Function
    End If
    '---------------------------
    Do While LoadFile(EachFile(CurrentFileIndex)) = False   '�����ȡ�ļ�����ͼ�����һ���ļ�
        CurrentFileIndex = CurrentFileIndex + 1                     '�ļ���� + 1
        If CurrentFileIndex > TotalFiles Then                       '�����ų����ļ�����
            NextFile = False
            frmFileTransfer.AddLog "�ϴ����������������" & TotalSent & "���ļ���"
            Exit Function
        End If
    Loop
    TotalSent = TotalSent + 1
    Me.labFile.Caption = "�����ϴ���" & EachFile(CurrentFileIndex)
    Me.Caption = "���� (" & CurrentFileIndex + 1 & "/" & TotalFiles & ")"
    frmFileTransfer.AddLog "�����ϴ���" & EachFile(CurrentFileIndex)
    NextFile = True
End Function

Public Function SplitPath()                                 '�����ÿ���ļ���·������
    EachFile = Split(FileList, vbCrLf)                          '�ָ�
    TotalFiles = UBound(EachFile)                               '������ļ�����
End Function

Public Function SendUploadMsg()                             '�����ϴ���Ϣ����
    '�����ϴ��ļ����� ������ͷ|�ļ�·��|�ļ���С��
    wsSendData frmFileTransfer.wsMessage, "UPL|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If IsWorking Then                   '�����ǰ�����ڽ����е�����
        Dim a
        a = MsgBox("ȷ�Ϲرձ����ڣ���ǰ���е������жϣ�", 32 + vbYesNo, "ȷ��")
        If a <> 6 Then                      'ȡ���ر�
            Cancel = True
        Else                                'ȷ�Ϲر�
            frmFileTransfer.AddLog "�ļ��ϴ��жϡ�������" & TotalSent & "���ļ���"
            wsSendData frmFileTransfer.wsMessage, "END"     '���ͽ����ϴ���Ϣ
            wsSendData frmFileTransfer.wsMessage, "REF"     '����ˢ������
        End If
    End If
End Sub

Private Sub tmrRefresh_Timer()
    If IsUpload Then
        BytesPerSec = lSent - oldSent                           '����ÿ�뷢�͵��ֽ���
        oldSent = lSent                                         '��һ����ֽ��������һ���
    Else
        BytesPerSec = lRec - lOldRec                            '����ÿ����յ��ֽ���
        lOldRec = lRec                                          '��һ����ֽ��������һ���
    End If
End Sub

Private Sub wsFile_ConnectionRequest(ByVal requestID As Long)
    '���յ���������
    Me.wsFile.Close
    Me.wsFile.Accept requestID                              '��������
    If IsUpload Then                                        '�ж��Ƿ�Ϊ�ϴ��ļ�״̬
        Me.wsFile.SendData "CNT"                                '���ӽ���ȷ����Ϣ
    Else
        Me.wsFile.SendData "DNT"                                '�öԷ�׼������
    End If
End Sub

Private Sub wsFile_DataArrival(ByVal bytesTotal As Long)
    Dim recTemp() As Byte               '���յ�����Ϣ�Ļ���
    Dim a As VbMsgBoxResult             'Msgbox ����ֵ
    '------------------------
    Me.wsFile.GetData recTemp
    '�õ�ǰ��λ
    Select Case Chr(recTemp(0)) & Chr(recTemp(1)) & Chr(recTemp(2))
        Case "TSN"                                          '�ļ�����ȷ��
            a = MsgBox("�ļ�������" & EachFile(CurrentFileIndex) & "���Ƿ�ȷ�ϸ��ǣ�" & vbCrLf & "������ǡ��������ļ���������񡿽�ȡ�����͸��ļ���", 32 + vbYesNo, "�ļ�����ȷ��")
            If a = 6 Then
                frmFileTransfer.AddLog "ȷ�ϸ����ļ� " & EachFile(CurrentFileIndex)
                Me.wsFile.SendData "YES"                        '���� ȷ�ϸ��� ��Ϣ
            Else
                frmFileTransfer.AddLog "ȡ�������ļ� " & EachFile(CurrentFileIndex)
                If NextFile = True Then
                    '����ɹ�����һ���ļ����ϴ���һ���ļ�����
                    wsSendData frmFileTransfer.wsMessage, "NXT|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
                Else
                    wsSendData frmFileTransfer.wsMessage, "END"     '��ʧ�ܾͷ��ͽ����ϴ���Ϣ
                    wsSendData frmFileTransfer.wsMessage, "REF"     '����ˢ������
                    IsWorking = False                               '��ǰû������
                    Unload Me
                End If
            End If
        
        Case "IRD"                                          '�Է�׼������
            Me.tmrRefresh.Enabled = True                        '����ͳ�Ƽ�ʱ��
            Me.wsFile.SendData Temp                             '��ʼ������ (�s�F����)�s��ߩ���
            
        Case "DFL"                                          '����ʧ��
            frmFileTransfer.AddLog "�ļ� " & EachFile(CurrentFileIndex) & " �ϴ�ʧ�ܡ�ԭ�򣺸����ļ�ʧ��"
            If NextFile = True Then
                '����ɹ�����һ���ļ����ϴ���һ���ļ�����
                wsSendData frmFileTransfer.wsMessage, "NXT|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
            Else
                wsSendData frmFileTransfer.wsMessage, "END"     '��ʧ�ܾͷ��ͽ����ϴ���Ϣ
                wsSendData frmFileTransfer.wsMessage, "REF"     '����ˢ������
                IsWorking = False                               '��ǰû������
                Unload Me
            End If
        
        Case "NXT"                                          '��һ�ļ�����
            If NextFile = True Then
                '����ɹ�����һ���ļ��ͷ����ϴ���һ���ļ�����
                wsSendData frmFileTransfer.wsMessage, "NXT|" & EachFile(CurrentFileIndex) & "|" & CStr(CurrentSize)
            Else
                wsSendData frmFileTransfer.wsMessage, "END"     '��ʧ�ܾͷ��ͽ����ϴ���Ϣ
                wsSendData frmFileTransfer.wsMessage, "REF"     '����ˢ������
                IsWorking = False                               '��ǰû������
                Unload Me
            End If
        
        Case "NNT"                                          '׼���ý�����һ���ļ�
            Me.wsFile.SendData "CNT"
        
        Case "DNT"
            Dim nFileTitle As String, nFileSize As String       '�ļ����� �ļ���С
            Dim SplitTmp() As String                            '�ָ��
            Dim recString As String                             '�ѽ��յ�������ת�����ַ����ı���
            Dim recTotal As Integer, recCurrent As Integer      '��Ҫ���ܵ��ļ����� ��ǰ���յ��ļ����
            '----------------------------------
            recString = StrConv(recTemp, vbUnicode)                 '������������ת�����ַ���
            '----------------------------------
            SplitTmp = Split(Replace(recString, "DNT|", ""), "|")   'ȥ������ͷ
            nFileTitle = SplitTmp(0)                                '������ļ�����
            nFileSize = SplitTmp(1)                                 '������ļ���С
            recCurrent = SplitTmp(2)                                '�������ǰ���յ��ļ����
            recTotal = SplitTmp(3)                                  '�������Ҫ���ܵ��ļ�����
            Me.Caption = "���� (" & recCurrent + 1 & "/" & recTotal & ")"   '���±���
            lRec = 0                                                '��ս����ֽ���
            lOldRec = 0
            frmFileTransfer.AddLog "��ʼ�����ļ���" & nFileTitle
            Me.wsFile.Tag = nFileTitle & "|" & nFileSize            '��ֵ���ؼ���Tag
            Me.labFile.Caption = "�������أ�" & nFileTitle          '���´���״̬
            '----------------------------------
            If IsPathExists(DownloadPath & nFileTitle, True) = True Then    '����ļ��Ѵ���
                a = MsgBox("��⵽�ļ��Ѵ��ڣ�" & DownloadPath & nFileTitle & "���Ƿ�ȷ�ϸ��ǣ�" & vbCrLf & "������ǡ��������ļ���������񡿽�ȡ�����͸��ļ���", 32 + vbYesNo, "�ļ�����ȷ��")
                If a = 6 Then                                                   'ȷ�ϸ���
                    Kill DownloadPath & nFileTitle                                  '����ɾ���ļ�
                    If Err.Number = 0 Then                                          'ɾ���ɹ�
                        frmFileTransfer.AddLog "�����ļ� " & DownloadPath & nFileTitle
                        Open DownloadPath & nFileTitle For Binary As #2                 '���ļ�
                        Me.wsFile.SendData "IRD"                                        '���߶Է�׼������
                    Else                                                            'ɾ��ʧ��
                        frmFileTransfer.AddLog "�����ļ� " & DownloadPath & nFileTitle & " ʧ�ܣ���ȡ�������ļ� " & nFileTitle
                        Me.wsFile.SendData "NXT"                                        '���߶Է���һ�ļ�
                    End If
                Else
                    frmFileTransfer.AddLog "ȡ�������ļ� " & DownloadPath & nFileTitle
                    Me.wsFile.SendData "NXT"                                        '���߶Է���һ�ļ�
                End If
            Else                                                            '����ļ�������
                Open DownloadPath & nFileTitle For Binary As #2                 '���ļ�
                Me.wsFile.SendData "IRD"                                        '�����Է�׼������
            End If
            
        Case "END"                                                      '�������
            frmFileTransfer.AddLog "�������������������" & Split(StrConv(recTemp, vbUnicode), "|")(1) & "���ļ���"
            IsWorking = False                                           '��ǰû������
            frmFileTransfer.cmdRefreshLocal_Click                       'ˢ�±���Ŀ¼
            Unload Me
            
        Case Else                                                       '�������������
            '�õ�����λ
            '����ǡ�FIN�����������ļ��������
            If recTemp(UBound(recTemp) - 2) = 70 And recTemp(UBound(recTemp) - 1) = 73 And recTemp(UBound(recTemp)) = 78 Then
                frmFileTransfer.AddLog "�ļ� " & Split(Me.wsFile.Tag, "|")(0) & " ������ϡ�"
                Close #2                                                    '�ر��ļ�
                Me.wsFile.SendData "NXT"                                    '���߶Է���һ�ļ�
            Else                                                        '�����Ȼ���������� ˵����������
                Put #2, LOF(2) + 1, recTemp                                 'д���ļ�
                lRec = lRec + UBound(recTemp)
                '-----------------------------------------
                '���´���״̬
                Dim tmpShow As String                                       '��ǩ����ʾ�����ݵĻ���
                Dim CurrentFileSize As Long                                 '��ǰ�ļ���С
                CurrentFileSize = CLng(Split(Me.wsFile.Tag, "|")(1))                            '��Tag��ȡ��ǰ�ļ���С
                tmpShow = SizeWithFormat(BytesPerSec) & "/s"                                    'ÿ���ֽ���
                tmpShow = tmpShow & " " & SizeWithFormat(lRec)                                  '�Ѿ����͵Ĵ�С
                tmpShow = tmpShow & "/" & SizeWithFormat(Split(Me.wsFile.Tag, "|")(1))          '�ļ��Ĵ�С
                tmpShow = tmpShow & " " & Format((lRec / CurrentFileSize) * 100, "0.00") & "%"  '����ٷֱ�
                Me.ProgressBar.Value = (lRec / CurrentFileSize) * 100                           '����������ɵİٷֱ�
                Me.labStatus.Caption = tmpShow                                                  '��ʾ����ǩ��
                Me.labStatus.Refresh                                                            'ˢ�±�ǩ
            End If
        
    End Select
End Sub

Private Sub wsFile_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If IsUpload Then
        lSent = lSent + bytesSent
        If lSent >= CurrentSize Then
            lSent = 0
            frmFileTransfer.AddLog "�ļ�" & EachFile(CurrentFileIndex) & "�ϴ��ɹ���"
            Me.wsFile.SendData "FIN"
            Exit Sub
        End If
        '=======================================================
        '���´���״̬
        Dim tmpShow As String                                                       '��ʾ���ı�����
        tmpShow = SizeWithFormat(BytesPerSec) & "/s"                                'ÿ���ֽ��� = ���ڷ��͵������ֽ��� - ��һ�뷢�͵������ֽ���
        tmpShow = tmpShow & " " & SizeWithFormat(lSent)                             '�Ѿ����͵Ĵ�С
        tmpShow = tmpShow & "/" & SizeWithFormat(CurrentSize)                       '�ļ��Ĵ�С
        tmpShow = tmpShow & " " & Format(lSent / CurrentSize * 100, "0.00") & "%"   '����ٷֱ�
        Me.ProgressBar.Value = lSent / CurrentSize * 100
        Me.labStatus.Caption = tmpShow                                              '��ʾ����ǩ��
        Me.labStatus.Refresh
    End If
End Sub
