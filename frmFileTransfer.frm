VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFileTransfer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "�ļ�����"
   ClientHeight    =   6432
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10248
   Icon            =   "frmFileTransfer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   10248
   StartUpPosition =   2  '��Ļ����
   Begin XtremeSuiteControls.ListView lstFileLocal 
      Height          =   3375
      Left            =   120
      TabIndex        =   22
      Top             =   1320
      Width           =   4815
      _Version        =   786432
      _ExtentX        =   8493
      _ExtentY        =   5953
      _StockProps     =   77
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      BackColor       =   16777215
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ListView lstFileRemote 
      Height          =   3375
      Left            =   5280
      TabIndex        =   23
      Top             =   1320
      Width           =   4815
      _Version        =   786432
      _ExtentX        =   8493
      _ExtentY        =   5953
      _StockProps     =   77
      Checkboxes      =   -1  'True
      View            =   3
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      BackColor       =   16777215
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin VB.FileListBox File 
      Height          =   252
      Hidden          =   -1  'True
      Left            =   2040
      Pattern         =   "*"
      System          =   -1  'True
      TabIndex        =   30
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.DirListBox Dir 
      Height          =   300
      Left            =   1440
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.DriveListBox Drive 
      Height          =   300
      Left            =   960
      TabIndex        =   28
      Top             =   120
      Visible         =   0   'False
      Width           =   270
   End
   Begin MSWinsockLib.Winsock wsMessage 
      Left            =   9600
      Top             =   4920
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.PictureBox picRemoteToolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5280
      ScaleHeight     =   372
      ScaleWidth      =   4812
      TabIndex        =   15
      Top             =   840
      Width           =   4815
      Begin XtremeSuiteControls.PushButton cmdRefreshRemote 
         Height          =   375
         Left            =   4320
         TabIndex        =   16
         ToolTipText     =   "ˢ��"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":0CCA
      End
      Begin XtremeSuiteControls.PushButton cmdDeleteRemote 
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         ToolTipText     =   "ɾ��ѡ�е��ļ�"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":1064
      End
      Begin XtremeSuiteControls.PushButton cmdMkdirRemote 
         Height          =   375
         Left            =   3285
         TabIndex        =   18
         ToolTipText     =   "�½��ļ���"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":13FE
      End
      Begin XtremeSuiteControls.PushButton cmdBackRemote 
         Height          =   375
         Left            =   2760
         TabIndex        =   19
         ToolTipText     =   "�����ϲ�Ŀ¼"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":1798
      End
      Begin XtremeSuiteControls.PushButton cmdHomeRemote 
         Height          =   375
         Left            =   2260
         TabIndex        =   20
         ToolTipText     =   "���ظ�Ŀ¼"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":1B32
      End
      Begin XtremeSuiteControls.PushButton cmdDownload 
         Height          =   375
         Left            =   0
         TabIndex        =   21
         ToolTipText     =   "�ӶԷ����������ļ�������"
         Top             =   0
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "����"
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":1ECC
         ImageAlignment  =   1
      End
   End
   Begin VB.PictureBox picLocalToolbar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   372
      ScaleWidth      =   4812
      TabIndex        =   8
      Top             =   840
      Width           =   4815
      Begin XtremeSuiteControls.PushButton cmdRefreshLocal 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "ˢ��"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":2266
      End
      Begin XtremeSuiteControls.PushButton cmdDeleteLocal 
         Height          =   375
         Left            =   500
         TabIndex        =   10
         ToolTipText     =   "ɾ��ѡ�е��ļ�"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":2600
      End
      Begin XtremeSuiteControls.PushButton cmdMkdirLocal 
         Height          =   375
         Left            =   1000
         TabIndex        =   11
         ToolTipText     =   "�½��ļ���"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":299A
      End
      Begin XtremeSuiteControls.PushButton cmdBackLocal 
         Height          =   375
         Left            =   1540
         TabIndex        =   12
         ToolTipText     =   "�����ϲ�Ŀ¼"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":2D34
      End
      Begin XtremeSuiteControls.PushButton cmdHomeLocal 
         Height          =   375
         Left            =   2040
         TabIndex        =   13
         ToolTipText     =   "���ظ�Ŀ¼"
         Top             =   0
         Width           =   495
         _Version        =   786432
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":30CE
      End
      Begin XtremeSuiteControls.PushButton cmdUpload 
         Height          =   375
         Left            =   3960
         TabIndex        =   14
         ToolTipText     =   "�ϴ�ѡ����ļ����Է�����"
         Top             =   0
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "�ϴ�"
         FlatStyle       =   -1  'True
         Appearance      =   4
         Picture         =   "frmFileTransfer.frx":3468
         ImageAlignment  =   1
         TextImageRelation=   4
      End
   End
   Begin XtremeSuiteControls.FlatEdit edLocalPath 
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   480
      Width           =   3975
      _Version        =   786432
      _ExtentX        =   7011
      _ExtentY        =   450
      _StockProps     =   77
      Text            =   "��Ŀ¼"
      BackColor       =   16777215
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit edRemotePath 
      Height          =   255
      Left            =   5880
      TabIndex        =   5
      Top             =   480
      Width           =   3975
      _Version        =   786432
      _ExtentX        =   7011
      _ExtentY        =   450
      _StockProps     =   77
      Text            =   "��Ŀ¼"
      BackColor       =   16777215
      Appearance      =   4
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdOpenLocalDir 
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "��"
      Top             =   480
      Width           =   255
      _Version        =   786432
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   ">"
      ForeColor       =   16777215
      BackColor       =   16416003
      Appearance      =   3
      DrawFocusRect   =   0   'False
      ImageGap        =   1
   End
   Begin XtremeSuiteControls.PushButton cmdOpenRemoteDir 
      Height          =   255
      Left            =   9840
      TabIndex        =   7
      ToolTipText     =   "��"
      Top             =   480
      Width           =   255
      _Version        =   786432
      _ExtentX        =   450
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   ">"
      ForeColor       =   16777215
      BackColor       =   16416003
      Appearance      =   3
      DrawFocusRect   =   0   'False
      ImageGap        =   1
   End
   Begin XtremeSuiteControls.FlatEdit edLog 
      Height          =   855
      Left            =   120
      TabIndex        =   27
      Top             =   5400
      Width           =   9975
      _Version        =   786432
      _ExtentX        =   17595
      _ExtentY        =   1508
      _StockProps     =   77
      BackColor       =   16777215
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��־"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   5160
      Width           =   420
   End
   Begin VB.Label labSelectedRemote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��0������0������ѡ������0byte"
      Height          =   180
      Left            =   5280
      TabIndex        =   25
      Top             =   4800
      Width           =   2970
   End
   Begin VB.Label labSelectedLocal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��0������0������ѡ������0byte"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   2970
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ¼��"
      Height          =   180
      Index           =   3
      Left            =   5280
      TabIndex        =   3
      Top             =   480
      Width           =   540
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ŀ¼��"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   540
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Զ�̼����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   5280
      TabIndex        =   1
      Top             =   120
      Width           =   1050
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "frmFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sDataRec As String          '���ݷֶ��õĻ��棺��ʱ��������̫��Winsockǿ�Ʒְ�ʱ��Ҫ�õ�����~
Dim bMultiData As Boolean       '�Ƿ�Ϊ������ݷֶ�ģʽ

Public Sub AddLog(strLog As String)             '��־��¼����
    Me.edLog.Text = Me.edLog.Text & Time & " " & strLog & vbCrLf
    Me.edLog.SelStart = Len(Me.edLog.Text)
    Me.Refresh
End Sub

'===================================================================================================
'���ɱ����ļ��б�Ĺ���
Private Sub MakeList(DirPath As String)         '����Ŀ¼�ļ���Ĺ���
    Dim tmp() As String                         '�����ַ���
    Dim FilePath As String                      '�ļ�·���ַ���
    On Error Resume Next
    '=================================
    Me.lstFileLocal.ListItems.Clear
    Me.lstFileLocal.ListItems.Add , , "..."
    Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "�ϲ�Ŀ¼"
    Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = ""
    '=================================
    Me.Dir.Path = DirPath
    Me.Dir.Refresh
    If Err.Number <> 0 Then
        MakeRootList
        Exit Sub
    End If
    For i = 0 To Me.Dir.ListCount - 1
        tmp = Split(Me.Dir.List(i), "\")
        Me.lstFileLocal.ListItems.Add , , tmp(UBound(tmp))
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "Ŀ¼"
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = ""
    Next i
    '=================================
    Me.File.Path = DirPath
    Me.File.Refresh
    FilePath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
    For i = 0 To Me.File.ListCount - 1
        tmp = Split(Me.File.List(i), "\")
        Me.lstFileLocal.ListItems.Add , , tmp(UBound(tmp))
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "�ļ�"
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = GetFileSize(FilePath & Me.File.List(i))
    Next i
    Call RefreshTip
End Sub

Private Sub MakeRootList()                       '���ɴ��̸�Ŀ¼��Ĺ���
    Dim tmpString As String
    Dim lSPC                         'Sectors Per Cluster        ��ÿ�ص���������
    Dim lBPS                         'Bytes Per Sector           ��ÿ�������ֽ�����
    Dim lF                           'Number Of Free Clusters    �����дص�������
    Dim lT                           'Total Number Of Clusters   ���ص�������
    Dim IsExists As Long             '�����Ƿ���Ч
    '===============================
    Me.lstFileLocal.ListItems.Clear
    Me.Drive.Refresh
    For i = 0 To frmMain.Drive.ListCount - 1
        tmpString = frmMain.Drive.List(i)        '�̷�/����
        IsExists = GetDiskFreeSpace(Left(tmpString, 2), lSPC, lBPS, lF, lT)          '��ȡӲ�̿ռ�
        If IsExists <> 0 Then                    '����ܻ�ȡ��������Ϣ˵��������Ч
            Me.lstFileLocal.ListItems.Add , , frmMain.Drive.List(i)
            Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "������"
            Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = "������:" & SizeWithFormat(lSPC * lBPS * lF) & " ��:" & SizeWithFormat(lSPC * lBPS * lT) & "��"
        End If
    Next i
    Call RefreshTip
End Sub

Public Function GetListSelCount(lstTarget As ListView)
    Dim Total As Integer        'ѡ�����Ŀ��
    '=========================================
    For i = 1 To lstTarget.ListItems.Count
        If lstTarget.ListItems(i).Checked = True Then      '���˹��ľͼӽ�ȥ~
            Total = Total + 1
        End If
    Next i
    GetListSelCount = Total
End Function

Public Function RefreshTip()
    On Error Resume Next
    Dim aSize, bSize                '�����б���ѡ�����Ŀ�Ĵ�С���ܺ�
    Dim tmp                         '�����С�Ļ������
    '====================================
    aSize = 0
    bSize = 0
    '---------------------
    For i = 1 To Me.lstFileLocal.ListItems.Count
        If Me.lstFileLocal.ListItems(i).Checked = True Then
            If Me.lstFileLocal.ListItems(i).Text <> "..." Then
                tmp = Me.lstFileLocal.ListItems(i).ListSubItems(2).Text     '��ȡ��С
                If InStr(tmp, "����") = 0 Then                              '������Ǵ��̸�Ŀ¼
                    If tmp = "" Then                                            '�����Ŀ¼
                        tmp = 0
                    End If
                    If InStr(tmp, " Byte") <> 0 Then                            '�����λ�� Byte
                        tmp = Replace(tmp, " Byte", "")                             'ȥ��Byte��λ
                    End If
                    If InStr(tmp, " KB") <> 0 Then                              '�����λ�� KB
                        tmp = Replace(tmp, " KB", "")                               'ȥ��KB��λ
                        tmp = tmp * 1024                                            '���� 1024
                    End If
                    If InStr(tmp, " MB") <> 0 Then                              '�����λ�� MB
                        tmp = Replace(tmp, " MB", "")                               'ȥ��MB��λ
                        tmp = tmp * 1024 ^ 2                                        '���� 1024^2
                    End If
                    If InStr(tmp, " GB") <> 0 Then                              '�����λ�� GB
                        tmp = Replace(tmp, " GB", "")                               'ȥ��GB��λ
                        tmp = tmp * 1024 ^ 3                                        '���� 1024^3
                    End If
                Else
                    tmp = 0
                End If
                aSize = aSize + tmp
            End If
        End If
    Next i
    '==========================
    For i = 1 To Me.lstFileRemote.ListItems.Count
        If Me.lstFileRemote.ListItems(i).Checked = True Then
            If Me.lstFileRemote.ListItems(i).Text <> "..." Then
                tmp = Me.lstFileRemote.ListItems(i).ListSubItems(2).Text    '��ȡ��С
                If InStr(tmp, "����") = 0 Then                              '������Ǵ��̸�Ŀ¼
                    If tmp = "" Then                                            '�����Ŀ¼
                        tmp = 0
                    End If
                    If InStr(tmp, " Byte") <> 0 Then                            '�����λ�� Byte
                        tmp = Replace(tmp, " Byte", "")                             'ȥ��Byte��λ
                    End If
                    If InStr(tmp, " KB") <> 0 Then                              '�����λ�� KB
                        tmp = Replace(tmp, " KB", "")                               'ȥ��KB��λ
                        tmp = tmp * 1024                                            '���� 1024
                    End If
                    If InStr(tmp, " MB") <> 0 Then                              '�����λ�� MB
                        tmp = Replace(tmp, " MB", "")                               'ȥ��MB��λ
                        tmp = tmp * 1024 ^ 2                                        '���� 1024^2
                    End If
                    If InStr(tmp, " GB") <> 0 Then                              '�����λ�� GB
                        tmp = Replace(tmp, " GB", "")                               'ȥ��GB��λ
                        tmp = tmp * 1024 ^ 3                                        '���� 1024^3
                    End If
                Else
                    tmp = 0
                End If
                bSize = bSize + tmp
            End If
        End If
    Next i
    '==========================
    If Me.lstFileLocal.ListItems(1).Text <> "..." Then
        Me.labSelectedLocal.Caption = "��" & Me.lstFileLocal.ListItems.Count & "������" & GetListSelCount(Me.lstFileLocal) & "������ѡ������" & SizeWithFormat(aSize)
    Else
        Me.labSelectedLocal.Caption = "��" & Me.lstFileLocal.ListItems.Count - 1 & "������" & GetListSelCount(Me.lstFileLocal) & "������ѡ������" & SizeWithFormat(aSize)
    End If
    If Me.lstFileRemote.ListItems(1).Text <> "..." Then
        Me.labSelectedRemote.Caption = "��" & Me.lstFileRemote.ListItems.Count & "������" & GetListSelCount(Me.lstFileRemote) & "������ѡ������" & SizeWithFormat(bSize)
    Else
        Me.labSelectedRemote.Caption = "��" & Me.lstFileRemote.ListItems.Count - 1 & "������" & GetListSelCount(Me.lstFileRemote) & "������ѡ������" & SizeWithFormat(bSize)
    End If
End Function

Private Function IsIPExists(strIP As String) As Integer
    For i = 0 To frmMain.comIP.ListCount - 1                'ɨһ���������IP�б�
        If frmMain.comIP.List(i) = strIP Then                   '�ҵ�����ͬ��IP�Ļ�
            IsIPExists = i                                      '����Index
            Exit Function                                       '�˳�����
        End If
    Next i
    IsIPExists = -1                                      '�Ҳ�����ͬ�Ļ��ͷ���-1
End Function

Private Sub ConnectedHandler()                   '���Ӻ�Ĵ������
    Dim iLstIP As Integer                                             '�б��е�IPλ��
    Me.Caption = "�ļ�����: " & Me.wsMessage.RemoteHostIP             '���ı���
    iLstIP = IsIPExists(frmMain.comIP.Text)
    If iLstIP <> -1 Then                                                                '�����⵽���IP���ӹ�
        wsSendData Me.wsMessage, "CNT|" & frmMain.lstPassword.List(iLstIP)                  '���ʹ����������
        Exit Sub
    End If
    If frmMain.lstPassword.List(frmMain.comIP.ListIndex) <> "" Then                         '�����¼��������
        wsSendData Me.wsMessage, "CNT|" & frmMain.lstPassword.List(frmMain.comIP.ListIndex)     '���ʹ����������
    Else
        wsSendData Me.wsMessage, "CNT"                                                          'û����ͷ������ӽ�������
    End If
End Sub

Private Sub cmdBackLocal_Click()
    Dim SplitTmp() As String            '�ָ��
    Dim tmpPath As String               '����·���ַ����Ļ���
    '=============================================================
    If Me.lstFileLocal.ListItems(1).SubItems(1) <> "������" Then
        If Right(Me.Dir.Path, 2) <> ":\" Then               '�����ǰ���ǴӸ�Ŀ¼����һ���Ļ�
            SplitTmp = Split(Me.File.Path, "\")             '��ʾ��һ��Ŀ¼
            For i = 0 To UBound(SplitTmp) - 1
                tmpPath = tmpPath & SplitTmp(i) & "\"
            Next i
            Call MakeList(tmpPath)
        Else                                                '����Ӹ�Ŀ¼����һ��
            Call MakeRootList                               '��ʾ���̸�Ŀ¼
        End If
    Else
        Call MakeRootList                               '��ʾ���̸�Ŀ¼
    End If
    '��ʾ��ǰ��·��
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "������" Then
        Me.edLocalPath.Text = "��Ŀ¼"
    Else
        Me.edLocalPath.Text = Me.Dir.Path
    End If
End Sub

Private Sub cmdBackRemote_Click()
    If Me.lstFileRemote.ListItems(1).SubItems(1) <> "������" Then
        wsSendData Me.wsMessage, "GUD"                            '���ʹ��ϲ�Ŀ¼����
    Else
        wsSendData Me.wsMessage, "GRT"
    End If
End Sub

Private Sub cmdDeleteLocal_Click()
    On Error Resume Next
    Dim tmpPath As String                   '���ɵ�·���Ļ���
    Dim SelectedPath As String              'ѡ���·��
    Dim SplitPath() As String               '�ָ������ÿ��·��������
    '=================================
    If GetListSelCount(Me.lstFileLocal) = 0 Then
        Exit Sub
    End If
    a = MsgBox("�����Ҫɾ��ѡ���" & GetListSelCount(Me.lstFileLocal) & "���ļ����У��𣿸ò������ɳ�����", 32 + vbYesNo, "ȷ��")
    If a <> 6 Then
        Exit Sub
    End If
    '----------------------------------------
    AddLog "�����ء���ʼɾ��ѡ���" & GetListSelCount(Me.lstFileLocal) & "���ļ����У���"
    For i = 1 To Me.lstFileLocal.ListItems.Count                   '��ȡ����ѡ����б���
        If Me.lstFileLocal.ListItems(i).Checked = True Then
            SelectedPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
            SelectedPath = SelectedPath & Me.lstFileLocal.ListItems(i)
            tmpPath = tmpPath & SelectedPath & "|" & Me.lstFileLocal.ListItems(i).SubItems(1) & vbCrLf
        End If
    Next i
    SplitPath = Split(tmpPath, vbCrLf)      '�ָ�
    For i = 0 To UBound(SplitPath)          '�õ�����ѡ����ļ�
        If Split(SplitPath(i), "|")(1) = "�ļ�" Then               '������ļ�����
            If bDeleteFiles Then
                Kill Split(SplitPath(i), "|")(0)
            Else
                MsgBox Split(SplitPath(i), "|")(0)
            End If
        End If
        If Split(SplitPath(i), "|")(1) = "Ŀ¼" Then               '�����Ŀ¼����
            If bDeleteFiles Then
                Shell "cmd /c rd " & Chr(34) & Split(SplitPath(i), "|")(0) & Chr(34) & " /s /q", vbHide
            Else
                MsgBox "cmd /c rd " & Chr(34) & Split(SplitPath(i), "|")(0) & Chr(34) & " /s /q"
            End If
        End If
    Next i
    Call cmdRefreshLocal_Click
End Sub

Private Sub cmdDeleteRemote_Click()
    Dim tmpSendString As String                     '�������ݻ���
    '----------------------------------
    If GetListSelCount(Me.lstFileRemote) = 0 Then
        Exit Sub
    End If
    a = MsgBox("�����Ҫɾ��ѡ���" & GetListSelCount(Me.lstFileRemote) & "���ļ����У��𣿸ò������ɳ�����", 32 + vbYesNo, "ȷ��")
    If a <> 6 Then
        Exit Sub
    End If
    AddLog "��Զ�̡���ʼɾ��ѡ���" & GetListSelCount(Me.lstFileRemote) & "���ļ����У���"
    tmpSendString = "DEL|"
    For i = 1 To Me.lstFileRemote.ListItems.Count
        If Me.lstFileRemote.ListItems(i).Checked = True Then      '������������Ӵ��˹����б���
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '����ɾ���ļ�����
End Sub

Private Sub cmdDownload_Click()
    Dim tmpSend As String                           '��Ҫ���ص��ļ�����б�
    Dim Reminded As Boolean                         '�Ƿ����ѹ��ļ��з�����
    '============================================================================
    If IsWorking Then                               '�����ǰ���������ڽ���
        AddLog "��ǰ�Ѿ������ڽ����е�������ȴ���������ɺ��ټ���������"
        Exit Sub
    End If
    Reminded = False                                'δ���ѹ�
    If GetListSelCount(Me.lstFileRemote) = 0 Then                            '���δѡ���ļ�
        MsgBox "δѡ���ļ���", 48, "��ʾ"                                       '��ʾ��ʾ
        Exit Sub                                                                '�˳�����
    End If
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "������" Then       '���Ŀ���Ǹ�Ŀ¼
        MsgBox "�޷����ص���Ŀ¼��", 48, "��ʾ"                                 '��ʾ��ʾ
        Exit Sub                                                                '�˳�����
    End If
    '============================================================================
    For i = 1 To Me.lstFileRemote.ListItems.Count                           '��ȡ����ѡ����б���
        If Me.lstFileRemote.ListItems(i).Checked = True Then                     '����Ŀ¼�ͼ����б�
            If Me.lstFileRemote.ListItems(i).SubItems(1) <> "Ŀ¼" Then
                tmpSend = tmpSend & CStr(i) & vbCrLf
            Else
                If Reminded = False Then                                    '���δ���ѹ�����ʾĿ¼���ܷ���
                    MsgBox "��ѡ��Ķ���������ļ��У��ļ��в�֧�ַ��ͣ�����ʱ���Զ����ԡ�", 48, "��ʾ"
                    Reminded = True                 '���ѹ���
                End If
            End If
        End If
    Next i
    If tmpSend = "" Then                            'û��һ�����ļ���
        MsgBox "û���ļ������͡�", 48, "��ʾ"
        Exit Sub
    End If
    '==============================================
    frmDownload.IsUpload = False                                        '����״̬
    frmDownload.DownloadPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\") '��������·��
    frmDownload.wsFile.Bind 20236
    frmDownload.wsFile.Listen                                           '��ʼ����
    IsWorking = True
    frmDownload.Show
    wsSendData Me.wsMessage, "DWL|" & tmpSend                           '���������ļ�����
End Sub

Private Sub cmdHomeLocal_Click()
    Call MakeRootList                               '���ɸ�Ŀ¼
End Sub

Private Sub cmdHomeRemote_Click()
    wsSendData Me.wsMessage, "GRT"                  '��ȡ��Ŀ¼�б�����
End Sub

Private Sub cmdMkdirLocal_Click()
    Dim fName As String                             '�ļ������ֻ���
    On Error Resume Next
    If Me.lstFileLocal.ListItems(1).SubItems(1) <> "������" Then      '����Ǹ�Ŀ¼�б�
        fName = InputBox("���������ļ��е����֣�", "����")
        If Trim(fName) = "" Then
            Exit Sub
        End If
        fName = IIf(Right(Me.Dir.Path, 1) = "\", Me.Dir.Path, Me.Dir.Path & "\") & fName
        MkDir fName
        If Err.Number <> 0 Then
            AddLog "�����ء������ļ���ʧ�ܡ�"
        End If
    Else
        AddLog "�����ء������ڸ�Ŀ¼�ﴴ���ļ��С�"
    End If
End Sub

Private Sub cmdMkdirRemote_Click()
    Dim fName As String                             '�ļ������ֻ���
    If Me.lstFileRemote.ListItems(1).SubItems(1) <> "������" Then      '����Ǹ�Ŀ¼�б�
        fName = InputBox("���������ļ��е����֣�", "����")
        If Trim(fName) <> "" Then
            If InStr(fName, "|") = 0 Then
                wsSendData Me.wsMessage, "MKD|" & fName         '���ʹ����ļ�������
            Else
                AddLog "��Զ�̡������ļ���ʧ�ܡ�"
            End If
        End If
    Else
        AddLog "��Զ�̡������ڸ�Ŀ¼�ﴴ���ļ��С�"
    End If
End Sub

Private Sub cmdOpenLocalDir_Click()
    Dim vPath As String                         'Ҫ�����Ŀ¼
    '===========================================================
    If Me.edLocalPath.Text = "��Ŀ¼" Then
        Call MakeRootList
    Else
        vPath = Me.edLocalPath.Text
        If IsPathExists(vPath) = False Then                 '��⵽�����Ŀ¼��Ч
            AddLog "�����ء���Ч��Ŀ¼�������ļ���" & vPath & "ʧ�ܡ�"
        Else
            Call MakeList(vPath)
        End If
    End If
End Sub

Private Sub cmdOpenRemoteDir_Click()
    wsSendData Me.wsMessage, "VPH|" & Me.edRemotePath.Text    '���ʹ�Ŀ¼����
End Sub

Public Sub cmdRefreshLocal_Click()
    'ˢ����~
    Me.Dir.Refresh
    Me.File.Refresh
    Me.Drive.Refresh
    '------------------------
    If Me.lstFileLocal.ListItems(1).SubItems(1) = "������" Then      '����Ǹ�Ŀ¼�б�
        Call MakeRootList                                           '���ɴ��̸�Ŀ¼�б�
    Else                                                            '�������ͨĿ¼������Ŀ¼�б�
        Call MakeList(Me.Dir.Path)
    End If
End Sub

Private Sub cmdRefreshRemote_Click()
    wsSendData Me.wsMessage, "REF"                  '����ˢ������
End Sub

Private Sub cmdUpload_Click()
    Dim tmpPath As String                           '��Ҫ�ϴ����ļ�
    Dim SelectedPath As String                      'ÿ���ļ���Ŀ¼
    Dim Reminded As Boolean                         '�Ƿ����ѹ��ļ��з�����
    '============================================================================
    If IsWorking Then                               '�����ǰ���������ڽ���
        AddLog "��ǰ�Ѿ������ڽ����е�������ȴ���������ɺ��ټ���������"
        Exit Sub
    End If
    Reminded = False                                'δ���ѹ�
    If GetListSelCount(Me.lstFileLocal) = 0 Then                            '���δѡ���ļ�
        MsgBox "δѡ���ļ���", 48, "��ʾ"                                       '��ʾ��ʾ
        Exit Sub                                                                '�˳�����
    End If
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "������" Then       '���Ŀ���Ǹ�Ŀ¼
        MsgBox "�޷����ص���Ŀ¼��", 48, "��ʾ"                                 '��ʾ��ʾ
        Exit Sub                                                                '�˳�����
    End If
    '============================================================================
    For i = 1 To Me.lstFileLocal.ListItems.Count                           '��ȡ����ѡ����б���
        If Me.lstFileLocal.ListItems(i).Checked = True Then                     '����Ŀ¼�ͼ����б�
            If Me.lstFileLocal.ListItems(i).SubItems(1) <> "Ŀ¼" Then
                SelectedPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                SelectedPath = SelectedPath & Me.lstFileLocal.ListItems(i)
                tmpPath = tmpPath & SelectedPath & vbCrLf
            Else
                If Reminded = False Then                                    '���δ���ѹ�����ʾĿ¼���ܷ���
                    MsgBox "��ѡ��Ķ���������ļ��У��ļ��в�֧�ַ��ͣ�����ʱ���Զ����ԡ�", 48, "��ʾ"
                    Reminded = True                 '���ѹ���
                End If
            End If
        End If
    Next i
    If tmpPath = "" Then                            'û��һ�����ļ���
        MsgBox "û���ļ������͡�", 48, "��ʾ"
        Exit Sub
    End If
    '==============================================
    frmDownload.FileList = tmpPath                                      '�����ļ��б�
    frmDownload.IsUpload = True                                         '�ϴ�״̬
    frmDownload.wsFile.Bind 20236
    frmDownload.wsFile.Listen                                           '��ʼ����
    '����Ϊ״̬��ʼ��
    frmDownload.CurrentFileIndex = -1                                   '��ʼ���ļ����
    Call frmDownload.SplitPath                                          '�����ÿ���ļ�
    If frmDownload.NextFile = False Then                                '��һ���ļ�
        Exit Sub                                                        '�����ȡʧ�ܾ�ֱ���˳�����
    End If
    IsWorking = True
    frmDownload.Show
    Call frmDownload.SendUploadMsg                                      '�����ϴ��ļ�����
End Sub

Private Sub edLocalPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOpenLocalDir_Click                   '���ð�ť���¹���
    End If
End Sub

Private Sub edRemotePath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOpenRemoteDir_Click                   '���ð�ť���¹���
    End If
End Sub

Private Sub Form_Load()
    '��¼�Է�IP���б���
    If bAutoRecord Then                 '������Զ���¼IP���㿩 �r(�s_�t)�q
        If frmRemoteControl.IsIPExists(Me.wsMessage.RemoteHostIP) = -1 Then       '���IP�����ھͼ�¼������
            frmMain.comIP.AddItem Me.wsMessage.RemoteHostIP
            frmMain.lstPassword.AddItem frmEnterPassword.edPassword.Text
            Call SaveConfig                         '���������ļ�
        End If
    End If
    '����б�ͷ
    Me.lstFileLocal.ColumnHeaders.Add , , "����", 2000
    Me.lstFileLocal.ColumnHeaders.Add , , "����", 800
    Me.lstFileLocal.ColumnHeaders.Add , , "��С", 3000
    Me.lstFileRemote.ColumnHeaders.Add , , "����", 2000
    Me.lstFileRemote.ColumnHeaders.Add , , "����", 800
    Me.lstFileRemote.ColumnHeaders.Add , , "��С", 3000
    '��ȡ��Ŀ¼�б�
    Dim sRootMsg() As String                        '���б������ϸ��Ϣ�и��
    For i = 0 To frmMain.lstTemp.ListCount - 1
        sRootMsg = Split(frmMain.lstTemp.List(i), "|")
        Me.lstFileLocal.ListItems.Add , , sRootMsg(0)
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "������"
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = sRootMsg(1)
    Next i
    Call MakeRootList
    IsWorking = False                               'û���������ڽ���
    AddLog "�ļ�������������"
End Sub

Private Sub Form_Resize()
    Dim lMid As Single
    On Error Resume Next
    '-------------------------
    If Me.Width < 7350 Then
        Me.Width = 7350
    End If
    If Me.Height < 4550 Then
        Me.Height = 4550
    End If
    lMid = Me.Width / 2             '��������λ��
    '-------------------------
    '�±�
    Me.edLog.Width = Me.Width - Me.edLog.Left - 360
    Me.edLog.Top = Me.Height - Me.edLog.Height * 2 + 120
    Me.labTip(4).Top = Me.edLog.Top - Me.labTip(4).Height - 120
    '--------------------------------------------------------------------------------
    '���
    Me.edLocalPath.Width = lMid - 240 - Me.cmdOpenLocalDir.Width - Me.edLocalPath.Left
    Me.cmdOpenLocalDir.Left = Me.edLocalPath.Left + Me.edLocalPath.Width
    Me.labTip(1).Left = lMid
    Me.picLocalToolbar.Width = lMid - Me.picLocalToolbar.Left - 240
    Me.cmdUpload.Left = Me.picLocalToolbar.Width - Me.cmdUpload.Width
    Me.lstFileLocal.Width = lMid - Me.lstFileLocal.Left - 240
    Me.labSelectedLocal.Top = Me.labTip(4).Top - Me.labSelectedLocal.Height - 120
    Me.lstFileLocal.Height = Me.labSelectedLocal.Top - Me.lstFileLocal.Top - 120
    '--------------------------------------------------------------------------------
    '�ұ�
    Me.labTip(3).Left = lMid
    Me.edRemotePath.Left = Me.labTip(3).Left + Me.labTip(3).Width
    Me.edRemotePath.Width = Me.Width - Me.edRemotePath.Left - 360 - Me.cmdOpenRemoteDir.Width
    Me.cmdOpenRemoteDir.Left = Me.edRemotePath.Left + Me.edRemotePath.Width
    Me.lstFileRemote.Left = lMid
    Me.lstFileRemote.Width = Me.Width - Me.lstFileRemote.Left - 360
    Me.lstFileRemote.Height = Me.lstFileLocal.Height
    Me.picRemoteToolbar.Left = lMid
    Me.picRemoteToolbar.Width = Me.Width - Me.picRemoteToolbar.Left - 360
    Me.cmdRefreshRemote.Left = Me.picRemoteToolbar.Width - Me.cmdRefreshRemote.Width
    Me.labSelectedRemote.Top = Me.labSelectedLocal.Top
    Me.labSelectedRemote.Left = lMid
    Me.cmdDeleteRemote.Left = Me.cmdRefreshRemote.Left - 500
    Me.cmdMkdirRemote.Left = Me.cmdDeleteRemote.Left - 500
    Me.cmdBackRemote.Left = Me.cmdMkdirRemote.Left - 500
    Me.cmdHomeRemote.Left = Me.cmdBackRemote.Left - 500
End Sub

Private Sub lstFileLocal_DblClick()
    '��ȡ���Ͳ������ͷ�������
    On Error Resume Next
    Dim SplitTmp() As String            '�ָ��
    Dim tmpPath As String               '����·���ַ����Ļ���
    '===============================================================
    Select Case Me.lstFileLocal.ListItems.Item(Me.lstFileLocal.SelectedItem.Index).SubItems(1)
        Case "������"
            tmpPath = Left(Me.lstFileLocal.ListItems.Item(Me.lstFileLocal.SelectedItem.Index), 2) & "\"
            Call MakeList(tmpPath)
            
        Case "Ŀ¼"
            tmpPath = Me.Dir.List(Me.lstFileLocal.SelectedItem.Index - 2)
            Call MakeList(tmpPath)
            
        Case "�ϲ�Ŀ¼"
            If Right(Me.Dir.Path, 2) <> ":\" Then               '�����ǰ���ǴӸ�Ŀ¼����һ���Ļ�
                SplitTmp = Split(Me.File.Path, "\")             '��ʾ��һ��Ŀ¼
                For i = 0 To UBound(SplitTmp) - 1
                    tmpPath = tmpPath & SplitTmp(i) & "\"
                Next i
                Call MakeList(tmpPath)
            Else                                                '����Ӹ�Ŀ¼����һ��
                Call MakeRootList                               '��ʾ���̸�Ŀ¼
            End If
            
    End Select
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "������" Then
        Me.edLocalPath.Text = "��Ŀ¼"
    Else
        Me.edLocalPath.Text = Me.Dir.Path
    End If
End Sub

Private Sub lstFileLocal_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
    On Error Resume Next
    For i = 1 To Me.lstFileLocal.ListItems.Count
        '����Ǵ��� ���� �ϲ�Ŀ¼ѡ�� ���� Ŀ¼
        If InStr(Me.lstFileLocal.ListItems(i).SubItems(2), "����:") <> 0 Or Me.lstFileLocal.ListItems(i).SubItems(1) = "�ϲ�Ŀ¼" Then
            Me.lstFileLocal.ListItems(i).Checked = False
        End If
        Call RefreshTip
    Next i
End Sub

Private Sub lstFileLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then        '�����˸��
        cmdBackLocal_Click                  '�����ϲ��ļ���
    End If
    If KeyAscii = 13 Then               '���»س���
        Call lstFileLocal_DblClick          '��ѡ����ļ�
        KeyAscii = 0
    End If
End Sub

Private Sub lstFileRemote_DblClick()
    '��ȡ���Ͳ������ͷ�������
    On Error Resume Next
    Select Case Me.lstFileRemote.ListItems.Item(Me.lstFileRemote.SelectedItem.Index).SubItems(1)
        Case "������"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFileRemote.SelectedItem.Index
            
        Case "Ŀ¼"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFileRemote.SelectedItem.Index
            
        Case "�ϲ�Ŀ¼"
            wsSendData Me.wsMessage, "GUD"
            
    End Select
End Sub

Private Sub lstFileRemote_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
    On Error Resume Next
    For i = 1 To Me.lstFileRemote.ListItems.Count
        '����Ǵ��� ���� �ϲ�Ŀ¼ѡ��
        If InStr(Me.lstFileRemote.ListItems(i).SubItems(2), "����:") <> 0 Or Me.lstFileRemote.ListItems(i).SubItems(1) = "�ϲ�Ŀ¼" Then
            Me.lstFileRemote.ListItems(i).Checked = False
        End If
        Call RefreshTip
    Next i
End Sub

Private Sub lstFileRemote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then        '�����˸��
        cmdBackRemote_Click                 '�����ϲ��ļ���
    End If
    If KeyAscii = 13 Then               '���»س���
        Call lstFileRemote_DblClick          '��ѡ����ļ�
        KeyAscii = 0
    End If
End Sub

Private Sub wsMessage_Connect()
    '���ӳɹ�
    Call ConnectedHandler       '���Ӵ���
End Sub

Private Sub wsMessage_DataArrival(ByVal bytesTotal As Long)
    Dim Temp As String                      '��������
    Dim sTemp() As String                   '�����и������
    Dim sTempPara() As String               '���ݸ����Ĳ����и�����Ļ�������
    Dim TaskSplit() As String               'ÿһ��������Ϣ�ķָ�棨���ڽ��̷���ģ�
    '====================================
    '��������
    Me.wsMessage.GetData Temp               '��ȡ����
    Temp = Decrypt(Temp)                            '��������
    If InStr(Temp, "{END}") = 0 Then                '�Ҳ����������˵�����Ƿְ����Ȳ�������Ϣ
        sDataRec = sDataRec & Temp
        bMultiData = True
        Exit Sub
    Else
        If bMultiData = True Then                           '����Ƕ�����ݰ�����ƴ����������
            sDataRec = Replace(sDataRec, "{END}", "")       '����ҵ�������Ǿ�˵�����������ݣ���ʼ������Ϣ ���ѽ������ɾ����
        Else
            sDataRec = Replace(Temp, "{END}", "")           '����ǵ������ݰ����õ������ܵ�������
        End If
    End If
    sTemp = Split(sDataRec, "{S}")              '�и�����
    For i = 0 To UBound(sTemp)
        Select Case Left(sTemp(i), 3)       '��������
            Case "NAC"                                      '�ܾ���������
                Me.wsMessage.Close
                frmMain.labState.Caption = "�Է�������Զ�̿���"
                frmMain.shpState.FillColor = vbRed
                frmMain.tmrReturn.Enabled = True
        
            Case "CNT"                                      '����������
                frmMain.Hide
                sTempPara = Split(sTemp(i), "|")
                Select Case sTempPara(1)                    '�������۷�������
                    Case "True"                                 '��������
                        '���п��Ƴ�ʼ��
                        '��¼�Է�IP���б���
                        If bAutoRecord Then                 '������Զ���¼IP���㿩 �r(�s_�t)�q
                            If IsIPExists(Me.wsMessage.RemoteHostIP) = -1 Then               '���IP�����ھͼ�¼������
                                frmMain.comIP.AddItem Me.wsMessage.RemoteHostIP
                                frmMain.lstPassword.AddItem frmEnterPassword.edPassword.Text
                                Call SaveConfig                         '���������ļ�
                            End If
                        End If
                        '===================================================
                        '����/��ʾ ����
                        frmEnterPassword.Hide
                        Me.Show
                        '===================================================
                        '������ϵͳ��Ϣ
                        wsSendData Me.wsMessage, "GRT"            '���ͻ�ȡ��Ŀ¼����
                        
                    Case "Wrong"                                '��������µľܾ�����
                        Dim iIpList As Integer                          '���IP�Ƿ���ڵĻ���
                        If frmEnterPassword.Visible = False Then
                            frmEnterPassword.Show
                        End If
                        iIpList = IsIPExists(Me.wsMessage.RemoteHostIP)
                        If iIpList <> -1 Then                           '�����⵽IP����
                            frmMain.comIP.RemoveItem iIpList                'ɾ�������������б���
                            frmMain.lstPassword.RemoveItem iIpList
                            Call SaveConfig                                 '���������ļ�
                        End If
                        frmEnterPassword.labIncorrect.Visible = True
                        frmEnterPassword.Height = 2400
                        frmEnterPassword.edPassword.SelStart = 0
                        frmEnterPassword.edPassword.SelLength = Len(frmEnterPassword.edPassword.Text)
                        frmEnterPassword.edPassword.SetFocus
                    
                    Case "False"                                '�����뵼�µľܾ�����
                        frmEnterPassword.Show
                        frmEnterPassword.Height = 2000
                        
                End Select
                        
            Case "GRT"                                      '��ȡ��Ŀ¼������Ӧ
                Dim sRootList() As String                       '�б����и��
                Dim sRootMsg() As String                        '���б������ϸ��Ϣ�и��
                '-------------------------------------
                sRootList() = Split(Replace(sDataRec, "GRT|", ""), "||")    '�и��ַ���
                Me.lstFileRemote.ListItems.Clear
                For k = 0 To UBound(sRootList) - 1
                    sRootMsg = Split(sRootList(k), "|")
                    Me.lstFileRemote.ListItems.Add , , sRootMsg(0)
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(1) = "������"
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(2) = sRootMsg(1)
                Next k
                Me.edRemotePath.Text = "��Ŀ¼"
                Call RefreshTip
            
            Case "GPH"                                      '��ȡĿ¼���ļ�������Ӧ
                Dim sDirList() As String                        '�б����и��
                Dim sDirMsg() As String                         '���б������ϸ��Ϣ�и��
                On Error Resume Next
                '-------------------------------------
                sDirList = Split(Replace(sDataRec, "GPH|", ""), "||")       '�и��ַ���
                Me.lstFileRemote.ListItems.Clear
                For k = 0 To UBound(sDirList) - 1
                    sDirMsg = Split(sDirList(k), "|")
                    Me.lstFileRemote.ListItems.Add , , sDirMsg(0)
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(1) = sDirMsg(1)
                    sDirMsg(2) = Replace(sDirMsg(2), ":N:", "")         '����Ƿ�ֹ�ļ��л�ȡ���˴�С���µĿհ����������
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(2) = sDirMsg(2)
                Next k
                Call RefreshTip
                wsSendData Me.wsMessage, "NPH"                  '���ͻ�ȡ��ǰĿ¼·��������
            
            Case "NPH"
                Me.edRemotePath.Text = Replace(Replace(sDataRec, "NPH|", ""), "{S}", "")      '��ʾ���ص��ļ���·��
            
            Case "VPH"
                AddLog "��Զ�̡���Ч��Ŀ¼�������ļ���" & Me.edRemotePath.Text & "ʧ�ܡ�"      '��ʾ������Ϣ
                wsSendData Me.wsMessage, "NPH"                  '���ͻ�ȡ��ǰĿ¼·��������
            
            Case "MKD"
                Dim bSuccess As String
                bSuccess = Replace(Replace(sDataRec, "MKD|", ""), "{S}", "")
                If bSuccess = "0" Then
                    AddLog "��Զ�̡������ļ���ʧ�ܡ�"
                Else
                    AddLog "��Զ�̡������ļ��гɹ���"
                End If
                wsSendData Me.wsMessage, "REF"                  '����ˢ������
            
        End Select
    Next i
    '-----------------------------
    sDataRec = ""                       'ִ��������������������
    bMultiData = False
End Sub
