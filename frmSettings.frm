VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����"
   ClientHeight    =   6564
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   7320
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6564
   ScaleWidth      =   7320
   StartUpPosition =   2  '��Ļ����
   Begin XtremeSuiteControls.GroupBox fraSettings1 
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   7095
      _Version        =   786432
      _ExtentX        =   12515
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "�������"
      Appearance      =   6
      BorderStyle     =   1
      Begin XtremeSuiteControls.RadioButton optExit 
         Height          =   255
         Left            =   2160
         TabIndex        =   4
         Top             =   1800
         Width           =   2200
         _Version        =   786432
         _ExtentX        =   3881
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "�˳����"
         Appearance      =   6
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkAutoStart 
         Height          =   255
         Left            =   0
         TabIndex        =   0
         Top             =   360
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "����ʱ�Զ�����"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkTraytWhenStart 
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "����ʱ��С����������"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkHideWhenStart 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "�����Զ��������Ժ�̨ģʽ���� ������ʾ����ͼ�꣩"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkAvoidExit 
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   1440
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "����������˳�"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton optTray 
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         Top             =   1800
         Width           =   2205
         _Version        =   786432
         _ExtentX        =   3881
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "��С����������"
         Appearance      =   6
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������尴ť����ϣ��"
         Height          =   180
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   1800
         Width           =   1980
      End
   End
   Begin XtremeSuiteControls.GroupBox fraSettings2 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   7095
      _Version        =   786432
      _ExtentX        =   12515
      _ExtentY        =   1931
      _StockProps     =   79
      Caption         =   "Զ�̿�������"
      Appearance      =   6
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkHideMode 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   360
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "����ģʽ (û���κ���ʾ����Ϣ)"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkAutoResize 
         Height          =   255
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "�Զ�����Զ����Ļ����"
         Appearance      =   6
      End
   End
   Begin XtremeSuiteControls.GroupBox fraSettings3 
      Height          =   2895
      Left            =   120
      TabIndex        =   18
      Top             =   3600
      Width           =   7095
      _Version        =   786432
      _ExtentX        =   12515
      _ExtentY        =   5106
      _StockProps     =   79
      Caption         =   "���뼰��¼����"
      Appearance      =   6
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkUserPassword 
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   720
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ʹ�ø�������(ͨ�������������Ҳ�ܷ��ʴ˵���)"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ListBox lstIP 
         Height          =   975
         Left            =   0
         TabIndex        =   13
         Top             =   1845
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   1720
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   4
      End
      Begin XtremeSuiteControls.CheckBox chkAutoRecord 
         CausesValidation=   0   'False
         Height          =   255
         Left            =   0
         TabIndex        =   8
         Top             =   360
         Width           =   6975
         _Version        =   786432
         _ExtentX        =   12303
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "�Զ���¼�ɹ����ӵ�IP������"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit edPassword1 
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   77
         Enabled         =   0   'False
         BackColor       =   16777215
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit edPassword2 
         Height          =   255
         Left            =   4080
         TabIndex        =   11
         Top             =   1080
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   77
         Enabled         =   0   'False
         BackColor       =   16777215
         PasswordChar    =   "*"
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdDeleteSelected 
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   1845
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ɾ��ѡ��"
         ForeColor       =   16777215
         BackColor       =   16416003
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin XtremeSuiteControls.PushButton cmdDeleteAll 
         Height          =   375
         Left            =   3240
         TabIndex        =   15
         Top             =   2355
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ɾ��ȫ��"
         ForeColor       =   16777215
         BackColor       =   16416003
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin XtremeSuiteControls.PushButton cmdChangePassword 
         Height          =   300
         Left            =   6120
         TabIndex        =   12
         Top             =   1080
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "����"
         ForeColor       =   16777215
         BackColor       =   16416003
         Enabled         =   0   'False
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��¼ (���ӳɹ�����IP����ʾ�����������ɾ��������������)��"
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   21
         Top             =   1560
         Width           =   5310
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȷ�����룺"
         Enabled         =   0   'False
         Height          =   180
         Index           =   2
         Left            =   3120
         TabIndex        =   20
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���룺"
         Enabled         =   0   'False
         Height          =   180
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   540
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkAutoRecord_Click()
    bAutoRecord = Me.chkAutoRecord.Value            '�л��Ƿ��Զ���¼IP
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkAutoResize_Click()
    AutoResize = Me.chkAutoResize.Value             '�л��Ƿ��Զ�����ͼ��
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkAutoStart_Click()
    On Error Resume Next
    Dim ws
    Dim ProgramPath As String                       '��ǰ����·��
    '=========================================
    bAutoStart = Me.chkAutoStart.Value              '�л��Ƿ񿪻�����
    Set ws = CreateObject("wscript.shell")
    '=========================================
    If bAutoStart = True Then                       '�������������д�뿪��������
        '��ȡ��������·��
        ProgramPath = App.Path
        ProgramPath = IIf(Right(ProgramPath, 1) = "\", ProgramPath & App.EXEName & ".exe", ProgramPath & "\" & App.EXEName & ".exe /AutoStart")
        ProgramPath = Chr(34) & ProgramPath & Chr(34)
        '����д��ע���
        ws.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteControl", ProgramPath
        If Err.Number = 0 Then                      'д��ע���ɹ�
            Me.chkAutoStart.Value = 1
            bAutoStart = True
        Else                                        'д��ע���ʧ��
            MsgBox "д�뿪��������ʧ�ܣ�", 48, "����"       '��ʾ������ʾ
            Me.chkAutoStart.Value = 0
            bAutoStart = False
        End If
    Else
        '����ɾ��ע�����Ŀ���������Ŀ
        ws.RegDelete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteControl"
    End If
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkAvoidExit_Click()
    bNoExit = Me.chkAvoidExit.Value                 '�л�����ģʽ
    frmMain.mnuExit.Enabled = Not bNoExit           '���Ĳ˵�����״̬
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkHideMode_Click()
    bHideMode = Me.chkHideMode.Value                '�л�����ģʽ
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkHideWhenStart_Click()
    bHideWhenStart = Me.chkHideWhenStart.Value      '�л��Ƿ񿪻��Զ��������Ժ�̨ģʽ����
    If bHideWhenStart = False Then
        frmMain.Tray.Icon = frmMain.Icon
    End If
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkTraytWhenStart_Click()
    bTrayWhenStart = Me.chkTraytWhenStart.Value     '�л��Ƿ�����ʱ��С��
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub chkUserPassword_Click()
    On Error Resume Next
    Me.labTip(1).Enabled = Me.chkUserPassword.Value     '���Ŀؼ�״̬
    Me.labTip(2).Enabled = Me.chkUserPassword.Value
    Me.edPassword1.Enabled = Me.chkUserPassword.Value
    Me.edPassword2.Enabled = Me.chkUserPassword.Value
    Me.cmdChangePassword.Enabled = Me.chkUserPassword.Value
    bUseUserPassword = Me.chkUserPassword.Value         '�л��Ƿ�ʹ���û�����
    If Me.edPassword1.Enabled = True Then               '����ı�����þ��Զ������ı���
        Me.edPassword1.SelStart = 0
        Me.edPassword1.SelLength = Len(Me.edPassword1.Text)
        Me.edPassword1.SetFocus
    End If
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub cmdChangePassword_Click()
    If Me.edPassword1.Text = Me.edPassword2.Text Then               '��������������
        sUserPassword = Me.edPassword1.Text                             '��������
        Call SaveConfig                                                 '���������ļ�
        MsgBox "�����޸ĳɹ���", 64, "��ʾ"
    Else
        MsgBox "������������벻����", 48, "��ʾ"
        Me.edPassword1.SetFocus
    End If
End Sub

Private Sub cmdDeleteAll_Click()
    Me.lstIP.Clear                      'ȫ������~~
    frmMain.comIP.Clear
    frmMain.lstPassword.Clear
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub cmdDeleteSelected_Click()
    On Error Resume Next                '��ֹɾ�����б������ ����~���ô����ˣ�ֱ���Թ�~ (*^__^*)
    frmMain.comIP.RemoveItem Me.lstIP.ListIndex                     'ɾ����¼��IP������
    frmMain.lstPassword.RemoveItem Me.lstIP.ListIndex
    Me.lstIP.RemoveItem Me.lstIP.ListIndex
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub edPassword1_GotFocus()
    If Me.edPassword1.Text <> "" Then
        Me.edPassword1.SelStart = 0
        Me.edPassword1.SelLength = Len(Me.edPassword1.Text)
    End If
End Sub

Private Sub edPassword1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.edPassword2.SetFocus
    End If
End Sub

Private Sub edPassword2_GotFocus()
    If Me.edPassword2.Text <> "" Then
        Me.edPassword2.SelStart = 0
        Me.edPassword2.SelLength = Len(Me.edPassword2.Text)
    End If
End Sub

Private Sub edPassword2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdChangePassword_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub optExit_Click()
    bExitWhenClose = Me.optExit.Value               '�ر�ʱ�˳����
    Call SaveConfig                                 '���������ļ�
End Sub

Private Sub optTray_Click()
    bExitWhenClose = Me.optExit.Value               '�ر�ʱ�����С����������
    Call SaveConfig                                 '���������ļ�
End Sub
