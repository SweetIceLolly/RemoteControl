VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Begin VB.Form frmSettings 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "设置"
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
   StartUpPosition =   2  '屏幕中心
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
      Caption         =   "软件设置"
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
         Caption         =   "退出软件"
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
         Caption         =   "开机时自动启动"
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
         Caption         =   "启动时最小化到托盘区"
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
         Caption         =   "开机自动启动后以后台模式运行 （不显示托盘图标）"
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
         Caption         =   "软件不允许退出"
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
         Caption         =   "最小化到托盘区"
         Appearance      =   6
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "点击窗体按钮后，我希望"
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
      Caption         =   "远程控制设置"
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
         Caption         =   "安静模式 (没有任何提示和消息)"
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
         Caption         =   "自动拉伸远程屏幕画面"
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
      Caption         =   "密码及记录设置"
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
         Caption         =   "使用个人密码(通过输入这个密码也能访问此电脑)"
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
         Caption         =   "自动记录成功连接的IP和密码"
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
         Caption         =   "删除选定"
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
         Caption         =   "删除全部"
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
         Caption         =   "更改"
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
         Caption         =   "记录 (连接成功过的IP会显示在这里，您可以删除掉以忘记密码)："
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
         Caption         =   "确认密码："
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
         Caption         =   "密码："
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
    bAutoRecord = Me.chkAutoRecord.Value            '切换是否自动记录IP
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkAutoResize_Click()
    AutoResize = Me.chkAutoResize.Value             '切换是否自动拉伸图像
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkAutoStart_Click()
    On Error Resume Next
    Dim ws
    Dim ProgramPath As String                       '当前程序路径
    '=========================================
    bAutoStart = Me.chkAutoStart.Value              '切换是否开机启动
    Set ws = CreateObject("wscript.shell")
    '=========================================
    If bAutoStart = True Then                       '如果开机启动就写入开机启动项
        '获取程序自身路径
        ProgramPath = App.Path
        ProgramPath = IIf(Right(ProgramPath, 1) = "\", ProgramPath & App.EXEName & ".exe", ProgramPath & "\" & App.EXEName & ".exe /AutoStart")
        ProgramPath = Chr(34) & ProgramPath & Chr(34)
        '尝试写入注册表
        ws.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteControl", ProgramPath
        If Err.Number = 0 Then                      '写入注册表成功
            Me.chkAutoStart.Value = 1
            bAutoStart = True
        Else                                        '写入注册表失败
            MsgBox "写入开机启动项失败！", 48, "错误"       '显示错误提示
            Me.chkAutoStart.Value = 0
            bAutoStart = False
        End If
    Else
        '尝试删除注册表里的开机启动项目
        ws.RegDelete "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Run\RemoteControl"
    End If
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkAvoidExit_Click()
    bNoExit = Me.chkAvoidExit.Value                 '切换安静模式
    frmMain.mnuExit.Enabled = Not bNoExit           '更改菜单栏的状态
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkHideMode_Click()
    bHideMode = Me.chkHideMode.Value                '切换安静模式
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkHideWhenStart_Click()
    bHideWhenStart = Me.chkHideWhenStart.Value      '切换是否开机自动启动后以后台模式运行
    If bHideWhenStart = False Then
        frmMain.Tray.Icon = frmMain.Icon
    End If
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkTraytWhenStart_Click()
    bTrayWhenStart = Me.chkTraytWhenStart.Value     '切换是否启动时最小化
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub chkUserPassword_Click()
    On Error Resume Next
    Me.labTip(1).Enabled = Me.chkUserPassword.Value     '更改控件状态
    Me.labTip(2).Enabled = Me.chkUserPassword.Value
    Me.edPassword1.Enabled = Me.chkUserPassword.Value
    Me.edPassword2.Enabled = Me.chkUserPassword.Value
    Me.cmdChangePassword.Enabled = Me.chkUserPassword.Value
    bUseUserPassword = Me.chkUserPassword.Value         '切换是否使用用户密码
    If Me.edPassword1.Enabled = True Then               '如果文本框可用就自动跳至文本框
        Me.edPassword1.SelStart = 0
        Me.edPassword1.SelLength = Len(Me.edPassword1.Text)
        Me.edPassword1.SetFocus
    End If
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub cmdChangePassword_Click()
    If Me.edPassword1.Text = Me.edPassword2.Text Then               '如果两次密码相符
        sUserPassword = Me.edPassword1.Text                             '更改密码
        Call SaveConfig                                                 '保存配置文件
        MsgBox "密码修改成功！", 64, "提示"
    Else
        MsgBox "两次输入的密码不符！", 48, "提示"
        Me.edPassword1.SetFocus
    End If
End Sub

Private Sub cmdDeleteAll_Click()
    Me.lstIP.Clear                      '全部清光光~~
    frmMain.comIP.Clear
    frmMain.lstPassword.Clear
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub cmdDeleteSelected_Click()
    On Error Resume Next                '防止删除空列表项出错 我懒~懒得处理了，直接略过~ (*^__^*)
    frmMain.comIP.RemoveItem Me.lstIP.ListIndex                     '删除记录的IP及密码
    frmMain.lstPassword.RemoveItem Me.lstIP.ListIndex
    Me.lstIP.RemoveItem Me.lstIP.ListIndex
    Call SaveConfig                                 '保存配置文件
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
    bExitWhenClose = Me.optExit.Value               '关闭时退出软件
    Call SaveConfig                                 '保存配置文件
End Sub

Private Sub optTray_Click()
    bExitWhenClose = Me.optExit.Value               '关闭时软件最小化到托盘区
    Call SaveConfig                                 '保存配置文件
End Sub
