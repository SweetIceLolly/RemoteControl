VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFileTransfer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "文件管理"
   ClientHeight    =   6432
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10248
   Icon            =   "frmFileTransfer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6432
   ScaleWidth      =   10248
   StartUpPosition =   2  '屏幕中心
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
         ToolTipText     =   "刷新"
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
         ToolTipText     =   "删除选中的文件"
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
         ToolTipText     =   "新建文件夹"
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
         ToolTipText     =   "返回上层目录"
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
         ToolTipText     =   "返回根目录"
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
         ToolTipText     =   "从对方电脑下载文件到本地"
         Top             =   0
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "下载"
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
         ToolTipText     =   "刷新"
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
         ToolTipText     =   "删除选中的文件"
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
         ToolTipText     =   "新建文件夹"
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
         ToolTipText     =   "返回上层目录"
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
         ToolTipText     =   "返回根目录"
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
         ToolTipText     =   "上传选择的文件到对方电脑"
         Top             =   0
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "上传"
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
      Text            =   "根目录"
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
      Text            =   "根目录"
      BackColor       =   16777215
      Appearance      =   4
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdOpenLocalDir 
      Height          =   255
      Left            =   4680
      TabIndex        =   6
      ToolTipText     =   "打开"
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
      ToolTipText     =   "打开"
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
      Caption         =   "日志"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "共0个对象，0个对象被选定，共0byte"
      Height          =   180
      Left            =   5280
      TabIndex        =   25
      Top             =   4800
      Width           =   2970
   End
   Begin VB.Label labSelectedLocal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "共0个对象，0个对象被选定，共0byte"
      Height          =   180
      Left            =   120
      TabIndex        =   24
      Top             =   4800
      Width           =   2970
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "目录："
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
      Caption         =   "目录："
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
      Caption         =   "远程计算机"
      BeginProperty Font 
         Name            =   "宋体"
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
      Caption         =   "本机"
      BeginProperty Font 
         Name            =   "宋体"
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
Dim sDataRec As String          '数据分段用的缓存：有时候数据量太大Winsock强制分包时就要用到我啦~
Dim bMultiData As Boolean       '是否为多包数据分段模式

Public Sub AddLog(strLog As String)             '日志记录过程
    Me.edLog.Text = Me.edLog.Text & Time & " " & strLog & vbCrLf
    Me.edLog.SelStart = Len(Me.edLog.Text)
    Me.Refresh
End Sub

'===================================================================================================
'生成本地文件列表的过程
Private Sub MakeList(DirPath As String)         '生成目录文件表的过程
    Dim tmp() As String                         '缓存字符串
    Dim FilePath As String                      '文件路径字符串
    On Error Resume Next
    '=================================
    Me.lstFileLocal.ListItems.Clear
    Me.lstFileLocal.ListItems.Add , , "..."
    Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "上层目录"
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
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "目录"
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = ""
    Next i
    '=================================
    Me.File.Path = DirPath
    Me.File.Refresh
    FilePath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
    For i = 0 To Me.File.ListCount - 1
        tmp = Split(Me.File.List(i), "\")
        Me.lstFileLocal.ListItems.Add , , tmp(UBound(tmp))
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "文件"
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = GetFileSize(FilePath & Me.File.List(i))
    Next i
    Call RefreshTip
End Sub

Private Sub MakeRootList()                       '生成磁盘根目录表的过程
    Dim tmpString As String
    Dim lSPC                         'Sectors Per Cluster        【每簇的扇区数】
    Dim lBPS                         'Bytes Per Sector           【每扇区的字节数】
    Dim lF                           'Number Of Free Clusters    【空闲簇的数量】
    Dim lT                           'Total Number Of Clusters   【簇的总数】
    Dim IsExists As Long             '磁盘是否有效
    '===============================
    Me.lstFileLocal.ListItems.Clear
    Me.Drive.Refresh
    For i = 0 To frmMain.Drive.ListCount - 1
        tmpString = frmMain.Drive.List(i)        '盘符/名称
        IsExists = GetDiskFreeSpace(Left(tmpString, 2), lSPC, lBPS, lF, lT)          '获取硬盘空间
        If IsExists <> 0 Then                    '如果能获取到磁盘信息说明磁盘有效
            Me.lstFileLocal.ListItems.Add , , frmMain.Drive.List(i)
            Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "驱动器"
            Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = "（空闲:" & SizeWithFormat(lSPC * lBPS * lF) & " 共:" & SizeWithFormat(lSPC * lBPS * lT) & "）"
        End If
    Next i
    Call RefreshTip
End Sub

Public Function GetListSelCount(lstTarget As ListView)
    Dim Total As Integer        '选择的项目数
    '=========================================
    For i = 1 To lstTarget.ListItems.Count
        If lstTarget.ListItems(i).Checked = True Then      '打了勾的就加进去~
            Total = Total + 1
        End If
    Next i
    GetListSelCount = Total
End Function

Public Function RefreshTip()
    On Error Resume Next
    Dim aSize, bSize                '两个列表中选择的项目的大小的总和
    Dim tmp                         '计算大小的缓存变量
    '====================================
    aSize = 0
    bSize = 0
    '---------------------
    For i = 1 To Me.lstFileLocal.ListItems.Count
        If Me.lstFileLocal.ListItems(i).Checked = True Then
            If Me.lstFileLocal.ListItems(i).Text <> "..." Then
                tmp = Me.lstFileLocal.ListItems(i).ListSubItems(2).Text     '获取大小
                If InStr(tmp, "空闲") = 0 Then                              '如果不是磁盘根目录
                    If tmp = "" Then                                            '如果是目录
                        tmp = 0
                    End If
                    If InStr(tmp, " Byte") <> 0 Then                            '如果单位是 Byte
                        tmp = Replace(tmp, " Byte", "")                             '去掉Byte单位
                    End If
                    If InStr(tmp, " KB") <> 0 Then                              '如果单位是 KB
                        tmp = Replace(tmp, " KB", "")                               '去掉KB单位
                        tmp = tmp * 1024                                            '乘以 1024
                    End If
                    If InStr(tmp, " MB") <> 0 Then                              '如果单位是 MB
                        tmp = Replace(tmp, " MB", "")                               '去掉MB单位
                        tmp = tmp * 1024 ^ 2                                        '乘以 1024^2
                    End If
                    If InStr(tmp, " GB") <> 0 Then                              '如果单位是 GB
                        tmp = Replace(tmp, " GB", "")                               '去掉GB单位
                        tmp = tmp * 1024 ^ 3                                        '乘以 1024^3
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
                tmp = Me.lstFileRemote.ListItems(i).ListSubItems(2).Text    '获取大小
                If InStr(tmp, "空闲") = 0 Then                              '如果不是磁盘根目录
                    If tmp = "" Then                                            '如果是目录
                        tmp = 0
                    End If
                    If InStr(tmp, " Byte") <> 0 Then                            '如果单位是 Byte
                        tmp = Replace(tmp, " Byte", "")                             '去掉Byte单位
                    End If
                    If InStr(tmp, " KB") <> 0 Then                              '如果单位是 KB
                        tmp = Replace(tmp, " KB", "")                               '去掉KB单位
                        tmp = tmp * 1024                                            '乘以 1024
                    End If
                    If InStr(tmp, " MB") <> 0 Then                              '如果单位是 MB
                        tmp = Replace(tmp, " MB", "")                               '去掉MB单位
                        tmp = tmp * 1024 ^ 2                                        '乘以 1024^2
                    End If
                    If InStr(tmp, " GB") <> 0 Then                              '如果单位是 GB
                        tmp = Replace(tmp, " GB", "")                               '去掉GB单位
                        tmp = tmp * 1024 ^ 3                                        '乘以 1024^3
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
        Me.labSelectedLocal.Caption = "共" & Me.lstFileLocal.ListItems.Count & "个对象，" & GetListSelCount(Me.lstFileLocal) & "个对象被选定，共" & SizeWithFormat(aSize)
    Else
        Me.labSelectedLocal.Caption = "共" & Me.lstFileLocal.ListItems.Count - 1 & "个对象，" & GetListSelCount(Me.lstFileLocal) & "个对象被选定，共" & SizeWithFormat(aSize)
    End If
    If Me.lstFileRemote.ListItems(1).Text <> "..." Then
        Me.labSelectedRemote.Caption = "共" & Me.lstFileRemote.ListItems.Count & "个对象，" & GetListSelCount(Me.lstFileRemote) & "个对象被选定，共" & SizeWithFormat(bSize)
    Else
        Me.labSelectedRemote.Caption = "共" & Me.lstFileRemote.ListItems.Count - 1 & "个对象，" & GetListSelCount(Me.lstFileRemote) & "个对象被选定，共" & SizeWithFormat(bSize)
    End If
End Function

Private Function IsIPExists(strIP As String) As Integer
    For i = 0 To frmMain.comIP.ListCount - 1                '扫一遍主窗体的IP列表
        If frmMain.comIP.List(i) = strIP Then                   '找到了相同的IP的话
            IsIPExists = i                                      '返回Index
            Exit Function                                       '退出过程
        End If
    Next i
    IsIPExists = -1                                      '找不到相同的话就返回-1
End Function

Private Sub ConnectedHandler()                   '连接后的处理过程
    Dim iLstIP As Integer                                             '列表中的IP位置
    Me.Caption = "文件管理: " & Me.wsMessage.RemoteHostIP             '更改标题
    iLstIP = IsIPExists(frmMain.comIP.Text)
    If iLstIP <> -1 Then                                                                '如果检测到这个IP连接过
        wsSendData Me.wsMessage, "CNT|" & frmMain.lstPassword.List(iLstIP)                  '发送带密码的请求
        Exit Sub
    End If
    If frmMain.lstPassword.List(frmMain.comIP.ListIndex) <> "" Then                         '如果记录里有密码
        wsSendData Me.wsMessage, "CNT|" & frmMain.lstPassword.List(frmMain.comIP.ListIndex)     '发送带密码的请求
    Else
        wsSendData Me.wsMessage, "CNT"                                                          '没密码就发送连接建立请求
    End If
End Sub

Private Sub cmdBackLocal_Click()
    Dim SplitTmp() As String            '分割缓存
    Dim tmpPath As String               '生成路径字符串的缓存
    '=============================================================
    If Me.lstFileLocal.ListItems(1).SubItems(1) <> "驱动器" Then
        If Right(Me.Dir.Path, 2) <> ":\" Then               '如果当前不是从根目录往上一级的话
            SplitTmp = Split(Me.File.Path, "\")             '显示上一级目录
            For i = 0 To UBound(SplitTmp) - 1
                tmpPath = tmpPath & SplitTmp(i) & "\"
            Next i
            Call MakeList(tmpPath)
        Else                                                '如果从根目录往上一级
            Call MakeRootList                               '显示磁盘根目录
        End If
    Else
        Call MakeRootList                               '显示磁盘根目录
    End If
    '显示当前的路径
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "驱动器" Then
        Me.edLocalPath.Text = "根目录"
    Else
        Me.edLocalPath.Text = Me.Dir.Path
    End If
End Sub

Private Sub cmdBackRemote_Click()
    If Me.lstFileRemote.ListItems(1).SubItems(1) <> "驱动器" Then
        wsSendData Me.wsMessage, "GUD"                            '发送打开上层目录请求
    Else
        wsSendData Me.wsMessage, "GRT"
    End If
End Sub

Private Sub cmdDeleteLocal_Click()
    On Error Resume Next
    Dim tmpPath As String                   '生成的路径的缓存
    Dim SelectedPath As String              '选择的路径
    Dim SplitPath() As String               '分割出来的每个路径及参数
    '=================================
    If GetListSelCount(Me.lstFileLocal) = 0 Then
        Exit Sub
    End If
    a = MsgBox("您真的要删除选择的" & GetListSelCount(Me.lstFileLocal) & "个文件（夹）吗？该操作不可撤销。", 32 + vbYesNo, "确认")
    If a <> 6 Then
        Exit Sub
    End If
    '----------------------------------------
    AddLog "【本地】开始删除选择的" & GetListSelCount(Me.lstFileLocal) & "个文件（夹）。"
    For i = 1 To Me.lstFileLocal.ListItems.Count                   '获取所有选择的列表项
        If Me.lstFileLocal.ListItems(i).Checked = True Then
            SelectedPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
            SelectedPath = SelectedPath & Me.lstFileLocal.ListItems(i)
            tmpPath = tmpPath & SelectedPath & "|" & Me.lstFileLocal.ListItems(i).SubItems(1) & vbCrLf
        End If
    Next i
    SplitPath = Split(tmpPath, vbCrLf)      '分割
    For i = 0 To UBound(SplitPath)          '得到所有选择的文件
        If Split(SplitPath(i), "|")(1) = "文件" Then               '如果是文件类型
            If bDeleteFiles Then
                Kill Split(SplitPath(i), "|")(0)
            Else
                MsgBox Split(SplitPath(i), "|")(0)
            End If
        End If
        If Split(SplitPath(i), "|")(1) = "目录" Then               '如果是目录类型
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
    Dim tmpSendString As String                     '发送数据缓存
    '----------------------------------
    If GetListSelCount(Me.lstFileRemote) = 0 Then
        Exit Sub
    End If
    a = MsgBox("您真的要删除选择的" & GetListSelCount(Me.lstFileRemote) & "个文件（夹）吗？该操作不可撤销。", 32 + vbYesNo, "确认")
    If a <> 6 Then
        Exit Sub
    End If
    AddLog "【远程】开始删除选择的" & GetListSelCount(Me.lstFileRemote) & "个文件（夹）。"
    tmpSendString = "DEL|"
    For i = 1 To Me.lstFileRemote.ListItems.Count
        If Me.lstFileRemote.ListItems(i).Checked = True Then      '往缓存里面添加打了勾的列表项
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '发送删除文件请求
End Sub

Private Sub cmdDownload_Click()
    Dim tmpSend As String                           '需要下载的文件序号列表
    Dim Reminded As Boolean                         '是否提醒过文件夹发不了
    '============================================================================
    If IsWorking Then                               '如果当前有任务正在进行
        AddLog "当前已经有正在进行中的任务，请等待到任务完成后再继续操作！"
        Exit Sub
    End If
    Reminded = False                                '未提醒过
    If GetListSelCount(Me.lstFileRemote) = 0 Then                            '如果未选择文件
        MsgBox "未选择文件。", 48, "提示"                                       '显示提示
        Exit Sub                                                                '退出过程
    End If
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "驱动器" Then       '如果目标是根目录
        MsgBox "无法下载到根目录。", 48, "提示"                                 '显示提示
        Exit Sub                                                                '退出过程
    End If
    '============================================================================
    For i = 1 To Me.lstFileRemote.ListItems.Count                           '获取所有选择的列表项
        If Me.lstFileRemote.ListItems(i).Checked = True Then                     '不是目录就加入列表
            If Me.lstFileRemote.ListItems(i).SubItems(1) <> "目录" Then
                tmpSend = tmpSend & CStr(i) & vbCrLf
            Else
                If Reminded = False Then                                    '如果未提醒过就提示目录不能发送
                    MsgBox "您选择的对象里包含文件夹，文件夹不支持发送，发送时将自动忽略。", 48, "提示"
                    Reminded = True                 '提醒过了
                End If
            End If
        End If
    Next i
    If tmpSend = "" Then                            '没有一个是文件？
        MsgBox "没有文件被发送。", 48, "提示"
        Exit Sub
    End If
    '==============================================
    frmDownload.IsUpload = False                                        '下载状态
    frmDownload.DownloadPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\") '设置下载路径
    frmDownload.wsFile.Bind 20236
    frmDownload.wsFile.Listen                                           '开始监听
    IsWorking = True
    frmDownload.Show
    wsSendData Me.wsMessage, "DWL|" & tmpSend                           '发送下载文件请求
End Sub

Private Sub cmdHomeLocal_Click()
    Call MakeRootList                               '生成根目录
End Sub

Private Sub cmdHomeRemote_Click()
    wsSendData Me.wsMessage, "GRT"                  '获取根目录列表请求
End Sub

Private Sub cmdMkdirLocal_Click()
    Dim fName As String                             '文件夹名字缓存
    On Error Resume Next
    If Me.lstFileLocal.ListItems(1).SubItems(1) <> "驱动器" Then      '如果是根目录列表
        fName = InputBox("请输入新文件夹的名字：", "输入")
        If Trim(fName) = "" Then
            Exit Sub
        End If
        fName = IIf(Right(Me.Dir.Path, 1) = "\", Me.Dir.Path, Me.Dir.Path & "\") & fName
        MkDir fName
        If Err.Number <> 0 Then
            AddLog "【本地】创建文件夹失败。"
        End If
    Else
        AddLog "【本地】不能在根目录里创建文件夹。"
    End If
End Sub

Private Sub cmdMkdirRemote_Click()
    Dim fName As String                             '文件夹名字缓存
    If Me.lstFileRemote.ListItems(1).SubItems(1) <> "驱动器" Then      '如果是根目录列表
        fName = InputBox("请输入新文件夹的名字：", "输入")
        If Trim(fName) <> "" Then
            If InStr(fName, "|") = 0 Then
                wsSendData Me.wsMessage, "MKD|" & fName         '发送创建文件夹请求
            Else
                AddLog "【远程】创建文件夹失败。"
            End If
        End If
    Else
        AddLog "【远程】不能在根目录里创建文件夹。"
    End If
End Sub

Private Sub cmdOpenLocalDir_Click()
    Dim vPath As String                         '要浏览的目录
    '===========================================================
    If Me.edLocalPath.Text = "根目录" Then
        Call MakeRootList
    Else
        vPath = Me.edLocalPath.Text
        If IsPathExists(vPath) = False Then                 '检测到请求的目录无效
            AddLog "【本地】无效的目录名，打开文件夹" & vPath & "失败。"
        Else
            Call MakeList(vPath)
        End If
    End If
End Sub

Private Sub cmdOpenRemoteDir_Click()
    wsSendData Me.wsMessage, "VPH|" & Me.edRemotePath.Text    '发送打开目录请求
End Sub

Public Sub cmdRefreshLocal_Click()
    '刷新新~
    Me.Dir.Refresh
    Me.File.Refresh
    Me.Drive.Refresh
    '------------------------
    If Me.lstFileLocal.ListItems(1).SubItems(1) = "驱动器" Then      '如果是根目录列表
        Call MakeRootList                                           '生成磁盘根目录列表
    Else                                                            '如果是普通目录就生成目录列表
        Call MakeList(Me.Dir.Path)
    End If
End Sub

Private Sub cmdRefreshRemote_Click()
    wsSendData Me.wsMessage, "REF"                  '发送刷新请求
End Sub

Private Sub cmdUpload_Click()
    Dim tmpPath As String                           '想要上传的文件
    Dim SelectedPath As String                      '每个文件的目录
    Dim Reminded As Boolean                         '是否提醒过文件夹发不了
    '============================================================================
    If IsWorking Then                               '如果当前有任务正在进行
        AddLog "当前已经有正在进行中的任务，请等待到任务完成后再继续操作！"
        Exit Sub
    End If
    Reminded = False                                '未提醒过
    If GetListSelCount(Me.lstFileLocal) = 0 Then                            '如果未选择文件
        MsgBox "未选择文件。", 48, "提示"                                       '显示提示
        Exit Sub                                                                '退出过程
    End If
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "驱动器" Then       '如果目标是根目录
        MsgBox "无法下载到根目录。", 48, "提示"                                 '显示提示
        Exit Sub                                                                '退出过程
    End If
    '============================================================================
    For i = 1 To Me.lstFileLocal.ListItems.Count                           '获取所有选择的列表项
        If Me.lstFileLocal.ListItems(i).Checked = True Then                     '不是目录就加入列表
            If Me.lstFileLocal.ListItems(i).SubItems(1) <> "目录" Then
                SelectedPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                SelectedPath = SelectedPath & Me.lstFileLocal.ListItems(i)
                tmpPath = tmpPath & SelectedPath & vbCrLf
            Else
                If Reminded = False Then                                    '如果未提醒过就提示目录不能发送
                    MsgBox "您选择的对象里包含文件夹，文件夹不支持发送，发送时将自动忽略。", 48, "提示"
                    Reminded = True                 '提醒过了
                End If
            End If
        End If
    Next i
    If tmpPath = "" Then                            '没有一个是文件？
        MsgBox "没有文件被发送。", 48, "提示"
        Exit Sub
    End If
    '==============================================
    frmDownload.FileList = tmpPath                                      '生成文件列表
    frmDownload.IsUpload = True                                         '上传状态
    frmDownload.wsFile.Bind 20236
    frmDownload.wsFile.Listen                                           '开始监听
    '以下为状态初始化
    frmDownload.CurrentFileIndex = -1                                   '初始化文件序号
    Call frmDownload.SplitPath                                          '分离出每个文件
    If frmDownload.NextFile = False Then                                '下一个文件
        Exit Sub                                                        '如果读取失败就直接退出过程
    End If
    IsWorking = True
    frmDownload.Show
    Call frmDownload.SendUploadMsg                                      '发送上传文件请求
End Sub

Private Sub edLocalPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOpenLocalDir_Click                   '调用按钮按下过程
    End If
End Sub

Private Sub edRemotePath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOpenRemoteDir_Click                   '调用按钮按下过程
    End If
End Sub

Private Sub Form_Load()
    '记录对方IP到列表里
    If bAutoRecord Then                 '如果不自动记录IP就算咯 ╮(╯_╰)╭
        If frmRemoteControl.IsIPExists(Me.wsMessage.RemoteHostIP) = -1 Then       '如果IP不存在就记录并保存
            frmMain.comIP.AddItem Me.wsMessage.RemoteHostIP
            frmMain.lstPassword.AddItem frmEnterPassword.edPassword.Text
            Call SaveConfig                         '保存配置文件
        End If
    End If
    '添加列表头
    Me.lstFileLocal.ColumnHeaders.Add , , "名称", 2000
    Me.lstFileLocal.ColumnHeaders.Add , , "类型", 800
    Me.lstFileLocal.ColumnHeaders.Add , , "大小", 3000
    Me.lstFileRemote.ColumnHeaders.Add , , "名称", 2000
    Me.lstFileRemote.ColumnHeaders.Add , , "类型", 800
    Me.lstFileRemote.ColumnHeaders.Add , , "大小", 3000
    '获取根目录列表
    Dim sRootMsg() As String                        '各列表项的详细信息切割缓存
    For i = 0 To frmMain.lstTemp.ListCount - 1
        sRootMsg = Split(frmMain.lstTemp.List(i), "|")
        Me.lstFileLocal.ListItems.Add , , sRootMsg(0)
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(1) = "驱动器"
        Me.lstFileLocal.ListItems(Me.lstFileLocal.ListItems.Count).SubItems(2) = sRootMsg(1)
    Next i
    Call MakeRootList
    IsWorking = False                               '没有任务正在进行
    AddLog "文件管理已启动。"
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
    lMid = Me.Width / 2             '窗体中心位置
    '-------------------------
    '下边
    Me.edLog.Width = Me.Width - Me.edLog.Left - 360
    Me.edLog.Top = Me.Height - Me.edLog.Height * 2 + 120
    Me.labTip(4).Top = Me.edLog.Top - Me.labTip(4).Height - 120
    '--------------------------------------------------------------------------------
    '左边
    Me.edLocalPath.Width = lMid - 240 - Me.cmdOpenLocalDir.Width - Me.edLocalPath.Left
    Me.cmdOpenLocalDir.Left = Me.edLocalPath.Left + Me.edLocalPath.Width
    Me.labTip(1).Left = lMid
    Me.picLocalToolbar.Width = lMid - Me.picLocalToolbar.Left - 240
    Me.cmdUpload.Left = Me.picLocalToolbar.Width - Me.cmdUpload.Width
    Me.lstFileLocal.Width = lMid - Me.lstFileLocal.Left - 240
    Me.labSelectedLocal.Top = Me.labTip(4).Top - Me.labSelectedLocal.Height - 120
    Me.lstFileLocal.Height = Me.labSelectedLocal.Top - Me.lstFileLocal.Top - 120
    '--------------------------------------------------------------------------------
    '右边
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
    '获取类型并分类型发送请求
    On Error Resume Next
    Dim SplitTmp() As String            '分割缓存
    Dim tmpPath As String               '生成路径字符串的缓存
    '===============================================================
    Select Case Me.lstFileLocal.ListItems.Item(Me.lstFileLocal.SelectedItem.Index).SubItems(1)
        Case "驱动器"
            tmpPath = Left(Me.lstFileLocal.ListItems.Item(Me.lstFileLocal.SelectedItem.Index), 2) & "\"
            Call MakeList(tmpPath)
            
        Case "目录"
            tmpPath = Me.Dir.List(Me.lstFileLocal.SelectedItem.Index - 2)
            Call MakeList(tmpPath)
            
        Case "上层目录"
            If Right(Me.Dir.Path, 2) <> ":\" Then               '如果当前不是从根目录往上一级的话
                SplitTmp = Split(Me.File.Path, "\")             '显示上一级目录
                For i = 0 To UBound(SplitTmp) - 1
                    tmpPath = tmpPath & SplitTmp(i) & "\"
                Next i
                Call MakeList(tmpPath)
            Else                                                '如果从根目录往上一级
                Call MakeRootList                               '显示磁盘根目录
            End If
            
    End Select
    If Me.lstFileLocal.ListItems.Item(1).SubItems(1) = "驱动器" Then
        Me.edLocalPath.Text = "根目录"
    Else
        Me.edLocalPath.Text = Me.Dir.Path
    End If
End Sub

Private Sub lstFileLocal_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
    On Error Resume Next
    For i = 1 To Me.lstFileLocal.ListItems.Count
        '如果是磁盘 或者 上层目录选项 或者 目录
        If InStr(Me.lstFileLocal.ListItems(i).SubItems(2), "空闲:") <> 0 Or Me.lstFileLocal.ListItems(i).SubItems(1) = "上层目录" Then
            Me.lstFileLocal.ListItems(i).Checked = False
        End If
        Call RefreshTip
    Next i
End Sub

Private Sub lstFileLocal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then        '按下退格键
        cmdBackLocal_Click                  '返回上层文件夹
    End If
    If KeyAscii = 13 Then               '按下回车键
        Call lstFileLocal_DblClick          '打开选择的文件
        KeyAscii = 0
    End If
End Sub

Private Sub lstFileRemote_DblClick()
    '获取类型并分类型发送请求
    On Error Resume Next
    Select Case Me.lstFileRemote.ListItems.Item(Me.lstFileRemote.SelectedItem.Index).SubItems(1)
        Case "驱动器"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFileRemote.SelectedItem.Index
            
        Case "目录"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFileRemote.SelectedItem.Index
            
        Case "上层目录"
            wsSendData Me.wsMessage, "GUD"
            
    End Select
End Sub

Private Sub lstFileRemote_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
    On Error Resume Next
    For i = 1 To Me.lstFileRemote.ListItems.Count
        '如果是磁盘 或者 上层目录选项
        If InStr(Me.lstFileRemote.ListItems(i).SubItems(2), "空闲:") <> 0 Or Me.lstFileRemote.ListItems(i).SubItems(1) = "上层目录" Then
            Me.lstFileRemote.ListItems(i).Checked = False
        End If
        Call RefreshTip
    Next i
End Sub

Private Sub lstFileRemote_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then        '按下退格键
        cmdBackRemote_Click                 '返回上层文件夹
    End If
    If KeyAscii = 13 Then               '按下回车键
        Call lstFileRemote_DblClick          '打开选择的文件
        KeyAscii = 0
    End If
End Sub

Private Sub wsMessage_Connect()
    '连接成功
    Call ConnectedHandler       '连接处理
End Sub

Private Sub wsMessage_DataArrival(ByVal bytesTotal As Long)
    Dim Temp As String                      '缓存数据
    Dim sTemp() As String                   '数据切割缓存数据
    Dim sTempPara() As String               '数据附带的参数切割出来的缓存数据
    Dim TaskSplit() As String               '每一个反馈信息的分割缓存（用于进程方面的）
    '====================================
    '处理数据
    Me.wsMessage.GetData Temp               '获取数据
    Temp = Decrypt(Temp)                            '解码数据
    If InStr(Temp, "{END}") = 0 Then                '找不到结束标记说明这是分包，先不处理消息
        sDataRec = sDataRec & Temp
        bMultiData = True
        Exit Sub
    Else
        If bMultiData = True Then                           '如果是多个数据包就用拼起来的数据
            sDataRec = Replace(sDataRec, "{END}", "")       '如果找到结束标记就说明接收了数据，开始处理消息 （把结束标记删掉）
        Else
            sDataRec = Replace(Temp, "{END}", "")           '如果是单个数据包就用单独接受到的数据
        End If
    End If
    sTemp = Split(sDataRec, "{S}")              '切割数据
    For i = 0 To UBound(sTemp)
        Select Case Left(sTemp(i), 3)       '分析数据
            Case "NAC"                                      '拒绝连接请求
                Me.wsMessage.Close
                frmMain.labState.Caption = "对方不允许远程控制"
                frmMain.shpState.FillColor = vbRed
                frmMain.tmrReturn.Enabled = True
        
            Case "CNT"                                      '连接请求反馈
                frmMain.Hide
                sTempPara = Split(sTemp(i), "|")
                Select Case sTempPara(1)                    '分类讨论返回类型
                    Case "True"                                 '允许连接
                        '进行控制初始化
                        '记录对方IP到列表里
                        If bAutoRecord Then                 '如果不自动记录IP就算咯 ╮(╯_╰)╭
                            If IsIPExists(Me.wsMessage.RemoteHostIP) = -1 Then               '如果IP不存在就记录并保存
                                frmMain.comIP.AddItem Me.wsMessage.RemoteHostIP
                                frmMain.lstPassword.AddItem frmEnterPassword.edPassword.Text
                                Call SaveConfig                         '保存配置文件
                            End If
                        End If
                        '===================================================
                        '隐藏/显示 窗体
                        frmEnterPassword.Hide
                        Me.Show
                        '===================================================
                        '请求传送系统信息
                        wsSendData Me.wsMessage, "GRT"            '发送获取根目录请求
                        
                    Case "Wrong"                                '密码错误导致的拒绝连接
                        Dim iIpList As Integer                          '检测IP是否存在的缓存
                        If frmEnterPassword.Visible = False Then
                            frmEnterPassword.Show
                        End If
                        iIpList = IsIPExists(Me.wsMessage.RemoteHostIP)
                        If iIpList <> -1 Then                           '如果检测到IP存在
                            frmMain.comIP.RemoveItem iIpList                '删除掉密码错误的列表项
                            frmMain.lstPassword.RemoveItem iIpList
                            Call SaveConfig                                 '保存配置文件
                        End If
                        frmEnterPassword.labIncorrect.Visible = True
                        frmEnterPassword.Height = 2400
                        frmEnterPassword.edPassword.SelStart = 0
                        frmEnterPassword.edPassword.SelLength = Len(frmEnterPassword.edPassword.Text)
                        frmEnterPassword.edPassword.SetFocus
                    
                    Case "False"                                '空密码导致的拒绝连接
                        frmEnterPassword.Show
                        frmEnterPassword.Height = 2000
                        
                End Select
                        
            Case "GRT"                                      '获取根目录请求响应
                Dim sRootList() As String                       '列表项切割缓存
                Dim sRootMsg() As String                        '各列表项的详细信息切割缓存
                '-------------------------------------
                sRootList() = Split(Replace(sDataRec, "GRT|", ""), "||")    '切割字符串
                Me.lstFileRemote.ListItems.Clear
                For k = 0 To UBound(sRootList) - 1
                    sRootMsg = Split(sRootList(k), "|")
                    Me.lstFileRemote.ListItems.Add , , sRootMsg(0)
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(1) = "驱动器"
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(2) = sRootMsg(1)
                Next k
                Me.edRemotePath.Text = "根目录"
                Call RefreshTip
            
            Case "GPH"                                      '获取目录下文件请求响应
                Dim sDirList() As String                        '列表项切割缓存
                Dim sDirMsg() As String                         '各列表项的详细信息切割缓存
                On Error Resume Next
                '-------------------------------------
                sDirList = Split(Replace(sDataRec, "GPH|", ""), "||")       '切割字符串
                Me.lstFileRemote.ListItems.Clear
                For k = 0 To UBound(sDirList) - 1
                    sDirMsg = Split(sDirList(k), "|")
                    Me.lstFileRemote.ListItems.Add , , sDirMsg(0)
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(1) = sDirMsg(1)
                    sDirMsg(2) = Replace(sDirMsg(2), ":N:", "")         '这个是防止文件夹获取不了大小导致的空白引发错误的
                    Me.lstFileRemote.ListItems(Me.lstFileRemote.ListItems.Count).SubItems(2) = sDirMsg(2)
                Next k
                Call RefreshTip
                wsSendData Me.wsMessage, "NPH"                  '发送获取当前目录路径的请求
            
            Case "NPH"
                Me.edRemotePath.Text = Replace(Replace(sDataRec, "NPH|", ""), "{S}", "")      '显示返回的文件夹路径
            
            Case "VPH"
                AddLog "【远程】无效的目录名，打开文件夹" & Me.edRemotePath.Text & "失败。"      '显示错误消息
                wsSendData Me.wsMessage, "NPH"                  '发送获取当前目录路径的请求
            
            Case "MKD"
                Dim bSuccess As String
                bSuccess = Replace(Replace(sDataRec, "MKD|", ""), "{S}", "")
                If bSuccess = "0" Then
                    AddLog "【远程】创建文件夹失败。"
                Else
                    AddLog "【远程】创建文件夹成功！"
                End If
                wsSendData Me.wsMessage, "REF"                  '发送刷新请求
            
        End Select
    Next i
    '-----------------------------
    sDataRec = ""                       '执行完所有命令后清除缓存
    bMultiData = False
End Sub
