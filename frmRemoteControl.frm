VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRemoteControl 
   BackColor       =   &H00FFFFFF&
   Caption         =   "远程控制"
   ClientHeight    =   5592
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10860
   Icon            =   "frmRemoteControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5592
   ScaleWidth      =   10860
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrDataPerSecond 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7440
      Top             =   4920
   End
   Begin XtremeSuiteControls.TabControl TabMain 
      Height          =   3732
      Left            =   0
      TabIndex        =   1
      Top             =   760
      Width           =   9132
      _Version        =   786432
      _ExtentX        =   16113
      _ExtentY        =   6588
      _StockProps     =   68
      Appearance      =   10
      Color           =   32
      ItemCount       =   6
      Item(0).Caption =   "屏幕控制"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "ResizerPicture"
      Item(1).Caption =   "进程管理"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lstTask"
      Item(1).Control(1)=   "labTaskTotal"
      Item(1).Control(2)=   "cmdRefreshTasks"
      Item(1).Control(3)=   "cmdKillTasks"
      Item(2).Caption =   "文件管理"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "lstFile"
      Item(2).Control(1)=   "labTip"
      Item(2).Control(2)=   "edPath"
      Item(2).Control(3)=   "cmdOpenDir"
      Item(3).Caption =   "运行VBS脚本"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "edVBS"
      Item(3).Control(1)=   "cmdRunVBS"
      Item(4).Caption =   "命令行"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "edCommandLine"
      Item(4).Control(1)=   "cmdRunCommandline"
      Item(4).Control(2)=   "edCommandlineEnter"
      Item(5).Caption =   "系统信息"
      Item(5).ControlCount=   1
      Item(5).Control(0)=   "edInformation"
      Begin XtremeSuiteControls.ListView lstFile 
         Height          =   2775
         Left            =   -69880
         TabIndex        =   14
         Top             =   840
         Visible         =   0   'False
         Width           =   8895
         _Version        =   786432
         _ExtentX        =   15690
         _ExtentY        =   4895
         _StockProps     =   77
         BackColor       =   16777152
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         BackColor       =   16777152
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lstTask 
         Height          =   2775
         Left            =   -69880
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   8535
         _Version        =   786432
         _ExtentX        =   15055
         _ExtentY        =   4895
         _StockProps     =   77
         BackColor       =   16777152
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         BackColor       =   16777152
         Appearance      =   6
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox edCommandlineEnter 
         BackColor       =   &H00FFFFC0&
         Height          =   372
         Left            =   -69880
         TabIndex        =   13
         Top             =   3240
         Visible         =   0   'False
         Width           =   7452
      End
      Begin VB.TextBox edCommandLine 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -69880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   8895
      End
      Begin VB.TextBox edVBS 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -69880
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   8895
      End
      Begin XtremeSuiteControls.PushButton cmdRefreshTasks 
         Height          =   375
         Left            =   -69880
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "点我刷新~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin VB.TextBox edInformation 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -69880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "frmRemoteControl.frx":0CCA
         Top             =   360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin XtremeSuiteControls.Resizer ResizerPicture 
         Height          =   2055
         Left            =   0
         TabIndex        =   3
         Top             =   360
         Width           =   4095
         _Version        =   786432
         _ExtentX        =   7223
         _ExtentY        =   3625
         _StockProps     =   1
         BackColor       =   0
         Begin VB.PictureBox picKeyboard 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   840
            ScaleHeight     =   348
            ScaleWidth      =   348
            TabIndex        =   21
            Top             =   120
            Width           =   375
         End
         Begin VB.PictureBox picMain 
            AutoRedraw      =   -1  'True
            Height          =   612
            Left            =   0
            ScaleHeight     =   564
            ScaleWidth      =   564
            TabIndex        =   4
            Top             =   0
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.Image imgMain 
            BorderStyle     =   1  'Fixed Single
            Height          =   1215
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1575
         End
      End
      Begin XtremeSuiteControls.PushButton cmdKillTasks 
         Height          =   372
         Left            =   -68320
         TabIndex        =   8
         Top             =   3240
         Visible         =   0   'False
         Width           =   1332
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "点我结束~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin XtremeSuiteControls.PushButton cmdRunVBS 
         Height          =   375
         Left            =   -69880
         TabIndex        =   10
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "点我执行~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin XtremeSuiteControls.PushButton cmdRunCommandline 
         Height          =   372
         Left            =   -62320
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   1332
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "点我执行~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin XtremeSuiteControls.PushButton cmdOpenDir 
         Height          =   375
         Left            =   -61960
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "转到"
         ForeColor       =   16777215
         BackColor       =   16416003
         Appearance      =   3
         DrawFocusRect   =   0   'False
         ImageGap        =   1
      End
      Begin XtremeSuiteControls.FlatEdit edPath 
         Height          =   375
         Left            =   -68800
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   6735
         _Version        =   786432
         _ExtentX        =   11880
         _ExtentY        =   661
         _StockProps     =   77
         BackColor       =   16777152
         Text            =   "根目录"
         BackColor       =   16777152
         Appearance      =   4
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前目录："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -69880
         TabIndex        =   15
         Top             =   400
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Label labTaskTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "共0个进程。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   216
         Left            =   -66760
         TabIndex        =   7
         Top             =   3360
         Visible         =   0   'False
         Width           =   1152
      End
   End
   Begin XtremeSuiteControls.Resizer ResizerMain 
      Height          =   765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
      _Version        =   786432
      _ExtentX        =   17383
      _ExtentY        =   1349
      _StockProps     =   1
      BackColor       =   16777088
      VScrollLargeChange=   1000
      VScrollSmallChange=   200
      HScrollLargeChange=   1000
      HScrollSmallChange=   200
      BorderStyle     =   4
      Begin XtremeSuiteControls.PushButton cmdBlockInput 
         Height          =   720
         Left            =   0
         TabIndex        =   18
         ToolTipText     =   "锁定对方的鼠标键盘"
         Top             =   0
         Width           =   2600
         _Version        =   786432
         _ExtentX        =   4586
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "锁定远程鼠标键盘"
         BackColor       =   16777152
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmRemoteControl.frx":0CD6
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton cmdIsControlling 
         Height          =   720
         Left            =   2640
         TabIndex        =   19
         ToolTipText     =   "开启控制对方的鼠标键盘"
         Top             =   0
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4586
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "开启控制对方鼠标键盘"
         BackColor       =   16777152
         TextAlignment   =   1
         Appearance      =   6
         Checked         =   -1  'True
         Picture         =   "frmRemoteControl.frx":19B0
         ImageAlignment  =   0
      End
      Begin XtremeSuiteControls.PushButton cmdAutoResize 
         Height          =   720
         Left            =   5280
         TabIndex        =   20
         ToolTipText     =   "自动拉伸屏幕画面"
         Top             =   0
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4586
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "自动拉伸屏幕画面"
         BackColor       =   16777152
         TextAlignment   =   1
         Appearance      =   6
         Picture         =   "frmRemoteControl.frx":268A
         ImageAlignment  =   0
      End
   End
   Begin MSWinsockLib.Winsock wsPicture 
      Left            =   8040
      Top             =   4920
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsMessage 
      Left            =   8640
      Top             =   4920
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuRefresh 
         Caption         =   "刷新(&R)"
      End
      Begin VB.Menu mnuMkdir 
         Caption         =   "新建文件夹(&N)"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "重命名(&M)"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "剪切(&U)"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "粘贴(&P)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "删除(&D)"
      End
   End
End
Attribute VB_Name = "frmRemoteControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0                                                '     color   table   in   RGBs

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Type BITMAPINFOHEADER                                                   '40 bytes
    biSize As Long                                                              'BITMAPINFOHEADER结构的大小
    biWidth As Long
    biHeight As Long
    biPlanes As Integer                                                         '设备的为平面数，现在都是1
    biBitCount As Integer                                                       '图像的颜色位图
    biCompression As Long                                                       '压缩方式
    biSizeImage As Long                                                         '实际的位图数据所占字节
    biXPelsPerMeter As Long                                                     '目标设备的水平分辨率
    biYPelsPerMeter As Long                                                     '目标设备的垂直分辨率
    biClrUsed As Long                                                           '使用的颜色数
    biClrImportant As Long                                                      '重要的颜色数 如果该项为0，表示所有颜色都是重要的
End Type
  
Private Type RGBQUAD                                                            '只有bibitcount为1，2，4时才有调色板
    Blue As Byte                                                                '蓝色分量
    Green As Byte                                                               '绿色分量
    Red As Byte                                                                 '红色分量
    Reserved As Byte                                                            '保留值
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As RGBQUAD
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'=========================================================================
'以下为zlib压缩的函数
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Const OFFSET As Long = &H8

'以下为屏幕显示参数及各种屏幕数据
Dim W As Long, H As Long, M As Long
Dim x, y As Long, dwCurrentThreadID1 As Long
Dim J As Byte, g As Long
Dim T As Boolean
Dim bmpLen As Long, F As Byte
Dim CapName As String
Dim iBitmap2 As Long, iDC2 As Long
Dim iBitmap1 As Long, iDC1 As Long
Dim iBitmap3 As Long, iDC3 As Long
Dim bi24BitInfo As BITMAPINFO
Dim xData As Long

'==========================================
'以下为各种乱七八糟的变量
Dim TwipsX, TwipsY As Single    '屏幕每个像素的Twip数
Dim DataPerSecond As Long       '每秒钟的流量数
Dim DataPerSecondOld As Long    '上一次计数的每秒钟流量数
Dim sDataRec As String          '数据分段用的缓存：有时候数据量太大Winsock强制分包时就要用到我啦~
Dim bMultiData As Boolean       '是否为多包数据分段模式

Function UnCompressByte(ByteArray() As Byte)                '解压缩过程
    Dim origLen As Long
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    Call CopyMemory(origLen, ByteArray(0), OFFSET)
    BufferSize = origLen
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    UnCompressByte = (uncompress(TempBuffer(0), BufferSize, ByteArray(OFFSET), UBound(ByteArray) - OFFSET + 1) = 0)
    ReDim Preserve ByteArray(0 To BufferSize - 1)
    CopyMemory ByteArray(0), TempBuffer(0), BufferSize
End Function

Public Sub Del_dc()                                         '删除内存DC过程
    DeleteDC iDC1
    DeleteObject iBitmap1
    DeleteDC iDC2
    DeleteObject iBitmap2
End Sub

Public Function DCload()
    iDC1 = CreateCompatibleDC(Me.hdc)                                           '1
    iBitmap1 = CreateCompatibleBitmap(Me.hdc, W, H * F)                         '显示区域设置
    SelectObject iDC1, iBitmap1
    
    iDC2 = CreateCompatibleDC(Me.hdc)                                           '2
    iBitmap2 = CreateCompatibleBitmap(Me.hdc, W, H)                             '显示区域设置
    SelectObject iDC2, iBitmap2
End Function

Private Sub GetWindows()                                '获取接收到的屏幕
    Dim Bmp As BITMAP
    Dim hdc As Long, BD As Long
    Dim Tpicture As Boolean
    Dim bBytes() As Byte
    '=======
    g = J * H
    '============
    If T Then
        Me.wsPicture.GetData bBytes, vbArray Or vbByte
        If bCompress = True Then
            UnCompressByte bBytes                                                   '解压
        End If
        SetDIBitsToDevice iDC2, 0, 0, W, H, 0, 0, 0, H, bBytes(0), bi24BitInfo, DIB_RGB_COLORS
        BitBlt iDC1, 0, g, W, H, iDC2, 0, 0, vbSrcInvert                        'xor'相反
        BitBlt Me.picMain.hdc, 0, g, W, H, iDC1, 0, g, vbSrcCopy
        Me.picMain.Refresh
        Me.imgMain.Picture = Me.picMain.Image
    End If
    '============
    If Not T Then
        Call New_dc
        T = True
    End If
    
    Dim a(2) As Byte
    Me.wsPicture.SendData a
    '========
    
End Sub

Private Sub New_dc()                                    '创建内存DC过程
    Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte
    
    Me.wsPicture.GetData bBytes, vbArray Or vbByte
    
    If bCompress = True Then
        UnCompressByte bBytes                                                       '解压
    End If
    
    With bi24BitInfo.bmiHeader
        .biBitCount = 16
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = W
        .biHeight = H * F
    End With
    
    SetDIBitsToDevice iDC1, 0, 0, W, H * F, 0, 0, 0, H * F, bBytes(0), bi24BitInfo, DIB_RGB_COLORS
    
End Sub

Public Function IsIPExists(strIP As String) As Integer
    For i = 0 To frmMain.comIP.ListCount - 1                '扫一遍主窗体的IP列表
        If frmMain.comIP.List(i) = strIP Then                   '找到了相同的IP的话
            IsIPExists = i                                          '返回Index
            Exit Function                                           '退出过程
        End If
    Next i
    IsIPExists = -1                                      '找不到相同的话就返回-1
End Function

'=======================================================================

Public Sub ConnectedHandler()                 '连接后的处理过程
    Dim iLstIP As Integer                                             '列表中的IP位置
    Me.Caption = "远程控制: " & Me.wsMessage.RemoteHostIP             '更改标题
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

Public Function GetListSelCount() As Integer    '获取列表中选择的项目数
    Dim Total As Integer        '选择的项目数
    '=========================================
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '打了勾的就加进去~
            Total = Total + 1
        End If
    Next i
    GetListSelCount = Total
End Function

Private Sub cmdAutoResize_Click()
    Me.cmdAutoResize.Checked = Not Me.cmdAutoResize.Checked             '反选
    AutoResize = Me.cmdAutoResize.Checked                               '调整是否拉伸画面状态
    Call Form_Resize                                                    '窗体重新排版
    Call SaveConfig                                                     '保存配置文件
End Sub

Private Sub cmdBlockInput_Click()
    Me.cmdBlockInput.Checked = Not Me.cmdBlockInput.Checked             '反选
    If Me.cmdBlockInput.Checked = True Then                             '锁定状态
        wsSendData Me.wsMessage, "BLK"                                      '发送锁定鼠标键盘消息
    Else
        wsSendData Me.wsMessage, "ULK"                                      '发送解锁鼠标键盘消息
    End If
End Sub

Private Sub cmdIsControlling_Click()
    Me.cmdIsControlling.Checked = Not Me.cmdIsControlling.Checked       '反选
    IsControlling = Me.cmdIsControlling.Checked                         '调整控制状态
End Sub

Private Sub cmdKillTasks_Click()
    Dim SendTemp As String              '要发送的数据缓存
    Dim IsChecked As Boolean            '是否有选择东西
    '========================================
    IsChecked = False
    For i = 1 To Me.lstTask.ListItems.Count                 '遍历列表，看看要结束哪个
        If Me.lstTask.ListItems(i).Checked = True Then
            '加入到发送数据缓存中
            SendTemp = SendTemp & Me.lstTask.ListItems(i).Text & "|" & Me.lstTask.ListItems(i).SubItems(1) & vbCrLf
            IsChecked = True                                        '有发送的东西
        End If
    Next i
    If IsChecked = False Then                               '没有选择东西？！
        MsgBox "你都没选择进程让我结束什么？！", 64
    Else
        a = MsgBox("确认结束选择的进程？这可能造成数据丢失等后果！", 32 + vbYesNo, "确认")      '确认框
        If a <> 6 Then
            Exit Sub
        End If
        wsSendData Me.wsMessage, "KTS|" & SendTemp          '如果有选择进程就发送结束进程请求
    End If
End Sub

Private Sub cmdOpenDir_Click()
    wsSendData Me.wsMessage, "VPH|" & Me.edPath.Text    '发送打开目录请求
End Sub

Private Sub cmdRefreshTasks_Click()
    wsSendData Me.wsMessage, "TSK"                      '发送获取进程请求
End Sub

Private Sub cmdRunCommandline_Click()
    If Trim(Me.edCommandlineEnter.Text) = "" Then       '不接受空命令
        Exit Sub
    End If
    wsSendData Me.wsMessage, "CMD|" & Me.edCommandlineEnter.Text        '发送执行命令行请求
    Me.edCommandlineEnter.Text = "执行中，请稍后..."
    Me.edCommandlineEnter.Enabled = False
End Sub

Private Sub cmdRunVBS_Click()
    wsSendData Me.wsMessage, "VBS|" & Me.edVBS.Text     '发送执行VBS请求
End Sub

Private Sub edCommandlineEnter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then                   '按下回车键就执行命令
        Call cmdRunCommandline_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOpenDir_Click                   '调用按钮按下过程
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.imgMainIcon.Picture       '更改图标，为了更好看 (*^__^*)
    '==================================================
    Me.lstTask.ColumnHeaders.Add , , "进程名称", 1500
    Me.lstTask.ColumnHeaders.Add , , "PID", 600
    Me.lstTask.ColumnHeaders.Add , , "程序路径", 4000
    '==================================================
    Me.lstFile.ColumnHeaders.Add , , "名称", 2000
    Me.lstFile.ColumnHeaders.Add , , "类型", 800
    Me.lstFile.ColumnHeaders.Add , , "大小", 3000
    '==================================================
    Me.picKeyboard.Left = -10
    Me.picKeyboard.Top = -10
    Me.picKeyboard.Width = 1
    Me.picKeyboard.Height = 1
    IsControlling = True                        '控制模式
    '==================================================
    If AutoResize Then
        Me.cmdAutoResize.Checked = True
    End If
    '==================================================
    If bMouseWheelHook = True Then
        PrevWndProc = SetWindowLong(Me.picKeyboard.hwnd, GWL_WNDPROC, AddressOf WndProc)        '启动鼠标滚轮事件拦截
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Height < 3000 Then
        Me.Height = 3000
    End If
    If Me.Width < 4000 Then
        Me.Width = 4000
    End If
    '-------------------------------------------------
    '调整各种控件的大小和状态以适应窗体
    Me.ResizerMain.Width = Me.Width - 240
    '==========================================
    Me.TabMain.Width = Me.Width - 240                               '总选项卡
    Me.TabMain.Height = Me.Height - Me.ResizerMain.Height - 480
    '==========================================
    Me.ResizerPicture.Width = Me.TabMain.Width - 120                '屏幕显示部分
    Me.ResizerPicture.Height = Me.TabMain.Height - 480
    If AutoResize = True Then                                       '如果是自动拉伸画面的话就让图片框适应窗体
        Me.imgMain.Width = Me.ResizerPicture.Width
        Me.imgMain.Height = Me.ResizerPicture.Height
        Me.ResizerPicture.HScrollMaximum = 0                            '不加滚动条
        Me.ResizerPicture.VScrollMaximum = 0
    Else                                                            '否则让他适应原图
        Me.imgMain.Width = Me.picMain.Width
        Me.imgMain.Height = Me.picMain.Height
        Me.ResizerPicture.HScrollMaximum = Me.picMain.Width             '添加滚动条
        Me.ResizerPicture.VScrollMaximum = Me.picMain.Height
    End If
    ScreenScaleW = Me.imgMain.Width / (W * TwipsX)                  '计算压缩图和原图的比例尺
    ScreenScaleH = Me.imgMain.Height / (H * F * TwipsY)
    '==========================================
    Me.lstTask.Width = Me.TabMain.Width - 240                       '进程部分
    Me.lstTask.Height = Me.TabMain.Height - 1000
    Me.cmdRefreshTasks.Top = Me.lstTask.Top + Me.lstTask.Height + 120
    Me.cmdKillTasks.Top = Me.cmdRefreshTasks.Top
    Me.labTaskTotal.Top = Me.cmdRefreshTasks.Top
    '==========================================
    Me.cmdOpenDir.Left = Me.TabMain.Width - Me.cmdOpenDir.Width - 120   '文件管理部分
    Me.edPath.Width = Me.cmdOpenDir.Left - Me.edPath.Left - 120
    Me.lstFile.Height = Me.TabMain.Height - 240 - Me.lstFile.Top
    Me.lstFile.Width = Me.TabMain.Width - 240
    '==========================================
    Me.edVBS.Width = Me.TabMain.Width - 240                         'VBS代码部分
    Me.edVBS.Height = Me.TabMain.Height - 1000
    Me.cmdRunVBS.Top = Me.cmdRefreshTasks.Top
    '==========================================
    Me.edInformation.Width = Me.TabMain.Width - 120                 '系统信息部分
    Me.edInformation.Height = Me.TabMain.Height - 480
    '==========================================
    Me.edCommandLine.Width = Me.TabMain.Width - 120                 '命令行操作部分
    Me.edCommandLine.Height = Me.TabMain.Height - 1000
    Me.edCommandlineEnter.Width = Me.TabMain.Width - 360 - Me.cmdRunCommandline.Width
    Me.edCommandlineEnter.Top = Me.lstTask.Top + Me.lstTask.Height + 120
    Me.cmdRunCommandline.Top = Me.cmdRefreshTasks.Top
    Me.cmdRunCommandline.Left = Me.edCommandlineEnter.Left + Me.edCommandlineEnter.Width + 120
End Sub

Private Sub Form_Unload(Cancel As Integer)      '关闭窗体即断开连接
    Me.wsMessage.Close
    Me.wsPicture.Close
    Dim OpenPath As String                  '程序自身的路径
    OpenPath = IIf(Right(App.Path, 1) = "\", App.Path & App.EXEName & ".exe", App.Path & "\" & App.EXEName & ".exe")
    Shell OpenPath, vbNormalFocus               '沉舟侧畔千帆过 病树前头万木春
    End                                         '我老了，不中用了，该上路了。
End Sub

Private Sub lstFile_DblClick()
    '获取类型并分类型发送请求
    On Error Resume Next
    Select Case Me.lstFile.ListItems.Item(Me.lstFile.SelectedItem.Index).SubItems(1)
        Case "驱动器"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFile.SelectedItem.Index
            
        Case "目录"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFile.SelectedItem.Index
            
        Case "上层目录"
            wsSendData Me.wsMessage, "GUD"
            
    End Select
End Sub

Private Sub lstFile_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
    On Error Resume Next
    For i = 1 To Me.lstFile.ListItems.Count
        '如果是磁盘 或者 上层目录选项
        If InStr(Me.lstFile.ListItems(i).SubItems(2), "空闲:") <> 0 Or Me.lstFile.ListItems(i).SubItems(1) = "上层目录" Then
            Me.lstFile.ListItems(i).Checked = False
        End If
    Next i
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sCount As Integer                           '选择了的列表项数
    '----------------------------------
    If Button = 2 Then                              '按下右键弹出菜单
        sCount = GetListSelCount
        Me.mnuMkdir.Visible = True
        Me.mnuPaste.Visible = True
        If sCount = 0 Then                     '分情形弹出菜单
            Me.mnuCopy.Visible = False                  '没有选择列表项则 复制、剪切、删除、重命名功能失效
            Me.mnuCut.Visible = False
            Me.mnuDelete.Visible = False
            Me.mnuRename.Visible = False
        Else
            Me.mnuCopy.Visible = True
            Me.mnuCut.Visible = True
            Me.mnuDelete.Visible = True
            Me.mnuRename.Visible = True
        End If
        If sCount > 1 Then                              '如果选择了多于一个列表项则 重命名功能失效
            Me.mnuRename.Visible = False
        End If
        If InStr(Me.lstFile.ListItems(1).ListSubItems(2).Text, "空闲:") <> 0 Then '如果是磁盘根目录则只留下刷新
            Me.mnuMkdir.Visible = False
            Me.mnuPaste.Visible = False
        End If
        PopupMenu Me.mnuPopup
    End If
End Sub

Private Sub mnuCopy_Click()
    Dim tmpSendString As String                     '发送数据缓存
    '----------------------------------
    tmpSendString = "CPY|"
    Me.mnuPaste.Enabled = True
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '往缓存里面添加打了勾的列表项
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '发送复制文件请求
End Sub

Private Sub mnuCut_Click()
    Dim tmpSendString As String                     '发送数据缓存
    '----------------------------------
    tmpSendString = "CUT|"
    Me.mnuPaste.Enabled = True
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '往缓存里面添加打了勾的列表项
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '发送剪切文件请求
End Sub

Private Sub mnuDelete_Click()
    Dim tmpSendString As String                     '发送数据缓存
    '----------------------------------
    If GetListSelCount = 0 Then
        Exit Sub
    End If
    a = MsgBox("您真的要删除选择的" & GetListSelCount & "个文件（夹）吗？该操作不可撤销。", 32 + vbYesNo, "确认")
    If a <> 6 Then
        Exit Sub
    End If
    tmpSendString = "DEL|"
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '往缓存里面添加打了勾的列表项
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '发送删除文件请求
End Sub

Private Sub mnuMkdir_Click()
    Dim fName As String                             '文件夹名字缓存
    fName = InputBox("请输入新文件夹的名字：", "输入")
    If Trim(fName) <> "" Then
        If InStr(fName, "|") = 0 Then
            wsSendData Me.wsMessage, "MKD|" & fName         '发送创建文件夹请求
        Else
            MsgBox "创建文件夹失败。", 64, "创建文件夹失败"
        End If
    End If
End Sub

Private Sub mnuPaste_Click()
    wsSendData Me.wsMessage, "PST"                  '发送粘贴请求
End Sub

Private Sub mnuRefresh_Click()
    wsSendData Me.wsMessage, "REF"                  '发送刷新请求
End Sub

Private Sub mnuRename_Click()
    Dim tmpName As String                           '文件名缓存
    tmpName = InputBox("请输入新的文件名：", "重命名")
    If Trim(tmpName) <> "" Then
        If InStr(tmpName, "|") = 0 Then
            wsSendData Me.wsMessage, "REN|" & Me.lstFile.SelectedItem.Index & "|" & tmpName         '发送文件重命名请求
        Else
            MsgBox "文件重命名失败。", 64, "文件重命名失败"
        End If
    End If
End Sub

Private Sub imgMain_DblClick()
    If IsControlling Then
        wsSendData Me.wsMessage, "DBC"                  '双击左键
    End If
End Sub

Private Sub PicKeyboard_KeyDown(KeyCode As Integer, Shift As Integer)
    If IsControlling Then
        wsSendData Me.wsMessage, "KBD|" & CStr(KeyCode)
    End If
End Sub

Private Sub PicKeyboard_KeyUp(KeyCode As Integer, Shift As Integer)
    If IsControlling Then
        wsSendData Me.wsMessage, "KBU|" & CStr(KeyCode)
    End If
End Sub

Private Sub imgMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsControlling Then
        Select Case Button
            Case 1                                  '左下
                wsSendData Me.wsMessage, "MLD"
                
            Case 2                                  '右下
                wsSendData Me.wsMessage, "MRD"
            
            Case 4                                  '中下
                wsSendData Me.wsMessage, "MMD"
            
        End Select
    End If
    Me.picKeyboard.SetFocus
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rX, rY As Single                        '在原图上的X坐标和Y坐标
    If IsControlling Then
        rX = x / ScreenScaleW / TwipsX
        rY = y / ScreenScaleH / TwipsY
        wsSendData Me.wsMessage, "MMP|" & CStr(rX) & "|" & CStr(rY)      '发送鼠标移动的坐标
    End If
End Sub

Private Sub imgMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsControlling Then
        Select Case Button
            Case 1                                  '左上
                wsSendData Me.wsMessage, "MLU"
                
            Case 2                                  '右上
                wsSendData Me.wsMessage, "MRU"
            
            Case 4                                  '中上
                wsSendData Me.wsMessage, "MMU"
            
        End Select
    End If
End Sub

Private Sub tmrDataPerSecond_Timer()
    '显示出每秒钟的流量数
    Dim Calc As Long
    Calc = DataPerSecond - DataPerSecondOld         '减去
    '更改标题
    Me.Caption = "远程控制: " & Me.wsMessage.RemoteHostIP & " 当前流量：" & Format((Abs(Calc)) / 1024, "0.00") & "kb/s"
    DataPerSecondOld = DataPerSecond                        '沉舟侧畔千帆过 病树前头万木春
    DataPerSecond = 0                                       '清空每秒钟的流量数
End Sub

Private Sub wsMessage_Close()
    Unload Me
End Sub

Private Sub wsMessage_Connect()
    frmMain.tmrTimeOut.Enabled = False          '连接成功后停止超时计时器
    Call ConnectedHandler
End Sub

Private Sub wsMessage_DataArrival(ByVal bytesTotal As Long)
    Dim Temp As String                      '缓存数据
    Dim sTemp() As String                   '数据切割缓存数据
    Dim sTempPara() As String               '数据附带的参数切割出来的缓存数据
    Dim TaskSplit() As String               '每一个反馈信息的分割缓存（用于进程方面的）
    '====================================
    '统计数据量
    DataPerSecond = DataPerSecond + bytesTotal
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
                Me.wsPicture.Close
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
                        wsSendData Me.wsMessage, "SIF"            '发送获取系统信息请求
                        
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
            
            Case "RES"                                      '对方分辨率反馈
                sTempPara = Split(sTemp(i), "|")
                '设置屏幕传输参数
                W = CLng(sTempPara(1))                      '给屏幕大小参数设置上接收到的对方的屏幕的分辨率
                H = CLng(sTempPara(2))
                TwipsX = CSng(sTempPara(3))
                TwipsY = CSng(sTempPara(4))
                '--------------------------------------
                '配置内存位图信息
                F = 6                                                                       '分多少行
                H = H \ F
                bmpLen = W * H * 2
                xData = 8
                With bi24BitInfo.bmiHeader
                    .biBitCount = 16
                    .biCompression = BI_RGB
                    .biPlanes = 1
                    .biSize = Len(bi24BitInfo.bmiHeader)
                    .biWidth = W
                    .biHeight = H
                End With
                '加载内存DC
                Call Del_dc
                Call DCload
                '调整图片框的宽度
                Me.picMain.Width = W * TwipsX
                Me.picMain.Height = H * F * TwipsY
                Me.ResizerPicture.HScrollMaximum = Me.picMain.Width
                Me.ResizerPicture.VScrollMaximum = Me.picMain.Height
                Me.tmrDataPerSecond.Enabled = True                      '激活流量统计计时器
                '一切就绪后发送屏幕传输请求
                wsSendData Me.wsMessage, "BEG"
                
            Case "SIF"                                      '系统信息
                sTempPara = Split(sTemp(i), "|")
                Me.edInformation.Text = sTempPara(1) & vbCrLf & vbCrLf & "以上信息仅供参考，一切请以实际为准！"
                '============================================
                '获取完系统信息后获取根目录
                If Me.lstFile.ListItems.Count = 0 Then
                    wsSendData Me.wsMessage, "GRT"
                End If
                
            Case "TSK"
                '-------------------------------------
                Me.lstTask.ListItems.Clear                  '清空列表
                sTempPara = Split(sTemp(i), vbCrLf)         '分割出每个回车间隔开的内容
                sTempPara(0) = Replace(sTempPara(0), "TSK|", "")
                For k = 0 To UBound(sTempPara) - 2            '获取每一行的内容
                    TaskSplit = Split(sTempPara(k), "|")    '分割出每个分隔符隔开的内容
                    Me.lstTask.ListItems.Add , , TaskSplit(0)                                       '进程名
                    Me.lstTask.ListItems(Me.lstTask.ListItems.Count).SubItems(1) = TaskSplit(1)     '进程PID
                    Me.lstTask.ListItems(Me.lstTask.ListItems.Count).SubItems(2) = TaskSplit(2)     '进程路径
                Next k
                Me.labTaskTotal.Caption = "共" & Me.lstTask.ListItems.Count + 1 & "个进程。"         '显示进程总数
            
            Case "KTS"                                      '结束进程的反馈信息
                Dim tmpString As String                     '显示消息缓存
                '-------------------------------------
                sTempPara = Split(Replace(Replace(sDataRec, "{S}", ""), "KTS|", ""), "||")  '去掉数据流分割符并进行分割
                tmpString = "结束进程反馈："
                For k = 0 To UBound(sTempPara) - 1                  '遍历分出来的内容
                    TaskSplit = Split(sTempPara(k), "|")
                    tmpString = tmpString & vbCrLf & "进程名：" & TaskSplit(0) & "（PID：" & TaskSplit(1) & _
                                            "） 【" & IIf(TaskSplit(2) = "T", "成功", "失败") & "】"
                Next k
                MsgBox tmpString, 64, "结束进程反馈"
                Call cmdRefreshTasks_Click                          '接收到反馈后再刷新
            
            Case "CMD"                                      '命令行执行的反馈消息
                Dim tmpResult As String                     '显示消息缓存
                '-------------------------------------
                tmpResult = Replace(Replace(sDataRec, "CMD|", ""), "{S}", "")
                Me.edCommandLine.Text = tmpResult
                Me.edCommandlineEnter.Enabled = True
                Me.edCommandlineEnter.Text = ""
                Me.edCommandlineEnter.SetFocus
                
            '===================================================================================================
            '===================================================================================================
            '文件管理反馈
            Case "GRT"                                      '获取根目录请求响应
                Dim sRootList() As String                       '列表项切割缓存
                Dim sRootMsg() As String                        '各列表项的详细信息切割缓存
                '-------------------------------------
                sRootList() = Split(Replace(sDataRec, "GRT|", ""), "||")    '切割字符串
                Me.lstFile.ListItems.Clear
                For k = 0 To UBound(sRootList) - 1
                    sRootMsg = Split(sRootList(k), "|")
                    Me.lstFile.ListItems.Add , , sRootMsg(0)
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(1) = "驱动器"
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(2) = sRootMsg(1)
                Next k
                Me.edPath.Text = "根目录"                       '显示路径名为“根目录”
            
            Case "GPH"                                      '获取目录下文件请求响应
                Dim sDirList() As String                        '列表项切割缓存
                Dim sDirMsg() As String                         '各列表项的详细信息切割缓存
                On Error Resume Next
                '-------------------------------------
                sDirList = Split(Replace(sDataRec, "GPH|", ""), "||")       '切割字符串
                Me.lstFile.ListItems.Clear
                For k = 0 To UBound(sDirList) - 1
                    sDirMsg = Split(sDirList(k), "|")
                    Me.lstFile.ListItems.Add , , sDirMsg(0)
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(1) = sDirMsg(1)
                    sDirMsg(2) = Replace(sDirMsg(2), ":N:", "")         '这个是防止文件夹获取不了大小导致的空白引发错误的
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(2) = sDirMsg(2)
                Next k
                wsSendData Me.wsMessage, "NPH"                  '发送获取当前目录路径的请求
            
            Case "MKD"
                Dim bSuccess As String
                bSuccess = Replace(Replace(sDataRec, "MKD|", ""), "{S}", "")
                If bSuccess = "0" Then
                    MsgBox "创建文件夹失败。", 64, "创建文件夹失败"
                Else
                    MsgBox "创建文件夹成功！", 64, "创建文件夹成功"
                End If
                wsSendData Me.wsMessage, "REF"                  '发送刷新请求
                
            Case "REN"
                Dim brSuccess As String
                brSuccess = Replace(Replace(sDataRec, "REN|", ""), "{S}", "")
                If brSuccess = "0" Then
                    MsgBox "文件重命名失败。", 64, "文件重命名失败"
                Else
                    MsgBox "文件重命名成功！", 64, "文件重命名成功"
                End If
                wsSendData Me.wsMessage, "REF"                  '发送刷新请求
            
            Case "NPH"
                Me.edPath.Text = Replace(Replace(sDataRec, "NPH|", ""), "{S}", "")      '显示返回的文件夹路径
            
            Case "VPH"
                MsgBox "目录名无效！", 48, "消息"                '显示错误消息
                wsSendData Me.wsMessage, "NPH"                  '发送获取当前目录路径的请求
            
        End Select
    Next i
    '-----------------------------
    sDataRec = ""                       '执行完所有命令后清除缓存
    bMultiData = False
End Sub

Private Sub wsPicture_Connect()
    frmMain.tmrTimeOut.Enabled = False          '连接成功后停止超时计时器
End Sub

Private Sub wsPicture_DataArrival(ByVal bytesTotal As Long)
    Dim dat() As Byte
    '统计数据量
    DataPerSecond = DataPerSecond + bytesTotal
    '处理数据
    If bytesTotal = 8 Then                                                      '[头] 对数据进行整理
        Me.wsPicture.GetData dat, vbArray Or vbByte
        xData = StrConv(dat, vbUnicode)
        J = Mid$(xData, 1, 1)
        If J = 9 Then J = 0
        xData = Mid$(xData, 2, 7) + 1
        If J <> 8 Then Me.wsPicture.SendData 0
    End If
    If bytesTotal = xData Then Call GetWindows
End Sub
