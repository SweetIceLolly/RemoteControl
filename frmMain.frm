VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "远程控制"
   ClientHeight    =   4596
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8424
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.8
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4596
   ScaleWidth      =   8424
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer tmrCalcPerSecond 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3120
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock wsFile 
      Left            =   7920
      Top             =   3960
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Timer tmrChangeIcon 
      Interval        =   10
      Left            =   7800
      Top             =   480
   End
   Begin VB.Timer tmrBlockInput 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2520
      Top             =   3360
   End
   Begin VB.ListBox lstPassword 
      Height          =   264
      ItemData        =   "frmMain.frx":0CCA
      Left            =   2400
      List            =   "frmMain.frx":0CCC
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Timer tmrForceRefresh 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1920
      Top             =   3360
   End
   Begin VB.DriveListBox Drive 
      Height          =   312
      Left            =   240
      TabIndex        =   20
      Top             =   4200
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.DirListBox Dir 
      Height          =   312
      Left            =   720
      TabIndex        =   19
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.FileListBox File 
      Height          =   288
      Hidden          =   -1  'True
      Left            =   1320
      Pattern         =   "*"
      System          =   -1  'True
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.ListBox lstClipboard 
      Height          =   264
      ItemData        =   "frmMain.frx":0CCE
      Left            =   1920
      List            =   "frmMain.frx":0CD0
      TabIndex        =   17
      Top             =   4200
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picRefresh 
      AutoRedraw      =   -1  'True
      DrawWidth       =   100
      Height          =   252
      Left            =   1320
      ScaleHeight     =   204
      ScaleWidth      =   324
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.ListBox lstTemp 
      Height          =   264
      ItemData        =   "frmMain.frx":0CD2
      Left            =   1920
      List            =   "frmMain.frx":0CD4
      TabIndex        =   16
      Top             =   3840
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.Timer tmrReturn 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1320
      Top             =   3360
   End
   Begin VB.Timer tmrTimeOut 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   3360
   End
   Begin VB.Timer tmrRetry 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock wsPic 
      Left            =   6960
      Top             =   3960
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00C86400&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   0
      ScaleHeight     =   732
      ScaleWidth      =   8412
      TabIndex        =   13
      Top             =   0
      Width           =   8415
      Begin VB.Image imgAbout 
         Height          =   360
         Left            =   8040
         Picture         =   "frmMain.frx":0CD6
         Stretch         =   -1  'True
         ToolTipText     =   "关于"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgSettings 
         Height          =   360
         Left            =   7620
         Picture         =   "frmMain.frx":19A0
         Stretch         =   -1  'True
         ToolTipText     =   "设置"
         Top             =   0
         Width           =   360
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "远程控制"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.8
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   216
         Index           =   5
         Left            =   720
         TabIndex        =   14
         Top             =   240
         Width           =   912
      End
      Begin VB.Image imgMainIcon 
         Height          =   500
         Left            =   120
         Picture         =   "frmMain.frx":266A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   500
      End
   End
   Begin XtremeSuiteControls.CheckBox chkAllow 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   3615
      _Version        =   786432
      _ExtentX        =   6376
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "允许远程控制"
      Appearance      =   4
      Value           =   1
   End
   Begin XtremeSuiteControls.FlatEdit edLocalIP 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   77
      Text            =   "666.666.666.666"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdConnect 
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   3720
      Width           =   2535
      _Version        =   786432
      _ExtentX        =   4471
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "点我连接咯！"
      ForeColor       =   16777215
      BackColor       =   16416003
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   3
      DrawFocusRect   =   0   'False
      ImageGap        =   1
   End
   Begin XtremeSuiteControls.RadioButton optControl 
      Height          =   372
      Left            =   4320
      TabIndex        =   1
      ToolTipText     =   "以远程控制方式连接对方"
      Top             =   2520
      Width           =   3852
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "远程控制模式"
      Appearance      =   4
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox comIP 
      Height          =   360
      Left            =   4320
      TabIndex        =   0
      Top             =   1800
      Width           =   3855
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   635
      _StockProps     =   77
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777215
      Appearance      =   6
      UseVisualStyle  =   -1  'True
      AutoComplete    =   -1  'True
      DropDownItemCount=   5
   End
   Begin XtremeSuiteControls.RadioButton optFile 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      ToolTipText     =   "以文件传输方式连接对方"
      Top             =   3000
      Width           =   3855
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "文件传输模式"
      Appearance      =   4
   End
   Begin XtremeSuiteControls.FlatEdit edPassword 
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   2160
      Width           =   2055
      _Version        =   786432
      _ExtentX        =   3625
      _ExtentY        =   661
      _StockProps     =   77
      Text            =   "123456"
      BackColor       =   16777215
      Locked          =   -1  'True
      Appearance      =   6
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      RightToLeft     =   -1  'True
   End
   Begin MSWinsockLib.Winsock wsMessage 
      Left            =   7440
      Top             =   3960
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.Image imgAbout1 
      Height          =   360
      Left            =   7680
      Picture         =   "frmMain.frx":3334
      Stretch         =   -1  'True
      Top             =   1080
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgSettings1 
      Height          =   384
      Left            =   6960
      Picture         =   "frmMain.frx":3FFE
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings2 
      Height          =   384
      Left            =   7080
      Picture         =   "frmMain.frx":4CC8
      Top             =   960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgSettings3 
      Height          =   384
      Left            =   7200
      Picture         =   "frmMain.frx":5992
      Top             =   840
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgAbout2 
      Height          =   384
      Left            =   7800
      Picture         =   "frmMain.frx":665C
      Top             =   960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgAbout3 
      Height          =   384
      Left            =   7920
      Picture         =   "frmMain.frx":7326
      Top             =   840
      Visible         =   0   'False
      Width           =   384
   End
   Begin XtremeSuiteControls.TrayIcon Tray 
      Left            =   3960
      Top             =   4080
      _Version        =   786432
      _ExtentX        =   339
      _ExtentY        =   339
      _StockProps     =   16
      Text            =   "远程控制"
      Picture         =   "frmMain.frx":7FF0
   End
   Begin VB.Image imgCopy 
      Height          =   360
      Index           =   1
      Left            =   3600
      Picture         =   "frmMain.frx":8CCA
      Stretch         =   -1  'True
      ToolTipText     =   "复制您的密码到剪贴板"
      Top             =   2160
      Width           =   360
   End
   Begin VB.Image imgCopy1 
      Height          =   360
      Left            =   2880
      Picture         =   "frmMain.frx":9434
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgCopy3 
      Height          =   288
      Left            =   3600
      Picture         =   "frmMain.frx":9B9E
      Top             =   3960
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgCopy2 
      Height          =   288
      Left            =   3240
      Picture         =   "frmMain.frx":A308
      Top             =   3960
      Visible         =   0   'False
      Width           =   288
   End
   Begin VB.Image imgCopy 
      Height          =   360
      Index           =   0
      Left            =   3600
      Picture         =   "frmMain.frx":AA72
      Stretch         =   -1  'True
      ToolTipText     =   "复制您的IP到剪贴板"
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label labState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "无连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   3960
      Width           =   540
   End
   Begin VB.Shape shpState 
      BorderColor     =   &H00FFFFC0&
      FillColor       =   &H0000FF00&
      FillStyle       =   0  'Solid
      Height          =   195
      Left            =   240
      Shape           =   3  'Circle
      Top             =   3960
      Width           =   195
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C000&
      Index           =   2
      X1              =   4080
      X2              =   0
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "您的IP"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   3
      Left            =   480
      TabIndex        =   10
      Top             =   1500
      Width           =   630
   End
   Begin VB.Image imgBack 
      Height          =   360
      Index           =   1
      Left            =   240
      Picture         =   "frmMain.frx":B1DC
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "随机密码"
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Index           =   4
      Left            =   360
      TabIndex        =   11
      Top             =   2205
      Width           =   840
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C000&
      Index           =   1
      X1              =   3960
      X2              =   120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line 
      BorderColor     =   &H00C0C000&
      Index           =   0
      X1              =   4080
      X2              =   4080
      Y1              =   840
      Y2              =   4560
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "允许远程控制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "对方IP地址"
      ForeColor       =   &H00404040&
      Height          =   210
      Index           =   2
      Left            =   4320
      TabIndex        =   8
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "控制远程电脑"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4320
      TabIndex        =   7
      Top             =   840
      Width           =   1890
   End
   Begin VB.Image imgBack 
      Height          =   360
      Index           =   2
      Left            =   240
      Picture         =   "frmMain.frx":B67C
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuShowWindow 
         Caption         =   "显示主窗体(&S)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
'    /\
'   /  \
'   |屠|
'   |B |
'   |U |
'   |G |
'   |宝|
'   |刀|　　　__
' ,_|  |_,　 /　)
'   (Oo　　/ _I_
'   +\ \　||· ·|
'  　 \ \||_ 0→
' 　　 \/.:.\- \
'   　　|.:. /-----\
'   　　|___|::oOo::|
'   　　/　 |:<_T_>:|
'       |   |::oOo::|
'       |    \-----/
'       |:_|__|
'       |===|==|
'       |   |  |
'       |&  \  \
'      *( ,  `'-.'-.
'       `"`"""""`""`
'
'       BUG挡路必死！

'================================================================
      
'图像端口：20234
'消息端口：20235
'文件端口：20236
'----------------------
'未连接：红色
'监听中：绿色
'有连接：蓝色
'正在连接：黄色
'----------------------
'通用数据流分割：{S}

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0                                                'color   table   in   RGBs

Private Type BITMAPINFOHEADER                                                   '40   bytes
    biSize   As Long
    biWidth   As Long
    biHeight   As Long
    biPlanes   As Integer
    biBitCount   As Integer
    biCompression   As Long
    biSizeImage   As Long
    biXPelsPerMeter   As Long
    biYPelsPerMeter   As Long
    biClrUsed   As Long
    biClrImportant   As Long
End Type

Private Type RGBQUAD
    rgbBlue   As Byte
    rgbGreen   As Byte
    rgbRed   As Byte
    rgbReserved   As Byte
End Type

Private Type BITMAPINFO
    bmiHeader   As BITMAPINFOHEADER
    bmiColors   As RGBQUAD
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef pointer As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursor Lib "user32" () As Long
Private Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function CompMemory Lib "ntdll.dll" Alias "RtlCompareMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'=========================================================================
'以下为zlib压缩及用到的函数
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Const OFFSET As Long = &H8

'键盘控制
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_KEYUP = &H2

'========
Dim W As Long, H As Long, M As Long
Dim x, y As Long, dwCurrentThreadID1 As Long
Dim J As Byte, Hei As Byte, g As Long
Dim T As Boolean
Dim bmpLen As Long, F As Byte
Dim CapName As String
Dim sdc As Long
Dim iBitmap1 As Long, iDC1 As Long, opeinter1 As Long
Dim iBitmap2 As Long, iDC2 As Long, opeinter2 As Long
Dim iBitmap3 As Long, iDC3 As Long, opeinter3 As Long
Dim iBitmap4 As Long, iDC4 As Long, opeinter4 As Long
Dim r1() As Byte
Dim XH As Boolean

'========================================================

Dim bMouseDown As Boolean                           '是否按下鼠标
Dim bTray As Boolean                                '当前是否处于托盘区状态

Public IsUpload As Boolean                          '对方是否为上传状态 【True:上传 False:下载】

Public recBytes As Long                             '接收到的文件数据字节数
Public oldBytes As Long                             '上一秒接收到的文件数据字节数
Public BytesPerSec As Long                          '每一秒的字节数
Public lSent As Long, oldSent As Long               '已发送的字节数 上一秒发送的字节数

Public FileRec As Integer                           '成功接收的文件数
Public FileSent As Integer                          '成功发送的文件数

Dim EachFile() As String                            '文件列表里的每个文件
Public FileList As String                           '待发送的文件列表
Public CurrentFile As Integer                       '当前文件的Index
Public TotalFiles As Integer                        '待发送的文件总数
Public CurrentSize As Long                          '当前打开文件的大小
Dim dSendTemp() As Byte                             '准备发送的文件

'========================================================================================================

Public Function NextFile() As Boolean           '下一文件过程
    CurrentFile = CurrentFile + 1
    If CurrentFile > UBound(EachFile) - 1 Then          '如果序号超出文件总数
        NextFile = False
        Exit Function
    End If
    '---------------------------
    Do While LoadFile(EachFile(CurrentFile)) = False    '如果读取文件错误就继续下一个文件
        CurrentFile = CurrentFile + 1                       '文件序号 + 1
        If CurrentFile > UBound(EachFile) - 1 Then          '如果序号超出文件总数
            NextFile = False
            Exit Function
        End If
    Loop
    NextFile = True
End Function

Public Function LoadFile(sFileName As String) As Boolean    '加载文件过程
    On Error Resume Next
    '-----------------------------------
    LoadFile = True
    Open sFileName For Binary As #1                             '打开文件
        If Err.Number <> 0 Then                                 '文件不存在
            Close #1                                            '关闭文件
            LoadFile = False
            Exit Function
        End If
        '--------------------
        ReDim dSendTemp(LOF(1))                                      '分配内存
        CurrentSize = LOF(1)
        If Err.Number <> 0 Then                                 '文件过大
            MsgBox "文件过大：" & sFileName
            LoadFile = False
            Exit Function
        End If
        If UBound(dSendTemp) = 0 Then                                '如果读取到的字节数为0
            Close #1                                            '关闭文件
            LoadFile = False
            Exit Function
        End If
        Get #1, , dSendTemp                                          '读取文件
    Close #1
    lSent = 0                                               '清空文件发送字节数
    oldSent = 0
End Function

Sub SendLoop()                                  '循环发送屏幕过程
    Do While XH                                     '不要看到Do就怕。。。对方接收完才会发送下一张滴~所以这个不是死循环~
        Call Send_data
        Sleep 1
        If bKeepSending = False Then
            Exit Do
        End If
    Loop
End Sub

Private Function MyGetCursor() As Long          '绘制鼠标位置
    Dim hWindow As Long, dwThreadID As Long, dwCurrentThreadID As Long
    Dim Pt As POINTAPI
    GetCursorPos Pt
    x = Pt.x - 6     '为什么要减去6呢？因为获取到的指针位置会有所偏移，减去6可以在一定程度上矫正位置，但是还是不准...
    y = Pt.y - 6 - g
    hWindow = WindowFromPoint(Pt.x, Pt.y)
    dwThreadID = GetWindowThreadProcessId(hWindow, 0)
    dwCurrentThreadID = GetCurrentThreadId
    If dwCurrentThreadID1 <> dwThreadID Then
        If AttachThreadInput(dwCurrentThreadID, dwCurrentThreadID1, False) Then
        End If
        dwCurrentThreadID1 = dwThreadID
        If AttachThreadInput(dwCurrentThreadID, dwCurrentThreadID1, True) Then
            MyGetCursor = GetCursor
        End If
    Else
        MyGetCursor = GetCursor
    End If
End Function

Private Sub Send_data()             '发送屏幕主过程
    Dim Bmp As BITMAP
    Dim hdc As Long, BD As Long
    Dim Tpicture As Boolean
    '===========
    Dim b As String
    Dim a() As Byte
    '=======
    If g = 0 Then BitBlt iDC3, 0, 0, W, H * F, sdc, 0, 0, vbSrcCopy             '3DC内容
    If g = 0 Then DrawIcon iDC3, x, y, MyGetCursor
    '！注3和1、2不能搞反
    '================================
    BitBlt iDC1, 0, 0, W, H, iDC3, 0, g, vbSrcCopy                              '1DC内容
    '================
    BitBlt iDC2, 0, 0, W, H, iDC4, 0, g, vbSrcCopy                              '2DC内容
    '================================
    
    '================================
    BD = CompMemory(ByVal opeinter1, ByVal opeinter2, bmpLen)                   '屏幕对比
    Tpicture = CBool(BD = bmpLen)
    '================================
    
    If Not Tpicture Then                                                        '如果屏幕内容发生了变化
        
        XH = False
        
        BitBlt iDC4, 0, g, W, H, iDC1, 0, 0, vbSrcCopy                          '4DC内容
        '===============
        BitBlt iDC1, 0, 0, W, H, iDC2, 0, 0, vbSrcInvert                        '扫描
        '===============
        ReDim r1(bmpLen) As Byte                                                '按图像数据实际的大小分配缓冲区
        CopyMemory r1(0), ByVal opeinter1, bmpLen
        If bCompress = True Then
            CompressByte r1                                                         '压缩
        End If
        '=======================                  [头]对数据进行整理
        b = Format(UBound(r1), "00000000")
        Mid$(b, 1, 1) = J
        If J = 0 Then Mid$(b, 1, 1) = 9
        a = StrConv(b, vbFromUnicode)                                           '字符串转换为字节数组
        Me.wsPic.SendData a
        '=====
    End If
    J = J + 1
    If J >= F Then J = 0
    g = J * H
End Sub

Private Sub SendDat()               '发送空图像让对方准备开始接收
    XH = True
    Me.wsPic.SendData r1
End Sub

Function CompressByte(ByteArray() As Byte)          '数据压缩过程
    Dim BufferSize As Long
    Dim TempBuffer() As Byte
    BufferSize = UBound(ByteArray) + 1
    BufferSize = BufferSize + (BufferSize * 0.01) + 12
    ReDim TempBuffer(BufferSize)
    CompressByte = (Compress(TempBuffer(0), BufferSize, ByteArray(0), UBound(ByteArray) + 1) = 0)
    Call CopyMemory(ByteArray(0), CLng(UBound(ByteArray) + 1), OFFSET)
    ReDim Preserve ByteArray(0 To BufferSize + OFFSET - 1)
    CopyMemory ByteArray(OFFSET), TempBuffer(0), BufferSize
End Function

Private Sub Del_dc()                                '释放内存DC过程
    ReleaseDC 0, sdc
    DeleteDC iDC2
    DeleteObject iBitmap2
    DeleteDC iDC1
    DeleteObject iBitmap1
    DeleteDC iDC3
    DeleteObject iBitmap3
    DeleteDC iDC4
    DeleteObject iBitmap4
End Sub

Private Sub sLoadf()                                '加载内存DC过程，这里是绑定位图操作
    Dim bi24BitInfo As BITMAPINFO
    With bi24BitInfo.bmiHeader                      '设置位图属性
        .biBitCount = 16
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = W
        .biHeight = H
    End With
    '=======                                        '绑定到位图
    iDC1 = CreateCompatibleDC(0)                                                '1
    iBitmap1 = CreateDIBSection(iDC1, bi24BitInfo, DIB_RGB_COLORS, opeinter1, ByVal 0&, ByVal 0&)
    SelectObject iDC1, iBitmap1
    '=======
    iDC2 = CreateCompatibleDC(0)                                                '2
    iBitmap2 = CreateDIBSection(iDC2, bi24BitInfo, DIB_RGB_COLORS, opeinter2, ByVal 0&, ByVal 0&)
    SelectObject iDC2, iBitmap2
    '=======
End Sub

Private Sub SLoad()                                 '同样是加载内存DC的过程，这里是加载操作
    Dim bi24BitInfo1 As BITMAPINFO
    Dim scrw, scrh As Single
    '=======
    sdc = GetDC(0)
    With bi24BitInfo1.bmiHeader
        .biBitCount = 16
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo1.bmiHeader)
        .biWidth = W
        .biHeight = H * F
    End With
    '=======
    iDC3 = CreateCompatibleDC(0)                                                '3内存dc
    iBitmap3 = CreateDIBSection(iDC3, bi24BitInfo1, DIB_RGB_COLORS, opeinter3, ByVal 0&, ByVal 0&)
    SelectObject iDC3, iBitmap3
    '=======
    '=======
    iDC4 = CreateCompatibleDC(0)                                                '4内存dc
    iBitmap4 = CreateDIBSection(iDC4, bi24BitInfo1, DIB_RGB_COLORS, opeinter4, ByVal 0&, ByVal 0&)
    SelectObject iDC4, iBitmap4
    '=======
    
    '=======
    Dim sdc1 As Long
    sdc1 = GetDC(0)
    BitBlt iDC4, 0, 0, W, H * F, sdc1, 0, 0, vbSrcCopy                          '4DC内容
    DrawIcon iDC4, x, y, MyGetCursor
    ReleaseDC 0, sdc1
    '======
    Call sLoadf                 '加载完后进行绑定操作
End Sub

Private Sub New_dc()                        '创建新位图处理
    '===========
    Dim b As String
    Dim a() As Byte
    '================
    ReDim r1(bmpLen * F) As Byte                                                '按图像数据实际的大小分配缓冲区
    CopyMemory r1(0), ByVal opeinter4, bmpLen * F
    '=======================                  [头]对数据进行整理
    If bCompress = True Then
        CompressByte r1                                                         '压缩
    End If
    b = Format(UBound(r1), "00000000")
    Mid$(b, 1, 1) = J
    If J = 0 Then Mid$(b, 1, 1) = 9
    a = StrConv(b, vbFromUnicode)                                               '字符串转换为字节数组
    Me.wsPic.SendData a
End Sub

'========================================================================

Private Sub cmdConnect_Click()
    Me.shpState.FillColor = vbYellow            '更改状态
    Me.labState.Caption = "正在连接..."
    If Me.optControl.Value = True Then                  '远控模式
        frmRemoteControl.wsPicture.Close
        frmRemoteControl.wsMessage.Close
        frmRemoteControl.wsPicture.Connect Me.comIP.Text, 20234         '尝试连接
        frmRemoteControl.wsMessage.Connect Me.comIP.Text, 20235
    Else                                                '文件传输模式
        frmFileTransfer.wsMessage.Close
        frmFileTransfer.wsMessage.Connect Me.comIP.Text, 20235          '尝试连接
    End If
    Me.tmrTimeOut.Enabled = True
End Sub

Private Sub comIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdConnect_Click           '按下回车键就相当于按按钮
    End If
End Sub

Private Sub edLocalIP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(0).Picture = Me.imgCopy1.Picture         '改图标~为了美观 (*^__^*)
    Me.imgCopy(1).Picture = Me.imgCopy1.Picture
End Sub

Private Sub edPassword_LostFocus()
    sUserPassword = Me.edPassword.Text                  '调整用户密码
    Call SaveConfig                                     '密码一旦编辑完成就保存配置文件
End Sub

Private Sub edPassword_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(0).Picture = Me.imgCopy1.Picture         '改图标~为了美观 (*^__^*)
    Me.imgCopy(1).Picture = Me.imgCopy1.Picture
End Sub

Private Sub Form_Load()
    On Error Resume Next
    '=======================================
    Call LoadConfig                             '读取配置文件
    If bTrayWhenStart Then                      '如果启动的时候最小化到托盘区
        Me.Tray.MinimizeToTray Me.hwnd
        bTray = True
    End If
    If Trim(LCase(Command)) = "/autostart" And bHideWhenStart Then      '如果是自动启动模式并且以后台模式运行
        Me.Tray.Icon = Nothing
        Me.Hide
        bTray = True
    End If
    '=======================================
    Me.Icon = Me.imgMainIcon.Picture            '更改主图标，为了更美观 (*^__^*)
    Me.edLocalIP.Text = Me.wsMessage.LocalIP    '显示本机IP
    '=======================================
    Me.wsMessage.Close                          '开始监听
    Me.wsPic.Close
    Me.wsMessage.Bind 20235
    Me.wsMessage.Listen
    Me.wsPic.Bind 20234
    Me.wsPic.Listen
    '若监听成功
    If Me.wsMessage.state = sckListening And Me.wsPic.state = sckListening And Err.Number = 0 Then
        Me.labState.Caption = "准备就绪"
        Me.shpState.FillColor = vbGreen
    Else
        Me.labState.Caption = "无连接"
        Me.shpState.FillColor = vbRed
        Me.tmrRetry.Enabled = True
    End If
    '=======================================
    '生成六位随机密码
    Dim tmpPsw As String
    For i = 1 To 6
        Randomize
        tmpPsw = tmpPsw & Chr(Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a")))
    Next i
    Me.edPassword.Text = tmpPsw
    '=======================================
    '获取屏幕传输需要的参数
    W = Screen.Width \ Screen.TwipsPerPixelX
    H = Screen.Height \ Screen.TwipsPerPixelY
    F = 6
    H = H \ F
    bmpLen = W * H * 2
    '----------------------
    '加载内存DC
    Call Del_dc
    Call SLoad
    '=======================================
    If Not bTrayWhenStart And Not bHideWhenStart Then
        Me.Show                                         '一切都默默加载完后再展现自我！
        Me.Refresh
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(0).Picture = Me.imgCopy1.Picture         '改图标~为了美观 (*^__^*)
    Me.imgCopy(1).Picture = Me.imgCopy1.Picture
    Me.imgSettings.Picture = Me.imgSettings1.Picture
    Me.imgAbout.Picture = Me.imgAbout1.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bExitWhenClose And Not bNoExit Then                  '如果设置了点击关闭按钮 且 不阻止程序关闭 就关闭就直接自杀
        End
    Else
        Me.Tray.MinimizeToTray Me.hwnd                      '最小化到托盘区
        Cancel = True                                       '取消关闭
        bTray = True
    End If
End Sub

Private Sub imgAbout_Click()
    frmAbout.Show                       '显示关于窗口
    Me.Enabled = False
End Sub

Private Sub imgCopy_Click(Index As Integer)
    Clipboard.Clear
    If Index = 0 Then                                   '将指定内容复制到剪贴板
        Clipboard.SetText Me.edLocalIP.Text
    Else
        Clipboard.SetText Me.edPassword.Text
    End If
End Sub

Private Sub imgCopy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(Index).Picture = Me.imgCopy3.Picture         '改图标~为了美观 (*^__^*)
    bMouseDown = True
End Sub

Private Sub imgCopy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown = False Then
        Me.imgCopy(Index).Picture = Me.imgCopy2.Picture     '改图标~为了美观 (*^__^*)
    End If
End Sub

Private Sub imgCopy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(Index).Picture = Me.imgCopy1.Picture         '改图标~为了美观 (*^__^*)
    bMouseDown = False
End Sub

Private Sub imgMainIcon_Click()
    MsgBox "I Love CXT～　(*^__^*)", 64, "(*^__^*) "
End Sub

Private Sub imgSettings_Click()
    '调整窗体状态
    With frmSettings
        .chkAutoRecord.Value = bAutoRecord
        .chkAutoResize.Value = AutoResize
        .chkAutoStart.Value = bAutoStart
        .chkHideMode.Value = bHideMode
        .chkTraytWhenStart.Value = bTrayWhenStart
        .chkUserPassword.Value = bUseUserPassword
        .chkHideWhenStart.Value = bHideWhenStart
        .chkAvoidExit.Value = bNoExit
        .optExit.Value = bExitWhenClose
        .optTray.Value = Not bExitWhenClose
        .edPassword1.Text = sUserPassword
        .edPassword2.Text = sUserPassword
        If bUseUserPassword = True Then
            .labTip(1).Enabled = True
            .labTip(2).Enabled = True
            .edPassword1.Enabled = True
            .edPassword2.Enabled = True
            .cmdChangePassword.Enabled = True
        End If
        .lstIP.Clear
        For i = 0 To Me.comIP.ListCount - 1         '依次添加IP
            .lstIP.AddItem Me.comIP.List(i)
        Next i
    End With
    '=========================================
    frmSettings.Show                    '显示设置窗口
    Me.Enabled = False
End Sub

Private Sub imgSettings_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgSettings.Picture = Me.imgSettings3.Picture
End Sub

Private Sub imgSettings_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgSettings.Picture = Me.imgSettings2.Picture
    Me.imgAbout.Picture = Me.imgAbout1.Picture
End Sub

Private Sub imgSettings_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgSettings.Picture = Me.imgSettings1.Picture
End Sub

Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgAbout.Picture = Me.imgAbout3.Picture
    Me.imgSettings.Picture = Me.imgSettings1.Picture
End Sub

Private Sub imgAbout_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgAbout.Picture = Me.imgAbout2.Picture
End Sub

Private Sub imgAbout_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgAbout.Picture = Me.imgAbout1.Picture
End Sub

Private Sub mnuExit_Click()
    a = MsgBox("真的要退出吗？", 32 + vbYesNo, "确认")          '退出确认
    If a = 6 Then
        End
    End If
End Sub

Private Sub mnuShowWindow_Click()
    If bTray = True Then
        Me.Tray.MaximizeFromTray Me.hwnd        '恢复窗体
        bTray = False
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgSettings.Picture = Me.imgSettings1.Picture
    Me.imgAbout.Picture = Me.imgAbout1.Picture
End Sub

Private Sub tmrBlockInput_Timer()
    '啊啊啊这里很危险！
    BlockInput True
End Sub

Private Sub tmrCalcPerSecond_Timer()
    If IsUpload Then                        '如果是从本机上传状态
        BytesPerSec = recBytes - oldBytes       '计算出每秒数据字节数
        oldBytes = recBytes                     '上一秒数据字节数等于这一秒的数据字节数`
    Else                                    '如果是从本机下载状态
        BytesPerSec = lSent - oldSent           '计算出每秒数据字节数
        oldSent = lSent                         '上一秒数据字节数等于这一秒的数据字节数`
    End If
    frmBeingControlled.Refresh              '刷新被控窗体
End Sub

Private Sub tmrChangeIcon_Timer()
    Dim Cur As POINTAPI
    GetCursorPos Cur
    '如果鼠标脱离了窗体就恢复所有图片的原样
    If Cur.x * Screen.TwipsPerPixelX < Me.Left Or _
       Cur.x * Screen.TwipsPerPixelX > Me.Left + Me.Width Or _
       Cur.y * Screen.TwipsPerPixelY < Me.Top Or _
       Cur.y * Screen.TwipsPerPixelY > Me.Top + Me.imgSettings.Top + Me.Height Then
        Call Form_MouseMove(0, 0, 0, 0)
    End If
End Sub

Private Sub tmrForceRefresh_Timer()
    '不停地往屏幕最左上角绘制大小为一像素的点，以触发屏幕刷新事件
    Me.picRefresh.PSet (0, 0), RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
    BitBlt sdc, 0, 0, 1, 1, Me.picRefresh.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub tmrRetry_Timer()
    On Error Resume Next
    Me.wsMessage.Close              '尝试监听
    Me.wsPic.Close
    Me.wsMessage.Bind 20235
    Me.wsPic.Bind 20234
    Me.wsMessage.Listen
    Me.wsPic.Listen
    '若监听成功
    If Me.wsMessage.state = sckListening And Me.wsPic.state = sckListening And Err.Number = 0 Then
        Me.labState.Caption = "准备就绪"
        Me.shpState.FillColor = vbGreen
        Me.tmrRetry.Enabled = False
    Else
        Me.labState.Caption = "无连接"
        Me.shpState.FillColor = vbRed
    End If
End Sub

Private Sub tmrReturn_Timer()               '超时后的状态显示恢复
    Me.shpState.FillColor = vbGreen
    Me.labState = "准备就绪"
    Me.tmrReturn.Enabled = False
End Sub

Private Sub tmrTimeOut_Timer()              '判断是否连接超时
    Me.shpState.FillColor = vbRed
    Me.labState.Caption = "连接失败"
    Me.tmrReturn.Enabled = True
    Me.tmrTimeOut.Enabled = False
End Sub

Private Sub Tray_DblClick()
    If bTray = True And Me.mnuShowWindow.Enabled = True Then
        Me.Tray.MaximizeFromTray Me.hwnd        '双击任务栏图标时还原窗体
        bTray = False
    End If
End Sub

Private Sub Tray_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnuPopup                   '弹出菜单~
    End If
End Sub

Private Sub wsFile_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim Temp() As Byte              '接收到的二进制数据
    Dim recStr As String            '接收到的字符串数据
    '-------------
    Dim sFileName As String         '写入文件路径
    Dim sFileTitle As String        '文件标题
    Dim lFileSize As Long           '文件大小
    Dim TitleSplitTmp() As String   '文件标题分割缓存
    '-------------
    Dim nFileName As String         '发送文件源路径
    Dim nFileTitle As String        '发送文件的标题
    Dim nTitleSplitTmp() As String  '文件标题分割缓存
    '----------------------------
    Me.wsFile.GetData Temp          '获取数据
    '--------------
    If IsUpload Then                '如果是上传模式
        '下面从我的Tag里读取信息啦~
        sFileName = Split(Me.wsFile.Tag, "|")(0)                                                            '分离出源文件路径
        TitleSplitTmp = Split(sFileName, "\")                                                               '按照“\”来分离
        sFileTitle = TitleSplitTmp(UBound(TitleSplitTmp))                                                   '得到文件标题
        lFileSize = CLng(Split(Me.wsFile.Tag, "|")(1))                                                      '分离出文件大小
        sFileName = Split(Me.wsFile.Tag, "|")(2)                                                            '分离出写入文件路径
        sFileName = IIf(Right(sFileName, 1) = "\", sFileName & sFileTitle, sFileName & "\" & sFileTitle)    '得到写入文件的完整路径
        '--------------
        '获取尾部数据查看是否有结束标记
        '获得二进制数据的尾部三位数据
        Select Case Temp(UBound(Temp) - 2) & "|" & Temp(UBound(Temp) - 1) & "|" & Temp(UBound(Temp))    '看看是不是命令
            Case "70|73|78"                         '【FIN】发送完成消息
                Dim lstData() As Byte                       '结束标记前的数据
                frmBeingControlled.ProgressBar.Value = 100  '进度条调到满处
                If UBound(Temp) - 2 > 0 Then            '如果不是刚刚好到结束标记
                    ReDim lstData(UBound(Temp) - 2)         '分配比接收到数据少3字节的内存
                    For i = 0 To UBound(Temp) - 2           '将数据复制过去
                        lstData(i) = Temp(i)
                    Next i
                    Put #2, LOF(2) + 1, lstData             '写入文件
                End If
                FileRec = FileRec + 1                   '接收文件数 + 1
                Close #2                                '关闭文件
                Me.wsFile.SendData "NXT"                '请求发送下一文件
                Exit Sub
        
            Case "67|78|84"                         '【CNT】如果是连接建立确认消息
                oldBytes = 0                                                '初始化接收到的字节数
                recBytes = 0
                With frmBeingControlled                                     '更新被控窗体的状态
                    .edStatus1.Text = "对方正在上传文件到您的电脑"
                    .edStatus2.Text = "文件名：" & sFileTitle
                    .edStatus3.Text = "0 Byte/s 0 Byte/" & SizeWithFormat(lFileSize) & " 0%"
                    .edStatus2.Visible = True
                    .edStatus3.Visible = True
                    .Refresh
                End With
                '-----------------------
                '判断文件是否重名
                If IsPathExists(sFileName, True) = True Then                '检测到已存在文件
                    Me.wsFile.SendData "TSN"                         '发送覆盖重名文件确认
                    Exit Sub
                End If
                '如果没有重名则发送准备就绪消息
                Open sFileName For Binary As #2         '创建文件
                Me.wsFile.SendData "IRD"
                Exit Sub                                '如果是命令语句就退出过程，不写入文件
    
            Case "89|69|83"                         '【YES】覆盖同名文件消息
                Kill sFileName                          '删除掉同名的文件
                If Err.Number <> 0 Then                 '如果文件无法删除
                    Me.wsFile.SendData "DFL"     '告诉对方无法覆盖耶。。。
                    Exit Sub
                End If
                Open sFileName For Binary As #2         '创建文件
                recBytes = 0                            '清空接收到的字节数
                Me.wsFile.SendData "IRD"         '发送准备就绪消息
                Exit Sub                                '退出过程，不写入文件
    
            Case "78|82|70"                         '【NRF】取消覆盖同名文件消息
                Me.wsFile.SendData "NXT"         '请求发送下一文件
                Exit Sub                                '退出过程，不写入文件
    
        End Select
        Put #2, LOF(2) + 1, Temp                '写入文件
        recBytes = recBytes + UBound(Temp)      '接收到的字节数
        '更新标签和滚动条
        frmBeingControlled.ProgressBar.Value = (recBytes / lFileSize) * 100
        frmBeingControlled.edStatus3.Text = SizeWithFormat(BytesPerSec) & "/s " & SizeWithFormat(recBytes) & "/" & _
                                            SizeWithFormat(lFileSize) & " " & Format(recBytes / lFileSize * 100, "0.00") & "%"
    Else                            '如果是下载模式
        '下面从各种变量里读取信息啦~
        nFileName = EachFile(CurrentFile)                       '得到发送文件的完整路径
        nTitleSplitTmp = Split(nFileName, "\")                  '按照“\”分离
        nFileTitle = nTitleSplitTmp(UBound(nTitleSplitTmp))     '得到文件标题
        '--------------
        '获取尾部数据查看是否有结束标记
        '获得二进制数据的尾部三位数据
        Select Case Temp(UBound(Temp) - 2) & "|" & Temp(UBound(Temp) - 1) & "|" & Temp(UBound(Temp))    '看看是不是命令
            Case "68|78|84"                     '【DNT】连接建立确认
                lSent = 0                                       '清空发送的字节数
                oldSent = 0
                With frmBeingControlled                                     '更新被控窗体的状态
                    .edStatus1.Text = "对方正在从您的电脑下载文件"
                    .edStatus2.Text = "文件名：" & nFileTitle
                    .edStatus3.Text = "0 Byte/s 0 Byte/" & SizeWithFormat(CurrentSize) & " 0%"
                    .edStatus2.Visible = True
                    .edStatus3.Visible = True
                    .Refresh
                End With
                '-----------------------
                '【请求头|文件标题|文件大小|当前文件序号|文件总数】
                Me.wsFile.SendData "DNT|" & nFileTitle & "|" & CurrentSize & "|" & CurrentFile & "|" & UBound(EachFile)
                Exit Sub
            
            Case "78|88|84"                     '【NXT】下一文件
                If NextFile = True Then
                    '如果成功打开下一个文件就发送上传下一个文件请求
                    sFileName = Split(EachFile(CurrentFile), "|")(0)                            '分离出源文件路径
                    TitleSplitTmp = Split(sFileName, "\")                               '按照“\”来分离
                    sFileTitle = TitleSplitTmp(UBound(TitleSplitTmp))                   '得到文件标题
                    lSent = 0                                                           '清空发送的字节数
                    oldSent = 0
                    With frmBeingControlled                                             '更新被控窗体的状态
                        .edStatus2.Text = "文件名：" & sFileTitle
                        .edStatus3.Text = "0 Byte/s 0 Byte/" & SizeWithFormat(CurrentSize) & " 0%"
                        .Refresh
                    End With
                    '【请求头|文件标题|文件大小|当前文件序号|文件总数】
                    Me.wsFile.SendData "DNT|" & sFileTitle & "|" & CurrentSize & "|" & CurrentFile & "|" & UBound(EachFile)
                Else
                    Close #2                                                '关闭文件，防止有文件忘记关闭
                    With frmBeingControlled
                        .edStatus1.Text = "任务完成"
                        .edStatus2.Text = "共发送" & FileSent & "个文件"
                        .edStatus3.Visible = False
                    End With
                    Me.tmrCalcPerSecond.Enabled = False                     '关闭计算每秒字节数的计时器
                    Me.wsFile.SendData "END|" & CStr(FileSent)                          '发送下载完成消息
                End If
            
            Case "73|82|68"                     '【IRD】准备就绪
                Me.tmrCalcPerSecond.Enabled = True                                  '启动计时器，计算每秒发送字节数
                Me.wsFile.SendData dSendTemp                                        '给我狠狠的砸数据！！ (╯‵□′)╯︵┻━┻
                FileSent = FileSent + 1
            
        End Select
    End If
End Sub

Private Sub wsFile_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If Not IsUpload Then                                                            '如果是下载状态
        lSent = lSent + bytesSent
        If lSent >= CurrentSize Then                                                '如果已发送数据大于文件大小
            lSent = 0
            Me.wsFile.SendData "FIN"                                                    '告诉对方发送完成
            Exit Sub
        End If
        '=======================================================
        '更新被控窗体状态
        Dim tmpShow As String                                                       '显示的文本缓存
        tmpShow = SizeWithFormat(BytesPerSec) & "/s"                                '每秒字节数 = 现在发送的数据字节数 - 上一秒发送的数据字节数
        tmpShow = tmpShow & " " & SizeWithFormat(lSent)                             '已经发送的大小
        tmpShow = tmpShow & "/" & SizeWithFormat(CurrentSize)                       '文件的大小
        tmpShow = tmpShow & " " & Format(lSent / CurrentSize * 100, "0.00") & "%"   '计算百分比
        frmBeingControlled.ProgressBar.Value = lSent / CurrentSize * 100            '滚动条的值
        frmBeingControlled.edStatus3.Text = tmpShow                                 '显示到标签上
    End If
End Sub

Public Sub wsMessage_Close()
    Dim OpenPath As String                  '程序自身的路径
    OpenPath = IIf(Right(App.Path, 1) = "\", App.Path & App.EXEName & ".exe", App.Path & "\" & App.EXEName & ".exe")
    Shell OpenPath, vbNormalFocus               '沉舟侧畔千帆过 病树前头万木春
    End                                         '我老了，不中用了，该上路了。
End Sub

Private Sub wsMessage_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim Temp As String                      '缓存数据
    Dim sTemp() As String                   '数据切割缓存数据
    Dim sTempPara() As String               '数据附带的参数切割出来的缓存数据
    Dim Cur As POINTAPI                     '鼠标坐标对象（用于鼠标控制部分）
    '====================================
    Me.wsMessage.GetData Temp               '获取数据
    sTemp = Split(Replace(Decrypt(Temp), "{END}", ""), "{S}")            '切割数据
    For i = 0 To UBound(sTemp)
        Select Case Left(sTemp(i), 3)       '分析数据
            '===================================================================================================
            '===================================================================================================
            '杂项
            Case "CNT"                                     '连接请求
                Me.labState.Caption = "控制请求: " & Me.wsMessage.RemoteHostIP    '更改状态
                Me.shpState.FillColor = vbYellow
                If Me.chkAllow.Value = xtpUnchecked Then        '当前不允许连接
                    wsSendData Me.wsMessage, "NAC"                    '发送拒绝连接消息
                    Exit Sub
                End If
                If InStr(sTemp(i), "|") <> 0 Then               '如果是带密码参数的话
                    sTempPara = Split(sTemp(i), "|")
                    If (sTempPara(1) = Me.edPassword.Text) Or _
                       (sTempPara(1) = sUserPassword And bUseUserPassword) Then
                    '如果与随机密码相符 或者 使用用户密码且与用户密码相符 即 密码正确
                        Me.Hide                                                                 '隐藏主窗体
                        Me.labState.Caption = "当前连接:来自 " & Me.wsMessage.RemoteHostIP      '更改状态
                        Me.shpState.FillColor = vbBlue
                        If Me.wsPic.state <> sckConnected Then              '如果图片端口未连接则是文件传输模式
                            IsRemoteControl = False
                        Else                                                '否则就是远控模式
                            IsRemoteControl = True
                        End If
                        If Not bHideMode And Not bHideWhenStart Then        '如果不是隐藏状态
                            frmBeingControlled.Show                             '显示被控窗口
                            frmBeingControlled.Left = Screen.Width - frmBeingControlled.Width
                            frmBeingControlled.Top = Screen.Height - frmBeingControlled.Height * 2
                            frmBeingControlled.labMsg.Caption = "来自：" & Me.wsMessage.RemoteHostIP
                        End If
                        '--------------------------------------
                        If IsRemoteControl Then                                 '如果是远程控制模式
                            '发送本机分辨率
                                wsSendData Me.wsMessage, "RES|" & CStr(Screen.Width / Screen.TwipsPerPixelX) & "|" & _
                                                                  CStr(Screen.Height / Screen.TwipsPerPixelY) & "|" & _
                                                                  CStr(Screen.TwipsPerPixelX) & "|" & _
                                                                  CStr(Screen.TwipsPerPixelY)
                            Else                                                '如果是文件传输状态
                                frmBeingControlled.labMode.Caption = "当前为文件传输模式,对方可以管理您的文件,并可以上传下载文件。"
                                frmBeingControlled.labMsg.Caption = "来自：" & Me.wsMessage.RemoteHostIP
                                '-------------------------------------
                                Dim tmpStringToSend As String                       '要发送的缓存字符串
                                Call MakeRootList                                   '生成根目录表
                                tmpStringToSend = "GRT|"                            '数据头
                                For k = 0 To Me.lstTemp.ListCount - 1               '遍历列表框
                                    tmpStringToSend = tmpStringToSend & Me.lstTemp.List(k) & "||"
                                Next k
                                wsSendData Me.wsMessage, tmpStringToSend            '发送数据
                        End If
                        '发送连接确认
                        wsSendData Me.wsMessage, "CNT|True"
                        
                    Else                                       '如果密码错误的话
                        wsSendData Me.wsMessage, "CNT|Wrong"                  '拒绝连接
                    End If
                Else                                         '如果没有密码参数的话
                    wsSendData Me.wsMessage, "CNT|False"
                End If
                
            Case "BEG"                                       '客户端说他已经准备好啦！！那我就开始砸图片啦~
                Call New_dc                 '创建内存DC并启动屏幕传输
                XH = True
                
            Case "SIF"                      '发送系统信息
                '获取系统信息
                Dim tempString As String       '缓存字符串
                Set wmi = GetObject("winmgmts:\\.\root\CIMV2")      '调用WMI获取系统硬件信息
                '===============================================================
                Set oWMINameSpace = GetObject("winmgmts:")                            '获取版本
                Set SystemSet = oWMINameSpace.InstancesOf("Win32_OperatingSystem")
                For Each System In SystemSet
                    tempString = "系统版本：" & System.Caption & vbCrLf & vbCrLf
                Next
                '----------------------------
                Set Msg = wmi.ExecQuery("select * from win32_processor")              '获取CPU信息
                For Each k In Msg
                    tempString = tempString & "CPU名称：" & k.Name & vbCrLf
                Next
                tempString = tempString & vbCrLf
                '----------------------------
                Set Msg = wmi.ExecQuery("select * from win32_ComputerSystem")         '获取内存信息
                tempString = tempString & "内存大小"
                For Each k In Msg
                    tempString = tempString & vbCrLf & CStr(Format(k.TotalPhysicalMemory / 1024 ^ 2, "0.00")) & "MB"
                Next
                tempString = tempString & vbCrLf
                '----------------------------
                Dim n As Integer
                Set Msg = wmi.ExecQuery("select * from win32_DiskDrive")              '获取硬盘大小信息
                tempString = tempString & vbCrLf & "硬盘大小" & vbCrLf
                For Each k In Msg
                    n = n + 1
                    tempString = tempString & "硬盘" & n & "：" & CStr(Format((k.Size / 1024 ^ 3), "0.00")) & "GB" & vbCrLf
                Next
                '----------------------------                                       '获取各分区大小信息
                Set Msg = wmi.ExecQuery("select * from win32_LogicalDisk where DriveType='3'")
                For Each k In Msg
                    tempString = tempString & vbCrLf & "盘符：" & k.DeviceID & "  大小：" & CStr(Format((k.Size / 1024 ^ 3), "0.00")) & "GB"
                Next
                '----------------------------
                tempString = tempString & vbCrLf & vbCrLf
                Set Msg = wmi.ExecQuery("select * from win32_VideoController")      '获取显卡信息
                For Each k In Msg
                    tempString = tempString & "显卡型号：" & k.Name & "  显存：" & CStr(Format(Abs(k.AdapterRAM / 1024 ^ 2), "0.00")) & "MB" & vbCrLf
                Next
                '===============================================================
                wsSendData Me.wsMessage, "SIF|" & tempString
            
            Case "VBS"                      '执行VBS请求
                Dim tmpFile As String                                   '文件内容缓存
                '------------------------------
                tmpFile = Split(sTemp(i), "VBS|")(1)
                Open App.Path & "\temp.vbs" For Output As #1            '保存文件
                    Print #1, tmpFile
                Close #1
                '------------------------------
                Shell "wscript.exe " & Chr(34) & App.Path & "\temp.vbs" & Chr(34)          '运行VBS
            
            Case "CMD"                      '执行命令行请求
                Dim objCMD As Object                                    'CMD管道对象
                Dim tmpCMD As String                                    '命令内容
                Dim SendTemp As String                                  '反馈内容
                '------------------------------
                Set objCMD = New clsCMD                                 '建立CMD管道
                tmpCMD = Split(sTemp(i), "CMD|")(1)                         '分离出请求的命令
                Call objCMD.DosInput(tmpCMD)                                '执行CMD命令！
                SendTemp = objCMD.DosOutPutEx(10000)                        '执行获取结果！
                Set objCMD = Nothing                                    '关闭CMD管道
                wsSendData Me.wsMessage, "CMD|" & SendTemp              '发送执行结果
            
            '===================================================================================================
            '===================================================================================================
            '进程管理
            Case "TSK"                      '发送当前系统的进程信息
                '获取所有进程
                Dim myProcess As PROCESSENTRY32         '进程类型
                Dim mySnapshot As Long                  '快照句柄
                Dim ProcessPath As String               '进程路径
                Dim tmpString As String                 '缓存字符串
                '------------------------------
                Me.lstTemp.Clear                        '清空缓存列表框
                myProcess.dwSize = Len(myProcess)
                '创建进程快照
                mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
                '获取第一个进程名
                ProcessFirst mySnapshot, myProcess
                Me.lstTemp.AddItem Trim(myProcess.szexeFile)        '加入进程镜像名
                '加入进程PID和位置
                Me.lstTemp.List(Me.lstTemp.ListCount - 1) = Me.lstTemp.List(Me.lstTemp.ListCount - 1) & "|" & myProcess.th32ProcessID & "|" & GetFileName(myProcess.th32ProcessID)
                '添加所有进程
                While ProcessNext(mySnapshot, myProcess)
                    Me.lstTemp.AddItem Trim(myProcess.szexeFile)        '加入进程镜像名
                    '加入进程PID和位置
                    Me.lstTemp.List(Me.lstTemp.ListCount - 1) = Me.lstTemp.List(Me.lstTemp.ListCount - 1) & "|" & myProcess.th32ProcessID & "|" & GetFileName(myProcess.th32ProcessID)
                Wend
                '------------------------------
                For k = 0 To Me.lstTemp.ListCount
                    tmpString = tmpString & Me.lstTemp.List(k) & vbCrLf
                Next k
                wsSendData Me.wsMessage, "TSK|" & tmpString             '发送所有进程信息
                
            Case "KTS"                      '结束进程请求
                Dim tmpTask() As String                                 '字符串分割缓存变量
                Dim tmpSend As String                                   '发送反馈的缓存变量
                '------------------------------
                sTempPara = Split(sTemp(i), vbCrLf)                     '分割出每个回车间的内容
                For k = 0 To UBound(sTempPara) - 1
                    tmpTask = Split(Replace(sTempPara(k), "KTS|", ""), "|")
                    If KillPID(CLng(tmpTask(1))) = True Then    '杀杀杀！！
                        '成功！
                        tmpSend = tmpSend & tmpTask(0) & "|" & tmpTask(1) & "|" & "T||"
                    Else
                        '失败！
                        tmpSend = tmpSend & tmpTask(0) & "|" & tmpTask(1) & "|" & "F||"
                    End If
                Next k
                wsSendData Me.wsMessage, "KTS|" & tmpSend               '发送反馈信息
                
            '===================================================================================================
            '===================================================================================================
            '鼠标事件
            Case "MMP"                      '更改鼠标坐标
                sTempPara = Split(sTemp(i), "|")
                SetCursorPos CLng(sTempPara(1)), CLng(sTempPara(2))
            
            Case "MLD"                      '左下
                GetCursorPos Cur
                mouse_event LeftDown, Cur.x, Cur.y, 0, 0
            
            Case "MLU"                      '左上
                GetCursorPos Cur
                mouse_event LeftUp, Cur.x, Cur.y, 0, 0
            
            Case "MRD"                      '右下
                GetCursorPos Cur
                mouse_event RightDown, Cur.x, Cur.y, 0, 0
            
            Case "MRU"                      '右上
                GetCursorPos Cur
                mouse_event RightUp, Cur.x, Cur.y, 0, 0
            
            Case "MMD"                      '中下
                GetCursorPos Cur
                mouse_event MiddleDown, Cur.x, Cur.y, 0, 0
            
            Case "MMU"                      '中上
                GetCursorPos Cur
                mouse_event MiddleUp, Cur.x, Cur.y, 0, 0
                
            Case "DBC"                      '左键双击
                GetCursorPos Cur                                        '为什么这里只发送一次呢？因为双击的时候就包括了单击事件，
                mouse_event LeftDown Or LeftUp, Cur.x, Cur.y, 0, 0      '而双击的时候再发送两次的话就会发送三次按键了。
            
            Case "MWU"                      '滚轮向上
                mouse_event MOUSEEVENTF_WHEEL, 0, 0, 150, 0
            
            Case "MWD"                      '滚轮向下
                mouse_event MOUSEEVENTF_WHEEL, 0, 0, -150, 0
                
            '===================================================================================================
            '===================================================================================================
            '键盘事件
            Case "KBD"                      '键盘按下按键
                keybd_event Split(sTemp(i), "|")(1), 0, 0, 0
            
            Case "KBU"                      '键盘松开按键
                keybd_event Split(sTemp(i), "|")(1), 0, KEYEVENTF_KEYUP, 0
            
            '===================================================================================================
            '===================================================================================================
            '文件管理
            Case "GRT"                      '获取驱动器
                Dim tmpSendString As String                     '要发送的缓存字符串
                '-------------------------------------
                Call MakeRootList                               '生成根目录表
                tmpSendString = "GRT|"                          '数据头
                For k = 0 To Me.lstTemp.ListCount - 1           '遍历列表框
                    tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                Next k
                wsSendData Me.wsMessage, tmpSendString          '发送数据
             
            Case "GPH"                      '获取目录下的文件（夹）
                Dim gPath As String               '对方请求的目录名称缓存
                '-------------------------------------
                sTempPara = Split(sTemp(i), "|")
                gPath = Me.lstTemp.List(CInt(sTempPara(1) - 1))
                If InStr(gPath, "[System") <> 0 Then
                    Call MakeRootList
                    gPath = Me.lstTemp.List(CInt(sTempPara(1) - 1))
                End If
                If InStr(gPath, "空闲:") <> 0 Then  '请求目录为驱动器
                    MakeList Left(gPath, 2) & "\"           '生成目录表
                End If
                '-------------------------------------
                If InStr(gPath, "目录|") <> 0 Then  '请求目录为普通目录
                    '规范目录名
                    gPath = Split(gPath, "|")(0)
                    gPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path & gPath, Me.File.Path & "\" & gPath)
                    MakeList gPath                          '生成目录表
                End If
                '-------------------------------------
                tmpSendString = "GPH|"                          '数据头
                For k = 0 To Me.lstTemp.ListCount - 1           '遍历列表框
                    tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                Next k
                '==========
                wsSendData Me.wsMessage, tmpSendString          '发送数据
                
                
            Case "GUD"                      '获取上层目录的文件（夹）
                Dim sPath As String                 '目录缓存
                Dim tmpSplitPath() As String        '切割目录路径缓存
                '-------------------------------------
                If Right(Me.Dir.Path, 2) <> ":\" Then               '如果当前不是从根目录往上一级的话
                    tmpSplitPath = Split(Me.File.Path, "\")             '显示上一级目录
                    For k = 0 To UBound(tmpSplitPath) - 1
                        sPath = sPath & tmpSplitPath(k) & "\"
                    Next k
                    Call MakeList(sPath)                            '生成目录列表
                    '-------------------------------------
                    tmpSendString = "GPH|"                          '数据头
                    For k = 0 To Me.lstTemp.ListCount - 1           '遍历列表框
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '发送数据
                Else
                    Call MakeRootList                               '生成磁盘根目录列表
                    '-------------------------------------
                    tmpSendString = "GRT|"                          '数据头
                    For k = 0 To Me.lstTemp.ListCount - 1           '遍历列表框
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '发送数据
                End If
            
            Case "MKD"                      '新建文件夹
                Dim tmpPath As String                               '文件夹路径缓存
                sTempPara = Split(sTemp(i), "|")                    '分割出参数
                tmpPath = IIf(Right(Me.Dir.Path, 1) = "\", Me.Dir.Path, Me.Dir.Path & "\") & sTempPara(1)
                MkDir tmpPath
                If Err.Number = 0 Then                              '让我看看有没有出错捏~
                    wsSendData Me.wsMessage, "MKD|1"                    '发送成功反馈
                Else
                    wsSendData Me.wsMessage, "MKD|0"                    '发送失败反馈
                End If
            
            Case "REF"                      '刷新
                Me.Dir.Refresh
                Me.File.Refresh
                Me.Drive.Refresh
                '------------------------
                If InStr(Me.lstTemp.List(0), "空闲:") <> 0 Then     '如果是根目录列表
                    Call MakeRootList                               '生成磁盘根目录列表
                    '-------------------------------------
                    tmpSendString = "GRT|"                          '数据头
                    For k = 0 To Me.lstTemp.ListCount - 1           '遍历列表框
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '发送数据
                Else
                    Call MakeList(sPath)                            '生成目录列表
                    '-------------------------------------
                    tmpSendString = "GPH|"                          '数据头
                    For k = 0 To Me.lstTemp.ListCount - 1           '遍历列表框
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '发送数据
                End If
            
            Case "REN"                      '重命名
                Dim OldName As String                               '旧名字
                Dim NewName As String                               '新名字
                '------------------------
                sTempPara = Split(sTemp(i), "|")                    '分割出参数
                OldName = IIf(Right(Me.Dir.Path, 1) = "\", Me.Dir.Path, Me.Dir.Path & "\")  '生成缓存字符串
                NewName = OldName
                OldName = OldName & Split(Me.lstTemp.List(CInt(sTempPara(1)) - 1), "|")(0)
                NewName = NewName & sTempPara(2)
                Name OldName As NewName                             '重命名
                If Err.Number = 0 Then                              '让我看看有没有出错捏~
                    wsSendData Me.wsMessage, "REN|1"                    '发送成功反馈
                Else
                    wsSendData Me.wsMessage, "REN|0"                    '发送失败反馈
                End If
            
            Case "CPY"                      '复制
                Dim sIndex() As String                              '分割出来的Index缓存
                Dim cPath As String                                 '选择的文件的路径
                '------------------------
                Me.lstClipboard.Clear                               '清空“剪贴板”
                isCopy = True                                       '复制模式
                sIndex = Split(Replace(sTemp(i), "CPY|", ""), "|")
                cPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                For k = 0 To UBound(sIndex) - 1
                    Me.lstClipboard.AddItem cPath & Me.lstTemp.List(sIndex(k) - 1)    '添加到“剪贴板”里
                Next k
            
            Case "CUT"                      '剪切
                Dim uIndex() As String                              '分割出来的Index缓存
                Dim uPath As String                                 '选择的文件的路径
                '------------------------
                Me.lstClipboard.Clear                               '清空“剪贴板”
                isCopy = False                                      '剪切模式
                uIndex = Split(Replace(sTemp(i), "CPY|", ""), "|")
                uPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                For k = 0 To UBound(uIndex) - 1
                    Me.lstClipboard.AddItem uPath & Me.lstTemp.List(uIndex(k) - 1)    '添加到“剪贴板”里
                Next k
            
            Case "PST"                      '粘贴
                Dim tmpList As String                               '列表项里的内容缓存
                Dim tmpFilePath As String                           '“剪贴板”中的每个文件的路径
                Dim tmpFileTitle() As String                        '文件的名称
                Dim tmpTargetPath As String                         '目标路径
                Dim tmpTargetDirPath As String                      '目标所在的文件夹路径
                '------------------------
                tmpTargetDirPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                For k = 0 To Me.lstClipboard.ListCount - 1
                    tmpList = Me.lstClipboard.List(k)
                    tmpFilePath = Split(tmpList, "|")(0)                                '获取源文件路径
                    tmpFileTitle = Split(tmpFilePath, "\")                                              '分割出文件名
                    tmpTargetPath = tmpTargetDirPath & tmpFileTitle(UBound(tmpFileTitle))                     '生成目标路径
                    '======================================================
                    If isCopy = True Then                       '复制模式
                        If InStr(tmpList, "|目录") <> 0 Then            '目录就用xcopy复制
                            Shell "xcopy " & Chr(34) & tmpFilePath & Chr(34) & " " & Chr(34) & tmpTargetPath & Chr(34) & " /e /c /i /h /y", vbHide
                        End If
                        If InStr(tmpList, "|文件") <> 0 Then            '文件就用FileCopy复制
                            FileCopy tmpFilePath, tmpTargetPath
                        End If
                    Else                                        '剪切模式
                        Shell "cmd /c move /y " & Chr(34) & tmpFilePath & Chr(34) & " " & Chr(34) & tmpTargetPath & Chr(34), vbHide
                    End If
                Next k
                
            Case "DEL"                      '删除
                Dim dIndex() As String                              '分割出来的Index缓存
                Dim dPath As String                                 '选择的文件的路径
                Dim dList As String                                 '要删除的文件路径缓存
                '------------------------                                '剪切模式
                dIndex = Split(Replace(sTemp(i), "DEL|", ""), "|")
                dPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")     '规范化路径名
                For k = 0 To UBound(dIndex) - 1
                    dList = Me.lstTemp.List(dIndex(k) - 1)
                    If InStr(dList, "|目录") <> 0 Then         '如果是目录
                        If bDeleteFiles Then
                            Shell "cmd /c rd " & Chr(34) & dPath & Split(dList, "|")(0) & Chr(34) & " /s /q", vbHide
                        Else
                            MsgBox "cmd /c rd " & Chr(34) & dPath & Split(dList, "|")(0) & Chr(34) & " /s /q"
                        End If
                    End If
                    If InStr(dList, "|文件") <> 0 Then         '如果是文件
                        If bDeleteFiles Then
                            Kill dPath & Split(dList, "|")(0)
                        Else
                            MsgBox dPath & Split(dList, "|")(0)
                        End If
                    End If
                Next k
            
            Case "NPH"                      '获取当前浏览的路径
                Dim sNowPath As String      '要发送的路径的缓存
                '------------------------
                If InStr(Me.lstTemp.List(0), "空闲:") <> 0 Then     '如果是根目录列表
                    sNowPath = "根目录"             '返回“根目录”字符串
                Else
                    sNowPath = Me.Dir.Path          '否则就是目录列表当前的路径
                End If
                wsSendData Me.wsMessage, "NPH|" & sNowPath          '发送路径
            
            Case "VPH"                      '浏览指定路径请求
                Dim sViewPath As String     '请求的目录
                Dim svSend As String        '要发送的内容
                '------------------------
                sViewPath = Replace(sTemp(i), "VPH|", "")           '去掉请求头，分离出路径
                '------------------------
                If sViewPath = "根目录" Then                        '如果请求的是根目录列表
                    Call MakeRootList                                   '生成根目录表
                    svSend = "GRT|"                                     '数据头
                    For k = 0 To Me.lstTemp.ListCount - 1               '遍历列表框
                        svSend = svSend & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, svSend                     '发送数据
                Else                                                '如果请求的是其它文件夹路径
                    If Right(sViewPath, 1) <> "\" Then
                        sViewPath = sViewPath & "\"                     '规范化路径名
                    End If
                    If IsPathExists(sViewPath) = False Then                 '检测到请求的目录无效
                        wsSendData Me.wsMessage, "VPH|ERR"                  '发送错误消息
                    Else                                                '检测到请求的目录存在
                        MakeList sViewPath                                  '生成目录表
                        '-------------------------------------
                        svSend = "GPH|"                                     '数据头
                        For k = 0 To Me.lstTemp.ListCount - 1               '遍历列表框
                            svSend = svSend & Me.lstTemp.List(k) & "||"
                        Next k
                        '==========
                        wsSendData Me.wsMessage, svSend                     '发送数据
                    End If
                End If
                
            '===================================================================================================
            '===================================================================================================
            '锁定鼠标键盘
            Case "BLK"                          '锁定
                If bBlockInput = True Then
                    Me.tmrBlockInput.Enabled = True
                End If
                
            Case "ULK"                          '解锁
                If bBlockInput = True Then
                    Me.tmrBlockInput.Enabled = False
                    BlockInput False
                End If
            
            '===================================================================================================
            '===================================================================================================
            '文件传输
            Case "UPL"                          '上传文件请求
                Dim sFileName As String, lFileSize As Long              '文件名 & 文件大小
                '----------------------------------------
                IsUpload = True                                         '上传状态
                FileRec = 0                                             '接收文件数清空
                Me.tmrCalcPerSecond.Enabled = True                      '启动计算每秒字节数的计时器
                sFileName = Split(sTemp(i), "|")(1)                     '分离出文件名
                lFileSize = CLng(Split(sTemp(i), "|")(2))               '分离出文件大小
                Me.wsFile.Tag = sFileName & "|" & lFileSize & "|" & Me.File.Path            '赋值Tag 【源路径|文件大小|写入目录的路径】
                Me.wsFile.Close                                         '初始化Winsock
                Me.wsFile.Connect Me.wsMessage.RemoteHostIP, 20236      '连接到文件发送端
            
            Case "NXT"                          '上传下一个文件请求
                Dim nFileName As String, nFileSize As Long              '文件名 & 文件大小
                '----------------------------------------
                nFileName = Split(sTemp(i), "|")(1)                     '分离出文件名
                nFileSize = CLng(Split(sTemp(i), "|")(2))               '分离出文件大小
                Me.wsFile.Tag = nFileName & "|" & nFileSize & "|" & Split(Me.wsFile.Tag, "|")(2)    '赋值Tag 【源路径|文件大小|写入目录的路径】
                Me.wsFile.SendData "NNT"                                '告诉对方我已经准备好接收下一个文件啦！
            
            Case "END"                          '任务结束
                Close #2                                                '关闭文件，防止有文件忘记关闭
                With frmBeingControlled
                    .edStatus1.Text = "任务完成"
                    .edStatus2.Text = "共接收到" & FileRec & "个文件"
                    .edStatus3.Visible = False
                End With
                Me.tmrCalcPerSecond.Enabled = False                     '关闭计算每秒字节数的计时器
            
            Case "DWL"                          '下载文件请求
                Dim iFileIndex() As String          '所有文件的序号
                Dim iFilePath As String             '当前选择的文件的路径
                '-------------------------------------------------------
                IsUpload = False                                                                    '下载状态
                iFileIndex = Split(Replace(sTemp(i), "DWL|", ""), vbCrLf)                           '分割出每个文件的Index
                '---------------------------
                CurrentFile = -1                                                                    '一开始文件序号为-1
                TotalFiles = 0                                                                      '文件总数记为0
                lSent = 0                                                                           '清空发送数据字节数
                oldSent = 0
                ReDim EachFile(0)                                                                   '清空每个文件的缓存列表
                FileList = ""                                                                       '清空文件列表
                '---------------------------
                TotalFiles = UBound(iFileIndex) - 1                                                 '设置文件总数
                For k = 0 To UBound(iFileIndex) - 1                                                 '减去一是因为最后面有一个Split导致的空数据
                    iFilePath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")     '规范化路径名
                    iFilePath = iFilePath & Me.lstTemp.List(iFileIndex(k) - 1)                          '添加上文件名及文件大小
                    FileList = FileList & iFilePath & vbCrLf                                            '添加到文件列表里
                Next k
                '-----------------
                EachFile = Split(FileList, vbCrLf)                                                  '分离出每个文件 【完整路径|文件大小】
                For k = 0 To UBound(EachFile) - 1                                                   '仅仅保留完整路径
                    EachFile(k) = Split(EachFile(k), "|")(0)
                Next k
                Call NextFile                           '下一文件
                Me.wsFile.Close                                         '初始化Winsock
                Me.wsFile.Connect Me.wsMessage.RemoteHostIP, 20236      '连接到文件接收端
            
        End Select
    Next i
End Sub

Private Sub wsPic_Close()
    Me.wsPic.Close                  '断开自动重新监听
    Me.wsPic.Bind 20234
    Me.wsPic.Listen
End Sub

Private Sub wsPic_ConnectionRequest(ByVal requestID As Long)
    Me.wsPic.Close                  '接受连接
    Me.wsPic.Accept requestID
End Sub

Private Sub wsMessage_ConnectionRequest(ByVal requestID As Long)
    Me.wsMessage.Close              '接受连接
    Me.wsMessage.Accept requestID
    IsRemoteControl = Me.optControl.Value       '获取远控模式：远程控制模式 or 文件传输模式
    Me.mnuShowWindow.Enabled = False
End Sub

Private Sub wsPic_DataArrival(ByVal bytesTotal As Long)
    Dim dat() As Byte
    Me.wsPic.GetData dat, vbArray Or vbByte
    If bytesTotal - 1 = 1 Then                  '消息一：发送空图
        SendDat
    End If
    If bytesTotal - 1 = 2 Then                  '消息二：开始发送连续图片
        Me.tmrForceRefresh.Enabled = True           '开始不断使屏幕刷新
        Call SendLoop                               '连续发图
    End If
End Sub
