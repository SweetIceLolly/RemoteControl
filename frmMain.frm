VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Զ�̿���"
   ClientHeight    =   4596
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8424
   BeginProperty Font 
      Name            =   "����"
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
   StartUpPosition =   2  '��Ļ����
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
         ToolTipText     =   "����"
         Top             =   0
         Width           =   360
      End
      Begin VB.Image imgSettings 
         Height          =   360
         Left            =   7620
         Picture         =   "frmMain.frx":19A0
         Stretch         =   -1  'True
         ToolTipText     =   "����"
         Top             =   0
         Width           =   360
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Զ�̿���"
         BeginProperty Font 
            Name            =   "����"
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
      Caption         =   "����Զ�̿���"
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
      Caption         =   "�������ӿ���"
      ForeColor       =   16777215
      BackColor       =   16416003
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      ToolTipText     =   "��Զ�̿��Ʒ�ʽ���ӶԷ�"
      Top             =   2520
      Width           =   3852
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Զ�̿���ģʽ"
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
         Name            =   "����"
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
      ToolTipText     =   "���ļ����䷽ʽ���ӶԷ�"
      Top             =   3000
      Width           =   3855
      _Version        =   786432
      _ExtentX        =   6800
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "�ļ�����ģʽ"
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
      Text            =   "Զ�̿���"
      Picture         =   "frmMain.frx":7FF0
   End
   Begin VB.Image imgCopy 
      Height          =   360
      Index           =   1
      Left            =   3600
      Picture         =   "frmMain.frx":8CCA
      Stretch         =   -1  'True
      ToolTipText     =   "�����������뵽������"
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
      ToolTipText     =   "��������IP��������"
      Top             =   1440
      Width           =   360
   End
   Begin VB.Label labState 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����IP"
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
      Caption         =   "�������"
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
      Caption         =   "����Զ�̿���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�Է�IP��ַ"
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
      Caption         =   "����Զ�̵���"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "��ʾ������(&S)"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�˳�(&E)"
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
'   |��|
'   |B |
'   |U |
'   |G |
'   |��|
'   |��|������__
' ,_|  |_,�� /��)
'   (Oo����/ _I_
'   +\ \��||�� ��|
'  �� \ \||_ 0��
' ���� \/.:.\- \
'   ����|.:. /-----\
'   ����|___|::oOo::|
'   ����/�� |:<_T_>:|
'       |   |::oOo::|
'       |    \-----/
'       |:_|__|
'       |===|==|
'       |   |  |
'       |&  \  \
'      *( ,  `'-.'-.
'       `"`"""""`""`
'
'       BUG��·������

'================================================================
      
'ͼ��˿ڣ�20234
'��Ϣ�˿ڣ�20235
'�ļ��˿ڣ�20236
'----------------------
'δ���ӣ���ɫ
'�����У���ɫ
'�����ӣ���ɫ
'�������ӣ���ɫ
'----------------------
'ͨ���������ָ{S}

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
'����Ϊzlibѹ�����õ��ĺ���
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Compress Lib "zlib.dll" Alias "compress" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Const OFFSET As Long = &H8

'���̿���
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

Dim bMouseDown As Boolean                           '�Ƿ������
Dim bTray As Boolean                                '��ǰ�Ƿ���������״̬

Public IsUpload As Boolean                          '�Է��Ƿ�Ϊ�ϴ�״̬ ��True:�ϴ� False:���ء�

Public recBytes As Long                             '���յ����ļ������ֽ���
Public oldBytes As Long                             '��һ����յ����ļ������ֽ���
Public BytesPerSec As Long                          'ÿһ����ֽ���
Public lSent As Long, oldSent As Long               '�ѷ��͵��ֽ��� ��һ�뷢�͵��ֽ���

Public FileRec As Integer                           '�ɹ����յ��ļ���
Public FileSent As Integer                          '�ɹ����͵��ļ���

Dim EachFile() As String                            '�ļ��б����ÿ���ļ�
Public FileList As String                           '�����͵��ļ��б�
Public CurrentFile As Integer                       '��ǰ�ļ���Index
Public TotalFiles As Integer                        '�����͵��ļ�����
Public CurrentSize As Long                          '��ǰ���ļ��Ĵ�С
Dim dSendTemp() As Byte                             '׼�����͵��ļ�

'========================================================================================================

Public Function NextFile() As Boolean           '��һ�ļ�����
    CurrentFile = CurrentFile + 1
    If CurrentFile > UBound(EachFile) - 1 Then          '�����ų����ļ�����
        NextFile = False
        Exit Function
    End If
    '---------------------------
    Do While LoadFile(EachFile(CurrentFile)) = False    '�����ȡ�ļ�����ͼ�����һ���ļ�
        CurrentFile = CurrentFile + 1                       '�ļ���� + 1
        If CurrentFile > UBound(EachFile) - 1 Then          '�����ų����ļ�����
            NextFile = False
            Exit Function
        End If
    Loop
    NextFile = True
End Function

Public Function LoadFile(sFileName As String) As Boolean    '�����ļ�����
    On Error Resume Next
    '-----------------------------------
    LoadFile = True
    Open sFileName For Binary As #1                             '���ļ�
        If Err.Number <> 0 Then                                 '�ļ�������
            Close #1                                            '�ر��ļ�
            LoadFile = False
            Exit Function
        End If
        '--------------------
        ReDim dSendTemp(LOF(1))                                      '�����ڴ�
        CurrentSize = LOF(1)
        If Err.Number <> 0 Then                                 '�ļ�����
            MsgBox "�ļ�����" & sFileName
            LoadFile = False
            Exit Function
        End If
        If UBound(dSendTemp) = 0 Then                                '�����ȡ�����ֽ���Ϊ0
            Close #1                                            '�ر��ļ�
            LoadFile = False
            Exit Function
        End If
        Get #1, , dSendTemp                                          '��ȡ�ļ�
    Close #1
    lSent = 0                                               '����ļ������ֽ���
    oldSent = 0
End Function

Sub SendLoop()                                  'ѭ��������Ļ����
    Do While XH                                     '��Ҫ����Do���¡������Է�������Żᷢ����һ�ŵ�~�������������ѭ��~
        Call Send_data
        Sleep 1
        If bKeepSending = False Then
            Exit Do
        End If
    Loop
End Sub

Private Function MyGetCursor() As Long          '�������λ��
    Dim hWindow As Long, dwThreadID As Long, dwCurrentThreadID As Long
    Dim Pt As POINTAPI
    GetCursorPos Pt
    x = Pt.x - 6     'ΪʲôҪ��ȥ6�أ���Ϊ��ȡ����ָ��λ�û�����ƫ�ƣ���ȥ6������һ���̶��Ͻ���λ�ã����ǻ��ǲ�׼...
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

Private Sub Send_data()             '������Ļ������
    Dim Bmp As BITMAP
    Dim hdc As Long, BD As Long
    Dim Tpicture As Boolean
    '===========
    Dim b As String
    Dim a() As Byte
    '=======
    If g = 0 Then BitBlt iDC3, 0, 0, W, H * F, sdc, 0, 0, vbSrcCopy             '3DC����
    If g = 0 Then DrawIcon iDC3, x, y, MyGetCursor
    '��ע3��1��2���ܸ㷴
    '================================
    BitBlt iDC1, 0, 0, W, H, iDC3, 0, g, vbSrcCopy                              '1DC����
    '================
    BitBlt iDC2, 0, 0, W, H, iDC4, 0, g, vbSrcCopy                              '2DC����
    '================================
    
    '================================
    BD = CompMemory(ByVal opeinter1, ByVal opeinter2, bmpLen)                   '��Ļ�Ա�
    Tpicture = CBool(BD = bmpLen)
    '================================
    
    If Not Tpicture Then                                                        '�����Ļ���ݷ����˱仯
        
        XH = False
        
        BitBlt iDC4, 0, g, W, H, iDC1, 0, 0, vbSrcCopy                          '4DC����
        '===============
        BitBlt iDC1, 0, 0, W, H, iDC2, 0, 0, vbSrcInvert                        'ɨ��
        '===============
        ReDim r1(bmpLen) As Byte                                                '��ͼ������ʵ�ʵĴ�С���仺����
        CopyMemory r1(0), ByVal opeinter1, bmpLen
        If bCompress = True Then
            CompressByte r1                                                         'ѹ��
        End If
        '=======================                  [ͷ]�����ݽ�������
        b = Format(UBound(r1), "00000000")
        Mid$(b, 1, 1) = J
        If J = 0 Then Mid$(b, 1, 1) = 9
        a = StrConv(b, vbFromUnicode)                                           '�ַ���ת��Ϊ�ֽ�����
        Me.wsPic.SendData a
        '=====
    End If
    J = J + 1
    If J >= F Then J = 0
    g = J * H
End Sub

Private Sub SendDat()               '���Ϳ�ͼ���öԷ�׼����ʼ����
    XH = True
    Me.wsPic.SendData r1
End Sub

Function CompressByte(ByteArray() As Byte)          '����ѹ������
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

Private Sub Del_dc()                                '�ͷ��ڴ�DC����
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

Private Sub sLoadf()                                '�����ڴ�DC���̣������ǰ�λͼ����
    Dim bi24BitInfo As BITMAPINFO
    With bi24BitInfo.bmiHeader                      '����λͼ����
        .biBitCount = 16
        .biCompression = BI_RGB
        .biPlanes = 1
        .biSize = Len(bi24BitInfo.bmiHeader)
        .biWidth = W
        .biHeight = H
    End With
    '=======                                        '�󶨵�λͼ
    iDC1 = CreateCompatibleDC(0)                                                '1
    iBitmap1 = CreateDIBSection(iDC1, bi24BitInfo, DIB_RGB_COLORS, opeinter1, ByVal 0&, ByVal 0&)
    SelectObject iDC1, iBitmap1
    '=======
    iDC2 = CreateCompatibleDC(0)                                                '2
    iBitmap2 = CreateDIBSection(iDC2, bi24BitInfo, DIB_RGB_COLORS, opeinter2, ByVal 0&, ByVal 0&)
    SelectObject iDC2, iBitmap2
    '=======
End Sub

Private Sub SLoad()                                 'ͬ���Ǽ����ڴ�DC�Ĺ��̣������Ǽ��ز���
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
    iDC3 = CreateCompatibleDC(0)                                                '3�ڴ�dc
    iBitmap3 = CreateDIBSection(iDC3, bi24BitInfo1, DIB_RGB_COLORS, opeinter3, ByVal 0&, ByVal 0&)
    SelectObject iDC3, iBitmap3
    '=======
    '=======
    iDC4 = CreateCompatibleDC(0)                                                '4�ڴ�dc
    iBitmap4 = CreateDIBSection(iDC4, bi24BitInfo1, DIB_RGB_COLORS, opeinter4, ByVal 0&, ByVal 0&)
    SelectObject iDC4, iBitmap4
    '=======
    
    '=======
    Dim sdc1 As Long
    sdc1 = GetDC(0)
    BitBlt iDC4, 0, 0, W, H * F, sdc1, 0, 0, vbSrcCopy                          '4DC����
    DrawIcon iDC4, x, y, MyGetCursor
    ReleaseDC 0, sdc1
    '======
    Call sLoadf                 '���������а󶨲���
End Sub

Private Sub New_dc()                        '������λͼ����
    '===========
    Dim b As String
    Dim a() As Byte
    '================
    ReDim r1(bmpLen * F) As Byte                                                '��ͼ������ʵ�ʵĴ�С���仺����
    CopyMemory r1(0), ByVal opeinter4, bmpLen * F
    '=======================                  [ͷ]�����ݽ�������
    If bCompress = True Then
        CompressByte r1                                                         'ѹ��
    End If
    b = Format(UBound(r1), "00000000")
    Mid$(b, 1, 1) = J
    If J = 0 Then Mid$(b, 1, 1) = 9
    a = StrConv(b, vbFromUnicode)                                               '�ַ���ת��Ϊ�ֽ�����
    Me.wsPic.SendData a
End Sub

'========================================================================

Private Sub cmdConnect_Click()
    Me.shpState.FillColor = vbYellow            '����״̬
    Me.labState.Caption = "��������..."
    If Me.optControl.Value = True Then                  'Զ��ģʽ
        frmRemoteControl.wsPicture.Close
        frmRemoteControl.wsMessage.Close
        frmRemoteControl.wsPicture.Connect Me.comIP.Text, 20234         '��������
        frmRemoteControl.wsMessage.Connect Me.comIP.Text, 20235
    Else                                                '�ļ�����ģʽ
        frmFileTransfer.wsMessage.Close
        frmFileTransfer.wsMessage.Connect Me.comIP.Text, 20235          '��������
    End If
    Me.tmrTimeOut.Enabled = True
End Sub

Private Sub comIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdConnect_Click           '���»س������൱�ڰ���ť
    End If
End Sub

Private Sub edLocalIP_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(0).Picture = Me.imgCopy1.Picture         '��ͼ��~Ϊ������ (*^__^*)
    Me.imgCopy(1).Picture = Me.imgCopy1.Picture
End Sub

Private Sub edPassword_LostFocus()
    sUserPassword = Me.edPassword.Text                  '�����û�����
    Call SaveConfig                                     '����һ���༭��ɾͱ��������ļ�
End Sub

Private Sub edPassword_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(0).Picture = Me.imgCopy1.Picture         '��ͼ��~Ϊ������ (*^__^*)
    Me.imgCopy(1).Picture = Me.imgCopy1.Picture
End Sub

Private Sub Form_Load()
    On Error Resume Next
    '=======================================
    Call LoadConfig                             '��ȡ�����ļ�
    If bTrayWhenStart Then                      '���������ʱ����С����������
        Me.Tray.MinimizeToTray Me.hwnd
        bTray = True
    End If
    If Trim(LCase(Command)) = "/autostart" And bHideWhenStart Then      '������Զ�����ģʽ�����Ժ�̨ģʽ����
        Me.Tray.Icon = Nothing
        Me.Hide
        bTray = True
    End If
    '=======================================
    Me.Icon = Me.imgMainIcon.Picture            '������ͼ�꣬Ϊ�˸����� (*^__^*)
    Me.edLocalIP.Text = Me.wsMessage.LocalIP    '��ʾ����IP
    '=======================================
    Me.wsMessage.Close                          '��ʼ����
    Me.wsPic.Close
    Me.wsMessage.Bind 20235
    Me.wsMessage.Listen
    Me.wsPic.Bind 20234
    Me.wsPic.Listen
    '�������ɹ�
    If Me.wsMessage.state = sckListening And Me.wsPic.state = sckListening And Err.Number = 0 Then
        Me.labState.Caption = "׼������"
        Me.shpState.FillColor = vbGreen
    Else
        Me.labState.Caption = "������"
        Me.shpState.FillColor = vbRed
        Me.tmrRetry.Enabled = True
    End If
    '=======================================
    '������λ�������
    Dim tmpPsw As String
    For i = 1 To 6
        Randomize
        tmpPsw = tmpPsw & Chr(Int((Asc("z") - Asc("a") + 1) * Rnd + Asc("a")))
    Next i
    Me.edPassword.Text = tmpPsw
    '=======================================
    '��ȡ��Ļ������Ҫ�Ĳ���
    W = Screen.Width \ Screen.TwipsPerPixelX
    H = Screen.Height \ Screen.TwipsPerPixelY
    F = 6
    H = H \ F
    bmpLen = W * H * 2
    '----------------------
    '�����ڴ�DC
    Call Del_dc
    Call SLoad
    '=======================================
    If Not bTrayWhenStart And Not bHideWhenStart Then
        Me.Show                                         'һ�ж�ĬĬ���������չ�����ң�
        Me.Refresh
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(0).Picture = Me.imgCopy1.Picture         '��ͼ��~Ϊ������ (*^__^*)
    Me.imgCopy(1).Picture = Me.imgCopy1.Picture
    Me.imgSettings.Picture = Me.imgSettings1.Picture
    Me.imgAbout.Picture = Me.imgAbout1.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If bExitWhenClose And Not bNoExit Then                  '��������˵���رհ�ť �� ����ֹ����ر� �͹رվ�ֱ����ɱ
        End
    Else
        Me.Tray.MinimizeToTray Me.hwnd                      '��С����������
        Cancel = True                                       'ȡ���ر�
        bTray = True
    End If
End Sub

Private Sub imgAbout_Click()
    frmAbout.Show                       '��ʾ���ڴ���
    Me.Enabled = False
End Sub

Private Sub imgCopy_Click(Index As Integer)
    Clipboard.Clear
    If Index = 0 Then                                   '��ָ�����ݸ��Ƶ�������
        Clipboard.SetText Me.edLocalIP.Text
    Else
        Clipboard.SetText Me.edPassword.Text
    End If
End Sub

Private Sub imgCopy_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(Index).Picture = Me.imgCopy3.Picture         '��ͼ��~Ϊ������ (*^__^*)
    bMouseDown = True
End Sub

Private Sub imgCopy_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown = False Then
        Me.imgCopy(Index).Picture = Me.imgCopy2.Picture     '��ͼ��~Ϊ������ (*^__^*)
    End If
End Sub

Private Sub imgCopy_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgCopy(Index).Picture = Me.imgCopy1.Picture         '��ͼ��~Ϊ������ (*^__^*)
    bMouseDown = False
End Sub

Private Sub imgMainIcon_Click()
    MsgBox "I Love CXT����(*^__^*)", 64, "(*^__^*) "
End Sub

Private Sub imgSettings_Click()
    '��������״̬
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
        For i = 0 To Me.comIP.ListCount - 1         '�������IP
            .lstIP.AddItem Me.comIP.List(i)
        Next i
    End With
    '=========================================
    frmSettings.Show                    '��ʾ���ô���
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
    a = MsgBox("���Ҫ�˳���", 32 + vbYesNo, "ȷ��")          '�˳�ȷ��
    If a = 6 Then
        End
    End If
End Sub

Private Sub mnuShowWindow_Click()
    If bTray = True Then
        Me.Tray.MaximizeFromTray Me.hwnd        '�ָ�����
        bTray = False
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.imgSettings.Picture = Me.imgSettings1.Picture
    Me.imgAbout.Picture = Me.imgAbout1.Picture
End Sub

Private Sub tmrBlockInput_Timer()
    '�����������Σ�գ�
    BlockInput True
End Sub

Private Sub tmrCalcPerSecond_Timer()
    If IsUpload Then                        '����Ǵӱ����ϴ�״̬
        BytesPerSec = recBytes - oldBytes       '�����ÿ�������ֽ���
        oldBytes = recBytes                     '��һ�������ֽ���������һ��������ֽ���`
    Else                                    '����Ǵӱ�������״̬
        BytesPerSec = lSent - oldSent           '�����ÿ�������ֽ���
        oldSent = lSent                         '��һ�������ֽ���������һ��������ֽ���`
    End If
    frmBeingControlled.Refresh              'ˢ�±��ش���
End Sub

Private Sub tmrChangeIcon_Timer()
    Dim Cur As POINTAPI
    GetCursorPos Cur
    '�����������˴���ͻָ�����ͼƬ��ԭ��
    If Cur.x * Screen.TwipsPerPixelX < Me.Left Or _
       Cur.x * Screen.TwipsPerPixelX > Me.Left + Me.Width Or _
       Cur.y * Screen.TwipsPerPixelY < Me.Top Or _
       Cur.y * Screen.TwipsPerPixelY > Me.Top + Me.imgSettings.Top + Me.Height Then
        Call Form_MouseMove(0, 0, 0, 0)
    End If
End Sub

Private Sub tmrForceRefresh_Timer()
    '��ͣ������Ļ�����Ͻǻ��ƴ�СΪһ���صĵ㣬�Դ�����Ļˢ���¼�
    Me.picRefresh.PSet (0, 0), RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
    BitBlt sdc, 0, 0, 1, 1, Me.picRefresh.hdc, 0, 0, vbSrcCopy
End Sub

Private Sub tmrRetry_Timer()
    On Error Resume Next
    Me.wsMessage.Close              '���Լ���
    Me.wsPic.Close
    Me.wsMessage.Bind 20235
    Me.wsPic.Bind 20234
    Me.wsMessage.Listen
    Me.wsPic.Listen
    '�������ɹ�
    If Me.wsMessage.state = sckListening And Me.wsPic.state = sckListening And Err.Number = 0 Then
        Me.labState.Caption = "׼������"
        Me.shpState.FillColor = vbGreen
        Me.tmrRetry.Enabled = False
    Else
        Me.labState.Caption = "������"
        Me.shpState.FillColor = vbRed
    End If
End Sub

Private Sub tmrReturn_Timer()               '��ʱ���״̬��ʾ�ָ�
    Me.shpState.FillColor = vbGreen
    Me.labState = "׼������"
    Me.tmrReturn.Enabled = False
End Sub

Private Sub tmrTimeOut_Timer()              '�ж��Ƿ����ӳ�ʱ
    Me.shpState.FillColor = vbRed
    Me.labState.Caption = "����ʧ��"
    Me.tmrReturn.Enabled = True
    Me.tmrTimeOut.Enabled = False
End Sub

Private Sub Tray_DblClick()
    If bTray = True And Me.mnuShowWindow.Enabled = True Then
        Me.Tray.MaximizeFromTray Me.hwnd        '˫��������ͼ��ʱ��ԭ����
        bTray = False
    End If
End Sub

Private Sub Tray_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu Me.mnuPopup                   '�����˵�~
    End If
End Sub

Private Sub wsFile_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim Temp() As Byte              '���յ��Ķ���������
    Dim recStr As String            '���յ����ַ�������
    '-------------
    Dim sFileName As String         'д���ļ�·��
    Dim sFileTitle As String        '�ļ�����
    Dim lFileSize As Long           '�ļ���С
    Dim TitleSplitTmp() As String   '�ļ�����ָ��
    '-------------
    Dim nFileName As String         '�����ļ�Դ·��
    Dim nFileTitle As String        '�����ļ��ı���
    Dim nTitleSplitTmp() As String  '�ļ�����ָ��
    '----------------------------
    Me.wsFile.GetData Temp          '��ȡ����
    '--------------
    If IsUpload Then                '������ϴ�ģʽ
        '������ҵ�Tag���ȡ��Ϣ��~
        sFileName = Split(Me.wsFile.Tag, "|")(0)                                                            '�����Դ�ļ�·��
        TitleSplitTmp = Split(sFileName, "\")                                                               '���ա�\��������
        sFileTitle = TitleSplitTmp(UBound(TitleSplitTmp))                                                   '�õ��ļ�����
        lFileSize = CLng(Split(Me.wsFile.Tag, "|")(1))                                                      '������ļ���С
        sFileName = Split(Me.wsFile.Tag, "|")(2)                                                            '�����д���ļ�·��
        sFileName = IIf(Right(sFileName, 1) = "\", sFileName & sFileTitle, sFileName & "\" & sFileTitle)    '�õ�д���ļ�������·��
        '--------------
        '��ȡβ�����ݲ鿴�Ƿ��н������
        '��ö��������ݵ�β����λ����
        Select Case Temp(UBound(Temp) - 2) & "|" & Temp(UBound(Temp) - 1) & "|" & Temp(UBound(Temp))    '�����ǲ�������
            Case "70|73|78"                         '��FIN�����������Ϣ
                Dim lstData() As Byte                       '�������ǰ������
                frmBeingControlled.ProgressBar.Value = 100  '��������������
                If UBound(Temp) - 2 > 0 Then            '������Ǹոպõ��������
                    ReDim lstData(UBound(Temp) - 2)         '����Ƚ��յ�������3�ֽڵ��ڴ�
                    For i = 0 To UBound(Temp) - 2           '�����ݸ��ƹ�ȥ
                        lstData(i) = Temp(i)
                    Next i
                    Put #2, LOF(2) + 1, lstData             'д���ļ�
                End If
                FileRec = FileRec + 1                   '�����ļ��� + 1
                Close #2                                '�ر��ļ�
                Me.wsFile.SendData "NXT"                '��������һ�ļ�
                Exit Sub
        
            Case "67|78|84"                         '��CNT����������ӽ���ȷ����Ϣ
                oldBytes = 0                                                '��ʼ�����յ����ֽ���
                recBytes = 0
                With frmBeingControlled                                     '���±��ش����״̬
                    .edStatus1.Text = "�Է������ϴ��ļ������ĵ���"
                    .edStatus2.Text = "�ļ�����" & sFileTitle
                    .edStatus3.Text = "0 Byte/s 0 Byte/" & SizeWithFormat(lFileSize) & " 0%"
                    .edStatus2.Visible = True
                    .edStatus3.Visible = True
                    .Refresh
                End With
                '-----------------------
                '�ж��ļ��Ƿ�����
                If IsPathExists(sFileName, True) = True Then                '��⵽�Ѵ����ļ�
                    Me.wsFile.SendData "TSN"                         '���͸��������ļ�ȷ��
                    Exit Sub
                End If
                '���û����������׼��������Ϣ
                Open sFileName For Binary As #2         '�����ļ�
                Me.wsFile.SendData "IRD"
                Exit Sub                                '��������������˳����̣���д���ļ�
    
            Case "89|69|83"                         '��YES������ͬ���ļ���Ϣ
                Kill sFileName                          'ɾ����ͬ�����ļ�
                If Err.Number <> 0 Then                 '����ļ��޷�ɾ��
                    Me.wsFile.SendData "DFL"     '���߶Է��޷�����Ү������
                    Exit Sub
                End If
                Open sFileName For Binary As #2         '�����ļ�
                recBytes = 0                            '��ս��յ����ֽ���
                Me.wsFile.SendData "IRD"         '����׼��������Ϣ
                Exit Sub                                '�˳����̣���д���ļ�
    
            Case "78|82|70"                         '��NRF��ȡ������ͬ���ļ���Ϣ
                Me.wsFile.SendData "NXT"         '��������һ�ļ�
                Exit Sub                                '�˳����̣���д���ļ�
    
        End Select
        Put #2, LOF(2) + 1, Temp                'д���ļ�
        recBytes = recBytes + UBound(Temp)      '���յ����ֽ���
        '���±�ǩ�͹�����
        frmBeingControlled.ProgressBar.Value = (recBytes / lFileSize) * 100
        frmBeingControlled.edStatus3.Text = SizeWithFormat(BytesPerSec) & "/s " & SizeWithFormat(recBytes) & "/" & _
                                            SizeWithFormat(lFileSize) & " " & Format(recBytes / lFileSize * 100, "0.00") & "%"
    Else                            '���������ģʽ
        '����Ӹ��ֱ������ȡ��Ϣ��~
        nFileName = EachFile(CurrentFile)                       '�õ������ļ�������·��
        nTitleSplitTmp = Split(nFileName, "\")                  '���ա�\������
        nFileTitle = nTitleSplitTmp(UBound(nTitleSplitTmp))     '�õ��ļ�����
        '--------------
        '��ȡβ�����ݲ鿴�Ƿ��н������
        '��ö��������ݵ�β����λ����
        Select Case Temp(UBound(Temp) - 2) & "|" & Temp(UBound(Temp) - 1) & "|" & Temp(UBound(Temp))    '�����ǲ�������
            Case "68|78|84"                     '��DNT�����ӽ���ȷ��
                lSent = 0                                       '��շ��͵��ֽ���
                oldSent = 0
                With frmBeingControlled                                     '���±��ش����״̬
                    .edStatus1.Text = "�Է����ڴ����ĵ��������ļ�"
                    .edStatus2.Text = "�ļ�����" & nFileTitle
                    .edStatus3.Text = "0 Byte/s 0 Byte/" & SizeWithFormat(CurrentSize) & " 0%"
                    .edStatus2.Visible = True
                    .edStatus3.Visible = True
                    .Refresh
                End With
                '-----------------------
                '������ͷ|�ļ�����|�ļ���С|��ǰ�ļ����|�ļ�������
                Me.wsFile.SendData "DNT|" & nFileTitle & "|" & CurrentSize & "|" & CurrentFile & "|" & UBound(EachFile)
                Exit Sub
            
            Case "78|88|84"                     '��NXT����һ�ļ�
                If NextFile = True Then
                    '����ɹ�����һ���ļ��ͷ����ϴ���һ���ļ�����
                    sFileName = Split(EachFile(CurrentFile), "|")(0)                            '�����Դ�ļ�·��
                    TitleSplitTmp = Split(sFileName, "\")                               '���ա�\��������
                    sFileTitle = TitleSplitTmp(UBound(TitleSplitTmp))                   '�õ��ļ�����
                    lSent = 0                                                           '��շ��͵��ֽ���
                    oldSent = 0
                    With frmBeingControlled                                             '���±��ش����״̬
                        .edStatus2.Text = "�ļ�����" & sFileTitle
                        .edStatus3.Text = "0 Byte/s 0 Byte/" & SizeWithFormat(CurrentSize) & " 0%"
                        .Refresh
                    End With
                    '������ͷ|�ļ�����|�ļ���С|��ǰ�ļ����|�ļ�������
                    Me.wsFile.SendData "DNT|" & sFileTitle & "|" & CurrentSize & "|" & CurrentFile & "|" & UBound(EachFile)
                Else
                    Close #2                                                '�ر��ļ�����ֹ���ļ����ǹر�
                    With frmBeingControlled
                        .edStatus1.Text = "�������"
                        .edStatus2.Text = "������" & FileSent & "���ļ�"
                        .edStatus3.Visible = False
                    End With
                    Me.tmrCalcPerSecond.Enabled = False                     '�رռ���ÿ���ֽ����ļ�ʱ��
                    Me.wsFile.SendData "END|" & CStr(FileSent)                          '�������������Ϣ
                End If
            
            Case "73|82|68"                     '��IRD��׼������
                Me.tmrCalcPerSecond.Enabled = True                                  '������ʱ��������ÿ�뷢���ֽ���
                Me.wsFile.SendData dSendTemp                                        '���Һݺݵ������ݣ��� (�s�F����)�s��ߩ���
                FileSent = FileSent + 1
            
        End Select
    End If
End Sub

Private Sub wsFile_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    If Not IsUpload Then                                                            '���������״̬
        lSent = lSent + bytesSent
        If lSent >= CurrentSize Then                                                '����ѷ������ݴ����ļ���С
            lSent = 0
            Me.wsFile.SendData "FIN"                                                    '���߶Է��������
            Exit Sub
        End If
        '=======================================================
        '���±��ش���״̬
        Dim tmpShow As String                                                       '��ʾ���ı�����
        tmpShow = SizeWithFormat(BytesPerSec) & "/s"                                'ÿ���ֽ��� = ���ڷ��͵������ֽ��� - ��һ�뷢�͵������ֽ���
        tmpShow = tmpShow & " " & SizeWithFormat(lSent)                             '�Ѿ����͵Ĵ�С
        tmpShow = tmpShow & "/" & SizeWithFormat(CurrentSize)                       '�ļ��Ĵ�С
        tmpShow = tmpShow & " " & Format(lSent / CurrentSize * 100, "0.00") & "%"   '����ٷֱ�
        frmBeingControlled.ProgressBar.Value = lSent / CurrentSize * 100            '��������ֵ
        frmBeingControlled.edStatus3.Text = tmpShow                                 '��ʾ����ǩ��
    End If
End Sub

Public Sub wsMessage_Close()
    Dim OpenPath As String                  '���������·��
    OpenPath = IIf(Right(App.Path, 1) = "\", App.Path & App.EXEName & ".exe", App.Path & "\" & App.EXEName & ".exe")
    Shell OpenPath, vbNormalFocus               '���۲���ǧ���� ����ǰͷ��ľ��
    End                                         '�����ˣ��������ˣ�����·�ˡ�
End Sub

Private Sub wsMessage_DataArrival(ByVal bytesTotal As Long)
    'On Error Resume Next
    Dim Temp As String                      '��������
    Dim sTemp() As String                   '�����и������
    Dim sTempPara() As String               '���ݸ����Ĳ����и�����Ļ�������
    Dim Cur As POINTAPI                     '�������������������Ʋ��֣�
    '====================================
    Me.wsMessage.GetData Temp               '��ȡ����
    sTemp = Split(Replace(Decrypt(Temp), "{END}", ""), "{S}")            '�и�����
    For i = 0 To UBound(sTemp)
        Select Case Left(sTemp(i), 3)       '��������
            '===================================================================================================
            '===================================================================================================
            '����
            Case "CNT"                                     '��������
                Me.labState.Caption = "��������: " & Me.wsMessage.RemoteHostIP    '����״̬
                Me.shpState.FillColor = vbYellow
                If Me.chkAllow.Value = xtpUnchecked Then        '��ǰ����������
                    wsSendData Me.wsMessage, "NAC"                    '���;ܾ�������Ϣ
                    Exit Sub
                End If
                If InStr(sTemp(i), "|") <> 0 Then               '����Ǵ���������Ļ�
                    sTempPara = Split(sTemp(i), "|")
                    If (sTempPara(1) = Me.edPassword.Text) Or _
                       (sTempPara(1) = sUserPassword And bUseUserPassword) Then
                    '��������������� ���� ʹ���û����������û�������� �� ������ȷ
                        Me.Hide                                                                 '����������
                        Me.labState.Caption = "��ǰ����:���� " & Me.wsMessage.RemoteHostIP      '����״̬
                        Me.shpState.FillColor = vbBlue
                        If Me.wsPic.state <> sckConnected Then              '���ͼƬ�˿�δ���������ļ�����ģʽ
                            IsRemoteControl = False
                        Else                                                '�������Զ��ģʽ
                            IsRemoteControl = True
                        End If
                        If Not bHideMode And Not bHideWhenStart Then        '�����������״̬
                            frmBeingControlled.Show                             '��ʾ���ش���
                            frmBeingControlled.Left = Screen.Width - frmBeingControlled.Width
                            frmBeingControlled.Top = Screen.Height - frmBeingControlled.Height * 2
                            frmBeingControlled.labMsg.Caption = "���ԣ�" & Me.wsMessage.RemoteHostIP
                        End If
                        '--------------------------------------
                        If IsRemoteControl Then                                 '�����Զ�̿���ģʽ
                            '���ͱ����ֱ���
                                wsSendData Me.wsMessage, "RES|" & CStr(Screen.Width / Screen.TwipsPerPixelX) & "|" & _
                                                                  CStr(Screen.Height / Screen.TwipsPerPixelY) & "|" & _
                                                                  CStr(Screen.TwipsPerPixelX) & "|" & _
                                                                  CStr(Screen.TwipsPerPixelY)
                            Else                                                '������ļ�����״̬
                                frmBeingControlled.labMode.Caption = "��ǰΪ�ļ�����ģʽ,�Է����Թ��������ļ�,�������ϴ������ļ���"
                                frmBeingControlled.labMsg.Caption = "���ԣ�" & Me.wsMessage.RemoteHostIP
                                '-------------------------------------
                                Dim tmpStringToSend As String                       'Ҫ���͵Ļ����ַ���
                                Call MakeRootList                                   '���ɸ�Ŀ¼��
                                tmpStringToSend = "GRT|"                            '����ͷ
                                For k = 0 To Me.lstTemp.ListCount - 1               '�����б��
                                    tmpStringToSend = tmpStringToSend & Me.lstTemp.List(k) & "||"
                                Next k
                                wsSendData Me.wsMessage, tmpStringToSend            '��������
                        End If
                        '��������ȷ��
                        wsSendData Me.wsMessage, "CNT|True"
                        
                    Else                                       '����������Ļ�
                        wsSendData Me.wsMessage, "CNT|Wrong"                  '�ܾ�����
                    End If
                Else                                         '���û����������Ļ�
                    wsSendData Me.wsMessage, "CNT|False"
                End If
                
            Case "BEG"                                       '�ͻ���˵���Ѿ�׼�������������ҾͿ�ʼ��ͼƬ��~
                Call New_dc                 '�����ڴ�DC��������Ļ����
                XH = True
                
            Case "SIF"                      '����ϵͳ��Ϣ
                '��ȡϵͳ��Ϣ
                Dim tempString As String       '�����ַ���
                Set wmi = GetObject("winmgmts:\\.\root\CIMV2")      '����WMI��ȡϵͳӲ����Ϣ
                '===============================================================
                Set oWMINameSpace = GetObject("winmgmts:")                            '��ȡ�汾
                Set SystemSet = oWMINameSpace.InstancesOf("Win32_OperatingSystem")
                For Each System In SystemSet
                    tempString = "ϵͳ�汾��" & System.Caption & vbCrLf & vbCrLf
                Next
                '----------------------------
                Set Msg = wmi.ExecQuery("select * from win32_processor")              '��ȡCPU��Ϣ
                For Each k In Msg
                    tempString = tempString & "CPU���ƣ�" & k.Name & vbCrLf
                Next
                tempString = tempString & vbCrLf
                '----------------------------
                Set Msg = wmi.ExecQuery("select * from win32_ComputerSystem")         '��ȡ�ڴ���Ϣ
                tempString = tempString & "�ڴ��С"
                For Each k In Msg
                    tempString = tempString & vbCrLf & CStr(Format(k.TotalPhysicalMemory / 1024 ^ 2, "0.00")) & "MB"
                Next
                tempString = tempString & vbCrLf
                '----------------------------
                Dim n As Integer
                Set Msg = wmi.ExecQuery("select * from win32_DiskDrive")              '��ȡӲ�̴�С��Ϣ
                tempString = tempString & vbCrLf & "Ӳ�̴�С" & vbCrLf
                For Each k In Msg
                    n = n + 1
                    tempString = tempString & "Ӳ��" & n & "��" & CStr(Format((k.Size / 1024 ^ 3), "0.00")) & "GB" & vbCrLf
                Next
                '----------------------------                                       '��ȡ��������С��Ϣ
                Set Msg = wmi.ExecQuery("select * from win32_LogicalDisk where DriveType='3'")
                For Each k In Msg
                    tempString = tempString & vbCrLf & "�̷���" & k.DeviceID & "  ��С��" & CStr(Format((k.Size / 1024 ^ 3), "0.00")) & "GB"
                Next
                '----------------------------
                tempString = tempString & vbCrLf & vbCrLf
                Set Msg = wmi.ExecQuery("select * from win32_VideoController")      '��ȡ�Կ���Ϣ
                For Each k In Msg
                    tempString = tempString & "�Կ��ͺţ�" & k.Name & "  �Դ棺" & CStr(Format(Abs(k.AdapterRAM / 1024 ^ 2), "0.00")) & "MB" & vbCrLf
                Next
                '===============================================================
                wsSendData Me.wsMessage, "SIF|" & tempString
            
            Case "VBS"                      'ִ��VBS����
                Dim tmpFile As String                                   '�ļ����ݻ���
                '------------------------------
                tmpFile = Split(sTemp(i), "VBS|")(1)
                Open App.Path & "\temp.vbs" For Output As #1            '�����ļ�
                    Print #1, tmpFile
                Close #1
                '------------------------------
                Shell "wscript.exe " & Chr(34) & App.Path & "\temp.vbs" & Chr(34)          '����VBS
            
            Case "CMD"                      'ִ������������
                Dim objCMD As Object                                    'CMD�ܵ�����
                Dim tmpCMD As String                                    '��������
                Dim SendTemp As String                                  '��������
                '------------------------------
                Set objCMD = New clsCMD                                 '����CMD�ܵ�
                tmpCMD = Split(sTemp(i), "CMD|")(1)                         '��������������
                Call objCMD.DosInput(tmpCMD)                                'ִ��CMD���
                SendTemp = objCMD.DosOutPutEx(10000)                        'ִ�л�ȡ�����
                Set objCMD = Nothing                                    '�ر�CMD�ܵ�
                wsSendData Me.wsMessage, "CMD|" & SendTemp              '����ִ�н��
            
            '===================================================================================================
            '===================================================================================================
            '���̹���
            Case "TSK"                      '���͵�ǰϵͳ�Ľ�����Ϣ
                '��ȡ���н���
                Dim myProcess As PROCESSENTRY32         '��������
                Dim mySnapshot As Long                  '���վ��
                Dim ProcessPath As String               '����·��
                Dim tmpString As String                 '�����ַ���
                '------------------------------
                Me.lstTemp.Clear                        '��ջ����б��
                myProcess.dwSize = Len(myProcess)
                '�������̿���
                mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
                '��ȡ��һ��������
                ProcessFirst mySnapshot, myProcess
                Me.lstTemp.AddItem Trim(myProcess.szexeFile)        '������̾�����
                '�������PID��λ��
                Me.lstTemp.List(Me.lstTemp.ListCount - 1) = Me.lstTemp.List(Me.lstTemp.ListCount - 1) & "|" & myProcess.th32ProcessID & "|" & GetFileName(myProcess.th32ProcessID)
                '������н���
                While ProcessNext(mySnapshot, myProcess)
                    Me.lstTemp.AddItem Trim(myProcess.szexeFile)        '������̾�����
                    '�������PID��λ��
                    Me.lstTemp.List(Me.lstTemp.ListCount - 1) = Me.lstTemp.List(Me.lstTemp.ListCount - 1) & "|" & myProcess.th32ProcessID & "|" & GetFileName(myProcess.th32ProcessID)
                Wend
                '------------------------------
                For k = 0 To Me.lstTemp.ListCount
                    tmpString = tmpString & Me.lstTemp.List(k) & vbCrLf
                Next k
                wsSendData Me.wsMessage, "TSK|" & tmpString             '�������н�����Ϣ
                
            Case "KTS"                      '������������
                Dim tmpTask() As String                                 '�ַ����ָ�����
                Dim tmpSend As String                                   '���ͷ����Ļ������
                '------------------------------
                sTempPara = Split(sTemp(i), vbCrLf)                     '�ָ��ÿ���س��������
                For k = 0 To UBound(sTempPara) - 1
                    tmpTask = Split(Replace(sTempPara(k), "KTS|", ""), "|")
                    If KillPID(CLng(tmpTask(1))) = True Then    'ɱɱɱ����
                        '�ɹ���
                        tmpSend = tmpSend & tmpTask(0) & "|" & tmpTask(1) & "|" & "T||"
                    Else
                        'ʧ�ܣ�
                        tmpSend = tmpSend & tmpTask(0) & "|" & tmpTask(1) & "|" & "F||"
                    End If
                Next k
                wsSendData Me.wsMessage, "KTS|" & tmpSend               '���ͷ�����Ϣ
                
            '===================================================================================================
            '===================================================================================================
            '����¼�
            Case "MMP"                      '�����������
                sTempPara = Split(sTemp(i), "|")
                SetCursorPos CLng(sTempPara(1)), CLng(sTempPara(2))
            
            Case "MLD"                      '����
                GetCursorPos Cur
                mouse_event LeftDown, Cur.x, Cur.y, 0, 0
            
            Case "MLU"                      '����
                GetCursorPos Cur
                mouse_event LeftUp, Cur.x, Cur.y, 0, 0
            
            Case "MRD"                      '����
                GetCursorPos Cur
                mouse_event RightDown, Cur.x, Cur.y, 0, 0
            
            Case "MRU"                      '����
                GetCursorPos Cur
                mouse_event RightUp, Cur.x, Cur.y, 0, 0
            
            Case "MMD"                      '����
                GetCursorPos Cur
                mouse_event MiddleDown, Cur.x, Cur.y, 0, 0
            
            Case "MMU"                      '����
                GetCursorPos Cur
                mouse_event MiddleUp, Cur.x, Cur.y, 0, 0
                
            Case "DBC"                      '���˫��
                GetCursorPos Cur                                        'Ϊʲô����ֻ����һ���أ���Ϊ˫����ʱ��Ͱ����˵����¼���
                mouse_event LeftDown Or LeftUp, Cur.x, Cur.y, 0, 0      '��˫����ʱ���ٷ������εĻ��ͻᷢ�����ΰ����ˡ�
            
            Case "MWU"                      '��������
                mouse_event MOUSEEVENTF_WHEEL, 0, 0, 150, 0
            
            Case "MWD"                      '��������
                mouse_event MOUSEEVENTF_WHEEL, 0, 0, -150, 0
                
            '===================================================================================================
            '===================================================================================================
            '�����¼�
            Case "KBD"                      '���̰��°���
                keybd_event Split(sTemp(i), "|")(1), 0, 0, 0
            
            Case "KBU"                      '�����ɿ�����
                keybd_event Split(sTemp(i), "|")(1), 0, KEYEVENTF_KEYUP, 0
            
            '===================================================================================================
            '===================================================================================================
            '�ļ�����
            Case "GRT"                      '��ȡ������
                Dim tmpSendString As String                     'Ҫ���͵Ļ����ַ���
                '-------------------------------------
                Call MakeRootList                               '���ɸ�Ŀ¼��
                tmpSendString = "GRT|"                          '����ͷ
                For k = 0 To Me.lstTemp.ListCount - 1           '�����б��
                    tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                Next k
                wsSendData Me.wsMessage, tmpSendString          '��������
             
            Case "GPH"                      '��ȡĿ¼�µ��ļ����У�
                Dim gPath As String               '�Է������Ŀ¼���ƻ���
                '-------------------------------------
                sTempPara = Split(sTemp(i), "|")
                gPath = Me.lstTemp.List(CInt(sTempPara(1) - 1))
                If InStr(gPath, "[System") <> 0 Then
                    Call MakeRootList
                    gPath = Me.lstTemp.List(CInt(sTempPara(1) - 1))
                End If
                If InStr(gPath, "����:") <> 0 Then  '����Ŀ¼Ϊ������
                    MakeList Left(gPath, 2) & "\"           '����Ŀ¼��
                End If
                '-------------------------------------
                If InStr(gPath, "Ŀ¼|") <> 0 Then  '����Ŀ¼Ϊ��ͨĿ¼
                    '�淶Ŀ¼��
                    gPath = Split(gPath, "|")(0)
                    gPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path & gPath, Me.File.Path & "\" & gPath)
                    MakeList gPath                          '����Ŀ¼��
                End If
                '-------------------------------------
                tmpSendString = "GPH|"                          '����ͷ
                For k = 0 To Me.lstTemp.ListCount - 1           '�����б��
                    tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                Next k
                '==========
                wsSendData Me.wsMessage, tmpSendString          '��������
                
                
            Case "GUD"                      '��ȡ�ϲ�Ŀ¼���ļ����У�
                Dim sPath As String                 'Ŀ¼����
                Dim tmpSplitPath() As String        '�и�Ŀ¼·������
                '-------------------------------------
                If Right(Me.Dir.Path, 2) <> ":\" Then               '�����ǰ���ǴӸ�Ŀ¼����һ���Ļ�
                    tmpSplitPath = Split(Me.File.Path, "\")             '��ʾ��һ��Ŀ¼
                    For k = 0 To UBound(tmpSplitPath) - 1
                        sPath = sPath & tmpSplitPath(k) & "\"
                    Next k
                    Call MakeList(sPath)                            '����Ŀ¼�б�
                    '-------------------------------------
                    tmpSendString = "GPH|"                          '����ͷ
                    For k = 0 To Me.lstTemp.ListCount - 1           '�����б��
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '��������
                Else
                    Call MakeRootList                               '���ɴ��̸�Ŀ¼�б�
                    '-------------------------------------
                    tmpSendString = "GRT|"                          '����ͷ
                    For k = 0 To Me.lstTemp.ListCount - 1           '�����б��
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '��������
                End If
            
            Case "MKD"                      '�½��ļ���
                Dim tmpPath As String                               '�ļ���·������
                sTempPara = Split(sTemp(i), "|")                    '�ָ������
                tmpPath = IIf(Right(Me.Dir.Path, 1) = "\", Me.Dir.Path, Me.Dir.Path & "\") & sTempPara(1)
                MkDir tmpPath
                If Err.Number = 0 Then                              '���ҿ�����û�г�����~
                    wsSendData Me.wsMessage, "MKD|1"                    '���ͳɹ�����
                Else
                    wsSendData Me.wsMessage, "MKD|0"                    '����ʧ�ܷ���
                End If
            
            Case "REF"                      'ˢ��
                Me.Dir.Refresh
                Me.File.Refresh
                Me.Drive.Refresh
                '------------------------
                If InStr(Me.lstTemp.List(0), "����:") <> 0 Then     '����Ǹ�Ŀ¼�б�
                    Call MakeRootList                               '���ɴ��̸�Ŀ¼�б�
                    '-------------------------------------
                    tmpSendString = "GRT|"                          '����ͷ
                    For k = 0 To Me.lstTemp.ListCount - 1           '�����б��
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '��������
                Else
                    Call MakeList(sPath)                            '����Ŀ¼�б�
                    '-------------------------------------
                    tmpSendString = "GPH|"                          '����ͷ
                    For k = 0 To Me.lstTemp.ListCount - 1           '�����б��
                        tmpSendString = tmpSendString & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, tmpSendString          '��������
                End If
            
            Case "REN"                      '������
                Dim OldName As String                               '������
                Dim NewName As String                               '������
                '------------------------
                sTempPara = Split(sTemp(i), "|")                    '�ָ������
                OldName = IIf(Right(Me.Dir.Path, 1) = "\", Me.Dir.Path, Me.Dir.Path & "\")  '���ɻ����ַ���
                NewName = OldName
                OldName = OldName & Split(Me.lstTemp.List(CInt(sTempPara(1)) - 1), "|")(0)
                NewName = NewName & sTempPara(2)
                Name OldName As NewName                             '������
                If Err.Number = 0 Then                              '���ҿ�����û�г�����~
                    wsSendData Me.wsMessage, "REN|1"                    '���ͳɹ�����
                Else
                    wsSendData Me.wsMessage, "REN|0"                    '����ʧ�ܷ���
                End If
            
            Case "CPY"                      '����
                Dim sIndex() As String                              '�ָ������Index����
                Dim cPath As String                                 'ѡ����ļ���·��
                '------------------------
                Me.lstClipboard.Clear                               '��ա������塱
                isCopy = True                                       '����ģʽ
                sIndex = Split(Replace(sTemp(i), "CPY|", ""), "|")
                cPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                For k = 0 To UBound(sIndex) - 1
                    Me.lstClipboard.AddItem cPath & Me.lstTemp.List(sIndex(k) - 1)    '��ӵ��������塱��
                Next k
            
            Case "CUT"                      '����
                Dim uIndex() As String                              '�ָ������Index����
                Dim uPath As String                                 'ѡ����ļ���·��
                '------------------------
                Me.lstClipboard.Clear                               '��ա������塱
                isCopy = False                                      '����ģʽ
                uIndex = Split(Replace(sTemp(i), "CPY|", ""), "|")
                uPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                For k = 0 To UBound(uIndex) - 1
                    Me.lstClipboard.AddItem uPath & Me.lstTemp.List(uIndex(k) - 1)    '��ӵ��������塱��
                Next k
            
            Case "PST"                      'ճ��
                Dim tmpList As String                               '�б���������ݻ���
                Dim tmpFilePath As String                           '�������塱�е�ÿ���ļ���·��
                Dim tmpFileTitle() As String                        '�ļ�������
                Dim tmpTargetPath As String                         'Ŀ��·��
                Dim tmpTargetDirPath As String                      'Ŀ�����ڵ��ļ���·��
                '------------------------
                tmpTargetDirPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")
                For k = 0 To Me.lstClipboard.ListCount - 1
                    tmpList = Me.lstClipboard.List(k)
                    tmpFilePath = Split(tmpList, "|")(0)                                '��ȡԴ�ļ�·��
                    tmpFileTitle = Split(tmpFilePath, "\")                                              '�ָ���ļ���
                    tmpTargetPath = tmpTargetDirPath & tmpFileTitle(UBound(tmpFileTitle))                     '����Ŀ��·��
                    '======================================================
                    If isCopy = True Then                       '����ģʽ
                        If InStr(tmpList, "|Ŀ¼") <> 0 Then            'Ŀ¼����xcopy����
                            Shell "xcopy " & Chr(34) & tmpFilePath & Chr(34) & " " & Chr(34) & tmpTargetPath & Chr(34) & " /e /c /i /h /y", vbHide
                        End If
                        If InStr(tmpList, "|�ļ�") <> 0 Then            '�ļ�����FileCopy����
                            FileCopy tmpFilePath, tmpTargetPath
                        End If
                    Else                                        '����ģʽ
                        Shell "cmd /c move /y " & Chr(34) & tmpFilePath & Chr(34) & " " & Chr(34) & tmpTargetPath & Chr(34), vbHide
                    End If
                Next k
                
            Case "DEL"                      'ɾ��
                Dim dIndex() As String                              '�ָ������Index����
                Dim dPath As String                                 'ѡ����ļ���·��
                Dim dList As String                                 'Ҫɾ�����ļ�·������
                '------------------------                                '����ģʽ
                dIndex = Split(Replace(sTemp(i), "DEL|", ""), "|")
                dPath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")     '�淶��·����
                For k = 0 To UBound(dIndex) - 1
                    dList = Me.lstTemp.List(dIndex(k) - 1)
                    If InStr(dList, "|Ŀ¼") <> 0 Then         '�����Ŀ¼
                        If bDeleteFiles Then
                            Shell "cmd /c rd " & Chr(34) & dPath & Split(dList, "|")(0) & Chr(34) & " /s /q", vbHide
                        Else
                            MsgBox "cmd /c rd " & Chr(34) & dPath & Split(dList, "|")(0) & Chr(34) & " /s /q"
                        End If
                    End If
                    If InStr(dList, "|�ļ�") <> 0 Then         '������ļ�
                        If bDeleteFiles Then
                            Kill dPath & Split(dList, "|")(0)
                        Else
                            MsgBox dPath & Split(dList, "|")(0)
                        End If
                    End If
                Next k
            
            Case "NPH"                      '��ȡ��ǰ�����·��
                Dim sNowPath As String      'Ҫ���͵�·���Ļ���
                '------------------------
                If InStr(Me.lstTemp.List(0), "����:") <> 0 Then     '����Ǹ�Ŀ¼�б�
                    sNowPath = "��Ŀ¼"             '���ء���Ŀ¼���ַ���
                Else
                    sNowPath = Me.Dir.Path          '�������Ŀ¼�б�ǰ��·��
                End If
                wsSendData Me.wsMessage, "NPH|" & sNowPath          '����·��
            
            Case "VPH"                      '���ָ��·������
                Dim sViewPath As String     '�����Ŀ¼
                Dim svSend As String        'Ҫ���͵�����
                '------------------------
                sViewPath = Replace(sTemp(i), "VPH|", "")           'ȥ������ͷ�������·��
                '------------------------
                If sViewPath = "��Ŀ¼" Then                        '���������Ǹ�Ŀ¼�б�
                    Call MakeRootList                                   '���ɸ�Ŀ¼��
                    svSend = "GRT|"                                     '����ͷ
                    For k = 0 To Me.lstTemp.ListCount - 1               '�����б��
                        svSend = svSend & Me.lstTemp.List(k) & "||"
                    Next k
                    wsSendData Me.wsMessage, svSend                     '��������
                Else                                                '���������������ļ���·��
                    If Right(sViewPath, 1) <> "\" Then
                        sViewPath = sViewPath & "\"                     '�淶��·����
                    End If
                    If IsPathExists(sViewPath) = False Then                 '��⵽�����Ŀ¼��Ч
                        wsSendData Me.wsMessage, "VPH|ERR"                  '���ʹ�����Ϣ
                    Else                                                '��⵽�����Ŀ¼����
                        MakeList sViewPath                                  '����Ŀ¼��
                        '-------------------------------------
                        svSend = "GPH|"                                     '����ͷ
                        For k = 0 To Me.lstTemp.ListCount - 1               '�����б��
                            svSend = svSend & Me.lstTemp.List(k) & "||"
                        Next k
                        '==========
                        wsSendData Me.wsMessage, svSend                     '��������
                    End If
                End If
                
            '===================================================================================================
            '===================================================================================================
            '����������
            Case "BLK"                          '����
                If bBlockInput = True Then
                    Me.tmrBlockInput.Enabled = True
                End If
                
            Case "ULK"                          '����
                If bBlockInput = True Then
                    Me.tmrBlockInput.Enabled = False
                    BlockInput False
                End If
            
            '===================================================================================================
            '===================================================================================================
            '�ļ�����
            Case "UPL"                          '�ϴ��ļ�����
                Dim sFileName As String, lFileSize As Long              '�ļ��� & �ļ���С
                '----------------------------------------
                IsUpload = True                                         '�ϴ�״̬
                FileRec = 0                                             '�����ļ������
                Me.tmrCalcPerSecond.Enabled = True                      '��������ÿ���ֽ����ļ�ʱ��
                sFileName = Split(sTemp(i), "|")(1)                     '������ļ���
                lFileSize = CLng(Split(sTemp(i), "|")(2))               '������ļ���С
                Me.wsFile.Tag = sFileName & "|" & lFileSize & "|" & Me.File.Path            '��ֵTag ��Դ·��|�ļ���С|д��Ŀ¼��·����
                Me.wsFile.Close                                         '��ʼ��Winsock
                Me.wsFile.Connect Me.wsMessage.RemoteHostIP, 20236      '���ӵ��ļ����Ͷ�
            
            Case "NXT"                          '�ϴ���һ���ļ�����
                Dim nFileName As String, nFileSize As Long              '�ļ��� & �ļ���С
                '----------------------------------------
                nFileName = Split(sTemp(i), "|")(1)                     '������ļ���
                nFileSize = CLng(Split(sTemp(i), "|")(2))               '������ļ���С
                Me.wsFile.Tag = nFileName & "|" & nFileSize & "|" & Split(Me.wsFile.Tag, "|")(2)    '��ֵTag ��Դ·��|�ļ���С|д��Ŀ¼��·����
                Me.wsFile.SendData "NNT"                                '���߶Է����Ѿ�׼���ý�����һ���ļ�����
            
            Case "END"                          '�������
                Close #2                                                '�ر��ļ�����ֹ���ļ����ǹر�
                With frmBeingControlled
                    .edStatus1.Text = "�������"
                    .edStatus2.Text = "�����յ�" & FileRec & "���ļ�"
                    .edStatus3.Visible = False
                End With
                Me.tmrCalcPerSecond.Enabled = False                     '�رռ���ÿ���ֽ����ļ�ʱ��
            
            Case "DWL"                          '�����ļ�����
                Dim iFileIndex() As String          '�����ļ������
                Dim iFilePath As String             '��ǰѡ����ļ���·��
                '-------------------------------------------------------
                IsUpload = False                                                                    '����״̬
                iFileIndex = Split(Replace(sTemp(i), "DWL|", ""), vbCrLf)                           '�ָ��ÿ���ļ���Index
                '---------------------------
                CurrentFile = -1                                                                    'һ��ʼ�ļ����Ϊ-1
                TotalFiles = 0                                                                      '�ļ�������Ϊ0
                lSent = 0                                                                           '��շ��������ֽ���
                oldSent = 0
                ReDim EachFile(0)                                                                   '���ÿ���ļ��Ļ����б�
                FileList = ""                                                                       '����ļ��б�
                '---------------------------
                TotalFiles = UBound(iFileIndex) - 1                                                 '�����ļ�����
                For k = 0 To UBound(iFileIndex) - 1                                                 '��ȥһ����Ϊ�������һ��Split���µĿ�����
                    iFilePath = IIf(Right(Me.File.Path, 1) = "\", Me.File.Path, Me.File.Path & "\")     '�淶��·����
                    iFilePath = iFilePath & Me.lstTemp.List(iFileIndex(k) - 1)                          '������ļ������ļ���С
                    FileList = FileList & iFilePath & vbCrLf                                            '��ӵ��ļ��б���
                Next k
                '-----------------
                EachFile = Split(FileList, vbCrLf)                                                  '�����ÿ���ļ� ������·��|�ļ���С��
                For k = 0 To UBound(EachFile) - 1                                                   '������������·��
                    EachFile(k) = Split(EachFile(k), "|")(0)
                Next k
                Call NextFile                           '��һ�ļ�
                Me.wsFile.Close                                         '��ʼ��Winsock
                Me.wsFile.Connect Me.wsMessage.RemoteHostIP, 20236      '���ӵ��ļ����ն�
            
        End Select
    Next i
End Sub

Private Sub wsPic_Close()
    Me.wsPic.Close                  '�Ͽ��Զ����¼���
    Me.wsPic.Bind 20234
    Me.wsPic.Listen
End Sub

Private Sub wsPic_ConnectionRequest(ByVal requestID As Long)
    Me.wsPic.Close                  '��������
    Me.wsPic.Accept requestID
End Sub

Private Sub wsMessage_ConnectionRequest(ByVal requestID As Long)
    Me.wsMessage.Close              '��������
    Me.wsMessage.Accept requestID
    IsRemoteControl = Me.optControl.Value       '��ȡԶ��ģʽ��Զ�̿���ģʽ or �ļ�����ģʽ
    Me.mnuShowWindow.Enabled = False
End Sub

Private Sub wsPic_DataArrival(ByVal bytesTotal As Long)
    Dim dat() As Byte
    Me.wsPic.GetData dat, vbArray Or vbByte
    If bytesTotal - 1 = 1 Then                  '��Ϣһ�����Ϳ�ͼ
        SendDat
    End If
    If bytesTotal - 1 = 2 Then                  '��Ϣ������ʼ��������ͼƬ
        Me.tmrForceRefresh.Enabled = True           '��ʼ����ʹ��Ļˢ��
        Call SendLoop                               '������ͼ
    End If
End Sub
