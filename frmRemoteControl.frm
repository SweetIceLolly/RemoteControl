VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmRemoteControl 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Զ�̿���"
   ClientHeight    =   5592
   ClientLeft      =   120
   ClientTop       =   456
   ClientWidth     =   10860
   Icon            =   "frmRemoteControl.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5592
   ScaleWidth      =   10860
   StartUpPosition =   2  '��Ļ����
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
      Item(0).Caption =   "��Ļ����"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "ResizerPicture"
      Item(1).Caption =   "���̹���"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "lstTask"
      Item(1).Control(1)=   "labTaskTotal"
      Item(1).Control(2)=   "cmdRefreshTasks"
      Item(1).Control(3)=   "cmdKillTasks"
      Item(2).Caption =   "�ļ�����"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "lstFile"
      Item(2).Control(1)=   "labTip"
      Item(2).Control(2)=   "edPath"
      Item(2).Control(3)=   "cmdOpenDir"
      Item(3).Caption =   "����VBS�ű�"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "edVBS"
      Item(3).Control(1)=   "cmdRunVBS"
      Item(4).Caption =   "������"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "edCommandLine"
      Item(4).Control(1)=   "cmdRunCommandline"
      Item(4).Control(2)=   "edCommandlineEnter"
      Item(5).Caption =   "ϵͳ��Ϣ"
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
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "����ˢ��~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "���ҽ���~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "����ִ��~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "����ִ��~"
         ForeColor       =   16777215
         BackColor       =   16416003
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "ת��"
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
         Text            =   "��Ŀ¼"
         BackColor       =   16777152
         Appearance      =   4
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ǰĿ¼��"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "��0�����̡�"
         BeginProperty Font 
            Name            =   "����"
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
         ToolTipText     =   "�����Է���������"
         Top             =   0
         Width           =   2600
         _Version        =   786432
         _ExtentX        =   4586
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "����Զ��������"
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
         ToolTipText     =   "�������ƶԷ���������"
         Top             =   0
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4586
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "�������ƶԷ�������"
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
         ToolTipText     =   "�Զ�������Ļ����"
         Top             =   0
         Width           =   2595
         _Version        =   786432
         _ExtentX        =   4586
         _ExtentY        =   1270
         _StockProps     =   79
         Caption         =   "�Զ�������Ļ����"
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
         Caption         =   "ˢ��(&R)"
      End
      Begin VB.Menu mnuMkdir 
         Caption         =   "�½��ļ���(&N)"
      End
      Begin VB.Menu mnuRename 
         Caption         =   "������(&M)"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "����(&C)"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "����(&U)"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "ճ��(&P)"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "ɾ��(&D)"
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
    biSize As Long                                                              'BITMAPINFOHEADER�ṹ�Ĵ�С
    biWidth As Long
    biHeight As Long
    biPlanes As Integer                                                         '�豸��Ϊƽ���������ڶ���1
    biBitCount As Integer                                                       'ͼ�����ɫλͼ
    biCompression As Long                                                       'ѹ����ʽ
    biSizeImage As Long                                                         'ʵ�ʵ�λͼ������ռ�ֽ�
    biXPelsPerMeter As Long                                                     'Ŀ���豸��ˮƽ�ֱ���
    biYPelsPerMeter As Long                                                     'Ŀ���豸�Ĵ�ֱ�ֱ���
    biClrUsed As Long                                                           'ʹ�õ���ɫ��
    biClrImportant As Long                                                      '��Ҫ����ɫ�� �������Ϊ0����ʾ������ɫ������Ҫ��
End Type
  
Private Type RGBQUAD                                                            'ֻ��bibitcountΪ1��2��4ʱ���е�ɫ��
    Blue As Byte                                                                '��ɫ����
    Green As Byte                                                               '��ɫ����
    Red As Byte                                                                 '��ɫ����
    Reserved As Byte                                                            '����ֵ
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
'����Ϊzlibѹ���ĺ���
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Const OFFSET As Long = &H8

'����Ϊ��Ļ��ʾ������������Ļ����
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
'����Ϊ�������߰���ı���
Dim TwipsX, TwipsY As Single    '��Ļÿ�����ص�Twip��
Dim DataPerSecond As Long       'ÿ���ӵ�������
Dim DataPerSecondOld As Long    '��һ�μ�����ÿ����������
Dim sDataRec As String          '���ݷֶ��õĻ��棺��ʱ��������̫��Winsockǿ�Ʒְ�ʱ��Ҫ�õ�����~
Dim bMultiData As Boolean       '�Ƿ�Ϊ������ݷֶ�ģʽ

Function UnCompressByte(ByteArray() As Byte)                '��ѹ������
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

Public Sub Del_dc()                                         'ɾ���ڴ�DC����
    DeleteDC iDC1
    DeleteObject iBitmap1
    DeleteDC iDC2
    DeleteObject iBitmap2
End Sub

Public Function DCload()
    iDC1 = CreateCompatibleDC(Me.hdc)                                           '1
    iBitmap1 = CreateCompatibleBitmap(Me.hdc, W, H * F)                         '��ʾ��������
    SelectObject iDC1, iBitmap1
    
    iDC2 = CreateCompatibleDC(Me.hdc)                                           '2
    iBitmap2 = CreateCompatibleBitmap(Me.hdc, W, H)                             '��ʾ��������
    SelectObject iDC2, iBitmap2
End Function

Private Sub GetWindows()                                '��ȡ���յ�����Ļ
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
            UnCompressByte bBytes                                                   '��ѹ
        End If
        SetDIBitsToDevice iDC2, 0, 0, W, H, 0, 0, 0, H, bBytes(0), bi24BitInfo, DIB_RGB_COLORS
        BitBlt iDC1, 0, g, W, H, iDC2, 0, 0, vbSrcInvert                        'xor'�෴
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

Private Sub New_dc()                                    '�����ڴ�DC����
    Dim bi24BitInfo As BITMAPINFO, bBytes() As Byte
    
    Me.wsPicture.GetData bBytes, vbArray Or vbByte
    
    If bCompress = True Then
        UnCompressByte bBytes                                                       '��ѹ
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
    For i = 0 To frmMain.comIP.ListCount - 1                'ɨһ���������IP�б�
        If frmMain.comIP.List(i) = strIP Then                   '�ҵ�����ͬ��IP�Ļ�
            IsIPExists = i                                          '����Index
            Exit Function                                           '�˳�����
        End If
    Next i
    IsIPExists = -1                                      '�Ҳ�����ͬ�Ļ��ͷ���-1
End Function

'=======================================================================

Public Sub ConnectedHandler()                 '���Ӻ�Ĵ������
    Dim iLstIP As Integer                                             '�б��е�IPλ��
    Me.Caption = "Զ�̿���: " & Me.wsMessage.RemoteHostIP             '���ı���
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

Public Function GetListSelCount() As Integer    '��ȡ�б���ѡ�����Ŀ��
    Dim Total As Integer        'ѡ�����Ŀ��
    '=========================================
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '���˹��ľͼӽ�ȥ~
            Total = Total + 1
        End If
    Next i
    GetListSelCount = Total
End Function

Private Sub cmdAutoResize_Click()
    Me.cmdAutoResize.Checked = Not Me.cmdAutoResize.Checked             '��ѡ
    AutoResize = Me.cmdAutoResize.Checked                               '�����Ƿ����컭��״̬
    Call Form_Resize                                                    '���������Ű�
    Call SaveConfig                                                     '���������ļ�
End Sub

Private Sub cmdBlockInput_Click()
    Me.cmdBlockInput.Checked = Not Me.cmdBlockInput.Checked             '��ѡ
    If Me.cmdBlockInput.Checked = True Then                             '����״̬
        wsSendData Me.wsMessage, "BLK"                                      '����������������Ϣ
    Else
        wsSendData Me.wsMessage, "ULK"                                      '���ͽ�����������Ϣ
    End If
End Sub

Private Sub cmdIsControlling_Click()
    Me.cmdIsControlling.Checked = Not Me.cmdIsControlling.Checked       '��ѡ
    IsControlling = Me.cmdIsControlling.Checked                         '��������״̬
End Sub

Private Sub cmdKillTasks_Click()
    Dim SendTemp As String              'Ҫ���͵����ݻ���
    Dim IsChecked As Boolean            '�Ƿ���ѡ����
    '========================================
    IsChecked = False
    For i = 1 To Me.lstTask.ListItems.Count                 '�����б�����Ҫ�����ĸ�
        If Me.lstTask.ListItems(i).Checked = True Then
            '���뵽�������ݻ�����
            SendTemp = SendTemp & Me.lstTask.ListItems(i).Text & "|" & Me.lstTask.ListItems(i).SubItems(1) & vbCrLf
            IsChecked = True                                        '�з��͵Ķ���
        End If
    Next i
    If IsChecked = False Then                               'û��ѡ��������
        MsgBox "�㶼ûѡ��������ҽ���ʲô����", 64
    Else
        a = MsgBox("ȷ�Ͻ���ѡ��Ľ��̣������������ݶ�ʧ�Ⱥ����", 32 + vbYesNo, "ȷ��")      'ȷ�Ͽ�
        If a <> 6 Then
            Exit Sub
        End If
        wsSendData Me.wsMessage, "KTS|" & SendTemp          '�����ѡ����̾ͷ��ͽ�����������
    End If
End Sub

Private Sub cmdOpenDir_Click()
    wsSendData Me.wsMessage, "VPH|" & Me.edPath.Text    '���ʹ�Ŀ¼����
End Sub

Private Sub cmdRefreshTasks_Click()
    wsSendData Me.wsMessage, "TSK"                      '���ͻ�ȡ��������
End Sub

Private Sub cmdRunCommandline_Click()
    If Trim(Me.edCommandlineEnter.Text) = "" Then       '�����ܿ�����
        Exit Sub
    End If
    wsSendData Me.wsMessage, "CMD|" & Me.edCommandlineEnter.Text        '����ִ������������
    Me.edCommandlineEnter.Text = "ִ���У����Ժ�..."
    Me.edCommandlineEnter.Enabled = False
End Sub

Private Sub cmdRunVBS_Click()
    wsSendData Me.wsMessage, "VBS|" & Me.edVBS.Text     '����ִ��VBS����
End Sub

Private Sub edCommandlineEnter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then                   '���»س�����ִ������
        Call cmdRunCommandline_Click
        KeyAscii = 0
    End If
End Sub

Private Sub edPath_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdOpenDir_Click                   '���ð�ť���¹���
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.imgMainIcon.Picture       '����ͼ�꣬Ϊ�˸��ÿ� (*^__^*)
    '==================================================
    Me.lstTask.ColumnHeaders.Add , , "��������", 1500
    Me.lstTask.ColumnHeaders.Add , , "PID", 600
    Me.lstTask.ColumnHeaders.Add , , "����·��", 4000
    '==================================================
    Me.lstFile.ColumnHeaders.Add , , "����", 2000
    Me.lstFile.ColumnHeaders.Add , , "����", 800
    Me.lstFile.ColumnHeaders.Add , , "��С", 3000
    '==================================================
    Me.picKeyboard.Left = -10
    Me.picKeyboard.Top = -10
    Me.picKeyboard.Width = 1
    Me.picKeyboard.Height = 1
    IsControlling = True                        '����ģʽ
    '==================================================
    If AutoResize Then
        Me.cmdAutoResize.Checked = True
    End If
    '==================================================
    If bMouseWheelHook = True Then
        PrevWndProc = SetWindowLong(Me.picKeyboard.hwnd, GWL_WNDPROC, AddressOf WndProc)        '�����������¼�����
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
    '�������ֿؼ��Ĵ�С��״̬����Ӧ����
    Me.ResizerMain.Width = Me.Width - 240
    '==========================================
    Me.TabMain.Width = Me.Width - 240                               '��ѡ�
    Me.TabMain.Height = Me.Height - Me.ResizerMain.Height - 480
    '==========================================
    Me.ResizerPicture.Width = Me.TabMain.Width - 120                '��Ļ��ʾ����
    Me.ResizerPicture.Height = Me.TabMain.Height - 480
    If AutoResize = True Then                                       '������Զ����컭��Ļ�����ͼƬ����Ӧ����
        Me.imgMain.Width = Me.ResizerPicture.Width
        Me.imgMain.Height = Me.ResizerPicture.Height
        Me.ResizerPicture.HScrollMaximum = 0                            '���ӹ�����
        Me.ResizerPicture.VScrollMaximum = 0
    Else                                                            '����������Ӧԭͼ
        Me.imgMain.Width = Me.picMain.Width
        Me.imgMain.Height = Me.picMain.Height
        Me.ResizerPicture.HScrollMaximum = Me.picMain.Width             '��ӹ�����
        Me.ResizerPicture.VScrollMaximum = Me.picMain.Height
    End If
    ScreenScaleW = Me.imgMain.Width / (W * TwipsX)                  '����ѹ��ͼ��ԭͼ�ı�����
    ScreenScaleH = Me.imgMain.Height / (H * F * TwipsY)
    '==========================================
    Me.lstTask.Width = Me.TabMain.Width - 240                       '���̲���
    Me.lstTask.Height = Me.TabMain.Height - 1000
    Me.cmdRefreshTasks.Top = Me.lstTask.Top + Me.lstTask.Height + 120
    Me.cmdKillTasks.Top = Me.cmdRefreshTasks.Top
    Me.labTaskTotal.Top = Me.cmdRefreshTasks.Top
    '==========================================
    Me.cmdOpenDir.Left = Me.TabMain.Width - Me.cmdOpenDir.Width - 120   '�ļ�������
    Me.edPath.Width = Me.cmdOpenDir.Left - Me.edPath.Left - 120
    Me.lstFile.Height = Me.TabMain.Height - 240 - Me.lstFile.Top
    Me.lstFile.Width = Me.TabMain.Width - 240
    '==========================================
    Me.edVBS.Width = Me.TabMain.Width - 240                         'VBS���벿��
    Me.edVBS.Height = Me.TabMain.Height - 1000
    Me.cmdRunVBS.Top = Me.cmdRefreshTasks.Top
    '==========================================
    Me.edInformation.Width = Me.TabMain.Width - 120                 'ϵͳ��Ϣ����
    Me.edInformation.Height = Me.TabMain.Height - 480
    '==========================================
    Me.edCommandLine.Width = Me.TabMain.Width - 120                 '�����в�������
    Me.edCommandLine.Height = Me.TabMain.Height - 1000
    Me.edCommandlineEnter.Width = Me.TabMain.Width - 360 - Me.cmdRunCommandline.Width
    Me.edCommandlineEnter.Top = Me.lstTask.Top + Me.lstTask.Height + 120
    Me.cmdRunCommandline.Top = Me.cmdRefreshTasks.Top
    Me.cmdRunCommandline.Left = Me.edCommandlineEnter.Left + Me.edCommandlineEnter.Width + 120
End Sub

Private Sub Form_Unload(Cancel As Integer)      '�رմ��弴�Ͽ�����
    Me.wsMessage.Close
    Me.wsPicture.Close
    Dim OpenPath As String                  '���������·��
    OpenPath = IIf(Right(App.Path, 1) = "\", App.Path & App.EXEName & ".exe", App.Path & "\" & App.EXEName & ".exe")
    Shell OpenPath, vbNormalFocus               '���۲���ǧ���� ����ǰͷ��ľ��
    End                                         '�����ˣ��������ˣ�����·�ˡ�
End Sub

Private Sub lstFile_DblClick()
    '��ȡ���Ͳ������ͷ�������
    On Error Resume Next
    Select Case Me.lstFile.ListItems.Item(Me.lstFile.SelectedItem.Index).SubItems(1)
        Case "������"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFile.SelectedItem.Index
            
        Case "Ŀ¼"
            wsSendData Me.wsMessage, "GPH|" & Me.lstFile.SelectedItem.Index
            
        Case "�ϲ�Ŀ¼"
            wsSendData Me.wsMessage, "GUD"
            
    End Select
End Sub

Private Sub lstFile_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
    On Error Resume Next
    For i = 1 To Me.lstFile.ListItems.Count
        '����Ǵ��� ���� �ϲ�Ŀ¼ѡ��
        If InStr(Me.lstFile.ListItems(i).SubItems(2), "����:") <> 0 Or Me.lstFile.ListItems(i).SubItems(1) = "�ϲ�Ŀ¼" Then
            Me.lstFile.ListItems(i).Checked = False
        End If
    Next i
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sCount As Integer                           'ѡ���˵��б�����
    '----------------------------------
    If Button = 2 Then                              '�����Ҽ������˵�
        sCount = GetListSelCount
        Me.mnuMkdir.Visible = True
        Me.mnuPaste.Visible = True
        If sCount = 0 Then                     '�����ε����˵�
            Me.mnuCopy.Visible = False                  'û��ѡ���б����� ���ơ����С�ɾ��������������ʧЧ
            Me.mnuCut.Visible = False
            Me.mnuDelete.Visible = False
            Me.mnuRename.Visible = False
        Else
            Me.mnuCopy.Visible = True
            Me.mnuCut.Visible = True
            Me.mnuDelete.Visible = True
            Me.mnuRename.Visible = True
        End If
        If sCount > 1 Then                              '���ѡ���˶���һ���б����� ����������ʧЧ
            Me.mnuRename.Visible = False
        End If
        If InStr(Me.lstFile.ListItems(1).ListSubItems(2).Text, "����:") <> 0 Then '����Ǵ��̸�Ŀ¼��ֻ����ˢ��
            Me.mnuMkdir.Visible = False
            Me.mnuPaste.Visible = False
        End If
        PopupMenu Me.mnuPopup
    End If
End Sub

Private Sub mnuCopy_Click()
    Dim tmpSendString As String                     '�������ݻ���
    '----------------------------------
    tmpSendString = "CPY|"
    Me.mnuPaste.Enabled = True
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '������������Ӵ��˹����б���
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '���͸����ļ�����
End Sub

Private Sub mnuCut_Click()
    Dim tmpSendString As String                     '�������ݻ���
    '----------------------------------
    tmpSendString = "CUT|"
    Me.mnuPaste.Enabled = True
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '������������Ӵ��˹����б���
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '���ͼ����ļ�����
End Sub

Private Sub mnuDelete_Click()
    Dim tmpSendString As String                     '�������ݻ���
    '----------------------------------
    If GetListSelCount = 0 Then
        Exit Sub
    End If
    a = MsgBox("�����Ҫɾ��ѡ���" & GetListSelCount & "���ļ����У��𣿸ò������ɳ�����", 32 + vbYesNo, "ȷ��")
    If a <> 6 Then
        Exit Sub
    End If
    tmpSendString = "DEL|"
    For i = 1 To Me.lstFile.ListItems.Count
        If Me.lstFile.ListItems(i).Checked = True Then      '������������Ӵ��˹����б���
            tmpSendString = tmpSendString & i & "|"
        End If
    Next i
    wsSendData Me.wsMessage, tmpSendString          '����ɾ���ļ�����
End Sub

Private Sub mnuMkdir_Click()
    Dim fName As String                             '�ļ������ֻ���
    fName = InputBox("���������ļ��е����֣�", "����")
    If Trim(fName) <> "" Then
        If InStr(fName, "|") = 0 Then
            wsSendData Me.wsMessage, "MKD|" & fName         '���ʹ����ļ�������
        Else
            MsgBox "�����ļ���ʧ�ܡ�", 64, "�����ļ���ʧ��"
        End If
    End If
End Sub

Private Sub mnuPaste_Click()
    wsSendData Me.wsMessage, "PST"                  '����ճ������
End Sub

Private Sub mnuRefresh_Click()
    wsSendData Me.wsMessage, "REF"                  '����ˢ������
End Sub

Private Sub mnuRename_Click()
    Dim tmpName As String                           '�ļ�������
    tmpName = InputBox("�������µ��ļ�����", "������")
    If Trim(tmpName) <> "" Then
        If InStr(tmpName, "|") = 0 Then
            wsSendData Me.wsMessage, "REN|" & Me.lstFile.SelectedItem.Index & "|" & tmpName         '�����ļ�����������
        Else
            MsgBox "�ļ�������ʧ�ܡ�", 64, "�ļ�������ʧ��"
        End If
    End If
End Sub

Private Sub imgMain_DblClick()
    If IsControlling Then
        wsSendData Me.wsMessage, "DBC"                  '˫�����
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
            Case 1                                  '����
                wsSendData Me.wsMessage, "MLD"
                
            Case 2                                  '����
                wsSendData Me.wsMessage, "MRD"
            
            Case 4                                  '����
                wsSendData Me.wsMessage, "MMD"
            
        End Select
    End If
    Me.picKeyboard.SetFocus
End Sub

Private Sub imgMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim rX, rY As Single                        '��ԭͼ�ϵ�X�����Y����
    If IsControlling Then
        rX = x / ScreenScaleW / TwipsX
        rY = y / ScreenScaleH / TwipsY
        wsSendData Me.wsMessage, "MMP|" & CStr(rX) & "|" & CStr(rY)      '��������ƶ�������
    End If
End Sub

Private Sub imgMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsControlling Then
        Select Case Button
            Case 1                                  '����
                wsSendData Me.wsMessage, "MLU"
                
            Case 2                                  '����
                wsSendData Me.wsMessage, "MRU"
            
            Case 4                                  '����
                wsSendData Me.wsMessage, "MMU"
            
        End Select
    End If
End Sub

Private Sub tmrDataPerSecond_Timer()
    '��ʾ��ÿ���ӵ�������
    Dim Calc As Long
    Calc = DataPerSecond - DataPerSecondOld         '��ȥ
    '���ı���
    Me.Caption = "Զ�̿���: " & Me.wsMessage.RemoteHostIP & " ��ǰ������" & Format((Abs(Calc)) / 1024, "0.00") & "kb/s"
    DataPerSecondOld = DataPerSecond                        '���۲���ǧ���� ����ǰͷ��ľ��
    DataPerSecond = 0                                       '���ÿ���ӵ�������
End Sub

Private Sub wsMessage_Close()
    Unload Me
End Sub

Private Sub wsMessage_Connect()
    frmMain.tmrTimeOut.Enabled = False          '���ӳɹ���ֹͣ��ʱ��ʱ��
    Call ConnectedHandler
End Sub

Private Sub wsMessage_DataArrival(ByVal bytesTotal As Long)
    Dim Temp As String                      '��������
    Dim sTemp() As String                   '�����и������
    Dim sTempPara() As String               '���ݸ����Ĳ����и�����Ļ�������
    Dim TaskSplit() As String               'ÿһ��������Ϣ�ķָ�棨���ڽ��̷���ģ�
    '====================================
    'ͳ��������
    DataPerSecond = DataPerSecond + bytesTotal
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
                Me.wsPicture.Close
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
                        wsSendData Me.wsMessage, "SIF"            '���ͻ�ȡϵͳ��Ϣ����
                        
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
            
            Case "RES"                                      '�Է��ֱ��ʷ���
                sTempPara = Split(sTemp(i), "|")
                '������Ļ�������
                W = CLng(sTempPara(1))                      '����Ļ��С���������Ͻ��յ��ĶԷ�����Ļ�ķֱ���
                H = CLng(sTempPara(2))
                TwipsX = CSng(sTempPara(3))
                TwipsY = CSng(sTempPara(4))
                '--------------------------------------
                '�����ڴ�λͼ��Ϣ
                F = 6                                                                       '�ֶ�����
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
                '�����ڴ�DC
                Call Del_dc
                Call DCload
                '����ͼƬ��Ŀ��
                Me.picMain.Width = W * TwipsX
                Me.picMain.Height = H * F * TwipsY
                Me.ResizerPicture.HScrollMaximum = Me.picMain.Width
                Me.ResizerPicture.VScrollMaximum = Me.picMain.Height
                Me.tmrDataPerSecond.Enabled = True                      '��������ͳ�Ƽ�ʱ��
                'һ�о���������Ļ��������
                wsSendData Me.wsMessage, "BEG"
                
            Case "SIF"                                      'ϵͳ��Ϣ
                sTempPara = Split(sTemp(i), "|")
                Me.edInformation.Text = sTempPara(1) & vbCrLf & vbCrLf & "������Ϣ�����ο���һ������ʵ��Ϊ׼��"
                '============================================
                '��ȡ��ϵͳ��Ϣ���ȡ��Ŀ¼
                If Me.lstFile.ListItems.Count = 0 Then
                    wsSendData Me.wsMessage, "GRT"
                End If
                
            Case "TSK"
                '-------------------------------------
                Me.lstTask.ListItems.Clear                  '����б�
                sTempPara = Split(sTemp(i), vbCrLf)         '�ָ��ÿ���س������������
                sTempPara(0) = Replace(sTempPara(0), "TSK|", "")
                For k = 0 To UBound(sTempPara) - 2            '��ȡÿһ�е�����
                    TaskSplit = Split(sTempPara(k), "|")    '�ָ��ÿ���ָ�������������
                    Me.lstTask.ListItems.Add , , TaskSplit(0)                                       '������
                    Me.lstTask.ListItems(Me.lstTask.ListItems.Count).SubItems(1) = TaskSplit(1)     '����PID
                    Me.lstTask.ListItems(Me.lstTask.ListItems.Count).SubItems(2) = TaskSplit(2)     '����·��
                Next k
                Me.labTaskTotal.Caption = "��" & Me.lstTask.ListItems.Count + 1 & "�����̡�"         '��ʾ��������
            
            Case "KTS"                                      '�������̵ķ�����Ϣ
                Dim tmpString As String                     '��ʾ��Ϣ����
                '-------------------------------------
                sTempPara = Split(Replace(Replace(sDataRec, "{S}", ""), "KTS|", ""), "||")  'ȥ���������ָ�������зָ�
                tmpString = "�������̷�����"
                For k = 0 To UBound(sTempPara) - 1                  '�����ֳ���������
                    TaskSplit = Split(sTempPara(k), "|")
                    tmpString = tmpString & vbCrLf & "��������" & TaskSplit(0) & "��PID��" & TaskSplit(1) & _
                                            "�� ��" & IIf(TaskSplit(2) = "T", "�ɹ�", "ʧ��") & "��"
                Next k
                MsgBox tmpString, 64, "�������̷���"
                Call cmdRefreshTasks_Click                          '���յ���������ˢ��
            
            Case "CMD"                                      '������ִ�еķ�����Ϣ
                Dim tmpResult As String                     '��ʾ��Ϣ����
                '-------------------------------------
                tmpResult = Replace(Replace(sDataRec, "CMD|", ""), "{S}", "")
                Me.edCommandLine.Text = tmpResult
                Me.edCommandlineEnter.Enabled = True
                Me.edCommandlineEnter.Text = ""
                Me.edCommandlineEnter.SetFocus
                
            '===================================================================================================
            '===================================================================================================
            '�ļ�������
            Case "GRT"                                      '��ȡ��Ŀ¼������Ӧ
                Dim sRootList() As String                       '�б����и��
                Dim sRootMsg() As String                        '���б������ϸ��Ϣ�и��
                '-------------------------------------
                sRootList() = Split(Replace(sDataRec, "GRT|", ""), "||")    '�и��ַ���
                Me.lstFile.ListItems.Clear
                For k = 0 To UBound(sRootList) - 1
                    sRootMsg = Split(sRootList(k), "|")
                    Me.lstFile.ListItems.Add , , sRootMsg(0)
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(1) = "������"
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(2) = sRootMsg(1)
                Next k
                Me.edPath.Text = "��Ŀ¼"                       '��ʾ·����Ϊ����Ŀ¼��
            
            Case "GPH"                                      '��ȡĿ¼���ļ�������Ӧ
                Dim sDirList() As String                        '�б����и��
                Dim sDirMsg() As String                         '���б������ϸ��Ϣ�и��
                On Error Resume Next
                '-------------------------------------
                sDirList = Split(Replace(sDataRec, "GPH|", ""), "||")       '�и��ַ���
                Me.lstFile.ListItems.Clear
                For k = 0 To UBound(sDirList) - 1
                    sDirMsg = Split(sDirList(k), "|")
                    Me.lstFile.ListItems.Add , , sDirMsg(0)
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(1) = sDirMsg(1)
                    sDirMsg(2) = Replace(sDirMsg(2), ":N:", "")         '����Ƿ�ֹ�ļ��л�ȡ���˴�С���µĿհ����������
                    Me.lstFile.ListItems(Me.lstFile.ListItems.Count).SubItems(2) = sDirMsg(2)
                Next k
                wsSendData Me.wsMessage, "NPH"                  '���ͻ�ȡ��ǰĿ¼·��������
            
            Case "MKD"
                Dim bSuccess As String
                bSuccess = Replace(Replace(sDataRec, "MKD|", ""), "{S}", "")
                If bSuccess = "0" Then
                    MsgBox "�����ļ���ʧ�ܡ�", 64, "�����ļ���ʧ��"
                Else
                    MsgBox "�����ļ��гɹ���", 64, "�����ļ��гɹ�"
                End If
                wsSendData Me.wsMessage, "REF"                  '����ˢ������
                
            Case "REN"
                Dim brSuccess As String
                brSuccess = Replace(Replace(sDataRec, "REN|", ""), "{S}", "")
                If brSuccess = "0" Then
                    MsgBox "�ļ�������ʧ�ܡ�", 64, "�ļ�������ʧ��"
                Else
                    MsgBox "�ļ��������ɹ���", 64, "�ļ��������ɹ�"
                End If
                wsSendData Me.wsMessage, "REF"                  '����ˢ������
            
            Case "NPH"
                Me.edPath.Text = Replace(Replace(sDataRec, "NPH|", ""), "{S}", "")      '��ʾ���ص��ļ���·��
            
            Case "VPH"
                MsgBox "Ŀ¼����Ч��", 48, "��Ϣ"                '��ʾ������Ϣ
                wsSendData Me.wsMessage, "NPH"                  '���ͻ�ȡ��ǰĿ¼·��������
            
        End Select
    Next i
    '-----------------------------
    sDataRec = ""                       'ִ��������������������
    bMultiData = False
End Sub

Private Sub wsPicture_Connect()
    frmMain.tmrTimeOut.Enabled = False          '���ӳɹ���ֹͣ��ʱ��ʱ��
End Sub

Private Sub wsPicture_DataArrival(ByVal bytesTotal As Long)
    Dim dat() As Byte
    'ͳ��������
    DataPerSecond = DataPerSecond + bytesTotal
    '��������
    If bytesTotal = 8 Then                                                      '[ͷ] �����ݽ�������
        Me.wsPicture.GetData dat, vbArray Or vbByte
        xData = StrConv(dat, vbUnicode)
        J = Mid$(xData, 1, 1)
        If J = 9 Then J = 0
        xData = Mid$(xData, 2, 7) + 1
        If J <> 8 Then Me.wsPicture.SendData 0
    End If
    If bytesTotal = xData Then Call GetWindows
End Sub
