VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Begin VB.Form frmBeingControlled 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   0  'None
   ClientHeight    =   3468
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4536
   ControlBox      =   0   'False
   Icon            =   "frmBeingControlled.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3468
   ScaleWidth      =   4536
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer tmrChangeIcon 
      Interval        =   10
      Left            =   4080
      Top             =   2760
   End
   Begin XtremeSuiteControls.GroupBox fraFileTransfer 
      Height          =   1452
      Left            =   600
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   3372
      _Version        =   786432
      _ExtentX        =   5948
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "�ļ�����"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit edStatus2 
         Height          =   252
         Left            =   0
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   3132
         _Version        =   786432
         _ExtentX        =   5524
         _ExtentY        =   444
         _StockProps     =   77
         Text            =   "�ļ�����I Love CXT.exe"
         BackColor       =   16777088
         Locked          =   -1  'True
         Appearance      =   6
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar ProgressBar 
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   16777215
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit edStatus1 
         Height          =   252
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   3132
         _Version        =   786432
         _ExtentX        =   5524
         _ExtentY        =   444
         _StockProps     =   77
         Text            =   "��ǰ���ļ���������"
         BackColor       =   16777088
         Locked          =   -1  'True
         Appearance      =   6
         FlatStyle       =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit edStatus3 
         Height          =   252
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Visible         =   0   'False
         Width           =   3132
         _Version        =   786432
         _ExtentX        =   5524
         _ExtentY        =   444
         _StockProps     =   77
         Text            =   "2333 GB/s 2333 GB/2333 GB 100%"
         BackColor       =   16777088
         Locked          =   -1  'True
         Appearance      =   6
         FlatStyle       =   -1  'True
      End
   End
   Begin VB.Timer tmrMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   1800
   End
   Begin VB.Timer tmrHide 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   2280
   End
   Begin VB.Timer tmrShow 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4080
      Top             =   1320
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00C86400&
      BorderStyle     =   0  'None
      Height          =   550
      Left            =   0
      ScaleHeight     =   552
      ScaleWidth      =   4452
      TabIndex        =   2
      Top             =   0
      Width           =   4455
      Begin VB.Image imgMainIcon 
         Height          =   500
         Left            =   0
         Picture         =   "frmBeingControlled.frx":0CCA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   500
      End
      Begin VB.Image imgClose 
         Height          =   480
         Left            =   3960
         Picture         =   "frmBeingControlled.frx":1994
         Stretch         =   -1  'True
         ToolTipText     =   "�Ͽ�����"
         Top             =   0
         Width           =   480
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Զ�̿���"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   3
         Top             =   120
         Width           =   1008
      End
   End
   Begin VB.Label labMode 
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰΪԶ�̿���ģʽ���Է����Կ������������ĵ��ԡ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Image imgShow1 
      Height          =   384
      Left            =   120
      Picture         =   "frmBeingControlled.frx":265E
      Top             =   960
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgShow2 
      Height          =   384
      Left            =   120
      Picture         =   "frmBeingControlled.frx":3328
      Top             =   1200
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgShow3 
      Height          =   384
      Left            =   120
      Picture         =   "frmBeingControlled.frx":3FF2
      Top             =   1080
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����رհ�ť���ԶϿ����ӡ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   180
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   3120
      Width           =   2532
   End
   Begin VB.Label labMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ԣ�233.233.233.233"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.8
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2475
   End
   Begin VB.Image imgClose1 
      Height          =   384
      Left            =   4080
      Picture         =   "frmBeingControlled.frx":4CBC
      Top             =   840
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgClose2 
      Height          =   384
      Left            =   4080
      Picture         =   "frmBeingControlled.frx":5986
      Top             =   720
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgClose3 
      Height          =   384
      Left            =   4080
      Picture         =   "frmBeingControlled.frx":6650
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHide3 
      Height          =   384
      Left            =   120
      Picture         =   "frmBeingControlled.frx":731A
      Top             =   840
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHide2 
      Height          =   384
      Left            =   120
      Picture         =   "frmBeingControlled.frx":7FE4
      Top             =   720
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHide1 
      Height          =   384
      Left            =   120
      Picture         =   "frmBeingControlled.frx":8CAE
      Top             =   600
      Visible         =   0   'False
      Width           =   384
   End
   Begin VB.Image imgHide 
      Height          =   384
      Left            =   0
      Picture         =   "frmBeingControlled.frx":9978
      Top             =   1920
      Width           =   384
   End
End
Attribute VB_Name = "frmBeingControlled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type POINTAPI
    x As Long
    y As Long
End Type

'��ȡ���λ�õ�API
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'�ı䴰����״�õ���API
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long

'���ô�����ǰ�ε�API
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST = -1
Private Const SWP_SHOWWINDOW = &H40
Private Const RGN_OR = 2

Dim P() As POINTAPI             '��λ��
Public IsInForm As Boolean      '��ǰ����Ƿ��ڴ�����
Public IsDown As Boolean        '����Ƿ���
Public IsDown2 As Boolean
Public n As Integer             '����ʱ�ļ��ٶ�
Public bHide As Boolean         '��ǰ�����Ƿ�����
Public bMoving As Boolean       '�����Ƿ�����ִ�в���
Public dY As Single             '��갴�´���λ�õ�Y����

Private Sub Form_Load()
    Dim n As Long
    ReDim P(6)                                                  '������Ҫ6������Χͼ��
    Me.Width = 3600                                             '�̶���С
    Me.Height = 2000
    n = 0                                                       '��ʼ�ٶ�Ϊ0
    bHide = False                                               '����Ϊ ��ʾ ״̬
    '---------------
    If IsRemoteControl = False Then                             '�������Զ�̿���ģʽ�����ļ�����ģʽ��
        Me.Height = 3500
        Me.fraFileTransfer.Visible = True
    End If
    '=========================================
    P(0).x = 500 / Screen.TwipsPerPixelX                        '�滮��״
    P(0).y = 0
    '---------------
    P(1).x = 500 / Screen.TwipsPerPixelX
    P(1).y = (Me.Height - 450 - 450) / Screen.TwipsPerPixelY
    '---------------
    P(2).x = 0
    P(2).y = (Me.Height - 450 - 450) / Screen.TwipsPerPixelY
    '---------------
    P(3).x = 0
    P(3).y = (Me.Height - 450) / Screen.TwipsPerPixelY
    '---------------
    P(4).x = 500 / Screen.TwipsPerPixelX
    P(4).y = Me.Height / Screen.TwipsPerPixelY
    '---------------
    P(5).x = Me.Width / Screen.TwipsPerPixelX
    P(5).y = Me.Height / Screen.TwipsPerPixelY
    '---------------
    P(6).x = Me.Width / Screen.TwipsPerPixelX
    P(6).y = 0
    '=========================================
    n = CreatePolygonRgn(P(0), 7, RGN_OR)                       '����������״
    SetWindowRgn Me.hwnd, n, True
    '=========================================
    Me.imgHide.Top = Me.Height - 450 - 450                      '�����ؼ�λ��
    Me.imgHide.Height = Me.Height
    Me.imgHide.Width = 500
    '---------------
    Me.imgMainIcon.Top = 50
    Me.imgMainIcon.Left = 500 + 50
    '---------------
    Me.imgClose.Left = Me.Width - Me.imgClose.Width
    Me.imgClose.Top = 0
    '---------------
    Me.labTip(0).Left = Me.imgMainIcon.Left + Me.imgMainIcon.Width + 50
    Me.labTip(0).Top = Me.imgMainIcon.Top + 50
    '---------------
    Me.labMode.Width = Me.Width - 500
    Me.labMode.Height = Me.Height - Me.labMode.Top - Me.labTip(1).Height
    '---------------
    Me.picBack.Width = Me.Width
    '---------------
    Me.labTip(1).Top = Me.Height - Me.labTip(1).Height - 120
    '---------------
    If IsRemoteControl = False Then                             '������ļ�����ģʽ������״̬��
        Me.fraFileTransfer.Height = Me.Height - Me.fraFileTransfer.Top - Me.labTip(1).Height - 50 - 120
    End If
    '=========================================
    '���ô�����ǰ��
    SetWindowPos Me.hwnd, HWND_TOPMOST, Me.CurrentX, Me.CurrentY, Me.Width / Screen.TwipsPerPixelX, Me.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW
    Me.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bHide = True Then
        Me.imgHide.Picture = Me.imgShow1.Picture
    Else
        Me.imgHide.Picture = Me.imgHide1.Picture
    End If
    Me.imgClose.Picture = Me.imgClose1.Picture
End Sub

Private Sub imgClose_Click()
    '�Ͽ�����
    frmMain.wsMessage.Close
    frmMain.wsPic.Close
    Call frmMain.wsMessage_Close
End Sub

Private Sub imgClose_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsDown2 = True
    Me.imgClose.Picture = Me.imgClose3.Picture
End Sub

Private Sub imgClose_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsInForm And Not IsDown2 Then                    '�������ڴ�����δ�������ٸı���ɫ
        Me.imgClose.Picture = Me.imgClose2.Picture
    End If
End Sub

Private Sub imgClose_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsDown2 = False
    Me.imgClose.Picture = Me.imgClose1.Picture
End Sub

Private Sub imgHide_Click()
    '================================================
    bHide = Not bHide                                   '����ģʽ���л�
    If Not bMoving Then                                         '��ǰû��չ����������
        If bHide = True Then
            bMoving = True
            Me.tmrHide.Enabled = True                           '��ʼ����
            Me.imgHide.Picture = Me.imgShow1.Picture
        Else
            bMoving = True
            Me.tmrShow.Enabled = True                           '��ʼչ��
            Me.imgHide.Picture = Me.imgShow2.Picture
        End If
    End If
End Sub

Private Sub imgHide_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsDown = True
    If bHide = True Then
        Me.imgHide.Picture = Me.imgShow3.Picture
    Else
        Me.imgHide.Picture = Me.imgHide3.Picture            '�л�ͼƬ
    End If
End Sub

Private Sub imgHide_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If IsInForm And Not IsDown Then                     '�������ڴ�����δ�������ٸı���ɫ
        If bHide = True Then
            Me.imgHide.Picture = Me.imgShow2.Picture
        Else
            Me.imgHide.Picture = Me.imgHide2.Picture
        End If
    End If
End Sub

Private Sub imgHide_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsDown = False
    If bHide = True Then
        Me.imgHide.Picture = Me.imgShow1.Picture
    Else
        Me.imgHide.Picture = Me.imgHide1.Picture
    End If
End Sub

Private Sub imgMainIcon_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim c As POINTAPI
    GetCursorPos c
    dY = y
    Me.tmrMove.Enabled = True
End Sub

Private Sub imgMainIcon_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.tmrMove.Enabled = False
End Sub

Private Sub labTip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then
        Dim c As POINTAPI
        GetCursorPos c
        dY = y
        Me.tmrMove.Enabled = True
    End If
End Sub

Private Sub labTip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Index = 0 Then
        Me.tmrMove.Enabled = False
    End If
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim c As POINTAPI
    GetCursorPos c
    dY = y
    Me.tmrMove.Enabled = True
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call Form_MouseMove(Button, Shift, x, y)
End Sub

Private Sub picBack_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.tmrMove.Enabled = False
End Sub

Private Sub tmrChangeIcon_Timer()
    Dim Cur As POINTAPI
    Dim hWindow As Long         '��ǰ��ý���Ĵ���
    '================================
    GetCursorPos Cur
    hWindow = GetForegroundWindow
    '�����������˴���ͻָ�����ͼƬ��ԭ��
    If Cur.x * Screen.TwipsPerPixelX < Me.Left Or _
       Cur.x * Screen.TwipsPerPixelX > Me.Left + Me.Width Or _
       Cur.y * Screen.TwipsPerPixelY < Me.Top Or _
       Cur.y * Screen.TwipsPerPixelY > Me.Top + Me.Height Then
        IsInForm = False
        Call Form_MouseMove(0, 0, 0, 0)
    Else
        IsInForm = True
    End If
End Sub

Private Sub tmrHide_Timer()
    n = n + 25                              '���ٶ�
    Me.Left = Me.Left + n                   '����Ļ�ұ��ƶ�
    If Me.Left > Screen.Width - 500 Then
        Me.Left = Screen.Width - 500            '����϶�ոպ�ʱ��ֹͣ�ƶ�
        bMoving = False                         '�����Ѿ�ִֹͣ�в���
        n = 0                                   '��ʼ���ٶ�
        Me.tmrHide.Enabled = False              '�ص���ʱ��
    End If
End Sub

Private Sub tmrMove_Timer()
    Dim c As POINTAPI
    GetCursorPos c
    If c.y * Screen.TwipsPerPixelY - dY < Screen.Height - Me.Height - 480 And c.y * Screen.TwipsPerPixelY - dY > 0 Then
        Me.Top = c.y * Screen.TwipsPerPixelY - dY
    Else
        If c.y * Screen.TwipsPerPixelY - dY <= 0 Then
            Me.Top = 0
        Else
            Me.Top = Screen.Height - Me.Height - 480
        End If
    End If
End Sub

Private Sub tmrShow_Timer()
    n = n + 25                              '���ٶ�
    Me.Left = Me.Left - n                   '����Ļ����ƶ�
    If Me.Left < Screen.Width - Me.Width Then
        Me.Left = Screen.Width - Me.Width       '����϶�ոպ�ʱ��ֹͣ�ƶ�
        bMoving = False                         '�����Ѿ�ִֹͣ�в���
        n = 0                                   '��ʼ���ٶ�
        Me.tmrShow.Enabled = False              '�ص���ʱ��
    End If
End Sub
