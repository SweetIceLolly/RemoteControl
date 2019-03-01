VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Begin VB.Form frmEnterPassword 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "输入密码"
   ClientHeight    =   1536
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   6132
   Icon            =   "frmEnterPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1536
   ScaleWidth      =   6132
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSuiteControls.FlatEdit edPassword 
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   2655
      _Version        =   786432
      _ExtentX        =   4683
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   16777215
      PasswordChar    =   "*"
      Appearance      =   4
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdEnter 
      Height          =   375
      Left            =   1530
      TabIndex        =   3
      Top             =   1080
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "确认"
      ForeColor       =   16777215
      BackColor       =   16416003
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
   Begin XtremeSuiteControls.PushButton cmdCancel 
      Height          =   375
      Left            =   3330
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "取消"
      ForeColor       =   16777215
      BackColor       =   16416003
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
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
   Begin VB.Label labIncorrect 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "密码错误"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请输入对方的控制密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.8
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2310
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmEnterPassword.frx":0CCA
      Stretch         =   -1  'True
      Top             =   240
      Width           =   480
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "需要密码才能控制哟~"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   240
      Index           =   1
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2430
   End
End
Attribute VB_Name = "frmEnterPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEnter_Click()
    If frmMain.optControl.Value = True Then           '如果是远控模式
        wsSendData frmRemoteControl.wsMessage, "CNT|" & Me.edPassword.Text
    Else
        wsSendData frmFileTransfer.wsMessage, "CNT|" & Me.edPassword.Text
    End If
End Sub

Private Sub edPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call cmdEnter_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '断开连接
    If IsRemoteControl Then
        frmRemoteControl.wsMessage.Close
        frmRemoteControl.wsPicture.Close
        frmMain.Show
    Else
        frmFileTransfer.wsMessage.Close
        frmMain.Show
    End If
End Sub
