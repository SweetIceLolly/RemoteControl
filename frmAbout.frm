VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Controls.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   6192
   ClientLeft      =   36
   ClientTop       =   360
   ClientWidth     =   5856
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6192
   ScaleWidth      =   5856
   StartUpPosition =   2  '屏幕中心
   Begin XtremeSuiteControls.GroupBox fraThanksTo 
      Height          =   2292
      Left            =   120
      TabIndex        =   5
      Top             =   3360
      Width           =   5652
      _Version        =   786432
      _ExtentX        =   9970
      _ExtentY        =   4043
      _StockProps     =   79
      Caption         =   "特别感谢"
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "界面是仿 TeamViewer 的"
         Height          =   180
         Index           =   7
         Left            =   0
         TabIndex        =   12
         Top             =   2040
         Width           =   2052
      End
      Begin VB.Label labURL 
         Caption         =   "http://www.ithov.com/soft/119746.shtml"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   2
         Left            =   0
         TabIndex        =   11
         Top             =   1680
         Width           =   5532
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "Codejoke Xtreme Suit Controls 控件集 来自"
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   10
         Top             =   1440
         Width           =   3876
      End
      Begin VB.Label labURL 
         Caption         =   "洛小羽 （QQ: 603678397） 指导"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   9
         Top             =   1080
         Width           =   5532
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "获取大文件的文件大小 由"
         Height          =   180
         Index           =   4
         Left            =   0
         TabIndex        =   8
         Top             =   840
         Width           =   2076
      End
      Begin VB.Label labURL 
         Caption         =   "http://blog.csdn.net/qwer430401/article/details/46636249"
         ForeColor       =   &H00C00000&
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   5532
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         Caption         =   "VB远程屏幕逐行扫描算法 来自"
         Height          =   180
         Index           =   3
         Left            =   0
         TabIndex        =   6
         Top             =   240
         Width           =   2448
      End
   End
   Begin VB.PictureBox picBack 
      BackColor       =   &H00C86400&
      BorderStyle     =   0  'None
      Height          =   732
      Left            =   0
      ScaleHeight     =   732
      ScaleWidth      =   5892
      TabIndex        =   0
      Top             =   0
      Width           =   5892
      Begin VB.Image imgMainIcon 
         Height          =   500
         Left            =   120
         Picture         =   "frmAbout.frx":0CCA
         Stretch         =   -1  'True
         Top             =   120
         Width           =   500
      End
      Begin VB.Label labTip 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "关于本远程控制"
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
         TabIndex        =   1
         Top             =   240
         Width           =   1596
      End
   End
   Begin VB.Label labURL 
      AutoSize        =   -1  'True
      Caption         =   "1257472418"
      ForeColor       =   &H00C00000&
      Height          =   180
      Index           =   3
      Left            =   4440
      TabIndex        =   14
      Top             =   5880
      Width           =   960
   End
   Begin VB.Label labTip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "感谢使用本软件！欢迎提出Bug和宝贵意见，作者QQ："
      Height          =   180
      Index           =   8
      Left            =   120
      TabIndex        =   13
      Top             =   5880
      Width           =   4260
   End
   Begin VB.Label labTip 
      BackStyle       =   0  'Transparent
      Caption         =   "  最后向我们热爱的VB 致敬！"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2652
   End
   Begin VB.Label labTip 
      BackStyle       =   0  'Transparent
      Caption         =   "致 敬爱的用户"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1452
   End
   Begin VB.Label labTip 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":1994
      Height          =   1740
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5616
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub labURL_Click(Index As Integer)
    Select Case Index
        Case 0
            Shell "Explorer http://blog.csdn.net/qwer430401/article/details/46636249", vbNormalFocus
        
        Case 1
            Clipboard.Clear
            Clipboard.SetText "603678397"
            MsgBox "已将QQ号复制到剪贴板。", 64, "提示"
        
        Case 2
            Shell "Explorer http://www.ithov.com/soft/119746.shtml", vbNormalFocus
        
        Case 3
            Clipboard.Clear
            Clipboard.SetText "1257472418"
            MsgBox "已将QQ号复制到剪贴板。", 64, "提示"
            
    End Select
End Sub
